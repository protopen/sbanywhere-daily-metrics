"""
Surebright Anywhere Traffic & Campaign Insights - Supabase Streamlit App

Run locally:
  pip install -r requirements_streamlit_d2c.txt
  streamlit run streamlit_d2c_metrics_app.py

Keep this file in the same folder as d2c_clean_external_metrics_report.py.
"""

from __future__ import annotations

import base64
import io
import json
import re
import tempfile
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Optional, Tuple
from urllib.parse import parse_qs, urlencode, urlparse, urlunparse

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st
from supabase import create_client

try:
    from zoneinfo import ZoneInfo
except Exception:  # pragma: no cover
    ZoneInfo = None

from d2c_clean_external_metrics_report import (
    build_daily_metrics,
    build_totals,
    build_attribution_campaign_stats,
    build_attribution_adset_stats,
    build_attribution_ad_stats,
    build_utm_event_breakdown,
    build_product_stats,
    build_sales_stats,
    build_high_intent_dropoffs,
    build_retailer_stats,
    normalize_all_events,
)

APP_TITLE = "Surebright Anywhere Traffic & Campaign Insights"
DEFAULT_START_DATE = date(2026, 5, 21)
DEFAULT_TIMEZONE = "Asia/Kolkata"
DEFAULT_SUPABASE_TABLE = "d2c_raw_events"
DEFAULT_SUPABASE_JSON_COLUMN = "raw"
DEFAULT_SUPABASE_TIMESTAMP_COLUMN = "occurred_at"
DEFAULT_SUPABASE_SELECT_COLUMNS = "*"
DEFAULT_SUPABASE_PAGE_SIZE = 1000
LOGO_FILENAME = "surebright_logo_homepage.webp"


def get_logo_path() -> Optional[Path]:
    possible_paths = [
        Path(__file__).with_name(LOGO_FILENAME),
        Path.cwd() / LOGO_FILENAME,
        Path("/mnt/data") / LOGO_FILENAME,
    ]
    for path in possible_paths:
        if path.exists():
            return path
    return None


def safe_stem(filename: str) -> str:
    stem = Path(filename).stem or "d2c_source"
    return re.sub(r"[^A-Za-z0-9_\-]+", "_", stem).strip("_") or "d2c_source"


def date_to_str(value: Optional[date]) -> Optional[str]:
    return value.isoformat() if value else None


def dataframe_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")




def add_or_replace_query_param(url: str, key: str, value: str) -> str:
    parsed = urlparse(url)
    query = parse_qs(parsed.query, keep_blank_values=True)
    query[key] = [value]
    new_query = urlencode(query, doseq=True)
    return urlunparse(parsed._replace(query=new_query))


def encode_onedrive_share_url(url: str) -> str:
    """Encode a OneDrive/SharePoint sharing URL for the Microsoft sharing API."""
    encoded = base64.urlsafe_b64encode(url.encode("utf-8")).decode("utf-8").rstrip("=")
    return "u!" + encoded


def build_onedrive_download_candidates(url: str) -> list[str]:
    """Return likely direct-download endpoints for public OneDrive/SharePoint sharing links."""
    url = url.strip()
    candidates = []
    if not url:
        return candidates

    # 1) Many OneDrive share links download when download=1 is present.
    candidates.append(add_or_replace_query_param(url, "download", "1"))

    # 2) Personal OneDrive links sometimes expose a download.aspx endpoint from resid/authkey.
    parsed = urlparse(url)
    qs = parse_qs(parsed.query)
    resid = (qs.get("resid") or qs.get("id") or [None])[0]
    authkey = (qs.get("authkey") or [None])[0]
    if resid and "onedrive.live.com" in parsed.netloc:
        params = {"resid": resid}
        if authkey:
            params["authkey"] = authkey
        candidates.append("https://onedrive.live.com/download.aspx?" + urlencode(params))

    # 3) Public sharing API. This is useful for some anonymous/shared links.
    share_token = encode_onedrive_share_url(url)
    candidates.append(f"https://api.onedrive.com/v1.0/shares/{share_token}/root/content")
    candidates.append(f"https://graph.microsoft.com/v1.0/shares/{share_token}/driveItem/content")

    # Preserve order while removing duplicates.
    deduped = []
    for candidate in candidates:
        if candidate not in deduped:
            deduped.append(candidate)
    return deduped


def guess_filename_from_response(response: requests.Response, fallback_url: str) -> str:
    cd = response.headers.get("content-disposition", "")
    match = re.search(r'filename\*?=(?:UTF-8\'\')?"?([^";]+)', cd, re.IGNORECASE)
    if match:
        return Path(match.group(1)).name
    path_name = Path(urlparse(response.url or fallback_url).path).name
    if path_name and "." in path_name:
        return path_name
    fallback_name = Path(urlparse(fallback_url).path).name
    return fallback_name if fallback_name and "." in fallback_name else "onedrive_source.xlsx"


def looks_like_html_or_json_error(content: bytes, content_type: str) -> tuple[bool, str]:
    """Return True when OneDrive returned a web page or API error, not the actual data file."""
    head = content[:4096].lstrip()
    lower_head = head[:1000].lower()

    if b"<html" in lower_head or b"<!doctype html" in lower_head or b"<body" in lower_head:
        return True, "download returned an HTML page instead of the file"

    # Some OneDrive/Graph endpoints return application/json error/metadata bodies.
    # Do not pass those to pandas as CSV; it creates ParserError: ',' expected after '"'.
    if "application/json" in content_type or head.startswith(b"{"):
        text = head.decode("utf-8", errors="ignore").lower()
        if '"error"' in text or '"message"' in text or '"@odata' in text:
            return True, "download returned JSON metadata/error instead of file content"

    return False, ""


def fetch_onedrive_file(url: str, access_token: Optional[str] = None) -> Tuple[bytes, str, str]:
    """
    Download an Excel/CSV/text source from OneDrive/SharePoint as raw bytes.

    For public/anyone links, leave access_token blank. For private business files,
    paste a Microsoft Graph access token or set it in Streamlit secrets.
    """
    headers = {
        "User-Agent": "SureBright-D2C-Metrics-App/1.0",
        "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel, text/csv, text/plain, application/octet-stream, */*",
    }
    if access_token:
        headers["Authorization"] = f"Bearer {access_token.strip()}"

    last_error = None
    for candidate in build_onedrive_download_candidates(url):
        try:
            response = requests.get(candidate, headers=headers, timeout=60, allow_redirects=True)
            content_type = response.headers.get("content-type", "").lower()
            content = response.content or b""
            bad, reason = looks_like_html_or_json_error(content, content_type)
            if response.ok and content and not bad:
                filename = guess_filename_from_response(response, url)

                # If the URL says .xlsx but the returned content is really CSV/text, use a text suffix.
                head = content[:16].lstrip()
                if not head.startswith(b"PK") and filename.lower().endswith((".xlsx", ".xls")):
                    if b"," in content[:1024] or head.startswith(b"{"):
                        filename = Path(filename).with_suffix(".csv").name

                return content, filename, candidate

            last_error = f"{response.status_code} from {candidate}; content-type={content_type or 'unknown'}; {reason or 'empty/unusable response'}"
        except Exception as exc:  # pragma: no cover
            last_error = f"{candidate}: {exc}"

    raise RuntimeError(
        "Could not download the actual OneDrive file bytes. Make sure the link points directly to a file, "
        "not a folder or preview page. Set sharing to 'Anyone with the link can view', or provide a Microsoft Graph token. "
        f"Last error: {last_error}"
    )


def build_excel_bytes(
    daily: pd.DataFrame,
    totals: pd.DataFrame,
    attribution_campaign: pd.DataFrame,
    attribution_adset: pd.DataFrame,
    attribution_ad: pd.DataFrame,
    breakdown_campaign: pd.DataFrame,
    breakdown_adset: pd.DataFrame,
    breakdown_ad: pd.DataFrame,
    product: pd.DataFrame,
    sales: pd.DataFrame,
    high_intent_signup: pd.DataFrame,
    high_intent_invoice: pd.DataFrame,
    retailer: pd.DataFrame,
    audit_df: pd.DataFrame,
    metadata: pd.DataFrame,
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        daily.to_excel(writer, index=False, sheet_name="Daily Metrics")
        totals.to_excel(writer, index=False, sheet_name="Totals")
        attribution_campaign.to_excel(writer, index=False, sheet_name="Campaign Attribution")
        attribution_adset.to_excel(writer, index=False, sheet_name="Ad Set Attribution")
        attribution_ad.to_excel(writer, index=False, sheet_name="Ad Attribution")
        breakdown_campaign.to_excel(writer, index=False, sheet_name="Campaign Metric Breakdown")
        breakdown_adset.to_excel(writer, index=False, sheet_name="Ad Set Metric Breakdown")
        breakdown_ad.to_excel(writer, index=False, sheet_name="Ad Metric Breakdown")
        product.to_excel(writer, index=False, sheet_name="Product Stats")
        sales.to_excel(writer, index=False, sheet_name="Sales")
        high_intent_signup.to_excel(writer, index=False, sheet_name="Signup Dropoffs")
        high_intent_invoice.to_excel(writer, index=False, sheet_name="Invoice Dropoffs")
        retailer.to_excel(writer, index=False, sheet_name="Retailer Stats")
        audit_df.to_excel(writer, index=False, sheet_name="Event Audit")
        metadata.to_excel(writer, index=False, sheet_name="Run Metadata")

        wb = writer.book
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True, color="FFFFFF")
                cell.fill = cell.fill.copy(fill_type="solid", fgColor="1F4E78")
                cell.alignment = cell.alignment.copy(horizontal="center")

            for column_cells in ws.columns:
                values = [str(cell.value or "") for cell in column_cells[:300]]
                width = min(max(len(v) for v in values) + 2, 55)
                ws.column_dimensions[column_cells[0].column_letter].width = max(10, width)

            for cell in ws[1]:
                if cell.value and ("GWP" in str(cell.value) or "Price" in str(cell.value)):
                    for data_cell in ws[cell.column_letter][1:]:
                        data_cell.number_format = "$#,##0.00"
                if cell.value and "%" in str(cell.value):
                    for data_cell in ws[cell.column_letter][1:]:
                        data_cell.number_format = "0.00%"
            if ws.title == "Totals":
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                    if row[0].value == "Gross GWP $":
                        row[1].number_format = "$#,##0.00"

    output.seek(0)
    return output.read()


def get_secret_value(*keys: str, default: str = "") -> str:
    """Read nested Streamlit secrets safely, returning default when missing."""
    try:
        current = st.secrets
        for key in keys:
            current = current[key]
        return str(current)
    except Exception:
        return default


def get_secret_int(*keys: str, default: int = 1000) -> int:
    try:
        return int(get_secret_value(*keys, default=str(default)))
    except Exception:
        return default


@st.cache_data(show_spinner=False, ttl=300)
def fetch_supabase_rows(
    supabase_url: str,
    supabase_key: str,
    table_name: str,
    select_columns: str,
    timestamp_column: str,
    start_date: Optional[str],
    end_date: Optional[str],
    page_size: int,
) -> list[dict]:
    """Fetch rows from Supabase with simple pagination.

    Expected table shape: one row per event, with a JSON/JSONB column containing the
    full event payload. Default JSON column is configured separately as `raw`.
    """
    if not supabase_url or not supabase_key:
        raise RuntimeError("Supabase URL/key are missing. Add them to .streamlit/secrets.toml or Streamlit Cloud secrets.")
    if not table_name:
        raise RuntimeError("Supabase table name is missing.")

    client = create_client(supabase_url, supabase_key)
    rows: list[dict] = []
    offset = 0
    page_size = max(1, min(int(page_size or 1000), 5000))

    while True:
        query = client.table(table_name).select(select_columns or "*")
        if timestamp_column:
            # UTC boundaries are only used to reduce DB volume. Add a one-day buffer
            # on both sides so timezone bucketing in the app does not drop edge events.
            if start_date:
                start_buffer = (date.fromisoformat(start_date) - timedelta(days=1)).isoformat()
                query = query.gte(timestamp_column, f"{start_buffer}T00:00:00+00:00")
            if end_date:
                end_buffer = (date.fromisoformat(end_date) + timedelta(days=1)).isoformat()
                query = query.lte(timestamp_column, f"{end_buffer}T23:59:59+00:00")
            query = query.order(timestamp_column)
        response = query.range(offset, offset + page_size - 1).execute()
        data = response.data or []
        rows.extend(data)
        if len(data) < page_size:
            break
        offset += page_size

    return rows


def supabase_rows_to_jsonl_bytes(rows: list[dict], json_column: str) -> bytes:
    """Convert Supabase rows into JSONL bytes compatible with the existing parser."""
    lines: list[str] = []
    json_column = (json_column or "").strip()

    for row in rows:
        payload = None
        if json_column and json_column in row:
            payload = row.get(json_column)
        elif "Raw" in row:
            payload = row.get("Raw")
        elif "raw" in row:
            payload = row.get("raw")
        elif "payload" in row:
            payload = row.get("payload")
        elif "event_payload" in row:
            payload = row.get("event_payload")
        else:
            payload = row

        if payload is None:
            continue
        if isinstance(payload, str):
            payload = payload.strip()
            if not payload:
                continue
            lines.append(payload)
        else:
            lines.append(json.dumps(payload, default=str, ensure_ascii=False))

    return ("\n".join(lines) + "\n").encode("utf-8") if lines else b""


@st.cache_data(show_spinner=False)
def process_uploaded_file(
    file_bytes: bytes,
    filename: str,
    timezone_name: str,
    start_date: Optional[str],
    end_date: Optional[str],
) -> dict:
    suffix = Path(filename).suffix.lower() or ".txt"
    with tempfile.TemporaryDirectory() as temp_dir:
        input_path = Path(temp_dir) / f"source{suffix}"
        input_path.write_bytes(file_bytes)

        if ZoneInfo:
            tz = ZoneInfo(timezone_name)
        else:
            tz = None

        all_events = normalize_all_events(input_path, tz)
        clean_events = [event for event in all_events if not event.excluded]
        if start_date:
            clean_events = [event for event in clean_events if event.date >= start_date]
        if end_date:
            clean_events = [event for event in clean_events if event.date <= end_date]

        daily = build_daily_metrics(clean_events, start_date=None, end_date=None)
        totals = build_totals(daily)
        attribution_campaign = build_attribution_campaign_stats(clean_events)
        attribution_adset = build_attribution_adset_stats(clean_events)
        attribution_ad = build_attribution_ad_stats(clean_events)
        breakdown_campaign = build_utm_event_breakdown(clean_events, level="campaign")
        breakdown_adset = build_utm_event_breakdown(clean_events, level="ad_set")
        breakdown_ad = build_utm_event_breakdown(clean_events, level="ad")
        product = build_product_stats(clean_events)
        product = product.drop(columns=["Product Events"], errors="ignore")
        sales = build_sales_stats(clean_events)
        high_intent_signup, high_intent_invoice, high_intent_unique_dropoffs = build_high_intent_dropoffs(clean_events)
        retailer = build_retailer_stats(clean_events)
        audit_df = pd.DataFrame([event.__dict__ for event in all_events])
        clean_audit_df = pd.DataFrame([event.__dict__ for event in clean_events])

        metadata = pd.DataFrame(
            [
                ["source_file", filename],
                ["timezone", timezone_name],
                ["start_date_inclusive", start_date or "all"],
                ["end_date_inclusive", end_date or "all"],
                ["parseable_events", len(all_events)],
                ["clean_external_events_in_window", len(clean_events)],
                ["gross_gwp_total", round(float(daily["Gross GWP $"].sum()), 2) if not daily.empty else 0.0],
                ["clean_external_definition", "prod surebrightanywhere.com only, excluding internal/test identities, test domains, and test URLs"],
                ["generated_at_local", datetime.now().isoformat(timespec="seconds")],
            ],
            columns=["Field", "Value"],
        )

        excel_bytes = build_excel_bytes(
            daily, totals, attribution_campaign, attribution_adset, attribution_ad,
            breakdown_campaign, breakdown_adset, breakdown_ad, product, sales, high_intent_signup, high_intent_invoice, retailer, audit_df, metadata
        )

    return {
        "daily": daily,
        "totals": totals,
        "attribution_campaign": attribution_campaign,
        "attribution_adset": attribution_adset,
        "attribution_ad": attribution_ad,
        "breakdown_campaign": breakdown_campaign,
        "breakdown_adset": breakdown_adset,
        "breakdown_ad": breakdown_ad,
        "product": product,
        "sales": sales,
        "high_intent_signup": high_intent_signup,
        "high_intent_invoice": high_intent_invoice,
        "high_intent_unique_dropoffs": high_intent_unique_dropoffs,
        "retailer": retailer,
        "audit": audit_df,
        "clean_audit": clean_audit_df,
        "metadata": metadata,
        "excel_bytes": excel_bytes,
        "parseable_count": len(all_events),
        "clean_count": len(clean_events),
    }



# -----------------------------
# Daily Metrics Sankey
# -----------------------------

SANKEY_STAGE_EVENTS = {
    "Enquiry Attempted": {"homepage_form_submit", "enquiry_attempted"},
    "Sign Up_total": {"lead_signup", "quote_lead_captured", "signup_completed", "sign_up_completed"},
    "First Quote_Success": {"quote_generated", "first_quote_success", "offer_generation_success"},
    "Offer_Selected": {"plan_selected", "offer_selected"},
    "Invoice Upload_Success": {"invoice_uploaded", "invoice_upload_success"},
    "Invoice Upload_Failure": {"invoice_upload_failed", "invoice_upload_failure"},
    "Revised Offer": {"revised_offer_shown", "revised_offer"},
    "Additional Product": {"additional_product_detected", "additional_product"},
    "Add to Cart_Success": {"cart_confirmed", "add_to_cart"},
    "Payment Attempted": {"pay_now_clicked", "initiate_checkout"},
    "Payment Success": {"payment_completed", "payment_success", "payment_succeeded", "purchase"},
    "Payment Failed": {"payment_failed"},
}

SANKEY_MAIN_PATH = [
    "Enquiry Attempted",
    "Sign Up_total",
    "First Quote_Success",
    "Offer_Selected",
    "Invoice Upload_Success",
    "Add to Cart_Success",
    "Payment Attempted",
    "Payment Success",
]


def _stage_for_event_name(event_name: str) -> Optional[str]:
    event_name = str(event_name or "").strip().lower()
    for stage, event_names in SANKEY_STAGE_EVENTS.items():
        if event_name in event_names:
            return stage
    return None


def build_sankey_links_from_audit(clean_audit: pd.DataFrame) -> pd.DataFrame:
    """Build user-progression Sankey links from clean event-level audit rows."""
    if clean_audit.empty or "event_name" not in clean_audit.columns:
        return pd.DataFrame(columns=["source", "target", "value"])

    df = clean_audit.copy()
    df["stage"] = df["event_name"].map(_stage_for_event_name)
    df = df[df["stage"].notna()]
    if df.empty:
        return pd.DataFrame(columns=["source", "target", "value"])

    if "identity_key" not in df.columns:
        df["identity_key"] = "unknown"
    if "event_time" in df.columns:
        df = df.sort_values(["identity_key", "event_time"])

    link_counts: dict[tuple[str, str], int] = {}

    for _, g in df.groupby("identity_key", dropna=False):
        reached = set(g["stage"].dropna().astype(str).tolist())
        ordered_reached = [stage for stage in SANKEY_MAIN_PATH if stage in reached]

        # Connect each reached business stage to the next reached business stage for that user.
        for source, target in zip(ordered_reached, ordered_reached[1:]):
            link_counts[(source, target)] = link_counts.get((source, target), 0) + 1

        # Add terminal failure branches without double-counting users who eventually succeeded.
        if "Invoice Upload_Failure" in reached and "Invoice Upload_Success" not in reached:
            previous = None
            for stage in ["Offer_Selected", "First Quote_Success", "Sign Up_total", "Enquiry Attempted"]:
                if stage in reached:
                    previous = stage
                    break
            if previous:
                link_counts[(previous, "Invoice Upload_Failure")] = link_counts.get((previous, "Invoice Upload_Failure"), 0) + 1

        if "Payment Failed" in reached and "Payment Success" not in reached:
            previous = None
            for stage in ["Payment Attempted", "Add to Cart_Success", "Invoice Upload_Success", "Offer_Selected"]:
                if stage in reached:
                    previous = stage
                    break
            if previous:
                link_counts[(previous, "Payment Failed")] = link_counts.get((previous, "Payment Failed"), 0) + 1

    rows = [{"source": s, "target": t, "value": v} for (s, t), v in link_counts.items() if v > 0]
    if not rows:
        return pd.DataFrame(columns=["source", "target", "value"])

    out = pd.DataFrame(rows)
    stage_order = {stage: i for i, stage in enumerate(SANKEY_MAIN_PATH + ["Invoice Upload_Failure", "Payment Failed"])}
    out["_sort"] = out["source"].map(stage_order).fillna(999)
    out = out.sort_values(["_sort", "source", "target"]).drop(columns=["_sort"]).reset_index(drop=True)
    return out



def build_daily_metrics_summary_table(daily: pd.DataFrame) -> pd.DataFrame:
    """Daily compact table requested for the Daily Metrics tab."""
    if daily.empty:
        return pd.DataFrame(
            columns=[
                "Date",
                "Enquiry Attempted",
                "Sign Up_total",
                "Add to cart",
                "Invoice Upload_Success",
                "revised_offer_shown",
                "Initiate Checkout",
                "Payment Success",
            ]
        )

    out = pd.DataFrame()
    out["Date"] = daily["Date"] if "Date" in daily.columns else ""
    out["Enquiry Attempted"] = daily.get("Enquiry Attempted", 0)
    out["Sign Up_total"] = daily.get("Sign Up_total", 0)
    out["Add to cart"] = daily.get("Add to Cart_Success", 0)
    out["Invoice Upload_Success"] = daily.get("Invoice Upload_Success", 0)
    out["revised_offer_shown"] = daily.get("Revised Offer", 0)
    out["Initiate Checkout"] = daily.get("Payment Attempted", 0)
    out["Payment Success"] = daily.get("Payment Success", 0)
    return out


def build_detailed_total_table(daily: pd.DataFrame, clean_audit: pd.DataFrame) -> pd.DataFrame:
    """Detailed total table requested for the Daily Metrics tab."""
    def total(col: str) -> float:
        if daily.empty or col not in daily.columns:
            return 0
        return float(daily[col].sum())

    if clean_audit is not None and not clean_audit.empty and "identity_key" in clean_audit.columns:
        user_count = int(clean_audit["identity_key"].dropna().nunique())
    else:
        user_count = 0

    invoice_upload_attempted = total("Invoice Upload_Success") + total("Invoice Upload_Failure")
    payment_attempted = total("Payment Attempted")

    rows = [
        ("User Count", user_count),
        ("Enquiry Attempted", total("Enquiry Attempted")),
        ("Sign Up_total", total("Sign Up_total")),
        ("Add to cart", total("Add to Cart_Success")),
        ("Invoice Upload_Attempted", invoice_upload_attempted),
        ("Invoice Upload_Success", total("Invoice Upload_Success")),
        ("Invoice Upload_Failure", total("Invoice Upload_Failure")),
        ("revised_offer_shown", total("Revised Offer")),
        ("Initiate Checkout", payment_attempted),
        ("Payment Attempted", payment_attempted),
        ("Payment Success", total("Payment Success")),
        ("Gross GWP $", round(total("Gross GWP $"), 2)),
    ]

    out = pd.DataFrame(rows, columns=["Metric", "Value"])
    # Keep integer-looking values clean while preserving currency decimals.
    out["Value"] = out.apply(
        lambda r: f"${float(r['Value']):,.2f}" if r["Metric"] == "Gross GWP $" else int(r["Value"]),
        axis=1,
    )
    return out


def build_product_count_table(daily: pd.DataFrame) -> pd.DataFrame:
    """Product count table requested for the Daily Metrics tab."""
    def total(col: str) -> int:
        if daily.empty or col not in daily.columns:
            return 0
        return int(daily[col].sum())

    return pd.DataFrame(
        [
            ("Invoice Success_Product Count", total("Invoice Success_Product Count")),
            ("Add to Cart_Success Count", total("Add to Cart_Success Count")),
            ("Payment Success_ Count", total("Payment Success_Count")),
        ],
        columns=["Metric", "Value"],
    )


def render_daily_sankey(clean_audit: pd.DataFrame) -> None:
    sankey_links = build_sankey_links_from_audit(clean_audit)
    if sankey_links.empty:
        st.info("Not enough clean external event progression data to draw a Sankey chart for this date range.")
        return

    labels = list(dict.fromkeys(sankey_links["source"].tolist() + sankey_links["target"].tolist()))
    label_to_idx = {label: idx for idx, label in enumerate(labels)}

    fig = go.Figure(
        data=[
            go.Sankey(
                arrangement="snap",
                node=dict(
                    pad=22,
                    thickness=18,
                    line=dict(width=0.5),
                    label=labels,
                ),
                link=dict(
                    source=[label_to_idx[x] for x in sankey_links["source"]],
                    target=[label_to_idx[x] for x in sankey_links["target"]],
                    value=sankey_links["value"].tolist(),
                ),
            )
        ]
    )
    fig.update_layout(
        title_text="Clean External funnel journey",
        height=620,
        font=dict(size=13, color="#FFFFFF"),
        margin=dict(l=10, r=10, t=50, b=10),
    )
    st.plotly_chart(fig, use_container_width=True)

    with st.expander("View Sankey source data"):
        st.dataframe(sankey_links, use_container_width=True, hide_index=True)


def show_kpis(daily: pd.DataFrame) -> None:
    gross_gwp = float(daily["Gross GWP $"].sum()) if not daily.empty and "Gross GWP $" in daily else 0.0
    payment_success = int(daily["Payment Success"].sum()) if not daily.empty and "Payment Success" in daily else 0
    payment_attempted = int(daily["Payment Attempted"].sum()) if not daily.empty and "Payment Attempted" in daily else 0
    enquiry_attempted = int(daily["Enquiry Attempted"].sum()) if not daily.empty and "Enquiry Attempted" in daily else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Enquiry Attempted", f"{enquiry_attempted:,}")
    c2.metric("Payment Attempted", f"{payment_attempted:,}")
    c3.metric("Payment Success", f"{payment_success:,}")
    c4.metric("Gross GWP", f"${gross_gwp:,.2f}")

    if payment_attempted:
        st.caption(f"Payment success rate: {payment_success / payment_attempted:.1%} based on clean external payment-attempt users.")


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="📊", layout="wide")

    logo_path = get_logo_path()
    if logo_path is not None:
        st.image(str(logo_path), width=360)
    st.title(APP_TITLE)

    with st.sidebar:
        st.header("Supabase connection")

        supabase_url = get_secret_value("supabase", "url")
        supabase_key = get_secret_value("supabase", "key")
        table_name = get_secret_value("supabase", "table", default=DEFAULT_SUPABASE_TABLE)
        json_column = get_secret_value("supabase", "json_column", default=DEFAULT_SUPABASE_JSON_COLUMN)
        timestamp_column = get_secret_value("supabase", "timestamp_column", default=DEFAULT_SUPABASE_TIMESTAMP_COLUMN)
        select_columns = get_secret_value("supabase", "select_columns", default=DEFAULT_SUPABASE_SELECT_COLUMNS)
        page_size = get_secret_int("supabase", "page_size", default=DEFAULT_SUPABASE_PAGE_SIZE)

        if supabase_url and supabase_key:
            st.success("Supabase secrets loaded")
        else:
            st.warning("Add Supabase credentials in `.streamlit/secrets.toml`.")

        with st.expander("Advanced table settings", expanded=False):
            table_name = st.text_input("Table name", value=table_name)
            json_column = st.text_input("JSON payload column", value=json_column)
            timestamp_column = st.text_input("Timestamp column for DB filtering", value=timestamp_column)
            select_columns = st.text_input("Select columns", value=select_columns)
            page_size = st.number_input("Fetch page size", min_value=100, max_value=5000, value=int(page_size), step=100)

        st.divider()
        st.header("Report settings")
        timezone_name = st.text_input("Timezone for date bucketing", value=DEFAULT_TIMEZONE)
        start_date_value = st.date_input("Start date", value=DEFAULT_START_DATE)
        use_end_date = st.checkbox("Use end date", value=False)
        end_date_value = st.date_input("End date", value=date.today()) if use_end_date else None

        st.divider()
        st.caption("Clean External filter")
        st.markdown(
            "- Includes `surebrightanywhere.com` production traffic\n"
            "- Excludes Abhishek, Santosh\n"
            "- Excludes `@surebright.com`, `@surerbright.com`, `@example.com`\n"
            "- Excludes localhost, Webflow, Amplify, and `_meta_test=1`"
        )

    start_date_str = date_to_str(start_date_value)
    end_date_str = date_to_str(end_date_value)

    try:
        with st.spinner("Fetching source events from Supabase..."):
            rows = fetch_supabase_rows(
                supabase_url=supabase_url,
                supabase_key=supabase_key,
                table_name=table_name,
                select_columns=select_columns,
                timestamp_column=timestamp_column,
                start_date=start_date_str,
                end_date=end_date_str,
                page_size=int(page_size),
            )
            file_bytes = supabase_rows_to_jsonl_bytes(rows, json_column=json_column)
            source_filename = f"{table_name}_supabase.jsonl"

        if not file_bytes:
            st.warning("Supabase query returned no event payloads for the selected settings/date range.")
            st.stop()

        with st.spinner("Processing clean external metrics..."):
            result = process_uploaded_file(
                file_bytes=file_bytes,
                filename=source_filename,
                timezone_name=timezone_name,
                start_date=start_date_str,
                end_date=end_date_str,
            )
    except Exception as exc:
        st.error("Could not fetch or process Supabase data. Check secrets, table name, JSON column, timestamp column, and table permissions.")
        st.exception(exc)
        st.stop()

    daily = result["daily"]
    totals = result["totals"]
    audit = result["audit"]
    clean_audit = result["clean_audit"]

    show_kpis(daily)

    attribution_campaign = result["attribution_campaign"]
    attribution_adset = result["attribution_adset"]
    attribution_ad = result["attribution_ad"]
    breakdown_campaign = result["breakdown_campaign"]
    breakdown_adset = result["breakdown_adset"]
    breakdown_ad = result["breakdown_ad"]
    product = result["product"]
    sales = result["sales"]
    high_intent_signup = result["high_intent_signup"]
    high_intent_invoice = result["high_intent_invoice"]
    high_intent_unique_dropoffs = result["high_intent_unique_dropoffs"]
    retailer = result["retailer"]

    tab_daily, tab_totals, tab_attribution, tab_product, tab_sales, tab_high_intent, tab_retailer, tab_audit, tab_downloads = st.tabs(
        ["Daily Metrics", "Totals", "Attribution", "Product", "Sales", "High Intent", "Retailer", "Event Audit", "Downloads"]
    )

    with tab_daily:
        st.subheader("Daily Metrics")
        if daily.empty:
            st.warning("No clean external events matched the selected date range.")
        else:
            daily_tables_tab, daily_sankey_tab = st.tabs(["Tables", "Sankey"])

            with daily_tables_tab:
                st.markdown("#### Daily")
                st.dataframe(
                    build_daily_metrics_summary_table(daily),
                    use_container_width=True,
                    hide_index=True,
                )

                st.markdown("#### Detailed / Total")
                st.dataframe(
                    build_detailed_total_table(daily, clean_audit),
                    use_container_width=True,
                    hide_index=True,
                )

                st.markdown("#### Product Count")
                st.dataframe(
                    build_product_count_table(daily),
                    use_container_width=True,
                    hide_index=True,
                )

                with st.expander("View full Daily Metrics table"):
                    st.dataframe(daily, use_container_width=True, hide_index=True)

            with daily_sankey_tab:
                st.markdown("#### Funnel Sankey")
                st.caption("Shows unique clean external users progressing between funnel stages in the selected date range. Node label text uses a bright color for readability.")
                render_daily_sankey(clean_audit)

    with tab_totals:
        st.subheader("Totals")
        st.dataframe(totals, use_container_width=True, hide_index=True)

    with tab_attribution:
        st.subheader("Attribution")
        st.caption("UTM/ad-click traffic only. Campaign = `utm_campaign`, Ad set = ad-set fields or `utm_term`, Ad = ad fields or `utm_content`.")

        campaign_tab, adset_tab, ad_tab = st.tabs(["Campaign level", "Ad set level", "Ads level"])

        def show_attribution_level(summary_df: pd.DataFrame, breakdown_df: pd.DataFrame, label_col: str) -> None:
            st.markdown("#### Summary")
            if summary_df.empty:
                st.warning("No UTM/ad-click attribution rows available for this date range.")
            else:
                st.dataframe(summary_df, use_container_width=True, hide_index=True)

            st.markdown("#### Metric breakdown")
            st.caption("Uses the same business metric labels as Daily Metrics, but only for UTM/ad-click attributed traffic.")
            if breakdown_df.empty:
                st.warning("No UTM/ad-click metric breakdown rows available for this date range.")
            else:
                st.dataframe(breakdown_df, use_container_width=True, hide_index=True)
                event_cols = [
                    c for c in [
                        "Enquiry Attempted",
                        "Sign Up_total",
                        "First Quote_Success",
                        "Offer_Selected",
                        "Add to Cart_Success",
                        "Payment Attempted",
                        "Payment Success",
                        "Payment Failed",
                    ]
                    if c in breakdown_df.columns
                ]
                if event_cols and label_col in breakdown_df.columns:
                    st.bar_chart(breakdown_df.set_index(label_col)[event_cols])

        with campaign_tab:
            show_attribution_level(attribution_campaign, breakdown_campaign, "Campaign Label")

        with adset_tab:
            show_attribution_level(attribution_adset, breakdown_adset, "Ad set Label")

        with ad_tab:
            show_attribution_level(attribution_ad, breakdown_ad, "Ad Label")

    with tab_product:
        st.subheader("Product Stats")
        st.caption("Explodes `event_data.line_items` and also includes form-only enquiry category rows.")
        if product.empty:
            st.warning("No product rows available for this date range.")
        else:
            st.dataframe(product, use_container_width=True, hide_index=True)
            if "Product Category" in product.columns and "Enquiry Product Count" in product.columns:
                st.bar_chart(product.set_index("Product Category")[["Enquiry Product Count"]])


    with tab_sales:
        st.subheader("Sales")
        st.caption("Payment Success events only. This table excludes PII other than email ID.")
        if sales.empty:
            st.warning("No payment success rows available for this date range.")
        else:
            st.dataframe(sales, use_container_width=True, hide_index=True)


    with tab_high_intent:
        st.subheader("High Intent Dropoffs")
        st.caption("Session-based dropoffs. This tab intentionally includes PII for follow-up workflows. User Key is hidden from both tables.")
        st.metric("Unique high-intent dropoff users", high_intent_unique_dropoffs)

        high_signup_tab, high_invoice_tab = st.tabs(["Dropped after Sign Up_total", "Dropped after invoice upload or later"])

        with high_signup_tab:
            st.markdown("#### Dropped after Sign Up_total")
            st.caption("Sessions that reached Sign Up_total but did not upload an invoice, add to cart, initiate checkout, or complete payment. Shows only first event date, last event date, Name, Email, last stage, product category, and product price.")
            if high_intent_signup.empty:
                st.warning("No Sign Up_total dropoffs found for this date range.")
            else:
                st.dataframe(high_intent_signup, use_container_width=True, hide_index=True)

        with high_invoice_tab:
            st.markdown("#### Dropped after invoice upload or later")
            st.caption("Sessions that uploaded an invoice or reached a later step but did not complete payment. Shows first event date, last event date, Name, Email, last stage, and remaining invoice/product fields. Keeps full line_items object only in Line Items Data.")
            if high_intent_invoice.empty:
                st.warning("No invoice-or-later dropoffs found for this date range.")
            else:
                st.dataframe(high_intent_invoice, use_container_width=True, hide_index=True)

    with tab_retailer:
        st.subheader("Retailer Stats")
        st.caption("Uses revised `event_data.invoice.retailer` fields and legacy `retailer_name` / `retailer_detected` fallbacks.")
        if retailer.empty:
            st.warning("No retailer rows available for this date range.")
        else:
            st.dataframe(retailer, use_container_width=True, hide_index=True)
            if "Retailer Name" in retailer.columns and "Invoice Upload_Success" in retailer.columns:
                st.bar_chart(retailer.set_index("Retailer Name")[["Invoice Upload_Success"]])

    with tab_audit:
        st.subheader("Clean External Event Audit")
        st.caption("Use this to verify exactly which events were included after filters and date selection.")
        if clean_audit.empty:
            st.warning("No clean external audit rows available for this date range.")
        else:
            display_cols = [
                "date", "event_time", "event_name", "identity_key", "email", "name", "lead_id",
                "session_id", "source_component", "product_count", "eligible_product_count", "gwp", "page_url"
            ]
            display_cols = [col for col in display_cols if col in clean_audit.columns]
            st.dataframe(clean_audit[display_cols], use_container_width=True, hide_index=True)

        with st.expander("All parseable events, including exclusions"):
            if not audit.empty:
                display_cols = [
                    "date", "event_name", "excluded", "exclusion_reason", "identity_key", "email", "name", "page_url"
                ]
                display_cols = [col for col in display_cols if col in audit.columns]
                st.dataframe(audit[display_cols], use_container_width=True, hide_index=True)

    with tab_downloads:
        st.subheader("Downloads")
        base = safe_stem(source_filename)
        st.download_button(
            "Download Excel workbook",
            data=result["excel_bytes"],
            file_name=f"{base}_clean_external_metrics.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.download_button(
            "Download daily metrics CSV",
            data=dataframe_to_csv_bytes(daily),
            file_name=f"{base}_clean_external_daily_metrics.csv",
            mime="text/csv",
        )
        st.download_button(
            "Download campaign attribution CSV",
            data=dataframe_to_csv_bytes(attribution_campaign),
            file_name=f"{base}_campaign_attribution.csv",
            mime="text/csv",
        )
        st.download_button(
            "Download ad set attribution CSV",
            data=dataframe_to_csv_bytes(attribution_adset),
            file_name=f"{base}_adset_attribution.csv",
            mime="text/csv",
        )
        st.download_button(
            "Download ads attribution CSV",
            data=dataframe_to_csv_bytes(attribution_ad),
            file_name=f"{base}_ads_attribution.csv",
            mime="text/csv",
        )
        st.download_button(
            "Download campaign event breakdown CSV",
            data=dataframe_to_csv_bytes(breakdown_campaign),
            file_name=f"{base}_campaign_event_breakdown.csv",
            mime="text/csv",
        )
        st.download_button(
            "Download ad set event breakdown CSV",
            data=dataframe_to_csv_bytes(breakdown_adset),
            file_name=f"{base}_adset_event_breakdown.csv",
            mime="text/csv",
        )
        st.download_button(
            "Download ads event breakdown CSV",
            data=dataframe_to_csv_bytes(breakdown_ad),
            file_name=f"{base}_ads_event_breakdown.csv",
            mime="text/csv",
        )
        st.download_button(
            "Download product stats CSV",
            data=dataframe_to_csv_bytes(product),
            file_name=f"{base}_product_stats.csv",
            mime="text/csv",
        )
        st.download_button(
            "Download sales CSV",
            data=dataframe_to_csv_bytes(sales),
            file_name=f"{base}_sales.csv",
            mime="text/csv",
        )
        st.download_button(
            "Download high intent signup dropoffs CSV",
            data=dataframe_to_csv_bytes(high_intent_signup),
            file_name=f"{base}_high_intent_signup_dropoffs.csv",
            mime="text/csv",
        )
        st.download_button(
            "Download high intent invoice dropoffs CSV",
            data=dataframe_to_csv_bytes(high_intent_invoice),
            file_name=f"{base}_high_intent_invoice_dropoffs.csv",
            mime="text/csv",
        )
        st.download_button(
            "Download retailer stats CSV",
            data=dataframe_to_csv_bytes(retailer),
            file_name=f"{base}_retailer_stats.csv",
            mime="text/csv",
        )
        st.download_button(
            "Download clean event audit CSV",
            data=dataframe_to_csv_bytes(clean_audit),
            file_name=f"{base}_clean_external_event_audit.csv",
            mime="text/csv",
        )
        st.download_button(
            "Download full audit CSV",
            data=dataframe_to_csv_bytes(audit),
            file_name=f"{base}_full_event_audit.csv",
            mime="text/csv",
        )

    st.caption("Metric logic matches the Clean External reporting script: unique-user counts for user metrics, product counts for product-level metrics, GWP from successful payment checkout totals, and campaign, ad set, ads attribution, UTM event breakdown, product, sales, high-intent dropoff, and retailer rollups from revised payload fields.")


if __name__ == "__main__":
    main()
