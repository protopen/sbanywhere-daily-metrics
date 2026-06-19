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
DEFAULT_V2_START_DATE = date(2026, 6, 18)
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
    "Sign Up_total": {"lead_signup", "quote_lead_captured", "signup_completed", "sign_up_completed", "sign_up_total", "signup_total", "sign_up"},
    "First Quote_Success": {"quote_generated", "first_quote_success", "offer_generation_success"},
    "Add to Cart_Success": {"cart_confirmed", "add_to_cart"},
    "Invoice Upload_Success": {"invoice_uploaded", "invoice_upload_success"},
    "Revised_offer_shown": {"revised_offer_shown", "revised_offer"},
    "Plan_selected": {"plan_selected", "offer_selected"},
    "Initiate_checkout": {"initiate_checkout"},
    "Payment Attempted": {"pay_now_clicked", "payment_attempted"},
    "Payment Success": {"payment_completed", "payment_success", "payment_succeeded", "purchase"},
    "Payment Failed": {"payment_failed"},
}

SANKEY_MAIN_PATH = [
    "Enquiry Attempted",
    "Sign Up_total",
    "First Quote_Success",
    "Add to Cart_Success",
]

SANKEY_TAIL_PATH = [
    "Initiate_checkout",
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
    """Build the requested branch-aware session funnel Sankey.

    Flow:
    Enquiry Attempted -> Sign Up_total -> First Quote_Success -> Add to Cart_Success
    -> Invoice Upload_Success OR Revised_offer_shown
    -> Plan_selected if Revised_offer_shown happened
    -> Initiate_checkout -> Payment Attempted -> Payment Success
    """
    if clean_audit.empty or "event_name" not in clean_audit.columns:
        return pd.DataFrame(columns=["source", "target", "value"])

    df = clean_audit.copy()
    df["stage"] = df["event_name"].map(_stage_for_event_name)
    df = df[df["stage"].notna()]
    if df.empty:
        return pd.DataFrame(columns=["source", "target", "value"])

    if "session_id" in df.columns and "identity_key" in df.columns:
        session_vals = df["session_id"].fillna("").astype(str).str.strip()
        identity_vals = df["identity_key"].fillna("").astype(str).str.strip()
        df["journey_key"] = session_vals.where(session_vals != "", identity_vals)
    elif "session_id" in df.columns:
        df["journey_key"] = df["session_id"].fillna("").astype(str).str.strip()
    elif "identity_key" in df.columns:
        df["journey_key"] = df["identity_key"].fillna("").astype(str).str.strip()
    else:
        df["journey_key"] = "unknown"

    df["journey_key"] = df["journey_key"].replace("", "unknown")

    link_counts: dict[tuple[str, str], int] = {}
    prefix = ["Enquiry Attempted", "Sign Up_total", "First Quote_Success", "Add to Cart_Success"]

    def add(source: str, target: str) -> None:
        link_counts[(source, target)] = link_counts.get((source, target), 0) + 1

    for _, g in df.groupby("journey_key", dropna=False):
        reached = set(g["stage"].dropna().astype(str).tolist())

        strict_prefix = []
        for stage in prefix:
            if stage in reached:
                strict_prefix.append(stage)
            else:
                break

        for source, target in zip(strict_prefix, strict_prefix[1:]):
            add(source, target)

        if len(strict_prefix) != len(prefix):
            continue

        branch_sources = []

        if "Invoice Upload_Success" in reached:
            add("Add to Cart_Success", "Invoice Upload_Success")
            branch_sources.append("Invoice Upload_Success")

        if "Revised_offer_shown" in reached:
            add("Add to Cart_Success", "Revised_offer_shown")
            if "Plan_selected" in reached:
                add("Revised_offer_shown", "Plan_selected")
                branch_sources.append("Plan_selected")
            else:
                branch_sources.append("Revised_offer_shown")

        if not branch_sources:
            continue

        if "Initiate_checkout" in reached:
            for source in branch_sources:
                add(source, "Initiate_checkout")

            if "Payment Attempted" in reached:
                add("Initiate_checkout", "Payment Attempted")
                if "Payment Success" in reached:
                    add("Payment Attempted", "Payment Success")
                elif "Payment Failed" in reached:
                    add("Payment Attempted", "Payment Failed")
            elif "Payment Failed" in reached:
                add("Initiate_checkout", "Payment Failed")

    rows = [{"source": s, "target": t, "value": v} for (s, t), v in link_counts.items() if v > 0]
    if not rows:
        return pd.DataFrame(columns=["source", "target", "value"])

    out = pd.DataFrame(rows)
    stage_order_list = [
        "Enquiry Attempted",
        "Sign Up_total",
        "First Quote_Success",
        "Add to Cart_Success",
        "Invoice Upload_Success",
        "Revised_offer_shown",
        "Plan_selected",
        "Initiate_checkout",
        "Payment Attempted",
        "Payment Success",
        "Payment Failed",
    ]
    stage_order = {stage: i for i, stage in enumerate(stage_order_list)}
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
        title_text="Clean External requested funnel journey",
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




# -----------------------------
# V2 Dashboard
# -----------------------------

V2_EVENT_SEQUENCE = [
    "enquiry_attempted",
    "enquiry_success",
    "sign_up",
    "initiate_checkout",
    "payment_attempted",
    "payment_success",
    "payment_failure",
]

V2_DAILY_COLUMNS = [
    "Date",
    "Enquiry Attempted_Total",
    "Sign Up_ Total",
    "Initiate Checkout",
    "Payment Success",
    "Payment Failure",
    "Gross GWP $",
]

V2_TAB2_COLUMNS = [
    "Date",
    "Enquiry Attempted",
    "Enquiry Success",
    "Sign Up_ Total",
    "Innitiate Checkout_User Count",
    "Innitiate Checkout_Product Count",
    "Innitiate Checkout_Cart Value",
    "Payment Success_ User Count",
    "Payment Failure_ User Count",
    "Payment Success_ Product Count",
    "Gross GWP $",
]

V2_TAB3_SOURCES = [
    "Meta",
    "Google",
    "Tiktok",
    "Direct",
    "Referral (surebright.com)",
    "Referral (Others)",
    "Organic",
    "Others",
]

V2_TAB3_COLUMNS = [
    "Source",
    "Traffic (Total)",
    "Enquiry Attempted_Total",
    "Sign Up_ Total",
    "Innitiate Checkout",
    "Payment Success",
    "Payment Failure",
    "Gross GWP $",
]

V2_PAID_SOURCE_OPTIONS = [
    "Meta",
    "Google",
    "Tiktok",
    "Influencers",
    "Others",
]

V2_TAB4_CAMPAIGN_COLUMNS = [
    "Campaign Name",
    "Traffic (Total)",
    "Enquiry Attempted_Total",
    "Sign Up_ Total",
    "Innitiate Checkout",
    "Payment Success",
    "Payment Failure",
    "Gross GWP $",
]

V2_TAB4_ADSET_COLUMNS = [
    "Campaign Name",
    "Adset Name",
    "Traffic (Total)",
    "Enquiry Attempted_Total",
    "Sign Up_ Total",
    "Innitiate Checkout",
    "Payment Success",
    "Payment Failure",
    "Gross GWP $",
]

V2_TAB4_AD_COLUMNS = [
    "Campaign Name",
    "Adset Name",
    "Ad Name",
    "Traffic (Total)",
    "Enquiry Attempted_Total",
    "Sign Up_ Total",
    "Innitiate Checkout",
    "Payment Success",
    "Payment Failure",
    "Gross GWP $",
]

V2_TAB5_COLUMNS = [
    "Category",
    "Enquiry Attempted_Total",
    "Sign Up_ Total",
    "Innitiate Checkout",
    "Payment Success",
    "Payment Failure",
    "Gross GWP $",
]

V2_TRAFFIC_SOURCE_OPTIONS = [
    "Meta",
    "Google",
    "Tiktok",
    "Direct",
    "Referral (surebright.com)",
    "Referral (Others)",
    "Organic",
    "Others",
]

V2_TAB6_COLUMNS = [
    "Date",
    "Name",
    "Email",
    "Last Event",
    "Product Sub Category",
    "Price",
    "Waranty Type",
    "Warranty Tenure",
    "Plan Price",
    "UTM Campaign Name",
    "UTM Adset Name",
    "UTM Ad Name",
]

V2_TAB7_COLUMNS = [
    "Date",
    "Name",
    "Email",
    "Journey Flow",
    "Event",
    "Order / Payment ID",
    "Product Category",
    "Product Title",
    "Product Brand",
    "Manufacturer",
    "Model Number",
    "Product Condition",
    "Quantity",
    "Product Unit Price",
    "Product Total Price",
    "Warranty Type",
    "Warranty / Plan Name",
    "Warranty Price",
    "Warranty Term Months",
    "Warranty Provider",
    "Manufacturer Warranty",
    "Eligible",
    "Gross GWP $",
    "UTM Source",
    "UTM Medium",
    "UTM Campaign",
]


def _v2_safe_json(raw_value):
    if isinstance(raw_value, dict):
        return raw_value
    try:
        return json.loads(str(raw_value or "{}"))
    except Exception:
        return {}


def _v2_nested_get(obj: dict, *paths: str):
    for path in paths:
        cur = obj
        ok = True
        for part in path.split("."):
            if isinstance(cur, dict) and part in cur:
                cur = cur[part]
            else:
                ok = False
                break
        if ok:
            return cur
    return None


def _v2_norm_value(value) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[\s\-]+", "_", text)
    return text


def _v2_label(value: str) -> str:
    value = _v2_norm_value(value)
    if not value:
        return "No Status"
    mapping = {
        "manual": "Manual",
        "invoice_upload": "Invoice Upload",
        "invoice_upload_success": "Invoice Upload Success",
        "invoice_upload_failure": "Invoice Upload Failure",
        "success": "Invoice Upload Success",
        "failure": "Invoice Upload Failure",
    }
    return mapping.get(value, value.replace("_", " ").title())


def _v2_parse_money(value) -> float:
    if value in (None, ""):
        return 0.0
    if isinstance(value, (int, float)):
        try:
            return float(value)
        except Exception:
            return 0.0
    text = str(value)
    text = re.sub(r"[^0-9.\-]", "", text)
    if text in ("", ".", "-", "-."):
        return 0.0
    try:
        return float(text)
    except Exception:
        return 0.0


def _v2_collect_line_items(obj: dict) -> list[dict]:
    candidates = [
        _v2_nested_get(obj, "event_data.line_items"),
        _v2_nested_get(obj, "event_data.items"),
        _v2_nested_get(obj, "data.line_items"),
        _v2_nested_get(obj, "data.items"),
        _v2_nested_get(obj, "line_items"),
        _v2_nested_get(obj, "items"),
        _v2_nested_get(obj, "raw.line_items"),
        _v2_nested_get(obj, "raw.items"),
    ]
    for candidate in candidates:
        if isinstance(candidate, list):
            return [x for x in candidate if isinstance(x, dict)]
    return []


def _v2_item_quantity(item: dict) -> int:
    qty = _v2_nested_get(item, "quantity", "qty", "purchase.quantity")
    try:
        qty_int = int(float(qty))
        return qty_int if qty_int > 0 else 1
    except Exception:
        return 1


def _v2_item_total_value(item: dict) -> float:
    direct = _v2_nested_get(item, "purchase.total_price", "total_price", "item_total", "line_total", "amount")
    direct_val = _v2_parse_money(direct)
    if direct_val:
        return direct_val

    unit = _v2_parse_money(_v2_nested_get(item, "purchase.unit_price", "unit_price", "price", "product.price"))
    return unit * _v2_item_quantity(item)


def _v2_product_category_from_obj(obj: dict) -> str:
    line_items = _v2_collect_line_items(obj)
    for item in line_items:
        product = item.get("product") if isinstance(item.get("product"), dict) else {}
        category = product.get("category") if isinstance(product.get("category"), dict) else None
        value = (
            (category or {}).get("name") if isinstance(category, dict) else None
        ) or product.get("category") or item.get("item_category") or item.get("category")
        if value not in (None, ""):
            return str(value).strip()

    value = _v2_nested_get(
        obj,
        "event_data.form.fields.product_category",
        "event_data.form.fields.category",
        "event_data.product.category.name",
        "event_data.product.category",
        "data.categoryName",
        "data.sbCategoryName",
        "data.product.category.name",
        "data.product.category",
        "raw.product_category",
        "raw.category",
        "product.category.name",
        "product.category",
    )
    return str(value).strip() if value not in (None, "") else "Unknown"


def _v2_product_count_from_obj(obj: dict, fallback=0) -> int:
    line_items = _v2_collect_line_items(obj)
    if line_items:
        return int(sum(_v2_item_quantity(item) for item in line_items))
    try:
        val = int(float(fallback or 0))
        return val if val > 0 else 1
    except Exception:
        return 1


def _v2_cart_value_from_obj(obj: dict, fallback=0) -> float:
    line_items = _v2_collect_line_items(obj)
    if line_items:
        return float(sum(_v2_item_total_value(item) for item in line_items))
    return _v2_parse_money(fallback)


def _v2_url_domain(value) -> str:
    if value in (None, ""):
        return ""
    text = str(value).strip().lower()
    try:
        parsed = urlparse(text if "://" in text else f"https://{text}")
        return parsed.netloc.lower().replace("www.", "")
    except Exception:
        return ""


def _v2_url_query_value(value: str, key: str) -> str:
    if not value:
        return ""
    try:
        parsed = urlparse(str(value))
        qs = parse_qs(parsed.query)
        vals = qs.get(key) or []
        return str(vals[0]).strip().lower() if vals else ""
    except Exception:
        return ""


def _v2_source_bucket(obj: dict) -> str:
    utm_source = str(
        _v2_nested_get(
            obj,
            "event_data.utm.source",
            "event_data.utm_source",
            "utm.source",
            "utm_source",
            "data.utm.source",
            "data.utm_source",
            "raw.utm_source",
        )
        or ""
    ).strip().lower()

    utm_medium = str(
        _v2_nested_get(
            obj,
            "event_data.utm.medium",
            "event_data.utm_medium",
            "utm.medium",
            "utm_medium",
            "data.utm.medium",
            "data.utm_medium",
            "raw.utm_medium",
        )
        or ""
    ).strip().lower()

    page_url = str(
        _v2_nested_get(
            obj,
            "source.page_url",
            "source.url",
            "page_url",
            "url",
            "context.page.url",
            "event_data.page_url",
        )
        or ""
    )

    referrer = str(
        _v2_nested_get(
            obj,
            "source.referrer",
            "source.referrer_url",
            "referrer",
            "referer",
            "context.page.referrer",
            "event_data.referrer",
            "event_data.referrer_url",
        )
        or ""
    )

    ref_domain = _v2_url_domain(referrer)
    page_has_gclid = bool(_v2_url_query_value(page_url, "gclid"))
    page_has_fbclid = bool(_v2_url_query_value(page_url, "fbclid"))
    page_has_ttclid = bool(_v2_url_query_value(page_url, "ttclid"))

    source_text = f"{utm_source} {utm_medium} {ref_domain}".lower()

    if any(token in source_text for token in ["facebook", "instagram", "meta", "fb", "ig"]) or page_has_fbclid:
        return "Meta"
    if "google" in source_text or page_has_gclid:
        return "Google"
    if "tiktok" in source_text or page_has_ttclid:
        return "Tiktok"

    search_domains = ("google.", "bing.", "yahoo.", "duckduckgo.", "ecosia.", "baidu.", "yandex.")
    if "organic" in utm_medium or any(domain in ref_domain for domain in search_domains):
        return "Organic"

    if ref_domain:
        if ref_domain == "surebright.com" or ref_domain.endswith(".surebright.com"):
            return "Referral (surebright.com)"
        if "surebrightanywhere.com" not in ref_domain:
            return "Referral (Others)"

    if not utm_source and not ref_domain:
        return "Direct"

    return "Others"



def _v2_attr_value(obj: dict, *paths: str) -> str:
    for path in paths:
        value = _v2_nested_get(obj, path)
        if value not in (None, ""):
            return str(value).strip()
    return ""


def _v2_url_param_from_obj(obj: dict, key: str) -> str:
    page_url = str(
        _v2_nested_get(
            obj,
            "source.page_url",
            "source.url",
            "page_url",
            "url",
            "context.page.url",
            "event_data.page_url",
        )
        or ""
    )
    return _v2_url_query_value(page_url, key)


def _v2_campaign_values(obj: dict) -> dict:
    campaign = _v2_attr_value(
        obj,
        "event_data.utm.campaign",
        "event_data.utm_campaign",
        "utm.campaign",
        "utm_campaign",
        "data.utm.campaign",
        "data.utm_campaign",
        "raw.utm_campaign",
    ) or _v2_url_param_from_obj(obj, "utm_campaign") or "Unknown"

    adset = _v2_attr_value(
        obj,
        "event_data.utm.adset",
        "event_data.utm.ad_set",
        "event_data.utm_adset",
        "event_data.utm_ad_set",
        "event_data.adset_name",
        "event_data.ad_set_name",
        "utm.adset",
        "utm.ad_set",
        "utm_adset",
        "utm_ad_set",
        "adset_name",
        "ad_set_name",
        "data.utm.adset",
        "data.utm.ad_set",
        "data.adset_name",
        "data.ad_set_name",
        "raw.utm_adset",
        "raw.utm_ad_set",
    ) or _v2_url_param_from_obj(obj, "utm_adset") or _v2_url_param_from_obj(obj, "utm_ad_set") or "Unknown"

    ad = _v2_attr_value(
        obj,
        "event_data.utm.ad",
        "event_data.utm_ad",
        "event_data.ad_name",
        "utm.ad",
        "utm_ad",
        "ad_name",
        "data.utm.ad",
        "data.ad_name",
        "raw.utm_ad",
    ) or _v2_url_param_from_obj(obj, "utm_ad") or "Unknown"

    return {
        "Campaign Name": campaign,
        "Adset Name": adset,
        "Ad Name": ad,
    }



def _v2_detail_value(obj: dict, *paths: str) -> str:
    for path in paths:
        value = _v2_nested_get(obj, path)
        if value not in (None, ""):
            return str(value).strip()
    return ""


def _v2_name_from_obj(obj: dict) -> str:
    direct = _v2_detail_value(
        obj,
        "event_data.form.fields.name",
        "event_data.form.fields.full_name",
        "event_data.form.fields.fullName",
        "event_data.customer.name",
        "event_data.user.name",
        "actor.name",
        "customer.name",
        "user.name",
        "data.name",
        "data.full_name",
        "name",
    )
    if direct:
        return direct

    first = _v2_detail_value(
        obj,
        "event_data.form.fields.first_name",
        "event_data.form.fields.firstname",
        "actor.first_name",
        "customer.first_name",
        "data.first_name",
    )
    last = _v2_detail_value(
        obj,
        "event_data.form.fields.last_name",
        "event_data.form.fields.lastname",
        "actor.last_name",
        "customer.last_name",
        "data.last_name",
    )
    return " ".join([x for x in [first, last] if x]).strip()


def _v2_email_from_obj(obj: dict) -> str:
    return _v2_detail_value(
        obj,
        "event_data.form.fields.email",
        "event_data.form.fields.email_address",
        "event_data.customer.email",
        "event_data.user.email",
        "actor.email",
        "customer.email",
        "user.email",
        "data.email",
        "email",
    ).lower()


def _v2_product_detail_values(obj: dict) -> dict:
    line_items = _v2_collect_line_items(obj)
    first_item = line_items[0] if line_items else {}
    product = first_item.get("product") if isinstance(first_item.get("product"), dict) else {}
    purchase = first_item.get("purchase") if isinstance(first_item.get("purchase"), dict) else {}
    protection = first_item.get("protection") if isinstance(first_item.get("protection"), dict) else {}

    category_obj = product.get("category") if isinstance(product.get("category"), dict) else {}
    subcategory = (
        _v2_detail_value(
            obj,
            "event_data.form.fields.product_sub_category",
            "event_data.form.fields.product_subcategory",
            "event_data.form.fields.sub_category",
            "event_data.form.fields.subcategory",
            "event_data.product.sub_category",
            "event_data.product.subcategory",
            "data.product.sub_category",
            "data.product.subcategory",
            "raw.product_sub_category",
            "raw.product_subcategory",
        )
        or str(product.get("sub_category") or product.get("subcategory") or first_item.get("item_subcategory") or first_item.get("sub_category") or first_item.get("subcategory") or category_obj.get("sub_category") or category_obj.get("subcategory") or "")
    )

    price = (
        _v2_parse_money(_v2_detail_value(
            obj,
            "event_data.form.fields.product_price",
            "event_data.form.fields.price",
            "event_data.product.price",
            "data.product.price",
            "raw.product_price",
            "raw.price",
        ))
        or _v2_parse_money(purchase.get("unit_price") or first_item.get("unit_price") or first_item.get("price") or purchase.get("total_price") or first_item.get("total_price"))
    )

    warranty_type = str(
        protection.get("plan_name")
        or protection.get("name")
        or protection.get("type")
        or first_item.get("item_variant")
        or _v2_detail_value(obj, "event_data.warranty.type", "event_data.plan.type", "data.warranty.type", "data.plan.type")
        or ""
    ).strip()

    warranty_tenure = str(
        protection.get("term_months")
        or protection.get("term")
        or protection.get("tenure")
        or _v2_detail_value(obj, "event_data.warranty.term_months", "event_data.warranty.tenure", "event_data.plan.term_months", "event_data.plan.tenure", "data.warranty.term_months")
        or ""
    ).strip()

    plan_price = (
        _v2_parse_money(protection.get("plan_price") or protection.get("price") or first_item.get("plan_price") or first_item.get("warrantyPrice") or first_item.get("warranty_price"))
        or _v2_parse_money(_v2_detail_value(obj, "event_data.warranty.price", "event_data.plan.price", "data.warranty.price", "data.plan.price"))
    )

    return {
        "Product Sub Category": subcategory or "",
        "Price": round(float(price or 0), 2),
        "Waranty Type": warranty_type,
        "Warranty Tenure": warranty_tenure,
        "Plan Price": round(float(plan_price or 0), 2),
    }



def _v2_order_payment_id(obj: dict) -> str:
    return _v2_detail_value(
        obj,
        "event_data.order.id",
        "event_data.order.order_id",
        "event_data.order_id",
        "event_data.payment.id",
        "event_data.payment.payment_id",
        "event_data.payment_id",
        "event_data.checkout.id",
        "event_data.checkout_id",
        "data.order.id",
        "data.order_id",
        "data.payment.id",
        "data.payment_id",
        "order.id",
        "order_id",
        "payment.id",
        "payment_id",
        "id",
    )


def _v2_utm_values(obj: dict) -> dict:
    source = _v2_attr_value(
        obj,
        "event_data.utm.source",
        "event_data.utm_source",
        "utm.source",
        "utm_source",
        "data.utm.source",
        "data.utm_source",
        "raw.utm_source",
    ) or _v2_url_param_from_obj(obj, "utm_source")

    medium = _v2_attr_value(
        obj,
        "event_data.utm.medium",
        "event_data.utm_medium",
        "utm.medium",
        "utm_medium",
        "data.utm.medium",
        "data.utm_medium",
        "raw.utm_medium",
    ) or _v2_url_param_from_obj(obj, "utm_medium")

    campaign = _v2_attr_value(
        obj,
        "event_data.utm.campaign",
        "event_data.utm_campaign",
        "utm.campaign",
        "utm_campaign",
        "data.utm.campaign",
        "data.utm_campaign",
        "raw.utm_campaign",
    ) or _v2_url_param_from_obj(obj, "utm_campaign")

    return {"UTM Source": source, "UTM Medium": medium, "UTM Campaign": campaign}


def _v2_full_product_values(obj: dict, gwp=0) -> dict:
    line_items = _v2_collect_line_items(obj)
    item = line_items[0] if line_items else {}
    product = item.get("product") if isinstance(item.get("product"), dict) else {}
    purchase = item.get("purchase") if isinstance(item.get("purchase"), dict) else {}
    protection = item.get("protection") if isinstance(item.get("protection"), dict) else {}
    category_obj = product.get("category") if isinstance(product.get("category"), dict) else {}

    product_category = (
        category_obj.get("name")
        or product.get("category")
        or item.get("item_category")
        or item.get("category")
        or _v2_product_category_from_obj(obj)
    )

    product_title = (
        product.get("title")
        or product.get("name")
        or product.get("description")
        or item.get("item_name")
        or item.get("name")
        or _v2_detail_value(obj, "event_data.product.title", "event_data.product.name", "data.product.title", "data.product.name")
    )

    product_brand = (
        product.get("brand")
        or _v2_detail_value(obj, "event_data.product.brand", "data.product.brand", "raw.product_brand")
    )

    manufacturer = (
        _v2_detail_value(obj, "manufacturer.name", "event_data.manufacturer.name", "data.manufacturer.name")
        or product_brand
    )

    model_number = (
        _v2_detail_value(obj, "manufacturer.model_number", "event_data.manufacturer.model_number", "data.manufacturer.model_number")
        or str(product.get("sku") or product.get("model_number") or item.get("model_number") or "")
    )

    product_condition = str(
        product.get("condition")
        or item.get("condition")
        or _v2_detail_value(obj, "event_data.product.condition", "data.product.condition")
        or ""
    ).strip()

    quantity = _v2_item_quantity(item) if item else _v2_product_count_from_obj(obj, 1)

    unit_price = _v2_parse_money(
        purchase.get("unit_price")
        or item.get("unit_price")
        or item.get("price")
        or _v2_detail_value(obj, "event_data.product.unit_price", "event_data.product.price", "data.product.price", "raw.product_price")
    )

    total_price = _v2_parse_money(purchase.get("total_price") or item.get("total_price"))
    if not total_price and unit_price:
        total_price = unit_price * quantity

    warranty_type = str(
        protection.get("type")
        or _v2_detail_value(obj, "event_data.warranty.type", "event_data.plan.type", "data.warranty.type", "data.plan.type")
        or ""
    ).strip()

    warranty_plan_name = str(
        protection.get("plan_name")
        or protection.get("name")
        or item.get("item_variant")
        or _v2_detail_value(obj, "event_data.warranty.plan_name", "event_data.plan.name", "data.warranty.plan_name", "data.plan.name")
        or ""
    ).strip()

    warranty_price = (
        _v2_parse_money(protection.get("plan_price") or protection.get("price") or item.get("plan_price") or item.get("warrantyPrice") or item.get("warranty_price"))
        or _v2_parse_money(_v2_detail_value(obj, "event_data.warranty.price", "event_data.plan.price", "data.warranty.price", "data.plan.price"))
        or _v2_parse_money(gwp)
    )

    warranty_term = str(
        protection.get("term_months")
        or protection.get("term")
        or protection.get("tenure")
        or _v2_detail_value(obj, "event_data.warranty.term_months", "event_data.plan.term_months", "data.warranty.term_months", "data.plan.term_months")
        or ""
    ).strip()

    warranty_provider = str(
        protection.get("provider")
        or protection.get("underwriter")
        or protection.get("administrator")
        or _v2_detail_value(obj, "event_data.warranty.provider", "event_data.plan.provider", "data.warranty.provider", "data.plan.provider")
        or ""
    ).strip()

    manufacturer_warranty = _v2_detail_value(
        obj,
        "manufacturer.warranty",
        "event_data.manufacturer.warranty",
        "data.manufacturer.warranty",
        "event_data.product.manufacturer_warranty",
        "data.product.manufacturer_warranty",
    )

    eligible_value = (
        item.get("eligible")
        if isinstance(item, dict) and "eligible" in item
        else _v2_nested_get(obj, "event_data.eligible", "data.eligible", "eligible")
    )
    if isinstance(eligible_value, bool):
        eligible = eligible_value
    elif eligible_value in (None, ""):
        eligible = ""
    else:
        eligible = str(eligible_value).strip()

    return {
        "Product Category": str(product_category or "Unknown"),
        "Product Title": str(product_title or ""),
        "Product Brand": str(product_brand or ""),
        "Manufacturer": str(manufacturer or ""),
        "Model Number": str(model_number or ""),
        "Product Condition": product_condition,
        "Quantity": quantity,
        "Product Unit Price": round(float(unit_price or 0), 2),
        "Product Total Price": round(float(total_price or 0), 2),
        "Warranty Type": warranty_type,
        "Warranty / Plan Name": warranty_plan_name,
        "Warranty Price": round(float(warranty_price or 0), 2),
        "Warranty Term Months": warranty_term,
        "Warranty Provider": warranty_provider,
        "Manufacturer Warranty": manufacturer_warranty,
        "Eligible": eligible,
    }


def _v2_paid_source_bucket(obj: dict, source_bucket: str) -> str:
    if source_bucket in {"Meta", "Google", "Tiktok"}:
        return source_bucket

    utm_source = str(
        _v2_nested_get(
            obj,
            "event_data.utm.source",
            "event_data.utm_source",
            "utm.source",
            "utm_source",
            "data.utm.source",
            "data.utm_source",
            "raw.utm_source",
        )
        or ""
    ).strip().lower()

    utm_medium = str(
        _v2_nested_get(
            obj,
            "event_data.utm.medium",
            "event_data.utm_medium",
            "utm.medium",
            "utm_medium",
            "data.utm.medium",
            "data.utm_medium",
            "raw.utm_medium",
        )
        or ""
    ).strip().lower()

    combined = f"{utm_source} {utm_medium}".lower()
    if any(token in combined for token in ["influencer", "creator", "affiliate", "collab"]):
        return "Influencers"

    return "Others"


def _v2_clean_events_frame(clean_audit: pd.DataFrame) -> pd.DataFrame:
    if clean_audit is None or clean_audit.empty:
        return pd.DataFrame(columns=["Date", "event_name", "flow_method", "flow_status", "Product Category", "gwp"])

    rows = []
    for _, row in clean_audit.iterrows():
        obj = _v2_safe_json(row.get("raw_json", "{}"))
        flow_method = _v2_nested_get(
            obj,
            "event_data.flow.method",
            "data.flow.method",
            "flow.method",
        )
        flow_status = _v2_nested_get(
            obj,
            "event_data.flow.status",
            "data.flow.status",
            "flow.status",
        )

        event_name = _v2_norm_value(row.get("event_name", ""))
        gwp = float(row.get("gwp", 0) or 0)
        product_count_fallback = row.get("product_count", 0)
        identity = str(row.get("session_id", "") or "").strip() or str(row.get("identity_key", "") or "").strip()

        source_bucket = _v2_source_bucket(obj)
        campaign_values = _v2_campaign_values(obj)
        product_detail_values = _v2_product_detail_values(obj)
        full_product_values = _v2_full_product_values(obj, gwp)
        utm_values = _v2_utm_values(obj)

        rows.append(
            {
                "Date": str(row.get("date", "") or ""),
                "event_name": event_name,
                "flow_method": _v2_norm_value(flow_method),
                "flow_status": _v2_norm_value(flow_status),
                "Journey Type": _v2_label(flow_method),
                "Invoice Status": _v2_label(flow_status),
                "Product Category": _v2_product_category_from_obj(obj),
                "Source": source_bucket,
                "Traffic Source": source_bucket,
                "Paid Campaign Source": _v2_paid_source_bucket(obj, source_bucket),
                "Campaign Name": campaign_values["Campaign Name"],
                "Adset Name": campaign_values["Adset Name"],
                "Ad Name": campaign_values["Ad Name"],
                "Name": _v2_name_from_obj(obj),
                "Email": _v2_email_from_obj(obj),
                "Product Sub Category": product_detail_values["Product Sub Category"],
                "Price": product_detail_values["Price"],
                "Waranty Type": product_detail_values["Waranty Type"],
                "Warranty Tenure": product_detail_values["Warranty Tenure"],
                "Plan Price": product_detail_values["Plan Price"],
                "Journey Flow": _v2_label(flow_method),
                "Event": event_name,
                "Order / Payment ID": _v2_order_payment_id(obj),
                "Product Title": full_product_values["Product Title"],
                "Product Brand": full_product_values["Product Brand"],
                "Manufacturer": full_product_values["Manufacturer"],
                "Model Number": full_product_values["Model Number"],
                "Product Condition": full_product_values["Product Condition"],
                "Quantity": full_product_values["Quantity"],
                "Product Unit Price": full_product_values["Product Unit Price"],
                "Product Total Price": full_product_values["Product Total Price"],
                "Warranty Type": full_product_values["Warranty Type"],
                "Warranty / Plan Name": full_product_values["Warranty / Plan Name"],
                "Warranty Price": full_product_values["Warranty Price"],
                "Warranty Term Months": full_product_values["Warranty Term Months"],
                "Warranty Provider": full_product_values["Warranty Provider"],
                "Manufacturer Warranty": full_product_values["Manufacturer Warranty"],
                "Eligible": full_product_values["Eligible"],
                "Gross GWP $": round(float(gwp or 0), 2),
                "UTM Source": utm_values["UTM Source"],
                "UTM Medium": utm_values["UTM Medium"],
                "UTM Campaign": utm_values["UTM Campaign"],
                "identity": identity,
                "product_count": _v2_product_count_from_obj(obj, product_count_fallback),
                "cart_value": _v2_cart_value_from_obj(obj, row.get("gwp", 0)),
                "gwp": gwp,
            }
        )

    return pd.DataFrame(rows)


def _v2_build_daily_tab1(events: pd.DataFrame) -> pd.DataFrame:
    if events is None or events.empty:
        return pd.DataFrame(columns=V2_DAILY_COLUMNS)

    df = events[events["event_name"].isin(V2_EVENT_SEQUENCE)].copy()
    if df.empty:
        return pd.DataFrame(columns=V2_DAILY_COLUMNS)

    rows = []
    for day, g in df.groupby("Date", dropna=False):
        rows.append(
            {
                "Date": day,
                "Enquiry Attempted_Total": int((g["event_name"] == "enquiry_attempted").sum()),
                "Sign Up_ Total": int((g["event_name"] == "sign_up").sum()),
                "Initiate Checkout": int((g["event_name"] == "initiate_checkout").sum()),
                "Payment Success": int((g["event_name"] == "payment_success").sum()),
                "Payment Failure": int((g["event_name"] == "payment_failure").sum()),
                "Gross GWP $": round(float(g.loc[g["event_name"] == "payment_success", "gwp"].sum()), 2),
            }
        )

    out = pd.DataFrame(rows)
    if out.empty:
        return pd.DataFrame(columns=V2_DAILY_COLUMNS)
    out = out.sort_values("Date", ascending=False).reset_index(drop=True)
    return out[V2_DAILY_COLUMNS]




def _v2_build_daily_tab2(events: pd.DataFrame) -> pd.DataFrame:
    if events is None or events.empty:
        return pd.DataFrame(columns=V2_TAB2_COLUMNS)

    df = events.copy()
    if df.empty:
        return pd.DataFrame(columns=V2_TAB2_COLUMNS)

    rows = []
    for day, g in df.groupby("Date", dropna=False):
        initiate = g[g["event_name"] == "initiate_checkout"]
        payment_success = g[g["event_name"] == "payment_success"]
        payment_failure = g[g["event_name"] == "payment_failure"]

        rows.append(
            {
                "Date": day,
                "Enquiry Attempted": int((g["event_name"] == "enquiry_attempted").sum()),
                "Enquiry Success": int((g["event_name"] == "enquiry_success").sum()),
                "Sign Up_ Total": int((g["event_name"] == "sign_up").sum()),
                "Innitiate Checkout_User Count": int(initiate["identity"].dropna().astype(str).replace("", pd.NA).dropna().nunique()),
                "Innitiate Checkout_Product Count": int(initiate["product_count"].fillna(0).sum()),
                "Innitiate Checkout_Cart Value": round(float(initiate["cart_value"].fillna(0).sum()), 2),
                "Payment Success_ User Count": int(payment_success["identity"].dropna().astype(str).replace("", pd.NA).dropna().nunique()),
                "Payment Failure_ User Count": int(payment_failure["identity"].dropna().astype(str).replace("", pd.NA).dropna().nunique()),
                "Payment Success_ Product Count": int(payment_success["product_count"].fillna(0).sum()),
                "Gross GWP $": round(float(payment_success["gwp"].fillna(0).sum()), 2),
            }
        )

    out = pd.DataFrame(rows)
    if out.empty:
        return pd.DataFrame(columns=V2_TAB2_COLUMNS)
    out = out.sort_values("Date", ascending=False).reset_index(drop=True)
    return out[V2_TAB2_COLUMNS]




def _v2_build_source_tab3(events: pd.DataFrame) -> pd.DataFrame:
    rows = []
    if events is None or events.empty:
        return pd.DataFrame([{col: (source if col == "Source" else 0) for col in V2_TAB3_COLUMNS} for source in V2_TAB3_SOURCES])

    df = events.copy()
    for source in V2_TAB3_SOURCES:
        g = df[df["Source"] == source]
        payment_success = g[g["event_name"] == "payment_success"]
        rows.append(
            {
                "Source": source,
                "Traffic (Total)": int(g["identity"].dropna().astype(str).replace("", pd.NA).dropna().nunique()),
                "Enquiry Attempted_Total": int((g["event_name"] == "enquiry_attempted").sum()),
                "Sign Up_ Total": int((g["event_name"] == "sign_up").sum()),
                "Innitiate Checkout": int((g["event_name"] == "initiate_checkout").sum()),
                "Payment Success": int((g["event_name"] == "payment_success").sum()),
                "Payment Failure": int((g["event_name"] == "payment_failure").sum()),
                "Gross GWP $": round(float(payment_success["gwp"].fillna(0).sum()), 2),
            }
        )

    out = pd.DataFrame(rows)
    return out[V2_TAB3_COLUMNS]




def _v2_aggregate_campaign_table(events: pd.DataFrame, group_cols: list[str], output_cols: list[str]) -> pd.DataFrame:
    if events is None or events.empty:
        return pd.DataFrame(columns=output_cols)

    rows = []
    for key, g in events.groupby(group_cols, dropna=False):
        if not isinstance(key, tuple):
            key = (key,)
        payment_success = g[g["event_name"] == "payment_success"]
        base = {col: (str(val) if val not in (None, "") else "Unknown") for col, val in zip(group_cols, key)}
        base.update(
            {
                "Traffic (Total)": int(g["identity"].dropna().astype(str).replace("", pd.NA).dropna().nunique()),
                "Enquiry Attempted_Total": int((g["event_name"] == "enquiry_attempted").sum()),
                "Sign Up_ Total": int((g["event_name"] == "sign_up").sum()),
                "Innitiate Checkout": int((g["event_name"] == "initiate_checkout").sum()),
                "Payment Success": int((g["event_name"] == "payment_success").sum()),
                "Payment Failure": int((g["event_name"] == "payment_failure").sum()),
                "Gross GWP $": round(float(payment_success["gwp"].fillna(0).sum()), 2),
            }
        )
        rows.append(base)

    out = pd.DataFrame(rows)
    if out.empty:
        return pd.DataFrame(columns=output_cols)
    out = out.sort_values(["Traffic (Total)", "Payment Success", "Gross GWP $"], ascending=[False, False, False]).reset_index(drop=True)
    return out.reindex(columns=output_cols)


def _v2_build_paid_campaign_tab4(events: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    campaign = _v2_aggregate_campaign_table(events, ["Campaign Name"], V2_TAB4_CAMPAIGN_COLUMNS)
    adset = _v2_aggregate_campaign_table(events, ["Campaign Name", "Adset Name"], V2_TAB4_ADSET_COLUMNS)
    ad = _v2_aggregate_campaign_table(events, ["Campaign Name", "Adset Name", "Ad Name"], V2_TAB4_AD_COLUMNS)
    return campaign, adset, ad




def _v2_build_category_tab5(events: pd.DataFrame) -> pd.DataFrame:
    if events is None or events.empty:
        return pd.DataFrame(columns=V2_TAB5_COLUMNS)

    rows = []
    for category, g in events.groupby("Product Category", dropna=False):
        payment_success = g[g["event_name"] == "payment_success"]
        rows.append(
            {
                "Category": str(category) if str(category or "").strip() else "Unknown",
                "Enquiry Attempted_Total": int((g["event_name"] == "enquiry_attempted").sum()),
                "Sign Up_ Total": int((g["event_name"] == "sign_up").sum()),
                "Innitiate Checkout": int((g["event_name"] == "initiate_checkout").sum()),
                "Payment Success": int((g["event_name"] == "payment_success").sum()),
                "Payment Failure": int((g["event_name"] == "payment_failure").sum()),
                "Gross GWP $": round(float(payment_success["gwp"].fillna(0).sum()), 2),
            }
        )

    out = pd.DataFrame(rows)
    if out.empty:
        return pd.DataFrame(columns=V2_TAB5_COLUMNS)
    out = out.sort_values(["Payment Success", "Gross GWP $", "Enquiry Attempted_Total"], ascending=[False, False, False]).reset_index(drop=True)
    return out[V2_TAB5_COLUMNS]




def _v2_first_nonempty(series: pd.Series) -> str:
    for value in series.tolist():
        if value not in (None, "") and str(value).strip():
            return value
    return ""


def _v2_build_detail_tab6(events: pd.DataFrame) -> pd.DataFrame:
    if events is None or events.empty:
        return pd.DataFrame(columns=V2_TAB6_COLUMNS)

    df = events.copy()
    if "event_name" in df.columns:
        df = df[df["event_name"].isin(V2_EVENT_SEQUENCE)].copy()

    if df.empty:
        return pd.DataFrame(columns=V2_TAB6_COLUMNS)

    # Keep one row per session/identity. The displayed Date and Last Event come
    # from the latest event after filtering; details use the first available
    # non-empty value from the same session.
    rows = []
    for identity, g in df.groupby("identity", dropna=False):
        if not str(identity or "").strip():
            continue

        g = g.sort_values("Date")
        latest = g.iloc[-1]
        rows.append(
            {
                "Date": latest.get("Date", ""),
                "Name": _v2_first_nonempty(g["Name"]) if "Name" in g.columns else "",
                "Email": _v2_first_nonempty(g["Email"]) if "Email" in g.columns else "",
                "Last Event": str(latest.get("event_name", "")).replace("_", " ").title(),
                "Product Sub Category": _v2_first_nonempty(g["Product Sub Category"]) if "Product Sub Category" in g.columns else "",
                "Price": _v2_first_nonempty(g["Price"]) if "Price" in g.columns else 0,
                "Waranty Type": _v2_first_nonempty(g["Waranty Type"]) if "Waranty Type" in g.columns else "",
                "Warranty Tenure": _v2_first_nonempty(g["Warranty Tenure"]) if "Warranty Tenure" in g.columns else "",
                "Plan Price": _v2_first_nonempty(g["Plan Price"]) if "Plan Price" in g.columns else 0,
                "UTM Campaign Name": _v2_first_nonempty(g["Campaign Name"]) if "Campaign Name" in g.columns else "",
                "UTM Adset Name": _v2_first_nonempty(g["Adset Name"]) if "Adset Name" in g.columns else "",
                "UTM Ad Name": _v2_first_nonempty(g["Ad Name"]) if "Ad Name" in g.columns else "",
            }
        )

    out = pd.DataFrame(rows)
    if out.empty:
        return pd.DataFrame(columns=V2_TAB6_COLUMNS)
    out = out.sort_values(["Date", "Last Event"], ascending=[False, True]).reset_index(drop=True)
    return out[V2_TAB6_COLUMNS]




def _v2_build_order_event_tab7(events: pd.DataFrame) -> pd.DataFrame:
    if events is None or events.empty:
        return pd.DataFrame(columns=V2_TAB7_COLUMNS)

    df = events.copy()
    if "event_name" in df.columns:
        df = df[df["event_name"].isin(V2_EVENT_SEQUENCE)].copy()

    if df.empty:
        return pd.DataFrame(columns=V2_TAB7_COLUMNS)

    out = df.copy()
    out["Event"] = out["event_name"].astype(str)
    for col in V2_TAB7_COLUMNS:
        if col not in out.columns:
            out[col] = ""

    out = out[V2_TAB7_COLUMNS].copy()
    out = out.sort_values(["Date", "Email", "Event"], ascending=[False, True, True]).reset_index(drop=True)
    return out


def render_v2_dashboard(clean_audit: pd.DataFrame) -> None:
    events = _v2_clean_events_frame(clean_audit)

    journey_options = ["Manual", "Invoice Upload"]
    present_journey = sorted([x for x in events.get("Journey Type", pd.Series(dtype=str)).dropna().unique().tolist() if x and x != "No Status"])
    combined_journey = list(dict.fromkeys(journey_options + present_journey))

    invoice_options = ["Invoice Upload Success", "Invoice Upload Failure"]
    present_status = sorted([x for x in events.get("Invoice Status", pd.Series(dtype=str)).dropna().unique().tolist() if x and x != "No Status"])
    combined_status = list(dict.fromkeys(invoice_options + present_status))

    present_categories = sorted([x for x in events.get("Product Category", pd.Series(dtype=str)).dropna().unique().tolist() if x])
    combined_categories = list(dict.fromkeys(["All"] + present_categories))

    present_paid_sources = sorted([x for x in events.get("Paid Campaign Source", pd.Series(dtype=str)).dropna().unique().tolist() if x])
    combined_paid_sources = list(dict.fromkeys(V2_PAID_SOURCE_OPTIONS + present_paid_sources))

    present_traffic_sources = sorted([x for x in events.get("Traffic Source", pd.Series(dtype=str)).dropna().unique().tolist() if x])
    combined_traffic_sources = list(dict.fromkeys(V2_TRAFFIC_SOURCE_OPTIONS + present_traffic_sources))

    st.markdown("##### Filters")
    c1, c2, c3, c4, c5 = st.columns([1, 1, 1, 1, 1])
    selected_journey = c1.multiselect("Journey Type", combined_journey, default=combined_journey, label_visibility="collapsed", placeholder="Journey Type")
    selected_status = c2.multiselect("Invoice Status", combined_status, default=combined_status, label_visibility="collapsed", placeholder="Invoice Status")
    selected_categories = c3.multiselect("Product Category", combined_categories, default=["All"], label_visibility="collapsed", placeholder="Product Category")
    selected_paid_sources = c4.multiselect("Paid Campaign Source", combined_paid_sources, default=combined_paid_sources, label_visibility="collapsed", placeholder="Paid Campaign Source")
    selected_traffic_sources = c5.multiselect("Traffic Source", combined_traffic_sources, default=combined_traffic_sources, label_visibility="collapsed", placeholder="Traffic Source")

    filtered = events.copy()
    if selected_journey:
        filtered = filtered[filtered["Journey Type"].isin(selected_journey)]
    if selected_status:
        filtered = filtered[filtered["Invoice Status"].isin(selected_status)]

    filtered_tab2 = filtered.copy()
    if selected_categories and "All" not in selected_categories:
        filtered_tab2 = filtered_tab2[filtered_tab2["Product Category"].isin(selected_categories)]

    filtered_tab4 = filtered_tab2.copy()
    if selected_paid_sources:
        filtered_tab4 = filtered_tab4[filtered_tab4["Paid Campaign Source"].isin(selected_paid_sources)]

    filtered_tab6 = filtered_tab2.copy()
    if selected_traffic_sources:
        filtered_tab6 = filtered_tab6[filtered_tab6["Traffic Source"].isin(selected_traffic_sources)]

    tab1_view, tab2_view, tab3_view, tab4_view, tab5_view, tab6_view, tab7_view = st.tabs(["Tab 1: Daily Metrics", "Tab 2: Product Category Metrics", "Tab 3: Source Metrics", "Tab 4: Paid Campaign Metrics", "Tab 5: Category Metrics", "Tab 6: Detail View", "Tab 7: Order Event Detail"])

    with tab1_view:
        tab1 = _v2_build_daily_tab1(filtered)
        st.dataframe(tab1, use_container_width=True, hide_index=True)
        st.download_button(
            "Download V2 Tab 1 CSV",
            data=dataframe_to_csv_bytes(tab1),
            file_name="v2_daily_metrics_tab1.csv",
            mime="text/csv",
        )

    with tab2_view:
        tab2 = _v2_build_daily_tab2(filtered_tab2)
        st.dataframe(tab2, use_container_width=True, hide_index=True)
        st.download_button(
            "Download V2 Tab 2 CSV",
            data=dataframe_to_csv_bytes(tab2),
            file_name="v2_product_category_metrics_tab2.csv",
            mime="text/csv",
        )

    with tab3_view:
        tab3 = _v2_build_source_tab3(filtered_tab2)
        st.dataframe(tab3, use_container_width=True, hide_index=True)
        st.download_button(
            "Download V2 Tab 3 CSV",
            data=dataframe_to_csv_bytes(tab3),
            file_name="v2_source_metrics_tab3.csv",
            mime="text/csv",
        )

    with tab4_view:
        campaign_table, adset_table, ad_table = _v2_build_paid_campaign_tab4(filtered_tab4)

        st.markdown("###### 1. Campaign")
        st.dataframe(campaign_table, use_container_width=True, hide_index=True)
        st.download_button(
            "Download V2 Tab 4 Campaign CSV",
            data=dataframe_to_csv_bytes(campaign_table),
            file_name="v2_paid_campaign_metrics_campaign.csv",
            mime="text/csv",
        )

        st.markdown("###### 2. Campaign + Adset")
        st.dataframe(adset_table, use_container_width=True, hide_index=True)
        st.download_button(
            "Download V2 Tab 4 Adset CSV",
            data=dataframe_to_csv_bytes(adset_table),
            file_name="v2_paid_campaign_metrics_adset.csv",
            mime="text/csv",
        )

        st.markdown("###### 3. Campaign + Adset + Ad")
        st.dataframe(ad_table, use_container_width=True, hide_index=True)
        st.download_button(
            "Download V2 Tab 4 Ad CSV",
            data=dataframe_to_csv_bytes(ad_table),
            file_name="v2_paid_campaign_metrics_ad.csv",
            mime="text/csv",
        )

    with tab5_view:
        tab5 = _v2_build_category_tab5(filtered_tab2)
        st.dataframe(tab5, use_container_width=True, hide_index=True)
        st.download_button(
            "Download V2 Tab 5 CSV",
            data=dataframe_to_csv_bytes(tab5),
            file_name="v2_category_metrics_tab5.csv",
            mime="text/csv",
        )

    with tab6_view:
        tab6 = _v2_build_detail_tab6(filtered_tab6)
        st.dataframe(tab6, use_container_width=True, hide_index=True)
        st.download_button(
            "Download V2 Tab 6 CSV",
            data=dataframe_to_csv_bytes(tab6),
            file_name="v2_detail_view_tab6.csv",
            mime="text/csv",
        )

    with tab7_view:
        tab7 = _v2_build_order_event_tab7(filtered_tab4)
        st.dataframe(tab7, use_container_width=True, hide_index=True)
        st.download_button(
            "Download V2 Tab 7 CSV",
            data=dataframe_to_csv_bytes(tab7),
            file_name="v2_order_event_detail_tab7.csv",
            mime="text/csv",
        )


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="📊", layout="wide")
    st.markdown(
        """
        <style>
            .block-container {
                padding-top: 1.0rem;
                padding-bottom: 1.0rem;
            }
            [data-testid="stSidebar"] .block-container {
                padding-top: 0.75rem;
            }
            div[data-testid="stVerticalBlock"] {
                gap: 0.45rem;
            }
            h1 {
                margin-top: 0rem;
                margin-bottom: 0.35rem;
            }
            h2, h3 {
                margin-top: 0.45rem;
                margin-bottom: 0.35rem;
            }
            .stMarkdown p {
                margin-bottom: 0.35rem;
            }
            div[data-testid="stDataFrame"] {
                margin-top: 0.25rem;
            }
            img {
                object-fit: contain !important;
                max-height: none !important;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )

    logo_path = get_logo_path()
    if logo_path is not None:
        st.image(str(logo_path), width=190)
    st.markdown(f"### {APP_TITLE}")

    with st.sidebar:
        st.markdown("#### Supabase")

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
        st.markdown("#### Dashboard")
        dashboard_version = st.radio(
            "Open dashboard",
            ["V2: New Dashboard", "V1: Current Dashboard"],
            index=0,
            horizontal=True,
        )

        st.divider()
        st.markdown("#### Report settings")
        timezone_name = st.text_input("Timezone for date bucketing", value=DEFAULT_TIMEZONE)
        default_start = DEFAULT_V2_START_DATE if dashboard_version.startswith("V2") else DEFAULT_START_DATE
        start_date_value = st.date_input("Start date", value=default_start)
        use_end_date = st.checkbox("Use end date", value=False)
        end_date_value = st.date_input("End date", value=date.today()) if use_end_date else None

        st.divider()
        st.caption("Clean External filter")
        st.markdown(
            "- Includes `surebrightanywhere.com` production traffic\n"
            "- Excludes Abhishek, Santosh, Ankita, and soumyaramtri@gmail.com\n"
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

    if dashboard_version.startswith("V2"):
        render_v2_dashboard(clean_audit)
        return

    st.info("You are viewing V1: Current Dashboard. Use the sidebar to switch back to V2.")

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
                st.caption("Session-based requested funnel: Enquiry -> Sign Up -> First Quote -> Add to Cart -> Invoice Upload or Revised Offer -> Plan Selected -> Initiate Checkout -> Payment Attempted -> Payment Success. Node label text uses a bright color for readability.")
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
