#!/usr/bin/env python3
"""
SureBright D2C Clean External Metrics Report

Reads a raw D2C event export from Excel, CSV, or JSON-lines text and produces:
  1) Daily clean-external metrics Excel workbook
  2) Daily metrics CSV
  3) Clean-external event audit CSV

Clean External definition used by default:
  - production website traffic only: surebrightanywhere.com
  - excludes internal/test identities containing: abhishek, santosh
  - excludes email domains containing: surebright.com, surerbright.com, example.com
  - excludes obvious test URLs/domains: localhost, webflow, amplifyapp, _meta_test=1

Examples:
  python d2c_clean_external_metrics_report.py --input Surebright_D2C_Master_Data.xlsx --out-dir reports
  python d2c_clean_external_metrics_report.py --input Surebright_D2C_Master_Data.csv --out-dir reports --start-date 2026-05-21
  python d2c_clean_external_metrics_report.py --input raw_events.txt --out-dir reports --start-date 2026-06-02

Install dependencies:
  pip install pandas openpyxl
"""

from __future__ import annotations

import argparse
import csv
import json
import re
from copy import copy
from collections import Counter, defaultdict
from dataclasses import dataclass, asdict
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import urlparse

import pandas as pd

try:
    from zoneinfo import ZoneInfo
except Exception:  # pragma: no cover
    ZoneInfo = None

# -----------------------------
# Configuration
# -----------------------------

PROD_DOMAIN_KEYWORDS = ("surebrightanywhere.com",)
EXCLUDED_IDENTITY_TERMS = ("abhishek", "santosh")
EXCLUDED_EMAIL_DOMAINS = ("surebright.com", "surerbright.com", "example.com")
EXCLUDED_URL_TERMS = ("localhost", "webflow.io", "amplifyapp.com", "_meta_test=1")

USER_METRIC_EVENTS = {
    "Enquiry Attempted": {"homepage_form_submit"},
    "Sign Up_total": {"lead_signup", "quote_lead_captured"},
    "First Quote_Success": {"quote_generated"},
    "Offer_Selected": {"plan_selected"},
    "Invoice Upload_Success": {"invoice_uploaded"},
    "Invoice Upload_Failure": {"invoice_upload_failed"},
    "Revised Offer": {"revised_offer_shown"},
    "Additional Product": {"additional_product_detected"},
    "Add to Cart_Success": {"cart_confirmed"},
    "Payment Attempted": {"pay_now_clicked"},
    "Payment Success": {"payment_completed", "payment_success"},
    "Payment Failed": {"payment_failed"},
}

DAILY_COLUMN_ORDER = [
    "Date",
    "Enquiry Attempted",
    "Sign Up_total",
    "Sign Up_new User",
    "Sign Up_repeat user",
    "First Quote_Success",
    "Offer_Selected",
    "Invoice Upload_Success",
    "Invoice Upload_Failure",
    "Add details Later",
    "Revised Offer",
    "Additional Product",
    "Add to Cart_Success",
    "Payment Attempted",
    "Payment Success",
    "Payment Failed",
    "Email Verified_Success",
    "Email Verified_Failed",
    "Gross GWP $",
    "Invoice Success_Product Count",
    "Offer Generation_Success Count",
    "Add to Cart_Success Count",
    "Payment Success_Count",
]

# -----------------------------
# Helpers
# -----------------------------


def nested_get(obj: dict, *paths: str) -> Any:
    for path in paths:
        cur: Any = obj
        ok = True
        for part in path.split("."):
            if isinstance(cur, dict) and part in cur:
                cur = cur[part]
            else:
                ok = False
                break
        if ok and cur not in (None, ""):
            return cur
    return None


def collect_values_by_key(obj: Any, wanted_keys: set[str]) -> list[str]:
    out: list[str] = []
    if isinstance(obj, dict):
        for key, val in obj.items():
            norm_key = key.lower().replace("_", "")
            if norm_key in wanted_keys and val not in (None, ""):
                if isinstance(val, (str, int, float, bool)):
                    out.append(str(val))
            if isinstance(val, (dict, list)):
                out.extend(collect_values_by_key(val, wanted_keys))
    elif isinstance(obj, list):
        for item in obj:
            out.extend(collect_values_by_key(item, wanted_keys))
    return out


def parse_money(value: Any) -> float:
    if value in (None, ""):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = re.sub(r"[^0-9.\-]", "", str(value))
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def parse_dt(value: Any) -> datetime | None:
    if value in (None, ""):
        return None
    # numeric epoch values: legacy ts is generally milliseconds
    if isinstance(value, (int, float)):
        v = float(value)
        if v > 10_000_000_000:  # milliseconds
            return datetime.fromtimestamp(v / 1000, tz=timezone.utc)
        if v > 1_000_000_000:  # seconds
            return datetime.fromtimestamp(v, tz=timezone.utc)
    s = str(value).strip()
    if not s:
        return None
    if re.fullmatch(r"\d+(\.\d+)?", s):
        return parse_dt(float(s))
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00"))
    except ValueError:
        pass
    for fmt in (
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
    ):
        try:
            return datetime.strptime(s, fmt).replace(tzinfo=timezone.utc)
        except ValueError:
            continue
    return None


def json_dumps(obj: Any) -> str:
    return json.dumps(obj, ensure_ascii=False, separators=(",", ":"))


def event_type(obj: dict) -> str:
    return str(obj.get("event_type") or obj.get("event") or "").strip()


def page_url(obj: dict) -> str:
    return str(
        nested_get(
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


def page_domain(obj: dict) -> str:
    domain = nested_get(obj, "source.domain", "page.domain", "context.page.domain", "domain")
    if domain:
        return str(domain).lower().strip()
    url = page_url(obj)
    if url:
        return urlparse(url).netloc.lower().strip()
    return ""


def occurred_at(obj: dict, row_time: Any, tz) -> datetime | None:
    val = nested_get(
        obj,
        "occurred_at",
        "received_at",
        "timestamp",
        "event_timestamp",
        "created_at",
        "ts",
        "data.ts",
    )
    dt = parse_dt(val) or parse_dt(row_time)
    if not dt:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(tz)


def emails_in(obj: dict) -> list[str]:
    return [x.lower().strip() for x in collect_values_by_key(obj, {"email", "emailaddress"})]


def names_in(obj: dict) -> list[str]:
    return [
        x.lower().strip()
        for x in collect_values_by_key(
            obj,
            {"name", "firstname", "lastname", "fullname", "first", "last"},
        )
    ]


def identity_tokens(obj: dict) -> set[str]:
    tokens: set[str] = set()
    for path in (
        "actor.lead_id",
        "actor.customer_id",
        "customer.customer_id",
        "lead_id",
        "leadId",
        "data.leadId",
        "data.lead_id",
        "event_data.lead_id",
        "event_data.leadId",
        "event_data.checkout.lead_id",
        "session.anonymous_id",
        "anonymous_id",
        "session.session_id",
        "session_id",
    ):
        v = nested_get(obj, path)
        if v not in (None, ""):
            tokens.add(f"{path}:{str(v).lower().strip()}")
    for email in emails_in(obj):
        tokens.add(f"email:{email}")
    return tokens


def identity_key(obj: dict, row_num: int) -> str:
    # Prefer stable customer/user identifiers over session.
    for path in (
        "actor.customer_id",
        "customer.customer_id",
        "actor.email",
        "customer.email",
        "data.email",
        "email",
        "actor.lead_id",
        "lead_id",
        "leadId",
        "data.leadId",
        "data.lead_id",
        "event_data.lead_id",
        "event_data.checkout.lead_id",
    ):
        v = nested_get(obj, path)
        if v not in (None, ""):
            return f"{path}:{str(v).lower().strip()}"

    # Some post-invoice events can have a customer name/address but no email/lead ID.
    # Use that before anonymous/session to dedupe the same user across repeated checkout sessions.
    name = nested_get(obj, "customer.full_name", "actor.name", "data.fullName", "data.full_name", "name")
    phone = nested_get(obj, "customer.phone", "actor.phone", "phone")
    addr = nested_get(obj, "customer.address.line_1", "customer.address.postal_code")
    if name not in (None, "") and (phone not in (None, "") or addr not in (None, "")):
        return "name_contact:" + re.sub(r"\s+", " ", f"{name}|{phone or ''}|{addr or ''}".lower()).strip()

    for path in (
        "session.anonymous_id",
        "anonymous_id",
        "session.session_id",
        "session_id",
    ):
        v = nested_get(obj, path)
        if v not in (None, ""):
            return f"{path}:{str(v).lower().strip()}"
    return f"row:{row_num}"


def is_prod_website(obj: dict) -> bool:
    text = f"{page_url(obj)} {page_domain(obj)}".lower()
    return any(d in text for d in PROD_DOMAIN_KEYWORDS)


def has_test_url(obj: dict) -> bool:
    text = f"{page_url(obj)} {page_domain(obj)}".lower()
    return any(term in text for term in EXCLUDED_URL_TERMS)


def has_excluded_identity(obj: dict) -> bool:
    names_text = " ".join(names_in(obj))
    emails_text = " ".join(emails_in(obj))
    if any(term in names_text for term in EXCLUDED_IDENTITY_TERMS):
        return True
    if any(term in emails_text for term in EXCLUDED_IDENTITY_TERMS):
        return True
    if any(domain in emails_text for domain in EXCLUDED_EMAIL_DOMAINS):
        return True
    return False


def get_line_items(obj: dict) -> list[dict]:
    for path in (
        "event_data.line_items",
        "event_data.lineItems",
        "line_items",
        "lineItems",
        "data.line_items",
        "data.lineItems",
        "products",
        "data.products",
        "event_data.products",
    ):
        val = nested_get(obj, path)
        if isinstance(val, list):
            return [x for x in val if isinstance(x, dict)]
    return []


def item_quantity(item: dict) -> int:
    q = nested_get(item, "purchase.quantity", "quantity", "qty")
    try:
        return max(1, int(float(q))) if q not in (None, "") else 1
    except Exception:
        return 1


def is_eligible_item(item: dict) -> bool:
    val = nested_get(item, "protection.eligible", "eligible", "isEligible", "is_eligible")
    if isinstance(val, bool):
        return val
    if isinstance(val, str):
        return val.strip().lower() in {"true", "yes", "1", "eligible"}
    # If no eligibility flag exists, treat items in successful invoice/cart/payment contexts as eligible.
    return True


def invoice_success_product_count(obj: dict) -> int:
    return sum(item_quantity(i) for i in get_line_items(obj) if is_eligible_item(i))


def cart_or_payment_product_count(obj: dict) -> int:
    items = get_line_items(obj)
    if items:
        return sum(item_quantity(i) for i in items)
    raw_items = nested_get(obj, "raw.items", "data.items")
    if isinstance(raw_items, list) and raw_items:
        total = 0
        for item in raw_items:
            if isinstance(item, dict):
                total += item_quantity(item)
        return total or len(raw_items)
    item_count = nested_get(obj, "event_data.checkout.item_count", "checkout.item_count", "item_count")
    try:
        return int(float(item_count)) if item_count not in (None, "") else 1
    except Exception:
        return 1


def offer_generation_product_count(obj: dict) -> int:
    # Business definition is product/enquiry count, not number of plans shown.
    return 1


def gwp_for_payment_success(obj: dict) -> float:
    if event_type(obj) not in {"payment_completed", "payment_success"}:
        return 0.0
    # Prefer warranty-specific and checkout fields. Avoid legacy product totalAmount unless it is clearly warranty revenue.
    candidates = [
        nested_get(obj, "event_data.checkout.subtotal_amount"),
        nested_get(obj, "event_data.checkout.total_amount"),
        nested_get(obj, "checkout.subtotal_amount"),
        nested_get(obj, "checkout.total_amount"),
        nested_get(obj, "raw.value"),
        nested_get(obj, "event_data.gwp"),
        nested_get(obj, "gwp"),
        nested_get(obj, "warrantyPrice"),
        nested_get(obj, "warranty_price"),
        nested_get(obj, "event_data.warrantyPrice"),
        nested_get(obj, "event_data.warranty_price"),
    ]
    for c in candidates:
        amount = parse_money(c)
        if amount > 0:
            return amount
    total = 0.0
    for item in get_line_items(obj):
        total += parse_money(nested_get(item, "protection.plan_price", "plan_price", "warrantyPrice")) * item_quantity(item)
    return total


@dataclass
class NormEvent:
    source_row: int
    event_id: str
    event_time: str
    date: str
    event_name: str
    identity_key: str
    domain: str
    page_url: str
    email: str
    name: str
    session_id: str
    lead_id: str
    source_component: str
    product_count: int
    eligible_product_count: int
    gwp: float
    excluded: bool
    exclusion_reason: str
    raw_json: str


# -----------------------------
# Input loading
# -----------------------------


def iter_raw_events(input_path: Path) -> Iterable[tuple[int, Any, str]]:
    """Yield (row_number, row_time, raw_json_text)."""
    suffix = input_path.suffix.lower()

    if suffix in {".xlsx", ".xls"}:
        df = pd.read_excel(input_path, dtype=str)
        raw_col = next((c for c in df.columns if str(c).strip().lower() == "raw"), None)
        time_col = next((c for c in df.columns if str(c).strip().lower() in {"time", "timestamp", "occurred_at"}), None)
        if raw_col is None:
            # Try the first column as JSON text if no Raw column exists.
            raw_col = df.columns[0]
        for idx, row in df.iterrows():
            raw = row.get(raw_col)
            if pd.isna(raw) or str(raw).strip() == "":
                continue
            yield int(idx) + 2, row.get(time_col) if time_col else None, str(raw).strip()
        return

    if suffix == ".csv":
        df = pd.read_csv(input_path, dtype=str)
        raw_col = next((c for c in df.columns if str(c).strip().lower() == "raw"), None)
        time_col = next((c for c in df.columns if str(c).strip().lower() in {"time", "timestamp", "occurred_at"}), None)
        if raw_col is None:
            raw_col = df.columns[0]
        for idx, row in df.iterrows():
            raw = row.get(raw_col)
            if pd.isna(raw) or str(raw).strip() == "":
                continue
            yield int(idx) + 2, row.get(time_col) if time_col else None, str(raw).strip()
        return

    # JSONL / pasted text: one JSON object per line.
    with input_path.open("r", encoding="utf-8-sig") as f:
        for line_num, line in enumerate(f, start=1):
            line = line.strip()
            if not line:
                continue
            yield line_num, None, line


# -----------------------------
# Event normalization + filtering
# -----------------------------


def normalize_all_events(input_path: Path, tz) -> list[NormEvent]:
    events: list[NormEvent] = []
    seen_event_ids: set[str] = set()

    # First pass: normalize each parseable event, but do not yet apply linked identity exclusions.
    tmp: list[tuple[NormEvent, set[str], bool, str]] = []
    token_to_indices: dict[str, list[int]] = defaultdict(list)

    for row_num, row_time, raw_text in iter_raw_events(input_path):
        try:
            obj = json.loads(raw_text)
        except json.JSONDecodeError:
            continue
        ev = event_type(obj)
        if not ev:
            continue
        event_id = str(obj.get("event_id") or "")
        if event_id and event_id in seen_event_ids:
            continue
        if event_id:
            seen_event_ids.add(event_id)

        dt = occurred_at(obj, row_time, tz)
        if not dt:
            continue

        tokens = identity_tokens(obj)
        ident = identity_key(obj, row_num)
        if ident:
            tokens.add(ident)

        email = ", ".join(sorted(set(emails_in(obj))))
        name = ", ".join(sorted(set(names_in(obj))))
        source_component = str(nested_get(obj, "source.component", "raw.source", "data.source") or "")
        lead_id = str(nested_get(obj, "actor.lead_id", "lead_id", "leadId", "data.leadId", "event_data.lead_id", "event_data.checkout.lead_id") or "")
        session_id = str(nested_get(obj, "session.session_id", "session_id") or "")

        base_excluded = False
        reason_parts = []
        if not is_prod_website(obj):
            base_excluded = True
            reason_parts.append("not_prod_domain")
        if has_test_url(obj):
            base_excluded = True
            reason_parts.append("test_url_or_domain")
        if has_excluded_identity(obj):
            base_excluded = True
            reason_parts.append("excluded_identity_or_email")

        if ev == "invoice_uploaded":
            eligible_product_count = invoice_success_product_count(obj)
        else:
            eligible_product_count = 0

        if ev in {"cart_confirmed", "payment_completed", "payment_success", "pay_now_clicked", "payment_failed", "plan_selected"}:
            prod_count = cart_or_payment_product_count(obj)
        elif ev == "quote_generated":
            prod_count = offer_generation_product_count(obj)
        elif ev == "invoice_uploaded":
            prod_count = invoice_success_product_count(obj)
        else:
            prod_count = 0

        norm = NormEvent(
            source_row=row_num,
            event_id=event_id,
            event_time=dt.isoformat(),
            date=dt.date().isoformat(),
            event_name=ev,
            identity_key=ident,
            domain=page_domain(obj),
            page_url=page_url(obj),
            email=email,
            name=name,
            session_id=session_id,
            lead_id=lead_id,
            source_component=source_component,
            product_count=prod_count,
            eligible_product_count=eligible_product_count,
            gwp=round(gwp_for_payment_success(obj), 2),
            excluded=base_excluded,
            exclusion_reason=";".join(reason_parts),
            raw_json=json_dumps(obj),
        )
        idx = len(tmp)
        tmp.append((norm, tokens, base_excluded, ";".join(reason_parts)))
        for token in tokens:
            token_to_indices[token].append(idx)

    # Second pass: propagate identity exclusions to linked events.
    excluded_indices = {i for i, (_, _, excluded, _) in enumerate(tmp) if excluded}
    changed = True
    while changed:
        changed = False
        excluded_tokens = set()
        for i in excluded_indices:
            excluded_tokens.update(tmp[i][1])
        for token in excluded_tokens:
            for i in token_to_indices.get(token, []):
                if i not in excluded_indices:
                    excluded_indices.add(i)
                    changed = True

    for i, (norm, _tokens, _excluded, _reason) in enumerate(tmp):
        if i in excluded_indices:
            if not norm.excluded:
                norm.excluded = True
                norm.exclusion_reason = "linked_to_excluded_identity_or_session"
        events.append(norm)

    return events


# -----------------------------
# Metrics
# -----------------------------


def is_repeat_signup_map(clean_events: list[NormEvent]) -> dict[tuple[str, str], str]:
    """Returns per (date, identity) signup type: new or repeat."""
    signup_events = sorted(
        [e for e in clean_events if e.event_name in USER_METRIC_EVENTS["Sign Up_total"]],
        key=lambda e: e.event_time,
    )
    first_seen: dict[str, str] = {}
    result: dict[tuple[str, str], str] = {}
    for e in signup_events:
        key = e.identity_key
        if key not in first_seen:
            first_seen[key] = e.date
            result[(e.date, key)] = "new"
        else:
            # Same user signing up again later in the same file/window.
            result[(e.date, key)] = "repeat"
    return result


def count_unique(clean_events: list[NormEvent], day: str, event_names: set[str]) -> int:
    return len({e.identity_key for e in clean_events if e.date == day and e.event_name in event_names})


def build_daily_metrics(clean_events: list[NormEvent], start_date: str | None = None, end_date: str | None = None) -> pd.DataFrame:
    if start_date:
        clean_events = [e for e in clean_events if e.date >= start_date]
    if end_date:
        clean_events = [e for e in clean_events if e.date <= end_date]

    dates = sorted({e.date for e in clean_events})
    signup_type = is_repeat_signup_map(clean_events)
    rows: list[dict[str, Any]] = []

    for d in dates:
        day_events = [e for e in clean_events if e.date == d]
        row: dict[str, Any] = {"Date": d}
        for metric, event_names in USER_METRIC_EVENTS.items():
            row[metric] = count_unique(clean_events, d, event_names)

        signup_identities = {e.identity_key for e in day_events if e.event_name in USER_METRIC_EVENTS["Sign Up_total"]}
        row["Sign Up_new User"] = sum(1 for ident in signup_identities if signup_type.get((d, ident)) == "new")
        row["Sign Up_repeat user"] = sum(1 for ident in signup_identities if signup_type.get((d, ident)) == "repeat")

        row["Add details Later"] = 0
        row["Email Verified_Success"] = 0
        row["Email Verified_Failed"] = 0
        row["Gross GWP $"] = round(sum(e.gwp for e in day_events if e.event_name in {"payment_completed", "payment_success"}), 2)

        row["Invoice Success_Product Count"] = sum(e.eligible_product_count for e in day_events if e.event_name == "invoice_uploaded")
        row["Offer Generation_Success Count"] = sum(e.product_count for e in day_events if e.event_name == "quote_generated")
        row["Add to Cart_Success Count"] = sum(e.product_count for e in day_events if e.event_name == "cart_confirmed")
        row["Payment Success_Count"] = sum(e.product_count for e in day_events if e.event_name in {"payment_completed", "payment_success"})

        rows.append(row)

    df = pd.DataFrame(rows)
    if df.empty:
        df = pd.DataFrame(columns=DAILY_COLUMN_ORDER)
    else:
        for col in DAILY_COLUMN_ORDER:
            if col not in df.columns:
                df[col] = 0
        df = df[DAILY_COLUMN_ORDER]
    return df


def build_totals(daily: pd.DataFrame) -> pd.DataFrame:
    if daily.empty:
        return pd.DataFrame(columns=["Metric", "Value"])
    rows = []
    for col in DAILY_COLUMN_ORDER:
        if col == "Date":
            continue
        rows.append({"Metric": col, "Value": round(float(daily[col].sum()), 2) if "GWP" in col else int(daily[col].sum())})
    return pd.DataFrame(rows)


# -----------------------------
# Output
# -----------------------------


def write_outputs(
    out_dir: Path,
    base_name: str,
    daily: pd.DataFrame,
    totals: pd.DataFrame,
    all_events: list[NormEvent],
    clean_events: list[NormEvent],
    input_path: Path,
    start_date: str | None,
    end_date: str | None,
    timezone_name: str,
) -> tuple[Path, Path, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    metrics_csv = out_dir / f"{base_name}_daily_metrics.csv"
    audit_csv = out_dir / f"{base_name}_event_audit.csv"
    xlsx_path = out_dir / f"{base_name}_metrics.xlsx"

    daily.to_csv(metrics_csv, index=False)

    audit_rows = [asdict(e) for e in all_events]
    audit_df = pd.DataFrame(audit_rows)
    audit_df.to_csv(audit_csv, index=False)

    metadata = pd.DataFrame(
        [
            ["input_file", str(input_path)],
            ["timezone", timezone_name],
            ["start_date_inclusive", start_date or "all"],
            ["end_date_inclusive", end_date or "all"],
            ["total_parseable_events", len(all_events)],
            ["clean_external_events_in_window", len(clean_events)],
            ["clean_external_definition", "prod surebrightanywhere.com excluding internal/test identities, email domains, and test URLs"],
            ["generated_at_utc", datetime.now(timezone.utc).isoformat()],
        ],
        columns=["Field", "Value"],
    )

    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        daily.to_excel(writer, index=False, sheet_name="Daily Metrics")
        totals.to_excel(writer, index=False, sheet_name="Totals")
        audit_df.to_excel(writer, index=False, sheet_name="Event Audit")
        metadata.to_excel(writer, index=False, sheet_name="Run Metadata")

        wb = writer.book
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
            for cell in ws[1]:
                new_font = copy(cell.font)
                new_font.bold = True
                new_font.color = "FFFFFF"
                cell.font = new_font

                new_fill = copy(cell.fill)
                new_fill.fill_type = "solid"
                new_fill.fgColor = "1F4E78"
                cell.fill = new_fill

                new_alignment = copy(cell.alignment)
                new_alignment.horizontal = "center"
                cell.alignment = new_alignment
            for col_cells in ws.columns:
                max_len = min(max(len(str(c.value or "")) for c in col_cells[:300]) + 2, 55)
                ws.column_dimensions[col_cells[0].column_letter].width = max(10, max_len)
            if ws.title == "Daily Metrics":
                for cell in ws[1]:
                    if cell.value == "Gross GWP $":
                        col = cell.column_letter
                        for c in ws[f"{col}2":f"{col}{ws.max_row}"]:
                            c[0].number_format = '$#,##0.00'
            if ws.title == "Totals":
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                    if row[0].value == "Gross GWP $":
                        row[1].number_format = '$#,##0.00'

    return xlsx_path, metrics_csv, audit_csv


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Build SureBright D2C Clean External metrics from raw event Excel/CSV/TXT source.")
    p.add_argument("--input", required=True, help="Path to source .xlsx, .csv, or JSON-lines .txt file")
    p.add_argument("--out-dir", default="d2c_metrics_output", help="Output folder")
    p.add_argument("--start-date", default="2026-05-21", help="Inclusive local start date YYYY-MM-DD. Use empty string for all dates.")
    p.add_argument("--end-date", default=None, help="Inclusive local end date YYYY-MM-DD")
    p.add_argument("--timezone", default="Asia/Kolkata", help="Timezone for date bucketing")
    p.add_argument("--base-name", default=None, help="Output filename prefix. Default derived from input name")
    return p.parse_args()


def main() -> int:
    args = parse_args()
    input_path = Path(args.input)
    if not input_path.exists():
        raise FileNotFoundError(input_path)
    tz = ZoneInfo(args.timezone) if ZoneInfo else timezone.utc
    start_date = args.start_date or None
    end_date = args.end_date or None

    all_events = normalize_all_events(input_path, tz)
    clean_events = [e for e in all_events if not e.excluded]
    if start_date:
        clean_events = [e for e in clean_events if e.date >= start_date]
    if end_date:
        clean_events = [e for e in clean_events if e.date <= end_date]

    daily = build_daily_metrics(clean_events, start_date=None, end_date=None)
    totals = build_totals(daily)

    base_name = args.base_name or re.sub(r"[^A-Za-z0-9_\-]+", "_", input_path.stem).strip("_") + "_clean_external"
    xlsx, metrics_csv, audit_csv = write_outputs(
        Path(args.out_dir),
        base_name,
        daily,
        totals,
        all_events,
        clean_events,
        input_path,
        start_date,
        end_date,
        args.timezone,
    )

    print(f"Parseable events: {len(all_events)}")
    print(f"Clean external events in window: {len(clean_events)}")
    print(f"Daily rows: {len(daily)}")
    print(f"Gross GWP total: ${daily['Gross GWP $'].sum() if not daily.empty else 0:,.2f}")
    print(f"Excel: {xlsx}")
    print(f"Daily CSV: {metrics_csv}")
    print(f"Audit CSV: {audit_csv}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
