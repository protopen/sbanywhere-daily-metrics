#!/usr/bin/env python3
"""
SureBright D2C Clean External Metrics Report

Reads a raw D2C event export from Excel, CSV, or JSON-lines text and produces:
  1) Daily clean-external metrics Excel workbook
  2) Daily metrics CSV
  3) Clean-external event audit CSV
  4) Attribution, Product, and Retailer statistics tabs

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
from urllib.parse import urlparse, parse_qs

import pandas as pd

try:
    from zoneinfo import ZoneInfo
except Exception:  # pragma: no cover
    ZoneInfo = None

# -----------------------------
# Configuration
# -----------------------------

PROD_DOMAIN_KEYWORDS = ("surebrightanywhere.com",)
EXCLUDED_IDENTITY_TERMS = ("abhishek", "santosh", "ankita",
    "soumya ramtri", "soumyaramtri@gmail.com",
    "ramtrisoumya11@gmail.com")
EXCLUDED_EMAIL_DOMAINS = ("surebright.com", "surerbright.com", "example.com")
EXCLUDED_URL_TERMS = ("localhost", "webflow.io", "amplifyapp.com", "_meta_test=1")

# Metric event mapping supports both legacy names and the revised payload names.
USER_METRIC_EVENTS = {
    "Enquiry Attempted": {"homepage_form_submit", "enquiry_attempted"},
    "Sign Up_total": {"lead_signup", "quote_lead_captured", "signup_completed", "sign_up_completed", "sign_up_total", "signup_total", "sign_up"},
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

# Compatibility aliases used by Product Stats and older app builds.
SIGNUP_EVENTS = USER_METRIC_EVENTS["Sign Up_total"]
PAYMENT_ATTEMPT_EVENTS = USER_METRIC_EVENTS["Payment Attempted"]

INVOICE_SUCCESS_EVENTS = USER_METRIC_EVENTS["Invoice Upload_Success"]
INVOICE_FAILURE_EVENTS = USER_METRIC_EVENTS["Invoice Upload_Failure"]
QUOTE_SUCCESS_EVENTS = USER_METRIC_EVENTS["First Quote_Success"]
OFFER_SELECTED_EVENTS = USER_METRIC_EVENTS["Offer_Selected"]
ADD_TO_CART_EVENTS = USER_METRIC_EVENTS["Add to Cart_Success"]
PAYMENT_ATTEMPTED_EVENTS = USER_METRIC_EVENTS["Payment Attempted"]
PAYMENT_SUCCESS_EVENTS = USER_METRIC_EVENTS["Payment Success"]
PAYMENT_FAILED_EVENTS = USER_METRIC_EVENTS["Payment Failed"]
REVISED_OFFER_EVENTS = USER_METRIC_EVENTS["Revised Offer"]
ENQUIRY_EVENTS = USER_METRIC_EVENTS["Enquiry Attempted"]

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
    if event_type(obj) not in PAYMENT_SUCCESS_EVENTS:
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




def read_csv_with_fallback(input_path: Path) -> pd.DataFrame:
    """Read CSV exports that may be UTF-8, Windows encoded, or slightly malformed.

    OneDrive sometimes serves CSV-like content with Windows characters or quoting
    differences. We first try strict parsing, then the Python CSV engine, then a
    conservative manual parser for the expected two-column Time,Raw export.
    """
    import csv

    data = input_path.read_bytes()
    encodings = ["utf-8-sig", "utf-8", "cp1252", "latin1"]
    last_error = None

    for encoding in encodings:
        try:
            return pd.read_csv(input_path, dtype=str, encoding=encoding)
        except (UnicodeDecodeError, pd.errors.ParserError, csv.Error) as exc:
            last_error = exc

    for encoding in encodings:
        try:
            return pd.read_csv(
                input_path,
                dtype=str,
                encoding=encoding,
                engine="python",
                sep=",",
                quotechar='"',
                escapechar="\\",
                on_bad_lines="warn",
            )
        except Exception as exc:
            last_error = exc

    # Last resort for the common D2C export: each row is Time,<JSON payload>.
    # This avoids pandas failing on embedded quotes when OneDrive rewrites the CSV.
    for encoding in encodings:
        try:
            text = data.decode(encoding)
        except UnicodeDecodeError as exc:
            last_error = exc
            continue

        rows = []
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        if lines and lines[0].lower().replace('"', '').startswith("time,raw"):
            lines = lines[1:]
        for line in lines:
            comma = line.find(",")
            if comma == -1:
                continue
            time_part = line[:comma].strip().strip('"')
            raw_part = line[comma + 1 :].strip()
            if raw_part.startswith('"') and raw_part.endswith('"'):
                raw_part = raw_part[1:-1].replace('""', '"')
            rows.append({"Time": time_part, "Raw": raw_part})
        if rows:
            return pd.DataFrame(rows)

    raise last_error if last_error else RuntimeError(f"Could not parse CSV: {input_path}")


def iter_text_lines_with_fallback(input_path: Path):
    """Yield lines from text/jsonl exports with robust encoding fallback."""
    data = input_path.read_bytes()
    encodings = ["utf-8-sig", "utf-8", "cp1252", "latin1"]
    last_error = None
    for encoding in encodings:
        try:
            text = data.decode(encoding)
            for line in text.splitlines():
                yield line
            return
        except UnicodeDecodeError as exc:
            last_error = exc
    # Final defensive fallback replaces only undecodable characters.
    text = data.decode("utf-8", errors="replace")
    for line in text.splitlines():
        yield line

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
        df = read_csv_with_fallback(input_path)
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
    for line_num, line in enumerate(iter_text_lines_with_fallback(input_path), start=1):
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

        if ev in INVOICE_SUCCESS_EVENTS:
            eligible_product_count = invoice_success_product_count(obj)
        else:
            eligible_product_count = 0

        if ev in (ADD_TO_CART_EVENTS | PAYMENT_SUCCESS_EVENTS | PAYMENT_ATTEMPTED_EVENTS | PAYMENT_FAILED_EVENTS | OFFER_SELECTED_EVENTS):
            prod_count = cart_or_payment_product_count(obj)
        elif ev in QUOTE_SUCCESS_EVENTS:
            prod_count = offer_generation_product_count(obj)
        elif ev in INVOICE_SUCCESS_EVENTS:
            prod_count = invoice_success_product_count(obj)
        elif ev in ENQUIRY_EVENTS:
            prod_count = 1
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
        row["Gross GWP $"] = round(sum(e.gwp for e in day_events if e.event_name in PAYMENT_SUCCESS_EVENTS), 2)

        row["Invoice Success_Product Count"] = sum(e.eligible_product_count for e in day_events if e.event_name in INVOICE_SUCCESS_EVENTS)
        row["Offer Generation_Success Count"] = sum(e.product_count for e in day_events if e.event_name in QUOTE_SUCCESS_EVENTS)
        row["Add to Cart_Success Count"] = sum(e.product_count for e in day_events if e.event_name in ADD_TO_CART_EVENTS)
        row["Payment Success_Count"] = sum(e.product_count for e in day_events if e.event_name in PAYMENT_SUCCESS_EVENTS)

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



def primary_email(e: NormEvent) -> str:
    return (e.email.split(",")[0].strip().lower() if e.email else e.identity_key)


def _safe_json(e: NormEvent) -> dict:
    try:
        return json.loads(e.raw_json)
    except Exception:
        return {}


def _query_params_from_url(url: str) -> dict[str, str]:
    if not url:
        return {}
    try:
        qs = parse_qs(urlparse(url).query, keep_blank_values=True)
        return {k: (v[0] if v else "") for k, v in qs.items()}
    except Exception:
        return {}


def _first_non_empty(*values: Any) -> str:
    for value in values:
        if value not in (None, ""):
            return str(value).strip()
    return ""


def attribution_values(obj: dict) -> dict[str, str]:
    attr = nested_get(obj, "source.attribution")
    if not isinstance(attr, dict):
        attr = {}
    utm = nested_get(obj, "source.utm")
    if not isinstance(utm, dict):
        utm = {}

    source_referrer = str(nested_get(obj, "source.referrer") or "").strip()
    page = page_url(obj)
    landing_page_raw = str(attr.get("landing_page") or page or "").strip()
    landing_qs = _query_params_from_url(landing_page_raw)
    page_qs = _query_params_from_url(page)
    referrer_qs = _query_params_from_url(source_referrer)

    def p(name: str) -> str:
        return _first_non_empty(
            attr.get(name),
            utm.get(name.replace("utm_", "")),
            landing_qs.get(name),
            page_qs.get(name),
            referrer_qs.get(name),
        )

    referrer = _first_non_empty(attr.get("referrer"), source_referrer)
    landing_page = landing_page_raw
    gclid = _first_non_empty(attr.get("gclid"), landing_qs.get("gclid"), page_qs.get("gclid"), referrer_qs.get("gclid"))
    fbclid = _first_non_empty(attr.get("fbclid"), landing_qs.get("fbclid"), page_qs.get("fbclid"), referrer_qs.get("fbclid"))
    gbraid = _first_non_empty(attr.get("gbraid"), landing_qs.get("gbraid"), page_qs.get("gbraid"), referrer_qs.get("gbraid"))
    wbraid = _first_non_empty(attr.get("wbraid"), landing_qs.get("wbraid"), page_qs.get("wbraid"), referrer_qs.get("wbraid"))

    utm_source = p("utm_source")
    utm_medium = p("utm_medium")
    utm_campaign = p("utm_campaign")
    utm_term = p("utm_term")
    utm_content = p("utm_content")

    # Extra fields for Meta/paid social hierarchy. Fallbacks map common UTM conventions:
    # campaign = utm_campaign, ad set = utm_term, ad = utm_content.
    adset = _first_non_empty(
        attr.get("utm_adset"), attr.get("adset"), attr.get("ad_set"),
        attr.get("adset_name"), attr.get("ad_set_name"), attr.get("adgroup"), attr.get("ad_group"),
        utm.get("adset"), utm.get("ad_set"), utm.get("adset_name"), utm.get("ad_group"),
        landing_qs.get("utm_adset"), landing_qs.get("adset"), landing_qs.get("ad_set"), landing_qs.get("adset_name"), landing_qs.get("ad_set_name"),
        page_qs.get("utm_adset"), page_qs.get("adset"), page_qs.get("ad_set"), page_qs.get("adset_name"), page_qs.get("ad_set_name"),
        referrer_qs.get("utm_adset"), referrer_qs.get("adset"), referrer_qs.get("ad_set"), referrer_qs.get("adset_name"), referrer_qs.get("ad_set_name"),
        utm_term,
    )
    ad = _first_non_empty(
        attr.get("utm_ad"), attr.get("ad"), attr.get("ad_name"), attr.get("creative"), attr.get("creative_name"),
        utm.get("ad"), utm.get("ad_name"), utm.get("creative"), utm.get("creative_name"),
        landing_qs.get("utm_ad"), landing_qs.get("ad"), landing_qs.get("ad_name"), landing_qs.get("creative"), landing_qs.get("creative_name"),
        page_qs.get("utm_ad"), page_qs.get("ad"), page_qs.get("ad_name"), page_qs.get("creative"), page_qs.get("creative_name"),
        referrer_qs.get("utm_ad"), referrer_qs.get("ad"), referrer_qs.get("ad_name"), referrer_qs.get("creative"), referrer_qs.get("creative_name"),
        utm_content,
    )

    if not utm_source:
        if gclid or gbraid or wbraid:
            utm_source = "google_ads"
        elif fbclid:
            utm_source = "fb"
        elif referrer:
            try:
                host = urlparse(referrer).netloc.lower().replace("www.", "")
                utm_source = host or "referral"
            except Exception:
                utm_source = "referral"
        else:
            utm_source = "direct"
    if not utm_medium:
        if gclid or gbraid or wbraid or fbclid:
            utm_medium = "paid"
        elif referrer:
            utm_medium = "referral"
        else:
            utm_medium = "none"

    return {
        "channel": str(nested_get(obj, "source.channel") or "unknown"),
        "utm_source": utm_source or "unknown",
        "utm_medium": utm_medium or "unknown",
        "utm_campaign": utm_campaign or "(none)",
        "utm_term": utm_term or "",
        "utm_content": utm_content or "",
        "ad_set": adset or "(not set)",
        "ad": ad or "(not set)",
        "gclid": gclid,
        "gbraid": gbraid,
        "wbraid": wbraid,
        "fbclid": fbclid,
        "referrer": referrer,
        "landing_page": landing_page,
        "first_seen_at": str(attr.get("first_seen_at") or ""),
        "source_component": str(nested_get(obj, "source.component", "raw.source", "data.source") or ""),
    }

def retailer_values(obj: dict) -> dict[str, str]:
    retailer = nested_get(obj, "event_data.invoice.retailer", "invoice.retailer")
    if not isinstance(retailer, dict):
        retailer = {}
    address = retailer.get("address") if isinstance(retailer.get("address"), dict) else {}
    name = retailer.get("name") or nested_get(obj, "event_data.invoice.retailer_name", "event_data.invoice.retailerName", "raw.retailer_detected") or "Unknown"
    return {
        "retailer_name": str(name or "Unknown"),
        "retailer_email": str(retailer.get("email") or ""),
        "retailer_phone": str(retailer.get("phone") or ""),
        "retailer_city": str(address.get("city") or ""),
        "retailer_state": str(address.get("state") or ""),
        "retailer_country": str(address.get("country") or ""),
    }



def product_email_for_event(e: NormEvent) -> str:
    """Prefer actor.email for Product Stats, then fall back to normalized email extraction."""
    obj = _safe_json(e)
    actor_email = nested_get(obj, "actor.email")
    if actor_email:
        return str(actor_email).strip()
    return primary_email(e)


def product_rows_for_event(e: NormEvent) -> list[dict[str, Any]]:
    obj = _safe_json(e)
    rows: list[dict[str, Any]] = []
    attr = attribution_values(obj)
    retailer = retailer_values(obj)

    items = get_line_items(obj)
    if not items:
        raw_items = nested_get(obj, "raw.items", "data.items")
        if isinstance(raw_items, list):
            items = [i for i in raw_items if isinstance(i, dict)]

    if items:
        for item in items:
            product = item.get("product") if isinstance(item.get("product"), dict) else {}
            purchase = item.get("purchase") if isinstance(item.get("purchase"), dict) else {}
            protection = item.get("protection") if isinstance(item.get("protection"), dict) else {}
            category = product.get("category") if isinstance(product.get("category"), dict) else {}
            qty = item_quantity(item)
            rows.append({
                "date": e.date,
                "event_name": e.event_name,
                "identity_key": e.identity_key,
                "email": product_email_for_event(e),
                "product_category": str(category.get("name") or product.get("category") or item.get("item_category") or item.get("category") or "Unknown"),
                "product_title": str(product.get("title") or item.get("item_name") or item.get("name") or product.get("description") or "Unknown"),
                "product_brand": str(product.get("brand") or nested_get(obj, "manufacturer.name") or ""),
                "manufacturer_name": str(nested_get(obj, "manufacturer.name") or product.get("brand") or ""),
                "model_number": str(nested_get(obj, "manufacturer.model_number") or product.get("sku") or ""),
                "product_condition": str(product.get("condition") or ""),
                "quantity": qty,
                "product_unit_price": parse_money(purchase.get("unit_price") or item.get("unit_price") or item.get("price")),
                "product_total_price": parse_money(purchase.get("total_price") or item.get("total_price")),
                "plan_name": str(protection.get("plan_name") or item.get("item_variant") or ""),
                "plan_price": parse_money(protection.get("plan_price") or item.get("price")),
                "plan_term_months": protection.get("term_months") or "",
                "eligible": is_eligible_item(item),
                "retailer_name": retailer["retailer_name"],
                "utm_source": attr["utm_source"],
                "utm_medium": attr["utm_medium"],
                "utm_campaign": attr["utm_campaign"],
                "gwp": e.gwp,
            })
        return rows

    # Form-only enquiry rows still matter for product/category attribution.
    category = nested_get(obj, "event_data.form.fields.product_category", "raw.product_category", "data.categoryName", "data.sbCategoryName")
    price = nested_get(obj, "event_data.form.fields.product_price", "raw.product_price", "data.productPrice", "raw.value")
    if category or e.event_name in ENQUIRY_EVENTS:
        rows.append({
            "date": e.date,
            "event_name": e.event_name,
            "identity_key": e.identity_key,
            "email": product_email_for_event(e),
            "product_category": str(category or "Unknown"),
            "product_title": str(category or "Unknown"),
            "product_brand": str(nested_get(obj, "manufacturer.name") or ""),
            "manufacturer_name": str(nested_get(obj, "manufacturer.name") or ""),
            "model_number": str(nested_get(obj, "manufacturer.model_number") or ""),
            "product_condition": "",
            "quantity": 1,
            "product_unit_price": parse_money(price),
            "product_total_price": parse_money(price),
            "plan_name": "",
            "plan_price": 0.0,
            "plan_term_months": "",
            "eligible": True,
            "retailer_name": retailer["retailer_name"],
            "utm_source": attr["utm_source"],
            "utm_medium": attr["utm_medium"],
            "utm_campaign": attr["utm_campaign"],
            "gwp": e.gwp,
        })
    return rows


def _metric_counts_for_group(events: list[NormEvent]) -> dict[str, Any]:
    out: dict[str, Any] = {
        "Events": len(events),
        "Unique Users": len({e.identity_key for e in events}),
        "Enquiry Attempted": len({e.identity_key for e in events if e.event_name in ENQUIRY_EVENTS}),
        "Sign Up_total": len({e.identity_key for e in events if e.event_name in USER_METRIC_EVENTS["Sign Up_total"]}),
        "First Quote_Success": len({e.identity_key for e in events if e.event_name in QUOTE_SUCCESS_EVENTS}),
        "Offer_Selected": len({e.identity_key for e in events if e.event_name in OFFER_SELECTED_EVENTS}),
        "Invoice Upload_Success": len({e.identity_key for e in events if e.event_name in INVOICE_SUCCESS_EVENTS}),
        "Add to Cart_Success": len({e.identity_key for e in events if e.event_name in ADD_TO_CART_EVENTS}),
        "Payment Attempted": len({e.identity_key for e in events if e.event_name in PAYMENT_ATTEMPTED_EVENTS}),
        "Payment Success": len({e.identity_key for e in events if e.event_name in PAYMENT_SUCCESS_EVENTS}),
        "Payment Failed": len({e.identity_key for e in events if e.event_name in PAYMENT_FAILED_EVENTS}),
        "Gross GWP $": round(sum(e.gwp for e in events if e.event_name in PAYMENT_SUCCESS_EVENTS), 2),
        "Products Quoted": sum(e.product_count for e in events if e.event_name in QUOTE_SUCCESS_EVENTS),
        "Products Added To Cart": sum(e.product_count for e in events if e.event_name in ADD_TO_CART_EVENTS),
        "Products Purchased": sum(e.product_count for e in events if e.event_name in PAYMENT_SUCCESS_EVENTS),
    }
    out["Enquiry to Payment Success %"] = round(out["Payment Success"] / out["Enquiry Attempted"], 4) if out["Enquiry Attempted"] else 0.0
    out["Payment Attempt to Success %"] = round(out["Payment Success"] / out["Payment Attempted"], 4) if out["Payment Attempted"] else 0.0
    return out




def _ad_click_ids(obj: dict) -> dict[str, str]:
    """Return common paid-social/search click ids from revised attribution payloads and fallbacks."""
    a = attribution_values(obj)
    return {
        "fbclid": str(a.get("fbclid") or nested_get(obj, "fbclid", "raw.fbclid") or "").strip(),
        "gclid": str(a.get("gclid") or nested_get(obj, "gclid", "raw.gclid") or "").strip(),
        "gbraid": str(a.get("gbraid") or nested_get(obj, "gbraid", "raw.gbraid") or "").strip(),
        "wbraid": str(a.get("wbraid") or nested_get(obj, "wbraid", "raw.wbraid") or "").strip(),
    }


def _is_utm_traffic(a: dict[str, str], ids: dict[str, str]) -> bool:
    """Return True for explicit campaign/ad-click traffic, excluding plain direct/referral rows."""
    campaign = str(a.get("utm_campaign") or "").strip()
    explicit_campaign = campaign not in ("", "(none)")
    has_click_id = any(v for v in ids.values())
    return explicit_campaign or has_click_id


def _attribution_group_key(a: dict[str, str], level: str) -> tuple[str, ...]:
    source = str(a.get("utm_source") or "unknown").strip() or "unknown"
    campaign = str(a.get("utm_campaign") or "(none)").strip() or "(none)"
    ad_set = str(a.get("ad_set") or "(not set)").strip() or "(not set)"
    ad = str(a.get("ad") or "(not set)").strip() or "(not set)"
    if level == "campaign":
        return (source, campaign)
    if level == "ad_set":
        return (source, campaign, ad_set)
    if level == "ad":
        return (source, campaign, ad_set, ad)
    raise ValueError(f"Unknown attribution level: {level}")


def _attribution_columns(level: str) -> list[str]:
    if level == "campaign":
        return ["Source", "Campaign", "Events", "Sessions", "Leads captured", "Unique fbclicks"]
    if level == "ad_set":
        return ["Source", "Campaign", "Ad set", "Events", "Sessions", "Leads captured", "Unique fbclicks"]
    if level == "ad":
        return ["Source", "Campaign", "Ad set", "Ad", "Events", "Sessions", "Leads captured", "Unique fbclicks"]
    raise ValueError(f"Unknown attribution level: {level}")


def build_attribution_level_stats(clean_events: list[NormEvent], level: str = "campaign") -> pd.DataFrame:
    """Build attribution summary at campaign, ad set, or ad level."""
    groups: dict[tuple[str, ...], list[tuple[NormEvent, dict[str, str], dict[str, str]]]] = defaultdict(list)
    for e in clean_events:
        obj = _safe_json(e)
        a = attribution_values(obj)
        ids = _ad_click_ids(obj)
        if not _is_utm_traffic(a, ids):
            continue
        groups[_attribution_group_key(a, level)].append((e, a, ids))

    rows = []
    for key, items in groups.items():
        events = [x[0] for x in items]
        row = dict(zip(_attribution_columns(level)[:-4], key))
        row.update({
            "Events": len(events),
            "Sessions": len({e.session_id for e in events if e.session_id}),
            "Leads captured": len({e.identity_key for e in events if e.event_name in USER_METRIC_EVENTS["Sign Up_total"]}),
            "Unique fbclicks": len({ids.get("fbclid", "") for _, _, ids in items if ids.get("fbclid", "")}),
        })
        rows.append(row)

    columns = _attribution_columns(level)
    df = pd.DataFrame(rows, columns=columns)
    if not df.empty:
        df = df.sort_values(["Events", "Sessions", "Leads captured"], ascending=[False, False, False]).reset_index(drop=True)
    return df


def build_attribution_campaign_stats(clean_events: list[NormEvent]) -> pd.DataFrame:
    return build_attribution_level_stats(clean_events, level="campaign")


def build_attribution_adset_stats(clean_events: list[NormEvent]) -> pd.DataFrame:
    return build_attribution_level_stats(clean_events, level="ad_set")


def build_attribution_ad_stats(clean_events: list[NormEvent]) -> pd.DataFrame:
    return build_attribution_level_stats(clean_events, level="ad")


def build_utm_event_breakdown(clean_events: list[NormEvent], level: str = "campaign") -> pd.DataFrame:
    """
    Build attribution breakdown for UTM/ad-click traffic at the requested level,
    using the same business metric labels as Daily Metrics.

    Scope note:
      - Daily Metrics = all Clean External events
      - This table = only Clean External events with UTM/ad-click attribution
    """
    groups: dict[tuple[str, ...], list[NormEvent]] = defaultdict(list)
    for e in clean_events:
        obj = _safe_json(e)
        a = attribution_values(obj)
        ids = _ad_click_ids(obj)
        if not _is_utm_traffic(a, ids):
            continue
        groups[_attribution_group_key(a, level)].append(e)

    rows = []
    key_cols = _attribution_columns(level)[:-4]
    metric_cols = [
        "Enquiry Attempted",
        "Sign Up_total",
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
    ]

    for key, evs in groups.items():
        row = dict(zip(key_cols, key))
        if level == "campaign":
            row["Campaign Label"] = f"{row['Source']} / {row['Campaign']}"
            label_col = "Campaign Label"
        elif level == "ad_set":
            row["Ad set Label"] = f"{row['Source']} / {row['Campaign']} / {row['Ad set']}"
            label_col = "Ad set Label"
        else:
            row["Ad Label"] = f"{row['Source']} / {row['Campaign']} / {row['Ad set']} / {row['Ad']}"
            label_col = "Ad Label"

        # User-level metrics use unique identities, matching the Daily Metrics convention.
        row["Enquiry Attempted"] = len({e.identity_key for e in evs if e.event_name in ENQUIRY_EVENTS})
        row["Sign Up_total"] = len({e.identity_key for e in evs if e.event_name in USER_METRIC_EVENTS["Sign Up_total"]})
        row["First Quote_Success"] = len({e.identity_key for e in evs if e.event_name in QUOTE_SUCCESS_EVENTS})
        row["Offer_Selected"] = len({e.identity_key for e in evs if e.event_name in OFFER_SELECTED_EVENTS})
        row["Invoice Upload_Success"] = len({e.identity_key for e in evs if e.event_name in INVOICE_SUCCESS_EVENTS})
        row["Invoice Upload_Failure"] = len({e.identity_key for e in evs if e.event_name in INVOICE_FAILURE_EVENTS})
        row["Add details Later"] = 0
        row["Revised Offer"] = len({e.identity_key for e in evs if e.event_name in REVISED_OFFER_EVENTS})
        row["Additional Product"] = len({e.identity_key for e in evs if e.event_name in USER_METRIC_EVENTS["Additional Product"]})
        row["Add to Cart_Success"] = len({e.identity_key for e in evs if e.event_name in ADD_TO_CART_EVENTS})
        row["Payment Attempted"] = len({e.identity_key for e in evs if e.event_name in PAYMENT_ATTEMPTED_EVENTS})
        row["Payment Success"] = len({e.identity_key for e in evs if e.event_name in PAYMENT_SUCCESS_EVENTS})
        row["Payment Failed"] = len({e.identity_key for e in evs if e.event_name in PAYMENT_FAILED_EVENTS})
        row["Email Verified_Success"] = 0
        row["Email Verified_Failed"] = 0
        row["Gross GWP $"] = round(sum(e.gwp for e in evs if e.event_name in PAYMENT_SUCCESS_EVENTS), 2)
        rows.append(row)

    label_col = "Campaign Label" if level == "campaign" else "Ad set Label" if level == "ad_set" else "Ad Label"
    columns = key_cols + [label_col] + metric_cols
    df = pd.DataFrame(rows, columns=columns)
    if not df.empty:
        df["_sort_events"] = df[[c for c in metric_cols if c != "Gross GWP $"]].sum(axis=1)
        df = (
            df.sort_values(["Gross GWP $", "_sort_events", "Enquiry Attempted"], ascending=[False, False, False])
            .drop(columns=["_sort_events"])
            .reset_index(drop=True)
        )
    return df


def build_attribution_stats(clean_events: list[NormEvent]) -> pd.DataFrame:
    """Backward-compatible wrapper."""
    return build_attribution_campaign_stats(clean_events)



def build_sales_stats(clean_events: list[NormEvent]) -> pd.DataFrame:
    """Build a non-PII sales table from Payment Success events.

    Includes email ID as the only customer identifier. Excludes names, phone
    numbers, billing/shipping addresses, and raw JSON.
    """
    rows: list[dict[str, Any]] = []

    for e in clean_events:
        if e.event_name not in PAYMENT_SUCCESS_EVENTS:
            continue

        obj = _safe_json(e)
        attr = attribution_values(obj)
        items = get_line_items(obj)

        if not items:
            raw_items = nested_get(obj, "raw.items", "data.items")
            if isinstance(raw_items, list):
                items = [i for i in raw_items if isinstance(i, dict)]

        # If a payment event has no line_items, still keep one order-level row.
        if not items:
            items = [{}]

        checkout_subtotal = parse_money(
            nested_get(
                obj,
                "event_data.checkout.subtotal_amount",
                "event_data.checkout.total_amount",
                "checkout.subtotal_amount",
                "checkout.total_amount",
                "raw.value",
                "event_data.gwp",
                "gwp",
                "warrantyPrice",
                "warranty_price",
            )
        )

        for item in items:
            product = item.get("product") if isinstance(item.get("product"), dict) else {}
            purchase = item.get("purchase") if isinstance(item.get("purchase"), dict) else {}
            protection = item.get("protection") if isinstance(item.get("protection"), dict) else {}
            category = product.get("category") if isinstance(product.get("category"), dict) else {}

            row = {
                "Date": e.date,
                "Email": product_email_for_event(e),
                "Event": e.event_name,
                "Order / Payment ID": str(
                    nested_get(
                        obj,
                        "event_data.checkout.order_id",
                        "event_data.checkout.payment_id",
                        "checkout.order_id",
                        "checkout.payment_id",
                        "payment_id",
                        "order_id",
                        "event_data.order_id",
                        "event_data.payment_id",
                    )
                    or e.event_id
                    or ""
                ),
                "Product Category": str(category.get("name") or product.get("category") or item.get("item_category") or item.get("category") or "Unknown"),
                "Product Title": str(product.get("title") or item.get("item_name") or item.get("name") or product.get("description") or "Unknown"),
                "Product Brand": str(product.get("brand") or nested_get(obj, "manufacturer.name") or ""),
                "Manufacturer": str(nested_get(obj, "manufacturer.name") or product.get("brand") or ""),
                "Model Number": str(nested_get(obj, "manufacturer.model_number") or product.get("sku") or ""),
                "Product Condition": str(product.get("condition") or ""),
                "Quantity": item_quantity(item),
                "Product Unit Price": parse_money(purchase.get("unit_price") or item.get("unit_price") or item.get("price")),
                "Product Total Price": parse_money(purchase.get("total_price") or item.get("total_price")),
                "Warranty / Plan Name": str(protection.get("plan_name") or protection.get("name") or item.get("item_variant") or ""),
                "Warranty Price": parse_money(
                    protection.get("plan_price")
                    or protection.get("price")
                    or item.get("plan_price")
                    or item.get("warrantyPrice")
                    or item.get("warranty_price")
                    or item.get("price")
                    or e.gwp
                    or checkout_subtotal
                ),
                "Warranty Term Months": protection.get("term_months") or protection.get("term") or "",
                "Warranty Provider": str(protection.get("provider") or protection.get("underwriter") or protection.get("administrator") or ""),
                "Manufacturer Warranty": str(nested_get(obj, "manufacturer.warranty") or product.get("warranty") or ""),
                "Eligible": bool(is_eligible_item(item)) if item else "",
                "Gross GWP $": round(float(e.gwp or checkout_subtotal or 0.0), 2),
                "UTM Source": attr["utm_source"],
                "UTM Medium": attr["utm_medium"],
                "UTM Campaign": attr["utm_campaign"],
            }
            rows.append(row)

    if not rows:
        return pd.DataFrame(
            columns=[
                "Date", "Email", "Event", "Order / Payment ID", "Product Category", "Product Title",
                "Product Brand", "Manufacturer", "Model Number", "Product Condition", "Quantity",
                "Product Unit Price", "Product Total Price", "Warranty / Plan Name", "Warranty Price",
                "Warranty Term Months", "Warranty Provider", "Manufacturer Warranty", "Eligible",
                "Gross GWP $", "UTM Source", "UTM Medium", "UTM Campaign",
            ]
        )

    df = pd.DataFrame(rows)
    df = df.sort_values(["Date", "Email", "Product Category", "Product Title"], ascending=[False, True, True, True]).reset_index(drop=True)
    return df




# -----------------------------
# High Intent dropoff reporting
# -----------------------------

def pii_value(obj: dict, *paths: str) -> str:
    for path in paths:
        val = nested_get(obj, path)
        if val not in (None, ""):
            return str(val).strip()
    return ""


def phone_for_event_obj(obj: dict) -> str:
    for key in ("phone", "phone_number", "phonenumber", "mobile", "mobile_number", "tel"):
        vals = collect_values_by_key(obj, {key})
        for val in vals:
            s = str(val).strip()
            if s:
                return s
    return ""


def address_for_event_obj(obj: dict) -> str:
    parts = []
    for path in (
        "actor.address.line1",
        "actor.address.line2",
        "customer.address.line1",
        "customer.address.line2",
        "billing_address.line1",
        "billing_address.line2",
        "shipping_address.line1",
        "shipping_address.line2",
        "event_data.invoice.retailer.address.line1",
        "event_data.invoice.retailer.address.line2",
    ):
        val = nested_get(obj, path)
        if val not in (None, ""):
            parts.append(str(val).strip())
    city = pii_value(obj, "actor.address.city", "customer.address.city", "billing_address.city", "shipping_address.city")
    state = pii_value(obj, "actor.address.state", "customer.address.state", "billing_address.state", "shipping_address.state")
    postal = pii_value(obj, "actor.address.postal_code", "customer.address.postal_code", "billing_address.postal_code", "shipping_address.postal_code", "zip", "zipcode")
    country = pii_value(obj, "actor.address.country", "customer.address.country", "billing_address.country", "shipping_address.country")
    for val in (city, state, postal, country):
        if val:
            parts.append(val)
    return ", ".join(dict.fromkeys(parts))


def name_for_event_obj(obj: dict) -> str:
    direct = pii_value(obj, "actor.name", "customer.name", "name", "event_data.customer.name")
    if direct:
        return direct
    first = pii_value(obj, "actor.first_name", "customer.first_name", "first_name", "firstname", "event_data.customer.first_name")
    last = pii_value(obj, "actor.last_name", "customer.last_name", "last_name", "lastname", "event_data.customer.last_name")
    return " ".join([p for p in (first, last) if p]).strip()


def high_intent_email_for_event(e: NormEvent) -> str:
    obj = _safe_json(e)
    return pii_value(obj, "actor.email", "customer.email", "event_data.customer.email", "data.email", "email") or primary_email(e)


HIGH_INTENT_STAGE_ORDER = [
    ("Sign Up_total", USER_METRIC_EVENTS["Sign Up_total"]),
    ("First Quote_Success", QUOTE_SUCCESS_EVENTS),
    ("Offer_Selected", OFFER_SELECTED_EVENTS),
    ("Invoice Upload_Success", INVOICE_SUCCESS_EVENTS),
    ("Invoice Upload_Failure", INVOICE_FAILURE_EVENTS),
    ("Revised Offer", REVISED_OFFER_EVENTS),
    ("Additional Product", USER_METRIC_EVENTS["Additional Product"]),
    ("Add to cart", ADD_TO_CART_EVENTS),
    ("Initiate Checkout", {"initiate_checkout", "pay_now_clicked"}),
    ("Payment Failed", PAYMENT_FAILED_EVENTS),
    ("Payment Success", PAYMENT_SUCCESS_EVENTS),
]


def latest_stage_for_events(events: list[NormEvent]) -> str:
    latest = ""
    latest_rank = -1
    for e in events:
        for rank, (stage, event_set) in enumerate(HIGH_INTENT_STAGE_ORDER):
            if e.event_name in event_set and rank >= latest_rank:
                latest = stage
                latest_rank = rank
    return latest


def product_summary_for_session(events: list[NormEvent]) -> dict[str, Any]:
    categories: list[str] = []
    titles: list[str] = []
    brands: list[str] = []
    manufacturers: list[str] = []
    models: list[str] = []
    warranty_names: list[str] = []
    warranty_prices: list[float] = []
    retailers: list[str] = []
    gwp_total = 0.0

    for e in events:
        obj = _safe_json(e)
        retailer = retailer_values(obj).get("retailer_name") or ""
        if retailer and retailer != "Unknown":
            retailers.append(retailer)
        gwp_total += float(e.gwp or 0.0)

        items = get_line_items(obj)
        if not items:
            raw_items = nested_get(obj, "raw.items", "data.items")
            if isinstance(raw_items, list):
                items = [i for i in raw_items if isinstance(i, dict)]

        for item in items or []:
            product = item.get("product") if isinstance(item.get("product"), dict) else {}
            protection = item.get("protection") if isinstance(item.get("protection"), dict) else {}
            category = product.get("category") if isinstance(product.get("category"), dict) else {}

            cat = str(category.get("name") or product.get("category") or item.get("item_category") or item.get("category") or "").strip()
            title = str(product.get("title") or item.get("item_name") or item.get("name") or product.get("description") or "").strip()
            brand = str(product.get("brand") or "").strip()
            manufacturer = str(nested_get(obj, "manufacturer.name") or product.get("brand") or "").strip()
            model = str(nested_get(obj, "manufacturer.model_number") or product.get("sku") or "").strip()
            warranty_name = str(protection.get("plan_name") or protection.get("name") or item.get("item_variant") or "").strip()
            warranty_price = parse_money(
                protection.get("plan_price")
                or protection.get("price")
                or item.get("plan_price")
                or item.get("warrantyPrice")
                or item.get("warranty_price")
                or item.get("price")
            )

            if cat:
                categories.append(cat)
            if title:
                titles.append(title)
            if brand:
                brands.append(brand)
            if manufacturer:
                manufacturers.append(manufacturer)
            if model:
                models.append(model)
            if warranty_name:
                warranty_names.append(warranty_name)
            if warranty_price > 0:
                warranty_prices.append(warranty_price)

    def joined(vals: list[Any]) -> str:
        return ", ".join(str(v) for v in dict.fromkeys(vals) if str(v).strip())

    return {
        "Product Category": joined(categories),
        "Product Title": joined(titles),
        "Product Brand": joined(brands),
        "Manufacturer": joined(manufacturers),
        "Model Number": joined(models),
        "Warranty / Plan Name": joined(warranty_names),
        "Warranty Price": round(float(sum(warranty_prices)), 2),
        "Retailer": joined(retailers),
        "Gross GWP $": round(float(gwp_total), 2),
    }



def compact_json_for_table(value: Any) -> str:
    if value in (None, ""):
        return ""
    try:
        return json.dumps(value, ensure_ascii=False, separators=(",", ":"), default=str)
    except Exception:
        return str(value)


def flatten_for_table(value: Any, prefix: str = "") -> dict[str, str]:
    """Flatten dict/list data into readable table columns."""
    out: dict[str, str] = {}

    def walk(v: Any, key: str) -> None:
        if isinstance(v, dict):
            if not v:
                out[key] = ""
            for k, child in v.items():
                child_key = f"{key}.{k}" if key else str(k)
                walk(child, child_key)
        elif isinstance(v, list):
            out[key] = compact_json_for_table(v)
        else:
            out[key] = "" if v is None else str(v)

    walk(value, prefix)
    return out


def form_object_for_event_obj(obj: dict) -> Any:
    """Return the form/forms object from the event, with broad schema support."""
    candidates = [
        nested_get(obj, "event_data.forms"),
        nested_get(obj, "event_data.form"),
        nested_get(obj, "forms"),
        nested_get(obj, "form"),
        nested_get(obj, "event_data.form.fields"),
        nested_get(obj, "raw.form"),
        nested_get(obj, "raw.forms"),
        nested_get(obj, "data.form"),
        nested_get(obj, "data.forms"),
    ]
    for candidate in candidates:
        if candidate not in (None, "", {}, []):
            return candidate
    return {}


def flatten_form_columns(obj: dict, column_prefix: str = "Form") -> dict[str, str]:
    form_obj = form_object_for_event_obj(obj)
    if not form_obj:
        return {}
    flat = flatten_for_table(form_obj)
    return {f"{column_prefix}: {k}": v for k, v in flat.items()}


def form_value(obj: dict, *names: str) -> str:
    """Find a value inside form/forms object by key name, case-insensitive."""
    form_obj = form_object_for_event_obj(obj)
    if not form_obj:
        return ""
    wanted = {n.lower().replace("_", "").replace(" ", "") for n in names}
    found = ""

    def walk(v: Any) -> None:
        nonlocal found
        if found:
            return
        if isinstance(v, dict):
            for k, child in v.items():
                nk = str(k).lower().replace("_", "").replace(" ", "")
                if nk in wanted and child not in (None, ""):
                    found = str(child).strip()
                    return
                walk(child)
        elif isinstance(v, list):
            for child in v:
                walk(child)

    walk(form_obj)
    return found


def first_event_with_any(events: list[NormEvent], event_sets: list[set[str]]) -> NormEvent | None:
    wanted: set[str] = set()
    for s in event_sets:
        wanted |= set(s)
    for e in sorted(events, key=lambda x: x.event_time or ""):
        if e.event_name in wanted:
            return e
    return None


def line_items_full_data_for_session(events: list[NormEvent]) -> list[Any]:
    all_items: list[Any] = []
    for e in sorted(events, key=lambda x: x.event_time or ""):
        obj = _safe_json(e)
        items = get_line_items(obj)
        if not items:
            raw_items = nested_get(obj, "event_data.line_items", "line_items", "raw.items", "data.items")
            if isinstance(raw_items, list):
                items = [i for i in raw_items if isinstance(i, dict)]
        for item in items or []:
            all_items.append(item)
    return all_items


def build_high_intent_dropoffs(clean_events: list[NormEvent]) -> tuple[pd.DataFrame, pd.DataFrame, int]:
    """Build session-based high-intent dropoff tables.

    Table 1: sessions that reached Sign Up_total but did not upload an invoice, add to cart,
    initiate checkout, or purchase. It includes all form/forms fields from the signup event.
    Table 2: sessions that uploaded an invoice or reached a later stage but did not purchase.
    It includes name/email from the form object and full line_items object data.

    Uses session_id as the primary stitching key, with identity_key fallback if session_id is missing.
    """
    session_events: dict[str, list[NormEvent]] = defaultdict(list)
    for e in clean_events:
        sid = str(e.session_id or "").strip()
        key = f"session:{sid}" if sid else f"identity:{e.identity_key}"
        session_events[key].append(e)

    signup_rows: list[dict[str, Any]] = []
    invoice_rows: list[dict[str, Any]] = []
    unique_user_ids: set[str] = set()

    for session_key, events in session_events.items():
        event_names = {e.event_name for e in events}
        has_signup = bool(event_names & USER_METRIC_EVENTS["Sign Up_total"])
        has_invoice_success = bool(event_names & (INVOICE_SUCCESS_EVENTS | REVISED_OFFER_EVENTS))
        has_invoice_failure = bool(event_names & INVOICE_FAILURE_EVENTS)
        has_invoice_or_later = bool(
            event_names
            & (
                INVOICE_SUCCESS_EVENTS
                | INVOICE_FAILURE_EVENTS
                | REVISED_OFFER_EVENTS
                | USER_METRIC_EVENTS["Additional Product"]
                | ADD_TO_CART_EVENTS
                | {"initiate_checkout", "pay_now_clicked"}
                | PAYMENT_FAILED_EVENTS
            )
        )
        has_downstream_after_signup = bool(
            event_names
            & (
                INVOICE_SUCCESS_EVENTS
                | INVOICE_FAILURE_EVENTS
                | ADD_TO_CART_EVENTS
                | {"initiate_checkout", "pay_now_clicked"}
                | PAYMENT_SUCCESS_EVENTS
            )
        )
        has_payment_success = bool(event_names & PAYMENT_SUCCESS_EVENTS)

        if has_payment_success:
            continue

        events_sorted = sorted(events, key=lambda x: x.event_time or "")
        first = events_sorted[0]
        last = events_sorted[-1]
        first_obj = _safe_json(first)

        signup_event = first_event_with_any(events, [USER_METRIC_EVENTS["Sign Up_total"]])
        invoice_event = first_event_with_any(events, [INVOICE_SUCCESS_EVENTS, INVOICE_FAILURE_EVENTS]) or first_event_with_any(
            events,
            [REVISED_OFFER_EVENTS, USER_METRIC_EVENTS["Additional Product"], ADD_TO_CART_EVENTS, {"initiate_checkout", "pay_now_clicked"}, PAYMENT_FAILED_EVENTS],
        )

        identity_events = sorted(events, key=lambda x: (0 if high_intent_email_for_event(x) else 1, x.event_time or ""))
        pii_event = identity_events[0] if identity_events else last
        pii_obj = _safe_json(pii_event)

        email = high_intent_email_for_event(pii_event)
        name = name_for_event_obj(pii_obj)
        phone = phone_for_event_obj(pii_obj)
        address = address_for_event_obj(pii_obj)

        product_bits = product_summary_for_session(events)
        attr = attribution_values(first_obj)

        base_row = {
            "Session ID": str(first.session_id or session_key.replace("session:", "")),
            "Email": email,
            "Name": name,
            "Phone": phone,
            "Address": address,
            "Lead ID": str(first.lead_id or ""),
            "First Event Date": first.date,
            "Last Event Date": last.date,
            "Last Stage": latest_stage_for_events(events),
            "Event Count": int(len(events)),
            "Invoice Uploaded": "Yes" if has_invoice_success else "No",
            "Invoice Upload Failed": "Yes" if has_invoice_failure else "No",
            "UTM Source": attr["utm_source"],
            "UTM Medium": attr["utm_medium"],
            "UTM Campaign": attr["utm_campaign"],
        }
        base_row.update(product_bits)

        if has_signup and not has_downstream_after_signup:
            signup_obj = _safe_json(signup_event) if signup_event else {}

            form_name_signup = form_value(signup_obj, "name", "full_name", "fullname", "first_name", "firstname", "last_name", "lastname")
            form_email_signup = form_value(signup_obj, "email", "email_address", "emailaddress")
            form_product_category = form_value(
                signup_obj,
                "product_category",
                "productcategory",
                "category",
                "category_name",
                "categoryname",
                "sb_category_name",
                "sbcategoryname",
            )
            form_product_price = form_value(
                signup_obj,
                "product_price",
                "productprice",
                "price",
                "value",
                "product_value",
                "productvalue",
            )

            row = {
                "First Event Date": first.date,
                "Last Event Date": last.date,
                "Name": form_name_signup,
                "Email": form_email_signup,
                "Last Stage": latest_stage_for_events(events),
                "Product Category": form_product_category or product_bits.get("Product Category", ""),
                "Product Price": form_product_price,
            }
            signup_rows.append(row)
            unique_user_ids.add(form_email_signup or email or first.identity_key or session_key)

        elif has_invoice_or_later:
            invoice_obj = _safe_json(invoice_event) if invoice_event else {}

            form_name = form_value(invoice_obj, "name", "full_name", "fullname", "first_name", "firstname", "last_name", "lastname")
            form_email = form_value(invoice_obj, "email", "email_address", "emailaddress")

            line_items_data = line_items_full_data_for_session(events)

            row = {
                "First Event Date": first.date,
                "Last Event Date": last.date,
                "Name": form_name,
                "Email": form_email,
                "Last Stage": latest_stage_for_events(events),
                "Invoice Uploaded": "Yes" if has_invoice_success else "No",
                "Invoice Upload Failed": "Yes" if has_invoice_failure else "No",
                "Product Category": product_bits.get("Product Category", ""),
                "Product Brand": product_bits.get("Product Brand", ""),
                "Manufacturer": product_bits.get("Manufacturer", ""),
                "Model Number": product_bits.get("Model Number", ""),
                "Warranty / Plan Name": product_bits.get("Warranty / Plan Name", ""),
                "Warranty Price": product_bits.get("Warranty Price", ""),
                "Retailer": product_bits.get("Retailer", ""),
                "Line Items Data": compact_json_for_table(line_items_data),
                "UTM Source": attr["utm_source"],
                "UTM Medium": attr["utm_medium"],
                "UTM Campaign": attr["utm_campaign"],
            }

            invoice_rows.append(row)
            unique_user_ids.add((form_email or email) or first.identity_key or session_key)

    base_columns = [
        "Session ID",
        "Email",
        "Name",
        "Phone",
        "Address",
        "Lead ID",
        "First Event Date",
        "Last Event Date",
        "Last Stage",
        "Event Count",
        "Invoice Uploaded",
        "Invoice Upload Failed",
        "Product Category",
        "Product Title",
        "Product Brand",
        "Manufacturer",
        "Model Number",
        "Warranty / Plan Name",
        "Warranty Price",
        "Retailer",
        "Gross GWP $",
        "UTM Source",
        "UTM Medium",
        "UTM Campaign",
    ]

    signup_columns = [
        "First Event Date",
        "Last Event Date",
        "Name",
        "Email",
        "Last Stage",
        "Product Category",
        "Product Price",
    ]
    invoice_columns = [
        "First Event Date",
        "Last Event Date",
        "Name",
        "Email",
        "Last Stage",
        "Invoice Uploaded",
        "Invoice Upload Failed",
        "Product Category",
        "Product Brand",
        "Manufacturer",
        "Model Number",
        "Warranty / Plan Name",
        "Warranty Price",
        "Retailer",
        "Line Items Data",
        "UTM Source",
        "UTM Medium",
        "UTM Campaign",
    ]

    signup_df = pd.DataFrame(signup_rows)
    invoice_df = pd.DataFrame(invoice_rows)

    signup_df = signup_df.reindex(columns=signup_columns)
    invoice_df = invoice_df.reindex(columns=invoice_columns)

    if not signup_df.empty:
        signup_df.sort_values(["Last Event Date", "Email"], ascending=[False, True], inplace=True)
        signup_df.reset_index(drop=True, inplace=True)

    if not invoice_df.empty:
        invoice_df.sort_values(["Last Event Date", "Email"], ascending=[False, True], inplace=True)
        invoice_df.reset_index(drop=True, inplace=True)

    unique_dropoff_users = int(len({u for u in unique_user_ids if str(u).strip()}))

    return signup_df, invoice_df, unique_dropoff_users


def build_product_stats(clean_events: list[NormEvent]) -> pd.DataFrame:
    rows = []
    for e in clean_events:
        rows.extend(product_rows_for_event(e))
    if not rows:
        return pd.DataFrame()
    detail = pd.DataFrame(rows)
    group_cols = ["product_category", "product_title", "product_brand", "manufacturer_name", "product_condition"]
    out_rows = []
    for key, g in detail.groupby(group_cols, dropna=False):
        evs_by_key = []
        # Metric counts here are product-level counts, not unique user counts.
        row = dict(zip(["Product Category", "Product Title", "Product Brand", "Manufacturer", "Condition"], key))
        unique_emails = sorted({str(x).strip() for x in g["email"].dropna().tolist() if str(x).strip()})
        row["Email"] = ", ".join(unique_emails)
        row["Enquiry Attempted"] = int(g.loc[g["event_name"].isin(ENQUIRY_EVENTS), "quantity"].sum())
        row["Sign Up_total"] = int(g.loc[g["event_name"].isin({"sign_up_total"}), "quantity"].sum())
        row["Add to cart"] = int(g.loc[g["event_name"].isin(ADD_TO_CART_EVENTS), "quantity"].sum())
        row["Invoice Upload_Success"] = int(g.loc[g["event_name"].isin(INVOICE_SUCCESS_EVENTS) & (g["eligible"] == True), "quantity"].sum())
        row["Initiate Checkout"] = int(g.loc[g["event_name"].isin({"initiate_checkout"}), "quantity"].sum())
        row["Payment Success"] = int(g.loc[g["event_name"].isin(PAYMENT_SUCCESS_EVENTS), "quantity"].sum())
        out_rows.append(row)
    df = pd.DataFrame(out_rows)
    if not df.empty:
        df = df.drop(columns=["Product Events"], errors="ignore")
        df = df.sort_values(["Payment Success", "Add to cart", "Enquiry Attempted"], ascending=[False, False, False])
        preferred_cols = [
            "Product Category",
            "Email",
            "Product Title",
            "Product Brand",
            "Manufacturer",
            "Condition",
            "Enquiry Attempted",
            "Sign Up_total",
            "Add to cart",
            "Invoice Upload_Success",
            "Initiate Checkout",
            "Payment Success",
        ]
        df = df[[c for c in preferred_cols if c in df.columns] + [c for c in df.columns if c not in preferred_cols]]
    return df


def build_retailer_stats(clean_events: list[NormEvent]) -> pd.DataFrame:
    groups: dict[tuple, list[NormEvent]] = defaultdict(list)
    retailer_lookup: dict[tuple, dict[str, str]] = {}
    for e in clean_events:
        obj = _safe_json(e)
        rv = retailer_values(obj)
        # Only include events where a retailer is known from invoice/revised-offer payloads.
        if not rv["retailer_name"] or rv["retailer_name"] == "Unknown":
            continue
        key = (rv["retailer_name"], rv["retailer_email"], rv["retailer_city"], rv["retailer_country"])
        groups[key].append(e)
        retailer_lookup[key] = rv
    rows = []
    for key, evs in groups.items():
        rv = retailer_lookup[key]
        row = {
            "Retailer Name": rv["retailer_name"],
            "Retailer Email": rv["retailer_email"],
            "Retailer City": rv["retailer_city"],
            "Retailer State": rv["retailer_state"],
            "Retailer Country": rv["retailer_country"],
        }
        row.update(_metric_counts_for_group(evs))
        row["Invoice Success_Product Count"] = sum(e.eligible_product_count for e in evs if e.event_name in INVOICE_SUCCESS_EVENTS)
        row["Revised Offer Users"] = len({e.identity_key for e in evs if e.event_name in REVISED_OFFER_EVENTS})
        rows.append(row)
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values(["Gross GWP $", "Invoice Upload_Success", "Events"], ascending=[False, False, False])
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
    attribution: pd.DataFrame,
    utm_breakdown: pd.DataFrame,
    product: pd.DataFrame,
    retailer: pd.DataFrame,
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
        attribution.to_excel(writer, index=False, sheet_name="Attribution by Campaign")
        utm_breakdown.to_excel(writer, index=False, sheet_name="UTM Event Breakdown")
        product.to_excel(writer, index=False, sheet_name="Product Stats")
        retailer.to_excel(writer, index=False, sheet_name="Retailer Stats")
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
            for cell in ws[1]:
                if cell.value and ("GWP" in str(cell.value) or "Price" in str(cell.value)):
                    col = cell.column_letter
                    for c in ws[f"{col}2":f"{col}{ws.max_row}"]:
                        c[0].number_format = '$#,##0.00'
                if cell.value and "%" in str(cell.value):
                    col = cell.column_letter
                    for c in ws[f"{col}2":f"{col}{ws.max_row}"]:
                        c[0].number_format = '0.00%'
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
    attribution = build_attribution_campaign_stats(clean_events)
    utm_breakdown = build_utm_event_breakdown(clean_events)
    product = build_product_stats(clean_events)
    retailer = build_retailer_stats(clean_events)

    base_name = args.base_name or re.sub(r"[^A-Za-z0-9_\-]+", "_", input_path.stem).strip("_") + "_clean_external"
    xlsx, metrics_csv, audit_csv = write_outputs(
        Path(args.out_dir),
        base_name,
        daily,
        totals,
        attribution,
        utm_breakdown,
        product,
        retailer,
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
