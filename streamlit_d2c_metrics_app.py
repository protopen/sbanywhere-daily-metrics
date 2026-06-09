"""
SureBright D2C Clean External Metrics - Streamlit App

Run locally:
  pip install -r requirements_streamlit_d2c.txt
  streamlit run streamlit_d2c_metrics_app.py

Keep this file in the same folder as d2c_clean_external_metrics_report.py.
"""

from __future__ import annotations

import io
import re
import tempfile
from datetime import date, datetime
from pathlib import Path
from typing import Optional

import pandas as pd
import streamlit as st

try:
    from zoneinfo import ZoneInfo
except Exception:  # pragma: no cover
    ZoneInfo = None

from d2c_clean_external_metrics_report import (
    build_daily_metrics,
    build_totals,
    normalize_all_events,
)

APP_TITLE = "SureBright D2C Clean External Metrics"
DEFAULT_START_DATE = date(2026, 5, 21)
DEFAULT_TIMEZONE = "Asia/Kolkata"


def safe_stem(filename: str) -> str:
    stem = Path(filename).stem or "d2c_source"
    return re.sub(r"[^A-Za-z0-9_\-]+", "_", stem).strip("_") or "d2c_source"


def date_to_str(value: Optional[date]) -> Optional[str]:
    return value.isoformat() if value else None


def dataframe_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


def build_excel_bytes(
    daily: pd.DataFrame,
    totals: pd.DataFrame,
    audit_df: pd.DataFrame,
    metadata: pd.DataFrame,
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        daily.to_excel(writer, index=False, sheet_name="Daily Metrics")
        totals.to_excel(writer, index=False, sheet_name="Totals")
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

            if ws.title == "Daily Metrics":
                for cell in ws[1]:
                    if cell.value == "Gross GWP $":
                        for data_cell in ws[cell.column_letter][1:]:
                            data_cell.number_format = "$#,##0.00"
            if ws.title == "Totals":
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                    if row[0].value == "Gross GWP $":
                        row[1].number_format = "$#,##0.00"

    output.seek(0)
    return output.read()


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

        excel_bytes = build_excel_bytes(daily, totals, audit_df, metadata)

    return {
        "daily": daily,
        "totals": totals,
        "audit": audit_df,
        "clean_audit": clean_audit_df,
        "metadata": metadata,
        "excel_bytes": excel_bytes,
        "parseable_count": len(all_events),
        "clean_count": len(clean_events),
    }


def show_kpis(daily: pd.DataFrame, parseable_count: int, clean_count: int) -> None:
    gross_gwp = float(daily["Gross GWP $"].sum()) if not daily.empty and "Gross GWP $" in daily else 0.0
    payment_success = int(daily["Payment Success"].sum()) if not daily.empty and "Payment Success" in daily else 0
    payment_attempted = int(daily["Payment Attempted"].sum()) if not daily.empty and "Payment Attempted" in daily else 0
    enquiry_attempted = int(daily["Enquiry Attempted"].sum()) if not daily.empty and "Enquiry Attempted" in daily else 0

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Parseable events", f"{parseable_count:,}")
    c2.metric("Clean external events", f"{clean_count:,}")
    c3.metric("Enquiry Attempted", f"{enquiry_attempted:,}")
    c4.metric("Payment Success", f"{payment_success:,}")
    c5.metric("Gross GWP", f"${gross_gwp:,.2f}")

    if payment_attempted:
        st.caption(f"Payment success rate: {payment_success / payment_attempted:.1%} based on clean external payment-attempt users.")


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="📊", layout="wide")
    st.title(APP_TITLE)
    st.write(
        "Upload the D2C raw export and generate Clean External daily metrics, totals, and an event audit. "
        "Supported formats: `.xlsx`, `.csv`, `.txt`, `.jsonl`."
    )

    with st.sidebar:
        st.header("Inputs")
        uploaded_file = st.file_uploader(
            "Upload source file",
            type=["xlsx", "xls", "csv", "txt", "jsonl"],
            help="Use your raw export with Time/Raw columns or a JSON-lines raw event dump.",
        )

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

    if uploaded_file is None:
        st.info("Upload a source file to generate metrics.")
        st.stop()

    file_bytes = uploaded_file.getvalue()
    start_date_str = date_to_str(start_date_value)
    end_date_str = date_to_str(end_date_value)

    try:
        with st.spinner("Processing clean external metrics..."):
            result = process_uploaded_file(
                file_bytes=file_bytes,
                filename=uploaded_file.name,
                timezone_name=timezone_name,
                start_date=start_date_str,
                end_date=end_date_str,
            )
    except Exception as exc:
        st.error("Could not process the file. Check that it is a raw D2C export with JSON payloads.")
        st.exception(exc)
        st.stop()

    daily = result["daily"]
    totals = result["totals"]
    audit = result["audit"]
    clean_audit = result["clean_audit"]

    show_kpis(daily, result["parseable_count"], result["clean_count"])

    tab_daily, tab_totals, tab_audit, tab_downloads = st.tabs(
        ["Daily Metrics", "Totals", "Event Audit", "Downloads"]
    )

    with tab_daily:
        st.subheader("Daily Metrics")
        if daily.empty:
            st.warning("No clean external events matched the selected date range.")
        else:
            st.dataframe(daily, use_container_width=True, hide_index=True)
            chart_cols = ["Enquiry Attempted", "Sign Up_total", "Payment Attempted", "Payment Success"]
            available_chart_cols = [col for col in chart_cols if col in daily.columns]
            if available_chart_cols:
                chart_df = daily.set_index("Date")[available_chart_cols]
                st.line_chart(chart_df)

    with tab_totals:
        st.subheader("Totals")
        st.dataframe(totals, use_container_width=True, hide_index=True)

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
        base = safe_stem(uploaded_file.name)
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

    st.caption("Metric logic matches the Clean External reporting script: unique-user counts for user metrics, product counts for product-level metrics, and GWP from successful payment checkout totals.")


if __name__ == "__main__":
    main()
