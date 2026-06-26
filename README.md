# V2 page_view separation update

Built on top of V2 anomaly logs build.

Changes:
- `page_view` is treated as a separate traffic event.
- `page_view` is documented as coming before `enquiry_attempted` in the journey sequence.
- `page_view` is intentionally excluded from all current V2 dashboard calculations and tables for now.
- `enquiry_attempted` remains a separate conversion/funnel event.
- Existing V2 traffic object parsing support is preserved for future use.

Preserved:
- V2 starts from 26 Jun 2026.
- V1 stops before 26 Jun 2026.
- New `traffic` object parsing support.
- Journey Type anomaly logs.
- Conditional per-tab filters.
- Visible filter names.
- Clean tab names.
- Logo lowered spacing.
- All V2 tabs 1 through 7 plus Logs.

Replace these files in GitHub:
- streamlit_d2c_metrics_app_supabase.py
- d2c_clean_external_metrics_report.py
- requirements.txt

After pushing, reboot the Streamlit app and clear cache if needed.
