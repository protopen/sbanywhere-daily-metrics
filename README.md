# V2 traffic/page_view + last-click update

Built on top of V2 KPI alignment build.

Changes:
- Invoice Status filter is now unselected by default across V2 tabs.
  - If no invoice status is selected, the dashboard does not filter by invoice status.
- Source Metrics now uses last-click attribution for source bucketing.
  - It prefers `traffic.attribution.last_touch` paths.
  - It falls back to `traffic.marketing`, then older payload paths.
- Any table column named `Traffic (Total)` now counts unique traffic identities from `page_view` events only.
  - Source Metrics traffic uses page_view by source.
  - Paid Campaign Metrics traffic uses page_view by campaign/adset/ad.
- `page_view` remains separate from conversion events and is not counted as enquiry_attempted.

Preserved:
- KPI alignment in V2 Daily Metrics.
- V2 starts from 26 Jun 2026.
- V1 stops before 26 Jun 2026.
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
