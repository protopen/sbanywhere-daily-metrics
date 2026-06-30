# V2 Eastern Time dashboard update

Built on top of V2 Paid Campaign Metrics source.attribution-only build.

Change:
- Default dashboard timezone changed to America/New_York.
- V2 date grouping now uses Eastern Time dates.
- V2 detail/order/log views include a `Date/Time (ET)` column.
- Raw UTC timestamps are converted to America/New_York for display.
- Date filters continue to fetch a buffered UTC window from Supabase, then the app applies the selected dashboard timezone date filter.

Notes:
- America/New_York handles EST/EDT automatically. June events display as EDT because daylight savings is active.
- Paid Campaign Metrics still uses only source.attribution.

Preserved:
- Paid Campaign Metrics uses only source.attribution.
- No traffic/source.utm/URL/old UTM fallback in Paid Campaign Metrics.
- All missing V2 helpers are defined.
- Soumya Ramtri / ramtrisoumya11@gmail.com exclusion.
- Product Sub Category removed from V2 Detail View.
- All V2 event-count metrics use unique session/identity counts.
- Traffic (Total) still counts only unique page_view identities.
- Sign Up_ Total counts event_name = sign_up_total.
- Enquiry Success remains removed from Product Category Metrics.
- Invoice Status filter is unselected by default.
- page_view remains separate from conversion events.
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
