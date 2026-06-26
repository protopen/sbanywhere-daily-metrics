# V2 Jun 26 + traffic payload update

Built on top of V2 logo-lowered build.

Changes:
- V2 start date default is now 26 June 2026.
- V1 is capped to stop before 26 June 2026.
  - V1 effective end date is forced to 25 June 2026 if no earlier end date is selected.
  - Sidebar includes a caption explaining the V1 cutoff.
- Added V2 support for the new `page_view` event.
  - `page_view` is treated as traffic.
  - Existing conversion metrics remain counted from their specific events.
- Added V2 support for the new `traffic` payload object:
  - traffic.page.url
  - traffic.page.referrer
  - traffic.marketing.utm_source
  - traffic.marketing.utm_medium
  - traffic.marketing.utm_campaign
  - traffic.marketing.utm_term
  - traffic.marketing.utm_content
  - traffic.marketing click IDs
  - traffic.attribution first_touch / last_touch / last_non_direct_touch
- Existing fallback paths from the older payload remain preserved.

Preserved:
- Conditional per-tab filters
- Visible filter names above dropdowns
- Clean tab names
- Logo lowered spacing fix
- All V2 tabs 1 through 7
- V1 sidebar access

Replace these files in GitHub:
- streamlit_d2c_metrics_app_supabase.py
- d2c_clean_external_metrics_report.py
- requirements.txt

After pushing, reboot the Streamlit app and clear cache if needed.
