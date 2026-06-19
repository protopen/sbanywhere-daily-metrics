# V2 logo spacing fix

Built on top of V2 start date hard-coded build.

Changes:
- Increased top padding so the logo is not clipped.
- Added safe image CSS so the logo uses full height and does not get cut.
- Added a little spacing below the logo.
- Moved the page content slightly lower vertically.

Preserved:
- V2 start date default/hard-code: 18 June 2026
- Conditional per-tab filters
- Visible filter names above dropdowns
- Clean tab names
- All V2 tabs 1 through 7
- V1 sidebar access

Replace these files in GitHub:
- streamlit_d2c_metrics_app_supabase.py
- d2c_clean_external_metrics_report.py
- requirements.txt

After pushing, reboot the Streamlit app and clear cache if needed.
