# V2 logo lowered

Built on top of reverted top-navbar logo build.

Changes:
- Increased page top padding so the logo sits below the Streamlit top navbar.
- Added extra logo top margin to prevent clipping.
- Preserved regular in-page logo placement, without fixed custom navbar.

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
