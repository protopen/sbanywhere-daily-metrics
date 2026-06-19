# V2 start date hard-coded

Built on top of V2 filter names visible build.

Change:
- V2 start date is hard-coded/defaulted to 18 June 2026.

Preserved:
- V2 opens by default.
- V1 remains accessible from the sidebar.
- Conditional per-tab filters.
- Visible filter names above dropdowns.
- Clean tab names.
- All V2 tabs 1 through 7.

Replace these files in GitHub:
- streamlit_d2c_metrics_app_supabase.py
- d2c_clean_external_metrics_report.py
- requirements.txt

After pushing, reboot the Streamlit app and clear cache if needed.
