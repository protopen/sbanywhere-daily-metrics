# V2 reverted top-navbar logo change

This reverts the last change that moved the logo into a fixed top navbar.

Restored state:
- Logo is back in the regular page body.
- Previous logo spacing/clipping fix is preserved.
- V2 start date remains 18 June 2026.
- Conditional per-tab filters remain preserved.
- Visible filter names above dropdowns remain preserved.
- Clean tab names remain preserved.
- All V2 tabs 1 through 7 remain preserved.
- V1 sidebar access remains preserved.

Replace these files in GitHub:
- streamlit_d2c_metrics_app_supabase.py
- d2c_clean_external_metrics_report.py
- requirements.txt

After pushing, reboot the Streamlit app and clear cache if needed.
