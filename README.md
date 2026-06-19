# V2 compact UI no intro text

Built on top of V2 compact UI.

Change:
- Removed the V2 intro heading and caption:
  - V2 Dashboard
  - V2 uses the new funnel logic and reads Journey Type / Invoice Status from flow.method and flow.status inside the JSON payload.

Preserved:
- V2 opens by default.
- V1 remains accessible from the sidebar.
- Compact filter bar.
- V2 Tab 1 Daily Metrics.

Replace these files in GitHub:
- streamlit_d2c_metrics_app_supabase.py
- d2c_clean_external_metrics_report.py
- requirements.txt

After pushing, reboot the Streamlit app and clear cache if needed.
