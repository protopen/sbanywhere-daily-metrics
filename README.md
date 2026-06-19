# V2 clean tab names

Built on top of V2 filter label cleanup.

Change:
- Removed `Tab 1:`, `Tab 2:`, etc. prefixes from V2 tab labels.

New tab labels:
- Daily Metrics
- Product Category Metrics
- Source Metrics
- Paid Campaign Metrics
- Category Metrics
- Detail View
- Order Event Detail

Replace these files in GitHub:
- streamlit_d2c_metrics_app_supabase.py
- d2c_clean_external_metrics_report.py
- requirements.txt

After pushing, reboot the Streamlit app and clear cache if needed.
