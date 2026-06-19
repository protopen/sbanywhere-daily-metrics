# V2 visible filter names fix

Built on top of V2 conditional filters.

Change:
- Filter labels now render using Streamlit-native markdown above each dropdown.
- This avoids the previous custom HTML/CSS label being hidden or not visible.
- Dropdown placeholder remains generic `Select`.
- Conditional per-tab filter rendering is preserved.

Replace these files in GitHub:
- streamlit_d2c_metrics_app_supabase.py
- d2c_clean_external_metrics_report.py
- requirements.txt

After pushing, reboot the Streamlit app and clear cache if needed.
