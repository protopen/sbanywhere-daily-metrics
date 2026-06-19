# V2 dropdown checkbox-style filter update

Built on top of V2 Tab 7.

Changes:
- Reverted the always-open checkbox panels.
- Filters are now compact placeholder-style dropdowns again.
- Filter names remain visible above each dropdown.
- The closed placeholder/dropdown height is fixed.
- The opened dropdown menu keeps the checkbox-style multiselect behavior through Streamlit's native multiselect UI.

Preserved:
- All V2 tabs 1 through 7.
- V1 sidebar access.

Replace these files in GitHub:
- streamlit_d2c_metrics_app_supabase.py
- d2c_clean_external_metrics_report.py
- requirements.txt

After pushing, reboot the Streamlit app and clear cache if needed.
