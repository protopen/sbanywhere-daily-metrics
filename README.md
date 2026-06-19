# V2 checkbox filter UI update

Built on top of V2 Tab 7.

Changes:
- Added visible filter labels above each filter:
  - Journey Type
  - Invoice Status
  - Product Category
  - Paid Campaign Source
  - Traffic Source
- Replaced dropdown-style multiselects with compact checkbox-style filter panels.
- Filter panel/header height is fixed, with internal scroll for longer option lists.
- Preserved all V2 tabs 1 through 7 and V1 sidebar access.

Replace these files in GitHub:
- streamlit_d2c_metrics_app_supabase.py
- d2c_clean_external_metrics_report.py
- requirements.txt

After pushing, reboot the Streamlit app and clear cache if needed.
