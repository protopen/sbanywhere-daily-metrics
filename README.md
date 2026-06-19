# V2 logo and stale text fix

Built on top of V2 compact UI no intro.

Changes:
- Reduced logo width and added object-fit CSS so the logo does not crop/cut.
- Removed stale filter helper text:
  Journey Type = flow.method • Invoice Status = flow.status
- Changed V2 filter bar from 3 columns to 2 columns after removing the caption.

Replace these files in GitHub:
- streamlit_d2c_metrics_app_supabase.py
- d2c_clean_external_metrics_report.py
- requirements.txt

After pushing, reboot the Streamlit app and clear cache if needed.
