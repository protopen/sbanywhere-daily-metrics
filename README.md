# Product table email update

Adds an `Email` column as the second column in Product Stats.

This version also preserves:
- Sign Up_total mapping fix for `sign_up_total`
- Clean External exclusion for abhishek, santosh, and ankita

Replace these files in GitHub:
- streamlit_d2c_metrics_app_supabase.py
- d2c_clean_external_metrics_report.py
- requirements.txt

Streamlit secrets do not change.
