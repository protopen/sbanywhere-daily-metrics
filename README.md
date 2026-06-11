# Product Events removed

Removes the `Product Events` column from Product Stats.

This version preserves:
- Product Stats metrics aligned to the Daily Metrics table
- Email as the second Product Stats column, sourced from actor.email first
- Sign Up_total mapping fix for sign_up_total
- Clean External exclusions for abhishek, santosh, and ankita

Replace these files in GitHub:
- streamlit_d2c_metrics_app_supabase.py
- d2c_clean_external_metrics_report.py
- requirements.txt

Streamlit secrets do not change.
