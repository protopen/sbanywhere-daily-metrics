# Product Stats actor.email update

Updates Product Stats so the `Email` column prefers `actor.email`.

Fallback:
- If `actor.email` is blank, it falls back to the existing normalized email extraction.

This version also preserves:
- Sign Up_total mapping fix for `sign_up_total`
- Clean External exclusions for abhishek, santosh, and ankita
- Email as the second Product Stats column

Replace these files in GitHub:
- streamlit_d2c_metrics_app_supabase.py
- d2c_clean_external_metrics_report.py
- requirements.txt

Streamlit secrets do not change.
