# Product Stats Daily Metrics alignment

Updates Product Stats to show the same funnel metric labels as the Daily Metrics table:
- Enquiry Attempted
- Sign Up_total
- Add to cart
- Invoice Upload_Success
- Initiate Checkout
- Payment Success

This version also preserves:
- Product Stats Email as the second column, sourced from actor.email first
- Sign Up_total mapping fix for sign_up_total
- Clean External exclusions for abhishek, santosh, and ankita

Replace these files in GitHub:
- streamlit_d2c_metrics_app_supabase.py
- d2c_clean_external_metrics_report.py
- requirements.txt

Streamlit secrets do not change.
