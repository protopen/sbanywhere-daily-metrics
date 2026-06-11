# Product Stats signup compatibility fix

Fixes:
- NameError: name 'SIGNUP_EVENTS' is not defined

Changes:
- Product Stats `Sign Up_total` counts `sign_up_total` directly.
- Product Stats `Initiate Checkout` counts `initiate_checkout` directly.
- Adds compatibility aliases so older references to SIGNUP_EVENTS or PAYMENT_ATTEMPT_EVENTS will not crash.
- Keeps Product Events removed.
- Keeps Email as the second Product Stats column, sourced from actor.email first.
- Keeps Clean External exclusions for abhishek, santosh, and ankita.

Replace these files in GitHub:
- streamlit_d2c_metrics_app_supabase.py
- d2c_clean_external_metrics_report.py
- requirements.txt

After pushing, reboot the Streamlit app and clear cache if needed.
