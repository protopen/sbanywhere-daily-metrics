# High Intent table update

Built on top of the High Intent version.

Changes:
- Removes `User Key` from both High Intent tables.
- Sign Up_total dropoff table now includes all fields from the event form/forms object.
- Invoice-or-later dropoff table now uses name/email from the event form object when available.
- Invoice-or-later dropoff table now includes full line_items object data in `Line Items Data`.
- Invoice-or-later dropoff table also flattens the first line item into `Line Item: ...` columns for easier scanning.

Preserved:
- `soumyaramtri@gmail.com` exclusion.
- High Intent unique dropoff ticker.
- Sales tab.
- Product table aligned to Daily Metrics.
- Product Events removed.
- Product Email sourced from actor.email first.
- Sign Up_total product metric counts sign_up_total.
- Initiate Checkout product metric counts initiate_checkout.
- Clean External exclusions for abhishek, santosh, ankita, and soumyaramtri@gmail.com.

Replace these files in GitHub:
- streamlit_d2c_metrics_app_supabase.py
- d2c_clean_external_metrics_report.py
- requirements.txt

After pushing, reboot the Streamlit app and clear cache if needed.
