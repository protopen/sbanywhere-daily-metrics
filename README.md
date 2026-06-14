# High Intent Invoice Uploaded update

Built on top of the current Daily revised_offer_shown version.

Change:
- In High Intent > Dropped after invoice upload or later, the `Invoice Uploaded` column now says `Yes` if either of these events happened in the session:
  - invoice_upload_success / invoice_uploaded
  - revised_offer_shown / revised_offer

Preserved:
- Daily Metrics revised_offer_shown column.
- High Intent invoice slim table.
- High Intent signup slim table.
- soumyaramtri@gmail.com exclusion.
- Sales tab.
- Product table changes and prior fixes.

Replace these files in GitHub:
- streamlit_d2c_metrics_app_supabase.py
- d2c_clean_external_metrics_report.py
- requirements.txt

After pushing, reboot the Streamlit app and clear cache if needed.
