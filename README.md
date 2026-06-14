# Daily Metrics revised_offer_shown column update

Built on top of the current High Intent invoice-slim version.

Change:
- Adds `revised_offer_shown` to the Daily Metrics date-wise compact table.
- The new column appears immediately after `Invoice Upload_Success`.
- It maps to the existing Daily Metrics column `Revised Offer`.

Preserved:
- High Intent tab changes.
- Invoice dropoff slim table.
- Sign Up dropoff slim table.
- soumyaramtri@gmail.com exclusion.
- Sales tab.
- Product table changes and prior fixes.

Replace these files in GitHub:
- streamlit_d2c_metrics_app_supabase.py
- d2c_clean_external_metrics_report.py
- requirements.txt

After pushing, reboot the Streamlit app and clear cache if needed.
