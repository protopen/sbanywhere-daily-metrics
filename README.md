# Strict Sankey funnel fix

Built on top of the current invoice_uploaded + revised_offer version.

Problem fixed:
- Sankey could show First Quote_Success higher than Enquiry Attempted because it connected any reached stages inside a user/session even if earlier stages were missing from the selected date range.

Change:
- Sankey is now a strict session-based funnel.
- Later stages only count if the same session reached all previous main funnel stages.
- Journey stitching uses session_id first, then identity_key fallback.
- Sign Up_total Sankey mapping now includes sign_up_total / signup_total / sign_up.

Main funnel:
- Enquiry Attempted
- Sign Up_total
- First Quote_Success
- Offer_Selected
- Invoice Upload_Success
- Add to Cart_Success
- Payment Attempted
- Payment Success

Preserved:
- Daily Metrics revised_offer_shown column.
- Invoice Uploaded = Yes for invoice_upload_success or revised_offer_shown in High Intent.
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
