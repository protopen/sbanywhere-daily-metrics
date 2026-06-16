# Requested Sankey flow update

Sankey flow changed to:
- Enquiry Attempted
- Sign Up_total
- First Quote_Success
- Add to Cart_Success
- Invoice Upload_Success OR Revised_offer_shown
- Plan_selected if Revised_offer_shown happened and Plan_selected is present
- Initiate_checkout
- Payment Attempted
- Payment Success

Journey stitching uses session_id first, then identity_key fallback.
Downstream stages only count if the same session reached the required earlier stages.

Replace these files in GitHub:
- streamlit_d2c_metrics_app_supabase.py
- d2c_clean_external_metrics_report.py
- requirements.txt

After pushing, reboot the Streamlit app and clear cache if needed.
