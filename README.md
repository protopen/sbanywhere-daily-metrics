# V2 Tab 1 build

This build keeps the existing dashboard as V1 and adds the new V2 dashboard as the default.

Sidebar:
- Dashboard version selector
  - V2: New Dashboard opens by default
  - V1: Current Dashboard remains accessible
- V2 default start date is 18 June 2026
- V1 default start date remains 21 May 2026

V2 Tab 1:
- Filters:
  - Date From / Till via the global date filter
  - Journey Type from JSON `flow.method`
  - Invoice Status from JSON `flow.status`
- Event sequence:
  - enquiry_attempted
  - sign_up
  - initiate_checkout
  - payment_attempted
  - payment_success
  - payment_failure
- Table:
  - Date
  - Enquiry Attempted_Total
  - Sign Up_ Total
  - Initiate Checkout
  - Payment Success
  - Payment Failure
  - Gross GWP $

Notes:
- JSON payload structure is the same as before.
- Journey Type reads from `event_data.flow.method`, with fallbacks to `data.flow.method` and `flow.method`.
- Invoice Status reads from `event_data.flow.status`, with fallbacks to `data.flow.status` and `flow.status`.

Replace these files in GitHub:
- streamlit_d2c_metrics_app_supabase.py
- d2c_clean_external_metrics_report.py
- requirements.txt

After pushing, reboot the Streamlit app and clear cache if needed.
