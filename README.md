# V2 Tab 7 build

Built on top of V2 Tab 6.

Changes:
- Adds V2 Tab 7: Order Event Detail.
- Uses same filters as Tab 4:
  - Date From / Till via global date filter
  - Journey Type from flow.method
  - Invoice Status from flow.status
  - Product Category, default All
  - Paid Campaign Source

Tab 7 table columns:
- Date
- Name
- Email
- Journey Flow
- Event
- Order / Payment ID
- Product Category
- Product Title
- Product Brand
- Manufacturer
- Model Number
- Product Condition
- Quantity
- Product Unit Price
- Product Total Price
- Warranty Type
- Warranty / Plan Name
- Warranty Price
- Warranty Term Months
- Warranty Provider
- Manufacturer Warranty
- Eligible
- Gross GWP $
- UTM Source
- UTM Medium
- UTM Campaign

Replace these files in GitHub:
- streamlit_d2c_metrics_app_supabase.py
- d2c_clean_external_metrics_report.py
- requirements.txt

After pushing, reboot the Streamlit app and clear cache if needed.
