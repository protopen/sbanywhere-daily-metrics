# High Intent Invoice table cleanup

Built on top of the High Intent signup-slim version.

Changes to `Dropped after invoice upload or later` table:
- Removes base PII fields from older row structure:
  - email
  - name
  - phone
  - address
  - lead id
- Removes:
  - product title
  - gross gwp
  - all flattened `Line Item: ...` fields
- Keeps only full line_items object in:
  - Line Items Data

Invoice table column order is now:
1. First Event Date
2. Last Event Date
3. Name
4. Email
5. Last Stage
6. Invoice Uploaded
7. Invoice Upload Failed
8. Product Category
9. Product Brand
10. Manufacturer
11. Model Number
12. Warranty / Plan Name
13. Warranty Price
14. Retailer
15. Line Items Data
16. UTM Source
17. UTM Medium
18. UTM Campaign

Notes:
- Name is sourced from `fields.name` / form name equivalents.
- Email is sourced from `fields.email` / form email equivalents.
- Sign Up dropoff table remains as previously slimmed:
  First Event Date, Last Event Date, Name, Email, Last Stage, Product Category, Product Price.

Preserved:
- User Key removed from High Intent tables.
- High Intent unique dropoff ticker.
- soumyaramtri@gmail.com exclusion.
- Sales tab.
- Product table changes and prior fixes.

Replace these files in GitHub:
- streamlit_d2c_metrics_app_supabase.py
- d2c_clean_external_metrics_report.py
- requirements.txt

After pushing, reboot the Streamlit app and clear cache if needed.
