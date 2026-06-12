# High Intent Sign Up table cleanup

Built on top of the High Intent forms + line_items version.

Changes to `Dropped after Sign Up_total` table:
- Removes unwanted fields:
  - form_name
  - form_id
  - fields.method
  - retailer
  - gross gwp
  - product brand
  - manufacturer
  - warranty / plan name
  - product title
  - invoice uploaded
  - invoice upload failed
  - event count
  - name / phone / address / lead ID / email / session id from base row
- Final displayed columns are exactly:
  1. First Event Date
  2. Last Event Date
  3. Name
  4. Email
  5. Last Stage
  6. Product Category
  7. Product Price

Notes:
- Name is sourced from `fields.name` / form name equivalents.
- Email is sourced from `fields.email` / form email equivalents.
- Product Category and Product Price are sourced from form fields where available.

Preserved:
- Invoice dropoff table keeps name/email from form object and line_items object data.
- User Key remains removed from High Intent tables.
- High Intent unique dropoff ticker.
- soumyaramtri@gmail.com exclusion.
- Sales tab.
- Product table changes and prior fixes.

Replace these files in GitHub:
- streamlit_d2c_metrics_app_supabase.py
- d2c_clean_external_metrics_report.py
- requirements.txt

After pushing, reboot the Streamlit app and clear cache if needed.
