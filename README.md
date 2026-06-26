# V2 missing helpers hotfix

Built on top of V2 detail value hotfix.

Hotfix:
- Added all missing V2 helpers referenced at runtime:
  - `_v2_line_items_from_obj`
  - `_v2_paid_source_bucket`
  - `_v2_full_product_values`
- Confirmed all referenced `_v2_*` helper functions are now defined.
- Fixes runtime errors such as:
  - NameError: name '_v2_full_product_values' is not defined
  - NameError: name '_v2_paid_source_bucket' is not defined
  - NameError: name '_v2_line_items_from_obj' is not defined

Preserved:
- V2 source/campaign/UTM attribution ignores the `traffic` object.
- Soumya Ramtri / ramtrisoumya11@gmail.com exclusion.
- Product Sub Category removed from V2 Detail View.
- All V2 event-count metrics use unique session/identity counts.
- Traffic (Total) still counts only unique page_view identities.
- Sign Up_ Total counts event_name = sign_up_total.
- Enquiry Success remains removed from Product Category Metrics.
- Invoice Status filter is unselected by default.
- page_view remains separate from conversion events.
- V2 starts from 26 Jun 2026.
- V1 stops before 26 Jun 2026.
- Journey Type anomaly logs.
- Conditional per-tab filters.
- Visible filter names.
- Clean tab names.
- Logo lowered spacing.
- All V2 tabs 1 through 7 plus Logs.

Replace these files in GitHub:
- streamlit_d2c_metrics_app_supabase.py
- d2c_clean_external_metrics_report.py
- requirements.txt

After pushing, reboot the Streamlit app and clear cache if needed.
