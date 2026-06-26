# V2 Paid Campaign Metrics attribution-object update

Built on top of V2 missing helpers hotfix.

Change:
- Paid Campaign Metrics now explicitly uses `source.attribution` first.
- No `traffic.*` fields are used for campaign/adset/ad extraction or paid source bucketing.

Paid Campaign Metrics mapping:
- Campaign Name:
  - source.attribution.utm_campaign
  - source.attribution.campaign
  - source.utm.campaign / source.utm_campaign
  - URL utm_campaign from non-traffic page URL fallbacks
  - older non-traffic fallback UTM fields
- Adset Name:
  - source.attribution.utm_term
  - source.attribution.term
  - source.utm.term / source.utm_term
  - URL utm_term / utm_adset from non-traffic page URL fallbacks
  - older non-traffic fallback adset fields
- Ad Name:
  - source.attribution.utm_content
  - source.attribution.content
  - source.utm.content / source.utm_content
  - URL utm_content / utm_ad from non-traffic page URL fallbacks
  - older non-traffic fallback ad fields
- Paid Campaign Source:
  - source.attribution source/medium/click IDs first
  - source.utm fallback
  - older non-traffic fallback UTM fields

Preserved:
- V2 source/campaign/UTM attribution ignores the traffic object.
- All missing V2 helpers are defined.
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
