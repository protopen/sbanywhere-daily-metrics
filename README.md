# V2 ignore traffic object update

Built on top of V2 Soumya exclusion update.

Change:
- V2 source/campaign/UTM attribution now ignores the `traffic` object.
- Source Metrics, Paid Campaign Metrics, Detail View UTM fields, and Order Event Detail UTM fields now use:
  - `source.attribution.*`
  - `source.utm.*`
  - page URL query params
  - older non-traffic fallback fields
- The `traffic.*` paths are no longer used for source bucketing or campaign/adset/ad extraction.

Current mapping:
- Source bucket uses `source.attribution`, `source.utm`, URL params, and referrer.
- Campaign Name uses `source.attribution.utm_campaign`, then `source.utm.campaign`, then older UTM fields / URL params.
- Adset Name uses `source.attribution.utm_term`, then `source.utm.term`, then older term/adset fields / URL params.
- Ad Name uses `source.attribution.utm_content`, then `source.utm.content`, then older content/ad fields / URL params.

Preserved:
- Soumya Ramtri / ramtrisoumya11@gmail.com exclusion.
- Product Sub Category removed from V2 Detail View.
- All V2 event-count metrics use unique session/identity counts.
- Traffic (Total) still comes only from unique page_view identities.
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
