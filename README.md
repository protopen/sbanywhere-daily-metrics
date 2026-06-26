# V2 Paid Campaign Metrics source.attribution only

Built on top of V2 Paid Campaign Metrics attribution-object update.

Change:
- Paid Campaign Metrics uses ONLY source.attribution.
- No traffic object.
- No source.utm fallback.
- No URL parsing fallback.
- No older UTM fallback fields.

Paid Campaign Metrics mapping:
- Campaign Name = source.attribution.utm_campaign, fallback source.attribution.campaign, else Unknown
- Adset Name = source.attribution.utm_term, fallback source.attribution.term, else Unknown
- Ad Name = source.attribution.utm_content, fallback source.attribution.content, else Unknown
- Paid Campaign Source = source.attribution source/medium/referrer/click IDs only

Also:
- Normalizes campaign/adset/ad display text by replacing + with spaces and collapsing repeated spaces.
- Avoids false Meta classification from loose ig substring matching.

Preserved:
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
