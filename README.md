# V2 ignore traffic object hotfix

Built on top of V2 ignore traffic object update.

Hotfix:
- Removed Python 3.10-only pipe type hints that can break Streamlit Cloud if the app runs on an older Python version.
- Removed remaining `traffic.*` URL parameter fallback from campaign/adset/ad extraction.
- Source/campaign/UTM attribution now ignores `traffic` object paths completely.

Attribution now uses:
- source.attribution.*
- source.utm.*
- source.page_url / source.url
- root page_url / url
- context.page.url
- event_data.page_url
- older non-traffic fallback UTM fields

Preserved:
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
