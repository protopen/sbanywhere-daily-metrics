# SureBright D2C Metrics Streamlit App v4

This version adds the requested attribution tables:

1. **By campaign**
   - Source
   - Campaign
   - Events
   - Sessions
   - Leads captured
   - Unique fbclicks

2. **Event breakdown for UTM traffic**
   - Campaign
   - Homepage form submit
   - Quote generated
   - Plan selected
   - Quote lead captured

The attribution tab uses `source.attribution.utm_source`, `source.attribution.utm_campaign`, `session.session_id`, and unique `source.attribution.fbclid` values. It only includes campaign/ad-click traffic, so plain direct/referral traffic with campaign `(none)` is not included in these two tables.

## Run locally

```bash
pip install -r requirements.txt
streamlit run streamlit_d2c_metrics_app_onedrive.py
```

## Files

- `streamlit_d2c_metrics_app_onedrive.py` - Streamlit UI
- `d2c_clean_external_metrics_report.py` - reporting/parsing backend
- `requirements.txt` - dependencies
