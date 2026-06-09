# Surebright Anywhere Traffic & Campaign Insights - Supabase version

This version removes file upload and OneDrive inputs. The app reads event payloads directly from a Supabase table, then keeps the same Clean External metrics, attribution, product, retailer, audit, and download tabs.

## Files

- `streamlit_d2c_metrics_app_supabase.py` - Streamlit UI
- `d2c_clean_external_metrics_report.py` - metrics backend
- `surebright_logo_homepage.webp` - homepage logo
- `requirements.txt` - Python dependencies
- `.streamlit/secrets.toml.example` - copy to `.streamlit/secrets.toml`
- `supabase_setup.sql` - optional recommended table/view structure

## Local setup

```bash
pip install -r requirements.txt
mkdir -p .streamlit
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
```

Edit `.streamlit/secrets.toml` with your Supabase details.

```toml
[supabase]
url = "https://YOUR_PROJECT_REF.supabase.co"
key = "YOUR_SUPABASE_SERVICE_ROLE_OR_READ_ONLY_KEY"
table = "d2c_raw_events"
json_column = "raw"
timestamp_column = "occurred_at"
select_columns = "*"
page_size = 1000
```

Then run:

```bash
streamlit run streamlit_d2c_metrics_app_supabase.py
```

## Expected Supabase table

The app expects one row per event and a JSON/JSONB column containing the full event payload. By default that column is named `raw`.

Recommended columns:

- `occurred_at` timestamptz
- `event_id` text
- `event_type` text
- `raw` jsonb

If your table uses a different JSON column name, update `json_column` in secrets or in the app sidebar under Advanced table settings.

## Security note

For local/internal use, a service-role key works, but do not commit it to GitHub. For hosted/public deployments, prefer a restricted read-only key or a Supabase view/policy that only exposes the required event rows.
