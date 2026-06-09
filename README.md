# SureBright D2C Clean External Metrics Streamlit App

This Streamlit app turns the D2C raw event export into daily Clean External metrics, totals, and an audit file.

## Files included

- `streamlit_d2c_metrics_app.py` - Streamlit UI
- `d2c_clean_external_metrics_report.py` - shared metric logic
- `requirements_streamlit_d2c.txt` - Python dependencies

## Install

```bash
pip install -r requirements_streamlit_d2c.txt
```

## Run

```bash
streamlit run streamlit_d2c_metrics_app.py
```

Then upload one of these source formats:

- `.xlsx` or `.csv` with a `Raw` column containing JSON payloads
- `.txt` or `.jsonl` with one JSON event per line

## Clean External definition

The app includes only production `surebrightanywhere.com` traffic and excludes:

- Abhishek / Santosh identity rows
- `@surebright.com`
- `@surerbright.com`
- `@example.com`
- localhost, Webflow, Amplify preview URLs
- `_meta_test=1`

The exclusion also propagates to linked events by email, lead ID, session ID, and anonymous ID.

## Outputs

Inside the app you can download:

- Excel workbook with `Daily Metrics`, `Totals`, `Event Audit`, and `Run Metadata`
- Daily metrics CSV
- Clean event audit CSV
- Full audit CSV including excluded rows

## Metric notes

- User-level metrics count unique users per day.
- Product-level metrics count product/item quantities.
- Gross GWP comes from successful payment checkout/warranty amounts.
- Dates are bucketed by the selected timezone, default `Asia/Kolkata`.
