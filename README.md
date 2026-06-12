# High Intent tab update

Built on top of the stable Sales-tab version.

Changes:
- Adds `soumyaramtri@gmail.com` to Clean External exclusions.
- Adds a new `High Intent` tab.
- Adds a top-level ticker: Unique high-intent dropoff users.
- Adds table 1: Dropped after Sign Up_total.
- Adds table 2: Dropped after invoice upload or later.
- Uses session_id as the primary key for dropoff detection, with identity_key fallback if session_id is missing.
- Includes PII in High Intent tables for follow-up workflows.

Dropoff logic:
- Sign Up_total dropoff: session reached Sign Up_total but did not upload invoice, add to cart, initiate checkout, or complete payment.
- Invoice-or-later dropoff: session uploaded invoice or reached a later stage but did not complete payment.

High Intent tables include:
- Session ID
- User Key
- Email
- Name
- Phone
- Address
- Lead ID
- First Event Date
- Last Event Date
- Last Stage
- Event Count
- Invoice Uploaded
- Invoice Upload Failed
- Product / warranty details
- UTM fields

Preserved from stable version:
- Sales tab
- Product table aligned to Daily Metrics
- Product Events removed
- Product Email sourced from actor.email first
- Sign Up_total product metric counts sign_up_total
- Initiate Checkout product metric counts initiate_checkout
- Compatibility aliases for older helper names
- Clean External exclusions for abhishek, santosh, ankita

Replace these files in GitHub:
- streamlit_d2c_metrics_app_supabase.py
- d2c_clean_external_metrics_report.py
- requirements.txt

After pushing, reboot the Streamlit app and clear cache if needed.
