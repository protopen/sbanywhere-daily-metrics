-- Recommended event table shape if you do not already have one.
-- If your existing table already has a JSONB payload column, you can skip this.

create table if not exists public.d2c_raw_events (
  id bigserial primary key,
  event_id text unique,
  event_type text,
  occurred_at timestamptz,
  raw jsonb not null,
  inserted_at timestamptz not null default now()
);

create index if not exists d2c_raw_events_occurred_at_idx
  on public.d2c_raw_events (occurred_at);

create index if not exists d2c_raw_events_event_type_idx
  on public.d2c_raw_events (event_type);

-- Optional read-only view for the Streamlit app.
create or replace view public.d2c_streamlit_events as
select id, event_id, event_type, occurred_at, raw
from public.d2c_raw_events;
