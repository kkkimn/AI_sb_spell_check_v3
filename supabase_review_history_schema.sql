create table if not exists public.review_history (
  history_id text primary key,
  created_at timestamptz not null,
  saved_at timestamptz not null,
  original_name text not null,
  correction_count integer not null default 0,
  metadata jsonb not null
);

create index if not exists review_history_created_at_idx
  on public.review_history (created_at desc);

create index if not exists review_history_original_name_idx
  on public.review_history (original_name);

insert into storage.buckets (id, name, public)
values ('review-history', 'review-history', false)
on conflict (id) do nothing;

create table if not exists public.app_settings (
  setting_key text primary key,
  value jsonb not null,
  saved_at timestamptz not null
);
