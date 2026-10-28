-- Run once in Supabase: SQL Editor > New query > paste > Run.
-- Adds: notes + completed date on tasks, and an optional sync URL on habits.

alter table public.tasks
  add column notes text not null default '' check (char_length(notes) <= 2000),
  add column done_at timestamptz;

alter table public.habits
  add column source_url text check (source_url is null or source_url like 'https://%');
