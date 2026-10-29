-- Run once in Supabase SQL Editor (paste the code, not the filename).
-- Rest days for weekly habits (e.g. the gym): a day you chose not to train.
-- A rest day is NOT a tick, so it never changes the 3-a-week counter, streaks or stats.

create table public.habit_rest (
  id         uuid primary key default gen_random_uuid(),
  owner      uuid not null default auth.uid() references auth.users(id) on delete cascade,
  habit_id   text not null,          -- the habit's id, stored as text so it works whatever type the habits table uses
  rest_on    date not null,
  created_at timestamptz not null default now(),
  unique (owner, habit_id, rest_on)
);

alter table public.habit_rest enable row level security;

create policy "own rest days: read"   on public.habit_rest for select using (owner = auth.uid());
create policy "own rest days: add"    on public.habit_rest for insert with check (owner = auth.uid());
create policy "own rest days: delete" on public.habit_rest for delete using (owner = auth.uid());

grant select, insert, delete on public.habit_rest to authenticated;
