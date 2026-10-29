-- Run once in Supabase SQL Editor (paste the code, not the filename).
-- 1) Tasks may have NO due date ("someday" tasks: still bubbles, never shown in the calendar).
-- 2) A small table of your own calendar events (work, personal, hers) typed in with "+ Event".

alter table public.tasks alter column due drop not null;

create table public.events (
  id         uuid primary key default gen_random_uuid(),
  owner      uuid not null default auth.uid() references auth.users(id) on delete cascade,
  title      text not null check (char_length(title) between 1 and 80),
  on_date    date not null,
  start_time time,                     -- empty = an all-day event
  end_time   time,
  tag        text not null default 'personal' check (tag in ('work', 'personal', 'her')),
  created_at timestamptz not null default now()
);

alter table public.events enable row level security;

create policy "own events: read"   on public.events for select using (owner = auth.uid());
create policy "own events: add"    on public.events for insert with check (owner = auth.uid());
create policy "own events: change" on public.events for update using (owner = auth.uid()) with check (owner = auth.uid());
create policy "own events: delete" on public.events for delete using (owner = auth.uid());

grant select, insert, update, delete on public.events to authenticated;
