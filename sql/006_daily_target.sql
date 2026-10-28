-- Run once in Supabase SQL Editor (paste the code, not the filename).
-- Lets a habit need several ticks a day (e.g. drink water 4 times).

alter table public.habits
  add column daily_target int not null default 1 check (daily_target between 1 and 20);

alter table public.habit_log
  add column count int not null default 1 check (count between 1 and 50);

update public.habits set daily_target = 4 where name = 'Drink water';
