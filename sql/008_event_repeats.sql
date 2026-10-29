-- Run once in Supabase SQL Editor (paste the code, not the filename).
-- Only for people who already ran 007: adds repeating events (and moving/cancelling single days).

alter table public.events
  add column repeat text not null default 'none'
    check (repeat in ('none', 'daily', 'weekdays', 'weekly', 'biweekly', 'monthly')),
  add column repeat_until date,                          -- empty = repeats for ever
  add column skip_dates date[] not null default '{}';    -- single days taken out of a repeating event
