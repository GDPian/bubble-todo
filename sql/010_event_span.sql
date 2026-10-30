-- Run once in Supabase SQL Editor (paste the code, not the filename).
-- Adds trips / multi-day events to "+ Event": a last day, plus a free-text box for flight details.

alter table public.events
  add column end_date date,                                                -- empty = a one-day event
  add column details  text check (details is null or char_length(details) <= 300);
