-- Run once in Supabase SQL Editor (paste the code, not the filename).
-- Trips: a last day, notes, and an outbound + return flight (flight, departs, lands).
-- Safe to run even if you ran an earlier version of this file.

alter table public.events
  add column if not exists end_date   date,                                  -- empty = a one-day event
  add column if not exists details    text check (details is null or char_length(details) <= 300),
  add column if not exists out_date   date,                                  -- flight dates (empty = first / last day of the trip)
  add column if not exists ret_date   date,
  add column if not exists out_flight text check (out_flight is null or char_length(out_flight) <= 60),
  add column if not exists out_dep    time,
  add column if not exists out_arr    time,
  add column if not exists ret_flight text check (ret_flight is null or char_length(ret_flight) <= 60),
  add column if not exists ret_dep    time,
  add column if not exists ret_arr    time;
