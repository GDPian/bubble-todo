-- Run once in Supabase SQL Editor (paste the code, not the filename).
-- Adds a note to each habit tick, and a per-habit "ask what I did" switch.

alter table public.habit_log
  add column note text not null default '' check (char_length(note) <= 500);

alter table public.habits
  add column ask_note boolean not null default false;

-- Send a CV asks "what did you do?" every time you tick it
update public.habits set ask_note = true where name = 'Send a CV';
