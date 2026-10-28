-- Run once in Supabase SQL Editor. Remembers where you drag a bubble sideways (0 = far left, 1 = far right).
alter table public.tasks
  add column pos_x real check (pos_x >= 0 and pos_x <= 1);
