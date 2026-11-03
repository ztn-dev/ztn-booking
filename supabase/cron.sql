-- Hourly auto-close: Confirmed bookings whose end time has passed -> Completed.
-- First enable pg_cron: Dashboard -> Database -> Extensions -> pg_cron.
-- Then run this once in the SQL Editor.

select cron.unschedule(jobid) from cron.job where jobname = 'ztn-close-completed';

select cron.schedule(
  'ztn-close-completed',
  '5 * * * *',                       -- five past every hour
  $$ select public._close_completed() $$
);

-- Weekly archive: every Sunday at 23:55 Harare time (21:55 UTC), save a
-- snapshot of the week that is ending (Monday to Sunday).
select cron.unschedule(jobid) from cron.job where jobname = 'ztn-weekly-archive';

select cron.schedule(
  'ztn-weekly-archive',
  '55 21 * * 0',                     -- Sunday 21:55 UTC = 23:55 Harare
  $$ select public._close_completed();
     select public._archive_week((now() at time zone public.ztn_tz())::date) $$
);
