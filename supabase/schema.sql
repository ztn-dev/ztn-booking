-- =============================================================================
--  ZTN Studio Booking - Supabase schema
--  Run once in Supabase: Dashboard -> SQL Editor -> paste -> Run.
--  Safe to re-run: functions/policies are replaced, data is kept.
--
--  The database enforces every booking rule, so the web pages stay thin:
--    * Statuses: Requested | Pending Approval | Conflict | Confirmed |
--                Declined | Rescheduled | Cancelled | Completed
--    * Clash = same date + same studio + overlapping time, where setup access
--      time pulls the effective start earlier.
--    * A clashing booking is saved as Conflict; confirming a clashing booking
--      is refused and the row is set back to Conflict.
--    * Editing date/time/studio re-checks that row; declining/cancelling or
--      moving a booking re-checks the Conflict rows it was blocking.
--    * Confirmed bookings whose end time has passed become Completed (hourly).
--    * Regular weekly shows never block a booking. An overlap is noted and the
--      engineer decides; confirming over a show is allowed and recorded.
-- =============================================================================

create extension if not exists pg_net;

-- ---------------------------------------------------------------- constants --
create or replace function public.ztn_tz() returns text
language sql immutable as $$ select 'Africa/Harare' $$;

-- Monday of the current week in Harare. Weeks before it have ended and are locked.
create or replace function public.ztn_week_open() returns date
language sql stable as $$ select date_trunc('week', now() at time zone public.ztn_tz())::date $$;

create or replace function public.ztn_statuses() returns text[]
language sql immutable as $$
  select array['Requested','Pending Approval','Conflict','Confirmed',
               'Declined','Rescheduled','Cancelled','Completed']
$$;

-- Statuses that occupy a slot when checking for clashes
create or replace function public.ztn_blocking() returns text[]
language sql immutable as $$
  select array['Requested','Pending Approval','Conflict','Confirmed','Rescheduled','Completed']
$$;

-- Open bookings that get re-checked
create or replace function public.ztn_live() returns text[]
language sql immutable as $$
  select array['Requested','Pending Approval','Conflict','Confirmed','Rescheduled']
$$;

-- Shown on the weekly overview / studio monitor
create or replace function public.ztn_firm() returns text[]
language sql immutable as $$ select array['Confirmed','Completed'] $$;

-- ------------------------------------------------------------------- tables --
create table if not exists public.engineers (
  user_id    uuid primary key references auth.users(id) on delete cascade,
  full_name  text,
  created_at timestamptz not null default now()
);

create table if not exists public.app_settings (
  key   text primary key,
  value text not null default ''
);

create table if not exists public.bookings (
  id              bigint generated always as identity primary key,
  ref             text unique,
  booking_date    date not null,
  start_time      time not null,
  end_time        time not null,
  setup_time      time,
  studio          text not null
                  check (studio in ('Studio 1','Studio 2','Audio Studio','MCR','Other')),
  title           text not null check (length(title) between 1 and 200),
  requester_name  text not null check (length(requester_name) between 1 and 120),
  department      text check (length(department) <= 120),
  requester_email text not null check (requester_email ~* '^[^@\s]+@[^@\s]+\.[^@\s]+$'),
  contact_number  text check (length(contact_number) <= 60),
  producer        text check (length(producer) <= 120),
  programme_type  text check (length(programme_type) <= 60),
  live_recorded   text check (live_recorded in ('Live','Recorded','Live-to-tape')),
  crew            text check (length(crew) <= 1000),
  priority        text check (priority in ('High','Urgent')),
  status          text not null default 'Requested' check (status = any (public.ztn_statuses())),
  notes           text not null default '' check (length(notes) <= 8000),
  details         jsonb not null default '{}'::jsonb check (pg_column_size(details) <= 16000),
  approved_by     uuid references auth.users(id) on delete set null,
  approved_at     timestamptz,
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now(),
  check (end_time > start_time),
  check (setup_time is null or setup_time <= start_time)
);

-- Last status the email function handled; stops duplicate emails.
alter table public.bookings add column if not exists notified_status text;

create index if not exists bookings_date_studio_idx on public.bookings (booking_date, studio);
create index if not exists bookings_status_idx on public.bookings (status);

-- Regular programmes that repeat every week (the Excel "Weekly Shows" sheet).
-- days: ISO weekdays, 1 = Monday ... 7 = Sunday.
create table if not exists public.weekly_shows (
  id            bigint generated always as identity primary key,
  name          text not null check (length(name) between 1 and 200),
  days          smallint[] not null check (cardinality(days) > 0 and days <@ array[1,2,3,4,5,6,7]::smallint[]),
  start_time    time not null,
  end_time      time not null,
  studio        text not null default 'Studio 1'
                check (studio in ('Studio 1','Studio 2','Audio Studio','MCR','Other')),
  live_recorded text check (live_recorded in ('Live','Recorded','Live-to-tape')),
  presenter     text check (length(presenter) <= 200),
  active_from   date,
  active_to     date,
  notes         text not null default '' check (length(notes) <= 2000),
  created_at    timestamptz not null default now(),
  check (end_time > start_time),
  check (active_to is null or active_from is null or active_to >= active_from)
);

-- Keep the studio list current on tables that already exist.
alter table public.bookings drop constraint if exists bookings_studio_check;
alter table public.bookings add constraint bookings_studio_check
  check (studio in ('Studio 1','Studio 2','Audio Studio','MCR','Other'));
alter table public.weekly_shows drop constraint if exists weekly_shows_studio_check;
alter table public.weekly_shows add constraint weekly_shows_studio_check
  check (studio in ('Studio 1','Studio 2','Audio Studio','MCR','Other'));

-- Your current line-up (from the workbook). Only inserted into an empty table.
insert into public.weekly_shows (name, days, start_time, end_time)
select v.name, v.days::smallint[], v.s::time, v.e::time
  from (values
    ('Kwayedza Show',                          '{5}',         '09:00', '09:30'),
    ('GAME CHANGERS',                          '{2}',         '10:30', '11:00'),
    ('INSIDE THE SCRUM',                       '{3}',         '10:30', '11:00'),
    ('Beyond the Boundary',                    '{4}',         '10:30', '11:00'),
    ('NEWSBLITZ / NHAU / INDABA / WORLD IN 5', '{1,2,3,4,5}', '12:00', '13:00'),
    ('NEWSBLITZ',                              '{1,2,3,4,5}', '16:00', '16:30'),
    ('ULTIMATE SPORTS SHOW',                   '{1}',         '18:00', '18:30'),
    ('NEWSPLUS',                               '{1,2,3,4,5}', '19:00', '19:30')
  ) as v(name, days, s, e)
 where not exists (select 1 from public.weekly_shows);

-- ------------------------------------------------------------------ helpers --
create or replace function public.is_engineer() returns boolean
language sql stable security definer set search_path = public as $$
  select exists (select 1 from public.engineers where user_id = auth.uid())
$$;

-- Refs of bookings that clash with the given slot. firm_only = only
-- Confirmed/Completed (used when re-checking an already-confirmed booking).
create or replace function public.find_conflicts(
  p_date date, p_start time, p_end time, p_setup time, p_studio text,
  p_exclude bigint, p_firm_only boolean default false)
returns text[]
language sql stable security definer set search_path = public as $$
  select coalesce(array_agg(b.ref order by b.ref), '{}')
  from public.bookings b
  where b.booking_date = p_date
    and b.studio = p_studio
    and (p_exclude is null or b.id <> p_exclude)
    and b.status = any (case when p_firm_only then public.ztn_firm() else public.ztn_blocking() end)
    and least(coalesce(b.setup_time, b.start_time), b.start_time) < p_end
    and least(coalesce(p_setup, p_start), p_start) < b.end_time
$$;

-- Regular shows on that date/studio overlapping the slot, as 'NAME 19:00–19:30'.
-- These never block a booking: the engineer decides.
create or replace function public.find_show_overlaps(
  p_date date, p_start time, p_end time, p_setup time, p_studio text)
returns text[]
language sql stable security definer set search_path = public as $$
  select coalesce(array_agg(s.name || ' ' || to_char(s.start_time, 'HH24:MI') || chr(8211)
                            || to_char(s.end_time, 'HH24:MI') order by s.start_time, s.name), '{}')
  from public.weekly_shows s
  where s.studio = p_studio
    and extract(isodow from p_date)::smallint = any (s.days)
    and (s.active_from is null or p_date >= s.active_from)
    and (s.active_to is null or p_date <= s.active_to)
    and s.start_time < p_end
    and least(coalesce(p_setup, p_start), p_start) < s.end_time
$$;

-- Append a timestamped line to notes unless the same text is already there.
create or replace function public.stamp_note(p_notes text, p_text text) returns text
language sql stable as $$
  select case
    when position(lower(p_text) in lower(coalesce(p_notes, ''))) > 0 then coalesce(p_notes, '')
    when coalesce(p_notes, '') = '' then
      '[' || to_char(now() at time zone public.ztn_tz(), 'DD-Mon HH24:MI') || '] ' || p_text
    else p_notes || E'\n[' || to_char(now() at time zone public.ztn_tz(), 'DD-Mon HH24:MI') || '] ' || p_text
  end
$$;

create or replace function public.next_booking_ref() returns text
language plpgsql security definer set search_path = public as $$
declare
  y text := to_char(now() at time zone public.ztn_tz(), 'YYYY');
  n int;
begin
  perform pg_advisory_xact_lock(hashtext('ztn_ref_' || y));
  select coalesce(max(substring(ref from '^ZTN-' || y || '-(\d+)$')::int), 0) + 1
    into n
    from public.bookings
   where ref like 'ZTN-' || y || '-%';
  return 'ZTN-' || y || '-' || lpad(n::text, 4, '0');
end $$;

-- ----------------------------------------------------------------- triggers --
create or replace function public.bookings_before() returns trigger
language plpgsql security definer set search_path = public as $$
declare
  hits  text[];
  firm  text[];
  shows text[];
begin
  -- The email function only records notified_status: leave the booking alone.
  if tg_op = 'UPDATE' and (to_jsonb(new) - 'notified_status') = (to_jsonb(old) - 'notified_status') then
    return new;
  end if;
  new.updated_at := now();
  new.notes := coalesce(new.notes, '');

  if tg_op = 'INSERT' then
    if coalesce(new.ref, '') = '' then
      new.ref := public.next_booking_ref();
    end if;
    if new.status = any (public.ztn_live()) then
      hits := public.find_conflicts(new.booking_date, new.start_time, new.end_time,
                                    new.setup_time, new.studio, null);
      if cardinality(hits) > 0 then
        new.status := 'Conflict';
        new.notes := public.stamp_note(new.notes, 'CONFLICT with ' || array_to_string(hits, ', '));
      end if;
      if new.status <> 'Confirmed' then
        shows := public.find_show_overlaps(new.booking_date, new.start_time, new.end_time, new.setup_time, new.studio);
        if cardinality(shows) > 0 then
          new.notes := public.stamp_note(new.notes, 'Overlaps regular show ' || array_to_string(shows, ', ')
                                                    || ' ' || chr(8212) || ' engineer to decide');
        end if;
      end if;
    end if;
    return new;
  end if;

  -- UPDATE: the reference never changes
  new.ref := old.ref;

  if new.status = 'Confirmed' and old.status is distinct from 'Confirmed' then
    hits := public.find_conflicts(new.booking_date, new.start_time, new.end_time,
                                  new.setup_time, new.studio, new.id);
    if cardinality(hits) > 0 then
      new.status := 'Conflict';
      new.notes := public.stamp_note(new.notes,
                     'Cannot confirm ' || chr(8212) || ' clash with ' || array_to_string(hits, ', '));
      return new;
    end if;
    new.approved_at := now();
    new.approved_by := auth.uid();
    shows := public.find_show_overlaps(new.booking_date, new.start_time, new.end_time, new.setup_time, new.studio);
    if cardinality(shows) > 0 then
      new.notes := public.stamp_note(new.notes, 'Confirmed over regular show ' || array_to_string(shows, ', ')
                                                || ' at engineer''s discretion');
    end if;

  elsif new.status = 'Confirmed' then
    -- Already confirmed: only another Confirmed/Completed booking can knock it back.
    firm := public.find_conflicts(new.booking_date, new.start_time, new.end_time,
                                  new.setup_time, new.studio, new.id, true);
    if cardinality(firm) > 0 then
      new.status := 'Conflict';
      new.notes := public.stamp_note(new.notes, 'Clash with confirmed ' || array_to_string(firm, ', '));
    else
      hits := public.find_conflicts(new.booking_date, new.start_time, new.end_time,
                                    new.setup_time, new.studio, new.id);
      if cardinality(hits) > 0 then
        new.notes := public.stamp_note(new.notes, 'Overlaps unconfirmed ' || array_to_string(hits, ', '));
      end if;
    end if;

  elsif new.status = any (public.ztn_live()) then
    hits := public.find_conflicts(new.booking_date, new.start_time, new.end_time,
                                  new.setup_time, new.studio, new.id);
    if cardinality(hits) > 0 then
      new.status := 'Conflict';
      new.notes := public.stamp_note(new.notes, 'Clash with ' || array_to_string(hits, ', '));
    elsif new.status = 'Conflict' then
      new.status := 'Requested';
      new.notes := public.stamp_note(new.notes, 'Clash cleared');
    end if;
    shows := public.find_show_overlaps(new.booking_date, new.start_time, new.end_time, new.setup_time, new.studio);
    if cardinality(shows) > 0 then
      new.notes := public.stamp_note(new.notes, 'Overlaps regular show ' || array_to_string(shows, ', ')
                                                || ' ' || chr(8212) || ' engineer to decide');
    end if;
  end if;

  return new;
end $$;

-- When a booking stops blocking its old slot, re-check the Conflict rows there.
create or replace function public.bookings_after() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  if pg_trigger_depth() > 1 then
    return null;
  end if;
  if (new.status in ('Declined','Cancelled') and old.status is distinct from new.status)
     or (old.booking_date, old.studio, old.start_time, old.end_time, old.setup_time)
        is distinct from (new.booking_date, new.studio, new.start_time, new.end_time, new.setup_time)
  then
    update public.bookings
       set updated_at = now()
     where status = 'Conflict'
       and id <> new.id
       and booking_date = old.booking_date
       and studio = old.studio;
  end if;
  return null;
end $$;

-- Email hook: tells the "notify" edge function which booking changed, via pg_net.
-- Only the id is sent; the function reads the booking itself, so a forged call
-- cannot invent content. URL comes from app_settings.notify_url and the
-- project's public anon key from app_settings.notify_key, which passes the
-- function's JWT check (see README). Never blocks a save.
create or replace function public.bookings_notify() returns trigger
language plpgsql security definer set search_path = public as $$
declare
  v_url text;
  v_key text;
begin
  if tg_op = 'UPDATE' and new.status is not distinct from old.status then
    return null;
  end if;
  begin
    select value into v_url from public.app_settings where key = 'notify_url';
    select value into v_key from public.app_settings where key = 'notify_key';
    if coalesce(v_url, '') <> '' then
      perform net.http_post(
        url     := v_url,
        headers := jsonb_build_object('Content-Type', 'application/json',
                                      'Authorization', 'Bearer ' || coalesce(v_key, ''),
                                      'apikey', coalesce(v_key, '')),
        body    := jsonb_build_object('id', new.id));
    end if;
  exception when others then
    raise warning 'ZTN notify skipped: %', sqlerrm;
  end;
  return null;
end $$;

-- Engineer addresses for new-request emails (edge function only).
create or replace function public.engineer_emails() returns setof text
language sql stable security definer set search_path = public as $$
  select u.email::text from public.engineers e join auth.users u on u.id = e.user_id
   where u.email is not null order by 1
$$;

-- Ended weeks are read-only. Applies to everything that comes through the
-- website or API (session user "authenticator"); the hourly job and the SQL
-- editor can still correct records. Runs before bookings_before (alphabetical).
create or replace function public.bookings_lock_past() returns trigger
language plpgsql set search_path = public as $$
declare
  v_open date := public.ztn_week_open();
begin
  if session_user <> 'authenticator' then
    return case when tg_op = 'DELETE' then old else new end;
  end if;
  if tg_op = 'UPDATE' then
    -- The email sender only records what it has sent.
    if (to_jsonb(new) - 'notified_status') = (to_jsonb(old) - 'notified_status') then return new; end if;
    -- Engineers' "close completed" button finishing last week's bookings.
    if old.status = 'Confirmed' and new.status = 'Completed'
       and current_setting('ztn.actor', true) = 'Automatic (hourly auto-complete)' then
      return new;
    end if;
  end if;
  if (tg_op <> 'INSERT' and old.booking_date < v_open) or (tg_op <> 'DELETE' and new.booking_date < v_open) then
    raise exception 'That week has ended and its schedule is locked';
  end if;
  return case when tg_op = 'DELETE' then old else new end;
end $$;

drop trigger if exists bookings_a_lock_past on public.bookings;
create trigger bookings_a_lock_past before insert or update or delete on public.bookings
  for each row execute function public.bookings_lock_past();

drop trigger if exists bookings_before on public.bookings;
create trigger bookings_before before insert or update on public.bookings
  for each row execute function public.bookings_before();

drop trigger if exists bookings_after on public.bookings;
create trigger bookings_after after update on public.bookings
  for each row execute function public.bookings_after();

drop trigger if exists bookings_notify on public.bookings;
create trigger bookings_notify after insert or update of status on public.bookings
  for each row execute function public.bookings_notify();

-- Instant sync: broadcast "changed" on the public Realtime topic "ztn-schedule"
-- whenever bookings or weekly shows change. The message carries no data; the
-- monitor just reloads its week. Never blocks a save.
create or replace function public.schedule_changed() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  begin
    perform realtime.send(jsonb_build_object('table', tg_table_name), 'changed', 'ztn-schedule', false);
  exception when others then
    raise warning 'ZTN sync signal skipped: %', sqlerrm;
  end;
  return null;
end $$;

drop trigger if exists bookings_changed on public.bookings;
create trigger bookings_changed after insert or update or delete on public.bookings
  for each statement execute function public.schedule_changed();

drop trigger if exists weekly_shows_changed on public.weekly_shows;
create trigger weekly_shows_changed after insert or update or delete on public.weekly_shows
  for each statement execute function public.schedule_changed();

-- ------------------------------------------------------------ public RPCs --
-- Booking form submission (anyone with the site link).
create or replace function public.submit_booking(p jsonb) returns jsonb
language plpgsql security definer set search_path = public as $$
declare
  b        public.bookings;
  v_today  date := (now() at time zone public.ztn_tz())::date;
  v_recent int;
begin
  perform set_config('ztn.actor', 'Public booking form', true);
  if coalesce(trim(p->>'requester_name'), '') = '' then raise exception 'Requester name is required'; end if;
  if coalesce(trim(p->>'requester_email'), '') = '' then raise exception 'Requester email is required'; end if;
  if coalesce(trim(p->>'title'), '') = '' then raise exception 'Programme / production title is required'; end if;
  if coalesce(p->>'booking_date', '') = '' then raise exception 'Requested date is required'; end if;
  if coalesce(p->>'start_time', '') = '' or coalesce(p->>'end_time', '') = '' then
    raise exception 'Start and end time are required';
  end if;
  if (p->>'booking_date')::date < v_today then raise exception 'Requested date is in the past'; end if;
  if (p->>'end_time')::time <= (p->>'start_time')::time then
    raise exception 'End time must be later than start time';
  end if;

  select count(*) into v_recent
    from public.bookings
   where lower(requester_email) = lower(trim(p->>'requester_email'))
     and created_at > now() - interval '1 day';
  if v_recent >= 20 then raise exception 'Too many requests from this email today'; end if;

  insert into public.bookings (
    booking_date, start_time, end_time, setup_time, studio, title,
    requester_name, department, requester_email, contact_number, producer,
    programme_type, live_recorded, crew, priority, details)
  values (
    (p->>'booking_date')::date,
    (p->>'start_time')::time,
    (p->>'end_time')::time,
    nullif(p->>'setup_time', '')::time,
    p->>'studio',
    trim(p->>'title'),
    trim(p->>'requester_name'),
    nullif(trim(p->>'department'), ''),
    trim(p->>'requester_email'),
    nullif(trim(p->>'contact_number'), ''),
    nullif(trim(p->>'producer'), ''),
    nullif(trim(p->>'programme_type'), ''),
    nullif(p->>'live_recorded', ''),
    nullif(trim(p->>'crew'), ''),
    nullif(p->>'priority', ''),
    coalesce(p->'details', '{}'::jsonb))
  returning * into b;

  return jsonb_build_object('ref', b.ref, 'status', b.status, 'notes', b.notes,
                            'show_overlaps', to_jsonb(public.find_show_overlaps(
                              b.booking_date, b.start_time, b.end_time, b.setup_time, b.studio)));
end $$;

-- Busy times for one studio/day (availability hint on the form). Bookings show
-- no titles or names; regular shows are public programming, so they are named.
drop function if exists public.day_busy(date, text);
create or replace function public.day_busy(p_date date, p_studio text)
returns table (start_time time, end_time time, status text, label text)
language sql stable security definer set search_path = public as $$
  select * from (
    select least(coalesce(b.setup_time, b.start_time), b.start_time), b.end_time,
           case when b.status = any (public.ztn_firm()) then 'Confirmed' else 'Pending' end, null::text
      from public.bookings b
     where b.booking_date = p_date and b.studio = p_studio
       and b.status = any (public.ztn_blocking())
    union all
    select s.start_time, s.end_time, 'Show', s.name
      from public.weekly_shows s
     where s.studio = p_studio
       and extract(isodow from p_date)::smallint = any (s.days)
       and (s.active_from is null or p_date >= s.active_from)
       and (s.active_to is null or p_date <= s.active_to)
  ) x
  order by 1, 2
$$;

-- Studio monitor feed. If app_settings.monitor_key is set, the key must match.
-- kind = 'show' (regular weekly show) or 'booking' (Confirmed / Completed).
drop function if exists public.week_overview(date, text);
create or replace function public.week_overview(p_start date, p_key text default null)
returns table (kind text, ref text, booking_date date, start_time time, end_time time,
               studio text, title text, status text)
language plpgsql stable security definer set search_path = public as $$
declare
  v_key text;
  v_arch jsonb;
begin
  select value into v_key from public.app_settings where key = 'monitor_key';
  if coalesce(v_key, '') <> '' and coalesce(p_key, '') <> v_key and not public.is_engineer() then
    raise exception 'Invalid monitor key';
  end if;
  -- A week that has ended shows exactly what was archived, so later edits to
  -- weekly shows can't change it. Weeks without a snapshot use live data.
  select a.data into v_arch from public.weekly_archives a
   where a.week_start = p_start and p_start < public.ztn_week_open();
  if v_arch is not null then
    return query
      select x.* from (
        select 'show'::text, null::text, (j->>'date')::date, (j->>'start')::time, (j->>'end')::time,
               j->>'studio', j->>'title', 'Show'::text
          from jsonb_array_elements(v_arch->'shows') j
        union all
        select 'booking'::text, j->>'ref', (j->>'date')::date, (j->>'start')::time, (j->>'end')::time,
               j->>'studio', j->>'title', j->>'status'
          from jsonb_array_elements(v_arch->'bookings') j
         where j->>'status' = any (public.ztn_firm())
      ) x
      order by 3, 4, 1 desc;
    return;
  end if;
  return query
    select x.* from (
      select 'show'::text as kind, null::text as ref, d::date as booking_date, s.start_time, s.end_time,
             s.studio, s.name as title, 'Show'::text as status
        from public.weekly_shows s
        cross join generate_series(p_start, p_start + 6, interval '1 day') d
       where extract(isodow from d)::smallint = any (s.days)
         and (s.active_from is null or d::date >= s.active_from)
         and (s.active_to is null or d::date <= s.active_to)
      union all
      select 'booking'::text, b.ref, b.booking_date, b.start_time, b.end_time, b.studio, b.title, b.status
        from public.bookings b
       where b.status = any (public.ztn_firm())
         and b.booking_date between p_start and p_start + 6
    ) x
    order by x.booking_date, x.start_time, x.kind desc;
end $$;

-- ----------------------------------------------------------- engineer RPCs --
create or replace function public.recheck_all_conflicts() returns int
language plpgsql security definer set search_path = public as $$
declare
  n int;
begin
  if not public.is_engineer() then raise exception 'Engineers only'; end if;
  with before as (
    select id, status from public.bookings
     where status = any (public.ztn_live()) and booking_date >= public.ztn_week_open()
  ), upd as (
    update public.bookings b set updated_at = now()
      from before where b.id = before.id
    returning b.id, b.status
  )
  select count(*) into n from upd join before using (id) where upd.status <> before.status;
  return n;
end $$;

create or replace function public._close_completed() returns int
language plpgsql security definer set search_path = public as $$
declare
  n int;
begin
  perform set_config('ztn.actor', 'Automatic (hourly auto-complete)', true);
  update public.bookings
     set status = 'Completed',
         notes = public.stamp_note(notes, 'Auto-completed')
   where status = 'Confirmed'
     and (booking_date + end_time) < (now() at time zone public.ztn_tz());
  get diagnostics n = row_count;
  return n;
end $$;

create or replace function public.close_completed() returns int
language plpgsql security definer set search_path = public as $$
begin
  if not public.is_engineer() then raise exception 'Engineers only'; end if;
  return public._close_completed();
end $$;

-- ----------------------------------------------------------- weekly archive --
-- One compact snapshot per week (Monday-Sunday): the regular shows that ran and
-- every booking in that week, whatever its status. Saved automatically every
-- Sunday night (cron.sql) and on demand by engineers. Read-only from the site;
-- sha256 is a fingerprint of the stored data so it can be shown to be unaltered.
create table if not exists public.weekly_archives (
  week_start    date primary key check (extract(isodow from week_start) = 1),
  archived_at   timestamptz not null default now(),
  archived_by   text not null default 'automatic',
  booking_count int not null default 0,
  show_count    int not null default 0,
  data          jsonb not null,
  sha256        text not null
);

create or replace function public._archive_week(p_start date, p_by text default 'automatic')
returns public.weekly_archives
language plpgsql security definer set search_path = public as $$
declare
  v_start date := date_trunc('week', p_start)::date;
  v_shows jsonb;
  v_bookings jsonb;
  v_data jsonb;
  r public.weekly_archives;
begin
  perform set_config('ztn.actor', 'Automatic (weekly archive)', true);
  select coalesce(jsonb_agg(jsonb_build_object(
           'date', d::date, 'start', to_char(s.start_time, 'HH24:MI'), 'end', to_char(s.end_time, 'HH24:MI'),
           'studio', s.studio, 'title', s.name) order by d, s.start_time, s.name), '[]')
    into v_shows
    from public.weekly_shows s
    cross join generate_series(v_start, v_start + 6, interval '1 day') d
   where extract(isodow from d)::smallint = any (s.days)
     and (s.active_from is null or d::date >= s.active_from)
     and (s.active_to is null or d::date <= s.active_to);

  select coalesce(jsonb_agg(jsonb_strip_nulls(jsonb_build_object(
           'ref', b.ref, 'date', b.booking_date, 'start', to_char(b.start_time, 'HH24:MI'), 'end', to_char(b.end_time, 'HH24:MI'),
           'setup', to_char(b.setup_time, 'HH24:MI'), 'studio', b.studio, 'title', b.title, 'status', b.status,
           'requester', b.requester_name, 'department', b.department, 'email', b.requester_email,
           'contact', b.contact_number, 'producer', b.producer, 'live_recorded', b.live_recorded, 'crew', b.crew,
           'priority', b.priority, 'notes', nullif(b.notes, ''), 'submitted', b.created_at, 'approved', b.approved_at))
           order by b.booking_date, b.start_time, b.ref), '[]')
    into v_bookings
    from public.bookings b
   where b.booking_date between v_start and v_start + 6;

  v_data := jsonb_build_object('week_start', v_start, 'week_end', v_start + 6,
                               'shows', v_shows, 'bookings', v_bookings);

  insert into public.weekly_archives as a (week_start, archived_at, archived_by, booking_count, show_count, data, sha256)
  values (v_start, now(), coalesce(nullif(p_by, ''), 'automatic'), jsonb_array_length(v_bookings), jsonb_array_length(v_shows),
          v_data, encode(sha256(convert_to(v_data::text, 'UTF8')), 'hex'))
  on conflict (week_start) do update
     set archived_at = excluded.archived_at, archived_by = excluded.archived_by,
         booking_count = excluded.booking_count, show_count = excluded.show_count,
         data = excluded.data, sha256 = excluded.sha256
  returning * into r;
  return r;
end $$;

-- Engineers: save (or refresh) the archive for a week now.
create or replace function public.archive_week(p_start date) returns public.weekly_archives
language plpgsql security definer set search_path = public as $$
declare
  v_email text;
begin
  if not public.is_engineer() then raise exception 'Engineers only'; end if;
  if date_trunc('week', p_start)::date < public.ztn_week_open()
     and exists (select 1 from public.weekly_archives where week_start = date_trunc('week', p_start)::date) then
    raise exception 'That week has ended and its snapshot is locked';
  end if;
  select email into v_email from auth.users where id = auth.uid();
  return public._archive_week(p_start, coalesce(v_email, 'engineer'));
end $$;

-- -------------------------------------------------------------- audit trail --
-- Who did what. Database triggers record every change to bookings, weekly shows,
-- engineers, settings and archives, so nothing can skip the log. The page adds
-- activity that changes no data (sign-ins, downloads) through audit_event().
-- Engineers can read it; nobody can edit or delete it.
create table if not exists public.audit_log (
  id          bigint generated always as identity primary key,
  at          timestamptz not null default now(),
  actor_id    uuid,
  actor       text not null,
  action      text not null,
  target_type text not null,
  target      text,
  summary     text,
  changes     jsonb
);
create index if not exists audit_log_at_idx on public.audit_log (at desc);

create or replace function public.audit_row() returns trigger
language plpgsql security definer set search_path = public as $$
declare
  v_old     jsonb := case when tg_op <> 'INSERT' then to_jsonb(old) end;
  v_new     jsonb := case when tg_op <> 'DELETE' then to_jsonb(new) end;
  v_row     jsonb := coalesce(v_new, v_old);
  v_uid     uuid  := auth.uid();
  v_actor   text;
  v_action  text  := case tg_op when 'INSERT' then 'Created' when 'UPDATE' then 'Updated' else 'Deleted' end;
  v_type    text;
  v_target  text;
  v_summary text;
  v_changes jsonb;
begin
  if tg_op = 'UPDATE' then
    select jsonb_object_agg(e.key, jsonb_build_object('from', v_old -> e.key, 'to', e.value))
      into v_changes
      from jsonb_each(v_new) e
     where e.key not in ('updated_at', 'notified_status', 'approved_at', 'approved_by', 'data', 'sha256', 'archived_at')
       and e.value is distinct from v_old -> e.key;
    -- Bookkeeping only (e.g. the email function marking a status as sent): nothing to record.
    if v_changes is null and tg_table_name <> 'weekly_archives' then
      return null;
    end if;
  elsif tg_op = 'DELETE' then
    v_changes := v_old - 'data';   -- keep what was deleted
  end if;

  if v_uid is not null then
    select email into v_actor from auth.users where id = v_uid;
  end if;
  v_actor := coalesce(v_actor, nullif(current_setting('ztn.actor', true), ''), 'System / database admin');

  case tg_table_name
    when 'bookings' then
      v_type := 'Booking';
      v_target := (v_row ->> 'ref') || ' ' || chr(8212) || ' ' || (v_row ->> 'title');
      if tg_op = 'INSERT' then
        v_summary := (v_row ->> 'status') || ', ' || (v_row ->> 'booking_date') || ' ' || left(v_row ->> 'start_time', 5)
                     || chr(8211) || left(v_row ->> 'end_time', 5) || ', ' || (v_row ->> 'studio');
      elsif tg_op = 'UPDATE' and v_changes ? 'status' then
        v_action := 'Status changed';
        v_summary := (v_old ->> 'status') || ' ' || chr(8594) || ' ' || (v_new ->> 'status');
      end if;
      if pg_trigger_depth() > 1 then
        v_summary := concat_ws(' ', v_summary, '(automatic re-check)');
      end if;
    when 'weekly_shows' then
      v_type := 'Weekly show';
      v_target := v_row ->> 'name';
    when 'engineers' then
      v_type := 'Engineer';
      select email into v_target from auth.users where id = (v_row ->> 'user_id')::uuid;
      v_target := coalesce(v_target, v_row ->> 'full_name', v_row ->> 'user_id');
      v_action := case tg_op when 'INSERT' then 'Added' when 'DELETE' then 'Removed' else v_action end;
    when 'app_settings' then
      v_type := 'Setting';
      v_target := v_row ->> 'key';
      v_changes := null;           -- values can be private: record only that it changed
    when 'weekly_archives' then
      v_type := 'Archive';
      v_target := 'Week of ' || (v_row ->> 'week_start');
      v_action := 'Snapshot saved';
      v_summary := (v_row ->> 'booking_count') || ' booking(s), ' || (v_row ->> 'show_count') || ' show slot(s)';
      v_changes := null;
  end case;

  insert into public.audit_log (actor_id, actor, action, target_type, target, summary, changes)
  values (v_uid, v_actor, v_action, v_type, v_target, v_summary, v_changes);
  return null;
end $$;

drop trigger if exists bookings_audit on public.bookings;
create trigger bookings_audit after insert or update or delete on public.bookings
  for each row execute function public.audit_row();
drop trigger if exists weekly_shows_audit on public.weekly_shows;
create trigger weekly_shows_audit after insert or update or delete on public.weekly_shows
  for each row execute function public.audit_row();
drop trigger if exists engineers_audit on public.engineers;
create trigger engineers_audit after insert or update or delete on public.engineers
  for each row execute function public.audit_row();
drop trigger if exists app_settings_audit on public.app_settings;
create trigger app_settings_audit after insert or update or delete on public.app_settings
  for each row execute function public.audit_row();
drop trigger if exists weekly_archives_audit on public.weekly_archives;
create trigger weekly_archives_audit after insert or update on public.weekly_archives
  for each row execute function public.audit_row();

-- The trail itself cannot be rewritten, not even from the SQL Editor.
create or replace function public.audit_log_locked() returns trigger
language plpgsql as $$
begin
  raise exception 'The audit trail cannot be changed or deleted';
end $$;

drop trigger if exists audit_log_no_change on public.audit_log;
create trigger audit_log_no_change before update or delete on public.audit_log
  for each row execute function public.audit_log_locked();
drop trigger if exists audit_log_no_truncate on public.audit_log;
create trigger audit_log_no_truncate before truncate on public.audit_log
  for each statement execute function public.audit_log_locked();

-- Activity that changes no data, reported by the engineer page.
create or replace function public.audit_event(p_action text, p_detail text default null) returns void
language plpgsql security definer set search_path = public as $$
declare
  v_email text;
begin
  if not public.is_engineer() then raise exception 'Engineers only'; end if;
  if p_action not in ('Signed in', 'Signed out', 'Password set', 'Downloaded', 'Printed') then
    raise exception 'Unknown activity';
  end if;
  select email into v_email from auth.users where id = auth.uid();
  insert into public.audit_log (actor_id, actor, action, target_type, target)
  values (auth.uid(), coalesce(v_email, 'engineer'), p_action, 'Activity', left(p_detail, 300));
end $$;

-- ------------------------------------------------------- row level security --
alter table public.bookings enable row level security;
alter table public.engineers enable row level security;
alter table public.app_settings enable row level security;

drop policy if exists "engineers select bookings" on public.bookings;
drop policy if exists "engineers insert bookings" on public.bookings;
drop policy if exists "engineers update bookings" on public.bookings;
drop policy if exists "engineers delete bookings" on public.bookings;
create policy "engineers select bookings" on public.bookings for select to authenticated using (public.is_engineer());
create policy "engineers insert bookings" on public.bookings for insert to authenticated with check (public.is_engineer());
create policy "engineers update bookings" on public.bookings for update to authenticated
  using (public.is_engineer()) with check (public.is_engineer());
create policy "engineers delete bookings" on public.bookings for delete to authenticated using (public.is_engineer());

alter table public.weekly_shows enable row level security;
drop policy if exists "engineers manage shows" on public.weekly_shows;
create policy "engineers manage shows" on public.weekly_shows for all to authenticated
  using (public.is_engineer()) with check (public.is_engineer());

-- Audit trail: engineers can read; only the database writes to it.
alter table public.audit_log enable row level security;
drop policy if exists "engineers read audit log" on public.audit_log;
create policy "engineers read audit log" on public.audit_log for select to authenticated using (public.is_engineer());
revoke insert, update, delete, truncate on public.audit_log from anon, authenticated;

-- Archives: engineers can read; nobody can change or delete them from the site.
alter table public.weekly_archives enable row level security;
drop policy if exists "engineers read archives" on public.weekly_archives;
create policy "engineers read archives" on public.weekly_archives for select to authenticated using (public.is_engineer());
revoke insert, update, delete on public.weekly_archives from anon, authenticated;

drop policy if exists "engineers select engineers" on public.engineers;
create policy "engineers select engineers" on public.engineers for select to authenticated using (public.is_engineer());

drop policy if exists "engineers manage settings" on public.app_settings;
create policy "engineers manage settings" on public.app_settings for all to authenticated
  using (public.is_engineer()) with check (public.is_engineer());

-- --------------------------------------------------------------- privileges --
revoke all on function public.find_conflicts(date, time, time, time, text, bigint, boolean) from public, anon, authenticated;
revoke all on function public.next_booking_ref() from public, anon, authenticated;
revoke all on function public.find_show_overlaps(date, time, time, time, text) from public, anon, authenticated;
revoke all on function public._close_completed() from public, anon, authenticated;
revoke all on function public.bookings_before() from public, anon, authenticated;
revoke all on function public.bookings_lock_past() from public, anon, authenticated;
revoke all on function public.bookings_after() from public, anon, authenticated;
revoke all on function public.bookings_notify() from public, anon, authenticated;
revoke all on function public.engineer_emails() from public, anon, authenticated;
revoke all on function public.schedule_changed() from public, anon, authenticated;
revoke all on function public._archive_week(date, text) from public, anon, authenticated;
revoke all on function public.audit_row() from public, anon, authenticated;
revoke all on function public.audit_event(text, text) from public, anon;
revoke all on function public.archive_week(date) from public, anon;
do $$
begin
  if exists (select 1 from pg_roles where rolname = 'service_role') then
    grant execute on function public.engineer_emails() to service_role;
  end if;
end $$;

grant execute on function public.submit_booking(jsonb) to anon, authenticated;
grant execute on function public.day_busy(date, text) to anon, authenticated;
grant execute on function public.week_overview(date, text) to anon, authenticated;
grant execute on function public.is_engineer() to anon, authenticated;
grant execute on function public.recheck_all_conflicts() to authenticated;
grant execute on function public.close_completed() to authenticated;
grant execute on function public.archive_week(date) to authenticated;
grant execute on function public.audit_event(text, text) to authenticated;

-- Live updates in the engineer console
do $$
begin
  if exists (select 1 from pg_publication where pubname = 'supabase_realtime')
     and not exists (select 1 from pg_publication_tables
                      where pubname = 'supabase_realtime' and tablename = 'bookings') then
    alter publication supabase_realtime add table public.bookings;
  end if;
  if exists (select 1 from pg_publication where pubname = 'supabase_realtime')
     and not exists (select 1 from pg_publication_tables
                      where pubname = 'supabase_realtime' and tablename = 'weekly_shows') then
    alter publication supabase_realtime add table public.weekly_shows;
  end if;
end $$;
