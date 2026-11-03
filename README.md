# ZTN Studio Booking

A small website for booking ZTN studios. It has three pages:

| Page | Who uses it | What it does |
|---|---|---|
| `index.html` | Anyone with the link | Booking request form. Shows what is already booked in the chosen studio that day, including regular shows. |
| `engineer.html` | Studio engineers (sign-in) | Approve, decline, cancel, reschedule and edit bookings; manage the weekly shows; re-check conflicts; export CSV for Excel. |
| `monitor.html` | Studio TV | Full-screen weekly overview: regular shows in green, bookings in blue, red where the same studio is used twice. Refreshes every minute. |

The rules match the Excel workbook (`ZTN_Booking_System.bas`):

- Statuses: Requested, Pending Approval, Conflict, Confirmed, Declined, Rescheduled, Cancelled, Completed.
- A clash means the same date, the same studio and overlapping times. Setup access time counts as an earlier start.
- A clashing request is saved as **Conflict**. Confirming a booking that clashes is refused.
- **Regular weekly shows never block a booking.** The overlap is noted and the engineer decides. Confirming over a show is allowed and recorded in the notes.
- Every hour, Confirmed bookings that have finished become **Completed**.
- Emails: the engineer gets one for every new request. The requester gets one on Confirmed, Declined or Cancelled (with the reason), and Rescheduled.

The database enforces all of this, so nobody can get around it from the browser.

**Hosting:** the pages run on GitHub Pages (free). Data, sign-in and the rules live in Supabase (free tier). Email goes through Resend (free tier).

---

## Setup (about 30 minutes, one time)

### 1. Supabase project
1. Go to <https://supabase.com>, sign up and choose **New project**. Pick the region closest to Zimbabwe, which is usually *Frankfurt* or *London*. Save the database password somewhere safe.
2. Open **SQL Editor → New query**. Paste the whole of [`supabase/schema.sql`](supabase/schema.sql) and click **Run**. This creates the tables and rules, and loads your current weekly shows.
3. Open **Database → Extensions** and turn on **pg_cron**. Then run [`supabase/cron.sql`](supabase/cron.sql) in the SQL Editor. This sets up the hourly auto-complete.
4. Open **Authentication → Sign In / Providers → Email** and turn **off** "Allow new users to sign up". Only accounts you create can then sign in.

### 2. Engineer accounts
1. Open **Authentication → URL Configuration**. Set **Site URL** to `https://<your-site>/engineer.html` and add the same address under **Redirect URLs**. Invite and password-reset links then open the engineer page.
2. Open **Authentication → Users → Add user**, then either:
   - **Send invitation:** the engineer gets an email, clicks the link and chooses their own password on the engineer page. This needs email set up in step 5 (Resend as custom SMTP). Supabase's built-in mailer only delivers to members of your Supabase team.
   - **Create new user:** you set a password, tick *Auto confirm* and pass the password on.
3. Make that account an engineer by running this in the SQL Editor (change the email):
   ```sql
   insert into public.engineers (user_id, full_name)
   select id, 'Engineer Name' from auth.users where email = 'engineer@ztn.co.zw';
   ```
   Repeat for each engineer. To remove one, delete their row from `public.engineers`.

### 3. Connect the website
1. In Supabase, open **Project Settings → API**. Copy the **Project URL** and the **anon public** key.
2. Put them in [`assets/config.js`](assets/config.js). The anon key is safe to publish, because the database decides what each visitor can see.

### 4. Publish on GitHub Pages
1. Create a repository on GitHub and upload these files. The workbooks are not uploaded: `.gitignore` leaves them out.
2. Open **Settings → Pages → Build and deployment**. Choose *Deploy from a branch*, then `main` and `/ (root)`, and save.
3. After a minute the site is live at `https://<your-account>.github.io/<repo>/`.
   - Booking form: `…/index.html` (share this link)
   - Engineers: `…/engineer.html`
   - Monitor: `…/monitor.html`

**Optional: your own address, such as `booking.ztn.co.zw`.** In **Settings → Pages → Custom domain**, enter the address. Then ask whoever manages the ztn.co.zw DNS to add a `CNAME` record pointing `booking` to `<your-account>.github.io`.

### 5. Email (Resend)
1. Sign up at <https://resend.com>. Under **Domains**, add a sending subdomain such as `notify.ztn.co.zw` and add the DNS records it shows you. Until the domain is verified, Resend only delivers to your own address.
2. Create an **API key** with *Sending access*.
3. In Supabase, open **Edge Functions → Deploy a new function → Via editor**. Name it `notify` and paste [`supabase/functions/notify/index.ts`](supabase/functions/notify/index.ts). Deploy it and leave "Verify JWT" on: the database calls it with the public anon key (below).
4. Under **Edge Functions → Secrets**, add `RESEND_API_KEY` with the Resend key. That's the only secret: the function reads each booking itself, and new-request emails go to everyone on the engineers list.
5. Tell the database where the function is and who emails come from. Run this in the SQL Editor:
   ```sql
   insert into public.app_settings (key, value) values
     ('notify_url', 'https://<project-ref>.supabase.co/functions/v1/notify'),
     ('notify_key', '<the anon public key, same as in assets/config.js>'),
     ('from_email', 'ZTN Studio <studio@notify.ztn.co.zw>'),
     ('site_url',   'https://<your-site>')
   on conflict (key) do update set value = excluded.value;
   ```
   If email isn't set up, bookings still work. They just aren't emailed.
6. **Invite emails (engineer accounts):** in Supabase open **Authentication → Emails → SMTP Settings**, turn on custom SMTP and enter host `smtp.resend.com`, port `465`, username `resend`, the Resend API key as the password, and the same sender address.

### 6. Studio monitor
- **Recommended:** a mini PC or TV stick running Chrome or Edge in kiosk mode. It opens full-screen at startup and recovers on its own after a power cut:
  ```
  msedge --kiosk "https://<your-site>/monitor.html?key=<monitor key>" --edge-kiosk-type=fullscreen
  ```
  A smart TV browser also works.
- **Optional monitor key:** stops strangers from viewing the overview. Set it once:
  ```sql
  insert into public.app_settings values ('monitor_key', 'pick-a-long-random-word')
  on conflict (key) do update set value = excluded.value;
  ```
  Then add `?key=pick-a-long-random-word` to the monitor address. Engineers who are signed in don't need the key.
- Other options: `?week=2026-11-02` fixes the week shown, and `?refresh=30` changes the refresh time in seconds.

---

## Day to day

- **Weekly shows:** open the Engineers page, then the **Weekly shows** tab, to add or edit them. When a programme ends, set **Active to** rather than deleting it.
- **Approvals:** change the Status drop-down in the Schedule tab. Declining or cancelling asks for a reason, which is emailed to the requester.
- **Excel:** **Export CSV** downloads the schedule with the same columns as the workbook's Studio Schedule sheet.

## Costs and limits

- **GitHub Pages:** free for a public repository. This is fine because no secrets are stored in it.
- **Supabase free tier:** 500 MB of database storage. It pauses after a week with no activity, but the monitor's one-minute refresh keeps it active. The Pro plan ($25/month) adds daily backups and never pauses.
- **Resend free tier:** 3,000 emails a month, 100 a day.

## Files

```
index.html, engineer.html, monitor.html   the three pages
assets/app.js       shared code (week grid, dates, helpers)
assets/config.js    your Supabase URL + anon key
assets/style.css    styling (light/dark, monitor layout)
supabase/schema.sql tables, rules, security - safe to re-run
supabase/cron.sql   hourly auto-complete job
supabase/functions/notify/index.ts   email sender
ZTN_Booking_System.bas               the Excel version (VBA module)
```
