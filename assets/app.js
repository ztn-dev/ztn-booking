// Shared helpers for the booking form, engineer console and studio monitor.
import { createClient } from 'https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2.45.4/+esm';
import { SUPABASE_URL, SUPABASE_ANON_KEY, TIMEZONE } from './config.js';

export const STATUSES = ['Requested', 'Pending Approval', 'Conflict', 'Confirmed',
  'Declined', 'Rescheduled', 'Cancelled', 'Completed'];
export const FIRM_STATUSES = ['Confirmed', 'Completed'];
export const STUDIOS = ['Studio 1', 'Studio 2', 'Audio Studio', 'MCR', 'Other'];
export const PRIORITIES = ['High', 'Urgent'];
export const LIVE_RECORDED = ['Live', 'Recorded', 'Live-to-tape'];
// Minutes since midnight, matching the workbook: 06:00-08:00, half-hours to 19:30, then 19:30-22:00.
export const SLOT_BOUNDS = [360, ...Array.from({ length: 24 }, (_, i) => 480 + i * 30), 1320];
export const DAY_NAMES = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
export const DAY_SHORT = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];

// [1,2,3,4,5] -> "Mon-Fri", [2,4] -> "Tue, Thu" (ISO weekdays, 1 = Monday).
export function daysLabel(days) {
  const d = [...new Set(days || [])].sort((a, b) => a - b);
  if (d.length === 7) return 'Daily';
  const runs = [];
  for (const n of d) {
    const last = runs[runs.length - 1];
    if (last && n === last[1] + 1) last[1] = n; else runs.push([n, n]);
  }
  return runs.map(([a, b]) => (b - a >= 2 ? `${DAY_SHORT[a - 1]}-${DAY_SHORT[b - 1]}`
    : a === b ? DAY_SHORT[a - 1] : `${DAY_SHORT[a - 1]}, ${DAY_SHORT[b - 1]}`)).join(', ');
}

// Read before the client consumes the link: invite / password-reset emails land
// here with #...&type=invite or type=recovery.
export const AUTH_LINK_TYPE = new URLSearchParams(location.hash.slice(1)).get('type');

export const isConfigured = !/YOUR-PROJECT|YOUR-ANON/.test(SUPABASE_URL + SUPABASE_ANON_KEY);
export const db = isConfigured ? createClient(SUPABASE_URL, SUPABASE_ANON_KEY) : null;

const pad = (n) => String(n).padStart(2, '0');
export const mm = (min) => `${pad(Math.floor(min / 60))}:${pad(min % 60)}`;
export const toMin = (t) => {
  if (!t) return null;
  const [h, m] = String(t).split(':').map(Number);
  return h * 60 + (m || 0);
};
export const hm = (t) => (t ? String(t).slice(0, 5) : '');

// Date and minutes "now" in the studio's time zone.
export function nowParts() {
  const parts = Object.fromEntries(new Intl.DateTimeFormat('en-CA', {
    timeZone: TIMEZONE, year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23',
  }).formatToParts(new Date()).map((p) => [p.type, p.value]));
  return {
    date: `${parts.year}-${parts.month}-${parts.day}`,
    minutes: Number(parts.hour) * 60 + Number(parts.minute),
    time: `${parts.hour}:${parts.minute}:${parts.second}`,
  };
}
export const todayISO = () => nowParts().date;

export function addDays(iso, n) {
  const d = new Date(`${iso}T00:00:00Z`);
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
}
export function mondayOf(iso) {
  const dow = (new Date(`${iso}T00:00:00Z`).getUTCDay() + 6) % 7;
  return addDays(iso, -dow);
}
export function fmtDate(iso, opts = { day: '2-digit', month: 'short', year: 'numeric' }) {
  if (!iso) return '';
  return new Date(`${iso}T00:00:00Z`).toLocaleDateString('en-GB', { ...opts, timeZone: 'UTC' });
}
export const dayName = (iso) => fmtDate(iso, { weekday: 'long' });

export const esc = (s) => String(s ?? '').replace(/[&<>"']/g,
  (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
export const statusSlug = (s) => String(s || '').toLowerCase().replace(/\s+/g, '-');
export const statusPill = (s) => `<span class="pill s-${statusSlug(s)}">${esc(s)}</span>`;
export const errMsg = (e) => (e && (e.message || e.error_description)) || String(e);

export function showNotConfigured(container) {
  container.innerHTML = `<div class="notice warn"><strong>Not connected yet.</strong>
    Add your Supabase project URL and anon key to <code>assets/config.js</code> (see README).</div>`;
}

// Keeps a page that stays open up to date. Every few minutes it compares the files
// this page is running with the published ones. After a release it reloads as soon as
// nobody is mid-task: no box open, nothing being typed, and isBusy() false.
export function watchForUpdate({ every = 300000, isBusy = () => false } = {}) {
  const files = [location.href.split(/[?#]/)[0], import.meta.url,
    new URL('./config.js', import.meta.url).href, new URL('./style.css', import.meta.url).href];
  const read = (cache) => Promise.all(files.map((u) => fetch(u, { cache })
    .then((r) => (r.ok ? r.text() : Promise.reject(new Error(r.status)))))).then((t) => t.join(''));
  const busy = () => isBusy() || !!document.querySelector('dialog[open]')
    || !!document.activeElement?.matches('input, textarea, select');

  // force-cache gives the copies this page loaded; reload fetches the published
  // ones and stores them, so the reload below can't pick up old files.
  read('force-cache').then((running) => {
    let timer;
    const check = () => read('reload').then((latest) => {
      if (latest === running) return;
      clearInterval(timer);
      document.removeEventListener('visibilitychange', onShow);
      let idle = 0; // two quiet checks in a row, so a save that just started can finish
      setInterval(() => {
        idle = busy() ? 0 : idle + 1;
        if (idle < 2) return;
        try { sessionStorage.setItem('ztn-updated', '1'); } catch { /* the note is optional */ }
        location.reload();
      }, 10000);
    }, () => {}); // offline or a failed fetch: try again next time
    const onShow = () => { if (!document.hidden) check(); };
    timer = setInterval(check, every);
    document.addEventListener('visibilitychange', onShow);
    check();
  }, () => {});

  try {
    if (sessionStorage.getItem('ztn-updated')) {
      sessionStorage.removeItem('ztn-updated');
      toast('Page updated to the latest version.', 'ok');
    }
  } catch { /* the note is optional */ }
}

let toastHost;
export function toast(msg, kind = 'info', ms = 4500) {
  if (!toastHost) {
    toastHost = document.createElement('div');
    toastHost.className = 'toasts';
    toastHost.setAttribute('role', 'status');
    document.body.append(toastHost);
  }
  const t = document.createElement('div');
  t.className = `toast ${kind}`;
  t.textContent = msg;
  toastHost.append(t);
  setTimeout(() => t.remove(), ms);
}

// Weekly grid shared by the monitor and the engineer console, drawn like the
// workbook's Weekly Overview: regular shows green, bookings blue, and a cell
// red only when the same studio is used twice in it. Consecutive slots that
// hold exactly the same items are merged into a single tall cell. A cell lists
// what starts in it; only when nothing starts there does it name what carries on.
export function weekGridHTML(rows, weekStart, { now = null } = {}) {
  const slots = SLOT_BOUNDS.length - 1;
  const cells = Array.from({ length: slots }, () => Array.from({ length: 7 },
    () => ({ keys: [], starts: [], carries: [], studios: new Set(), clash: false, booked: false })));
  let seq = 0;
  for (const b of rows) {
    const isShow = b.kind === 'show';
    if (!isShow && !FIRM_STATUSES.includes(b.status)) continue;
    const day = Math.round((Date.parse(`${b.booking_date}T00:00:00Z`) - Date.parse(`${weekStart}T00:00:00Z`)) / 86400000);
    if (day < 0 || day > 6) continue;
    const s = toMin(b.start_time), e = toMin(b.end_time);
    if (!(e > s)) continue;
    const key = ++seq;
    const first = `<div class="bk ${b.status === 'Completed' ? 'done' : ''}"><div class="bk-title">${esc(b.title)}</div><div class="bk-time">${mm(s)}–${mm(e)} · ${esc(b.studio)}</div></div>`;
    const cont = `<div class="bk cont">${esc(b.title)}</div>`;
    let n = 0;
    for (let k = 0; k < slots; k++) {
      if (s < SLOT_BOUNDS[k + 1] && e > SLOT_BOUNDS[k]) {
        const c = cells[k][day];
        if (c.studios.has(b.studio)) c.clash = true;
        c.studios.add(b.studio);
        c.keys.push(key);
        if (n === 0) c.starts.push(first); else c.carries.push(cont);
        if (!isShow) c.booked = true;
        n++;
      }
    }
  }

  // rowspan[k][d]: 0 = covered by a merged cell above, otherwise the span.
  const span = Array.from({ length: slots }, () => Array(7).fill(1));
  for (let d = 0; d < 7; d++) {
    for (let k = 0; k < slots;) {
      let m = k;
      const same = cells[k][d].keys.join();
      if (same) while (m + 1 < slots && cells[m + 1][d].keys.join() === same) m++;
      span[k][d] = m - k + 1;
      for (let j = k + 1; j <= m; j++) span[j][d] = 0;
      k = m + 1;
    }
  }

  const todayIdx = now ? Math.round((Date.parse(`${now.date}T00:00:00Z`) - Date.parse(`${weekStart}T00:00:00Z`)) / 86400000) : -1;
  let html = '<table class="week"><thead><tr><th class="slot-h">Time</th>';
  DAY_NAMES.forEach((d, i) => {
    html += `<th class="${i === todayIdx ? 'today' : ''}">${d}<small>${fmtDate(addDays(weekStart, i), { day: '2-digit', month: 'short' })}</small></th>`;
  });
  html += '</tr></thead><tbody>';
  for (let k = 0; k < slots; k++) {
    const nowRow = now && now.minutes >= SLOT_BOUNDS[k] && now.minutes < SLOT_BOUNDS[k + 1];
    const tall = SLOT_BOUNDS[k + 1] - SLOT_BOUNDS[k] > 30 ? ' tall' : '';
    html += `<tr class="${nowRow ? 'now' : ''}${tall}"><th class="slot">${mm(SLOT_BOUNDS[k])}–${mm(SLOT_BOUNDS[k + 1])}</th>`;
    for (let d = 0; d < 7; d++) {
      if (!span[k][d]) continue;
      const c = cells[k][d];
      const fill = !c.keys.length ? '' : c.clash ? 'clash' : c.booked ? 'busy' : 'show';
      const cls = [d === todayIdx ? 'today' : '', fill].join(' ').trim();
      const shown = c.starts.length ? c.starts : c.carries;
      html += `<td class="${cls}" data-day="${d}" data-slot="${k}"${span[k][d] > 1 ? ` rowspan="${span[k][d]}"` : ''}>${shown.join('<div class="sep">— — —</div>')}</td>`;
    }
    html += '</tr>';
  }
  return html + '</tbody></table>';
}
