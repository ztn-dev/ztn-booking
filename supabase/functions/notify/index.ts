// ZTN Studio Booking - email notifications (Supabase Edge Function).
// The bookings_notify trigger posts {"id": <booking id>} here through pg_net.
// The function reads the booking itself with the service role, so the only
// thing a caller controls is which booking id to look at. Each status is
// emailed once: notified_status is claimed atomically before sending.
//
// Function secret (Dashboard -> Edge Functions -> Secrets):
//   RESEND_API_KEY   Resend API key
// Settings (public.app_settings, set in the SQL Editor):
//   from_email       e.g. "ZTN Studio <studio@notify.ztn.co.zw>" (domain verified in Resend)
//   site_url         e.g. https://ztn-dev.github.io/ztn-booking (adds links)
// Engineers' addresses come from public.engineers, so new engineers are
// included automatically.
// SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY are provided by Supabase.

type Booking = {
  id: number;
  ref: string;
  booking_date: string;
  start_time: string;
  end_time: string;
  setup_time: string | null;
  studio: string;
  title: string;
  requester_name: string;
  department: string | null;
  requester_email: string;
  contact_number: string | null;
  producer: string | null;
  live_recorded: string | null;
  crew: string | null;
  priority: string | null;
  status: string;
  notes: string;
  notified_status: string | null;
};

type Message = { to: string[]; subject: string; text: string };

const env = (k: string) => Deno.env.get(k) ?? "";
const DASH = "—";
const API = `${env("SUPABASE_URL")}/rest/v1`;
const DB_HEADERS = {
  apikey: env("SUPABASE_SERVICE_ROLE_KEY"),
  Authorization: `Bearer ${env("SUPABASE_SERVICE_ROLE_KEY")}`,
  "Content-Type": "application/json",
};

const hm = (t: string | null) => (t ? t.slice(0, 5) : "");

function longDate(iso: string): string {
  const [y, m, d] = iso.split("-").map(Number);
  return new Date(Date.UTC(y, m - 1, d)).toLocaleDateString("en-GB", {
    weekday: "long", day: "2-digit", month: "long", year: "numeric", timeZone: "UTC",
  });
}

function summary(b: Booking): string {
  const lines = [
    ["Reference", b.ref],
    ["Date / time", `${longDate(b.booking_date)}, ${hm(b.start_time)} - ${hm(b.end_time)}`],
    ["Setup access", hm(b.setup_time)],
    ["Studio", b.studio],
    ["Programme", b.title],
    ["Requester", [b.requester_name, b.department, b.requester_email, b.contact_number].filter(Boolean).join(` ${DASH} `)],
    ["Producer", b.producer ?? ""],
    ["Live/Rec", b.live_recorded ?? ""],
    ["Crew", b.crew ?? ""],
    ["Priority", b.priority ?? ""],
    ["Status", b.status],
    ["Notes", b.notes],
  ];
  return lines.filter(([, v]) => v).map(([k, v]) => `${(k + ":").padEnd(14)}${v}`).join("\n");
}

async function db(path: string, init: RequestInit = {}) {
  const res = await fetch(`${API}/${path}`, { ...init, headers: { ...DB_HEADERS, ...(init.headers ?? {}) } });
  if (!res.ok) throw new Error(`${path}: ${res.status} ${await res.text()}`);
  return res.status === 204 ? null : res.json();
}

function buildMessage(b: Booking, isNew: boolean, engineers: string[], site: string): Message | null {
  const tag = `${DASH} ${b.ref} ${DASH} ${b.title}`;
  const sign = "\n\nZTN Studio / Technical Department";

  if (isNew) {
    if (!engineers.length) return null;
    const conflict = b.status === "Conflict";
    const show = /Overlaps regular show/.test(b.notes)
      ? "NOTE: this time overlaps a regular weekly show. Approve or decline at your discretion.\n\n"
      : "";
    return {
      to: engineers,
      subject: conflict ? `[ZTN Studio] CONFLICT ${tag}` : `[ZTN Studio] New booking request ${tag}`,
      text: (conflict
        ? "A new booking request CLASHES with an existing booking.\n\n"
        : "A new studio booking request has been submitted.\n\n") + show +
        summary(b) + (site ? `\n\nReview: ${site}/engineer.html` : ""),
    };
  }

  const to = [b.requester_email];
  switch (b.status) {
    case "Confirmed":
      return {
        to,
        subject: `[ZTN Studio] Booking confirmed ${tag}`,
        text: "Your studio booking has been CONFIRMED.\n\n" + summary(b) + sign,
      };
    case "Declined":
    case "Cancelled":
      return {
        to,
        subject: `[ZTN Studio] Booking ${b.status.toLowerCase()} ${tag}`,
        text: `Your studio booking has been ${b.status.toUpperCase()}.\n\nReason / notes:\n${b.notes || "(none given)"}\n\n` +
          summary(b) + sign,
      };
    case "Rescheduled":
      return {
        to,
        subject: `[ZTN Studio] Booking rescheduled ${tag}`,
        text: "Your studio booking has been rescheduled. Current details:\n\n" + summary(b) + sign,
      };
    default:
      return null;
  }
}

Deno.serve(async (req) => {
  if (req.method !== "POST") return new Response("Method not allowed", { status: 405 });

  let id: number;
  try {
    id = Number((await req.json()).id);
  } catch {
    return new Response("Bad JSON", { status: 400 });
  }
  if (!Number.isInteger(id) || id <= 0) return new Response("Bad id", { status: 400 });

  try {
    const [b] = (await db(`bookings?id=eq.${id}&select=*`)) as Booking[];
    if (!b || b.notified_status === b.status) return new Response("skipped", { status: 200 });

    // Claim this status so a repeated call cannot send the same email twice.
    const prev = b.notified_status;
    const claimed = await db(
      `bookings?id=eq.${id}&notified_status=${prev === null ? "is.null" : `eq.${encodeURIComponent(prev)}`}`,
      { method: "PATCH", headers: { Prefer: "return=representation" }, body: JSON.stringify({ notified_status: b.status }) },
    ) as Booking[];
    if (!claimed.length) return new Response("already handled", { status: 200 });

    const settings = Object.fromEntries(
      ((await db("app_settings?key=in.(from_email,site_url)&select=key,value")) as { key: string; value: string }[])
        .map((s) => [s.key, s.value]),
    );
    const engineers = prev === null ? ((await db("rpc/engineer_emails", { method: "POST", body: "{}" })) as string[]) : [];
    const msg = buildMessage(b, prev === null, engineers, (settings.site_url ?? "").replace(/\/$/, ""));
    if (!msg) return new Response("no email for this change", { status: 200 });
    if (!env("RESEND_API_KEY")) return new Response("RESEND_API_KEY not set", { status: 200 });

    const res = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: { Authorization: `Bearer ${env("RESEND_API_KEY")}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        from: settings.from_email || "ZTN Studio <onboarding@resend.dev>",
        to: msg.to,
        subject: msg.subject,
        text: msg.text,
      }),
    });
    return new Response(await res.text(), { status: res.ok ? 200 : 502 });
  } catch (e) {
    console.error(e);
    return new Response("error", { status: 500 });
  }
});
