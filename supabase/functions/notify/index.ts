// ZTN Studio Booking - email notifications (Supabase Edge Function).
// The bookings_notify trigger posts {"id": <booking id>} here through pg_net.
// The function reads the booking itself with the service role, so the only
// thing a caller controls is which booking id to look at. Each status is
// emailed once: notified_status is claimed atomically before sending.
//
// Mail goes out through one of two senders:
//   SMTP (e.g. Gmail)  used when the SMTP_PASSWORD secret is set
//   Resend             used otherwise, with the RESEND_API_KEY secret
// Function secrets (Dashboard -> Edge Functions -> Secrets):
//   SMTP_PASSWORD    e.g. a Gmail App Password
//   RESEND_API_KEY   Resend API key
// Settings (public.app_settings, set in the SQL Editor):
//   from_email       e.g. "ZTN Studio <ztnget@gmail.com>"
//   site_url         e.g. https://ztn-dev.github.io/ztn-booking (adds links)
//   smtp_host        e.g. smtp.gmail.com   (port 465, TLS; port 587 is blocked on Supabase)
//   smtp_user        e.g. ztnget@gmail.com
// Engineers' addresses come from public.engineers, so new engineers are
// included automatically.
// SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY are provided by Supabase.

import { SMTPClient } from "https://deno.land/x/denomailer@1.6.0/mod.ts";

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

// Booking facts as "Label: value" lines. Engineers also get the working details
// and internal notes; requesters only get what concerns them.
function details(b: Booking, forEngineer: boolean): string {
  const person = [...new Set([b.requester_name, b.department, b.requester_email, b.contact_number].filter(Boolean))];
  const lines = [
    ["Reference", b.ref],
    ["Programme", b.title],
    ["Date", longDate(b.booking_date)],
    ["Time", `${hm(b.start_time)} to ${hm(b.end_time)}`],
    ["Setup access from", hm(b.setup_time)],
    ["Studio", b.studio],
    ...(forEngineer ? [
      ["Requested by", person.join(", ")],
      ["Producer", b.producer ?? ""],
      ["Live or recorded", b.live_recorded ?? ""],
      ["Crew", b.crew ?? ""],
      ["Priority", b.priority ?? ""],
      ["Status", b.status],
      ["Notes", b.notes],
    ] : []),
  ];
  return lines.filter(([, v]) => v).map(([k, v]) => `${k}: ${v}`).join("\n");
}

// The reason an engineer typed when declining or cancelling.
function reasonFrom(notes: string): string {
  const found = notes.split("\n").map((l) => l.match(/^(?:\[[^\]]*\]\s*)?Reason:\s*(.+)$/)?.[1]).filter(Boolean);
  return found[found.length - 1] ?? "";
}

// Hello line; quick-added bookings use an email address as the name.
const hello = (b: Booking) => (b.requester_name && !b.requester_name.includes("@") ? `Hello ${b.requester_name},` : "Hello,");

// Subjects stay plain ASCII: the SMTP library wraps encoded subjects badly, which
// ends the headers early and shows the raw email to the reader.
const plainSubject = (s: string) => s.replace(/[\u2013\u2014]/g, "-").replace(/[^\x20-\x7E]/g, "").replace(/\s+/g, " ").trim();

async function db(path: string, init: RequestInit = {}) {
  const res = await fetch(`${API}/${path}`, { ...init, headers: { ...DB_HEADERS, ...(init.headers ?? {}) } });
  if (!res.ok) throw new Error(`${path}: ${res.status} ${await res.text()}`);
  return res.status === 204 ? null : res.json();
}

function buildMessage(b: Booking, isNew: boolean, engineers: string[], site: string): Message | null {
  const tag = `${b.ref} - ${b.title}`;
  const sign = "\n\nIf you have any questions, please contact the studio team.\n\nRegards,\nZTN Studio / Technical Department";
  const toRequester = (subject: string, opening: string, extra = "") => ({
    to: [b.requester_email],
    subject,
    text: `${hello(b)}\n\n${opening}\n\n${extra}Booking details\n${details(b, false)}${sign}`,
  });

  if (isNew) {
    if (!engineers.length) return null;
    const conflict = b.status === "Conflict";
    const show = /Overlaps regular show/.test(b.notes)
      ? "This time overlaps a regular weekly show. Approve or decline it at your discretion.\n\n"
      : "";
    return {
      to: engineers,
      subject: conflict ? `[ZTN Studio] CONFLICT: ${tag}` : `[ZTN Studio] New booking request: ${tag}`,
      text: "Hello,\n\n" + (conflict
        ? "A new booking request clashes with an existing booking.\n\n"
        : "A new studio booking request has been submitted.\n\n") + show +
        `Booking details\n${details(b, true)}` + (site ? `\n\nReview it here: ${site}/engineer.html` : ""),
    };
  }

  switch (b.status) {
    case "Confirmed":
      return toRequester(`[ZTN Studio] Booking confirmed: ${tag}`, "Your studio booking has been confirmed.");
    case "Declined":
    case "Cancelled": {
      const word = b.status.toLowerCase();
      const reason = reasonFrom(b.notes);
      return toRequester(`[ZTN Studio] Booking ${word}: ${tag}`, `Your studio booking has been ${word}.`,
        reason ? `Reason: ${reason}\n\n` : "");
    }
    case "Rescheduled":
      return toRequester(`[ZTN Studio] Booking rescheduled: ${tag}`, "Your studio booking has been rescheduled. The current details are below.");
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
      ((await db("app_settings?key=in.(from_email,site_url,smtp_host,smtp_user)&select=key,value")) as { key: string; value: string }[])
        .map((s) => [s.key, s.value]),
    );
    const engineers = prev === null ? ((await db("rpc/engineer_emails", { method: "POST", body: "{}" })) as string[]) : [];
    const msg = buildMessage(b, prev === null, engineers, (settings.site_url ?? "").replace(/\/$/, ""));
    if (!msg) return new Response("no email for this change", { status: 200 });

    if (env("SMTP_PASSWORD") && settings.smtp_host && settings.smtp_user) {
      const client = new SMTPClient({
        connection: {
          hostname: settings.smtp_host,
          port: 465,
          tls: true,
          auth: { username: settings.smtp_user, password: env("SMTP_PASSWORD") },
        },
      });
      try {
        await client.send({
          from: settings.from_email || settings.smtp_user,
          to: msg.to,
          subject: plainSubject(msg.subject),
          content: msg.text,
        });
      } finally {
        await client.close();
      }
      return new Response("sent via SMTP", { status: 200 });
    }

    if (!env("RESEND_API_KEY")) return new Response("no email sender configured", { status: 200 });
    const res = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: { Authorization: `Bearer ${env("RESEND_API_KEY")}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        from: settings.from_email || "ZTN Studio <onboarding@resend.dev>",
        to: msg.to,
        subject: plainSubject(msg.subject),
        text: msg.text,
      }),
    });
    return new Response(await res.text(), { status: res.ok ? 200 : 502 });
  } catch (e) {
    console.error(e);
    return new Response("error", { status: 500 });
  }
});
