# DNS records for ZTN studio booking emails

**For:** whoever manages the **ztn.co.zw** zone in **Cloudflare**
**Why:** the studio booking system sends emails from `studio@notify.ztn.co.zw` through Resend. These records prove we're allowed to send from that address, so emails don't land in spam. They only touch the new subdomain `notify.ztn.co.zw`; ZTN's normal email is not affected.

In Cloudflare, open **ztn.co.zw → DNS → Records → Add record** and add these. Names are as Cloudflare shows them (it adds `.ztn.co.zw` itself).

| # | Type | Name | Content / Target | Proxy | TTL |
|---|---|---|---|---|---|
| 1 | TXT | `resend._domainkey.notify` | `p=MIGfMA0GCSqGSIb3DQEBAQUAA4GNADCBiQKBgQC9/8wZhNlFxDo4ZuY/fMpR63qwGGqYqQscHjII4+hEs3lpSt/cIiknXMqxqq1vKb6Hb1RFE5YQsfIWOX5qoeaG5+w/gdg6x/bG0E56GxMk9FxpKeAsdm79PYTohrrjF4Cw8a+Fqrn9FWDGXO1ojHSW++NSWMcJYFq4S2wJTzNp2wIDAQAB` | – | Auto |
| 2 | CNAME | `rsend.notify` | `rsend-euw1.forge.rmta.net` | **DNS only** (grey cloud) | Auto |
| 3 | CNAME | `send.notify` | `send.forge.rmta.net` | **DNS only** (grey cloud) | Auto |
| 4 | TXT | `_dmarc` | `v=DMARC1; p=none;` | – | Auto |

**Notes:**
- Records 2 and 3 must be **DNS only**. If the orange cloud (proxied) is on, verification fails.
- Record 4 is optional, but recommended. ztn.co.zw has no DMARC record today, and `p=none` only adds reporting; it doesn't block any mail. Skip it if ZTN already manages DMARC elsewhere.
- Copy record 1's value exactly, on one line, with no spaces.

When they're added, tell the booking system admin. They'll click **Verify** in Resend; this usually takes a few minutes, at most a few hours.
