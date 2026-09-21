# Legacy Classic ASP callers (APS / SPR sites)

Drop-in replacements for the old CDO/SMTP includes on the legacy society sites, so those pages send
through the MailGateway (customer instance, `/mailgateway/`) instead of SMTP.

| file | goes to | purpose |
|------|---------|---------|
| `send_email.asp` | `/includes/send_email.asp` (APS also: `send_email1.asp`, `sendmail/send_email.asp`) | plain email from the page's session variables |
| `send_email_ad.asp` | `/includes/send_email_ad.asp` | same, plus one attachment (`session("pdf_localpath")` or `session("pdf1")`) |
| `aps/uploads.asp`, `aps/reupload-pdf.asp` | `www.aps1888.org/ad-recruit/PDF/` | APS Career Center upload pages converted from inline CDO to the include |

Contract (unchanged from the CDO days): `session("subject_line")`, `session("from_line")`, `session("bcc_line")`,
`session("email")` (To), `session("message")`; result in `session("smtperror")` and written to the page.

Design notes, learned the hard way:

* The includes contain no `Const`/`Function`/`Sub` and declare variables via `ExecuteGlobal`, because
  `upload-pdf.asp` includes the file twice and the sites run `Option Explicit`.
* `MG_URL` / `MG_API_KEY` come from `Application()` (global.asa) when present, else from the constants in the
  file. Pick one convention per site — a stale `MG_URL` in global.asa silently overrides the code.
* MSXML's base64 output has line breaks every 76 chars; they are stripped before building the JSON.
* From addresses must be mailboxes (or Send-As-granted) in the customer tenant: `noreply@aps1888.org`,
  `noreply@societyforpediatricresearch.org`, etc. are shared mailboxes.
* `session("pdf1")` on the APS pages is now computed from `Server.MapPath("/")` instead of a fixed
  character offset, so the stored URL is right on any server.
