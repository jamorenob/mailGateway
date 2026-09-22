# MailGateway — handoff for PediaSphere development

Status: **v2.0.0 in production since 2026-09-19** (v2.0.1 built, adds diagnostics; deploy pending).
Source: https://github.com/jamorenob/mailGateway (private; owner jamorenob). Dev copy: `C:\Dev\MailGateway`.

## What it is

An HTTP → Microsoft Graph mail relay (.NET 8 minimal API, zero NuGet dependencies). Callers POST JSON with
an API key; the gateway sends through Graph with the app registration of its tenant. One binary, two live
instances on the same IIS server (site `apps-thanos.central-office.info`):

| instance | URL | tenant / default sender | may send from |
|----------|-----|-------------------------|---------------|
| customer (APS / SPR) | `https://apps-thanos.central-office.info/mailgateway/api/mail/send` | customer tenant, `mail-sender@central-office.info` | aps1888.org, societyforpediatricresearch.org, pas-meeting.org, pasmeeting.org, central-office.info, aps-spr.org, midwestspr.org |
| provider (PediaSphere / piruli / hta / liceojavier) | `https://apps-thanos.central-office.info/mailgateway_AI/api/mail/send` | Antonio's tenant, `noreply@pediasphere.ai` | pediasphere.ai, piruli.com, hta-tech.com, liceojavier.com |

Health: same base URL + `/health` (shows version, sender, limits, config errors). Logs: `C:\Apps\MailGateway\Logs`
and `C:\Apps\MailGateway_AI\Logs`, one file per day, one `requestId` per request (also returned to the caller).
Each instance is an IIS application (`/mailgateway` → `C:\Apps\mailgateway_customer`, `/mailgateway_AI` →
`C:\Apps\mailgateway_provider`), app pools `mailgateway_customer` / `mailgateway_provider`. Config lives in
`appsettings.json` next to the DLL on the server (never in git).

PediaSphere (pas./spr.pediasphere.ai) already uses the **customer** instance — `MG_URL` and `MG_API_KEY` are
`Application()` values in each site's `global.asa`. Verified working on v2 (receipt from spr.pediasphere.ai, 2026-09-21).

## The API (v1 contract unchanged; everything new is optional)

```
POST {url}      X-API-Key: <key>      Content-Type: application/json
{
  "to":       "a@x.org; Name <b@y.org>",     // string or array; ; or , separated
  "cc": "…", "bcc": "…", "replyTo": "…",     // optional, same forms
  "from":     "noreply@pasmeeting.org",      // optional; must be allow-listed for that instance
  "fromName": "PAS Meeting",                 // optional
  "subject":  "…",
  "htmlBody": "…",  "textBody": "…",         // one required
  "appName":  "pas.pediasphere.ai",          // optional, appears in the gateway log
  "attachments": [                           // optional
    { "name": "Receipt-123.pdf", "contentBase64": "JVBERi0…" },              // bytes in the request (use this)
    { "name": "logo.png", "contentBase64": "…", "contentId": "logo", "isInline": true },  // <img src="cid:logo">
    { "path": "C:\\inetpub\\wwwroot\\x.pdf" }                                 // file on the gateway server (legacy)
  ]
}
```

Responses: `200 {success:true, requestId, message, mode, from}` · `400` validation (error text says what) ·
`401` bad key · `413` body too large · `502` Graph rejected it (Graph's own error code in the text) · `500` other.

Limits (per instance, `MailLimits`): currently 5 attachments / 5 MB each on both live instances (raise to 10 / 25 MB
by editing appsettings.json + recycling the pool). Base64 must be on one line — no CR/LF inside the JSON string.
Messages over ~3 MB go through a Graph draft + chunked upload automatically; nothing changes for the caller.

## Send-as rules (the part that bit us)

`SendMode` is `auto`: the gateway first posts to `/users/{from}/sendMail` (works only if `from` is a real or
**shared mailbox** in that tenant, needs just `Mail.Send`), and if Graph says it isn't a mailbox it falls back to
sending from the instance's default sender with a `from` header (needs Exchange *Send As* rights, otherwise
`ErrorSendAsDenied`). Practical rule: **every From address must be a shared mailbox** in its tenant.
Shared mailboxes created 2026-09-20 in the customer tenant: `noreply@aps1888.org`,
`noreply@societyforpediatricresearch.org`, `noreply@pas-meeting.org`, `noreply@pasmeeting.org`,
`noreply@midwestspr.org`. The `info@…` addresses in the allow-list have NOT been verified as mailboxes.
Mailbox aliases must be unique tenant-wide (`noreply-aps1888`, `noreply-spr`, …) even though the SMTP
addresses can repeat the local part across domains. Display name on the direct path follows `fromName`;
on the Send-As fallback Exchange substitutes the address's registered display name.

## Using it from Classic ASP (PediaSphere)

Reference client: `examples/classic-asp/mailgateway.inc.asp` in the repo. Include it once per page and call:

```vbscript
' plain
ok = MG_Send(toList, subject, htmlBody, textBody, fromAddr, fromName, "")
' with files (any server, any size within limits)
atts = MG_AttachFile(Server.MapPath("/receipts/123.pdf"), "Receipt-123.pdf", "")
atts = atts & "," & MG_AttachFile(Server.MapPath("/img/logo.png"), "logo.png", "logo")   ' inline, cid:logo
ok = MG_SendEx(toList, ccList, bccList, replyTo, subject, htmlBody, textBody, fromAddr, fromName, atts, "pas.pediasphere.ai")
' afterwards: MG_LastStatus, MG_LastError, MG_LastRequestId
```

It reads `Application("MG_URL")` / `Application("MG_API_KEY")` from global.asa. Helpers: `MG_JsonEscape`,
`MG_FileToBase64` (strips MSXML's 76-char line breaks), `MG_AttachBytes` for in-memory files.
Pages that already post `{to, subject, htmlBody, textBody}` by hand keep working; to add attachments, add the
`attachments` array — no other change.

Gotchas from the legacy conversion (in `legacy-asp/` in the repo): `#include` is a textual paste, so an include
used twice on one page cannot contain `Const`/`Function`/`Dim` under `Option Explicit`; a stale `MG_URL` in
global.asa silently overrides anything in code; MSXML base64 has line breaks; `Server.MapPath` beats hard-coded
character offsets for building paths.

## Deploying a new gateway version

```powershell
cd C:\Dev\MailGateway\src\MailGateway
dotnet build -c Release
Remove-Item C:\Dev\MailGateway\publish -Recurse -Force
dotnet publish -c Release -r win-x64 --self-contained false -o C:\Dev\MailGateway\publish
```
Copy `publish\` to the server as `C:\Apps\publish-v2`, then per instance:
`Stop-WebAppPool`, back up the folder, copy `MailGateway.dll/.pdb/.deps.json/.runtimeconfig.json` over
`C:\Apps\mailgateway_<x>` (never `appsettings.json` or `web.config`), `Start-WebAppPool`, check `/health`.
Test with `tools\Send-TestMail.ps1`. Dev box is Windows-on-ARM (`-r win-x64` matters); NuGet source must be
`nuget.org` for that publish; the project itself has no packages.

## Open items

* Deploy 2.0.1 (validation failures now log To/From/App/IP — needed to identify a caller rejected 2026-09-20 04:46 local for a non-allow-listed From).
* Verify the `info@…` addresses in the customer allow-list are mailboxes or Send-As-granted.
* Recent `tblAPS_Recruit_ads.PDFfile` URLs were built with a wrong path offset before 2026-09-21; check.
* Later: WordPress plugin on `examples/php-wordpress/mailgateway.php`; single multi-tenant endpoint routed by From-domain is an easy addition if ever wanted.
