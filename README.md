# MailGateway

Small HTTP → Microsoft Graph mail relay (.NET 8 minimal API, zero NuGet dependencies). One binary, many deployments: WordPress, the legacy admin site,
PediaSphere (PAS / SPR), piruli, liceojavier … anything that can POST JSON.

```
POST /api/mail/send      X-API-Key: <key>      Content-Type: application/json
GET  /health
```

## Repository layout

```
src/MailGateway/            the app (Program.cs, GraphMailer.cs, AttachmentPreparer.cs, ...)
deploy/customer/            appsettings.Production.example.json for the customer tenant (port 5000)
deploy/provider/            appsettings.Production.example.json for the provider tenant (port 5001)
deploy/iis-site/web.config  IIS reverse-proxy site folder (ARR → localhost:port)
deploy/iis-inprocess/       alternative: host inside IIS instead of as a Windows service
examples/classic-asp/       mailgateway.inc.asp  – VBScript client with base64 attachments
examples/php-wordpress/     mailgateway.php      – PHP client + optional wp_mail() override
tools/Send-TestMail.ps1     end-to-end test from PowerShell
```

`MailGateway_Customer/` and `MailGateway_Provider/` (the published folders that live next to this
repo on the dev box) are git-ignored on purpose: they contain real keys.

## Request

| field            | type                    | notes |
|------------------|-------------------------|-------|
| `to`             | string or array         | required. `"a@x.org; b@y.org"` or `["a@x.org","Name <b@y.org>"]` |
| `cc`, `bcc`      | string or array         | optional |
| `replyTo`        | string or array         | optional |
| `subject`        | string                  | required |
| `htmlBody`       | string                  | one of `htmlBody` / `textBody` required |
| `textBody`       | string                  | |
| `from`           | string                  | optional; must pass `MailSecurity` allow-list. `"Name <addr>"` accepted |
| `fromName`       | string                  | optional display name |
| `appName`        | string                  | optional, logs only |
| `attachments`    | array                   | see below. `attachmentItems` is still accepted as an alias |

### Attachments

Each element is either a **string** (a file path on the gateway server, original behaviour, must be under
`MailLimits:AllowedAttachmentRoot`) or an **object**:

```json
{
  "name": "Receipt-12345.pdf",
  "contentBase64": "JVBERi0xLjcK...",        // the file bytes on one line (a data:...;base64, prefix is tolerated)
  "contentType": "application/pdf",          // optional, guessed from the extension
  "contentId": "logo",                       // optional: inline image referenced as <img src="cid:logo">
  "isInline": true                           // optional, default false
}
```

or `{ "path": "C:\\inetpub\\wwwroot\\files\\x.pdf", "name": "optional override" }`.

PDF, JPG, PNG, GIF, Word, Excel, PowerPoint, CSV, ZIP … anything goes; the MIME type comes from the
extension or from `contentType`. Limits are per deployment (`MailLimits`). Messages whose estimated
request size (base64 attachments + body) exceeds ~3 MB go through a Graph draft + chunked upload sessions automatically, so
20-MB scans work the same way as a 50-KB receipt from the caller's point of view.

### Response

```json
{ "success": true,  "requestId": "3f1c...", "message": "Mail sent successfully", "mode": "direct", "from": "noreply@pasmeeting.org" }
{ "success": false, "requestId": "3f1c...", "error": "Requested From address is not allowed." }
```

HTTP 200 sent · 400 validation · 401 bad key · 413 body too large · 502 Graph rejected it (error text
includes Graph's own code, e.g. `ErrorSendAsDenied`) · 500 unexpected. The `requestId` matches the
line in the daily log file (`Logging:Folder\mailgateway-yyyyMMdd.log`).

## Send As — how it works and what Entra/Exchange needs

`Graph:SendMode`:

* **`direct`** – POST `/users/{from}/sendMail`. Needs only the app permission **Mail.Send**
  (application, admin-consented). The `from` address must be a real user or **shared mailbox**
  in that tenant. This is the reliable path: create a free shared mailbox for each sender
  (`noreply@…`, `info@…`) and nothing else is required.
* **`sendAs`** – POST `/users/{Graph:Sender}/sendMail` with a `from` header (what the old gateway
  always did). Works only if the `Sender` mailbox has Exchange **Send As** rights on that address:
  `Add-RecipientPermission noreply@x.org -Trustee mail-sender@central-office.info -AccessRights SendAs`.
  Otherwise Graph answers `ErrorSendAsDenied` — or silently rewrites the From.
* **`auto`** (default) – tries `direct`; if Graph says the address is not a mailbox (403/404) it falls
  back to `sendAs` and remembers that for 6 hours. Existing configs keep working, and every address
  you turn into a shared mailbox automatically moves to the reliable path.

Optional hardening: an Exchange **Application Access Policy** can restrict the app registration to the
mailboxes it may send from (`New-ApplicationAccessPolicy -AppId <clientId> -PolicyScopeGroupId <mail-enabled group> -AccessRight RestrictAccess`).

## Configuration

Copy `deploy/<env>/appsettings.Production.example.json` to `appsettings.Production.json` next to the
deployed `MailGateway.dll` and fill in the secrets. Keys can also come from environment variables
(`Security__ApiKey`, `Graph__ClientSecret`, …). The committed `src/MailGateway/appsettings.json` is a
blank template and is the only settings file that gets published.

| key | meaning |
|-----|---------|
| `Security:ApiKey` / `Security:ApiKeys[]` | one or more accepted keys (give each project its own) |
| `Graph:TenantId/ClientId/ClientSecret` | app registration with `Mail.Send` (application) |
| `Graph:Sender` | default From and the Send-As mailbox |
| `Graph:SendMode` | `auto` / `direct` / `sendAs` |
| `Graph:SaveToSentItems` | keep a copy in Sent Items (small messages only; drafts always land there) |
| `MailSecurity:AllowCustomFrom` | allow callers to set `from` |
| `MailSecurity:AllowedFromAddresses[]`, `AllowedFromDomains[]` | allow-list |
| `MailLimits:MaxAttachments` | count per message |
| `MailLimits:MaxAttachmentBytes` | per file (decoded) |
| `MailLimits:MaxTotalAttachmentBytes` | per message (decoded); Exchange Online caps ~150 MB, most tenants 25–35 MB |
| `MailLimits:AllowedAttachmentRoot` | folder for path attachments; empty = base64 only |
| `MailLimits:MaxRequestBodyBytes` | Kestrel/IIS body limit — keep ≥ 1.4 × total attachment bytes |
| `MailLimits:InlineSendThresholdBytes` | estimated request size (base64 + body) above which the draft + upload-session flow is used (Graph's /sendMail limit ≈ 4 MB) |
| `MailLimits:UploadChunkBytes` | multiple of 327,680, ≤ 4 MiB |
| `Logging:Folder` | daily log files |
| `Urls` | listen address when running as a service/console (omit under IIS) |

## Build, run, deploy

```powershell
cd C:\Dev\MailGateway\src\MailGateway
dotnet restore
dotnet build -c Release

# run locally (uses appsettings.Development.json if present, else appsettings.json)
dotnet run

# publish for the customer / provider service folders
dotnet publish -c Release -o C:\Dev\MailGateway\publish
```

Then stop the service, copy the contents of `publish\` over `C:\Apps\MailGateway` (or
`C:\Apps\MailGateway_AI`), make sure `appsettings.Production.json` is still there, start the service.
Check `GET /health` — `configErrors` must be empty, and the response shows the version.

The app is a plain console process, so it runs as a service the same way the old one does (NSSM,
a scheduled task, or whatever wrapper is already in place) — just replace the files. For a brand-new
box, NSSM is the simplest: `nssm install MailGateway C:\Apps\MailGateway\MailGateway.exe`.

If the public URL goes through IIS + ARR (current setup), copy `deploy/iis-site/web.config` into the
site folder — it raises IIS's 30 MB request limit so large base64 bodies are not cut off. The Classic ASP
proxy `send.asp` in the provider folder is limited by `AspMaxRequestEntityAllowed` (200 KB default);
prefer the rewrite rule for attachment traffic.

## Callers

* Classic ASP: `<!--#include virtual="/inc/mailgateway.inc.asp" -->` then `MG_Send(...)` /
  `MG_AttachFile(...)`. Payload stays identical to the old `{to, subject, htmlBody, textBody}` call when no
  extras are used.
* PHP / WordPress: `require 'mailgateway.php'; mg_send([...])`, or enable the `pre_wp_mail` filter to route all
  WordPress mail through the gateway.
* PowerShell smoke test: `tools\Send-TestMail.ps1`.

## What changed vs. the original (v1)

* Attachments can be sent **in the request** (`contentBase64`), not only as paths on the gateway's disk.
* Large attachments use Graph upload sessions instead of failing above 3 MB.
* Send-As: direct-mailbox send with automatic fallback; Graph's error code and message are now logged and returned.
* OAuth token cached until expiry; `IHttpClientFactory` instead of a new `HttpClient` per request; 401 retry.
* Multiple API keys; recipients accept arrays or `Name <addr>`; constant-time key comparison.
* Secrets separated from the repo (`appsettings.Production.json`, git-ignored, never published).
