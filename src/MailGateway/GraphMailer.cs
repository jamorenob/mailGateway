using System.Collections.Concurrent;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace MailGateway;

/// <summary>
/// Sends mail through Microsoft Graph.
///
/// Send-as strategy (Graph:SendMode):
///   direct  - POST /users/{from}/sendMail. Works for any user or shared mailbox with only the Mail.Send application permission.
///   sendAs  - POST /users/{Graph:Sender}/sendMail with a From header. Needs Exchange "Send As" rights on the Sender mailbox for {from}.
///   auto    - direct first; if Graph says {from} is not a mailbox (403/404) fall back to sendAs and remember that for a few hours.
///
/// Attachment strategy:
///   small total  - single /sendMail call with fileAttachment items inline (Graph limit ≈ 3 MB for the whole request)
///   large total  - create a draft, add each attachment (inline if ≤ 3 MB, otherwise an upload session in chunks), then /send
/// </summary>
public sealed class GraphMailer
{
    private const long GraphInlineAttachmentLimit = 3_000_000; // Graph rejects single fileAttachment POSTs above ~3 MB
    private static readonly TimeSpan NotMailboxMemory = TimeSpan.FromHours(6);

    private static readonly JsonSerializerOptions JsonOpts = new(JsonSerializerDefaults.Web)
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
    };

    private readonly IHttpClientFactory _http;
    private readonly GraphTokenProvider _tokens;
    private readonly MailGatewayOptions _options;
    private readonly FileLogger _log;

    /// <summary>Addresses that turned out not to be mailboxes (auto mode), with the time we learned it.</summary>
    private readonly ConcurrentDictionary<string, DateTimeOffset> _notMailbox = new(StringComparer.OrdinalIgnoreCase);

    public GraphMailer(IHttpClientFactory http, GraphTokenProvider tokens, MailGatewayOptions options, FileLogger log)
    {
        _http = http;
        _tokens = tokens;
        _options = options;
        _log = log;
    }

    public sealed record SendResult(string Mode, string Mailbox);

    public async Task<SendResult> SendAsync(
        string requestId,
        string effectiveFrom,
        string? fromName,
        List<Recipient> to,
        List<Recipient> cc,
        List<Recipient> bcc,
        List<Recipient> replyTo,
        string subject,
        string? htmlBody,
        string? textBody,
        IReadOnlyList<PreparedAttachment> attachments,
        CancellationToken ct)
    {
        var sender = _options.Graph.Sender.Trim();
        var mode = _options.Graph.NormalizedSendMode;
        var sameMailbox = string.Equals(effectiveFrom, sender, StringComparison.OrdinalIgnoreCase);

        // Decide the first attempt.
        bool useSendAs;
        if (sameMailbox) useSendAs = false;                    // sending from the configured mailbox itself
        else if (mode == "sendas") useSendAs = true;
        else if (mode == "direct") useSendAs = false;
        else useSendAs = IsKnownNotMailbox(effectiveFrom);     // auto

        try
        {
            return await SendCoreAsync(requestId, useSendAs, sender, effectiveFrom, fromName, to, cc, bcc, replyTo, subject, htmlBody, textBody, attachments, ct);
        }
        catch (GraphException ex) when (mode == "auto" && !useSendAs && !sameMailbox && IsNotMailboxError(ex))
        {
            await _log.WarnAsync(requestId, $"Direct send from {effectiveFrom} rejected ({ex.StatusCode} {ex.GraphErrorCode}); falling back to Send-As via {sender}.");
            _notMailbox[effectiveFrom] = DateTimeOffset.UtcNow;
            return await SendCoreAsync(requestId, true, sender, effectiveFrom, fromName, to, cc, bcc, replyTo, subject, htmlBody, textBody, attachments, ct);
        }
    }

    private async Task<SendResult> SendCoreAsync(
        string requestId, bool useSendAs, string sender, string effectiveFrom, string? fromName,
        List<Recipient> to, List<Recipient> cc, List<Recipient> bcc, List<Recipient> replyTo,
        string subject, string? htmlBody, string? textBody,
        IReadOnlyList<PreparedAttachment> attachments, CancellationToken ct)
    {
        var mailbox = useSendAs ? sender : effectiveFrom;
        var modeName = useSendAs ? "sendAs" : "direct";
        var includeFromHeader = useSendAs || !string.IsNullOrWhiteSpace(fromName);

        long totalBytes = 0;
        foreach (var a in attachments) totalBytes += a.Bytes.LongLength;
        var anyLarge = attachments.Any(a => a.Bytes.LongLength > GraphInlineAttachmentLimit);

        // Graph's /sendMail limit applies to the whole JSON request: base64 attachments (+33%) plus the body text.
        var estimatedRequestBytes = totalBytes * 4 / 3 + (htmlBody?.Length ?? 0) + (textBody?.Length ?? 0);
        var useDraftFlow = anyLarge || estimatedRequestBytes > _options.MailLimits.InlineSendThresholdBytes;

        var message = BuildMessage(effectiveFrom, fromName, includeFromHeader, to, cc, bcc, replyTo, subject, htmlBody, textBody,
                                   useDraftFlow ? Array.Empty<PreparedAttachment>() : attachments);

        var userPath = $"users/{Uri.EscapeDataString(mailbox)}";

        if (!useDraftFlow)
        {
            var payload = new Dictionary<string, object?>
            {
                ["message"] = message,
                ["saveToSentItems"] = _options.Graph.SaveToSentItems,
            };
            await CallAsync(requestId, HttpMethod.Post, $"{userPath}/sendMail", payload, ct);
            await _log.InfoAsync(requestId, $"Mail sent via Graph ({modeName}, mailbox={mailbox}, attachments={attachments.Count}, bytes={totalBytes}).");
            return new SendResult(modeName, mailbox);
        }

        // ---- Draft + attachments + send (large messages) ----
        await _log.InfoAsync(requestId, $"Large message ({totalBytes} bytes across {attachments.Count} attachments); using draft + upload sessions.");

        var (_, draftBody) = await CallAsync(requestId, HttpMethod.Post, $"{userPath}/messages", message, ct);
        string draftId;
        using (var doc = JsonDocument.Parse(draftBody))
        {
            if (!doc.RootElement.TryGetProperty("id", out var idEl) || string.IsNullOrWhiteSpace(idEl.GetString()))
                throw new GraphException("Graph draft creation did not return a message id.");
            draftId = idEl.GetString()!;
        }

        var draftPath = $"{userPath}/messages/{Uri.EscapeDataString(draftId)}";

        try
        {
            foreach (var att in attachments)
            {
                if (att.Bytes.LongLength <= GraphInlineAttachmentLimit)
                {
                    await CallAsync(requestId, HttpMethod.Post, $"{draftPath}/attachments", ToFileAttachment(att), ct);
                }
                else
                {
                    await UploadLargeAttachmentAsync(requestId, draftPath, att, ct);
                }
            }

            await CallAsync(requestId, HttpMethod.Post, $"{draftPath}/send", null, ct);
        }
        catch
        {
            // Best effort: do not leave half-built drafts in the mailbox.
            try
            {
                using var cleanupCts = new CancellationTokenSource(TimeSpan.FromSeconds(30));
                await CallAsync(requestId, HttpMethod.Delete, draftPath, null, cleanupCts.Token);
            }
            catch (Exception cleanupEx) { await _log.WarnAsync(requestId, $"Could not delete draft {draftId}: {cleanupEx.Message}"); }
            throw;
        }

        // Graph's /send on a draft ignores saveToSentItems; drafts sent this way always land in Sent Items unless removed.
        await _log.InfoAsync(requestId, $"Mail sent via Graph draft flow ({modeName}, mailbox={mailbox}, attachments={attachments.Count}, bytes={totalBytes}).");
        return new SendResult(modeName, mailbox);
    }

    private async Task UploadLargeAttachmentAsync(string requestId, string draftPath, PreparedAttachment att, CancellationToken ct)
    {
        var sessionRequest = new Dictionary<string, object?>
        {
            ["AttachmentItem"] = new Dictionary<string, object?>
            {
                ["attachmentType"] = "file",
                ["name"] = att.Name,
                ["size"] = att.Bytes.LongLength,
                ["contentType"] = att.ContentType,
                ["isInline"] = att.IsInline,
                ["contentId"] = att.IsInline ? att.ContentId : null,
            }
        };

        var (_, sessionBody) = await CallAsync(requestId, HttpMethod.Post, $"{draftPath}/attachments/createUploadSession", sessionRequest, ct);

        string uploadUrl;
        using (var doc = JsonDocument.Parse(sessionBody))
        {
            if (!doc.RootElement.TryGetProperty("uploadUrl", out var urlEl) || string.IsNullOrWhiteSpace(urlEl.GetString()))
                throw new GraphException("Graph upload session did not return an uploadUrl.");
            uploadUrl = urlEl.GetString()!;
        }

        var chunkSize = _options.MailLimits.UploadChunkBytes;
        var total = att.Bytes.LongLength;
        var uploader = _http.CreateClient("upload"); // no Authorization header: the uploadUrl is pre-authenticated

        long offset = 0;
        while (offset < total)
        {
            var length = (int)Math.Min(chunkSize, total - offset);
            var end = offset + length - 1;

            using var content = new ByteArrayContent(att.Bytes, (int)offset, length);
            content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
            content.Headers.ContentRange = new ContentRangeHeaderValue(offset, end, total);
            content.Headers.ContentLength = length;

            using var request = new HttpRequestMessage(HttpMethod.Put, uploadUrl) { Content = content };
            using var response = await uploader.SendAsync(request, ct);

            if (!response.IsSuccessStatusCode)
            {
                var body = await response.Content.ReadAsStringAsync(ct);
                throw new GraphException($"Upload of '{att.Name}' failed at bytes {offset}-{end}/{total} ({(int)response.StatusCode}): {Truncate(body)}", (int)response.StatusCode);
            }

            offset += length;
        }

        await _log.InfoAsync(requestId, $"Uploaded large attachment '{att.Name}' ({total} bytes) in {Math.Ceiling(total / (double)chunkSize)} chunk(s).");
    }

    // ------------------------------------------------------------------ helpers

    private static Dictionary<string, object?> BuildMessage(
        string from, string? fromName, bool includeFromHeader,
        List<Recipient> to, List<Recipient> cc, List<Recipient> bcc, List<Recipient> replyTo,
        string subject, string? htmlBody, string? textBody, IReadOnlyList<PreparedAttachment> attachments)
    {
        var useHtml = !string.IsNullOrWhiteSpace(htmlBody);

        var message = new Dictionary<string, object?>
        {
            ["subject"] = subject,
            ["body"] = new Dictionary<string, object?>
            {
                ["contentType"] = useHtml ? "HTML" : "Text",
                ["content"] = useHtml ? htmlBody : (textBody ?? ""),
            },
            ["toRecipients"] = to.Select(ToGraphRecipient).ToList(),
        };

        if (cc.Count > 0) message["ccRecipients"] = cc.Select(ToGraphRecipient).ToList();
        if (bcc.Count > 0) message["bccRecipients"] = bcc.Select(ToGraphRecipient).ToList();
        if (replyTo.Count > 0) message["replyTo"] = replyTo.Select(ToGraphRecipient).ToList();
        if (includeFromHeader) message["from"] = ToGraphRecipient(new Recipient(from, fromName));
        if (attachments.Count > 0) message["attachments"] = attachments.Select(ToFileAttachment).ToList();

        return message;
    }

    private static Dictionary<string, object?> ToGraphRecipient(Recipient r) => new()
    {
        ["emailAddress"] = new Dictionary<string, object?>
        {
            ["address"] = r.Address,
            ["name"] = string.IsNullOrWhiteSpace(r.Name) ? null : r.Name,
        }
    };

    private static Dictionary<string, object?> ToFileAttachment(PreparedAttachment a) => new()
    {
        ["@odata.type"] = "#microsoft.graph.fileAttachment",
        ["name"] = a.Name,
        ["contentType"] = a.ContentType,
        ["contentBytes"] = Convert.ToBase64String(a.Bytes),
        ["contentId"] = a.IsInline ? a.ContentId : null,
        ["isInline"] = a.IsInline,
    };

    /// <summary>Calls Graph with a bearer token; retries once with a fresh token on 401; throws GraphException on non-2xx.</summary>
    private async Task<(int Status, string Body)> CallAsync(string requestId, HttpMethod method, string relativeUrl, object? payload, CancellationToken ct)
    {
        for (var attempt = 1; ; attempt++)
        {
            var token = await _tokens.GetAccessTokenAsync(ct);
            var client = _http.CreateClient("graph");

            using var request = new HttpRequestMessage(method, relativeUrl);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            if (payload is not null)
                request.Content = new StringContent(JsonSerializer.Serialize(payload, JsonOpts), Encoding.UTF8, "application/json");

            using var response = await client.SendAsync(request, ct);
            var body = await response.Content.ReadAsStringAsync(ct);
            var status = (int)response.StatusCode;

            if (response.IsSuccessStatusCode)
                return (status, body);

            if (response.StatusCode == HttpStatusCode.Unauthorized && attempt == 1)
            {
                _tokens.Invalidate();
                await _log.WarnAsync(requestId, "Graph returned 401; refreshing token and retrying once.");
                continue;
            }

            var (code, msg) = ParseGraphError(body);
            throw new GraphException($"Graph {method} {relativeUrl} failed ({status} {code}): {msg}", status, code);
        }
    }

    private static (string Code, string Message) ParseGraphError(string body)
    {
        try
        {
            using var doc = JsonDocument.Parse(body);
            if (doc.RootElement.TryGetProperty("error", out var err))
            {
                var code = err.TryGetProperty("code", out var c) ? c.GetString() ?? "" : "";
                var msg = err.TryGetProperty("message", out var m) ? m.GetString() ?? "" : "";
                return (code, Truncate(msg));
            }
        }
        catch { /* not JSON */ }
        return ("", Truncate(body));
    }

    private static bool IsNotMailboxError(GraphException ex) =>
        ex.StatusCode is 404 or 403;

    private bool IsKnownNotMailbox(string address)
    {
        if (!_notMailbox.TryGetValue(address, out var when)) return false;
        if (DateTimeOffset.UtcNow - when > NotMailboxMemory)
        {
            _notMailbox.TryRemove(address, out _);
            return false;
        }
        return true;
    }

    private static string Truncate(string s, int max = 2000) =>
        string.IsNullOrEmpty(s) ? "" : (s.Length <= max ? s : s[..max] + "…");
}

public sealed record Recipient(string Address, string? Name);
