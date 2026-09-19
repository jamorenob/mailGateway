using System.Net.Mail;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using MailGateway;
using Microsoft.AspNetCore.HttpOverrides;

// ============================================================================
//  MailGateway  -  HTTP -> Microsoft Graph mail relay
//
//  GET  /health              liveness + effective configuration summary
//  POST /api/mail/send       X-Api-Key header, JSON body (see README.md)
//
//  One binary, many deployments: every environment keeps its own
//  appsettings.Production.json (never committed, never published).
// ============================================================================

var builder = WebApplication.CreateBuilder(args);

// ---- configuration ---------------------------------------------------------
var options = new MailGatewayOptions();
builder.Configuration.Bind(options);
var configProblems = options.Validate();

builder.Services.AddSingleton(options);
builder.Services.AddSingleton<FileLogger>();
builder.Services.AddSingleton<GraphTokenProvider>();
builder.Services.AddSingleton<GraphMailer>();
builder.Services.AddSingleton<AttachmentPreparer>();

// Pooled HttpClients (the original created a new HttpClient per request, which exhausts sockets under load).
builder.Services.AddHttpClient("graph", c =>
{
    c.BaseAddress = new Uri("https://graph.microsoft.com/v1.0/");
    c.Timeout = TimeSpan.FromMinutes(5);
});
builder.Services.AddHttpClient("token", c => c.Timeout = TimeSpan.FromSeconds(30));
builder.Services.AddHttpClient("upload", c => c.Timeout = TimeSpan.FromMinutes(5));

builder.Services.ConfigureHttpJsonOptions(o =>
{
    o.SerializerOptions.PropertyNamingPolicy = JsonNamingPolicy.CamelCase;
    o.SerializerOptions.PropertyNameCaseInsensitive = true;
    o.SerializerOptions.WriteIndented = true;
});

// Base64 attachments make request bodies large; raise the hosting limits accordingly.
builder.WebHost.ConfigureKestrel(k => k.Limits.MaxRequestBodySize = options.MailLimits.MaxRequestBodyBytes);
builder.Services.Configure<IISServerOptions>(o => o.MaxRequestBodySize = options.MailLimits.MaxRequestBodyBytes);

builder.Services.Configure<ForwardedHeadersOptions>(o =>
{
    o.ForwardedHeaders = ForwardedHeaders.XForwardedFor | ForwardedHeaders.XForwardedProto;
    // Default KnownProxies/KnownNetworks already trust loopback, which is where the IIS reverse-proxy rule forwards from.
});

var app = builder.Build();
app.UseForwardedHeaders();

var log = app.Services.GetRequiredService<FileLogger>();
var preparer = app.Services.GetRequiredService<AttachmentPreparer>();
var mailer = app.Services.GetRequiredService<GraphMailer>();
var apiKeys = options.Security.AllKeys();
var version = Assembly.GetExecutingAssembly().GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
              ?? Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "unknown";

await log.InfoAsync("startup", $"MailGateway {version} starting. Environment={app.Environment.EnvironmentName} Sender={options.Graph.Sender} SendMode={options.Graph.NormalizedSendMode} LogFolder={log.Folder}");
foreach (var problem in configProblems)
    await log.ErrorAsync("startup", $"Configuration problem: {problem}");

// ---- GET /health -----------------------------------------------------------
app.MapGet("/health", () => Results.Json(new
{
    success = configProblems.Count == 0,
    app = "MailGateway",
    version,
    environment = app.Environment.EnvironmentName,
    graphSender = options.Graph.Sender,
    sendMode = options.Graph.NormalizedSendMode,
    allowCustomFrom = options.MailSecurity.AllowCustomFrom,
    maxAttachments = options.MailLimits.MaxAttachments,
    maxAttachmentBytes = options.MailLimits.MaxAttachmentBytes,
    maxTotalAttachmentBytes = options.MailLimits.MaxTotalAttachmentBytes,
    pathAttachmentsEnabled = !string.IsNullOrWhiteSpace(options.MailLimits.AllowedAttachmentRoot),
    configErrors = configProblems,
    timestampUtc = DateTime.UtcNow,
}));

// ---- POST /api/mail/send ---------------------------------------------------
app.MapPost("/api/mail/send", async (HttpContext context, CancellationToken ct) =>
{
    var requestId = Guid.NewGuid().ToString("N");

    // -- authentication
    var providedKey = context.Request.Headers["X-Api-Key"].FirstOrDefault()
                      ?? context.Request.Headers["X-API-Key"].FirstOrDefault();
    if (!ApiKeyMatches(providedKey, apiKeys))
    {
        await log.WarnAsync(requestId, $"Unauthorized request from {context.Connection.RemoteIpAddress}.");
        return Results.Json(new ApiError(false, requestId, "Unauthorized"), statusCode: 401);
    }

    if (configProblems.Count > 0)
    {
        await log.ErrorAsync(requestId, "Rejected: gateway configuration is incomplete.");
        return Results.Json(new ApiError(false, requestId, "Gateway configuration is incomplete. See /health."), statusCode: 500);
    }

    // -- payload
    MailSendRequest? req;
    try
    {
        if (!context.Request.HasJsonContentType())
        {
            await log.WarnAsync(requestId, $"Unsupported Content-Type: {context.Request.ContentType}");
            return Results.Json(new ApiError(false, requestId, "Content-Type must be application/json"), statusCode: 415);
        }
        if (context.Request.ContentLength == 0)
        {
            await log.WarnAsync(requestId, "Empty request body.");
            return Results.Json(new ApiError(false, requestId, "Empty request body"), statusCode: 400);
        }
        req = await context.Request.ReadFromJsonAsync<MailSendRequest>(ct);
    }
    catch (JsonException ex)
    {
        await log.WarnAsync(requestId, $"Invalid JSON payload. {ex.Message}");
        return Results.Json(new ApiError(false, requestId, "Invalid JSON payload"), statusCode: 400);
    }
    catch (BadHttpRequestException ex)
    {
        await log.WarnAsync(requestId, $"Bad request: {ex.Message}");
        return Results.Json(new ApiError(false, requestId, ex.StatusCode == 413 ? "Request body too large" : "Bad request"), statusCode: ex.StatusCode);
    }

    if (req is null)
    {
        await log.WarnAsync(requestId, "Empty request body.");
        return Results.Json(new ApiError(false, requestId, "Empty request body"), statusCode: 400);
    }

    // -- validation
    var to = ParseRecipients(req.To, out var toError);
    var cc = ParseRecipients(req.Cc, out var ccError);
    var bcc = ParseRecipients(req.Bcc, out var bccError);
    var replyTo = ParseRecipients(req.ReplyTo, out var replyToError);

    string? validationError =
        to.Count == 0 ? "The 'to' field is required." :
        toError ?? ccError ?? bccError ?? replyToError ??
        (string.IsNullOrWhiteSpace(req.Subject) ? "The 'subject' field is required." :
        (string.IsNullOrWhiteSpace(req.HtmlBody) && string.IsNullOrWhiteSpace(req.TextBody)) ? "Either 'htmlBody' or 'textBody' is required." : null);

    string effectiveFrom = options.Graph.Sender.Trim();
    if (validationError is null && !string.IsNullOrWhiteSpace(req.From))
    {
        var requested = req.From.Trim();
        if (!MailAddress.TryCreate(requested, out var parsedFrom))
            validationError = "The 'from' field is not a valid email address.";
        else
        {
            requested = parsedFrom.Address;
            if (!string.Equals(requested, effectiveFrom, StringComparison.OrdinalIgnoreCase))
            {
                if (!options.MailSecurity.AllowCustomFrom) validationError = "Custom From address is not allowed.";
                else if (!IsAllowedFromAddress(requested, options.MailSecurity)) validationError = "Requested From address is not allowed.";
                else effectiveFrom = requested;
            }
            // Allow "Display Name <addr>" in 'from' to supply the name when fromName is empty.
            if (string.IsNullOrWhiteSpace(req.FromName) && !string.IsNullOrWhiteSpace(parsedFrom.DisplayName))
                req.FromName = parsedFrom.DisplayName;
        }
    }

    if (validationError is not null)
    {
        await log.WarnAsync(requestId, $"Validation failed: {validationError}");
        return Results.Json(new ApiError(false, requestId, validationError), statusCode: 400);
    }

    // -- attachments
    List<PreparedAttachment> attachments;
    try
    {
        attachments = await preparer.PrepareAsync(req.AllAttachments(), ct);
    }
    catch (AttachmentPreparer.ValidationException ex)
    {
        await log.WarnAsync(requestId, $"Attachment rejected: {ex.Message}");
        return Results.Json(new ApiError(false, requestId, ex.Message), statusCode: 400);
    }
    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
    {
        await log.ErrorAsync(requestId, $"Attachment preparation failed: {ex}");
        return Results.Json(new ApiError(false, requestId, "Attachment could not be read: " + ex.Message), statusCode: 400);
    }

    // -- send
    var attachmentSummary = attachments.Count == 0 ? "none"
        : string.Join(", ", attachments.Select(a => $"{a.Name} ({a.Bytes.LongLength} B{(a.IsInline ? ", inline" : "")})"));

    await log.InfoAsync(requestId,
        $"Sending mail via Graph. To={Join(to)} Cc={Join(cc)} Bcc={Join(bcc)} Subject={req.Subject!.Trim()} From={effectiveFrom} App={req.AppName ?? ""} Attachments={attachmentSummary}");

    try
    {
        var result = await mailer.SendAsync(requestId, effectiveFrom, req.FromName?.Trim(), to, cc, bcc, replyTo,
                                            req.Subject!.Trim(), req.HtmlBody, req.TextBody, attachments, ct);

        return Results.Json(new ApiOk(true, requestId, "Mail sent successfully", result.Mode, effectiveFrom));
    }
    catch (GraphException ex)
    {
        await log.ErrorAsync(requestId, $"Graph send failed: {ex.Message}");
        return Results.Json(new ApiError(false, requestId, $"Graph send failed: {ex.Message}"), statusCode: 502);
    }
    catch (OperationCanceledException) when (context.RequestAborted.IsCancellationRequested)
    {
        await log.WarnAsync(requestId, "Client disconnected before the send completed.");
        return Results.StatusCode(499);
    }
    catch (Exception ex)
    {
        await log.ErrorAsync(requestId, $"Unhandled error: {ex}");
        return Results.Json(new ApiError(false, requestId, "Unhandled error: " + ex.Message), statusCode: 500);
    }
});

app.Run();

// ============================================================================
//  local helpers
// ============================================================================

static bool ApiKeyMatches(string? provided, List<string> validKeys)
{
    if (string.IsNullOrWhiteSpace(provided)) return false;
    var providedBytes = Encoding.UTF8.GetBytes(provided.Trim());
    foreach (var key in validKeys)
    {
        var keyBytes = Encoding.UTF8.GetBytes(key);
        if (keyBytes.Length == providedBytes.Length && CryptographicOperations.FixedTimeEquals(keyBytes, providedBytes))
            return true;
    }
    return false;
}

static bool IsAllowedFromAddress(string address, MailSecurityOptions security)
{
    var normalized = address.Trim().ToLowerInvariant();

    foreach (var allowed in security.AllowedFromAddresses)
        if (string.Equals(allowed?.Trim(), normalized, StringComparison.OrdinalIgnoreCase)) return true;

    var at = normalized.LastIndexOf('@');
    if (at < 0 || at == normalized.Length - 1) return false;
    var domain = normalized[(at + 1)..];

    foreach (var allowedDomain in security.AllowedFromDomains)
        if (string.Equals(allowedDomain?.Trim(), domain, StringComparison.OrdinalIgnoreCase)) return true;

    return false;
}

/// <summary>Accepts "a@x.com; Name &lt;b@y.com&gt;, c@z.com" or a list; returns validated recipients.</summary>
static List<Recipient> ParseRecipients(List<string>? raw, out string? error)
{
    error = null;
    var result = new List<Recipient>();
    if (raw is null) return result;

    foreach (var entry in raw)
    {
        if (string.IsNullOrWhiteSpace(entry)) continue;

        // A single RFC-style entry such as "Doe, John" <john@x.org> contains a comma; try it whole first.
        if (entry.Contains('<') && MailAddress.TryCreate(entry.Trim(), out var whole))
        {
            result.Add(new Recipient(whole.Address, string.IsNullOrWhiteSpace(whole.DisplayName) ? null : whole.DisplayName));
            continue;
        }

        foreach (var piece in entry.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (!MailAddress.TryCreate(piece, out var addr))
            {
                error ??= $"Invalid email address: {piece}";
                continue;
            }
            result.Add(new Recipient(addr.Address, string.IsNullOrWhiteSpace(addr.DisplayName) ? null : addr.DisplayName));
        }
    }
    return result;
}

static string Join(List<Recipient> list) => string.Join(";", list.Select(r => r.Address));
