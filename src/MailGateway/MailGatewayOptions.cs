namespace MailGateway;

/// <summary>
/// Strongly-typed view of appsettings.json. Section and key names are kept
/// identical to the original gateway so existing deployments keep working.
/// </summary>
public sealed class MailGatewayOptions
{
    public SecurityOptions Security { get; set; } = new();
    public GraphOptions Graph { get; set; } = new();
    public MailSecurityOptions MailSecurity { get; set; } = new();
    public LoggingOptions Logging { get; set; } = new();
    public MailLimitsOptions MailLimits { get; set; } = new();

    /// <summary>Returns a human-readable list of configuration problems (empty when valid).</summary>
    public List<string> Validate()
    {
        var problems = new List<string>();

        if (Security.AllKeys().Count == 0)
            problems.Add("Security:ApiKey (or Security:ApiKeys) is required.");

        if (string.IsNullOrWhiteSpace(Graph.TenantId)) problems.Add("Graph:TenantId is required.");
        if (string.IsNullOrWhiteSpace(Graph.ClientId)) problems.Add("Graph:ClientId is required.");
        if (string.IsNullOrWhiteSpace(Graph.ClientSecret)) problems.Add("Graph:ClientSecret is required.");
        if (string.IsNullOrWhiteSpace(Graph.Sender)) problems.Add("Graph:Sender is required.");

        var mode = Graph.NormalizedSendMode;
        if (mode != "auto" && mode != "direct" && mode != "sendas")
            problems.Add("Graph:SendMode must be 'auto', 'direct' or 'sendAs'.");

        if (MailLimits.MaxAttachments < 0) problems.Add("MailLimits:MaxAttachments cannot be negative.");
        if (MailLimits.MaxAttachmentBytes <= 0) problems.Add("MailLimits:MaxAttachmentBytes must be positive.");
        if (MailLimits.MaxTotalAttachmentBytes <= 0) problems.Add("MailLimits:MaxTotalAttachmentBytes must be positive.");
        if (MailLimits.MaxRequestBodyBytes <= 0) problems.Add("MailLimits:MaxRequestBodyBytes must be positive.");
        if (MailLimits.UploadChunkBytes <= 0 || MailLimits.UploadChunkBytes % 327_680 != 0 || MailLimits.UploadChunkBytes > 4 * 1024 * 1024)
            problems.Add("MailLimits:UploadChunkBytes must be a multiple of 327680 and at most 4194304.");

        return problems;
    }
}

public sealed class SecurityOptions
{
    /// <summary>Single API key (original setting). Still supported.</summary>
    public string? ApiKey { get; set; }

    /// <summary>Optional additional keys, so each calling project can have its own.</summary>
    public List<string> ApiKeys { get; set; } = new();

    public List<string> AllKeys()
    {
        var keys = new List<string>();
        if (!string.IsNullOrWhiteSpace(ApiKey)) keys.Add(ApiKey.Trim());
        foreach (var k in ApiKeys)
            if (!string.IsNullOrWhiteSpace(k)) keys.Add(k.Trim());
        return keys;
    }
}

public sealed class GraphOptions
{
    public string TenantId { get; set; } = "";
    public string ClientId { get; set; } = "";
    public string ClientSecret { get; set; } = "";

    /// <summary>Mailbox used when the caller does not specify a From, and the mailbox used for Send-As delivery.</summary>
    public string Sender { get; set; } = "";

    /// <summary>
    /// How to deliver when the effective From differs from <see cref="Sender"/>:
    ///   direct  - POST /users/{from}/sendMail (requires From to be a real user or shared mailbox; needs only Mail.Send)
    ///   sendAs  - POST /users/{Sender}/sendMail with a From header (requires Exchange "Send As" rights on Sender for that address)
    ///   auto    - try direct, fall back to sendAs when Graph reports the address is not a mailbox (default)
    /// </summary>
    public string SendMode { get; set; } = "auto";

    public string NormalizedSendMode => (SendMode ?? "auto").Trim().ToLowerInvariant();

    /// <summary>Whether Graph should keep a copy in the sending mailbox's Sent Items.</summary>
    public bool SaveToSentItems { get; set; } = false;

    /// <summary>Kept for backward compatibility with older provider configs; not used for validation (MailSecurity is).</summary>
    public List<string> AllowedSenders { get; set; } = new();
}

public sealed class MailSecurityOptions
{
    public bool AllowCustomFrom { get; set; } = false;
    public List<string> AllowedFromAddresses { get; set; } = new();
    public List<string> AllowedFromDomains { get; set; } = new();
}

public sealed class LoggingOptions
{
    /// <summary>Folder for daily log files. Defaults to {app}/logs when empty.</summary>
    public string? Folder { get; set; }
}

public sealed class MailLimitsOptions
{
    public int MaxAttachments { get; set; } = 5;

    /// <summary>Maximum size of a single attachment (decoded bytes).</summary>
    public long MaxAttachmentBytes { get; set; } = 5_000_000;

    /// <summary>Maximum combined size of all attachments (decoded bytes). Exchange Online rejects messages above ~150 MB; most tenants cap at 25-35 MB.</summary>
    public long MaxTotalAttachmentBytes { get; set; } = 25_000_000;

    /// <summary>Only files under this folder may be referenced by path. Leave empty to disable path attachments entirely.</summary>
    public string? AllowedAttachmentRoot { get; set; }

    /// <summary>Kestrel request-body limit. Base64 inflates attachments by ~37%, so keep this above MaxTotalAttachmentBytes * 1.4.</summary>
    public long MaxRequestBodyBytes { get; set; } = 64L * 1024 * 1024;

    /// <summary>Estimated /sendMail request size (base64 attachments + body text) above which the gateway switches to a draft + upload sessions. Graph's simple sendMail request limit is ~4 MB.</summary>
    public long InlineSendThresholdBytes { get; set; } = 3_000_000;

    /// <summary>Upload-session chunk size. Must be a multiple of 327,680 bytes and at most 4 MiB.</summary>
    public int UploadChunkBytes { get; set; } = 3_276_800;
}
