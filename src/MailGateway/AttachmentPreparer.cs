namespace MailGateway;

/// <summary>
/// Turns the caller's attachment items into byte arrays, enforcing MailLimits.
/// Two sources are supported:
///   path           - a file on the gateway server under MailLimits:AllowedAttachmentRoot (original behaviour)
///   contentBase64  - the file bytes sent in the request (new; works from WordPress, PediaSphere, or any remote caller)
/// </summary>
public sealed class AttachmentPreparer
{
    private readonly MailLimitsOptions _limits;

    public AttachmentPreparer(MailGatewayOptions options)
    {
        _limits = options.MailLimits;
    }

    public sealed class ValidationException : Exception
    {
        public ValidationException(string message) : base(message) { }
    }

    public async Task<List<PreparedAttachment>> PrepareAsync(IEnumerable<MailAttachmentItem> items, CancellationToken ct)
    {
        var list = items.Where(i => i is not null).ToList();
        var result = new List<PreparedAttachment>(list.Count);

        if (list.Count > _limits.MaxAttachments)
            throw new ValidationException($"Too many attachments. Max allowed: {_limits.MaxAttachments}.");

        long total = 0;
        foreach (var item in list)
        {
            var prepared = !string.IsNullOrWhiteSpace(item.ContentBase64)
                ? FromBase64(item)
                : await FromPathAsync(item, ct);

            total += prepared.Bytes.LongLength;
            if (total > _limits.MaxTotalAttachmentBytes)
                throw new ValidationException($"Attachments too large in total. Max allowed: {_limits.MaxTotalAttachmentBytes} bytes.");

            result.Add(prepared);
        }

        return result;
    }

    private PreparedAttachment FromBase64(MailAttachmentItem item)
    {
        if (string.IsNullOrWhiteSpace(item.Name))
            throw new ValidationException("Attachment 'name' is required when 'contentBase64' is used.");

        var name = SanitizeFileName(item.Name!);
        var data = item.ContentBase64!.Trim();

        // Tolerate data URIs: data:application/pdf;base64,JVBERi0...
        var comma = data.IndexOf(',');
        if (data.StartsWith("data:", StringComparison.OrdinalIgnoreCase) && comma > 0)
            data = data[(comma + 1)..];

        // Tolerate whitespace/newlines inserted by some encoders (VBScript's MSXML wraps at 76 chars).
        if (data.IndexOfAny(new[] { '\r', '\n', ' ', '\t' }) >= 0)
            data = string.Concat(data.Where(c => !char.IsWhiteSpace(c)));

        // Reject oversized payloads before decoding them into memory (base64 length ≈ 4/3 of the decoded size).
        var estimatedBytes = (long)data.Length / 4 * 3 - 2;
        if (estimatedBytes > _limits.MaxAttachmentBytes)
            throw new ValidationException($"Attachment too large: {name} (about {estimatedBytes} bytes; max {_limits.MaxAttachmentBytes}).");

        byte[] bytes;
        try { bytes = Convert.FromBase64String(data); }
        catch (FormatException) { throw new ValidationException($"Attachment '{name}': contentBase64 is not valid base64."); }

        if (bytes.LongLength == 0)
            throw new ValidationException($"Attachment '{name}' is empty.");
        if (bytes.LongLength > _limits.MaxAttachmentBytes)
            throw new ValidationException($"Attachment too large: {name} ({bytes.LongLength} bytes; max {_limits.MaxAttachmentBytes}).");

        return new PreparedAttachment
        {
            Name = name,
            ContentType = string.IsNullOrWhiteSpace(item.ContentType) ? MimeTypes.FromFileName(name) : item.ContentType!.Trim(),
            Bytes = bytes,
            ContentId = item.ContentId,
            IsInline = item.IsInline,
        };
    }

    private async Task<PreparedAttachment> FromPathAsync(MailAttachmentItem item, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(item.Path))
            throw new ValidationException("Each attachment needs either 'path' or 'contentBase64'.");

        if (string.IsNullOrWhiteSpace(_limits.AllowedAttachmentRoot))
            throw new ValidationException("Attachment path is not allowed (MailLimits:AllowedAttachmentRoot is not configured; send contentBase64 instead).");

        var root = Path.GetFullPath(_limits.AllowedAttachmentRoot!);
        if (!root.EndsWith(Path.DirectorySeparatorChar)) root += Path.DirectorySeparatorChar;

        var full = Path.GetFullPath(item.Path!.Trim());
        if (!full.StartsWith(root, StringComparison.OrdinalIgnoreCase))
            throw new ValidationException("Attachment path is not allowed.");

        if (!File.Exists(full))
            throw new ValidationException($"Attachment file not found: {item.Path}");

        var fi = new FileInfo(full);
        if (fi.Length == 0)
            throw new ValidationException($"Attachment '{fi.Name}' is empty.");
        if (fi.Length > _limits.MaxAttachmentBytes)
            throw new ValidationException($"Attachment too large: {fi.Name} ({fi.Length} bytes; max {_limits.MaxAttachmentBytes}).");

        var bytes = await File.ReadAllBytesAsync(full, ct);
        var name = string.IsNullOrWhiteSpace(item.Name) ? fi.Name : SanitizeFileName(item.Name!);

        return new PreparedAttachment
        {
            Name = name,
            ContentType = string.IsNullOrWhiteSpace(item.ContentType) ? MimeTypes.FromFileName(name) : item.ContentType!.Trim(),
            Bytes = bytes,
            ContentId = item.ContentId,
            IsInline = item.IsInline,
        };
    }

    private static string SanitizeFileName(string name)
    {
        // Strip any directory component and characters Windows/Exchange reject.
        var trimmed = name.Trim();
        var slash = trimmed.LastIndexOfAny(new[] { '/', '\\' });
        if (slash >= 0) trimmed = trimmed[(slash + 1)..];
        foreach (var c in Path.GetInvalidFileNameChars())
            trimmed = trimmed.Replace(c, '_');
        return string.IsNullOrWhiteSpace(trimmed) ? "attachment" : trimmed;
    }
}

public static class MimeTypes
{
    private static readonly Dictionary<string, string> Map = new(StringComparer.OrdinalIgnoreCase)
    {
        [".txt"] = "text/plain",
        [".csv"] = "text/csv",
        [".htm"] = "text/html",
        [".html"] = "text/html",
        [".xml"] = "application/xml",
        [".json"] = "application/json",
        [".ics"] = "text/calendar",
        [".pdf"] = "application/pdf",
        [".doc"] = "application/msword",
        [".docx"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        [".xls"] = "application/vnd.ms-excel",
        [".xlsx"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        [".ppt"] = "application/vnd.ms-powerpoint",
        [".pptx"] = "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        [".jpg"] = "image/jpeg",
        [".jpeg"] = "image/jpeg",
        [".png"] = "image/png",
        [".gif"] = "image/gif",
        [".bmp"] = "image/bmp",
        [".webp"] = "image/webp",
        [".svg"] = "image/svg+xml",
        [".tif"] = "image/tiff",
        [".tiff"] = "image/tiff",
        [".zip"] = "application/zip",
        [".7z"] = "application/x-7z-compressed",
        [".rtf"] = "application/rtf",
        [".mp3"] = "audio/mpeg",
        [".mp4"] = "video/mp4",
    };

    public static string FromFileName(string fileName)
    {
        var ext = Path.GetExtension(fileName);
        return !string.IsNullOrEmpty(ext) && Map.TryGetValue(ext, out var mime) ? mime : "application/octet-stream";
    }
}
