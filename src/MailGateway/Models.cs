using System.Text.Json;
using System.Text.Json.Serialization;

namespace MailGateway;

/// <summary>
/// Inbound payload for POST /api/mail/send. Property names are case-insensitive.
/// Recipient fields accept either a single string ("a@x.com; b@y.com") or a JSON array.
/// Attachment fields accept strings (file paths) or objects (see <see cref="MailAttachmentItem"/>).
/// </summary>
public sealed class MailSendRequest
{
    /// <summary>Free-text name of the calling application; only used in logs.</summary>
    public string? AppName { get; set; }

    public string? From { get; set; }
    public string? FromName { get; set; }

    [JsonConverter(typeof(FlexibleStringListConverter))]
    public List<string>? To { get; set; }

    [JsonConverter(typeof(FlexibleStringListConverter))]
    public List<string>? Cc { get; set; }

    [JsonConverter(typeof(FlexibleStringListConverter))]
    public List<string>? Bcc { get; set; }

    [JsonConverter(typeof(FlexibleStringListConverter))]
    public List<string>? ReplyTo { get; set; }

    public string? Subject { get; set; }
    public string? HtmlBody { get; set; }
    public string? TextBody { get; set; }

    /// <summary>Original field: list of file paths (strings). Now also accepts attachment objects.</summary>
    [JsonConverter(typeof(FlexibleAttachmentListConverter))]
    public List<MailAttachmentItem>? Attachments { get; set; }

    /// <summary>Original field: list of attachment objects. Kept for compatibility; merged with <see cref="Attachments"/>.</summary>
    [JsonConverter(typeof(FlexibleAttachmentListConverter))]
    public List<MailAttachmentItem>? AttachmentItems { get; set; }

    public IEnumerable<MailAttachmentItem> AllAttachments()
    {
        if (Attachments is not null) foreach (var a in Attachments) yield return a;
        if (AttachmentItems is not null) foreach (var a in AttachmentItems) yield return a;
    }
}

/// <summary>
/// One attachment. Provide EITHER <see cref="Path"/> (file on the gateway server, must be under MailLimits:AllowedAttachmentRoot)
/// OR <see cref="ContentBase64"/> (the file bytes, base64-encoded; a "data:...;base64," prefix is tolerated).
/// </summary>
public sealed class MailAttachmentItem
{
    public string? Path { get; set; }

    /// <summary>File name shown to the recipient. Defaults to the file name of <see cref="Path"/>. Required with <see cref="ContentBase64"/>.</summary>
    public string? Name { get; set; }

    /// <summary>MIME type. Guessed from the extension when omitted.</summary>
    public string? ContentType { get; set; }

    public string? ContentBase64 { get; set; }

    /// <summary>For inline images: the value referenced by &lt;img src="cid:..."&gt; in htmlBody.</summary>
    public string? ContentId { get; set; }

    public bool IsInline { get; set; }
}

/// <summary>Attachment after validation: decoded bytes ready for Graph.</summary>
public sealed class PreparedAttachment
{
    public required string Name { get; init; }
    public required string ContentType { get; init; }
    public required byte[] Bytes { get; init; }
    public string? ContentId { get; init; }
    public bool IsInline { get; init; }
}

public sealed record ApiOk(bool Success, string RequestId, string Message, string Mode, string From);
public sealed record ApiError(bool Success, string RequestId, string Error);

/// <summary>Accepts "a; b, c" or ["a","b"] (or null) and yields a list of strings.</summary>
public sealed class FlexibleStringListConverter : JsonConverter<List<string>?>
{
    public override List<string>? Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        switch (reader.TokenType)
        {
            case JsonTokenType.Null:
                return null;
            case JsonTokenType.String:
                return new List<string> { reader.GetString() ?? "" };
            case JsonTokenType.StartArray:
                var list = new List<string>();
                while (reader.Read())
                {
                    if (reader.TokenType == JsonTokenType.EndArray) break;
                    if (reader.TokenType == JsonTokenType.String) list.Add(reader.GetString() ?? "");
                    else if (reader.TokenType == JsonTokenType.Null) continue;
                    else throw new JsonException("Expected a string inside the array.");
                }
                return list;
            default:
                throw new JsonException("Expected a string or an array of strings.");
        }
    }

    public override void Write(Utf8JsonWriter writer, List<string>? value, JsonSerializerOptions options)
    {
        if (value is null) { writer.WriteNullValue(); return; }
        writer.WriteStartArray();
        foreach (var s in value) writer.WriteStringValue(s);
        writer.WriteEndArray();
    }
}

/// <summary>Accepts "path", ["path", {...}], or {...} and yields attachment items.</summary>
public sealed class FlexibleAttachmentListConverter : JsonConverter<List<MailAttachmentItem>?>
{
    public override List<MailAttachmentItem>? Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        switch (reader.TokenType)
        {
            case JsonTokenType.Null:
                return null;
            case JsonTokenType.String:
                return new List<MailAttachmentItem> { new() { Path = reader.GetString() } };
            case JsonTokenType.StartObject:
                return new List<MailAttachmentItem> { ReadItem(ref reader, options) };
            case JsonTokenType.StartArray:
                var list = new List<MailAttachmentItem>();
                while (reader.Read())
                {
                    if (reader.TokenType == JsonTokenType.EndArray) break;
                    if (reader.TokenType == JsonTokenType.String) list.Add(new MailAttachmentItem { Path = reader.GetString() });
                    else if (reader.TokenType == JsonTokenType.StartObject) list.Add(ReadItem(ref reader, options));
                    else if (reader.TokenType == JsonTokenType.Null) continue;
                    else throw new JsonException("Expected a string or an object inside the attachments array.");
                }
                return list;
            default:
                throw new JsonException("Expected a string, an object, or an array for attachments.");
        }
    }

    private static MailAttachmentItem ReadItem(ref Utf8JsonReader reader, JsonSerializerOptions options)
    {
        // Deserialize the object without this converter to avoid recursion.
        var item = JsonSerializer.Deserialize<MailAttachmentItem>(ref reader, options);
        return item ?? new MailAttachmentItem();
    }

    public override void Write(Utf8JsonWriter writer, List<MailAttachmentItem>? value, JsonSerializerOptions options)
    {
        if (value is null) { writer.WriteNullValue(); return; }
        writer.WriteStartArray();
        foreach (var item in value) JsonSerializer.Serialize(writer, item, options);
        writer.WriteEndArray();
    }
}
