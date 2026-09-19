using System.Text.Json;

namespace MailGateway;

/// <summary>
/// Client-credentials token for Microsoft Graph, cached in memory until shortly before expiry.
/// The original gateway requested a new token on every send; this removes that round-trip.
/// </summary>
public sealed class GraphTokenProvider
{
    private readonly IHttpClientFactory _http;
    private readonly GraphOptions _graph;
    private readonly SemaphoreSlim _gate = new(1, 1);

    private string? _token;
    private DateTimeOffset _expiresAt = DateTimeOffset.MinValue;

    public GraphTokenProvider(IHttpClientFactory http, MailGatewayOptions options)
    {
        _http = http;
        _graph = options.Graph;
    }

    public async Task<string> GetAccessTokenAsync(CancellationToken ct = default)
    {
        // Fast path: cached and still valid for at least 2 minutes.
        var cached = _token;
        if (cached is not null && DateTimeOffset.UtcNow.AddMinutes(2) < _expiresAt)
            return cached;

        await _gate.WaitAsync(ct);
        try
        {
            cached = _token;
            if (cached is not null && DateTimeOffset.UtcNow.AddMinutes(2) < _expiresAt)
                return cached;

            var client = _http.CreateClient("token");
            var url = $"https://login.microsoftonline.com/{Uri.EscapeDataString(_graph.TenantId)}/oauth2/v2.0/token";

            using var request = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new FormUrlEncodedContent(new Dictionary<string, string>
                {
                    ["client_id"] = _graph.ClientId,
                    ["client_secret"] = _graph.ClientSecret,
                    ["scope"] = "https://graph.microsoft.com/.default",
                    ["grant_type"] = "client_credentials",
                })
            };

            using var response = await client.SendAsync(request, ct);
            var body = await response.Content.ReadAsStringAsync(ct);

            if (!response.IsSuccessStatusCode)
                throw new GraphException($"Graph token request failed ({(int)response.StatusCode}): {Truncate(body)}");

            using var doc = JsonDocument.Parse(body);
            var root = doc.RootElement;

            if (!root.TryGetProperty("access_token", out var tokenEl))
                throw new GraphException("Graph token response did not include access_token.");

            var token = tokenEl.GetString();
            if (string.IsNullOrWhiteSpace(token))
                throw new GraphException("Graph access token was empty.");

            var expiresIn = 3600;
            if (root.TryGetProperty("expires_in", out var expEl))
            {
                if (expEl.ValueKind == JsonValueKind.Number && expEl.TryGetInt32(out var n)) expiresIn = n;
                else if (expEl.ValueKind == JsonValueKind.String && int.TryParse(expEl.GetString(), out var s)) expiresIn = s;
            }

            _token = token;
            _expiresAt = DateTimeOffset.UtcNow.AddSeconds(expiresIn);
            return _token;
        }
        finally
        {
            _gate.Release();
        }
    }

    /// <summary>Drops the cached token (used after a 401 from Graph so the next call fetches a fresh one).</summary>
    public void Invalidate()
    {
        _expiresAt = DateTimeOffset.MinValue;
        _token = null;
    }

    private static string Truncate(string s, int max = 1000) => s.Length <= max ? s : s[..max] + "…";
}

public sealed class GraphException : Exception
{
    public int? StatusCode { get; }
    public string? GraphErrorCode { get; }

    public GraphException(string message, int? statusCode = null, string? graphErrorCode = null) : base(message)
    {
        StatusCode = statusCode;
        GraphErrorCode = graphErrorCode;
    }
}
