using System.Text;

namespace MailGateway;

/// <summary>
/// Minimal daily-rolling file logger, same line format as the original gateway:
///   2026-06-19 17:43:23.347 UTC | INFO | {requestId} | message
/// Writes are serialized so concurrent requests never interleave half-lines.
/// </summary>
public sealed class FileLogger
{
    private readonly string _folder;
    private readonly SemaphoreSlim _gate = new(1, 1);
    private readonly ILogger<FileLogger> _console;

    public FileLogger(MailGatewayOptions options, ILogger<FileLogger> console)
    {
        _console = console;
        _folder = string.IsNullOrWhiteSpace(options.Logging.Folder)
            ? Path.Combine(AppContext.BaseDirectory, "logs")
            : options.Logging.Folder!;
        try { Directory.CreateDirectory(_folder); }
        catch (Exception ex) { _console.LogWarning(ex, "Could not create log folder {Folder}", _folder); }
    }

    public string Folder => _folder;

    public Task InfoAsync(string requestId, string message) => WriteAsync("INFO", requestId, message);
    public Task WarnAsync(string requestId, string message) => WriteAsync("WARN", requestId, message);
    public Task ErrorAsync(string requestId, string message) => WriteAsync("ERROR", requestId, message);

    public async Task WriteAsync(string level, string requestId, string message)
    {
        var now = DateTime.UtcNow;
        var line = $"{now:yyyy-MM-dd HH:mm:ss.fff} UTC | {level} | {requestId} | {message}{Environment.NewLine}";
        var file = Path.Combine(_folder, $"mailgateway-{now:yyyyMMdd}.log");

        await _gate.WaitAsync();
        try
        {
            await File.AppendAllTextAsync(file, line, Encoding.UTF8);
        }
        catch (Exception ex)
        {
            _console.LogWarning(ex, "Could not write to log file {File}", file);
        }
        finally
        {
            _gate.Release();
        }

        switch (level)
        {
            case "ERROR": _console.LogError("{RequestId} {Message}", requestId, message); break;
            case "WARN": _console.LogWarning("{RequestId} {Message}", requestId, message); break;
            default: _console.LogInformation("{RequestId} {Message}", requestId, message); break;
        }
    }
}
