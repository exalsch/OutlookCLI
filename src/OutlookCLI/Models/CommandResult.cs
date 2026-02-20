using System.Text.Json.Serialization;

namespace OutlookCLI.Models;

public class CommandResult<T>
{
    public bool Success { get; init; }
    public string Command { get; init; } = string.Empty;
    public T? Data { get; init; }
    public ErrorInfo? Error { get; init; }
    public ResultMetadata? Metadata { get; init; }

    public static CommandResult<T> Ok(string command, T data, ResultMetadata? metadata = null)
    {
        return new CommandResult<T>
        {
            Success = true,
            Command = command,
            Data = data,
            Error = null,
            Metadata = metadata
        };
    }

    public static CommandResult<T> Fail(string command, string code, string message)
    {
        return new CommandResult<T>
        {
            Success = false,
            Command = command,
            Data = default,
            Error = new ErrorInfo(code, message),
            Metadata = null
        };
    }
}

public record ErrorInfo(string Code, string Message);

public record ResultMetadata
{
    public int? Count { get; init; }
    public DateTime Timestamp { get; init; } = DateTime.UtcNow;

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? AdditionalInfo { get; init; }
}
