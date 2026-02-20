using System.Text.Json;
using System.Text.Json.Serialization;
using OutlookCLI.Models;

namespace OutlookCLI.Output;

public class JsonOutputFormatter : IOutputFormatter
{
    private static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    public string Format<T>(CommandResult<T> result)
    {
        return JsonSerializer.Serialize(result, Options);
    }

    public string FormatMailList(CommandResult<IEnumerable<MailMessageSummary>> result)
    {
        return Format(result);
    }

    public string FormatMailListFull(CommandResult<IEnumerable<MailMessage>> result)
    {
        return Format(result);
    }

    public string FormatMail(CommandResult<MailMessage> result)
    {
        return Format(result);
    }

    public string FormatEventList(CommandResult<IEnumerable<CalendarEventSummary>> result)
    {
        return Format(result);
    }

    public string FormatEventListFull(CommandResult<IEnumerable<CalendarEvent>> result)
    {
        return Format(result);
    }

    public string FormatEvent(CommandResult<CalendarEvent> result)
    {
        return Format(result);
    }

    public string FormatFolders(CommandResult<IEnumerable<FolderInfo>> result)
    {
        return Format(result);
    }
}
