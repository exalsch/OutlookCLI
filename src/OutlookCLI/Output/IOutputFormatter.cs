using OutlookCLI.Models;

namespace OutlookCLI.Output;

public interface IOutputFormatter
{
    string Format<T>(CommandResult<T> result);
    string FormatMailList(CommandResult<IEnumerable<MailMessageSummary>> result);
    string FormatMailListFull(CommandResult<IEnumerable<MailMessage>> result);
    string FormatMail(CommandResult<MailMessage> result);
    string FormatEventList(CommandResult<IEnumerable<CalendarEventSummary>> result);
    string FormatEventListFull(CommandResult<IEnumerable<CalendarEvent>> result);
    string FormatEvent(CommandResult<CalendarEvent> result);
    string FormatFolders(CommandResult<IEnumerable<FolderInfo>> result);
}
