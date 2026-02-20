using System.Text;
using OutlookCLI.Models;

namespace OutlookCLI.Output;

public class HumanOutputFormatter : IOutputFormatter
{
    public string Format<T>(CommandResult<T> result)
    {
        if (!result.Success)
        {
            return $"Error: [{result.Error?.Code}] {result.Error?.Message}";
        }

        return result.Data?.ToString() ?? "No data";
    }

    public string FormatMailList(CommandResult<IEnumerable<MailMessageSummary>> result)
    {
        if (!result.Success)
        {
            return $"Error: [{result.Error?.Code}] {result.Error?.Message}";
        }

        var sb = new StringBuilder();
        var messages = result.Data?.ToList() ?? [];

        sb.AppendLine($"Mail ({messages.Count} messages)");
        sb.AppendLine(new string('─', 50));

        int index = 1;
        foreach (var msg in messages)
        {
            var unreadMarker = msg.IsUnread ? "[UNREAD] " : "";
            var attachmentMarker = msg.HasAttachments ? " 📎" : "";
            sb.AppendLine($"{index}. {unreadMarker}{msg.Subject}{attachmentMarker}");
            sb.AppendLine($"   From: {msg.SenderName} <{msg.SenderEmail}>");
            sb.AppendLine($"   Date: {msg.ReceivedTime:MMM dd, yyyy h:mm tt}");
            sb.AppendLine($"   ID: {msg.EntryId[..Math.Min(20, msg.EntryId.Length)]}...");
            sb.AppendLine();
            index++;
        }

        return sb.ToString().TrimEnd();
    }

    public string FormatMailListFull(CommandResult<IEnumerable<MailMessage>> result)
    {
        if (!result.Success)
        {
            return $"Error: [{result.Error?.Code}] {result.Error?.Message}";
        }

        var sb = new StringBuilder();
        var messages = result.Data?.ToList() ?? [];

        sb.AppendLine($"Mail ({messages.Count} messages)");
        sb.AppendLine(new string('─', 50));

        int index = 1;
        foreach (var msg in messages)
        {
            var unreadMarker = msg.IsUnread ? "[UNREAD] " : "";
            sb.AppendLine($"{index}. {unreadMarker}{msg.Subject}");
            sb.AppendLine($"   From: {msg.SenderName} <{msg.SenderEmail}>");
            sb.AppendLine($"   To: {string.Join(", ", msg.To)}");
            if (msg.Cc.Count > 0)
                sb.AppendLine($"   Cc: {string.Join(", ", msg.Cc)}");
            sb.AppendLine($"   Date: {msg.ReceivedTime:MMM dd, yyyy h:mm tt}");
            sb.AppendLine($"   ID: {msg.EntryId[..Math.Min(20, msg.EntryId.Length)]}...");
            sb.AppendLine();
            sb.AppendLine($"   Body:");
            var bodyLines = msg.Body.Split('\n').Take(10);
            foreach (var line in bodyLines)
            {
                sb.AppendLine($"   {line.TrimEnd()}");
            }
            sb.AppendLine();
            sb.AppendLine(new string('─', 50));
            index++;
        }

        return sb.ToString().TrimEnd();
    }

    public string FormatMail(CommandResult<MailMessage> result)
    {
        if (!result.Success)
        {
            return $"Error: [{result.Error?.Code}] {result.Error?.Message}";
        }

        var msg = result.Data;
        if (msg == null) return "No message found";

        var sb = new StringBuilder();
        sb.AppendLine(new string('─', 50));
        sb.AppendLine($"Subject: {msg.Subject}");
        sb.AppendLine($"From: {msg.SenderName} <{msg.SenderEmail}>");
        sb.AppendLine($"To: {string.Join(", ", msg.To)}");
        if (msg.Cc.Count > 0)
            sb.AppendLine($"Cc: {string.Join(", ", msg.Cc)}");
        sb.AppendLine($"Date: {msg.ReceivedTime:MMM dd, yyyy h:mm tt}");
        sb.AppendLine($"Folder: {msg.Folder}");
        sb.AppendLine($"Status: {(msg.IsUnread ? "Unread" : "Read")}");

        if (msg.HasAttachments)
        {
            sb.AppendLine($"Attachments ({msg.AttachmentCount}):");
            foreach (var att in msg.Attachments)
            {
                sb.AppendLine($"  - {att.FileName} ({FormatSize(att.Size)})");
            }
        }

        sb.AppendLine(new string('─', 50));
        sb.AppendLine();
        sb.AppendLine(msg.Body);

        return sb.ToString().TrimEnd();
    }

    public string FormatEventList(CommandResult<IEnumerable<CalendarEventSummary>> result)
    {
        if (!result.Success)
        {
            return $"Error: [{result.Error?.Code}] {result.Error?.Message}";
        }

        var sb = new StringBuilder();
        var events = result.Data?.ToList() ?? [];

        sb.AppendLine($"Calendar ({events.Count} events)");
        sb.AppendLine(new string('─', 50));

        int index = 1;
        foreach (var evt in events)
        {
            var recurringMarker = evt.IsRecurring ? " 🔁" : "";
            var allDayMarker = evt.IsAllDay ? " [All Day]" : "";
            sb.AppendLine($"{index}. {evt.Subject}{recurringMarker}{allDayMarker}");
            if (evt.IsAllDay)
            {
                sb.AppendLine($"   Date: {evt.Start:MMM dd, yyyy}");
            }
            else
            {
                sb.AppendLine($"   Start: {evt.Start:MMM dd, yyyy h:mm tt}");
                sb.AppendLine($"   End: {evt.End:MMM dd, yyyy h:mm tt}");
            }
            if (!string.IsNullOrEmpty(evt.Location))
                sb.AppendLine($"   Location: {evt.Location}");
            sb.AppendLine($"   ID: {evt.EntryId[..Math.Min(20, evt.EntryId.Length)]}...");
            sb.AppendLine();
            index++;
        }

        return sb.ToString().TrimEnd();
    }

    public string FormatEventListFull(CommandResult<IEnumerable<CalendarEvent>> result)
    {
        if (!result.Success)
        {
            return $"Error: [{result.Error?.Code}] {result.Error?.Message}";
        }

        var sb = new StringBuilder();
        var events = result.Data?.ToList() ?? [];

        sb.AppendLine($"Calendar ({events.Count} events)");
        sb.AppendLine(new string('─', 50));

        int index = 1;
        foreach (var evt in events)
        {
            sb.AppendLine($"{index}. {evt.Subject}");
            if (evt.IsAllDay)
            {
                sb.AppendLine($"   Date: {evt.Start:MMM dd, yyyy} [All Day]");
            }
            else
            {
                sb.AppendLine($"   Start: {evt.Start:MMM dd, yyyy h:mm tt}");
                sb.AppendLine($"   End: {evt.End:MMM dd, yyyy h:mm tt}");
            }
            if (!string.IsNullOrEmpty(evt.Location))
                sb.AppendLine($"   Location: {evt.Location}");
            if (!string.IsNullOrEmpty(evt.Organizer))
                sb.AppendLine($"   Organizer: {evt.Organizer}");
            if (evt.Attendees.Count > 0)
                sb.AppendLine($"   Attendees: {string.Join(", ", evt.Attendees)}");
            if (evt.IsRecurring && !string.IsNullOrEmpty(evt.RecurrencePattern))
                sb.AppendLine($"   Recurrence: {evt.RecurrencePattern}");
            sb.AppendLine($"   ID: {evt.EntryId[..Math.Min(20, evt.EntryId.Length)]}...");
            if (!string.IsNullOrEmpty(evt.Body))
            {
                sb.AppendLine();
                sb.AppendLine($"   Description:");
                var bodyLines = evt.Body.Split('\n').Take(5);
                foreach (var line in bodyLines)
                {
                    sb.AppendLine($"   {line.TrimEnd()}");
                }
            }
            sb.AppendLine();
            sb.AppendLine(new string('─', 50));
            index++;
        }

        return sb.ToString().TrimEnd();
    }

    public string FormatEvent(CommandResult<CalendarEvent> result)
    {
        if (!result.Success)
        {
            return $"Error: [{result.Error?.Code}] {result.Error?.Message}";
        }

        var evt = result.Data;
        if (evt == null) return "No event found";

        var sb = new StringBuilder();
        sb.AppendLine(new string('─', 50));
        sb.AppendLine($"Subject: {evt.Subject}");
        if (evt.IsAllDay)
        {
            sb.AppendLine($"Date: {evt.Start:MMM dd, yyyy} [All Day]");
        }
        else
        {
            sb.AppendLine($"Start: {evt.Start:MMM dd, yyyy h:mm tt}");
            sb.AppendLine($"End: {evt.End:MMM dd, yyyy h:mm tt}");
        }
        if (!string.IsNullOrEmpty(evt.Location))
            sb.AppendLine($"Location: {evt.Location}");
        if (!string.IsNullOrEmpty(evt.Organizer))
            sb.AppendLine($"Organizer: {evt.Organizer}");
        if (evt.Attendees.Count > 0)
            sb.AppendLine($"Attendees: {string.Join(", ", evt.Attendees)}");
        if (evt.IsRecurring && !string.IsNullOrEmpty(evt.RecurrencePattern))
            sb.AppendLine($"Recurrence: {evt.RecurrencePattern}");

        sb.AppendLine(new string('─', 50));
        if (!string.IsNullOrEmpty(evt.Body))
        {
            sb.AppendLine();
            sb.AppendLine(evt.Body);
        }

        return sb.ToString().TrimEnd();
    }

    public string FormatFolders(CommandResult<IEnumerable<FolderInfo>> result)
    {
        if (!result.Success)
        {
            return $"Error: [{result.Error?.Code}] {result.Error?.Message}";
        }

        var sb = new StringBuilder();
        var folders = result.Data?.ToList() ?? [];

        sb.AppendLine("Mail Folders");
        sb.AppendLine(new string('─', 50));

        foreach (var folder in folders)
        {
            var unreadInfo = folder.UnreadCount > 0 ? $" ({folder.UnreadCount} unread)" : "";
            sb.AppendLine($"  {folder.FullPath} [{folder.ItemCount} items]{unreadInfo}");
        }

        return sb.ToString().TrimEnd();
    }

    private static string FormatSize(long bytes)
    {
        string[] sizes = ["B", "KB", "MB", "GB"];
        int order = 0;
        double size = bytes;
        while (size >= 1024 && order < sizes.Length - 1)
        {
            order++;
            size /= 1024;
        }
        return $"{size:0.##} {sizes[order]}";
    }
}
