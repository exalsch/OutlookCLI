using System.CommandLine;
using OutlookCLI.Configuration;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;

namespace OutlookCLI.Commands.Mail;

public class ListMailCommand : Command
{
    public ListMailCommand() : base("list", "List emails from a folder. Returns entryId, subject, from, date, isUnread and follow-up flag state (isFlagged, flagStatus, flagRequest, flagDueDate) for each message. Use --full to include body.")
    {
        var folderOption = new Option<string?>(
            ["--folder", "-f"],
            "Folder name to list from (default: Inbox). Use 'mail folders' to discover available folder names");

        var unreadOption = new Option<bool>(
            ["--unread", "-u"],
            "Show only unread messages");

        var limitOption = new Option<int>(
            ["--limit", "-l"],
            () => 20,
            "Maximum number of messages to return (most recent first)");

        var flaggedOption = new Option<bool>(
            ["--flagged"],
            "Show only messages flagged for follow-up (flag set, not completed). Same as --flag-status flagged");

        var flagStatusOption = new Option<string?>(
            ["--flag-status"],
            "Show only messages with this follow-up flag state: none, flagged, or complete");
        flagStatusOption.FromAmong(MailFlags.StatusNames);

        AddOption(folderOption);
        AddOption(unreadOption);
        AddOption(limitOption);
        AddOption(flaggedOption);
        AddOption(flagStatusOption);

        this.SetHandler(Execute, folderOption, unreadOption, limitOption, flaggedOption, flagStatusOption);
    }

    private void Execute(string? folder, bool unreadOnly, int limit, bool flagged, string? flagStatus)
    {
        var options = GlobalOptionsAccessor.Current;
        IOutputFormatter formatter = options.Human ? new HumanOutputFormatter() : new JsonOutputFormatter();

        if (!MailFlags.TryResolveFilter(flagged, flagStatus, out var flagFilter, out var flagError))
        {
            var errorResult = CommandResult<IEnumerable<MailMessageSummary>>.Fail(
                "mail list",
                "INVALID_ARGS",
                flagError!
            );
            Console.WriteLine(formatter.FormatMailList(errorResult));
            Environment.Exit(1);
            return;
        }

        using var service = new OutlookService();
        try
        {
            service.Initialize();

            if (options.Full)
            {
                var messages = service.GetMailListFull(folder, unreadOnly, limit, flagFilter).ToList();
                var result = CommandResult<IEnumerable<MailMessage>>.Ok(
                    "mail list",
                    messages,
                    new ResultMetadata { Count = messages.Count }
                );
                Console.WriteLine(formatter.FormatMailListFull(result));
            }
            else
            {
                var messages = service.GetMailList(folder, unreadOnly, limit, flagFilter).ToList();
                var result = CommandResult<IEnumerable<MailMessageSummary>>.Ok(
                    "mail list",
                    messages,
                    new ResultMetadata { Count = messages.Count }
                );
                Console.WriteLine(formatter.FormatMailList(result));
            }
        }
        catch (Exception ex)
        {
            var result = CommandResult<IEnumerable<MailMessageSummary>>.Fail(
                "mail list",
                "OUTLOOK_ERROR",
                ex.Message
            );
            Console.WriteLine(formatter.FormatMailList(result));
            Environment.Exit(1);
        }
    }
}
