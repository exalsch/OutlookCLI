using System.CommandLine;
using OutlookCLI.Configuration;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;

namespace OutlookCLI.Commands.Mail;

public class SearchMailCommand : Command
{
    public SearchMailCommand() : base("search", "Search emails by query, sender, date range, or follow-up flag. At least one filter is required. Returns same fields as 'mail list'.")
    {
        var queryOption = new Option<string?>(
            ["--query", "-q"],
            "Search text in subject and body (case-insensitive substring match)");

        var fromOption = new Option<string?>(
            ["--from"],
            "Filter by sender email address (e.g. user@example.com)");

        var afterOption = new Option<DateTime?>(
            ["--after"],
            "Show emails received after this date (format: yyyy-MM-dd, e.g. 2024-01-15)");

        var beforeOption = new Option<DateTime?>(
            ["--before"],
            "Show emails received before this date (format: yyyy-MM-dd, e.g. 2024-06-30)");

        var folderOption = new Option<string?>(
            ["--folder", "-f"],
            "Folder to search in (default: Inbox). Use 'mail folders' for available names");

        var flaggedOption = new Option<bool>(
            ["--flagged"],
            "Show only messages flagged for follow-up (flag set, not completed). Same as --flag-status flagged");

        var flagStatusOption = new Option<string?>(
            ["--flag-status"],
            "Show only messages with this follow-up flag state: none, flagged, or complete");
        flagStatusOption.FromAmong(MailFlags.StatusNames);

        AddOption(queryOption);
        AddOption(fromOption);
        AddOption(afterOption);
        AddOption(beforeOption);
        AddOption(folderOption);
        AddOption(flaggedOption);
        AddOption(flagStatusOption);

        this.SetHandler(Execute, queryOption, fromOption, afterOption, beforeOption, folderOption, flaggedOption, flagStatusOption);
    }

    private void Execute(string? query, string? from, DateTime? after, DateTime? before, string? folder, bool flagged, string? flagStatus)
    {
        var options = GlobalOptionsAccessor.Current;
        IOutputFormatter formatter = options.Human ? new HumanOutputFormatter() : new JsonOutputFormatter();

        if (!MailFlags.TryResolveFilter(flagged, flagStatus, out var flagFilter, out var flagError))
        {
            var errorResult = CommandResult<IEnumerable<MailMessageSummary>>.Fail(
                "mail search",
                "INVALID_ARGS",
                flagError!
            );
            Console.WriteLine(formatter.FormatMailList(errorResult));
            Environment.Exit(1);
            return;
        }

        if (string.IsNullOrEmpty(query) && string.IsNullOrEmpty(from) && !after.HasValue && !before.HasValue && !flagFilter.HasValue)
        {
            var errorResult = CommandResult<IEnumerable<MailMessageSummary>>.Fail(
                "mail search",
                "INVALID_ARGS",
                "At least one search criteria is required: --query, --from, --after, --before, --flagged, or --flag-status"
            );
            Console.WriteLine(formatter.FormatMailList(errorResult));
            Environment.Exit(1);
            return;
        }

        using var service = new OutlookService();
        try
        {
            service.Initialize();
            var messages = service.SearchMail(query, from, after, before, folder, flagFilter).ToList();

            var result = CommandResult<IEnumerable<MailMessageSummary>>.Ok(
                "mail search",
                messages,
                new ResultMetadata { Count = messages.Count }
            );

            Console.WriteLine(formatter.FormatMailList(result));
        }
        catch (Exception ex)
        {
            var result = CommandResult<IEnumerable<MailMessageSummary>>.Fail(
                "mail search",
                "OUTLOOK_ERROR",
                ex.Message
            );
            Console.WriteLine(formatter.FormatMailList(result));
            Environment.Exit(1);
        }
    }
}
