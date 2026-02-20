using System.CommandLine;
using OutlookCLI.Configuration;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;

namespace OutlookCLI.Commands.Calendar;

public class ListEventsCommand : Command
{
    public ListEventsCommand() : base("list", "List calendar events in a date range. Returns entryId, subject, start, end, location, isAllDay. Use --full to include body.")
    {
        var startOption = new Option<DateTime?>(
            ["--start", "-s"],
            "Start date for event range (format: yyyy-MM-dd, e.g. 2024-01-15. Default: today)");

        var endOption = new Option<DateTime?>(
            ["--end", "-e"],
            "End date for event range (format: yyyy-MM-dd, e.g. 2024-02-15. Default: 1 month from start)");

        var limitOption = new Option<int>(
            ["--limit", "-l"],
            () => 50,
            "Maximum number of events to return (chronological order)");

        AddOption(startOption);
        AddOption(endOption);
        AddOption(limitOption);

        this.SetHandler(Execute, startOption, endOption, limitOption);
    }

    private void Execute(DateTime? start, DateTime? end, int limit)
    {
        var options = GlobalOptionsAccessor.Current;
        IOutputFormatter formatter = options.Human ? new HumanOutputFormatter() : new JsonOutputFormatter();

        using var service = new OutlookService();
        try
        {
            service.Initialize();

            if (options.Full)
            {
                var events = service.GetEventListFull(start, end, limit).ToList();
                var result = CommandResult<IEnumerable<CalendarEvent>>.Ok(
                    "calendar list",
                    events,
                    new ResultMetadata { Count = events.Count }
                );
                Console.WriteLine(formatter.FormatEventListFull(result));
            }
            else
            {
                var events = service.GetEventList(start, end, limit).ToList();
                var result = CommandResult<IEnumerable<CalendarEventSummary>>.Ok(
                    "calendar list",
                    events,
                    new ResultMetadata { Count = events.Count }
                );
                Console.WriteLine(formatter.FormatEventList(result));
            }
        }
        catch (Exception ex)
        {
            var result = CommandResult<IEnumerable<CalendarEventSummary>>.Fail(
                "calendar list",
                "OUTLOOK_ERROR",
                ex.Message
            );
            Console.WriteLine(formatter.FormatEventList(result));
            Environment.Exit(1);
        }
    }
}
