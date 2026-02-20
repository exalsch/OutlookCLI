using System.CommandLine;
using OutlookCLI.Configuration;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;

namespace OutlookCLI.Commands.Calendar;

public class GetEventCommand : Command
{
    public GetEventCommand() : base("get", "Get full details of a specific calendar event including body, attendees, and organizer.")
    {
        var entryIdArg = new Argument<string>(
            "entry-id",
            "The Entry ID of the event (from 'calendar list' output)");

        AddArgument(entryIdArg);

        this.SetHandler(Execute, entryIdArg);
    }

    private void Execute(string entryId)
    {
        var options = GlobalOptionsAccessor.Current;
        IOutputFormatter formatter = options.Human ? new HumanOutputFormatter() : new JsonOutputFormatter();

        using var service = new OutlookService();
        try
        {
            service.Initialize();
            var evt = service.GetEvent(entryId);

            if (evt == null)
            {
                var errorResult = CommandResult<CalendarEvent>.Fail(
                    "calendar get",
                    "NOT_FOUND",
                    $"Event with Entry ID '{entryId}' not found"
                );
                Console.WriteLine(formatter.FormatEvent(errorResult));
                Environment.Exit(1);
                return;
            }

            var result = CommandResult<CalendarEvent>.Ok("calendar get", evt);
            Console.WriteLine(formatter.FormatEvent(result));
        }
        catch (Exception ex)
        {
            var result = CommandResult<CalendarEvent>.Fail(
                "calendar get",
                "OUTLOOK_ERROR",
                ex.Message
            );
            Console.WriteLine(formatter.FormatEvent(result));
            Environment.Exit(1);
        }
    }
}
