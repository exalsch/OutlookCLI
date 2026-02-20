using System.CommandLine;
using OutlookCLI.Configuration;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;

namespace OutlookCLI.Commands.Calendar;

public class UpdateEventCommand : Command
{
    public UpdateEventCommand() : base("update", "Update an existing calendar event. Only specified fields are changed; omitted fields keep their current value.")
    {
        var entryIdArg = new Argument<string>(
            "entry-id",
            "The Entry ID of the event to update (from 'calendar list' output)");

        var subjectOption = new Option<string?>(
            ["--subject", "-s"],
            "New event subject/title");

        var startOption = new Option<DateTime?>(
            ["--start"],
            "New start date/time (format: \"yyyy-MM-dd HH:mm\")");

        var endOption = new Option<DateTime?>(
            ["--end"],
            "New end date/time (format: \"yyyy-MM-dd HH:mm\")");

        var locationOption = new Option<string?>(
            ["--location", "-l"],
            "New event location");

        var bodyOption = new Option<string?>(
            ["--body", "-b"],
            "New event description/body");

        AddArgument(entryIdArg);
        AddOption(subjectOption);
        AddOption(startOption);
        AddOption(endOption);
        AddOption(locationOption);
        AddOption(bodyOption);

        this.SetHandler(Execute, entryIdArg, subjectOption, startOption, endOption, locationOption, bodyOption);
    }

    private void Execute(string entryId, string? subject, DateTime? start, DateTime? end, string? location, string? body)
    {
        var options = GlobalOptionsAccessor.Current;
        IOutputFormatter formatter = options.Human ? new HumanOutputFormatter() : new JsonOutputFormatter();

        if (subject == null && start == null && end == null && location == null && body == null)
        {
            var errorResult = CommandResult<object>.Fail(
                "calendar update",
                "NO_UPDATES",
                "At least one update parameter is required: --subject, --start, --end, --location, or --body"
            );
            Console.WriteLine(formatter.Format(errorResult));
            Environment.Exit(1);
            return;
        }

        using var service = new OutlookService();
        try
        {
            service.Initialize();

            // First verify the event exists
            var existingEvent = service.GetEvent(entryId);
            if (existingEvent == null)
            {
                var errorResult = CommandResult<object>.Fail(
                    "calendar update",
                    "NOT_FOUND",
                    $"Event with Entry ID '{entryId}' not found"
                );
                Console.WriteLine(formatter.Format(errorResult));
                Environment.Exit(1);
                return;
            }

            service.UpdateEvent(entryId, subject, start, end, location, body);

            var result = CommandResult<object>.Ok(
                "calendar update",
                new
                {
                    message = "Event updated successfully",
                    entryId,
                    updates = new
                    {
                        subject = subject != null,
                        start = start != null,
                        end = end != null,
                        location = location != null,
                        body = body != null
                    }
                },
                new ResultMetadata()
            );

            Console.WriteLine(formatter.Format(result));
        }
        catch (Exception ex)
        {
            var result = CommandResult<object>.Fail(
                "calendar update",
                "OUTLOOK_ERROR",
                ex.Message
            );
            Console.WriteLine(formatter.Format(result));
            Environment.Exit(1);
        }
    }
}
