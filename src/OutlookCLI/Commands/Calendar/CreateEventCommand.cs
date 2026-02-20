using System.CommandLine;
using OutlookCLI.Configuration;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;

namespace OutlookCLI.Commands.Calendar;

public class CreateEventCommand : Command
{
    public CreateEventCommand() : base("create", "Create a new calendar event. Returns the entryId of the created event. End must be after start.")
    {
        var subjectOption = new Option<string>(
            ["--subject", "-s"],
            "Event subject/title")
        { IsRequired = true };

        var startOption = new Option<DateTime>(
            ["--start"],
            "Event start date/time (format: \"yyyy-MM-dd HH:mm\" e.g. \"2024-01-15 09:00\", or yyyy-MM-dd for all-day)")
        { IsRequired = true };

        var endOption = new Option<DateTime>(
            ["--end"],
            "Event end date/time (format: \"yyyy-MM-dd HH:mm\" e.g. \"2024-01-15 10:00\"). Must be after --start")
        { IsRequired = true };

        var locationOption = new Option<string?>(
            ["--location", "-l"],
            "Event location (room name, address, or Teams/Zoom link)");

        var bodyOption = new Option<string?>(
            ["--body", "-b"],
            "Event description/body (plain text)");

        var allDayOption = new Option<bool>(
            ["--all-day"],
            "Create as an all-day event (only date part of --start/--end is used)");

        AddOption(subjectOption);
        AddOption(startOption);
        AddOption(endOption);
        AddOption(locationOption);
        AddOption(bodyOption);
        AddOption(allDayOption);

        this.SetHandler(Execute, subjectOption, startOption, endOption, locationOption, bodyOption, allDayOption);
    }

    private void Execute(string subject, DateTime start, DateTime end, string? location, string? body, bool allDay)
    {
        var options = GlobalOptionsAccessor.Current;
        IOutputFormatter formatter = options.Human ? new HumanOutputFormatter() : new JsonOutputFormatter();

        if (end <= start)
        {
            var errorResult = CommandResult<object>.Fail(
                "calendar create",
                "INVALID_DATES",
                "End date/time must be after start date/time"
            );
            Console.WriteLine(formatter.Format(errorResult));
            Environment.Exit(1);
            return;
        }

        using var service = new OutlookService();
        try
        {
            service.Initialize();
            var entryId = service.CreateEvent(subject, start, end, location, body, allDay);

            var result = CommandResult<object>.Ok(
                "calendar create",
                new
                {
                    message = "Event created successfully",
                    entryId,
                    subject,
                    start,
                    end,
                    location,
                    isAllDay = allDay
                },
                new ResultMetadata()
            );

            Console.WriteLine(formatter.Format(result));
        }
        catch (Exception ex)
        {
            var result = CommandResult<object>.Fail(
                "calendar create",
                "OUTLOOK_ERROR",
                ex.Message
            );
            Console.WriteLine(formatter.Format(result));
            Environment.Exit(1);
        }
    }
}
