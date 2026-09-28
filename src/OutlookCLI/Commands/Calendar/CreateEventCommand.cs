using System.CommandLine;
using OutlookCLI.Configuration;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;

namespace OutlookCLI.Commands.Calendar;

public class CreateEventCommand : Command
{
    public CreateEventCommand() : base("create", "Create a new calendar event. Returns the entryId of the created event. End must be after start. With --attendees/--optional it is a PLACEHOLDER: a plain appointment whose body lists the invitees. Nothing is ever sent; a human adds the attendees in Outlook and presses Send.")
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

        var attendeesOption = new Option<string[]?>(
            ["--attendees"],
            "Required attendees (email addresses). Space-, comma- or semicolon-separated. Resolved against the address book and listed in the body, never added as recipients")
        { AllowMultipleArgumentsPerToken = true };

        var optionalOption = new Option<string[]?>(
            ["--optional"],
            "Optional attendees (email addresses). Space-, comma- or semicolon-separated. Resolved against the address book and listed in the body, never added as recipients")
        { AllowMultipleArgumentsPerToken = true };

        AddOption(subjectOption);
        AddOption(startOption);
        AddOption(endOption);
        AddOption(locationOption);
        AddOption(bodyOption);
        AddOption(allDayOption);
        AddOption(attendeesOption);
        AddOption(optionalOption);

        this.SetHandler(Execute, subjectOption, startOption, endOption, locationOption, bodyOption, allDayOption, attendeesOption, optionalOption);
    }

    /// <summary>
    /// Flattens "--attendees a@x.com,b@x.com c@x.com" into single addresses.
    /// </summary>
    public static string[] SplitAddresses(string[]? values) =>
        (values ?? [])
            .SelectMany(v => v.Split([',', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            .ToArray();

    private void Execute(string subject, DateTime start, DateTime end, string? location, string? body, bool allDay,
        string[]? attendees, string[]? optional)
    {
        var required = SplitAddresses(attendees);
        var optionalAttendees = SplitAddresses(optional);
        var isPlaceholder = required.Length + optionalAttendees.Length > 0;

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
            var entryId = service.CreateEvent(subject, start, end, location, body, allDay, required, optionalAttendees);

            var result = CommandResult<object>.Ok(
                "calendar create",
                new
                {
                    message = isPlaceholder
                        ? "Placeholder created. Nothing was sent: add the attendees listed in the body in Outlook, then press Send"
                        : "Event created successfully",
                    entryId,
                    subject,
                    start,
                    end,
                    location,
                    isAllDay = allDay,
                    isPlaceholder,
                    invitationsSent = false,
                    requiredAttendees = required,
                    optionalAttendees
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
