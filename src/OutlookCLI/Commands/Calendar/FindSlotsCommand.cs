using System.CommandLine;
using OutlookCLI.Configuration;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;

namespace OutlookCLI.Commands.Calendar;

public class FindSlotsCommand : Command
{
    public FindSlotsCommand() : base("find-slots", "Find common free meeting slots across multiple people. Filters to business hours (09:00-17:00) on weekdays.")
    {
        var emailsOption = new Option<string>(
            ["--emails", "-e"],
            "Comma-separated email addresses to check")
        { IsRequired = true };

        var startOption = new Option<DateTime?>(
            ["--start", "-s"],
            "Start date (format: yyyy-MM-dd. Default: tomorrow)");

        var endOption = new Option<DateTime?>(
            ["--end"],
            "End date (format: yyyy-MM-dd. Default: start + 5 weekdays)");

        var durationOption = new Option<int>(
            ["--duration", "-d"],
            () => 60,
            "Required meeting duration in minutes");

        var includeSelfOption = new Option<bool>(
            ["--include-self"],
            () => true,
            "Include your own calendar in availability check");

        AddOption(emailsOption);
        AddOption(startOption);
        AddOption(endOption);
        AddOption(durationOption);
        AddOption(includeSelfOption);

        this.SetHandler(Execute, emailsOption, startOption, endOption, durationOption, includeSelfOption);
    }

    private void Execute(string emails, DateTime? start, DateTime? end, int duration, bool includeSelf)
    {
        var options = GlobalOptionsAccessor.Current;
        IOutputFormatter formatter = options.Human ? new HumanOutputFormatter() : new JsonOutputFormatter();

        using var service = new OutlookService();
        try
        {
            service.Initialize();

            var emailList = emails.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();

            var effectiveStart = start ?? DateTime.Today.AddDays(1);
            // Default end: start + 5 weekdays
            var effectiveEnd = end ?? AddWeekdays(effectiveStart, 5);

            var slots = service.FindAvailableSlots(emailList, effectiveStart, effectiveEnd, duration, includeSelf);
            var result = CommandResult<List<AvailableSlot>>.Ok(
                "calendar find-slots",
                slots,
                new ResultMetadata { Count = slots.Count }
            );

            if (options.Human)
                Console.WriteLine(FormatSlotsHuman(result));
            else
                Console.WriteLine(formatter.Format(result));
        }
        catch (Exception ex)
        {
            var result = CommandResult<List<AvailableSlot>>.Fail(
                "calendar find-slots",
                "OUTLOOK_ERROR",
                ex.Message
            );
            Console.WriteLine(formatter.Format(result));
            Environment.Exit(1);
        }
    }

    private static DateTime AddWeekdays(DateTime start, int days)
    {
        var date = start;
        int added = 0;
        while (added < days)
        {
            date = date.AddDays(1);
            if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday)
                added++;
        }
        return date;
    }

    private static string FormatSlotsHuman(CommandResult<List<AvailableSlot>> result)
    {
        if (!result.Success)
            return $"Error: [{result.Error?.Code}] {result.Error?.Message}";

        var slots = result.Data!;
        var sb = new System.Text.StringBuilder();
        sb.AppendLine($"Available Slots ({slots.Count} found)");
        sb.AppendLine($"Attendees: {string.Join(", ", slots.FirstOrDefault()?.Attendees ?? [])}");
        sb.AppendLine(new string('─', 50));

        if (slots.Count == 0)
        {
            sb.AppendLine("  No common free slots found in the given range.");
        }
        else
        {
            DateTime? lastDate = null;
            int index = 1;
            foreach (var slot in slots)
            {
                if (lastDate == null || slot.Start.Date != lastDate.Value.Date)
                {
                    if (lastDate != null) sb.AppendLine();
                    sb.AppendLine($"  {slot.Start:ddd, MMM dd yyyy}");
                    lastDate = slot.Start.Date;
                }
                sb.AppendLine($"    {index}. {slot.Start:h:mm tt} - {slot.End:h:mm tt} ({slot.DurationMinutes} min)");
                index++;
            }
        }

        return sb.ToString().TrimEnd();
    }
}
