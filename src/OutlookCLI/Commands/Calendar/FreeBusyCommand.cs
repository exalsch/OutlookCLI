using System.CommandLine;
using OutlookCLI.Configuration;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;

namespace OutlookCLI.Commands.Calendar;

public class FreeBusyCommand : Command
{
    public FreeBusyCommand() : base("free-busy", "Check free/busy availability for a person. Returns time slots with status (Free, Busy, Tentative, OOF).")
    {
        var emailOption = new Option<string>(
            ["--email", "-e"],
            "Email address to check availability for")
        { IsRequired = true };

        var startOption = new Option<DateTime?>(
            ["--start", "-s"],
            "Start date (format: yyyy-MM-dd. Default: today)");

        var endOption = new Option<DateTime?>(
            ["--end"],
            "End date (format: yyyy-MM-dd. Default: start + 1 day)");

        AddOption(emailOption);
        AddOption(startOption);
        AddOption(endOption);

        this.SetHandler(Execute, emailOption, startOption, endOption);
    }

    private void Execute(string email, DateTime? start, DateTime? end)
    {
        var options = GlobalOptionsAccessor.Current;
        IOutputFormatter formatter = options.Human ? new HumanOutputFormatter() : new JsonOutputFormatter();

        using var service = new OutlookService();
        try
        {
            service.Initialize();

            var effectiveStart = start ?? DateTime.Today;
            var effectiveEnd = end ?? effectiveStart.AddDays(1);

            var freeBusy = service.GetFreeBusy(email, effectiveStart, effectiveEnd);
            var result = CommandResult<FreeBusyResult>.Ok(
                "calendar free-busy",
                freeBusy,
                new ResultMetadata { Count = freeBusy.BusySlots.Count }
            );

            if (options.Human)
                Console.WriteLine(FormatFreeBusyHuman(result));
            else
                Console.WriteLine(formatter.Format(result));
        }
        catch (Exception ex)
        {
            var result = CommandResult<FreeBusyResult>.Fail(
                "calendar free-busy",
                "OUTLOOK_ERROR",
                ex.Message
            );
            Console.WriteLine(formatter.Format(result));
            Environment.Exit(1);
        }
    }

    private static string FormatFreeBusyHuman(CommandResult<FreeBusyResult> result)
    {
        if (!result.Success)
            return $"Error: [{result.Error?.Code}] {result.Error?.Message}";

        var fb = result.Data!;
        var sb = new System.Text.StringBuilder();
        sb.AppendLine($"Free/Busy for {fb.Email}");
        sb.AppendLine($"Range: {fb.RangeStart:MMM dd, yyyy} - {fb.RangeEnd:MMM dd, yyyy}");
        sb.AppendLine(new string('─', 50));

        if (fb.BusySlots.Count == 0)
        {
            sb.AppendLine("  No busy slots found (entirely free)");
        }
        else
        {
            foreach (var slot in fb.BusySlots)
            {
                var indicator = slot.Status switch
                {
                    "Busy" => "[BUSY]     ",
                    "Tentative" => "[TENTATIVE]",
                    "OOF" => "[OOF]      ",
                    _ => "[FREE]     "
                };
                sb.AppendLine($"  {indicator} {slot.Start:MMM dd h:mm tt} - {slot.End:h:mm tt}");
            }
        }

        return sb.ToString().TrimEnd();
    }
}
