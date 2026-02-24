using System.CommandLine;
using OutlookCLI.Configuration;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;

namespace OutlookCLI.Commands.Calendar;

public class OpenEventCommand : Command
{
    public OpenEventCommand() : base("open", "Open a calendar event in the Outlook desktop window.")
    {
        var entryIdArg = new Argument<string>(
            "entry-id",
            "The Entry ID of the event to open (from 'calendar list' output)");

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
            service.OpenEvent(entryId);

            var result = CommandResult<object>.Ok(
                "calendar open",
                new { message = "Event opened in Outlook" },
                new ResultMetadata()
            );
            Console.WriteLine(formatter.Format(result));
        }
        catch (Exception ex)
        {
            var result = CommandResult<object>.Fail(
                "calendar open",
                "OUTLOOK_ERROR",
                ex.Message
            );
            Console.WriteLine(formatter.Format(result));
            Environment.Exit(1);
        }
    }
}
