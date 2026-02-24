using System.CommandLine;
using OutlookCLI.Configuration;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;

namespace OutlookCLI.Commands.Mail;

public class OpenMailCommand : Command
{
    public OpenMailCommand() : base("open", "Open an email in the Outlook desktop window.")
    {
        var entryIdArg = new Argument<string>(
            "entry-id",
            "The Entry ID of the email to open (from 'mail list' or 'mail search' output)");

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
            service.OpenMail(entryId);

            var result = CommandResult<object>.Ok(
                "mail open",
                new { message = "Email opened in Outlook" },
                new ResultMetadata()
            );
            Console.WriteLine(formatter.Format(result));
        }
        catch (Exception ex)
        {
            var result = CommandResult<object>.Fail(
                "mail open",
                "OUTLOOK_ERROR",
                ex.Message
            );
            Console.WriteLine(formatter.Format(result));
            Environment.Exit(1);
        }
    }
}
