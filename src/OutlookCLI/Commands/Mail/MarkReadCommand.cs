using System.CommandLine;
using OutlookCLI.Configuration;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;

namespace OutlookCLI.Commands.Mail;

public class MarkReadCommand : Command
{
    public MarkReadCommand() : base("mark-read", "Mark an email as read. Useful after processing an email in an automated workflow.")
    {
        var entryIdArg = new Argument<string>(
            "entry-id",
            "The Entry ID of the email to mark as read (from 'mail list' or 'mail search' output)");

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
            var success = service.MarkAsRead(entryId, true);

            if (success)
            {
                var result = CommandResult<object>.Ok(
                    "mail mark-read",
                    new { message = "Email marked as read", entryId },
                    new ResultMetadata()
                );
                Console.WriteLine(formatter.Format(result));
            }
            else
            {
                var errorResult = CommandResult<object>.Fail(
                    "mail mark-read",
                    "NOT_FOUND",
                    $"Email with Entry ID '{entryId}' not found"
                );
                Console.WriteLine(formatter.Format(errorResult));
                Environment.Exit(1);
            }
        }
        catch (Exception ex)
        {
            var result = CommandResult<object>.Fail(
                "mail mark-read",
                "OUTLOOK_ERROR",
                ex.Message
            );
            Console.WriteLine(formatter.Format(result));
            Environment.Exit(1);
        }
    }
}
