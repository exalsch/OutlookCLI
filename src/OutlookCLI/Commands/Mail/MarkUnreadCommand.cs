using System.CommandLine;
using OutlookCLI.Configuration;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;

namespace OutlookCLI.Commands.Mail;

public class MarkUnreadCommand : Command
{
    public MarkUnreadCommand() : base("mark-unread", "Mark an email as unread. Useful to flag an email for later processing.")
    {
        var entryIdArg = new Argument<string>(
            "entry-id",
            "The Entry ID of the email to mark as unread (from 'mail list' or 'mail search' output)");

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
            var success = service.MarkAsRead(entryId, false);

            if (success)
            {
                var result = CommandResult<object>.Ok(
                    "mail mark-unread",
                    new { message = "Email marked as unread", entryId },
                    new ResultMetadata()
                );
                Console.WriteLine(formatter.Format(result));
            }
            else
            {
                var errorResult = CommandResult<object>.Fail(
                    "mail mark-unread",
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
                "mail mark-unread",
                "OUTLOOK_ERROR",
                ex.Message
            );
            Console.WriteLine(formatter.Format(result));
            Environment.Exit(1);
        }
    }
}
