using System.CommandLine;
using OutlookCLI.Configuration;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;

namespace OutlookCLI.Commands.Mail;

public class ReadMailCommand : Command
{
    public ReadMailCommand() : base("read", "Read a specific email by Entry ID. Returns full details: subject, from, to, cc, date, body (HTML), attachments, isRead.")
    {
        var entryIdArg = new Argument<string>(
            "entry-id",
            "The Entry ID of the email to read (from 'mail list' or 'mail search' output)");

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
            var message = service.GetMail(entryId);

            if (message == null)
            {
                var errorResult = CommandResult<MailMessage>.Fail(
                    "mail read",
                    "NOT_FOUND",
                    $"Email with Entry ID '{entryId}' not found"
                );
                Console.WriteLine(formatter.FormatMail(errorResult));
                Environment.Exit(1);
                return;
            }

            var result = CommandResult<MailMessage>.Ok("mail read", message);
            Console.WriteLine(formatter.FormatMail(result));
        }
        catch (Exception ex)
        {
            var result = CommandResult<MailMessage>.Fail(
                "mail read",
                "OUTLOOK_ERROR",
                ex.Message
            );
            Console.WriteLine(formatter.FormatMail(result));
            Environment.Exit(1);
        }
    }
}
