using System.CommandLine;
using OutlookCLI.Configuration;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;

namespace OutlookCLI.Commands.Mail;

public class ReplyMailCommand : Command
{
    public ReplyMailCommand() : base("reply", "Reply to an email. Sends immediately. The original message is included in the reply thread.")
    {
        var entryIdArg = new Argument<string>(
            "entry-id",
            "The Entry ID of the email to reply to (from 'mail list' or 'mail search' output)");

        var bodyOption = new Option<string>(
            ["--body", "-b"],
            "Reply body text (plain text)")
        { IsRequired = true };

        var replyAllOption = new Option<bool>(
            ["--reply-all", "-a"],
            "Reply to all recipients (To + CC) instead of just the sender");

        AddArgument(entryIdArg);
        AddOption(bodyOption);
        AddOption(replyAllOption);

        this.SetHandler(Execute, entryIdArg, bodyOption, replyAllOption);
    }

    private void Execute(string entryId, string body, bool replyAll)
    {
        var options = GlobalOptionsAccessor.Current;
        IOutputFormatter formatter = options.Human ? new HumanOutputFormatter() : new JsonOutputFormatter();

        using var service = new OutlookService();
        try
        {
            service.Initialize();

            // First verify the email exists
            var originalMail = service.GetMail(entryId);
            if (originalMail == null)
            {
                var errorResult = CommandResult<object>.Fail(
                    "mail reply",
                    "NOT_FOUND",
                    $"Email with Entry ID '{entryId}' not found"
                );
                Console.WriteLine(formatter.Format(errorResult));
                Environment.Exit(1);
                return;
            }

            service.ReplyToMail(entryId, body, replyAll);

            var result = CommandResult<object>.Ok(
                "mail reply",
                new
                {
                    message = replyAll ? "Reply to all sent successfully" : "Reply sent successfully",
                    originalSubject = originalMail.Subject,
                    replyAll
                },
                new ResultMetadata()
            );

            Console.WriteLine(formatter.Format(result));
        }
        catch (Exception ex)
        {
            var result = CommandResult<object>.Fail(
                "mail reply",
                "OUTLOOK_ERROR",
                ex.Message
            );
            Console.WriteLine(formatter.Format(result));
            Environment.Exit(1);
        }
    }
}
