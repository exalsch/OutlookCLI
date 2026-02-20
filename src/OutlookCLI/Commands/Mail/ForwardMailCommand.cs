using System.CommandLine;
using OutlookCLI.Configuration;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;

namespace OutlookCLI.Commands.Mail;

public class ForwardMailCommand : Command
{
    public ForwardMailCommand() : base("forward", "Forward an email to new recipients. Sends immediately with original message and attachments.")
    {
        var entryIdArg = new Argument<string>(
            "entry-id",
            "The Entry ID of the email to forward (from 'mail list' or 'mail search' output)");

        var toOption = new Option<string[]>(
            ["--to", "-t"],
            "Recipient email addresses. Multiple: --to a@x.com b@x.com (space-separated)")
        { IsRequired = true, AllowMultipleArgumentsPerToken = true };

        var bodyOption = new Option<string?>(
            ["--body", "-b"],
            "Additional message to prepend above the forwarded content");

        AddArgument(entryIdArg);
        AddOption(toOption);
        AddOption(bodyOption);

        this.SetHandler(Execute, entryIdArg, toOption, bodyOption);
    }

    private void Execute(string entryId, string[] to, string? body)
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
                    "mail forward",
                    "NOT_FOUND",
                    $"Email with Entry ID '{entryId}' not found"
                );
                Console.WriteLine(formatter.Format(errorResult));
                Environment.Exit(1);
                return;
            }

            service.ForwardMail(entryId, to, body);

            var result = CommandResult<object>.Ok(
                "mail forward",
                new
                {
                    message = "Email forwarded successfully",
                    originalSubject = originalMail.Subject,
                    forwardedTo = to
                },
                new ResultMetadata()
            );

            Console.WriteLine(formatter.Format(result));
        }
        catch (Exception ex)
        {
            var result = CommandResult<object>.Fail(
                "mail forward",
                "OUTLOOK_ERROR",
                ex.Message
            );
            Console.WriteLine(formatter.Format(result));
            Environment.Exit(1);
        }
    }
}
