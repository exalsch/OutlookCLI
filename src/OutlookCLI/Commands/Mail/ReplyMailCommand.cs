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
            "Reply body text (HTML by default)")
        { IsRequired = true };

        var replyAllOption = new Option<bool>(
            ["--reply-all", "-a"],
            "Reply to all recipients (To + CC) instead of just the sender");

        var draftOption = new Option<bool>(
            ["--draft", "-d"],
            "Save as draft instead of sending immediately");

        var htmlOption = new Option<bool>(
            ["--html"],
            () => true,
            "Treat body as HTML (default: true). Use --html false for plain text");

        var signatureFileOption = new Option<FileInfo?>(
            ["--signature-file"],
            "Append HTML signature from file. Extract one first with 'mail extract-signature'");

        AddArgument(entryIdArg);
        AddOption(bodyOption);
        AddOption(replyAllOption);
        AddOption(draftOption);
        AddOption(htmlOption);
        AddOption(signatureFileOption);

        this.SetHandler(Execute, entryIdArg, bodyOption, replyAllOption, draftOption, htmlOption, signatureFileOption);
    }

    private void Execute(string entryId, string body, bool replyAll, bool draft, bool isHtml, FileInfo? signatureFile)
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

            // Append signature if provided
            if (signatureFile != null)
            {
                if (!signatureFile.Exists)
                {
                    var errorResult = CommandResult<object>.Fail(
                        "mail reply",
                        "FILE_NOT_FOUND",
                        $"Signature file not found: {signatureFile.FullName}"
                    );
                    Console.WriteLine(formatter.Format(errorResult));
                    Environment.Exit(1);
                    return;
                }
                var signature = File.ReadAllText(signatureFile.FullName);
                body = body + signature;
                isHtml = true;
            }

            var draftEntryId = service.ReplyToMail(entryId, body, replyAll, draft, isHtml);

            var result = CommandResult<object>.Ok(
                "mail reply",
                new
                {
                    message = draft
                        ? "Reply saved as draft"
                        : replyAll ? "Reply to all sent successfully" : "Reply sent successfully",
                    originalSubject = originalMail.Subject,
                    replyAll,
                    draft,
                    entryId = draft ? draftEntryId : (string?)null
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
