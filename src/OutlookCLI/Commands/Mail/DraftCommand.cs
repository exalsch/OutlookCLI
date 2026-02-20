using System.CommandLine;
using OutlookCLI.Configuration;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;

namespace OutlookCLI.Commands.Mail;

public class DraftCommand : Command
{
    public DraftCommand() : base("draft", "Create a draft email in Outlook Drafts folder (saved but not sent). Returns the entryId of the new draft.")
    {
        var toOption = new Option<string[]>(
            ["--to", "-t"],
            "Recipient email addresses. Multiple: --to a@x.com b@x.com (space-separated)")
        { IsRequired = true, AllowMultipleArgumentsPerToken = true };

        var ccOption = new Option<string[]?>(
            ["--cc"],
            "CC recipients. Multiple: --cc a@x.com b@x.com (space-separated)")
        { AllowMultipleArgumentsPerToken = true };

        var subjectOption = new Option<string>(
            ["--subject", "-s"],
            "Email subject line")
        { IsRequired = true };

        var bodyOption = new Option<string?>(
            ["--body", "-b"],
            "Email body text (plain text, or HTML if --html is set). Mutually exclusive with --body-file");

        var bodyFileOption = new Option<FileInfo?>(
            ["--body-file"],
            "Read email body from a file path. Mutually exclusive with --body");

        var htmlOption = new Option<bool>(
            ["--html"],
            "Treat body as HTML content (auto-enabled when --signature-file is used)");

        var signatureFileOption = new Option<FileInfo?>(
            ["--signature-file"],
            "Append HTML signature from file. Extract one first with 'mail extract-signature'");

        var attachmentOption = new Option<FileInfo[]?>(
            ["--attachment", "-a"],
            "File(s) to attach. Multiple: --attachment file1.pdf file2.docx (space-separated)")
        { AllowMultipleArgumentsPerToken = true };

        AddOption(toOption);
        AddOption(ccOption);
        AddOption(subjectOption);
        AddOption(bodyOption);
        AddOption(bodyFileOption);
        AddOption(htmlOption);
        AddOption(signatureFileOption);
        AddOption(attachmentOption);

        this.SetHandler(Execute, toOption, ccOption, subjectOption, bodyOption, bodyFileOption, htmlOption, signatureFileOption, attachmentOption);
    }

    private void Execute(string[] to, string[]? cc, string subject, string? body, FileInfo? bodyFile, bool isHtml, FileInfo? signatureFile, FileInfo[]? attachments)
    {
        var options = GlobalOptionsAccessor.Current;
        IOutputFormatter formatter = options.Human ? new HumanOutputFormatter() : new JsonOutputFormatter();

        string emailBody;
        if (bodyFile != null)
        {
            if (!bodyFile.Exists)
            {
                var errorResult = CommandResult<object>.Fail(
                    "mail draft",
                    "FILE_NOT_FOUND",
                    $"Body file not found: {bodyFile.FullName}"
                );
                Console.WriteLine(formatter.Format(errorResult));
                Environment.Exit(1);
                return;
            }
            emailBody = File.ReadAllText(bodyFile.FullName);
        }
        else
        {
            emailBody = body ?? "";
        }

        // Append signature if provided
        if (signatureFile != null)
        {
            if (!signatureFile.Exists)
            {
                var errorResult = CommandResult<object>.Fail(
                    "mail draft",
                    "FILE_NOT_FOUND",
                    $"Signature file not found: {signatureFile.FullName}"
                );
                Console.WriteLine(formatter.Format(errorResult));
                Environment.Exit(1);
                return;
            }
            var signature = File.ReadAllText(signatureFile.FullName);
            emailBody = emailBody + "\n\n" + signature;
            isHtml = true; // Signatures are typically HTML
        }

        // Validate attachments exist
        var attachmentPaths = new List<string>();
        if (attachments != null)
        {
            foreach (var att in attachments)
            {
                if (!att.Exists)
                {
                    var errorResult = CommandResult<object>.Fail(
                        "mail draft",
                        "FILE_NOT_FOUND",
                        $"Attachment file not found: {att.FullName}"
                    );
                    Console.WriteLine(formatter.Format(errorResult));
                    Environment.Exit(1);
                    return;
                }
                attachmentPaths.Add(att.FullName);
            }
        }

        using var service = new OutlookService();
        try
        {
            service.Initialize();
            var entryId = service.CreateDraft(to, cc, subject, emailBody, isHtml, attachmentPaths.ToArray());

            var result = CommandResult<object>.Ok(
                "mail draft",
                new
                {
                    message = "Draft created successfully",
                    entryId,
                    to,
                    subject,
                    attachmentCount = attachmentPaths.Count
                },
                new ResultMetadata()
            );

            Console.WriteLine(formatter.Format(result));
        }
        catch (Exception ex)
        {
            var result = CommandResult<object>.Fail(
                "mail draft",
                "OUTLOOK_ERROR",
                ex.Message
            );
            Console.WriteLine(formatter.Format(result));
            Environment.Exit(1);
        }
    }
}
