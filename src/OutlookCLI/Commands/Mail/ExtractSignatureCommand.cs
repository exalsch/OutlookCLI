using System.CommandLine;
using System.Text.RegularExpressions;
using OutlookCLI.Configuration;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;

namespace OutlookCLI.Commands.Mail;

public class ExtractSignatureCommand : Command
{
    public ExtractSignatureCommand() : base("extract-signature", "Extract signature HTML from a sent email. Images are embedded as base64. Save to file and reuse with --signature-file in send/draft.")
    {
        var entryIdArg = new Argument<string>(
            "entry-id",
            "The Entry ID of a sent email containing a signature (use 'mail list --folder \"Sent Items\"' to find one)");

        var outputOption = new Option<FileInfo?>(
            ["--output", "-o"],
            "Save signature HTML to file (default: print to stdout). Reuse with 'mail send --signature-file <path>'");

        AddArgument(entryIdArg);
        AddOption(outputOption);

        this.SetHandler(Execute, entryIdArg, outputOption);
    }

    private void Execute(string entryId, FileInfo? output)
    {
        var options = GlobalOptionsAccessor.Current;
        IOutputFormatter formatter = options.Human ? new HumanOutputFormatter() : new JsonOutputFormatter();

        using var service = new OutlookService();
        try
        {
            service.Initialize();

            // Get signature HTML with embedded images converted to base64
            var signature = service.ExtractSignatureWithImages(entryId);

            if (signature == null)
            {
                var errorResult = CommandResult<object>.Fail(
                    "mail extract-signature",
                    "NOT_FOUND",
                    $"Email with Entry ID '{entryId}' not found"
                );
                Console.WriteLine(formatter.Format(errorResult));
                Environment.Exit(1);
                return;
            }

            if (string.IsNullOrEmpty(signature))
            {
                var errorResult = CommandResult<object>.Fail(
                    "mail extract-signature",
                    "NO_SIGNATURE",
                    "No signature found in the email. Look for emails with a signature block."
                );
                Console.WriteLine(formatter.Format(errorResult));
                Environment.Exit(1);
                return;
            }

            if (output != null)
            {
                File.WriteAllText(output.FullName, signature);
                var result = CommandResult<object>.Ok(
                    "mail extract-signature",
                    new
                    {
                        message = "Signature extracted and saved (images embedded as base64)",
                        outputFile = output.FullName,
                        signatureLength = signature.Length
                    },
                    new ResultMetadata()
                );
                Console.WriteLine(formatter.Format(result));
            }
            else
            {
                var result = CommandResult<object>.Ok(
                    "mail extract-signature",
                    new
                    {
                        signature,
                        signatureLength = signature.Length
                    },
                    new ResultMetadata()
                );
                Console.WriteLine(formatter.Format(result));
            }
        }
        catch (Exception ex)
        {
            var result = CommandResult<object>.Fail(
                "mail extract-signature",
                "OUTLOOK_ERROR",
                ex.Message
            );
            Console.WriteLine(formatter.Format(result));
            Environment.Exit(1);
        }
    }
}
