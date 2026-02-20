using System.CommandLine;
using OutlookCLI.Configuration;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;

namespace OutlookCLI.Commands.Mail;

public class SaveAttachmentsCommand : Command
{
    public SaveAttachmentsCommand() : base("save-attachments", "Save email attachments to disk. Creates output directory if it doesn't exist. Returns list of saved file paths.")
    {
        var entryIdArg = new Argument<string>(
            "entry-id",
            "The Entry ID of the email with attachments (from 'mail list' or 'mail search' output)");

        var outputOption = new Option<DirectoryInfo>(
            ["--output", "-o"],
            () => new DirectoryInfo(Environment.CurrentDirectory),
            "Output directory for attachments (default: current directory, created if missing)");

        AddArgument(entryIdArg);
        AddOption(outputOption);

        this.SetHandler(Execute, entryIdArg, outputOption);
    }

    private void Execute(string entryId, DirectoryInfo output)
    {
        var options = GlobalOptionsAccessor.Current;
        IOutputFormatter formatter = options.Human ? new HumanOutputFormatter() : new JsonOutputFormatter();

        using var service = new OutlookService();
        try
        {
            service.Initialize();

            if (!output.Exists)
            {
                output.Create();
            }

            var savedFiles = service.SaveAttachments(entryId, output.FullName);

            if (savedFiles == null)
            {
                var errorResult = CommandResult<object>.Fail(
                    "mail save-attachments",
                    "NOT_FOUND",
                    $"Email with Entry ID '{entryId}' not found"
                );
                Console.WriteLine(formatter.Format(errorResult));
                Environment.Exit(1);
                return;
            }

            if (savedFiles.Count == 0)
            {
                var errorResult = CommandResult<object>.Fail(
                    "mail save-attachments",
                    "NO_ATTACHMENTS",
                    "Email has no attachments"
                );
                Console.WriteLine(formatter.Format(errorResult));
                return;
            }

            var result = CommandResult<object>.Ok(
                "mail save-attachments",
                new
                {
                    message = $"Saved {savedFiles.Count} attachment(s)",
                    outputDirectory = output.FullName,
                    files = savedFiles
                },
                new ResultMetadata { Count = savedFiles.Count }
            );
            Console.WriteLine(formatter.Format(result));
        }
        catch (Exception ex)
        {
            var result = CommandResult<object>.Fail(
                "mail save-attachments",
                "OUTLOOK_ERROR",
                ex.Message
            );
            Console.WriteLine(formatter.Format(result));
            Environment.Exit(1);
        }
    }
}
