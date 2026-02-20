using System.CommandLine;
using OutlookCLI.Configuration;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;

namespace OutlookCLI.Commands.Mail;

public class MoveMailCommand : Command
{
    public MoveMailCommand() : base("move", "Move an email to a different folder. Note: the entry-id changes after moving.")
    {
        var entryIdArg = new Argument<string>(
            "entry-id",
            "The Entry ID of the email to move (from 'mail list' or 'mail search' output)");

        var toFolderOption = new Option<string>(
            ["--to-folder", "-t"],
            "Target folder name (use 'mail folders' to discover available names)")
        { IsRequired = true };

        AddArgument(entryIdArg);
        AddOption(toFolderOption);

        this.SetHandler(Execute, entryIdArg, toFolderOption);
    }

    private void Execute(string entryId, string toFolder)
    {
        var options = GlobalOptionsAccessor.Current;
        IOutputFormatter formatter = options.Human ? new HumanOutputFormatter() : new JsonOutputFormatter();

        using var service = new OutlookService();
        try
        {
            service.Initialize();

            // First verify the email exists
            var mail = service.GetMail(entryId);
            if (mail == null)
            {
                var errorResult = CommandResult<object>.Fail(
                    "mail move",
                    "NOT_FOUND",
                    $"Email with Entry ID '{entryId}' not found"
                );
                Console.WriteLine(formatter.Format(errorResult));
                Environment.Exit(1);
                return;
            }

            service.MoveMail(entryId, toFolder);

            var result = CommandResult<object>.Ok(
                "mail move",
                new
                {
                    message = $"Email moved to '{toFolder}'",
                    movedSubject = mail.Subject,
                    fromFolder = mail.Folder,
                    toFolder
                },
                new ResultMetadata()
            );

            Console.WriteLine(formatter.Format(result));
        }
        catch (DirectoryNotFoundException)
        {
            var result = CommandResult<object>.Fail(
                "mail move",
                "FOLDER_NOT_FOUND",
                $"Target folder '{toFolder}' not found"
            );
            Console.WriteLine(formatter.Format(result));
            Environment.Exit(1);
        }
        catch (Exception ex)
        {
            var result = CommandResult<object>.Fail(
                "mail move",
                "OUTLOOK_ERROR",
                ex.Message
            );
            Console.WriteLine(formatter.Format(result));
            Environment.Exit(1);
        }
    }
}
