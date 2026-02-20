using System.CommandLine;
using OutlookCLI.Configuration;
using OutlookCLI.Guards;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;

namespace OutlookCLI.Commands.Mail;

public class DeleteMailCommand : Command
{
    public DeleteMailCommand() : base("delete", "Delete an email (moves to Deleted Items, not permanent). Shows confirmation dialog unless --no-confirm is set. Cannot delete from Deleted Items.")
    {
        var entryIdArg = new Argument<string>(
            "entry-id",
            "The Entry ID of the email to delete (from 'mail list' or 'mail search' output)");

        AddArgument(entryIdArg);

        this.SetHandler(Execute, entryIdArg);
    }

    private void Execute(string entryId)
    {
        var options = GlobalOptionsAccessor.Current;
        IOutputFormatter formatter = options.Human ? new HumanOutputFormatter() : new JsonOutputFormatter();
        IConfirmationGuard guard = options.NoConfirm ? new NoOpGuard() : new MessageBoxGuard();

        using var service = new OutlookService();
        try
        {
            service.Initialize();

            // First get the email to show in confirmation
            var mail = service.GetMail(entryId);
            if (mail == null)
            {
                var errorResult = CommandResult<object>.Fail(
                    "mail delete",
                    "NOT_FOUND",
                    $"Email with Entry ID '{entryId}' not found"
                );
                Console.WriteLine(formatter.Format(errorResult));
                Environment.Exit(1);
                return;
            }

            // Check if already in Deleted Items
            if (service.IsInDeletedItems(entryId))
            {
                var errorResult = CommandResult<object>.Fail(
                    "mail delete",
                    "DELETED_ITEMS_PROTECTED",
                    "Cannot delete items from Deleted Items folder. Use Outlook directly for permanent deletion."
                );
                Console.WriteLine(formatter.Format(errorResult));
                Environment.Exit(1);
                return;
            }

            // Confirm deletion
            if (!guard.Confirm($"Delete email: '{mail.Subject}'?", "Confirm Delete"))
            {
                var cancelledResult = CommandResult<object>.Fail(
                    "mail delete",
                    "CANCELLED",
                    "Delete operation cancelled by user"
                );
                Console.WriteLine(formatter.Format(cancelledResult));
                return;
            }

            var success = service.DeleteMail(entryId);
            if (success)
            {
                var result = CommandResult<object>.Ok(
                    "mail delete",
                    new
                    {
                        message = "Email moved to Deleted Items",
                        deletedSubject = mail.Subject
                    },
                    new ResultMetadata()
                );
                Console.WriteLine(formatter.Format(result));
            }
            else
            {
                var errorResult = CommandResult<object>.Fail(
                    "mail delete",
                    "DELETE_FAILED",
                    "Failed to delete the email"
                );
                Console.WriteLine(formatter.Format(errorResult));
                Environment.Exit(1);
            }
        }
        catch (Exception ex)
        {
            var result = CommandResult<object>.Fail(
                "mail delete",
                "OUTLOOK_ERROR",
                ex.Message
            );
            Console.WriteLine(formatter.Format(result));
            Environment.Exit(1);
        }
    }
}
