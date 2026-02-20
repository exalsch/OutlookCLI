using System.CommandLine;
using OutlookCLI.Configuration;
using OutlookCLI.Guards;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;

namespace OutlookCLI.Commands.Calendar;

public class DeleteEventCommand : Command
{
    public DeleteEventCommand() : base("delete", "Delete a calendar event. Shows confirmation dialog unless --no-confirm is set.")
    {
        var entryIdArg = new Argument<string>(
            "entry-id",
            "The Entry ID of the event to delete (from 'calendar list' output)");

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

            // First get the event to show in confirmation
            var evt = service.GetEvent(entryId);
            if (evt == null)
            {
                var errorResult = CommandResult<object>.Fail(
                    "calendar delete",
                    "NOT_FOUND",
                    $"Event with Entry ID '{entryId}' not found"
                );
                Console.WriteLine(formatter.Format(errorResult));
                Environment.Exit(1);
                return;
            }

            // Confirm deletion
            if (!guard.Confirm($"Delete event: '{evt.Subject}' ({evt.Start:g})?", "Confirm Delete"))
            {
                var cancelledResult = CommandResult<object>.Fail(
                    "calendar delete",
                    "CANCELLED",
                    "Delete operation cancelled by user"
                );
                Console.WriteLine(formatter.Format(cancelledResult));
                return;
            }

            var success = service.DeleteEvent(entryId);
            if (success)
            {
                var result = CommandResult<object>.Ok(
                    "calendar delete",
                    new
                    {
                        message = "Event deleted successfully",
                        deletedSubject = evt.Subject,
                        deletedStart = evt.Start
                    },
                    new ResultMetadata()
                );
                Console.WriteLine(formatter.Format(result));
            }
            else
            {
                var errorResult = CommandResult<object>.Fail(
                    "calendar delete",
                    "DELETE_FAILED",
                    "Failed to delete the event"
                );
                Console.WriteLine(formatter.Format(errorResult));
                Environment.Exit(1);
            }
        }
        catch (Exception ex)
        {
            var result = CommandResult<object>.Fail(
                "calendar delete",
                "OUTLOOK_ERROR",
                ex.Message
            );
            Console.WriteLine(formatter.Format(result));
            Environment.Exit(1);
        }
    }
}
