using System.CommandLine;
using OutlookCLI.Configuration;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;

namespace OutlookCLI.Commands.Calendar;

public class RespondCommand : Command
{
    public RespondCommand() : base("respond", "Respond to a meeting invitation. Exactly one of --accept, --decline, or --tentative must be specified.")
    {
        var entryIdArg = new Argument<string>(
            "entry-id",
            "The Entry ID of the meeting to respond to (from 'calendar list' output)");

        var acceptOption = new Option<bool>(
            ["--accept"],
            "Accept the meeting invitation (mutually exclusive with --decline/--tentative)");

        var declineOption = new Option<bool>(
            ["--decline"],
            "Decline the meeting invitation (mutually exclusive with --accept/--tentative)");

        var tentativeOption = new Option<bool>(
            ["--tentative"],
            "Tentatively accept the meeting (mutually exclusive with --accept/--decline)");

        var messageOption = new Option<string?>(
            ["--message", "-m"],
            "Optional message to include with the response (sent to organizer)");

        AddArgument(entryIdArg);
        AddOption(acceptOption);
        AddOption(declineOption);
        AddOption(tentativeOption);
        AddOption(messageOption);

        this.SetHandler(Execute, entryIdArg, acceptOption, declineOption, tentativeOption, messageOption);
    }

    private void Execute(string entryId, bool accept, bool decline, bool tentative, string? message)
    {
        var options = GlobalOptionsAccessor.Current;
        IOutputFormatter formatter = options.Human ? new HumanOutputFormatter() : new JsonOutputFormatter();

        // Validate exactly one response type is selected
        var responseCount = (accept ? 1 : 0) + (decline ? 1 : 0) + (tentative ? 1 : 0);
        if (responseCount != 1)
        {
            var errorResult = CommandResult<object>.Fail(
                "calendar respond",
                "INVALID_ARGS",
                "Exactly one response type must be specified: --accept, --decline, or --tentative"
            );
            Console.WriteLine(formatter.Format(errorResult));
            Environment.Exit(1);
            return;
        }

        var responseType = accept ? "accept" : (decline ? "decline" : "tentative");

        using var service = new OutlookService();
        try
        {
            service.Initialize();
            var success = service.RespondToMeeting(entryId, responseType, message);

            if (success)
            {
                var result = CommandResult<object>.Ok(
                    "calendar respond",
                    new
                    {
                        message = $"Meeting {responseType}ed successfully",
                        entryId,
                        responseType
                    },
                    new ResultMetadata()
                );
                Console.WriteLine(formatter.Format(result));
            }
            else
            {
                var errorResult = CommandResult<object>.Fail(
                    "calendar respond",
                    "NOT_FOUND",
                    $"Meeting with Entry ID '{entryId}' not found"
                );
                Console.WriteLine(formatter.Format(errorResult));
                Environment.Exit(1);
            }
        }
        catch (Exception ex)
        {
            var result = CommandResult<object>.Fail(
                "calendar respond",
                "OUTLOOK_ERROR",
                ex.Message
            );
            Console.WriteLine(formatter.Format(result));
            Environment.Exit(1);
        }
    }
}
