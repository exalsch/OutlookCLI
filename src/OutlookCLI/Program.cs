using System.CommandLine;
using System.CommandLine.Builder;
using System.CommandLine.Parsing;
using OutlookCLI.Commands.Calendar;
using OutlookCLI.Commands.Mail;
using OutlookCLI.Configuration;

namespace OutlookCLI;

class Program
{
    static async Task<int> Main(string[] args)
    {
        var rootCommand = new RootCommand(
            "OutlookCLI - AI-agent-friendly CLI for managing Outlook emails and calendar via COM Interop.\n\n" +
            "Output: JSON by default (use --human for formatted text). All commands return {success, command, data, error, metadata}.\n" +
            "Entry IDs: Most commands require an entry-id obtained from 'mail list', 'mail search', or 'calendar list' output.\n" +
            "Requires: Outlook desktop installed and configured on Windows.");

        // Global options
        var humanOption = new Option<bool>(
            ["--human", "-H"],
            "Output in human-readable format instead of JSON");

        var noConfirmOption = new Option<bool>(
            ["--no-confirm"],
            "Skip confirmation dialogs for destructive actions");

        var fullOption = new Option<bool>(
            ["--full"],
            "Include full details (e.g., mail body) in list operations");

        rootCommand.AddGlobalOption(humanOption);
        rootCommand.AddGlobalOption(noConfirmOption);
        rootCommand.AddGlobalOption(fullOption);

        // Mail command group
        var mailCommand = new Command("mail", "Email operations (list, read, search, send, reply, forward, draft, delete, move, mark-read/unread, attachments, signatures)");
        mailCommand.AddCommand(new FoldersCommand());
        mailCommand.AddCommand(new ListMailCommand());
        mailCommand.AddCommand(new ReadMailCommand());
        mailCommand.AddCommand(new SearchMailCommand());
        mailCommand.AddCommand(new SendMailCommand());
        mailCommand.AddCommand(new DraftCommand());
        mailCommand.AddCommand(new ReplyMailCommand());
        mailCommand.AddCommand(new ForwardMailCommand());
        mailCommand.AddCommand(new DeleteMailCommand());
        mailCommand.AddCommand(new MoveMailCommand());
        mailCommand.AddCommand(new MarkReadCommand());
        mailCommand.AddCommand(new MarkUnreadCommand());
        mailCommand.AddCommand(new SaveAttachmentsCommand());
        mailCommand.AddCommand(new ExtractSignatureCommand());
        mailCommand.AddCommand(new CategorizeCommand());
        mailCommand.AddCommand(new ConversationCommand());
        rootCommand.AddCommand(mailCommand);

        // Calendar command group
        var calendarCommand = new Command("calendar", "Calendar operations (list, get, create, update, delete, respond to meetings)");
        calendarCommand.AddCommand(new ListEventsCommand());
        calendarCommand.AddCommand(new GetEventCommand());
        calendarCommand.AddCommand(new CreateEventCommand());
        calendarCommand.AddCommand(new UpdateEventCommand());
        calendarCommand.AddCommand(new DeleteEventCommand());
        calendarCommand.AddCommand(new RespondCommand());
        rootCommand.AddCommand(calendarCommand);

        // Build parser with middleware to capture global options
        var parser = new CommandLineBuilder(rootCommand)
            .UseDefaults()
            .AddMiddleware(async (context, next) =>
            {
                var parseResult = context.ParseResult;

                GlobalOptionsAccessor.Current = new GlobalOptions
                {
                    Human = parseResult.GetValueForOption(humanOption),
                    NoConfirm = parseResult.GetValueForOption(noConfirmOption),
                    Full = parseResult.GetValueForOption(fullOption)
                };

                await next(context);
            })
            .Build();

        return await parser.InvokeAsync(args);
    }
}
