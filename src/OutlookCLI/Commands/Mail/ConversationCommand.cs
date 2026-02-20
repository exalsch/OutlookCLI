using System.CommandLine;
using OutlookCLI.Configuration;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;

namespace OutlookCLI.Commands.Mail;

public class ConversationCommand : Command
{
    public ConversationCommand() : base("conversation", "Get all emails in the same conversation thread as the given email. Searches Inbox, Sent Mail, and Drafts.")
    {
        var entryIdArg = new Argument<string>(
            "entry-id",
            "The Entry ID of any email in the conversation (from 'mail list' or 'mail search' output)");

        var limitOption = new Option<int>(
            "--limit",
            () => 50,
            "Maximum number of messages to return");

        AddArgument(entryIdArg);
        AddOption(limitOption);

        this.SetHandler(Execute, entryIdArg, limitOption);
    }

    private void Execute(string entryId, int limit)
    {
        var options = GlobalOptionsAccessor.Current;
        IOutputFormatter formatter = options.Human ? new HumanOutputFormatter() : new JsonOutputFormatter();

        using var service = new OutlookService();
        try
        {
            service.Initialize();

            if (options.Full)
            {
                var messages = service.GetConversationFull(entryId, limit).ToList();
                var result = CommandResult<IEnumerable<MailMessage>>.Ok(
                    "mail conversation", messages, new ResultMetadata { Count = messages.Count });
                Console.WriteLine(formatter.FormatMailListFull(result));
            }
            else
            {
                var messages = service.GetConversation(entryId, limit).ToList();
                var result = CommandResult<IEnumerable<MailMessageSummary>>.Ok(
                    "mail conversation", messages, new ResultMetadata { Count = messages.Count });
                Console.WriteLine(formatter.FormatMailList(result));
            }
        }
        catch (Exception ex)
        {
            var result = CommandResult<IEnumerable<MailMessageSummary>>.Fail(
                "mail conversation",
                "OUTLOOK_ERROR",
                ex.Message
            );
            Console.WriteLine(formatter.FormatMailList(result));
            Environment.Exit(1);
        }
    }
}
