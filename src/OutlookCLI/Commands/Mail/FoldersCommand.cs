using System.CommandLine;
using OutlookCLI.Configuration;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;

namespace OutlookCLI.Commands.Mail;

public class FoldersCommand : Command
{
    public FoldersCommand() : base("folders", "List available mail folders. Returns folder names and unread counts. Use these names with --folder in list/search/move.")
    {
        this.SetHandler(Execute);
    }

    private void Execute()
    {
        var options = GlobalOptionsAccessor.Current;
        IOutputFormatter formatter = options.Human ? new HumanOutputFormatter() : new JsonOutputFormatter();

        using var service = new OutlookService();
        try
        {
            service.Initialize();
            var folders = service.GetMailFolders().ToList();

            var result = CommandResult<IEnumerable<FolderInfo>>.Ok(
                "mail folders",
                folders,
                new ResultMetadata { Count = folders.Count }
            );

            Console.WriteLine(formatter.FormatFolders(result));
        }
        catch (Exception ex)
        {
            var result = CommandResult<IEnumerable<FolderInfo>>.Fail(
                "mail folders",
                "OUTLOOK_ERROR",
                ex.Message
            );
            Console.WriteLine(formatter.FormatFolders(result));
            Environment.Exit(1);
        }
    }
}
