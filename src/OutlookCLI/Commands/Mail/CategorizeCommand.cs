using System.CommandLine;
using OutlookCLI.Configuration;
using OutlookCLI.Models;
using OutlookCLI.Output;
using OutlookCLI.Services;

namespace OutlookCLI.Commands.Mail;

public class CategorizeCommand : Command
{
    public CategorizeCommand() : base("categorize", "View or modify categories on an email. Use --list to view, --add/--remove/--set/--clear to modify.")
    {
        var entryIdArg = new Argument<string>(
            "entry-id",
            "The Entry ID of the email (from 'mail list' or 'mail search' output)");

        var addOption = new Option<string?>(
            "--add",
            "Add a category to existing categories");

        var removeOption = new Option<string?>(
            "--remove",
            "Remove a category from existing categories");

        var setOption = new Option<string[]?>(
            "--set",
            "Replace all categories with the specified list")
        { AllowMultipleArgumentsPerToken = true };

        var clearOption = new Option<bool>(
            "--clear",
            "Remove all categories");

        var listOption = new Option<bool>(
            "--list",
            "List current categories");

        AddArgument(entryIdArg);
        AddOption(addOption);
        AddOption(removeOption);
        AddOption(setOption);
        AddOption(clearOption);
        AddOption(listOption);

        this.SetHandler(Execute, entryIdArg, addOption, removeOption, setOption, clearOption, listOption);
    }

    private void Execute(string entryId, string? add, string? remove, string[]? set, bool clear, bool list)
    {
        var options = GlobalOptionsAccessor.Current;
        IOutputFormatter formatter = options.Human ? new HumanOutputFormatter() : new JsonOutputFormatter();

        // Validate mutually exclusive options
        int actionCount = 0;
        if (add != null) actionCount++;
        if (remove != null) actionCount++;
        if (set != null && set.Length > 0) actionCount++;
        if (clear) actionCount++;
        if (list) actionCount++;

        if (actionCount == 0)
        {
            // Default to --list
            list = true;
        }
        else if (actionCount > 1)
        {
            var errorResult = CommandResult<object>.Fail(
                "mail categorize",
                "INVALID_ARGS",
                "Specify exactly one of: --add, --remove, --set, --clear, --list"
            );
            Console.WriteLine(formatter.Format(errorResult));
            Environment.Exit(1);
            return;
        }

        using var service = new OutlookService();
        try
        {
            service.Initialize();

            if (list)
            {
                var categories = service.GetCategories(entryId);
                var result = CommandResult<object>.Ok(
                    "mail categorize",
                    new { entryId, categories },
                    new ResultMetadata { Count = categories.Count }
                );
                Console.WriteLine(formatter.Format(result));
                return;
            }

            var current = service.GetCategories(entryId);
            string newCategories;

            if (add != null)
            {
                if (!current.Contains(add, StringComparer.OrdinalIgnoreCase))
                    current.Add(add);
                newCategories = string.Join(", ", current);
            }
            else if (remove != null)
            {
                current.RemoveAll(c => c.Equals(remove, StringComparison.OrdinalIgnoreCase));
                newCategories = string.Join(", ", current);
            }
            else if (set != null && set.Length > 0)
            {
                newCategories = string.Join(", ", set);
                current = set.ToList();
            }
            else // clear
            {
                newCategories = "";
                current = new List<string>();
            }

            var success = service.SetCategories(entryId, newCategories);

            if (success)
            {
                var result = CommandResult<object>.Ok(
                    "mail categorize",
                    new { entryId, categories = current },
                    new ResultMetadata { Count = current.Count }
                );
                Console.WriteLine(formatter.Format(result));
            }
            else
            {
                var errorResult = CommandResult<object>.Fail(
                    "mail categorize",
                    "NOT_FOUND",
                    $"Email with Entry ID '{entryId}' not found"
                );
                Console.WriteLine(formatter.Format(errorResult));
                Environment.Exit(1);
            }
        }
        catch (Exception ex)
        {
            var result = CommandResult<object>.Fail(
                "mail categorize",
                "OUTLOOK_ERROR",
                ex.Message
            );
            Console.WriteLine(formatter.Format(result));
            Environment.Exit(1);
        }
    }
}
