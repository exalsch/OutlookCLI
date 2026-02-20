namespace OutlookCLI.Guards;

public class MessageBoxGuard : IConfirmationGuard
{
    public bool Confirm(string message, string title)
    {
        Console.Error.Write($"{title}: {message} [y/N] ");
        var key = Console.ReadKey(intercept: false);
        Console.Error.WriteLine();
        return key.KeyChar is 'y' or 'Y';
    }
}
