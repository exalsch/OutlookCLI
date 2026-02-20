namespace OutlookCLI.Guards;

public class NoOpGuard : IConfirmationGuard
{
    public bool Confirm(string message, string title)
    {
        return true;
    }
}
