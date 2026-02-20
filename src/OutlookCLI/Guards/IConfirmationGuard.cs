namespace OutlookCLI.Guards;

public interface IConfirmationGuard
{
    bool Confirm(string message, string title);
}
