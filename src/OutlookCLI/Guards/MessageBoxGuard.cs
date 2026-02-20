using System.Windows.Forms;

namespace OutlookCLI.Guards;

public class MessageBoxGuard : IConfirmationGuard
{
    public bool Confirm(string message, string title)
    {
        var result = MessageBox.Show(
            message,
            title,
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question,
            MessageBoxDefaultButton.Button2);

        return result == DialogResult.Yes;
    }
}
