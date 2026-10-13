namespace OutlookCLI.Services;

/// <summary>
/// Follow-up flag state of a mail item (OlFlagStatus enum values).
/// </summary>
public enum MailFlagStatus
{
    None = 0,      // olNoFlag
    Complete = 1,  // olFlagComplete
    Flagged = 2    // olFlagMarked
}

/// <summary>
/// Follow-up flag details read from a mail item. Dates are null when Outlook reports "none" (4501-01-01).
/// </summary>
public record MailFlagInfo(
    MailFlagStatus Status,
    string? Request,
    DateTime? DueDate,
    DateTime? StartDate,
    DateTime? CompletedDate)
{
    public static readonly MailFlagInfo Empty = new(MailFlagStatus.None, null, null, null, null);

    public bool IsFlagged => Status == MailFlagStatus.Flagged;

    public string StatusName => MailFlags.ToName(Status);
}

/// <summary>
/// Pure helpers for mapping Outlook follow-up flag values. No COM access, so these are unit-testable.
/// </summary>
public static class MailFlags
{
    /// <summary>Accepted values for the --flag-status option.</summary>
    public static readonly string[] StatusNames = ["none", "flagged", "complete"];

    /// <summary>Outlook uses 4501-01-01 as the "no date" sentinel for task/flag dates.</summary>
    private const int OutlookNoDateYear = 4501;

    public static string ToName(MailFlagStatus status) => status switch
    {
        MailFlagStatus.Flagged => "flagged",
        MailFlagStatus.Complete => "complete",
        _ => "none"
    };

    public static MailFlagStatus FromOutlook(int flagStatus) => flagStatus switch
    {
        2 => MailFlagStatus.Flagged,
        1 => MailFlagStatus.Complete,
        _ => MailFlagStatus.None
    };

    public static bool TryParse(string? name, out MailFlagStatus status)
    {
        switch (name?.Trim().ToLowerInvariant())
        {
            case "none":
                status = MailFlagStatus.None;
                return true;
            case "flagged":
                status = MailFlagStatus.Flagged;
                return true;
            case "complete":
                status = MailFlagStatus.Complete;
                return true;
            default:
                status = MailFlagStatus.None;
                return false;
        }
    }

    /// <summary>
    /// Combines the --flagged switch and the --flag-status option into a single filter.
    /// Returns false with an error message when the two conflict or the status is unknown.
    /// </summary>
    public static bool TryResolveFilter(bool flagged, string? flagStatus, out MailFlagStatus? filter, out string? error)
    {
        filter = null;
        error = null;

        if (!string.IsNullOrWhiteSpace(flagStatus))
        {
            if (!TryParse(flagStatus, out var parsed))
            {
                error = $"Invalid --flag-status '{flagStatus}'. Use one of: {string.Join(", ", StatusNames)}";
                return false;
            }
            if (flagged && parsed != MailFlagStatus.Flagged)
            {
                error = $"--flagged conflicts with --flag-status {ToName(parsed)}";
                return false;
            }
            filter = parsed;
            return true;
        }

        if (flagged)
            filter = MailFlagStatus.Flagged;

        return true;
    }

    /// <summary>
    /// Items.Restrict filter for a flag status.
    /// Note: on Exchange cached-mode stores "[FlagStatus] = 2" (and the equivalent DASL filter on
    /// PR_FLAG_STATUS 0x10900003) returns no items even when flagged items exist, while ">= 2" matches
    /// exactly the flagged items. PR_FLAG_STATUS only takes values 0-2, so ">= 2" is equivalent.
    /// "= 0" also matches items that lack the property.
    /// </summary>
    public static string GetRestrictFilter(MailFlagStatus status) => status switch
    {
        MailFlagStatus.Flagged => "[FlagStatus] >= 2",
        MailFlagStatus.Complete => "[FlagStatus] = 1",
        _ => "[FlagStatus] = 0"
    };

    public static DateTime? NormalizeDate(DateTime? value)
    {
        if (value == null) return null;
        var date = value.Value;
        if (date.Year >= OutlookNoDateYear || date == DateTime.MinValue) return null;
        return date;
    }

    public static string? NormalizeRequest(string? request) =>
        string.IsNullOrWhiteSpace(request) ? null : request.Trim();
}
