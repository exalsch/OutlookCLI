namespace OutlookCLI.Models;

public record CalendarEventSummary(
    string EntryId,
    string Subject,
    DateTime Start,
    DateTime End,
    string Location,
    bool IsAllDay,
    bool IsRecurring
);
