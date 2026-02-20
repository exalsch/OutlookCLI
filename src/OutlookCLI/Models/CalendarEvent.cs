namespace OutlookCLI.Models;

public record CalendarEvent(
    string EntryId,
    string Subject,
    DateTime Start,
    DateTime End,
    string Location,
    bool IsAllDay,
    bool IsRecurring,
    string Body,
    List<string> Attendees,
    string Organizer,
    string? RecurrencePattern
);
