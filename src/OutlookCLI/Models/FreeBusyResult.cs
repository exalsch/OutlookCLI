namespace OutlookCLI.Models;

public record FreeBusySlot(DateTime Start, DateTime End, string Status);

public record FreeBusyResult(string Email, DateTime RangeStart, DateTime RangeEnd, List<FreeBusySlot> BusySlots);

public record AvailableSlot(DateTime Start, DateTime End, int DurationMinutes, List<string> Attendees);
