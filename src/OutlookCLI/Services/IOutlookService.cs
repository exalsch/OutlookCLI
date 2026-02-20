using OutlookCLI.Models;

namespace OutlookCLI.Services;

public interface IOutlookService : IDisposable
{
    void Initialize();

    // Mail operations
    IEnumerable<FolderInfo> GetMailFolders();
    IEnumerable<MailMessageSummary> GetMailList(string? folderName, bool unreadOnly, int limit);
    IEnumerable<MailMessage> GetMailListFull(string? folderName, bool unreadOnly, int limit);
    MailMessage? GetMail(string entryId);
    IEnumerable<MailMessageSummary> SearchMail(string? query, string? from, DateTime? after, DateTime? before, string? folderName);
    IEnumerable<MailMessageSummary> GetConversation(string entryId, int limit);
    IEnumerable<MailMessage> GetConversationFull(string entryId, int limit);
    void SendMail(string[] to, string[]? cc, string subject, string body, bool isHtml = false, string[]? attachments = null);
    string CreateDraft(string[] to, string[]? cc, string subject, string body, bool isHtml = false, string[]? attachments = null);
    void ReplyToMail(string entryId, string body, bool replyAll);
    void ForwardMail(string entryId, string[] to, string? body);
    bool DeleteMail(string entryId);
    bool IsInDeletedItems(string entryId);
    void MoveMail(string entryId, string targetFolderName);
    bool MarkAsRead(string entryId, bool read);
    List<string> GetCategories(string entryId);
    bool SetCategories(string entryId, string categories);
    List<string>? SaveAttachments(string entryId, string outputDirectory);
    string? ExtractSignatureWithImages(string entryId);

    // Calendar operations
    IEnumerable<CalendarEventSummary> GetEventList(DateTime? start, DateTime? end, int limit);
    IEnumerable<CalendarEvent> GetEventListFull(DateTime? start, DateTime? end, int limit);
    CalendarEvent? GetEvent(string entryId);
    string CreateEvent(string subject, DateTime start, DateTime end, string? location, string? body, bool isAllDay);
    void UpdateEvent(string entryId, string? subject, DateTime? start, DateTime? end, string? location, string? body);
    bool DeleteEvent(string entryId);
    bool RespondToMeeting(string entryId, string responseType, string? message);
}
