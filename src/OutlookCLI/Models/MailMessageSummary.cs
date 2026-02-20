namespace OutlookCLI.Models;

public record MailMessageSummary(
    string EntryId,
    string Subject,
    string SenderName,
    string SenderEmail,
    DateTime ReceivedTime,
    bool IsUnread,
    string Folder,
    bool HasAttachments,
    int AttachmentCount,
    List<string> Categories
);
