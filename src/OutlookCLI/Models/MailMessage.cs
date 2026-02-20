namespace OutlookCLI.Models;

public record MailMessage(
    string EntryId,
    string Subject,
    string SenderName,
    string SenderEmail,
    DateTime ReceivedTime,
    bool IsUnread,
    string Folder,
    bool HasAttachments,
    int AttachmentCount,
    string Body,
    string BodyHtml,
    List<string> To,
    List<string> Cc,
    List<Attachment> Attachments,
    List<string> Categories,
    string? ConversationTopic = null
);
