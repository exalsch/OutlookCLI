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
    List<string> Categories,
    string? ConversationTopic = null,
    bool IsFlagged = false,
    string FlagStatus = "none",
    string? FlagRequest = null,
    DateTime? FlagDueDate = null,
    DateTime? FlagStartDate = null,
    DateTime? FlagCompletedDate = null
);
