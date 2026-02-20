namespace OutlookCLI.Models;

public record FolderInfo(
    string Name,
    string FullPath,
    int ItemCount,
    int UnreadCount
);
