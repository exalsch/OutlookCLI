namespace OutlookCLI.Models;

public record Attachment(
    string FileName,
    long Size,
    string ContentType
);
