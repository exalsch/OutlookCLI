# OutlookCLI

A command-line tool for managing Outlook emails and calendar events via local COM Interop. Designed to be AI-agent friendly with JSON output by default.

## Why OutlookCLI?

- **No Azure registration** - Uses your existing Outlook session, no app registration or admin consent
- **AI-agent ready** - Consistent JSON output parseable by any automation tool
- **Full offline access** - Works with cached mail, no internet required
- **Simple CLI** - Intuitive commands for mail and calendar operations

## Requirements

- Windows 10/11
- .NET 8.0+ Runtime
- Microsoft Outlook (desktop) installed and configured

## Quick Start

```bash
# Build
dotnet build

# Run
dotnet run --project src/OutlookCLI -- mail list --limit 5

# Or publish a standalone executable
dotnet publish -c Release -r win-x64 --self-contained
```

## Usage

```
OutlookCLI [command] [options]

Global Options:
  -H, --human     Human-readable output instead of JSON
  --no-confirm    Skip confirmation dialogs (for automation)
  --full          Include full details (e.g., mail body) in list operations
```

### Mail Commands

```bash
# Browse
outlook mail folders                              # List folder names and unread counts
outlook mail list [--folder Inbox] [--unread] [--limit 20]
outlook mail read <entry-id>                      # Full email with body
outlook mail search --query "text" [--from user@example.com] [--after 2024-01-01]

# Compose
outlook mail send --to a@x.com --subject "Hi" --body "Hello" [--html] [--signature-file sig.html]
outlook mail draft --to a@x.com --subject "Review" --body "Draft" [--attachment file.pdf]
outlook mail reply <entry-id> --body "Thanks" [--reply-all]
outlook mail forward <entry-id> --to b@x.com [--body "FYI"]

# Organize
outlook mail move <entry-id> --to-folder Archive
outlook mail delete <entry-id> [--no-confirm]
outlook mail mark-read <entry-id>
outlook mail mark-unread <entry-id>

# Utilities
outlook mail save-attachments <entry-id> --output ./downloads
outlook mail extract-signature <entry-id> --output signature.html
```

### Calendar Commands

```bash
outlook calendar list [--start 2024-01-01] [--end 2024-02-01] [--limit 50]
outlook calendar get <entry-id>
outlook calendar create --subject "Meeting" --start "2024-01-15 09:00" --end "2024-01-15 10:00" [--location "Room A"]
outlook calendar update <entry-id> --subject "Updated Meeting"
outlook calendar delete <entry-id> [--no-confirm]
outlook calendar respond <entry-id> --accept [--message "See you there!"]
```

## JSON Output Format

All commands return a consistent JSON structure:

```json
{
  "success": true,
  "command": "mail list",
  "data": [ ... ],
  "error": null,
  "metadata": {
    "count": 10,
    "timestamp": "2024-01-15T10:30:00Z"
  }
}
```

On error:

```json
{
  "success": false,
  "command": "mail delete",
  "data": null,
  "error": "DELETED_ITEMS_PROTECTED",
  "metadata": {
    "message": "Cannot delete items from Deleted Items folder."
  }
}
```

## AI Agent Workflow Examples

### Process unread emails

```bash
# List unread, process each, mark as read
outlook mail list --unread --limit 10
outlook mail read <entry-id>
outlook mail mark-read <entry-id>
```

### Send email with corporate signature

```bash
# One-time: extract signature from a sent email
outlook mail list --folder "Sent Items" --limit 1
outlook mail extract-signature <entry-id> --output signature.html

# Reuse signature on every send
outlook mail send --to user@example.com --subject "Hello" \
  --body "<p>Message body</p>" --html --signature-file signature.html
```

### Create draft for human review

```bash
outlook mail draft --to user@example.com --subject "Proposal" \
  --body "Please review" --attachment report.pdf
```

### Accept a meeting

```bash
outlook calendar list --start 2024-01-15 --end 2024-01-16
outlook calendar respond <entry-id> --accept --message "Looking forward to it!"
```

## Entry IDs

Most commands require an `entry-id` argument. These are opaque Outlook identifiers returned by `mail list`, `mail search`, and `calendar list` in the `entryId` field of each item. Entry IDs may change when items are moved between folders.

## Safety Features

- **Confirmation guard** - Delete operations show a confirmation dialog (bypass with `--no-confirm`)
- **Deleted Items protection** - Cannot delete items already in Deleted Items (prevents accidental permanent deletion)
- **Non-destructive defaults** - JSON output, no side effects on read operations

## Building from Source

```bash
dotnet build                                        # Debug build
dotnet test                                         # Run tests
dotnet publish -c Release -r win-x64 --self-contained  # Standalone exe
```

## License

[MIT](LICENSE)
