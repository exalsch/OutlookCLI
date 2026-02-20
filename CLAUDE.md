# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

OutlookCLI is a .NET command-line tool for managing Outlook emails and calendar events via local Outlook COM Interop. It is designed to be AI-agent friendly with JSON output by default.

## Technology Stack

- **.NET 8+** - Target framework (Windows only)
- **Late-binding COM Interop** - No PIAs required, uses `dynamic` for Outlook automation
- **System.CommandLine** - CLI argument parsing and command structure

## Architecture Decisions

### Why Outlook COM Interop (not Microsoft Graph)
- No Azure app registration required
- No admin consent needed
- Uses existing Outlook credentials/session
- Full offline access to cached mail
- Tradeoff: Windows-only, requires Outlook installed

### COM Interop Patterns

Uses late-binding COM interop (no PIAs required), which avoids assembly version conflicts:
```csharp
var outlookType = Type.GetTypeFromProgID("Outlook.Application");
dynamic app = Activator.CreateInstance(outlookType);
dynamic ns = app.GetNamespace("MAPI");
try
{
    // work with Outlook objects using dynamic
}
finally
{
    Marshal.ReleaseComObject(ns);
    Marshal.ReleaseComObject(app);
}
```

Key folder constants (OlDefaultFolders values):
- `olFolderInbox = 6`
- `olFolderCalendar = 9`
- `olFolderSentMail = 5`
- `olFolderDrafts = 16`
- `olFolderDeletedItems = 3`
- `olFolderContacts = 10`

Item class constants (for type checking):
- `olMail = 43`
- `olAppointment = 26`

## CLI Commands

### Global Options
| Option | Description |
|--------|-------------|
| `--human` / `-H` | Human-readable output instead of JSON |
| `--no-confirm` | Skip confirmation dialogs for destructive actions |
| `--full` | Include full details (e.g., mail body) in list operations |

### Mail Commands
```bash
# List and read
outlook mail folders                                    # List available mail folders
outlook mail list [--folder <name>] [--unread] [--limit <n>] [--full]
outlook mail read <entry-id>                            # Read specific email
outlook mail search --query <text> [--from <email>] [--after <date>] [--before <date>] [--folder <name>]

# Send and draft
outlook mail send --to <emails> [--cc <emails>] --subject <text> [--body <text>] [--body-file <path>] [--html] [--signature-file <path>] [--attachment <files>]
outlook mail draft --to <emails> --subject <text> [--body <text>] [--html] [--signature-file <path>] [--attachment <files>]

# Reply and forward
outlook mail reply <entry-id> --body <text> [--reply-all]
outlook mail forward <entry-id> --to <emails> [--body <text>]

# Organize
outlook mail delete <entry-id> [--no-confirm]           # Moves to Deleted Items
outlook mail move <entry-id> --to-folder <name>
outlook mail mark-read <entry-id>                       # Mark as read
outlook mail mark-unread <entry-id>                     # Mark as unread

# Attachments and signatures
outlook mail save-attachments <entry-id> [--output <dir>]
outlook mail extract-signature <entry-id> [--output <file>]
```

### Calendar Commands
```bash
outlook calendar list [--start <date>] [--end <date>] [--limit <n>] [--full]
outlook calendar get <entry-id>
outlook calendar create --subject <text> --start <datetime> --end <datetime> [--location <text>] [--body <text>] [--all-day]
outlook calendar update <entry-id> [--subject <text>] [--start <datetime>] [--end <datetime>] [--location <text>] [--body <text>]
outlook calendar delete <entry-id> [--no-confirm]
outlook calendar respond <entry-id> --accept|--decline|--tentative [--message <text>]
```

## AI Assistant Workflow Examples

### Process unread emails
```bash
# List unread emails
outlook mail list --unread --limit 10

# Mark as read after processing
outlook mail mark-read <entry-id>
```

### Send email with signature
```bash
# Extract signature once from a sent email
outlook mail extract-signature <sent-email-id> --output signature.html

# Send emails with signature
outlook mail send --to user@example.com --subject "Hello" --body "<p>Message</p>" --html --signature-file signature.html
```

### Create draft for human review
```bash
outlook mail draft --to user@example.com --subject "Review needed" --body "Draft content" --attachment report.pdf
```

### Accept meeting invitation
```bash
outlook calendar respond <meeting-id> --accept --message "Looking forward to it!"
```

### Save attachments for processing
```bash
outlook mail save-attachments <entry-id> --output ./downloads
```

## Build Commands

```bash
dotnet build
dotnet run --project src/OutlookCLI -- <command> [options]
dotnet test
dotnet publish -c Release -r win-x64 --self-contained
```

## Output Format

### JSON Output (Default)
All commands return consistent JSON structure:
```json
{
  "success": true,
  "command": "mail list",
  "data": [...],
  "error": null,
  "metadata": {
    "count": 10,
    "timestamp": "2024-01-15T10:30:00Z"
  }
}
```

### Human Output (--human flag)
Human-readable formatted output for terminal use.

## Safety Features

1. **Confirmation Guard**: Destructive actions (delete) show a MessageBox confirmation unless `--no-confirm` is set
2. **Deleted Items Protection**: Cannot delete items already in Deleted Items folder (returns `DELETED_ITEMS_PROTECTED` error)

## Important Notes

- Outlook must be installed and configured on the machine
- The CLI will use whatever account is active in Outlook
- If Outlook prompts for security (accessing address book, sending mail), consider running Outlook as admin or adjusting Trust Center settings
- Signature extraction looks for `<div id="Signature">` or common greeting patterns
