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
outlook mail open <entry-id>                            # Open email in Outlook GUI
outlook mail search --query <text> [--from <email>] [--after <date>] [--before <date>] [--folder <name>]

# Send and draft
outlook mail send --to <emails> [--cc <emails>] --subject <text> [--body <text>] [--body-file <path>] [--html] [--signature-file <path>] [--attachment <files>]
outlook mail draft --to <emails> --subject <text> [--body <text>] [--html] [--signature-file <path>] [--attachment <files>]

# Reply, forward, and conversation
outlook mail reply <entry-id> --body <text> [--reply-all] [--draft] [--html] [--signature-file <path>]
outlook mail forward <entry-id> --to <emails> [--body <text>] [--draft] [--html] [--signature-file <path>]
outlook mail conversation <entry-id> [--limit <n>]            # Get all emails in same thread

# Organize
outlook mail delete <entry-id> [--no-confirm]           # Moves to Deleted Items
outlook mail move <entry-id> --to-folder <name>
outlook mail mark-read <entry-id>                       # Mark as read
outlook mail mark-unread <entry-id>                     # Mark as unread
outlook mail categorize <entry-id> --list               # Show categories
outlook mail categorize <entry-id> --add <name>         # Add a category
outlook mail categorize <entry-id> --remove <name>      # Remove a category
outlook mail categorize <entry-id> --set <names>        # Replace all categories
outlook mail categorize <entry-id> --clear              # Remove all categories

# Attachments and signatures
outlook mail save-attachments <entry-id> [--output <dir>]
outlook mail extract-signature <entry-id> [--output <file>]
```

### Calendar Commands
```bash
outlook calendar list [--start <date>] [--end <date>] [--limit <n>] [--full]
outlook calendar get <entry-id>
outlook calendar open <entry-id>                        # Open event in Outlook GUI
outlook calendar create --subject <text> --start <datetime> --end <datetime> [--location <text>] [--body <text>] [--all-day]
outlook calendar update <entry-id> [--subject <text>] [--start <datetime>] [--end <datetime>] [--location <text>] [--body <text>]
outlook calendar delete <entry-id> [--no-confirm]
outlook calendar respond <entry-id> --accept|--decline|--tentative [--message <text>]

# Availability
outlook calendar free-busy --email <email> [--start <date>] [--end <date>]
outlook calendar find-slots --emails "a@co.com,b@co.com" [--start <date>] [--end <date>] [--duration <min>] [--include-self]
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
dotnet publish -c Release -r win-x64               # Framework-dependent (~1 MB, requires .NET 8 Runtime)
dotnet publish -c Release -r win-x64 --self-contained  # Self-contained (~150 MB, no runtime needed)
```

## Release Process

Releases are automated via GitHub Actions (`.github/workflows/release.yml`).

### To create a new release:
```bash
git tag v1.0.0
git push origin v1.0.0
```

The workflow will:
1. Build and test on `windows-latest`
2. Publish a **framework-dependent** single-file exe (requires .NET 8 Runtime on user's machine)
3. Zip it as `OutlookCLI-<version>-win-x64.zip`
4. Create a GitHub Release with auto-generated release notes and the zip attached

### CI
- `.github/workflows/ci.yml` runs build + test on every push to `main` and on PRs

### Branching
- `master` - stable releases, CI runs here
- `dev` - active development branch

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

1. **Confirmation Guard**: Destructive actions (delete) show a console prompt `[y/N]` unless `--no-confirm` is set
2. **Deleted Items Protection**: Cannot delete items already in Deleted Items folder (returns `DELETED_ITEMS_PROTECTED` error)

## Important Notes

- Outlook must be installed and configured on the machine
- The CLI will use whatever account is active in Outlook
- If Outlook prompts for security (accessing address book, sending mail), consider running Outlook as admin or adjusting Trust Center settings
- Signature extraction looks for `<div id="Signature">` or common greeting patterns
