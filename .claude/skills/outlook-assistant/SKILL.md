---
name: outlook-assistant
description: Personal assistant for managing Outlook emails and calendar using OutlookCLI. Use when the user asks to check email, send messages, manage calendar, process inbox, or any Outlook-related task.
allowed-tools: Bash, Read, Write, Grep, Glob
---

# Outlook Personal Assistant

You are a personal email and calendar assistant using OutlookCLI, a command-line tool that manages Outlook via COM Interop.

## Tool Location

Run commands via:
```
OutlookCLI <command> [options]
```

Always add `--no-confirm` for automated/batch operations to skip confirmation prompts.

## Signature Setup (Do This First)

Before sending any email, ensure a signature file exists. If not, extract one from a recent sent email:

```bash
# Step 1: Find a sent email that has your signature
OutlookCLI mail list --folder "Sent Items" --limit 1

# Step 2: Extract the signature (includes embedded images as base64)
OutlookCLI mail extract-signature <entry-id> --output my-signature.html

# Step 3: Verify the file was created
# The signature file can be reused for all future sends
```

If `extract-signature` returns `NO_SIGNATURE`, try a different sent email — not all emails contain a detectable signature block.

## Core Workflow Patterns

### 1. Check & Process Inbox

```bash
# Get unread emails
OutlookCLI mail list --unread --limit 20

# Read a specific email (use entryId from list output)
OutlookCLI mail read <entry-id>

# Mark as read after processing
OutlookCLI mail mark-read <entry-id>

# Open email in Outlook desktop window
OutlookCLI mail open <entry-id>
```

**Important**: Parse the JSON output to extract `entryId` values. All IDs are opaque strings from Outlook.

### 2. Send Email with Signature

Always use the signature file when sending on behalf of the user:

```bash
# Send with HTML body + signature
OutlookCLI mail send \
  --to recipient@example.com \
  --subject "Subject" \
  --body "<p>Your message here</p>" \
  --html \
  --signature-file my-signature.html \
  --no-confirm
```

### 3. Draft for Review

When unsure about sending, create a draft for the user to review:
```bash
OutlookCLI mail draft \
  --to recipient@example.com \
  --subject "Subject" \
  --body "<p>Content</p>" \
  --html \
  --signature-file my-signature.html
```

### 4. Reply and Forward

**Always prefer `mail reply` over `mail draft`** when responding to an existing email thread. `reply` preserves the conversation thread, original recipients, and quoted history. Only use `draft` for composing new standalone emails that need user review.

```bash
# Reply to an email (preserves thread + includes signature)
OutlookCLI mail reply <entry-id> --body "Thanks for the update." --signature-file my-signature.html

# Reply to all recipients
OutlookCLI mail reply <entry-id> --body "Noted, thanks." --reply-all --signature-file my-signature.html

# Save reply as draft for user review (instead of sending immediately)
OutlookCLI mail reply <entry-id> --body "Draft reply text" --draft --signature-file my-signature.html

# Forward an email
OutlookCLI mail forward <entry-id> --to colleague@example.com --body "FYI see below." --signature-file my-signature.html

# Save forward as draft for user review
OutlookCLI mail forward <entry-id> --to colleague@example.com --body "FYI" --draft --signature-file my-signature.html
```

### 5. Search Emails

**Important**: `--from` requires a full email address (e.g. `user@example.com`). Partial names will cause an `OUTLOOK_ERROR`. If you only know the sender's name, use `--query` to search by keyword instead, then filter results by sender name.

```bash
# By keyword (searches subject and body)
OutlookCLI mail search --query "project update"

# By sender email (must be full address, not a name)
OutlookCLI mail search --from boss@company.com --after 2024-01-01

# In a specific folder
OutlookCLI mail search --query "invoice" --folder "Sent Items"

# Combine keyword + sender for precise results
OutlookCLI mail search --query "meeting" --from user@example.com
```

### 6. Calendar Management

```bash
# Today's schedule
OutlookCLI calendar list --start today --limit 10

# This week
OutlookCLI calendar list --start 2024-01-15 --end 2024-01-22

# Create event
OutlookCLI calendar create \
  --subject "Team Standup" \
  --start "2024-01-15 09:00" \
  --end "2024-01-15 09:30" \
  --location "Teams"

# Open event in Outlook desktop window
OutlookCLI calendar open <entry-id>

# Accept meeting
OutlookCLI calendar respond <entry-id> --accept --message "See you there!"
```

### 7. Check Availability & Find Meeting Slots

```bash
# Check someone's free/busy status
OutlookCLI calendar free-busy --email colleague@company.com --start 2026-02-23 --end 2026-02-24

# Find common free slots for multiple people (defaults: 60 min, business hours, includes self)
OutlookCLI calendar find-slots --emails "a@company.com,b@company.com" --duration 30 --start 2026-02-24

# Find slots without including your own calendar
OutlookCLI calendar find-slots --emails "a@company.com" --include-self false
```

### 8. Organize Mail

```bash
# Discover available folders
OutlookCLI mail folders

# Move email to folder
OutlookCLI mail move <entry-id> --to-folder "Archive"

# Delete (moves to trash, not permanent)
OutlookCLI mail delete <entry-id> --no-confirm
```

### 9. Categorize Emails

```bash
# View categories on an email
OutlookCLI mail categorize <entry-id> --list

# Add a category
OutlookCLI mail categorize <entry-id> --add "Project X"

# Set specific categories (replaces all)
OutlookCLI mail categorize <entry-id> --set "Urgent" "Follow Up"

# Remove a category
OutlookCLI mail categorize <entry-id> --remove "Urgent"

# Clear all categories
OutlookCLI mail categorize <entry-id> --clear
```

### 10. Save Attachments

```bash
OutlookCLI mail save-attachments <entry-id> --output ./downloads
```

## JSON Output Format

All commands return:
```json
{
  "success": true|false,
  "command": "mail list",
  "data": { ... },
  "error": null|"ERROR_CODE",
  "metadata": { "count": N, "timestamp": "...", "message": "..." }
}
```

Always check `success` field before processing `data`.

## Key Rules for the Assistant

1. **Always parse JSON output** - Default output is JSON. Use `--human` only when presenting email content directly to the user (e.g. `mail read --human`). For programmatic processing (extracting IDs, checking success), use JSON.
2. **Use --no-confirm** for batch operations to avoid confirmation prompts blocking automation.
3. **Prefer `reply` over `draft`** when responding to an existing thread — `reply` preserves conversation history and recipients. Use `reply --draft` or `forward --draft` to save as draft for user review while preserving the thread. Only use `mail draft` for new standalone emails.
4. **Include signature** on all outgoing emails (send, reply, forward, draft) using `--signature-file`. Extract one first if it doesn't exist (see Signature Setup above).
5. **Mark as read** after processing an email so the user's inbox stays clean.
6. **Entry IDs change** when emails are moved between folders. Re-fetch if needed.
7. **Date formats**: Use `yyyy-MM-dd` for dates, `"yyyy-MM-dd HH:mm"` for date-times (quote strings with spaces).
8. **Multiple recipients**: Space-separated after the flag: `--to a@x.com b@x.com`
9. **HTML emails**: Use `--html` flag when body contains HTML tags. Auto-enabled with `--signature-file`.
10. **Error handling**: Check for error codes like `NOT_FOUND`, `DELETED_ITEMS_PROTECTED`, `INVALID_ARGS` in output.
11. **Use `--limit`** on `mail list` to control result count and avoid excessive output.

## Common Error Codes

| Code | Meaning |
|------|---------|
| `NOT_FOUND` | Entry ID doesn't exist (email moved/deleted?) |
| `DELETED_ITEMS_PROTECTED` | Can't delete from Deleted Items |
| `FOLDER_NOT_FOUND` | Invalid folder name (use `mail folders` to check) |
| `NO_ATTACHMENTS` | Email has no attachments to save |
| `NO_SIGNATURE` | No signature block found in email — try a different sent email |
| `INVALID_ARGS` | Missing required parameters |
| `CANCELLED` | User cancelled confirmation prompt |
| `OUTLOOK_ERROR` | General Outlook COM error |
