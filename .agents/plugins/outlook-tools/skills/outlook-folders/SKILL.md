---
name: outlook-folders
description: Lists Microsoft Outlook accounts, data stores, and folder hierarchy with item counts and unread badges. Use when the user asks to "list outlook folders", "check unread mail folders", "find outlook folder path", or explore the mailbox tree.
---

# Outlook Folders Skill

This skill allows inspecting Outlook mailboxes, accounts, and directory trees, providing counts of total and unread messages.

## Execution

All scripts must be executed via `uv run` with isolated dependencies:

```bash
uv run scripts/outlook_folders.py [options]
```

## CLI Parameters

- `--account <name>`: Filter by specific mailbox or account name (case-insensitive substring).
- `--unread-only`: Display only folders containing unread items (or folders with subfolders having unread items).
- `--depth <int>`: Traversal depth limit (default: `4`).
- `--format <text|json>`: Output format. Use `json` for programmatic agent processing, `text` for readable user summaries.

## Examples

### 1. View Entire Mailbox Folder Tree
```bash
uv run scripts/outlook_folders.py --format text
```

### 2. Check Only Folders with Unread Messages
```bash
uv run scripts/outlook_folders.py --unread-only --format text
```

### 3. Get JSON Hierarchy for a Specific Account
```bash
uv run scripts/outlook_folders.py --account "Corporate" --format json
```

## Output Structure (JSON)

```json
{
  "status": "success",
  "accounts": [
    {
      "account": "user@company.com",
      "folders": [
        {
          "name": "Inbox",
          "path": "\\\\user@company.com\\Inbox",
          "entry_id": "00000000...",
          "total_items": 154,
          "unread_items": 3,
          "type": "Mail",
          "has_unread": true,
          "folders": []
        }
      ]
    }
  ]
}
```
