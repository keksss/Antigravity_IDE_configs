# Microsoft Teams Tools Plugin (`teams-tools`)

A modular, portable Antigravity plugin for interacting with the desktop **Microsoft Teams** client via the local WebView2 Chromium DevTools Protocol (CDP) debug port (`127.0.0.1:9222`).

---

## Skills in this Plugin

| Skill | Directory | Description |
| :--- | :--- | :--- |
| **`teams-chat-reader`** | [`skills/teams-chat-reader/`](./skills/teams-chat-reader/SKILL.md) | Lists recent personal and group chats, reads conversation history with timestamps and senders, exports to Markdown/JSON, and downloads media attachments. |
| **`teams-chat-writer`** | [`skills/teams-chat-writer/`](./skills/teams-chat-writer/SKILL.md) | Safely sends messages to personal (1:1) and group chats with mandatory human confirmation (`--confirmed`). |

---

## 🛡️ Multi-Level Safety Architecture

The plugin enforces strict safety policies:

1. **Absolute No-Delete Policy (`rules/teams-safety.md`):**
   - The plugin has zero deletion endpoints or commands.
   - Deletion of messages, chats, threads, participants, or attachments is strictly prohibited.
2. **Mandatory Human-in-the-Loop Confirmation:**
   - Every message dispatch requires explicit user confirmation.
   - Running `message_send.py` without `--confirmed` produces a dry-run preview and exits with code 2.
3. **Antigravity Hook Interception (`hooks/scripts/teams_safety_gate.py`):**
   - A `PreToolUse` hook intercepts any command executing `message_send.py` without the `--confirmed` parameter and forces an approval prompt (`force_ask`).

---

## Directory Structure

```text
teams-tools/
├── plugin.json                         # Plugin manifest
├── hooks.json                          # PreToolUse hook configuration
├── README.md                           # Plugin documentation
├── hooks/
│   └── scripts/
│       └── teams_safety_gate.py        # PreToolUse interception script
├── rules/
│   └── teams-safety.md                 # No-delete & human confirmation rules
├── shared/
│   └── teams_client.py                 # Shared CDP WebSocket client and IndexedDB reader
└── skills/
    ├── teams-chat-reader/
    │   ├── SKILL.md
    │   └── scripts/
    │       ├── chats_list.py           # List and search recent conversations
    │       ├── chat_history.py         # Read history, export to Markdown, download media
    │       └── attachment_download.py  # Download individual attachments/images
    └── teams-chat-writer/
        ├── SKILL.md
        └── scripts/
            └── message_send.py         # Send message with mandatory human confirmation
```

---

## Quick Usage Guide

### 1. List Recent Chats
```bash
# List 10 most recent chats
uv run skills/teams-chat-reader/scripts/chats_list.py --limit 10

# Search chats by contact or query
uv run skills/teams-chat-reader/scripts/chats_list.py --query "Project"
```

### 2. Read Conversation History
```bash
# Read chat history with a contact
uv run skills/teams-chat-reader/scripts/chat_history.py --contact "Alexey Smirnov" --limit 20

# Export full chat to Markdown and download media
uv run skills/teams-chat-reader/scripts/chat_history.py --contact "Alexey Smirnov" --limit 100 --format markdown --output "./chat_history.md" --download-media "./chat_media"
```

### 3. Send Messages (Human Confirmation Required)
```bash
# 1. Preview dispatch (dry-run)
uv run skills/teams-chat-writer/scripts/message_send.py --contact "alexey.smirnov@company.com" --message "Привет! Отчет готов."

# 2. Actual sending (ONLY after user confirmation)
uv run skills/teams-chat-writer/scripts/message_send.py --contact "alexey.smirnov@company.com" --message "Привет! Отчет готов." --confirmed
```

---

## Portability & Requirements

- **Runtime:** Microsoft Teams desktop client launched with remote debugging enabled (`--remote-debugging-port=9222`).
- **Dependencies:** Python scripts declare inline dependencies via PEP 723 and run with `uv run`.
- **Portability:** Strictly adheres to `.agents/rules/plugin-architecture.md` with relative paths and zero hardcoded credentials.
