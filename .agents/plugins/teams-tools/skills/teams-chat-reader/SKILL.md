---
name: teams-chat-reader
description: Lists Microsoft Teams personal and group chats, reads message history with full timestamps and sender info, and downloads attachments and images locally. Use when the user asks to "check teams chats", "find message in teams", "read teams conversation", or "download file from teams".
---

# Teams Chat Reader Skill

This skill allows listing chats, reading conversation history, exporting chats to Markdown/JSON, and downloading attachments/images from Microsoft Teams via the local WebView2 debug interface (`127.0.0.1:9222`).

---

## 1. List Recent Chats

Use `chats_list.py` to view recent chats or search for a specific conversation:

```bash
# List 10 most recent chats
uv run scripts/chats_list.py --limit 10

# Search chats by contact name, group title, or message keyword
uv run scripts/chats_list.py --query "Ivanov"

# Output as JSON
uv run scripts/chats_list.py --limit 5 --format json
```

---

## 2. Read & Export Chat History

Use `chat_history.py` to retrieve messages, export to Markdown, and download all media attachments:

```bash
# Read messages by contact name or email (console text)
uv run scripts/chat_history.py --contact "Ivanov, Konstantin" --limit 20

# Read messages by exact Conversation ID
uv run scripts/chat_history.py --chat-id "19:b527dc1b-...@unq.gbl.spaces" --limit 30

# Output as JSON
uv run scripts/chat_history.py --contact "Ivanov" --format json

# Export chat to pure text (Markdown without images/media)
uv run scripts/chat_history.py --contact "Konstantin Ustinov" --limit 500 --format markdown --no-media --output "./chat_history.md"

# Export chat to Markdown and download all images/attachments into a media folder
uv run scripts/chat_history.py --contact "Olga Mihailova" --limit 500 --format markdown --output "./chat_history.md" --download-media "./chat_media"
```

---

## 3. Download Specific Attachments

Use `attachment_download.py` when you need to save an individual document or image:

```bash
# Download by filename from a specific chat
uv run scripts/attachment_download.py --chat-id "19:b527dc1b-...@unq.gbl.spaces" --file-name "Report.pdf" --output "./downloads/Report.pdf"

# Download directly by attachment URL
uv run scripts/attachment_download.py --url "<url_from_chat_history>" --output "./downloads/image.png"
```

---

## Safety Guidelines

* **NO DELETE:** This skill is strictly read-only. It cannot delete messages, chats, or files.
* **Local Execution:** Operates entirely over local loopback (`127.0.0.1:9222`).
