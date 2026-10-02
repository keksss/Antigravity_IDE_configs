---
name: outlook-mail
description: Searches emails, reads messages with attachments, creates and iteratively edits drafts by EntryID, and safely sends emails via Microsoft Outlook. Use when the user asks to "find email", "read mail", "check inbox", "draft email", "edit draft", or "send email".
---

# Outlook Mail Skill

This skill provides full lifecycle email operations via Microsoft Outlook COM automation: searching, reading content and attachments, drafting messages, editing existing drafts by ID, and controlled sending.

## Safety & Human Confirmation Rules

> [!IMPORTANT]
> **CRITICAL RULE**: The agent must **NEVER** run `mail_compose.py --send` without prior explicit approval from the user in chat.
> 1. Always create a draft first (`mail_compose.py --draft` or default).
> 2. Present a clear card to the user: Recipients, Subject, Body summary, and EntryID.
> 3. Ask: *"Черновик создан в Outlook. Вы подтверждаете отправку данного сообщения?"*
> 4. If the user wants changes, update the draft via `--id <entry_id>`.
> 5. Only after affirmative confirmation ("Да", "Отправляй"), execute the `--send` command.

---

## 1. Searching Emails (`scripts/mail_search.py`)

Search and filter emails across folders, dates, senders, subjects, and text queries.

```bash
uv run scripts/mail_search.py [options]
```

### CLI Options
- `--folder <name|path>`: Target folder (default: `Inbox`, supports Russian names like `Входящие`).
- `--query <text>`: Substring search across subject and body.
- `--sender <email|name>`: Filter by sender.
- `--subject <text>`: Filter by subject line.
- `--unread`: Only unread emails.
- `--has-attachments`: Only emails with attachments.
- `--since <date>`: Start date (`YYYY-MM-DD`, `today`, `yesterday`, `7d`).
- `--until <date>`: End date (`YYYY-MM-DD`).
- `--limit <int>`: Max results (default: `20`).
- `--delete`: Preview or perform moving matching emails to 'Deleted Items' (destructive action, requires explicit confirmation).
- `--confirm-delete-ids <ids>`: Comma-separated exact EntryIDs from prior search preview; required to execute `--delete`.
- `--apply`: Execute confirmed deletion; without `--apply`, `--delete` outputs a dry-run preview.
- `--format <text|json>`: Format output (`json` for agent processing, `text` for user display).

### Example
```bash
# Search and preview matching emails
uv run scripts/mail_search.py --query "Отчет" --since 7d --format json

# Delete matching emails after confirmation (preview first, then apply with confirmed IDs)
uv run scripts/mail_search.py --query "Спам" --delete --confirm-delete-ids "00000000..." --apply
```

---

## 2. Reading Emails (`scripts/mail_read.py`)

Retrieve complete details, headers, full body, and optionally save attachments.

```bash
uv run scripts/mail_read.py --id <entry_id> [options]
```

### CLI Options
- `--id <entry_id>`: (Required) MAPI EntryID of the email.
- `--save-attachments <dir>`: Directory where attachments should be saved (e.g. `scratch/downloads/`).
- `--max-body-chars <int>`: Body length limit (default: `50000`).
- `--format <text|json>`

### Example
```bash
uv run scripts/mail_read.py --id "00000000..." --save-attachments "scratch/attachments" --format text
```

---

## 3. Composing and Editing Drafts (`scripts/mail_compose.py`)

Handles creating drafts, iterative modification by ID, and sending.
By default, running without `--apply` prints a **dry-run preview** without touching Outlook. Pass `--apply` to commit changes.

```bash
uv run scripts/mail_compose.py [options]
```

### CLI Options
- `--id <entry_id>`: EntryID of an **existing draft** to edit, send, or delete.
- `--last-draft`: (Recommended) Automatically target the most recently created or edited draft without needing the long hex EntryID.
- `--to <emails>`: Recipient emails (comma-separated).
- `--cc <emails>`: CC recipients.
- `--bcc <emails>`: BCC recipients.
- `--subject <text>`: Email subject.
- `--body <text|filepath>`: Message body (or path to text/markdown/html file).
- `--append-body <text>`: Append text to existing draft body.
- `--html`: Treat body as HTML.
- `--attachment <path>`: File path to attach (repeatable).
- `--inline-image <path>`: Path to image to embed inline via CID (repeatable).
- `--remove-attachment <filename>`: Remove an existing attachment by name.
- `--draft`: Save as draft (default behavior).
- `--send`: Send the email (**STRICT CONFIRMATION REQUIRED**).
- `--delete`: Delete the draft (**STRICT CONFIRMATION REQUIRED**).
- `--apply`: Apply planned changes to Outlook; otherwise runs preview (dry-run).
- `--format <text|json>`

### Common Workflows

#### Workflow A: Create Draft
```bash
# Preview draft (dry-run)
uv run scripts/mail_compose.py \
  --to "partner@example.com" \
  --subject "Встреча по проекту" \
  --body "Добрый день! Направляю статус по задачам..." \
  --format json

# Save draft in Outlook
uv run scripts/mail_compose.py \
  --to "partner@example.com" \
  --subject "Встреча по проекту" \
  --body "Добрый день! Направляю статус по задачам..." \
  --apply \
  --format json
```

#### Workflow B: Edit Existing Draft (Using `--last-draft`)
```bash
uv run scripts/mail_compose.py \
  --last-draft \
  --append-body "P.S. Прошу подтвердить участие до пятницы." \
  --cc "manager@example.com" \
  --apply \
  --format json
```

#### Workflow C: Send (Only After User Confirmation)
```bash
uv run scripts/mail_compose.py --last-draft --send --apply
```

#### Workflow D: Delete Draft (Only After User Confirmation)
```bash
uv run scripts/mail_compose.py --last-draft --delete --apply
```
