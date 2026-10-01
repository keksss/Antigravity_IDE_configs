# Microsoft Teams Safety and Integrity Rules

When reading, downloading, or sending messages in Microsoft Teams using the `teams-tools` plugin, adhere to the following safety policies without exception.

---

## 1. Absolute Prohibition on Deletion (Strict No-Delete Policy)

1. **No Deletion Operations Whatsoever:**
   - Under NO circumstances may the agent delete, purge, or remove any message, chat, thread, attachment, team, channel, or participant.
   - The plugin contains NO delete methods, flags, or endpoints.
   - The agent MUST NEVER attempt to write, execute, or simulate deletion commands for Microsoft Teams data.

---

## 2. Mandatory Human Confirmation Before Sending (Human-in-the-Loop)

1. **Explicit Permission Required for Every Message:**
   - The agent MUST NEVER send a message autonomously or silently.
   - Before executing `message_send.py` with the `--confirmed` flag, the agent MUST:
     1. Formulate the exact message text.
     2. Identify the target recipient / chat name and ID.
     3. Present a clear confirmation request to the user with the recipient and the full message text.
     4. Wait for explicit affirmative confirmation from the user (e.g. "Да, отправляй", "Отправь", "Подтверждаю").
   - Running `message_send.py` without `--confirmed` will intentionally fail or produce a dry-run draft preview.

---

## 3. Scope Limitation: Chats Only

1. **Only Personal and Group Chats:**
   - All operations are restricted to direct 1-on-1 chats and group/meeting chats.
   - Modifying team structures, channel configurations, or organization-level policies is strictly out of scope.

---

## 4. Attachment Handling Safety

1. **Non-Destructive Local Downloads:**
   - When downloading attachments or images from chats, files must be saved locally without overwriting existing critical files.
   - Avoid executing or opening downloaded binary files without user verification.
