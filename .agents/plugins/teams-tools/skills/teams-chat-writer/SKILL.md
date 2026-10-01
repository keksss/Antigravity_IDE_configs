---
name: teams-chat-writer
description: Safely sends messages to Microsoft Teams personal (1:1) and group chats with MANDATORY human confirmation. The agent MUST ALWAYS show the draft to the user and obtain explicit permission before executing with --confirmed. Use when the user asks to "send message in teams", "write to colleague in teams", "reply in teams chat", or "notify in teams".
---

# Teams Chat Writer Skill

This skill safely sends messages to Microsoft Teams personal and group chats via the local WebView2 debug interface (`127.0.0.1:9222`).

---

## ⚠️ MANDATORY SAFETY POLICY: HUMAN CONFIRMATION

**CRITICAL RULE:**
You must **NEVER** run `message_send.py --confirmed` without explicit, prior user approval using the `ask_question` tool.

1. **Step 1 (Formulate Draft):**  
   Determine target recipient, chat ID, and message text.
2. **Step 2 (Ask Confirmation via `ask_question`):**  
   Call `ask_question` with a clear, readable Markdown question:
   ```json
   {
     "questions": [
       {
         "question": "Подтвердите отправку сообщения в Microsoft Teams:\n\n**Кому:** Иванов Константин (konstantin.ivanov02@sap.com)\n**Текст:**\n> Привет! Отчет готов.\n",
         "options": [
           "(Recommended) Отправить сообщение",
           "Отменить / Изменить текст"
         ],
         "is_multi_select": false
       }
     ]
   }
   ```
3. **Step 3 (Execute):**  
   ONLY when the user selects "Отправить сообщение", execute `message_send.py --confirmed`.
   Because the user already approved via `ask_question`, the execution runs smoothly without bulky IDE system modals.

---

## Usage Examples

```bash
# 1. Dry-run preview (verifies chat resolution without sending)
uv run scripts/message_send.py --contact "konstantin.ivanov02@sap.com" --message "Привет! Отчет готов."

# 2. Actual sending (ONLY after user confirmation!)
uv run scripts/message_send.py --contact "konstantin.ivanov02@sap.com" --message "Привет! Отчет готов." --confirmed

# 3. Sending by exact Conversation ID
uv run scripts/message_send.py --chat-id "19:b527dc1b-...@unq.gbl.spaces" --message "Согласовано." --confirmed
```

---

## Strict Prohibitions

* **NO DELETE:** This skill only sends new messages. It CANNOT delete, unsend, or modify past messages.
* **NO SILENT SENDING:** Every message must be reviewed and approved by the user.
