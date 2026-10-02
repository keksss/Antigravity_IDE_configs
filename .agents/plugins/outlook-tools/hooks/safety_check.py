#!/usr/bin/env python3
"""
PreToolUse safety hook for outlook-tools plugin.
Intercepts `run_command` calls that involve sending emails, cancelling/deleting meetings,
or deleting items, enforcing an explicit confirmation prompt ('force_ask').
"""

import sys
import json
import re

DANGEROUS_ACTIONS = [
    # Sending emails
    (r"mail_compose(\.py)?\b.*--send\b", "📤 Отправка письма в Outlook"),
    # Deleting drafts or emails
    (r"mail_compose(\.py)?\b.*--delete\b", "🗑️ Удаление черновика в Outlook"),
    (r"mail_search(\.py)?\b.*--delete\b", "🗑️ Удаление писем в Outlook"),
    # Deleting or cancelling calendar meetings/appointments
    (r"calendar_manage(\.py)?\b.*(--cancel|--delete)\b", "❌ Отмена встречи календаря в Outlook"),
    # Sending calendar invitations
    (r"calendar_manage(\.py)?\b.*--send\b", "📅 Рассылка приглашений на встречу в Outlook"),
    # General delete
    (r"outlook_.*\b.*--delete\b", "🗑️ Удаление данных в Outlook"),
]

def extract_context(cmd_line: str) -> str:
    parts = []
    # Extract --folder if present
    folder_match = re.search(r'--folder\s+["\']?([^"\'\s]+)', cmd_line)
    if folder_match:
        parts.append(f"Папка: {folder_match.group(1)}")
    # Extract --limit or --all if present
    limit_match = re.search(r'--limit\s+(\d+)', cmd_line)
    if limit_match:
        parts.append(f"Количество: {limit_match.group(1)}")
    elif re.search(r'\b--all\b', cmd_line):
        parts.append("Все сообщения в папке")
    # Extract --to if present
    to_match = re.search(r'--to\s+["\']?([^"\'\s]+)', cmd_line)
    if to_match:
        parts.append(f"Кому: {to_match.group(1)}")
    # Extract --subject if present
    sub_match = re.search(r'--subject\s+["\']([^"\']+)["\']', cmd_line)
    if sub_match:
        parts.append(f'"{sub_match.group(1)}"')
    return f" ({', '.join(parts)})" if parts else ""

def check_command(cmd_line: str):
    for pattern, title in DANGEROUS_ACTIONS:
        if re.search(pattern, cmd_line, re.IGNORECASE):
            ctx = extract_context(cmd_line)
            return True, f"{title}{ctx}"
    return False, ""

def main():
    try:
        if sys.stdout.encoding != 'utf-8':
            try:
                sys.stdout.reconfigure(encoding='utf-8')
            except (AttributeError, OSError, ValueError) as exc:
                print(f"Warning: UTF-8 console configuration failed: {exc}", file=sys.stderr)

        raw_input = sys.stdin.read()
        if not raw_input.strip():
            raise ValueError("Empty hook input")

        payload = json.loads(raw_input)
        if not isinstance(payload, dict) or not isinstance(payload.get("toolCall"), dict):
            raise ValueError("Missing toolCall")
        tool_call = payload.get("toolCall", {})
        tool_name = tool_call.get("name", "")
        args = tool_call.get("args", {})

        if tool_name == "run_command":
            cmd_line = args.get("CommandLine", "")
            is_dangerous, reason_desc = check_command(cmd_line)
            if is_dangerous:
                output = {
                    "decision": "force_ask",
                    "reason": reason_desc
                }
                print(json.dumps(output, ensure_ascii=True))
                return

        print(json.dumps({"decision": "allow"}))
    except Exception as e:
        print(json.dumps({"decision": "deny", "reason": f"Invalid safety hook input: {e}"}, ensure_ascii=True))

if __name__ == "__main__":
    main()
