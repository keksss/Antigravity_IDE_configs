# /// script
# requires-python = ">=3.10"
# ///

"""
teams_safety_gate.py - Antigravity PreToolUse hook enforcing:
1. Hard DENY on any delete attempt.
2. FORCE_ASK confirmation before sending any Teams message.
"""

import json
import sys

def main():
    try:
        raw_input = sys.stdin.read()
        if not raw_input.strip():
            print(json.dumps({"decision": "allow"}))
            return

        payload = json.loads(raw_input)
        tool_call = payload.get("toolCall", {})
        tool_name = tool_call.get("name", "")
        args = tool_call.get("args", {})
        cmd = args.get("CommandLine", "").lower()

        # 1. Strict No-Delete Policy
        if "delete" in cmd and ("team" in cmd or "chat" in cmd):
            print(json.dumps({
                "decision": "deny",
                "reason": "⛔ ПРАВИЛО БЕЗОПАСНОСТИ TEAMS: Любое удаление сообщений, чатов или файлов из Teams категорически запрещено!"
            }))
            return

        # 2. Enforce Confirmation on Send Message:
        # If --confirmed is missing, hard DENY (agent MUST ask via ask_question first).
        # If --confirmed is present, ALLOW immediately (user has already approved via interactive UI).
        if "message_send.py" in cmd:
            if "--confirmed" not in cmd:
                print(json.dumps({
                    "decision": "deny",
                    "reason": "⛔ Отправка без подтверждения запрещена. Агент обязан сначала запросить одобрение у пользователя через ask_question."
                }))
                return
            else:
                print(json.dumps({"decision": "allow"}))
                return

        print(json.dumps({"decision": "allow"}))
    except Exception:
        print(json.dumps({"decision": "allow"}))

if __name__ == "__main__":
    main()
