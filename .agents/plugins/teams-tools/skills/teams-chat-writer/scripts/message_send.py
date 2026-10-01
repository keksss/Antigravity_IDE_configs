# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "websockets>=12.0",
# ]
# ///

"""
message_send.py - Safely sends a message to a Microsoft Teams chat with mandatory confirmation.
"""

import argparse
import asyncio
import json
import sys
from pathlib import Path

# Setup local shared import
SHARED_DIR = Path(__file__).resolve().parents[3] / "shared"
if str(SHARED_DIR) not in sys.path:
    sys.path.insert(0, str(SHARED_DIR))

import teams_client

def parse_args():
    parser = argparse.ArgumentParser(description="Safely send a message to a Teams chat.")
    parser.add_argument("--message", type=str, required=True, help="Message text to send.")
    parser.add_argument("--contact", type=str, default=None, help="Target contact name or email (e.g. 'konstantin.ivanov02@sap.com').")
    parser.add_argument("--chat-id", type=str, default=None, help="Target conversation ID.")
    parser.add_argument("--confirmed", action="store_true", help="Explicit human confirmation flag. Required to actually send.")
    return parser.parse_args()

async def resolve_recipient(chat_id: str | None, contact: str | None) -> tuple[str, str]:
    if chat_id:
        return chat_id, chat_id

    if not contact:
        # Check active window title
        target = teams_client.get_active_target()
        title = target.get("title", "Active Chat")
        return "active", title

    chats = await teams_client.list_chats(limit=10, query=contact)
    if not chats:
        raise ValueError(f"No chat found matching '{contact}'.")

    best = chats[0]
    return best["id"], best.get("title", best["id"])

async def main():
    args = parse_args()

    try:
        target_id, target_title = await resolve_recipient(args.chat_id, args.contact)
    except Exception as err:
        print(f"Error resolving recipient: {err}", file=sys.stderr)
        sys.exit(1)

    # 1. MANDATORY SAFETY CHECK: HUMAN CONFIRMATION
    if not args.confirmed:
        print("\n" + "=" * 80)
        print("⚠️  ЗАПРОС ПОДТВЕРЖДЕНИЯ НА ОТПРАВКУ СООБЩЕНИЯ В MICROSOFT TEAMS")
        print("=" * 80)
        print(f"Получатель / Чат: {target_title}")
        print(f"ID чата:         {target_id}")
        print("-" * 80)
        print("Текст сообщения:")
        print(f"\"{args.message}\"")
        print("-" * 80)
        print("⛔ СТАТУС: СООБЩЕНИЕ НЕ ОТПРАВЛЕНО (Режим Dry-Run).")
        print("Для отправки необходимо явно запросить подтверждение пользователя.")
        print("После получения согласия выполните команду повторно с флагом --confirmed.")
        print("=" * 80 + "\n")
        # Exit code 2 indicates confirmation pending
        sys.exit(2)

    # 2. EXECUTION AFTER CONFIRMATION
    print(f"\n[ПОДТВЕРЖДЕНО] Отправка сообщения в '{target_title}'...")

    try:
        # If target is specific chat and not already active, navigate to it
        if target_id != "active" and "@" in target_id:
            nav_js = f"""
            (() => {{
                window.location.hash = '#/conversations/' + encodeURIComponent({json.dumps(target_id)});
                return true;
            }})()
            """
            await teams_client.evaluate_js(nav_js, await_promise=False)
            await asyncio.sleep(0.5)

        res = await teams_client.send_message_to_active_chat(args.message)
        print("✅ Сообщение успешно отправлено в Microsoft Teams!")
        print(f"   Длина текста: {len(args.message)} символов.")
    except teams_client.TeamsConnectionError as err:
        print(f"Connection Error: {err}", file=sys.stderr)
        sys.exit(1)
    except Exception as err:
        print(f"❌ Ошибка отправки: {err}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    asyncio.run(main())
