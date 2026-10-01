# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "websockets>=12.0",
# ]
# ///

"""
chats_list.py - Lists and searches recent Microsoft Teams chats.
"""

import argparse
import asyncio
import json
import sys
from datetime import datetime
from pathlib import Path

# Setup local shared import
SHARED_DIR = Path(__file__).resolve().parents[3] / "shared"
if str(SHARED_DIR) not in sys.path:
    sys.path.insert(0, str(SHARED_DIR))

import teams_client

def parse_args():
    parser = argparse.ArgumentParser(description="List recent Microsoft Teams chats.")
    parser.add_argument("--limit", type=int, default=15, help="Number of chats to retrieve (default: 15).")
    parser.add_argument("--query", type=str, default=None, help="Filter chats by contact name, group title, or message text.")
    parser.add_argument("--format", choices=["text", "json"], default="text", help="Output format (default: text).")
    return parser.parse_args()

def format_timestamp(ts) -> str:
    if not ts:
        return "N/A"
    try:
        # Check if milliseconds or seconds
        val = float(ts)
        if val > 1e11:
            val /= 1000.0
        return datetime.fromtimestamp(val).strftime("%Y-%m-%d %H:%M")
    except Exception:
        return str(ts)

async def main():
    args = parse_args()
    try:
        chats = await teams_client.list_chats(limit=args.limit, query=args.query)
    except teams_client.TeamsConnectionError as err:
        print(f"Connection Error: {err}", file=sys.stderr)
        sys.exit(1)
    except Exception as err:
        print(f"Error fetching chats: {err}", file=sys.stderr)
        sys.exit(1)

    if args.format == "json":
        print(json.dumps(chats, indent=2, ensure_ascii=False))
        return

    if not chats:
        print("No matching chats found.")
        return

    print(f"\nFound {len(chats)} active chats:\n" + "=" * 80)
    for idx, c in enumerate(chats, 1):
        title = c.get("title", "Без названия")
        ctype = c.get("type", "Chat")
        sender = c.get("last_sender", "")
        last_msg = c.get("last_message", "")
        time_str = format_timestamp(c.get("last_timestamp"))
        chat_id = c.get("id", "")

        print(f"[{idx}] {title} ({ctype})")
        print(f"    ID:        {chat_id}")
        print(f"    Updated:   {time_str}")
        if sender or last_msg:
            print(f"    Last Msg:  {sender + ': ' if sender else ''}{last_msg}")
        print("-" * 80)

if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    asyncio.run(main())
