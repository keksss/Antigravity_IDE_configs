# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "websockets>=12.0",
# ]
# ///

"""
attachment_download.py - Downloads attachments and images from Microsoft Teams chats using authenticated session.
"""

import argparse
import asyncio
import os
import sys
from pathlib import Path

# Setup local shared import
SHARED_DIR = Path(__file__).resolve().parents[3] / "shared"
if str(SHARED_DIR) not in sys.path:
    sys.path.insert(0, str(SHARED_DIR))

import teams_client

def parse_args():
    parser = argparse.ArgumentParser(description="Download an attachment or image from a Teams chat.")
    parser.add_argument("--url", type=str, default=None, help="Direct download URL or objectUrl from message attachment.")
    parser.add_argument("--chat-id", type=str, default=None, help="Chat ID to search attachment in.")
    parser.add_argument("--file-name", type=str, default=None, help="Attachment file name to look up in the chat.")
    parser.add_argument("--output", type=str, required=True, help="Destination file path (e.g. './downloads/document.pdf').")
    return parser.parse_args()

async def resolve_download_url(url: str | None, chat_id: str | None, file_name: str | None) -> str:
    if url:
        return url

    if not chat_id or not file_name:
        raise ValueError("Must specify either --url, or both --chat-id and --file-name.")

    messages = await teams_client.get_chat_history(chat_id, limit=100)
    for m in messages:
        for att in m.get("attachments", []):
            if file_name.lower() in att.get("name", "").lower():
                found_url = att.get("url")
                if found_url:
                    return found_url

    raise ValueError(f"File '{file_name}' not found in chat '{chat_id}' attachments.")

async def main():
    args = parse_args()
    try:
        download_url = await resolve_download_url(args.url, args.chat_id, args.file_name)
        out_path = Path(args.output).resolve()
        print(f"Downloading: {download_url[:80]}...")
        print(f"Destination: {out_path}")
        saved_file = await teams_client.download_attachment_to_file(download_url, str(out_path))
        file_size = os.path.getsize(saved_file)
        print(f"✅ Successfully downloaded {file_size} bytes to: {saved_file}")
    except teams_client.TeamsConnectionError as err:
        print(f"Connection Error: {err}", file=sys.stderr)
        sys.exit(1)
    except Exception as err:
        print(f"Error downloading attachment: {err}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    asyncio.run(main())
