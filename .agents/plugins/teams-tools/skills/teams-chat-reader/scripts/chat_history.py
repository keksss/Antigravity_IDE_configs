# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "websockets>=12.0",
# ]
# ///

"""
chat_history.py - Reads message history, exports to Markdown/JSON/text, and optionally downloads attachments/images.
"""

import argparse
import asyncio
import json
import os
import sys
from datetime import datetime
from pathlib import Path

# Setup local shared import
SHARED_DIR = Path(__file__).resolve().parents[3] / "shared"
if str(SHARED_DIR) not in sys.path:
    sys.path.insert(0, str(SHARED_DIR))

import teams_client

def parse_args():
    parser = argparse.ArgumentParser(
        description="Read message history, export to text/json/markdown, and download attachments from Microsoft Teams."
    )
    parser.add_argument("--chat-id", type=str, default=None, help="The conversation ID (e.g. '19:...@unq.gbl.spaces').")
    parser.add_argument("--contact", type=str, default=None, help="Contact name or email to auto-resolve chat ID.")
    parser.add_argument("--limit", type=int, default=50, help="Number of messages to retrieve (default: 50).")
    parser.add_argument(
        "--format",
        choices=["text", "json", "markdown", "md"],
        default="text",
        help="Output format: text, json, or markdown (default: text)."
    )
    parser.add_argument("--output", "-o", type=str, default=None, help="Optional output file path. If omitted, prints to stdout.")
    parser.add_argument(
        "--download-media",
        type=str,
        default=None,
        help="Optional directory path to download all attached media/images and link them locally."
    )
    parser.add_argument(
        "--no-media",
        action="store_true",
        help="Exclude all images, attachments, and media links completely (pure text)."
    )
    return parser.parse_args()

def format_timestamp(ts) -> str:
    if not ts:
        return "N/A"
    try:
        val = float(ts)
        if val > 1e11:
            val /= 1000.0
        return datetime.fromtimestamp(val).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return str(ts)

async def resolve_chat_id(target_chat_id: str | None, contact_query: str | None) -> tuple[str, str]:
    if target_chat_id:
        return target_chat_id, target_chat_id

    if not contact_query:
        raise ValueError("Either --chat-id or --contact must be specified.")

    chats = await teams_client.list_chats(limit=10, query=contact_query)
    if not chats:
        raise ValueError(f"No chat found matching contact '{contact_query}'.")

    best = chats[0]
    return best["id"], best.get("title", best["id"])

def sanitize_filename(name: str) -> str:
    for ch in ['\\', '/', ':', '*', '?', '"', '<', '>', '|']:
        name = name.replace(ch, "_")
    return name.strip()

async def download_attachments(messages: list[dict], media_dir: Path, output_file: Path | None):
    media_dir.mkdir(parents=True, exist_ok=True)
    total_downloaded = 0

    for m in messages:
        attachments = m.get("attachments", [])
        if not attachments:
            continue

        for idx, att in enumerate(attachments, 1):
            url = att.get("url", "")
            if not url:
                continue

            raw_name = att.get("name", "")
            is_image = "image" in att.get("type", "").lower()
            if not raw_name or raw_name == "image":
                raw_name = f"img_{m['id']}_{idx}.png" if is_image else f"file_{m['id']}_{idx}.dat"

            safe_name = sanitize_filename(raw_name)
            local_target = media_dir / safe_name
            
            # Relative path from output file if available, otherwise from CWD
            if output_file:
                try:
                    rel_link = os.path.relpath(local_target, output_file.parent).replace("\\", "/")
                except Exception:
                    rel_link = str(local_target).replace("\\", "/")
            else:
                rel_link = os.path.relpath(local_target, Path.cwd()).replace("\\", "/")

            # Avoid re-downloading identical files if already present
            if not local_target.exists():
                try:
                    print(f"Downloading media [{safe_name}]...", file=sys.stderr)
                    await teams_client.download_attachment_to_file(url, str(local_target))
                    total_downloaded += 1
                except Exception as exc:
                    print(f"  Warning: failed to download {safe_name}: {exc}", file=sys.stderr)
                    att["download_error"] = str(exc)
            else:
                total_downloaded += 1

            if local_target.exists():
                att["local_path"] = str(local_target)
                att["rel_link"] = rel_link
            else:
                att["local_path"] = None
                att["rel_link"] = None

    return total_downloaded

def format_as_markdown(chat_title: str, chat_id: str, messages: list[dict], no_media: bool = False) -> str:
    lines = [
        f"# История переписки: {chat_title}",
        "",
        f"- **Чат / Собеседник:** {chat_title}",
        f"- **ID чата:** `{chat_id}`",
        f"- **Дата экспорта:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"- **Всего сообщений:** {len(messages)}",
        "",
        "---",
        ""
    ]

    for m in messages:
        sender = m.get("sender", "Unknown")
        time_str = format_timestamp(m.get("timestamp"))
        text = m.get("text", "")
        attachments = m.get("attachments", [])

        lines.append(f"### {sender} — *{time_str}*")
        lines.append("")
        if text:
            lines.append(text)
            lines.append("")

        if not no_media and attachments:
            for att in attachments:
                if att.get("download_error"):
                    continue
                name = att.get("name", "attachment")
                is_image = "image" in att.get("type", "").lower()
                rel_link = att.get("rel_link")
                url = att.get("url")
                target = rel_link or url

                if target:
                    if is_image:
                        lines.append(f"![{name}]({target})")
                    else:
                        lines.append(f"[📎 {name}]({target})")
                    lines.append("")

        lines.append("---")
        lines.append("")

    return "\n".join(lines)

def format_as_text(chat_title: str, chat_id: str, messages: list[dict], no_media: bool = False) -> str:
    lines = [
        f"\nChat History: {chat_title}",
        f"ID: {chat_id}",
        f"Messages displayed: {len(messages)}\n" + "=" * 80
    ]

    if not messages:
        lines.append("No cached messages found for this chat in the local store.")
        lines.append("(Tip: Open this chat in Teams to populate the local cache if it is older).")
        return "\n".join(lines)

    for m in messages:
        sender = m.get("sender", "Unknown")
        time_str = format_timestamp(m.get("timestamp"))
        text = m.get("text", "")
        attachments = m.get("attachments", [])

        lines.append(f"[{time_str}] {sender}:")
        if text:
            lines.append(f"  {text}")
        if not no_media and attachments:
            lines.append("  📎 Вложения:")
            for a in attachments:
                name = a.get("name", "unnamed")
                size = a.get("size", 0)
                url = a.get("url", "")
                local_path = a.get("local_path", "")
                size_kb = f" ({size // 1024} KB)" if size else ""
                loc_info = f" -> {local_path}" if local_path else f" [URL: {url[:60]}...]"
                lines.append(f"    - {name}{size_kb}{loc_info}")
        lines.append("-" * 80)

    return "\n".join(lines)

async def main():
    args = parse_args()
    try:
        chat_id, chat_title = await resolve_chat_id(args.chat_id, args.contact)
        messages = await teams_client.get_chat_history(chat_id, limit=args.limit)
    except teams_client.TeamsConnectionError as err:
        print(f"Connection Error: {err}", file=sys.stderr)
        sys.exit(1)
    except Exception as err:
        print(f"Error fetching history: {err}", file=sys.stderr)
        sys.exit(1)

    out_file = Path(args.output).resolve() if args.output else None

    # Handle media downloads if requested
    if args.download_media and not args.no_media:
        media_path = Path(args.download_media).resolve()
        downloaded = await download_attachments(messages, media_path, out_file)
        print(f"Downloaded/verified {downloaded} media files in {media_path}", file=sys.stderr)

    fmt = args.format.lower()
    if fmt == "json":
        content = json.dumps({
            "chat_id": chat_id,
            "chat_title": chat_title,
            "count": len(messages),
            "messages": messages
        }, indent=2, ensure_ascii=False)
    elif fmt in ["markdown", "md"]:
        content = format_as_markdown(chat_title, chat_id, messages, no_media=args.no_media)
    else:
        content = format_as_text(chat_title, chat_id, messages, no_media=args.no_media)

    if out_file:
        out_file.parent.mkdir(parents=True, exist_ok=True)
        out_file.write_text(content, encoding="utf-8")
        print(f"Exported {len(messages)} messages to {out_file}", file=sys.stderr)
    else:
        print(content)

if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    asyncio.run(main())
