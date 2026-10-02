# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "pywin32>=306",
#     "python-dateutil>=2.8.2",
# ]
# ///

"""
mail_read.py - Reads full content and metadata of a specific email by EntryID, with optional attachment downloading.
"""

import sys
import os
import argparse
import json
import gc
from pathlib import Path

def parse_args():
    parser = argparse.ArgumentParser(description="Read email details and optionally download attachments.")
    parser.add_argument("--id", type=str, required=True, help="MAPI EntryID of the email to read.")
    parser.add_argument("--save-attachments", type=str, default="", help="Directory to save downloaded attachments.")
    parser.add_argument("--max-body-chars", type=int, default=50000, help="Maximum characters of body text to return.")
    parser.add_argument("--format", choices=["text", "json"], default="text", help="Output format.")
    return parser.parse_args()

def main():
    if sys.stdout.encoding != 'utf-8':
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except (AttributeError, OSError, ValueError) as exc:
            print(f"Warning: UTF-8 console configuration failed: {exc}", file=sys.stderr)

    args = parse_args()
    outlook = None
    namespace = None

    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()

        try:
            outlook = win32com.client.GetActiveObject("Outlook.Application")
        except Exception:
            outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        try:
            item = namespace.GetItemFromID(args.id)
        except Exception as e:
            raise ValueError(f"Could not retrieve email with EntryID '{args.id}': {e}")

        subject = getattr(item, "Subject", "") or ""
        sender_name = getattr(item, "SenderName", "") or ""
        sender_email = ""
        try:
            sender_email = getattr(item, "SenderEmailAddress", "") or ""
        except Exception as exc:
            print(f"Warning: sender address unavailable: {exc}", file=sys.stderr)

        to_recipients = getattr(item, "To", "") or ""
        cc_recipients = getattr(item, "CC", "") or ""
        bcc_recipients = getattr(item, "BCC", "") or ""
        received_time = str(getattr(item, "ReceivedTime", ""))
        importance = getattr(item, "Importance", 1)  # 0: Low, 1: Normal, 2: High
        unread = getattr(item, "UnRead", False)
        categories = getattr(item, "Categories", "") or ""
        body_text = getattr(item, "Body", "") or ""

        if len(body_text) > args.max_body_chars:
            body_text = body_text[:args.max_body_chars] + f"\n... [Truncated: {len(body_text) - args.max_body_chars} characters omitted]"

        # Process attachments
        attachments_info = []
        save_dir = None
        if args.save_attachments:
            save_dir = Path(args.save_attachments).resolve()
            save_dir.mkdir(parents=True, exist_ok=True)

        if hasattr(item, "Attachments"):
            for i in range(1, item.Attachments.Count + 1):
                att = item.Attachments.Item(i)
                filename = att.FileName
                size = getattr(att, "Size", 0)
                saved_path = None

                if save_dir:
                    target_file = save_dir / filename
                    # Prevent overwriting by appending counter if needed
                    counter = 1
                    stem = target_file.stem
                    suffix = target_file.suffix
                    while target_file.exists():
                        target_file = save_dir / f"{stem}_{counter}{suffix}"
                        counter += 1

                    att.SaveAsFile(str(target_file))
                    saved_path = str(target_file)

                attachments_info.append({
                    "name": filename,
                    "size_bytes": size,
                    "saved_path": saved_path
                })

        output_data = {
            "status": "success",
            "entry_id": args.id,
            "subject": subject,
            "from": f"{sender_name} <{sender_email}>".strip(),
            "to": to_recipients,
            "cc": cc_recipients,
            "bcc": bcc_recipients,
            "date": received_time,
            "unread": unread,
            "importance": {0: "Low", 1: "Normal", 2: "High"}.get(importance, "Normal"),
            "categories": categories,
            "attachments_count": len(attachments_info),
            "attachments": attachments_info,
            "body": body_text
        }

        if args.format == "json":
            print(json.dumps(output_data, indent=2, ensure_ascii=False))
        else:
            print("=" * 60)
            print(f"Subject: {subject}")
            print(f"From:    {output_data['from']}")
            print(f"To:      {to_recipients}")
            if cc_recipients:
                print(f"CC:      {cc_recipients}")
            print(f"Date:    {received_time}")
            if categories:
                print(f"Tags:    {categories}")
            if attachments_info:
                print(f"Attachments ({len(attachments_info)}):")
                for att in attachments_info:
                    saved_msg = f" -> saved to {att['saved_path']}" if att['saved_path'] else ""
                    print(f"  - {att['name']} ({att['size_bytes']} bytes){saved_msg}")
            print("=" * 60)
            print("\n" + body_text)

    except Exception as e:
        err_payload = {"status": "error", "message": str(e)}
        if args.format == "json":
            print(json.dumps(err_payload, ensure_ascii=False))
        else:
            print(f"Error reading email: {e}", file=sys.stderr)
        sys.exit(1)
    finally:
        del namespace
        del outlook
        gc.collect()
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception as exc:
            print(f"Warning: Outlook COM cleanup failed: {exc}", file=sys.stderr)

if __name__ == "__main__":
    main()
