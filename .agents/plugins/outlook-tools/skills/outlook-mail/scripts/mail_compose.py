# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "pywin32>=306",
#     "python-dateutil>=2.8.2",
# ]
# ///

"""
mail_compose.py - Creates drafts, edits existing drafts by EntryID, and safely sends emails via Outlook COM.
"""

import sys
import os
import argparse
import json
import gc
from pathlib import Path

def get_state_file() -> Path:
    # plugins/outlook-tools/.state/last_draft_id.txt
    base = Path(__file__).resolve().parent.parent.parent.parent
    state_dir = base / ".state"
    state_dir.mkdir(parents=True, exist_ok=True)
    return state_dir / "last_draft_id.txt"

def save_last_draft_id(entry_id: str):
    try:
        get_state_file().write_text(entry_id.strip(), encoding="utf-8")
    except Exception:
        pass

def get_last_draft_id() -> str:
    try:
        f = get_state_file()
        if f.exists():
            return f.read_text(encoding="utf-8").strip()
    except Exception:
        pass
    return ""

def clear_last_draft_id():
    try:
        f = get_state_file()
        if f.exists():
            f.unlink()
    except Exception:
        pass

def parse_args():
    parser = argparse.ArgumentParser(description="Compose, edit drafts, or send emails via Outlook.")
    parser.add_argument("--id", type=str, default="", help="EntryID of an existing draft to edit or send.")
    parser.add_argument("--last-draft", action="store_true", default=False, help="Use the most recently created or modified draft.")
    parser.add_argument("--to", type=str, default="", help="Recipient emails separated by comma or semicolon.")
    parser.add_argument("--cc", type=str, default="", help="CC recipients separated by comma or semicolon.")
    parser.add_argument("--bcc", type=str, default="", help="BCC recipients separated by comma or semicolon.")
    parser.add_argument("--subject", type=str, default="", help="Email subject line.")
    parser.add_argument("--body", type=str, default="", help="Email body text (or path to text/markdown/html file).")
    parser.add_argument("--append-body", type=str, default="", help="Text to append to existing draft body.")
    parser.add_argument("--html", action="store_true", help="Treat body as HTML.")
    parser.add_argument("--attachment", action="append", default=[], help="Path to file to attach (can be repeated).")
    parser.add_argument("--inline-image", action="append", default=[], help="Path to image file to embed inline via CID (repeatable).")
    parser.add_argument("--remove-attachment", type=str, default="", help="Filename of attachment to remove from draft.")
    parser.add_argument("--draft", action="store_true", default=False, help="Explicitly save as draft (default behavior).")
    parser.add_argument("--send", action="store_true", default=False, help="Send the email (requires user confirmation!).")
    parser.add_argument("--delete", action="store_true", default=False, help="Delete the draft email (requires user confirmation!).")
    parser.add_argument("--format", choices=["text", "json"], default="text", help="Output format.")
    return parser.parse_args()

def read_body_content(body_arg: str) -> str:
    if not body_arg:
        return ""
    if body_arg == "-":
        # Read directly from standard input (stdin)
        try:
            if hasattr(sys.stdin, "reconfigure"):
                sys.stdin.reconfigure(encoding="utf-8")
            content = sys.stdin.read()
        except Exception:
            content = ""
    else:
        # Check if body is an existing file path
        p = Path(body_arg)
        if p.is_file():
            try:
                content = p.read_text(encoding="utf-8")
            except Exception:
                try:
                    content = p.read_text(encoding="cp1251")
                except Exception:
                    content = body_arg
            # Auto-cleanup temporary file if stored inside .scratch directory
            if ".scratch" in p.parts:
                try:
                    p.unlink()
                except Exception:
                    pass
        else:
            content = body_arg

    # Normalize literal escaped newlines and standard newlines to Windows CRLF
    normalized = content.replace(r"\r\n", "\n").replace(r"\n", "\n").replace(r"\t", "\t")
    return normalized.replace("\r\n", "\n").replace("\n", "\r\n")

def main():
    if sys.stdout.encoding != 'utf-8':
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except Exception:
            pass

    args = parse_args()
    outlook = None
    namespace = None
    mail_item = None

    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()

        try:
            outlook = win32com.client.GetActiveObject("Outlook.Application")
        except Exception:
            outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        if args.last_draft and not args.id:
            args.id = get_last_draft_id()
            if not args.id:
                raise ValueError("No recently saved draft found in local state. Please specify --id explicitly.")

        is_update = bool(args.id)
        if is_update:
            try:
                mail_item = namespace.GetItemFromID(args.id)
            except Exception as e:
                raise ValueError(f"Could not find draft with EntryID '{args.id}': {e}")
        else:
            # Create new mail item (0 = olMailItem)
            mail_item = outlook.CreateItem(0)

        # Handle deletion
        if args.delete:
            sub = getattr(mail_item, "Subject", "")
            mail_item.Delete()
            clear_last_draft_id()
            del_res = {
                "status": "deleted",
                "message": f"Draft '{sub}' was successfully deleted from Outlook."
            }
            if args.format == "json":
                print(json.dumps(del_res, ensure_ascii=True))
            else:
                print(f"=== Draft Deleted ===\n{del_res['message']}")
            return

        # Update recipients if provided
        if args.to:
            mail_item.To = args.to.replace(";", ",")
        if args.cc:
            mail_item.CC = args.cc.replace(";", ",")
        if args.bcc:
            mail_item.BCC = args.bcc.replace(";", ",")
        if args.subject:
            mail_item.Subject = args.subject

        # Update body
        new_body = read_body_content(args.body)
        if new_body:
            if args.html:
                mail_item.HTMLBody = new_body
            else:
                mail_item.Body = new_body

        if args.append_body:
            app_text = read_body_content(args.append_body)
            if args.html or getattr(mail_item, "BodyFormat", 1) == 2:
                # HTML append
                mail_item.HTMLBody = (getattr(mail_item, "HTMLBody", "") or "") + f"<br>{app_text.replace(chr(13)+chr(10), '<br>')}"
            else:
                mail_item.Body = (getattr(mail_item, "Body", "") or "") + f"\r\n\r\n{app_text}"

        # Attachments handling
        for att_path in args.attachment:
            clean_path = str(Path(att_path).resolve())
            if not os.path.exists(clean_path):
                raise FileNotFoundError(f"Attachment file not found: {clean_path}")
            mail_item.Attachments.Add(clean_path)

        for img_path in args.inline_image:
            clean_path = str(Path(img_path).resolve())
            if not os.path.exists(clean_path):
                raise FileNotFoundError(f"Inline image file not found: {clean_path}")
            att = mail_item.Attachments.Add(clean_path)
            cid_name = Path(clean_path).name
            try:
                # 0x3712001F is PR_ATTACH_CONTENT_ID
                att.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid_name)
            except Exception:
                pass

        if args.remove_attachment:
            target_name = args.remove_attachment.lower()
            if hasattr(mail_item, "Attachments"):
                for i in range(mail_item.Attachments.Count, 0, -1):
                    att = mail_item.Attachments.Item(i)
                    if att.FileName.lower() == target_name:
                        att.Delete()

        # Handle Action: Send vs Save Draft
        if args.send:
            # Perform sending
            current_to = getattr(mail_item, "To", "")
            current_cc = getattr(mail_item, "CC", "")
            current_sub = getattr(mail_item, "Subject", "")
            mail_item.Send()
            clear_last_draft_id()
            status = "sent"
            entry_id = ""  # Once sent, it moves to Sent Items and entry ID changes
            action_desc = f"Email sent successfully to '{current_to}'."
        else:
            # Save draft
            mail_item.Save()
            status = "draft_updated" if is_update else "draft_created"
            entry_id = mail_item.EntryID
            save_last_draft_id(entry_id)
            action_desc = f"Draft saved successfully (EntryID: {entry_id})."

        # Prepare summary
        attachments_list = []
        if status != "sent" and hasattr(mail_item, "Attachments"):
            for i in range(1, mail_item.Attachments.Count + 1):
                attachments_list.append(mail_item.Attachments.Item(i).FileName)

        body_preview = ""
        if status != "sent":
            raw_body = getattr(mail_item, "Body", "") or ""
            body_preview = " ".join(raw_body.split())[:200]

        result_payload = {
            "status": status,
            "entry_id": entry_id,
            "message": action_desc,
            "to": getattr(mail_item, "To", "") if status != "sent" else current_to,
            "cc": getattr(mail_item, "CC", "") if status != "sent" else current_cc,
            "subject": getattr(mail_item, "Subject", "") if status != "sent" else current_sub,
            "attachments": attachments_list,
            "body_preview": body_preview,
        }

        if args.format == "json":
            print(json.dumps(result_payload, indent=2, ensure_ascii=False))
        else:
            print(f"=== Status: {status.upper()} ===")
            print(action_desc)
            if entry_id:
                print(f"EntryID:     {entry_id}")
            print(f"To:          {result_payload['to']}")
            if result_payload['cc']:
                print(f"CC:          {result_payload['cc']}")
            print(f"Subject:     {result_payload['subject']}")
            if attachments_list:
                print(f"Attachments: {', '.join(attachments_list)}")
            if body_preview:
                print(f"Preview:     {body_preview}...")

    except Exception as e:
        err_payload = {"status": "error", "message": str(e)}
        if args.format == "json":
            print(json.dumps(err_payload, ensure_ascii=False))
        else:
            print(f"Error composing/sending email: {e}", file=sys.stderr)
        sys.exit(1)
    finally:
        del mail_item
        del namespace
        del outlook
        gc.collect()
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception:
            pass

if __name__ == "__main__":
    main()
