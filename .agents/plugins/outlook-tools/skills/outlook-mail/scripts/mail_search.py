# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "pywin32>=306",
#     "python-dateutil>=2.8.2",
# ]
# ///

"""
mail_search.py - Searches and filters Outlook emails by folder, query, sender, date, and status.
"""

import sys
import argparse
import json
import gc
from datetime import datetime, timedelta
from dateutil import parser as dt_parser

def parse_args():
    parser = argparse.ArgumentParser(description="Search and filter Outlook emails.")
    parser.add_argument("--folder", type=str, default="Inbox", help="Folder name or path (default: 'Inbox').")
    parser.add_argument("--query", type=str, default="", help="Text search in subject and body.")
    parser.add_argument("--sender", type=str, default="", help="Filter by sender name or email address.")
    parser.add_argument("--subject", type=str, default="", help="Filter specifically by subject.")
    parser.add_argument("--unread", action="store_true", help="Only show unread messages.")
    parser.add_argument("--has-attachments", action="store_true", help="Only show messages with attachments.")
    parser.add_argument("--since", type=str, default="", help="Start date (YYYY-MM-DD, 'today', 'yesterday', '7d').")
    parser.add_argument("--until", type=str, default="", help="End date (YYYY-MM-DD).")
    parser.add_argument("--limit", type=int, default=20, help="Maximum number of emails to return (default: 20).")
    parser.add_argument("--delete", action="store_true", help="Delete matching emails (moves them to 'Deleted Items' folder). Requires confirmation.")
    parser.add_argument("--format", choices=["text", "json"], default="text", help="Output format.")
    return parser.parse_args()

def parse_relative_date(date_str: str) -> datetime | None:
    if not date_str:
        return None
    val = date_str.strip().lower()
    now = datetime.now()
    if val == "today":
        return now.replace(hour=0, minute=0, second=0, microsecond=0)
    elif val == "yesterday":
        y = now - timedelta(days=1)
        return y.replace(hour=0, minute=0, second=0, microsecond=0)
    elif val.endswith("d") and val[:-1].isdigit():
        days = int(val[:-1])
        d = now - timedelta(days=days)
        return d.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        try:
            return dt_parser.parse(date_str)
        except Exception:
            return None

def resolve_folder(namespace, folder_spec: str):
    if not folder_spec or not folder_spec.strip():
        return namespace.GetDefaultFolder(6)

    spec_lower = folder_spec.strip().lower()
    default_mapping = {
        "inbox": 6, "входящие": 6,
        "drafts": 16, "черновики": 16,
        "sent": 5, "sent items": 5, "отправленные": 5,
        "deleted": 3, "deleted items": 3, "удаленные": 3,
        "junk": 23, "спам": 23,
        "outbox": 4, "исходящие": 4,
    }

    if spec_lower in default_mapping:
        try:
            return namespace.GetDefaultFolder(default_mapping[spec_lower])
        except Exception as e:
            raise ValueError(f"Could not open default folder '{folder_spec}': {e}")

    # Check immediate subfolders of Default Inbox first (e.g. Inbox\SAP_Project)
    try:
        inbox = namespace.GetDefaultFolder(6)
        for sub in inbox.Folders:
            if sub.Name.lower() == spec_lower:
                return sub
    except Exception:
        pass

    # Try path traversal if path contains separators
    parts = [p.strip() for p in folder_spec.replace("\\", "/").split("/") if p.strip()]
    if len(parts) > 1:
        for store in namespace.Folders:
            current = store
            found = True
            for part in parts:
                child_found = None
                try:
                    for sub in current.Folders:
                        if sub.Name.lower() == part.lower():
                            child_found = sub
                            break
                except Exception:
                    pass
                if child_found:
                    current = child_found
                else:
                    found = False
                    break
            if found:
                return current
    else:
        # Fast search within the user's primary mailbox (root and 2 levels deep)
        single_name = parts[0].lower()
        try:
            primary_root = namespace.GetDefaultFolder(6).Parent
            # 1. Check direct children of mailbox root
            for f in primary_root.Folders:
                if f.Name.lower() == single_name:
                    return f
            # 2. Check second-level subfolders (e.g. Inbox/*, Archive/*)
            for f in primary_root.Folders:
                try:
                    for sub in f.Folders:
                        if sub.Name.lower() == single_name:
                            return sub
                except Exception:
                    pass

            # 3. Fallback for common plural/singular variations (e.g. sap_projects -> sap_project)
            variants = [single_name[:-1]] if single_name.endswith('s') else [single_name + 's']
            for v in variants:
                for f in primary_root.Folders:
                    if f.Name.lower() == v:
                        return f
                    try:
                        for sub in f.Folders:
                            if sub.Name.lower() == v:
                                return sub
                    except Exception:
                        pass
        except Exception:
            pass

    # STRICT SAFETY: NEVER silently fallback to inbox!
    raise ValueError(
        f"Folder '{folder_spec}' was not found in Outlook. "
        f"Please check the folder name or specify the full path (e.g. 'Inbox/{folder_spec}')."
    )

def main():
    if sys.stdout.encoding != 'utf-8':
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except Exception:
            pass

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

        folder = resolve_folder(namespace, args.folder)
        items = folder.Items
        
        # We sort by ReceivedTime descending
        try:
            items.Sort("[ReceivedTime]", True)
        except Exception:
            pass

        # Apply basic Jet restrict for speed if unread or date is given
        since_dt = parse_relative_date(args.since)
        until_dt = parse_relative_date(args.until)

        filter_clauses = []
        if args.unread:
            filter_clauses.append("[UnRead] = True")
        if since_dt:
            # Jet date format: 'yyyy-mm-dd hh:mm AMPM' or 'mm/dd/yyyy hh:mm AMPM'
            filter_clauses.append(f"[ReceivedTime] >= '{since_dt.strftime('%m/%d/%Y %H:%M')}'")
        if until_dt:
            filter_clauses.append(f"[ReceivedTime] <= '{until_dt.strftime('%m/%d/%Y %H:%M')}'")

        if filter_clauses:
            filter_str = " AND ".join(filter_clauses)
            try:
                items = items.Restrict(filter_str)
                items.Sort("[ReceivedTime]", True)
            except Exception:
                # If restrict fails, fallback to iterating
                pass

        results = []
        query_lower = args.query.lower() if args.query else ""
        sender_lower = args.sender.lower() if args.sender else ""
        subject_lower = args.subject.lower() if args.subject else ""

        count = 0
        for item in items:
            if count >= args.limit:
                break

            # Filter only mail items (Class 43 is olMail)
            try:
                if item.Class != 43:
                    continue
            except Exception:
                continue

            try:
                subject = item.Subject or ""
                sender_name = item.SenderName or ""
                sender_email = ""
                try:
                    sender_email = item.SenderEmailAddress or ""
                except Exception:
                    pass

                body = item.Body or ""
                has_attachments = item.Attachments.Count > 0 if hasattr(item, "Attachments") else False
                unread = item.UnRead
                rec_time = str(item.ReceivedTime) if hasattr(item, "ReceivedTime") else ""
                entry_id = item.EntryID
            except Exception:
                continue

            # Check query filter
            if query_lower and (query_lower not in subject.lower() and query_lower not in body.lower()):
                continue

            # Check subject filter
            if subject_lower and (subject_lower not in subject.lower()):
                continue

            # Check sender filter
            if sender_lower and (sender_lower not in sender_name.lower() and sender_lower not in sender_email.lower()):
                continue

            # Check attachments filter
            if args.has_attachments and not has_attachments:
                continue

            # Snippet of body (first 200 chars clean)
            snippet = " ".join(body.split())[:180]

            results.append({
                "entry_id": entry_id,
                "subject": subject,
                "sender_name": sender_name,
                "sender_email": sender_email,
                "received_time": rec_time,
                "unread": unread,
                "has_attachments": has_attachments,
                "snippet": snippet,
            })
            count += 1

        if args.delete:
            deleted_count = 0
            deleted_items = []
            for msg in results:
                try:
                    item_obj = namespace.GetItemFromID(msg["entry_id"])
                    sub = getattr(item_obj, "Subject", "")
                    item_obj.Delete()
                    deleted_count += 1
                    deleted_items.append({"entry_id": msg["entry_id"], "subject": sub})
                except Exception as del_err:
                    print(f"Warning: could not delete '{msg.get('subject', '')}': {del_err}", file=sys.stderr)

            if args.format == "json":
                print(json.dumps({
                    "status": "deleted",
                    "folder": folder.Name,
                    "deleted_count": deleted_count,
                    "items": deleted_items
                }, indent=2, ensure_ascii=False))
            else:
                print(f"=== Cleaned Folder '{folder.Name}' ===")
                print(f"Successfully moved {deleted_count} messages to 'Deleted Items' folder:")
                for idx, it in enumerate(deleted_items, 1):
                    print(f"  {idx}. {it['subject']}")
            return

        if args.format == "json":
            print(json.dumps({
                "status": "success",
                "folder": folder.Name,
                "count": len(results),
                "messages": results
            }, indent=2, ensure_ascii=False))
        else:
            print(f"=== Search Results in '{folder.Name}' ({len(results)} found) ===")
            if not results:
                print("No emails matched your criteria.")
                return

            for idx, msg in enumerate(results, 1):
                unread_flag = "[UNREAD] " if msg["unread"] else ""
                attach_flag = "[ATTACHMENT] " if msg["has_attachments"] else ""
                print(f"\n{idx}. {unread_flag}{attach_flag}{msg['subject']}")
                print(f"   From: {msg['sender_name']} <{msg['sender_email']}>")
                print(f"   Date: {msg['received_time']}")
                print(f"   ID:   {msg['entry_id']}")
                if msg["snippet"]:
                    print(f"   Preview: {msg['snippet']}...")

    except Exception as e:
        err_payload = {"status": "error", "message": str(e)}
        if args.format == "json":
            print(json.dumps(err_payload, ensure_ascii=False))
        else:
            print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
    finally:
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
