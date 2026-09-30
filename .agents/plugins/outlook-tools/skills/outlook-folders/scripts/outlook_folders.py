# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "pywin32>=306",
#     "python-dateutil>=2.8.2",
# ]
# ///

"""
outlook_folders.py - Lists Outlook accounts and folder hierarchy with item & unread counts.
"""

import sys
import argparse
import json
import gc

def parse_args():
    parser = argparse.ArgumentParser(description="List Outlook mailboxes and folder structure.")
    parser.add_argument("--account", type=str, default="", help="Filter by specific mailbox/account name.")
    parser.add_argument("--unread-only", action="store_true", help="Show only folders with unread items.")
    parser.add_argument("--depth", type=int, default=4, help="Maximum folder nesting depth (default: 4).")
    parser.add_argument("--format", choices=["text", "json"], default="text", help="Output format.")
    return parser.parse_args()

def get_outlook_namespace():
    import win32com.client
    import pythoncom

    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    return outlook, namespace

ITEM_TYPE_NAMES = {
    0: "Mail",
    1: "Appointment/Calendar",
    2: "Contact",
    3: "Task",
    4: "Journal",
    5: "Note",
    6: "Post",
}

def inspect_folder(folder, current_depth, max_depth, unread_only):
    try:
        name = folder.Name
        folder_path = folder.FolderPath
        entry_id = folder.EntryID
        total_items = folder.Items.Count
        unread_items = folder.UnReadItemCount
        item_type = ITEM_TYPE_NAMES.get(folder.DefaultItemType, "Other")
    except Exception as e:
        return None

    children = []
    if current_depth < max_depth:
        try:
            for sub in folder.Folders:
                child_data = inspect_folder(sub, current_depth + 1, max_depth, unread_only)
                if child_data:
                    children.append(child_data)
        except Exception:
            pass

    # If unread_only is True, include this folder if it has unread items or any of its children have unread
    has_unread = unread_items > 0 or any(c["has_unread"] for c in children)
    if unread_only and not has_unread:
        return None

    return {
        "name": name,
        "path": folder_path,
        "entry_id": entry_id,
        "total_items": total_items,
        "unread_items": unread_items,
        "type": item_type,
        "has_unread": has_unread,
        "folders": children,
    }

def print_text_tree(folder_data, indent=0):
    prefix = "  " * indent
    unread_str = f" [UNREAD: {folder_data['unread_items']}]" if folder_data['unread_items'] > 0 else ""
    type_str = f" ({folder_data['type']})" if folder_data.get('type') else ""
    print(f"{prefix}- {folder_data['name']}{type_str}: {folder_data['total_items']} items{unread_str}")
    for child in folder_data.get("folders", []):
        print_text_tree(child, indent + 1)

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
        import pythoncom
        outlook, namespace = get_outlook_namespace()

        accounts_data = []
        for store in namespace.Folders:
            store_name = store.Name
            if args.account and args.account.lower() not in store_name.lower():
                continue

            store_info = {
                "account": store_name,
                "folders": []
            }

            try:
                for sub in store.Folders:
                    folder_info = inspect_folder(sub, current_depth=1, max_depth=args.depth, unread_only=args.unread_only)
                    if folder_info:
                        store_info["folders"].append(folder_info)
            except Exception as e:
                store_info["error"] = str(e)

            accounts_data.append(store_info)

        if args.format == "json":
            print(json.dumps({"status": "success", "accounts": accounts_data}, indent=2, ensure_ascii=False))
        else:
            if not accounts_data:
                print("No accounts or folders found matching the criteria.")
                return

            print("=== Outlook Mailboxes & Folders ===")
            for acc in accounts_data:
                print(f"\nAccount / Store: {acc['account']}")
                for fld in acc.get("folders", []):
                    print_text_tree(fld, indent=1)

    except Exception as e:
        err_payload = {"status": "error", "message": f"Error accessing Outlook COM: {str(e)}"}
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
