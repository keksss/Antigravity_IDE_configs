# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "pywin32>=306",
#     "python-dateutil>=2.8.2",
# ]
# ///

"""
calendar_manage.py - Creates, updates, sends invitations, and cancels Outlook calendar appointments and Teams meetings.
"""

import sys
import argparse
import json
import gc
from pathlib import Path
from datetime import datetime, timedelta
from dateutil import parser as dt_parser

def get_state_file() -> Path:
    # plugins/outlook-tools/.state/last_event_id.txt
    base = Path(__file__).resolve().parent.parent.parent.parent
    state_dir = base / ".state"
    state_dir.mkdir(parents=True, exist_ok=True)
    return state_dir / "last_event_id.txt"

def save_last_event_id(entry_id: str):
    try:
        get_state_file().write_text(entry_id.strip(), encoding="utf-8")
    except Exception:
        pass

def get_last_event_id() -> str:
    try:
        f = get_state_file()
        if f.exists():
            return f.read_text(encoding="utf-8").strip()
    except Exception:
        pass
    return ""

def clear_last_event_id():
    try:
        f = get_state_file()
        if f.exists():
            f.unlink()
    except Exception:
        pass

def read_body_content(body_arg: str) -> str:
    if not body_arg:
        return ""
    if body_arg == "-":
        try:
            if hasattr(sys.stdin, "reconfigure"):
                sys.stdin.reconfigure(encoding="utf-8")
            content = sys.stdin.read()
        except Exception:
            content = ""
    else:
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

    normalized = content.replace(r"\r\n", "\n").replace(r"\n", "\n").replace(r"\t", "\t")
    return normalized.replace("\r\n", "\n").replace("\n", "\r\n")

def parse_args():
    parser = argparse.ArgumentParser(description="Create, edit, or cancel Outlook calendar appointments.")
    parser.add_argument("--id", type=str, default="", help="EntryID of existing appointment to edit, send, or cancel.")
    parser.add_argument("--last-event", action="store_true", default=False, help="Use the most recently created or modified event.")
    parser.add_argument("--subject", type=str, default="", help="Meeting subject / title.")
    parser.add_argument("--start", type=str, default="", help="Start time (YYYY-MM-DD HH:MM).")
    parser.add_argument("--duration", type=int, default=30, help="Duration in minutes (default: 30).")
    parser.add_argument("--end", type=str, default="", help="End time (alternative to --duration).")
    parser.add_argument("--attendees", type=str, default="", help="Comma-separated emails of required attendees.")
    parser.add_argument("--optional-attendees", type=str, default="", help="Comma-separated emails of optional attendees.")
    parser.add_argument("--location", type=str, default="", help="Meeting room or location.")
    parser.add_argument("--body", type=str, default="", help="Meeting agenda / description.")
    parser.add_argument("--teams", action="store_true", help="Create as Microsoft Teams online meeting.")
    parser.add_argument("--reminder", type=int, default=15, help="Reminder in minutes before start (default: 15).")
    parser.add_argument("--save-draft", action="store_true", default=False, help="Save to calendar without sending invitations (default).")
    parser.add_argument("--send", action="store_true", default=False, help="Send invitations to attendees (requires user confirmation!).")
    parser.add_argument("--cancel", action="store_true", default=False, help="Cancel/delete the meeting (requires user confirmation!).")
    parser.add_argument("--delete", action="store_true", default=False, help="Alias for --cancel.")
    parser.add_argument("--format", choices=["text", "json"], default="text", help="Output format.")
    return parser.parse_args()

def enable_teams_meeting(appt):
    """
    Attempts to trigger the Microsoft Teams meeting add-in via Outlook Inspector Ribbon.
    """
    try:
        inspector = appt.GetInspector
        # Execute the built-in MSO command for the Teams meeting add-in
        inspector.CommandBars.ExecuteMso("TeamsMeetingAddin")
        return True
    except Exception:
        # If command bar fails (e.g. headless without active UI), set location and note
        if not appt.Location:
            appt.Location = "Microsoft Teams Meeting"
        return False

def main():
    if sys.stdout.encoding != 'utf-8':
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except Exception:
            pass

    args = parse_args()
    outlook = None
    namespace = None
    appt = None

    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()

        try:
            outlook = win32com.client.GetActiveObject("Outlook.Application")
        except Exception:
            outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        if args.last_event and not args.id:
            args.id = get_last_event_id()
            if not args.id:
                raise ValueError("No recently saved calendar event found in local state. Please specify --id explicitly.")

        is_update = bool(args.id)
        if is_update:
            try:
                appt = namespace.GetItemFromID(args.id)
            except Exception as e:
                raise ValueError(f"Could not find appointment with EntryID '{args.id}': {e}")
        else:
            if not args.subject:
                raise ValueError("Subject is required when creating a new calendar event.")
            if not args.start:
                raise ValueError("Start date/time (--start) is required when creating a new event.")
            # 1 = olAppointmentItem
            appt = outlook.CreateItem(1)

        # Handle cancellation / deletion
        if args.cancel or args.delete:
            sub = getattr(appt, "Subject", "")
            if getattr(appt, "MeetingStatus", 0) == 1:
                # 5 = olMeetingCanceled
                try:
                    appt.MeetingStatus = 5
                    if args.send:
                        appt.Send()
                except Exception:
                    pass
            appt.Delete()
            clear_last_event_id()
            res = {
                "status": "cancelled",
                "message": f"Appointment '{sub}' was successfully removed from calendar."
            }
            if args.format == "json":
                print(json.dumps(res, ensure_ascii=False))
            else:
                print(f"=== Meeting Cancelled ===\n{res['message']}")
            return

        # Update properties
        if args.subject:
            appt.Subject = args.subject

        if args.start:
            start_dt = dt_parser.parse(args.start)
            appt.Start = start_dt.strftime("%Y-%m-%d %H:%M:%S")

        if args.end:
            end_dt = dt_parser.parse(args.end)
            appt.End = end_dt.strftime("%Y-%m-%d %H:%M:%S")
        elif args.duration:
            appt.Duration = args.duration

        if args.location:
            appt.Location = args.location

        if args.body:
            appt.Body = read_body_content(args.body)

        if args.reminder >= 0:
            appt.ReminderSet = True
            appt.ReminderMinutesBeforeStart = args.reminder

        # Manage attendees
        has_attendees = bool(args.attendees or args.optional_attendees)
        if has_attendees:
            # 1 = olMeeting
            appt.MeetingStatus = 1
            if args.attendees:
                for att in args.attendees.split(","):
                    clean = att.strip()
                    if clean:
                        rec = appt.Recipients.Add(clean)
                        rec.Type = 1  # 1 = olRequired
            if args.optional_attendees:
                for att in args.optional_attendees.split(","):
                    clean = att.strip()
                    if clean:
                        rec = appt.Recipients.Add(clean)
                        rec.Type = 2  # 2 = olOptional
            appt.Recipients.ResolveAll()

        # Handle Teams meeting
        teams_activated = False
        if args.teams:
            teams_activated = enable_teams_meeting(appt)

        # Save or Send
        if args.send:
            if getattr(appt, "MeetingStatus", 0) == 1:
                appt.Send()
                action_status = "invitations_sent"
                action_msg = "Meeting saved and invitations sent to all attendees."
            else:
                appt.Save()
                action_status = "saved"
                action_msg = "Appointment saved (no attendees specified to send to)."
            clear_last_event_id()
            entry_id = getattr(appt, "EntryID", "")
        else:
            appt.Save()
            action_status = "updated" if is_update else "created"
            entry_id = appt.EntryID
            save_last_event_id(entry_id)
            action_msg = f"Meeting saved to calendar (EntryID: {entry_id})."

        result = {
            "status": action_status,
            "entry_id": entry_id,
            "message": action_msg,
            "subject": getattr(appt, "Subject", ""),
            "start": str(getattr(appt, "Start", "")),
            "end": str(getattr(appt, "End", "")),
            "duration": getattr(appt, "Duration", 0),
            "location": getattr(appt, "Location", ""),
            "is_teams": args.teams or "teams" in getattr(appt, "Location", "").lower(),
            "teams_addin_triggered": teams_activated,
            "attendees": [r.Name for r in appt.Recipients] if hasattr(appt, "Recipients") else []
        }

        if args.format == "json":
            print(json.dumps(result, indent=2, ensure_ascii=False))
        else:
            print(f"=== Calendar Event: {action_status.upper()} ===")
            print(action_msg)
            print(f"Subject:   {result['subject']}")
            print(f"Time:      {result['start']} -> {result['end']} ({result['duration']} min)")
            print(f"Location:  {result['location']}")
            if result['is_teams']:
                print(f"Online:    Microsoft Teams Meeting")
            if result['attendees']:
                print(f"Attendees: {', '.join(result['attendees'])}")
            if entry_id:
                print(f"EntryID:   {entry_id}")

    except Exception as e:
        err_payload = {"status": "error", "message": str(e)}
        if args.format == "json":
            print(json.dumps(err_payload, ensure_ascii=False))
        else:
            print(f"Error managing calendar event: {e}", file=sys.stderr)
        sys.exit(1)
    finally:
        del appt
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
