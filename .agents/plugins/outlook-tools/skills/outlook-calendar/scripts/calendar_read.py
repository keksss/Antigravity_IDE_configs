# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "pywin32>=306",
#     "python-dateutil>=2.8.2",
# ]
# ///

"""
calendar_read.py - Reads Outlook calendar appointments and meetings with recurrence and online link parsing.
"""

import sys
import re
import argparse
import json
import gc
from datetime import datetime, timedelta
from dateutil import parser as dt_parser

BUSY_STATUS_MAP = {
    0: "Free",
    1: "Tentative",
    2: "Busy",
    3: "OutOfOffice",
    4: "WorkingElsewhere"
}

def parse_args():
    parser = argparse.ArgumentParser(description="Read Outlook calendar events.")
    parser.add_argument("--period", choices=["today", "tomorrow", "this_week", "next_week", "month"], default="", help="Preset time window.")
    parser.add_argument("--start", type=str, default="", help="Start date/time (YYYY-MM-DD or YYYY-MM-DD HH:MM).")
    parser.add_argument("--end", type=str, default="", help="End date/time (YYYY-MM-DD or YYYY-MM-DD HH:MM).")
    parser.add_argument("--query", type=str, default="", help="Filter by text in subject, location, or body.")
    parser.add_argument("--limit", type=int, default=50, help="Maximum events to return (default: 50).")
    parser.add_argument("--format", choices=["text", "json"], default="text", help="Output format.")
    return parser.parse_args()

def calculate_time_window(period_arg: str, start_arg: str, end_arg: str):
    now = datetime.now()
    start_dt = None
    end_dt = None

    if period_arg == "today":
        start_dt = now.replace(hour=0, minute=0, second=0, microsecond=0)
        end_dt = start_dt + timedelta(days=1)
    elif period_arg == "tomorrow":
        start_dt = (now + timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
        end_dt = start_dt + timedelta(days=1)
    elif period_arg == "this_week":
        # Monday to Sunday of current week
        start_dt = (now - timedelta(days=now.weekday())).replace(hour=0, minute=0, second=0, microsecond=0)
        end_dt = start_dt + timedelta(days=7)
    elif period_arg == "next_week":
        start_dt = (now - timedelta(days=now.weekday()) + timedelta(days=7)).replace(hour=0, minute=0, second=0, microsecond=0)
        end_dt = start_dt + timedelta(days=7)
    elif period_arg == "month":
        start_dt = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        end_dt = (start_dt + timedelta(days=32)).replace(day=1)

    if start_arg:
        start_dt = dt_parser.parse(start_arg)
    if end_arg:
        end_dt = dt_parser.parse(end_arg)

    if not start_dt:
        start_dt = now.replace(hour=0, minute=0, second=0, microsecond=0)
    if not end_dt:
        end_dt = start_dt + timedelta(days=7)

    return start_dt, end_dt

def extract_meeting_links(text: str) -> str:
    if not text:
        return ""
    # Check for Teams meeting link
    teams_pattern = r"(https://teams\.microsoft\.com/l/meetup-join/[^\s\>\<\"]+)"
    zoom_pattern = r"(https://[a-zA-Z0-9\.\-]*zoom\.us/j/[^\s\>\<\"]+)"
    meet_pattern = r"(https://meet\.google\.com/[a-zA-Z0-9\-]+)"

    for pattern in [teams_pattern, zoom_pattern, meet_pattern]:
        match = re.search(pattern, text)
        if match:
            return match.group(1)
    return ""

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

        # olFolderCalendar = 9
        calendar_folder = namespace.GetDefaultFolder(9)
        items = calendar_folder.Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")

        start_dt, end_dt = calculate_time_window(args.period, args.start, args.end)

        # Restrict items between start and end
        # Jet format for date filtering: 'mm/dd/yyyy hh:mm AMPM'
        start_str = start_dt.strftime("%m/%d/%Y %H:%M")
        end_str = end_dt.strftime("%m/%d/%Y %H:%M")
        filter_str = f"[Start] < '{end_str}' AND [End] > '{start_str}'"

        try:
            restricted_items = items.Restrict(filter_str)
            restricted_items.Sort("[Start]")
        except Exception:
            restricted_items = items

        events = []
        query_lower = args.query.lower() if args.query else ""
        count = 0

        for item in restricted_items:
            if count >= args.limit:
                break

            try:
                # Class 26 is olAppointment
                if item.Class != 26:
                    continue
            except Exception as exc:
                print(f"Warning: calendar item class unavailable: {exc}", file=sys.stderr)
                continue

            try:
                item_start = dt_parser.parse(str(item.Start))
                item_end = dt_parser.parse(str(item.End))
            except Exception as exc:
                print(f"Warning: calendar item dates unavailable: {exc}", file=sys.stderr)
                continue

            # Double check boundary in Python
            # Ensure naive datetime comparison
            if item_start.tzinfo and not start_dt.tzinfo:
                item_start = item_start.replace(tzinfo=None)
            if item_end.tzinfo and not end_dt.tzinfo:
                item_end = item_end.replace(tzinfo=None)

            if item_end <= start_dt or item_start >= end_dt:
                continue

            subject = getattr(item, "Subject", "") or "(No Subject)"
            location = getattr(item, "Location", "") or ""
            body = getattr(item, "Body", "") or ""
            organizer = getattr(item, "Organizer", "") or ""
            duration = getattr(item, "Duration", 0)
            busy_status_val = getattr(item, "BusyStatus", 2)
            busy_status = BUSY_STATUS_MAP.get(busy_status_val, "Busy")
            is_recurring = getattr(item, "IsRecurring", False)
            entry_id = getattr(item, "EntryID", "")

            # Match query
            if query_lower:
                combined = f"{subject} {location} {body} {organizer}".lower()
                if query_lower not in combined:
                    continue

            # Detect online meeting link
            meeting_link = extract_meeting_links(location) or extract_meeting_links(body)
            is_teams = "teams.microsoft.com" in meeting_link or "teams" in location.lower()

            attendees = []
            if hasattr(item, "Recipients"):
                for r in range(1, item.Recipients.Count + 1):
                    rec = item.Recipients.Item(r)
                    attendees.append(rec.Name)

            events.append({
                "entry_id": entry_id,
                "subject": subject,
                "start": item_start.strftime("%Y-%m-%d %H:%M"),
                "end": item_end.strftime("%Y-%m-%d %H:%M"),
                "duration_minutes": duration,
                "location": location,
                "meeting_link": meeting_link,
                "is_teams": is_teams,
                "organizer": organizer,
                "attendees": attendees,
                "status": busy_status,
                "is_recurring": is_recurring,
                "body_preview": " ".join(body.split())[:150]
            })
            count += 1

        if args.format == "json":
            print(json.dumps({
                "status": "success",
                "period_start": start_dt.strftime("%Y-%m-%d %H:%M"),
                "period_end": end_dt.strftime("%Y-%m-%d %H:%M"),
                "count": len(events),
                "events": events
            }, indent=2, ensure_ascii=False))
        else:
            print(f"=== Calendar Schedule ({start_dt.strftime('%d.%m.%Y')} - {end_dt.strftime('%d.%m.%Y')}) ===")
            if not events:
                print("No events scheduled for this period.")
                return

            for ev in events:
                teams_mark = " [TEAMS MEETING]" if ev["is_teams"] else ""
                recur_mark = " [RECURRING]" if ev["is_recurring"] else ""
                print(f"\n• {ev['start']} - {ev['end']} ({ev['duration_minutes']} min): {ev['subject']}{teams_mark}{recur_mark}")
                if ev["location"]:
                    print(f"  Location: {ev['location']}")
                if ev["meeting_link"]:
                    print(f"  Link:     {ev['meeting_link']}")
                if ev["organizer"]:
                    print(f"  Organizer: {ev['organizer']}")
                if ev["attendees"]:
                    print(f"  Attendees: {', '.join(ev['attendees'])}")
                print(f"  Status:   {ev['status']} | ID: {ev['entry_id']}")

    except Exception as e:
        err_payload = {"status": "error", "message": str(e)}
        if args.format == "json":
            print(json.dumps(err_payload, ensure_ascii=False))
        else:
            print(f"Error accessing calendar: {e}", file=sys.stderr)
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
