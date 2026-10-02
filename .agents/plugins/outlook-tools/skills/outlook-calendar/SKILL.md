---
name: outlook-calendar
description: Reads Outlook calendar events, inspects schedules, creates appointments, manages Microsoft Teams meetings, and updates or cancels calendar events. Use when the user asks to "check calendar", "see schedule", "schedule meeting", "create teams call", or "cancel appointment".
---

# Outlook Calendar Skill

This skill allows viewing schedule, finding free time slots, creating calendar events, scheduling Microsoft Teams meetings, and safely updating or cancelling appointments.

## Safety & Human Confirmation Rules

> [!IMPORTANT]
> **CRITICAL RULE**: The agent must **NEVER** run `calendar_manage.py --send` (sending invitations to participants) or `calendar_manage.py --cancel` / `--delete` without prior explicit approval from the user in chat.
> 1. By default, create meetings with `--save-draft` (or omit `--send`) to place them on the calendar without disturbing attendees.
> 2. Present event details in chat: Subject, Time, Duration, Location/Teams, Attendees.
> 3. Only after the user confirms sending or cancelling, run the respective command with `--send` or `--cancel`.

---

## 1. Reading Calendar & Schedule (`scripts/calendar_read.py`)

Inspect upcoming appointments, meetings, recurrence, and extract meeting links (Teams, Zoom, Google Meet).

```bash
uv run scripts/calendar_read.py [options]
```

### CLI Options
- `--period <today|tomorrow|this_week|next_week|month>`: Preset time window.
- `--start <datetime>`: Start date/time (`YYYY-MM-DD` or `YYYY-MM-DD HH:MM`).
- `--end <datetime>`: End date/time.
- `--query <text>`: Filter by keyword in title, location, or body.
- `--limit <int>`: Max events (default: `50`).
- `--format <text|json>`: Output format.

### Examples

#### View Today's Schedule
```bash
uv run scripts/calendar_read.py --period today --format text
```

#### View This Week's Schedule with Online Meeting Links
```bash
uv run scripts/calendar_read.py --period this_week --format json
```

#### Query Specific Date Range
```bash
uv run scripts/calendar_read.py --start "2026-10-05 09:00" --end "2026-10-09 18:00" --format text
```

---

## 2. Creating and Managing Meetings (`scripts/calendar_manage.py`)

Create appointments, schedule online Teams meetings, update existing events by EntryID, or cancel meetings.
By default, running without `--apply` prints a **dry-run preview** without touching Outlook. Pass `--apply` to commit changes.

```bash
uv run scripts/calendar_manage.py [options]
```

### CLI Options
- `--id <entry_id>`: EntryID of an existing appointment to edit, send, or cancel.
- `--last-event`: (Recommended) Target the most recently created or modified event without entering the long EntryID.
- `--subject <text>`: (Required for new) Meeting title.
- `--start <datetime>`: (Required for new) Start date/time (`YYYY-MM-DD HH:MM`).
- `--duration <minutes>`: Meeting length in minutes (default: `30`).
- `--end <datetime>`: Alternative to duration.
- `--attendees <emails>`: Comma-separated required attendee email addresses.
- `--optional-attendees <emails>`: Optional attendee email addresses.
- `--location <text>`: Physical room or note.
- `--teams`: **Flag to create as a Microsoft Teams online meeting**.
- `--body <text>`: Agenda or meeting description.
- `--reminder <minutes>`: Reminder alert before start (default: `15`).
- `--save-draft`: Save to calendar without broadcasting invitations (default).
- `--send`: Send invitations to all attendees (**CONFIRMATION REQUIRED**).
- `--cancel` / `--delete`: Cancel/remove the meeting (**CONFIRMATION REQUIRED**).
- `--confirm-cancel-id <entry_id>`: Exact existing EntryID to confirm `--cancel`/`--delete`.
- `--apply`: Apply planned changes to Outlook; otherwise runs preview (dry-run).
- `--format <text|json>`

### Common Workflows

#### Workflow A: Schedule a Personal Appointment
```bash
# Preview appointment (dry-run)
uv run scripts/calendar_manage.py \
  --subject "Подготовка квартального отчета" \
  --start "2026-10-02 14:00" \
  --duration 60 \
  --format json

# Create appointment in Outlook
uv run scripts/calendar_manage.py \
  --subject "Подготовка квартального отчета" \
  --start "2026-10-02 14:00" \
  --duration 60 \
  --apply \
  --format json
```

#### Workflow B: Schedule a Microsoft Teams Meeting with Attendees (Draft)
```bash
uv run scripts/calendar_manage.py \
  --subject "Синхронизация по релизу" \
  --start "2026-10-03 11:00" \
  --duration 30 \
  --attendees "developer@example.com, tester@example.com" \
  --teams \
  --body "Обсуждение блокеров и плана тестирования." \
  --apply \
  --format json
```

#### Workflow C: Send Invitations (Using `--last-event`)
```bash
uv run scripts/calendar_manage.py --last-event --send --apply
```

#### Workflow D: Cancel a Meeting (Requires Exact ID and Confirmation)
```bash
uv run scripts/calendar_manage.py --id "<entry_id>" --confirm-cancel-id "<entry_id>" --cancel --apply
```
