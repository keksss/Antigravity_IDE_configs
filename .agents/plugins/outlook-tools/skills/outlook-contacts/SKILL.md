---
name: outlook-contacts
description: Searches local Outlook contacts and corporate Exchange Global Address List (GAL) by name, email, department, or phone. Use when the user asks to "find contact", "look up email", "search address book", "get phone number", or "find colleague in GAL".
---

# Outlook Contacts Skill

This skill allows searching through personal Outlook contacts and the corporate Exchange Global Address List (GAL) to retrieve email addresses, phone numbers, job titles, and departments.

## Execution

All scripts must be executed via `uv run` with isolated dependencies:

```bash
uv run scripts/contacts_search.py [options]
```

## CLI Parameters

- `--query <text>`: (Required) Search term (person name, surname, email snippet, department, or company).
- `--source <all|local|gal>`: Where to look:
  - `local`: Search personal Contacts folder only.
  - `gal`: Search corporate Exchange Global Address List only.
  - `all`: Search both (default).
- `--limit <int>`: Max results to return (default: `10`).
- `--format <text|json>`: Output format.

## Examples

### 1. Search for a Colleague Across All Sources
```bash
uv run scripts/contacts_search.py --query "Иванов" --format text
```

### 2. Search GAL Exclusively for Corporate Directory Lookup
```bash
uv run scripts/contacts_search.py --query "dev-lead@company.com" --source gal --format json
```

### 3. Search Personal Contacts Only
```bash
uv run scripts/contacts_search.py --query "Alex" --source local --format text
```

## Output Structure (JSON)

```json
{
  "status": "success",
  "query": "Иванов",
  "count": 1,
  "contacts": [
    {
      "name": "Иванов Иван Иванович",
      "email": "ivanov@company.com",
      "secondary_email": "",
      "company": "Company LLC",
      "job_title": "Lead Software Engineer",
      "department": "Platform Core",
      "business_phone": "+7 (495) 000-00-00",
      "mobile_phone": "+7 (999) 000-00-00",
      "source": "Exchange GAL",
      "entry_id": "00000000..."
    }
  ]
}
```
