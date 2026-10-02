# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "pywin32>=306",
#     "python-dateutil>=2.8.2",
# ]
# ///

"""
contacts_search.py - Searches local Outlook contacts and corporate Exchange Global Address List (GAL).
"""

import sys
import argparse
import json
import gc

def parse_args():
    parser = argparse.ArgumentParser(description="Search Outlook local contacts and Exchange GAL.")
    parser.add_argument("--query", type=str, required=True, help="Search query (name, surname, email, or department).")
    parser.add_argument("--source", choices=["all", "local", "gal"], default="all", help="Contact source (default: 'all').")
    parser.add_argument("--limit", type=int, default=10, help="Maximum results to return (default: 10).")
    parser.add_argument("--format", choices=["text", "json"], default="text", help="Output format.")
    return parser.parse_args()

CYR_TO_LAT_BASE = {
    'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e', 'ё': 'e',
    'ж': 'zh', 'з': 'z', 'и': 'i', 'к': 'k', 'л': 'l', 'м': 'm',
    'н': 'n', 'о': 'o', 'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u',
    'ф': 'f', 'х': 'kh', 'ц': 'ts', 'ч': 'ch', 'ш': 'sh', 'щ': 'shch',
    'ъ': '', 'ы': 'y', 'ь': '', 'э': 'e', 'ю': 'yu', 'я': 'ya'
}

def transliterate(text: str, j_char: str = 'i') -> str:
    res = []
    for ch in text:
        low = ch.lower()
        if low == 'й':
            mapped = j_char
            res.append(mapped.capitalize() if ch.isupper() else mapped)
        elif low in CYR_TO_LAT_BASE:
            trans = CYR_TO_LAT_BASE[low]
            res.append(trans.capitalize() if ch.isupper() else trans)
        else:
            res.append(ch)
    return "".join(res)

def get_query_variants(query: str) -> list[str]:
    """Generates standard transliteration variants (ISO Slavic 'i' vs ICAO Passport 'y') without hardcoded names."""
    variants = [query]
    has_cyrillic = any('\u0400' <= ch <= '\u04FF' for ch in query)
    if has_cyrillic:
        # Standard Slavic / Corporate directory convention (й -> i, e.g. Andrei, Sergei, Nikolai)
        v_i = transliterate(query, j_char='i')
        if v_i not in variants:
            variants.append(v_i)
        # Standard Passport / ICAO convention (й -> y, e.g. Andrey, Sergey, Nikolay)
        v_y = transliterate(query, j_char='y')
        if v_y not in variants:
            variants.append(v_y)
    return variants

def resolve_smtp_address(namespace, raw_address: str, fallback_name: str = "") -> str:
    if not raw_address and not fallback_name:
        return ""
    if raw_address and "@" in raw_address and not raw_address.startswith("/o="):
        return raw_address

    lookup_targets = [t for t in [raw_address, fallback_name] if t]
    for target in lookup_targets:
        try:
            rec = namespace.CreateRecipient(target)
            if rec.Resolve():
                ae = rec.AddressEntry
                ex_user = None
                try:
                    ex_user = ae.GetExchangeUser()
                except Exception as exc:
                    print(f"Warning: Exchange user lookup failed: {exc}", file=sys.stderr)
                if ex_user:
                    smtp = getattr(ex_user, "PrimarySmtpAddress", "")
                    if smtp and "@" in smtp:
                        return smtp
                try:
                    smtp = ae.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001F")
                    if smtp and "@" in smtp:
                        return smtp
                except Exception as exc:
                    print(f"Warning: SMTP property lookup failed: {exc}", file=sys.stderr)
        except Exception as exc:
            print(f"Warning: recipient lookup failed: {exc}", file=sys.stderr)
    return raw_address

def search_local_contacts(namespace, query: str, limit: int):
    results = []
    seen = set()
    try:
        # olFolderContacts = 10
        contacts_folder = namespace.GetDefaultFolder(10)
        folders_to_scan = [contacts_folder]
        try:
            for sub in contacts_folder.Folders:
                folders_to_scan.append(sub)
        except Exception as exc:
            print(f"Warning: contact subfolders unavailable: {exc}", file=sys.stderr)

        query_variants = [v.lower() for v in get_query_variants(query)]

        for folder in folders_to_scan:
            folder_name = folder.Name
            try:
                items = folder.Items
                for item in items:
                    if len(results) >= limit:
                        break
                    try:
                        # Class 40 is olContact
                        if getattr(item, "Class", 0) != 40:
                            continue

                        full_name = getattr(item, "FullName", "") or ""
                        email1 = getattr(item, "Email1Address", "") or ""
                        email2 = getattr(item, "Email2Address", "") or ""
                        company = getattr(item, "CompanyName", "") or ""
                        job_title = getattr(item, "JobTitle", "") or ""
                        dept = getattr(item, "Department", "") or ""
                        biz_phone = getattr(item, "BusinessTelephoneNumber", "") or ""
                        mob_phone = getattr(item, "MobileTelephoneNumber", "") or ""
                        entry_id = getattr(item, "EntryID", "")

                        haystack = f"{full_name} {email1} {email2} {company} {job_title} {dept}".lower()
                        if any(v in haystack for v in query_variants):
                            smtp_email = resolve_smtp_address(namespace, email1 or email2, full_name)
                            key = (full_name.lower(), smtp_email.lower())
                            if key not in seen:
                                seen.add(key)
                                results.append({
                                    "name": full_name,
                                    "email": smtp_email,
                                    "secondary_email": resolve_smtp_address(namespace, email2, full_name) if email1 and email2 != email1 else "",
                                    "company": company,
                                    "job_title": job_title,
                                    "department": dept,
                                    "business_phone": biz_phone,
                                    "mobile_phone": mob_phone,
                                    "source": f"Local Contacts ({folder_name})",
                                    "entry_id": entry_id
                                })
                    except Exception as exc:
                        print(f"Warning: contact item skipped: {exc}", file=sys.stderr)
                        continue
            except Exception as exc:
                print(f"Warning: contact folder skipped: {exc}", file=sys.stderr)
                continue
    except Exception as exc:
        print(f"Warning: contact search unavailable: {exc}", file=sys.stderr)
    return results

def search_gal(namespace, query: str, limit: int):
    """
    Searches the Exchange Global Address List (GAL) efficiently.
    Uses Recipient resolution and AddressLists lookup.
    """
    results = []
    seen_emails = set()

    # Strategy 1: Resolve directly as a recipient using standard transliteration variants
    lookup_candidates = get_query_variants(query)

    for cand in lookup_candidates:
        if len(results) >= limit:
            break
        try:
            rec = namespace.CreateRecipient(cand)
            if rec.Resolve():
                ae = rec.AddressEntry
                ex_user = None
                try:
                    ex_user = ae.GetExchangeUser()
                except Exception as exc:
                    print(f"Warning: Exchange user lookup failed: {exc}", file=sys.stderr)

                if ex_user:
                    name = getattr(ex_user, "Name", "") or ae.Name
                    email = getattr(ex_user, "PrimarySmtpAddress", "") or ae.Address
                    if email and email.lower() not in seen_emails:
                        seen_emails.add(email.lower())
                        results.append({
                            "name": name,
                            "email": email,
                            "secondary_email": "",
                            "company": getattr(ex_user, "CompanyName", "") or "",
                            "job_title": getattr(ex_user, "JobTitle", "") or "",
                            "department": getattr(ex_user, "Department", "") or "",
                            "business_phone": getattr(ex_user, "BusinessTelephoneNumber", "") or "",
                            "mobile_phone": getattr(ex_user, "MobileTelephoneNumber", "") or "",
                            "source": "Exchange GAL",
                            "entry_id": ae.ID
                        })
        except Exception as exc:
            print(f"Warning: GAL recipient lookup failed: {exc}", file=sys.stderr)

    # Strategy 2: If we still need more results or query is partial, search AddressLists
    if len(results) < limit:
        q_lower = query.lower()
        try:
            for al in namespace.AddressLists:
                # Look for Global Address List or All Users
                al_name_lower = al.Name.lower()
                if "global" in al_name_lower or "глобальн" in al_name_lower or "all users" in al_name_lower:
                    entries = al.AddressEntries
                    # Note: we cap iterations to 150 entries to avoid hanging on large organizations
                    count_scanned = 0
                    for entry in entries:
                        if len(results) >= limit or count_scanned > 150:
                            break
                        count_scanned += 1
                        try:
                            entry_name = entry.Name
                            if q_lower in entry_name.lower():
                                ex_user = None
                                try:
                                    ex_user = entry.GetExchangeUser()
                                except Exception as exc:
                                    print(f"Warning: GAL Exchange user lookup failed: {exc}", file=sys.stderr)

                                if ex_user:
                                    email = getattr(ex_user, "PrimarySmtpAddress", "") or entry.Address
                                    if email and email.lower() not in seen_emails:
                                        seen_emails.add(email.lower())
                                        results.append({
                                            "name": ex_user.Name,
                                            "email": email,
                                            "secondary_email": "",
                                            "company": getattr(ex_user, "CompanyName", "") or "",
                                            "job_title": getattr(ex_user, "JobTitle", "") or "",
                                            "department": getattr(ex_user, "Department", "") or "",
                                            "business_phone": getattr(ex_user, "BusinessTelephoneNumber", "") or "",
                                            "mobile_phone": getattr(ex_user, "MobileTelephoneNumber", "") or "",
                                            "source": "Exchange GAL",
                                            "entry_id": entry.ID
                                        })
                        except Exception as exc:
                            print(f"Warning: GAL entry skipped: {exc}", file=sys.stderr)
                            continue
        except Exception as exc:
            print(f"Warning: GAL address list unavailable: {exc}", file=sys.stderr)

    return results

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

        all_contacts = []
        q = args.query.strip()

        if args.source in ["all", "local"]:
            local_res = search_local_contacts(namespace, q, args.limit)
            all_contacts.extend(local_res)

        if args.source in ["all", "gal"] and len(all_contacts) < args.limit:
            remaining_limit = args.limit - len(all_contacts)
            gal_res = search_gal(namespace, q, remaining_limit)
            all_contacts.extend(gal_res)

        if args.format == "json":
            print(json.dumps({
                "status": "success",
                "query": q,
                "count": len(all_contacts),
                "contacts": all_contacts
            }, indent=2, ensure_ascii=False))
        else:
            print(f"=== Contact Search for '{q}' ({len(all_contacts)} found) ===")
            if not all_contacts:
                print("No matching contacts found in local address book or Exchange GAL.")
                return

            for idx, c in enumerate(all_contacts, 1):
                print(f"\n{idx}. {c['name']} [{c['source']}]")
                if c['email']:
                    print(f"   Email:      {c['email']}")
                if c['secondary_email']:
                    print(f"   Alt Email:  {c['secondary_email']}")
                if c['job_title'] or c['department']:
                    parts = [p for p in [c['job_title'], c['department']] if p]
                    print(f"   Role:       {', '.join(parts)}")
                if c['company']:
                    print(f"   Company:    {c['company']}")
                if c['mobile_phone'] or c['business_phone']:
                    phones = [p for p in [f"Mobile: {c['mobile_phone']}" if c['mobile_phone'] else "",
                                          f"Work: {c['business_phone']}" if c['business_phone'] else ""] if p]
                    print(f"   Phone:      {', '.join(phones)}")

    except Exception as e:
        err_payload = {"status": "error", "message": str(e)}
        if args.format == "json":
            print(json.dumps(err_payload, ensure_ascii=False))
        else:
            print(f"Error searching contacts: {e}", file=sys.stderr)
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
