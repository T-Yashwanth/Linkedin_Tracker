import argparse
import json
import os

from dateutil import parser as dateparser

from src.gmail_client import get_gmail_service
from src.recruiter_directory import (
    ContactBook, is_bot_headers, parse_recipients, ensure_directory,
    read_directory_rows, merge_contact, write_directory,
)
from src.phone_lookup import (
    scan_contact_signature, get_own_phone_numbers, get_own_linkedin_slugs,
)
from update_tracker import fetch_all_messages

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROCESSED_FILE = os.path.join(BASE_DIR, 'data', 'directory_processed_ids.json')
DEFAULT_DIRECTORY = os.path.join(BASE_DIR, 'data', 'Recruiters.xlsx')

METADATA_HEADERS = ['To', 'Cc', 'From', 'Subject', 'List-Unsubscribe', 'Precedence', 'Auto-Submitted']


def load_processed():
    if os.path.exists(PROCESSED_FILE):
        with open(PROCESSED_FILE) as f:
            return set(json.load(f))
    return set()


def save_processed(ids):
    os.makedirs(os.path.dirname(PROCESSED_FILE), exist_ok=True)
    with open(PROCESSED_FILE, 'w') as f:
        json.dump(sorted(ids), f, indent=2)


def fetch_headers(service, msg_id):
    msg = service.users().messages().get(
        userId='me', id=msg_id, format='metadata', metadataHeaders=METADATA_HEADERS
    ).execute(num_retries=3)
    return {h['name']: h['value'] for h in msg['payload'].get('headers', [])}


def discover_contacts(service, since_date, max_results, processed, own_email):
    """Scan Sent (To/Cc) and Inbox (From) mail for eligible human contacts,
    marking every scanned message ID as processed regardless of outcome.
    own_email (lowercased) is never added -- guards against the user
    showing up as their own contact (e.g. CC'd on their own outgoing mail
    from a non-generic-domain address)."""
    book = ContactBook()

    sent_query = 'in:sent'
    inbox_query = 'in:inbox'
    if since_date:
        suffix = f' after:{since_date.strftime("%Y/%m/%d")}'
        sent_query += suffix
        inbox_query += suffix

    for m in fetch_all_messages(service, sent_query, max_results):
        msg_id = m['id']
        if msg_id in processed:
            continue
        headers = fetch_headers(service, msg_id)
        subject = headers.get('Subject', '')
        for name, email in parse_recipients(headers.get('To', '')):
            if email != own_email:
                book.add(email, name=name, subject=subject)
        for name, email in parse_recipients(headers.get('Cc', '')):
            if email != own_email:
                book.add(email, name=name, subject=subject)
        processed.add(msg_id)

    for m in fetch_all_messages(service, inbox_query, max_results):
        msg_id = m['id']
        if msg_id in processed:
            continue
        headers = fetch_headers(service, msg_id)
        processed.add(msg_id)
        if is_bot_headers(headers):
            continue
        subject = headers.get('Subject', '')
        for name, email in parse_recipients(headers.get('From', '')):
            if email != own_email:
                book.add(email, name=name, subject=subject)

    return book


def main():
    ap = argparse.ArgumentParser(description='Build/update the recruiter directory from Gmail.')
    ap.add_argument('--directory', default=DEFAULT_DIRECTORY, help='Path to the directory xlsx to update')
    ap.add_argument('--max-results', type=int, default=2000, help='Max Gmail messages to scan per mailbox')
    ap.add_argument('--since', default=None, help='Only scan mail on/after this date, e.g. 2026-06-24')
    ap.add_argument('--rebuild', action='store_true', help='Start the directory fresh instead of merging')
    ap.add_argument('--dry-run', action='store_true', help='Discover contacts and print a summary without writing or scanning signatures')
    args = ap.parse_args()

    since_date = dateparser.parse(args.since).date() if args.since else None

    service = get_gmail_service()
    processed = set() if args.rebuild else load_processed()
    own_email = service.users().getProfile(userId='me').execute().get('emailAddress', '').lower()

    print('Scanning Sent and Inbox mail for recruiter/contact emails...')
    book = discover_contacts(service, since_date, args.max_results, processed, own_email)
    contacts = dict(book.items())
    print(f'Found {len(contacts)} eligible contact(s) after filtering out generic/job-board/automated senders.')

    if args.dry_run:
        by_domain = {}
        for email, c in contacts.items():
            by_domain.setdefault(c['domain'], []).append((email, c))
        for domain in sorted(by_domain):
            entries = by_domain[domain]
            print(f'{domain} ({len(entries)}):')
            for email, c in sorted(entries):
                positions = '; '.join(sorted(c['subjects'])) or '(no subject captured)'
                print(f"  {c['name'] or '(no name)'} <{email}> | {positions}")
        print(f'{len(contacts)} contact(s) would be added/updated (dry run, nothing saved, no signature scans performed).')
        return

    wb = ensure_directory(args.directory, args.rebuild)
    ws = wb.active
    existing_by_email, no_email_rows = read_directory_rows(ws)

    own_numbers = get_own_phone_numbers(service, fetch_all_messages)
    own_slugs = get_own_linkedin_slugs(service, fetch_all_messages)

    signature_cache = {}
    scanned = 0

    def lookup_signature(email):
        nonlocal scanned
        if email not in signature_cache:
            signature_cache[email] = scan_contact_signature(
                service, fetch_all_messages, email,
                exclude_digits=own_numbers, exclude_slugs=own_slugs,
            )
            scanned += 1
        return signature_cache[email]

    new_count = 0
    updated_count = 0

    for email, c in contacts.items():
        existing = existing_by_email.get(email)
        needs_scan = existing is None or not all([existing[1], existing[3], existing[5]])
        sig = lookup_signature(email) if needs_scan else {'phone': None, 'title': None, 'linkedin': None}

        merged = merge_contact(
            existing, c['company'], sig['title'], c['name'],
            sig['phone'], email, sig['linkedin'], c['subjects'],
        )
        if existing is None:
            new_count += 1
        else:
            updated_count += 1
        existing_by_email[email] = merged

    print(f'Signature-scanned {scanned} contact(s) for title/phone/LinkedIn (skipped contacts already fully filled in).')

    try:
        write_directory(args.directory, existing_by_email, no_email_rows)
    except PermissionError:
        print(f'Could not save {args.directory} -- it looks like the file is open in Excel. Close it and re-run.')
        raise

    save_processed(processed)
    print(f'Added {new_count} new contact(s), updated {updated_count} existing entr{"y" if updated_count == 1 else "ies"}; '
          f'directory now has {len(existing_by_email) + len(no_email_rows)} row(s), saved to {args.directory}')


if __name__ == '__main__':
    main()
