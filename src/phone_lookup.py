import re

from src.parser import get_text_body

PHONE_RE = re.compile(
    r'(?:\+?1[\s.-]?)?\(?(\d{3})\)?[\s.-]?(\d{3})[\s.-]?(\d{4})'
)
EXT_RE = re.compile(r'(?:ext\.?|extension|x)\s*\.?\s*(\d{1,6})', re.IGNORECASE)
SEPARATOR_CHARS = set('()-. +')


def extract_phone(text):
    """Best-effort extraction of a US phone number (with optional
    extension) from plain text, e.g. '(415) 555-0132 x204'. Returns None
    if no confident match is found."""
    if not text:
        return None

    for m in PHONE_RE.finditer(text):
        matched_text = m.group(0)
        if not any(c in SEPARATOR_CHARS for c in matched_text):
            # A bare 10-digit run with no separators at all is more often a
            # tracking/ID number than a phone number in a signature block.
            continue

        phone = f'({m.group(1)}) {m.group(2)}-{m.group(3)}'

        tail = text[m.end():m.end() + 20]
        ext_m = EXT_RE.search(tail)
        if ext_m:
            phone += f' x{ext_m.group(1)}'

        return phone

    return None


def find_phone_for_email(service, fetch_all_messages_fn, email, max_messages=3):
    """Search the user's Inbox for messages from `email`, and return the
    first phone number found in their plain-text content (most recent
    message first). Returns None if no messages exist from this sender, or
    none contain a recognizable phone number."""
    query = f'from:{email} in:inbox'
    messages = fetch_all_messages_fn(service, query, max_messages)

    for m in messages:
        msg = service.users().messages().get(
            userId='me', id=m['id'], format='full'
        ).execute(num_retries=3)
        text = get_text_body(msg['payload'])
        phone = extract_phone(text)
        if phone:
            return phone

    return None
