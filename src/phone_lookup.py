import re

from src.parser import get_text_body

PHONE_RE = re.compile(
    r'(?:\+?1[\s.-]?)?\(?(\d{3})\)?[\s.-]?(\d{3})[\s.-]?(\d{4})'
)
EXT_RE = re.compile(r'(?:ext\.?|extension|x)\s*\.?\s*(\d{1,6})', re.IGNORECASE)
SEPARATOR_CHARS = set('()-. +')

# Markers that begin the quoted copy of an earlier message inside a reply.
# Anything after the first marker is someone else's text (usually the
# user's own original email, signature included), not the sender's.
QUOTE_MARKERS = [
    re.compile(r'^\s*-{2,}\s*Original Message\s*-{2,}', re.IGNORECASE | re.MULTILINE),
    re.compile(r'^\s*On .{5,120}wrote:', re.IGNORECASE | re.MULTILINE),
    re.compile(r'^\s*From:\s\S.*$', re.MULTILINE),
    re.compile(r'^\s*_{6,}\s*$', re.MULTILINE),
]


def strip_quoted_text(text):
    """Return only the sender's own (top) part of a reply: cut at the first
    quoted-message marker and drop '>'-prefixed quote lines."""
    if not text:
        return text
    cut = len(text)
    for pat in QUOTE_MARKERS:
        m = pat.search(text)
        if m:
            cut = min(cut, m.start())
    lines = [l for l in text[:cut].splitlines() if not l.lstrip().startswith('>')]
    return '\n'.join(lines)


def extract_phone(text, exclude_digits=None):
    """Best-effort extraction of a US phone number (with optional
    extension) from plain text, e.g. '(415) 555-0132 x204'. Numbers whose
    10 digits appear in exclude_digits (e.g. the user's own numbers) are
    skipped. Returns None if no confident match is found."""
    if not text:
        return None

    for m in PHONE_RE.finditer(text):
        matched_text = m.group(0)
        if not any(c in SEPARATOR_CHARS for c in matched_text):
            # A bare 10-digit run with no separators at all is more often a
            # tracking/ID number than a phone number in a signature block.
            continue

        digits = m.group(1) + m.group(2) + m.group(3)
        if exclude_digits and digits in exclude_digits:
            continue

        phone = f'({m.group(1)}) {m.group(2)}-{m.group(3)}'

        tail = text[m.end():m.end() + 20]
        ext_m = EXT_RE.search(tail)
        if ext_m:
            phone += f' x{ext_m.group(1)}'

        return phone

    return None


def get_own_phone_numbers(service, fetch_all_messages_fn, sample=8):
    """Detect the user's own phone number(s) by sampling their recent Sent
    mail: any number appearing in at least half the sampled messages is
    almost certainly part of their signature. Returns a set of 10-digit
    strings (empty if none found)."""
    messages = fetch_all_messages_fn(service, 'in:sent', sample)
    counts = {}
    for m in messages:
        msg = service.users().messages().get(
            userId='me', id=m['id'], format='full'
        ).execute(num_retries=3)
        text = get_text_body(msg['payload']) or ''
        seen_here = set()
        for pm in PHONE_RE.finditer(text):
            if any(c in SEPARATOR_CHARS for c in pm.group(0)):
                seen_here.add(pm.group(1) + pm.group(2) + pm.group(3))
        for num in seen_here:
            counts[num] = counts.get(num, 0) + 1

    threshold = max(2, len(messages) // 2)
    return {num for num, c in counts.items() if c >= threshold}


def find_phone_for_email(service, fetch_all_messages_fn, email, max_messages=3, exclude_digits=None):
    """Search the user's Inbox for messages from `email`, and return the
    first phone number found in the sender's own text (most recent message
    first). Quoted reply history is stripped first and the user's own
    numbers are excluded, so a recruiter reply that merely quotes the
    user's signature does not yield the user's number. Returns None if no
    messages exist from this sender, or none contain a recognizable
    phone number."""
    query = f'from:{email} in:inbox'
    messages = fetch_all_messages_fn(service, query, max_messages)

    for m in messages:
        msg = service.users().messages().get(
            userId='me', id=m['id'], format='full'
        ).execute(num_retries=3)
        text = strip_quoted_text(get_text_body(msg['payload']))
        phone = extract_phone(text, exclude_digits=exclude_digits)
        if phone:
            return phone

    return None
