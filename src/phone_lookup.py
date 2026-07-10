import re

from src.parser import get_text_body, get_html_body

PHONE_RE = re.compile(
    r'(?:\+?1[\s.-]?)?\(?(\d{3})\)?[\s.-]?(\d{3})[\s.-]?(\d{4})'
)
EXT_RE = re.compile(r'(?:ext\.?|extension|x)\s*\.?\s*(\d{1,6})', re.IGNORECASE)
SEPARATOR_CHARS = set('()-. +')

LINKEDIN_PROFILE_RE = re.compile(
    r'(?:https?://)?(?:www\.)?linkedin\.com/in/([A-Za-z0-9\-_%.]+)', re.IGNORECASE
)

TITLE_KEYWORDS_RE = re.compile(
    r'recruit(?:er|ing|ment)?|talent|staffing|sourc(?:er|ing)|\bh\.?r\.?\b|'
    r'human resources|account manager|bench sales|resource manager|'
    r'delivery manager|acquisition|people operations|hiring',
    re.IGNORECASE,
)

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


def extract_title(text):
    """Best-effort extraction of a job title from the signature zone of an
    already-quote-stripped reply (the last ~12 non-empty lines). Returns
    None if nothing confidently looks like a title."""
    if not text:
        return None

    lines = [l for l in text.splitlines() if l.strip()]
    zone = lines[-12:]

    for line in zone:
        for seg in re.split(r'[|•·–,]', line):
            seg = seg.strip()
            if not (3 <= len(seg) <= 60):
                continue
            if '@' in seg or 'http' in seg.lower():
                continue
            if sum(c.isdigit() for c in seg) >= 4:
                continue
            if TITLE_KEYWORDS_RE.search(seg):
                return seg

    return None


def _normalize_linkedin_url(slug):
    slug = slug.split('?')[0].rstrip('/')
    return f'https://www.linkedin.com/in/{slug.lower()}'


def extract_linkedin_url(text, html=None, exclude_slugs=None):
    """Best-effort extraction of a recruiter's LinkedIn profile URL
    (linkedin.com/in/... only -- never /company/ or /jobs/view/ links).
    Prefers a match in the already-quote-stripped text; falls back to the
    raw HTML body (quoting can't be reliably stripped there), excluding
    any slug in exclude_slugs (the user's own profile). Returns None if
    nothing is found."""
    exclude_slugs = {s.lower() for s in (exclude_slugs or set())}

    if text:
        m = LINKEDIN_PROFILE_RE.search(text)
        if m and m.group(1).lower() not in exclude_slugs:
            return _normalize_linkedin_url(m.group(1))

    if html:
        for m in LINKEDIN_PROFILE_RE.finditer(html):
            if m.group(1).lower() not in exclude_slugs:
                return _normalize_linkedin_url(m.group(1))

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


def get_own_linkedin_slugs(service, fetch_all_messages_fn, sample=8):
    """Detect the user's own LinkedIn profile slug(s) by sampling recent
    Sent mail (text and HTML), the same way get_own_phone_numbers detects
    phone numbers. Returns a set of lowercase slugs (empty if none found)."""
    messages = fetch_all_messages_fn(service, 'in:sent', sample)
    counts = {}
    for m in messages:
        msg = service.users().messages().get(
            userId='me', id=m['id'], format='full'
        ).execute(num_retries=3)
        text = get_text_body(msg['payload']) or ''
        html = get_html_body(msg['payload']) or ''
        seen_here = set()
        for pm in LINKEDIN_PROFILE_RE.finditer(text + '\n' + html):
            seen_here.add(pm.group(1).lower())
        for slug in seen_here:
            counts[slug] = counts.get(slug, 0) + 1

    threshold = max(2, len(messages) // 2)
    return {slug for slug, c in counts.items() if c >= threshold}


def scan_contact_signature(service, fetch_all_messages_fn, email, max_messages=3,
                            exclude_digits=None, exclude_slugs=None):
    """Search the user's Inbox for messages from `email` and extract a
    phone number, job title, and LinkedIn profile URL from the sender's
    own text (quoted reply history stripped first), stopping as soon as
    all three are found. Returns {'phone', 'title', 'linkedin'}, each
    None if not found. Returns all-None if no messages exist from this
    sender."""
    result = {'phone': None, 'title': None, 'linkedin': None}
    query = f'from:{email} in:inbox'
    messages = fetch_all_messages_fn(service, query, max_messages)

    for m in messages:
        msg = service.users().messages().get(
            userId='me', id=m['id'], format='full'
        ).execute(num_retries=3)
        raw_text = get_text_body(msg['payload'])
        stripped = strip_quoted_text(raw_text)
        html = get_html_body(msg['payload'])

        if result['phone'] is None:
            result['phone'] = extract_phone(stripped, exclude_digits=exclude_digits)
        if result['title'] is None:
            result['title'] = extract_title(stripped)
        if result['linkedin'] is None:
            result['linkedin'] = extract_linkedin_url(stripped, html, exclude_slugs=exclude_slugs)

        if all(result.values()):
            break

    return result


def find_phone_for_email(service, fetch_all_messages_fn, email, max_messages=3, exclude_digits=None):
    """Search the user's Inbox for messages from `email`, and return the
    first phone number found in the sender's own text (most recent message
    first). Quoted reply history is stripped first and the user's own
    numbers are excluded, so a recruiter reply that merely quotes the
    user's signature does not yield the user's number. Returns None if no
    messages exist from this sender, or none contain a recognizable
    phone number."""
    return scan_contact_signature(
        service, fetch_all_messages_fn, email,
        max_messages=max_messages, exclude_digits=exclude_digits,
    )['phone']
