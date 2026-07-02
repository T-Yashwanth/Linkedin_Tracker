import re
from datetime import datetime
from email.utils import parseaddr

STOPWORDS = {
    'inc', 'llc', 'ltd', 'corp', 'corporation', 'group', 'consulting',
    'solutions', 'solution', 'staffing', 'systems', 'system', 'technologies',
    'technology', 'tech', 'services', 'service', 'resources', 'resource',
    'company', 'co', 'the', 'and', 'llp', 'pvt', 'private', 'limited',
    'partners', 'partner', 'associates', 'global', 'international', 'usa',
    'us', 'staffing', 'holdings', 'enterprises', 'enterprise', 'it',
}

GENERIC_DOMAINS = {
    'gmail.com', 'yahoo.com', 'outlook.com', 'hotmail.com', 'icloud.com',
    'aol.com', 'live.com', 'msn.com', 'linkedin.com', 'me.com',
    'protonmail.com', 'ymail.com',
}


def normalize_words(name):
    name = re.sub(r'[^a-z0-9 ]', ' ', name.lower())
    return [w for w in name.split() if w not in STOPWORDS and len(w) >= 3]


def name_from_email_local(local_part):
    parts = re.split(r'[._\-0-9]+', local_part)
    parts = [p for p in parts if p.isalpha() and len(p) > 1]
    if len(parts) >= 2:
        return ' '.join(p.capitalize() for p in parts[:2])
    if len(parts) == 1 and len(parts[0]) >= 4:
        return parts[0].capitalize()
    return None


def fetch_sent_index(service, fetch_all_messages_fn, since_query_date, max_results=2000):
    """Build an index of {date, name, email, domain} for every To/Cc
    recipient of every message sent since since_query_date (a date object).
    'name' prefers the display name on the header, falling back to a name
    guessed from the email address itself (e.g. dipankar.k@x.com ->
    'Dipankar K')."""
    query = f'in:sent after:{since_query_date.strftime("%Y/%m/%d")}'
    messages = fetch_all_messages_fn(service, query, max_results)

    index = []
    for m in messages:
        msg = service.users().messages().get(
            userId='me', id=m['id'], format='metadata', metadataHeaders=['To', 'Cc', 'Date']
        ).execute()
        headers = {h['name']: h['value'] for h in msg['payload'].get('headers', [])}
        recipients_header = ', '.join(filter(None, [headers.get('To', ''), headers.get('Cc', '')]))
        if not recipients_header:
            continue

        seen_addrs = set()
        for recipient in recipients_header.split(','):
            disp_name, addr = parseaddr(recipient.strip())
            if not addr or '@' not in addr:
                continue
            addr = addr.lower().strip()
            if addr in seen_addrs:
                continue
            seen_addrs.add(addr)

            disp_name = disp_name.strip()
            if disp_name.lower() == addr:
                disp_name = ''

            name = disp_name or name_from_email_local(addr.split('@')[0]) or ''

            domain = addr.split('@')[-1]
            sent_dt = datetime.fromtimestamp(int(msg['internalDate']) / 1000)
            index.append({
                'date': sent_dt.date(),
                'name': name.strip(),
                'email': addr,
                'domain': domain,
            })

    return index


def find_hiring_managers(company, applied_date, sent_index, min_score=0.3):
    """Match a company name against sent emails (To/Cc) on the same calendar
    date by comparing significant company-name words against each
    recipient's email domain. Returns a list of {'name': ..., 'email': ...}
    for every distinct matching address, best score first (empty list if
    none match)."""
    company_words = normalize_words(company)
    if not company_words:
        return []

    candidates = [s for s in sent_index if s['date'] == applied_date]
    scored = []
    seen_emails = set()

    for c in candidates:
        if c['domain'] in GENERIC_DOMAINS or c['email'] in seen_emails:
            continue
        domain_base = c['domain'].split('.')[0]
        matches = sum(1 for w in company_words if len(w) >= 4 and w in domain_base)
        if matches == 0:
            continue
        score = matches / len(company_words)
        if score < min_score:
            continue
        seen_emails.add(c['email'])
        scored.append((score, c))

    scored.sort(key=lambda x: x[0], reverse=True)
    return [{'name': c['name'], 'email': c['email']} for _, c in scored]
