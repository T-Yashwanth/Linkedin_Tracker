import base64
import re

from bs4 import BeautifulSoup


def get_html_body(payload):
    """Recursively find the text/html part of a Gmail message payload."""
    if payload.get('mimeType') == 'text/html' and payload.get('body', {}).get('data'):
        data = payload['body']['data']
        return base64.urlsafe_b64decode(data).decode('utf-8', errors='ignore')

    for part in payload.get('parts', []) or []:
        html = get_html_body(part)
        if html:
            return html

    return None


def get_text_body(payload):
    """Recursively find plain text content for a Gmail message payload.
    Prefers an explicit text/plain part; falls back to stripping tags from
    text/html if no text/plain part exists. Returns None if no text
    content can be found."""
    if payload.get('mimeType') == 'text/plain' and payload.get('body', {}).get('data'):
        data = payload['body']['data']
        return base64.urlsafe_b64decode(data).decode('utf-8', errors='ignore')

    for part in payload.get('parts', []) or []:
        text = get_text_body(part)
        if text:
            return text

    html = get_html_body(payload)
    if html:
        return BeautifulSoup(html, 'lxml').get_text('\n', strip=True)

    return None


def parse_application_email(subject, html):
    """Extract company, title, location, applied date and job link from a
    LinkedIn 'your application was sent to X' notification email."""

    soup = BeautifulSoup(html, 'lxml') if html else None
    text = soup.get_text('\n', strip=True) if soup else ''

    company = None
    m = re.search(r'application was sent to (.+?)\s*$', subject.strip())
    if m:
        company = m.group(1).strip()

    title = None
    job_link = None
    if soup:
        for a in soup.find_all('a', href=True):
            href = a['href']
            if '/jobs/view/' in href or '/comm/jobs/view/' in href:
                candidate = a.get_text(strip=True)
                if candidate and len(candidate) > 1:
                    title = candidate
                    m2 = re.match(r'(https?://[^?]*?/jobs/view/\d+)', href)
                    job_link = m2.group(1) if m2 else href
                    break

    applied_date = None
    m = re.search(r'Applied on ([A-Za-z]+ \d{1,2},?\s*\d{4})', text)
    if m:
        applied_date = m.group(1).replace(',', ',')

    location = None
    if company:
        m = re.search(re.escape(company) + r'\s*[·•]\s*([^\n]+)', text)
        if m:
            location = m.group(1).strip()

    return {
        'company': company,
        'title': title,
        'location': location,
        'applied_date': applied_date,
        'job_link': job_link,
    }
