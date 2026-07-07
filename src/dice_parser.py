import re

SUBJECT_RE = re.compile(r'^Application for (.+?) at (.+?) sent$')


def parse_dice_subject(subject):
    """Extract job title and company from a Dice 'Application for X at Y
    sent' confirmation email subject. Dice's email body has no job-specific
    link (only opaque click-tracking redirects), so there's no job link to
    capture for this source."""
    m = SUBJECT_RE.match(subject.strip())
    if not m:
        return {'title': None, 'company': None}
    return {'title': m.group(1).strip(), 'company': m.group(2).strip()}
