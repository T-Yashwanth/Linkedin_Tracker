import argparse
import functools
import os
import subprocess
import sys
from collections import Counter

import openpyxl
from dateutil import parser as dateparser

print = functools.partial(print, flush=True)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPDATE_TRACKER = os.path.join(BASE_DIR, 'update_tracker.py')


def discover_accounts(base_dir):
    """Any immediate subdirectory containing secrets/credentials.json and
    data/Job_Tracker.xlsx is treated as an account. The project's own
    top-level secrets/ and data/ never match this (they're not nested
    inside a subdirectory), so the shared swap slot is naturally excluded."""
    accounts = []
    for name in sorted(os.listdir(base_dir)):
        path = os.path.join(base_dir, name)
        if not os.path.isdir(path):
            continue
        creds = os.path.join(path, 'secrets', 'credentials.json')
        tracker = os.path.join(path, 'data', 'Job_Tracker.xlsx')
        if os.path.exists(creds) and os.path.exists(tracker):
            accounts.append({'name': name, 'dir': path, 'secrets_dir': os.path.join(path, 'secrets'),
                              'tracker': tracker})
    return accounts


def excel_is_running():
    result = subprocess.run(
        ['tasklist', '/FI', 'IMAGENAME eq EXCEL.EXE'],
        capture_output=True, text=True,
    )
    return 'EXCEL.EXE' in result.stdout


def read_rows(tracker_path):
    wb = openpyxl.load_workbook(tracker_path)
    ws = wb.active
    return list(ws.iter_rows(min_row=2, values_only=True))


def get_last_date(tracker_path):
    rows = read_rows(tracker_path)
    dates = [dateparser.parse(r[1]).date() for r in rows if r[1]]
    return max(dates) if dates else None


def verify(tracker_path):
    rows = read_rows(tracker_path)
    dates = [dateparser.parse(r[1]).date() for r in rows if r[1]]
    links = [r[5] for r in rows if r[5]]
    dupe_links = {k: v for k, v in Counter(links).items() if v > 1}
    reach_emails = [r[7] for r in rows if r[4] == 'Email Reach Out' and r[7]]
    dupe_emails = {k: v for k, v in Counter(e.lower() for e in reach_emails).items() if v > 1}
    snos = [r[0] for r in rows]
    return {
        'row_count': len(rows),
        'date_range': (min(dates), max(dates)) if dates else (None, None),
        'dupe_job_links': len(dupe_links),
        'dupe_reachout_emails': len(dupe_emails),
        'sno_sequential': snos == list(range(1, len(rows) + 1)),
        'platforms': dict(Counter(r[4] for r in rows)),
        'with_phone': sum(1 for r in rows if r[8]),
    }


def run_account(account, dry_run):
    print(f"\n{'=' * 60}\nAccount: {account['name']}\n{'=' * 60}")

    rows_before = len(read_rows(account['tracker']))
    since_date = get_last_date(account['tracker'])
    since_str = since_date.isoformat() if since_date else None
    print(f"Last date in tracker: {since_str or '(none, full history)'}")

    cmd = [
        sys.executable, UPDATE_TRACKER,
        '--tracker', account['tracker'],
        '--secrets-dir', account['secrets_dir'],
        '--include-reachout', '--include-phone',
    ]
    if since_str:
        cmd += ['--since', since_str]
    if dry_run:
        cmd += ['--dry-run']

    env = dict(os.environ)
    env['PYTHONIOENCODING'] = 'utf-8'

    result = subprocess.run(cmd, env=env)

    if result.returncode != 0:
        print(f"!! Update FAILED for {account['name']} (exit code {result.returncode}) -- skipping verification.")
        return {'name': account['name'], 'ok': False, 'rows_before': rows_before}

    if dry_run:
        return {'name': account['name'], 'ok': True, 'dry_run': True, 'rows_before': rows_before}

    stats = verify(account['tracker'])
    print(f"Verification: {stats['row_count']} rows, date range {stats['date_range'][0]} -> {stats['date_range'][1]}, "
          f"duplicate job links: {stats['dupe_job_links']}, duplicate reach-out emails: {stats['dupe_reachout_emails']}, "
          f"S.no sequential: {stats['sno_sequential']}, phones found: {stats['with_phone']}")
    return {'name': account['name'], 'ok': True, 'dry_run': False, 'rows_before': rows_before, **stats}


def main():
    ap = argparse.ArgumentParser(description='Update every account folder\'s tracker in place, one after another.')
    ap.add_argument('--dry-run', action='store_true', help='Preview only; passes --dry-run through to every account')
    args = ap.parse_args()

    accounts = discover_accounts(BASE_DIR)
    if not accounts:
        print('No account folders found (looked for <folder>/secrets/credentials.json + <folder>/data/Job_Tracker.xlsx).')
        return

    print(f"Discovered {len(accounts)} account folder(s): {', '.join(a['name'] for a in accounts)}")

    if not args.dry_run and excel_is_running():
        print('\nExcel appears to be running. Close it first so tracker files can be saved, then re-run.')
        return

    summaries = [run_account(a, args.dry_run) for a in accounts]

    print(f"\n{'=' * 60}\nSummary\n{'=' * 60}")
    for s in summaries:
        if not s.get('ok'):
            print(f"{s['name']}: FAILED")
        elif s.get('dry_run'):
            print(f"{s['name']}: dry run only (rows before: {s['rows_before']})")
        else:
            print(f"{s['name']}: {s['rows_before']} -> {s['row_count']} rows, "
                  f"dupe job links: {s['dupe_job_links']}, dupe reach-out emails: {s['dupe_reachout_emails']}, "
                  f"phones found: {s['with_phone']}")


if __name__ == '__main__':
    main()
