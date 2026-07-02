# LinkedIn Application Tracker

Scans Gmail for LinkedIn "your application was sent to X" confirmation
emails, matches a hiring manager name/email from your Sent folder when
possible, and keeps an Excel tracker updated (sorted by date, no duplicates,
manual edits preserved).

## Project structure

```
Linkedin_Tracker/
├── tracker_venv/            # virtual environment (not committed)
├── linkedin_tracker/        # package: Gmail auth, email parsing, matching
│   ├── gmail_client.py
│   ├── parser.py
│   └── sent_matcher.py
├── data/
│   ├── LinkedIn_Job_Tracker.xlsx              # the tracker you open/edit
│   └── LinkedIn_Job_Tracker_full_history_backup.xlsx
├── secrets/
│   ├── credentials.json     # OAuth client (not committed)
│   └── token.json           # created after first login (not committed)
├── update_tracker.py        # entry point — run this
├── processed_ids.json       # tracks which emails were already imported
└── requirements.txt
```

## One-time setup

1. Get an OAuth client:
   - In [Google Cloud Console](https://console.cloud.google.com/apis/library/gmail.googleapis.com),
     enable the **Gmail API** for your project, create an **OAuth client ID**
     (type: Desktop app), and download it as `secrets/credentials.json`.
   - If the OAuth consent screen is in "Testing" status, add your own Google
     account as a test user under that screen.
2. Create the virtual environment and install dependencies.

### Command Prompt (cmd)
```cmd
cd path\to\Linkedin_Tracker
python -m venv tracker_venv
tracker_venv\Scripts\activate
pip install -r requirements.txt
```

### Git Bash
```bash
cd /c/path/to/Linkedin_Tracker
python -m venv tracker_venv
source tracker_venv/Scripts/activate
pip install -r requirements.txt
```

(Only needed once — or again after `requirements.txt` changes.)

## Running it

### Command Prompt (cmd)
```cmd
tracker_venv\Scripts\python update_tracker.py
```

### Git Bash
```bash
tracker_venv/Scripts/python update_tracker.py
```

- First run opens a browser window to sign in with your Google account and
  approve read-only Gmail access. This creates `secrets/token.json`, which is
  reused on future runs (no repeated logins).
- If `data/LinkedIn_Job_Tracker.xlsx` doesn't exist yet, it's created
  automatically with the correct headers — no template file needed.
- Every run reads the whole sheet (including anything you filled in by hand),
  merges in new applications, re-sorts everything by date ascending, and
  rewrites the sheet — your manual edits and manually-added rows are kept.
- Already-processed emails are recorded in `processed_ids.json`, so re-running
  only pulls new applications since the last run.
- **Close the xlsx in Excel before running** — an open file can't be saved to.

Useful flags:
- `--dry-run` — parse and print what would happen, without touching the xlsx.
- `--tracker <path>` — write to a different xlsx file.
- `--since YYYY-MM-DD` — only include applications on/after this date.
- `--rebuild` — start the tracker fresh (blank headers) instead of merging.
- `--no-hiring-manager` — skip matching hiring manager name/email from Sent mail.
- `--max-results N` — limit how many Gmail messages to scan (default 2000).

Example:
```bash
tracker_venv/Scripts/python update_tracker.py --since 2026-06-24 --rebuild
```

## What gets filled in

For each application email:
`S.no`, `Date`, `Title`, `Company Full Name`, `Applied In Linkedin` (`Yes`),
`Website Applied` (shortened job link, opens the application directly).

`Hiring Manager In Linkedin` and `Company Email` are auto-filled when a
same-day Sent email to a matching company domain is found (name is taken from
the email's display name, or guessed from the address itself, e.g.
`dipankar.k@x.com` → "Dipankar K"). `Contact Number For Job Post` and
`Comment Section` are left blank for you.

## Re-running periodically

Run it whenever you want to sync new applications — safe to run repeatedly.
You can also schedule it (Windows Task Scheduler) to run daily.

## Security notes

- `secrets/credentials.json` is your OAuth client secret; `secrets/token.json`
  (created after first login) holds your personal access/refresh token. Keep
  both private — don't commit them to a public repo (`.gitignore` already
  excludes `secrets/`).
- The app only requests `gmail.readonly` scope — it can read email, not send
  or delete anything.
