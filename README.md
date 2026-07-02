# Job Application Tracker

This tool reads your Gmail inbox, finds the application-confirmation emails
sent by **LinkedIn** and **Dice** every time you apply to a job, and
automatically fills an Excel spreadsheet with the company name, job title,
date, platform, and (for LinkedIn) a link back to the job posting. If it can
also find an email you personally sent to someone at that company on the
same day, it fills in that person's name and email address too.

You run it from a terminal (a command-line window) on your own computer.
It never sends or deletes anything in your Gmail — it only reads.

---

## Before you start: things you need installed

1. **Python** (version 3.9 or newer). Check if you already have it by
   opening a terminal and running:
   ```
   python --version
   ```
   If that fails, download Python from [python.org/downloads](https://www.python.org/downloads/)
   and install it (on Windows, tick "Add Python to PATH" during install).

2. **Git** (only needed if you're cloning from GitHub rather than
   downloading a ZIP). Get it from [git-scm.com](https://git-scm.com/downloads).

3. **A terminal.** On Windows you can use **Command Prompt (cmd)** or
   **Git Bash** (installed alongside Git). This README gives commands for
   both — use whichever one you have open.

4. **Excel** (or any spreadsheet app that opens `.xlsx` files), to view the
   tracker once it's filled in.

---

## Step 1 — Get the project onto your computer

**Option A — Download as ZIP (easiest, no Git knowledge needed):**
On the GitHub page for this project, click the green **Code** button →
**Download ZIP**, then extract it anywhere on your computer (e.g. your
Documents folder).

**Option B — Fork and clone (if you want your own copy on GitHub too):**
1. Click **Fork** on the GitHub page — this creates a copy of the project
   under your own GitHub account.
2. On your fork's page, click **Code** → copy the HTTPS URL.
3. In a terminal, run:
   ```bash
   git clone <paste the URL you copied>
   cd Linkedin_Tracker
   ```

**Option C — Just clone the original (if you don't need your own fork):**
```bash
git clone https://github.com/T-Yashwanth/Linkedin_Tracker.git
cd Linkedin_Tracker
```

From here on, every command assumes your terminal is open **inside the
`Linkedin_Tracker` folder**.

---

## Step 2 — Create a Gmail API credential (one-time, ~5 minutes)

This tool needs your permission to read your Gmail. Google requires every
app to register itself first — this step creates that registration.

1. Go to [Google Cloud Console](https://console.cloud.google.com/) and sign
   in with the Google account whose Gmail you want to scan.
2. Create a new project (top-left project dropdown → **New Project** → give
   it any name, e.g. "Job Tracker").
3. Go to [APIs & Services → Library](https://console.cloud.google.com/apis/library/gmail.googleapis.com),
   search for **Gmail API**, and click **Enable**.
4. Go to **APIs & Services → OAuth consent screen**. Choose **External**,
   fill in the required fields (app name, your email), and save. When asked
   for "test users," add your own Gmail address.
5. Go to **APIs & Services → Credentials** → **Create Credentials** →
   **OAuth client ID**. For "Application type," choose **Desktop app**, give
   it any name, and click **Create**.
6. Click **Download JSON** on the credential you just created.
7. Inside the project folder, create a folder called `secrets` (if it
   doesn't already exist), and put the downloaded file there, renamed to
   exactly `credentials.json`. It should end up at:
   ```
   Linkedin_Tracker/secrets/credentials.json
   ```

You only need to do this once. `credentials.json` is specific to you —
never share it or upload it anywhere public.

---

## Step 3 — Set up the Python environment

A "virtual environment" is just an isolated folder that holds the exact
Python packages this project needs, so it doesn't interfere with anything
else on your computer. We create one called `tracker_venv`.

**Command Prompt (cmd):**
```cmd
python -m venv tracker_venv
tracker_venv\Scripts\activate
pip install -r requirements.txt
```

**Git Bash:**
```bash
python -m venv tracker_venv
source tracker_venv/Scripts/activate
pip install -r requirements.txt
```

You'll know it worked if your terminal prompt now shows `(tracker_venv)` at
the start of the line, and the `pip install` finishes with no red error
text. You only need to do this once — or again if `requirements.txt`
changes in a future update.

---

## Step 4 — Run it

**Command Prompt (cmd):**
```cmd
tracker_venv\Scripts\python update_tracker.py
```

**Git Bash:**
```bash
tracker_venv/Scripts/python update_tracker.py
```

What happens the first time:
1. Your web browser opens automatically, asking you to sign in to the
   Google account you want to scan and approve **read-only** Gmail access.
2. After you approve, the browser tab can be closed — the terminal takes
   over automatically. This creates `secrets/token.json`, so you won't be
   asked to log in again on future runs.
3. If `data/LinkedIn_Job_Tracker.xlsx` doesn't exist yet, it's created
   automatically with the right column headers — you don't need to make
   this file yourself.
4. It searches your Gmail, parses each application email, tries to find a
   matching hiring-manager email from your Sent folder, and writes
   everything into the spreadsheet — sorted by date.

**Important:** close the tracker `.xlsx` file in Excel before running the
script. Excel locks the file while it's open, so the script can't save
changes to it and will show a permission error.

---

## Understanding the flags

A "flag" is an extra option you type after `update_tracker.py` to change
how it behaves. You can combine several at once.

### `--dry-run`
Preview only — prints what it *would* add to the spreadsheet, without
actually writing anything or changing any files. Use this to sanity-check
before a real run, especially the first time or after changing something.
```bash
tracker_venv/Scripts/python update_tracker.py --dry-run
```

### `--since YYYY-MM-DD`
Only look at applications on or after this date. Speeds things up a lot on
a large mailbox, since it also narrows the Gmail search itself (not just
what gets written to the spreadsheet).
```bash
tracker_venv/Scripts/python update_tracker.py --since 2026-06-24
```

### `--rebuild`
**Use this whenever the tracker file is missing, was deleted, or you want
to force everything to be re-imported from scratch.** Normally, once an
email has been read by the script, its ID gets saved in `processed_ids.json`
so it's never re-imported again — that file, not the spreadsheet, is the
script's memory of "what's already been done." This means:

> ⚠️ **If you delete `data/LinkedIn_Job_Tracker.xlsx` but leave
> `processed_ids.json` untouched, and then run the script normally, you'll
> end up with an empty spreadsheet.** The script will recreate a blank
> file, look at Gmail, see that every application email is already marked
> "processed," and skip all of them — nothing gets written.

`--rebuild` fixes this: it deletes/recreates the spreadsheet **and**
temporarily ignores `processed_ids.json` for that run, so every matching
email is treated as new again. At the end, `processed_ids.json` is
rewritten to match whatever ended up in the spreadsheet.
```bash
tracker_venv/Scripts/python update_tracker.py --rebuild
```
If you only want to rebuild a recent window (not your entire Gmail
history), combine it with `--since`:
```bash
tracker_venv/Scripts/python update_tracker.py --since 2026-06-24 --rebuild
```

### `--no-hiring-manager`
Skips the step that scans your Sent folder to guess hiring-manager names
and emails. This makes the run noticeably faster, but the "Hiring Manager
In Linkedin" and "Company Email" columns will be left blank. Use this if
you don't need that lookup, or just want a quick sync.
```bash
tracker_venv/Scripts/python update_tracker.py --no-hiring-manager
```

### `--tracker <path>`
Read/write a different spreadsheet file instead of the default
`data/LinkedIn_Job_Tracker.xlsx`. Handy for testing on a throwaway copy
before trusting a real run.
```bash
tracker_venv/Scripts/python update_tracker.py --tracker data/test_copy.xlsx
```

### `--max-results N`
Caps how many Gmail messages the script will scan in one run (default
2000). You'll rarely need to change this — it's mainly a safety limit.
```bash
tracker_venv/Scripts/python update_tracker.py --max-results 500
```

### `--source {all|linkedin|dice}`
Restrict this run to a single platform instead of scanning both. Only
changes what's *fetched* — your existing rows from the other platform are
never touched or removed. Defaults to `all`.
```bash
tracker_venv/Scripts/python update_tracker.py --source dice
tracker_venv/Scripts/python update_tracker.py --source linkedin
```

### Common combinations

```bash
# Safe test: see what a full rebuild would produce, without saving anything
tracker_venv/Scripts/python update_tracker.py --since 2026-06-24 --rebuild --dry-run

# Actually do that rebuild for real
tracker_venv/Scripts/python update_tracker.py --since 2026-06-24 --rebuild

# Normal day-to-day use — pulls in whatever's new from both LinkedIn and Dice
tracker_venv/Scripts/python update_tracker.py

# Only sync new Dice applications
tracker_venv/Scripts/python update_tracker.py --source dice

# Quick sync, skip the slower hiring-manager lookup
tracker_venv/Scripts/python update_tracker.py --no-hiring-manager
```

---

## What each spreadsheet column means

| Column | Filled automatically? | Notes |
|---|---|---|
| S.no | Yes | Renumbered every run, in date order |
| Date | Yes | The date you applied, from the email |
| Title | Yes | Job title |
| Company Full Name | Yes | Company name |
| Platform | Yes | `LinkedIn` or `Dice`, depending on which email it came from |
| Website Applied | LinkedIn only | Shortened link straight to the job posting. Dice's emails don't include a job-specific link, so this is blank for Dice rows |
| Hiring Manager In Linkedin | Sometimes | Only if a matching Sent email was found on the same date, for either platform. If you emailed multiple people at that company that day (To **or** Cc), all of them are listed here, separated by `; ` |
| Company Email | Sometimes | Same matches as above, in the same order, also `; `-separated |
| Contact Number For Job Post | No | Left blank for you to fill in |
| Comment Section | No | Left blank for you to fill in |

Every run **preserves anything you've typed in by hand** — manual notes,
contact numbers, corrected names, even entire rows you added yourself.
The script merges your existing data with newly found applications and
re-sorts by date; it never throws away what's already there (unless you
use `--rebuild`, which starts over).

---

## Running it again later

Just run the same command again whenever you want to pull in new
applications — it's safe to run as often as you like, duplicates are never
added twice. If you want it to run automatically (e.g. every morning), you
can set it up as a scheduled task using Windows Task Scheduler.

---

## Troubleshooting

**"PermissionError" / can't save the file** — Close the tracker `.xlsx` in
Excel first; Excel locks the file while open.

**Spreadsheet came out empty after a run** — See the `--rebuild`
explanation above; you likely deleted the xlsx without also passing
`--rebuild`.

**Browser doesn't open / login fails** — Double check
`secrets/credentials.json` exists and is the file you downloaded from
Google Cloud Console (see Step 2), and that your Google account was added
as a test user on the OAuth consent screen.

**"Gmail API has not been used in project..." error** — Go back to Google
Cloud Console and make sure you clicked **Enable** on the Gmail API
(Step 2, item 3).

---

## Project structure

```
Linkedin_Tracker/
├── tracker_venv/            # virtual environment (not committed to Git)
├── linkedin_tracker/        # the actual code: Gmail login, email parsing, matching
│   ├── gmail_client.py
│   ├── parser.py            # LinkedIn email parsing
│   ├── dice_parser.py       # Dice email parsing
│   └── sent_matcher.py
├── data/
│   ├── LinkedIn_Job_Tracker.xlsx              # the tracker you open/edit
│   └── LinkedIn_Job_Tracker_full_history_backup.xlsx
├── secrets/
│   ├── credentials.json     # your OAuth client (not committed to Git)
│   └── token.json           # created after first login (not committed to Git)
├── update_tracker.py        # the script you actually run
├── processed_ids.json       # the script's memory of which emails it already imported
└── requirements.txt         # list of Python packages this project needs
```

---

## Security & privacy notes

- `secrets/credentials.json` and `secrets/token.json` are personal to your
  Google account. Never commit them to a public repository or share them —
  `.gitignore` already excludes the whole `secrets/` folder for this reason.
- The app only ever requests **read-only** Gmail access
  (`gmail.readonly`) — it cannot send, delete, or modify any email.
- Your tracker spreadsheet may contain personal application data (company
  contacts, emails) — the `.gitignore` also excludes `data/*.xlsx` so it
  isn't accidentally committed to Git either.
