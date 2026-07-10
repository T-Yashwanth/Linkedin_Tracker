# Job Application Tracker

This tool reads your Gmail inbox, finds the application-confirmation emails
sent by **LinkedIn** and **Dice** every time you apply to a job, and
automatically fills an Excel spreadsheet with the company name, job title,
date, platform, and (for LinkedIn) a link back to the job posting. If it can
also find an email you personally sent to someone at that company on the
same day, it fills in that person's name and email address too.

Optionally, it can also log every recruiter/contact you've personally
emailed on a real company domain that *isn't* already tied to one of those
applications — e.g. someone you reached out to about a role you heard about
from a LinkedIn post, a friend, or anywhere else. See
[Reach-out contact tracking](#reach-out-contact-tracking) below.

It can also try to fill in a recruiter's phone number, by searching for
replies they've actually sent you and pattern-matching a phone number out
of the plain text (no AI/LLM involved — just regex). See
[Phone number extraction](#phone-number-extraction) below.

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
3. If `data/Job_Tracker.xlsx` doesn't exist yet, it's created
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
how it behaves. You can combine several at once — order doesn't matter.

Quick reference:

| Flag | Takes a value? | Default | What it does |
|---|---|---|---|
| `--dry-run` | no | off | Preview only, writes nothing |
| `--since` | yes, `YYYY-MM-DD` | none (all history) | Only include applications on/after this date |
| `--rebuild` | no | off | Start the spreadsheet fresh, ignore `processed_ids.json` for this run |
| `--no-hiring-manager` | no | off | Skip the Sent-mail lookup (faster, leaves those columns blank) |
| `--tracker` | yes, a file path | `data/Job_Tracker.xlsx` | Which spreadsheet file to read/write |
| `--max-results` | yes, a number | `2000` | Cap on how many Gmail messages to scan |
| `--source` | yes, `all`/`linkedin`/`dice`/`reachout` | `all` | Only fetch from one platform (or only reach-out contacts) |
| `--include-reachout` | no | off | Also log reach-out contacts alongside LinkedIn/Dice, in the same run |
| `--include-phone` | no | off | Look up recruiter phone numbers and backfill blank Contact Number cells |

Full explanation of each, with examples, below.

> ⚠️ **`--since` does not limit `--include-phone`'s backfill.** Filling in
> phone numbers for your *existing* rows always scans every row with a
> blank Contact Number, regardless of `--since` — there's no need to drop
> `--since` to "widen" the backfill. `--since` only controls which *new*
> LinkedIn/Dice applications get fetched. Dropping `--since` entirely
> re-fetches your **entire** Gmail history for LinkedIn/Dice/reach-out too —
> and if `processed_ids.json` has ever been narrowed by an earlier
> `--rebuild --since <date>` run, old applications from before that date can
> silently reappear as "new" and get re-imported as duplicates. If you just
> want phone numbers backfilled, keep whatever `--since` you'd normally use
> (or drop it entirely only if you deliberately want a full history refetch).

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

> ⚠️ **If you delete `data/Job_Tracker.xlsx` but leave
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
`data/Job_Tracker.xlsx`. Handy for testing on a throwaway copy
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

### `--source {all|linkedin|dice|reachout}`
Restrict this run to a single platform instead of scanning both LinkedIn and
Dice. Only changes what's *fetched* — your existing rows from other
platforms are never touched or removed. Defaults to `all` (LinkedIn + Dice,
but **not** reach-out contacts — that one is opt-in only, see below).
```bash
tracker_venv/Scripts/python update_tracker.py --source dice
tracker_venv/Scripts/python update_tracker.py --source linkedin
tracker_venv/Scripts/python update_tracker.py --source reachout
```
`--source reachout` runs *only* the reach-out contact scan (see next
section) — it skips the LinkedIn/Dice email search entirely.

### `--include-reachout`
Adds the reach-out contact scan on top of whatever `--source` you chose
(default `all`, so normally this means "LinkedIn + Dice + reach-out
contacts, all in one run"). Off by default — you have to opt in. See
[Reach-out contact tracking](#reach-out-contact-tracking) for the full
explanation of what this does and its limitations before turning it on.
```bash
tracker_venv/Scripts/python update_tracker.py --include-reachout
```

### `--include-phone`
Looks up recruiter phone numbers and fills the "Contact Number For Job
Post" column — both for contacts found in this run, and by backfilling any
existing row that has a matched recruiter email but a blank Contact
Number. Off by default. See
[Phone number extraction](#phone-number-extraction) below for the full
explanation, including real limitations worth reading before turning it on.
```bash
tracker_venv/Scripts/python update_tracker.py --include-phone
```

### Common combinations

**Day-to-day use — no flags needed:**
```bash
# Pulls in whatever's new from both LinkedIn and Dice since last time,
# merges with existing rows, re-sorts by date. This is what you'll run most often.
tracker_venv/Scripts/python update_tracker.py
```

**Previewing before you trust a run:**
```bash
# See exactly what a normal run would add, without touching the spreadsheet
tracker_venv/Scripts/python update_tracker.py --dry-run

# Preview just the Dice side of things
tracker_venv/Scripts/python update_tracker.py --source dice --dry-run

# Preview a full rebuild before committing to it
tracker_venv/Scripts/python update_tracker.py --since 2026-06-24 --rebuild --dry-run
```

**Rebuilding from scratch:**
```bash
# Full rebuild, entire Gmail history, both platforms
tracker_venv/Scripts/python update_tracker.py --rebuild

# Rebuild, but only recent applications (faster, smaller sheet)
tracker_venv/Scripts/python update_tracker.py --since 2026-06-24 --rebuild

# Rebuild only the Dice rows' worth of data — e.g. after fixing a Dice
# parsing bug, without forcing LinkedIn to be re-scanned too
tracker_venv/Scripts/python update_tracker.py --source dice --since 2026-06-24 --rebuild
```
Note: `--rebuild` always deletes/recreates the *entire* spreadsheet file
first (see the `--rebuild` section above), so combining it with `--source
dice` still wipes LinkedIn rows too — they just won't get re-added in that
run. Use this combo only if you plan to run `--source linkedin --rebuild`
right after, or don't mind re-running the other source separately.

**Working with just one platform:**
```bash
# Only sync new Dice applications, leave LinkedIn alone
tracker_venv/Scripts/python update_tracker.py --source dice

# Only sync new LinkedIn applications, leave Dice alone
tracker_venv/Scripts/python update_tracker.py --source linkedin

# Preview only recent Dice applications, skip hiring-manager lookup for speed
tracker_venv/Scripts/python update_tracker.py --source dice --since 2026-06-29 --no-hiring-manager --dry-run
```

**Faster, lighter runs:**
```bash
# Skip the Sent-mail hiring-manager lookup entirely (much faster)
tracker_venv/Scripts/python update_tracker.py --no-hiring-manager

# Combine with --source and --since for the fastest possible sync
tracker_venv/Scripts/python update_tracker.py --source dice --since 2026-07-01 --no-hiring-manager
```

**Testing safely on a throwaway file:**
```bash
# Try a full rebuild against a test copy instead of your real tracker
tracker_venv/Scripts/python update_tracker.py --tracker data/test_copy.xlsx --rebuild

# Test only the Dice source against that same throwaway file
tracker_venv/Scripts/python update_tracker.py --tracker data/test_copy.xlsx --source dice
```

**Reach-out contact tracking:**
```bash
# Preview only, only recent contacts, no LinkedIn/Dice fetching
tracker_venv/Scripts/python update_tracker.py --source reachout --since 2026-06-24 --dry-run

# Write those reach-out contacts for real
tracker_venv/Scripts/python update_tracker.py --source reachout --since 2026-06-24

# One run that does everything: LinkedIn + Dice applications AND reach-out contacts
tracker_venv/Scripts/python update_tracker.py --include-reachout

# Same, but scoped to recent activity only
tracker_venv/Scripts/python update_tracker.py --include-reachout --since 2026-06-24
```

**Phone number lookup:**
```bash
# Preview only: backfill existing blank Contact Numbers + look up phones
# for any new contacts found this run. --since here only limits what NEW
# LinkedIn/Dice applications get fetched -- the backfill itself always
# covers every existing row regardless.
tracker_venv/Scripts/python update_tracker.py --include-phone --since 2026-06-24 --dry-run

# Write it for real
tracker_venv/Scripts/python update_tracker.py --include-phone --since 2026-06-24

# Cheapest way to test just the phone lookup against a small, known set of
# contacts before trusting it more broadly
tracker_venv/Scripts/python update_tracker.py --source reachout --since 2026-06-24 --include-phone --dry-run

# Full run: LinkedIn + Dice + reach-out + phone lookup/backfill, all at once
tracker_venv/Scripts/python update_tracker.py --include-reachout --include-phone --since 2026-06-24
```

---

## Reach-out contact tracking

Beyond application-confirmation emails, you might also personally email
recruiters about roles you heard about from a LinkedIn post, a friend, a
referral, or anywhere else — with no automated confirmation email to detect.
This feature logs those contacts too, using your **Sent** mail as the
source instead of an inbox confirmation.

**How it decides what counts as a reach-out contact:**
- It looks at everyone you've sent an email **To** or **Cc'd** on a real
  company domain (personal providers like Gmail/Yahoo/Outlook are always
  excluded, same as the hiring-manager matching for LinkedIn/Dice).
- It skips anyone whose email address is **already** in the spreadsheet's
  "Company Email" column — whether that's from an existing LinkedIn/Dice
  hiring-manager match, a previous reach-out run, or something you typed in
  by hand. This is how re-runs avoid creating duplicate rows for people
  you've already logged.
- Each remaining, unique email address becomes **one row**, using the
  earliest date you emailed them (even if you emailed them again on a later
  date, it won't create a second row).

**What gets filled in for these rows:**
- **Platform** — always `Email Reach Out` (distinct from `LinkedIn`/`Dice`,
  so you can filter/sort to see just these).
- **Hiring Manager In Linkedin** / **Company Email** — the contact's name
  (from the display name on your email, or guessed from their address) and
  email address.
- **Company Full Name** — a rough guess from their email domain (e.g.
  `gravityitresources.com` → "Gravityitresources"). This is **not**
  reliable — it's just a starting point. Expect to clean these up by hand.
- **Title**, **Website Applied** — always blank; there's no job posting
  tied to a reach-out contact.
- **Date** — the day you first emailed them, not necessarily the day you
  learned about the role.

**Known limitations — read before turning this on:**
- It has no way to know the role, or how you found out about it (LinkedIn
  post, friend, referral, etc.) — that part is always on you to fill in
  manually in the Comment Section.
- It can't tell a recruiter apart from any other professional contact on a
  company domain — vendors, colleagues, anyone. The domain filter cuts out
  personal email providers, but not false positives within real companies.
- High-volume staffing firms can generate a lot of rows — every individual
  recruiter you've ever emailed at, say, TekSystems becomes its own row.
- This is newer and less tested than the LinkedIn/Dice matching. Run it
  with `--dry-run` first and review the output before writing it for real.

---

## Phone number extraction

`--include-phone` fills the "Contact Number For Job Post" column where
possible, for both hiring-manager and reach-out contacts. **No AI/LLM is
involved anywhere in this** — it's plain regular-expression pattern
matching against real text pulled from Gmail.

**How it works:**
- For each contact's email address, it searches your **Inbox** (not Sent
  mail) for `from:<email>` — since a recruiter's phone number can only
  appear in something *they* actually sent you, not in your own outgoing
  messages.
- It reads the plain text of up to 3 of their most recent messages to you,
  and runs two regex patterns against it: one for common US phone formats
  (`(415) 555-0132`, `415-555-0132`, `415.555.0132`, `+1 415 555 0132`,
  `+14155550132`), and one for an extension marker right after it (`ext`,
  `ext.`, `extension`, or `x`, e.g. `x204`).
- A bare 10-digit run with **no** separators at all (no `()`, `-`, `.`,
  space, or leading `+`) is deliberately rejected — those are more often
  tracking/order numbers than phone numbers in a signature block.
- **Quoted reply history is stripped first, and your own number is never
  recorded.** Recruiter replies usually quote your original email — with
  your signature and your phone number — underneath their message. The
  extractor cuts the body at the first quoted-message marker ("On ...
  wrote:", "-----Original Message-----", `>`-prefixed lines, etc.), and it
  also auto-detects your own number(s) from your Sent-mail signature and
  excludes them from every match, so a reply that contains no recruiter
  phone yields a blank instead of *your* number.
- The first valid match found is used, formatted as e.g.
  `"(415) 555-0132 x204"` in a single cell (no separate extension column).

**Backfill behavior:** unlike the rest of the pipeline, this doesn't only
apply to new rows — it also scans every *existing* row with a matched
recruiter email but a blank Contact Number, and fills it in if a phone can
be found. This only ever writes into cells that are currently blank; a
number you've typed in by hand is never touched or overwritten. Note this
backfill is **not** limited by `--since` — it always covers your entire
sheet (see the warning under "Understanding the flags" above).

**Known limitations:**
- **Most recruiters never reply**, so most lookups will find nothing —
  that's expected, not a bug. The run summary reports how many unique
  contacts were checked vs. how many phone numbers were actually found.
- **Signatures that are images are not read at all.** If a phone number
  only exists as pixels in an embedded signature graphic (common with some
  signature-generator tools), there's no text for the regex to match — this
  would require OCR, which isn't implemented.
- Adds one extra Gmail search (plus up to 3 message fetches) per unique
  contact email. The first run with `--include-phone` on a large tracker
  can take several minutes; later runs are much faster since most cells
  will already be filled in.

---

## Recruiter directory (`update_directory.py`)

A **separate script and a separate file** (`data/Recruiters.xlsx`) from
the main tracker — a standalone address book of every recruiter/contact
you've exchanged email with, grouped by company domain. It does not read
`Job_Tracker.xlsx` at all; it scans Gmail directly and keeps its own state
file, `directory_processed_ids.json`, exactly like `processed_ids.json`
does for the main tracker.

```bash
# Command Prompt
tracker_venv\Scripts\python update_directory.py --since 2026-06-29 --dry-run

# Git Bash
tracker_venv/Scripts/python update_directory.py --since 2026-06-29 --dry-run
```

### What it does

- Scans your **Sent** mail (To/Cc recipients) and your **Inbox** (senders)
  for both directions of contact — someone you emailed, or someone who
  emailed you (including interview/meeting invitations).
- Filters out anything that isn't a real, domain-specific human contact:
  personal mail providers (Gmail, Yahoo, etc.), job-board/ATS domains
  (Dice, Indeed, LinkedIn, Greenhouse, Workday, etc.), and automated
  senders — detected both by header (`List-Unsubscribe`, bulk
  `Precedence`, `Auto-Submitted`) and by local-part patterns
  (`no-reply`, `notifications`, `newsletter`, and similar).
- Columns: `S.no | Company Name | Job Title | Recruiter Name | Phone |
  Email | LinkedIn Profile | Positions`. Rows are sorted by the domain of
  each email address (derived from Email, not a separate column) so every
  contact at the same company sits together. **Email is mandatory and
  unique** — one row per address, never duplicated.
- **Positions** are the subject lines of every email exchanged with that
  contact (in either direction), cleaned of repeated `Re:`/`Fwd:`
  prefixes and `;`-joined.
- Job Title, Phone, and LinkedIn Profile are filled in the same way as
  `--include-phone`: a best-effort regex scan of the contact's own Inbox
  replies (quoted text stripped, your own phone number/LinkedIn profile
  excluded automatically).

### Unlike the main tracker — this one preserves your edits

This is the opposite of `--rebuild`-style behavior: **the directory is
never regenerated from scratch on a normal run.** Existing rows are only
ever merged into — a blank cell gets filled in if new data is found, but
anything you've already typed in (or a value the script already filled)
is never overwritten. Positions accumulate (union, not replace) rather
than being reset. Manually added rows, even ones with no email at all,
are always kept. Only `--rebuild` wipes the file and starts over.

### Flags

Mirrors the main tracker's style: `--since YYYY-MM-DD` (narrows both the
Sent and Inbox scan), `--dry-run` (discovers and prints contacts grouped
by domain — **no signature scans, no file writes**, so it's fast),
`--rebuild` (wipes `directory_processed_ids.json` and the xlsx, starts
fresh), `--max-results N`, `--directory <path>`.

```bash
# Safe first look — no writes, no slow signature scans
tracker_venv/Scripts/python update_directory.py --since 2026-06-29 --dry-run

# Build/update it for real, scoped to a recent window
tracker_venv/Scripts/python update_directory.py --since 2026-06-29

# Full history (first run establishes the whole directory; can take a
# while since almost every contact needs a signature scan the first time)
tracker_venv/Scripts/python update_directory.py
```

### Known limitations

- Same phone/LinkedIn/title extraction caveats as `--include-phone`
  above: most contacts never replied (so most cells stay blank), image
  signatures aren't read, and best-effort regex matching can occasionally
  pick up the wrong thing — e.g. a LinkedIn link the sender shared for
  someone *else* (like a candidate's profile) can get attributed to them
  instead. Treat these three columns as a helpful starting point, not
  ground truth.
- Automated-sender filtering is heuristic; if something automated slips
  through, or a real contact is wrongly excluded, it's a filter to refine
  in `src/directory.py` (`is_eligible_contact`, `BOT_LOCAL_PART_RE`), not
  something to fix by hand-editing rows.
- Close `Recruiters.xlsx` in Excel before running — same file-lock
  behavior as the main tracker.

---

## What each spreadsheet column means

| Column | Filled automatically? | Notes |
|---|---|---|
| S.no | Yes | Renumbered every run, in date order |
| Date | Yes | Application rows: the date you applied. Reach-out rows: the date you first emailed that contact |
| Title | LinkedIn/Dice only | Job title. Always blank for reach-out rows (no job tied to them) |
| Company Full Name | Yes | Application rows: company name from the email. Reach-out rows: a rough guess from the contact's email domain — verify/correct by hand |
| Platform | Yes | `LinkedIn`, `Dice`, or `Email Reach Out` |
| Website Applied | LinkedIn only | Shortened link straight to the job posting. Always blank for Dice and reach-out rows (no job-specific link available) |
| Hiring Manager In Linkedin | Sometimes | Application rows: only if a matching Sent email was found on the same date. If you emailed multiple people at that company that day (To **or** Cc), all of them are listed here, separated by `; `. Reach-out rows: always filled — the contact's name |
| Company Email | Sometimes | Application rows: same matches as above, `; `-separated. Reach-out rows: always filled — the contact's email |
| Contact Number For Job Post | Sometimes (opt-in) | Blank unless you run with `--include-phone` — see [Phone number extraction](#phone-number-extraction). Otherwise left blank for you to fill in |
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

**Too many irrelevant reach-out rows (`--include-reachout` / `--source
reachout`)** — This scan can't distinguish recruiters from any other
company-domain contact you've emailed. Delete the rows you don't want; that
alone stops them from reappearing, since already-present "Company Email"
values are skipped on future runs. If it's consistently noisy, don't use
`--include-reachout` on your regular runs — only run `--source reachout`
occasionally and review with `--dry-run` first.

---

## Project structure

```
Linkedin_Tracker/
├── tracker_venv/            # virtual environment (not committed to Git)
├── src/                     # the actual code: Gmail login, email parsing, matching
│   ├── gmail_client.py
│   ├── parser.py            # LinkedIn email parsing
│   ├── dice_parser.py       # Dice email parsing
│   ├── sent_matcher.py      # hiring-manager & reach-out contact matching
│   ├── phone_lookup.py      # recruiter phone/title/LinkedIn signature scan
│   └── directory.py         # recruiter directory aggregation + sheet I/O
├── data/
│   ├── Job_Tracker.xlsx              # the tracker you open/edit
│   ├── Recruiters.xlsx                # the recruiter directory you open/edit
│   └── LinkedIn_Job_Tracker_full_history_backup.xlsx
├── secrets/
│   ├── credentials.json     # your OAuth client (not committed to Git)
│   └── token.json           # created after first login (not committed to Git)
├── update_tracker.py        # main tracker script you run
├── update_directory.py      # recruiter directory script you run
├── processed_ids.json       # main tracker's memory of which emails it already imported
├── directory_processed_ids.json  # directory's own memory (separate from the above)
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
