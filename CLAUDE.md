# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A CLI tool that reads the user's Gmail via the Gmail API (read-only scope), finds LinkedIn and Dice job-application confirmation emails, and writes them into an Excel tracker (`data/Job_Tracker.xlsx`). It also cross-references the user's Sent folder to guess hiring-manager contacts for each application, can separately log "reach-out" contacts (people emailed on a company domain with no matching application), and can look up recruiter phone numbers from their Inbox replies via regex (no LLM/AI involved).

There is a **second, independent entry point** — `update_directory.py` — which builds a recruiter/contact address book (`data/Recruiters.xlsx`) by scanning Sent (To/Cc) and Inbox (From) mail directly. It does not read `Job_Tracker.xlsx` at all and keeps its own state file (`directory_processed_ids.json`). See "The recruiter directory" section below before touching it.

## Running it

```bash
# Windows Command Prompt
tracker_venv\Scripts\python update_tracker.py [flags]

# Git Bash
tracker_venv/Scripts/python update_tracker.py [flags]
```

There is no test suite, linter, or build step in this repo — verification is done via `--dry-run` against real Gmail data (see below).

First run requires `secrets/credentials.json` (a Google OAuth Desktop-app client, downloaded from Google Cloud Console) to exist; it opens a browser for OAuth consent and writes `secrets/token.json`. Both files are gitignored and personal to the user's Google account — never fabricate, log, or commit their contents.

### Key flags (see README.md for full details)
- `--dry-run` — parse and print without writing to the xlsx or `processed_ids.json`. Use this to validate any change to parsing/matching logic before a real run.
- `--rebuild` — deletes/recreates the tracker and ignores `processed_ids.json` for that run (otherwise already-processed message IDs are skipped forever). Combine with `--since` to limit scope.
- `--source {all|linkedin|dice|reachout}` — restrict which source(s) are fetched.
- `--since YYYY-MM-DD` — narrows both the Gmail query and what gets written.
- `--no-hiring-manager` — skips the Sent-mail indexing step entirely (faster).
- `--include-reachout` — also run the reach-out contact scan alongside whatever `--source` is fetching.
- `--include-phone` — look up recruiter phone numbers (regex against Inbox replies, no LLM) and backfill any existing row with a blank Contact Number. **Not** limited by `--since` — the backfill pass always scans the whole sheet regardless (see gotcha below).

When testing changes manually, prefer `--tracker data/test_copy.xlsx` (or `--dry-run`) over the real tracker file to avoid corrupting the user's actual data.

## Architecture

Single entry point: `update_tracker.py`. Everything else lives in `src/` as focused, single-responsibility modules with no shared state beyond what's passed explicitly (imported as `from src.gmail_client import ...`, etc. — note the root project folder is `Linkedin_Tracker/`, a different name/case from the `src/` package, so there's no ambiguity):

- `gmail_client.py` — OAuth flow and `get_gmail_service()`, the only place that touches `secrets/`.
- `parser.py` — parses LinkedIn's HTML email body (via BeautifulSoup) to extract company, title, applied date, and job link. Company is extracted from the *subject* line, not the body.
- `dice_parser.py` — parses Dice's plain-text subject line (no HTML body worth parsing; no job link is ever available for Dice).
- `sent_matcher.py` — builds an index of every To/Cc recipient across the user's Sent mail (`fetch_sent_index`), then does two independent things with it:
  - `find_hiring_managers`: fuzzy-matches a company name against recipient domains on the *same calendar date* as the application (word-overlap scoring against `GENERIC_DOMAINS`-filtered, non-personal email domains).
  - `find_reachout_contacts`: every distinct non-generic-domain recipient *not* already known (i.e., not already a hiring-manager match), one row per unique email using the earliest date seen.
- `phone_lookup.py` — `find_phone_for_email` searches the user's **Inbox** (not Sent mail — a recruiter's phone can only appear in something *they* sent) for `from:<email>`, pulls plain text via `get_text_body` (in `parser.py`), and runs `extract_phone` (pure regex, no LLM) to find a US-format phone number + optional extension. Returns `None` if no messages exist from that sender or none contain a recognizable number — this is the common case, since most recruiters never reply.

`update_tracker.py` orchestrates: fetch messages per source (LinkedIn/Dice queries are Gmail search strings) → parse → resolve applied date (parsed date, falling back to message internal timestamp) → optionally match hiring managers → merge with existing spreadsheet rows → re-sort by date → rewrite the whole sheet.

### Data flow / state model

Two persistent files gate behavior across runs:

1. **`processed_ids.json`** (gitignored) — the set of Gmail message IDs already imported. This, not the spreadsheet, is the source of truth for "already handled." Deleting the xlsx without also passing `--rebuild` produces an empty sheet, because every source email is still marked processed and gets skipped. `--rebuild` clears this set for the run and rewrites it at the end to match whatever's in the sheet.
2. **`data/Job_Tracker.xlsx`** — columns are fixed (see `HEADERS` in `update_tracker.py`). Every run reads existing rows (`read_existing_rows`), merges in new rows, sorts the combined set by date, and rewrites rows 2..N in place — manual edits (comments, contact numbers, hand-added rows) are always preserved except under `--rebuild`. `LEGACY_PLATFORM_VALUES` silently migrates old `'Yes'` platform values (from when LinkedIn was the only source) to `'LinkedIn'` on every read.

`--dry-run` mirrors the same read/merge logic (including reading existing rows for non-rebuild runs, so previews accurately reflect what a real run would add) but never opens the workbook for writing and never touches `processed_ids.json`.

### ⚠️ Gotcha: `--since` scope is not what it looks like

`--since` narrows the Gmail **fetch** query for LinkedIn/Dice/reach-out (i.e. what counts as "new"), and separately filters reach-out contacts by date. It does **not** limit the `--include-phone` backfill pass, which always scans every existing row regardless.

This means: **never drop `--since` "just to widen" a phone backfill** — the backfill doesn't need it. Omitting `--since` entirely re-fetches the user's **entire** Gmail history for LinkedIn/Dice/reach-out. If `processed_ids.json` has ever been narrowed by an earlier `--rebuild --since <date>` run (which resets the processed set to just that run's window), old applications from before that date are no longer marked "processed" and will silently reappear as "new" on the next unscoped run — producing duplicate rows if any of that data is still in the sheet from before. This happened once in this project's history and required a manual dedup pass (by job link for LinkedIn/Dice rows, by exact email for reach-out rows) to recover. Before running any command without `--since` against the user's real tracker, confirm that's actually intended, and consider a defensive copy of the xlsx first.

### Adding a new source

Follow the `SOURCES` list pattern in `update_tracker.py`: a Gmail search query, a `parse_*` function returning `{company, title, applied_date, job_link}`, and a `source_value` for the Platform column. The rest of the pipeline (dedup via `processed_ids.json`, hiring-manager matching, merge/sort/write) is source-agnostic.

## The recruiter directory (`update_directory.py`)

Separate script, separate file, separate state (`directory_processed_ids.json`) — does not read `Job_Tracker.xlsx`. Scans Sent (To/Cc) and Inbox (From) mail for real human contacts on company domains, filtering out personal-mail providers, job-board/ATS domains, and automated senders (both by header signal — `List-Unsubscribe`, bulk `Precedence`, `Auto-Submitted` — and by local-part pattern, see `BOT_LOCAL_PART_RE` in `src/recruiter_directory.py`). The user's own email address (fetched via `getProfile`) is explicitly excluded from every scan, so CC'ing themselves never adds them as their own contact.

Unlike the main tracker, this is **merge-only, never rebuilt on a normal run** — `merge_contact()` fills blank cells and unions the Positions column (subject lines, `Re:`/`Fwd:`-stripped) but never overwrites a non-blank value. Manually added rows (even ones missing an email) are always preserved. Only `--rebuild` wipes the file.

Job Title, Phone, and LinkedIn Profile URL come from `scan_contact_signature()` in `src/phone_lookup.py` — one Inbox fetch per contact extracts all three (quoted reply text stripped first via `strip_quoted_text`). None of the three is a *verified* match to the contact specifically — each is just the first regex match found in their own text, so a shared mailbox or a message that mentions someone else's number/profile can attribute the wrong one.

### ⚠️ Gotcha: default `--max-results` (2000) is too low for a full-history run

Confirmed empirically in this project's history: a real account had 2,148 Sent + 3,626 Inbox messages — both exceed the default. Passing `--max-results` too low **silently truncates** the scan (no error, just fewer messages processed) rather than failing loudly. Before a first full-history `update_directory.py` run (no `--since`), check actual mailbox size and pass an explicit `--max-results` well above it (e.g. `fetch_all_messages(service, 'in:sent', 10000)` / `'in:inbox'` to get a real count first).
