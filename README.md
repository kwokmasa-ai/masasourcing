# MasaSourcing Lead Finder

Finds people online who need a China sourcing agent or quality inspection, and
adds them to `leads.xlsx` for daily review.

It collects posts from Reddit and web results from Brave Search, filters them by
keyword, classifies the matches with Claude, and writes the leads to a
spreadsheet sorted by buyer quality × urgency.

**It never posts, comments, or messages anyone.** It only finds leads and drafts
replies. You send them yourself.

## Setup

Python 3.10+.

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env   # then fill in the keys below
```

### 1. Reddit API (PRAW)

1. Log in to Reddit and open https://www.reddit.com/prefs/apps.
2. Click **create another app**, choose type **script**, and set any redirect
   URI (for example `http://localhost:8080`).
3. Copy the string under the app name into `REDDIT_CLIENT_ID`, and the
   **secret** into `REDDIT_CLIENT_SECRET`.
4. Set `REDDIT_USER_AGENT` to something descriptive, such as
   `MasaSourcingLeadFinder/1.0 by your_reddit_username`.

### 2. Brave Search API

1. Sign up at https://brave.com/search/api/ and create a key for the
   **Search** plan.
2. Put it in `BRAVE_API_KEY`.

The free tier allows roughly one request per second. The script waits between
searches automatically.

### 3. Anthropic API

1. Create a key at https://console.anthropic.com/.
2. Put it in `ANTHROPIC_API_KEY`.

The model is set in `config.yaml` under `claude.model`.

### Contact details for web fetching

Set `web.user_agent` in `config.yaml` to a real contact address or URL. The
script identifies itself with that string when it fetches pages.

## Usage

```bash
python leadfinder.py                       # last 7 days, reddit + web
python leadfinder.py --days 2 --sources reddit
python leadfinder.py --dry-run             # print the top 10, write nothing
```

| Flag | Default | Meaning |
| --- | --- | --- |
| `--days N` | `days` in config (7) | Look-back window |
| `--sources` | `reddit,web` | Which sources to run |
| `--dry-run` | off | Print the top leads and skip all writes |
| `--config` | `config.yaml` | Alternate config file |

A dry run still calls the Claude API, so it costs the same as a real run.

## What gets written

`leads.xlsx` (path set by `output.xlsx_path`) with these columns: date found,
source, subreddit/site, title, URL, author, product, lead type, buyer quality,
urgency, summary, suggested reply, and status (default `new`).

- Rows are sorted by buyer quality × urgency, highest first.
- A URL already in the spreadsheet is never added twice.
- Your edits to the **Status** column are kept across runs.

`leadfinder.sqlite3` remembers every URL already classified, so the same post is
never sent to Claude twice, and non-leads stay out of the spreadsheet.

## Editing what it looks for

All lists live in `config.yaml`:

- `reddit.subreddits`: subreddits to scan.
- `web.queries`: search queries for the web source.
- `keywords`: the first-pass filter. An item must contain at least one of these
  words to be sent to Claude.
- `claude.max_items_per_run`: caps the number of Claude calls per run.

## Daily schedule (cron)

Run it every morning at 07:30. Use `--days 2` so consecutive runs overlap. The
database prevents duplicates, so the overlap costs nothing.

```bash
crontab -e
```

```cron
30 7 * * * cd /path/to/masasourcing && .venv/bin/python leadfinder.py --days 2 >> /path/to/masasourcing/cron.log 2>&1
```

Check `cron.log` if the spreadsheet stops updating. Cron does not load `.env`
on its own, so the script reads it from the project folder.

## Tests

```bash
python -m unittest discover -s tests -v
```

The tests cover keyword matching, parsing of Claude's replies, the dedupe
database, and spreadsheet sorting and status preservation. They do not call any
live API.

## Files

- `leadfinder.py`: the script.
- `config.yaml`: lists and settings.
- `.env` (not committed): API keys. Copy from `.env.example`.
- `leads.xlsx`, `leadfinder.sqlite3`, `cron.log` (not committed): output.
