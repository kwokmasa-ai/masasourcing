#!/usr/bin/env python3
"""MasaSourcing lead finder.

Finds people online who need a China sourcing agent or quality inspection,
classifies them with Claude, and appends the leads to leads.xlsx.

Pipeline:
  1. Collect posts and comments from Reddit (PRAW) and web results (Brave Search).
  2. Keyword pre-filter from config.yaml.
  3. Skip anything already seen (SQLite) and anything older than the window.
  4. Classify each remaining item with the Claude API.
  5. Append leads to leads.xlsx, sorted by buyer_quality x urgency.

The bot never posts, comments, or messages anyone. It only drafts replies.

Usage:
  python leadfinder.py --days 7 --sources reddit,web
  python leadfinder.py --dry-run
"""

import argparse
import json
import logging
import os
import re
import sqlite3
import sys
import time
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from urllib import robotparser
from urllib.parse import urlparse

import anthropic
import openpyxl
import praw
import requests
import yaml
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from openpyxl.styles import Font

HERE = os.path.dirname(os.path.abspath(__file__))
log = logging.getLogger("leadfinder")

LEAD_TYPES = {"inspection", "sourcing", "freight", "other"}
XLSX_HEADERS = [
    "Date found", "Source", "Subreddit/site", "Title", "URL", "Author",
    "Product", "Lead type", "Buyer quality", "Urgency", "Summary",
    "Suggested reply", "Status",
]
# Column positions (0-based) used for sorting and status preservation.
COL_URL = XLSX_HEADERS.index("URL")
COL_BUYER = XLSX_HEADERS.index("Buyer quality")
COL_URGENCY = XLSX_HEADERS.index("Urgency")
COL_STATUS = XLSX_HEADERS.index("Status")

SYSTEM_PROMPT = """You screen online posts for MasaSourcing, a sourcing and quality-inspection company based in Guangzhou, China, run by Masa Kwok. Services: finding and vetting factories, factory visits and verification, pre-shipment inspections with photo and video, sample checks, and help with freight.

A LEAD is an established business (Amazon or Shopify seller, importer, brand, wholesaler) that has a supplier problem: overpaying, an unreliable supplier, a quality issue, a suspected scam, or no one on the ground in China. The author should be asking for help, or describing a problem that MasaSourcing could solve.

NOT A LEAD:
- Someone selling sourcing, inspection, or agent services (a competitor).
- Dropshipping beginners with no budget, or hobbyists.
- Job seekers looking for work in China.
- Posts that only mention a keyword in passing and show no real need.

Score the lead:
- buyer_quality (1-5): 5 = already sells or has distribution with real volume; 1 = no evident business.
- urgency (1-5): 5 = active problem right now (shipment in transit, scam in progress, deadline); 1 = vague future interest.

Write suggested_reply as a helpful first reply, not salesy: answer the question or offer a concrete next step, and mention MasaSourcing only lightly. Keep it under 80 words. The owner sends it manually.

Respond with ONLY a JSON object, no prose, with exactly these keys:
{"is_lead": true|false, "lead_type": "inspection"|"sourcing"|"freight"|"other", "buyer_quality": 1-5, "urgency": 1-5, "product": "short product name or empty string", "summary": "one line", "suggested_reply": "under 80 words"}"""


@dataclass
class Item:
    source: str       # "reddit" or "web"
    site: str         # "r/Alibaba" or the web domain
    title: str
    text: str
    url: str
    author: str
    created: datetime  # timezone-aware UTC


# ---------------------------------------------------------------- config

def load_config(path):
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def resolve_path(p):
    return p if os.path.isabs(p) else os.path.join(HERE, p)


def normalize_url(url):
    return url.strip().rstrip("/")


# ---------------------------------------------------------------- sources: Reddit

def fetch_reddit(cfg, since):
    """Newest posts and comments from each subreddit, inside the window."""
    reddit = praw.Reddit(
        client_id=os.environ["REDDIT_CLIENT_ID"],
        client_secret=os.environ["REDDIT_CLIENT_SECRET"],
        user_agent=os.environ["REDDIT_USER_AGENT"],
    )
    limit = cfg["reddit"]["limit_per_subreddit"]
    items = []
    for name in cfg["reddit"]["subreddits"]:
        sub = reddit.subreddit(name)
        site = f"r/{name}"
        try:
            for post in sub.new(limit=limit):
                created = datetime.fromtimestamp(post.created_utc, timezone.utc)
                if created < since:
                    break  # listing is newest first
                items.append(Item("reddit", site, post.title, post.selftext or "",
                                  f"https://www.reddit.com{post.permalink}",
                                  str(post.author), created))
            for comment in sub.comments(limit=limit):
                created = datetime.fromtimestamp(comment.created_utc, timezone.utc)
                if created < since:
                    break
                items.append(Item("reddit", site, f"Comment in {site}", comment.body or "",
                                  f"https://www.reddit.com{comment.permalink}",
                                  str(comment.author), created))
        except Exception as exc:  # one bad subreddit should not stop the run
            log.warning("Reddit %s failed: %s", site, exc)
    return items


# ---------------------------------------------------------------- sources: web

def freshness_for(days):
    """Map the window to Brave's freshness codes."""
    if days <= 1:
        return "pd"
    if days <= 7:
        return "pw"
    if days <= 31:
        return "pm"
    return "py"


def brave_search(query, count, freshness, api_key, timeout):
    resp = requests.get(
        "https://api.search.brave.com/res/v1/web/search",
        params={"q": query, "count": count, "freshness": freshness},
        headers={"X-Subscription-Token": api_key, "Accept": "application/json"},
        timeout=timeout,
    )
    resp.raise_for_status()
    results = resp.json().get("web", {}).get("results", [])
    return [
        {
            "url": r.get("url", ""),
            "title": r.get("title", ""),
            "description": r.get("description", ""),
            "page_age": r.get("page_age"),
        }
        for r in results
        if r.get("url")
    ]


class PoliteFetcher:
    """Fetches pages while respecting robots.txt and a per-domain rate limit."""

    def __init__(self, user_agent, delay_seconds, timeout, max_chars):
        self.user_agent = user_agent
        self.delay = delay_seconds
        self.timeout = timeout
        self.max_chars = max_chars
        self._robots = {}
        self._last_hit = {}

    def _robots_for(self, base):
        if base in self._robots:
            return self._robots[base]
        rp = robotparser.RobotFileParser()
        try:
            resp = requests.get(f"{base}/robots.txt", headers={"User-Agent": self.user_agent},
                                timeout=self.timeout)
            if resp.status_code in (401, 403):
                rp.disallow_all = True  # RFC 9309: treat as fully disallowed
            elif resp.status_code >= 400:
                rp.allow_all = True     # no robots.txt means no restrictions
            else:
                rp.parse(resp.text.splitlines())
        except requests.RequestException:
            rp.disallow_all = True      # unknown means skip, to be safe
        self._robots[base] = rp
        return rp

    def allowed(self, url):
        p = urlparse(url)
        base = f"{p.scheme}://{p.netloc}"
        return self._robots_for(base).can_fetch(self.user_agent, url)

    def fetch_text(self, url):
        """Return (title, main_text) or None if the page cannot be used."""
        p = urlparse(url)
        if p.scheme not in ("http", "https"):
            return None
        wait = self.delay - (time.monotonic() - self._last_hit.get(p.netloc, 0))
        if wait > 0:
            time.sleep(wait)
        self._last_hit[p.netloc] = time.monotonic()
        try:
            resp = requests.get(url, headers={"User-Agent": self.user_agent}, timeout=self.timeout)
        except requests.RequestException as exc:
            log.debug("fetch failed %s: %s", url, exc)
            return None
        if resp.status_code != 200 or "html" not in resp.headers.get("Content-Type", ""):
            return None
        soup = BeautifulSoup(resp.text, "html.parser")
        title = soup.title.get_text(strip=True) if soup.title else ""
        for tag in soup(["script", "style", "nav", "header", "footer", "aside", "form", "noscript"]):
            tag.decompose()
        text = re.sub(r"\s+", " ", soup.get_text(" ", strip=True))
        return title, text[: self.max_chars]


def parse_page_age(value):
    if not value:
        return None
    try:
        dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None
    return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)


def fetch_web(cfg, since, days, api_key):
    web = cfg["web"]
    fetcher = PoliteFetcher(web["user_agent"], web["rate_limit_seconds"],
                            web["request_timeout_seconds"], web["max_text_chars"])
    freshness = freshness_for(days)
    seen_urls = set()
    items = []
    for query in web["queries"]:
        try:
            results = brave_search(query, web["results_per_query"], freshness, api_key,
                                   web["request_timeout_seconds"])
        except requests.RequestException as exc:
            log.warning("Search failed for %r: %s", query, exc)
            continue
        time.sleep(1.1)  # Brave's free tier allows about one request per second
        for res in results:
            url = normalize_url(res["url"])
            if url in seen_urls:
                continue
            seen_urls.add(url)
            created = parse_page_age(res.get("page_age")) or datetime.now(timezone.utc)
            if created < since:
                continue
            if not fetcher.allowed(url):
                log.info("robots.txt disallows %s, skipping", url)
                continue
            page = fetcher.fetch_text(url)
            if not page:
                continue
            page_title, text = page
            items.append(Item("web", urlparse(url).netloc, page_title or res["title"],
                              text or res["description"], url, "", created))
    return items


# ---------------------------------------------------------------- keyword pass

def keyword_hits(text, keywords):
    hits = []
    for kw in keywords:
        if re.search(r"\b" + re.escape(kw) + r"\b", text, re.IGNORECASE):
            hits.append(kw)
    return hits


# ---------------------------------------------------------------- Claude pass

def _as_bool(value):
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() == "true"


def _as_score(value):
    try:
        return min(5, max(1, int(value)))
    except (TypeError, ValueError):
        return 1


def parse_classification(raw_text):
    """Pull the JSON object out of the model's reply. Returns None if unusable."""
    match = re.search(r"\{.*\}", raw_text, re.DOTALL)
    if not match:
        return None
    try:
        data = json.loads(match.group(0))
    except json.JSONDecodeError:
        return None
    lead_type = str(data.get("lead_type", "other")).strip().lower()
    return {
        "is_lead": _as_bool(data.get("is_lead", False)),
        "lead_type": lead_type if lead_type in LEAD_TYPES else "other",
        "buyer_quality": _as_score(data.get("buyer_quality")),
        "urgency": _as_score(data.get("urgency")),
        "product": str(data.get("product") or ""),
        "summary": str(data.get("summary") or ""),
        "suggested_reply": str(data.get("suggested_reply") or ""),
    }


def classify(client, model, max_tokens, item):
    user_msg = (
        f"Source: {item.source} ({item.site})\n"
        f"Title: {item.title}\n"
        f"URL: {item.url}\n\n"
        f"Text:\n{item.text[:3000]}"
    )
    try:
        resp = client.messages.create(
            model=model,
            max_tokens=max_tokens,
            system=[{"type": "text", "text": SYSTEM_PROMPT,
                     "cache_control": {"type": "ephemeral"}}],
            messages=[{"role": "user", "content": user_msg}],
        )
    except anthropic.APIError as exc:
        log.warning("Claude call failed for %s: %s", item.url, exc)
        return None
    text = "".join(block.text for block in resp.content if block.type == "text")
    result = parse_classification(text)
    if result is None:
        log.warning("Unparseable classification for %s", item.url)
    return result


# ---------------------------------------------------------------- SQLite dedupe

def open_db(path):
    conn = sqlite3.connect(path)
    conn.execute(
        "CREATE TABLE IF NOT EXISTS seen ("
        " url TEXT PRIMARY KEY, first_seen TEXT NOT NULL, is_lead INTEGER NOT NULL)"
    )
    conn.commit()
    return conn


def is_seen(conn, url):
    return conn.execute("SELECT 1 FROM seen WHERE url = ?", (url,)).fetchone() is not None


def mark_seen(conn, url, is_lead):
    conn.execute(
        "INSERT OR REPLACE INTO seen (url, first_seen, is_lead) VALUES (?, ?, ?)",
        (url, datetime.now(timezone.utc).isoformat(), int(is_lead)),
    )
    conn.commit()


# ---------------------------------------------------------------- XLSX output

def lead_row(item, result, found_on):
    return [
        found_on, item.source, item.site, item.title, item.url, item.author,
        result["product"], result["lead_type"], result["buyer_quality"],
        result["urgency"], result["summary"], result["suggested_reply"], "new",
    ]


def score_of(row):
    try:
        return int(row[COL_BUYER]) * int(row[COL_URGENCY])
    except (TypeError, ValueError):
        return 0


def append_leads(path, new_rows):
    """Add rows to the workbook, keep existing statuses, and sort by score."""
    if os.path.exists(path):
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        existing = [list(r) for r in ws.iter_rows(min_row=2, values_only=True) if r[COL_URL]]
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Leads"
        existing = []

    known_urls = {row[COL_URL] for row in existing}
    fresh = [r for r in new_rows if r[COL_URL] not in known_urls]
    all_rows = existing + fresh
    all_rows.sort(key=lambda r: (-score_of(r), str(r[0])))

    ws.delete_rows(1, ws.max_row)
    ws.append(XLSX_HEADERS)
    for row in all_rows:
        ws.append(row)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)
    return len(fresh)


# ---------------------------------------------------------------- main

def require_env(names):
    missing = [n for n in names if not os.environ.get(n)]
    return missing


def main(argv=None):
    load_dotenv(os.path.join(HERE, ".env"))
    parser = argparse.ArgumentParser(description="Find sourcing and inspection leads.")
    parser.add_argument("--days", type=int, help="look-back window in days (default: config)")
    parser.add_argument("--sources", default="reddit,web", help="comma list: reddit,web")
    parser.add_argument("--dry-run", action="store_true",
                        help="print the top leads and write nothing")
    parser.add_argument("--config", default=os.path.join(HERE, "config.yaml"))
    args = parser.parse_args(argv)

    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    cfg = load_config(args.config)
    days = args.days or cfg["days"]
    since = datetime.now(timezone.utc) - timedelta(days=days)
    sources = {s.strip() for s in args.sources.split(",") if s.strip()}

    anthropic_key = os.environ.get("ANTHROPIC_API_KEY")
    if not anthropic_key:
        log.error("ANTHROPIC_API_KEY is not set. Add it to .env.")
        return 1

    items = []
    if "reddit" in sources:
        missing = require_env(["REDDIT_CLIENT_ID", "REDDIT_CLIENT_SECRET", "REDDIT_USER_AGENT"])
        if missing:
            log.error("Skipping reddit, missing: %s", ", ".join(missing))
        else:
            items += fetch_reddit(cfg, since)
    if "web" in sources:
        brave_key = os.environ.get("BRAVE_API_KEY")
        if not brave_key:
            log.error("Skipping web, BRAVE_API_KEY is not set.")
        else:
            items += fetch_web(cfg, since, days, brave_key)
    log.info("Collected %d items from the last %d days", len(items), days)

    keywords = cfg["keywords"]
    candidates = []
    for item in items:
        if keyword_hits(f"{item.title}\n{item.text}", keywords):
            candidates.append(item)
    log.info("%d items passed the keyword filter", len(candidates))

    conn = open_db(resolve_path(cfg["output"]["db_path"]))
    unique = {}
    for item in candidates:
        url = normalize_url(item.url)
        if url not in unique and not is_seen(conn, url):
            unique[url] = item
    to_classify = list(unique.values())[: cfg["claude"]["max_items_per_run"]]
    log.info("%d new items to classify", len(to_classify))

    client = anthropic.Anthropic(api_key=anthropic_key)
    found_on = datetime.now().strftime("%Y-%m-%d")
    leads = []
    for item in to_classify:
        result = classify(client, cfg["claude"]["model"], cfg["claude"]["max_tokens"], item)
        if result is None:
            continue  # not marked seen, so a later run can retry it
        if not args.dry_run:
            mark_seen(conn, normalize_url(item.url), result["is_lead"])
        if result["is_lead"]:
            leads.append((result["buyer_quality"] * result["urgency"],
                          lead_row(item, result, found_on)))
    conn.close()

    leads.sort(key=lambda pair: -pair[0])
    if args.dry_run:
        top = leads[: cfg["output"]["top_n_dry_run"]]
        print(f"\nTop {len(top)} leads (dry run, nothing written):")
        for score, row in top:
            print(f"\n[score {score}] {row[COL_BUYER]}x{row[COL_URGENCY]}  {row[7]}  {row[2]}")
            print(f"  {row[3]}\n  {row[4]}\n  {row[10]}\n  Reply: {row[11]}")
        return 0

    added = 0
    if leads:
        added = append_leads(resolve_path(cfg["output"]["xlsx_path"]), [row for _, row in leads])
    log.info("Added %d leads to %s", added, cfg["output"]["xlsx_path"])
    return 0


if __name__ == "__main__":
    sys.exit(main())
