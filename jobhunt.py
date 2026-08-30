#!/usr/bin/env python3
"""
jobhunt.py v2 - a personal job search engine that runs itself.

New in v2:
  * Every link is opened and verified before it reaches your Excel file.
    Dead and closed postings are quarantined, not shown as applyable.
  * US-only filtering (roles outside the US are dropped).
  * Ghost-job risk scoring - flags listings that look like pipeline filler.
  * Far more detail per row: salary range, team, employment type,
    description snippet, how long it has been listed.

Commands
--------
  python jobhunt.py setup      Interview you, then write config.json
  python jobhunt.py run        Fetch -> score -> verify links -> write Excel
  python jobhunt.py add NAME   Add a company (auto-detects its ATS)
  python jobhunt.py check      Test every company token, prune dead ones
  python jobhunt.py schedule   Install the daily background run

Requires: pip install requests openpyxl

Location-filter audit fixtures (is_us_location):
  must return False: "Remote UK", "Remote - Europe", "Cambridge, UK",
                     "San Jose, Costa Rica", "London", "Toronto, ON",
                     "Remote EMEA", "Bengaluru"
  must return True:  "New York, NY", "Remote - US", "Springfield, MO",
                     "Boise, ID", "United States", "SF Bay Area"
"""

import argparse
import hashlib
import html
import json
import os
import platform
import re
import subprocess
import sys
import threading
import time
import urllib.parse
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone

try:
    import requests
except ImportError:
    sys.exit("Missing dependency. Run:  pip install requests openpyxl")

HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(HERE, "config.json")
STATE_PATH = os.path.join(HERE, "state.json")
LOG_PATH = os.path.join(HERE, "jobhunt.log")
LOCK_PATH = os.path.join(HERE, "jobhunt.lock")

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/122.0 Safari/537.36"}
TIMEOUT = 25
TODAY = datetime.now().strftime("%Y-%m-%d")


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------

def log(msg):
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    print(line, flush=True)
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except OSError:
        pass


def strip_html(text):
    if not text:
        return ""
    text = re.sub(r"<br\s*/?>", " ", text, flags=re.I)
    text = re.sub(r"</(p|li|div|h\d)>", " ", text, flags=re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    text = html.unescape(text)
    return re.sub(r"\s+", " ", text).strip()


def clean_html(text):
    """Unescape THEN strip, for APIs that return HTML-escaped HTML.

    Greenhouse's job["content"] arrives as "&lt;p&gt;...&lt;/p&gt;" - running
    strip_html alone finds no tags to remove and raw HTML leaks into Excel
    (the audit counted 198 such snippets). html.unescape() first turns it back
    into real markup so strip_html can flatten it to plain text.
    """
    if not text:
        return ""
    return strip_html(html.unescape(text))


def to_date(value):
    if value in (None, "", 0):
        return ""
    try:
        if isinstance(value, (int, float)):
            ts = float(value)
            if ts > 1e11:
                ts /= 1000.0
            return datetime.fromtimestamp(ts, tz=timezone.utc).strftime("%Y-%m-%d")
        s = str(value).replace("Z", "+00:00")
        return datetime.fromisoformat(s).strftime("%Y-%m-%d")
    except (ValueError, OSError, OverflowError):
        return ""


def days_since(datestr):
    if not datestr:
        return ""
    try:
        return (datetime.now() - datetime.strptime(datestr, "%Y-%m-%d")).days
    except ValueError:
        return ""


def job_key(company, title, url):
    return hashlib.md5(f"{company}|{title}|{url}".lower().encode()).hexdigest()[:12]


def ask(prompt, default=""):
    suffix = f" [{default}]" if default else ""
    try:
        return input(f"{prompt}{suffix}: ").strip() or default
    except EOFError:
        return default


def ask_yes(prompt, default=True):
    answer = ask(f"{prompt} ({'Y/n' if default else 'y/N'})").lower()
    return default if not answer else answer.startswith("y")


def ask_list(prompt, default_list):
    print(f"\n{prompt}")
    print(f"  default: {', '.join(default_list)}")
    raw = ask("  Comma-separated values, or Enter to keep default")
    return [x.strip() for x in raw.split(",") if x.strip()] if raw else default_list


# --------------------------------------------------------------------------
# US location filtering
# --------------------------------------------------------------------------

US_STATES = {
    "alabama": "AL", "alaska": "AK", "arizona": "AZ", "arkansas": "AR",
    "california": "CA", "colorado": "CO", "connecticut": "CT", "delaware": "DE",
    "florida": "FL", "georgia": "GA", "hawaii": "HI", "idaho": "ID",
    "illinois": "IL", "indiana": "IN", "iowa": "IA", "kansas": "KS",
    "kentucky": "KY", "louisiana": "LA", "maine": "ME", "maryland": "MD",
    "massachusetts": "MA", "michigan": "MI", "minnesota": "MN", "mississippi": "MS",
    "missouri": "MO", "montana": "MT", "nebraska": "NE", "nevada": "NV",
    "new hampshire": "NH", "new jersey": "NJ", "new mexico": "NM", "new york": "NY",
    "north carolina": "NC", "north dakota": "ND", "ohio": "OH", "oklahoma": "OK",
    "oregon": "OR", "pennsylvania": "PA", "rhode island": "RI",
    "south carolina": "SC", "south dakota": "SD", "tennessee": "TN", "texas": "TX",
    "utah": "UT", "vermont": "VT", "virginia": "VA", "washington": "WA",
    "west virginia": "WV", "wisconsin": "WI", "wyoming": "WY",
    "district of columbia": "DC",
}
US_ABBREV = set(US_STATES.values())

US_CITIES = {
    "new york", "nyc", "brooklyn", "manhattan", "queens", "san francisco",
    "los angeles", "chicago", "boston", "seattle", "austin", "denver", "atlanta",
    "miami", "dallas", "houston", "philadelphia", "phoenix", "san diego",
    "san jose", "palo alto", "mountain view", "sunnyvale", "menlo park",
    "cupertino", "redmond", "bellevue", "portland", "minneapolis", "detroit",
    "nashville", "charlotte", "raleigh", "pittsburgh", "washington dc", "arlington",
    "cambridge", "jersey city", "hoboken", "santa monica", "san mateo", "oakland",
    "boulder", "salt lake city", "columbus", "indianapolis", "kansas city",
    "st. louis", "tampa", "orlando", "las vegas", "sacramento", "irvine",
}

NON_US = {
    "india", "canada", "united kingdom", "london", "ireland", "dublin",
    "germany", "berlin", "munich", "france", "paris", "spain", "madrid",
    "barcelona", "netherlands", "amsterdam", "poland", "warsaw", "krakow",
    "australia", "sydney", "melbourne", "singapore", "japan", "tokyo",
    "china", "shanghai", "beijing", "hong kong", "korea", "seoul", "israel",
    "tel aviv", "brazil", "sao paulo", "mexico", "argentina",
    "toronto", "vancouver", "montreal", "ottawa", "waterloo", "ontario",
    "bangalore", "bengaluru", "hyderabad", "mumbai", "delhi", "gurgaon",
    "noida", "pune", "chennai", "sweden", "stockholm", "denmark", "copenhagen",
    "norway", "oslo", "switzerland", "zurich", "italy", "milan", "rome",
    "portugal", "lisbon", "philippines", "manila", "vietnam", "indonesia",
    "new zealand", "south africa", "nigeria", "kenya", "egypt", "dubai",
    "emea", "apac", "latam", "emirates", "belgium", "brussels", "austria",
    "czech", "prague", "romania", "bucharest", "hungary", "budapest", "turkey",
    "uk", "u.k.", "europe", "european", "costa rica", "england", "scotland",
    "wales", "britain",
}


def _term_regex(terms):
    """Compile a phrase set into one word-bounded pattern for lowercased text.

    Lookarounds instead of \\b so terms ending in punctuation ("u.k.") still
    bound correctly; internal spaces match any whitespace run. Word bounding
    is the point: the old substring checks let "india" match "Indiana" and
    "san jose" match "San Jose, Costa Rica".
    """
    parts = sorted((r"\s+".join(re.escape(w) for w in t.split()) for t in terms),
                   key=len, reverse=True)
    return re.compile(r"(?<![a-z0-9])(?:" + "|".join(parts) + r")(?![a-z0-9])")


_NON_US_RE = _term_regex(NON_US)
_US_STATE_RE = _term_regex(US_STATES)
_US_CITY_RE = _term_regex(US_CITIES | {"sf", "bay area", "sf bay area"})
_US_WORD_RE = re.compile(r"(?<![a-z0-9])(?:usa|u\.s\.a\.?|u\.s\.?|us)(?![a-z0-9])")
_STATE_ABBREV_RE = re.compile(r"\b([A-Z]{2})\b")


def _classify_segment(segment):
    """Classify ONE location segment: True (US), False (foreign), None."""
    seg = segment.strip()
    if not seg:
        return None
    low = seg.lower()

    # 1) Explicit foreign country/city/region markers win immediately, so a
    #    US-city lookalike later in the string can never override them
    #    ("San Jose, Costa Rica", "Cambridge, UK").
    if _NON_US_RE.search(low):
        return False

    # 2) Explicit US markers. State abbreviations are matched case-sensitively
    #    on the ORIGINAL string BEFORE any lowercasing - "Boise, ID" is a
    #    state, lowercase "id"/"in"/"or" inside prose is not.
    for m in _STATE_ABBREV_RE.finditer(seg):
        if m.group(1) in US_ABBREV:
            return True
    if "united states" in low or _US_WORD_RE.search(low):
        return True
    if _US_STATE_RE.search(low) or _US_CITY_RE.search(low):
        return True
    # A remote mention with nothing foreign in the segment (foreign was
    # checked first) is presumed US - most tracked boards are US companies.
    if re.search(r"\bremote\b", low):
        return True

    # 3) No evidence either way.
    return None


def is_us_location(loc):
    """
    True if the location looks US-based, False if clearly not, None if unknown.

    Multi-location strings are split on | ; / (and " or "): if ANY segment is
    US the posting counts as US; if every segment is explicitly foreign it is
    False; anything else is None (unknown). Audit fixtures for both directions
    are listed in the module docstring.
    """
    if not loc:
        return None
    verdicts = [_classify_segment(s)
                for s in re.split(r"[|;/]|\s+or\s+", str(loc)) if s.strip()]
    if not verdicts:
        return None
    if any(v is True for v in verdicts):
        return True
    if all(v is False for v in verdicts):
        return False
    return None


# --------------------------------------------------------------------------
# Link verification - the anti-dead-link layer
# --------------------------------------------------------------------------

DEAD_PAGE_PHRASES = [
    "no longer accepting applications", "this job is no longer",
    "this position is no longer", "position has been filled",
    "role has been filled", "job posting is closed",
    "posting is no longer available", "this job is closed",
    "applications are closed", "job not found", "position not found",
    "page not found", "opening is closed", "this opening is no longer",
    "sorry, this job", "posting has expired", "this role is closed",
]


def verify_link(url):
    """
    Open the posting and decide whether it is really applyable.
    Returns: LIVE, DEAD, CLOSED, BLOCKED, or UNREACHABLE
    """
    if not url or not url.startswith("http"):
        return "DEAD"
    try:
        r = requests.get(url, headers=UA, timeout=15, allow_redirects=True)
    except requests.exceptions.RequestException:
        return "UNREACHABLE"

    if r.status_code in (404, 410):
        return "DEAD"
    if r.status_code in (401, 403, 429):
        return "BLOCKED"
    if r.status_code != 200:
        return "UNREACHABLE"

    body = r.text[:120000].lower()
    for phrase in DEAD_PAGE_PHRASES:
        if phrase in body:
            return "CLOSED"
    if len(body.strip()) < 400:
        return "CLOSED"
    return "LIVE"


def verify_all(rows, workers=24):
    total = len(rows)
    done = [0]

    def work(row):
        status = verify_link(row["Apply Link"])
        done[0] += 1
        if done[0] % 25 == 0 or done[0] == total:
            log(f"  verified {done[0]}/{total}")
        return status

    with ThreadPoolExecutor(max_workers=workers) as pool:
        statuses = list(pool.map(work, rows))
    for row, status in zip(rows, statuses):
        row["Link Status"] = status

    _mark_dead_hosts(rows)
    return rows


# How many links from one host we need to see before a run of failures is
# evidence about the host rather than about our own rate limiting.
HOST_MIN_SAMPLE = 5
HOST_FAIL_RATIO = 0.9


def _host_of(url):
    try:
        return urllib.parse.urlsplit(url).netloc.lower()
    except (ValueError, AttributeError):
        return ""


def _mark_dead_hosts(rows):
    """Re-label UNREACHABLE rows when their whole host is down.

    A single UNREACHABLE means our own checker timed out or got rate-limited,
    so those rows are kept -- deleting them silently loses real jobs. But when
    every link on one host fails, that is not our rate limit, it is their site.
    Shipping 122 lifeattiktok.com links that all answer 503 as "ready to apply"
    is the same "looks clean, is not" failure in the other direction.

    These rows stay in the workbook (the jobs are real, the host is just down)
    but move off the apply list into their own sheet.
    """
    seen, failed = {}, {}
    for row in rows:
        host = _host_of(row.get("Apply Link", ""))
        if not host:
            continue
        seen[host] = seen.get(host, 0) + 1
        if row["Link Status"] in ("UNREACHABLE", "DEAD", "CLOSED"):
            failed[host] = failed.get(host, 0) + 1

    down = {h for h, n in seen.items()
            if n >= HOST_MIN_SAMPLE and failed.get(h, 0) / n >= HOST_FAIL_RATIO}
    if not down:
        return

    hit = 0
    for row in rows:
        if (row["Link Status"] == "UNREACHABLE"
                and _host_of(row.get("Apply Link", "")) in down):
            row["Link Status"] = "SITE DOWN"
            note = row.get("Risk Notes") or ""
            row["Risk Notes"] = (note + "; " if note else "") + \
                "whole careers site unreachable during this run - retry later"
            hit += 1
    for host in sorted(down):
        log(f"  ! {host} failed {failed[host]}/{seen[host]} link checks "
            f"- treating as site-wide outage, not our rate limit")
    log(f"  {hit} roles moved off the apply list into 'Needs Link Check'")


# --------------------------------------------------------------------------
# Ghost job risk
# --------------------------------------------------------------------------

def ghost_risk(row, first_seen_date):
    """
    How likely is this pipeline filler rather than a real opening?
    Signals from ATS outcome research: long-lived listings, no salary band.
    """
    reasons, points = [], 0

    age = row.get("Days Old")
    if isinstance(age, int):
        if age >= 90:
            points += 3
            reasons.append("posted 90+ days ago")
        elif age >= 60:
            points += 2
            reasons.append("posted 60+ days ago")
        elif age >= 45:
            points += 1
            reasons.append("posted 45+ days ago")

    listed = days_since(first_seen_date)
    if isinstance(listed, int) and listed >= 45:
        points += 2
        reasons.append(f"on your list {listed} days, still open")

    if not row.get("Salary Range"):
        points += 1
        reasons.append("no salary posted")

    if points >= 4:
        return "HIGH", "; ".join(reasons)
    if points >= 2:
        return "MEDIUM", "; ".join(reasons)
    if points == 1:
        return "LOW", "; ".join(reasons)
    return "FRESH", "recent, salary published"


# --------------------------------------------------------------------------
# Starter company pack
# --------------------------------------------------------------------------

STARTER_COMPANIES = [
    {"name": "Stripe", "ats": "greenhouse", "token": "stripe"},
    {"name": "Databricks", "ats": "greenhouse", "token": "databricks"},
    {"name": "Ramp", "ats": "greenhouse", "token": "ramp"},
    {"name": "Robinhood", "ats": "greenhouse", "token": "robinhood"},
    {"name": "DoorDash", "ats": "greenhouse", "token": "doordash"},
    {"name": "Instacart", "ats": "greenhouse", "token": "instacart"},
    {"name": "Affirm", "ats": "greenhouse", "token": "affirm"},
    {"name": "Samsara", "ats": "greenhouse", "token": "samsara"},
    {"name": "Datadog", "ats": "greenhouse", "token": "datadog"},
    {"name": "Cloudflare", "ats": "greenhouse", "token": "cloudflare"},
    {"name": "Peloton", "ats": "greenhouse", "token": "peloton"},
    {"name": "Betterment", "ats": "greenhouse", "token": "betterment"},
    {"name": "Oscar Health", "ats": "greenhouse", "token": "oscarhealth"},
    {"name": "Squarespace", "ats": "greenhouse", "token": "squarespace"},
    {"name": "Etsy", "ats": "greenhouse", "token": "etsy"},
    {"name": "Warby Parker", "ats": "greenhouse", "token": "warbyparker"},
    {"name": "Flatiron Health", "ats": "greenhouse", "token": "flatironhealth"},
    {"name": "MongoDB", "ats": "greenhouse", "token": "mongodb"},
    {"name": "Yext", "ats": "greenhouse", "token": "yext"},
    {"name": "Chainalysis", "ats": "greenhouse", "token": "chainalysis"},
    {"name": "Attentive", "ats": "greenhouse", "token": "attentive"},
    {"name": "Justworks", "ats": "greenhouse", "token": "justworks"},
    {"name": "SeatGeek", "ats": "greenhouse", "token": "seatgeek"},
    {"name": "Palantir", "ats": "lever", "token": "palantir"},
    {"name": "Plaid", "ats": "lever", "token": "plaid"},
    {"name": "Netflix", "ats": "lever", "token": "netflix"},
    {"name": "Spotify", "ats": "lever", "token": "spotify"},
    {"name": "OpenAI", "ats": "ashby", "token": "openai"},
    {"name": "Anthropic", "ats": "ashby", "token": "anthropic"},
    {"name": "Linear", "ats": "ashby", "token": "linear"},
    {"name": "Vanta", "ats": "ashby", "token": "vanta"},
    {"name": "Clay", "ats": "ashby", "token": "clay"},
    {"name": "Deel", "ats": "ashby", "token": "deel"},
]


# --------------------------------------------------------------------------
# Salary parsing
# --------------------------------------------------------------------------

def parse_salary_range(*texts):
    """Return (min, max, display) parsed from any of the given strings."""
    for text in texts:
        if not text:
            continue
        nums = []
        for m in re.finditer(
                r"\$\s?(\d{1,3}(?:,\d{3})+|\d{2,3}(?:\.\d)?\s?[kK]\b|\d{5,6})", str(text)):
            raw = m.group(1).replace(",", "").strip()
            try:
                val = int(float(raw[:-1].strip()) * 1000) if raw.lower().endswith("k") \
                    else int(float(raw))
            except ValueError:
                continue
            if 30000 <= val <= 900000:
                nums.append(val)
        if nums:
            lo, hi = min(nums), max(nums)
            return lo, hi, (f"${lo:,} - ${hi:,}" if hi > lo else f"${lo:,}")
    return 0, 0, ""


# --------------------------------------------------------------------------
# ATS fetchers
# --------------------------------------------------------------------------

# Shared requests.Session installed by cmd_run so the parallel board fetch
# reuses connections. When unset (tests, cmd_check), plain requests is used.
_SESSION = None


def _get(url):
    client = _SESSION if _SESSION is not None else requests
    r = client.get(url, headers=UA, timeout=TIMEOUT)
    r.raise_for_status()
    return r.json()


def _base(company, title, location, url, posted, desc, source,
          team="", emp_type="", salary_text=""):
    lo, hi, disp = parse_salary_range(salary_text, desc)
    return {"company": company, "title": title, "location": location, "url": url,
            "posted": posted, "description": desc, "source": source, "team": team,
            "emp_type": emp_type, "salary_min": lo, "salary_max": hi,
            "salary_display": disp, "sponsorship": ""}


def fetch_greenhouse(name, token):
    url = f"https://boards-api.greenhouse.io/v1/boards/{token}/jobs?content=true"
    out = []
    for j in _get(url).get("jobs", []):
        depts = j.get("departments") or []
        pay = ""
        for md in (j.get("metadata") or []):
            label = str(md.get("name", "")).lower()
            if "salary" in label or "pay" in label or "compensation" in label:
                pay = str(md.get("value", ""))
        out.append(_base(name, j.get("title", ""),
                         (j.get("location") or {}).get("name", ""),
                         j.get("absolute_url", ""),
                         to_date(j.get("first_published") or j.get("updated_at")),
                         clean_html(j.get("content", ""))[:8000], "Greenhouse",
                         depts[0].get("name", "") if depts else "", "", pay))
    return out


def fetch_lever(name, token):
    url = f"https://api.lever.co/v0/postings/{token}?mode=json"
    out = []
    for j in _get(url):
        cat = j.get("categories") or {}
        sr = j.get("salaryRange") or {}
        pay = f"${sr.get('min','')} - ${sr.get('max','')}" if sr.get("min") else ""
        out.append(_base(name, j.get("text", ""), cat.get("location", ""),
                         j.get("hostedUrl", ""), to_date(j.get("createdAt")),
                         strip_html(j.get("descriptionPlain") or j.get("description", ""))[:8000],
                         "Lever", cat.get("team") or cat.get("department", ""),
                         cat.get("commitment", ""), pay))
    return out


def fetch_ashby(name, token):
    url = f"https://api.ashbyhq.com/posting-api/job-board/{token}?includeCompensation=true"
    out = []
    for j in _get(url).get("jobs", []):
        comp = j.get("compensation") or {}
        out.append(_base(name, j.get("title", ""), j.get("location", ""),
                         j.get("jobUrl", ""), to_date(j.get("publishedAt")),
                         strip_html(j.get("descriptionPlain") or j.get("descriptionHtml", ""))[:8000],
                         "Ashby", j.get("department", ""), j.get("employmentType", ""),
                         comp.get("compensationTierSummary") or ""))
    return out


def fetch_smartrecruiters(name, token):
    url = f"https://api.smartrecruiters.com/v1/companies/{token}/postings?limit=100"
    out = []
    for j in _get(url).get("content", []):
        loc = j.get("location") or {}
        city = ", ".join(x for x in [loc.get("city"), loc.get("region"), loc.get("country")] if x)
        out.append(_base(name, j.get("name", ""), city,
                         j.get("applyUrl") or
                         f"https://jobs.smartrecruiters.com/{token}/{j.get('id','')}",
                         to_date(j.get("releasedDate")), "", "SmartRecruiters",
                         (j.get("department") or {}).get("label", ""),
                         (j.get("typeOfEmployment") or {}).get("label", ""), ""))
    return out


def fetch_workable(name, token):
    url = f"https://apply.workable.com/api/v1/widget/accounts/{token}?details=true"
    out = []
    for j in _get(url).get("jobs", []):
        out.append(_base(name, j.get("title", ""),
                         ", ".join(x for x in [j.get("city"), j.get("state"),
                                               j.get("country")] if x),
                         j.get("url", ""), to_date(j.get("published_on")),
                         strip_html(j.get("description", ""))[:8000], "Workable",
                         j.get("department", ""), j.get("type", ""), ""))
    return out


ATS_FETCHERS = {
    "greenhouse": fetch_greenhouse, "lever": fetch_lever, "ashby": fetch_ashby,
    "smartrecruiters": fetch_smartrecruiters, "workable": fetch_workable,
}


def detect_ats(token):
    for ats, fn in ATS_FETCHERS.items():
        try:
            jobs = fn(token, token)
            if jobs:
                return ats, len(jobs)
        except Exception:
            continue
    return None, 0


SIMPLIFY_FEEDS = {
    "Simplify-Internships": "https://raw.githubusercontent.com/SimplifyJobs/"
                            "Summer2027-Internships/dev/.github/scripts/listings.json",
    "Simplify-NewGrad": "https://raw.githubusercontent.com/SimplifyJobs/"
                        "New-Grad-Positions/dev/.github/scripts/listings.json",
}


def fetch_simplify():
    """Fetch both Simplify feeds. Returns (rows, outcomes).

    Each feed gets up to 3 attempts (backoff 2s/5s). A feed that still fails
    is reported loudly and recorded as FAILED in its outcome entry so the run
    can be flagged INCOMPLETE - these two feeds carry most of the coverage,
    and a silent skip used to look identical to a quiet day.
    """
    out, outcomes = [], []
    for label, url in SIMPLIFY_FEEDS.items():
        data, err = None, ""
        for attempt, wait in enumerate((0, 2, 5), start=1):
            if wait:
                time.sleep(wait)
            try:
                data = _get(url)
                break
            except Exception as e:
                err = f"{type(e).__name__}: {e}"[:200]
                if attempt < 3:
                    log(f"  ! {label} attempt {attempt}/3 failed "
                        f"({type(e).__name__}) - retrying...")
        if data is None:
            log(f"! SOURCE FAILED: {label} — this run is INCOMPLETE "
                f"(rerun to recover)")
            outcomes.append({"name": label, "ats": "Simplify",
                             "status": "FAILED", "count": 0, "error": err})
            continue
        n = 0
        for j in data:
            if not j.get("active") or not j.get("is_visible"):
                continue
            rec = _base(j.get("company_name", ""), j.get("title", ""),
                        "; ".join(j.get("locations") or []), j.get("url", ""),
                        to_date(j.get("date_posted")),
                        " ".join(j.get("terms") or []) + " " +
                        " ".join(j.get("degrees") or []), label, "", "", "")
            rec["sponsorship"] = j.get("sponsorship", "")
            out.append(rec)
            n += 1
        log(f"  {label}: {n} active postings")
        outcomes.append({"name": label, "ats": "Simplify",
                         "status": "OK" if n else "EMPTY",
                         "count": n, "error": ""})
    return out, outcomes


# --------------------------------------------------------------------------
# Scoring
# --------------------------------------------------------------------------

SENIOR_BLOCK = ["senior", "sr.", "sr ", "staff", "principal", "lead ", "director",
                "head of", "vp ", "vice president", "group product", "chief",
                "distinguished", "architect", " iii", " iv"]

# Word-bounded on the TITLE. Plain substring checks turned "Product Manager |
# International" into an Internship (bonus + senior-filter bypass) and matched
# "apm" inside unrelated words; "international", "internal", "internet" must
# never count as intern.
INTERN_TITLE_RE = re.compile(r"\bintern(?:ship)?s?\b|\bco-?op\b")
NEWGRAD_TITLE_RE = re.compile(
    r"\bnew\s+grad(?:uate)?s?\b|\bgraduates?\b|\bassociate\s+product\b|\bapm\b|"
    r"\brotational\b|\bentry[\s-]+level\b|\buniversity\b|\bearly\s+career\b")

VISA_GREEN_FLAGS = ["will sponsor", "visa sponsorship available", "offers sponsorship",
                    "sponsor visas", "h-1b sponsorship", "sponsorship is available",
                    "we sponsor", "sponsor h-1b", "sponsors h-1b", "sponsor h1b"]
# Negation patterns checked FIRST inside any visa-topic sentence; one negated
# sentence beats any number of positive-sounding ones. "Unable to provide
# H-1B sponsorship" must classify as NO_SPONSORSHIP even though "h-1b
# sponsorship" alone looks like a green flag.
VISA_NEGATIONS = ["unable to", "cannot", "can't", "not able to", "no sponsorship",
                  "without sponsorship", "will not", "won't", "not offer",
                  "not provide", "not available", "not eligible",
                  "do not sponsor", "does not sponsor", "no visa sponsorship"]
# Hard requirements that block regardless of sponsorship wording.
VISA_CITIZEN_FLAGS = [
    "us citizen", "u.s. citizen", "citizenship is required", "security clearance",
    "green card holder", "permanent resident only",
    "must be authorized to work in the united states without",
]
_VISA_TOPIC_RE = re.compile(r"visa|sponsor|h[- ]?1\s?b", re.I)
_VISA_OPT_RE = re.compile(r"\b(?:OPT|CPT|F-?1)\b")  # case-sensitive on purpose


def classify_visa_signal(text):
    """Return "NO_SPONSORSHIP", "SPONSORS", or "" from posting text.

    Sentence-level: only sentences mentioning visa/sponsorship/H-1B/OPT/CPT
    are considered, negation is checked before positive wording, and any
    negated sentence wins over any positive one.
    """
    if not text:
        return ""
    positive = False
    for sentence in re.split(r"(?<=[.!?;])\s+|\n+", text):
        if not (_VISA_TOPIC_RE.search(sentence) or _VISA_OPT_RE.search(sentence)):
            continue
        low = sentence.lower()
        if any(neg in low for neg in VISA_NEGATIONS):
            return "NO_SPONSORSHIP"
        if any(g in low for g in VISA_GREEN_FLAGS):
            positive = True
    if any(f in text.lower() for f in VISA_CITIZEN_FLAGS):
        return "NO_SPONSORSHIP"
    return "SPONSORS" if positive else ""


def score_job(job, cfg):
    title = (job.get("title") or "").lower()
    loc = (job.get("location") or "").lower()
    desc = (job.get("description") or "").lower()
    blob = f"{title} {desc}"
    reasons, score = [], 0

    for bad in cfg.get("exclude_title_words", []):
        if bad.lower() in title:
            return None

    us = is_us_location(job.get("location"))
    if us is False:
        return None
    if us is None and cfg.get("us_only", True) and job.get("location"):
        return None

    is_intern = bool(INTERN_TITLE_RE.search(title))
    is_newgrad = bool(NEWGRAD_TITLE_RE.search(title))

    if not cfg.get("want_internships", True) and is_intern:
        return None
    if not cfg.get("want_fulltime", True) and not is_intern:
        return None
    if not is_intern and not is_newgrad:
        for word in SENIOR_BLOCK:
            if word in title:
                return None

    title_hit = False
    for kw in cfg.get("target_titles", []):
        if kw.lower() in title:
            score += 40
            reasons.append(f"title match: {kw}")
            title_hit = True
            break
    if not title_hit:
        for kw in cfg.get("secondary_titles", []):
            if kw.lower() in title:
                score += 22
                reasons.append(f"related role: {kw}")
                title_hit = True
                break
    if not title_hit:
        return None

    if is_intern:
        level, add = "Internship", 18
        reasons.append("internship")
    elif is_newgrad:
        level, add = "New Grad / APM", 20
        reasons.append("new-grad level")
    else:
        level, add = "Mid-level", 5
    score += add

    loc_hit = False
    for want in cfg.get("locations", []):
        if want.lower() in loc:
            score += 18
            reasons.append(f"location: {want}")
            loc_hit = True
            break
    if not loc_hit and "remote" in loc:
        score += 10
        reasons.append("US remote")
        loc_hit = True
    if not loc_hit:
        if cfg.get("strict_location", False):
            return None
        score -= 12
        reasons.append("elsewhere in US")

    visa_flag = ""
    declared = (job.get("sponsorship") or "").lower()
    if "does not offer" in declared or "citizenship is required" in declared:
        return None
    if "offers sponsorship" in declared:
        score += 25
        visa_flag = "SPONSORS (declared)"
        reasons.append("declared sponsor")
    if not visa_flag:
        signal = classify_visa_signal(
            f"{job.get('title') or ''}. {job.get('description') or ''}")
        if signal == "NO_SPONSORSHIP":
            if cfg.get("drop_no_sponsorship", True):
                return None
            score -= 40
            visa_flag = "NO SPONSORSHIP"
        elif signal == "SPONSORS":
            score += 20
            visa_flag = "SPONSORS (in posting)"
            reasons.append("sponsorship mentioned")

    # Word-bounded skill matching: plain substrings credited "Java" inside
    # "JavaScript" and "RAG" inside "leverage". Multi-word skills keep their
    # internal spaces flexible.
    matched = []
    for s in cfg.get("skills", []):
        pat = r"\s+".join(re.escape(w) for w in s.split())
        if pat and re.search(rf"(?<![a-z0-9]){pat}(?![a-z0-9])", blob, re.IGNORECASE):
            matched.append(s)
    if matched:
        score += min(len(matched) * 4, 20)
        reasons.append("your skills: " + ", ".join(matched[:6]))

    smax = job.get("salary_max") or job.get("salary_min") or 0
    if smax:
        if smax >= cfg.get("salary_target", 140000):
            score += 12
            reasons.append("strong pay band (better H-1B lottery tier)")
        elif smax >= cfg.get("salary_floor", 90000):
            score += 6
        else:
            score -= 5

    age = days_since(job.get("posted"))
    if isinstance(age, int):
        if age <= 7:
            score += 14
            reasons.append("posted this week")
        elif age <= 21:
            score += 7
            reasons.append("posted recently")
        elif age > cfg.get("max_age_days", 45):
            return None

    return max(score, 0), "; ".join(reasons), visa_flag, level


# --------------------------------------------------------------------------
# Excel
# --------------------------------------------------------------------------

COLUMNS = ["AI Fit", "Score", "Status", "Company", "Title", "Level", "Location",
           "Salary Range", "Visa Signal", "Link Status", "Ghost Risk",
           "Posted", "Days Old", "Team", "Type", "Source", "AI Verdict",
           "Eligibility", "Why It Matched",
           "Risk Notes", "What The Role Says", "Apply Link", "Applied Date",
           "Notes", "Key"]

WIDTHS = [8, 7, 14, 20, 42, 15, 26, 20, 20, 12, 12, 11, 10, 20, 14, 18, 46,
          12, 46, 30, 60, 12, 13, 30, 12]

# Long-text columns, resolved by NAME. The old literal index set was written
# before "AI Fit"/"AI Verdict" were prepended to COLUMNS and had drifted onto
# the wrong fields; deriving it means reordering COLUMNS can never do that again.
WRAP_COLS = frozenset(COLUMNS.index(n) + 1 for n in
                      ("Title", "Location", "AI Verdict", "Why It Matched",
                       "Risk Notes", "What The Role Says"))


def col(name):
    """Excel column letter for a COLUMNS header, looked up by name.

    Every Dashboard formula goes through this. Hard-coded letters silently
    pointed at the neighbouring field once "AI Fit" was inserted at column A,
    so the whole Dashboard counted the wrong columns and read as zeros.
    """
    from openpyxl.utils import get_column_letter
    return get_column_letter(COLUMNS.index(name) + 1)

# Active "AI Fit" cutoff for the apply list. cmd_run overwrites this from
# config.json ("min_ai_fit") before writing the workbook; it lives at module
# level so write_excel can label the Dashboard without growing its signature.
MIN_AI_FIT = 55

# Freshness window in days for the "Fresh (48h)" sheet. cmd_run overwrites
# this from config.json ("fresh_days", default 2) before writing the workbook.
FRESH_DAYS = 2

# Set by cmd_run when the run begins, so the Run Report sheet can show both
# the start and finish timestamps without growing write_excel's signature.
RUN_STARTED = ""


def ai_fit_value(row):
    """Return a row's "AI Fit" as an int, or None if the ranker never scored it.

    A blank AI Fit means the AI never saw the row - a failed batch, or the API
    was down. That is UNKNOWN, not bad. Callers must keep such rows on the apply
    list; only a real integer below the cutoff is grounds for demotion. Dropping
    rows the ranker never scored is the same class of bug as quarantining
    UNREACHABLE links, which cost this project ~29% of its live roles once.
    """
    v = row.get("AI Fit")
    if isinstance(v, bool) or v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return int(v)
    try:
        return int(str(v).strip())
    except (TypeError, ValueError):
        return None


def load_prior_tracker(path):
    if not os.path.exists(path):
        return {}
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, data_only=True)
        ws = wb["All Matches"] if "All Matches" in wb.sheetnames else wb.active
        headers = [c.value for c in ws[1]]
        idx = {h: i for i, h in enumerate(headers) if h}
        if not all(k in idx for k in ("Key", "Status", "Applied Date", "Notes")):
            return {}
        prior = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            key = row[idx["Key"]]
            if key:
                prior[key] = {"Status": row[idx["Status"]] or "",
                              "Applied Date": row[idx["Applied Date"]] or "",
                              "Notes": row[idx["Notes"]] or ""}
        return prior
    except Exception as e:
        log(f"  ! could not read prior tracker ({e})")
        return {}


def write_excel(all_rows, ready, new_keys, quarantine, path, lowfit=(),
                sources=(), incomplete=False, needs_check=()):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import ColorScaleRule, CellIsRule
    from openpyxl.worksheet.datavalidation import DataValidation

    # Defensive enforcement of the demotion contract, so no caller can put an
    # ineligible or below-cutoff row on the apply sheets by mistake. A row in
    # `ready` with Eligibility "NO" (hard blocker stated in the posting) or a
    # real integer AI Fit under MIN_AI_FIT is moved to Low Fit here. A blank
    # AI Fit is UNKNOWN, never demoted - the ranker simply never saw the row.
    # cmd_run pre-splits identically, so in production this is a no-op.
    kept, demoted = [], []
    for r in ready:
        fit = ai_fit_value(r)
        if r.get("Eligibility") == "NO" or (fit is not None and fit < MIN_AI_FIT):
            demoted.append(r)
        else:
            kept.append(r)
    ready = kept
    lowfit = list(lowfit) + demoted

    wb = Workbook()
    head_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="1F3864")
    body = Font(name="Arial", size=10)
    link_font = Font(name="Arial", size=10, color="0563C1", underline="single")
    new_fill = PatternFill("solid", fgColor="FFF2CC")
    thin = Side(style="thin", color="D9D9D9")

    def build(ws, data, mark_new=False):
        ws.append(COLUMNS)
        for i, c in enumerate(ws[1], start=1):
            c.font, c.fill = head_font, head_fill
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(i)].width = WIDTHS[i - 1]
        ws.row_dimensions[1].height = 32

        for r in data:
            ws.append([r.get(c, "") for c in COLUMNS])
            row = ws.max_row
            is_new = mark_new and r["Key"] in new_keys
            for i in range(1, len(COLUMNS) + 1):
                cell = ws.cell(row=row, column=i)
                cell.font = body
                cell.border = Border(bottom=thin)
                cell.alignment = Alignment(vertical="top",
                                           wrap_text=(i in WRAP_COLS))
                if is_new:
                    cell.fill = new_fill
            lc = ws.cell(row=row, column=COLUMNS.index("Apply Link") + 1)
            if r.get("Apply Link"):
                lc.hyperlink = r["Apply Link"]
                lc.value = "Apply"
                lc.font = link_font
                lc.alignment = Alignment(horizontal="center")

        last = ws.max_row
        if last > 1:
            ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{last}"
            ws.conditional_formatting.add(f"A2:A{last}", ColorScaleRule(
                start_type="min", start_color="F8696B",
                mid_type="percentile", mid_value=50, mid_color="FFEB84",
                end_type="max", end_color="63BE7B"))
            g = get_column_letter(COLUMNS.index("Ghost Risk") + 1)
            ws.conditional_formatting.add(f"{g}2:{g}{last}", CellIsRule(
                operator="equal", formula=['"HIGH"'],
                fill=PatternFill("solid", fgColor="FFC7CE"),
                font=Font(name="Arial", size=10, color="9C0006")))
            ws.conditional_formatting.add(f"{g}2:{g}{last}", CellIsRule(
                operator="equal", formula=['"FRESH"'],
                fill=PatternFill("solid", fgColor="C6EFCE"),
                font=Font(name="Arial", size=10, color="006100")))
            k = get_column_letter(COLUMNS.index("Link Status") + 1)
            ws.conditional_formatting.add(f"{k}2:{k}{last}", CellIsRule(
                operator="equal", formula=['"LIVE"'],
                fill=PatternFill("solid", fgColor="C6EFCE"),
                font=Font(name="Arial", size=10, color="006100")))
            dv = DataValidation(type="list", allow_blank=True,
                                formula1='"To Review,Interested,Applied,Referral Ask,'
                                         'Interviewing,Offer,Rejected,Skip"')
            ws.add_data_validation(dv)
            dv.add(f"{get_column_letter(COLUMNS.index('Status') + 1)}2:{get_column_letter(COLUMNS.index('Status') + 1)}{last}")
        ws.freeze_panes = "D2"

    ws_ready = wb.active
    ws_ready.title = "Ready to Apply"
    build(ws_ready, ready, True)
    build(wb.create_sheet("New This Run"), [r for r in ready if r["Key"] in new_keys])
    build(wb.create_sheet("All Matches"), all_rows, True)

    if needs_check:
        ws_nc = wb.create_sheet("Needs Link Check")
        build(ws_nc, needs_check)
        hosts = sorted({_host_of(r.get("Apply Link", "")) for r in needs_check} - {""})
        ws_nc.cell(row=ws_nc.max_row + 2, column=1, value=(
            f"{len(needs_check)} roles whose careers site was completely "
            f"unreachable during this run ({', '.join(hosts)}). Every link on "
            f"that host failed, so this is their outage, not a bad match - the "
            f"jobs are likely real. They are kept off Ready to Apply so you do "
            f"not click dead links. Re-run later to recover them."))

    ws_low = wb.create_sheet("Low Fit")
    build(ws_low, lowfit)
    note = ws_low.cell(row=ws_low.max_row + 2, column=1, value=(
        (f"{len(lowfit)} roles scored below AI Fit {MIN_AI_FIT} and were kept off "
         f"Ready to Apply. Nothing was deleted - every one of them is still on "
         f"All Matches. ")
        if lowfit else
        (f"No roles scored below AI Fit {MIN_AI_FIT} this run. ")) + (
        "Rows the AI never scored (blank AI Fit) count as unknown and are never "
        "demoted - they stay on Ready to Apply."))
    note.font = Font(name="Arial", size=10, italic=True, color="808080")

    # --- Run Report: one row per source, so a failed board or feed is ------
    # --- visible in the workbook itself, not only in a scrolled-away log ---
    ws_rep = wb.create_sheet("Run Report")
    rep_row = 1
    if incomplete:
        warn = ws_rep.cell(row=1, column=1, value=(
            "INCOMPLETE — a source failed; counts below show what was "
            "actually covered"))
        warn.font = Font(name="Arial", size=12, bold=True, color="9C0006")
        warn.fill = PatternFill("solid", fgColor="FFC7CE")
        rep_row = 3
    rep_headers = ["Source", "ATS", "Status", "Postings", "Kept", "Error"]
    for i, h in enumerate(rep_headers, start=1):
        c = ws_rep.cell(row=rep_row, column=i, value=h)
        c.font, c.fill = head_font, head_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
    for w, letter in zip([28, 16, 10, 10, 8, 60], "ABCDEF"):
        ws_rep.column_dimensions[letter].width = w
    fail_font = Font(name="Arial", size=10, bold=True, color="9C0006")
    for s in sources:
        rep_row += 1
        vals = [s.get("name", ""), s.get("ats", ""), s.get("status", ""),
                s.get("count", 0), s.get("kept", 0), s.get("error", "")]
        for i, v in enumerate(vals, start=1):
            c = ws_rep.cell(row=rep_row, column=i, value=v)
            c.font = fail_font if s.get("status") == "FAILED" else body
            c.border = Border(bottom=thin)
    covered = sum(1 for s in sources if s.get("status") in ("OK", "EMPTY"))
    rep_row += 2
    summary = ws_rep.cell(row=rep_row, column=1, value=(
        f"Run started {RUN_STARTED or '?'}, finished "
        f"{datetime.now():%Y-%m-%d %H:%M:%S}. "
        f"Sources covered: {covered} of {len(sources)}. "
        f"INCOMPLETE: {'yes' if incomplete else 'no'}."))
    summary.font = fail_font if incomplete else Font(
        name="Arial", size=10, italic=True, color="808080")
    ws_rep.freeze_panes = f"A{2 if not incomplete else 4}"

    build(wb.create_sheet("Quarantine (dead links)"), quarantine)

    # --- Fresh sheet: newest of the apply list, best AI Fit first ----------
    fresh_title = f"Fresh ({FRESH_DAYS * 24}h)"
    fresh = [r for r in ready
             if isinstance(r.get("Days Old"), int)
             and not isinstance(r.get("Days Old"), bool)
             and r["Days Old"] <= FRESH_DAYS]
    fresh.sort(key=lambda r: (-(ai_fit_value(r) if ai_fit_value(r) is not None
                                else -1),
                              r.get("Days Old", 0), r.get("Company", "")))

    ws_d = wb.create_sheet("Dashboard", 0)
    build(wb.create_sheet(fresh_title, 1), fresh, True)
    # Never let a range end above row 2: "D2:D1" is a reversed range, and
    # Excel normalises it to D1:D2, which counts the header as a data row.
    nr = max(len(ready) + 1, 2)
    na = max(len(all_rows) + 1, 2)
    nq = max(len(quarantine) + 1, 2)
    nl = max(len(lowfit) + 1, 2)
    ws_d["A1"] = "Job Search Dashboard"
    ws_d["A1"].font = Font(name="Arial", size=16, bold=True, color="1F3864")
    ws_d["A2"] = (f"Last run: {datetime.now():%A %d %B %Y, %H:%M}   |   "
                  f"US roles only   |   every link opened and checked   |   "
                  f"Ready to Apply filtered at AI Fit >= {MIN_AI_FIT}")
    ws_d["A2"].font = Font(name="Arial", size=9, italic=True, color="808080")
    if sources:
        ws_d["A3"] = f"Sources covered: {covered} of {len(sources)}" + (
            "   |   INCOMPLETE — a source failed; see Run Report and rerun "
            "to recover" if incomplete else "   |   full coverage this run")
        ws_d["A3"].font = (Font(name="Arial", size=10, bold=True,
                                color="9C0006") if incomplete else
                           Font(name="Arial", size=9, italic=True,
                                color="808080"))

    fit_c = col("AI Fit")
    stat_c = col("Status")
    comp_c = col("Company")
    lvl_c = col("Level")
    sal_c = col("Salary Range")
    visa_c = col("Visa Signal")
    ghost_c = col("Ghost Risk")
    age_c = col("Days Old")

    nf = max(len(fresh) + 1, 2)
    metrics = [
        ("APPLYABLE NOW", None),
        ("Verified live roles", f"=COUNTA('Ready to Apply'!{comp_c}2:{comp_c}{nr})"),
        (f"Fresh (last {FRESH_DAYS * 24}h)",
         f"=COUNTA('{fresh_title}'!{comp_c}2:{comp_c}{nf})"),
        ("New since last run", f"=COUNTA('New This Run'!{comp_c}2:{comp_c}{nr})"),
        ("Strong matches (70+)",
         f"=COUNTIF('Ready to Apply'!{fit_c}2:{fit_c}{nr},\">=70\")"),
        ("Low ghost-job risk",
         f"=COUNTIF('Ready to Apply'!{ghost_c}2:{ghost_c}{nr},\"FRESH\")"),
        ("Minimum AI Fit kept", MIN_AI_FIT),
        ("", None),
        ("BREAKDOWN", None),
        ("Internships", f"=COUNTIF('Ready to Apply'!{lvl_c}2:{lvl_c}{nr},\"Internship\")"),
        ("New grad / APM", f"=COUNTIF('Ready to Apply'!{lvl_c}2:{lvl_c}{nr},\"New Grad*\")"),
        ("Confirmed visa sponsors",
         f"=COUNTIF('Ready to Apply'!{visa_c}2:{visa_c}{nr},\"SPONSORS*\")"),
        ("Salary published", f"=COUNTIF('Ready to Apply'!{sal_c}2:{sal_c}{nr},\"$*\")"),
        ("Posted within 7 days",
         f"=COUNTIFS('Ready to Apply'!{age_c}2:{age_c}{nr},\"<=7\","
         f"'Ready to Apply'!{age_c}2:{age_c}{nr},\">=0\")"),
        ("", None),
        ("YOUR PIPELINE", None),
        ("Applied", f"=COUNTIF('All Matches'!{stat_c}2:{stat_c}{na},\"Applied\")"),
        ("Interviewing", f"=COUNTIF('All Matches'!{stat_c}2:{stat_c}{na},\"Interviewing\")"),
        ("Still to review",
         f"=COUNTIF('Ready to Apply'!{stat_c}2:{stat_c}{nr},\"To Review\")"),
        ("", None),
        ("FILTERED OUT", None),
        ("Dead / closed links removed",
         f"=COUNTA('Quarantine (dead links)'!{comp_c}2:{comp_c}{nq})"),
        (f"Weak matches (AI Fit < {MIN_AI_FIT})",
         f"=COUNTA('Low Fit'!{comp_c}2:{comp_c}{nl})"),
    ]
    r = 5
    for label, formula in metrics:
        if label and formula is None:
            c = ws_d.cell(row=r, column=1, value=label)
            c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F3864")
            ws_d.cell(row=r, column=2).fill = PatternFill("solid", fgColor="1F3864")
        elif label:
            ws_d.cell(row=r, column=1, value=label).font = Font(name="Arial", size=10)
            c = ws_d.cell(row=r, column=2, value=formula)
            c.font = Font(name="Arial", size=11, bold=True)
            c.alignment = Alignment(horizontal="center")
            c.fill = PatternFill("solid", fgColor="EAF1FB")
        r += 1

    ws_d.column_dimensions["A"].width = 30
    ws_d.column_dimensions["B"].width = 14
    ws_d.column_dimensions["D"].width = 78
    ws_d["D4"] = "How to use this file"
    ws_d["D4"].font = Font(name="Arial", size=12, bold=True, color="1F3864")
    tips = [
        "READY TO APPLY - start here. Every link on this sheet was opened and",
        "confirmed live at run time. Yellow rows are new since your last run.",
        "",
        f"LOW FIT - roles the AI scored below {MIN_AI_FIT} out of 100 against your",
        "resume. They are parked on that sheet rather than deleted, and All Matches",
        "still holds every row found this run. A blank AI Fit means the ranker never",
        "saw that row, so it counts as unknown and stays on Ready to Apply.",
        "",
        "GHOST RISK - FRESH means recent with a published salary. HIGH means the",
        "posting has been open a long time with no salary band, which research links",
        "to pipeline-filler listings. Deprioritise HIGH; you need not skip it entirely.",
        "",
        "LINK STATUS - LIVE was verified during this run. BLOCKED means the site",
        "refused an automated check, so open it yourself. DEAD and CLOSED postings",
        "are moved to the Quarantine sheet and kept out of your way.",
        "",
        "VISA SIGNAL - SPONSORS means the posting says so explicitly. Blank means",
        "unknown, not no. Salary is shown because the H-1B lottery is wage-weighted:",
        "higher wage levels get more entries in the draw.",
        "",
        "Set Status from the dropdown and write in Notes. Both carry over into",
        "tomorrow's file automatically.",
    ]
    for i, t in enumerate(tips):
        ws_d.cell(row=5 + i, column=4, value=t).font = Font(name="Arial", size=10)

    ws_d.sheet_view.showGridLines = False
    # Atomic write: save to a sibling .tmp, then swap it over the target.
    # A reader (app.py, Excel) never sees a half-written file, and a crash
    # mid-save leaves yesterday's workbook untouched. os.replace over a file
    # Excel holds open still raises PermissionError, which the caller's
    # existing OSError guard turns into a "close it in Excel" message.
    tmp = path + ".tmp"
    wb.save(tmp)
    try:
        os.replace(tmp, path)
    except OSError:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise


# --------------------------------------------------------------------------
# Commands
# --------------------------------------------------------------------------

def load_config():
    if not os.path.exists(CONFIG_PATH):
        sys.exit("No config.json yet. Run:  python jobhunt.py setup")
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)


def save_config(cfg):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)


def cmd_setup(_args):
    print("\n" + "=" * 64)
    print("  JOB HUNT SETUP")
    print("  Press Enter to accept any default.")
    print("=" * 64)
    cfg = {"us_only": True}

    print("\n--- 1. What roles? ---")
    cfg["target_titles"] = ask_list("Primary job titles:", [
        "Product Manager", "Product Management", "Associate Product Manager",
        "APM", "Product Analyst", "Technical Product Manager"])
    cfg["secondary_titles"] = ask_list("Backup titles (scored lower):", [
        "Product Operations", "Business Analyst", "Strategy & Operations",
        "Data Analyst", "Program Manager", "Solutions Engineer"])
    cfg["exclude_title_words"] = ask_list("Never show titles containing:", [
        "sales", "account executive", "recruiter", "marketing manager",
        "product marketing", "designer", "clinical"])

    print("\n--- 2. What stage? ---")
    cfg["want_internships"] = ask_yes("Include internships?", True)
    cfg["want_fulltime"] = ask_yes("Include full-time / new-grad roles?", True)

    print("\n--- 3. Where in the US? ---")
    cfg["locations"] = ask_list("Locations to prioritise:", [
        "New York", "NYC", "Manhattan", "Brooklyn", "Jersey City", "Remote"])
    cfg["strict_location"] = ask_yes(
        "Show ONLY those locations (hide rest of the US)?", False)
    print("  Note: non-US roles are always excluded.")

    print("\n--- 4. Visa ---")
    cfg["drop_no_sponsorship"] = ask_yes(
        "Hide postings that explicitly refuse sponsorship?", True)

    print("\n--- 5. Pay ---")
    cfg["salary_floor"] = int(ask("Minimum acceptable salary (USD)", "90000"))
    cfg["salary_target"] = int(ask("Salary you'd call strong (USD)", "140000"))

    print("\n--- 6. Your skills ---")
    cfg["skills"] = ask_list("Skills to look for in descriptions:", [
        "SQL", "A/B testing", "experimentation", "product analytics", "roadmap",
        "user research", "Python", "LLM", "RAG", "machine learning", "API",
        "pricing", "e-commerce", "payments", "growth", "consumer", "marketplace"])

    print("\n--- 7. Freshness and verification ---")
    cfg["max_age_days"] = int(ask("Ignore postings older than N days", "45"))
    cfg["min_ai_fit"] = int(ask(
        "Minimum AI Fit (0-100) for a role to reach Ready to Apply", "55"))
    cfg["verify_links"] = ask_yes(
        "Open and verify every link before it reaches Excel? (slower, recommended)", True)
    cfg["output_dir"] = ask("Where should the Excel files go?",
                            os.path.join(os.path.expanduser("~"), "JobHunt"))
    cfg["companies"] = STARTER_COMPANIES
    cfg["use_simplify"] = ask_yes("Also use the SimplifyJobs GitHub feeds?", True)

    os.makedirs(cfg["output_dir"], exist_ok=True)
    save_config(cfg)
    print("\n" + "=" * 64)
    print(f"  Saved to {CONFIG_PATH}")
    print(f"  Tracking {len(cfg['companies'])} companies to start.\n")
    print("  Next:")
    print("    python jobhunt.py check      <- prune dead company tokens")
    print("    python jobhunt.py run        <- first search")
    print("    python jobhunt.py schedule   <- make it run daily")
    print("=" * 64 + "\n")


def cmd_check(_args):
    cfg = load_config()
    good, bad = [], []
    print(f"\nTesting {len(cfg['companies'])} company boards...\n")
    for c in cfg["companies"]:
        fn = ATS_FETCHERS.get(c["ats"])
        try:
            jobs = fn(c["name"], c["token"])
            if jobs:
                good.append(c)
                print(f"  OK    {c['name']:<22} {c['ats']:<16} {len(jobs)} open roles")
            else:
                bad.append(c)
                print(f"  EMPTY {c['name']:<22} {c['ats']}")
        except Exception:
            bad.append(c)
            print(f"  DEAD  {c['name']:<22} {c['ats']}  (wrong token or moved ATS)")
        time.sleep(0.2)
    print(f"\n{len(good)} working, {len(bad)} not.")
    if bad and ask_yes("Remove the broken ones from config?", True):
        cfg["companies"] = good
        save_config(cfg)
        print("Cleaned up.")
    print("\nAdd a company:  python jobhunt.py add <token-from-careers-URL>\n")


def cmd_add(args):
    cfg = load_config()
    token = args.company.strip().lower()
    print(f"\nProbing every ATS for '{token}'...")
    ats, count = detect_ats(token)
    if not ats:
        print("  Not found. Check the company's careers page URL:")
        print("    job-boards.greenhouse.io/TOKEN   jobs.lever.co/TOKEN")
        print("    jobs.ashbyhq.com/TOKEN           apply.workable.com/TOKEN")
        return
    name = ask(f"  Found on {ats} ({count} roles). Display name", token.title())
    cfg["companies"].append({"name": name, "ats": ats, "token": token})
    save_config(cfg)
    print(f"  Added {name}. Now tracking {len(cfg['companies'])} companies.\n")


def _pid_alive(pid):
    """Is the process holding the lock still running?"""
    if platform.system() != "Windows":
        try:
            os.kill(pid, 0)
            return True
        except OSError:
            return False
    try:
        out = subprocess.run(["tasklist", "/FI", f"PID eq {pid}"],
                             capture_output=True, text=True, timeout=10).stdout
    except Exception:
        return True  # cannot tell - assume alive rather than steal the lock
    return re.search(rf"\b{pid}\b", out) is not None


def _acquire_lock():
    """Take jobhunt.lock (PID inside) or return False if a live run holds it.

    O_CREAT|O_EXCL is atomic, so CLI, web (app.py subprocess) and scheduled
    runs cannot slip past each other no matter which started first. A lock
    left by a crashed process (PID no longer alive) is treated as stale,
    removed, and taken over.
    """
    for _ in range(3):
        try:
            fd = os.open(LOCK_PATH, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, str(os.getpid()).encode("ascii"))
            os.close(fd)
            return True
        except FileExistsError:
            pass
        try:
            with open(LOCK_PATH, encoding="utf-8") as f:
                pid = int(f.read().strip() or 0)
        except (OSError, ValueError):
            pid = 0
        if pid and pid != os.getpid() and _pid_alive(pid):
            log(f"Another jobhunt run is already in progress (PID {pid}).")
            log("Two runs would write the same workbook - exiting. "
                "Wait for it to finish and try again.")
            return False
        try:  # stale lock - its owner is gone
            os.remove(LOCK_PATH)
            log("Removed stale jobhunt.lock from a crashed run")
        except OSError:
            return False
    return False


def _release_lock():
    try:
        os.remove(LOCK_PATH)
    except OSError:
        pass


def cmd_run(_args):
    if not _acquire_lock():
        sys.exit(1)
    global _SESSION
    _SESSION = requests.Session()
    _SESSION.headers.update(UA)
    try:
        _run_locked(_args)
    finally:
        _release_lock()
        try:
            _SESSION.close()
        except Exception:
            pass
        _SESSION = None


def _run_locked(_args):
    global RUN_STARTED
    run_started = datetime.now()
    RUN_STARTED = f"{run_started:%Y-%m-%d %H:%M:%S}"
    cfg = load_config()
    out_dir = cfg.get("output_dir") or HERE
    os.makedirs(out_dir, exist_ok=True)
    log("=" * 58)
    log("Starting job search run (US only, verified links)")
    preflight(cfg)

    raw, source_outcomes = [], []
    boards = list(cfg["companies"])
    log(f"Pulling {len(boards)} company boards (8 in parallel)...")
    progress = [0]
    plock = threading.Lock()

    def fetch_board(c):
        name, ats = c.get("name", "?"), c.get("ats", "")
        fn = ATS_FETCHERS.get(ats)
        jobs = []
        if not fn:
            outcome = {"name": name, "ats": ats, "status": "FAILED",
                       "count": 0, "error": f"unknown ATS '{ats}'"}
        else:
            try:
                jobs = fn(name, c["token"])
                outcome = {"name": name, "ats": ats,
                           "status": "OK" if jobs else "EMPTY",
                           "count": len(jobs), "error": ""}
            except Exception as e:
                outcome = {"name": name, "ats": ats, "status": "FAILED",
                           "count": 0,
                           "error": f"{type(e).__name__}: {e}"[:200]}
        with plock:
            progress[0] += 1
            if outcome["status"] == "FAILED":
                log(f"  ! {name} failed: "
                    f"{outcome['error'].split(':', 1)[0]}")
            elif progress[0] % 10 == 0 or progress[0] == len(boards):
                log(f"  boards {progress[0]}/{len(boards)} fetched")
        return jobs, outcome

    with ThreadPoolExecutor(max_workers=8) as pool:
        for jobs, outcome in pool.map(fetch_board, boards):
            raw.extend(jobs)
            source_outcomes.append(outcome)
    log(f"  {len(raw)} postings from company boards")

    if cfg.get("use_simplify", True):
        log("Pulling SimplifyJobs GitHub feeds...")
        simplify_rows, simplify_outcomes = fetch_simplify()
        raw.extend(simplify_rows)
        source_outcomes.extend(simplify_outcomes)

    run_incomplete = any(s["status"] == "FAILED" for s in source_outcomes)
    if run_incomplete:
        failed = [s["name"] for s in source_outcomes if s["status"] == "FAILED"]
        log(f"! {len(failed)} source(s) failed ({', '.join(failed[:6])}"
            f"{'...' if len(failed) > 6 else ''}) - this run is INCOMPLETE")
    log(f"{len(raw)} total postings fetched. Filtering to US + scoring...")

    state = {}
    if os.path.exists(STATE_PATH):
        try:
            with open(STATE_PATH, encoding="utf-8") as f:
                state = json.load(f)
        except (OSError, json.JSONDecodeError):
            state = {}
    first_seen = state.get("first_seen", {})
    known = set(first_seen.keys())

    seen, rows = set(), []
    for j in raw:
        if not j.get("title") or not j.get("url"):
            continue
        key = job_key(j["company"], j["title"], j["url"])
        if key in seen:
            continue
        seen.add(key)
        result = score_job(j, cfg)
        if not result:
            continue
        score, reasons, visa, level = result
        if key not in first_seen:
            first_seen[key] = TODAY
        desc_clean = re.sub(r"\s+", " ", j.get("description") or "").strip()
        row = {
            "Score": score, "Status": "To Review", "Company": j["company"],
            "Title": j["title"], "Level": level, "Location": j.get("location", ""),
            "Salary Range": j.get("salary_display", ""), "Visa Signal": visa,
            "Link Status": "", "Ghost Risk": "", "Posted": j.get("posted", ""),
            "Days Old": days_since(j.get("posted")), "Team": j.get("team", ""),
            "Type": j.get("emp_type", ""), "Source": j.get("source", ""),
            "Eligibility": "",  # "" = ranker never saw it; ai_rank fills YES/NO/UNKNOWN
            "Why It Matched": reasons, "Risk Notes": "",
            "What The Role Says": desc_clean[:400],
            "Apply Link": j["url"], "Applied Date": "", "Notes": "", "Key": key,
            # Private key for ai_rank._batch_prompt - cleaned plain text,
            # NOT in COLUMNS, never written to Excel.
            "_rank_text": desc_clean[:1500],
        }
        row["Ghost Risk"], row["Risk Notes"] = ghost_risk(row, first_seen[key])
        rows.append(row)

    rows.sort(key=lambda r: (-r["Score"], r["Company"]))
    log(f"{len(rows)} US matches after scoring")

    # Per-source "Kept" for the Run Report. Board rows carry the ATS name in
    # "Source", so board outcomes are matched by company; Simplify rows carry
    # the feed label itself.
    kept_by_source = {}
    for r in rows:
        src = r.get("Source", "")
        k = src if src.startswith("Simplify") else r.get("Company", "")
        kept_by_source[k] = kept_by_source.get(k, 0) + 1
    for s in source_outcomes:
        s["kept"] = kept_by_source.get(s["name"], 0)

    if cfg.get("verify_links", True) and rows:
        log(f"Opening and verifying {len(rows)} links (this is the slow part)...")
        verify_all(rows)
        counts = {}
        for r in rows:
            counts[r["Link Status"]] = counts.get(r["Link Status"], 0) + 1
        log("  " + ", ".join(f"{k}: {v}" for k, v in sorted(counts.items())))
    else:
        for r in rows:
            r["Link Status"] = "UNVERIFIED"

    # UNREACHABLE means OUR check failed (timeout / the ATS rate-limited us),
    # not that the job is gone. Quarantining it silently deletes live roles --
    # on a busy run that was ~29% of everything found. Only a definite DEAD/CLOSED
    # signal from the site itself is grounds for removal.
    # "SITE DOWN" is set by _mark_dead_hosts: every link on that host failed, so
    # it is their outage, not our rate limit. Real jobs, unusable links today --
    # they keep their row everywhere except the apply list.
    ready = [r for r in rows
             if r["Link Status"] in ("LIVE", "BLOCKED", "UNVERIFIED", "UNREACHABLE")]
    needs_check = [r for r in rows if r["Link Status"] == "SITE DOWN"]
    quarantine = [r for r in rows if r["Link Status"] in ("DEAD", "CLOSED")]
    log(f"{len(ready)} applyable, {len(needs_check)} awaiting a site that is down, "
        f"{len(quarantine)} dead/closed removed")

    if ready:
        log("AI relevance ranking (Gemini)...")
        try:
            import ai_rank
            if ai_rank.rerank(ready, log):
                def by_fit(r):
                    fit = ai_fit_value(r)
                    return (-(fit if fit is not None else -1),
                            -r.get("Score", 0), r.get("Company", ""))
                rows.sort(key=by_fit)
                ready.sort(key=by_fit)
        except Exception as e:
            log(f"  ! AI ranking unavailable ({type(e).__name__}) - keyword order kept")

    # Demote weak matches off the apply list. A blank "AI Fit" is UNKNOWN, not
    # bad - the ranker never scored it - so it is always kept. Only a real int
    # under the cutoff is demoted, and demoted rows still appear on Low Fit and
    # All Matches, so nothing found this run becomes invisible.
    global MIN_AI_FIT, FRESH_DAYS
    try:
        MIN_AI_FIT = max(0, min(100, int(cfg.get("min_ai_fit", 55))))
    except (TypeError, ValueError):
        MIN_AI_FIT = 55
    try:
        FRESH_DAYS = max(0, min(30, int(cfg.get("fresh_days", 2))))
    except (TypeError, ValueError):
        FRESH_DAYS = 2
    keep, lowfit = [], []
    ineligible = 0
    for r in ready:
        fit = ai_fit_value(r)
        if r.get("Eligibility") == "NO":
            # Hard blocker stated in the posting text (no sponsorship, requires
            # citizenship/clearance, 4+ yrs) - demote regardless of fit score.
            # YES/UNKNOWN/"" are never demoted for eligibility reasons.
            lowfit.append(r)
            ineligible += 1
        elif fit is not None and fit < MIN_AI_FIT:
            lowfit.append(r)
        else:
            keep.append(r)
    if ineligible:
        log(f"Demoted {ineligible} roles marked Eligibility NO "
            f"(hard blocker in posting) into the Low Fit sheet")
    if lowfit:
        log(f"Filtered {len(lowfit)} roles below AI Fit {MIN_AI_FIT} "
            f"into the Low Fit sheet ({len(keep)} kept)")
    else:
        log(f"No roles scored below AI Fit {MIN_AI_FIT} ({len(keep)} kept)")

    new_keys = {r["Key"] for r in ready} - known
    log(f"{len(new_keys)} are new since the last run")

    latest = os.path.join(out_dir, "job_matches_latest.xlsx")
    prior = load_prior_tracker(latest)
    carried = 0
    for r in rows:
        p = prior.get(r["Key"])
        if p:
            if p["Status"]:
                r["Status"] = p["Status"]
            r["Applied Date"], r["Notes"] = p["Applied Date"], p["Notes"]
            carried += 1
    if carried:
        log(f"Carried your notes forward on {carried} rows")

    # Both saves are guarded. On Windows, Excel holding either workbook open
    # raises PermissionError, and an unguarded save killed the run right here -
    # losing the whole fetch/verify/rank pass AND the first_seen state below,
    # which would then report every role as new again on the next run.
    dated = os.path.join(out_dir, f"job_matches_{TODAY}.xlsx")
    saved = []
    for target in (dated, latest):
        try:
            write_excel(rows, keep, new_keys, quarantine, target, lowfit,
                        source_outcomes, run_incomplete, needs_check)
            saved.append(target)
        except OSError as e:
            log(f"  ! could not save {os.path.basename(target)} "
                f"({type(e).__name__}) - close it in Excel and re-run")

    state["first_seen"] = first_seen
    state["last_run"] = datetime.now().isoformat()
    try:
        with open(STATE_PATH, "w", encoding="utf-8") as f:
            json.dump(state, f)
    except OSError as e:
        log(f"  ! could not save state.json ({type(e).__name__})")

    if saved:
        log(f"Wrote {saved[0]}")
    else:
        log("  ! nothing was saved this run - no spreadsheet was updated")
    if keep:
        log("Top verified matches:")
        for r in keep[:5]:
            log(f"   {r['Score']:>3}  [{r['Ghost Risk']:<6}] {r['Company']} - {r['Title']}")
    if run_incomplete:
        log("! Run finished INCOMPLETE - a source failed; see the Run Report "
            "sheet and rerun to recover")

    publish_results(keep, saved[0] if saved else "")

    record_run_history(
        rows=rows, keep=keep, lowfit=lowfit, needs_check=needs_check,
        quarantine=quarantine, new_keys=new_keys,
        source_outcomes=source_outcomes, incomplete=run_incomplete,
        started=run_started, workbook=(saved[0] if saved else ""))  # noqa: F821

    log("Run complete")
    log("=" * 58)


# --------------------------------------------------------------------------
# Preflight
# --------------------------------------------------------------------------

def preflight(cfg):
    """Say what this run is about to do, and check the two things that can
    silently spoil it: the Gemini key and the shared store.

    Checked BEFORE the 7-minute fetch, not after. Discovering a dead key at the
    ranking stage means the whole sweep produced an unranked list.
    """
    log("-" * 58)

    # 1. Gemini - the key is what decides whether ranking happens at all
    try:
        import ai_rank
        key = ai_rank._api_key()
    except Exception:
        key = None
    if not key:
        log("  GEMINI    no API key found - roles will be keyword-ranked only")
    else:
        try:
            reply = ai_rank._call(key, 'Reply with exactly: OK', retries=1, timeout=20)
            ok = "OK" in (reply or "")
            log(f"  GEMINI    key valid, model {ai_rank.MODEL} responded"
                if ok else
                f"  GEMINI    key answered oddly ({str(reply)[:40]!r}) - continuing")
        except Exception as e:
            name = type(e).__name__
            log(f"  GEMINI    key present but unreachable ({name})")
            log("            roles will be keyword-ranked only this run")

    # 2. Shared storage - whether results reach the website
    try:
        import storage
        if storage.enabled():
            ok, msg = storage.ping()
            log(f"  STORAGE   {msg}" if ok
                else f"  STORAGE   configured but unreachable - {msg}")
            if ok:
                log("            results will publish to the website when this finishes")
        else:
            log("  STORAGE   local files only - the website will not see this run")
            log("            set KV_REST_API_URL and KV_REST_API_TOKEN to publish")
    except Exception as e:
        log(f"  STORAGE   could not check ({type(e).__name__})")

    # 3. What this run will actually keep
    log(f"  PROFILE   {_profile_name()}")
    log(f"  CUTOFF    fit {MIN_AI_FIT}+ reaches the apply list"
        f"   |  fresh = last {cfg.get('fresh_days', 2)} days"
        f"   |  max age {cfg.get('max_age_days', 45)} days")
    log(f"  BOARDS    {len(cfg.get('companies', []))} company job boards")
    log("-" * 58)


def _profile_name():
    try:
        with open(os.path.join(HERE, "profile.json"), encoding="utf-8-sig") as f:
            p = json.load(f)
        who = p.get("name") or "unnamed"
        res = p.get("resume_file")
        return f"{who}" + (f"  (from {res})" if res else "")
    except (OSError, ValueError):
        return "no profile.json - using the built-in default rubric"


# --------------------------------------------------------------------------
# Publishing to shared storage
# --------------------------------------------------------------------------

def publish_results(keep, workbook_path):
    """Push this run's rows (and the workbook) to the shared store.

    This is the whole point of the store: the search needs 7-15 minutes, which
    no serverless function will give it, so it runs here and the hosted site
    reads what it left behind. Best effort throughout -- the spreadsheet on
    this machine is already written by now, and a store being down must never
    turn a successful run into a failed one.
    """
    try:
        import storage
    except ImportError:
        return
    if not storage.enabled():
        return

    try:
        rows = []
        for r in keep:
            rows.append({
                "fit": r.get("AI Fit") if isinstance(r.get("AI Fit"), int) else None,
                "score": r.get("Score", 0),
                "company": r.get("Company", ""),
                "title": r.get("Title", ""),
                "location": r.get("Location", ""),
                "why": r.get("AI Verdict", ""),
                "link": r.get("Apply Link", ""),
                "level": r.get("Level", ""),
                "days": days_since(r.get("Posted")),
                "fresh": bool(r.get("Fresh")),
                "eligibility": r.get("Eligibility", "UNKNOWN"),
            })
        payload = {
            "rows": rows,
            "total": len(rows),
            "threshold": MIN_AI_FIT,
            "generated": f"{datetime.now():%Y-%m-%d %H:%M}",
            "machine": platform.node(),
        }
        log("-" * 58)
        log(f"PUBLISHING to {storage.describe()}")
        if storage.set_json("results", payload):
            scored = sum(1 for r in rows if r["fit"] is not None)
            log(f"  roles      {len(rows)} sent  ({scored} AI-scored)")
        else:
            log("  roles      FAILED - the website keeps the previous set")
            log("             (nothing is lost; this machine's Excel is fine)")

        # history and profile travel with them so the site's trends and the
        # attached-resume panel are not stuck on whatever it last saw
        hist = load_history()
        log(f"  history    {len(hist)} runs sent"
            if storage.set_json("history", {"runs": hist})
            else "  history    FAILED")
        try:
            with open(os.path.join(HERE, "profile.json"), encoding="utf-8-sig") as f:
                prof = json.load(f)
            log(f"  profile    {prof.get('name', 'sent')}"
                if storage.set_json("profile", prof) else "  profile    FAILED")
        except (OSError, ValueError):
            log("  profile    skipped - no readable profile.json")

        if workbook_path and os.path.exists(workbook_path):
            with open(workbook_path, "rb") as f:
                blob = f.read()
            ok, why = storage.set_bytes("workbook", blob)
            if ok:
                log(f"  workbook   {len(blob)//1024} KB sent - downloadable from the site")
            else:
                log(f"  workbook   NOT sent - {why}")
                log("             roles are still on the site; only the download "
                    "needs this machine")
        else:
            log("  workbook   skipped - no file was saved this run")
        log(f"  website    https://jobhunt-tau-seven.vercel.app  is now showing this run")
        log("-" * 58)
    except Exception as e:
        log(f"  ! publishing failed ({type(e).__name__}) - local results are fine")


# --------------------------------------------------------------------------
# Run history
# --------------------------------------------------------------------------

def _load_dotenv():
    """Read .env into the environment.

    app.py already did this; the CLI did not, so `python jobhunt.py run` could
    not see KV_REST_API_URL and silently skipped publishing - the run looked
    perfect and the website stayed on yesterday's roles. Real environment
    variables win, so a scheduled task or CI can override the file.
    """
    try:
        with open(os.path.join(HERE, ".env"), encoding="utf-8-sig",
                  errors="replace") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                if line.startswith("export "):
                    line = line[7:].lstrip()
                name, _, val = line.partition("=")
                name = name.strip()
                if name and name not in os.environ:
                    os.environ[name] = val.strip().strip('"').strip("'")
    except OSError:
        pass


_load_dotenv()


HISTORY_PATH = os.path.join(HERE, "history.json")
HISTORY_MAX = 400


def _fit_buckets(rows):
    """Fit distribution for one run. Unscored rows are counted separately --
    folding them into '0-29' would make a Gemini outage look like a day when
    every job was a bad match."""
    b = {"90-100": 0, "75-89": 0, "55-74": 0, "30-54": 0, "0-29": 0, "unscored": 0}
    for r in rows:
        f = ai_fit_value(r)
        if f is None:
            b["unscored"] += 1
        elif f >= 90:
            b["90-100"] += 1
        elif f >= 75:
            b["75-89"] += 1
        elif f >= 55:
            b["55-74"] += 1
        elif f >= 30:
            b["30-54"] += 1
        else:
            b["0-29"] += 1
    return b


def load_history():
    try:
        with open(HISTORY_PATH, encoding="utf-8-sig") as f:
            data = json.load(f)
        runs = data.get("runs") if isinstance(data, dict) else data
        return [r for r in runs if isinstance(r, dict)] if isinstance(runs, list) else []
    except (OSError, ValueError):
        return []


def record_run_history(rows, keep, lowfit, needs_check, quarantine, new_keys,
                       source_outcomes, incomplete, started, workbook):
    """Append one row to history.json so today can be compared with yesterday.

    Appends -- never replaces. Two runs on the same day are two entries, because
    collapsing them would hide that the morning run found 8 roles and the evening
    one found 600. Failure here must never fail the run: the spreadsheet is
    already written by this point.
    """
    try:
        finished = datetime.now()
        ok = sum(1 for s in source_outcomes if s.get("status") == "OK")
        failed = [s.get("name", "?") for s in source_outcomes
                  if s.get("status") == "FAILED"]
        profile = {}
        try:
            with open(os.path.join(HERE, "profile.json"), encoding="utf-8-sig") as f:
                profile = json.load(f) or {}
        except (OSError, ValueError):
            pass

        entry = {
            "id": finished.strftime("%Y-%m-%dT%H:%M:%S"),
            "date": finished.strftime("%Y-%m-%d"),
            "time": finished.strftime("%H:%M"),
            "started": started.strftime("%Y-%m-%d %H:%M:%S") if started else "",
            "finished": finished.strftime("%Y-%m-%d %H:%M:%S"),
            "duration_sec": int((finished - started).total_seconds()) if started else None,
            "total": len(rows),
            "ready": len(keep),
            "new": len(new_keys),
            "lowfit": len(lowfit),
            "needs_check": len(needs_check),
            "quarantine": len(quarantine),
            "buckets": _fit_buckets(keep),
            "threshold": MIN_AI_FIT,
            "incomplete": bool(incomplete),
            "sources_ok": ok,
            "sources_total": len(source_outcomes),
            "failed_sources": failed[:12],
            "workbook": os.path.basename(workbook) if workbook else "",
            "profile_name": str(profile.get("name") or ""),
            "resume_file": str(profile.get("resume_file") or ""),
        }
        runs = load_history()
        runs.append(entry)
        runs = runs[-HISTORY_MAX:]
        tmp = HISTORY_PATH + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump({"runs": runs}, f, indent=2, ensure_ascii=False)
        os.replace(tmp, HISTORY_PATH)
        log(f"Logged this run to history.json ({len(runs)} runs kept)")
    except Exception as e:
        log(f"  ! could not write run history ({type(e).__name__}) - "
            f"the spreadsheet is unaffected")


def cmd_schedule(_args):
    cfg = load_config()
    py, script = sys.executable, os.path.abspath(__file__)
    hour = ask("What hour should it run each day? (0-23)", "7")
    print()

    if platform.system() == "Windows":
        cmd = (f'schtasks /Create /SC DAILY /TN "JobHunt" /ST {int(hour):02d}:00 '
               f'/TR "\'{py}\' \'{script}\' run" /F')
        print("Run this once in PowerShell:\n\n  " + cmd + "\n")
        print('To remove later:  schtasks /Delete /TN "JobHunt" /F')
        return

    entry = (f"0 {int(hour)} * * * cd {os.path.dirname(script)} && "
             f"{py} {script} run >> {LOG_PATH} 2>&1")
    try:
        existing = subprocess.run(["crontab", "-l"], capture_output=True, text=True).stdout
    except FileNotFoundError:
        print("cron not found. Add this to your scheduler manually:\n  " + entry)
        return
    if script in existing:
        print("Already scheduled. Current crontab:\n\n" + existing)
        return
    if not ask_yes(f"Add a daily {int(hour):02d}:00 run to your crontab?", True):
        print("Skipped. To do it by hand:  crontab -e   then add:\n  " + entry)
        return
    subprocess.run(["crontab", "-"],
                   input=(existing.rstrip("\n") + "\n" + entry + "\n").lstrip("\n"),
                   text=True, check=True)
    print(f"\nScheduled. Every day at {int(hour):02d}:00 a verified Excel lands in:")
    print(f"  {cfg['output_dir']}\nLog: {LOG_PATH}")
    print("To remove:  crontab -e   and delete the jobhunt line")
    if platform.system() == "Darwin":
        print("\nmacOS: if it doesn't fire, add /usr/sbin/cron to")
        print("System Settings > Privacy & Security > Full Disk Access.")


def main():
    p = argparse.ArgumentParser(description="Personal job search engine")
    sub = p.add_subparsers(dest="cmd")
    for name, helptext in [("setup", "interactive configuration"),
                           ("run", "search, verify, write Excel"),
                           ("check", "validate company tokens"),
                           ("schedule", "install daily background run")]:
        sub.add_parser(name, help=helptext)
    a = sub.add_parser("add", help="add a company by ATS token")
    a.add_argument("company")
    args = p.parse_args()

    handlers = {"setup": cmd_setup, "run": cmd_run, "check": cmd_check,
                "add": cmd_add, "schedule": cmd_schedule}
    if args.cmd not in handlers:
        p.print_help()
        return
    handlers[args.cmd](args)


if __name__ == "__main__":
    main()
