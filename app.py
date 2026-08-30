"""
Local web UI for the job search engine.

This is a thin control panel around the two things that already work:
  jobhunt.py  - fetches, filters, verifies and writes the Excel
  ai_rank.py  - the one and only Gemini code path in this project

Everything here is glue: serve one HTML page, test/save the API key, take a
resume upload, launch `python jobhunt.py run` as a subprocess, stream its log
lines back to the browser, and hand over the finished spreadsheet.

Run it with:  python app.py     then open http://127.0.0.1:5000
"""

import hashlib
import hmac
import json
import os
import re
import subprocess
import sys
import threading
import urllib.error
import webbrowser
from datetime import datetime, timedelta
from urllib.parse import quote

from flask import (Flask, jsonify, redirect, request, send_from_directory,
                   session)
from werkzeug.utils import secure_filename

import openpyxl

import storage

# --------------------------------------------------------------------------
# paths
# --------------------------------------------------------------------------

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(PROJECT_DIR, "static")

# On Vercel (and any serverless host) the deployment directory is READ-ONLY and
# only /tmp accepts writes. Creating uploads/ beside the code at import time is
# what crashed the function with FUNCTION_INVOCATION_FAILED before it served a
# single request. Everything writable therefore moves under /tmp there.
# /tmp is also wiped between invocations, so nothing written there survives --
# see DEPLOY.md. This keeps the app *running*; it does not make it persistent.
IS_SERVERLESS = bool(os.environ.get("VERCEL") or os.environ.get("AWS_LAMBDA_FUNCTION_NAME"))
WRITABLE_DIR = os.path.join("/tmp", "resumify") if IS_SERVERLESS else PROJECT_DIR

UPLOAD_DIR = os.path.join(WRITABLE_DIR, "uploads")
ENV_PATH = os.path.join(PROJECT_DIR, ".env")            # read-only is fine
CONFIG_PATH = os.path.join(PROJECT_DIR, "config.json")  # shipped defaults
CONFIG_WRITE_PATH = (os.path.join(WRITABLE_DIR, "config.json")
                     if IS_SERVERLESS else CONFIG_PATH)
PROFILE_PATH = os.path.join(WRITABLE_DIR, "profile.json")
EXCEL_NAME = "job_matches_latest.xlsx"
READY_SHEET = "Ready to Apply"

ALLOWED_EXT = {".pdf", ".docx", ".txt"}
MAX_RESUME_BYTES = 10 * 1024 * 1024          # 10 MB
CREATE_NO_WINDOW = 0x08000000                # Windows: no console flash


def _ensure_dir(path):
    """Best effort. A read-only filesystem must not stop the app importing."""
    try:
        os.makedirs(path, exist_ok=True)
        return True
    except OSError:
        return False


_ensure_dir(UPLOAD_DIR)

app = Flask(__name__, static_folder=None)
# Bigger than our own limit so an oversize file reaches our friendly message
# instead of werkzeug's 413 page.
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024

# ---- accounts -------------------------------------------------------------
# Load .env before auth reads ADMIN_* / GOOGLE_* out of the environment.
def _load_dotenv():
    try:
        with open(ENV_PATH, encoding="utf-8-sig", errors="replace") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                if line.startswith("export "):
                    line = line[7:].lstrip()
                name, _, val = line.partition("=")
                name = name.strip()
                # Real environment variables win: a deployment sets its own
                # secrets and must not be overridden by a checked-out .env.
                if name and name not in os.environ:
                    os.environ[name] = val.strip().strip('"').strip("'")
    except OSError:
        pass


_load_dotenv()

import auth  # noqa: E402  (must follow _load_dotenv)

app.secret_key = auth.secret_key()
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    # Only send the cookie over HTTPS once deployed; localhost stays plain http.
    SESSION_COOKIE_SECURE=bool(os.environ.get("VERCEL") or
                               os.environ.get("FORCE_HTTPS")),
    PERMANENT_SESSION_LIFETIME=timedelta(days=14),
)
auth.ensure_admin()
auth.ensure_local_users()

# Set to "1" to run the old single-user way with no login at all.
AUTH_DISABLED = os.environ.get("DISABLE_AUTH", "").strip() == "1"


def login_required(fn):
    return fn if AUTH_DISABLED else auth.login_required(fn)


def admin_required(fn):
    return fn if AUTH_DISABLED else auth.admin_required(fn)


def me():
    """The signed-in user, or a stand-in when auth is switched off."""
    if AUTH_DISABLED:
        return {"id": "local", "email": "", "name": "You", "is_admin": True}
    return auth.current_user()


def track(kind, detail=None):
    if not AUTH_DISABLED:
        auth.log_event(auth.current_user(), kind, detail)


# --------------------------------------------------------------------------
# small helpers: config, .env, excel paths
# --------------------------------------------------------------------------

def _store_key(path):
    """File path -> shared-store key. Mirrors auth._store_key."""
    path = os.path.abspath(path)
    stem = os.path.splitext(os.path.basename(path))[0]
    parent = os.path.dirname(path)
    if os.path.basename(os.path.dirname(parent)) == "users":
        return "user:%s:%s" % (os.path.basename(parent), stem)
    return stem


def _read_json(path, default=None):
    # utf-8-sig, not utf-8: a config.json or profile.json re-saved by Notepad
    # on Windows carries a BOM, which makes json.load raise and would silently
    # blank out output_dir / the whole profile.
    if storage.enabled():
        got = storage.get_json(_store_key(path), storage.MISSING)
        if got is not storage.MISSING:
            return got
    try:
        with open(path, encoding="utf-8-sig", errors="replace") as f:
            return json.load(f)
    except (OSError, ValueError):
        return default


def _write_json_atomic(path, data):
    """Write a temp file beside `path`, then swap it in.

    A plain open(path, "w") truncates first, so a crash - or two requests
    writing at once - can leave config.json empty and take companies,
    output_dir and the saved threshold with it. os.replace is only atomic on
    one filesystem, hence the temp file living in the same directory.
    """
    wrote_remote = storage.set_json(_store_key(path), data) if storage.enabled() else False
    tmp = path + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
            f.write("\n")
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, path)
    except OSError:
        try:
            os.remove(tmp)
        except OSError:
            pass
        # A read-only disk is expected on serverless; only a failure on BOTH
        # backends is a real error.
        if not wrote_remote:
            raise


# Both writers of config.json - the resume upload's keyword merge and the fit
# threshold - do a read-modify-write, so they must not interleave. Without
# this, saving a threshold while a resume finishes parsing loses whichever
# change read the file first.
_config_lock = threading.Lock()


def _load_config():
    """Prefer the writable copy, fall back to the one shipped in the repo.

    On serverless these are two different files: config.json in the deployment
    is read-only, so any saved threshold lands in the /tmp copy instead."""
    cfg = _read_json(CONFIG_WRITE_PATH, None)
    if not isinstance(cfg, dict) or not cfg:
        cfg = _read_json(CONFIG_PATH, {})
    return cfg if isinstance(cfg, dict) else {}


def _save_config(cfg):
    _ensure_dir(os.path.dirname(CONFIG_WRITE_PATH))
    _write_json_atomic(CONFIG_WRITE_PATH, cfg)


def _output_dir():
    """Where jobhunt.py drops the spreadsheets."""
    out = _load_config().get("output_dir")
    if isinstance(out, str) and out.strip() and os.path.isdir(out):
        return out
    return PROJECT_DIR


def _excel_path():
    """The NEWEST workbook we can find, not just job_matches_latest.xlsx.

    When Excel has job_matches_latest.xlsx open, Windows locks it and the run
    can only write job_matches_<date>.xlsx. Returning "latest" blindly then
    serves a workbook hours older than the search that just finished, with no
    sign anything is wrong -- the user downloads yesterday's roles believing
    they are today's. Picking by mtime means a locked file degrades into
    "slightly odd filename", never into silently stale data.
    """
    import glob

    candidates = []
    for folder in {_output_dir(), PROJECT_DIR}:
        candidates.append(os.path.join(folder, EXCEL_NAME))
        candidates.extend(glob.glob(os.path.join(folder, "job_matches_*.xlsx")))

    best, best_mtime = None, -1.0
    for path in candidates:
        # ~$job_matches_*.xlsx are Excel's own lock files, not workbooks.
        if os.path.basename(path).startswith("~$"):
            continue
        try:
            mtime = os.path.getmtime(path)
        except OSError:
            continue
        if mtime > best_mtime:
            best, best_mtime = path, mtime

    return best or os.path.join(_output_dir(), EXCEL_NAME)


ENV_NAME = "GEMINI_API_KEY"


def _env_line_value(line):
    """The value if this .env line assigns GEMINI_API_KEY, else None.

    Exact-name matching, the same rule ai_rank._api_key() uses. A prefix test
    would make GEMINI_API_KEY_OLD=... win here while ai_rank kept using the
    real key, so the two halves of the app must parse .env identically.
    """
    text = line.strip()
    if not text or text.startswith("#"):
        return None
    if text.startswith("export "):
        text = text[len("export "):].lstrip()
    name, sep, val = text.partition("=")
    if not sep or name.strip() != ENV_NAME:
        return None
    return val.strip().strip('"').strip("'").strip()


def _env_key():
    """The saved Gemini key, from the live process env or from .env."""
    key = os.environ.get("GEMINI_API_KEY", "").strip()
    if key:
        return key
    try:
        # utf-8-sig: a .env rewritten by PowerShell/Notepad carries a BOM,
        # which would glue itself to the key name and hide the key.
        with open(ENV_PATH, encoding="utf-8-sig", errors="replace") as f:
            for line in f:
                val = _env_line_value(line)
                if val:
                    return val
    except OSError:
        pass
    return ""


def _save_env_key(key):
    """Rewrite only the GEMINI_API_KEY line; leave every other line alone."""
    try:
        with open(ENV_PATH, encoding="utf-8-sig", errors="replace") as f:
            lines = f.read().splitlines()
    except OSError:
        lines = []

    new_line = ENV_NAME + "=" + key
    out, replaced = [], False
    for line in lines:
        if _env_line_value(line) is None:
            out.append(line)
        elif not replaced:                 # first assignment wins its place
            out.append(new_line)
            replaced = True
        # any later duplicate assignment is dropped, so the key we just saved
        # is the only one _env_key()/ai_rank can find
    if not replaced:
        out.append(new_line)

    with open(ENV_PATH, "w", encoding="utf-8") as f:
        f.write("\n".join(out).rstrip("\n") + "\n")


# --------------------------------------------------------------------------
# reading the spreadsheet
# --------------------------------------------------------------------------

def _as_int(value):
    """Excel hands back ints, floats, strings or None. Return an int or None."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        try:
            return int(value)
        except (ValueError, OverflowError):
            return None
    m = re.search(r"-?\d+", str(value))
    return int(m.group(0)) if m else None


def _bucket_for(fit):
    if fit >= 90:
        return "90-100"
    if fit >= 75:
        return "75-89"
    if fit >= 55:
        return "55-74"
    if fit >= 30:
        return "30-54"
    return "0-29"


def _ready_sheet(wb):
    """The Ready to Apply sheet, or a sheet that is plainly the same thing.

    Never wb.sheetnames[0] as a blind fallback: sheet 0 is the Dashboard, and
    reading it by header name yields rows of empty strings that both /api/state
    and /api/results would hand over as if they were real jobs. Falling back on
    the header row instead means a renamed sheet still works and a missing one
    is reported as an error.
    """
    if READY_SHEET in wb.sheetnames:
        return wb[READY_SHEET]
    for name in wb.sheetnames:
        ws = wb[name]
        try:
            header = next(ws.iter_rows(min_row=1, max_row=1, values_only=True),
                          None)
        except Exception:
            continue
        idx = _header_index(header)
        if "company" in idx and "title" in idx:
            return ws
    return None


def _header_index(header_row):
    """Map lower-cased header text -> column position. Never a fixed index."""
    idx = {}
    for i, value in enumerate(header_row or ()):
        if value is None:
            continue
        name = str(value).strip().lower()
        if name and name not in idx:
            idx[name] = i
    return idx


def _cell(row, idx, *names):
    """First header that exists wins; a missing column degrades to ''."""
    for name in names:
        i = idx.get(name)
        if i is not None and i < len(row):
            value = row[i]
            if value is not None:
                return value
    return ""


def _blank_row(row):
    return row is None or all(
        c is None or (isinstance(c, str) and not c.strip()) for c in row)


def _read_excel_summary(path):
    """{'total', 'top', 'buckets'} from the Ready to Apply sheet, or None."""
    if not path or not os.path.exists(path):
        return None
    wb = None
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = _ready_sheet(wb)
        if ws is None:
            return None

        rows = ws.iter_rows(values_only=True)
        idx = _header_index(next(rows, None))

        buckets = {"90-100": 0, "75-89": 0, "55-74": 0, "30-54": 0, "0-29": 0}
        top, total = [], 0

        for row in rows:
            if _blank_row(row):
                continue
            total += 1
            raw_fit = _cell(row, idx, "ai fit", "score")
            fit = _as_int(raw_fit)
            if fit is not None:
                buckets[_bucket_for(max(0, min(100, fit)))] += 1
            if len(top) < 10:
                top.append({
                    "fit": fit if fit is not None else str(raw_fit),
                    "company": str(_cell(row, idx, "company")),
                    "title": str(_cell(row, idx, "title")),
                    "location": str(_cell(row, idx, "location")),
                    "why": str(_cell(row, idx, "ai verdict", "why it matched")),
                })
        return {"total": total, "top": top, "buckets": buckets}
    except Exception:
        return None
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass


_excel_cache = {"stamp": None, "info": None}


def _excel_info():
    """{'filename', 'rows', 'modified'} for /api/state, cached on mtime."""
    global _excel_cache
    path = _excel_path()
    try:
        st = os.stat(path)
    except OSError:
        return None

    stamp = (path, st.st_mtime, st.st_size)
    if _excel_cache["stamp"] == stamp and _excel_cache["info"]:
        return _excel_cache["info"]

    summary = _read_excel_summary(path)
    stamp_text = datetime.fromtimestamp(st.st_mtime).strftime(
        "%b %d, %Y at %I:%M %p")
    info = {
        "filename": os.path.basename(path),
        "rows": summary["total"] if summary else 0,
        "modified": stamp_text.replace(" 0", " ").replace("at  ", "at "),
    }
    # One rebind, not two key writes: another request thread can never see a
    # fresh stamp paired with the previous run's info.
    _excel_cache = {"stamp": stamp, "info": info}
    return info


# --------------------------------------------------------------------------
# tracker storage (tracker.json)
# --------------------------------------------------------------------------

TRACKER_PATH = os.path.join(PROJECT_DIR, "tracker.json")


def _tracker_path(uid=None):
    """Each account gets its own tracker. The legacy root tracker.json is
    migrated into the first account that asks for it, so nothing saved before
    logins existed is lost."""
    if AUTH_DISABLED:
        return TRACKER_PATH
    u = uid or (me() or {}).get("id")
    if not u:
        return TRACKER_PATH
    path = auth.user_file(u, "tracker.json")
    if not os.path.exists(path) and os.path.exists(TRACKER_PATH):
        try:
            legacy = _read_json(TRACKER_PATH, None)
            if legacy:
                _write_json_atomic(path, legacy)
        except OSError:
            pass
    return path

STATUSES = ("saved", "applied", "interview", "offer", "rejected")
TEXT_CAP = 200                 # company / role / link / location
NOTE_CAP = 1000                # the free-text note

# Its own lock, deliberately not _lock: the reader thread holds _lock for
# every log line of a running search, so sharing it would stall the tracker
# behind a subprocess. Reads take it too, because os.replace() on Windows
# fails while another handle has the file open.
#
# Every read-modify-write below happens inside this lock as one step. Two adds
# arriving together would otherwise both start from the same stale list and
# the second write would drop the first one's row.
_tracker_lock = threading.Lock()


def _clean(value, cap=TEXT_CAP):
    """One trimmed, collapsed, length-capped line out of anything.

    The cap is what stops a pasted job description from becoming a permanent
    200 KB "role" in the user's tracker.
    """
    if value is None or isinstance(value, bool):
        return ""
    text = value if isinstance(value, str) else str(value)
    return " ".join(text.split())[:cap]


def _clean_note(value):
    """Same idea, but line breaks survive - notes are written in paragraphs."""
    if value is None or isinstance(value, bool):
        return ""
    text = value if isinstance(value, str) else str(value)
    return text.replace("\r\n", "\n").replace("\r", "\n").strip()[:NOTE_CAP]


def _clean_fit(value):
    """0-100, or None when there was nothing usable there."""
    fit = _as_int(value)
    return None if fit is None else max(0, min(100, fit))


def _today():
    # At request time, not import time: a control panel left open overnight
    # would otherwise stamp tomorrow's saves with yesterday's date.
    return datetime.now().strftime("%Y-%m-%d")


def _item_id(company, role):
    """A short id derived from the job itself, never random.

    The same company+role always hashes to the same id, which is what makes
    the /bulk dedupe a set lookup and makes re-saving a job you already have
    a no-op instead of a second row.
    """
    seed = company.strip().lower() + "|" + role.strip().lower()
    return hashlib.sha1(seed.encode("utf-8")).hexdigest()[:12]


def _new_item(company, role, link="", location="", fit=None, status="saved"):
    return {
        "id": _item_id(company, role),
        "company": company,
        "role": role,
        "link": _clean(link),
        "location": _clean(location),
        "fit": _clean_fit(fit),
        "status": status,
        "note": "",
        "added": _today(),
    }


def _shape_item(raw):
    """Force one stored row into the shape the API promises, or drop it.

    tracker.json is a plain file the user can open and edit, so nothing read
    back out of it is trusted to have the right keys or the right types.
    """
    if not isinstance(raw, dict):
        return None
    company = _clean(raw.get("company"))
    role = _clean(raw.get("role"))
    if not company and not role:
        return None
    # Half a row is still the user's row. Dropping it here would not just hide
    # it from the board, it would erase it for good on the very next write, so
    # the missing half is labelled instead.
    company = company or "(unknown company)"
    role = role or "(unknown role)"
    status = _clean(raw.get("status"), 20).lower()
    return {
        "id": _clean(raw.get("id"), 40) or _item_id(company, role),
        "company": company,
        "role": role,
        "link": _clean(raw.get("link")),
        "location": _clean(raw.get("location")),
        "fit": _clean_fit(raw.get("fit")),
        "status": status if status in STATUSES else "saved",
        "note": _clean_note(raw.get("note")),
        "added": _clean(raw.get("added"), 10) or _today(),
    }


def _tracker_items(uid=None):
    """The saved list, cleaned. Callers hold _tracker_lock around this."""
    data = _read_json(_tracker_path(uid), None)
    raw = data.get("items") if isinstance(data, dict) else data
    if not isinstance(raw, list):
        return []
    items, seen = [], set()
    for row in raw:
        item = _shape_item(row)
        if item is None:
            continue
        # A repeated id must not silently delete the row behind it: two rows
        # sharing an id (a hand-copied block in tracker.json, say) would lose
        # one for good the next time this list is written back. Re-id the
        # later one instead so update/delete can still address it.
        if item["id"] in seen:
            base, n = item["id"][:10], 2
            while f"{base}-{n}" in seen and n < 1000:
                n += 1
            item["id"] = f"{base}-{n}"
            if item["id"] in seen:
                continue
        seen.add(item["id"])
        items.append(item)
    return items


def _write_tracker(items):
    """Write a temp file, then swap it in. Caller holds _tracker_lock.

    Losing power halfway through leaves the old tracker intact instead of a
    truncated one. The temp file sits beside the real file on purpose:
    os.replace is only atomic within a single filesystem.
    """
    target = _tracker_path()
    tmp = target + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump({"items": items}, f, indent=2, ensure_ascii=False)
            f.write("\n")
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, target)
    except OSError:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise


def _tracker_counts(items):
    counts = {name: 0 for name in STATUSES}
    for item in items:
        if item["status"] in counts:
            counts[item["status"]] += 1
    return counts


def _tracker_count():
    """For /api/state. Never raises: no file at all simply means zero."""
    try:
        with _tracker_lock:
            return len(_tracker_items())
    except Exception:
        return 0


# --------------------------------------------------------------------------
# run state (one search at a time)
# --------------------------------------------------------------------------

_lock = threading.Lock()
_lines = []            # every log line the subprocess has printed
_proc = None
_run = {
    "running": False,
    "done": False,
    "error": None,
    "result": None,
    "pct": 0,           # high-water mark, never moves backwards
    "stage": "Idle",
}

LOCK_PATH = os.path.join(PROJECT_DIR, "jobhunt.lock")


def _pid_alive(pid):
    """True when a process with this PID exists right now.

    tasklist rather than os.kill on Windows: kill(pid, 0) is unreliable
    there, and tasklist ships with every Windows. The CSV form quotes the
    PID, so a check for `"123"` cannot accidentally match PID 1234.
    """
    if not isinstance(pid, int) or pid <= 0:
        return False
    if os.name == "nt":
        try:
            out = subprocess.run(
                ["tasklist", "/FI", f"PID eq {pid}", "/NH", "/FO", "CSV"],
                capture_output=True, text=True, timeout=10,
                creationflags=CREATE_NO_WINDOW,
            )
        except (OSError, subprocess.SubprocessError):
            return False
        return f'"{pid}"' in (out.stdout or "")
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except OSError:
        return True                  # exists, just not ours to signal
    return True


def _external_run_pid():
    """The PID in jobhunt.lock when that run is alive and is not our own.

    jobhunt.py writes jobhunt.lock (containing its PID) for the duration of
    every run - including CLI and scheduled runs this panel never started
    and whose state the in-memory "running" flag knows nothing about. A lock
    whose PID is dead is a leftover from a crash and is ignored; a lock that
    belongs to the subprocess this UI itself launched is not "elsewhere".
    """
    try:
        with open(LOCK_PATH, encoding="utf-8-sig", errors="replace") as f:
            pid = _as_int(f.read(64))
    except OSError:
        return None
    if pid is None:
        return None
    proc = _proc
    if proc is not None and proc.poll() is None and proc.pid == pid:
        return None                  # the run this UI is already streaming
    return pid if _pid_alive(pid) else None


VERIFY_RE = re.compile(r"verified\s+(\d+)\s*/\s*(\d+)", re.I)
RANKED_RE = re.compile(r"ranked\s+(\d+)\s*/\s*(\d+)", re.I)
SCORING_RE = re.compile(r"scoring\s+\d+\s+postings", re.I)

STAGE_FETCH = "Reading company job boards"
STAGE_FILTER = "Filtering to US roles"
STAGE_SCORE = "Scoring against your profile"
STAGE_VERIFY = "Verifying every link is live"
STAGE_RANK = "Ranking with Gemini"
STAGE_WRITE = "Writing your Excel file"


def _progress_for(line):
    """(pct, stage) for one log line, or None when the line says nothing.

    Ordered most specific first, because jobhunt.py's markers overlap
    ("AI ranked 572/572" is also a "ranked N/M" line).
    """
    low = line.lower()

    if "run complete" in low:
        return 100, "Done"
    if "wrote " in low:
        return 98, STAGE_WRITE

    m = RANKED_RE.search(low)
    if m:
        done, total = int(m.group(1)), int(m.group(2))
        frac = min(1.0, done / total) if total else 0.0
        return 74 + int(22 * frac), STAGE_RANK
    if SCORING_RE.search(low):
        return 74, STAGE_RANK
    if "applyable," in low:
        return 72, STAGE_RANK

    m = VERIFY_RE.search(low)
    if m:
        done, total = int(m.group(1)), int(m.group(2))
        frac = min(1.0, done / total) if total else 0.0
        return 38 + int(32 * frac), STAGE_VERIFY
    if "opening and verifying" in low:
        return 38, STAGE_VERIFY

    if "us matches after scoring" in low:
        return 34, STAGE_SCORE
    if "total postings fetched" in low:
        return 28, STAGE_FILTER
    if "postings from company boards" in low:
        return 20, STAGE_FETCH
    if "pulling" in low and "boards" in low:
        return 8, STAGE_FETCH
    return None


def _note_progress_locked(line):
    """Caller holds _lock. pct only ever moves up."""
    hit = _progress_for(line)
    if not hit:
        return
    pct, stage = hit
    if pct >= _run["pct"]:
        _run["pct"] = pct
        _run["stage"] = stage


def _reader(proc):
    """Drain the subprocess stdout line by line until it exits."""
    try:
        for raw in proc.stdout:
            line = raw.rstrip("\r\n")
            if not line.strip():
                continue
            with _lock:
                _lines.append(line)
                _note_progress_locked(line)
    except Exception as e:                        # the pipe died mid-run
        with _lock:
            _lines.append(f"! log stream ended: {type(e).__name__}: {e}")
    finally:
        try:
            proc.stdout.close()
        except Exception:
            pass

        # Everything from here must not raise: if it did, "running" would stay
        # True forever and both /api/run and the UI would be wedged until the
        # server is restarted.
        error, result = None, None
        try:
            code = proc.wait()
        except Exception as e:
            code = -1
            error = f"Lost track of the search process ({type(e).__name__}: {e})."

        if error is None and code != 0:
            with _lock:
                tail = list(_lines[-6:])
            detail = "\n".join(tail).strip()
            error = (f"The search stopped (exit code {code})."
                     + (f"\n{detail}" if detail else ""))

        if error is None:
            try:
                result = _read_excel_summary(_excel_path())
            except Exception:
                result = None
            if result is None:
                error = ("The run finished but no spreadsheet could be read. "
                         "Check the log above.")

        with _lock:
            _run["running"] = False
            _run["done"] = True
            _run["error"] = error
            _run["result"] = result
            if error:
                _run["stage"] = "Stopped"
            else:
                _run["pct"] = 100
                _run["stage"] = "Done"


def _start_run():
    """Launch jobhunt.py. Returns an error string, or None on success."""
    global _proc

    env = dict(os.environ)
    env["PYTHONUNBUFFERED"] = "1"          # so log lines arrive as they happen
    env["PYTHONIOENCODING"] = "utf-8"
    key = _env_key()
    if key:
        env["GEMINI_API_KEY"] = key

    kwargs = {}
    if os.name == "nt":
        kwargs["creationflags"] = CREATE_NO_WINDOW

    try:
        proc = subprocess.Popen(
            [sys.executable, "jobhunt.py", "run"],
            cwd=PROJECT_DIR,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            env=env,
            **kwargs,
        )
    except Exception as e:
        return f"Could not start the search: {type(e).__name__}: {e}"

    _proc = proc
    try:
        threading.Thread(target=_reader, args=(proc,), daemon=True).start()
    except Exception as e:
        # Nobody would ever clear "running" without the reader, so refuse the
        # run outright instead of wedging the UI.
        try:
            proc.kill()
        except Exception:
            pass
        _proc = None
        return f"Could not watch the search: {type(e).__name__}: {e}"
    return None


# --------------------------------------------------------------------------
# routes
# --------------------------------------------------------------------------

@app.after_request
def _no_cache(resp):
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    return resp


@app.errorhandler(413)
def _too_large(_e):
    return jsonify(ok=False, error="That file is too large. Keep it under 10 MB."), 200


def _send_page(name):
    path = os.path.join(STATIC_DIR, name)
    if not os.path.exists(path):
        return (f"<h1>static/{name} is missing</h1>"
                f"<p>The backend is running. Put the page in "
                f"<code>{STATIC_DIR}</code> and refresh.</p>"), 200
    resp = send_from_directory(STATIC_DIR, name)
    resp.headers["Cache-Control"] = "no-store"
    return resp


def _tracker_items_for(uid):
    """Another account's tracker, for the admin view only."""
    try:
        with _tracker_lock:
            return _tracker_items(uid)
    except Exception:
        return []


@app.route("/")
@login_required
def index():
    if not os.path.exists(os.path.join(STATIC_DIR, "index.html")):
        return ("<h1>static/index.html is missing</h1>"
                "<p>The backend is running. Put the page in "
                f"<code>{STATIC_DIR}</code> and refresh.</p>"), 200
    return send_from_directory(STATIC_DIR, "index.html")


@app.route("/static/<path:filename>")
def static_files(filename):
    return send_from_directory(STATIC_DIR, filename)


# ---- key ------------------------------------------------------------------

_key_test_lock = threading.Lock()

TEST_PROMPT = (
    'Reply with JSON exactly like {"reply": "<one short friendly sentence>"} '
    "where the sentence confirms, in the first person, that you are connected "
    "and ready to rank job postings. Keep it under 20 words."
)


def _http_error_message(e):
    if e.code in (401, 403):
        return "Key rejected or this Google account is blocked from Gemini."
    if e.code == 404:
        return "That model is not available to this key."
    if e.code == 429:
        return "Rate limited - wait a moment and try again."
    if e.code == 400:
        return "Key rejected - check you pasted the whole key with no spaces."
    detail = ""
    try:
        body = json.loads(e.read().decode("utf-8", "replace"))
        detail = str(body.get("error", {}).get("message", ""))[:160]
    except Exception:
        detail = ""
    return f"Google returned {e.code}." + (f" {detail}" if detail else "")


def _unwrap_reply(raw):
    """_call asks Gemini for JSON, so pull the sentence back out of it."""
    reply = (raw or "").strip()
    try:
        parsed = json.loads(reply)
    except ValueError:
        return reply
    if isinstance(parsed, dict):
        for k in ("reply", "message", "text", "response"):
            v = parsed.get(k)
            if isinstance(v, str) and v.strip():
                return v.strip()
        first = next((v for v in parsed.values() if isinstance(v, str)), "")
        return first.strip() or reply
    if isinstance(parsed, list) and parsed and isinstance(parsed[0], str):
        return parsed[0].strip()
    if isinstance(parsed, str):
        return parsed.strip()
    return reply


@app.route("/api/key/test", methods=["POST"])
@login_required
def key_test():
    key = ((request.get_json(silent=True) or {}).get("key") or "").strip()
    if not key:
        return jsonify(ok=False, error="Paste a key first.")

    try:
        import ai_rank                   # the single Gemini code path
    except Exception as e:
        # The contract says this endpoint never 500s, so a broken import has
        # to come back as a readable ok:false too.
        return jsonify(ok=False,
                       error=f"ai_rank.py could not be loaded ({type(e).__name__}: {e}).")

    with _key_test_lock:
        # _call retries rate limits for about 90 seconds and waits 180s per
        # attempt, far too long for a button press. Ask for a short, cheap
        # probe instead; the RETRIES patch is the fallback for an ai_rank that
        # does not take those keywords.
        original = getattr(ai_rank, "RETRIES", 5)
        try:
            ai_rank.RETRIES = 2
            try:
                raw = ai_rank._call(key, TEST_PROMPT, retries=2, timeout=30)
            except TypeError:
                raw = ai_rank._call(key, TEST_PROMPT)
        except urllib.error.HTTPError as e:
            return jsonify(ok=False, error=_http_error_message(e))
        except urllib.error.URLError as e:
            return jsonify(ok=False,
                           error=f"Could not reach Google ({e.reason}). "
                                 "Check your internet connection.")
        except TimeoutError:
            return jsonify(ok=False, error="Google did not answer in time. Try again.")
        except Exception as e:
            return jsonify(ok=False, error=f"{type(e).__name__}: {e}")
        finally:
            ai_rank.RETRIES = original

    reply = _unwrap_reply(raw)
    if not reply:
        return jsonify(ok=False, error="The model answered with nothing. Try again.")
    return jsonify(ok=True, reply=reply)


@app.route("/api/key/save", methods=["POST"])
@login_required
def key_save():
    key = ((request.get_json(silent=True) or {}).get("key") or "").strip()
    if not key:
        return jsonify(ok=False, error="Paste a key first.")
    try:
        _save_env_key(key)
    except OSError as e:
        return jsonify(ok=False, error=f"Could not write .env: {e}")
    os.environ["GEMINI_API_KEY"] = key
    track("key_saved")        # live process picks it up now
    return jsonify(ok=True)


# ---- resume ---------------------------------------------------------------

# "locations" is deliberately NOT here: config.json's locations encode where
# the user chose to search, which no resume can know - the model has invented
# San Francisco/Seattle "preferences" before. They are never merged.
PROFILE_CONFIG_KEYS = ("target_titles", "secondary_titles",
                       "exclude_title_words", "skills")


def _merge_profile_into_config(profile):
    """Let the new resume ADD scorer keywords - never replace the user's own.

    Case-insensitive union per list: whatever config.json already has stays
    first, in its original order, and only genuinely new resume entries are
    appended after it. "companies" and "output_dir" are never touched, and
    every other key in config.json is left exactly as is.
    """
    # Read and write inside one lock, and swap the file in atomically: this
    # used to truncate config.json in place, so a failure mid-write - or a
    # threshold save arriving at the same moment - could destroy the whole
    # config, companies list and all.
    with _config_lock:
        cfg = _read_json(CONFIG_PATH, None)
        if not isinstance(cfg, dict):
            return
        changed = False
        for key in PROFILE_CONFIG_KEYS:
            values = profile.get(key)
            if not isinstance(values, list):
                continue
            existing = cfg.get(key)
            merged, seen = [], set()
            for source in (existing if isinstance(existing, list) else [],
                           values):
                for v in source:
                    if not isinstance(v, str):
                        continue
                    v = v.strip()
                    if v and v.lower() not in seen:
                        seen.add(v.lower())
                        merged.append(v)
            if merged and merged != existing:
                cfg[key] = merged
                changed = True
        if not changed:
            return
        _save_config( cfg)


@app.route("/api/resume", methods=["POST"])
@login_required
def resume_upload():
    upload = request.files.get("resume")
    if upload is None or not (upload.filename or "").strip():
        return jsonify(ok=False, error="No file received. Pick a resume first.")

    ext = os.path.splitext(upload.filename)[1].lower()
    if ext not in ALLOWED_EXT:
        return jsonify(ok=False, error="Only .pdf, .docx or .txt files work here.")

    try:
        upload.stream.seek(0, os.SEEK_END)
        size = upload.stream.tell()
        upload.stream.seek(0)
    except (OSError, ValueError):
        size = 0
    if size > MAX_RESUME_BYTES:
        return jsonify(ok=False, error="That file is too large. Keep it under 10 MB.")

    safe = secure_filename(upload.filename) or ("resume" + ext)
    if not safe.lower().endswith(ext):
        safe += ext
    path = os.path.join(UPLOAD_DIR, safe)
    try:
        upload.save(path)
    except OSError as e:
        return jsonify(ok=False, error=f"Could not save the upload: {e}")

    try:
        import resume_profile             # written alongside this file
    except Exception as e:
        return jsonify(ok=False,
                       error=f"resume_profile.py could not be loaded ({e}).")

    try:
        profile = resume_profile.build_profile(path)
    except Exception as e:
        return jsonify(ok=False, error=f"Could not read that resume: {e}")

    if not isinstance(profile, dict) or not profile:
        return jsonify(ok=False, error="The resume reader returned nothing usable.")

    # Which resume is currently driving the ranking. The uploaded file stays in
    # uploads/ until the next upload replaces it, so "what am I applying with?"
    # is answerable later -- a profile alone cannot tell you which PDF produced it.
    track("resume_upload", {"file": safe})
    profile["resume_file"] = safe
    profile["resume_size"] = int(size)
    profile["resume_uploaded_at"] = datetime.now().strftime("%Y-%m-%d %H:%M")

    _ensure_dir(os.path.dirname(PROFILE_PATH))
    try:
        with open(PROFILE_PATH, "w", encoding="utf-8") as f:
            json.dump(profile, f, indent=2, ensure_ascii=False)
            f.write("\n")
    except OSError as e:
        return jsonify(ok=False, error=f"Could not write profile.json: {e}")

    try:
        _merge_profile_into_config(profile)
    except Exception as e:
        # A failed merge must not lose the profile we just built.
        return jsonify(ok=True, profile=profile,
                       warning=f"Profile saved, but config.json was not updated: {e}")

    return jsonify(ok=True, profile=profile)


# ---- run / progress -------------------------------------------------------

@app.route("/api/run", methods=["POST"])
@login_required
def run_search():
    if IS_SERVERLESS:
        # A search takes 7-15 minutes and writes files. A serverless function is
        # capped at 60-300s on an ephemeral filesystem, so this cannot work here
        # however it is dressed up. Say so plainly rather than starting a
        # subprocess that gets killed halfway with a confusing error.
        return jsonify(ok=False, error=(
            "Searches cannot run on this hosted deployment - one takes 7-15 "
            "minutes and a serverless function is cut off after 60-300 seconds. "
            "Run it on your own machine with `python jobhunt.py run`, or a VM, "
            "and the results appear here. See DEPLOY.md."))
    with _lock:
        if _run["running"]:
            return jsonify(ok=False, error="already running")

    # A CLI `python jobhunt.py run` or a scheduled task never sets our
    # in-memory flag, so jobhunt.lock is the only warning that a second run
    # would trample the same spreadsheet. Checked outside _lock: the check
    # shells out to tasklist, and the reader thread takes _lock per log line.
    pid = _external_run_pid()
    if pid is not None:
        return jsonify(
            ok=False,
            error=f"a search is already running elsewhere (PID {pid})")

    with _lock:
        if _run["running"]:          # somebody else won the race just now
            return jsonify(ok=False, error="already running")
        _lines.clear()
        _run.update(running=True, done=False, error=None, result=None,
                    pct=0, stage="Starting up")

    track("run_started")
    error = _start_run()
    if error:
        with _lock:
            _run.update(running=False, done=True, error=error, stage="Stopped")
        return jsonify(ok=False, error=error)
    return jsonify(ok=True)


@app.route("/api/progress")
@login_required
def progress():
    try:
        since = max(0, int(request.args.get("since", 0)))
    except (TypeError, ValueError):
        since = 0
    with _lock:
        total = len(_lines)
        # A cursor past the end means the buffer was reset under this client
        # (a new run started, or a second tab kicked one off). Replaying from
        # the top is the only way it does not silently skip the new run's log.
        start = since if since <= total else 0
        return jsonify(
            lines=_lines[start:],
            next=total,
            running=_run["running"],
            done=_run["done"],
            error=_run["error"],
            pct=_run["pct"],
            stage=_run["stage"],
            result=_run["result"],
        )


# ---- state / download -----------------------------------------------------

@app.route("/api/state")
@login_required
def state():
    profile = _read_json(PROFILE_PATH, None)
    if not isinstance(profile, dict) or not profile:
        profile = None
    excel = _excel_info()
    with _lock:
        running = _run["running"]
    u = me() or {}
    # The setup steps (API key, resume, run) only make sense where a search can
    # actually happen: the machine with the code, driven by whoever owns it.
    # On the hosted site nobody can run one, so showing three dead steps just
    # invites clicks that cannot work.
    show_setup = (not IS_SERVERLESS) and bool(u.get("is_admin"))
    return jsonify(
        show_setup=show_setup,
        serverless=IS_SERVERLESS,
        is_admin=bool(u.get("is_admin")),
        has_key=bool(_env_key()),
        has_profile=profile is not None,
        has_excel=excel is not None,
        running=running,
        external_run=_external_run_pid() is not None,
        profile=profile,
        excel=excel,
        tracker_count=_tracker_count(),
    )


@app.route("/api/download")
@login_required
def download():
    path = _excel_path()
    if not os.path.exists(path):
        blob = storage.get_bytes("workbook") if storage.enabled() else None
        if blob:
            from flask import Response
            return Response(blob, mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"),
                headers={"Content-Disposition":
                         f'attachment; filename="{EXCEL_NAME}"'})
        return jsonify(ok=False,
                       error="No spreadsheet yet - run a search first."), 404
    try:
        return send_from_directory(os.path.dirname(path), os.path.basename(path),
                                   as_attachment=True, download_name=EXCEL_NAME)
    except OSError as e:
        # Windows: jobhunt.py is mid-rewrite, or Excel has the file locked.
        return jsonify(ok=False,
                       error=f"The spreadsheet is busy right now ({e.strerror or e}). "
                             "Close it in Excel or wait for the run to finish, "
                             "then try again."), 503


# ---- tracker --------------------------------------------------------------

BAD_STATUS = "Status must be one of: " + ", ".join(STATUSES) + "."
NOT_FOUND = "That job is not in the tracker - it may already be deleted."


@app.route("/api/tracker")
@login_required
def tracker_list():
    with _tracker_lock:
        items = _tracker_items()
    return jsonify(ok=True, items=items, counts=_tracker_counts(items))


@app.route("/api/tracker", methods=["POST"])
@login_required
def tracker_add():
    body = request.get_json(silent=True) or {}
    company = _clean(body.get("company"))
    role = _clean(body.get("role"))
    if not company or not role:
        return jsonify(ok=False, error="A company and a role are both required.")

    status = _clean(body.get("status"), 20).lower() or "saved"
    if status not in STATUSES:
        return jsonify(ok=False, error=BAD_STATUS)

    fresh = _new_item(company, role, body.get("link"), body.get("location"),
                      body.get("fit"), status)

    with _tracker_lock:
        items = _tracker_items()
        item = next((i for i in items if i["id"] == fresh["id"]), None)
        if item is None:
            item = fresh
            items.append(item)
        else:
            # Saving the same job twice is idempotent: it must not wipe the
            # note already written on it or move its "added" date forward, so
            # only blanks get filled in.
            for key in ("link", "location"):
                if fresh[key] and not item[key]:
                    item[key] = fresh[key]
            if item["fit"] is None:
                item["fit"] = fresh["fit"]
            if body.get("status") is not None:
                item["status"] = status
        try:
            _write_tracker(items)
        except OSError as e:
            return jsonify(ok=False, error=f"Could not write tracker.json: {e}")
    track("job_tracked", {"company": item.get("company"), "role": item.get("role")})
    return jsonify(ok=True, item=item)


@app.route("/api/tracker/bulk", methods=["POST"])
@login_required
def tracker_bulk():
    """Save a whole shortlist at once, skipping anything already tracked."""
    rows = (request.get_json(silent=True) or {}).get("items")
    if not isinstance(rows, list):
        return jsonify(ok=False, error='Send {"items": [ ... ]}.')

    added, skipped = 0, 0
    with _tracker_lock:
        items = _tracker_items()
        # The id is built from the lower-cased company+role, so an id already
        # in this set IS the case-insensitive duplicate check.
        seen = {i["id"] for i in items}
        for raw in rows:
            company = _clean(raw.get("company")) if isinstance(raw, dict) else ""
            role = _clean(raw.get("role")) if isinstance(raw, dict) else ""
            if not company or not role or _item_id(company, role) in seen:
                skipped += 1
                continue
            item = _new_item(company, role, raw.get("link"),
                             raw.get("location"), raw.get("fit"))
            seen.add(item["id"])
            items.append(item)
            added += 1
        if added:
            try:
                _write_tracker(items)
            except OSError as e:
                return jsonify(ok=False,
                               error=f"Could not write tracker.json: {e}")
    track("job_tracked", {"bulk": added})
    return jsonify(ok=True, added=added, skipped=skipped)


@app.route("/api/tracker/<item_id>/update", methods=["POST"])
@login_required
def tracker_update(item_id):
    body = request.get_json(silent=True) or {}
    status, note = body.get("status"), body.get("note")
    if status is not None:
        status = _clean(status, 20).lower()
        if status not in STATUSES:
            return jsonify(ok=False, error=BAD_STATUS)
    if note is not None:
        note = _clean_note(note)
    if status is None and note is None:
        return jsonify(ok=False,
                       error="Nothing to update - send a status or a note.")

    wanted = _clean(item_id, 40)
    with _tracker_lock:
        items = _tracker_items()
        item = next((i for i in items if i["id"] == wanted), None)
        if item is None:
            return jsonify(ok=False, error=NOT_FOUND), 404
        if status is not None:
            item["status"] = status
        if note is not None:
            item["note"] = note
        try:
            _write_tracker(items)
        except OSError as e:
            return jsonify(ok=False, error=f"Could not write tracker.json: {e}")
    track("job_status", {"id": item_id, "status": item.get("status")})
    return jsonify(ok=True, item=item)


@app.route("/api/tracker/<item_id>/delete", methods=["POST"])
@login_required
def tracker_delete(item_id):
    wanted = _clean(item_id, 40)
    with _tracker_lock:
        items = _tracker_items()
        kept = [i for i in items if i["id"] != wanted]
        if len(kept) == len(items):
            return jsonify(ok=False, error=NOT_FOUND), 404
        try:
            _write_tracker(kept)
        except OSError as e:
            return jsonify(ok=False, error=f"Could not write tracker.json: {e}")
    track("job_removed", {"id": item_id})
    return jsonify(ok=True)


# ---- results / threshold --------------------------------------------------

THRESHOLD_KEY = "min_ai_fit"
# Must match jobhunt.py's own fallback - cfg.get("min_ai_fit", 55). Any other
# number here and the panel would report a cutoff the search does not use.
DEFAULT_THRESHOLD = 55

FRESH_DAYS_KEY = "fresh_days"
# Must match jobhunt.py's own fallback - cfg.get("fresh_days", 2) - so the
# "Fresh (48h)" sheet and the UI's fresh badge agree on what counts as fresh.
DEFAULT_FRESH_DAYS = 2
MAX_FRESH_DAYS = 30

DEFAULT_LIMIT = 100
MAX_LIMIT = 1000

BUSY_EXCEL = "Close the spreadsheet in Excel and try again."


def _threshold():
    value = _as_int(_load_config().get(THRESHOLD_KEY))
    return DEFAULT_THRESHOLD if value is None else max(0, min(100, value))


def _fresh_days():
    value = _as_int(_load_config().get(FRESH_DAYS_KEY))
    return (DEFAULT_FRESH_DAYS if value is None
            else max(0, min(MAX_FRESH_DAYS, value)))


def _result_order(row):
    """Sort key: AI Fit high to low, then Days Old low to high.

    A blank or unreadable fit sinks below a real 0, exactly as before; a row
    with no readable age sinks to the end of its fit band. The sort stays
    stable, so equal rows keep the order the sheet listed them in.
    """
    fit = row["fit"] if isinstance(row["fit"], int) else -1
    days = row["days"]
    return (-fit, days is None, days if days is not None else 0)


def _link_target(cells, idx):
    """The real apply URL for one row.

    write_excel() in jobhunt.py writes the word "Apply" into the cell and
    hangs the URL off it as a hyperlink, so the cell VALUE is useless to the
    browser - the hyperlink target is where the job actually lives. A sheet
    that stored a plain URL instead still works.
    """
    for name in ("apply link", "link", "url"):
        i = idx.get(name)
        if i is None or i >= len(cells):
            continue
        link = getattr(cells[i], "hyperlink", None)
        target = getattr(link, "target", None) or (link if isinstance(link, str) else None)
        if target:
            return str(target).strip()[:600]
        value = cells[i].value
        if isinstance(value, str) and value.strip().lower().startswith("http"):
            return value.strip()[:600]
    return ""


def _archive_dates():
    """Every day that has a stored snapshot, oldest first."""
    if not storage.enabled():
        return []
    d = storage.get_json("results:dates", [])
    if d is storage.MISSING or not isinstance(d, list):
        return []
    return sorted(str(x) for x in d)


def _published_rows(frm=None, to=None):
    """Merge the dated snapshots covering [frm, to].

    Defaults to today. If today has no run yet, falls back to the most recent
    day that does -- showing an empty page because nobody has run a search since
    midnight would look broken rather than accurate.

    A role open on several days is one row, tagged with the days it appeared,
    so a week's range is a list of distinct jobs and not the same job seven
    times.
    """
    empty = {"rows": [], "threshold": None, "generated": "", "dates": [],
             "covering": "", "available": []}
    if not storage.enabled():
        return empty

    available = _archive_dates()
    if not available:
        # Nothing archived yet: fall back to the single live set, so an install
        # that published before dated snapshots existed still shows something.
        live = storage.get_json("results", None)
        if live is storage.MISSING or not isinstance(live, dict):
            return empty
        rows = live.get("rows") or []
        return {"rows": rows, "threshold": live.get("threshold"),
                "generated": live.get("generated", ""), "dates": [],
                "covering": "latest", "available": []}

    today = datetime.now().strftime("%Y-%m-%d")
    frm = (frm or "").strip()
    to = (to or "").strip()
    if not frm and not to:
        frm = to = today if today in available else available[-1]
    else:
        frm = frm or available[0]
        to = to or today

    wanted = [d for d in available if frm <= d <= to]
    if not wanted:
        return {**empty, "available": available,
                "covering": f"{frm} to {to}" if frm != to else frm}

    snapshots = storage.mget_json([f"results:{d}" for d in wanted])

    merged, threshold, generated = {}, None, ""
    for d in wanted:
        snap = snapshots.get(f"results:{d}")
        if not isinstance(snap, dict):
            continue
        if snap.get("threshold") is not None:
            threshold = snap["threshold"]
        generated = snap.get("generated", generated) or generated
        for r in snap.get("rows") or []:
            # The apply link identifies a posting; company+title does not.
            # Red Ventures lists the same "Associate Product Manager - AI" in
            # NYC and in Charlotte -- two real jobs she could apply to
            # separately, which a company+title key silently merges into one.
            key = str(r.get("link") or "").strip().lower()
            if not key:
                key = (str(r.get("company", "")).lower().strip() + "|"
                       + str(r.get("title", "")).lower().strip() + "|"
                       + str(r.get("location", "")).lower().strip())
            prev = merged.get(key)
            if prev is None:
                row = dict(r)
                row["seen"] = [d]
                merged[key] = row
            else:
                prev["seen"].append(d)
                # keep the best score and the freshest posting age we ever saw
                if isinstance(r.get("fit"), int) and (
                        not isinstance(prev.get("fit"), int) or r["fit"] > prev["fit"]):
                    prev["fit"] = r["fit"]
                    prev["why"] = r.get("why", prev.get("why", ""))
                if isinstance(r.get("days"), int) and (
                        not isinstance(prev.get("days"), int) or r["days"] < prev["days"]):
                    prev["days"] = r["days"]

    rows = list(merged.values())
    rows.sort(key=lambda r: (-(r.get("fit") if isinstance(r.get("fit"), int) else -1),
                             r.get("days") if isinstance(r.get("days"), int) else 9999,
                             str(r.get("company", ""))))
    return {"rows": rows, "threshold": threshold, "generated": generated,
            "dates": wanted, "available": available,
            "covering": wanted[0] if len(wanted) == 1 else f"{wanted[0]} to {wanted[-1]}"}


@app.route("/api/dates")
@login_required
def archive_dates():
    """Which days have data, so the picker can bound itself."""
    available = _archive_dates()
    return jsonify(ok=True, dates=available,
                   today=datetime.now().strftime("%Y-%m-%d"),
                   latest=available[-1] if available else "")


@app.route("/api/results")
@login_required
def results():
    """The Ready to Apply sheet as JSON, best fit first."""
    try:
        limit = int(request.args.get("limit", DEFAULT_LIMIT))
    except (TypeError, ValueError):
        limit = DEFAULT_LIMIT
    limit = max(1, min(MAX_LIMIT, limit))
    fresh_days = _fresh_days()

    # The dated archive comes first whenever there is one. It is the only source
    # that can answer "what was open last Tuesday" -- the local workbook holds a
    # single run and would silently ignore a date range, returning today's rows
    # under yesterday's heading. The workbook is the fallback for a machine with
    # no shared storage configured.
    path = _excel_path()
    if _archive_dates() or not os.path.exists(path):
        merged = _published_rows(request.args.get("from"), request.args.get("to"))
        if merged["rows"]:
            return jsonify(ok=True, rows=merged["rows"][:limit],
                           total=len(merged["rows"]),
                           threshold=merged["threshold"],
                           generated=merged["generated"],
                           dates=merged["dates"],
                           covering=merged["covering"],
                           available=merged["available"],
                           source="published")
        if not os.path.exists(path):
            return jsonify(ok=False, error=(
                "No results yet. Run a search on the machine that has the app "
                "installed - `python jobhunt.py run` - and they appear here."))
        # An archive exists but this range is empty: say so plainly instead of
        # falling through and showing another day's rows as if they were these.
        if request.args.get("from") or request.args.get("to"):
            return jsonify(ok=True, rows=[], total=0,
                           dates=[], covering=merged["covering"],
                           available=merged["available"], source="published")

    wb, rows = None, []
    try:
        # Read the workbook fresh on every call, never _excel_info()'s cached
        # summary: the list has to change the moment a run rewrites the file.
        # Not read_only=True: that mode throws hyperlinks away, and the
        # hyperlink is the only place the apply URL exists.
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = _ready_sheet(wb)
        if ws is None:
            return jsonify(ok=False,
                           error='That spreadsheet has no "' + READY_SHEET +
                                 '" sheet. Run a search to rebuild it.')

        sheet_rows = ws.iter_rows()
        # Columns are found by header name, never by fixed position: the
        # COLUMNS list in jobhunt.py gets reordered, and a hard-coded index
        # would then quietly read salaries into the title field.
        header = next(sheet_rows, None)
        idx = _header_index([c.value for c in header] if header else ())

        for cells in sheet_rows:
            row = [c.value for c in cells]
            if _blank_row(row):
                continue
            days = _as_int(_cell(row, idx, "days old"))
            entry = {
                "fit": _clean_fit(_cell(row, idx, "ai fit")),
                "score": _as_int(_cell(row, idx, "score")) or 0,
                "company": str(_cell(row, idx, "company")),
                "title": str(_cell(row, idx, "title")),
                "location": str(_cell(row, idx, "location")),
                "why": str(_cell(row, idx, "ai verdict", "why it matched")),
                "link": _link_target(cells, idx),
                "level": str(_cell(row, idx, "level", "seniority")),
                "days": days,
                "fresh": days is not None and days <= fresh_days,
            }
            if "eligibility" in idx:
                entry["eligibility"] = str(_cell(row, idx, "eligibility"))
            rows.append(entry)
    except PermissionError:
        # Windows: the user has the workbook open, or a run is mid-rewrite.
        return jsonify(ok=False, error=BUSY_EXCEL)
    except Exception as e:
        # This endpoint never 500s - the page shows the message instead.
        return jsonify(ok=False,
                       error=f"Could not read the spreadsheet "
                             f"({type(e).__name__}: {e}).")
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass

    rows.sort(key=_result_order)
    return jsonify(ok=True, rows=rows[:limit], total=len(rows),
                   threshold=_threshold(), fresh_days=fresh_days)


# --------------------------------------------------------------------------
# sign in / sign up
# --------------------------------------------------------------------------

@app.route("/login")
def login_page():
    if AUTH_DISABLED or auth.current_user():
        return redirect("/")
    return _send_page("login.html")


@app.route("/api/auth/config")
def auth_config():
    """What the login page should offer. Never leaks the client secret."""
    # Without ADMIN_PASSWORD_HASH there is no local account at all, and the
    # login form would just say "wrong password" forever. Surface that instead.
    return jsonify(ok=True,
                   google=auth.google_configured(),
                   local=bool(auth.ensure_admin()),
                   serverless=IS_SERVERLESS,
                   auth_disabled=AUTH_DISABLED,
                   user=auth.public_user(auth.current_user()))


@app.route("/auth/google")
def auth_google():
    if not auth.google_configured():
        return redirect("/login?error=" + quote(
            "Google sign-in is not configured yet. Add GOOGLE_CLIENT_ID and "
            "GOOGLE_CLIENT_SECRET to .env, then restart."))
    try:
        return redirect(auth.google_auth_url())
    except Exception as e:
        return redirect("/login?error=" + quote(f"Could not start Google sign-in ({e})."))


@app.route("/auth/google/callback")
def auth_google_callback():
    err = request.args.get("error")
    if err:
        return redirect("/login?error=" + quote(f"Google returned: {err}"))

    state = request.args.get("state") or ""
    expected = session.pop("oauth_state", None)
    # Without this check any page could walk a signed-in user through a login
    # to an account the attacker controls.
    if not expected or not hmac.compare_digest(state, expected):
        return redirect("/login?error=" + quote(
            "That sign-in link expired or did not start here. Try again."))

    code = request.args.get("code")
    if not code:
        return redirect("/login?error=" + quote("Google sent no authorization code."))

    try:
        info = auth.google_exchange(code)
        user = auth.upsert_google_user(info)
    except ValueError as e:
        return redirect("/login?error=" + quote(str(e)))
    except Exception as e:
        return redirect("/login?error=" + quote(f"Sign-in failed ({type(e).__name__})."))

    auth.login_user(user)
    return redirect("/")


@app.route("/api/auth/local", methods=["POST"])
def auth_local():
    body = request.get_json(silent=True) or {}
    user = auth.check_local_login(body.get("username"), body.get("password"))
    if not user:
        # Deliberately one message for both cases: saying which half was wrong
        # tells an attacker when they have found a real username.
        return jsonify(ok=False, error="Wrong username or password.")
    auth.login_user(user)
    return jsonify(ok=True, user=auth.public_user(user))


@app.route("/api/auth/logout", methods=["POST"])
def auth_logout():
    auth.logout_user()
    return jsonify(ok=True)


@app.route("/api/me")
def api_me():
    u = me()
    if not u:
        return jsonify(ok=False, error="Not signed in.", auth_required=True), 401
    if not AUTH_DISABLED:
        auth.touch_seen(u["id"])
    return jsonify(ok=True, user=auth.public_user(u) if not AUTH_DISABLED else u)


@app.route("/api/activity", methods=["POST"])
@login_required
def api_activity():
    """Client-reported events (app opened, an apply link clicked)."""
    body = request.get_json(silent=True) or {}
    kind = str(body.get("type") or "").strip()
    allowed = {"open_app", "apply_click", "download"}
    if kind not in allowed:
        return jsonify(ok=False, error="Unknown event."), 400
    detail = body.get("detail")
    if isinstance(detail, dict):
        detail = {k: str(v)[:200] for k, v in list(detail.items())[:8]}
    track(kind, detail)
    return jsonify(ok=True)


# --------------------------------------------------------------------------
# admin
# --------------------------------------------------------------------------

@app.route("/admin")
@admin_required
def admin_page():
    return _send_page("admin.html")


@app.route("/api/admin/users")
@admin_required
def admin_users():
    users = auth.all_users()
    events = auth.read_events(limit=5000)
    by_user = {}
    for e in events:
        u = by_user.setdefault(e.get("user_id", ""), {"events": 0, "last": "", "runs": 0,
                                                      "tracked": 0, "applies": 0})
        u["events"] += 1
        u["last"] = u["last"] or e.get("ts", "")
        t = e.get("type")
        if t == "run_started":
            u["runs"] += 1
        elif t == "job_tracked":
            u["tracked"] += 1
        elif t == "apply_click":
            u["applies"] += 1

    out = []
    for u in users:
        stats = by_user.get(u.get("id"), {})
        pub = auth.public_user(u)
        pub.update(
            last_seen=u.get("last_seen", ""),
            login_count=int(u.get("login_count") or 0),
            events=stats.get("events", 0),
            runs=stats.get("runs", 0),
            tracked_events=stats.get("tracked", 0),
            apply_clicks=stats.get("applies", 0),
            tracker_count=len(_tracker_items_for(u.get("id"))),
        )
        out.append(pub)
    out.sort(key=lambda x: str(x.get("last_seen") or x.get("last_login") or ""), reverse=True)
    return jsonify(ok=True, users=out, total=len(out))


@app.route("/api/admin/activity")
@admin_required
def admin_activity():
    try:
        limit = int(request.args.get("limit", 300))
    except (TypeError, ValueError):
        limit = 300
    return jsonify(ok=True,
                   events=auth.read_events(limit=limit,
                                           user_id=request.args.get("user") or None,
                                           kind=request.args.get("type") or None,
                                           since=request.args.get("since") or None),
                   labels=auth.EVENT_LABELS)


@app.route("/api/admin/user/<uid>")
@admin_required
def admin_user(uid):
    u = auth.get_user(uid)
    if not u:
        return jsonify(ok=False, error="No such user."), 404
    return jsonify(ok=True,
                   user=auth.public_user(u),
                   tracker=_tracker_items_for(uid),
                   events=auth.read_events(limit=400, user_id=uid))


@app.route("/api/resume/file")
@login_required
def resume_file():
    """Open the resume that is currently attached, inline in the browser.

    The filename comes from profile.json and is reduced to a bare basename
    before it touches the filesystem: it originated in an upload, so treating
    it as a path would let "../../.env" walk out of uploads/.
    """
    profile = _read_json(PROFILE_PATH, None)
    name = (profile or {}).get("resume_file") if isinstance(profile, dict) else None
    if not name:
        return jsonify(ok=False, error="No resume is attached yet."), 404

    safe = os.path.basename(str(name))
    path = os.path.join(UPLOAD_DIR, safe)
    if not os.path.isfile(path):
        return jsonify(
            ok=False,
            error="That resume file is no longer on disk. Upload it again."), 404

    ext = os.path.splitext(safe)[1].lower()
    mime = {".pdf": "application/pdf",
            ".txt": "text/plain",
            ".docx": ("application/vnd.openxmlformats-officedocument"
                      ".wordprocessingml.document")}.get(ext, "application/octet-stream")
    # .docx cannot render in a browser tab, so let it download instead of
    # showing a wall of binary.
    inline = ext in (".pdf", ".txt")
    return send_from_directory(UPLOAD_DIR, safe, mimetype=mime,
                               as_attachment=not inline,
                               download_name=safe)


HISTORY_PATH = os.path.join(PROJECT_DIR, "history.json")


def _history_runs():
    data = _read_json(HISTORY_PATH, None)
    runs = data.get("runs") if isinstance(data, dict) else data
    if not isinstance(runs, list):
        return []
    return [r for r in runs if isinstance(r, dict)]


@app.route("/api/history")
@login_required
def history():
    """Every run ever, newest first, optionally windowed by date.

    ?from=YYYY-MM-DD & ?to=YYYY-MM-DD are inclusive; ?days=N is a shorthand for
    the last N days. Also returns per-day totals so the UI can show "yesterday
    600, today 8" without recomputing it client-side.
    """
    runs = _history_runs()

    frm = (request.args.get("from") or "").strip()
    to = (request.args.get("to") or "").strip()
    days = request.args.get("days")
    if days and not frm:
        try:
            n = max(0, min(3650, int(days)))
            frm = (datetime.now() - timedelta(days=n)).strftime("%Y-%m-%d")
        except (TypeError, ValueError):
            pass

    def in_window(r):
        d = str(r.get("date") or "")
        if frm and d < frm:
            return False
        if to and d > to:
            return False
        return True

    sel = [r for r in runs if in_window(r)]
    # newest first; ids are ISO timestamps so a string sort is chronological
    sel.sort(key=lambda r: str(r.get("id") or r.get("date") or ""), reverse=True)

    by_day = {}
    for r in sel:
        d = str(r.get("date") or "")
        if not d:
            continue
        day = by_day.setdefault(d, {"date": d, "runs": 0, "ready": 0, "new": 0,
                                    "total": 0, "best_ready": 0})
        day["runs"] += 1
        day["total"] += int(r.get("total") or 0)
        day["new"] += int(r.get("new") or 0)
        day["ready"] += int(r.get("ready") or 0)
        day["best_ready"] = max(day["best_ready"], int(r.get("ready") or 0))
    days_list = sorted(by_day.values(), key=lambda d: d["date"], reverse=True)

    today = datetime.now().strftime("%Y-%m-%d")
    return jsonify(ok=True, runs=sel, days=days_list,
                   total_runs=len(runs), today=today,
                   first_date=(min(str(r.get("date") or "") for r in runs) if runs else ""))


@app.route("/api/lastrun")
@login_required
def lastrun():
    """What the read-only view shows in place of the setup steps: how many
    roles the last search found, and when it ran."""
    runs = _history_runs()
    runs.sort(key=lambda r: str(r.get("id") or r.get("date") or ""), reverse=True)
    latest = runs[0] if runs else None

    published = storage.get_json("results", None) if storage.enabled() else None
    if published is storage.MISSING:
        published = None

    total = None
    if isinstance(published, dict):
        total = published.get("total")
    if total is None and latest:
        total = latest.get("ready")

    return jsonify(
        ok=True,
        roles=total,
        when=(published or {}).get("generated") or (latest or {}).get("finished", ""),
        date=(latest or {}).get("date", ""),
        time=(latest or {}).get("time", ""),
        new=(latest or {}).get("new"),
        threshold=(published or {}).get("threshold") or (latest or {}).get("threshold"),
        incomplete=bool((latest or {}).get("incomplete")),
        machine=(published or {}).get("machine", ""),
        runs=len(runs),
    )


@app.route("/api/threshold")
@login_required
def threshold_get():
    return jsonify(ok=True, value=_threshold(), fresh_days=_fresh_days())


@app.route("/api/threshold", methods=["POST"])
@login_required
def threshold_set():
    body = request.get_json(silent=True)
    if not isinstance(body, dict):
        body = {}
    has_value = "value" in body
    has_fresh = "fresh_days" in body
    if not has_value and not has_fresh:
        return jsonify(ok=False,
                       error='Send {"value": 0-100} and/or '
                             '{"fresh_days": 0-30}.')

    value = fresh = None
    if has_value:
        value = _as_int(body.get("value"))
        if value is None:
            return jsonify(ok=False, error="Send a number between 0 and 100.")
        value = max(0, min(100, value))
    if has_fresh:
        fresh = _as_int(body.get("fresh_days"))
        if fresh is None:
            return jsonify(ok=False,
                           error="fresh_days must be a number between 0 and 30.")
        fresh = max(0, min(MAX_FRESH_DAYS, fresh))

    # The whole read-modify-write is one step under the config lock, or a
    # resume upload landing at the same moment would write its own copy of
    # config.json from a snapshot taken before these keys were added.
    with _config_lock:
        cfg = _read_json(CONFIG_PATH, None)
        if not isinstance(cfg, dict):
            if os.path.exists(CONFIG_PATH):
                # It is there but unreadable. Refuse: overwriting it with a
                # near-empty file would throw away companies, output_dir and
                # the rest of the user's setup.
                return jsonify(ok=False,
                               error="config.json could not be read, so the "
                                     "setting was not saved. Fix or delete "
                                     "that file and try again.")
            cfg = {}          # no file at all - start one, nothing to lose
        # Only the posted keys change; companies, output_dir and every other
        # key survive.
        if value is not None:
            cfg[THRESHOLD_KEY] = value
        if fresh is not None:
            cfg[FRESH_DAYS_KEY] = fresh
        try:
            _save_config( cfg)
        except OSError as e:
            return jsonify(ok=False, error=f"Could not write config.json: {e}")
        saved_value = _as_int(cfg.get(THRESHOLD_KEY))
        saved_fresh = _as_int(cfg.get(FRESH_DAYS_KEY))
    saved_value = (DEFAULT_THRESHOLD if saved_value is None
                   else max(0, min(100, saved_value)))
    saved_fresh = (DEFAULT_FRESH_DAYS if saved_fresh is None
                   else max(0, min(MAX_FRESH_DAYS, saved_fresh)))
    return jsonify(ok=True, value=saved_value, fresh_days=saved_fresh)


# --------------------------------------------------------------------------

URL = "http://127.0.0.1:5000"


def _banner():
    line = "=" * 62
    print(line)
    print("  JOB HUNT  -  control panel")
    print(line)
    print(f"  Open this in your browser:   {URL}")
    print(f"  Project folder:              {PROJECT_DIR}")
    print(f"  Spreadsheets land in:        {_output_dir()}")
    print("  Leave this window open. Press Ctrl+C here to stop.")
    print(line, flush=True)


if __name__ == "__main__":
    _banner()
    try:
        threading.Timer(1.0, lambda: webbrowser.open(URL)).start()
    except Exception:
        pass

    # SECURITY: this app has no authentication and can replace the API key,
    # rewrite config, upload files and launch subprocesses. It must only ever
    # listen on localhost. To use it from a phone, tunnel it deliberately
    # (e.g. ssh) rather than binding to all interfaces.
    app.run(
        host="127.0.0.1",
        port=5000,
        threaded=True,
        debug=False
    )