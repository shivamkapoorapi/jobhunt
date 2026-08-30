"""
Where state lives.

Two backends behind one interface:

  files  - the default. JSON beside the code, exactly as before.
  redis  - Upstash over its HTTP REST API, used when KV_REST_API_URL and
           KV_REST_API_TOKEN are set.

The point is the split between the machine that CAN run a 15-minute search and
the host that can SERVE a page. Your PC runs the search and writes the results
here; the Vercel deployment reads them. Neither can see the other's disk, so a
shared store is the only thing that links them.

Design rules this module holds to:

- Redis is never assumed to work. Every call has a timeout and returns a
  sentinel on failure; callers fall back to the filesystem. A store being
  briefly unreachable must degrade to "slightly stale", never to "crashed" or,
  worse, "silently empty" -- an empty read is indistinguishable from real data
  loss to everything downstream, so `get_json` returns MISSING rather than the
  caller's default when the network fails.
- Writes go to BOTH backends when Redis is on and a local disk exists. The
  local copy stays authoritative for a machine working offline, and it means
  turning Redis off later loses nothing.
- Keys are namespaced under `jobhunt:` so this database can be shared.
"""

import base64
import json
import os
import threading

PREFIX = "jobhunt:"
TIMEOUT = 15

# Distinct from None: None is a legitimate stored value, and a caller's default
# is a legitimate answer to "not found". Neither should be confused with
# "the store did not answer", which is what this means.
MISSING = object()

# A value the store returned but that will not parse. Distinct from MISSING
# ("could not read") and from absent ("never written") -- silently skipping it
# would under-report a range as if the day had held nothing.
CORRUPT = object()

_session = None
_session_lock = threading.Lock()


def _creds():
    """Vercel's Upstash integration injects KV_*; a hand-rolled Upstash account
    uses UPSTASH_REDIS_REST_*. Accept either so this works both ways."""
    url = (os.environ.get("KV_REST_API_URL")
           or os.environ.get("UPSTASH_REDIS_REST_URL") or "").strip().rstrip("/")
    token = (os.environ.get("KV_REST_API_TOKEN")
             or os.environ.get("UPSTASH_REDIS_REST_TOKEN") or "").strip()
    return url, token


def enabled():
    url, token = _creds()
    return bool(url and token)


def describe():
    url, _ = _creds()
    if not url:
        return "local files"
    host = url.split("//", 1)[-1].split(".", 1)[0]
    return f"redis ({host})"


def _http():
    global _session
    with _session_lock:
        if _session is None:
            import requests
            _session = requests.Session()
        return _session


def _command(*args):
    """Run one Redis command over the REST API. Returns MISSING on any failure."""
    url, token = _creds()
    if not (url and token):
        return MISSING
    try:
        r = _http().post(
            url,
            headers={"Authorization": f"Bearer {token}",
                     "Content-Type": "application/json"},
            data=json.dumps([str(a) for a in args]).encode("utf-8"),
            timeout=TIMEOUT,
        )
        if r.status_code != 200:
            return MISSING
        return r.json().get("result")
    except Exception:
        return MISSING


def _set_raw(key, value):
    """SET with the value in the body, so large payloads are not URL-encoded."""
    url, token = _creds()
    if not (url and token):
        return False
    try:
        r = _http().post(
            f"{url}/set/{PREFIX}{key}",
            headers={"Authorization": f"Bearer {token}"},
            data=value.encode("utf-8") if isinstance(value, str) else value,
            timeout=TIMEOUT * 4,          # a workbook is ~600 KB
        )
        return r.status_code == 200
    except Exception:
        return False


def _get_raw(key):
    res = _command("GET", f"{PREFIX}{key}")
    return res


# --------------------------------------------------------------------------
# JSON
# --------------------------------------------------------------------------

def get_json(key, default=None):
    """Read one key. Returns `default` when genuinely absent, MISSING when the
    store could not be reached -- the caller decides what to do about that."""
    if not enabled():
        return MISSING
    raw = _get_raw(key)
    if raw is MISSING:
        return MISSING
    if raw is None:
        return default
    try:
        return json.loads(raw)
    except (ValueError, TypeError):
        return CORRUPT


VALUE_CAP = 1_000_000          # Upstash free tier caps a value at 1 MB


def set_json(key, value):
    """False on refusal or failure. Checks the size ceiling first, because a
    day's archive grows all run and 'the store said no' is otherwise opaque."""
    if not enabled():
        return False
    try:
        blob = json.dumps(value, ensure_ascii=False)
    except (TypeError, ValueError):
        return False
    if len(blob.encode("utf-8")) > VALUE_CAP:
        return False
    return _set_raw(key, blob)


def too_big(value):
    """(is_too_big, kilobytes) - so a caller can explain the refusal."""
    try:
        n = len(json.dumps(value, ensure_ascii=False).encode("utf-8"))
    except (TypeError, ValueError):
        return True, 0
    return n > VALUE_CAP, n // 1024


def delete(key):
    if not enabled():
        return False
    return _command("DEL", f"{PREFIX}{key}") is not MISSING


def keys(pattern="*"):
    res = _command("KEYS", f"{PREFIX}{pattern}")
    if res is MISSING or not isinstance(res, list):
        return []
    n = len(PREFIX)
    return [k[n:] if k.startswith(PREFIX) else k for k in res]


MGET_BATCH = 25


def mget_json(names):
    """Fetch many keys. Returns MISSING if ANY batch could not be read.

    Two things this must not do. It must not answer "unreachable" with an empty
    dict -- a caller cannot tell that from "the archive is empty", and would
    render a year of history as zero roles. And it must not ask for 400 keys in
    one request: that response is megabytes and blows the timeout, which is what
    produced the unreachable state to begin with. Hence batching.
    """
    if not enabled():
        return MISSING
    if not names:
        return {}
    out = {}
    for i in range(0, len(names), MGET_BATCH):
        chunk = names[i:i + MGET_BATCH]
        res = _command("MGET", *[f"{PREFIX}{n}" for n in chunk])
        if res is MISSING or not isinstance(res, list) or len(res) != len(chunk):
            # A short list would silently misalign the zip below.
            return MISSING
        for name, raw in zip(chunk, res):
            if raw is None:
                continue
            try:
                out[name] = json.loads(raw)
            except (ValueError, TypeError):
                # Corrupt, not absent. Report rather than quietly skip.
                out[name] = CORRUPT
    return out


# --------------------------------------------------------------------------
# binary (the .xlsx)
# --------------------------------------------------------------------------

def set_bytes(key, data):
    """Store bytes as base64. Upstash caps a value at 1 MB on the free plan and
    base64 inflates by ~33%, so this refuses early with a readable reason
    instead of letting the store return an opaque error."""
    if not enabled():
        return False, "storage not configured"
    encoded = base64.b64encode(data).decode("ascii")
    if len(encoded) > 1_000_000:
        return False, (f"workbook is {len(data)//1024} KB, too large for the "
                       f"1 MB free-tier limit once base64-encoded")
    ok = _set_raw(key, encoded)
    return ok, ("" if ok else "the store rejected the write")


def get_bytes(key):
    """MISSING if unreachable, None if genuinely absent - the download must not
    tell someone to re-run a search that already succeeded."""
    raw = _get_raw(key)
    if raw is MISSING:
        return MISSING
    if raw is None:
        return None
    try:
        return base64.b64decode(raw)
    except Exception:
        return None


# --------------------------------------------------------------------------
# key naming
# --------------------------------------------------------------------------

def key_for_user(uid, name):
    return f"user:{uid}:{name}"


def ping():
    """(ok, message) -- used by the CLI and the admin console."""
    if not enabled():
        return False, "not configured (set KV_REST_API_URL and KV_REST_API_TOKEN)"
    res = _command("PING")
    if res is MISSING:
        return False, "could not reach the store"
    return True, describe()
