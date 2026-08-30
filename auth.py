"""
Accounts, sessions and the activity log.

Two ways in:
  - Google sign-in (OAuth 2.0 authorization code flow, id_token verified
    against Google's public keys -- never trusted because it merely parsed).
  - A local admin account, for the owner.

Storage is flat JSON under data/. That is deliberate for a tool this size: it
is inspectable, diffable and needs no service to run. Every write goes through
_write_json_atomic under a lock, so a crash or two concurrent requests cannot
leave a half-written users.json behind.

ANALYTICS: this module records what each signed-in user does (page opens, runs,
jobs tracked, apply clicks) so the owner can see how the app is used. If this is
ever opened to real users, that belongs in a privacy policy -- most jurisdictions
expect the disclosure to exist somewhere, though not necessarily on the login
screen itself.
"""

import hashlib
import hmac
import json
import os
import secrets
import threading
import time

import storage
from datetime import datetime, timezone
from functools import wraps

from flask import jsonify, redirect, request, session

HERE = os.path.dirname(os.path.abspath(__file__))

# Serverless hosts give you a read-only deployment and a writable /tmp. Creating
# data/ beside the code at import time crashes the function before it can serve
# anything, so the store moves to /tmp there. /tmp does not survive between
# invocations -- accounts written there are lost. See DEPLOY.md.
IS_SERVERLESS = bool(os.environ.get("VERCEL") or os.environ.get("AWS_LAMBDA_FUNCTION_NAME"))
DATA_DIR = os.path.join("/tmp", "resumify", "data") if IS_SERVERLESS     else os.path.join(HERE, "data")
USERS_PATH = os.path.join(DATA_DIR, "users.json")
ACTIVITY_PATH = os.path.join(DATA_DIR, "activity.json")
SECRET_PATH = os.path.join(DATA_DIR, "secret.key")
USER_DIR = os.path.join(DATA_DIR, "users")

ACTIVITY_MAX = 20000          # keep the log bounded; oldest events drop first
PBKDF2_ROUNDS = 240_000

_lock = threading.Lock()

def _ensure_dir(path):
    """Best effort: a read-only filesystem must not stop this module importing."""
    try:
        os.makedirs(path, exist_ok=True)
        return True
    except OSError:
        return False


_ensure_dir(DATA_DIR)
_ensure_dir(USER_DIR)


# --------------------------------------------------------------------------
# json helpers
# --------------------------------------------------------------------------

def _store_key(path):
    """Map a file path to a store key, so both backends address the same thing.

    data/users.json            -> users
    data/activity.json         -> activity
    data/users/<uid>/x.json    -> user:<uid>:x
    """
    path = os.path.abspath(path)
    stem = os.path.splitext(os.path.basename(path))[0]
    parent = os.path.dirname(path)
    if os.path.basename(os.path.dirname(parent)) == "users":
        return "user:%s:%s" % (os.path.basename(parent), stem)
    return stem


def _read_json(path, default):
    """Shared store first, local file second.

    The fallback is deliberate and one-directional: if the store is unreachable
    we serve whatever this machine last wrote rather than reporting an empty
    account list, which downstream code cannot tell apart from "no accounts
    exist" and would happily overwrite.
    """
    if storage.enabled():
        got = storage.get_json(_store_key(path), storage.MISSING)
        if got is not storage.MISSING:
            return got
    try:
        with open(path, encoding="utf-8-sig") as f:
            return json.load(f)
    except (OSError, ValueError):
        return default


def _write_json_atomic(path, data):
    """Write both places. The local copy keeps an offline machine working and
    means switching the store off later loses nothing."""
    wrote_remote = storage.set_json(_store_key(path), data) if storage.enabled() else False
    try:
        tmp = path + ".tmp"
        _ensure_dir(os.path.dirname(path))
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, path)
    except OSError:
        # Read-only disk (serverless). Fine as long as the store took it.
        if not wrote_remote:
            raise


def _now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# --------------------------------------------------------------------------
# secret key
# --------------------------------------------------------------------------

def secret_key():
    """Signs the session cookie. Stable across restarts or everyone is logged
    out on every reload; generated once if absent so there is no default."""
    env = os.environ.get("FLASK_SECRET_KEY", "").strip()
    if env:
        return env
    existing = _read_json(SECRET_PATH, None)
    if isinstance(existing, dict) and existing.get("key"):
        return existing["key"]
    key = secrets.token_urlsafe(48)
    try:
        _write_json_atomic(SECRET_PATH, {"key": key})
    except OSError:
        pass
    return key


# --------------------------------------------------------------------------
# passwords
# --------------------------------------------------------------------------

def hash_password(password, salt=None):
    salt = salt or secrets.token_hex(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"),
                             salt.encode("utf-8"), PBKDF2_ROUNDS)
    return f"pbkdf2_sha256${PBKDF2_ROUNDS}${salt}${dk.hex()}"


def verify_password(password, stored):
    """Constant-time check. Returns False on any malformed hash rather than
    raising -- a corrupt record must read as 'wrong password', not a 500."""
    try:
        algo, rounds, salt, digest = str(stored).split("$", 3)
        if algo != "pbkdf2_sha256":
            return False
        dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"),
                                 salt.encode("utf-8"), int(rounds))
        return hmac.compare_digest(dk.hex(), digest)
    except (ValueError, TypeError):
        return False


# --------------------------------------------------------------------------
# users
# --------------------------------------------------------------------------

def _load_users():
    data = _read_json(USERS_PATH, None)
    users = data.get("users") if isinstance(data, dict) else data
    return [u for u in users if isinstance(u, dict)] if isinstance(users, list) else []


def _save_users(users):
    _write_json_atomic(USERS_PATH, {"users": users})


def all_users():
    with _lock:
        return _load_users()


def get_user(uid):
    for u in all_users():
        if u.get("id") == uid:
            return u
    return None


def _uid_for(provider, key):
    return hashlib.sha1(f"{provider}:{key}".encode("utf-8")).hexdigest()[:16]


def upsert_google_user(info):
    """Create or refresh a Google account. `info` is the verified id_token."""
    email = (info.get("email") or "").strip().lower()
    if not email:
        raise ValueError("Google did not return an email address.")
    uid = _uid_for("google", info.get("sub") or email)
    with _lock:
        users = _load_users()
        for u in users:
            if u.get("id") == uid:
                u["name"] = info.get("name") or u.get("name") or email
                u["picture"] = info.get("picture") or u.get("picture", "")
                u["last_login"] = _now()
                u["login_count"] = int(u.get("login_count") or 0) + 1
                try:
                    _save_users(users)
                except OSError:
                    pass
                return u
        user = {
            "id": uid,
            "email": email,
            "name": info.get("name") or email,
            "picture": info.get("picture") or "",
            "provider": "google",
            "is_admin": False,
            "created": _now(),
            "last_login": _now(),
            "last_seen": _now(),
            "login_count": 1,
        }
        users.append(user)
        try:
            _save_users(users)
        except OSError:
            pass
        return user


def ensure_admin():
    """The owner account. Password comes from ADMIN_PASSWORD_HASH (preferred)
    or ADMIN_PASSWORD; it is only ever stored hashed."""
    username = os.environ.get("ADMIN_USERNAME", "shivamkapoor").strip().lower()
    uid = _uid_for("local", username)
    with _lock:
        users = _load_users()
        for u in users:
            if u.get("id") == uid:
                return u
        pw_hash = os.environ.get("ADMIN_PASSWORD_HASH", "").strip()
        if not pw_hash:
            plain = os.environ.get("ADMIN_PASSWORD", "").strip()
            if not plain:
                return None          # nothing configured: no admin exists
            pw_hash = hash_password(plain)
        user = {
            "id": uid,
            "email": os.environ.get("ADMIN_EMAIL", username + "@local").strip(),
            "name": os.environ.get("ADMIN_NAME", "Shivam Kapoor").strip(),
            "picture": "",
            "provider": "local",
            "username": username,
            "password_hash": pw_hash,
            "is_admin": True,
            "created": _now(),
            "last_login": "",
            "last_seen": "",
            "login_count": 0,
        }
        users.append(user)
        try:
            _save_users(users)
        except OSError:
            # Read-only store (serverless cold start). The admin still works for
            # this invocation; it simply is not persisted.
            pass
        return user


def check_local_login(username, password):
    username = (username or "").strip().lower()
    ensure_admin()
    with _lock:
        users = _load_users()
        for u in users:
            if u.get("provider") == "local" and u.get("username") == username:
                if verify_password(password or "", u.get("password_hash") or ""):
                    u["last_login"] = _now()
                    u["login_count"] = int(u.get("login_count") or 0) + 1
                    try:
                        _save_users(users)
                    except OSError:
                        pass          # bookkeeping only - the login still stands
                    return u
                return None
    return None


def touch_seen(uid):
    with _lock:
        users = _load_users()
        changed = False
        for u in users:
            if u.get("id") == uid:
                u["last_seen"] = _now()
                changed = True
                break
        if changed:
            try:
                _save_users(users)
            except OSError:
                pass


def public_user(u):
    """What the browser is allowed to see. Never the password hash."""
    if not u:
        return None
    return {
        "id": u.get("id"), "email": u.get("email"), "name": u.get("name"),
        "picture": u.get("picture", ""), "provider": u.get("provider"),
        "is_admin": bool(u.get("is_admin")),
        "created": u.get("created", ""), "last_login": u.get("last_login", ""),
    }


# --------------------------------------------------------------------------
# per-user storage
# --------------------------------------------------------------------------

def user_dir(uid):
    d = os.path.join(USER_DIR, str(uid))
    _ensure_dir(d)
    return d


def user_file(uid, name):
    return os.path.join(user_dir(uid), name)


# --------------------------------------------------------------------------
# activity log
# --------------------------------------------------------------------------

EVENT_LABELS = {
    "login": "Signed in",
    "logout": "Signed out",
    "open_app": "Opened the app",
    "resume_upload": "Uploaded a resume",
    "run_started": "Started a search",
    "run_finished": "Search finished",
    "job_tracked": "Tracked a job",
    "job_removed": "Removed a tracked job",
    "job_status": "Moved a job",
    "apply_click": "Opened a job to apply",
    "download": "Downloaded the spreadsheet",
    "key_saved": "Saved an API key",
}


def log_event(user, kind, detail=None):
    """Append one activity event. Never raises: analytics must not be able to
    break the feature it is observing."""
    try:
        ev = {
            "ts": _now(),
            "epoch": int(time.time()),
            "user_id": (user or {}).get("id", ""),
            "email": (user or {}).get("email", ""),
            "name": (user or {}).get("name", ""),
            "type": str(kind),
            "detail": detail if isinstance(detail, (dict, str, int, float)) else str(detail),
        }
        with _lock:
            data = _read_json(ACTIVITY_PATH, None)
            events = data.get("events") if isinstance(data, dict) else data
            if not isinstance(events, list):
                events = []
            events.append(ev)
            if len(events) > ACTIVITY_MAX:
                events = events[-ACTIVITY_MAX:]
            _write_json_atomic(ACTIVITY_PATH, {"events": events})
        return ev
    except Exception:
        return None


def read_events(limit=500, user_id=None, kind=None, since=None):
    with _lock:
        data = _read_json(ACTIVITY_PATH, None)
    events = data.get("events") if isinstance(data, dict) else data
    if not isinstance(events, list):
        return []
    out = []
    for e in events:
        if not isinstance(e, dict):
            continue
        if user_id and e.get("user_id") != user_id:
            continue
        if kind and e.get("type") != kind:
            continue
        if since and str(e.get("ts", "")) < since:
            continue
        out.append(e)
    out.reverse()                                   # newest first
    return out[:max(1, min(int(limit or 500), 5000))]


# --------------------------------------------------------------------------
# session helpers + decorators
# --------------------------------------------------------------------------

def current_user():
    uid = session.get("uid")
    return get_user(uid) if uid else None


def login_user(user):
    session.clear()
    session["uid"] = user["id"]
    session.permanent = True
    log_event(user, "login", {"provider": user.get("provider")})
    return user


def logout_user():
    u = current_user()
    if u:
        log_event(u, "logout")
    session.clear()


def _wants_json():
    if request.path.startswith("/api/"):
        return True
    return "application/json" in (request.headers.get("Accept") or "")


def login_required(fn):
    @wraps(fn)
    def wrapper(*a, **kw):
        u = current_user()
        if not u:
            if _wants_json():
                return jsonify(ok=False, error="Sign in to continue.",
                               auth_required=True), 401
            return redirect("/login")
        return fn(*a, **kw)
    return wrapper


def admin_required(fn):
    @wraps(fn)
    def wrapper(*a, **kw):
        u = current_user()
        if not u:
            if _wants_json():
                return jsonify(ok=False, error="Sign in to continue.",
                               auth_required=True), 401
            return redirect("/login")
        if not u.get("is_admin"):
            if _wants_json():
                return jsonify(ok=False, error="Admins only."), 403
            return redirect("/")
        return fn(*a, **kw)
    return wrapper


# --------------------------------------------------------------------------
# Google OAuth 2.0
# --------------------------------------------------------------------------

GOOGLE_AUTH = "https://accounts.google.com/o/oauth2/v2/auth"
GOOGLE_TOKEN = "https://oauth2.googleapis.com/token"


def google_configured():
    return bool(os.environ.get("GOOGLE_CLIENT_ID", "").strip()
                and os.environ.get("GOOGLE_CLIENT_SECRET", "").strip())


def google_redirect_uri():
    configured = os.environ.get("GOOGLE_REDIRECT_URI", "").strip()
    if configured:
        return configured
    return request.url_root.rstrip("/") + "/auth/google/callback"


def google_auth_url():
    """Authorization URL plus the CSRF state we expect back."""
    state = secrets.token_urlsafe(24)
    session["oauth_state"] = state
    from urllib.parse import urlencode
    params = {
        "client_id": os.environ["GOOGLE_CLIENT_ID"].strip(),
        "redirect_uri": google_redirect_uri(),
        "response_type": "code",
        "scope": "openid email profile",
        "state": state,
        "access_type": "online",
        "prompt": "select_account",
    }
    return GOOGLE_AUTH + "?" + urlencode(params)


def google_exchange(code):
    """Trade the code for an id_token and verify it. Raises ValueError with a
    readable message on any failure -- the caller shows it on the login page."""
    import requests
    from google.oauth2 import id_token as google_id_token
    from google.auth.transport import requests as google_requests

    try:
        resp = requests.post(GOOGLE_TOKEN, timeout=20, data={
            "code": code,
            "client_id": os.environ["GOOGLE_CLIENT_ID"].strip(),
            "client_secret": os.environ["GOOGLE_CLIENT_SECRET"].strip(),
            "redirect_uri": google_redirect_uri(),
            "grant_type": "authorization_code",
        })
    except requests.RequestException as e:
        raise ValueError(f"Could not reach Google ({type(e).__name__}).")

    if resp.status_code != 200:
        detail = ""
        try:
            detail = resp.json().get("error_description") or resp.json().get("error", "")
        except ValueError:
            pass
        raise ValueError(f"Google rejected the sign-in. {detail}".strip())

    tok = resp.json().get("id_token")
    if not tok:
        raise ValueError("Google returned no id_token.")

    # Verify signature, issuer, audience and expiry. Decoding without this
    # would accept any token a caller cared to paste in.
    try:
        info = google_id_token.verify_oauth2_token(
            tok, google_requests.Request(),
            os.environ["GOOGLE_CLIENT_ID"].strip())
    except Exception as e:
        raise ValueError(f"Could not verify Google's token ({e}).")

    if info.get("iss") not in ("accounts.google.com", "https://accounts.google.com"):
        raise ValueError("Unexpected token issuer.")
    if not info.get("email_verified", False):
        raise ValueError("That Google account has no verified email address.")
    return info
