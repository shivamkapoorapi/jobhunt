"""
Gemini re-ranking layer for jobhunt.py.

The keyword scorer in jobhunt.py is good at precision (does the title match?)
but blind to eligibility. It happily ranks a mid-level PM role requiring 5 years
of experience above a Summer 2027 PM internship, and it over-weights whichever
company happens to have the most postings open.

This module reads each posting and scores it against Shruti's actual situation:
a first-semester Columbia MS student on an F-1, graduating Dec 2027, who can
realistically take a Summer 2027 internship now and a new-grad role later.

Falls back silently to keyword-only ranking if the API key is missing or blocked,
so `jobhunt.py run` never breaks because of this file.
"""

import concurrent.futures as cf
import json
import os
import re
import threading
import time
import urllib.error
import urllib.request

MODEL = "gemini-3.5-flash-lite"
ENDPOINT = ("https://generativelanguage.googleapis.com/v1beta/models/"
            f"{MODEL}:generateContent")
# 10, not 20: batches now carry up to 1500 chars of description text each.
BATCH = 10
WORKERS = 2
RETRIES = 5
# A generated candidate_block shorter than this is treated as a stub. Kept low:
# falling back means scoring someone else's resume against the default profile.
MIN_PROFILE_CHARS = 40

DEFAULT_CANDIDATE = """\
SHRUTI GUPTA — target profile

CURRENT SITUATION (this is the most important part):
- Started M.S. Management Science & Engineering at Columbia University in
  August 2026. Graduates December 2027. She is in her FIRST SEMESTER now.
- International student on F-1. Needs CPT for a summer internship, and
  H-1B sponsorship for any eventual full-time role.
- Based in New York, NY. Strongly prefers NYC; open to the rest of the US.

BACKGROUND:
- 3.5 years at Jubilant FoodWorks (Domino's / Popeyes India), rising from
  Graduate Engineer Trainee to Senior Software Engineer (Jul 2023 - Jul 2026).
- Engineer who did product work: owned discovery, defined MVPs, ran
  experimentation, partnered with Design/Ops/Risk stakeholders.
- Domain depth: pricing and revenue, payments and prepaid adoption, fraud and
  COD abuse, checkout and cart conversion, food delivery / e-commerce
  marketplace, order management systems at scale.
- Technical: SQL, Python, Java, Kotlin, Spring Boot, React, MongoDB, Kafka.
- AI/ML: built an LLM Text-to-Query platform with LangChain, LangGraph, RAG,
  ChromaDB, semantic retrieval. Vector embeddings, prompt engineering.
- Analytics tooling: A/B testing, Firebase Analytics, PostHog, Figma, Jira.
- B.Tech Computer Science, VIT, GPA 9.44/10.

WHAT SHE WANTS: Product Management in the US. Secondarily product analytics,
strategy & operations, or data/business analyst roles that lead into PM.

HOW TO SCORE FIT (0-100):
- 90-100: Summer 2027 PM / APM / product analyst internship, or an MBA-or-MS
  level product internship. NYC or remote is a bonus. This is her sweet spot.
- 75-89: Any 2027 summer internship in product, analytics, strategy, or
  business ops at a strong company; or a new-grad / APM rotational program
  with a start date after December 2027.
- 55-74: New-grad or entry-level product/analytics role with no stated
  experience requirement, or an internship adjacent to her domain
  (fintech, payments, pricing, e-commerce, AI/LLM tooling).
- 30-54: Full-time mid-level role she cannot start until Dec 2027 but which
  matches her domain unusually well — worth tracking, not applying to now.
- 10-29: Full-time role requiring meaningful post-grad experience, or a weak
  domain match.
- 0-9: Requires 5+ years, is senior/staff/principal, is a non-product function
  she did not ask for, explicitly refuses visa sponsorship, or is US-citizen /
  security-clearance only (she cannot get a clearance).

BOOST for: pricing, revenue management, payments, fintech, fraud/risk,
checkout/conversion, marketplace, food delivery, e-commerce, LLM/RAG/AI product,
internal developer tools, experimentation platforms, data/self-serve analytics.
These are things she has actually shipped and can talk about in an interview.

PENALISE: roles needing a clearance, hardware/robotics/biotech domain roles with
no software-product angle, and heavily quantitative trading roles.
"""

# Back-compat alias: other modules refer to the module-level profile prompt as
# CANDIDATE. Keep both names bound so `ai_rank.CANDIDATE` never AttributeErrors.
CANDIDATE = DEFAULT_CANDIDATE

INSTRUCTION = """\
You are screening job postings for the candidate described above.

For EACH posting in the list, return one JSON object:
  {"i": <the id given>, "fit": <integer 0-100>, "why": "<max 14 words>",
   "eligible": "yes" | "no" | "unknown"}

Each posting's "text" now includes the real qualifications/requirements
section of the description, not just the intro blurb. You MUST weigh it when
scoring fit: experience minimums, graduation windows, and work-authorization
requirements stated there all move the score.

"eligible" is "no" ONLY when the text states a hard blocker for this
candidate: it refuses visa sponsorship/OPT/CPT, is US-citizen-only, requires
a security clearance, or requires a minimum of 4+ years of experience.
Use "unknown" when the text does not contain enough evidence either way.
When "eligible" is "no", name the blocker in "why".

FUNCTION MISMATCH IS A HARD CAP, NOT A DEDUCTION. Work out the job family
this posting belongs to (product management, software engineering, data,
design, sales, marketing, operations...) and compare it with the family the
candidate is targeting. If they differ, the score is AT MOST 30 - however
well the tech stack, seniority, location or domain line up. A Software
Engineer role is not a Product Manager role for someone seeking product, even
when they can obviously do the work. Say "wrong function" in "why".
Titles like "Product Engineer", "Technical Product Manager" or "Product
Analyst" ARE in the product family; "Software Engineer, Product Platform" is
not. Judge the actual role, not a keyword.

The "why" must be concrete and specific to THIS posting - name the reason
(e.g. "Summer 2027 PM intern, NYC, pricing domain" or
"needs 5+ yrs, she is still in school"). Never write a generic phrase.

Return ONLY a JSON array. No markdown fences, no commentary.
"""


def _api_key():
    key = os.environ.get("GEMINI_API_KEY")
    if key and key.strip():
        return key.strip()
    here = os.path.dirname(os.path.abspath(__file__))
    for path in (".env", os.path.join(here, ".env")):
        try:
            # utf-8-sig, not utf-8: a .env rewritten by PowerShell/Notepad on
            # Windows carries a BOM, which would glue itself to the first key
            # name and silently make the key undiscoverable.
            with open(path, encoding="utf-8-sig", errors="replace") as f:
                for line in f:
                    line = line.strip()
                    if line.startswith("#"):
                        continue
                    if line.startswith("export "):
                        line = line[len("export "):].lstrip()
                    name, sep, val = line.partition("=")
                    # partition, not split(...)[1]: a bare "GEMINI_API_KEY" line
                    # used to raise IndexError. Exact name match so that
                    # GEMINI_API_KEY_OLD=... cannot win.
                    if not sep or name.strip() != "GEMINI_API_KEY":
                        continue
                    val = val.strip().strip('"').strip("'").strip()
                    if val:
                        return val
        except OSError:
            continue
    return None


def _call(key, prompt, retries=None, timeout=180):
    """One request, with exponential backoff on rate limits and 5xx.

    retries/timeout are optional so a caller that just wants to validate a key
    can ask for a single fast attempt instead of ~90s of backoff.
    """
    body = json.dumps({
        "contents": [{"role": "user", "parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0, "responseMimeType": "application/json"},
    }).encode("utf-8")
    try:
        attempts = max(1, int(RETRIES if retries is None else retries))
    except (TypeError, ValueError):
        attempts = RETRIES
    last = RuntimeError("request was never attempted")
    for attempt in range(attempts):
        req = urllib.request.Request(
            ENDPOINT, data=body,
            headers={"x-goog-api-key": key, "Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=timeout) as r:
                data = json.loads(r.read().decode("utf-8", "replace"))
        except urllib.error.HTTPError as e:
            last = e
            # 429 = free-tier RPM cap, 5xx = transient. Anything else is fatal.
            if e.code not in (429, 500, 502, 503, 504):
                raise
            time.sleep(min(2 ** attempt * 3, 45))
            continue
        except (urllib.error.URLError, TimeoutError, json.JSONDecodeError,
                OSError) as e:
            last = e
            time.sleep(min(2 ** attempt * 3, 45))
            continue
        # A 200 carrying no candidates (safety block, MAX_TOKENS, quota body)
        # used to escape as a bare KeyError/IndexError with no readable text.
        try:
            parts = data["candidates"][0]["content"]["parts"]
        except (KeyError, IndexError, TypeError):
            reason = ""
            if isinstance(data, dict):
                fb = data.get("promptFeedback")
                if isinstance(fb, dict):
                    reason = fb.get("blockReason") or ""
                err = data.get("error")
                if not reason and isinstance(err, dict):
                    reason = err.get("message") or ""
                cands = data.get("candidates")
                if (not reason and isinstance(cands, list) and cands
                        and isinstance(cands[0], dict)):
                    reason = cands[0].get("finishReason") or ""
            raise RuntimeError(
                "model returned no text" + (f" ({reason})" if reason else ""))
        if not isinstance(parts, list):
            raise RuntimeError("model returned no text (malformed response)")
        return "".join(p.get("text", "") for p in parts if isinstance(p, dict))
    raise last


def _parse(text):
    text = (text or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```[a-z]*\n?|```$", "", text, flags=re.M).strip()
    try:
        out = json.loads(text)
    except ValueError:  # covers json.JSONDecodeError
        m = re.search(r"\[.*\]", text, re.S)
        try:
            out = json.loads(m.group(0)) if m else []
        except ValueError:
            out = []
    # The caller iterates this as a list of objects. A bare object, number or
    # string here used to blow up (or silently iterate characters) downstream.
    if isinstance(out, dict):
        for k in ("results", "items", "postings"):
            if isinstance(out.get(k), list):
                return out[k]
        out = [out]
    return out if isinstance(out, list) else []


def load_candidate():
    """The profile text to score against: profile.json if usable, else default.

    Read at call time (never at import) so a resume uploaded while the server
    is running takes effect on the next run without a restart. Never raises -
    a missing, malformed or stub profile.json silently falls back.
    """
    here = os.path.dirname(os.path.abspath(__file__))
    # Both locations: the server may write profile.json beside this module or
    # into whatever directory it was launched from.
    for path in (os.path.join(here, "profile.json"),
                 os.path.abspath("profile.json")):
        try:
            # utf-8-sig: a BOM would make json.load raise and silently drop the
            # user's uploaded profile back to the built-in default.
            with open(path, encoding="utf-8-sig") as f:
                data = json.load(f)
            block = data.get("candidate_block") if isinstance(data, dict) else None
            if isinstance(block, str) and len(block.strip()) >= MIN_PROFILE_CHARS:
                return block.strip()
        except Exception:
            continue
    return DEFAULT_CANDIDATE


def _batch_prompt(chunk):
    lines = []
    for i, r in chunk:
        # Prefer the cleaned full-description excerpt jobhunt.py stashes in
        # the private "_rank_text" key (carries real qualifications /
        # requirements text); fall back to the short summary column, which is
        # usually company boilerplate.
        text = str(r.get("_rank_text") or "").strip()
        text = text[:1500] if text else str(r.get("What The Role Says") or "")[:400]
        lines.append(json.dumps({
            "i": i,
            "company": r.get("Company", ""),
            "title": r.get("Title", ""),
            "location": r.get("Location", ""),
            "salary": r.get("Salary Range", ""),
            "type": r.get("Type", ""),
            "posted": str(r.get("Posted", "")),
            "text": text,
        }, ensure_ascii=False))
    return f"{load_candidate()}\n{INSTRUCTION}\nPOSTINGS:\n" + "\n".join(lines)


def rerank(rows, log=print, limit=None):
    """Add 'AI Fit', 'AI Verdict' and 'Eligibility' to each row.

    Eligibility is "YES"/"NO"/"UNKNOWN" for rows the model returned, and stays
    "" for rows it never saw. Returns True if AI ranking ran.
    """
    key = _api_key()
    if not key:
        log("  ! no GEMINI_API_KEY found - keeping keyword ranking only")
        return False

    targets = rows if limit is None else rows[:limit]
    for r in rows:
        r.setdefault("AI Fit", "")
        r.setdefault("AI Verdict", "")
        r.setdefault("Eligibility", "")

    chunks = [list(enumerate(targets))[i:i + BATCH]
              for i in range(0, len(targets), BATCH)]
    log(f"  scoring {len(targets)} postings against her profile "
        f"({len(chunks)} batches, {MODEL})...")

    done = [0]
    failed = [0]
    # done/failed and the log stream are touched by every worker thread;
    # `+= 1` on a list slot is not atomic and interleaved log() calls corrupt
    # the progress lines the UI parses.
    lock = threading.Lock()

    def work(chunk):
        time.sleep(0.6)
        try:
            out = _parse(_call(key, _batch_prompt(chunk)))
        except Exception as e:
            with lock:
                failed[0] += 1
                log(f"  ! batch failed ({type(e).__name__}) - those keep keyword score")
            return
        try:
            chunk_ids = {i for i, _ in chunk}
            seen = set()
            for item in out:
                if not isinstance(item, dict):
                    continue
                try:
                    idx = int(item["i"])
                    fit = max(0, min(100, int(float(item["fit"]))))
                # OverflowError: json.loads accepts Infinity/NaN, and
                # int(float("Infinity")) raises it - without it here one bad
                # value abandoned every remaining item in the batch.
                except (KeyError, ValueError, TypeError, OverflowError):
                    continue
                # One response per job: only ids belonging to THIS batch
                # count, and the first answer per id wins - a duplicated or
                # foreign id must never overwrite already-written data.
                if idx not in chunk_ids or idx in seen:
                    continue
                seen.add(idx)
                elig = str(item.get("eligible", "")).strip().upper()
                if elig not in ("YES", "NO", "UNKNOWN"):
                    elig = "UNKNOWN"
                with lock:
                    targets[idx]["AI Fit"] = fit
                    targets[idx]["AI Verdict"] = str(item.get("why", ""))[:120]
                    targets[idx]["Eligibility"] = elig
            with lock:
                done[0] += len(chunk)
                if len(seen) < 0.8 * len(chunk):
                    log(f"  ! model answered only {len(seen)}/{len(chunk)} items "
                        "in a batch - the rest keep keyword score, blank eligibility")
                log(f"  ranked {done[0]}/{len(targets)}")
        except Exception as e:
            # Nothing may escape a worker: ex.map() re-raises on iteration,
            # which would abandon every remaining batch and lose the ranking.
            with lock:
                failed[0] += 1
                log(f"  ! batch not usable ({type(e).__name__}) - keyword score kept")

    with cf.ThreadPoolExecutor(WORKERS) as ex:
        list(ex.map(work, chunks))

    scored = sum(1 for r in targets if r.get("AI Fit", "") != "")
    if not scored:
        log("  ! AI ranking produced nothing - keeping keyword order")
        return False

    # Sort by AI fit first, keyword score as tiebreak. Unscored rows sink
    # below anything the model actually looked at, but keep their own order.
    def _order(r):
        # Every part is coerced: one None/str Score or Company in the sheet
        # would otherwise raise TypeError mid-sort and lose the whole ranking.
        try:
            fit = int(r.get("AI Fit", ""))
        except (TypeError, ValueError):
            fit = -1
        try:
            score = float(r.get("Score") or 0)
        except (TypeError, ValueError):
            score = 0.0
        return (-fit, -score, str(r.get("Company") or ""))

    rows.sort(key=_order)
    log(f"  AI ranked {scored}/{len(targets)} postings"
        + (f" ({failed[0]} batches failed)" if failed[0] else ""))
    return True
