# JobHunt — a job search that reads your resume

Drop in a resume. It becomes the scoring rubric. The app then sweeps ~83 company
job boards, keeps the US roles, checks every link is still live, ranks what
survives against *your* situation with Gemini, and hands back a filtered Excel
file plus a tracker.

Built for one specific problem: an international student needs Summer 2027
internships, and a keyword search happily puts a "Product Manager II, 5+ years"
role above them.

![Python](https://img.shields.io/badge/python-3.10%2B-blue)

---

## What it does

**Resume in → ranked shortlist out.** Four steps in the browser:

1. **Connect a Gemini key** — tested against the live model before anything is saved.
2. **Drop a resume** (PDF/DOCX/TXT) — it writes your titles, skills, locations
   and a 0–100 fit rubric derived from your actual timeline.
3. **Run the search** — live log console, ~7 minutes.
4. **Get the Excel** — filtered, verified, ranked.

Plus a kanban **tracker** (Saved → Applied → Interview → Offer → Rejected),
**run history** with per-run fit distributions, Google sign-in, and an admin
console.

## Why the ranking is two layers

Keyword matching answers *"does this title match?"* It cannot tell that a
mid-level PM role wanting five years is useless to someone still in school. So
every posting is also read by Gemini and scored against the candidate's real
eligibility window — graduation date, work authorization, experience level.

Roles that state a hard blocker (no sponsorship, citizens only, clearance
required, 4+ years minimum) are marked `Eligibility: NO` and leave the apply
list entirely.

The profile it scores against lives in `candidate_block` inside `profile.json`,
in plain English. Edit it and the whole ranking follows.

## The spreadsheet

| Sheet | What it holds |
|---|---|
| **Fresh (48h)** | posted today or yesterday — start here |
| **Ready to Apply** | verified-live, above your fit cutoff |
| New This Run | only what appeared since last time |
| All Matches | everything, your tracker of record |
| **Needs Link Check** | real jobs whose careers site was down |
| Low Fit | below your cutoff, kept for reference |
| **Run Report** | every source, its status, and whether the run was complete |
| Quarantine | links the site itself said were dead |

## Two design rules worth knowing

**Nothing disappears silently.** A row the ranker never scored is never treated
as a bad match. An `UNREACHABLE` link means *our* checker got rate-limited, not
that the job is gone — so it's kept. But when *every* link on one host fails
(a careers site being down), those roles move to their own sheet instead of
filling your apply list with dead ends.

**A run says when it's incomplete.** If a source times out, the Run Report and
the log say so, rather than quietly returning fewer jobs that look like a
complete answer.

---

## Running it

```bash
pip install -r requirements.txt
python app.py            # -> http://127.0.0.1:5000
```

Sign in, add a [Gemini API key](https://aistudio.google.com/apikey), drop a
resume, run.

Command line, if you prefer:

```bash
python jobhunt.py run          # search and write the Excel
python jobhunt.py add stripe   # track another company by ATS token
python jobhunt.py check        # prune dead company tokens
python jobhunt.py schedule     # install a daily run
```

### Configuration

`config.json` ships with 83 companies and sensible defaults. Worth knowing:

| Key | Default | Meaning |
|---|---|---|
| `min_ai_fit` | 55 | below this, roles go to the Low Fit sheet |
| `fresh_days` | 2 | what counts as "fresh" |
| `max_age_days` | 45 | ignore postings older than this |
| `verify_links` | true | open every link before it reaches Excel |

Both `min_ai_fit` and `fresh_days` are editable from the UI.

### Adding companies

The highest-leverage change you can make. Grab the token from a careers URL —
`job-boards.greenhouse.io/**stripe**` → `stripe` — then:

```bash
python jobhunt.py add stripe
```

Greenhouse, Lever, Ashby, SmartRecruiters and Workable are auto-detected.

## Sources

Public ATS APIs — first-party, no API key, no rate limiting, and every link goes
to the real application page. Layered with the
[SimplifyJobs](https://github.com/SimplifyJobs) feeds for internships and new-grad
roles.

Deliberately **not** LinkedIn or Indeed scraping: it rate-limits hard and puts
your own account at risk.

## Tests

```bash
python test_filters.py     # 48 regression tests
```

Covers the parsing that quietly gets things wrong: `"Remote UK"` is not US,
`"Product Manager | International"` is not an internship, `"unable to provide
H-1B sponsorship"` does not mean they sponsor, and `"leverage"` does not contain
the skill `RAG`.

## Accounts

Google sign-in plus a local admin account. Passwords are stored only as
PBKDF2-SHA256 hashes (240k rounds) — set `ADMIN_PASSWORD_HASH`, never a
plaintext password:

```bash
python -c "import auth; print(auth.hash_password('your-password'))"
```

Setup for Google OAuth is in [DEPLOY.md](DEPLOY.md).

## Deploying

See [DEPLOY.md](DEPLOY.md). Short version: the UI, accounts, tracker and admin
console can run on Vercel, but **the search cannot** — it takes 7–15 minutes and
writes files, while a serverless function is capped at 60–300s on an ephemeral
read-only filesystem. It runs on your machine, a VM, or a scheduled job.

## Security

`.gitignore` excludes `.env`, `data/`, `uploads/`, resumes, and the generated
spreadsheets — everything holding a secret or personal data.

The app binds to `127.0.0.1` on purpose. It accepts file uploads, writes config
and launches subprocesses, so putting it on `0.0.0.0` exposes all of that to
your network.

## Layout

```
app.py              Flask backend, HTTP API
auth.py             accounts, sessions, activity log
jobhunt.py          the search engine (fetch, filter, verify, write Excel)
ai_rank.py          Gemini re-ranker
resume_profile.py   resume -> scoring profile
static/             index.html, login.html, admin.html
test_filters.py     regression suite
```

## Licence

MIT
