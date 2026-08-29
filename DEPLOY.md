# Deploying

## Read this first: what Vercel can and cannot run

The scaffolding is in place (`vercel.json`, `api/index.py`, `requirements.txt`)
and the account layer is built to work there. But a straight "push to Vercel and
it all works" is not possible, and it is better to know why now than after a
deploy that half-works.

Vercel runs serverless functions. Three of its properties collide with this app:

| Vercel | This app |
|---|---|
| Functions time out at **60s** (Hobby) / **300s** (Pro) | A search takes **7–15 minutes** |
| Filesystem is **read-only** except `/tmp`, which is **wiped between invocations** | Writes `config.json`, `profile.json`, `tracker.json`, `history.json`, `.xlsx` |
| **No shared memory** between invocations | Live progress streaming keeps the log lines in module state |

So:

| Feature | On Vercel |
|---|---|
| Google sign-in / accounts | ✅ works |
| Admin console | ✅ works |
| Job tracker | ✅ works — **once storage moves off disk** |
| Run history + trends | ✅ works — same condition |
| Viewing ranked results | ✅ works — same condition |
| Uploading a resume | ⚠️ works, but the file vanishes; keep it in blob storage |
| **Running a search** | ❌ **cannot work** — it needs minutes and a filesystem |
| Live progress console | ❌ needs shared state |

## The two honest options

### Option A — keep it local (nothing to do, works today)

`python app.py` → `http://127.0.0.1:5000`. Everything works, including the
search. This is what the app is built for and what is tested.

### Option B — split it (what a deployed version really looks like)

- **Vercel** hosts the UI, accounts, tracker, admin and results viewer.
- **Storage** moves from JSON files to Postgres (Neon/Supabase) + Blob, so it
  survives between invocations. `auth.py` already funnels every write through
  `_write_json_atomic`, and `app.py` through `_read_json`/`_write_json_atomic`,
  so this is a contained change — swap those helpers, not the whole app.
- **The search** runs where it can take its time: your machine on a schedule
  (`python jobhunt.py run`), a small VM, or a GitHub Action. It uploads the
  workbook + history to the same storage the Vercel app reads.

Option B is a real piece of work — mostly the storage swap. Say the word and
I'll do it; I did not start it because you asked for Vercel-*ready*, and
pretending a 15-minute job runs in a 60-second function would have been the
wrong kind of "ready".

## Deploying what exists

```bash
npm i -g vercel
vercel link          # pick the project you want
vercel --prod
```

### Environment variables to set in Vercel (never commit these)

| Variable | Why |
|---|---|
| `FLASK_SECRET_KEY` | signs session cookies; without it every deploy logs everyone out |
| `ADMIN_USERNAME` | `shivamkapoor` |
| `ADMIN_PASSWORD_HASH` | the PBKDF2 hash from your local `.env` — **never the plain password** |
| `GOOGLE_CLIENT_ID` | from Google Cloud Console |
| `GOOGLE_CLIENT_SECRET` | same |
| `GOOGLE_REDIRECT_URI` | `https://<your-app>.vercel.app/auth/google/callback` |
| `GEMINI_API_KEY` | only if you rank on the server |

`SESSION_COOKIE_SECURE` turns itself on when `VERCEL` is set, so cookies become
HTTPS-only automatically.

## Setting up Google sign-in

1. [console.cloud.google.com](https://console.cloud.google.com) → new project.
2. **APIs & Services → OAuth consent screen** → External → add your email as a
   test user (or publish it once you are ready for real users).
3. **Credentials → Create credentials → OAuth client ID → Web application.**
4. Authorised redirect URIs — add both:
   - `http://127.0.0.1:5000/auth/google/callback`
   - `https://<your-app>.vercel.app/auth/google/callback`
5. Put the client ID and secret in `.env` locally and in Vercel's env vars.
6. Restart. The login page's Google button lights up on its own.

Until those are set, the button is visibly disabled and says so — the local
admin login still works.

## Pushing to GitHub

`.gitignore` already excludes `.env`, `data/`, `uploads/`, `profile.json`,
`tracker.json`, `history.json` and the spreadsheets — i.e. every file holding a
secret or someone's personal data.

```bash
git init
git add -A
git commit -m "Resumify: job search, tracker, accounts"
git remote add origin https://github.com/shivamkapoor172002/jobhunt.git
git push -u origin main
```

Before the first push, confirm nothing sensitive is staged:

```bash
git status --short          # .env, data/, uploads/ must NOT appear
git ls-files | grep -E "\.env|data/|uploads/"    # must print nothing
```

## Security notes

- The admin password is stored **only** as a PBKDF2-SHA256 hash (240k rounds).
  The plaintext is nowhere in the repo. Change it with:
  `python -c "import auth;print(auth.hash_password('new-password'))"`
  then update `ADMIN_PASSWORD_HASH`.
- The app binds to `127.0.0.1` locally, on purpose. It has file upload, config
  writes and subprocess launching, so exposing it on `0.0.0.0` puts all of that
  on your network.
- Users are told on the login page that their activity is recorded. Keep that
  notice — recording it without saying so is the part that would be wrong.
