"""Vercel entry point.

Vercel runs this as a serverless function. Read DEPLOY.md before relying on it:
the job SEARCH cannot run here (it takes 7-15 minutes and writes files, while a
Vercel function is capped at 60-300s on an ephemeral, read-only filesystem).
What does work is the account layer, the tracker, the admin console and viewing
results -- provided storage is pointed at something persistent.
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app  # noqa: E402

# Vercel's Python runtime looks for a WSGI callable named `app`.
application = app
