"""
Rules about a resume profile that the website and the search must agree on.

The website builds a profile when a resume is uploaded; the search on the PC
uses it. If each side had its own idea of which job titles count, or of which
profile is newer, they would disagree silently - so both import this module.
"""

import re
from collections import Counter

# A product-family title, checked first so "Product Manager" is never mistaken
# for anything else. Deliberately not "product market..." - Product Marketing
# is a marketing role.
_PRODUCT = re.compile(
    r"product\s+(manag|owner|analyst|lead|operations|ops|strateg)"
    r"|associate\s+product|\bapm\b|\bpm\b", re.I)

# Checked in order after the product test.
_FAMILIES = [
    ("engineering", re.compile(
        r"engineer|developer|\bswe\b|programmer|devops|\bsre\b|architect", re.I)),
    ("design", re.compile(r"designer|\bux\b|\bui\b|user experience", re.I)),
    ("sales", re.compile(
        r"\bsales\b|account executive|business development|\bbdr\b|\bsdr\b", re.I)),
    ("marketing", re.compile(r"marketing", re.I)),
    ("program", re.compile(r"program manag|project manag|\btpm\b", re.I)),
    ("analytics", re.compile(r"analyst|analytics|data scien", re.I)),
    ("strategy", re.compile(r"strateg|operations|bizops|consult", re.I)),
]

# Families far enough from one another that a title in one is noise in a
# search for another. Program, analytics and strategy sit close to product
# and stay.
FAR = {"engineering", "design", "sales", "marketing"}


def title_family(title):
    t = str(title or "")
    if _PRODUCT.search(t):
        return "product"
    for name, rx in _FAMILIES:
        if rx.search(t):
            return name
    return "product" if re.search(r"\bproduct\b", t, re.I) else "other"


def target_family(titles):
    """The job family most of the target titles belong to."""
    fams = [title_family(t) for t in titles or [] if str(t or "").strip()]
    fams = [f for f in fams if f != "other"]
    return Counter(fams).most_common(1)[0][0] if fams else None


def clean_titles(profile):
    """(profile, dropped): target/secondary titles outside the person's job
    family removed.

    Why this exists: reading a product manager's resume, Gemini lists
    "Software Engineer" as a secondary title because her background is
    engineering. Admitting that title let ~540 engineering postings into the
    search, 88% of the list, even though the ranker later capped them. The
    title list is the first gate, so it must stay in her family.
    """
    if not isinstance(profile, dict):
        return profile, []
    out = dict(profile)
    family = target_family(out.get("target_titles"))
    if not family:
        return out, []
    dropped = []
    for key in ("target_titles", "secondary_titles"):
        kept = []
        for t in out.get(key) or []:
            fam = title_family(t)
            if fam == family or fam not in FAR:
                kept.append(t)
            elif str(t) not in dropped:
                dropped.append(str(t))
        out[key] = kept
    return out, dropped


def profile_version(profile):
    """Upload time as epoch seconds. 0 for a profile made before versions
    existed, so any website upload counts as newer than it."""
    if not isinstance(profile, dict):
        return 0
    try:
        return int(profile.get("profile_version") or 0)
    except (TypeError, ValueError):
        return 0
