"""Which OPENING a JD document describes.

The JD library stores one opening as SEVERAL documents: typically a "Job Announcement"
plus a "Position Description", often with "Updated" / "v2" / "Comprehensive" revisions on
top, and sometimes a "LinkedIn Announcement" as well. Live count on 2026-08-03: 85 JD files
covering ~45 real openings.

Nothing in the source data linked those documents together, so the scorer used to compare
raw titles - and a single opening could therefore win Suggested Role 1, 2 AND 3 (a mobile
candidate scored "Mobile Application Lead Developer Job Announcement", "... Position
Description" and "... Position Description v2" as three separate suggestions). The client
sheet then showed one job three times instead of three real alternatives.

This module is the ONE place that answers "what opening is this document about". It runs at
INGEST (jd_sources stamps `opening` / `opening_id` / `doc_type` onto every role it builds),
so scoring simply reads a field instead of re-deriving identity on every comparison. The
derivation is still a filename heuristic - the source documents carry no opening ID of their
own - but it now runs once per document rather than once per candidate-vs-role comparison,
and there is exactly one implementation to correct when a new naming convention shows up.
"""

import re

#: Document-type markers. A JD title is structured "<Opening Name> <DocType> [qualifiers]",
#: so the opening name is everything BEFORE the first marker. Cutting at the marker (rather
#: than stripping a fixed list of trailing words) is what handles free-text qualifiers that
#: sit between the marker and the end: "Cybersecurity IT Admin Manager JA AI Tools Updated"
#: and "... PD Updated" are two revisions of one opening, and no suffix list would ever cover
#: every qualifier someone might type there.
_DOC_MARKER_RE = re.compile(
    r"\b(?:job\s+announcement|linkedin\s+announcement|job\s+description"
    r"|pos[it]+ion\s+(?:description|announcement)"      # tolerates the live 'Postion' typo
    r"|announcement|description|pos[it]+ion|ja|pd)\b",
    re.I,
)

#: Revision words for titles that carry a qualifier but NO document-type marker at all -
#: e.g. "Software Website Developer Intern Performance Metrics", whose sibling document is
#: "Software Website Developer Intern PD".
_REVISION_SUFFIX_RE = re.compile(
    r"(?:\bperformance\s+metrics|\bwith\s+overview"
    r"|\bupdated|\bupdate|\brevised|\brev|\bfinal|\bcomprehensive|\bv\d+)[\s\-_.,]*$",
    re.I,
)

#: Canonical labels for the document types we recognise, longest match first.
_DOC_TYPE_LABELS = (
    ("job announcement", "Job Announcement"),
    ("linkedin announcement", "LinkedIn Announcement"),
    ("job description", "Job Description"),
    ("position description", "Position Description"),
    ("postion description", "Position Description"),
    ("position announcement", "Position Announcement"),
    ("postion announcement", "Position Announcement"),
    ("announcement", "Job Announcement"),
    ("description", "Position Description"),
    ("position", "Position Description"),
    ("postion", "Position Description"),
    ("ja", "Job Announcement"),
    ("pd", "Position Description"),
)


def split_document_title(title: str) -> tuple[str, str]:
    """Split a JD FILE title into (opening name, document type).

    'Mobile Application Lead Developer Position Description v2'
        -> ('Mobile Application Lead Developer', 'Position Description')

    Only the document-type tail is removed, so genuinely different openings stay distinct
    ('Gaming Position' vs 'Graphics Position', 'AWS Lead' vs 'AWS Senior Engineer'). When
    cutting would leave nothing - a JD literally named 'Position Description' - the whole
    title is kept as the opening name, so it dedupes against itself rather than merging with
    every other marker-only title. Returns ('', '') for a blank title.
    """
    normalized = re.sub(r"\s+", " ", str(title or "").strip())
    if not normalized:
        return "", ""

    marker = _DOC_MARKER_RE.search(normalized)
    doc_type = ""
    if marker:
        matched = marker.group(0).lower()
        doc_type = next((label for token, label in _DOC_TYPE_LABELS if token == matched), "")
        opening = normalized[:marker.start()].strip(" -_.,")
    else:
        opening = normalized

    while True:
        trimmed = _REVISION_SUFFIX_RE.sub("", opening).strip(" -_.,")
        if trimmed == opening or not trimmed:
            break
        opening = trimmed

    return (opening or normalized), doc_type


def opening_id(title: str) -> str:
    """Stable identifier for the opening a JD document describes.

    Lowercased opening name with punctuation and spacing normalised away, so incidental
    formatting differences between two documents of the same opening ('Web Dev  Full Stack'
    vs 'Web Dev - Full Stack') do not split them into two openings.
    """
    opening, _ = split_document_title(title)
    slug = re.sub(r"[^a-z0-9]+", "-", opening.lower()).strip("-")
    return slug


def stamp_role_identity(role: dict) -> dict:
    """Add `opening`, `opening_id` and `doc_type` to a role dict, in place.

    Idempotent, and never overwrites an `opening_id` that is already set - a future JD source
    that knows its own opening ID (a real ATS, say) can supply one directly and this heuristic
    will step aside. Roles that arrive without these fields (a JD cache written before
    2026-08-03, or the built-in default_roles from config.yaml) get them backfilled on read,
    so every consumer can rely on the field being present.
    """
    if not isinstance(role, dict):
        return role
    if role.get("opening_id"):
        return role
    title = str(role.get("title", "") or "")
    opening, doc_type = split_document_title(title)
    role["opening"] = opening
    role["opening_id"] = opening_id(title)
    role["doc_type"] = doc_type
    return role


def stamp_roles(roles):
    """Backfill opening identity across a list of role dicts. Returns the same list."""
    for role in roles or []:
        stamp_role_identity(role)
    return roles


def role_opening_id(role) -> str:
    """The opening id for a role dict (preferred) or a bare title string (fallback)."""
    if isinstance(role, dict):
        stored = str(role.get("opening_id", "") or "").strip()
        if stored:
            return stored
        return opening_id(str(role.get("title", "") or ""))
    return opening_id(str(role or ""))


def role_display_name(role) -> str:
    """The human-facing role name for a role dict (preferred) or a bare title string.

    The JD library names its files after the DOCUMENT, not the opening, so a raw title
    reaches the client sheet carrying filing metadata: 'Data Analyst Job Announcement rev',
    'Cloud Database Engineer  PD', 'DriverAI Dual Business Analytics Marketing Intern Job
    Announcement (2)'. The opening name is the part a reader actually wants, and
    split_document_title already computes it for deduplication - this just exposes it for
    display so 'Suggested Role 1/2/3' read as job titles.

    Derived from the title rather than read off role['opening'], because stamp_role_identity
    short-circuits on a role that already carries an opening_id and can leave 'opening' unset.
    """
    title = str(role.get("title", "") or "") if isinstance(role, dict) else str(role or "")
    opening, _ = split_document_title(title)
    return re.sub(r"\s{2,}", " ", opening).strip()
