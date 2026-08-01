"""Resume text extraction, candidate detail parsing, name resolution, and skills."""

import io
import json
import re
from html import unescape

from hiring_agent.config import (
    SKILL_KEYWORDS, SKILL_DISPLAY, SECTION_WORDS, ROLE_WORDS,
    LOCATION_KEYWORDS, AI_TEXT_LIMIT,
    OLLAMA_ENABLED, OLLAMA_MODEL, OLLAMA_HOST, OLLAMA_TIMEOUT,
    SCORING_MAX_SKILLS, US_STATE_ABBREVS, US_STATE_NAMES,
    US_COUNTRY_TERMS, FOREIGN_COUNTRIES, FOREIGN_CITIES, logger,
)

# Fields the Ollama extractor returns (the offline parser fills the same shape).
_AI_PROMPT = (
    "You extract candidate details from a job application (email body + resume text). "
    "Return every field as a string; use an empty string when a field is not present. "
    "For location, extract the candidate's CURRENT contact/home location only. Prefer "
    "the resume header/contact section or an explicit current-location statement; never "
    "use a past school, employer, project, or hometown as the current location. "
    "'location' should be city and state/region only — do NOT include the country "
    "(e.g. 'Austin, Texas' or 'Bengaluru, Karnataka' or 'London'). "
    "'country' is the country name (e.g. 'United States' or 'India' or 'UK') — "
    "always infer the country from context clues (area codes, universities, "
    "addresses, currency, employer names). "
    "'skills' is a short comma-separated list of individual skills or tools ACTUALLY NAMED "
    "in the text below - never invent or add a skill that is not written in the resume, "
    "and never return section headings, categories, or descriptive phrases; "
    "'looking_for_role' is the job title/role they seek. "
    "'education' is the candidate's highest degree and field of study, plus school if given "
    "(e.g. 'B.S. Computer Science, Arizona State University') - empty string if the resume "
    "doesn't mention any degree or schooling."
)


# Values an extractor can return that mean "nothing was actually found" - never let one
# of these overwrite a real value, and always treat one as a gap that still needs filling.
_GAP_LITERALS = {"not extracted", "n/a", "na", "none", "-", "not found", "unknown", ""}


def _ai_fields(data: dict, text: str) -> dict:
    """Normalize an LLM's parsed dict into the candidate-detail shape."""
    def field(key):
        value = str(data.get(key, "") or "").strip()
        return value if value else "Not extracted"
    return {
        "full_name": field("full_name"),
        "phone": field("phone"),
        "location": field("location"),
        "country": field("country"),
        "skills": field("skills"),
        "looking_for_role": field("looking_for_role"),
        "education": normalize_education(field("education")),
        "notes": text[:800],
    }


# ── Name resolution ──────────────────────────────────────────────────────────

# Words that essentially never appear in a real person's first/last name but are
# extremely common in a company/entity name - added 2026-08-01, live case (Futurism
# Technologies / APP-20260717-0559-CCE1): "Transforming Business Models" (a marketing
# tagline from a vendor's company-brochure PDF, not a resume) is 3 alphabetic tokens with
# no digits or section-word overlap, so it passed the existing structural shape check
# outright and was stored as the candidate's Full Name. A word-list veto catches this
# class of mistake without needing a real name database: none of these words are
# plausible as a genuine first or last name.
_COMPANY_NAME_WORDS = {
    "inc", "llc", "ltd", "corp", "corporation", "company", "co",
    "technologies", "technology", "solutions", "systems", "software", "digital",
    "group", "enterprises", "industries", "consulting", "partners", "associates",
    "ventures", "holdings", "capital", "labs", "global", "international", "worldwide",
    "services", "transforming", "business", "models",
}


def _looks_like_company_name(s: str) -> bool:
    """True if any token in s is common company/entity vocabulary, never a real name."""
    tokens = [t.strip(".,'-").lower() for t in (s or "").split()]
    return any(t in _COMPANY_NAME_WORDS for t in tokens)


def _looks_like_name(s: str) -> bool:
    """True if s looks like a real person's name: 2-3 alpha tokens, none a section word.

    A single-character token is allowed as a MIDDLE INITIAL ('Christopher L. Feld',
    'Jane Q Public') - fixed 2026-08-01, found by the offline dry-run harness: the old
    `len(core) < 2` rule rejected every name carrying an initial, so the offline parser
    returned 'Not extracted' for them. Live rows looked fine only because Ollama (Tier 2)
    resolves the name independently; the bug was invisible until Ollama was disabled,
    which is exactly the documented fallback path. At least TWO full-length tokens are
    still required, so an initial can never carry the name on its own ('A B', 'I am' are
    still rejected).
    """
    s = (s or "").strip()
    if not s or "@" in s or any(ch.isdigit() for ch in s):
        return False
    if _looks_like_company_name(s):
        return False
    tokens = s.split()
    if not (2 <= len(tokens) <= 3):
        return False
    _cores = [t.replace(".", "").replace("'", "").replace("-", "") for t in tokens]
    if sum(1 for c in _cores if len(c) >= 2) < 2:
        return False
    for t in tokens:
        core = t.replace(".", "").replace("'", "").replace("-", "")
        if len(core) < 1 or not core.isalpha() or t.lower() in SECTION_WORDS:
            return False
    return True


def _normalize_name(s: str) -> str:
    """'YASH VERMA' -> 'Yash Verma'; leave already-mixed-case tokens unchanged."""
    return " ".join(t.title() if t.isupper() else t for t in (s or "").split())


def _plausible_name_shape(s: str) -> bool:
    """Looser structural check for an AI/regex-PARSED name candidate: 1-4 alpha tokens,
    none a section word, no digits/@.

    Deliberately wider than _looks_like_name's 2-3-token window (a parsed name may
    legitimately be a single name or include a middle name), but still catches an AI
    slip-up like echoing back a full sentence or a role title as the "name" - a gap
    found during a final review: the old check here only rejected on a section-word
    overlap, not on the string actually having name-like shape at all.
    """
    s = (s or "").strip()
    if not s or "@" in s or any(ch.isdigit() for ch in s):
        return False
    if _looks_like_company_name(s):
        return False
    tokens = s.split()
    if not (1 <= len(tokens) <= 4):
        return False
    for t in tokens:
        core = t.replace(".", "").replace("'", "").replace("-", "")
        if not core.isalpha() or t.lower() in SECTION_WORDS:
            return False
    return True


def resolve_full_name(parsed_name: str, sender_name: str, resume_text: str) -> str:
    """Decide the candidate's name, most-trustworthy source first."""
    p = (parsed_name or "").strip()
    if p and p != "Not extracted" and _plausible_name_shape(p):
        return _normalize_name(p)
    if _looks_like_name(sender_name):
        return _normalize_name(sender_name.strip())
    for line in (resume_text or "").splitlines()[:20]:
        if _looks_like_name(line):
            return _normalize_name(line.strip())
    s = (sender_name or "").strip()
    if s and "@" not in s and not any(c.isdigit() for c in s):
        return s
    return "Not extracted"


# ── Title / location helpers ─────────────────────────────────────────────────

def _normalize_title(s: str) -> str:
    """Title-case a role line but keep short acronyms upper: 'AI', 'ML', 'QA', 'CISO'.

    Cap raised 3 -> 4 chars - fixed 2026-08-01, live case (Brett Worker /
    APP-20260721-0122-E743): 'CISO' (a genuine 4-letter role acronym extracted from his
    mail body) was getting reduced to 'Ciso'.
    """
    out = []
    for w in (s or "").split():
        if w.isupper() and len(w) <= 4:
            out.append(w)
        else:
            out.append(w.title() if w.isupper() else w)
    return " ".join(out).strip()


def _extract_header_role(text: str) -> str | None:
    """Grab a professional title from the top lines of a resume."""
    lines = [ln.strip() for ln in (text or "").splitlines() if ln.strip()]
    for ln in lines[:6]:
        for seg in re.split(r"\s*[|*,]\s*|\s+-\s+", ln):
            seg = seg.strip()
            if not seg or _looks_like_name(seg):
                continue
            words = seg.split()
            # Cap raised 6 -> 8 words - fixed 2026-08-01, live case (Mindy Anderson /
            # APP-20260723-2034-12DF): her actual title line, "Fractional/Interim Chief
            # Marketing Officer & Marketing Advisor" (7 words), was rejected outright for
            # being one word over the old cap, so nothing caught her real title and
            # extraction fell through to a generic keyword-bucket guess instead ("Data
            # Scientist / Analyst" for a CMO). A compound "X & Y" / "X/Y" executive title
            # is common enough that 8 is still tight relative to a genuine descriptive
            # sentence, which routinely runs well past 8 words.
            if 1 <= len(words) <= 8 and any(w.lower().rstrip("s") in ROLE_WORDS for w in words):
                if seg.lower() not in ("work experience", "professional experience",
                                       "experience", "education", "employment history"):
                    return _normalize_title(seg)
    return None


_INSTITUTION_LINE_RE = re.compile(
    r"(?i)\b(university|college|institute|academy|polytechnic|school\s+of)\b"
)


def _extract_location(text: str) -> str | None:
    """Resume location, most reliable first.

    1. Header "City, XX" (2-letter state abbrev) in first 6 lines
    2. Header "City, Full State Name" in first 6 lines
    3. Full-text "City, XX" anywhere (work history addresses, etc.)
    4. LOCATION_KEYWORDS scan (header first, then full text)
    5. Full-text US state full name after a comma anywhere
    6. Foreign country/city scan (so the geo filter can reject non-USA)
    """
    lines = [ln.strip() for ln in (text or "").splitlines() if ln.strip()]
    header = lines[:6]

    # Steps 4 and 6 below scan the WHOLE document (not just the header) for a bare
    # city/country name, so a city that only appears as part of a school's own name
    # ('Daulat Ram College, University of Delhi', 'University of Phoenix', 'San Jose
    # State University') was being returned as the candidate's current location - fixed
    # 2026-08-01, live case (Prerna Saluja / APP-20260716-2052-6112): her only 'Delhi'
    # mention is her undergrad alma mater; her current, most recent affiliation is
    # Arizona State University. Skip any match whose line is naming an institution.

    _city_state_re = re.compile(
        r"([A-Z][A-Za-z.\-]+(?:\s+[A-Z][A-Za-z.\-]+){0,2}),\s*([A-Z]{2})\b"
    )

    for ln in header:
        for seg in re.split(r"\s*[|]\s*", ln):
            m = _city_state_re.match(seg.strip())
            if m:
                return f"{m.group(1)}, {m.group(2)}"

    # 1b. Header "City ST" with NO comma ('Phoenix AZ 85004', 'Dallas TX | 682-...').
    #     Extremely common in real resume headers and previously missed entirely, which
    #     is how live rows ended up with Location "Not extracted" and were then rejected.
    #     Kept deliberately tight to avoid false positives from prose like 'Java OR
    #     Python': header lines only, the 2-letter token must be a real state abbrev,
    #     and it must end the segment (optionally followed by a ZIP) rather than be
    #     followed by another word.
    _city_state_nocomma_re = re.compile(
        r"^([A-Z][A-Za-z.\-]+(?:\s+[A-Z][A-Za-z.\-]+){0,2})\s+([A-Z]{2})"
        r"(?:\s+\d{5}(?:-\d{4})?)?\s*$"
    )
    for ln in header:
        for seg in re.split(r"\s*[|•·]\s*", ln):
            m = _city_state_nocomma_re.match(seg.strip())
            if m and m.group(2).lower() in US_STATE_ABBREVS:
                return f"{m.group(1)}, {m.group(2)}"

    for ln in header:
        for seg in re.split(r"\s*[|]\s*", ln):
            seg_low = seg.strip().lower()
            # Drop a trailing ZIP before comparing, so 'Phoenix, Arizona 85004' still
            # resolves to the state instead of falling through to the fuzzy scan below.
            seg_low = re.sub(r"\s+\d{5}(?:-\d{4})?$", "", seg_low)
            parts = re.split(r",\s*", seg_low)
            if len(parts) >= 2 and parts[-1].strip() in US_STATE_NAMES:
                return ", ".join(p.strip().title() for p in parts)
            # "City, State, Country" — state in second-to-last part, strip trailing country
            if len(parts) >= 3 and parts[-2].strip() in US_STATE_NAMES:
                return ", ".join(p.strip().title() for p in parts[:-1])

    for ln in lines:
        for m in re.finditer(_city_state_re, ln):
            if m.group(2).lower() not in US_STATE_ABBREVS:
                continue
            # A genuine work-history "City, ST" address ends there (end of line, a date,
            # a pipe, punctuation) - it is never immediately followed by another bare
            # capitalized word continuing a comma-separated list. Fixed 2026-08-01, live
            # case (Syyed Nazir Ali / APP-20260727-1431-6F75): "Bloomberg Terminal, MS
            # Project, MS Office Suite" is a tools list ('MS' = Microsoft), not an
            # address, but matched "Bloomberg Terminal, MS" as a "City, ST" location.
            if re.match(r"\s+[A-Z][a-z]", ln[m.end():]):
                continue
            return f"{m.group(1)}, {m.group(2)}"

    for scope in (header, lines):
        joined = " \n ".join(scope).lower()
        for city in LOCATION_KEYWORDS:
            if re.search(r"(?<![a-z])" + re.escape(city) + r"(?![a-z])", joined):
                hit_line = next(
                    (ln for ln in scope if re.search(
                        r"(?<![a-z])" + re.escape(city) + r"(?![a-z])", ln.lower())),
                    "",
                )
                if _INSTITUTION_LINE_RE.search(hit_line):
                    continue
                return city.title()

    # Search line by line, NOT over " ".join(lines): joining lets the city pattern run
    # backwards across a line break and swallow whatever preceded it, which turned
    # "Jane Doe\nPhoenix, Arizona 85004" into the location "Jane Doe Phoenix, Arizona".
    for state in US_STATE_NAMES:
        state_re = re.compile(
            r"([A-Z][A-Za-z.\-]+(?:\s+[A-Z][A-Za-z.\-]+){0,2}),\s*"
            + re.escape(state.title()) + r"(?![a-z])")
        for ln in lines:
            if _INSTITUTION_LINE_RE.search(ln):
                continue
            m2 = state_re.search(ln)
            if m2:
                return f"{m2.group(1)}, {state.title()}"
        for ln in lines:
            if _INSTITUTION_LINE_RE.search(ln):
                continue
            if re.search(r",\s*" + re.escape(state) + r"(?![a-z])", ln.lower()):
                return state.title()

    # 6. Foreign country/city (header first, then full text) — so the geo filter can
    #    reject non-USA candidates instead of defaulting to "benefit of the doubt".
    def _location_like(seg: str) -> bool:
        # a real location segment is short ('Mumbai, Maharashtra, India');
        # a prose sentence that merely mentions a country is not
        return len(seg) <= 60 and len(seg.split()) <= 6

    for scope in (header, lines):
        joined = " \n ".join(scope).lower()
        for country in FOREIGN_COUNTRIES:
            if re.search(r"(?<![a-z])" + re.escape(country) + r"(?![a-z])", joined):
                for ln in scope:
                    if _INSTITUTION_LINE_RE.search(ln):
                        continue
                    for seg in re.split(r"\s*[|]\s*", ln):
                        if country in seg.lower():
                            seg = seg.strip()
                            return seg if _location_like(seg) else country.title()
        for city in FOREIGN_CITIES:
            if re.search(r"(?<![a-z])" + re.escape(city) + r"(?![a-z])", joined):
                for ln in scope:
                    if _INSTITUTION_LINE_RE.search(ln):
                        continue
                    for seg in re.split(r"\s*[|]\s*", ln):
                        if city in seg.lower():
                            seg = seg.strip()
                            return seg if _location_like(seg) else city.title()

    return None


# Degree keywords, longest/most-specific alternatives first so e.g. "Bachelor of Science"
# matches whole rather than stopping at a shorter overlapping alternative.
#
# Bachelor/Master/Associate require a qualifying suffix ('s / of X) rather than matching
# bare - fixed 2026-07-31 after a live case (Syyed Nazir Ali, APP-20260727-1431-6F75) where
# "Scrum Master" on line 2 (a job-title line right under the name) matched bare "Master"
# and got returned as the Education value, entirely preempting the real "Bachelor of
# Science" degree line near the end of the resume. A genuine degree mention essentially
# always includes "'s"/"s" or "of <field>"; a bare "Master"/"Bachelor"/"Associate" is far
# more likely to be a job title ("Scrum Master", "Associate Director", a bachelor party
# planner, etc.) than a degree.
_DEGREE_RE = re.compile(
    r"\b(Bachelor(?:'s|s)\b|Bachelor\s+of\s+\w+|Master(?:'s|s)\b|Master\s+of\s+\w+|"
    r"Associate(?:'s|s)\b|Associate\s+of\s+\w+|Ph\.?D\.?|MBA|B\.?Tech\.?|M\.?Tech\.?|"
    r"B\.?S\.?[Cc]?\.?|M\.?S\.?[Cc]?\.?|B\.?A\.?|M\.?A\.?|B\.?E\.?|M\.?E\.?)\b"
)


_EDUCATION_FUSED_RE = re.compile(
    r"(?i)(?:master|bachelor|associate|university|college|institute)"
    r"(?:of|in|at)(?:science|arts|technology|engineering|computer|business|information)"
)


def education_needs_repair(value: str) -> bool:
    """True when an Education cell is visibly fused, truncated, OCR-damaged, or garbage."""
    text = str(value or "").strip()
    if not text or text.lower() in _GAP_LITERALS:
        return True
    low = text.lower()
    return bool(
        _EDUCATION_FUSED_RE.search(text)
        or re.search(r"(?i)\b(?:master|bachelor)\s*['\u2019]?\s*s?in\b", text)
        or re.search(r"(?i)\b(?:universit|univer|unive|technical)\s*$", text)
        or re.search(r"(?i)\b[a-z]\s+[a-z]{4,}\b", text)
        or "educationuniversity" in low
        # A resume-formatting divider ('----...----', '====...====') mis-captured as the
        # Education value: no letters at all, so normalize_education's strip() would
        # already reduce it to "Not extracted" - flag it so a heal pass actually applies
        # that, instead of leaving the raw divider published.
        or not any(c.isalpha() for c in text)
    )


def normalize_education(value: str) -> str:
    """Conservatively format a degree/school string without inventing information."""
    text = str(value or "").strip()
    if not text or text.lower() in _GAP_LITERALS:
        return "Not extracted"
    text = text.replace("\u00a0", " ").replace("\u2019", "'")
    replacements = (
        (r"(?i)\bMaster\s*'?\s*s?\s*of\s*Science\s*in\b", "Master of Science in"),
        (r"(?i)\bMasterofSciencein", "Master of Science in "),
        (r"(?i)\bBachelorofSciencein", "Bachelor of Science in "),
        (r"(?i)\bBachelorofTechnologyin", "Bachelor of Technology in "),
        (r"(?i)\bBachelorofEngineeringin", "Bachelor of Engineering in "),
        (r"(?i)\bMasterofComputerScience\b", "Master of Computer Science"),
        (r"(?i)\bMasterofScience\b", "Master of Science"),
        (r"(?i)\bBachelorofScience\b", "Bachelor of Science"),
        (r"(?i)\bUniversityof\b", "University of "),
        (r"(?i)\bUniversityat\b", "University at "),
        (r"(?i)\bArizonaStateUniversity\b", "Arizona State University"),
        (r"(?i)\bSanJoseStateUniversity\b", "San Jose State University"),
        (r"(?i)\bOregonStateUniversity\b", "Oregon State University"),
        (r"(?i)\bNortheasternUniversity\b", "Northeastern University"),
        (r"(?i)\bComputerScience\b", "Computer Science"),
        (r"(?i)\bDataScience\b", "Data Science"),
        (r"(?i)\bBusinessAnalytics\b", "Business Analytics"),
        (r"(?i)\bInformationTechnology\b", "Information Technology"),
        (r"(?i)\bSoftwareEngineering\b", "Software Engineering"),
        (r"(?i)\bComputerEngineering\b", "Computer Engineering"),
    )
    for pattern, replacement in replacements:
        text = re.sub(pattern, replacement, text)
    text = re.sub(r"\s+([,;])", r"\1", text)
    text = re.sub(r"([,;])(?=\S)", r"\1 ", text)
    text = re.sub(r"\s{2,}", " ", text).strip(" ,;-|")
    if text.upper() in {"AND TRAINING", "EDUCATION", "TRAINING", "AND"}:
        return "Not extracted"
    return text or "Not extracted"


def _collapse_letter_spacing(line: str) -> str:
    """'E D U C A T I O N' -> 'EDUCATION'; 'M a s t e r  o f  S c i e n c e' -> 'Master of
    Science'. PDF extraction preserves a stylized letter-spaced layout (a common resume-
    template design) as literal single-character tokens separated by ONE space, with word
    boundaries marked by a DOUBLE space - fixed 2026-07-31, live case (Divy Parmar /
    APP-20260720-1013-09AC) where an ENTIRE resume was letter-spaced this way (name, phone,
    email, every line), not just a section header. A plain '^education\\b' style regex, or
    any word-boundary keyword scan, can't match single-character tokens at all - so
    extraction silently fails wherever this pattern appears. Only collapses when the line
    actually looks letter-spaced (at least 4 single-character tokens), so normal prose is
    never touched."""
    tokens = line.split(" ")
    single_char_tokens = [t for t in tokens if t]
    if not (len(single_char_tokens) >= 4 and all(len(t) == 1 for t in single_char_tokens)):
        return line
    words: list[str] = []
    current: list[str] = []
    for t in tokens:
        if t == "":
            if current:
                words.append("".join(current))
                current = []
        else:
            current.append(t)
    if current:
        words.append("".join(current))
    return " ".join(words)


def _normalize_letterspaced_document(text: str) -> str:
    """Collapse letter-spacing across an ENTIRE resume, not just a section header - fixed
    2026-07-31, live case (Divy Parmar / APP-20260720-1013-09AC): some PDF export/template
    pipelines apply the same stylistic letter-spacing to EVERY line, not just headers - name
    ('D I V Y  P A R M A R'), phone ('7 4 0 5 4 6 5 2 0 4'), even the email address
    ('d i v y p a r m a r 1 9 @ g m a i l . c o m'). _collapse_letter_spacing already fixed
    the narrower single-header case (Education); this applies the exact same per-line
    heuristic across every line up front, before any field-level extraction runs, so name/
    phone/email/skills/location all see clean text instead of failing independently. Called
    once at the top of extract_candidate_details[_smart] - idempotent on already-normal
    text, since a line that isn't actually letter-spaced is returned unchanged."""
    if not text:
        return text
    return "\n".join(_collapse_letter_spacing(ln) for ln in text.splitlines())


def _extract_education(text: str) -> str | None:
    """Best-effort degree/school line, most reliable first.

    1. An explicit 'Education' section header - same line if it has content after the
       label, else the next non-header line under it.
    2. A full-text scan for a degree-keyword line (Bachelor's, M.S., PhD, MBA, ...).
    Just a hint for Ollama to confirm/correct (same role as the location/name hints) -
    not meant to be authoritative on its own."""
    lines = [ln.strip() for ln in (text or "").splitlines() if ln.strip()]

    header_re = re.compile(r"^education\b[:\-]?\s*(.*)$", re.IGNORECASE)
    for i, ln in enumerate(lines):
        m = header_re.match(_collapse_letter_spacing(ln))
        if not m:
            continue
        same_line = m.group(1).strip()
        if len(same_line) > 3:
            candidate = normalize_education(same_line.rstrip(".,;"))
            # Require an actual degree keyword before trusting same-line content - fixed
            # 2026-07-31, live case (Brett Worker / APP-20260721-0122-E743): a compound
            # header "EDUCATION & EXECUTIVE DEVELOPMENT" had its "& EXECUTIVE DEVELOPMENT"
            # remainder returned as the Education value outright (not in the small hardcoded
            # exclusion set below), completely skipping the real degree lines just below it
            # ("Robert Morris University" / "Master of Information Systems... Bachelor of
            # Science..."). A header with descriptive trailing words but no real degree
            # content must fall through to scanning the lines under it, same as a bare header.
            if candidate != "Not extracted" and _DEGREE_RE.search(candidate):
                return candidate
        fallback = None
        for nxt in lines[i + 1:i + 6]:
            if nxt.lower().rstrip(":").strip() in SECTION_WORDS:
                break
            cleaned = normalize_education(nxt.rstrip(".,;"))
            if cleaned == "Not extracted":
                continue
            if _DEGREE_RE.search(cleaned):
                return cleaned
            fallback = fallback or cleaned
        if fallback:
            return fallback
        break

    for ln in lines:
        if _DEGREE_RE.search(ln):
            return normalize_education(ln.rstrip(".,;"))
    return None


def normalize_skills(value) -> str:
    """Normalize AI/list-like skill output into the workbook's comma-separated text form."""
    if value is None:
        return "Not extracted"
    if isinstance(value, (list, tuple, set)):
        pieces = [str(v).strip(" \t\r\n'\"") for v in value]
    else:
        text = str(value).strip()
        if not text or text.lower() in _GAP_LITERALS:
            return "Not extracted"
        list_match = re.fullmatch(r"\[\s*(.*?)\s*\]", text)
        if list_match:
            text = list_match.group(1)
        pieces = [p.strip(" \t\r\n'\"") for p in re.split(r"\s*,\s*", text)]
    seen, out = set(), []
    for piece in pieces:
        piece = re.sub(r"\s{2,}", " ", piece).strip(" ,;")
        if not piece or piece.lower() in _GAP_LITERALS:
            continue
        key = piece.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(piece)
    return ", ".join(out) if out else "Not extracted"


def split_location_country(raw: str) -> tuple[str, str]:
    """Split a combined location into (city_state, country)."""
    if not raw or raw.strip().lower() in ("", "not extracted"):
        return ("Not extracted", "")

    parts = [p.strip() for p in raw.split(",")]
    if len(parts) >= 2:
        last = parts[-1].lower().strip()
        # Check if last segment is a known country term
        for term in US_COUNTRY_TERMS:
            if term == last:
                # Infer US state from remaining parts for completeness
                return ", ".join(parts[:-1]), "United States"
        for fc in FOREIGN_COUNTRIES:
            if fc == last:
                # If this "foreign country" is actually a state/province, map to parent country
                # and keep the province as part of the location (e.g. "Mumbai, Maharashtra" → India).
                parent = _PROVINCE_TO_COUNTRY.get(fc)
                if parent:
                    return raw.strip(), parent
                return ", ".join(parts[:-1]), parts[-1].strip().title()

    # Direct province lookup: catches provinces that appear in _PROVINCE_TO_COUNTRY
    # but are NOT listed in the config's FOREIGN_COUNTRIES (e.g. Ontario, British Columbia).
    # This ensures Canadian/Australian/etc. cities get the right country without needing every
    # province name in config.
    if len(parts) >= 2:
        last = parts[-1].lower().strip()
        parent = _PROVINCE_TO_COUNTRY.get(last)
        if parent:
            return raw.strip(), parent

    # No explicit country found — infer "United States" if any part is a US state signal.
    # Keep the ENTIRE raw string as location (do not strip trailing parts that are not
    # recognised country names — e.g. "Maharashtra" is an Indian state, not a country).
    loc = raw.strip()
    for p in parts:
        p_low = p.strip().lower()
        if p_low in US_STATE_ABBREVS or p_low in US_STATE_NAMES:
            return loc, "United States"

    return loc, ""


# ── Skill scanning ───────────────────────────────────────────────────────────

# Skill keywords that collide with an unrelated ALL-CAPS acronym once the scan text is
# lowercased - fixed 2026-07-31, live case (Syyed Nazir Ali / APP-20260727-1431-6F75):
# 'SWIFT' (the banking payment-messaging standard, always written all-caps - "UPI, NEFT,
# RTGS, IMPS, SWIFT, ISO 20022") matched the "Swift" (Apple's language) skill keyword once
# both sides were lowercased to 'swift', which then triggered a mobile-developer Category
# override for a Business Analyst candidate who has never written a line of Swift. These
# keywords require their canonical mixed-case spelling in the ORIGINAL (non-lowercased)
# text - an all-caps-only hit doesn't count. Add more entries here as other collisions turn
# up; do not lowercase-blanket-match a short/acronym-prone keyword without checking first.
_CASE_SENSITIVE_SKILL_FORMS = {"swift": "Swift"}

# Skill keywords that collide with a common English phrase regardless of case - fixed
# 2026-07-31, live case (Mindy Anderson / APP-20260723-2034-12DF): 'Go' (the language)
# matched "Go To Market"/"go-to-market" (ubiquitous marketing/business jargon, and
# title-cased in headings just like the language name would be, so case-sensitivity alone
# can't disambiguate it the way it does for 'Swift') in a resume with zero mention of the
# Go language anywhere. Each pattern is stripped out of the scan text before that keyword's
# bare-word check runs, so a genuine standalone mention elsewhere is still caught.
_SKILL_EXCLUDE_PHRASES = {
    # "go-live"/"go live" added 2026-07-31, live case (Syyed Nazir Ali / APP-20260727-1431-
    # 6F75): "...business sign-off and go-live approval..." and "Release & Go-Live
    # Governance" - standard IT/project-management deployment terminology, not the Go
    # language, same false-positive class as "Go To Market".
    "go": (r"go[\s-]*to[\s-]*market", r"go[\s-]*live"),
    # Marketing-metric "SQL"/"MQL" phrases added 2026-08-01, live case (Sai Krishna Yallapu
    # / APP-20260717-1100-7BDA): "reported MQLs, SQLs, pipeline...", "MQL-to-SQL handoff",
    # "SQL acceptance rates" - Sales-/Marketing-Qualified-Lead counts, standard B2B
    # marketing terminology, not the SQL database language (this candidate's resume never
    # mentions a database anywhere). Only these marketing collocations are excluded; a
    # genuine standalone "SQL" mention (e.g. "SQL Server", "MySQL") is still caught.
    "sql": (r"sqls\b", r"mqls\b", r"mql[\s-]*to[\s-]*sql", r"sql\s+acceptance"),
    # "AWS re:Invent" added 2026-08-01, same Sai Krishna Yallapu row: "Led Oracle AI World,
    # AWS re:Invent, Cloud World, Ascend" names a conference he led sponsorship/marketing
    # for, not a personal cloud-computing skill - this candidate's resume never claims
    # hands-on AWS use anywhere. A genuine standalone "AWS" mention is still caught.
    "aws": (r"aws\s*re:?\s*invent",),
}


def _scan_skill_keywords(text: str) -> list:
    """Return the SKILL_KEYWORDS present in text (lowercased, word-bounded)."""
    raw = text or ""
    low = raw.lower()
    found = []
    for sk in SKILL_KEYWORDS:
        if sk in found:
            continue
        canonical = _CASE_SENSITIVE_SKILL_FORMS.get(sk)
        haystack, needle, boundary = (
            (raw, canonical, r'[a-zA-Z0-9]') if canonical else (low, sk, r'[a-z0-9]')
        )
        for phrase in _SKILL_EXCLUDE_PHRASES.get(sk, ()):
            haystack = re.sub(phrase, ' ', haystack, flags=re.IGNORECASE)
        if re.search(r'(?<!' + boundary + r')' + re.escape(needle) + r'(?!' + boundary + r')', haystack):
            found.append(sk)
    return found


def _strip_unconfirmed_ambiguous_skills(skills_str: str, text: str) -> str:
    """Remove an ambiguous skill (one with a _CASE_SENSITIVE_SKILL_FORMS or
    _SKILL_EXCLUDE_PHRASES entry) from a free-text skills string (e.g. Ollama's own
    extraction) unless _scan_skill_keywords independently confirms it against the same
    resume text. An LLM's free-text skill reading isn't bound by that function's
    disambiguation - it can independently misread the same all-caps acronym or common-
    phrase collision as the ambiguous skill, same root cause, different layer. Reuses
    _scan_skill_keywords as the single source of truth rather than re-implementing the
    same checks a second way."""
    if not skills_str:
        return skills_str
    ambiguous = set(_CASE_SENSITIVE_SKILL_FORMS) | set(_SKILL_EXCLUDE_PHRASES)
    confirmed = set(_scan_skill_keywords(text))
    kept = []
    for part in skills_str.split(","):
        p = part.strip()
        if not p:
            continue
        if p.lower() in ambiguous and p.lower() not in confirmed:
            continue
        kept.append(p)
    return ", ".join(kept)


# ── Portfolio URL extraction ─────────────────────────────────────────────
# Priority:
#   Portfolio 1 → LinkedIn > personal website / portfolio domain
#   Portfolio 2 → GitHub
#   Portfolio 3 → Figma > Behance > Dribbble

# Domains we never want as "portfolio" links
_SKIP_DOMAINS = frozenset({
    "gmail", "yahoo", "hotmail", "outlook", "live", "icloud",
    "google", "docs.google", "drive.google", "play", "apps", "mail", "protonmail",
    "microsoft", "office", "teams", "sharepoint",
    "linkedin", "github", "figma", "behance", "dribbble",
    "instagram", "facebook", "twitter", "x", "tiktok", "youtube",
    "amazon", "aws", "azure", "oracle", "salesforce", "notion",
    "slack", "zoom", "dropbox", "box", "medium", "substack",
    # job boards / ATS links are never personal portfolios
    "indeed", "ziprecruiter", "glassdoor", "monster", "dice",
    "greenhouse", "lever", "workable", "smartrecruiters", "jobvite",
    # document/QR utilities embedded by resume generators are not portfolios
    "qrcode",
    # tech/library brand names that happen to use a "personal-site-shaped" TLD (.io/.dev/
    # .me) are not portfolios either - fixed 2026-08-01, live case (Muhammad Ahsan Hussain
    # / APP-20260721-2030-3AA7): "Socket.io" (a real-time messaging library he uses, never
    # written as a link anywhere in his resume) matched the bare 'word.io' shape and the
    # '.io' personal-site heuristic below, and was stored as Portfolio 1.
    "socket",
})

# Embedded profile URLs (searched INSIDE a matched token — text layers often glue a
# whole contact line into one token, and the profile link is somewhere in the middle).
# The lookbehind stops the match starting mid-hostname, so 'gist.github.com/u' or
# 'docs.github.com/en' can never be misread as 'github.com/...' profiles.
_LINKEDIN_URL_RE = re.compile(
    r'(?<![A-Za-z0-9.])(?:https?://)?(?:www\.)?linkedin\.com/(?:in|pub)/[A-Za-z0-9\-_.%~]+/?', re.I)
_GITHUB_URL_RE = re.compile(
    r'(?<![A-Za-z0-9.])(?:https?://)?(?:www\.)?github\.com/[A-Za-z0-9\-_.]{2,}/?', re.I)

# Scan any URL-like token: full https:// URL or bare domain.tld[/path]
_RAW_URL_RE = re.compile(
    r'https?://[^\s<>"\')\]]+|'
    r'(?:[a-z0-9](?:[a-z0-9\-]{0,61}[a-z0-9])?\.)'      # domain labels
    r'+(?:com|io|dev|me|co|net|org|info|tech|design|art|'
    r'work|site|app|portfolio|edu|uk|in|au|ca|de|fr|jp)'
    r'(?:/[^\s<>"\')\]]*)?',
    re.I,
)


def _clean_url(url: str) -> str:
    url = url.rstrip(".,;:!?)>\"'")
    if not url.lower().startswith(("http://", "https://")):
        url = "https://" + url
    return url


def _domain_root(url: str) -> str:
    """Return the root domain label, e.g. 'github' from 'https://github.com/user'."""
    # Strip scheme first, then take hostname, then strip www prefix
    host = re.sub(r'^https?://', '', url.lower())
    host = host.split("/")[0].split("?")[0]
    host = re.sub(r'^www\.', '', host)
    return host.split(".")[0]


def _normalize_link_artifacts(value: str) -> str:
    """Undo the two text-layer artifacts PDF extraction keeps producing in URLs.

    1. Letter-spaced runs from styled PDFs: 'g i t h u b . c o m / user' -> 'github.com/user'.
    2. Display text glued to the href: 'LinkedInlinkedin.com/in/x' -> 'linkedin.com/in/x'
       (the doubled leading domain word is the anchor text, not part of the URL).
    """
    v = (value or "").strip()
    # 1) collapse letter-spacing only when it clearly IS letter-spacing: every chunk
    #    between spaces is a single character (so normal sentences are never touched).
    if " " in v:
        chunks = v.split(" ")
        if len(chunks) >= 8 and all(len(c) <= 1 for c in chunks):
            v = "".join(chunks)
    # 2) drop a doubled leading site word ('linkedinlinkedin.com' -> 'linkedin.com'),
    #    with or without an https://-prefix before it.
    v = re.sub(r'(?i)\b(linkedin|github|figma|behance|dribbble)(?=\1\.)', '', v)
    return v


def _looks_like_url(value: str) -> bool:
    """True when a string is plausibly a URL: a domain.tld shape, optionally schemed.

    Used to keep AI link *display text* ('linkedin', 'Github-SwiftUI') out of the
    Portfolio columns — presence-in-resume alone is not enough to accept a value.
    """
    v = (value or "").strip().lower()
    if not v or v in _GAP_LITERALS or " " in v:
        return False
    return bool(re.match(
        r'^(?:https?://)?(?:www\.)?[a-z0-9][a-z0-9.\-]*\.[a-z]{2,}(?:[/?#]|$)', v))


def extract_portfolios(text: str) -> tuple[str, str, str]:
    """Return (portfolio_1, portfolio_2, portfolio_3).

    P1 = LinkedIn URL  OR  personal website / portfolio domain
    P2 = GitHub URL
    P3 = Figma / Behance / Dribbble / other creative tool URL
    Missing slots → 'N/A'.
    """
    raw = text or ""
    # PDF text layers can wrap a profile handle at a line boundary.  Join only when
    # the first fragment ends in URL punctuation, which avoids merging an otherwise
    # complete profile with unrelated prose on the next line.
    raw = re.sub(
        r'(?i)((?:linkedin\.com/(?:in|pub)|github\.com)/'
        r'[A-Za-z0-9_.%~\-]*[._~\-])\s*\r?\n\s*'
        r'([A-Za-z0-9][A-Za-z0-9_.%~\-]*/?)',
        r'\1\2',
        raw,
    )
    linkedin = github = figma = behance = dribbble = personal = None

    for m in _RAW_URL_RE.finditer(raw):
        token = _normalize_link_artifacts(m.group(0))
        url = _clean_url(token)
        low = url.lower()
        root = _domain_root(url)
        # PDF text layers often mash a whole contact line into one 'URL' token
        # ('gmail.com/linkedin.com/in/x/github.com/y...'), so pull the embedded
        # profile URL out of the token instead of classifying the token wholesale.
        if not linkedin:
            lm = _LINKEDIN_URL_RE.search(token)
            if lm:
                linkedin = _clean_url(lm.group(0).rstrip("/"))
        if not github:
            gm = _GITHUB_URL_RE.search(token)
            if gm:
                github = _clean_url(gm.group(0).rstrip("/"))
        if root == "figma" and not figma:
            figma = url
        elif root == "behance" and not behance:
            behance = url
        elif root == "dribbble" and not dribbble:
            dribbble = url
        elif root not in _SKIP_DOMAINS and not personal:
            # Only accept URLs the candidate wrote with a scheme, or that are clearly
            # personal sites. Check the ORIGINAL token — _clean_url prepends https://
            # to bare domains, so checking `low` here would always be True.
            #
            # A bare TLD match (word.io/word.dev/word.me/...) with no scheme and no
            # "portfolio" nearby used to be accepted outright - removed 2026-08-01 after
            # TWO independent live false positives in the same afternoon: "Socket.io" (a
            # library Hussain mentions, never a link) and "Loquatinc.io" (a client company
            # name in Mindy Anderson's contracts list, "Contracts include: Loquatinc.io,
            # BNY Mellon, EY..."). Resumes are full of company/technology names that happen
            # to use these trendy TLDs; a bare match alone is not enough signal. Missing a
            # genuine bare "janesmith.io" mention (no scheme, no "portfolio" nearby) now
            # falls through to N/A instead - a safe default, not a fabricated wrong link.
            if token.lower().startswith(("http://", "https://")) or "portfolio" in low:
                personal = url

    p1 = linkedin or personal or "N/A"
    p2 = github or "N/A"
    # When LinkedIn occupies P1, retain a separate personal portfolio in P3.
    p3 = figma or behance or dribbble or (personal if linkedin else None) or "N/A"
    return p1, p2, p3


# ── Phone extraction ────────────────────────────────────────────────────

_PHONE_PATTERNS = [
    re.compile(r'\+1[\s.\-]?\(?\d{3}\)?[\s.\-]?\d{3}[\s.\-]?\d{4}'),
    re.compile(r'\+\d{1,3}[\s.\-]?\(?\d{2,5}\)?[\s.\-]?\d{3,5}[\s.\-]?\d{3,5}'),
    re.compile(r'\(?\d{3}\)?[\s.\-]\d{3}[\s.\-]\d{4}'),
    re.compile(r'\+?\(?\d[\d\s().\-]{8,16}\d'),
]


def format_phone(phone: str) -> str:
    """Canonicalize a US phone to '(XXX) XXX-XXXX' so the column reads uniformly.

    A US 10-digit number (or 11-digit starting with country code 1) is reformatted;
    the parens+dash form also stops Excel from treating a bare digit-run like
    '6677036994' as a NUMBER (which right-aligns it and drops leading zeros) — a
    formatted string is unambiguously text. International (+..) numbers, masked
    numbers ('732-4***'), gaps, and anything that isn't a clean 10/11-digit US
    number are returned unchanged.
    """
    raw = str(phone or "").strip()
    if not raw or raw.lower() in _GAP_LITERALS:
        return raw
    if "*" in raw:                       # deliberately masked — leave as the candidate wrote it
        return raw
    d = "".join(c for c in raw if c.isdigit())
    if len(d) == 11 and d.startswith("1"):        # US with country code (+1 / 1-...)
        return f"({d[1:4]}) {d[4:7]}-{d[7:]}"
    if len(d) == 10 and not raw.lstrip().startswith("+"):   # bare US 10-digit
        return f"({d[0:3]}) {d[3:6]}-{d[6:]}"
    return raw


def sanitize_phone(phone: str) -> str:
    """Blank a phone that isn't a real, complete number.

    A masked ('732-4***') or truncated/too-short value is worse than blank — it reads as
    broken. Anything with a '*' or fewer than 7 digits (and not already a gap literal)
    becomes 'Not extracted' so the column only ever shows a usable number or a clean gap.
    """
    raw = str(phone or "").strip()
    if raw.lower() in _GAP_LITERALS:
        return raw
    if re.fullmatch(r"(?:\d\s+){6,}\d", raw):
        return "Not extracted"
    if "*" in raw or sum(c.isdigit() for c in raw) < 7:
        return "Not extracted"
    return raw


def _extract_phone(text: str) -> str | None:
    """Extract the best phone number from text, preferring labeled numbers."""
    for line in (text or "").splitlines():
        low = line.lower().strip()
        if any(k in low for k in ("phone", "mobile", "cell", "tel:", "contact")):
            searchable = _RAW_URL_RE.sub(" ", line)
            for pat in _PHONE_PATTERNS:
                m = pat.search(searchable)
                if m:
                    return m.group(0).strip()
    searchable_text = _RAW_URL_RE.sub(" ", text or "")
    for pat in _PHONE_PATTERNS:
        m = pat.search(searchable_text)
        if m:
            return m.group(0).strip()
    return None


# ── Province / state → parent country map ─────────────────────────────
# FOREIGN_COUNTRIES mixes sovereign nations with states/provinces.
# This map converts state/province names to their parent country so the
# Country column gets "India" not "Maharashtra".
_PROVINCE_TO_COUNTRY: dict[str, str] = {
    # India — states and major UTs
    "andhra pradesh": "India", "arunachal pradesh": "India", "assam": "India",
    "bihar": "India", "chhattisgarh": "India", "goa": "India", "gujarat": "India",
    "haryana": "India", "himachal pradesh": "India", "jharkhand": "India",
    "karnataka": "India", "kerala": "India", "madhya pradesh": "India",
    "maharashtra": "India", "manipur": "India", "meghalaya": "India",
    "mizoram": "India", "nagaland": "India", "odisha": "India", "orissa": "India",
    "punjab": "India", "rajasthan": "India", "sikkim": "India",
    "tamil nadu": "India", "telangana": "India", "tripura": "India",
    "uttar pradesh": "India", "uttarakhand": "India", "west bengal": "India",
    "delhi": "India", "jammu and kashmir": "India", "ladakh": "India",
    # Canada — provinces and territories
    "alberta": "Canada", "british columbia": "Canada", "manitoba": "Canada",
    "new brunswick": "Canada", "newfoundland": "Canada", "labrador": "Canada",
    "nova scotia": "Canada", "ontario": "Canada", "prince edward island": "Canada",
    "quebec": "Canada", "saskatchewan": "Canada", "yukon": "Canada",
    "northwest territories": "Canada", "nunavut": "Canada",
    # Australia — states and territories
    "new south wales": "Australia", "queensland": "Australia",
    "south australia": "Australia", "tasmania": "Australia",
    "victoria": "Australia", "western australia": "Australia",
    "northern territory": "Australia", "australian capital territory": "Australia",
    # United Kingdom — constituent nations
    "england": "United Kingdom", "scotland": "United Kingdom",
    "wales": "United Kingdom", "northern ireland": "United Kingdom",
    # Germany — Lander (commonly seen on resumes)
    "bavaria": "Germany", "berlin": "Germany", "hamburg": "Germany",
    "hesse": "Germany", "north rhine": "Germany",
    # China — provinces
    "guangdong": "China", "beijing": "China", "shanghai": "China",
    "zhejiang": "China", "sichuan": "China", "hubei": "China",
    # Pakistan
    "punjab pakistan": "Pakistan", "sindh": "Pakistan", "khyber": "Pakistan",
    "balochistan": "Pakistan",
    # Other common cases
    "catalonia": "Spain", "sao paulo": "Brazil", "rio de janeiro": "Brazil",
    "buenos aires": "Argentina", "mexico city": "Mexico", "jalisco": "Mexico",
}

# ── Mail-body fallback extraction ──────────────────────────────────────

def extract_from_mail_body(body: str) -> dict:
    """Extract candidate details from the email body text as a fallback.

    Used when resume extraction returns 'Not extracted' for a field - the
    applicant may have put their details in the email body itself.
    """
    body = body or ""
    result = {}

    email_match = re.search(r'[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}', body)
    if email_match:
        result["email"] = email_match.group(0)

    phone = _extract_phone(body)
    if phone:
        result["phone"] = phone

    name_match = re.search(
        r"(?:(?i:my name is|i am|i'm|this is)[ ]+)"
        r"([A-Z][A-Za-z'\-]+(?:[ ]+[A-Z][A-Za-z'\-]+){1,2})",
        body,
    )
    if name_match:
        result["full_name"] = name_match.group(1).strip().rstrip(".,;")

    loc = _extract_location(body)
    if loc:
        result["location"], result["country"] = split_location_country(loc)

    for line in body.splitlines():
        low = line.lower().strip()
        if any(k in low for k in ("based in", "located in", "living in", "from")):
            loc2 = _extract_location(line)
            if loc2 and "location" not in result:
                result["location"], result["country"] = split_location_country(loc2)
                break

    found = _scan_skill_keywords(body)
    if found:
        from hiring_agent.config import SKILL_DISPLAY, SCORING_MAX_SKILLS
        seen, display = set(), []
        for s in found:
            label = SKILL_DISPLAY.get(s, s.title())
            if label.lower() not in seen:
                seen.add(label.lower())
                display.append(label)
        if display:
            result["skills"] = ", ".join(display[:SCORING_MAX_SKILLS])

    role = _extract_role_from_body(body)
    if role:
        result["looking_for_role"] = role

    return result


def _extract_role_from_body(body: str) -> str | None:
    """Extract a job title from the email body (signature lines, explicit statements)."""
    if not body:
        return None
    # Tried first - fixed 2026-08-01, live case (Brett Worker / APP-20260721-0122-E743):
    # "I came across Tracy Simon's LinkedIn post for the CISO position. The role caught my
    # attention because it combines..." - the role name here comes BEFORE "position", so
    # the keyword-then-role pattern below never had a letter to capture right after
    # "position" (a period followed it) and fell through to matching "role" in the NEXT
    # sentence instead, capturing "caught my attention because it combines" as the desired
    # role. "for/the X position/role" is at least as common a phrasing as "role: X" and
    # must be checked before the fallback pattern gets a chance to latch onto an unrelated
    # later sentence.
    before_match = re.search(
        r'\b(?:for|in)\s+(?:the\s+|an?\s+)?([A-Za-z][A-Za-z0-9 /&\-]{1,39}?)\s+'
        r'(?:position|role|opening|opportunity)\b',
        body, re.IGNORECASE,
    )
    if before_match:
        return _normalize_title(before_match.group(1).strip().rstrip(".,;"))

    role_match = re.search(
        r'\b(?:applying for|interested in|position|role|objective)\s*[:\-]?\s*'
        r'([A-Za-z][A-Za-z /&\-]{2,39})',
        body, re.IGNORECASE,
    )
    if role_match:
        return _normalize_title(role_match.group(1).strip().rstrip(".,;"))

    for line in body.splitlines():
        seg = line.strip()
        if not seg or len(seg) > 60:
            continue
        words = seg.split()
        if 2 <= len(words) <= 6 and any(w.lower().rstrip("s") in ROLE_WORDS for w in words):
            if seg.lower() not in ("work experience", "professional experience",
                                   "experience", "education", "employment history",
                                   "warm regards", "best regards", "kind regards"):
                return _normalize_title(seg)
    return None


def merge_mail_body_fallback(extracted: dict, body: str) -> dict:
    """Fill 'Not extracted' fields from mail body as fallback.

    For 'looking_for_role', the mail body ALWAYS wins when it has a value, because the
    email shows what the candidate is actively seeking (the resume just shows history).
    """
    if not body:
        return extracted
    body_fields = extract_from_mail_body(body)
    merged = dict(extracted)
    for key in ("full_name", "phone", "location", "country", "skills", "email", "looking_for_role"):
        body_val = body_fields.get(key)
        if not body_val:
            continue
        cur = merged.get(key)
        if key == "looking_for_role" and body_val:
            # Mail body role always takes priority — it's what the candidate wants NOW.
            if cur != body_val:
                logger.info(f"       Fallback  : mail body set '{key}': {body_val[:50]}")
            merged[key] = body_val
        elif cur in (None, "", "Not extracted"):
            merged[key] = body_val
            logger.info(f"       Fallback  : mail body filled '{key}': {body_val[:50]}")
    return merged


# ── Offline candidate parser ────────────────────────────────────────────────

def extract_candidate_details(text: str) -> dict:
    """Best-effort OFFLINE extraction from the combined email + resume text."""
    text = _normalize_letterspaced_document(text or "")
    low = text.lower()
    extracted = {
        "full_name": "Not extracted",
        "email": "Not extracted",
        "phone": "Not extracted",
        "location": "Not extracted",
        "country": "",
        "skills": "Not extracted",
        "looking_for_role": "Not extracted",
        "education": "Not extracted",
        "notes": text[:800],
    }

    name_match = re.search(
        r"(?:(?i:my name is)[ ]+|(?i:name)\s*[:\-]\s*)([A-Z][A-Za-z'\-]+(?:[ ]+[A-Z][A-Za-z'\-]+){1,2})",
        text,
    )
    if name_match:
        extracted["full_name"] = name_match.group(1).strip().rstrip(".,;")
    else:
        for line in text.splitlines()[:8]:
            if _looks_like_name(line):
                extracted["full_name"] = _normalize_name(line.strip())
                break

    email_match = re.search(r'[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}', text)
    if email_match:
        extracted["email"] = email_match.group(0)

    phone_match = _extract_phone(text)
    if phone_match:
        extracted["phone"] = phone_match

    header_role = _extract_header_role(text)
    if header_role:
        extracted["looking_for_role"] = header_role
    else:
        # Scoped to the header/summary area only (first 15 lines), not the whole document -
        # fixed 2026-07-31, live case (Prerna Saluja / APP-20260716-2052-6112): "position"
        # and "role" are common English words far beyond "job position" - an Experience
        # bullet reading "...identifying an investment position that appreciated
        # approximately 3x..." matched "position" and returned "that appreciated
        # approximately" as the candidate's desired role. A genuine "seeking X" / "Objective:
        # X" statement lives in the header/summary, same reasoning as the location/name
        # header-first heuristics already used elsewhere in this file.
        # \b after the alternation - fixed 2026-08-01, live case (Mindy Anderson /
        # APP-20260723-2034-12DF): without a trailing word boundary, "position" matched as
        # a bare substring inside "brand positioning" (a business term, not a job-role
        # statement), and the capture group grabbed the word's own leftover letters ("ing")
        # as the "desired role" - identical zero-separator collision to the "rpa" inside
        # "counterparties" bug fixed on the same day.
        _role_header_text = "\n".join(text.splitlines()[:15])
        role_match = re.search(
            r'(?:applying for|application for|position|role|objective)\b\s*[:\-]?\s*([A-Za-z][A-Za-z /&]{2,39})',
            _role_header_text, re.IGNORECASE,
        )
        # Word-boundary matching, not bare substring - fixed 2026-08-01, live case (Prerna
        # Saluja / APP-20260716-2052-6112): 'rpa' in low matched inside 'counterparties',
        # tagging a Financial/Audit Analyst as wanting an "Automation / RPA Engineer" role
        # she never mentioned. Short tokens ('rpa', 'sre', 'seo') are exactly the ones prone
        # to landing mid-word, same root cause as the Swift/Go skill collisions fixed above.
        def _kw_hit(kw: str) -> bool:
            return bool(re.search(r"(?<![a-z])" + re.escape(kw) + r"(?![a-z])", low))

        if role_match:
            extracted["looking_for_role"] = role_match.group(1).strip().rstrip(".")
        elif any(_kw_hit(kw) for kw in ["uipath", "power automate", "rpa", "automation engineer", "process automation"]):
            extracted["looking_for_role"] = "Automation / RPA Engineer"
        elif any(_kw_hit(kw) for kw in ["machine learning", "deep learning", "llm", "genai", "generative ai", "ai engineer", "ml engineer"]):
            extracted["looking_for_role"] = "AI / ML Engineer"
        elif any(_kw_hit(kw) for kw in ["data scientist", "data science", "analytics", "data analyst"]):
            extracted["looking_for_role"] = "Data Scientist / Analyst"
        elif any(_kw_hit(kw) for kw in ["devops", "kubernetes", "terraform", "ci/cd", "site reliability", "sre"]):
            extracted["looking_for_role"] = "DevOps / SRE Engineer"
        elif any(_kw_hit(kw) for kw in ["react", "angular", "vue", "frontend", "front-end", "ui/ux"]):
            extracted["looking_for_role"] = "Frontend Developer"
        elif any(_kw_hit(kw) for kw in ["full stack", "fullstack", "full-stack"]):
            extracted["looking_for_role"] = "Full-Stack Engineer"
        elif any(_kw_hit(kw) for kw in ["python", "django", "flask", "backend", "developer", "engineer", "software"]):
            extracted["looking_for_role"] = "Software Developer / Backend Engineer"
        elif any(_kw_hit(kw) for kw in ["marketing", "content", "seo", "growth"]):
            extracted["looking_for_role"] = "Marketing / Growth"
        elif any(_kw_hit(kw) for kw in ["product manager", "product owner", "scrum master", "agile"]):
            extracted["looking_for_role"] = "Product Manager"
        elif any(_kw_hit(kw) for kw in ["project manager", "program manager", "pmp"]):
            extracted["looking_for_role"] = "Project / Program Manager"

    loc = _extract_location(text)
    if loc:
        extracted["location"], extracted["country"] = split_location_country(loc)

    edu = _extract_education(text)
    if edu:
        extracted["education"] = edu

    found = _scan_skill_keywords(text)
    if found:
        seen, display = set(), []
        for s in found:
            label = SKILL_DISPLAY.get(s, s.title())
            if label.lower() not in seen:
                seen.add(label.lower())
                display.append(label)
        extracted["skills"] = ", ".join(display[:SCORING_MAX_SKILLS])

    p1, p2, p3 = extract_portfolios(text)
    extracted["portfolio_1"] = p1
    extracted["portfolio_2"] = p2
    extracted["portfolio_3"] = p3

    return extracted


# ── Resume text extraction (PDF / DOCX bytes) ────────────────────────────────

def find_tesseract() -> str | None:
    """Path to the Tesseract binary, or None. Checks PATH first, then the standard
    Windows install locations (the UB-Mannheim installer often does not add it to PATH)."""
    import os
    import shutil
    found = shutil.which("tesseract")
    if found:
        return found
    for cand in (r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                 r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                 os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe")):
        if os.path.exists(cand):
            return cand
    return None


def _ocr_pdf(raw: bytes) -> str:
    """OPTIONAL OCR for scanned/image-only PDFs. Returns '' (no-op) unless PyMuPDF +
    pytesseract + the Tesseract binary are all present. Renders each page and OCRs it."""
    try:
        import fitz  # PyMuPDF - renders PDF pages to images (no poppler needed)
        import pytesseract
        from PIL import Image
    except Exception:
        return ""  # OCR libs not installed -> silently fall back to "no text"
    # Point pytesseract at the binary even when it is installed but not on PATH (common on
    # Windows), so OCR works right after Launch.bat's winget install with no manual PATH edit.
    _tess = find_tesseract()
    if _tess:
        pytesseract.pytesseract.tesseract_cmd = _tess
    try:
        parts = []
        doc = fitz.open(stream=raw, filetype="pdf")
        for page in doc:
            pix = page.get_pixmap(dpi=200)
            parts.append(pytesseract.image_to_string(Image.open(io.BytesIO(pix.tobytes("png")))) or "")
        doc.close()
        return "\n".join(parts).strip()
    except Exception as e:
        logger.warning(f"   OCR unavailable ({e}); leaving scanned PDF unread.")
        return ""


def _pdf_hyperlinks(reader) -> list:
    """Collect URI link annotations from every page (best-effort, never raises).

    Resume PDFs very often show only display text ('LinkedIn', 'GitHub') while the real
    URL lives in the page's /Annots — the text layer never contains it. Harvesting the
    annotations gives the portfolio extractor the actual hrefs.
    """
    links = []
    for page in reader.pages:
        try:
            for annot in (page.get("/Annots") or []):
                try:
                    obj = annot.get_object()
                    action = obj.get("/A")
                    if action is None:
                        continue
                    uri = action.get_object().get("/URI")
                    if uri and str(uri).lower().startswith(("http://", "https://", "www.")):
                        links.append(str(uri))
                except Exception:
                    continue
        except Exception:
            continue
    return list(dict.fromkeys(links))   # de-dupe, keep order


def _docx_hyperlinks(document) -> list:
    """Collect external hyperlink targets from a python-docx Document (best-effort)."""
    links = []
    try:
        for rel in document.part.rels.values():
            if "hyperlink" in rel.reltype and rel.is_external:
                t = str(rel.target_ref or "")
                if t.lower().startswith(("http://", "https://", "www.")):
                    links.append(t)
    except Exception:
        pass
    return list(dict.fromkeys(links))


def extract_text_from_bytes(raw: bytes, filename: str) -> str:
    """Extract text from raw PDF/DOCX bytes (with optional OCR fallback for scanned PDFs)."""
    name = (filename or "").lower()
    if not raw:
        return ""
    try:
        if name.endswith(".pdf"):
            from pypdf import PdfReader
            reader = PdfReader(io.BytesIO(raw))
            # Extract page-by-page so one malformed page (pypdf/generator incompatibilities
            # are common on certain PDF producers) doesn't discard the whole document -
            # only that page's text is lost, everything else is kept.
            parts = []
            for i, page in enumerate(reader.pages):
                try:
                    parts.append(page.extract_text() or "")
                except Exception as page_err:
                    logger.warning(f"   page {i + 1} of '{filename or '?'}' failed to parse "
                                   f"({page_err}); skipping that page only.")
            text = "\n".join(parts)
            # Some PDFs technically have a text layer but it is unusable: either only a
            # few stray characters, or every visible word is extracted as spaced letters
            # ("B e n g a l u r u"). Treat both as OCR candidates instead of feeding bad
            # text to location/education extraction.
            _single = len(re.findall(r"\b[A-Za-z]\b", text))
            _words = len(re.findall(r"\b[A-Za-z]{2,}\b", text))
            _bad_text_layer = len(text.strip()) < 250 or _single > max(80, _words * 2)
            if _bad_text_layer:
                ocr_text = _ocr_pdf(raw)
                if len(ocr_text.strip()) > len(text.strip()) / 3:
                    text = ocr_text
                    logger.info("   used OCR (PDF text layer was empty or low quality)")
            if not text.strip():
                try:
                    import pdfplumber
                    with pdfplumber.open(io.BytesIO(raw)) as pdf:
                        plumber_text = "\n".join((page.extract_text() or "") for page in pdf.pages)
                    if plumber_text.strip():
                        text = plumber_text
                        logger.info("   used pdfplumber fallback (pypdf returned no readable text)")
                except Exception as plumber_err:
                    logger.warning(f"   pdfplumber fallback failed for '{filename or '?'}': {plumber_err}")
            # Append link-annotation URLs only when there IS readable text — a link-only
            # result must not make an unreadable resume look readable.
            if text.strip():
                links = _pdf_hyperlinks(reader)
                if links:
                    text += "\nLinks in document: " + " ".join(links)
            return text
        if name.endswith(".docx"):
            import docx
            document = docx.Document(io.BytesIO(raw))
            parts = [p.text for p in document.paragraphs]
            # Table cells too — plenty of resumes lay out contact info in tables,
            # which document.paragraphs alone never sees.
            for tbl in document.tables:
                for row in tbl.rows:
                    for cell in row.cells:
                        if cell.text.strip():
                            parts.append(cell.text)
            text = "\n".join(parts)
            if text.strip():
                links = _docx_hyperlinks(document)
                if links:
                    text += "\nLinks in document: " + " ".join(links)
            return text
    except Exception as e:
        logger.warning(f"   could not read '{filename or '?'}': {e}")
    return ""


# ── AI extraction (Ollama - the only AI brain actually configured in this deployment) ────

def extract_with_ollama(text: str, hints: dict | None = None) -> dict | None:
    """FREE local-LLM extraction via Ollama (no API key). None to fall back.

    Off unless ai_extraction.ollama.enabled (config.yaml) / HIRING_OLLAMA_ENABLED.
    Any failure (Ollama not running, model missing, bad JSON) returns None so the
    caller falls through to the offline parser - nothing breaks.

    `hints`, if given, are the deterministic parser's own guesses for full_name/
    location/country/looking_for_role - passed into the prompt so Ollama confirms or
    corrects a real candidate value instead of deriving the field from a blank slate
    (anchoring reduces hallucination/drift versus an unconstrained guess).
    """
    text = (text or "").strip()
    if not text or not OLLAMA_ENABLED:
        return None
    hint_block = ""
    if hints:
        lines = "\n".join(f"- {k}: {v}" for k, v in hints.items() if v)
        if lines:
            hint_block = (
                "\n\nA deterministic pattern-matcher already found these candidate values "
                "from the text below - confirm each is correct, or correct it if the text "
                "says otherwise:\n" + lines
            )
    try:
        import requests
        resp = requests.post(
            f"{OLLAMA_HOST}/api/chat",
            json={
                "model": OLLAMA_MODEL,
                "format": "json",
                "stream": False,
                "options": {"temperature": 0},
                "messages": [
                    {"role": "system",
                     "content": _AI_PROMPT + hint_block + " Respond ONLY with a JSON object "
                     "whose keys are full_name, phone, location, country, skills, "
                     "looking_for_role, education."},
                    {"role": "user", "content": text[:AI_TEXT_LIMIT]},
                ],
            },
            timeout=OLLAMA_TIMEOUT,
        )
        resp.raise_for_status()
        data = json.loads(resp.json()["message"]["content"])
    except Exception as e:
        logger.warning(f"Ollama extraction unavailable ({e}); using offline parser.")
        return None

    return _ai_fields(data, text)


def _merge_keyword_skills(result: dict, text: str) -> dict:
    """Fold the config skill-keywords found in `text` into an LLM result's `skills`.

    The role matcher intersects on the SKILL_KEYWORDS vocabulary, but LLMs return free-text
    skills that may not use those exact tokens (so match % can read low). Merging the
    keyword-scanned skills keeps the LLM's better name/location while restoring strong role
    scoring. De-duped on display label, capped at SCORING_MAX_SKILLS.
    """
    found = _scan_skill_keywords(text)
    if not found:
        return result
    base = [s.strip() for s in str(result.get("skills", "")).split(",")
            if s.strip() and s.strip().lower() != "not extracted"]
    have = {s.lower() for s in base}
    for sk in found:
        label = SKILL_DISPLAY.get(sk, sk.title())
        if label.lower() not in have:
            base.append(label)
            have.add(label.lower())
    if base:
        result = dict(result)
        if len(base) > SCORING_MAX_SKILLS:
            # Keyword-scanned skills are in the role vocabulary — role matching
            # depends on them, so trim the LLM free-text extras instead.
            kw = {SKILL_DISPLAY.get(sk, sk.title()).lower() for sk in found}
            keyword_part = [s for s in base if s.lower() in kw][:SCORING_MAX_SKILLS]
            free_part = [s for s in base if s.lower() not in kw]
            base = keyword_part + free_part[:SCORING_MAX_SKILLS - len(keyword_part)]
        result["skills"] = ", ".join(base)
    return result


def extract_candidate_details_smart(text: str) -> dict:
    """Deterministic-first extraction, tiered by how reliable each field's regex is.

    The offline/regex parser (extract_candidate_details) always runs first and
    unconditionally, for every field - it's cheap and has no downside.

    Tier 1 (skills, portfolio_1/2/3, phone) - a regex match here is virtually never
    wrong, so the deterministic value is used directly whenever the parser found one;
    Ollama is never even asked to override it (skills is the one exception where AI
    may still ADD extra terms on top, same as before - never replace).

    Tier 2 (full_name, location, country, looking_for_role) - the regex/heuristic only
    gives a *candidate*, not a verdict, so it's passed into Ollama's prompt as a hint
    and Ollama makes the actual call (confirm or correct it) - anchored to a real guess
    instead of deriving the field from a blank slate.

    Falls back to the pure offline parser if Ollama is disabled or the call fails -
    nothing here ever raises.
    """
    text = _normalize_letterspaced_document(text or "")   # see its own docstring
    baseline = extract_candidate_details(text)   # deterministic baseline - always computed

    result = {
        "portfolio_1": baseline.get("portfolio_1", "N/A"),
        "portfolio_2": baseline.get("portfolio_2", "N/A"),
        "portfolio_3": baseline.get("portfolio_3", "N/A"),
        "notes": baseline.get("notes", text[:800]),
    }
    tier2_fields = ("full_name", "location", "country", "looking_for_role", "education")

    if not OLLAMA_ENABLED:
        logger.info("   extractor: offline parser (Ollama disabled)")
        result["phone"] = baseline.get("phone", "Not extracted")
        result["skills"] = baseline.get("skills", "Not extracted")
        for f in tier2_fields:
            result[f] = baseline.get(f, "" if f == "country" else "Not extracted")
        return result

    hints = {f: baseline.get(f, "") for f in tier2_fields}
    hints = {k: v for k, v in hints.items() if v and v.strip().lower() not in _GAP_LITERALS}

    ollama = extract_with_ollama(text, hints)
    if ollama is None:
        logger.info("   extractor: offline parser (Ollama unavailable)")
        result["phone"] = baseline.get("phone", "Not extracted")
        result["skills"] = baseline.get("skills", "Not extracted")
        for f in tier2_fields:
            result[f] = baseline.get(f, "" if f == "country" else "Not extracted")
        return result

    logger.info(f"   extractor: Ollama ({OLLAMA_MODEL}, free local LLM) + deterministic Tier 1")

    # Tier 1 phone: deterministic wins outright if it found something; Ollama's answer
    # only fills a genuine gap, it never overrides a real regex match.
    baseline_phone = baseline.get("phone", "Not extracted")
    if baseline_phone and baseline_phone.strip().lower() not in _GAP_LITERALS:
        result["phone"] = baseline_phone
    else:
        ollama_phone = str(ollama.get("phone", "") or "").strip()
        result["phone"] = ollama_phone if ollama_phone and ollama_phone.lower() not in _GAP_LITERALS else baseline_phone

    # Tier 1 skills: deterministic vocabulary scan is the floor; Ollama's free-text
    # skills only ADD to it (same union pattern as before this plan).
    #
    # Ollama's own free-text skill reading isn't bound by _scan_skill_keywords' case-
    # sensitive disambiguation - it can independently misread the same all-caps acronym
    # (fixed 2026-07-31: 'SWIFT' the banking payment standard, in the presence of the
    # word 'Swift' as this prompt's OWN few-shot example) as the ambiguous skill, same
    # root cause as the regex scanner, just at the LLM layer. Strip any such term from
    # Ollama's output unless its canonical mixed-case spelling genuinely appears in the
    # resume text, before it ever reaches the union merge.
    ollama_skills = _strip_unconfirmed_ambiguous_skills(
        str(ollama.get("skills", "") or ""), text)
    result = _merge_keyword_skills({**result, "skills": ollama_skills or baseline.get("skills", "")}, text)

    # Tier 2: Ollama is the actual decision-maker, confirming or correcting the hint;
    # only fall back to the deterministic guess if Ollama's answer for that field is
    # itself a gap literal.
    for f in tier2_fields:
        v = str(ollama.get(f, "") or "").strip()
        if v and v.lower() not in _GAP_LITERALS:
            result[f] = v
        else:
            result[f] = baseline.get(f, "" if f == "country" else "Not extracted")

    # Safety net: if the resolved location came back combined without a separate
    # country (AI sometimes still does this despite the prompt), split it.
    if not result.get("country") or result.get("country") == "Not extracted":
        loc = result.get("location", "")
        if loc and loc != "Not extracted":
            result["location"], result["country"] = split_location_country(loc)

    # Consistency guard: Ollama can correct 'location' from real text evidence while
    # still echoing a now-stale 'country' hint verbatim - fixed 2026-08-01, live case
    # (Syyed Nazir Ali / APP-20260727-1431-6F75): a stray 'MS' in his skills list made
    # the baseline mis-hint location as "Bloomberg Terminal, MS" and country as "United
    # States"; Ollama correctly read the resume's own header ("Location: India") and
    # fixed location, but left country as the stale "United States" hint, producing a
    # self-contradictory India/United-States row that could pass a country-only geo
    # check. Location is grounded in a confirm/correct step against real text; when the
    # resolved location is ITSELF a bare recognized country name that conflicts with the
    # resolved country, location wins.
    loc_low = str(result.get("location", "")).strip().lower()
    country_low = str(result.get("country", "")).strip().lower()
    if loc_low in FOREIGN_COUNTRIES and country_low in US_COUNTRY_TERMS:
        result["country"] = str(result["location"]).strip().title()
    elif loc_low in US_COUNTRY_TERMS and country_low in FOREIGN_COUNTRIES:
        result["country"] = "United States"

    # Location must hold city/state only - Country is the dedicated field for the country
    # name itself. Fixed 2026-08-01, live case (Syyed Nazir Ali / APP-20260727-1431-6F75):
    # his resume's only location statement is "Location: India" - no city anywhere - so
    # 'India' ended up duplicated into both the Location AND Country columns. If nothing
    # more specific than a bare country name was ever found, don't leave the country name
    # sitting in Location too: use 'Remote' when the text itself says so (this candidate's
    # header also reads "Open to Remote / Global"), else 'N/A' rather than a guess.
    final_loc_low = str(result.get("location", "")).strip().lower()
    if final_loc_low and (final_loc_low in FOREIGN_COUNTRIES or final_loc_low in US_COUNTRY_TERMS):
        if re.search(r'\bremote\b|\bwork[\s-]*from[\s-]*home\b|\bwfh\b', text, re.IGNORECASE):
            result["location"] = "Remote"
        else:
            result["location"] = "N/A"

    return result


# ── AI recheck (post-extraction validation) ────────────────────────────────

_RECHECK_PROMPT = (
    "You are rechecking extracted candidate fields for accuracy. "
    "I will give you the MAIL BODY (the candidate's email), the RESUME TEXT, and the "
    "CURRENT extracted fields. Your job:\n"
    "1. Verify each field. Fix obvious errors (wrong name, partial phone, etc.).\n"
    "2. For 'looking_for_role': this is the role the candidate WANTS. First understand "
    "what they say in the MAIL BODY (e.g. 'applying for X', their email signature title, "
    "or the job they mention). If the mail body has no clue, infer from the resume what "
    "kind of role they are seeking based on their experience and skills.\n"
    "3. For 'location': city and state/region only (no country). "
    "For 'country': the country name.\n"
    "4. For 'skills': keep as comma-separated technical skills/tools only.\n"
    "5. For 'education': the candidate's highest degree and field of study, plus school "
    "if given (e.g. 'B.S. Computer Science, Arizona State University'); empty string if "
    "the resume doesn't mention any degree or schooling.\n"
    "Return ONLY a JSON object with keys: full_name, phone, location, country, skills, "
    "looking_for_role, education. "
    "Use the current value if it is already correct. Use empty string only if truly unknown."
)


def ai_recheck_fields(fields: dict, resume_text: str, mail_body: str) -> dict:
    """Send extracted fields back to Ollama for a final validation pass.

    Returns the improved fields dict, or the original if Ollama is unavailable.
    """
    if not OLLAMA_ENABLED:
        return fields

    context = (
        f"=== MAIL BODY ===\n{(mail_body or '')[:2000]}\n\n"
        f"=== RESUME TEXT ===\n{(resume_text or '')[:AI_TEXT_LIMIT]}\n\n"
        f"=== CURRENT EXTRACTED FIELDS ===\n"
        f"full_name: {fields.get('full_name', '')}\n"
        f"phone: {fields.get('phone', '')}\n"
        f"location: {fields.get('location', '')}\n"
        f"country: {fields.get('country', '')}\n"
        f"skills: {fields.get('skills', '')}\n"
        f"looking_for_role: {fields.get('looking_for_role', '')}\n"
        f"education: {fields.get('education', '')}\n"
    )

    try:
        import requests
        resp = requests.post(
            f"{OLLAMA_HOST}/api/chat",
            json={
                "model": OLLAMA_MODEL,
                "format": "json",
                "stream": False,
                "options": {"temperature": 0},
                "messages": [
                    {"role": "system", "content": _RECHECK_PROMPT},
                    {"role": "user", "content": context},
                ],
            },
            timeout=OLLAMA_TIMEOUT,
        )
        resp.raise_for_status()
        data = json.loads(resp.json()["message"]["content"])
    except Exception as e:
        logger.warning(f"       Recheck   : Ollama unavailable ({e}); keeping original fields.")
        return fields

    improved = dict(fields)
    changed = []
    for key in ("full_name", "phone", "location", "country", "skills", "looking_for_role", "education"):
        new_val = str(data.get(key, "") or "").strip()
        old_val = str(fields.get(key, "") or "").strip()
        # Reject gap-literal strings that Ollama sometimes returns off-prompt
        # (e.g. "Not extracted", "N/A") — never let these overwrite a good value.
        if new_val.lower() in _GAP_LITERALS:
            new_val = ""
        # Phone is a deterministic Tier-1 field the regex extracts reliably. The LLM
        # recheck sometimes returns a mangled fragment ("-8455") that isn't a gap
        # literal but can't be a real number — never let it replace a good phone.
        # Require >= 7 digits before accepting a recheck phone; otherwise keep old.
        if key == "phone" and new_val and sum(ch.isdigit() for ch in new_val) < 7:
            new_val = ""
        if new_val and new_val != old_val:
            changed.append(key)   # counts gap-fills too, so the log never lies
        if new_val:
            improved[key] = new_val
        elif old_val:
            improved[key] = old_val
    improved["education"] = normalize_education(improved.get("education", ""))
    if changed:
        logger.info(f"       Recheck   : AI updated {', '.join(changed)}")
    else:
        logger.info("       Recheck   : AI confirmed all fields OK")
    return improved


_ROLE_SUMMARY_PROMPT = (
    "You write a short note on what role a job candidate is likely seeking, for a recruiter "
    "reading a spreadsheet. First check the MAIL BODY: if the candidate says what role, team, "
    "or kind of work they want, summarize that in your own words. If the mail body gives no "
    "clue, ignore it and instead read the RESUME SKILLS and describe in 1-2 short sentences "
    "the type of role their background suggests they are suited for or seeking. "
    "Always write real content - never answer 'not extracted', 'unknown', 'unclear', or leave "
    "it blank, even if you have to guess from skills alone. Keep it to 1-2 plain sentences, "
    "no bullet points, no labels, under 200 characters. "
    "Return ONLY a JSON object: {\"summary\": \"...\"}"
)


def infer_looking_for_role(resume_text: str, mail_body: str, skills: str) -> str:
    """Guaranteed non-empty 1-2 line description of the role a candidate seems to want.

    Only meant to be called when extraction + ai_recheck_fields both came up empty for
    'looking_for_role' - this replaces the bare 'Not extracted' marker with something a
    recruiter can actually read. Tries Ollama first (reasoning over the mail body, then
    the resume skills); falls back to a deterministic sentence built from the top skills
    if Ollama is off or unreachable, so this never returns an empty/placeholder string.
    """
    if OLLAMA_ENABLED:
        context = (
            f"=== MAIL BODY ===\n{(mail_body or '')[:1500]}\n\n"
            f"=== RESUME SKILLS ===\n{(skills or '')[:600]}\n\n"
            f"=== RESUME TEXT (for context only) ===\n{(resume_text or '')[:AI_TEXT_LIMIT]}\n"
        )
        try:
            import requests
            resp = requests.post(
                f"{OLLAMA_HOST}/api/chat",
                json={
                    "model": OLLAMA_MODEL,
                    "format": "json",
                    "stream": False,
                    "options": {"temperature": 0},
                    "messages": [
                        {"role": "system", "content": _ROLE_SUMMARY_PROMPT},
                        {"role": "user", "content": context},
                    ],
                },
                timeout=OLLAMA_TIMEOUT,
            )
            resp.raise_for_status()
            data = json.loads(resp.json()["message"]["content"])
            summary = str(data.get("summary", "") or "").strip()
            if summary and summary.lower() not in _GAP_LITERALS:
                return summary[:300]
            logger.warning("       Role note  : Ollama returned no usable summary; using skill-based fallback.")
        except Exception as e:
            logger.warning(f"       Role note  : Ollama unavailable ({e}); using skill-based fallback.")

    # ── deterministic fallback (Ollama off, unreachable, or returned nothing usable) ──
    top_skills = [s.strip() for s in (skills or "").split(",") if s.strip()
                  and s.strip().lower() not in _GAP_LITERALS][:3]
    if top_skills:
        return f"No stated role preference; resume skills ({', '.join(top_skills)}) suggest a fit for related technical roles."
    return "No stated role preference and no clear skill signal in the resume to infer one from."


_PORTFOLIO_INFER_PROMPT = (
    "Look for a candidate's professional profile links in the resume text below: a LinkedIn "
    "profile, a GitHub profile, a personal website/portfolio, or a design-portfolio site "
    "(Behance/Dribbble/Figma). Only report a URL that is actually present in the text - "
    "never invent, guess, or complete a partial one. If a category has no real link in the "
    "text, use an empty string for that category. "
    "Return ONLY a JSON object: {\"linkedin\": \"...\", \"github\": \"...\", \"other\": \"...\"}"
)


def _url_grounded_in_text(url: str, text: str) -> bool:
    """Reject a URL the AI returned that doesn't actually appear (or whose identifying
    handle doesn't appear) anywhere in the source text.

    A deterministic anti-hallucination check: local models don't always honor a "never
    invent" instruction — confirmed live during this feature's own testing, where a
    throwaway, link-free input still came back with a plausible-looking fabricated
    LinkedIn URL. This never lets a URL through unless real evidence for it exists.
    """
    if not url:
        return False
    low_text = (text or "").lower()
    url_low = url.lower().rstrip("/")
    bare = url_low.replace("https://", "").replace("http://", "").replace("www.", "")
    if bare in low_text:
        return True
    handle = url_low.rsplit("/", 1)[-1]
    return len(handle) >= 3 and handle in low_text


def infer_missing_portfolios(resume_text: str, current_p1: str, current_p2: str,
                              current_p3: str) -> tuple[str, str, str]:
    """AI gap-fill for whichever Portfolio 1/2/3 slots the regex pass left as 'N/A'.

    Only ever asked about slots that are still empty - a slot the regex already filled
    is trusted outright and never re-examined here (Portfolio is a Tier 1 field; this is
    just its "gap -> ask AI" path, the same pattern as infer_looking_for_role). Gated on
    OLLAMA_ENABLED, wrapped in try/except, never fabricates a URL - 'N/A' stands if the
    AI also finds nothing. Returns the (possibly unchanged) three portfolio values.
    """
    p1 = current_p1 or "N/A"
    p2 = current_p2 or "N/A"
    p3 = current_p3 or "N/A"
    if not OLLAMA_ENABLED or (p1 != "N/A" and p2 != "N/A" and p3 != "N/A"):
        return p1, p2, p3
    try:
        import requests
        resp = requests.post(
            f"{OLLAMA_HOST}/api/chat",
            json={
                "model": OLLAMA_MODEL,
                "format": "json",
                "stream": False,
                "options": {"temperature": 0},
                "messages": [
                    {"role": "system", "content": _PORTFOLIO_INFER_PROMPT},
                    {"role": "user", "content": (resume_text or "")[:AI_TEXT_LIMIT]},
                ],
            },
            timeout=OLLAMA_TIMEOUT,
        )
        resp.raise_for_status()
        data = json.loads(resp.json()["message"]["content"])
    except Exception as e:
        logger.warning(f"       Portfolio  : Ollama unavailable ({e}); keeping N/A.")
        return p1, p2, p3

    def _clean(v):
        v = str(v or "").strip()
        return v if v and v.lower() not in _GAP_LITERALS else ""

    linkedin, github, other = (_clean(data.get("linkedin")), _clean(data.get("github")),
                               _clean(data.get("other")))

    def _accept(raw_value: str, require_host: str | None) -> str:
        """Ground, normalize, and shape-check one AI-returned link; '' if not usable.

        Grounding runs on the RAW value (that's what appears in the resume text);
        the normalized form is what gets stored. `require_host` pins the linkedin/
        github slots to their real domain so display text like 'Github-SwiftUI' —
        which IS in the text — can never pass. The 'other' slot must be URL-shaped
        and not one of the skip domains (mail providers, job boards, ...).
        """
        if not raw_value or not _url_grounded_in_text(raw_value, resume_text):
            return ""
        v = _normalize_link_artifacts(raw_value)
        if not _looks_like_url(v):
            return ""
        v = _clean_url(v)
        if require_host is not None:
            return v if f"{require_host}/" in v.lower() else ""
        return v if _domain_root(v) not in _SKIP_DOMAINS else ""

    taken = {u.lower().rstrip("/") for u in (p1, p2, p3) if u and u != "N/A"}

    def _fill(slot: str, raw_value: str, require_host: str | None) -> str:
        if slot != "N/A":
            return slot
        v = _accept(raw_value, require_host)
        if not v or v.lower().rstrip("/") in taken:   # never repeat a URL across slots
            return "N/A"
        taken.add(v.lower().rstrip("/"))
        return v

    p1 = _fill(p1, linkedin, "linkedin.com")
    p2 = _fill(p2, github, "github.com")
    p3 = _fill(p3, other, None)
    return p1, p2, p3


# ── HTML to text ─────────────────────────────────────────────────────────────

def html_to_text(raw: str) -> str:
    """Strip HTML tags to plain text (used by email body parsing and JD fetching)."""
    if not raw:
        return ""
    text = re.sub(r'(?is)<(script|style).*?</\1>', ' ', raw)
    text = re.sub(r'(?s)<[^>]+>', ' ', text)
    text = unescape(text)
    return re.sub(r'\s+', ' ', text).strip()
