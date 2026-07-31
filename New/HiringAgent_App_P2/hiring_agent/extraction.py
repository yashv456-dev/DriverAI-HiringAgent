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
    "'skills' is a short comma-separated list of individual technical skills or tools only "
    "(e.g. Python, AWS, Docker, React, Swift) - never section headings, categories, or "
    "descriptive phrases; 'looking_for_role' is the job title/role they seek. "
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

def _looks_like_name(s: str) -> bool:
    """True if s looks like a real person's name: 2-3 alpha tokens, none a section word."""
    s = (s or "").strip()
    if not s or "@" in s or any(ch.isdigit() for ch in s):
        return False
    tokens = s.split()
    if not (2 <= len(tokens) <= 3):
        return False
    for t in tokens:
        core = t.replace(".", "").replace("'", "").replace("-", "")
        if len(core) < 2 or not core.isalpha() or t.lower() in SECTION_WORDS:
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
    """Title-case a role line but keep short acronyms upper: 'AI', 'ML', 'QA'."""
    out = []
    for w in (s or "").split():
        if w.isupper() and len(w) <= 3:
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
            if 1 <= len(words) <= 6 and any(w.lower().rstrip("s") in ROLE_WORDS for w in words):
                if seg.lower() not in ("work experience", "professional experience",
                                       "experience", "education", "employment history"):
                    return _normalize_title(seg)
    return None


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

    _city_state_re = re.compile(
        r"([A-Z][A-Za-z.\-]+(?:\s+[A-Z][A-Za-z.\-]+){0,2}),\s*([A-Z]{2})\b"
    )

    for ln in header:
        for seg in re.split(r"\s*[|]\s*", ln):
            m = _city_state_re.match(seg.strip())
            if m:
                return f"{m.group(1)}, {m.group(2)}"

    for ln in header:
        for seg in re.split(r"\s*[|]\s*", ln):
            seg_low = seg.strip().lower()
            parts = re.split(r",\s*", seg_low)
            if len(parts) >= 2 and parts[-1].strip() in US_STATE_NAMES:
                return seg.strip().title()
            # "City, State, Country" — state in second-to-last part, strip trailing country
            if len(parts) >= 3 and parts[-2].strip() in US_STATE_NAMES:
                return ", ".join(p.strip().title() for p in parts[:-1])

    for ln in lines:
        m = _city_state_re.search(ln)
        if m and m.group(2).lower() in US_STATE_ABBREVS:
            return f"{m.group(1)}, {m.group(2)}"

    for scope in (header, lines):
        joined = " \n ".join(scope).lower()
        for city in LOCATION_KEYWORDS:
            if re.search(r"(?<![a-z])" + re.escape(city) + r"(?![a-z])", joined):
                return city.title()

    full_low = " ".join(lines).lower()
    for state in US_STATE_NAMES:
        if re.search(r",\s*" + re.escape(state) + r"(?![a-z])", full_low):
            m2 = re.search(
                r"([A-Z][A-Za-z.\-]+(?:\s+[A-Z][A-Za-z.\-]+){0,2}),\s*"
                + re.escape(state.title()), " ".join(lines))
            if m2:
                return f"{m2.group(1)}, {state.title()}"
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
                    for seg in re.split(r"\s*[|]\s*", ln):
                        if country in seg.lower():
                            seg = seg.strip()
                            return seg if _location_like(seg) else country.title()
        for city in FOREIGN_CITIES:
            if re.search(r"(?<![a-z])" + re.escape(city) + r"(?![a-z])", joined):
                for ln in scope:
                    for seg in re.split(r"\s*[|]\s*", ln):
                        if city in seg.lower():
                            seg = seg.strip()
                            return seg if _location_like(seg) else city.title()

    return None


# Degree keywords, longest/most-specific alternatives first so e.g. "Bachelor of Science"
# matches whole rather than stopping at a shorter overlapping alternative.
_DEGREE_RE = re.compile(
    r"\b(Bachelor(?:'s)?(?:\s+of\s+\w+)?|Master(?:'s)?(?:\s+of\s+\w+)?|"
    r"Associate(?:'s)?(?:\s+of\s+\w+)?|Ph\.?D\.?|MBA|B\.?Tech\.?|M\.?Tech\.?|"
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
        m = header_re.match(ln)
        if not m:
            continue
        same_line = m.group(1).strip()
        if len(same_line) > 3:
            candidate = normalize_education(same_line.rstrip(".,;"))
            if candidate != "Not extracted":
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

def _scan_skill_keywords(text: str) -> list:
    """Return the SKILL_KEYWORDS present in text (lowercased, word-bounded)."""
    low = (text or "").lower()
    found = []
    for sk in SKILL_KEYWORDS:
        if sk not in found and re.search(r'(?<![a-z0-9])' + re.escape(sk) + r'(?![a-z0-9])', low):
            found.append(sk)
    return found


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
            if (token.lower().startswith(("http://", "https://")) or "portfolio" in low
                    or low.split(".")[-1].split("/")[0] in ("io", "dev", "me", "design", "art")):
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

    found = _scan_skill_keywords(body.lower())
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
    role_match = re.search(
        r'(?:applying for|interested in|position|role|objective)\s*[:\-]?\s*'
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
    text = text or ""
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
        role_match = re.search(
            r'(?:applying for|application for|position|role|objective)\s*[:\-]?\s*([A-Za-z][A-Za-z /&]{2,39})',
            text, re.IGNORECASE,
        )
        if role_match:
            extracted["looking_for_role"] = role_match.group(1).strip().rstrip(".")
        elif any(kw in low for kw in ["uipath", "power automate", "rpa", "automation engineer", "process automation"]):
            extracted["looking_for_role"] = "Automation / RPA Engineer"
        elif any(kw in low for kw in ["machine learning", "deep learning", "llm", "genai", "generative ai", "ai engineer", "ml engineer"]):
            extracted["looking_for_role"] = "AI / ML Engineer"
        elif any(kw in low for kw in ["data scientist", "data science", "analytics", "data analyst"]):
            extracted["looking_for_role"] = "Data Scientist / Analyst"
        elif any(kw in low for kw in ["devops", "kubernetes", "terraform", "ci/cd", "site reliability", "sre"]):
            extracted["looking_for_role"] = "DevOps / SRE Engineer"
        elif any(kw in low for kw in ["react", "angular", "vue", "frontend", "front-end", "ui/ux"]):
            extracted["looking_for_role"] = "Frontend Developer"
        elif any(kw in low for kw in ["full stack", "fullstack", "full-stack"]):
            extracted["looking_for_role"] = "Full-Stack Engineer"
        elif any(kw in low for kw in ["python", "django", "flask", "backend", "developer", "engineer", "software"]):
            extracted["looking_for_role"] = "Software Developer / Backend Engineer"
        elif any(kw in low for kw in ["marketing", "content", "seo", "growth"]):
            extracted["looking_for_role"] = "Marketing / Growth"
        elif any(kw in low for kw in ["product manager", "product owner", "scrum master", "agile"]):
            extracted["looking_for_role"] = "Product Manager"
        elif any(kw in low for kw in ["project manager", "program manager", "pmp"]):
            extracted["looking_for_role"] = "Project / Program Manager"

    loc = _extract_location(text)
    if loc:
        extracted["location"], extracted["country"] = split_location_country(loc)

    edu = _extract_education(text)
    if edu:
        extracted["education"] = edu

    found = _scan_skill_keywords(low)
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
    result = _merge_keyword_skills({**result, "skills": ollama.get("skills", baseline.get("skills", ""))}, text)

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
