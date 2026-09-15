"""Resume text extraction, candidate detail parsing, name resolution, and skills."""

import io
import json
import re
import struct
from html import unescape

from hiring_agent.config import (
    SKILL_KEYWORDS, SKILL_DISPLAY, SECTION_WORDS, ROLE_WORDS,
    LOCATION_KEYWORDS, AI_TEXT_LIMIT,
    OLLAMA_ENABLED, OLLAMA_MODEL, OLLAMA_HOST, OLLAMA_TIMEOUT, OLLAMA_SEED,
    REQUIRE_AI,
    SCORING_MAX_SKILLS, US_STATE_ABBREVS, US_STATE_NAMES,
    US_STATE_ABBREV_TO_NAME, US_STATE_NAME_TO_ABBREV, US_TERRITORY_NAMES,
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
#: What a cell says when the CANDIDATE still owes us the value. Distinct from "N/A", which
#: means "asked and there genuinely is none" (portfolio slots). Added 2026-08-04 on client
#: instruction: a blank cell is ambiguous - a reader cannot tell whether the pipeline failed,
#: nobody looked yet, or the candidate simply has no such detail. "Missing" states plainly
#: that we need it and are chasing it by email.
MISSING_VALUE = "Missing"

#: Every literal that means "no real value here". MISSING_VALUE is deliberately INCLUDED, so
#: a cell reading "Missing" still counts as a gap everywhere downstream: healing keeps trying
#: to fill it from a newer resume, and _missing_fields_list keeps asking the candidate for it.
#: Were it treated as a real value, "Missing" would freeze the row permanently.
_GAP_LITERALS = {"not extracted", "n/a", "na", "none", "-", "not found", "unknown", "",
                 "missing", "not confirmed"}


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


#: Technology / discipline vocabulary that is never a personal name. Added 2026-08-06 after
#: live row APP-20260612-0905-DAA7 stored Full Name as 'Cross-Platform Mobile' - lifted off
#: the skills line of Abdurrahman's CV ('Flutter & Dart, Cross-Platform Mobile Apps, ...').
#: Two alpha tokens, no digits, no '@', no section word and no company word, so every
#: existing shape check passed it, and the garbage then propagated into the stored resume
#: filename ('CrossPlatform_Mobile_..._DAA7.pdf') exactly as 'for Scientific Computing.' did
#: on A455 - the same failure mode, a different vocabulary.
#:
#: Deliberately EXCLUDES occupational words that are also common surnames (Baker, Miller,
#: Cook, Taylor, Marshall, Porter, Mason, Fisher, Gardner, Walker, Wright, Hunter). Every
#: word here is a technology, a platform or a discipline noun, none of which appears in a
#: real person's name - which is what makes a single-token match safe enough to reject on.
_TECH_PHRASE_WORDS = {
    "platform", "mobile", "frontend", "backend", "fullstack", "stack", "android",
    "ios", "web", "app", "apps", "application", "applications", "api", "apis",
    "database", "cloud", "devops", "framework", "frameworks", "sdk", "native",
    "responsive", "microservices", "algorithms", "analytics", "cybersecurity",
    "blockchain", "automation", "scripting", "testing", "debugging", "deployment",
    "architecture", "integration", "javascript", "typescript", "python", "java",
    "kotlin", "swift", "flutter", "react", "angular", "node", "django", "sql",
    "nosql", "ui", "ux", "qa",
}


#: Document/scanner artifacts that arrive as a "name". Added 2026-08-06 after live row
#: APP-20260602-0447-0AE4 stored Full Name as 'CamScanner' - the scanning app writes itself
#: into the PDF's producer metadata and title, the CV had no usable name line, and the
#: artifact won. It reached the client sheet and the resume filename
#: ('CamScanner_CloudAndDevOps_0AE4.pdf'). Same failure shape as 'for Scientific Computing.'
#: (A455) and 'Cross-Platform Mobile' (DAA7): a non-name that satisfies every shape check.
#:
#: Matched on the WHOLE string, not per token, so a real name that merely contains one of
#: these words survives - 'Scott Word', 'Ana Documenta', 'Cam Newton' are all unaffected.
_DOCUMENT_ARTIFACT_NAMES = {
    "camscanner", "adobescan", "adobeacrobat", "tapscanner", "officelens",
    "microsoftword", "newmicrosoftworddocument", "newdocument", "worddocument",
    "word", "document", "documents", "untitled", "untitleddocument",
    "scanneddocument", "scanner", "scan", "resume", "myresume", "finalresume",
    "resumecopy", "cv", "mycv", "curriculumvitae", "doc", "doc1", "pdf", "docx",
}


def _is_document_artifact_name(s: str) -> bool:
    """True when the whole 'name' is a scanner/word-processor artifact, not a person.

    Punctuation and spacing are stripped before comparison, so 'Cam Scanner',
    'CamScanner' and 'cam-scanner' all resolve to the same key.
    """
    key = re.sub(r"[^a-z0-9]", "", str(s or "").lower())
    return bool(key) and key in _DOCUMENT_ARTIFACT_NAMES


def _looks_like_tech_phrase(s: str) -> bool:
    """True if any token in s is technology/discipline vocabulary rather than a name.

    Splits on non-letters as well as whitespace, so the hyphenated 'Cross-Platform' is seen
    as 'cross' + 'platform' and matches on 'platform' - a plain token comparison would miss
    it, since _looks_like_company_name only strips punctuation from the token EDGES.
    """
    for raw in (s or "").split():
        for part in re.split(r"[^A-Za-z]+", raw):
            if part and part.lower() in _TECH_PHRASE_WORDS:
                return True
    return False


#: Lowercase English function words. A person's name never STARTS with one, but a sentence
#: fragment scraped out of resume prose routinely does. Kept deliberately small and specific
#: so real name particles ("van", "de", "la", "bin", "al") are untouched - those appear mid-
#: name anyway ("Maria de la Cruz"), never as the first token of a stored full name.
_NAME_STOPWORD_START = {
    "for", "the", "and", "of", "in", "at", "to", "a", "an", "with", "by", "on",
    "from", "as", "is", "are", "was", "were", "or", "but", "via", "per",
}


def strip_parenthetical_nickname(s: str) -> str:
    """Drop a parenthesised preferred name: 'Tianxin (William) Wang' -> 'Tianxin Wang'.

    Added 2026-08-04, and this is the ROOT CAUSE behind live row APP-20260716-0312-A455.
    His resume's very first line reads 'Tianxin (William) Wang', but the token '(William)'
    fails .isalpha(), so every name-shape check rejected the correct name outright. Having
    rejected line 1, the header scan fell through to later lines and settled on the prose
    fragment 'for Scientific Computing.' - which then propagated into the stored resume
    filename on SharePoint. Fixing the placement of a nickname fixes the whole chain.

    A parenthesised alias is common in resumes from candidates who use an anglicised or
    shortened name ('Robert (Bob) Smith'), so this is not a one-off workaround.
    """
    s = re.sub(r"\s*\([^)]*\)\s*", " ", str(s or ""))
    return re.sub(r"\s{2,}", " ", s).strip()


def _name_is_sentence_fragment(s: str) -> bool:
    """True when a 'name' is really a scrap of prose the parser or the model picked up.

    Added 2026-08-04. Live case APP-20260716-0312-A455: Full Name was stored as
    'for Scientific Computing.' - three alpha tokens, no digits, no '@', so it satisfied
    every existing shape check - and the garbage then propagated into the resume FILE name
    on SharePoint ('forComputing_AIMLCVSIN2_A455.pdf'), so the bad value outlived the cell.
    The candidate is Tianxin Wang, who signs his mail "William".

    Two signals, both cheap and both safe:
      * the first token is a lowercase English function word - names do not begin "for";
      * the LAST token is a multi-letter word ending in '.', which reads as the end of a
        sentence rather than a middle initial ('Feld', 'L.' and 'Jr.' all stay valid).
    """
    s = (s or "").strip()
    if not s:
        return False
    tokens = s.split()
    if not tokens:
        return False
    if tokens[0].lower().strip(".,") in _NAME_STOPWORD_START:
        return True
    last = tokens[-1]
    core = last.replace(".", "")
    if last.endswith(".") and len(core) > 2 and core.lower() not in ("jr", "sr", "phd", "esq"):
        return True
    return False


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
    s = strip_parenthetical_nickname((s or "").strip())
    if not s or "@" in s or any(ch.isdigit() for ch in s):
        return False
    if (_looks_like_company_name(s) or _name_is_sentence_fragment(s)
            or _looks_like_tech_phrase(s) or _is_document_artifact_name(s)):
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


#: Name particles that are conventionally lowercase INSIDE a full name ('Maria de la Cruz',
#: 'Ludwig van Beethoven'). Exempted from the all-lowercase title-casing below so fixing
#: 'aashish sachaniya' does not turn 'de la' into 'De La'.
_NAME_PARTICLES = {
    "de", "del", "della", "di", "da", "dos", "das", "la", "le", "van", "von",
    "der", "den", "ter", "ten", "bin", "ibn", "al", "el", "op",
}


def _normalize_name(s: str) -> str:
    """'YASH VERMA' -> 'Yash Verma', 'aashish sachaniya' -> 'Aashish Sachaniya'; leave
    genuinely mixed-case tokens ('McDonald', 'DeSilva') unchanged.

    All-LOWERCASE tokens were left untouched until 2026-08-06, so two live rows reached the
    client sheet uncapitalised - APP-20260604-0818-44A6 stored 'aashish sachaniya' and
    APP-20260602-1043-B88E stored 'maftabsabir' - and both then propagated into the resume
    FILE name ('aashish_sachaniya_44A6.pdf'). An all-lower token is never a deliberate
    spelling the way a mixed-case one is; it is just an un-capitalised source, so it gets
    the same treatment all-upper already got.

    Also drops a parenthesised preferred name so the stored value is the legal name:
    'Tianxin (William) Wang' -> 'Tianxin Wang'. The shape checks already ignore the
    parenthetical, so normalising here keeps the stored cell consistent with what was
    validated - and keeps it out of the derived resume FILE name.
    """
    s = strip_parenthetical_nickname(s)
    out = []
    for t in (s or "").split():
        if t.islower() and t.strip(".'-").lower() in _NAME_PARTICLES:
            out.append(t)
        elif t.isupper() or t.islower():
            out.append(t.title())
        else:
            out.append(t)
    return " ".join(out)


def _plausible_name_shape(s: str) -> bool:
    """Looser structural check for an AI/regex-PARSED name candidate: 1-4 alpha tokens,
    none a section word, no digits/@.

    Deliberately wider than _looks_like_name's 2-3-token window (a parsed name may
    legitimately be a single name or include a middle name), but still catches an AI
    slip-up like echoing back a full sentence or a role title as the "name" - a gap
    found during a final review: the old check here only rejected on a section-word
    overlap, not on the string actually having name-like shape at all.
    """
    s = strip_parenthetical_nickname((s or "").strip())
    if not s or "@" in s or any(ch.isdigit() for ch in s):
        return False
    if (_looks_like_company_name(s) or _name_is_sentence_fragment(s)
            or _looks_like_tech_phrase(s) or _is_document_artifact_name(s)):
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
    """Decide the candidate's name, most-trustworthy source first.

    Precedence is deliberately CV-FIRST: both resume-derived sources are exhausted
    before the sender's display name is considered.

    1. `parsed_name`  - what the extractor read out of the CV (AI or regex)
    2. `resume_text`  - a name-shaped line in the CV's first 20 lines
    3. `sender_name`  - the mailbox display name, only if the CV yields nothing
    4. raw sender     - last resort

    Steps 2 and 3 were the other way round until 2026-08-01. That ordering is wrong
    whenever the sender is not the candidate: a staffing agency or vendor forwarding
    someone else's CV ("please find the attached resume of my developer") would have the
    AGENT's name written into Full Name, mis-attributing the row to a person who is not
    the applicant. Live case: an agency's BD manager submitted a Flutter developer's CV,
    and the row was created under the agent's name while the real candidate's name sat in
    the attached PDF. Such submissions are accepted as normal candidates as long as a real
    resume is attached, so the CV must always be the authority on who the candidate is.
    """
    p = (parsed_name or "").strip()
    if p and p != "Not extracted" and _plausible_name_shape(p):
        return _normalize_name(p)
    _cv_lines = [ln.strip() for ln in (resume_text or "").splitlines() if ln.strip()]
    for line in _cv_lines[:20]:
        if _looks_like_name(line):
            return _normalize_name(line.strip())
    # A mononym ("Shivam") is a single token, so _looks_like_name - which requires two
    # full words - can never match it. Accepted here only in the CV's first 3 non-blank
    # lines, where a lone capitalised word is the candidate's name rather than prose;
    # anywhere further down this would start matching section headings and job titles.
    for line in _cv_lines[:3]:
        core = line.replace(".", "").replace("'", "").replace("-", "")
        if (core.isalpha() and 2 <= len(core) <= 20 and line[:1].isupper()
                and line.lower() not in SECTION_WORDS
                and not _looks_like_company_name(line)
                and not _is_document_artifact_name(line)
                and not _looks_like_tech_phrase(line)):
            return _normalize_name(line)
    if _looks_like_name(sender_name):
        return _normalize_name(sender_name.strip())
    # Last resort: the raw sender display name. Normalised like every other source since
    # 2026-08-06 - it was returned verbatim before, which is how 'maftabsabir' (B88E) was
    # stored as a Full Name and baked into that row's resume filename. Company and tech
    # vocabulary is rejected outright here rather than capitalised: a vendor whose display
    # name reads 'Infilon Technologies' is not a candidate, and 'Missing' is the honest cell.
    s = (sender_name or "").strip()
    if (s and "@" not in s and not any(c.isdigit() for c in s)
            and not _looks_like_company_name(s) and not _looks_like_tech_phrase(s)
            and not _is_document_artifact_name(s)):
        return _normalize_name(s)
    return "Not extracted"


# ── Title / location helpers ─────────────────────────────────────────────────

def clean_role_text(s: str) -> str:
    """Cut trailing contact details off a role title.

    Added 2026-08-04. Live row APP-20260706-0752-8006 stored Looking For Role as
    'Mobile Application Developer Email: ashfaq.fullstackdev@gmail.com' - the resume header
    put the title and the contact block on one line, and everything after the title came
    with it. Truncates at the first contact marker ('Email:', 'Phone:', 'Mobile:', 'Tel:',
    'Contact:', a bare address, or a URL) and keeps what precedes it.
    """
    t = str(s or "").strip()
    if not t:
        return t
    t = re.split(r"(?i)\b(?:e-?mail|phone|mobile|tel|contact|linkedin|github)\b\s*[:\-|]",
                 t)[0]
    t = re.split(r"\S+@\S+|https?://\S+", t)[0]
    return t.strip(" ,;|-–—:").strip()


def _normalize_title(s: str) -> str:
    """Title-case a role line but keep short acronyms upper: 'AI', 'ML', 'QA', 'CISO'.

    Cap raised 3 -> 4 chars - fixed 2026-08-01, live case (Brett Worker /
    APP-20260721-0122-E743): 'CISO' (a genuine 4-letter role acronym extracted from his
    mail body) was getting reduced to 'Ciso'.
    """
    s = clean_role_text(s)   # drop any trailing 'Email: ...' / URL picked up off the header
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


#: A school's name, including the part that runs PAST the keyword: "University of Delhi",
#: "Institute of Industrial Technology". Without the "of X" tail, stripping "University"
#: off "University of Delhi" would leave "of Delhi" and hand back the city that is the
#: whole reason the guard exists.
_INSTITUTION_NAME_RE = re.compile(
    r"(?i)\b(?:university|college|institute|academy|polytechnic|school)"
    r"(?:\s+(?:of|at|for)\s+(?:the\s+)?[A-Z][A-Za-z.\-]*(?:\s+[A-Z][A-Za-z.\-]*){0,3})?")


def _past_institution_name(line: str) -> str:
    """The part of an education line that follows the school's own name, or "".

    A city on a school's line is where the candidate is only when it comes AFTER the name -
    "Arizona State University, Tempe, Arizona" is a campus address. A city INSIDE the name
    says nothing about where anyone lives: "University of Delhi" (live case Prerna Saluja /
    APP-20260716-2052-6112, whose only 'Delhi' is her undergrad alma mater) and "San Jose
    State University" are schools, not addresses.

    Everything up to and including the LAST school name is dropped, so a two-school line
    ("Daulat Ram College, University of Delhi") yields nothing at all.

    Splitting on commas is not enough: a school and its campus routinely share one comma
    segment ("Arizona State University (3.72/4) Arizona, USA"), and a school's name just as
    routinely runs PAST the keyword ("University of Delhi", "Institute of Industrial
    Technology"), so the name's own extent has to be measured either way.
    """
    end = 0
    for m in _INSTITUTION_NAME_RE.finditer(line):
        end = max(end, m.end())
    return line[end:].strip(" ,;:-\u2013\u2014") if end else line.strip()


#: US state names that are also common personal first/last names. "<Word> <StateName>" in a
#: resume header is far more often a person than a place for these, so the comma-less
#: full-state form (step 1c) only accepts them alongside a ZIP code.
_NAME_LIKE_STATE_NAMES = {"washington", "virginia", "georgia", "montana", "indiana",
                          "carolina", "dakota", "jersey"}


def _location_like(seg: str) -> bool:
    """A real location segment is short ('Mumbai, Maharashtra, India'); a prose sentence
    that merely mentions a country is not. Module-level (not nested in _extract_location)
    because both the bare-city pass and the foreign-country pass need it."""
    return len(seg) <= 60 and len(seg.split()) <= 6


#: Bullet/sentence markers. A resume address is a short label; a duty bullet is a sentence.
_PROSE_MARKER_RE = re.compile(
    r"(?i)(^\s*[•●▪·*\-–—]\s|"          # leading bullet glyph
    r"\b(analy[sz]|develop|led|built|design|implement|manag|improv|"  # duty verbs
    r"research|studi|report|present|deliver|support|creat|migrat)\w*\b|"
    r"'s\s|’s\s)"                                               # possessive: "Singapore's age"
)


def _mentions_place_as_prose(seg: str) -> bool:
    """True when a country/city name in this segment is the SUBJECT of a sentence rather
    than the candidate's address.

    _location_like already rejects a long segment from being returned verbatim, but the
    fallback branch below it used to return the bare country name anyway - so a place named
    anywhere in prose still became the candidate's Location. Live case 2026-08-04, Hetvi
    Shah (APP-20260601-1700-3A6E): her CV says "Arizona State University, Tempe, Arizona"
    and carries a (602) Phoenix phone, but one duty bullet reads "...analyzing Singapore's
    age, gender, and ethnicity trends...". 'Singapore' was returned as her Location, Ollama
    then split it into Location 'Remote' / Country 'Singapore', and a US-based candidate
    landed in location review with a foreign country on her row.

    A demographic dataset she analysed is a TOPIC. Requiring the match to sit in an
    address-shaped segment (short, no duty verb, no possessive, not a bullet) keeps real
    addresses working - 'Bangalore, India' and 'Mumbai, Maharashtra, India' are still
    returned by _location_like above, long before this guard is consulted.
    """
    return bool(_PROSE_MARKER_RE.search(str(seg or "")))


def _segment_names_a_country(seg: str) -> bool:
    """True when a comma-separated segment ENDS in an explicit country or known province.

    Used to decide whether a bare city-keyword hit should be widened to the whole segment
    the candidate actually wrote. Deliberately strict — only a real country/province tail
    qualifies — so an unrelated line that merely contains a city name ("Skills: Python,
    Boston Dynamics") is never mistaken for an address.
    """
    parts = [p.strip().lower() for p in str(seg or "").split(",")]
    if len(parts) < 2 or not parts[-1]:
        return False
    last = parts[-1]
    return (last in _PROVINCE_TO_COUNTRY
            or any(fc == last for fc in FOREIGN_COUNTRIES)
            or last in US_COUNTRY_TERMS)


#: Headings that END the contact block at the top of a resume. Everything above the first
#: of these is where a candidate states their OWN address; everything below is history.
_CONTACT_BLOCK_END_RE = re.compile(
    r"^(?:professional\s+|work\s+|technical\s+|career\s+|relevant\s+|key\s+)?"
    r"(?:summary|profile|objective|about(?:\s+me)?|overview|education|experience|"
    r"employment|skills|projects?|certifications?|publications?|awards?|achievements?|"
    # TOOLS/TECHNOLOGIES/LANGUAGES/COURSEWORK matter as much as the obvious ones: a tools
    # list left inside the contact block gives 'Bloomberg Terminal, MS Project' -> the
    # bogus city 'Project, MS' (Mississippi is a real state abbrev, so the state check
    # cannot catch it). Live regression, Syyed Nazir Ali / APP-20260727-1431-6F75.
    r"tools?|technolog(?:y|ies)|languages?|coursework|interests?|references?)"
    r"\b", re.IGNORECASE)

#: A run of 3+ spaces is a PDF LAYOUT GUTTER (a column break), not word spacing. Two spaces
#: is what a letter/word-spaced PDF puts between ordinary words, so the threshold sits above
#: it deliberately.
_GUTTER_RE = re.compile(r"[ \t]{3,}")


def _normalize_spacing(line: str) -> str:
    """Squeeze a PDF's doubled word spacing without gluing its columns together.

    Live 2026-09-05: 'Gulnara  Timokhina  Palo  Alto,  CA  94303   gtimokhina@...' produced
    the Location "Palo  Alto, CA" - the right place, unusable as a value - while
    'University of Massachusetts Amherst<gutter>Amherst,  MA  Master of Science' produced
    nothing at all, because no pass could see 'Amherst, MA' as its own field.

    Gutters become '|', the separator every segment split below already handles; only the
    narrow runs are collapsed.
    """
    return re.sub(r"\s{2,}", " ", _GUTTER_RE.sub(" | ", line)).strip()


#: Hard cap so a resume with no recognisable heading cannot turn the "contact block" into
#: the whole document - which would reintroduce the very bug this exists to prevent.
_CONTACT_BLOCK_MAX_LINES = 12

#: A duty bullet is never contact information. Some PDFs collapse a whole resume onto a
#: handful of very long lines with no heading at the start of any of them, and without this
#: the block ran to its cap and scanned twelve bullets for an address.
_BULLET_LINE_RE = re.compile(r"^\s*[\u2022\u25cf\u25aa\u25e6\u2023\u2219\u00b7*]|^\s*[-\u2013\u2014]\s")

#: The same headings as _CONTACT_BLOCK_END_RE but matched MID-LINE, and only in caps. A
#: collapsed PDF line reads '... |Portfolio SUMMARY | Full Stack Software Engineer ...',
#: where everything from SUMMARY on belongs to the body. Requiring caps is what keeps this
#: off a genuine contact line that happens to contain the word 'skills' or 'projects'.
_EMBEDDED_HEADING_RE = re.compile(
    r"(?<=\s)(?:PROFESSIONAL\s+|WORK\s+|TECHNICAL\s+|CAREER\s+|RELEVANT\s+|KEY\s+)?"
    r"(?:SUMMARY|PROFILE|OBJECTIVE|OVERVIEW|EDUCATION|EXPERIENCE|EMPLOYMENT|SKILLS|"
    r"PROJECTS?|CERTIFICATIONS?|PUBLICATIONS?|AWARDS?|ACHIEVEMENTS?)\b")


def _contact_block(lines: list) -> list:
    """The lines above the first section heading - a resume's own address, not its history.

    Replaces a flat lines[:6] slice, which silently degraded to six fragments on any PDF
    that extracts one word per line. Bounded by _CONTACT_BLOCK_MAX_LINES so an unheaded
    resume still yields a header, never the entire document.
    """
    out = []
    for ln in lines:
        if out and _CONTACT_BLOCK_END_RE.match(_collapse_letter_spacing(ln)):
            break
        if _BULLET_LINE_RE.match(ln):
            break
        cut = _EMBEDDED_HEADING_RE.search(ln)
        if cut and cut.start():
            out.append(ln[:cut.start()].rstrip())
            break
        out.append(ln)
        if len(out) >= _CONTACT_BLOCK_MAX_LINES:
            break
    return out


#: 'Present' outranks every real date. A resume that says "Oct 2024 - Present" is stating
#: where the candidate is NOW, which is the whole question this module is trying to answer.
_ENTRY_PRESENT_KEY = 9999 * 12

#: A place may follow the school/employer on the SAME line, or sit on the line just below
#: the dates - both layouts are common, and a resume never puts more than one entry between
#: them, so one line of look-ahead is enough and two would start borrowing the next job's
#: city.
_ENTRY_LOOKAHEAD = 1

#: Entries that end within a year of the most recent one are treated as concurrent - a
#: student finishing a degree in Dec while working through Jun is in ONE place, and the more
#: specific of the two statements is the useful one ("Northridge, USA" over a bare "USA").
_ENTRY_CONCURRENT_MONTHS = 12


def _entry_end_key(token: str):
    """Sortable end-of-range key for one resume date token, or None if it isn't a date."""
    t = _normalize_edu_date_token(token)
    if t == "Present":
        return _ENTRY_PRESENT_KEY
    m = re.fullmatch(rf"(?i)({_EDU_MONTH_RE})\s+(\d{{4}})", t)
    if m:
        return int(m.group(2)) * 12 + _EDU_MONTH_ORDINAL[m.group(1)[:3].lower()]
    m = re.fullmatch(r"(\d{1,2})/(\d{4})", t)
    if m:
        return int(m.group(2)) * 12 + int(m.group(1))
    m = re.fullmatch(r"(\d{4})", t)
    if m:
        return int(m.group(1)) * 12 + 12      # a bare year ends when the year does
    return None


def _place_in_entry(line: str) -> str | None:
    """The place named in one dated resume entry, most specific form first."""
    scope = _past_institution_name(line) if _INSTITUTION_LINE_RE.search(line) else line
    scope = _EDU_DATE_RANGE_RE.sub(" | ", scope)
    scope = _EDU_DATE_RANGE_RE_LOOSE.sub(" | ", scope)
    scope = re.sub(r"\([^)]*\)", " | ", scope)      # '(3.72/4)', '(Remote)', '(GPA: 3.8)'
    city_st = re.compile(r"([A-Z][A-Za-z.\-]+(?:\s+[A-Z][A-Za-z.\-]+){0,1}),\s*([A-Z]{2})\b")
    found = {}
    for seg in re.split(r"\s*[|\u2022\u00b7]\s*", scope):
        seg = seg.strip(" ,;:-\u2013\u2014")
        if not seg or len(seg) > 90:
            continue
        m = city_st.search(seg)
        if m and m.group(2).lower() in US_STATE_ABBREVS:
            found.setdefault(1, f"{m.group(1)}, {m.group(2)}")
        for state in US_STATE_NAMES:
            m = re.search(r"([A-Z][A-Za-z.\-]+(?:\s+[A-Z][A-Za-z.\-]+){0,1}),\s*"
                          + re.escape(state.title()) + r"(?![a-z])", seg)
            if m:
                found.setdefault(2, f"{m.group(1)}, {state.title()}")
            # A bare state name ENDING the segment. Allowed here even for the name-like
            # states ('Georgia', 'Washington') that the header passes refuse: a dated
            # work-history entry states a workplace, and the candidate's own name is in the
            # contact block, not down here. Live: Manasa Puli's current role reads
            # 'AdyahTech  Apr 2025 - Present  Data Analyst  Georgia (Remote)'.
            if seg.lower().endswith(" " + state):
                found.setdefault(4, state.title())
        if _segment_names_a_country(seg) and _location_like(seg):
            found.setdefault(3, seg)
        elif seg.lower() in US_COUNTRY_TERMS or any(fc == seg.lower() for fc in FOREIGN_COUNTRIES):
            found.setdefault(5, seg)
    for rank in (1, 2, 3, 4, 5):
        if rank in found:
            return found[rank]
    return None


def _recency_location(lines: list) -> str | None:
    """The place attached to the candidate's most recent dated entry - see the module note.

    Entries ending within _ENTRY_CONCURRENT_MONTHS of the winner are treated as concurrent,
    and the most specific place among them wins: Mishit Shah's latest entry says only 'USA'
    while the degree he finished six months earlier says 'Northridge, USA'.
    """
    entries = []
    for i, ln in enumerate(lines):
        m = _EDU_DATE_RANGE_RE.search(ln) or _EDU_DATE_RANGE_RE_LOOSE.search(ln)
        if not m:
            continue
        key = _entry_end_key(m.group(2))
        if key is None:
            continue
        place = _place_in_entry(ln)
        for nxt in lines[i + 1:i + 1 + _ENTRY_LOOKAHEAD]:
            if place:
                break
            if _EDU_DATE_RANGE_RE.search(nxt) or _EDU_DATE_RANGE_RE_LOOSE.search(nxt):
                break              # that is the next entry, not this one's second line
            place = _place_in_entry(nxt)
        if place:
            entries.append((key, place))
    if not entries:
        return None
    newest = max(k for k, _ in entries)
    recent = [p for k, p in entries if newest - k <= _ENTRY_CONCURRENT_MONTHS]
    return max(recent, key=lambda p: (p.count(","), len(p)))


def _extract_location(text: str) -> str | None:
    """Resume location, most reliable first.

    1. Contact block "City, XX" (2-letter state abbrev) - see _contact_block
    2. Contact block "City, Full State Name"
    3. Full-text "City, XX" anywhere (work history addresses, etc.)
    4. LOCATION_KEYWORDS scan (header first, then full text)
    5. Full-text US state full name after a comma anywhere
    6. Foreign country/city scan (so the geo filter can reject non-USA)
    """
    lines = [ln.strip() for ln in (text or "").splitlines() if ln.strip()]
    # Normalise BEFORE slicing the contact block. A PDF that extracts one WORD per line
    # ('Rajat' / 'Gade' / '|' / '413-472-4804') leaves lines[:6] holding fragments, so
    # steps 1/1b/2 below - the whole "most reliable first" ladder - can never match and
    # every such resume falls through to the full-text scan, which returns whichever city
    # appears first: typically a former employer's or a prior degree's.
    #
    # Live audit 2026-09-05: 15 US-resident candidates were stored with a city they had
    # left years earlier - Rajat Gade 'Pune' (Accenture, ended 2023) though his resume
    # says 'Amherst, MA' on line 2; Tanuka Majumder 'Delhi' (IGDTUW, ended 2022) though
    # her header says 'College Park, Maryland, USA'; Hitesh Malhotra 'Bengaluru' (Conga,
    # ended 2024). The identical fragmentation already broke the education dates - see
    # _extract_education_dates, which normalises for exactly this reason.
    lines = [_normalize_spacing(ln)
             for ln in _rejoin_word_fragmented_lines(_split_collapsed_sections(lines))]
    header = _contact_block(lines)

    # Steps 4 and 6 below scan the WHOLE document (not just the header) for a bare
    # city/country name, so a city that only appears as part of a school's own name
    # ('Daulat Ram College, University of Delhi', 'University of Phoenix', 'San Jose
    # State University') was being returned as the candidate's current location - fixed
    # 2026-08-01, live case (Prerna Saluja / APP-20260716-2052-6112): her only 'Delhi'
    # mention is her undergrad alma mater; her current, most recent affiliation is
    # Arizona State University. Skip any match whose line is naming an institution.

    # At most TWO capitalised words before the comma. Three used to be allowed, but once
    # word-fragment rejoining runs a contact line reads 'Tanuka Majumder College Park,
    # Maryland' and the third slot swallowed the surname ('Majumder College Park'). Two
    # still covers the real multi-word cities that occur here - Palo Alto, College Park,
    # New York, San Francisco, Los Angeles.
    _city_state_re = re.compile(
        r"([A-Z][A-Za-z.\-]+(?:\s+[A-Z][A-Za-z.\-]+){0,1}),\s*([A-Z]{2})\b"
    )

    for ln in header:
        for seg in re.split(r"\s*[|]\s*", ln):
            # LAST match, not an anchored one: after word-fragment rejoining a contact
            # line can read 'Tanuka Majumder College Park, Maryland', and anchoring at
            # the start swallows the surname into the city ('Majumder College Park').
            _ms = list(_city_state_re.finditer(seg.strip()))
            m = _ms[0] if _ms else None
            # The 2-letter token must be a real US state, exactly as the full-text pass
            # below already requires - fixed 2026-08-03. Without the check this step also
            # matched foreign province codes and, because it returns only the "City, XX"
            # prefix, TRUNCATED the country the candidate had actually written:
            # "Toronto, ON, Canada" was stored as Location "Toronto, ON" with a BLANK
            # Country, losing the one explicit non-US signal on the row. Falling through
            # instead lets the foreign scan (step 6) return the whole segment, so
            # split_location_country can pull "Canada" back out of it.
            if m and m.group(2).lower() in US_STATE_ABBREVS:
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

    # 1c. Header "City FullStateName" with NO comma ('San Francisco California 94102',
    #     'Columbus Ohio'). 1b above only covers the 2-letter form, so these fell all the
    #     way through to the bare-city keyword scan - "San Francisco California" was stored
    #     as just "San Francisco" (state lost, Country blank), and "Columbus Ohio" produced
    #     no location at all because Columbus is not in the city list.
    #
    #     A handful of state names double as ordinary personal names, where "<Word>
    #     <StateName>" is a person, not a place - "John Washington", "Mary Virginia",
    #     "Sarah Georgia". Those require a trailing ZIP before they count; every other state
    #     name is unambiguous enough to accept on its own. City words must be capitalised and
    #     the segment must be short, to keep prose out.
    _zip_tail_re = re.compile(r"\s+(\d{5}(?:-\d{4})?)$")
    _city_words_re = re.compile(r"[A-Z][A-Za-z.\-]*(?:\s+[A-Z][A-Za-z.\-]*){0,2}$")
    for ln in header:
        for seg in re.split(r"\s*[|•·]\s*", ln):
            seg = seg.strip()
            if "," in seg:
                continue                        # comma forms are handled elsewhere
            has_zip = bool(_zip_tail_re.search(seg))
            body = _zip_tail_re.sub("", seg).strip()
            low = body.lower()
            for state in US_STATE_NAMES:
                if not low.endswith(" " + state):
                    continue
                if state in _NAME_LIKE_STATE_NAMES and not has_zip:
                    break                       # almost certainly a person's name
                city = body[: len(body) - len(state) - 1].strip()
                if city and _city_words_re.fullmatch(city):
                    return f"{city}, {state.title()}"
                break

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
            # An education line is the exception: a campus address is routinely followed
            # by the degree or its dates - 'Amherst, MA  Master of Science in Computer
            # Science' (live case Rajat Gade / APP-20260518-1725-9BB6, whose resume names
            # no other US city, so this guard cost him his location entirely).
            tail = ln[m.end():]
            if re.match(r"\s+[A-Z][a-z]", tail) and not _ADDRESS_TAIL_OK_RE.match(tail):
                continue
            return f"{m.group(1)}, {m.group(2)}"

    for scope in (header,):
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
                # Widen a bare city hit to the full segment when the candidate actually
                # wrote a country after it - fixed 2026-08-03. This scan returns only
                # `city.title()`, so "Bengaluru, Karnataka, India" was stored as Location
                # "Bengaluru" with a BLANK Country, silently discarding the explicit
                # country. The foreign scan further down would have kept the whole
                # segment, but this bare-city pass runs first and preempts it.
                _prose_only = True
                for seg in re.split(r"\s*[|]\s*", hit_line):
                    seg = seg.strip()
                    if city not in seg.lower():
                        continue
                    if _segment_names_a_country(seg) and _location_like(seg):
                        return seg
                    if not _mentions_place_as_prose(seg):
                        _prose_only = False
                # Every segment naming this city is a SENTENCE, not an address - the place
                # is the topic of a duty bullet, not where the candidate lives. Fixed
                # 2026-08-04, live case Hetvi Shah (APP-20260601-1700-3A6E): her CV reads
                # "Arizona State University, Tempe, Arizona" with a (602) Phoenix phone,
                # but one bullet says "...analyzing Singapore's age, gender, and ethnicity
                # trends...". 'singapore' is in LOCATION_KEYWORDS, so this pass returned it
                # as her Location; Ollama then split that into Location 'Remote' / Country
                # 'Singapore' and a US-based candidate landed in location review with a
                # foreign country on her row. Keep scanning instead - the real
                # "City, State" passes below still find Tempe, Arizona.
                if _prose_only:
                    continue
                return city.title()

    # Nothing above found an address the candidate stated as their OWN. Everything below
    # is history - a school or an employer - so the dates decide which entry is current.
    _recent = _recency_location(lines)
    if _recent:
        return _recent

    # An explicit US STATE outranks a bare city keyword found anywhere in the document.
    # The bare-city scan above therefore only reads the contact block; its full-document
    # half runs below, after this. Live case Aniruddha Rajnekar (APP-20260428-2224-14E7):
    # 'Barclays, Pune, India' - a job he left in 2022 - was preempting 'Raleigh, North
    # Carolina', the degree he finished in 2025 and the state he actually lives in.
    #
    # Search line by line, NOT over " ".join(lines): joining lets the city pattern run
    # backwards across a line break and swallow whatever preceded it, which turned
    # "Jane Doe\nPhoenix, Arizona 85004" into the location "Jane Doe Phoenix, Arizona".
    for state in US_STATE_NAMES:
        state_re = re.compile(
            r"([A-Z][A-Za-z.\-]+(?:\s+[A-Z][A-Za-z.\-]+){0,2}),\s*"
            + re.escape(state.title()) + r"(?![a-z])")
        for ln in lines:
            if _INSTITUTION_LINE_RE.search(ln):
                # Keep only the campus address, if the line has one - see
                # _past_institution_name. The foreign scan (step 6) deliberately keeps the
                # blunt skip-the-whole-line guard: relaxing it there would start reading a
                # foreign alma mater as a foreign residence, which is the exact misread
                # that put 15 US-resident students in location review on 2026-09-05.
                ln = _past_institution_name(ln)
                if not ln:
                    continue
            m2 = state_re.search(ln)
            if m2:
                return f"{m2.group(1)}, {state.title()}"
        for ln in lines:
            if _INSTITUTION_LINE_RE.search(ln):
                ln = _past_institution_name(ln)
                if not ln:
                    continue
            if re.search(r",\s*" + re.escape(state) + r"(?![a-z])", ln.lower()):
                return state.title()

    # The full-document half of the bare-city scan - see the note above the state pass.
    for scope in (lines,):
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
                _prose_only = True
                for seg in re.split(r"\s*[|]\s*", hit_line):
                    seg = seg.strip()
                    if city not in seg.lower():
                        continue
                    if _segment_names_a_country(seg) and _location_like(seg):
                        return seg
                    if not _mentions_place_as_prose(seg):
                        _prose_only = False
                if _prose_only:
                    continue
                return city.title()

    # 6. Foreign country/city (header first, then full text) — so the geo filter can
    #    reject non-USA candidates instead of defaulting to "benefit of the doubt".
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
                            if _location_like(seg):
                                return seg
                            if _mentions_place_as_prose(seg):
                                continue      # a topic, not an address - see below
                            return country.title()
        for city in FOREIGN_CITIES:
            if re.search(r"(?<![a-z])" + re.escape(city) + r"(?![a-z])", joined):
                for ln in scope:
                    if _INSTITUTION_LINE_RE.search(ln):
                        continue
                    for seg in re.split(r"\s*[|]\s*", ln):
                        if city in seg.lower():
                            seg = seg.strip()
                            if _location_like(seg):
                                return seg
                            if _mentions_place_as_prose(seg):
                                continue
                            return city.title()

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
# The apostrophe class is ['’] on purpose: Word and most PDF generators emit a CURLY
# right single quote, not a straight one. With only the straight form, "Master's" (curly)
# failed to match while "Masters" and "Master's" (straight) both worked - so a real degree
# line was invisible to the degree scan. Live case 2026-08-11: APP-20260507-2235-813F, whose
# CV reads "Master’s in computer science, University of North Texas".
_DEGREE_RE = re.compile(
    r"\b(Bachelor(?:['’]s|s)\b|Bachelor\s+(?:of|in)\s+\w+|Master(?:['’]s|s)\b|"
    r"Master\s+(?:of|in)\s+\w+|Associate(?:['’]s|s)\b|Associate\s+of\s+\w+|"
    r"Ph\.?D\.?|MBA|B\.?Tech\.?|M\.?Tech\.?|"
    r"B\.?S\.?[Cc]?\.?|M\.?S\.?[Cc]?\.?|B\.?A\.?|M\.?A\.?|B\.?E\.?|M\.?E\.?)\b"
)


# Section headings used to re-split a PDF that extracted as a few giant blobs.
_SECTION_SPLIT_RE = re.compile(
    r"(?=\b(?:Education|Work\s+Experience|Experience|Skills|Projects|Certifications|"
    r"Publications|Summary|Other\s+Experiences)\b)")

# Exact section headings that terminate a bounded resume section.  Keep this narrower than
# config.SECTION_WORDS: that set contains ordinary title words such as "engineer" and
# "manager", which can appear inside an education entry.  This predicate is used by the
# education date reader so a nearby employment date can never be attached to a degree.
_SECTION_HEADING_RE = re.compile(
    r"(?i)^(?:education|(?:(?:professional|work|employment)\s+)?(?:experience|history)|"
    r"(?:hard|technical|core)\s+skills?|skills?|projects?|certifications?|"
    r"publications?|achievements?|awards?|languages?|references?|summary|profile|"
    r"objective|interests?|hobbies)\s*:?$")


def _is_section_heading(line: str) -> bool:
    return bool(_SECTION_HEADING_RE.fullmatch(
        _collapse_letter_spacing(str(line or "")).strip()))


def _split_collapsed_sections(lines: list) -> list:
    """Re-split resume lines that arrived as one enormous tab-joined blob.

    Added 2026-08-11. Some PDFs (typically Word exports with a table layout) extract with no
    usable line breaks at all - one live CV came back as 7 lines, two of them ~3,100 chars,
    with every word tab-separated. Every section-header heuristic here is line-oriented
    ('does this line START with Education?'), so on that shape they all silently fail and the
    candidate's real degree is never found even though the text is sitting right there.

    Only blobs are touched: a line has to be both very long AND tab-joined before it is split,
    so normally-extracted resumes pass through untouched. Tabs also collapse to spaces, which
    is what makes 'computer\\tscience' read as 'computer science' again.
    """
    out = []
    for ln in lines:
        if len(ln) > 400 and "\t" in ln:
            flat = re.sub(r"[ \t]+", " ", ln.replace("\t", " "))
            out.extend(p.strip() for p in _SECTION_SPLIT_RE.split(flat) if p.strip())
        else:
            out.append(ln)
    return out


# A run of consecutive fragments this short or shorter is a candidate for rejoining -
# genuine resume lines (a degree title, a school+date line) run well past this.
_FRAGMENT_MAX_LEN = 22
_FRAGMENT_MAX_WORDS = 3
#: Below this many consecutive short fragments, leave them alone - a genuine short bullet
#: list (a few one-word skills, say) must never get glued into a single line.
_FRAGMENT_MIN_RUN = 6
#: Deliberately the SAME narrow, curated header set _SECTION_SPLIT_RE already uses - not
#: the broader config.SECTION_WORDS, which also contains generic words like 'software'
#: and 'engineering' for an unrelated purpose. A degree title reading '... Software
#: Engineering ...' must never be mistaken for hitting a real section boundary mid-run.
_FRAGMENT_HEADER_WORDS = {
    "education", "work experience", "experience", "skills", "projects",
    "certifications", "publications", "summary", "other experiences",
}


def _rejoin_word_fragmented_lines(lines: list) -> list:
    """Rejoin a resume PDF that extracted with roughly one WORD (or short phrase) per
    line, back into normal multi-word lines.

    Added 2026-09-01: a live audit of every candidate still missing Education Start/End
    Date found this single shape behind well over a third of the misses - a school/degree
    block that reads on the page as 'Seattle University, MS in Computer Science
    Sept 2023 - June 2025' extracts as ten-plus separate one-or-two-word lines ('Seattle' /
    'University' / ',' / 'MS' / 'in' / 'Computer' / 'Science' / 'Sept' / '2023' / '-' /
    'June' / '2025'). It is the same PDF-layout defect class _collapse_letter_spacing
    already fixes at the single-CHARACTER level ('E D U C A T I O N') and
    _split_collapsed_sections fixes at the whole-blob level - this is the missing WORD-
    level case in between. Every line-window regex in this module ('does this line START
    with Education', 'is there a date range in the next 3 lines') sees only isolated
    single-word lines on this shape and can never match, even though _extract_education
    itself still succeeds (Ollama reads fragmented text semantically and isn't fooled).

    Deliberately conservative: only a RUN of at least _FRAGMENT_MIN_RUN consecutive short
    fragments is collapsed (into ONE line, space-joined), so a normal resume's genuinely
    short lines (a bullet, a section header, a one-word skill) are never touched - and a
    run breaks at a recognized section-header word so 'EDUCATION' immediately followed by
    fragmented content never gets glued INTO the heading itself. A normally-extracted
    resume (mostly full-sentence lines) is returned completely unchanged."""
    def is_fragment(ln: str) -> bool:
        return (len(ln) <= _FRAGMENT_MAX_LEN and len(ln.split()) <= _FRAGMENT_MAX_WORDS
                and ln.lower().rstrip(":").strip() not in _FRAGMENT_HEADER_WORDS)

    out = []
    i = 0
    n = len(lines)
    while i < n:
        if not is_fragment(lines[i]):
            out.append(lines[i])
            i += 1
            continue
        j = i
        while j < n and is_fragment(lines[j]):
            j += 1
        run = lines[i:j]
        if len(run) >= _FRAGMENT_MIN_RUN:
            out.append(" ".join(run))
        else:
            out.extend(run)
        i = j
    return out


# A school/awarding body. Used to decide whether a line under an 'Education' header that
# carries NO degree keyword is still education (e.g. a bare 'Arizona State University') or
# just the next unrelated line the old fallback swallowed - see _extract_education.
_INSTITUTION_RE = re.compile(
    r"(?i)\b(universit(?:y|e|at|à|ä)|college|institute|institut|school|academy|polytechnic|"
    r"seminary|conservatory|gymnasium|iit|nit|iiit)\b"
)

# Some resumes state a field of study but omit a degree label, then put the institution on
# the next line.  Consult this only inside an explicit Education section.
_EDUCATION_SUBJECT_RE = re.compile(
    r"(?i)\b(?:engineering|computer\s+science|information\s+technology|data\s+science|"
    r"business\s+administration|finance|accounting|marketing|graphic\s+design)\b")


_EDUCATION_FUSED_RE = re.compile(
    r"(?i)(?:master|bachelor|associate|university|college|institute)"
    r"(?:of|in|at)(?:science|arts|technology|engineering|computer|business|information)"
)


def _looks_letter_spaced(value: str) -> bool:
    """True when a value is mostly single-character tokens ('B a c h e l o r  o f  S c i...').

    Some PDF templates letter-space every line, which _normalize_letterspaced_document now
    fixes up front for FRESH extractions. Values stored before that fix (2026-07-31) are still
    sitting in the workbook in the damaged form, and none of the other repair heuristics can
    see them: _EDUCATION_FUSED_RE needs whole words, and the '\\b[a-z]\\s+[a-z]{4,}\\b' probe
    needs a 4+ letter word, which letter-spaced text never contains. Without this they would
    never be picked up by a heal pass.
    """
    tokens = [t for t in str(value or "").split() if t]
    if len(tokens) < 4:
        return False
    singles = sum(1 for t in tokens if len(t) == 1)
    return singles / len(tokens) > 0.6


def education_needs_repair(value: str) -> bool:
    """True when an Education cell is visibly fused, truncated, OCR-damaged, or garbage."""
    text = str(value or "").strip()
    if not text or text.lower() in _GAP_LITERALS:
        return True
    low = text.lower()
    return bool(
        _looks_letter_spaced(text)
        or _EDUCATION_FUSED_RE.search(text)
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
    # Single underscores are pymupdf4llm italic markers, not candidate data.
    text = re.sub(r"(?<!\w)_([^_\n]+)_(?!\w)", r"\1", text)
    text = text.replace("~~", "")
    text = _EDU_GLUED_MONTH_RE.sub(" ", text)
    # Undo stylistic letter-spacing before anything else reads the value, so a cell stored
    # damaged (pre-2026-07-31) is actually repaired rather than just flagged. Collapsing runs
    # the words together ("BachelorofScienceinCS"), which the fused-word replacements below
    # then split back apart - that path is already well covered.
    if _looks_letter_spaced(text):
        text = _collapse_letter_spacing(text)
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
    text = re.sub(r",\s*[-\u2013\u2014]\s*", ", ", text)
    # Multi-degree entries can carry a parenthesized range before the next degree, so the
    # trailing-date cleanup below never reaches it. Date columns own this data; remove only
    # an unmistakable four-digit year range wherever it appears.
    text = re.sub(
        r"\s*\(\s*(?:19|20)\d{2}\s*[-\u2013\u2014]\s*(?:19|20)\d{2}\s*\)",
        "", text)
    # A trailing date is not part of the degree: Education Start/End Date carry it, and live
    # row APP-20260902-2155-MCPA published 'B.Sc. (Design and Computing), BITS Pilani (WILP)
    # | 2025' with the same 2025 already in Education End Date (2026-09-14).
    # Education Start/End Date own month/year evidence. Remove it anywhere in a combined
    # multi-degree value, including "MS ... | May 2026 Bachelor ... | May 2024".
    text = re.sub(
        rf"(?i)\s*(?:\|\s*)?(?:{_EDU_MONTH_RE})\.?\s+\d{{4}}"
        rf"(?:\s*(?:[-\u2013\u2014]|to)\s*(?:(?:{_EDU_MONTH_RE})\.?\s+\d{{4}}|\d{{4}}|present|current))?"
        r"\s*(?:\(\s*expected\s*\))?",
        " ", text)
    text = re.sub(r"\(\s*\)", "", text)
    text = re.sub(r"\s+([,;])", r"\1", text)
    text = re.sub(r"\s{2,}", " ", text).strip(" ,;-|")
    text = _EDU_TRAILING_DATE_RE.sub("", text).strip(" ,;-|")
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
    # A blob-extracted PDF has no line starting with 'Education', so every check below would
    # miss a degree that is plainly in the text. Re-split those before looking. No-op on a
    # normally-extracted resume.
    lines = _split_collapsed_sections(lines)
    lines = _rejoin_word_fragmented_lines(lines)

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
        subject = None
        for nxt in lines[i + 1:i + 6]:
            if nxt.lower().rstrip(":").strip() in SECTION_WORDS:
                break
            cleaned = normalize_education(nxt.rstrip(".,;"))
            if cleaned == "Not extracted":
                continue
            if _DEGREE_RE.search(cleaned):
                return cleaned
            # The fallback used to accept ANY line under the header, which is how these live
            # values got stored as a candidate's Education (2026-08-11):
            #     'LANGUAGE SKILLS'                              (a section header)
            #     'Languages: Dart, Java (Basic), Kotlin (Basic)' (a skills line)
            #     'Manufacturing'                                 (an industry word)
            # A line with neither a degree nor an institution is not education, and a blank
            # cell is honest where a wrong value is not - the missing-info nudge also keys off
            # Education being blank, so garbage here actively suppresses that follow-up.
            if _INSTITUTION_RE.search(cleaned):
                if subject:
                    school = _institution_on_line(cleaned) or cleaned
                    return normalize_education(f"{subject}, {school}")
                if fallback is None:
                    fallback = cleaned
            elif subject is None and _EDUCATION_SUBJECT_RE.search(cleaned):
                subject = cleaned
        if fallback:
            return fallback
        break

    for ln in lines:
        if _DEGREE_RE.search(ln):
            return normalize_education(ln.rstrip(".,;"))
    return None


# Date tokens seen in an education line: 'Aug 2021', 'August 2021', '08/2021', or a bare
# '2021'. Deliberately does NOT match day-level dates - resumes state education dates by
# month/year at the finest, never by day.
_EDU_MONTH_RE = (r"(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|"
                 r"Jul(?:y)?|Aug(?:ust)?|Sep(?:t|tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)")
# The MM/YYYY alternative is guarded on both sides against an adjacent '/digit' so a
# day-precision date (05/10/2018, i.e. DD/MM/YYYY or MM/DD/YYYY) can't have its middle
# segment misread as a bogus month ('10/2018' out of '05/10/2018'). Education dates are
# never stated to day precision on a resume, so a 3-segment date falls through to the bare
# \d{4} alternative instead, which still recovers the (correct) year on each side.
_EDU_MONTH_ORDINAL = {m: i for i, m in enumerate(
    ("jan", "feb", "mar", "apr", "may", "jun",
     "jul", "aug", "sep", "oct", "nov", "dec"), start=1)}
_EDU_DATE_TOKEN = (rf"(?:{_EDU_MONTH_RE}\.?\s+\d{{4}}|"
                   rf"(?<![\d/])\d{{1,2}}/\d{{4}}(?![\d/])|\b\d{{4}}\b)")
_EDU_DATE_RANGE_RE = re.compile(
    rf"(?i)\b({_EDU_DATE_TOKEN})\s*(?:-|–|—|to)\s*"
    rf"({_EDU_DATE_TOKEN}|Present|Current|Currently|Ongoing|Now)\b")
_EDU_EXPECTED_RE = re.compile(
    rf"(?i)\b(?:expected|anticipated)(?:\s+graduation)?\s*[:\-]?\s*({_EDU_DATE_TOKEN})\b")
#: 'Graduated: 2025' / 'Graduated 01/2026' - same single-date-only shape as Expected/
#: Anticipated above (a candidate stating just the one date, past tense since they've
#: already finished), added 2026-09-01 after a live audit found this exact phrasing
#: unrecognized on two candidates whose actual date ('Graduated: 2025', 'Graduated
#: 01/2026') was otherwise sitting right there in the text.
_EDU_GRADUATED_RE = re.compile(
    rf"(?i)\bgraduat(?:ed|ing)\s*[:\-]?\s*({_EDU_DATE_TOKEN})\b")
_EDU_MONTH_YEAR_RE = re.compile(
    rf"(?i)\b({_EDU_MONTH_RE}\.?\s+\d{{4}})\b")
#: Any bare date token (month-name+year, MM/YYYY, or a lone 4-digit year) used as the
#: LAST-RESORT single-date fallback - added 2026-09-01. _EDU_MONTH_YEAR_RE above only
#: recognizes the month-NAME shape; a live audit found several resumes stating a single
#: END date as a bare year next to each degree ('KL University, 2018', 'Oklahoma
#: Christian University - 2017') or MM/YYYY ('Tempe, United States  05/2023') with no
#: month name and no 'Expected'/'Graduated' keyword at all.
#:
#: NOT simply _EDU_DATE_TOKEN reused standalone: that token's bare-\d{4}-year branch is
#: only safe INSIDE a range, where a second valid token is required alongside it. Used
#: alone, it matched the trailing '2018' straight out of a day-precision '05/10/2018' -
#: exactly the misread the range tier's DD/MM/YYYY guard exists to prevent (caught by the
#: existing regression test for that guard). The extra (?<!/) here blocks a bare year that
#: is itself the tail segment of a slash-separated date; a genuine standalone year is never
#: preceded by a slash in ordinary resume prose.
_EDU_ANY_SINGLE_DATE_RE = re.compile(
    rf"(?i)\b({_EDU_MONTH_RE}\.?\s+\d{{4}}|"
    rf"(?<![\d/])\d{{1,2}}/\d{{4}}(?![\d/])|"
    rf"(?<!/)\b\d{{4}}\b)")
# PDF text layers sometimes glue an institution name directly to the following month
# ("ManagementAug. 2025"). Restore that one missing boundary before searching for a date
# range; otherwise the parser falls back to the less useful bare year.
#: What may legitimately follow a "City, ST" on an education or work-history line: the
#: degree, or the dates. Anything else that looks like a continued comma-list is the tools
#: list this guard was built for ("Bloomberg Terminal, MS Project, MS Office Suite").
_ADDRESS_TAIL_OK_RE = re.compile(
    rf"(?i)^\s+(?:{_EDU_MONTH_RE}\b|\d|"
    r"(?:Masters?|Bachelors?|Associates?|Doctorate?|Diploma|Certificate|Present|"
    r"B\.?[AS]|M\.?[AS]|M\.?B\.?A|Ph\.?\s?D)\b)")


_EDU_GLUED_MONTH_RE = re.compile(
    rf"(?<=[A-Za-z])(?={_EDU_MONTH_RE}\.?\s+\d{{4}})", re.IGNORECASE)
# A 2-digit-year date ('Sep'24', 'Aug 23') used ONLY as a fallback RANGE token (added
# 2026-09-01, live audit: 57/160 scored rows had NO education dates at all despite the
# resume plainly stating them - a common resume-template style prints a whole block of
# school/degree lines then all their date RANGES together at the end, using an
# abbreviated 2-digit year). Deliberately kept OUT of _EDU_DATE_TOKEN itself and every
# single-date fallback below: a bare 'Aug 23' is genuinely ambiguous with a day-of-month
# mid-sentence, but a full RANGE - two of these tokens joined by a dash - is a far
# stronger, low-risk signal, so this is tried only as _EDU_DATE_RANGE_RE's fallback.
_EDU_YEAR2_TOKEN = rf"{_EDU_MONTH_RE}\.?['’]?\s?\d{{2}}\b"
_EDU_DATE_RANGE_RE_LOOSE = re.compile(
    rf"(?i)\b({_EDU_YEAR2_TOKEN})\s*(?:-|–|—|to)\s*"
    rf"({_EDU_YEAR2_TOKEN}|Present|Current|Currently|Ongoing|Now)\b")


#: A date (or date range) at the very END of an Education value, after a separator:
#: '… | 2025', '…, May 2023', '… – 2019 - 2023', '… (Expected 2026)'. Only a separated,
#: final date is removed; a year inside a school's name is never touched.
_EDU_TRAILING_DATE_RE = re.compile(
    # A full range needs no separator ('… Computer Science Sept 2023 - June 2025'): tried
    # first, or the dash INSIDE the range would be read as the separator and only
    # '- June 2025' removed. A single date must follow a separator.
    rf"(?i)(?:\s+{_EDU_DATE_TOKEN}\s*(?:-|–|—|to)\s*(?:{_EDU_DATE_TOKEN}|present|current|ongoing)"
    rf"|\s*(?:[|,;]|\s[-–—]\s?|\()\s*"
    rf"(?:(?:expected|anticipated|graduated|class\s+of)\s*[:\-]?\s*)?"
    rf"{_EDU_DATE_TOKEN}"
    rf"(?:\s*(?:-|–|—|to)\s*(?:{_EDU_DATE_TOKEN}|present|current|ongoing))?)"
    rf"\s*\)?\s*$")

#: Separators between a school's name and what follows it on its own line (city, country,
#: GPA, dates): commas, pipes, bullets, and a SPACED dash. An unspaced hyphen stays, so
#: acronyms like 'FAST-NUCES' survive intact.
_INSTITUTION_SPLIT_RE = re.compile(r"\s*(?:[,|•·;]|\s[-–—]\s)\s*")


def _institution_on_line(line: str) -> str:
    """The school named on one resume line ('National University of … (FAST-NUCES)'), or ""."""
    # Drop a leading bullet only ('o ', '• ', '- '): str.strip('o') would eat the O of 'Ohio'.
    line = re.sub(r"^\s*(?:[•·▪◦●*\-–—]|o(?=\s))\s*", "", str(line or ""))
    segments = _INSTITUTION_SPLIT_RE.split(line.strip())
    for index, segment in enumerate(segments):
        # A GPA ends the name: 'Arizona State University (3.72/4) Tempe, AZ' -> the school
        # only, not the campus city that follows the grade with no separator.
        gpa = re.search(r"\(\s*\d[\d.]*\s*/\s*\d[\d.]*\s*\)|\bGPA\b|\bCGPA\b", segment)
        segment = (segment[:gpa.start()] if gpa else segment).strip(" ,;")
        if _INSTITUTION_RE.search(segment) and not _DEGREE_RE.match(segment):
            # A business/engineering school and its parent university are often separated
            # by a comma. Keep the adjacent parent institution, but stop before a campus
            # city/country (live: W. P. Carey School of Business, Arizona State University).
            school_parts = [segment]
            for following in segments[index + 1:]:
                following = following.strip(" ,;")
                if not (_INSTITUTION_RE.search(following)
                        and not _DEGREE_RE.match(following)):
                    break
                school_parts.append(following)
            return ", ".join(school_parts)
    return ""


def complete_education(value, resume_text: str) -> str:
    """Add the school to an Education value that names only the degree.

    Live row APP-20260908-1304-MCLA published 'Bachelor of Science in Computer Science'
    while its resume reads that line followed directly by 'National University of Computer &
    Emerging Sciences (FAST-NUCES), Lahore, Pakistan' (2026-09-14). The extractors are asked
    for the school and a sibling row got it; this one did not, so the column was
    inconsistent. The school is taken only from the degree's own line or the two adjacent
    lines in the same section, and only its name - never the city or country after it,
    which say where the candidate studied, not where they live.
    """
    val = normalize_education(value)
    if val == "Not extracted" or _INSTITUTION_RE.search(val) or not resume_text:
        return val
    core = re.sub(r"[^a-z0-9]+", " ", val.lower()).strip()
    if len(core) < 4:
        return val
    lines = [ln.strip() for ln in str(resume_text).splitlines() if ln.strip()]
    for i, line in enumerate(lines):
        if core not in re.sub(r"[^a-z0-9]+", " ", line.lower()):
            continue
        forward, backward = [], []
        for candidate in lines[i + 1:i + 3]:
            if _is_section_heading(candidate) or _DEGREE_RE.search(candidate):
                break
            forward.append(candidate)
        for candidate in reversed(lines[max(0, i - 2):i]):
            if _is_section_heading(candidate) or _DEGREE_RE.search(candidate):
                break
            backward.append(candidate)
        # Immediate adjacency is stronger than direction: degree->school and
        # school->degree are both common. Do not skip over a coursework line to steal the
        # following degree's school (live: Nimish Goel, whose MS was paired with R.V. College
        # instead of the W. P. Carey / ASU line immediately above it).
        neighbors = [line]
        for offset in range(max(len(forward), len(backward))):
            if offset < len(forward):
                neighbors.append(forward[offset])
            if offset < len(backward):
                neighbors.append(backward[offset])
        for candidate in neighbors:
            school = _institution_on_line(candidate)
            if school and school.lower() not in val.lower():
                return normalize_education(f"{val}, {school}")
        break
    return val


def _normalize_edu_date_token(token: str) -> str:
    """'Currently'/'Current'/'Ongoing'/'Now' all read as 'Present' - that's the one signal
    the client actually wants (still enrolled vs. a finished date). Anything else is
    returned with whitespace collapsed, never reformatted, so the value stays traceable
    to the literal resume text."""
    t = re.sub(r"\s+", " ", str(token or "").strip())
    if t.lower() in {"present", "current", "currently", "ongoing", "now"}:
        return "Present"
    month_year = re.fullmatch(rf"(?i)({_EDU_MONTH_RE})\.?\s+(\d{{4}})", t)
    if month_year:
        # Keep the workbook's established compact month/year shape ("Aug 2025"),
        # including when the resume spells the month out or adds a trailing period.
        return f"{month_year.group(1)[:3].title()} {month_year.group(2)}"
    # 2-digit year ('Sep'24', 'Aug 23') -> full 'Aug 2023'. Every candidate this system
    # processes is a current/recent student, so the 2000s century is unambiguous.
    month_year2 = re.fullmatch(rf"(?i)({_EDU_MONTH_RE})\.?['’]?\s?(\d{{2}})", t)
    if month_year2:
        return f"{month_year2.group(1)[:3].title()} 20{month_year2.group(2)}"
    # 'MM/YYYY' ('05/2023') is deliberately left exactly as written, same as the existing
    # range-tier behavior ("08/2021 - 05/2025" stores literally, never reformatted to
    # "Aug 2021 - May 2025") - stays traceable to the literal resume text.
    return t


def _extract_education_dates(text: str) -> tuple[str, str]:
    """Best-effort (start, end) date for the SAME education entry _extract_education just
    read - scanned from the identical 'Education' header + next-5-lines window, so the
    dates returned always correspond to the degree/school already in the Education column,
    never to an unrelated date elsewhere in the resume (e.g. work history).

    Returns ("", "") when no explicit date range is present. Never fabricated, never
    inferred - a candidate resume that states no dates at all yields a blank pair, which is
    the honest answer, not a guess.

    An 'Expected <date>' / 'Anticipated <date>' phrase (no start date given) yields
    ("", "<date>") - a currently-enrolled candidate stating only a target graduation date.

    Each window is tried through four fallback tiers, strongest signal first: a strict
    (4-digit-year) date range, a loose (2-digit-year, e.g. "Sep'24 - May'26") date range,
    an "Expected <date>" phrase, then a single bare month/year anywhere in the window (the
    stated graduation date, start left blank rather than guessed). Added 2026-09-01 after
    a live audit found 57/160 scored candidates had NO education dates despite the resume
    plainly stating them - two common resume-template shapes were falling through every
    tier: (a) multiple degrees printed as stacked school/degree lines with ALL their date
    ranges grouped together afterward, using abbreviated 2-digit years; (b) a single date
    printed on the SCHOOL-name line immediately above the degree line, which the old
    degree-line-only single-date check couldn't see.
    """
    def _latest(pairs: list) -> tuple[str, str] | None:
        """The pair whose END date is furthest in the future, or None if none are dates."""
        dated = [(p, _entry_end_key(p[1])) for p in pairs]
        dated = [(p, k) for p, k in dated if k is not None]
        if not dated:
            return pairs[0] if pairs else None
        return max(dated, key=lambda pk: pk[1])[0]

    def _best_effort(window: str) -> tuple[str, str] | None:
        """The dates of the MOST RECENT education entry in the window.

        Tiers still run strongest-format first, because the weakest ones (a lone 4-digit
        year) match incidental numbers as readily as real dates. What changed on 2026-09-08
        is that a tier no longer returns its FIRST match: it returns its most recent one, and
        an explicit 'Expected <date>' is weighed against full ranges rather than sitting in a
        tier below them.

        Both halves were needed for live rows. 'University of Maryland ... Expected May 2026
        INDIRA GANDHI ... August 2018 - July 2022' published the Indian bachelor's range,
        because a strict range outranked 'Expected' no matter whose degree it belonged to.
        And "Bachelor's from JNTUK, India - 2017, Master's From Campbellsville University,
        USA - 2024" published 2017, because the single-date tier stopped at the first year it
        saw. This is the same rule Location already follows (see _recency_location): among
        dated entries, the current one describes the candidate.
        """
        # pymupdf4llm represents italic text as ``_August 2023 - May 2025_``. Underscore
        # is a regex word character, so the boundary before the month disappears and the
        # parser falls through to a less precise bare year.
        window = window.replace("_", " ")
        explicit = [(_normalize_edu_date_token(m.group(1)),
                     _normalize_edu_date_token(m.group(2)))
                    for m in _EDU_DATE_RANGE_RE.finditer(window)]
        # A stated graduation target is the same strength of claim as a printed range - it is
        # the degree the candidate is actually enrolled in - so it competes here, not later.
        explicit += [("", _normalize_edu_date_token(m.group(1)))
                     for m in _EDU_EXPECTED_RE.finditer(window)]
        explicit += [("", _normalize_edu_date_token(m.group(1)))
                     for m in _EDU_GRADUATED_RE.finditer(window)]
        if explicit:
            return _latest(explicit)
        loose = [(_normalize_edu_date_token(m.group(1)),
                  _normalize_edu_date_token(m.group(2)))
                 for m in _EDU_DATE_RANGE_RE_LOOSE.finditer(window)]
        if loose:
            return _latest(loose)
        # A single month/year anywhere in the window is conventionally the stated
        # graduation/end date. Do not infer a start date the candidate did not give.
        single = [("", _normalize_edu_date_token(m.group(1)))
                  for m in _EDU_MONTH_YEAR_RE.finditer(window)]
        if single:
            return _latest(single)
        # Last resort: a bare MM/YYYY or lone 4-digit year with no month name and no
        # Expected/Graduated keyword at all (still guarded against day-precision
        # DD/MM/YYYY by _EDU_DATE_TOKEN itself).
        any_dates = [("", _normalize_edu_date_token(m.group(1)))
                     for m in _EDU_ANY_SINGLE_DATE_RE.finditer(window)]
        if any_dates:
            return _latest(any_dates)
        return None

    lines = [ln.strip() for ln in (text or "").splitlines() if ln.strip()]
    lines = _split_collapsed_sections(lines)
    lines = _rejoin_word_fragmented_lines(lines)

    # A degree and date sharing one visual row is the strongest available association.
    # Consider all such rows before a damaged Markdown table can donate a nearby fellowship
    # or employment range (live: Pragya Mittal's 2023-2024 fellowship displaced the
    # Aug 2023-May 2025 dates printed on her master's row).
    degree_pairs = []
    for line in lines:
        if not _DEGREE_RE.search(line):
            continue
        found = _best_effort(_EDU_GLUED_MONTH_RE.sub(" ", line))
        if found:
            degree_pairs.append(found)
    same_row = _latest(degree_pairs)
    if same_row:
        return same_row

    # Anchor dates to the exact education value selected for the row whenever that entry can
    # be located.  Selecting the newest date in the whole section is wrong when the text
    # extractor chose a different entry.  APP-20260819-0140-MCRA selected the candidate's
    # B.Sc. line but attached a newer Japanese-diploma range from two lines above it.
    selected_education = _extract_education(text)
    if selected_education and selected_education != "Not extracted":
        selected_norm = normalize_education(selected_education)
        for i, line in enumerate(lines):
            line_norm = normalize_education(line.rstrip(".,;"))
            # A PDF layout supplement may repeat the same visible degree row with a date
            # restored from its right-hand column ("Bachelor ...                 2014").
            # Keep looking after a date-free occurrence and accept that longer row too.
            if (line_norm != selected_norm
                    and not line_norm.startswith(selected_norm + " ")):
                continue
            # Most layouts keep the range on the selected line.  This strongest association
            # also handles a collapsed two-column education row such as Victor's.
            found = _best_effort(_EDU_GLUED_MONTH_RE.sub(" ", line))
            if found:
                return found
            # A school/date line can precede the degree line in multi-column templates.
            if i:
                previous = lines[i - 1]
                if not _is_section_heading(previous):
                    found = _best_effort(_EDU_GLUED_MONTH_RE.sub(
                        " ", previous + "\n" + line))
                    if found:
                        return found
            # Or the school and range can follow it.  Stop at the next degree or section so
            # a sibling education entry cannot donate its dates.
            entry = [line]
            for candidate in lines[i + 1:i + 3]:
                if (_is_section_heading(candidate)
                        or _DEGREE_RE.search(candidate)):
                    break
                entry.append(candidate)
                found = _best_effort(_EDU_GLUED_MONTH_RE.sub(" ", "\n".join(entry)))
                if found:
                    return found
            # Do not stop at a date-free copy.  A coordinate-preserving supplement can
            # contain a second copy with the visually aligned date restored.

    header_re = re.compile(r"^education\b[:\-]?\s*(.*)$", re.IGNORECASE)
    for i, ln in enumerate(lines):
        m = header_re.match(_collapse_letter_spacing(ln))
        if not m:
            continue
        section = [m.group(1)]
        for candidate in lines[i + 1:i + 6]:
            if _is_section_heading(candidate):
                break
            section.append(candidate)
        window = "\n".join(section)
        window = _EDU_GLUED_MONTH_RE.sub(" ", window)
        found = _best_effort(window)
        if found:
            return found
        break

    # Multi-column PDFs can extract the EDUCATION heading after the degree entries. In
    # that layout the header window above is empty even though the degree line and its
    # dates are intact earlier in the text. Reuse the same first-degree rule as
    # _extract_education. The window includes ONE line before the matched degree line -
    # the common "School Name .......... Date" / "Degree Name .......... Location" resume
    # layout states the date on the school line, not the degree line - plus two lines
    # after, so a work history range elsewhere can never be mistaken for education dates.
    for i, line in enumerate(lines):
        if not _DEGREE_RE.search(line):
            continue
        degree_window = "\n".join(lines[max(0, i - 1):i + 3])
        degree_window = _EDU_GLUED_MONTH_RE.sub(" ", degree_window)
        found = _best_effort(degree_window)
        if found:
            return found
        break
    return "", ""


# -- Experience (total years) ------------------------------------------------
# Added 2026-09-04 on client instruction: one 'Experience' column reading "6 years", so a
# recruiter scanning the sheet can judge seniority without opening the CV. The SUMMARY is
# the stated source and is tried first for a reason - "Senior engineer with 6+ years of
# experience in ..." is the candidate's OWN total, and it already accounts for gaps,
# overlaps and part-time work in a way that adding up job date ranges cannot.
#
# Blank whenever the resume states no total - which reads as 'Missing' on the sheet like
# every other candidate-owed column. Deliberately NOT derived: not from graduation year,
# not from the number of listed jobs, and (since 2026-09-04, user instruction) not from
# the span of the dated work history either. Only a figure the candidate wrote down is
# ever published. Same "never fabricated" rule the education dates follow.

#: Section headers introducing the blurb at the top of a resume ('Professional Summary',
#: 'Career Objective', 'About Me', ...). Group 1 is any text on the header line itself,
#: for the common one-line "Summary: 6 years of ..." layout.
_SUMMARY_HEADER_RE = re.compile(
    r"^(?:professional\s+|career\s+|personal\s+|executive\s+|technical\s+)?"
    r"(?:summary|profile|objective|about(?:\s+me)?|overview|synopsis|snapshot)\b"
    r"[:\-–]?\s*(.*)$",
    re.IGNORECASE)

#: Spelled-out counts. Resumes write "over five years of experience" as often as "5+".
_EXP_WORD_NUMBERS = {
    "one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6, "seven": 7,
    "eight": 8, "nine": 9, "ten": 10, "eleven": 11, "twelve": 12, "thirteen": 13,
    "fourteen": 14, "fifteen": 15, "sixteen": 16, "seventeen": 17, "eighteen": 18,
    "nineteen": 19, "twenty": 20, "twenty-five": 25, "thirty": 30,
}
_EXP_NUM = (r"(?:(?P<w>twenty-five|twenty|thirty|nineteen|eighteen|seventeen|sixteen|"
            r"fifteen|fourteen|thirteen|twelve|eleven|ten|nine|eight|seven|six|five|"
            r"four|three|two|one)|(?P<n>\d{1,2}(?:\.\d+)?))")

#: "6+ years of experience", "over five years' experience", "6 to 8 years of hands-on
#: experience", "10 years of professional software development experience".
#: The bounded {0,4} word gap is what stops "5 years at Arizona State University" or
#: "3 years ago" from being read as a professional total - an experience/expertise/
#: background word has to follow closely.
_EXP_YEARS_FIRST_RE = re.compile(
    r"(?i)\b(?:over|above|more\s+than|nearly|almost|about|around|approx\.?|"
    r"approximately|close\s+to)?\s*" + _EXP_NUM + r"\s*(?:\+|plus)?\s*"
    r"(?:(?:-|–|—|to)\s*\d{1,2}\s*(?:\+|plus)?\s*)?"
    r"years?[’']?s?\s*(?:\+|plus)?\s*(?:of\s+|in\s+|with\s+|across\s+|within\s+|as\s+)?"
    # Filler tolerates a trailing comma, and runs a little longer. Resumes list
    # activities - "9+ years designing, developing, testing and supporting ..." - and a
    # comma stopped the filler dead, so the terminal word was never reached and a total
    # the candidate HAD stated was reported as Missing (live case, 2026-09-11).
    r"(?:[A-Za-z][A-Za-z/&+.\-]*,?\s+){0,6}?"
    # Gerunds sit alongside the nouns for the same reason: "9+ years designing enterprise
    # automation" states a total as plainly as "9+ years of experience". _exp_scan still
    # takes the LARGEST figure, so a per-skill breakdown cannot outrank an overall one.
    r"(?:experience|expertise|background|engineering|development|industry|software|owning|"
    r"production|systems|designing|developing|building|delivering|leading|managing|"
    r"supporting|architecting|automating)\b")

#: The mirrored phrasing: "Experience: 6 years", "total experience of 8+ years".
_EXP_YEARS_LAST_RE = re.compile(
    r"(?i)\b(?:experience|expertise)\b[^.\n]{0,24}?" + _EXP_NUM +
    r"\s*(?:\+|plus)?\s*years?\b")

def _exp_years_from_match(m) -> int | None:
    """Whole years out of an _EXP_YEARS_* match - the lower bound of any range."""
    if m.group("n"):
        try:
            value = int(float(m.group("n")))
        except ValueError:
            return None
    else:
        value = _EXP_WORD_NUMBERS.get((m.group("w") or "").lower())
    # 0 is not a claim of experience, and >60 is a parse accident (a stray year, say).
    if not value or not 1 <= value <= 60:
        return None
    return value


def _exp_scan(window: str) -> int | None:
    """Highest stated total in `window`.

    Resumes often print a per-skill breakdown ("3 years Python, 6 years Java") beneath one
    overall figure, and the overall figure is the answer - so the largest stated number
    wins rather than the first one seen.
    """
    best = None
    for regex in (_EXP_YEARS_FIRST_RE, _EXP_YEARS_LAST_RE):
        for m in regex.finditer(window or ""):
            years = _exp_years_from_match(m)
            if years is not None and (best is None or years > best):
                best = years
    return best


def _extract_experience(text: str) -> str:
    """Total professional experience as "6 years" / "1 year", or "" when unknown.

    Only ever reports a total the candidate STATED. Tiered strongest-signal-first,
    mirroring _extract_education_dates:
      1. an explicit total inside the Summary/Profile block (the client's stated source);
      2. an explicit total in the resume's opening lines - plenty of resumes state one in
         the headline under the name with no 'Summary' header at all;
      3. an explicit total anywhere else in the document.

    A resume that states no total yields "" (which the sheet renders as 'Missing'), even
    when its jobs are fully dated. An earlier revision spanned those dates into a number;
    that was REMOVED on user instruction 2026-09-04 ("just add years if its existing in
    summary") - a span is arithmetic the candidate never wrote, and a stated total already
    accounts for gaps, overlaps and part-time work in a way a span cannot. Same rule the
    education dates follow: blank beats derived.

    Normalized to whole years ("6+ years" and "6 to 8 years" both store "6 years"). The
    lower bound is used because it is the part the candidate actually committed to.
    """
    lines = [ln.strip() for ln in (text or "").splitlines() if ln.strip()]
    lines = _split_collapsed_sections(lines)
    lines = _rejoin_word_fragmented_lines(lines)

    years = None
    # 1. the summary block: the header line's own trailing text + the next 8 lines.
    for i, line in enumerate(lines):
        m = _SUMMARY_HEADER_RE.match(_collapse_letter_spacing(line))
        if not m:
            continue
        years = _exp_scan("\n".join([m.group(1)] + lines[i + 1:i + 9]))
        if years is not None:
            break
    # 2. the headline block above any section header.
    if years is None:
        years = _exp_scan("\n".join(lines[:12]))
    # 3. anywhere in the document.
    if years is None:
        years = _exp_scan("\n".join(lines))

    if years is None:
        return ""
    return "1 year" if years == 1 else "%d years" % years


def _strip_skill_section_label(piece: str) -> str:
    """Drop a resume SECTION HEADER that got glued to the first skill after it.

    Added 2026-08-04. Resumes group skills under headings, and when the whole block is
    flattened to one comma list the heading rides along on the first entry:

        'Programming Languages and Databases: Python'  -> 'Python'
        'Web Technologies: HTML'                       -> 'HTML'
        'Programming:Python'                           -> 'Python'
        'Tools & Software:Advanced Excel'              -> 'Advanced Excel'

    Live rows APP-20260716-0617-CB73 and APP-20260716-0312-A455 both stored these, which
    also broke scoring: _skill_set splits on commas, so the matcher was comparing JD skills
    against 'programming languages and databases: python' and matching nothing.

    Only strips when the text BEFORE the colon looks like a heading (letters, spaces, '&'
    or '/'), so a genuine skill containing a colon (a version or a ratio) survives.
    """
    s = str(piece or "").strip()
    m = re.match(r"^([A-Za-z][A-Za-z &/]{2,40}?)\s*:\s*(\S.*)$", s)
    if m and not re.search(r"\d", m.group(1)):
        return m.group(2).strip()
    return s


def _is_skill_prose(piece: str) -> bool:
    """True when an entry is a sentence lifted from a resume bullet, not a skill.

    Live row APP-20260720-1755-30E6 stored entries like 'Strategic planning for risk and
    cybersecurity at all levels of an organization' and 'Adaptable problem solving and the
    ability to create effective change'. They are accurate prose and useless as skills:
    nothing matches them, they crowd out real terms, and they make the column unreadable.

    Thresholds are deliberately loose so real multi-word skills stay ('Financial Planning &
    Analysis', 'AI Management System (AIMS ISO 42001:2023)', 'Natural Language Processing').
    """
    s = str(piece or "").strip()
    if not s:
        return True
    return len(s) > 70 or len(s.split()) > 8


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
        # Keep a bracketed group ONE skill: 'AWS (EC2, S3, IAM)' was split on its inner commas
        # into 'AWS (EC2' / 'S3' / 'IAM)', which only read correctly because they were joined
        # back with the same comma. Balanced, short groups only, so a stray '(' never swallows
        # the rest of the list.
        text = re.sub(r"\(([^()]{1,60})\)",
                      lambda m: "(" + re.sub(r"\s*[,;]\s*", "/", m.group(1).strip()) + ")", text)
        # Split on ';' as well as ',': resumes separate skill GROUPS with semicolons
        # ('...MongoDB, SQL; Web Technologies: HTML, CSS'), and without this the group
        # heading stays glued mid-piece where _strip_skill_section_label cannot see it.
        #
        # A FULL STOP followed by a capital is a group boundary too - live row
        # APP-20260716-0312-A455 reads '...MATLAB, LaTeX. Quantitative Toolkit:H ∞ Control
        # Theory', where 'LaTeX' is a real skill and 'Quantitative Toolkit:' is a heading, so
        # truncating would lose the skill and not splitting would keep the heading. Requiring
        # whitespace after the period leaves 'Node.js' and '.NET' untouched, and requiring a
        # capital next avoids splitting mid-sentence prose.
        pieces = [p.strip(" \t\r\n'\"")
                  for p in re.split(r"\s*[,;]\s*|(?<=[a-zA-Z])\.\s+(?=[A-Z])", text)]
    seen, out, kept_before_prose_filter = set(), [], []
    for piece in pieces:
        piece = re.sub(r"\s{2,}", " ", piece).strip(" ,;")
        if not piece or piece.lower() in _GAP_LITERALS:
            continue
        piece = _strip_skill_section_label(piece)
        if not piece:
            continue
        key = piece.lower()
        if key in seen:
            continue
        seen.add(key)
        kept_before_prose_filter.append(piece)
        if _is_skill_prose(piece):
            continue
        out.append(piece)
    # NEVER let the prose filter empty the column. A resume whose skills are ALL written as
    # sentences (APP-20260720-1755-30E6: 'Strategic planning for risk and cybersecurity at
    # all levels of an organization', ...) would otherwise come back as 'Not extracted' -
    # and an empty skill set makes every JD score zero overlap, which on 2026-08-04 caused
    # the scorer to discard all 11 ranked roles for that candidate. Unhelpful prose beats no
    # data at all, so fall back to the unfiltered list when filtering removed everything.
    if not out and kept_before_prose_filter:
        return ", ".join(kept_before_prose_filter)
    # 'AWS' next to 'AWS (EC2/S3/IAM)' is the same skill listed twice (live row
    # APP-20260908-1304-MCLA, 2026-09-14). Keep the more specific form; scoring._skill_set
    # still reads the base 'aws' out of it, so no role match is lost.
    detailed = {re.sub(r"\s*\(.*\)\s*$", "", p).strip().lower() for p in out if re.search(r"\(.+\)\s*$", p)}
    out = [p for p in out if "(" in p or p.lower() not in detailed]
    return ", ".join(out) if out else "Not extracted"


def clean_location_text(raw: str) -> str:
    """Strip the resume's own field LABEL and layout artifacts off a location string.

    Added 2026-08-04. Live row APP-20260716-0335-0C84 stored the location as
    'Address  -  San  Jose, California': the resume's 'Address -' label came along, and the
    PDF's column layout left doubled spaces inside the city name. Cosmetically ugly, but it
    also breaks matching - 'San  Jose' with two spaces is not the 'San Jose' in the US-city
    list, so a real US location can miss and drop the row into review.

        'Address  -  San  Jose, California'  ->  'San Jose, California'
        'Location: Austin, TX'               ->  'Austin, TX'
        'Based in Chicago, IL'               ->  'Chicago, IL'
    """
    s = str(raw or "").strip()
    if not s:
        return s
    # Leading field label, with or without a separator: 'Address -', 'Location:', 'City |'.
    s = re.sub(r"^\s*(?:current\s+)?(?:address|location|city|residence|based\s*(?:in|at))"
               r"\s*[-–—:|]*\s*", "", s, flags=re.I)
    # Collapse the doubled spaces PDF column extraction leaves behind.
    s = re.sub(r"\s{2,}", " ", s)
    # Tidy stray separators left at either end.
    s = s.strip(" -–—:|,")
    s = s.strip()

    # Collapse a segment that simply repeats one already present, e.g. the extractor
    # reading a header twice and storing 'Bikaner, Bikaner' (live row APP-20260602-1601-EF3F,
    # 2026-08-11). Cosmetically poor in the client-facing Location column, and a repeated
    # token is never additional information.
    #
    # GUARDED against the city-named-after-its-state case: 'New York, New York' and
    # 'Oklahoma City, Oklahoma'-style values are legitimate and must keep the state, so a
    # repeat is only dropped when it is NOT a US state/territory name. Comparison is
    # case-insensitive; the first occurrence keeps its original casing and position.
    segments = [seg.strip() for seg in s.split(",") if seg.strip()]
    if len(segments) > 1:
        kept: list[str] = []
        seen: set[str] = set()
        for seg in segments:
            low = seg.lower()
            if low in seen and low not in US_STATE_NAMES and low not in US_TERRITORY_NAMES:
                continue
            seen.add(low)
            kept.append(seg)
        s = ", ".join(kept)
    return s.strip()


#: Matches an institution name leaking into Location ('George Mason University',
#: 'Daulat Ram College'). _extract_location's own regex tier already guards against this
#: for ITS matches (see its 2026-08-01 comment), but the AI/Ollama tier has no equivalent
#: check - and the existing anti-hallucination gate (_location_grounded_in_text) only
#: verifies a claim genuinely APPEARS in the resume, not that it's plausibly a place. An
#: institution name is almost always genuinely present (it's the candidate's own school),
#: so it passes grounding cleanly while still being the wrong field entirely.
_LOCATION_INSTITUTION_RE = re.compile(
    r"\b(?:university|college|institute of technology)\b", re.I)
#: A degree-sentence fragment ('Bachelor's from JNTUK, India - 2017.') leaking in from the
#: Education section rather than a real address line.
_LOCATION_DEGREE_RE = re.compile(r"(?i)^(?:bachelor|master|associate|ph\.?d)'?s?\b")


def location_is_plausible(loc: str) -> bool:
    """False for a Location value that is clearly an institution name, a degree-sentence
    fragment, or a resume section-header artifact rather than a real place - added
    2026-09-01 after live rows stored 'George Mason University', 'ABOUT, ME', and
    'Bachelor's from JNTUK, India - 2017.' as Location, each of which is genuinely present
    in the resume text (so the grounding check alone lets it through) but is not a place.

    Deliberately narrow and keyword-based, matching this codebase's existing style of
    precedented, single-purpose guards rather than a general-purpose classifier - it only
    rejects unambiguous non-location shapes. A blank/gap value is not this function's
    concern (the existing _GAP_LITERALS machinery already handles that) - it returns True
    (plausible / not this function's business) for '' so callers can gate on it safely."""
    s = str(loc or "").strip()
    if not s or s.lower() in _GAP_LITERALS:
        return True
    if _LOCATION_INSTITUTION_RE.search(s):
        # A real "City, ST"/"City, State" address is never rejected just because the
        # CITY itself contains 'College'/'University' - 'College Park, MD', 'State
        # College, PA' and 'University City, MO' are all real places (caught live
        # 2026-09-01 as false positives before this guard existed). Only a BARE
        # institution name with no trailing state/territory qualifier is leakage.
        parts = [p.strip() for p in s.split(",")]
        last = parts[-1].lower() if len(parts) > 1 else ""
        if last not in US_STATE_ABBREVS and last not in US_STATE_NAMES \
                and last not in US_TERRITORY_NAMES:
            return False
    if _LOCATION_DEGREE_RE.match(s):
        return False
    if s.strip(" ,.").upper() in ("ABOUT ME", "ABOUT, ME", "ABOUT"):
        return False
    return True


def split_location_country(raw: str) -> tuple[str, str]:
    """Split a combined location into (city_state, country)."""
    if not raw or raw.strip().lower() in ("", "not extracted"):
        return ("Not extracted", "")

    raw = clean_location_text(raw) or raw
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

    # TRAILING country with no comma in front of it (added 2026-08-11). Everything above only
    # ever inspects the last COMMA-separated segment, so these all left Country blank on live
    # rows even though the country was sitting right there in the text:
    #     'Rawalpindi Punjab Pakistan'        -> no comma at all
    #     '133 B3 Johar Town, Lahore Pakistan'-> last segment is 'Lahore Pakistan', not 'pakistan'
    #     'Uk' / 'India'                      -> the whole value IS the country
    # A blank Country weakens classify_location_usa, which reads that field, and the bare-country
    # cases also violate the "Location is city/state only, never a country" rule.
    #
    # Anchored with (?<![a-z]) ... \s*$ rather than a plain `in` test: FOREIGN_COUNTRIES holds
    # 2-letter terms like 'uk', and an unanchored match is exactly the bug just fixed in
    # geo.is_strong_usa (it would fire on 'Columb-us', 'Padu-ka', etc). Longest name first so
    # 'south korea' wins over 'korea'.
    low = raw.strip().lower()
    for fc in sorted(FOREIGN_COUNTRIES, key=len, reverse=True):
        if not re.search(r"(?<![a-z])" + re.escape(fc) + r"\s*$", low):
            continue
        parent = _PROVINCE_TO_COUNTRY.get(fc)
        if parent:
            # A province, not a country - keep it in the location exactly as the comma
            # path above does, and report the parent country.
            return raw.strip(), parent
        head = re.sub(r"(?<![a-z])" + re.escape(fc) + r"\s*$", "", raw.strip(), flags=re.I)
        head = head.strip(" ,-–—")
        # Short forms are acronyms and must stay upper-case - .title() alone turns 'uk' into
        # the wrong-looking 'Uk' and 'uae' into 'Uae'.
        country = fc.upper() if len(fc) <= 3 else fc.title()
        # Nothing before the country means this was a bare country, which is not a Location.
        return (head or "N/A"), country

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
_CASE_SENSITIVE_SKILL_FORMS = {
    "swift": "Swift",
    # 'Maya' (Autodesk's 3D tool, added to the vocabulary 2026-08-21) is also a common given
    # name. Scanning the RAW text rather than the lowercased copy is what lets the
    # _SKILL_EXCLUDE_PHRASES entry below strip the "Maya <Surname>" name shape - that check
    # needs the surname's capital letter, which is gone from the lowercased haystack.
    "maya": "Maya",
}

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
    # ── Guards for the 2026-08-21 graphics/design vocabulary ──────────────────────────
    # 'Maya' scans the RAW text (see _CASE_SENSITIVE_SKILL_FORMS), so these patterns can use
    # capitalisation to tell the 3D tool from a person called Maya. "Maya Patel" / "Maya R.
    # Sharma" are stripped; "Maya, Blender, ZBrush" and "Autodesk Maya" still match.
    "maya": (r"Maya\s+(?:[A-Z]\.\s*)?[A-Z][a-z]+", r"\bMs\.?\s+Maya\b", r"\bMr\.?\s+Maya\b"),
    # 'Unity' the game engine vs. the ordinary noun ("unity of purpose", "team unity",
    # "unity and collaboration") - stock phrasing in soft-skill/leadership sections.
    "unity": (r"unity\s+of\b", r"team\s+unity", r"unity\s+and\s+(?:collaboration|purpose|teamwork)",
              r"in\s+unity\b", r"sense\s+of\s+unity"),
    # 'Sketch' the design tool vs. the verb ("sketch out a plan", "sketched wireframes").
    # Word-boundary matching already rejects "sketched"/"sketching"; this catches "sketch out".
    "sketch": (r"sketch(?:es)?\s+out\b", r"rough\s+sketch"),
    # 'Quota' as a sales metric vs. storage/system quotas in an infra resume.
    "quota": (r"(?:disk|storage|memory|api|rate)\s+quota", r"quota\s+(?:limit|exceeded)"),
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
    if not re.match(
            r'^(?:https?://)?(?:www\.)?[a-z0-9][a-z0-9.\-]*\.[a-z]{2,}(?:[/?#]|$)', v):
        return False
    # Degree abbreviations parse as perfectly valid domains: 'B.Sc' -> 'b.sc' matches the
    # pattern above because .sc is Seychelles. Live row APP-20260715-2257-7693 stored
    # 'http://b.sc' as Portfolio 3, harvested from the education line. Added 2026-08-04:
    # require the host label before the TLD to be at least two characters unless the URL
    # carries a real path - single-letter hosts are effectively never real portfolio links,
    # while degree strings ('b.sc', 'm.sc', 'b.tech', 'b.a') are exactly that shape.
    stripped = re.sub(r'^(?:https?://)?(?:www\.)?', '', v)
    host = stripped.split('/')[0].split('?')[0]
    has_path = bool(stripped[len(host):].strip('/'))
    labels = host.split('.')
    # Only reject a one-character host when there is NO path: 'b.sc' is a degree, but
    # 'x.com/someone' is a genuine profile link and must survive.
    if len(labels) >= 2 and len(labels[-2]) <= 1 and not has_path:
        return False
    if host in _DEGREE_LIKE_HOSTS and not has_path:
        return False
    return True


#: Degree abbreviations whose dotted form is also a syntactically valid domain.
_DEGREE_LIKE_HOSTS = {
    "b.sc", "m.sc", "b.tech", "m.tech", "b.a", "m.a", "b.e", "m.e", "b.ed", "m.ed",
    "b.com", "m.com", "ph.d", "b.s", "m.s", "b.arch",
}


def _is_meaningful_portfolio(url: str) -> bool:
    """True when a URL is plausibly the CANDIDATE'S OWN portfolio, not just any link.

    Added 2026-08-04 after auditing the live sheet. Two recurring non-portfolios:
      * a university/company HOME page with no path - APP-20260715-2231-25F9 stored
        'https://www.asu.edu/', which is his school, not his work;
      * a cloud FILE-SHARE link - APP-20260707-1727-0DF7 stored a Google Drive
        '/file/d/.../view' URL, which is almost certainly a copy of the resume we already
        have, not a portfolio.
    Both are honest extractions of real URLs, and both are useless in that column.
    """
    u = str(url or "").strip().lower()
    if not u or not _looks_like_url(u):
        return False
    host = re.sub(r'^(?:https?://)?(?:www\.)?', '', u).split('/')[0]
    path = u.split(host, 1)[-1].strip('/') if host in u else ""
    if re.search(r'(drive\.google\.com|dropbox\.com|onedrive\.live\.com|1drv\.ms)', host) \
            and re.search(r'/(file|s|document)/', u):
        return False
    # A CONTACT or MAP link is never a portfolio, with or without a path (added 2026-08-11).
    # Live values: 'https://maps.google.com/?q=Greater' (APP-20260630-1500-6A06) and
    # 'https://wa.me/+923170060278' (APP-20260603-0503-1245) - both are real URLs the
    # candidate put in their resume, neither is a sample of their work.
    if re.search(r'^(maps\.google\.[a-z.]+|goo\.gl|maps\.app\.goo\.gl|wa\.me|t\.me|'
                 r'api\.whatsapp\.com|m\.me)$', host):
        return False
    # A bare institutional home page (no path) identifies an organisation, not a person.
    # The country-coded forms matter as much as the bare ones: '.edu.pk' / '.edu.au' /
    # '.ac.in' end in the COUNTRY suffix, so the original '\.(edu|ac\.xx)$' never matched
    # them and both of these live values were stored as personal portfolios (2026-08-11):
    # 'https://www.uskt.edu.pk/' (his university) and 'https://pakaims.edu.pk/'.
    if not path and re.search(r'\.(edu|ac)(\.[a-z]{2,3})?$', host):
        return False
    return True


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
    from hiring_agent.geo import is_us_area_code
    raw = str(phone or "").strip()
    if not raw or raw.lower() in _GAP_LITERALS:
        return raw
    if "*" in raw:                       # deliberately masked — leave as the candidate wrote it
        return raw
    d = "".join(c for c in raw if c.isdigit())
    if len(d) == 11 and d.startswith("1"):        # US with country code (+1 / 1-...)
        if not is_us_area_code(d[1:4]):
            return raw
        return f"({d[1:4]}) {d[4:7]}-{d[7:]}"
    if len(d) == 10 and not raw.lstrip().startswith("+"):   # bare US 10-digit
        # ...but only when the candidate wrote it in a US SHAPE. A separator-free digit run
        # is indistinguishable from a foreign mobile written without its country code, and
        # geo._looks_like_us_phone refuses to trust one for exactly that reason (live case
        # Divy Parmar / APP-20260720-1013-09AC, whose Indian '7405465204' matched the NANP
        # area-code shape). That guard keys off "does this number carry a visual separator" -
        # so adding parens here MANUFACTURED the evidence it was looking for, and the guard
        # could never fire on any number this function had already touched.
        #
        # The parens also existed to stop Excel storing a bare digit run as a NUMBER; the
        # workbook-maintenance pass now forces the Phone column to Text, which covers that
        # without making a claim about the number's country.
        if not any(c in raw for c in " -.()"):
            return raw
        # An area code the NANP never issued to the US is a foreign number that happens to fit
        # the shape. Dressing it in US parens invents evidence the geo gate then trusts, which
        # is how a Pakistani mobile reached the main sheet as "(333) 333-8893" (2026-09-08).
        if not is_us_area_code(d[0:3]):
            return raw
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
    # A broad international-phone regex can otherwise accept an education range such as
    # "2023 - 2025" (eight digits plus a valid phone separator).  A pair of calendar years
    # is evidence about dates, never a callable number.
    if re.fullmatch(r"(?:19|20)\d{2}\s*[-–—/]\s*(?:19|20)\d{2}", raw):
        return "Not extracted"
    if re.fullmatch(r"(?:\d\s+){6,}\d", raw):
        return "Not extracted"
    if "*" in raw or sum(c.isdigit() for c in raw) < 7:
        return "Not extracted"
    return raw


def _extract_phone(text: str) -> str | None:
    """Extract the best phone number from text, preferring labeled numbers."""
    def acceptable(value: str) -> bool:
        return sanitize_phone(value).strip().lower() not in _GAP_LITERALS

    for line in (text or "").splitlines():
        low = line.lower().strip()
        if any(k in low for k in ("phone", "mobile", "cell", "tel:", "contact")):
            searchable = _RAW_URL_RE.sub(" ", line)
            for pat in _PHONE_PATTERNS:
                m = pat.search(searchable)
                if m and acceptable(m.group(0).strip()):
                    return m.group(0).strip()
    searchable_text = _RAW_URL_RE.sub(" ", text or "")
    for pat in _PHONE_PATTERNS:
        m = pat.search(searchable_text)
        if m and acceptable(m.group(0).strip()):
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
        "education_start_date": "",
        "education_end_date": "",
        "experience": "",
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

    edu_start, edu_end = _extract_education_dates(text)
    extracted["education_start_date"] = edu_start
    extracted["education_end_date"] = edu_end

    extracted["experience"] = _extract_experience(text)

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


_RAPID_OCR_ENGINE = None


def _get_rapidocr_engine():
    """Lazy-load and cache RapidOCR instance to avoid model reload overhead."""
    global _RAPID_OCR_ENGINE
    if _RAPID_OCR_ENGINE is None:
        from rapidocr_onnxruntime import RapidOCR
        _RAPID_OCR_ENGINE = RapidOCR()
    return _RAPID_OCR_ENGINE


def ocr_health() -> tuple[bool, str]:
    """Check if at least one OCR engine (RapidOCR or Tesseract) is available."""
    try:
        engine = _get_rapidocr_engine()
        if engine is not None:
            return True, "RapidOCR (ONNX) available"
    except Exception as e:
        logger.debug(f"RapidOCR health probe failed: {e}")

    try:
        tess = find_tesseract()
        if tess:
            return True, f"Tesseract available at {tess}"
    except Exception as e:
        logger.debug(f"Tesseract health probe failed: {e}")

    return False, "No OCR engine available (RapidOCR and Tesseract both unavailable)"


def _ocr_process(connection, raw):
    try:
        connection.send(_ocr_pdf_unbounded(raw))
    except Exception:
        connection.send("")
    finally:
        connection.close()


def _ocr_pdf(raw: bytes) -> str:
    """OCR runs in a killable process with a deadline, including native ONNX work."""
    import multiprocessing
    import os
    budget = float(os.getenv("HIRING_OCR_TIMEOUT", "120"))
    if budget <= 0:
        raise ValueError("HIRING_OCR_TIMEOUT must be positive")
    context = multiprocessing.get_context("spawn")
    reader, writer = context.Pipe(duplex=False)
    process = context.Process(target=_ocr_process, args=(writer, raw))
    process.start()
    writer.close()
    try:
        if not reader.poll(budget):
            raise TimeoutError(f"OCR exceeded {budget:g}s")
        return reader.recv()
    finally:
        reader.close()
        process.join(timeout=1)
        if process.is_alive():
            process.terminate()
            process.join(timeout=5)


def _ocr_pdf_unbounded(raw: bytes) -> str:
    """High-accuracy OCR for scanned/image-only PDFs.
    Primary: RapidOCR (PaddleOCR ONNX) - fast, layout-accurate, zero external binary dependency.
    Fallback: PyMuPDF + pytesseract + Tesseract binary if RapidOCR fails.
    """
    # 1. Primary: RapidOCR (pure ONNX runtime inside Python, no external .exe required)
    try:
        engine = _get_rapidocr_engine()
        import pymupdf
        from PIL import Image
        doc = pymupdf.open(stream=raw, filetype="pdf")
        parts = []
        for page in doc:
            pix = page.get_pixmap(dpi=150)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            ocr_res, _ = engine(img)
            if ocr_res:
                parts.extend([line[1] for line in ocr_res])
        doc.close()
        res = "\n".join(parts).strip()
        if res:
            return res
    except Exception as e:
        logger.debug(f"RapidOCR unavailable ({e}); trying Tesseract.")

    # 2. Fallback: Tesseract OCR
    try:
        import pymupdf as fitz
        import pytesseract
        from PIL import Image
    except Exception:
        try:
            import fitz
            import pytesseract
            from PIL import Image
        except Exception:
            return ""

    _tess = find_tesseract()
    if _tess:
        pytesseract.pytesseract.tesseract_cmd = _tess
    try:
        parts = []
        doc = fitz.open(stream=raw, filetype="pdf")
        for page in doc:
            pix = page.get_pixmap(dpi=200)
            parts.append(pytesseract.image_to_string(Image.open(io.BytesIO(pix.tobytes("png"))), timeout=30) or "")
        doc.close()
        return "\n".join(parts).strip()
    except Exception as e:
        logger.warning(f"   OCR unavailable ({e}); leaving scanned PDF unread.")
        return ""


def _pdf_hyperlinks(reader) -> list:
    """Collect URI link annotations from every page (best-effort, never raises).

    Resume PDFs very often show only display text ('LinkedIn', 'GitHub') while the real
    URL lives in the page's /Annots - the text layer never contains it. Harvesting the
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


def _demarkdown(text: str) -> str:
    """Strip Markdown emphasis and structure markers from extracted PDF text.

    pymupdf4llm is the PRIMARY extractor and it returns Markdown, not plain text. Every
    deterministic extractor downstream was written against plain text, so the markers
    leaked straight into the fields and broke pattern matching in ways that were easy to
    misread as "the resume does not say":

      * "# **YASH VERMA**" made the name extractor return 'Not extracted'.
      * "with **9+ years** designing" never matched the experience pattern, because the
        asterisks sit between "years" and the following word, so Years Exp read 'Missing'
        on a resume that states the total plainly (live case, 2026-09-11).
      * Education and the notes field stored the raw "**" markers verbatim.

    Fixing it here rather than in each regex means one rule instead of a dozen, and the
    LLM also sees cleaner input. Link text is kept and the target dropped, since resumes
    print the URL as the visible text anyway and extract_portfolios reads it from there.
    """
    if not text:
        return text
    import re as _re
    text = _re.sub(r"(?i)<br\s*/?>", "\n", text)
    text = _re.sub(r"(?i)</?(?:u|sup|sub|span|b|i|strong|em)(?:\s[^>]*)?>", "", text)
    text = _re.sub(r"!?\[([^\]]*)\]\([^)]*\)", r"\1", text)   # [label](url) -> label
    text = _re.sub(r"\*\*\*(.+?)\*\*\*", r"\1", text, flags=_re.S)
    text = _re.sub(r"\*\*(.+?)\*\*", r"\1", text, flags=_re.S)
    text = _re.sub(r"(?<!\w)_{2}(.+?)_{2}(?!\w)", r"\1", text, flags=_re.S)
    text = _re.sub(r"(?<!\w)_(.+?)_(?!\w)", r"\1", text, flags=_re.S)
    text = text.replace("~~", "")
    text = _re.sub(r"(?m)^\s{0,3}#{1,6}\s*", "", text)             # headings
    text = _re.sub(r"(?m)^\s{0,3}>\s?", "", text)                  # block quotes
    text = _re.sub(r"(?m)^\s*([-*+])\s+", "", text)                # bullets
    text = _re.sub(r"(?m)^\s*\|", " ", text)                       # table pipes
    text = _re.sub(r"(?m)^[\s|:-]{4,}$", "", text)                  # table rules
    text = text.replace("`", "")
    return text


def _pdf_layout_education_rows(doc, extracted_text: str) -> list[str]:
    """Recover education rows whose right-hand date was separated by Markdown ordering.

    PyMuPDF4LLM normally gives the best reading order, but a wide date column can be emitted
    after the following section.  PyMuPDF's coordinate-sorted plain text keeps words sharing
    a visual row together.  Add only degree rows that contain an explicit date and are absent
    from the primary text, so the scorer gains the missing evidence without duplicating the
    whole resume or changing its normal reading order.
    """
    existing = {re.sub(r"\s+", " ", line).strip()
                for line in str(extracted_text or "").splitlines() if line.strip()}
    recovered = []
    for page in doc:
        try:
            visual_lines = page.get_text("text", sort=True).splitlines()
        except Exception:
            continue
        for line in visual_lines:
            row = re.sub(r"\s+", " ", line).strip()
            if not row or row in existing or not _DEGREE_RE.search(row):
                continue
            if not (_EDU_DATE_RANGE_RE.search(row)
                    or _EDU_DATE_RANGE_RE_LOOSE.search(row)
                    or _EDU_EXPECTED_RE.search(row)
                    or _EDU_GRADUATED_RE.search(row)
                    or _EDU_MONTH_YEAR_RE.search(row)
                    or _EDU_ANY_SINGLE_DATE_RE.search(row)):
                continue
            recovered.append(row)
            existing.add(row)
    return recovered


def _docx_text(raw: bytes) -> str:
    """Text of a .docx, paragraphs plus table cells, with any hyperlinks appended."""
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


def _rtf_text(raw: bytes) -> str:
    """Plain text out of an RTF body.

    Several editors write RTF when asked to "save as .doc", so this is reached through
    the .doc branch rather than by extension. Deliberately small: drop control words,
    decode \\'xx hex escapes, and unwrap groups. Enough for a resume's prose, which is
    all the scorer reads.
    """
    body = raw.decode("cp1252", "replace")
    body = re.sub(r"\\\*\\[a-zA-Z]+(?:-?\d+)?[ ]?(?:\{[^{}]*\})?", "", body)
    body = re.sub(r"\{\\(?:fonttbl|colortbl|stylesheet|info|pict|object)[^{}]*"
                  r"(?:\{[^{}]*\}[^{}]*)*\}", "", body)
    body = re.sub(r"\\'([0-9a-fA-F]{2})",
                  lambda m: bytes([int(m.group(1), 16)]).decode("cp1252", "replace"), body)
    body = re.sub(r"\\(?:par|line|sect|page)\b[ ]?", "\n", body)
    body = re.sub(r"\\(?:tab)\b[ ]?", "\t", body)
    body = re.sub(r"\\[a-zA-Z]+(?:-?\d+)?[ ]?", "", body)
    body = body.replace("{", "").replace("}", "").replace("\\\n", "\n")
    return "\n".join(ln.rstrip() for ln in body.split("\n")).strip()


def _word97_text(raw: bytes) -> str:
    """Plain text out of a Word 97-2003 binary .doc (OLE2 compound file).

    The text is not stored contiguously: the FIB names a table stream holding a piece
    table, and each piece points at a run in the WordDocument stream that is either
    cp1252 single-byte or UTF-16LE, flagged by bit 30 of the piece's file offset.
    Reading the stream raw instead of following the pieces yields the document's text
    interleaved with deleted revisions and field codes, which is why this walks the
    table properly.

    Returns "" for anything it cannot read, which sends the row down P2's existing
    unreadable-resume path rather than publishing half a document as if it were whole.
    """
    try:
        import olefile
    except ImportError:
        logger.warning("   olefile is not installed; legacy .doc resumes cannot be read. "
                       "Fix: pip install olefile")
        return ""
    try:
        ole = olefile.OleFileIO(io.BytesIO(raw))
    except Exception as e:
        logger.warning(f"   legacy .doc is not a readable OLE2 file: {e}")
        return ""
    try:
        if not ole.exists("WordDocument"):
            return ""
        wd = ole.openstream("WordDocument").read()
        if len(wd) < 0x1A6 + 4:
            return ""
        # fibBase.flags bit 9 picks which of the two table streams is the live one.
        flags = struct.unpack_from("<H", wd, 0x0A)[0]
        table_name = "1Table" if flags & 0x0200 else "0Table"
        if not ole.exists(table_name):
            return ""
        table = ole.openstream(table_name).read()
        fc_clx, lcb_clx = struct.unpack_from("<II", wd, 0x1A2)
        clx = table[fc_clx:fc_clx + lcb_clx]

        # The Clx is zero or more Prc blocks (0x01) followed by the Pcdt (0x02).
        offset, pcdt = 0, b""
        while offset < len(clx):
            if clx[offset] == 0x01:
                offset += 3 + struct.unpack_from("<h", clx, offset + 1)[0]
            elif clx[offset] == 0x02:
                size = struct.unpack_from("<I", clx, offset + 1)[0]
                pcdt = clx[offset + 5:offset + 5 + size]
                break
            else:
                break
        if len(pcdt) < 16:
            return ""

        count = (len(pcdt) - 4) // 12
        cps = struct.unpack_from("<%dI" % (count + 1), pcdt, 0)
        chunks = []
        for index in range(count):
            base = 4 * (count + 1) + 8 * index
            fc = struct.unpack_from("<I", pcdt, base + 2)[0]
            length = cps[index + 1] - cps[index]
            if fc & 0x40000000:
                start = (fc & ~0x40000000) // 2
                chunks.append(wd[start:start + length].decode("cp1252", "replace"))
            else:
                chunks.append(wd[fc:fc + length * 2].decode("utf-16-le", "replace"))
        text = "".join(chunks)
    except Exception as e:
        logger.warning(f"   legacy .doc piece table could not be read: {e}")
        return ""
    finally:
        ole.close()

    # Word's in-band markers: \r ends a paragraph, \x07 a table cell/row, and
    # \x13-\x15 bracket field codes whose result is already in the text.
    for marker, replacement in (("\r", "\n"), ("\x07", "\n"), ("\x0b", "\n"),
                                ("\x0c", "\n"), ("\x13", ""), ("\x14", ""), ("\x15", "")):
        text = text.replace(marker, replacement)
    return "\n".join(line.rstrip() for line in text.split("\n")).strip()


def extract_text_from_bytes(raw: bytes, filename: str) -> str:
    """Extract text from raw PDF/DOC/DOCX bytes (layout-aware, with OCR fallback)."""
    name = (filename or "").lower()
    if not raw:
        return ""
    try:
        if name.endswith(".pdf"):
            reader = None
            try:
                from pypdf import PdfReader
                reader = PdfReader(io.BytesIO(raw))
            except Exception:
                reader = None

            # 1. Primary: PyMuPDF4LLM for structured, reading-order-aware Markdown extraction
            text = ""
            try:
                import pymupdf
                import pymupdf4llm
                doc = pymupdf.open(stream=raw, filetype="pdf")
                text = _demarkdown(pymupdf4llm.to_markdown(doc) or "")
                layout_rows = _pdf_layout_education_rows(doc, text)
                if layout_rows:
                    text = text.rstrip() + "\n\n" + "\n".join(layout_rows) + "\n"
                doc.close()
            except Exception as e:
                logger.debug(f"   pymupdf4llm extraction failed ({e}); trying pypdf.")

            # 2. Fallback to pypdf if pymupdf4llm returned empty or failed
            if not text.strip() and reader is not None:
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
                if len(ocr_text.strip()) > len(text.strip()) / 3 and len(ocr_text.strip()) >= 50:
                    text = ocr_text
                    logger.info("   used OCR (PDF text layer was empty or low quality)")
                else:
                    # Low quality or corrupt text layer, and OCR yielded no usable text.
                    # Clear text completely so regex extractors never run on garbage fragments or email bodies!
                    text = ""
                    logger.warning(f"   OCR failed or unavailable for '{filename or '?'}' (low quality text layer); "
                                   "marked unreadable (no regex fallback).")
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
            if text.strip() and reader is not None:
                links = _pdf_hyperlinks(reader)
                if links:
                    text += "\nLinks in document: " + " ".join(links)
            return text
        if name.endswith(".docx"):
            return _docx_text(raw)
        if name.endswith(".doc"):
            # Legacy Word. The extension is the least reliable thing about these files,
            # so route on the actual magic bytes: applicants rename a .docx to .doc, and
            # "Save as .doc" in several editors writes RTF. Only a genuine OLE2 compound
            # file gets the Word 97 binary reader.
            if raw[:4] == b"PK\x03\x04":
                return _docx_text(raw)
            if raw[:5] == b"{\\rtf":
                return _rtf_text(raw)
            if raw[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
                return _word97_text(raw)
            logger.warning(f"   '{filename or '?'}' is named .doc but is not Word or RTF; "
                           f"no text extracted.")
            return ""
        if name.endswith((".png", ".jpg", ".jpeg")):
            try:
                engine = _get_rapidocr_engine()
                from PIL import Image
                img = Image.open(io.BytesIO(raw))
                ocr_res, _ = engine(img)
                if ocr_res:
                    return "\n".join([line[1] for line in ocr_res]).strip()
            except Exception as img_err:
                logger.warning(f"   image OCR failed for '{filename or '?'}': {img_err}")
            return ""
    except TimeoutError:
        raise  # infrastructure timeout must not become an unreadable-CV rejection
    except Exception as e:
        logger.warning(f"   could not read '{filename or '?'}': {e}")
    return ""


# ── AI extraction (Ollama - the only AI brain actually configured in this deployment) ────

#: How the LAST extraction actually got its fields. Set by extract_with_ollama and read by
#: extract_candidate_details_smart, which copies it onto the details dict as
#: ``_extraction_source``. A leading underscore keeps it out of the 33-column contract:
#: _P2_COL_TO_KEY maps columns to keys explicitly, so an extra key is simply never written.
#:
#: Module-level because the two functions are a pair and the pipeline scores one candidate at
#: a time. It is reset at the start of every extraction, so a stale value cannot leak from the
#: previous candidate into this one. Do not read it from a worker thread.
EXTRACTION_SOURCE = {"value": "offline (Ollama disabled)"}

# Live qwen3-1.7b-p2 measurement, 2026-09-06: the largest non-thinking reply
# across these four calls was 121 tokens (421 characters), rising to 124 with
# schemas. 512 gives >4x headroom. Measurements: docs/benchmarks/stage2a/.
_EXTRACTION_NUM_PREDICT = 512


def _string_fields_schema(*fields: str) -> dict:
    """Require the existing string fields, including empty strings for gaps."""
    return {
        "type": "object",
        "properties": {field: {"type": "string"} for field in fields},
        "required": list(fields),
        "additionalProperties": False,
    }


_CANDIDATE_FORMAT = _string_fields_schema(
    "full_name", "phone", "location", "country", "skills", "looking_for_role", "education")
_ROLE_SUMMARY_FORMAT = _string_fields_schema("summary")
_PORTFOLIO_FORMAT = _string_fields_schema("linkedin", "github", "other")


def _post_ollama_extraction(url: str, **kwargs):
    """Retry a requests timeout once, independently of the row give-up counter.

    HTTP errors, bad JSON, and connection failures keep the caller's existing
    handling. Parsing stays outside the retry loop so it cannot trigger a retry.
    """
    import requests

    for attempt in range(2):
        try:
            return requests.post(url, **kwargs)
        except requests.Timeout:
            if attempt == 1:
                raise
            logger.warning(
                "       Ollama extraction timed out after %ss; retrying once.",
                kwargs.get("timeout"),
            )


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
        EXTRACTION_SOURCE["value"] = ("offline (Ollama disabled)" if not OLLAMA_ENABLED
                                      else "offline (no resume text)")
        return None
    EXTRACTION_SOURCE["value"] = "ollama"
    hint_block = ""
    if hints:
        lines = "\n".join(f"- {k}: {v}" for k, v in hints.items() if v)
        if lines:
            hint_block = (
                "\n\nA deterministic pattern-matcher already found these candidate values "
                "from the text below - confirm each is correct, or correct it if the text "
                "says otherwise:\n" + lines
            )
    payload = {
        "model": OLLAMA_MODEL,
        "format": _CANDIDATE_FORMAT,
        "think": False,
        "stream": False,
        "options": {"temperature": 0, "num_predict": _EXTRACTION_NUM_PREDICT,
                            "seed": OLLAMA_SEED},
        "messages": [
            {"role": "system",
             "content": _AI_PROMPT + hint_block + " Respond ONLY with a JSON object "
             "whose keys are full_name, phone, location, country, skills, "
             "looking_for_role, education."},
            {"role": "user", "content": text[:AI_TEXT_LIMIT]},
        ],
    }

    def _ask():
        r = _post_ollama_extraction(f"{OLLAMA_HOST}/api/chat", json=payload,
                                    timeout=OLLAMA_TIMEOUT)
        r.raise_for_status()
        return json.loads(r.json()["message"]["content"])

    try:
        try:
            data = _ask()
        except json.JSONDecodeError:
            # A truncated or malformed body is transient in exactly the way a timeout is,
            # but it was the one failure with no second attempt: _post_ollama_extraction
            # retries the REQUEST and deliberately keeps parsing outside that loop, so a
            # single bad response deferred the candidate outright. Seen live 2026-09-11,
            # "Unterminated string starting at line 8 column 13", on a resume that scored
            # fine on the retry. One more attempt, then fall through as before.
            logger.warning("       Ollama returned malformed JSON; retrying once.")
            data = _ask()
    except Exception as e:
        # requests is imported lazily inside the helper. Keep provenance classification
        # independent of that import, so a missing dependency also falls back cleanly.
        if type(e).__name__ in ("Timeout", "ReadTimeout", "ConnectTimeout"):
            # An ordinary two-page resume measured 62.1s against the former 60s budget on
            # the client's hardware (2026-09-06). Both attempts have now timed out. Falling
            # through to the regex parser silently published regex-quality fields as if the
            # model had read the resume, and the row still went out as 'Scored' with nothing
            # on it to say otherwise. Recorded loudly so a degraded run is visible and the
            # affected rows can be re-scored later.
            EXTRACTION_SOURCE["value"] = f"offline (Ollama timed out after {OLLAMA_TIMEOUT}s)"
            logger.warning(
                f"       DEGRADED : Ollama extraction TIMED OUT after two attempts "
                f"({OLLAMA_TIMEOUT}s each) - "
                f"falling back to the offline parser. This candidate's fields are "
                f"regex-quality, not model-read."
            )
            return None
        EXTRACTION_SOURCE["value"] = f"offline ({type(e).__name__})"
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
    # Group and de-duplicate BEFORE counting against the cap. A raw comma split shredded
    # 'AWS (EC2, S3, IAM)' into three entries and kept 'AWS' beside it, so those phantom
    # entries used up slots and genuine skills at the end of the list were cut: live row
    # APP-20260908-1304-MCLA lost Microservices, SQLite and n8n when three new vocabulary
    # terms were added (2026-09-14). normalize_skills keeps a bracketed group as one entry.
    raw_skills = normalize_skills(result.get("skills", ""))
    base = [s.strip() for s in raw_skills.split(", ")
            if s.strip() and s.strip().lower() != "not extracted"]
    # 'AWS (EC2/S3/IAM)' already covers the keyword 'aws'; do not append a plain 'AWS' again.
    have = {s.lower() for s in base} | {re.sub(r"\s*\(.*\)\s*$", "", s).lower() for s in base}
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


def settle_geography(current_location, current_country,
                     proposed_location, proposed_country,
                     resume_text: str, mail_body: str = "",
                     log_tag: str = "extractor") -> tuple[str, str]:
    """The ONE rule for whether a model's geography claim replaces the current one.

    Location and Country are a single claim: a model that invented the city has no more
    authority on the country it paired with it, so they are accepted or rejected together.

    Four checks, in order of how badly each was needed live:
      1. A location in the CONTACT BLOCK (resume header or mail signature) describes the
         candidate now. A model must not replace it with a city that merely appears in an
         older job or degree - APP-20260419-1833-C006, 'United States' -> 'Kerala'.
      2. A location that appears NOWHERE in the text is invented outright.
      3. Grounding only proves a string is present, not that it is a PLACE: an institution
         name or a section header passes grounding cleanly while being the wrong field.
      4. A country is acceptable only if the text says it, or the accepted location implies
         it. Anything else is the model reasoning from a dial code or a school's reputation
         - APP-20260716-1037-6CE3, Country 'India' on a resume where the word never appears.
    """
    both = f"{resume_text or ''}\n{mail_body or ''}"
    cur_loc = clean_location_text(current_location) or str(current_location or "").strip()
    new_loc = clean_location_text(proposed_location) or str(proposed_location or "").strip()
    cur_country = str(current_country or "").strip()
    new_country = str(proposed_country or "").strip()
    if new_loc.lower() in _GAP_LITERALS:
        new_loc = ""

    loc, country = (new_loc, new_country) if new_loc else (cur_loc, cur_country)

    def _in_contact(v: str) -> bool:
        return bool(v) and (_location_grounded_in_contact_block(v, resume_text)
                            or _location_grounded_in_contact_block(v, mail_body))

    if (new_loc and cur_loc and new_loc.lower() != cur_loc.lower()
            and _in_contact(cur_loc) and not _in_contact(new_loc)):
        logger.warning(f"   {log_tag}: AI location {new_loc!r} is outside the contact "
                       f"header - keeping current-location value {cur_loc!r}")
        loc, country = cur_loc, cur_country
    elif new_loc and not _location_grounded_in_text(new_loc, both):
        logger.warning(f"   {log_tag}: AI location {new_loc!r} is not supported by the "
                       f"resume/mail text - keeping {cur_loc or 'Not extracted'!r}")
        loc, country = cur_loc, cur_country

    if not location_is_plausible(loc):
        logger.warning(f"   {log_tag}: location {loc!r} is grounded in the resume but not "
                       f"plausibly a place (institution/degree-sentence/section-header) - "
                       f"falling back to {cur_loc or 'Not extracted'!r}")
        if location_is_plausible(cur_loc):
            loc, country = cur_loc, cur_country
        else:
            loc, country = "Not extracted", ""

    if country and country.lower() not in _GAP_LITERALS \
            and not _country_grounded_in_text(country, loc, both):
        logger.warning(f"   {log_tag}: AI country {country!r} is not supported by the "
                       f"resume/mail text and is not implied by {loc!r} - dropping it "
                       f"(will read 'Missing')")
        # An unsupported CURRENT value must not be preserved either - that is the value we
        # are trying to get rid of.
        country = (cur_country
                   if cur_country and _country_grounded_in_text(cur_country, loc, both)
                   else "")
    return loc or "Not extracted", country


def finalize_geography_shape(location, country, resume_text: str) -> tuple[str, str]:
    """The final SHAPE of the pair, once both values are settled.

    Ran only inside the first pass, which meant the recheck - which runs after it - could
    hand back a shape it had already normalised away. Now both passes end here.
    """
    loc = str(location or "").strip()
    country = str(country or "").strip()

    # A location that came back combined with its country (the model still does this
    # occasionally despite the prompt).
    if not country or country == "Not extracted":
        if loc and loc != "Not extracted":
            loc, country = split_location_country(loc)

    # Ollama can correct 'location' from real text evidence while still echoing a now-stale
    # 'country' hint verbatim - live case Syyed Nazir Ali / APP-20260727-1431-6F75, whose
    # row read India/United States at once. Location is the value that was grounded against
    # real text, so when the two name conflicting countries, location wins.
    loc_low, country_low = loc.lower(), country.lower()
    if loc_low in FOREIGN_COUNTRIES and country_low in US_COUNTRY_TERMS:
        country = loc.title()
    elif loc_low in US_COUNTRY_TERMS and country_low in FOREIGN_COUNTRIES:
        country = "United States"

    # Location holds city/state only - Country is the dedicated field for the country name.
    # Same live case: his resume's only location statement is "Location: India", so 'India'
    # ended up in BOTH columns.
    final_low = loc.lower()
    if final_low in US_COUNTRY_TERMS and _location_grounded_in_contact_block(loc, resume_text):
        # A candidate who states only "United States" in the contact header is telling us
        # something honest and current. Less specific than city/state, but it answers the
        # USA eligibility question, so it is kept rather than blanked to N/A.
        return "United States", "United States"
    if final_low and (final_low in FOREIGN_COUNTRIES or final_low in US_COUNTRY_TERMS):
        loc = ("Remote"
               if re.search(r"\bremote\b|\bwork[\s-]*from[\s-]*home\b|\bwfh\b",
                            resume_text or "", re.IGNORECASE)
               else "N/A")
    return loc, country


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
        # Deterministic-only, like the portfolios above - dates are a format-driven regex
        # match, not something an LLM should be asked to interpret/invent. Never gated
        # through the Tier 2 Ollama pass; always the same value the offline scan found.
        "education_start_date": baseline.get("education_start_date", ""),
        "education_end_date": baseline.get("education_end_date", ""),
        # Deterministic-only for the same reason as the education dates above: a stated
        # "6+ years of experience" is a literal read, not a judgement call, and a work-
        # history span is arithmetic. Neither is something to route through an LLM.
        "experience": baseline.get("experience", ""),
        "notes": baseline.get("notes", text[:800]),
    }
    tier2_fields = ("full_name", "location", "country", "looking_for_role", "education")

    if not OLLAMA_ENABLED:
        logger.info("   extractor: offline parser (Ollama disabled)")
        result["phone"] = baseline.get("phone", "Not extracted")
        result["skills"] = baseline.get("skills", "Not extracted")
        for f in tier2_fields:
            result[f] = baseline.get(f, "" if f == "country" else "Not extracted")
        result["education"] = complete_education(result.get("education"), text)
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
        result["education"] = complete_education(result.get("education"), text)
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

    # ONE geography rule, shared with ai_recheck_fields - see settle_geography.
    result["location"], result["country"] = settle_geography(
        baseline.get("location", "Not extracted"), baseline.get("country", ""),
        ollama.get("location", ""), result.get("country", ""), text)
    result["location"], result["country"] = finalize_geography_shape(
        result["location"], result["country"], text)
    # The model may drop the school even though the prompt asks for it; restore it from the
    # resume's own degree block so the column is consistent (see complete_education).
    result["education"] = complete_education(result.get("education"), text)

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
        resp = _post_ollama_extraction(
            f"{OLLAMA_HOST}/api/chat",
            json={
                "model": OLLAMA_MODEL,
                "format": _CANDIDATE_FORMAT,
                "think": False,
                "stream": False,
                "options": {"temperature": 0, "num_predict": _EXTRACTION_NUM_PREDICT,
                            "seed": OLLAMA_SEED},
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
        if globals().get("STRICT_AI_STAGES", False) or REQUIRE_AI:
            raise RuntimeError("Independent AI recheck failed; result deferred") from e
        logger.warning(f"       Recheck   : Ollama unavailable ({e}); keeping original fields.")
        return fields

    improved = dict(fields)
    changed = []
    _proposed: dict = {}
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
        _both_texts = f"{resume_text or ''}\n{mail_body or ''}"
        if key == "looking_for_role" and new_val:
            new_val = clean_role_text(new_val) or new_val
        # Geography is NOT decided here. Location and Country are one claim and they are
        # settled together, once, below - by the same settle_geography the first-pass merge
        # uses. This loop only records what the model proposed.
        if key in ("location", "country"):
            _proposed[key] = new_val
            continue
        if new_val and new_val != old_val:
            changed.append(key)   # counts gap-fills too, so the log never lies
        if new_val:
            improved[key] = new_val
        elif old_val:
            improved[key] = old_val
    # ONE geography rule, shared with the first-pass merge - see settle_geography. This
    # pass is the one that also gets to read the MAIL BODY: a candidate's signature block
    # is a contact block too.
    _settled = settle_geography(
        fields.get("location", ""), fields.get("country", ""),
        _proposed.get("location", ""), _proposed.get("country", ""),
        resume_text, mail_body, log_tag="Recheck")
    _settled = finalize_geography_shape(_settled[0], _settled[1], resume_text)
    for _k, _v in zip(("location", "country"), _settled):
        if _v and _v != str(fields.get(_k, "") or "").strip():
            changed.append(_k)
        improved[_k] = _v

    improved["education"] = complete_education(improved.get("education", ""), resume_text)
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
            resp = _post_ollama_extraction(
                f"{OLLAMA_HOST}/api/chat",
                json={
                    "model": OLLAMA_MODEL,
                    "format": _ROLE_SUMMARY_FORMAT,
                    "think": False,
                    "stream": False,
                    "options": {"temperature": 0, "num_predict": _EXTRACTION_NUM_PREDICT,
                            "seed": OLLAMA_SEED},
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


def _country_grounded_in_text(country: str, location: str, text: str) -> bool:
    """Is an AI-supplied Country actually supported, or reasoned from circumstantial hints?

    True when EITHER the country name appears in the source text, OR it is the deterministic
    consequence of a location we accepted ("Austin, TX" implies the United States). False
    when the model reached it from a dial code, a university's country of origin, or any
    other indirect signal - which is inference dressed as extraction.

    Added 2026-08-04. _location_grounded_in_text guarded the CITY, and the pair-gate that
    used it dropped Country too - but only when the LOCATION was the ungrounded part. Get the
    city right and an invented country sailed through. Live case APP-20260716-1037-6CE3
    (Peter Vishal): resume reads "Remote" and "Open to Remote & Global Opportunities", and
    India/Indian/Hyderabad/Telangana appear ZERO times, yet Country was stored as "India" -
    inferred from a +91 dial code and an Indian college name. "Remote" was genuinely
    grounded, so the pair-gate passed, and the invented country then fed the geo filter and
    pointed at a rejection nobody had evidence for.

    A dial code is not an address: numbers are kept across borders, and a degree is a fact
    about the past. Where the candidate lives NOW is theirs to state - so an unsupported
    country is dropped, the cell reads 'Missing', and we ask.
    """
    ctry = str(country or "").strip()
    if not ctry or ctry.lower() in _GAP_LITERALS:
        return False
    low_text = (text or "").lower()
    if ctry.lower() in low_text:
        return True
    loc = str(location or "").strip()
    if loc and loc.lower() not in _GAP_LITERALS:
        from hiring_agent.geo import is_strong_usa, normalize_country
        _, implied = split_location_country(loc)
        if implied and implied.strip().lower() == ctry.lower():
            return True
        if normalize_country(ctry) == "United States" and is_strong_usa(loc):
            return True
    return False


def _location_grounded_in_text(location: str, text: str) -> bool:
    """Reject a location the AI returned that doesn't actually appear in the source text.

    The deterministic twin of _url_grounded_in_text, added 2026-08-03 to close a real
    asymmetry: Portfolio (Tier 1) had a hard anti-hallucination gate, but Location - the
    single field with the longest incident history in this file, and the one that decides
    whether a candidate is geo-rejected - had only a prompt instruction telling the model
    not to invent. A local model that ignored it silently overwrote a correct regex match:
    a resume reading "Austin, TX 78701" could be stored as "Bengaluru, Karnataka" / "India"
    and routed toward rejection, with every downstream guard passing because the invented
    pair is perfectly self-consistent. Cross-field checks catch CONTRADICTIONS; nothing
    caught confident-and-wrong.

    Every comma-segment must be evidenced in the text. A US state may be written either way
    round on either side ("Phoenix AZ" in the resume -> "Phoenix, Arizona" from the model,
    and vice versa), because normalising the abbreviation is a correction we WANT to keep -
    only genuinely unsupported segments are rejected.
    """
    loc = str(location or "").strip()
    if not loc or loc.lower() in _GAP_LITERALS:
        return False
    low_text = (text or "").lower()
    if not low_text:
        return False
    for segment in loc.split(","):
        seg = segment.strip().lower()
        if not seg:
            continue
        if seg in low_text:
            continue
        # "AZ" <-> "Arizona": accept whichever form the resume actually uses.
        alt = US_STATE_ABBREV_TO_NAME.get(seg) or US_STATE_NAME_TO_ABBREV.get(seg)
        if alt and re.search(r"(?<![a-z])" + re.escape(alt) + r"(?![a-z])", low_text):
            continue
        return False
    return True


_CONTACT_SECTION_BREAK_RE = re.compile(
    r"^(?:professional\s+)?(?:experience|employment|work\s+history|education|skills|"
    r"projects?|certifications?)\b",
    re.IGNORECASE,
)


def _location_grounded_in_contact_block(location: str, text: str) -> bool:
    """True when ``location`` appears in the resume/email contact header.

    A place mentioned anywhere in a resume is not necessarily the candidate's current
    location: education and experience sections contain historical cities and regions.
    Keep this intentionally small and deterministic—the leading contact block ends at the
    first common section heading or after twelve non-empty lines.
    """
    lines = []
    for raw in str(text or "").splitlines():
        line = re.sub(r"\s+", " ", raw).strip()
        if not line:
            continue
        if lines and _CONTACT_SECTION_BREAK_RE.match(_collapse_letter_spacing(line)):
            break
        lines.append(line)
        if len(lines) >= 12:
            break
    return bool(lines) and _location_grounded_in_text(location, "\n".join(lines))


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
        resp = _post_ollama_extraction(
            f"{OLLAMA_HOST}/api/chat",
            json={
                "model": OLLAMA_MODEL,
                "format": _PORTFOLIO_FORMAT,
                "think": False,
                "stream": False,
                "options": {"temperature": 0, "num_predict": _EXTRACTION_NUM_PREDICT,
                            "seed": OLLAMA_SEED},
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
