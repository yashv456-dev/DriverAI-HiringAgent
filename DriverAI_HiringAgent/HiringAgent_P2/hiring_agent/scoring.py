"""Role matching, scoring, and rematch logic."""

import pandas as pd

import hiring_agent.config as _cfg
import re as _re

from hiring_agent.config import (
    COLUMNS, DEFAULT_ROLES, GENERIC_TITLE_WORDS,
    SCORING_MIN_MATCH, SCORING_TOP_N, SCORING_TITLE_BOOST,
    SCORING_MIN_ROLE_SKILLS,
    OLLAMA_ENABLED, OLLAMA_SCORING,
    ROLE_CATEGORY_RULES, ROLE_CATEGORY_DEFAULT,
    logger,
)
from hiring_agent.jd_identity import role_display_name, role_opening_id, stamp_role_identity
from hiring_agent.ollama_scorer import ai_score_roles


_AI_ROLE_LIMIT = 12
_NON_JOB_ROLE_TITLE_RE = _re.compile(r"\b(?:offer\s+letter|offerletter|onboarding|benefits?)\b", _re.I)


def get_open_roles() -> list[dict]:
    """Default open roles from config.yaml, each carrying its opening identity.

    Copied per call (rather than handing out the config dicts themselves) so stamping the
    identity fields never mutates module-level config state.
    """
    return [stamp_role_identity(dict(r)) for r in DEFAULT_ROLES]


#: Tool names that are essentially never used outside native mobile app development —
#: unlike a job-title word, these are specific enough that their presence in a
#: candidate's own Skills list is stronger domain evidence than whichever JD posting
#: happened to score highest (see assign_category's skill-based check below).
_MOBILE_TOOL_SKILLS = {
    "kotlin", "swift", "swiftui", "flutter", "dart", "jetpack compose",
    "android sdk", "react native", "objective-c", "xcode",
}
#: Game-engine/graphics-specific tool names. A candidate listing one of these alongside
#: a mobile tool is a genuine graphics/game-dev candidate (e.g. a Unity mobile game
#: developer), so the mobile-skill override below must not fire for them.
_GRAPHICS_TOOL_SKILLS = {
    "unity", "unreal", "unreal engine", "opengl", "directx", "blender",
    "maya", "3ds max", "cinema 4d", "c4d", "godot", "shader", "hlsl", "glsl",
}


_AI_TOOL_SKILLS = {
    "langchain", "llamaindex", "rag", "agentic ai", "agentic", "azure openai",
    "hugging face", "prompt engineering", "fine-tuning", "peft", "lora",
}


#: A QA / test-automation professional, from what they say they want or from their tools.
#: Added 2026-09-14 with _catalog_has_qa_opening below. The JD catalog (105 documents, 54
#: openings) has no QA/SDET/test role, so a QA engineer's best matches are always developer
#: roles that share only generic tooling (Java, Git, CI/CD): live row APP-20260902-2155-MCPA,
#: an SDET, was filed first under Web Team (Full Stack 60%) and then Mobile Apps (Mobile
#: Developer 80%). A category named after a role the candidate does not do is misleading;
#: 'General' is honest. Three distinct test tools are required so a developer who merely
#: lists JUnit is not caught - across 280 past candidates only this one profile matched.
_QA_PREF_RE = _re.compile(
    r"(?i)\b(?:qa|sdet|quality\s+assurance|test(?:ing)?\s+(?:automation|engineer|analyst|lead)|"
    r"automation\s+test(?:er|ing)?|software\s+test(?:er|ing)?)\b")
_QA_TOOL_SKILLS = {"selenium", "cypress", "testng", "junit", "testrail", "jmeter",
                   "restassured", "appium", "zephyr", "bugzilla", "playwright", "cucumber"}
_QA_OPENING_RE = _re.compile(r"(?i)\b(?:qa|sdet|quality\s+(?:assurance|engineer)|test(?:ing|er)?)\b")


def _catalog_has_qa_opening() -> bool:
    """True once the cached JD catalog contains a QA/test opening (so the guard retires itself)."""
    try:
        from hiring_agent.jd_sources import _read_role_cache
        cached = _read_role_cache() or {}
        titles = [str(r.get("title", "")) for r in (cached.get("roles") or [])]
    except Exception:
        titles = []
    if not titles:
        titles = [str(r.get("title", "")) for r in get_open_roles()]
    return any(_QA_OPENING_RE.search(t) for t in titles)


def is_unmatched_qa_profile(skills_str: str = "", role_pref: str = "") -> bool:
    """A clear QA/test candidate while no QA/test opening exists to match them against."""
    skills = {s.strip().lower() for s in _re.split(r"[,;]", str(skills_str or "")) if s.strip()}
    is_qa = bool(_QA_PREF_RE.search(str(role_pref or ""))) or len(skills & _QA_TOOL_SKILLS) >= 3
    return is_qa and not _catalog_has_qa_opening()


def assign_category(role_str: str, skills_str: str = "", role_pref: str = "") -> str:
    """Map 'Role Title (NN%)' → business category using config-driven rules.

    Rules are defined in config.yaml under role_categories.rules (first-match-wins,
    case-insensitive substring). Returns role_categories.default when no rule matches.
    Adding a new JD only requires a new rule entry in config.yaml — no code change.
    Every candidate always gets a non-blank category: an unmatched or role-less
    candidate falls back to role_categories.default ("General") so the client's
    Category filter never hides anyone.

    `skills_str`, if given, is checked FIRST for an unambiguous native-mobile tool stack
    (Kotlin/Swift/Flutter/...). Added 2026-07-24: the winning JD title (Suggested Role 1)
    is sometimes a differently-domain-named posting that a mobile candidate only weakly
    matched (e.g. 'Gaming Position (23%)', 'Mobile Application Lead Developer Position
    Description v2 (80%)' — the latter landing in 'Senior' via the 'lead'
    keyword, section 1 precedence, same as the intentional 'Logistics Manager' ->
    'Senior' behavior tested elsewhere). That precedence is correct when the
    JD title genuinely reflects the candidate's level/domain, but a candidate's own tool
    stack is more reliable than a loosely-matched title, so a clear, uncontested mobile
    signal in Skills is trusted over the title match. Optional and defaults to empty so
    any caller without skills in scope keeps working unchanged.

    `role_pref` (Looking For Role), if given, feeds the QA guard checked before everything
    else: a clear QA/test candidate is 'General' while the catalog has no QA/test opening,
    rather than being named after whichever developer role scored highest.
    """
    if is_unmatched_qa_profile(skills_str, role_pref):
        return ROLE_CATEGORY_DEFAULT
    skills = {s.strip().lower() for s in _re.split(r"[,;]", str(skills_str or "")) if s.strip()}
    if skills & _MOBILE_TOOL_SKILLS and not (skills & _GRAPHICS_TOOL_SKILLS):
        return "Mobile Apps (Android IOS)"

    title = _re.sub(r'\s*\(\d+%\)\s*$', '', (role_str or "")).strip().lower()

    # Precedence 1: Senior Manager and Executive titles
    for rule in ROLE_CATEGORY_RULES:
        if rule.get("category") not in ("Senior Manager", "Executive"):
            continue
        needle = rule.get("match", "").lower()
        if not needle:
            continue
        if rule.get("word"):
            if _re.search(rf"(?<![a-z0-9]){_re.escape(needle)}(?![a-z0-9])", title):
                return rule["category"]
        elif needle in title:
            return rule["category"]

    _AUTOMATION_SKILLS = {"uipath", "rpa", "blue prism", "power automate", "automation anywhere"}
    # Precedence 2: Uncontested AI / Agentic tool stack (3+ specialized AI skills)
    if len(skills & _AI_TOOL_SKILLS) >= 3 and not (skills & _AUTOMATION_SKILLS) and not any(k in title for k in ("automation", "rpa", "uipath", "process developer")):
        return "AI/ML/CV (SIN2)"

    for rule in ROLE_CATEGORY_RULES:
        needle = rule.get("match", "").lower()
        if not needle:
            continue
        if rule.get("word"):
            # Opt-in whole-word match (added 2026-08-21 for the C-suite acronyms). Plain
            # substring is unusable for short tokens: 'cto' is inside contractor/doctor/
            # factory/sector/director, 'coo' inside coordinator, 'cio' inside suspicious.
            # Lookarounds rather than \b because a needle may contain non-word characters
            # (e.g. 'ai/ml'), where \b would anchor in the wrong place.
            if _re.search(rf"(?<![a-z0-9]){_re.escape(needle)}(?![a-z0-9])", title):
                return rule["category"]
        elif needle in title:
            return rule["category"]

    # Precedence 3: Unaligned technical skills fallback (if title doesn't match any category rule)
    _CYBERSEC_SKILLS = {"siem", "soc", "wireshark", "metasploit", "penetration testing",
                        "incident response", "kali linux", "firewall", "cissp", "ceh", "cism", "owasp",
                        "vulnerability assessment", "threat intelligence", "edr", "xdr", "ids/ips"}
    _SDE_SKILLS = {"fastapi", "django", "flask", "spring boot", "express.js", "react", "angular",
                   "vue", "node.js", "typescript", "graphql", "rest api"}

    if not (skills & _AUTOMATION_SKILLS) and not any(k in title for k in ("automation", "rpa", "uipath", "process developer")):
        if len(skills & _CYBERSEC_SKILLS) >= 2:
            return "Cybersecurity and IT Admin"
        if len(skills & _SDE_SKILLS) >= 2:
            return "Web Team (Full stack/Back end & UI/UX)"
        if len(skills & _AI_TOOL_SKILLS) >= 2:
            return "AI/ML/CV (SIN2)"

    return ROLE_CATEGORY_DEFAULT


def _skill_set(skills_str: str) -> set[str]:
    """Turn a 'Skills' string ('Python, SQL, AWS') into a lowercase set."""
    out = set()
    import re as _re
    text = str(skills_str or "")
    if _re.fullmatch(r"\s*\[.*\]\s*", text):
        text = text.strip()[1:-1]
    for piece in _re.split(r"[,;&:]+", text):
        s = piece.strip(" \t\r\n'\"").lower()
        # 'nan'/'none' guard: blank Excel cells stringify to these, not real skills
        if s and s not in ("not extracted", "nan", "none", "n/a") and len(s) <= 40:
            out.add(s)
        # 'aws (ec2/s3/iam)' also means 'aws', 'ec2', 's3' and 'iam'. normalize_skills keeps a
        # bracketed group as one skill and drops the separate plain 'AWS', so the base and the
        # items must be read back out here or that match would silently disappear.
        m = _re.fullmatch(r"(.+?)\s*\((.+)\)", s)
        if m:
            for part in [m.group(1)] + _re.split(r"\s*/\s*", m.group(2)):
                part = part.strip()
                if part and len(part) <= 40:
                    out.add(part)
    return out

_SKILL_EXPANSIONS = {
    "agentic ai": {"agentic", "agentic ai"},
    "agentic": {"agentic", "agentic ai"},
    "llm": {"llm", "transformer", "transformers", "neural networks", "genai"},
    "gpt-4": {"llm", "transformer", "genai"},
    "openai": {"llm", "genai"},
    "azure openai": {"llm", "genai"},
    "hugging face": {"llm", "transformer", "transformers", "pytorch", "neural networks"},
    "fine-tuning": {"fine-tuning", "deep learning", "neural networks"},
    "peft": {"fine-tuning", "deep learning", "neural networks"},
    "lora": {"fine-tuning", "deep learning", "neural networks"},
    "rag": {"rag", "embeddings", "prompt engineering"},
    "langchain": {"langchain", "prompt engineering", "agentic"},
    "llamaindex": {"llamaindex", "prompt engineering", "rag"},
    "nlp": {"nlp", "natural language processing"},
    "machine learning": {"machine learning", "ml"},
    "deep learning": {"deep learning", "neural networks"},
    "github actions": {"github actions", "ci/cd", "automation"},
    "uipath": {"uipath", "rpa", "automation", "process automation"},
    "rpa": {"uipath", "rpa", "automation", "process automation"},
    "power automate": {"power automate", "automation", "rpa", "process automation"},
    "blue prism": {"blue prism", "rpa", "automation"},
}


def _expand_skills(skill_set: set[str]) -> set[str]:
    expanded = set(skill_set)
    for s in skill_set:
        if s in _SKILL_EXPANSIONS:
            expanded.update(_SKILL_EXPANSIONS[s])
    return expanded


def _skill_overlap_count(skills_str: str, role_skills) -> int:
    """How many of the candidate's skills the JD actually asks for.

    Returns a large sentinel when either side has NO parsed skills, so the zero-overlap guard
    only ever fires on a GENUINE mismatch - never on missing data.

    Both exemptions matter, and the second was learned the hard way on 2026-08-04: the
    candidate-side skills can come back empty (an AI recheck rewrote them as prose that the
    prose filter then stripped), and with an empty candidate set EVERY role scores 0, so the
    guard silently discarded all 11 ranked roles for APP-20260720-1755-30E6 and dropped a
    correct 'Cybersecurity IT Administrator PD (90%)' match down to 'CISO LinkedIn
    Announcement (40%)'. Zero overlap only means something when both sides actually parsed.
    """
    if not role_skills:
        return 999
    candidate = _expand_skills(_skill_set(skills_str))
    if not candidate:
        return 999
    return len(candidate & {str(s).lower() for s in role_skills})


def _is_scoring_role(role: dict) -> bool:
    """True for real candidate-match JDs; false for admin docs accidentally cached as JDs."""
    title = str(role.get("title", "") or "")
    return bool(title.strip()) and not _NON_JOB_ROLE_TITLE_RE.search(title)


def _dedupe_ranked_titles(scored: list[tuple[str, int, str]]) -> list[tuple[str, int]]:
    """Keep the best-scoring document per OPENING so Suggested Role 1/2/3 never repeat.

    `scored` carries (title, score, opening_id). The opening id is stamped onto every role at
    INGEST by hiring_agent.jd_identity - one opening is stored as several JD documents (Job
    Announcement + Position Description + revisions), and without grouping them a single
    opening could win all three Suggested Role slots. The surviving entry keeps its own full
    title, so the client still sees the exact JD document that scored best.
    """
    best: dict[str, tuple[str, int]] = {}
    order: list[str] = []
    for title, score, key in scored:
        key = key or _re.sub(r"\s+", " ", str(title or "").strip()).lower()
        if not key:
            continue
        if key not in best:
            order.append(key)
            best[key] = (title, score)
        elif score > best[key][1]:
            best[key] = (title, score)
    return [best[key] for key in order]


def _cell_str(row, col) -> str:
    """Cell as text; blank Excel cells read back as NaN, never the string 'nan'."""
    v = row.get(col, "")
    return "" if pd.isna(v) else str(v)


_PREF_FAMILIES = (
    (("data analyst", "business analyst", "business analytics", "reporting analyst",
      "data scientist / analyst"),
     ("data analyst", "business data analyst", "businessanalyst", "analytics")),
    (("data engineer",), ("data engineer",)),
    (("data scientist", "machine learning", "ml engineer", "ai engineer",
      "artificial intelligence", "computer vision", "forward deployed", "fde",
      "applied ai", "agentic", "agentic ai", "generative ai", "genai", "llm"),
     ("data scientist", "ai ml", "ai/ ml", "ai/", "computer vision", " cv ", "agai", "swdev ai ml", "ai")),
    (("rpa", "uipath", "automation", "process automation", "power automate", "blue prism",
      "robotic process automation"),
     ("automation", "process developer", "automation process developer", "power automate", "rpa")),
    (("software engineer", "sde", "swe", "software developer", "forward deployed", "fde",
      "full stack", "backend", "developer"),
     ("software developer", "software engineer", "swdev", "software", "developer", "backend", "full stack")),
    (("mobile", "ios", "android", "swift", "kotlin", "flutter"),
     ("mobile application", "mobile app")),
    (("cybersecurity", "cyber security", "security analyst", "ciso"),
     ("cybersecurity", "cyber security", "ciso")),
    (("cloud", "devops", "aws", "azure"),
     ("cloud", "devops", "aws")),
    (("product designer", "ux designer", "ui designer", "interaction designer",
      "visual designer"),
     ("ui ux", "ui/ux", "graphics", "design")),
    (("full stack", "fullstack", "front end", "frontend", "back end", "backend",
      "web developer", "ui/ux", "ui ux"),
     ("full stack", "web developer", "website developer", "ui ux", "ui/ux")),
    (("finance", "financial"), ("finance", "financial")),
    (("graphics", "gaming", "game developer", "3d"),
     ("graphics", "gaming", "game", "3d")),
)


def _preference_matches_title(role_pref: str, title: str) -> bool:
    """Token/phrase-safe role-family match used for the configured title boost."""
    pref = _re.sub(r"[^a-z0-9+#/]+", " ", (role_pref or "").lower()).strip()
    role = _re.sub(r"[^a-z0-9+#/]+", " ", (title or "").lower()).strip()
    if not pref or not role:
        return False

    # Domain specialization guard: If candidate specifies RPA/Automation, only match automation/RPA JDs.
    rpa_kw = ("rpa", "uipath", "automation", "process automation", "power automate", "blue prism", "robotic process automation")
    if any(k in pref for k in rpa_kw):
        return any(k in role for k in ("automation", "process developer", "rpa", "power automate"))

    # Domain specialization guard: If candidate specifies Mobile, only match mobile JDs.
    mobile_kw = ("mobile", "ios", "android", "swift", "kotlin", "flutter")
    if any(k in pref for k in mobile_kw):
        return any(k in role for k in ("mobile", "android", "ios"))

    for pref_terms, title_terms in _PREF_FAMILIES:
        if any(term in pref for term in pref_terms):
            if any(term in role for term in title_terms):
                return True
    # Exact tokens only: never let a short title token such as "it" match inside
    # unrelated text such as "with your organization".
    broad = {"data", "software", "application", "position", "job", "announcement",
             "intern", "role", "engineer", "developer", "analyst", "manager", "lead"}
    pref_tokens = set(pref.split()) - broad
    title_tokens = set(role.split()) - broad
    return bool(pref_tokens & title_tokens)


def _top_n_list(skills_str: str, role_pref: str = "", top_n: int = None, roles=None) -> list:
    """Internal: sorted list of (title, score) tuples above the minimum threshold."""
    if top_n is None:
        top_n = SCORING_TOP_N
    roles = [r for r in (roles if roles is not None else get_open_roles()) if _is_scoring_role(r)]
    candidate = _skill_set(skills_str)
    if not candidate:
        return []
    pref = (role_pref or "").lower()
    scored = []
    for role in roles:
        required = {s.lower() for s in role.get("skills", [])}
        if not required:
            continue
        score = round(100 * len(candidate & required) /
                      max(len(required), SCORING_MIN_ROLE_SKILLS))
        if pref and _preference_matches_title(pref, role["title"]):
            score = min(100, score + SCORING_TITLE_BOOST)
        scored.append((role["title"], score, role_opening_id(role)))
    ranked = sorted(_dedupe_ranked_titles(scored), key=lambda x: -x[1])
    return [(t, s) for t, s in ranked if s >= SCORING_MIN_MATCH][:top_n]


def _keyword_reason(skills_str: str, best_title: str, roles: list) -> str:
    """Build a human-readable reason for the top keyword-scored role match."""
    candidate = _skill_set(skills_str)
    if not candidate:
        return ""
    for role in roles:
        if role["title"] == best_title:
            required = {s.lower() for s in role.get("skills", [])}
            matched = sorted(candidate & required)
            if matched:
                return f"Matched {len(matched)}/{len(required)} skills: {', '.join(matched)}"
            return ""
    return ""


def _ai_role_shortlist(skills_str: str, role_pref: str, roles: list,
                       limit: int = _AI_ROLE_LIMIT) -> list:
    """Bound the LLM prompt while keeping the best candidates from every active JD.

    Asking a 4K-context local model to emit scores for 80+ roles is both too large and
    unreliable.  Deterministic overlap first narrows the live JD set; Ollama then adds
    semantic judgment only inside that evidence-based shortlist.

    Every slot goes to a DIFFERENT opening, so the model always sees real alternatives - the
    JD library holds several documents per opening, and an undeduplicated shortlist could
    spend all 12 slots on 4 openings. A catalog that is already both short enough and free of
    duplicate openings is handed through untouched, skipping the ranking work entirely.
    """
    if len(roles) <= limit and len({role_opening_id(r) for r in roles}) == len(roles):
        return roles
    candidate = _expand_skills(_skill_set(skills_str))
    pref = (role_pref or "").lower()
    ranked = []
    for position, role in enumerate(roles):
        required = {str(s).lower() for s in role.get("skills", [])}
        overlap = len(candidate & required)
        ratio = overlap / max(len(required), SCORING_MIN_ROLE_SKILLS)
        title = str(role.get("title", ""))
        pref_hit = _preference_matches_title(pref, title)
        # Intern penalty: if candidate did not request an intern role, deprioritize intern JDs
        is_intern = "intern" in title.lower() and "intern" not in pref
        intern_penalty = 0.3 if is_intern else 0.0
        # Score combines ratio (completeness) and overlap count (depth of match)
        # so large JDs with high skill counts aren't crowded out by 4-skill JDs
        score_metric = ratio * 0.5 + (min(overlap, 15) / 15.0) * 0.5 - intern_penalty
        ranked.append((pref_hit, score_metric, overlap, ratio, -position, role))
    ranked.sort(key=lambda item: (item[0], item[1], item[2], item[3], item[4]), reverse=True)
    # Ranking first, then keeping the first document seen per opening, means the best-scoring
    # document of each opening is the one that represents it.
    shortlist, seen = [], set()
    for item in ranked:
        key = role_opening_id(item[5])
        if key in seen:
            continue
        seen.add(key)
        shortlist.append(item[5])
        if len(shortlist) >= limit:
            break
    return shortlist


# ── AI-aware scoring dispatcher (Ollama brain -> keyword fallback) ────────────

def suggested_roles(skills_str: str, role_pref: str = "", roles=None,
                    resume_text: str = "") -> dict:
    """Best two role matches for a candidate, using the Ollama brain when enabled.

    Returns {'role_1', 'role_2', 'reason', 'source'} where role_1/role_2 are formatted
    'Title (NN%)' strings (empty if none clear the minimum). Falls back to the
    deterministic keyword scorer whenever Ollama is off, unreachable, or returns nothing.
    """
    roles = [r for r in (roles if roles is not None else get_open_roles()) if _is_scoring_role(r)]

    if OLLAMA_ENABLED and OLLAMA_SCORING:
        ai_roles = _ai_role_shortlist(skills_str, role_pref, roles)
        if len(ai_roles) < len(roles):
            logger.info(f"   scorer: shortlisted {len(ai_roles)}/{len(roles)} JDs "
                        "before Ollama ranking")
        ai = ai_score_roles(skills_str, role_pref, roles=ai_roles, resume_text=resume_text)
        if ai:
            # ai_score_roles only ever returns titles it was given, so the shortlist is
            # enough to look each one's opening up - no re-deriving identity from the title.
            openings = {str(r.get("title", "")).strip().lower(): role_opening_id(r)
                        for r in ai_roles}
            # Skills per JD title, for the zero-overlap guard below.
            role_skills_by_title = {str(r.get("title", "")).strip().lower():
                                    {str(s).lower() for s in (r.get("skills") or [])}
                                    for r in ai_roles}
            ranked = []
            seen_openings = set()
            for r in ai:
                if r["score"] < SCORING_MIN_MATCH:
                    continue
                title = str(r.get("title", "")).strip()
                # Same opening-level grouping as the keyword path (_dedupe_ranked_titles), so
                # an Ollama ranking can't fill all three slots with the JA/PD/v2 documents of
                # one opening either.
                key = openings.get(title.lower()) or role_opening_id(title)
                if not key or key in seen_openings or _NON_JOB_ROLE_TITLE_RE.search(title):
                    continue
                # ZERO-OVERLAP GUARD (added 2026-08-04). Ollama scores from its reading of
                # the text, so it can rate a JD highly that shares NOT ONE skill with the
                # candidate. Live case APP-20260721-0122-E743 (Brett Worker, CISSP applying
                # for CISO): 'SIN 2 AI ML ENg Job Announcement' was ranked #1 at 95% with a
                # keyword overlap of exactly 0/8, pushing 'Cybersecurity IT Admin Manager'
                # (4/4 matched, 80%) into second - and Category, which derives from Suggested
                # Role 1, then filed a security executive under AI/ML/CV.
                #
                # A confident model with zero evidence is worse than no model: drop any AI
                # pick that shares no skill at all with the candidate. Roles with a genuine
                # overlap keep the AI's ranking, which is what it is good at.
                if _skill_overlap_count(skills_str, role_skills_by_title.get(title.lower())) == 0:
                    logger.info(f"   scorer: dropped AI pick {title[:48]!r} - zero skill "
                                f"overlap with the candidate (AI said {r['score']}%)")
                    continue
                seen_openings.add(key)
                ranked.append(r)
            if ranked:
                # TOP-UP (added 2026-08-04). Ollama frequently returns a THIN ranking - one
                # or two roles clearing the threshold - and until now anything it didn't fill
                # was published as a blank Suggested Role 2/3, even when the deterministic
                # scorer had strong matches sitting right there. The old code only fell back
                # to keywords when `ranked` was COMPLETELY empty, so a partial AI answer
                # silently suppressed the rest.
                #
                # Live case APP-20260716-2052-6112 (Prerna Saluja, finance/FP&A): Ollama
                # returned exactly one role, 'Finance Intern Position Description (20%)', and
                # slots 2 and 3 shipped blank - while the keyword scorer ranks 'Business
                # Analytics Marketing Intern (52%)', 'Data Analyst PD (48%)' and 'Business
                # Data Analyst PD (45%)' for the same skills, all far better aligned with her
                # stated 'Data Scientist / Analyst' preference.
                #
                # Ollama keeps every slot it actually earned, in its own order; the keyword
                # scorer only fills what is left. Opening-level de-dupe still applies, so a
                # top-up can never repeat an opening the AI already picked.
                if len(ranked) < 3:
                    for _title, _score in _top_n_list(skills_str, role_pref,
                                                      top_n=len(roles), roles=roles):
                        if len(ranked) >= 3:
                            break
                        _key = openings.get(_title.strip().lower()) or role_opening_id(_title)
                        if not _key or _key in seen_openings:
                            continue
                        seen_openings.add(_key)
                        ranked.append({"title": _title, "score": _score, "reason": ""})
                        logger.info(f"   scorer: topped up slot {len(ranked)} with "
                                    f"{_title[:48]!r} ({_score}%) - Ollama returned only "
                                    f"{len(ai)} usable role(s)")

                # Publish the STRONGEST match first. Added 2026-08-06: Ollama returns its
                # picks in its own order and the top-up above APPENDS keyword matches after
                # them, so a strong keyword match could land in slot 3 behind weaker AI
                # picks. Five live rows shipped that way - APP-20260804-1856-2182 (Bharat
                # Gupta) read 20% / 20% / 100%, putting his best match, 'Mobile Application
                # Lead Developer (100%)', last on the client's sheet while slot 1 showed a
                # 20% AWS role; APP-20260602-0018-52C6 read 80/30/87, and 6112, 0C84 and
                # 55FC were out of order too.
                #
                # The sort is STABLE, so equal scores keep the order they were ranked in -
                # Ollama's own preference still wins a tie against a keyword top-up, which
                # is the whole point of the top-up keeping "every slot it actually earned".
                # _top_n_list (the keyword-only fallback below) already sorts descending, so
                # this is the only path that could publish out of order.
                ranked.sort(key=lambda r: -int(r.get("score", 0) or 0))

                r1 = ranked[0]
                r2 = ranked[1] if len(ranked) >= 2 else None
                r3 = ranked[2] if len(ranked) >= 3 else None
                return {
                    "role_1": f"{role_display_name(r1)} ({r1['score']}%)",
                    "role_2": f"{role_display_name(r2)} ({r2['score']}%)" if r2 else "",
                    "role_3": f"{role_display_name(r3)} ({r3['score']}%)" if r3 else "",
                    "reason": r1.get("reason", ""),
                    "source": "ollama",
                }

    # If Ollama AI scoring was enabled, but Ollama failed or timed out:
    # Under REQUIRE_AI, strictly forbid falling back to keyword scoring!
    if getattr(_cfg, "REQUIRE_AI", False) and (OLLAMA_ENABLED and OLLAMA_SCORING):
        logger.warning("   scorer: Ollama AI role scoring failed or timed out; "
                       "keyword fallback is disabled under REQUIRE_AI.")
        return {
            "role_1": "",
            "role_2": "",
            "role_3": "",
            "reason": "AI role scoring failed or timed out",
            "source": "unavailable",
        }

    # ── keyword fallback ──
    top = _top_n_list(skills_str, role_pref, top_n=3, roles=roles)
    reason = ""
    if top:
        reason = _keyword_reason(skills_str, top[0][0], roles)
    return {
        "role_1": f"{role_display_name(top[0][0])} ({top[0][1]}%)" if len(top) >= 1 else "",
        "role_2": f"{role_display_name(top[1][0])} ({top[1][1]}%)" if len(top) >= 2 else "",
        "role_3": f"{role_display_name(top[2][0])} ({top[2][1]}%)" if len(top) >= 3 else "",
        "reason": reason,
        "source": "keyword",
    }


def _rematch_one(target, roles=None) -> int:
    """Re-score 'Suggested Roles' in a single workbook. Returns rows updated (-1 on skip).

    `roles` is the role list to match against (e.g. the active JDs); None = built-in roles.
    """
    if not target.exists():
        logger.warning(f"No sheet to re-match yet ({target.name} does not exist).")
        return -1
    try:
        df = pd.read_excel(target)
    except PermissionError:
        logger.error(f"'{target.name}' is locked (open in Excel?). Close it and retry --rematch.")
        return -1
    if df.empty:
        logger.info(f"{target.name}: no rows to re-match.")
        return 0

    role1s, role2s, role3s, cats = [], [], [], []
    for _, row in df.iterrows():
        skills = _cell_str(row, "Current Skills")
        pref = _cell_str(row, "Looking For Role")
        res = suggested_roles(skills, pref, roles=roles)
        role1s.append(res["role_1"])
        role2s.append(res["role_2"])
        role3s.append(res.get("role_3", ""))
        cats.append(assign_category(res["role_1"], skills, pref))   # never blank — "General" fallback
        logger.info(f"   {row.get('Full Name', '?')}: cat={cats[-1]} | "
                    f"{res['role_1']} | {res['role_2']} | {res.get('role_3','')} "
                    f"[{res['source']}]")
    df["Suggested Role 1"] = role1s
    df["Suggested Role 2"] = role2s
    if "Suggested Role 3" in COLUMNS:
        df["Suggested Role 3"] = role3s
    if "Category" in COLUMNS:
        df["Category"] = cats
    # reindex to the canonical column set — drops any legacy ATS Score / Score
    # Reason / Sentiment columns from older workbooks on the next re-match.
    df = df.reindex(columns=COLUMNS)
    try:
        df.to_excel(target, index=False)
    except PermissionError:
        import datetime as _dt
        fb = target.with_name(
            f"{target.stem}_rematched_{_dt.datetime.now():%Y%m%d-%H%M%S}.xlsx")
        df.to_excel(fb, index=False)
        logger.warning(f"'{target.name}' is locked (open in Excel?) - "
                       f"wrote re-matched rows to '{fb.name}' instead.")
        return len(df)
    logger.info(f"Re-matched {len(df)} candidate(s) in {target.name}.")
    return len(df)


def rematch_sheet(excel_file=None, roles=None) -> None:
    """Re-score 'Suggested Roles' for every candidate.

    With an explicit/overridden workbook, re-matches just that one; otherwise iterates
    every monthly workbook under OUTPUT_DIR. `roles` is the role list to match against
    (callers pass the active JDs so Re-Score All honors the JD Manager); None = built-in.
    """
    if excel_file is not None:
        targets = [excel_file]
    elif _cfg.excel_override():
        targets = [_cfg.excel_override()]
    else:
        targets = _cfg.all_excel_files()

    if not targets:
        logger.warning("No candidate workbooks found to re-match yet.")
        return
    total = sum(max(_rematch_one(t, roles), 0) for t in targets)
    logger.info(f"Re-match complete: {total} candidate(s) across {len(targets)} workbook(s).")
