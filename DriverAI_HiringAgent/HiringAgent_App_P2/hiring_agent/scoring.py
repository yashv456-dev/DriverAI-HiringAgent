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
from hiring_agent.ollama_scorer import ai_score_roles


_AI_ROLE_LIMIT = 12
_NON_JOB_ROLE_TITLE_RE = _re.compile(r"\b(?:offer\s+letter|offerletter|onboarding|benefits?)\b", _re.I)


def get_open_roles() -> list[dict]:
    """Default open roles from config.yaml."""
    return list(DEFAULT_ROLES)


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


def assign_category(role_str: str, skills_str: str = "") -> str:
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
    Description v2 (80%)' — the latter landing in 'Senior & Executive' via the 'lead'
    keyword, section 1 precedence, same as the intentional 'Logistics Manager' ->
    'Senior & Executive' behavior tested elsewhere). That precedence is correct when the
    JD title genuinely reflects the candidate's level/domain, but a candidate's own tool
    stack is more reliable than a loosely-matched title, so a clear, uncontested mobile
    signal in Skills is trusted over the title match. Optional and defaults to empty so
    any caller without skills in scope keeps working unchanged.
    """
    skills = {s.strip().lower() for s in _re.split(r"[,;]", str(skills_str or "")) if s.strip()}
    if skills & _MOBILE_TOOL_SKILLS and not (skills & _GRAPHICS_TOOL_SKILLS):
        return "Mobile Apps (Android IOS)"

    title = _re.sub(r'\s*\(\d+%\)\s*$', '', (role_str or "")).strip().lower()
    for rule in ROLE_CATEGORY_RULES:
        if rule.get("match", "").lower() in title:
            return rule["category"]
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
    return out


def _is_scoring_role(role: dict) -> bool:
    """True for real candidate-match JDs; false for admin docs accidentally cached as JDs."""
    title = str(role.get("title", "") or "")
    return bool(title.strip()) and not _NON_JOB_ROLE_TITLE_RE.search(title)


def _dedupe_ranked_titles(scored: list[tuple[str, int]]) -> list[tuple[str, int]]:
    """Keep the best score per visible title so Suggested Role 1/2/3 never repeat."""
    best: dict[str, tuple[str, int]] = {}
    order: list[str] = []
    for title, score in scored:
        key = _re.sub(r"\s+", " ", str(title or "").strip()).lower()
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
      "artificial intelligence", "computer vision"),
     ("data scientist", "ai ml", "ai/ ml", "ai/", "computer vision", " cv ")),
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
    for pref_terms, title_terms in _PREF_FAMILIES:
        if any(term in pref for term in pref_terms):
            return any(term in role for term in title_terms)
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
        scored.append((role["title"], score))
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
    """
    if len(roles) <= limit:
        return roles
    candidate = _skill_set(skills_str)
    pref = (role_pref or "").lower()
    ranked = []
    for position, role in enumerate(roles):
        required = {str(s).lower() for s in role.get("skills", [])}
        overlap = len(candidate & required)
        ratio = overlap / max(len(required), SCORING_MIN_ROLE_SKILLS)
        title = str(role.get("title", ""))
        pref_hit = _preference_matches_title(pref, title)
        ranked.append((pref_hit, ratio, overlap, -position, role))
    ranked.sort(key=lambda item: item[:4], reverse=True)
    return [item[4] for item in ranked[:limit]]


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
            ranked = []
            seen_titles = set()
            for r in ai:
                if r["score"] < SCORING_MIN_MATCH:
                    continue
                key = _re.sub(r"\s+", " ", str(r.get("title", "")).strip()).lower()
                if not key or key in seen_titles or _NON_JOB_ROLE_TITLE_RE.search(key):
                    continue
                seen_titles.add(key)
                ranked.append(r)
            if ranked:
                r1 = ranked[0]
                r2 = ranked[1] if len(ranked) >= 2 else None
                r3 = ranked[2] if len(ranked) >= 3 else None
                return {
                    "role_1": f"{r1['title']} ({r1['score']}%)",
                    "role_2": f"{r2['title']} ({r2['score']}%)" if r2 else "",
                    "role_3": f"{r3['title']} ({r3['score']}%)" if r3 else "",
                    "reason": r1.get("reason", ""),
                    "source": "ollama",
                }

    # ── keyword fallback ──
    top = _top_n_list(skills_str, role_pref, top_n=3, roles=roles)
    reason = ""
    if top:
        reason = _keyword_reason(skills_str, top[0][0], roles)
    return {
        "role_1": f"{top[0][0]} ({top[0][1]}%)" if len(top) >= 1 else "",
        "role_2": f"{top[1][0]} ({top[1][1]}%)" if len(top) >= 2 else "",
        "role_3": f"{top[2][0]} ({top[2][1]}%)" if len(top) >= 3 else "",
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
        cats.append(assign_category(res["role_1"], skills))   # never blank — "General" fallback
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
