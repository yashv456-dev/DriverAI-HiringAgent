"""Local resume scoring: score files without SharePoint, match detail breakdowns."""

from pathlib import Path

import pandas as pd

from hiring_agent.config import SKILL_DISPLAY, logger
import hiring_agent.config as _cfg
from hiring_agent.extraction import extract_text_from_bytes, extract_candidate_details_smart
from hiring_agent.scoring import _skill_set
from hiring_agent.jd_sources import get_active_roles


def get_match_details(skills_str: str, role: dict) -> dict:
    """Skill-by-skill breakdown of a candidate vs. one role."""
    candidate = _skill_set(skills_str)
    required = {s.lower() for s in role.get("skills", [])}
    if not required:
        return {"role_title": role.get("title", "?"), "score": 0,
                "matched": [], "missing": [], "extra": sorted(candidate)}
    matched = sorted(candidate & required)
    missing = sorted(required - candidate)
    extra = sorted(candidate - required)
    score = round(100 * len(matched) / len(required))

    def display(s):
        return SKILL_DISPLAY.get(s, s.title())

    return {
        "role_title": role.get("title", "?"),
        "score": score,
        "matched": [display(s) for s in matched],
        "missing": [display(s) for s in missing],
        "extra": [display(s) for s in extra],
    }


def score_resume_file(file_path, roles=None) -> dict:
    """Score a single PDF/DOCX file against active roles. Returns details per role."""
    path = Path(file_path)
    if not path.exists():
        return {"file": str(path), "candidate": {}, "role_matches": [], "error": "File not found"}

    raw = path.read_bytes()
    text = extract_text_from_bytes(raw, path.name)
    if not text.strip():
        return {"file": path.name, "candidate": {}, "role_matches": [],
                "error": "Could not extract text"}

    from hiring_agent.extraction import EXTRACTION_SOURCE
    candidate = extract_candidate_details_smart(text)
    if getattr(_cfg, "REQUIRE_AI", False) and EXTRACTION_SOURCE.get("value") not in ("ollama", "gemini"):
        return {"file": path.name, "candidate": {}, "role_matches": [],
                "error": f"AI extraction unavailable ({EXTRACTION_SOURCE.get('value')}); regex fallback disabled under REQUIRE_AI"}

    if roles is None:
        roles = get_active_roles()

    details = [get_match_details(candidate.get("skills", ""), role) for role in roles]
    details.sort(key=lambda d: -d["score"])

    return {"file": path.name, "candidate": candidate, "role_matches": details}


def load_candidates(excel_file=None) -> tuple[list[dict], list[str]]:
    """Load candidates as a list of dicts.

    With an explicit/overridden workbook, reads just that one. Otherwise aggregates
    every monthly workbook under OUTPUT_DIR (all years/months) and de-dupes by
    Application ID (keeping the latest), so the dashboard shows everyone.

    Returns (records, errors) where errors is a list of human-readable strings for
    any workbooks that could not be read (locked, corrupted, etc.).
    """
    if excel_file:
        targets = [Path(excel_file)]
    elif _cfg.excel_override():
        targets = [_cfg.excel_override()]
    else:
        targets = _cfg.all_excel_files()

    frames = []
    errors: list[str] = []
    for target in targets:
        if not target.exists():
            continue
        try:
            frames.append(pd.read_excel(target))
        except Exception as e:
            msg = f"{target.name}: {e}"
            logger.warning(f"Could not load candidates from {msg}")
            errors.append(msg)
    if not frames:
        return [], errors

    df = pd.concat(frames, ignore_index=True)
    if "Application ID" in df.columns:
        df = df.drop_duplicates(subset="Application ID", keep="last")
    return df.fillna("").to_dict("records"), errors
