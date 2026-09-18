"""Ollama scoring brain: rate a candidate against the open roles using a local LLM.

This is the optional AI scorer. Given a candidate's skills (and optionally the full
resume text) plus the list of open roles, it asks Ollama to rate each role 0-100 with a
one-line reason. Any failure (Ollama off/not running, model missing, bad JSON) returns
None so callers fall back to the deterministic keyword scorer in scoring.py.
"""

import json

from hiring_agent.config import (
    OLLAMA_ENABLED, OLLAMA_SCORING, OLLAMA_MODEL, OLLAMA_HOST, OLLAMA_SCORING_TIMEOUT,
    GEMINI_ENABLED, GEMINI_MODEL,
    AI_TEXT_LIMIT, logger,
)
# ponytail: same recruiter prompt for both brains, one definition.
from hiring_agent.gemini_scorer import _SCORE_SYSTEM


def ollama_health() -> tuple[bool, str]:
    """Best-effort check that an AI brain (Gemini or Ollama) is reachable.

    Returns (ok, detail). Never raises — used only to log which scorer a run will use.
    """
    if GEMINI_ENABLED:
        from hiring_agent.gemini_scorer import gemini_health
        ok, msg = gemini_health()
        if ok:
            return True, msg
    if not (OLLAMA_ENABLED and OLLAMA_SCORING):
        return False, "AI scoring disabled - using keyword scorer"
    try:
        import requests
        resp = requests.get(f"{OLLAMA_HOST}/api/tags", timeout=5)
        resp.raise_for_status()
        names = [m.get("name", "") for m in resp.json().get("models", [])]
        have = any(n == OLLAMA_MODEL or n.split(":")[0] == OLLAMA_MODEL for n in names)
        if have:
            return True, f"Ollama brain: {OLLAMA_MODEL} reachable at {OLLAMA_HOST}"
        return False, (f"Ollama reachable but model '{OLLAMA_MODEL}' not pulled "
                       f"(have: {', '.join(names) or 'none'}) - using keyword scorer")
    except Exception as e:
        return False, f"Ollama not reachable ({e}) - using keyword scorer"


def _roles_brief(roles: list) -> str:
    """Compact 'Title :: skill, skill, ...' lines for the prompt."""
    lines = []
    for r in roles:
        title = r.get("title", "?")
        skills = ", ".join(r.get("skills", [])[:25])
        lines.append(f"- {title} :: {skills}")
    return "\n".join(lines)


def ai_score_roles(skills: str, role_pref: str = "", roles=None,
                   resume_text: str = "") -> list | None:
    """Ask an AI brain (Gemini or Ollama) to rate the candidate against each role.

    Returns a list of {"title", "score", "reason"} sorted high-to-low, or None to fall back.
    """
    if not roles:
        return None

    if GEMINI_ENABLED:
        try:
            from hiring_agent.gemini_scorer import gemini_score_roles
            res = gemini_score_roles(skills=skills, role_pref=role_pref, roles=roles, resume_text=resume_text)
            if res:
                logger.info(f"   scorer: Gemini Cloud AI ({GEMINI_MODEL}) rated {len(res)} role(s)")
                return res
        except Exception as e:
            logger.warning("Gemini role scoring error (%s); falling back to local scorer.", e)

    if not (OLLAMA_ENABLED and OLLAMA_SCORING):
        return None

    candidate_block = (
        f"Candidate skills: {skills or 'unknown'}\n"
        f"Role the candidate wants: {role_pref or 'unspecified'}\n"
    )
    if resume_text.strip():
        candidate_block += f"\nResume excerpt:\n{resume_text.strip()[:AI_TEXT_LIMIT]}\n"

    user = (
        f"{candidate_block}\n"
        f"Open roles (title :: required skills):\n{_roles_brief(roles)}\n\n"
        f"Rate every open role above for this candidate."
    )

    try:
        import requests
        resp = requests.post(
            f"{OLLAMA_HOST}/api/chat",
            json={
                "model": OLLAMA_MODEL,
                "format": "json",
                "think": False,
                "stream": False,
                "options": {"temperature": 0, "num_predict": 1536},
                "messages": [
                    {"role": "system", "content": _SCORE_SYSTEM},
                    {"role": "user", "content": user},
                ],
            },
            timeout=OLLAMA_SCORING_TIMEOUT,
        )
        resp.raise_for_status()
        content = resp.json()["message"]["content"]
        try:
            data = json.loads(content)
        except Exception:
            import re
            items = re.findall(r'\{[^{}]*?"title"\s*:\s*"([^"]+)"[^{}]*?"score"\s*:\s*(\d+)[^{}]*?\}', content)
            if items:
                data = {"scores": [{"title": t, "score": int(s), "reason": ""} for t, s in items]}
            else:
                raise
    except Exception as e:
        logger.warning(f"Ollama scoring unavailable ({e}); using keyword scorer.")
        return None

    raw = data if isinstance(data, list) else data.get("scores", [])
    valid_titles = {r.get("title", "").strip().lower() for r in roles}
    out = []
    for item in raw if isinstance(raw, list) else []:
        if not isinstance(item, dict):
            continue
        title = str(item.get("title", "")).strip()
        if not title:
            continue
        # keep only roles we actually asked about (LLMs sometimes invent extras)
        if title.lower() not in valid_titles:
            match = next((r["title"] for r in roles
                          if title.lower() in r.get("title", "").lower()
                          or r.get("title", "").lower() in title.lower()), None)
            if not match:
                continue
            title = match
        try:
            score = int(round(float(item.get("score", 0))))
        except (TypeError, ValueError):
            continue
        score = max(0, min(100, score))
        reason = str(item.get("reason", "")).strip()[:120]
        out.append({"title": title, "score": score, "reason": reason})

    if not out:
        logger.warning("Ollama scoring returned no usable rows; using keyword scorer.")
        return None

    out.sort(key=lambda d: -d["score"])
    logger.info(f"   scorer: Ollama ({OLLAMA_MODEL}) rated {len(out)} role(s)")
    return out
