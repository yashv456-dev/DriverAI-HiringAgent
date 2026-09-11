"""Ollama scoring brain: rate a candidate against the open roles using a local LLM.

This is the optional AI scorer. Given a candidate's skills (and optionally the full
resume text) plus the list of open roles, it asks Ollama to rate each role 0-100 with a
one-line reason. Any failure (Ollama off/not running, model missing, bad JSON) returns
None so callers fall back to the deterministic keyword scorer in scoring.py.
"""

import json

from hiring_agent.config import (
    OLLAMA_ENABLED, OLLAMA_SCORING, OLLAMA_MODEL, OLLAMA_HOST, OLLAMA_SCORING_TIMEOUT,
    AI_TEXT_LIMIT, logger,
)

_SCORE_SYSTEM = (
    "You are a precise technical recruiter. You rate how well a candidate fits each open "
    "role on a 0-100 scale, where 100 = an excellent match and 0 = no relevant overlap. "
    "Base the score on overlap between the candidate's skills/experience and each role's "
    "required skills, plus the role title the candidate says they want. Be strict and "
    "consistent. Respond ONLY with a JSON object of the form "
    '{"scores": [{"title": "<role title>", "score": <0-100 integer>, '
    '"reason": "<max 12 words>"}]} '
    "with one entry per role given, and nothing else."
)


def ollama_health() -> tuple[bool, str]:
    """Best-effort check that the Ollama brain is reachable and the model is present.

    Returns (ok, detail). Never raises — used only to log which scorer a run will use.
    """
    if not (OLLAMA_ENABLED and OLLAMA_SCORING):
        return False, "Ollama scoring disabled - using keyword scorer"
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
    """Ask Ollama to rate the candidate against each role.

    Returns a list of {"title", "score", "reason"} sorted high-to-low, or None to fall back.
    """
    if not roles or not (OLLAMA_ENABLED and OLLAMA_SCORING):
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
