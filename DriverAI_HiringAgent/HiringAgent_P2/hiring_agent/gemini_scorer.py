"""Gemini scoring & extraction brain: high-speed, intelligent AI recruitment assistant.

Uses Google Gemini API (gemini-2.5-flash / gemini-2.0-flash / gemini-1.5-flash) to:
1. Parse resumes (name, phone, location, skills, education, target role) with near-zero latency.
2. Score candidates against 105+ open roles with deep contextual reasoning.

Falls back smoothly if offline or unconfigured.
"""

import json
import logging
from typing import Any

from hiring_agent.config import (
    GEMINI_API_KEY, GEMINI_MODEL, GEMINI_ENABLED, AI_TEXT_LIMIT, logger,
)

_GEMINI_BASE_URL = "https://generativelanguage.googleapis.com/v1beta/models"
# Key travels in a header, never the URL: requests echoes the URL into every
# exception message, which then lands in logs and HTTP error bodies.
_HEADERS = {"x-goog-api-key": GEMINI_API_KEY}

_SCORE_SYSTEM = (
    "You are a precise technical recruiter. You rate how well a candidate fits each open "
    "role on a 0-100 scale, where 100 = an excellent match and 0 = no relevant overlap. "
    "Base the score on overlap between the candidate's skills/experience and each role's "
    "required skills, plus the role title the candidate says they want. Be strict and "
    "consistent. Respond ONLY with a JSON object of the form:\n"
    '{"scores": [{"title": "<role title>", "score": <0-100 integer>, "reason": "<max 12 words>"}]}'
)

_EXTRACTION_SYSTEM = (
    "You are an expert recruitment parser. Extract candidate details from the provided resume text. "
    "Return ONLY a valid JSON object with the following fields:\n"
    "- full_name: candidate full name\n"
    "- phone: phone number if present\n"
    "- location: city, state or region\n"
    "- country: country (e.g. USA, Canada, India)\n"
    "- skills: comma-separated list of technical and professional skills\n"
    "- looking_for_role: role or title candidate is applying for or best suited for\n"
    "- education: degree and university/college if present\n"
)


def gemini_health() -> tuple[bool, str]:
    """Verify Gemini API connectivity and credentials."""
    if not GEMINI_ENABLED or not GEMINI_API_KEY:
        return False, "Gemini API key not configured"
    try:
        import requests
        url = f"{_GEMINI_BASE_URL}/{GEMINI_MODEL}"
        resp = requests.get(url, timeout=5, headers=_HEADERS)
        if resp.status_code == 200:
            return True, f"Gemini Cloud AI: {GEMINI_MODEL} active"
        return False, f"Gemini API returned status {resp.status_code}: {resp.text[:100]}"
    except Exception as e:
        return False, f"Gemini API unreachable ({e})"


def _call_gemini_json(prompt: str, system_instruction: str = "", timeout: int = 30) -> dict | None:
    """Helper to send a JSON generation request to Gemini."""
    if not GEMINI_ENABLED or not GEMINI_API_KEY:
        return None
    try:
        import requests
        url = f"{_GEMINI_BASE_URL}/{GEMINI_MODEL}:generateContent"
        payload: dict[str, Any] = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {
                "responseMimeType": "application/json",
                "temperature": 0.1,
            },
        }
        if system_instruction:
            payload["systemInstruction"] = {
                "parts": [{"text": system_instruction}]
            }

        # ponytail: one retry on 5xx/timeout (503s seen 2 of 5 calls, 2026-09-14);
        # add backoff if the retry itself starts failing.
        for attempt in range(2):
            try:
                resp = requests.post(url, json=payload, timeout=timeout, headers=_HEADERS)
                if resp.status_code >= 500 and attempt == 0:
                    logger.warning("Gemini returned %s; retrying once.", resp.status_code)
                    continue
                resp.raise_for_status()
                break
            except requests.Timeout:
                if attempt == 1:
                    raise
                logger.warning("Gemini timed out after %ss; retrying once.", timeout)
        data = resp.json()
        candidates = data.get("candidates", [])
        if not candidates:
            return None
        parts = candidates[0].get("content", {}).get("parts", [])
        if not parts:
            return None
        raw_json = parts[0].get("text", "").strip()
        return json.loads(raw_json)
    except Exception as e:
        logger.warning("Gemini API call failed: %s", e)
        return None


def gemini_extract_candidate(text: str, hints: dict | None = None) -> dict | None:
    """Extract structured candidate information from resume text using Gemini."""
    text = (text or "").strip()
    if not text or not GEMINI_ENABLED:
        return None

    hint_block = ""
    if hints:
        lines = "\n".join(f"- {k}: {v}" for k, v in hints.items() if v)
        if lines:
            hint_block = (
                "\nPattern-matched hints from parser (verify or correct):\n" + lines
            )

    prompt = (
        f"{hint_block}\n\n"
        f"Candidate Resume Text:\n"
        f"{text[:AI_TEXT_LIMIT]}\n"
    )

    data = _call_gemini_json(prompt, system_instruction=_EXTRACTION_SYSTEM, timeout=30)
    if not isinstance(data, dict):
        return None

    # Standardize skills to comma-separated string if returned as list
    if isinstance(data.get("skills"), list):
        data["skills"] = ", ".join(str(s) for s in data["skills"])

    return data


def gemini_score_roles(skills: str, role_pref: str = "", roles: list | None = None,
                       resume_text: str = "") -> list | None:
    """Score candidate against open job descriptions using Gemini."""
    if not roles or not GEMINI_ENABLED:
        return None

    lines = []
    for r in roles:
        title = r.get("title", "?")
        req_skills = ", ".join(r.get("skills", [])[:25])
        lines.append(f"- {title} :: {req_skills}")
    roles_block = "\n".join(lines)

    candidate_block = (
        f"Candidate skills: {skills or 'unknown'}\n"
        f"Target role preference: {role_pref or 'unspecified'}\n"
    )
    if resume_text.strip():
        candidate_block += f"\nResume excerpt:\n{resume_text.strip()[:AI_TEXT_LIMIT]}\n"

    prompt = (
        f"{candidate_block}\n"
        f"Open roles (title :: required skills):\n{roles_block}\n\n"
        f"Rate every open role above for this candidate."
    )

    data = _call_gemini_json(prompt, system_instruction=_SCORE_SYSTEM, timeout=45)
    if isinstance(data, list):
        scores = data
    elif isinstance(data, dict):
        scores = data.get("scores", [])
    else:
        return None

    if not isinstance(scores, list):
        return None

    valid = []
    for s in scores:
        if isinstance(s, dict) and "title" in s and "score" in s:
            try:
                sc = int(s["score"])
                valid.append({
                    "title": str(s["title"]),
                    "score": max(0, min(100, sc)),
                    "reason": str(s.get("reason", ""))[:120],
                })
            except (ValueError, TypeError):
                continue

    valid.sort(key=lambda x: x["score"], reverse=True)
    return valid or None
