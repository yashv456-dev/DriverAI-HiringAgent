"""DriverAI Hiring Agent — High Performance Recruiter Web Server & API.

Provides:
- Real-time resume extraction & intelligent 105-role matching via Gemini AI
- Searchable candidate pipeline and SQLite persistence
- Open roles explorer & skill breakdown
- Gemini AI Recruiter Copilot chat
- Modern glassmorphic recruiter dashboard
"""

import datetime
import io
import json
import logging
import re
import os
import sqlite3
import traceback
import uuid
from pathlib import Path
from typing import Any, Optional

import uvicorn
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

# Import hiring agent modules
from hiring_agent.config import (
    BASE_DIR, GEMINI_API_KEY, GEMINI_MODEL, GEMINI_ENABLED,
    JD_CACHE_FILE, DEFAULT_ROLES, GEO_FILTER_USA_ONLY, logger,
    SHAREPOINT_CONFIGURED,   # ADDED 2026-09-18: needed by the mode switch
    SUPPRESS_EMAILS, ADMIN_EMAIL,   # ADDED 2026-09-18: real diagnostics values
)
from hiring_agent.extraction import (
    extract_text_from_bytes, extract_candidate_details_smart, EXTRACTION_SOURCE,
)
from hiring_agent.gemini_scorer import gemini_health, _call_gemini_json
from hiring_agent.jd_sources import get_active_roles
from hiring_agent.local_scorer import get_match_details
from hiring_agent.ollama_scorer import ai_score_roles
from hiring_agent.sqlite_store import SQLiteCandidateStore, database_path

app = FastAPI(title="DriverAI Hiring Agent", version="2.0.0")

# Enable CORS for local development flexibility
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

STATIC_DIR = BASE_DIR / "static"
STATIC_DIR.mkdir(parents=True, exist_ok=True)


# ── SQLite Database Helper ───────────────────────────────────────────────────

def get_store() -> SQLiteCandidateStore:
    return SQLiteCandidateStore()


def init_db():
    store = get_store()
    return store


# ── Pydantic Request Models ──────────────────────────────────────────────────

class ChatRequest(BaseModel):
    message: str
    context: Optional[dict] = None
    history: Optional[list[dict]] = None


class ManualCandidateRequest(BaseModel):
    full_name: str
    email: Optional[str] = ""
    phone: Optional[str] = ""
    location: Optional[str] = ""
    country: Optional[str] = "USA"
    skills: Optional[str] = ""
    looking_for_role: Optional[str] = ""
    education: Optional[str] = ""
    experience: Optional[str] = ""
    notes: Optional[str] = ""


# ── Health & System Diagnostics ──────────────────────────────────────────────

# ── Data-source mode (added 2026-09-18) ──────────────────────────────────────
# "testing"   -> the local SQLite store (sandbox or otherwise). No Graph calls.
# "developer" -> the live SharePoint master, read through the existing
#                SharePointClient. READ-ONLY: nothing in this file writes to Graph.
# Process-local on purpose. It is a view toggle for one operator's browser, not
# durable state, so it deliberately does not survive a restart.
_MODE = {"value": "testing"}


def current_mode() -> str:
    return _MODE["value"]


# The probe below costs a Graph round trip, and /api/health runs on every page
# load. Uncached it made health take ~16s and left the header badges stuck on
# "checking...". 60s is short enough that a tenant going down is noticed quickly
# and long enough that opening a few views does not re-probe each time.
_MS_CACHE: dict = {"at": 0.0, "value": None}
_MS_TTL_SECONDS = 60


def ms_status(force: bool = False) -> dict:
    """Live Microsoft reachability -- an actual Graph call, not a config flag.

    SHAREPOINT_CONFIGURED only proves the env vars are non-empty; an expired
    secret or a revoked grant still reads 'configured'. The only honest answer
    comes from a request that either returns or does not.

    Cached for _MS_TTL_SECONDS; pass force=True where the answer must be current
    (the mode switch, which refuses to move to Developer on a dead tenant).
    """
    import time
    if not force and _MS_CACHE["value"] is not None \
            and (time.time() - _MS_CACHE["at"]) < _MS_TTL_SECONDS:
        return _MS_CACHE["value"]

    if not SHAREPOINT_CONFIGURED:
        result = {"configured": False, "reachable": False,
                  "detail": "No Graph credentials in this environment (local mode only)."}
    else:
        try:
            from sharepoint_client import SharePointClient
            cols = SharePointClient().table_columns()
            result = {"configured": True, "reachable": True,
                      "detail": f"SharePoint table reachable ({len(cols)} columns)."}
        except Exception as e:
            result = {"configured": True, "reachable": False,
                      "detail": f"{type(e).__name__}: {str(e)[:160]}"}

    _MS_CACHE["at"], _MS_CACHE["value"] = time.time(), result
    return result



# ADDED 2026-09-18: get_active_roles() re-checks the SharePoint JD-folder
# fingerprint on EVERY call -- measured at 10.3s each time, with no in-process
# cache. /api/health and /api/stats both call it, so opening the dashboard cost
# ~20s and the header badges sat on "checking...". The JD catalog changes when
# someone edits the SharePoint folder, not between two clicks, so a 5-minute
# cache here is ample. jd_sources is left alone: the scoring pipeline wants its
# freshness check, this is only the dashboard's read path.
_ROLES_CACHE: dict = {"at": 0.0, "value": None}
_ROLES_TTL_SECONDS = 300


def cached_roles() -> list:
    import time
    if _ROLES_CACHE["value"] is not None \
            and (time.time() - _ROLES_CACHE["at"]) < _ROLES_TTL_SECONDS:
        return _ROLES_CACHE["value"]
    roles = get_active_roles()   # the real, uncached fetch -- NOT cached_roles()
    _ROLES_CACHE["at"], _ROLES_CACHE["value"] = time.time(), roles
    return roles


def load_candidate_rows() -> list[dict]:
    """Every candidate row from whichever source the mode selects.

    ADDED 2026-09-18. This exists because the mode switch was first wired into
    /api/candidates alone, which made the Candidate Pipeline show the live
    SharePoint rows while the Dashboard KPIs still counted the local store --
    the same screen reporting two different totals. Every endpoint that lists
    candidates calls this instead of opening SQLite itself.

    Shape is the local store's, so the existing parsing code works unchanged:
    {app_id, status, sheet, values_json, created_at, updated_at}.
    """
    if current_mode() == "developer":
        from sharepoint_client import SharePointClient
        return [{
            "app_id": r["values"].get("Application ID", ""),
            "status": r["values"].get("Status", ""),
            "sheet": "CandidateList",
            "values_json": json.dumps(r["values"]),
            "created_at": r["values"].get("Received Date", ""),
            "updated_at": r["values"].get("Last Updated Date", ""),
        } for r in SharePointClient().list_rows()]

    db_path = database_path()
    if not db_path.exists():
        return []
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    rows = conn.execute(
        "SELECT app_id, status, sheet, values_json, created_at, updated_at "
        "FROM candidates WHERE active = 1 ORDER BY id DESC"
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


# ADDED 2026-09-18: the pipeline has no numeric score column. suggested_roles()
# folds the percentage into the role STRING instead -- scoring.py:493 builds
# f"{role_display_name(r1)} ({r1['score']}%)" -- so the master stores e.g.
# "Senior ML Engineer (82%)". Without this, Developer mode reported 0% for every
# candidate because it looked for a "Match Score" field only the dashboard writes.
_ROLE_SCORE_RE = re.compile(r"\s*\((\d{1,3})\s*%\)\s*$")


def split_role_score(text: str) -> tuple[str, int]:
    """('Senior ML Engineer (82%)') -> ('Senior ML Engineer', 82).

    Returns (text_unchanged, 0) when no trailing percentage is present, which is
    the case for rows the pipeline deferred rather than scored.
    """
    s = str(text or "").strip()
    m = _ROLE_SCORE_RE.search(s)
    if not m:
        return s, 0
    score = int(m.group(1))
    if score > 100:
        # Not a percentage this pipeline produced. Leave the text alone rather than
        # silently editing a value we do not understand.
        return s, 0
    return s[:m.start()].strip(), score


class ModeRequest(BaseModel):
    mode: str


@app.get("/api/mode")
def api_get_mode():
    return {"mode": current_mode(), "microsoft": ms_status()}


@app.post("/api/mode")
def api_set_mode(req: ModeRequest):
    mode = (req.mode or "").strip().lower()
    if mode not in ("testing", "developer"):
        raise HTTPException(status_code=400, detail="mode must be 'testing' or 'developer'")
    ms = ms_status(force=True)   # never gate a switch on a cached probe
    # Refuse to pretend. Switching to developer with no reachable tenant would
    # show an empty table that looks like "no candidates" rather than "no cloud".
    if mode == "developer" and not ms["reachable"]:
        raise HTTPException(status_code=409,
                            detail=f"Cannot switch to Developer: {ms['detail']}")
    _MODE["value"] = mode
    return {"mode": mode, "microsoft": ms}


@app.get("/api/health")
def api_health():
    ai_ok, ai_msg = gemini_health()
    roles = cached_roles()
    db_path = database_path()
    
    candidate_count = 0
    if db_path.exists():
        try:
            conn = sqlite3.connect(str(db_path))
            c = conn.cursor()
            c.execute("SELECT count(*) FROM candidates WHERE active = 1")
            candidate_count = c.fetchone()[0]
            conn.close()
        except Exception:
            pass

    return {
        "status": "online",
        "ai_engine": {
            "enabled": GEMINI_ENABLED,
            "model": GEMINI_MODEL,
            "status": ai_msg,
            "connected": ai_ok,
        },
        "roles_loaded": len(roles),
        "total_candidates": candidate_count,
        "geo_filter": "USA-only" if GEO_FILTER_USA_ONLY else "Off",
        # ADDED 2026-09-18: mode + a real Microsoft reachability probe, so the
        # header badge reflects a Graph call instead of a hardcoded string.
        "mode": current_mode(),
        "microsoft": ms_status(),
        # ADDED 2026-09-18: real values behind the System Diagnostics panel, which
        # used to be hardcoded HTML. The mail flag especially: it read "Suppression
        # Active (Safe Mode)" unconditionally, so it would have shown green while
        # actually sending.
        "database": {
            "path": str(db_path),
            "name": db_path.name,
            "exists": db_path.exists(),
        },
        "mail": {
            "suppressed": SUPPRESS_EMAILS,
            "admin_email": ADMIN_EMAIL,
        },
        "timestamp": datetime.datetime.utcnow().isoformat(),
    }


# ── Dashboard Statistics ────────────────────────────────────────────────────

@app.get("/api/stats")
def api_stats():
    roles = cached_roles()

    total = 0
    scored = 0
    usa_count = 0
    avg_score = 0
    top_matched_roles = {}
    recent_candidates = []

    # CHANGED 2026-09-18: was hardwired to the local SQLite store, so the Dashboard
    # kept showing the sandbox count while the Candidate Pipeline showed live rows.
    try:
        rows = load_candidate_rows()

        total = len(rows)
        scores = []
        for r in rows:
            try:
                val = json.loads(r["values_json"])
                if r["status"] == "Scored":
                    scored += 1
                loc = str(val.get("Country", "")).strip().lower()
                if loc in ("usa", "us", "united states", "united states of america"):
                    usa_count += 1
                
                # CHANGED 2026-09-18: fall back to the percentage embedded in the
                # role string, so the average is not 0 for pipeline-scored rows.
                role_match, role_score = split_role_score(val.get("Suggested Role 1", ""))
                score_val = val.get("Match Score", 0) or role_score
                if score_val:
                    try:
                        scores.append(int(score_val))
                    except (ValueError, TypeError):
                        pass

                if role_match:
                    top_matched_roles[role_match] = top_matched_roles.get(role_match, 0) + 1
                
                if len(recent_candidates) < 6:
                    recent_candidates.append({
                        "app_id": r["app_id"],
                        "full_name": val.get("Full Name") or val.get("Name", "Unknown"),
                        "role_1": val.get("Suggested Role 1", "Unmatched"),
                        "score": val.get("Match Score", 0),
                        "location": val.get("Location", "Unknown"),
                        "status": r["status"],
                        "created_at": r["created_at"],
                    })
            except Exception:
                continue

        if scores:
            avg_score = round(sum(scores) / len(scores), 1)
    except Exception as e:
        logger.warning("Error computing stats from DB: %s", e)

    top_roles_list = sorted(
        [{"role": k, "count": v} for k, v in top_matched_roles.items()],
        key=lambda x: x["count"],
        reverse=True,
    )[:5]

    return {
        "total_candidates": total,
        "scored_candidates": scored,
        "usa_verified_pct": round((usa_count / total * 100) if total > 0 else 100, 1),
        "average_score": avg_score,
        "open_roles_count": len(roles),
        "top_roles": top_roles_list,
        "recent_candidates": recent_candidates,
    }


# ── Roles Catalog & Explorer ────────────────────────────────────────────────

@app.get("/api/roles")
def api_roles(search: Optional[str] = None, category: Optional[str] = None):
    roles = cached_roles()
    out = []
    
    for r in roles:
        title = r.get("title", "Untitled Role")
        skills = r.get("skills", [])
        
        # Categorize role
        t_low = title.lower()
        if any(w in t_low for w in ("ai", "ml", "learning", "data", "vision", "nlp", "llm")):
            cat = "AI & Machine Learning"
        elif any(w in t_low for w in ("system", "hardware", "rf", "embedded", "fpga", "electrical")):
            cat = "Systems & Hardware"
        elif any(w in t_low for w in ("cloud", "devops", "infra", "sre", "kubernetes", "security")):
            cat = "Cloud & Infrastructure"
        elif any(w in t_low for w in ("frontend", "react", "fullstack", "backend", "developer", "software", "python")):
            cat = "Software Engineering"
        elif any(w in t_low for w in ("product", "manager", "operations", "recruiter", "lead")):
            cat = "Product & Operations"
        else:
            cat = "Engineering & Other"

        if category and category.lower() != "all" and cat.lower() != category.lower():
            continue

        if search:
            q = search.lower()
            if q not in title.lower() and not any(q in s.lower() for s in skills):
                continue

        out.append({
            "title": title,
            "category": cat,
            "skills": skills,
            "skill_count": len(skills),
            "id": r.get("id") or hashlib_id(title),
        })

    # Group summary
    categories = sorted(list(set(item["category"] for item in out)))
    return {
        "roles": out,
        "total": len(out),
        "categories": categories,
    }


def hashlib_id(text: str) -> str:
    import hashlib
    return hashlib.md5(text.encode("utf-8")).hexdigest()[:8]


def _rank_roles(extracted: dict, roles: list, text: str) -> list[dict]:
    """AI role scores, or the keyword scorer when no brain answers.

    Always returns [{"title", "score", "reason"}] high-to-low: get_match_details()
    emits `role_title`, and both endpoints below index on `title`.
    """
    matches = ai_score_roles(
        skills=extracted.get("skills", ""),
        role_pref=extracted.get("looking_for_role", ""),
        roles=roles,
        resume_text=text,
    )
    if matches:
        return matches
    fallback = []
    for r in roles:
        d = get_match_details(extracted.get("skills", ""), r)
        fallback.append({"title": d["role_title"], "score": d["score"],
                         "reason": f"{len(d['matched'])} of {len(d['matched']) + len(d['missing'])} keywords matched."})
    fallback.sort(key=lambda d: -d["score"])
    return fallback


def _skills_list(skills) -> list[str]:
    if isinstance(skills, list):
        return skills
    return [s.strip() for s in str(skills or "").split(",") if s.strip()]


# ── Resume Upload & Real-Time Scoring ───────────────────────────────────────

@app.post("/api/score-resume")
async def api_score_resume(
    file: Optional[UploadFile] = File(None),
    resume_text: Optional[str] = Form(None),
    save_to_db: Optional[bool] = Form(True),
):
    text = ""
    filename = "pasted_text.txt"

    if file:
        filename = file.filename or "resume.pdf"
        content = await file.read()
        try:
            text = extract_text_from_bytes(content, filename)
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Could not parse file: {e}")
    elif resume_text:
        text = resume_text
    else:
        raise HTTPException(status_code=400, detail="Please upload a resume file or paste resume text.")

    if not text.strip():
        raise HTTPException(status_code=400, detail="Resume content was empty or unreadable.")

    # 1. AI Extraction (Gemini)
    extracted = extract_candidate_details_smart(text)
    extraction_source = EXTRACTION_SOURCE.get("value", "offline")

    # 2. Match against open roles
    roles = cached_roles()
    role_matches = _rank_roles(extracted, roles, text)

    top_matches = role_matches[:5]
    top_role_1 = top_matches[0]["title"] if top_matches else "Unmatched"
    top_score_1 = top_matches[0]["score"] if top_matches else 0

    top_role_2 = top_matches[1]["title"] if len(top_matches) > 1 else ""
    top_score_2 = top_matches[1]["score"] if len(top_matches) > 1 else 0

    top_role_3 = top_matches[2]["title"] if len(top_matches) > 2 else ""
    top_score_3 = top_matches[2]["score"] if len(top_matches) > 2 else 0

    # Determine status & location
    country = extracted.get("country", "").strip().upper()
    is_usa = country in ("USA", "US", "UNITED STATES", "UNITED STATES OF AMERICA") or "USA" in country
    
    if GEO_FILTER_USA_ONLY and not is_usa and country and country != "NOT EXTRACTED":
        status = "Rejected - Non-USA Location"
    elif top_score_1 >= 50:
        status = "Scored"
    else:
        status = "Needs Review"

    # Enrich matches with skill gaps
    enriched_matches = []
    cand_skills_raw = extracted.get("skills", "")
    cand_skills_set = {s.strip().lower() for s in cand_skills_raw.split(",") if s.strip()}

    for m in top_matches:
        title = m["title"]
        score = m["score"]
        reason = m.get("reason", "")
        
        # Find matching JD skills
        jd_obj = next((r for r in roles if r.get("title") == title), None)
        jd_skills = jd_obj.get("skills", []) if jd_obj else []
        
        matched_skills = [s for s in jd_skills if s.lower() in cand_skills_set or any(s.lower() in cs for cs in cand_skills_set)]
        missing_skills = [s for s in jd_skills if s not in matched_skills]

        enriched_matches.append({
            "title": title,
            "score": score,
            "reason": reason or f"{len(matched_skills)} of {len(jd_skills)} required skills matched.",
            "matched_skills": matched_skills[:8],
            "missing_skills": missing_skills[:6],
            "total_jd_skills": len(jd_skills),
        })

    # Generate unique Application ID
    app_id = f"APP-{datetime.datetime.now():%Y%m%d-%H%M%S}-{uuid.uuid4().hex[:6]}"

    record_data = {
        "Application ID": app_id,
        "Full Name": extracted.get("full_name", "Unknown Candidate"),
        "Email": extracted.get("email", ""),
        "Phone": extracted.get("phone", "Not extracted"),
        "Location": extracted.get("location", "Not extracted"),
        "Country": extracted.get("country", "USA"),
        "Skills": extracted.get("skills", ""),
        "Looking For Role": extracted.get("looking_for_role", ""),
        "Education": extracted.get("education", ""),
        "Experience": extracted.get("experience", ""),
        "Suggested Role 1": top_role_1,
        "Suggested Role 2": top_role_2,
        "Suggested Role 3": top_role_3,
        "Match Score": top_score_1,
        "Status": status,
        "Original Filename": filename,
        "Extraction Source": extraction_source,
        "Notes": extracted.get("notes", text[:400]),
    }

    # Save to SQLite DB if requested
    if save_to_db:
        try:
            store = get_store()
            store.add(record_data, sheet="main")
        except Exception as e:
            logger.warning("Could not persist candidate to SQLite: %s", e)

    return {
        "app_id": app_id,
        # app.js renders candidate.skills as chips and reads scores[].role/reasoning.
        "candidate": {**extracted, "skills": _skills_list(extracted.get("skills", ""))},
        "status": status,
        "match_score": top_score_1,
        "suggested_role_1": top_role_1,
        "suggested_role_2": top_role_2,
        "suggested_role_3": top_role_3,
        "matches": enriched_matches,
        "scores": [{"role": m["title"], "score": m["score"], "reasoning": m["reason"],
                    "matched_skills": m["matched_skills"], "missing_skills": m["missing_skills"]}
                   for m in enriched_matches],
        "extraction_source": extraction_source,
        "filename": filename,
    }


# ── Candidates Directory & CRUD ──────────────────────────────────────────────

@app.get("/api/candidates")
def api_candidates(
    search: Optional[str] = None,
    status: Optional[str] = None,
    min_score: Optional[int] = None,
    role: Optional[str] = None,
):
    # CHANGED 2026-09-18: source selection moved into load_candidate_rows() so the
    # Dashboard and this list can never disagree about which store they are reading.
    try:
        rows = load_candidate_rows()
    except Exception as e:
        raise HTTPException(status_code=502,
                            detail=f"SharePoint read failed: {type(e).__name__}: {str(e)[:160]}")
    if not rows:
        return {"candidates": [], "total": 0}

    results = []
    for r in rows:
        try:
            val = json.loads(r["values_json"])
            name = val.get("Full Name") or val.get("Name", "Unknown")
            loc = val.get("Location", "")
            # CHANGED 2026-09-18: the live master calls this column "Current Skills"
            # and carries no match score; the local store uses "Skills"/"Match Score".
            # Accept both so one code path serves either source.
            skills = val.get("Skills") or val.get("Current Skills", "")
            # CHANGED 2026-09-18: pull the percentage out of the role string when the
            # pipeline wrote one, and strip it from the displayed title so the score
            # is not shown twice. Dashboard-ingested rows keep their "Match Score".
            r1, r1_score = split_role_score(val.get("Suggested Role 1", ""))
            try:
                score = int(float(val.get("Match Score") or val.get("ATS Score") or 0))
            except (TypeError, ValueError):
                score = 0
            if not score:
                score = r1_score
            st = r["status"]

            if search:
                q = search.lower()
                if not (q in name.lower() or q in loc.lower() or q in skills.lower() or q in r1.lower() or q in r["app_id"].lower()):
                    continue

            if status and status.lower() != "all" and st.lower() != status.lower():
                continue

            if min_score is not None and score < min_score:
                continue

            if role and role.lower() != "all" and role.lower() not in r1.lower():
                continue

            results.append({
                "app_id": r["app_id"],
                "full_name": name,
                "email": val.get("Email", ""),
                "phone": val.get("Phone", ""),
                "location": loc,
                "country": val.get("Country", ""),
                "skills": skills,
                "suggested_role_1": r1,
                "suggested_role_2": split_role_score(val.get("Suggested Role 2", ""))[0],
                "suggested_role_3": split_role_score(val.get("Suggested Role 3", ""))[0],
                "match_score": score,
                "status": st,
                "education": val.get("Education", ""),
                "experience": val.get("Experience") or val.get("Years Exp", ""),  # CHANGED 2026-09-18: live master uses "Years Exp"
                "filename": val.get("Original Filename", ""),
                "created_at": r["created_at"],
            })
        except Exception:
            continue

    return {"candidates": results, "total": len(results)}


@app.get("/api/candidates/{app_id}")
def api_candidate_detail(app_id: str):
    # CHANGED 2026-09-18: this used to return {"values": {...Title Case...}} while the
    # dossier drawer reads flat snake_case (cand.email, cand.phone, cand.match_score).
    # Every field came back undefined and the drawer fell through to hardcoded
    # placeholders -- inventing a phone number, an @example.com address and an 80%
    # score, including on candidates that had been rejected. Now it returns the same
    # flat shape as /api/candidates and /api/compare, with `values` kept for any
    # caller that still wants the raw row.
    row = next((r for r in load_candidate_rows() if r["app_id"] == app_id), None)
    if not row:
        raise HTTPException(status_code=404, detail="Candidate not found")

    val = json.loads(row["values_json"])
    r1, r1_score = split_role_score(val.get("Suggested Role 1", ""))
    try:
        score = int(float(val.get("Match Score") or 0))
    except (TypeError, ValueError):
        score = 0

    return {
        "app_id": row["app_id"],
        "status": row["status"],
        "full_name": val.get("Full Name") or val.get("Name", ""),
        "email": val.get("Email", ""),
        "phone": val.get("Phone", ""),
        "location": val.get("Location", ""),
        "country": val.get("Country", ""),
        "skills": val.get("Skills") or val.get("Current Skills", ""),
        "suggested_role_1": r1,
        "suggested_role_2": split_role_score(val.get("Suggested Role 2", ""))[0],
        "suggested_role_3": split_role_score(val.get("Suggested Role 3", ""))[0],
        "match_score": score or r1_score,
        # The dossier drawer reads role_1/role_2/role_3 and score (the names
        # /api/compare uses), while the candidates table reads suggested_role_N
        # and match_score. Two vocabularies for one thing; both are served here
        # so neither caller breaks. Worth collapsing to one set later.
        "role_1": r1,
        "role_2": split_role_score(val.get("Suggested Role 2", ""))[0],
        "role_3": split_role_score(val.get("Suggested Role 3", ""))[0],
        "score": score or r1_score,
        "education": val.get("Education", ""),
        "experience": val.get("Experience") or val.get("Years Exp", ""),
        "filename": val.get("Original Filename", ""),
        "notes": val.get("Notes", ""),
        "values": val,
        "created_at": row["created_at"],
        "updated_at": row["updated_at"],
    }


# ADDED 2026-09-18: open the candidate's actual CV. The row carried a filename
# and (in Developer mode) a SharePoint link, but nothing in the UI could open
# either, so a recruiter could read extracted fields and never see the document
# they came from.
@app.get("/api/candidates/{app_id}/resume")
def api_candidate_resume(app_id: str):
    row = next((r for r in load_candidate_rows() if r["app_id"] == app_id), None)
    if not row:
        raise HTTPException(status_code=404, detail="Candidate not found")
    val = json.loads(row["values_json"])

    # Developer mode: the file lives in SharePoint. Hand back its link rather
    # than proxying the bytes -- the viewer is already authenticated there, and
    # streaming candidate CVs through this unauthenticated server would be worse.
    link = str(val.get("Resume URL") or val.get("Resume Link") or "").strip()
    if link.startswith("http"):
        return {"kind": "link", "url": link,
                "filename": val.get("Original Filename", "")}

    name = str(val.get("Original Filename") or "").strip()
    if not name:
        raise HTTPException(status_code=404, detail="No resume filename recorded for this candidate")

    # Local mode: the stored folder first, then the usual drop folders. Resolve
    # and confirm the result stays inside one of them -- the filename comes from
    # stored data, so a '../' in it must not escape into the filesystem.
    roots = []
    stored = str(val.get("Resume Folder Path") or "").strip()
    if stored:
        roots.append(Path(stored))
    roots += [BASE_DIR / "sandbox" / "resumes", BASE_DIR / "resumes"]

    for root in roots:
        try:
            root = root.resolve()
            candidate = (root / Path(name).name).resolve()
            if candidate.is_file() and candidate.is_relative_to(root):
                return FileResponse(str(candidate), filename=candidate.name)
        except (OSError, ValueError):
            continue

    raise HTTPException(status_code=404,
                        detail=f"Resume file '{name}' not found on disk")


@app.delete("/api/candidates/{app_id}")
def api_delete_candidate(app_id: str):
    # ADDED 2026-09-18: this writes to the LOCAL store unconditionally. In
    # Developer mode the operator is looking at live SharePoint rows, so a
    # "delete" would either silently deactivate an unrelated local row or appear
    # to succeed while the live candidate is untouched. Refuse instead: this
    # server is a read-only view of the live tenant.
    if current_mode() == "developer":
        raise HTTPException(
            status_code=409,
            detail="Developer mode is a read-only view of the live SharePoint data. "
                   "Switch to Testing to modify the local store.")

    db_path = database_path()
    if not db_path.exists():
        raise HTTPException(status_code=404, detail="Candidate not found")

    conn = sqlite3.connect(str(db_path))
    with conn:
        conn.execute("UPDATE candidates SET active = 0 WHERE app_id = ?", (app_id,))
    conn.close()
    return {"success": True, "app_id": app_id}


class CompareRequest(BaseModel):
    app_ids: list[str]
    target_role: Optional[str] = None


@app.post("/api/compare")
async def api_compare_candidates(req: CompareRequest):
    # CHANGED 2026-09-18: opened SQLite directly, so the Compare Matrix kept
    # showing local test candidates while the rest of the app was in Developer
    # mode. Routed through load_candidate_rows() like every other list endpoint.
    # The invented defaults below went too ("USA", "B.S.", "3+ years", 75) --
    # a comparison is exactly where fabricated values do the most damage.
    candidates = []
    by_id = {r["app_id"]: r for r in load_candidate_rows()}
    for app_id in req.app_ids:
        row = by_id.get(app_id)
        if not row:
            continue
        try:
            val = json.loads(row["values_json"])
            raw_skills = val.get("Skills") or val.get("Current Skills", [])
            if isinstance(raw_skills, str):
                raw_skills = [s.strip() for s in raw_skills.split(",") if s.strip()]
            role_1, role_score = split_role_score(val.get("Suggested Role 1", ""))
            try:
                score = int(float(val.get("Match Score") or 0))
            except (TypeError, ValueError):
                score = 0
            candidates.append({
                "app_id": row["app_id"],
                "full_name": val.get("Full Name") or val.get("Name", "Unknown"),
                "email": val.get("Email", ""),
                "phone": val.get("Phone", ""),
                "location": val.get("Location", ""),
                "country": val.get("Country", ""),
                "skills": raw_skills,
                "role_1": role_1,
                "score": score or role_score,
                "education": val.get("Education", ""),
                "experience": val.get("Experience") or val.get("Years Exp", ""),
                "status": row["status"],
            })
        except Exception:
            pass

    if not candidates:
        raise HTTPException(status_code=400, detail="No valid candidates found to compare.")

    # Compute 6-axis Radar chart metrics for each candidate
    axes = ["AI & Machine Learning", "Distributed Systems", "Cloud & Infra", "Languages & Core", "Engineering Depth", "JD Target Fit"]
    
    radar_series = []
    color_palette = ["#00f2fe", "#a855f7", "#10b981", "#f59e0b"]

    for idx, c in enumerate(candidates):
        sk_lower = [str(s).lower() for s in c.get("skills", [])]
        sk_text = " ".join(sk_lower)

        # AI & ML Axis
        ai_score = 40
        if any(w in sk_text for w in ["pytorch", "tensorflow", "llm", "rag", "transformer", "cuda", "deep learning"]):
            ai_score = 95 if "pytorch" in sk_text and "llm" in sk_text else 80

        # Distributed Systems Axis
        dist_score = 40
        if any(w in sk_text for w in ["distributed", "raft", "kafka", "grpc", "microservices", "redis", "concurrency"]):
            dist_score = 95 if "raft" in sk_text or "kafka" in sk_text else 80

        # Cloud & Infra Axis
        cloud_score = 45
        if any(w in sk_text for w in ["kubernetes", "docker", "aws", "gcp", "terraform", "linux", "ebpf"]):
            cloud_score = 90 if "kubernetes" in sk_text else 75

        # Languages & Core
        lang_score = 60
        if any(w in sk_text for w in ["rust", "c++", "go", "python", "cuda"]):
            lang_score = 90

        # Engineering Depth / Seniority
        depth_score = min(100, max(50, int(c.get("score", 75)) + (10 if "senior" in str(c.get("role_1", "")).lower() or "staff" in str(c.get("role_1", "")).lower() else 0)))

        # JD Target Fit
        fit_score = int(c.get("score", 75))

        radar_series.append({
            "name": c["full_name"],
            "color": color_palette[idx % len(color_palette)],
            "values": [ai_score, dist_score, cloud_score, lang_score, depth_score, fit_score]
        })

    # Query Gemini for Head-to-Head Comparative AI Verdict
    ai_verdict = ""
    if GEMINI_ENABLED and GEMINI_API_KEY:
        try:
            import requests
            target_role = req.target_role or (candidates[0].get("role_1") if candidates else "Senior Engineer")
            summary_cand = "\n".join(
                f"- Candidate {c['full_name']}: Score={c['score']}%, Skills={', '.join(c['skills'][:8])}, Experience={c['experience']}, Education={c['education']}"
                for c in candidates
            )
            prompt = (
                f"You are the DriverAI Senior Talent Evaluation Engine. Compare these candidates head-to-head for the role '{target_role}':\n\n"
                f"{summary_cand}\n\n"
                "Provide a concise, high-impact executive verdict:\n"
                "1. **Core Trade-off Summary** (1-2 sentences)\n"
                "2. **Candidate Strengths & Ideal Assignments**\n"
                "3. **Recommended Technical Probing Questions for Each**\n"
                "4. **Final Hiring Recommendation** (Which candidate is the primary pick and why)."
            )
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"
            resp = requests.post(url, json={"contents": [{"parts": [{"text": prompt}]}]}, timeout=25)
            if resp.status_code == 200:
                ai_verdict = resp.json()["candidates"][0]["content"]["parts"][0]["text"]
        except Exception as e:
            logger.warning("Gemini compare verdict error: %s", e)
            ai_verdict = "Both candidates demonstrate strong technical profiles. Review individual skill scores and portfolio projects to determine final alignment."

    return {
        "candidates": candidates,
        "axes": axes,
        "series": radar_series,
        "verdict": ai_verdict,
        "target_role": req.target_role or "Senior Engineering",
    }


# ── AI Recruiter Copilot (Gemini Chat) ───────────────────────────────────────

@app.post("/api/chat")
async def api_chat(req: ChatRequest):
    if not GEMINI_ENABLED or not GEMINI_API_KEY:
        raise HTTPException(status_code=503, detail="Gemini API is not configured.")

    roles = cached_roles()
    roles_summary = ", ".join(r.get("title", "") for r in roles[:20])

    system_instruction = (
        "You are 'DriverAI Recruiter Copilot', an elite, intelligent recruitment AI assistant for DriverAI. "
        "You assist recruiters in searching candidates, evaluating skill overlap against the 105 company roles, "
        "drafting personalized interview invitations, constructive feedback, or polite non-USA location decline letters.\n\n"
        f"Available roles at DriverAI include 105 specialized positions such as: {roles_summary}...\n"
        "Be concise, actionable, professional, and clear. Format responses in clean GitHub Markdown."
    )

    prompt = req.message
    if req.context:
        prompt = f"Context data:\n{json.dumps(req.context, indent=2)}\n\nUser Question:\n{req.message}"

    try:
        import requests
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent"
        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "systemInstruction": {"parts": [{"text": system_instruction}]},
            "generationConfig": {
                "temperature": 0.4,
                "maxOutputTokens": 1000,
            }
        }
        # Key in a header, not the URL: requests puts the URL in every error message.
        resp = requests.post(url, json=payload, timeout=30,
                             headers={"x-goog-api-key": GEMINI_API_KEY})
        resp.raise_for_status()
        data = resp.json()
        reply = data["candidates"][0]["content"]["parts"][0]["text"]
        return {"reply": reply}
    except Exception as e:
        logger.error("Chat error: %s", e)
        raise HTTPException(status_code=502, detail="AI chat request failed; see server log.")


# ── Local Directory Batch Ingestion ──────────────────────────────────────────

class LocalIngestRequest(BaseModel):
    directory_path: str = "./resumes"
    save_to_db: bool = True

@app.post("/api/ingest/local-dir")
def api_ingest_local_dir(req: LocalIngestRequest):
    expanded = Path(os.path.expanduser(req.directory_path)).resolve()
    if not expanded.exists() or not expanded.is_dir():
        raise HTTPException(status_code=400, detail=f"Directory '{req.directory_path}' not found or is not a directory.")

    valid_extensions = {".pdf", ".docx", ".txt", ".rtf", ".md"}
    files = [f for f in expanded.iterdir() if f.is_file() and f.suffix.lower() in valid_extensions]
    
    if not files:
        return {
            "status": "warning",
            "message": f"No resume files found in '{expanded}'. Supported: PDF, DOCX, TXT.",
            "processed_count": 0,
            "results": [],
        }

    roles = cached_roles()
    results = []
    store = get_store() if req.save_to_db else None

    for fpath in files:
        try:
            with open(fpath, "rb") as f:
                content = f.read()
            try:
                text = extract_text_from_bytes(content, fpath.name)
            except Exception:
                text = ""
            if not text or not text.strip():
                try:
                    text = content.decode('utf-8', errors='ignore')
                except Exception:
                    text = ""
            if not text.strip():
                continue
            
            extracted = extract_candidate_details_smart(text)
            role_matches = _rank_roles(extracted, roles, text)

            top = role_matches[0] if role_matches else {"title": "General", "score": 50, "reason": ""}
            country = extracted.get("country", "").strip().upper()
            is_usa = country in ("USA", "US", "UNITED STATES", "UNITED STATES OF AMERICA") or "USA" in country
            
            status = "Scored"
            if GEO_FILTER_USA_ONLY and not is_usa and country and country != "NOT EXTRACTED":
                status = "Location Ineligible"
            elif top.get("score", 0) < 60:
                status = "Under Review"

            app_id = f"APP-{datetime.datetime.utcnow().strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:6]}"
            row_dict = {
                "Application ID": app_id,
                "Full Name": extracted.get("full_name") or fpath.stem.replace("_", " ").title(),
                "Email": extracted.get("email", ""),
                "Phone": extracted.get("phone", ""),
                "Location": extracted.get("location", ""),
                "Country": extracted.get("country", "USA"),
                "Experience": extracted.get("experience_years", ""),
                "Education": extracted.get("education", ""),
                "Looking For Role": extracted.get("looking_for_role", ""),
                "Skills": extracted.get("skills", ""),
                "Suggested Role 1": top.get("title", ""),
                "Match Score": top.get("score", 0),
                "Status": status,
                "Original Filename": fpath.name,
                # ADDED 2026-09-18: remember WHERE the file came from. Only the
                # bare filename was stored, so nothing could reopen the original
                # document later -- see /api/candidates/{app_id}/resume.
                "Resume Folder Path": str(fpath.parent),
                "Notes": text[:600],
            }

            if store:
                store.add(row_dict, sheet="main")

            results.append({
                "app_id": app_id,
                "filename": fpath.name,
                "name": row_dict["Full Name"],
                "role": top.get("title", ""),
                "score": top.get("score", 0),
                "status": status,
            })
        except Exception as e:
            logger.warning("Failed processing file %s: %s", fpath.name, e)

    return {
        "status": "success",
        "message": f"Successfully ingested {len(results)} resumes from '{expanded.name}'",
        "processed_count": len(results),
        "results": results,
    }


# ── Markdown Scorecard & Comparison Export ───────────────────────────────────

class ExportMarkdownRequest(BaseModel):
    app_ids: list[str]
    target_role: Optional[str] = None

@app.post("/api/candidates/export-markdown")
def api_export_markdown(req: ExportMarkdownRequest):
    # CHANGED 2026-09-18: was SQLite-only, so a Markdown scorecard exported in
    # Developer mode described local test candidates. Same source as every other
    # list endpoint now. (It also 404'd outright under a non-SQLite backend.)
    candidates = []
    by_id = {r["app_id"]: r for r in load_candidate_rows()}
    for aid in req.app_ids:
        r = by_id.get(aid)
        if r and r["values_json"]:
            val = json.loads(r["values_json"])
            val["app_id"] = r["app_id"]
            val["status"] = r["status"]
            candidates.append(val)

    if not candidates:
        raise HTTPException(status_code=404, detail="No candidates found for specified IDs.")

    target = req.target_role or (candidates[0].get("Suggested Role 1") if candidates else "Open Position")
    
    # Generate Markdown Table
    lines = [
        f"# 💎 DriverAI Executive Candidate Scorecard",
        f"**Target Role Evaluation**: `{target}`",
        f"**Evaluation Date**: {datetime.datetime.utcnow().strftime('%B %d, %Y')}",
        f"**Source Engine**: DriverAI Hiring Agent (Gemini 3.1 Flash-Lite)",
        "",
        "---",
        "",
        "## 📊 Head-to-Head Candidate Summary",
        "",
        "| Candidate Name | Match Score | Top Role Fit | Location | Status | Key Stack |",
        "| :--- | :---: | :--- | :--- | :--- | :--- |",
    ]

    for c in candidates:
        name = c.get("Full Name", "Anonymous")
        score = c.get("Match Score", 0)
        role = c.get("Suggested Role 1", "N/A")
        loc = c.get("Location", "USA")
        status = c.get("Status", "Scored")
        skills = ", ".join(s.strip() for s in str(c.get("Skills", "")).split(",")[:4])
        lines.append(f"| **{name}** | **{score}%** | {role} | {loc} | `{status}` | {skills} |")

    lines.extend([
        "",
        "---",
        "",
        "## 🔍 Detailed Candidate Profiles",
        "",
    ])

    for c in candidates:
        name = c.get("Full Name", "Anonymous")
        score = c.get("Match Score", 0)
        role = c.get("Suggested Role 1", "N/A")
        exp = c.get("Experience", "N/A")
        edu = c.get("Education", "N/A")
        skills = c.get("Skills", "N/A")
        
        lines.extend([
            f"### 👤 {name} (Match: {score}%)",
            f"- **Primary Fit**: `{role}`",
            f"- **Experience**: {exp}",
            f"- **Education**: {edu}",
            f"- **Core Technical Stack**: {skills}",
            "",
        ])

    return {"markdown": "\n".join(lines)}


# ── Static Web App Serving ───────────────────────────────────────────────────

@app.api_route("/", methods=["GET", "HEAD"])
def serve_index():
    index_file = STATIC_DIR / "index.html"
    if index_file.exists():
        return FileResponse(str(index_file))
    return JSONResponse({"status": "DriverAI Hiring Agent Web API Running"})

app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


def run_web(host: str = "0.0.0.0", port: int = 8000):
    """Launch the Web Application Server."""
    print(f"\n🚀 DriverAI Hiring Agent Web Server launching on http://localhost:{port}")
    print(f"✨ Gemini Brain: {GEMINI_MODEL} (API connected)")
    uvicorn.run(app, host=host, port=port, log_level="info")


if __name__ == "__main__":
    run_web(port=8000)
