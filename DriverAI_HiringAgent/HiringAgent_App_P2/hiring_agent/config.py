"""Central configuration for the Hiring Agent app.

Loads business logic from config.yaml. Runs fully local by default (reads resumes, writes a
LOCAL Excel workbook, Ollama as the free local AI brain). An OPTIONAL online mode reads/writes
the Phase 1 SharePoint workbook — that path needs a .env with app-only credentials; everything
else ignores it.
"""

import datetime
import logging
import os
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path

import yaml

# ── Paths anchored to the app root (HiringAgent_App_P2/) ──────────────────────
BASE_DIR = Path(__file__).resolve().parent.parent

# ── Optional .env (only the SharePoint/online mode needs it) ──────────────────
# Local modes ignore this. If python-dotenv isn't installed or there's no .env, this is a
# silent no-op and the app stays fully offline.
try:
    from dotenv import load_dotenv
    load_dotenv(BASE_DIR / ".env")
except Exception:
    pass

# ── Serverless detection ──────────────────────────────────────────────────────
# True when running inside Azure Functions (or forced via HIRING_SERVERLESS). In that
# mode the app logs to stdout only, keeps scratch files in a temp dir, and never assumes
# a local Ollama is reachable.
import tempfile

_truthy_early = _truthy = {"1", "true", "yes", "on"}
IS_SERVERLESS = bool(
    os.getenv("FUNCTIONS_WORKER_RUNTIME")
    or os.getenv("WEBSITE_INSTANCE_ID")
    or (os.getenv("HIRING_SERVERLESS", "").strip().lower() in _truthy_early)
)

# ── Load config.yaml ──────────────────────────────────────────────────────────
# Guarded so a missing/broken config gives a clear one-line message instead of a raw
# traceback (this runs at import time, before the logger below exists).
_yaml_path = BASE_DIR / "config.yaml"
try:
    with open(_yaml_path, encoding="utf-8") as _f:
        _yaml = yaml.safe_load(_f)
    if not isinstance(_yaml, dict):
        raise ValueError("config.yaml did not parse to a mapping (is it empty?)")
except FileNotFoundError:
    print(f"[HiringAgent] config.yaml not found at {_yaml_path}. "
          f"Copy/restore it next to the app before running.", file=sys.stderr)
    sys.exit(1)
except (yaml.YAMLError, ValueError) as _e:
    print(f"[HiringAgent] config.yaml is invalid ({_e}). "
          f"Fix the YAML at {_yaml_path} and re-run.", file=sys.stderr)
    sys.exit(1)

# ── Company / branding ────────────────────────────────────────────────────────
COMPANY_NAME = _yaml.get("company_name", "DriverAI")

# ── Status values ─────────────────────────────────────────────────────────────
STATUS_SCORED = _yaml.get("status_values", {}).get("scored", "Scored")
STATUS_REJECTED = _yaml.get("status_values", {}).get("rejected", "Rejected - Non-USA Location")
STATUS_NEEDS_REVIEW = _yaml.get("status_values", {}).get("needs_review",
                                                         "Needs Review - Unreadable Resume")
STATUS_LOCATION_REVIEW = _yaml.get("status_values", {}).get(
    "location_review", "Needs Review - Location Confirmation")
STATUS_LOCATION_UNCONFIRMED = _yaml.get("status_values", {}).get(
    "location_unconfirmed", "Rejected - Location Not Confirmed")
STATUS_PROCESSING_FAILED = _yaml.get("status_values", {}).get(
    "processing_failed", "Rejected - Processing Error")
STATUS_NEEDS_REVIEW_SPAM = _yaml.get("status_values", {}).get(
    "needs_review_spam", "Needs Review - Possible Spam")

# P2's secondary content-safety net (see config.yaml content_safety) - checked against a
# row's resume text and mail body before normal extraction runs. P1 already screens
# incoming mail for this (HiringAgent_P1/flow/flow_config.json spam_filters); this only
# catches the rare row that still slips through.
SUSPICIOUS_CONTENT_PHRASES = [
    str(p).strip().lower()
    for p in (_yaml.get("content_safety", {}) or {}).get("suspicious_phrases", [])
    if str(p).strip()
]

# How many times a row may fail (unreadable resume OR any exception during scoring)
# before it's given up on and moved to Rejected - never retried forever.
SCORE_RETRY_MAX = int(os.getenv("HIRING_SCORE_RETRY_MAX") or _yaml.get("score_retry_max", 3))

# Admin recipient for the "row given up on after N failures" alert, and the P2-side alert
# toggle - separate from GEO_REJECT_EMAIL (applicant-facing) and P1's own Notify_failure.
ADMIN_EMAIL = os.getenv("HIRING_ADMIN_EMAIL") or _yaml.get("admin_email", "")

# ── Columns ───────────────────────────────────────────────────────────────────
# The candidate-list (main) schema. Every header is generated up front when the
# workbook is built; P1 fills 11 of them + stamps 'Mail Sent', P2 fills the rest.
COLUMNS = _yaml.get("columns", [])

# The Rejected sheet mirrors the main schema for every CONTENT column, but swaps the
# main-only 'Mail Sent' audit stamp and main-only 'Info Request Sent' marker (a rejected
# candidate is never asked for missing info) for the rejected-only 'Decline Sent' marker.
DECLINE_SENT_COLUMN = "Decline Sent"
MAIL_SENT_COLUMN = "Mail Sent"
INFO_REQUEST_SENT_COLUMN = "Info Request Sent"
REJECTED_COLUMNS = [c for c in COLUMNS if c not in (MAIL_SENT_COLUMN, INFO_REQUEST_SENT_COLUMN)] + [DECLINE_SENT_COLUMN]


# ── Directories (env-overridable) ─────────────────────────────────────────────
# In serverless mode the app dir may be read-only, so scratch defaults to a temp dir.
_data_root = Path(tempfile.gettempdir()) / "hiring_agent" if IS_SERVERLESS else BASE_DIR
OUTPUT_DIR = Path(os.getenv("HIRING_OUTPUT_DIR") or (_data_root / "P2_Output"))
INPUT_DIR = Path(os.getenv("HIRING_INPUT_DIR") or (_data_root / "P2_Input"))
LOGS_DIR = Path(os.getenv("HIRING_LOGS_DIR") or (_data_root / "P2_Logs"))
for _d in (OUTPUT_DIR, INPUT_DIR, LOGS_DIR):
    try:
        _d.mkdir(parents=True, exist_ok=True)
    except OSError:
        pass  # read-only FS (serverless) - file outputs simply won't be used


# ── Dated-folder helpers (Year / Month partitioning) ──────────────────────────
EXCEL_BASENAME = "HiringAgent_P1_CandidateList.xlsx"


def dated_subpath(when=None) -> str:
    """Relative date segments for partitioned storage, e.g. '2026/June'.

    No week-level subfolder (dropped 2026-07-08) — one folder per month."""
    when = when or datetime.date.today()
    return f"{when.year}/{when.strftime('%B')}"


def dated_dir(base: Path, when=None) -> Path:
    """Return base / <Year> / <Month> for the given date (no mkdir)."""
    p = Path(base)
    for seg in dated_subpath(when).split("/"):
        p = p / seg
    return p


def current_excel_file(when=None) -> Path:
    """The yearly master workbook for `when`: OUTPUT_DIR/<Year>/CandidateList.xlsx."""
    when = when or datetime.date.today()
    d = OUTPUT_DIR / f"{when.year}"
    try:
        d.mkdir(parents=True, exist_ok=True)
    except OSError:
        pass  # read-only FS (serverless) — caller handles missing dir
    return d / EXCEL_BASENAME


def all_excel_files() -> list:
    """Every master workbook under OUTPUT_DIR (across all years).

    If HIRING_EXCEL_FILE pins a single workbook, that file is the only one read.
    """
    if EXCEL_FILE_OVERRIDE is not None:
        return [EXCEL_FILE_OVERRIDE]
    return sorted(OUTPUT_DIR.rglob(EXCEL_BASENAME))


def excel_override() -> Path | None:
    """An explicit workbook pinned at runtime via HIRING_EXCEL_FILE (differs from default)."""
    ov = os.getenv("HIRING_EXCEL_FILE")
    if ov and Path(ov) != current_excel_file():
        return Path(ov)
    return None


def active_excel_file(when=None) -> Path:
    """Single workbook to read/write now: HIRING_EXCEL_FILE pin, else `when`'s year file."""
    return EXCEL_FILE_OVERRIDE or current_excel_file(when)


# An explicit HIRING_EXCEL_FILE pins ALL reads/writes to one workbook (no dated routing).
EXCEL_FILE_OVERRIDE = Path(os.environ["HIRING_EXCEL_FILE"]) if os.getenv("HIRING_EXCEL_FILE") else None
EXCEL_FILE = EXCEL_FILE_OVERRIDE or current_excel_file()
LOG_FILE = dated_dir(LOGS_DIR) / "P2.log"
try:
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
except OSError:
    pass

# ── Skills ────────────────────────────────────────────────────────────────────
SKILL_KEYWORDS = _yaml.get("skills", {}).get("keywords", [])
SKILL_DISPLAY = _yaml.get("skills", {}).get("display_names", {})

# ── Section / role / location words ───────────────────────────────────────────
SECTION_WORDS = set(_yaml.get("section_words", []))
ROLE_WORDS = set(_yaml.get("role_words", []))
LOCATION_KEYWORDS = _yaml.get("location_keywords", [])
GENERIC_TITLE_WORDS = set(_yaml.get("generic_title_words", []))

# ── Location filters ──────────────────────────────────────────────────────────
_filters = _yaml.get("filters", {})
KNOWN_US_CITIES = [loc.lower() for loc in _filters.get("known_us_cities", [])]
FOREIGN_COUNTRIES = [loc.lower() for loc in _filters.get("foreign_countries", [])]
FOREIGN_CITIES = [loc.lower() for loc in _filters.get("foreign_cities", [])]
FOREIGN_REGIONS = [loc.lower() for loc in _filters.get("foreign_regions", [])]
FOREIGN_DEMONYMS = [loc.lower() for loc in _filters.get("foreign_demonyms", [])]

US_STATE_ABBREVS = {
    "al", "ak", "az", "ar", "ca", "co", "ct", "de", "fl", "ga",
    "hi", "id", "il", "in", "ia", "ks", "ky", "la", "me", "md",
    "ma", "mi", "mn", "ms", "mo", "mt", "ne", "nv", "nh", "nj",
    "nm", "ny", "nc", "nd", "oh", "ok", "or", "pa", "ri", "sc",
    "sd", "tn", "tx", "ut", "vt", "va", "wa", "wv", "wi", "wy", "dc",
}

US_STATE_NAMES = {
    "alabama", "alaska", "arizona", "arkansas", "california", "colorado",
    "connecticut", "delaware", "florida", "georgia", "hawaii", "idaho",
    "illinois", "indiana", "iowa", "kansas", "kentucky", "louisiana",
    "maine", "maryland", "massachusetts", "michigan", "minnesota",
    "mississippi", "missouri", "montana", "nebraska", "nevada",
    "new hampshire", "new jersey", "new mexico", "new york", "north carolina",
    "north dakota", "ohio", "oklahoma", "oregon", "pennsylvania",
    "rhode island", "south carolina", "south dakota", "tennessee", "texas",
    "utah", "vermont", "virginia", "washington", "west virginia",
    "wisconsin", "wyoming", "district of columbia",
}

US_TERRITORY_NAMES = {
    "puerto rico", "guam", "u.s. virgin islands", "american samoa",
    "northern mariana islands",
}

US_COUNTRY_TERMS = {
    "united states", "united states of america", "usa", "u.s.a.", "u.s.", "us",
}

# ── Default roles ─────────────────────────────────────────────────────────────
DEFAULT_ROLES = _yaml.get("default_roles", [])

# ── AI extraction settings ────────────────────────────────────────────────────
_ai = _yaml.get("ai_extraction", {})
AI_TEXT_LIMIT = _ai.get("text_limit", 12000)

# Ollama: the FREE local-LLM brain (no API key). This app defaults it ON; any failure
# (Ollama not running, model missing, bad JSON) falls back automatically — nothing breaks.
_ollama = _ai.get("ollama", {})


def _envflag(name: str, default: bool) -> bool:
    raw = os.getenv(name)
    if raw is None or raw.strip() == "":
        return default
    return raw.strip().lower() in _truthy


# Ollama defaults ON locally but OFF in serverless (no local model in the cloud — trying
# would just burn the timeout before falling back). Set HIRING_OLLAMA_ENABLED to override.
_ollama_default = bool(_ollama.get("enabled", True)) and not IS_SERVERLESS
OLLAMA_ENABLED = _envflag("HIRING_OLLAMA_ENABLED", _ollama_default)
OLLAMA_MODEL = os.getenv("HIRING_OLLAMA_MODEL") or _ollama.get("model", "llama3.2")
OLLAMA_HOST = (os.getenv("HIRING_OLLAMA_HOST") or _ollama.get("host", "http://localhost:11434")).rstrip("/")
OLLAMA_TIMEOUT = int(_ollama.get("timeout", 60))
# Separate, longer timeout just for role-fit scoring (rates the candidate against every
# open role in one prompt - the heaviest Ollama call). Extraction's OLLAMA_TIMEOUT above
# is untouched - it was never the thing that timed out.
OLLAMA_SCORING_TIMEOUT = int(_ollama.get("scoring_timeout", 150))

# Use the Ollama brain for SCORING too (role-fit + reasoning), not just extraction.
# Falls back to the deterministic keyword scorer when off or unavailable.
OLLAMA_SCORING = _envflag("HIRING_OLLAMA_SCORING", bool(_ollama.get("scoring", True)) and not IS_SERVERLESS)

# Whether to save a local copy of each downloaded resume during scoring. Defaults OFF
# everywhere — resumes live in SharePoint, local copies waste disk and are deleted after
# scoring anyway. Set HIRING_SAVE_LOCAL_COPIES=true to keep them (debugging).
SAVE_LOCAL_COPIES = _envflag("HIRING_SAVE_LOCAL_COPIES", False)

# ── Scoring parameters ────────────────────────────────────────────────────────
_scoring = _yaml.get("scoring", {})
SCORING_MIN_MATCH = _scoring.get("min_match_percent", 20)
SCORING_TOP_N = _scoring.get("top_n", 3)
SCORING_TITLE_BOOST = _scoring.get("title_boost", 10)
SCORING_MIN_ROLE_SKILLS = int(_scoring.get("min_role_skills_denominator", 5))
SCORING_MAX_SKILLS = _scoring.get("max_skills_display", 40)
SCORING_BATCH_LIMIT = int(os.getenv("HIRING_SCORING_BATCH_LIMIT") or _scoring.get("batch_limit", 25))

# ── Role → Category mapping (config-driven; first-match-wins substring rules) ─
_cat_cfg = _yaml.get("role_categories", {})
ROLE_CATEGORY_RULES: list[dict] = _cat_cfg.get("rules", [])
ROLE_CATEGORY_DEFAULT: str = _cat_cfg.get("default", "General")

# ── Geo filter toggle (USA-only hiring; keeps USA, rejects non-USA) ───────────
GEO_FILTER_USA_ONLY = _envflag("HIRING_GEO_USA_ONLY", bool(_yaml.get("geo_filter_usa_only", True)))

# Email the polite non-USA decline when rejecting (needs Mail.Send app permission; see send_mail).
GEO_REJECT_EMAIL = _envflag("HIRING_GEO_REJECT_EMAIL", bool(_yaml.get("geo_reject_email", True)))

# Admin alert when a row is given up on after SCORE_RETRY_MAX failures (needs Mail.Send app
# permission; see send_mail). Independent of GEO_REJECT_EMAIL and P1's own Notify_failure -
# turn off during a bulk historical-replay test where no mail at all should go out.
ERROR_EMAIL_ENABLED = _envflag("HIRING_ERROR_EMAIL", bool(_yaml.get("error_email", True)))

# MASTER kill-switch for every outbound P2 email, mirroring P1's test_mode.suppress_emails.
# Enforced inside SharePointClient.send_mail() - the single choke point all four call sites
# pass through - so no future sender can bypass it the way the missing-info nudge bypassed
# GEO_REJECT_EMAIL (it was gated only on "is a field missing", not on any mail flag).
# Everything else still runs: rows update, resumes move, sheets reconcile. Markers are
# stamped 'TEST-MODE (suppressed) <ts>' instead of 'Sent <ts>' so replayed rows stay
# visibly distinguishable and are never mistaken for a real contact.
SUPPRESS_EMAILS = _envflag(
    "HIRING_SUPPRESS_EMAILS",
    bool((_yaml.get("test_mode") or {}).get("suppress_emails", False)),
)

# Email templates (acknowledgment / request_cv / rejection_non_usa).
EMAIL_TEMPLATES = _yaml.get("email_templates", {})

# ── SharePoint / online mode (optional; only the --score-sharepoint path uses it) ──
# The actual Graph client reads these from the environment directly; we expose them here
# only so the CLI/GUI can tell whether online mode is configured.
TENANT_ID = (os.getenv("TENANT_ID") or "").strip()
CLIENT_ID = (os.getenv("CLIENT_ID") or "").strip()
CLIENT_SECRET = (os.getenv("CLIENT_SECRET") or "").strip()
SHAREPOINT_HOSTNAME = (os.getenv("SHAREPOINT_HOSTNAME") or "").strip()
SHAREPOINT_SITE_PATH = (os.getenv("SHAREPOINT_SITE_PATH") or "").strip()
SHAREPOINT_TABLE = (os.getenv("SHAREPOINT_TABLE") or "").strip()
SHAREPOINT_CONFIGURED = all([
    TENANT_ID, CLIENT_ID, CLIENT_SECRET,
    SHAREPOINT_HOSTNAME, SHAREPOINT_SITE_PATH, SHAREPOINT_TABLE,
])

# ── JD sources file ───────────────────────────────────────────────────────────
JD_SOURCES_FILE = BASE_DIR / "jd_sources.json"
# Parsed-role cache built once by the app / `bot.py --refresh-jd`. Scoring runs read this
# instead of re-downloading every JD each run. Safe to commit (role titles + skills only).
JD_CACHE_FILE = BASE_DIR / "jd_roles_cache.json"

# ── Logging ───────────────────────────────────────────────────────────────────
# Console always (App Insights captures stdout in the cloud). A rotating file handler is
# added only when not serverless and the file is writable.
logger = logging.getLogger("HiringAgent")
logger.setLevel(logging.INFO)
if not logger.handlers:
    _fmt = logging.Formatter("%(asctime)s | %(levelname)-8s | %(message)s")
    _console_handler = logging.StreamHandler()
    _console_handler.setFormatter(_fmt)
    logger.addHandler(_console_handler)
    if not IS_SERVERLESS:
        try:
            _file_handler = RotatingFileHandler(
                LOG_FILE, maxBytes=1_000_000, backupCount=5, encoding="utf-8"
            )
            _file_handler.setFormatter(_fmt)
            logger.addHandler(_file_handler)
        except OSError:
            pass
logger.propagate = False
