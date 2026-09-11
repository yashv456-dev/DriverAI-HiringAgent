"""DriverAI Hiring Agent (local-only, Ollama brain) — public API re-exports."""

from hiring_agent.config import (
    BASE_DIR, OUTPUT_DIR, INPUT_DIR, LOGS_DIR, EXCEL_FILE, LOG_FILE,
    SKILL_KEYWORDS, SKILL_DISPLAY, SECTION_WORDS, COLUMNS,
    GENERIC_TITLE_WORDS, COMPANY_NAME, STATUS_SCORED, STATUS_REJECTED,
    STATUS_LOCATION_REVIEW,
    EXCEL_BASENAME, GEO_FILTER_USA_ONLY, SHAREPOINT_CONFIGURED,
    OLLAMA_ENABLED, OLLAMA_SCORING, OLLAMA_MODEL, OLLAMA_HOST,
    dated_dir, dated_subpath, current_excel_file, all_excel_files,
    active_excel_file, excel_override,
    logger,
)
from hiring_agent.extraction import (
    extract_candidate_details, extract_candidate_details_smart,
    extract_text_from_bytes, resolve_full_name,
    _looks_like_name, _normalize_name, _scan_skill_keywords, html_to_text,
    merge_mail_body_fallback, ai_recheck_fields,
    extract_with_ollama, ocr_health,
)
from hiring_agent.scoring import (
    get_open_roles, suggested_roles, assign_category, rematch_sheet, _skill_set,
)
from hiring_agent.ollama_scorer import ai_score_roles
from hiring_agent.jd_sources import (
    load_jd_sources, save_jd_sources, fetch_jd_role, get_active_roles,
    parse_jd_text,
)
from hiring_agent.local_scorer import (
    get_match_details, score_resume_file, load_candidates,
)
from hiring_agent.geo import (
    GeoDecision, classify_location_usa, is_usa_location, is_foreign_location,
    check_location_usa, verify_location_online,
)
from hiring_agent.intake import run_intake, score_existing_excel
from hiring_agent.sharepoint_scoring import score_from_sharepoint
