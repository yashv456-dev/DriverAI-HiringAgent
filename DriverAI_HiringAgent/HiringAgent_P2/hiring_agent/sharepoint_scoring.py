"""SharePoint scoring worker (online mode): pick up the Phase 1 queue, read each resume,
score it with the Ollama brain, and write the results back to the same SharePoint row.

Input  : the Phase 1 SharePoint workbook (Sharepoint_Master_File.xlsx, rows the PA flow left as 'New Email Received') +
         the resumes saved in the SharePoint resumes folder.
Output : the SAME rows, patched in place with Full Name, Phone, Location, Current Skills,
         Looking For Role, Suggested Role 1/2, Status.
Brain  : Ollama for reading + scoring (with AI recheck pass), keyword/offline fallbacks. Non-USA candidates
         are moved to the workbook's 'Rejected' sheet (when the geo filter is on).
"""

import datetime as dt
import os
import re
import time as _time
import pandas as pd
import hiring_agent.config as _cfg
from hiring_agent.config import (
    STATUS_LOCATION_REVIEW, STATUS_SCORED, GEO_FILTER_USA_ONLY,
    SAVE_LOCAL_COPIES, logger,
)

_temp_files: list = []
from hiring_agent.excel_output import _parse_received
from hiring_agent.extraction import (
    extract_text_from_bytes, extract_candidate_details, extract_candidate_details_smart,
    resolve_full_name,
    merge_mail_body_fallback, ai_recheck_fields, html_to_text,
    infer_looking_for_role, infer_missing_portfolios, _GAP_LITERALS, MISSING_VALUE,
    _looks_like_url, _normalize_link_artifacts, _clean_url, format_phone, sanitize_phone,
    _name_is_sentence_fragment, _is_meaningful_portfolio, clean_location_text,
    education_needs_repair, normalize_skills, clean_role_text,
    _extract_education_dates, _extract_phone, _extract_location, location_is_plausible,
    _looks_like_company_name, _scan_skill_keywords, _extract_education, _extract_experience,
    normalize_education,
    EXTRACTION_SOURCE,
)
from hiring_agent.config import SKILL_DISPLAY, SCORING_MAX_SKILLS
from hiring_agent.scoring import suggested_roles, assign_category
from hiring_agent.jd_sources import get_active_roles
from hiring_agent.geo import (
    GeoDecision, check_location_usa, classify_location_usa,
    normalize_country, reconcile_us_country,
)
from hiring_agent.store import (ExcelCandidateStore, RowVanished, APP_ID_COL, get_store,
                                CandidateRow, storage_backend)


def _store(client):
    """Row identity for this client — see hiring_agent/store.py.

    Stateless and cheap, so it is built where it is used rather than threaded through fifty
    function signatures. Every candidate write in this module goes through it, which is what
    keeps `itemAt(index=...)` from ever being addressed from up here again.
    """
    return get_store(client)


def _app_id_of(values) -> str:
    """The Application ID on a row dict, normalised. '' for a keyless structural row."""
    return str((values or {}).get(APP_ID_COL, "") or "").strip()


def _save_scan_row(client, r, patch: dict, *, sheet: str) -> bool:
    """Patch a row found by a whole-sheet maintenance scan. False if it was skipped.

    These passes used to call `client.update_row` through a `writer` variable, which is how
    they escaped the 2026-09-06 conversion to key-based addressing: a grep for
    `client.update_row(` cannot see an indirect call. They were still writing by position.

    Two cases, in order:

    1. **A structurally blank row is skipped.** The live sheets carry historical spacers -
       the Rejected sheet had drifted to roughly 14 before P2's separator writes were
       disabled on 2026-09-01. `list_rows()` already drops them for the main sheet;
       `list_rejected_rows()` does not, so a backfill pass would stamp 'Missing' into all ten
       columns of a cosmetic blank and turn it into a fake candidate every later pass then
       tries to process.

    2. **A row with an Application ID is addressed by it**, never by position.

    A row with real content but no Application ID still falls back to a positional write:
    there is no key to address it by, and silently skipping it would drop a genuine
    data-integrity repair. It is logged, because in production every row P1 writes has an ID,
    so hitting this path means something upstream is wrong.
    """
    vals = r["values"]
    if not any(str(v or "").strip() for v in vals.values()):
        return False
    app_id = _app_id_of(vals)
    if app_id:
        _store(client).save_by_id(app_id, patch, current_values=vals,
                                  sheet=sheet, hint=r["index"])
        return True
    logger.warning(
        f"       WARNING  : {sheet} row {r['index']} has content but no Application ID; "
        f"patching it by position, which cannot be verified. Investigate the row."
    )
    writer = client.update_rejected_row if sheet == "rejected" else client.update_row
    writer(r["index"], patch, current_values=vals)
    return True


def _row_key(app_id, values) -> str:
    """The key to address a row by: the caller's Application ID, else the row's own.

    Several helpers take `app_id` as an optional argument and are called without it — the
    authoritative value is the one on the row itself, so fall back to that rather than
    failing to locate a row whose ID was sitting in `vals` the whole time.
    """
    return str(app_id or "").strip() or _app_id_of(values)


def _na_if_gap(v) -> str:
    """Portfolio slots read 'N/A' when empty — never a bare blank cell on the sheet.

    'N/A' is the "asked, and there genuinely is none" label. Contrast _missing_if_gap.
    """
    return "N/A" if _is_gap(v) else str(v).strip()


def _compact_portfolios(p1, p2, p3) -> tuple:
    """Pack real links into slots 1..n, dropping any that aren't a candidate's own work.

    Two problems this fixes, both seen live on 2026-08-04:
      * ORDER - APP-20260716-1037-6CE3 had Portfolio 1 = 'N/A' with real GitHub and Vercel
        links sitting in slots 2 and 3. Reading down the column, the candidate looks like he
        has no portfolio until you scroll right.
      * JUNK - APP-20260715-2257-7693 stored 'http://b.sc' (his B.Sc degree text parsed as a
        domain), APP-20260715-2231-25F9 stored 'https://www.asu.edu/' (his university's home
        page) and APP-20260707-1727-0DF7 a Google Drive file link (almost certainly the
        resume we already hold). See _is_meaningful_portfolio.

    Empty trailing slots read 'N/A', which is the settled "there genuinely is none" label.
    """
    kept, seen = [], set()
    for v in (p1, p2, p3):
        s = str(v or "").strip()
        if _is_gap(s) or not _is_meaningful_portfolio(s):
            continue
        key = s.rstrip("/").lower()
        if key in seen:
            continue
        seen.add(key)
        kept.append(s)
    kept += ["N/A"] * (3 - len(kept))
    return tuple(kept[:3])


#: Candidate-supplied columns that read MISSING_VALUE rather than sitting blank, so a reader
#: can tell "we still need this" from "there is none". Portfolios are NOT here: an empty
#: portfolio slot keeps reading 'N/A' (see _na_if_gap), which is the long-standing convention.
#: 'Full Name' and 'Looking For Role' joined on 2026-08-04: without them a single row could
#: show THREE different words for the same state - 'Missing' (Phone), 'Not extracted' (Full
#: Name) and '' (blank) - which reads as inconsistency to the client. Both have their own
#: upstream fallbacks (sender display name; infer_looking_for_role) so reaching this point is
#: rare, but when it happens the cell must say the same thing every other empty cell says.
#: 'Education Start/End Date' joined 2026-09-01 (client instruction) for the same reason -
#: they were the one candidate-owed pair still sitting genuinely blank instead of reading
#: 'Missing' like every other column, once _reevaluate_missing_scored_fields's resume
#: re-check has already had its shot at recovering a real value.
#: 'Years Exp' joined 2026-09-04 with the column itself - a blank there is the same
#: ambiguity ("no experience?" vs "we could not read it?") the placeholder exists to end.
_MISSING_IF_BLANK_COLS = ("Phone", "Location", "Country", "Education", "Current Skills",
                          "Full Name", "Looking For Role",
                          "Education Start Date", "Education End Date",
                          "Years Exp")


def _missing_if_gap(v) -> str:
    """Candidate-owed fields read 'Missing' when empty — never a bare blank cell.

    Added 2026-08-04 on client instruction. A blank Location is ambiguous: it could mean the
    parser failed, or nobody has looked, or the resume genuinely carries no address. Writing
    'Missing' says which. It is still a gap by _is_gap (MISSING_VALUE is in _GAP_LITERALS),
    so healing keeps trying to fill it and the info-request email keeps asking for it.
    """
    return MISSING_VALUE if _is_gap(v) else str(v).strip()


def _apply_missing_placeholders(fields: dict) -> dict:
    """Stamp MISSING_VALUE onto every blank candidate-owed column in `fields`, in place."""
    for _col in _MISSING_IF_BLANK_COLS:
        if _col in fields:
            fields[_col] = _missing_if_gap(fields.get(_col))
    return fields
from hiring_agent.ollama_scorer import ollama_health

# Columns the original Phase 1 table may not have yet: the ones this worker writes,
# plus 'Mail Sent' (P1's post-reply audit stamp — created here so the column already
# exists when the updated PA package is imported; the Excel connector binds item
# columns against the live table schema at import time).
_EXTRA_COLUMNS = ["Country", "Suggested Role 3", "Category",
                  "Portfolio 1", "Portfolio 2", "Portfolio 3", "Mail Sent",
                  "Resume URL", "Resume Folder Path", "Resume Link", "Retry Count",
                  "Info Request Sent", "Education", "Last Updated Date",
                  "Education Start Date", "Education End Date",
                  "Years Exp"]

# Retired columns: deleted from both live sheets (not just stopped-writing-to) the next
# time _ensure_schema runs. 'Possible Duplicate' was replaced by actual row merging;
# 'Candidate Email' was a matching signal for that merge but is no longer needed now that
# merging matches on Phone alone.
_RETIRED_COLUMNS = ["Candidate Email", "Possible Duplicate"]

# 'Resume Link' is an Excel CALCULATED column: one uniform formula that Excel evaluates
# per row from that row's own 'Resume URL'. A per-row DIFFERENT formula can't survive
# because Excel auto-fills a table-column formula across all rows.
#
# The visible text is the STORED file's name ('Sadaf_Khan_7693.pdf'), not the candidate's
# original attachment name. Changed 2026-08-04 on client instruction: showing
# 'Ahsan Hussain Resume Docx.docx' while the link opened
# 'MuhammadHussain_MobileAppsAndroidIOS_3AA7.docx' meant the label never matched the file,
# and the messy originals ('cv (2).pdf', 'SADAF KHAN..pdf', 'PRAJWAL_ANAND_RESUME (18).pdf')
# read as broken links even though every one of them resolved correctly.
#
# Two URL shapes have to be handled, because Graph returns a DIFFERENT webUrl per file type:
#   PDF  -> .../Candidate_Resumes/2026/July/Sadaf_Khan_7693.pdf        (last path segment)
#   DOCX -> .../_layouts/15/Doc.aspx?sourcedoc={GUID}&file=Peter_Vishal_6CE3.docx&action=...
# Taking the last path segment of the Office-doc form yields the whole query string, so the
# '&file=' parameter is read out first and the path-tail split is the PDF/fallback branch.
# Any unexpected shape falls back through IFERROR to Original Filename rather than #VALUE!.
# 'Resume Folder Path' is a separate, unclickable, plain-text column showing only the folder.
_RESUME_LINK_FORMULA = (
    '=IF([@[Resume URL]]="","",'
    'HYPERLINK([@[Resume URL]],'
    'IFERROR('
    'IF(ISNUMBER(SEARCH("&file=",[@[Resume URL]])),'
    'MID([@[Resume URL]],SEARCH("&file=",[@[Resume URL]])+6,'
    'SEARCH("&action=",[@[Resume URL]])-SEARCH("&file=",[@[Resume URL]])-6),'
    'TRIM(RIGHT(SUBSTITUTE([@[Resume URL]],"/",REPT(" ",300)),300))),'
    'TRIM([@[Original Filename]]))))')

# Column Excel should render as TEXT, never a NUMBER — a bare digit-run phone (a foreign
# number deliberately left un-reformatted) would otherwise get auto-detected as a number and
# rendered in scientific notation / right-aligned.
# Columns Excel must treat as literal TEXT, never parse. Each holds something that merely
# LOOKS numeric/date-like, so Excel silently converts it and the original is lost on read.
#   Phone                -> bare digits became a float ('9.71546E+11')
#   Education Start/End  -> 'Aug 2024' became the date serial 45505 (added 2026-08-24; the
#                           columns arrived 2026-08-18 without this guard, so every scored
#                           row rendered two meaningless integers to the client). These are
#                           free text by design - a resume may say 'Present', 'Expected May
#                           2026', or nothing at all - so they must never be parsed as dates.
_TEXT_FORMAT_COLUMNS = ["Phone", "Education Start Date", "Education End Date"]

# Rejected-sheet-only marker column: blank = decline email still owed, otherwise a
# 'Sent <timestamp>' / 'Duplicate...' / 'No valid email...' stamp. The stamp is what
# makes decline sending idempotent across runs, crashes, and backlog re-ingests.
_DECLINE_COL = "Decline Sent"

# Main-sheet-only marker column: blank = a SCORED row still needs checking for a missing-
# info nudge, otherwise a 'Sent <timestamp>' / 'Nothing missing' / 'No valid email' stamp.
# Same idempotency role as _DECLINE_COL, just for the main (not Rejected) sheet.
_INFO_REQUEST_COL = "Info Request Sent"

# Client-facing workbook export: main Candidate List only, no Rejected sheet.
# Chosen by column number from the canonical 33-column SharePoint schema:
# 1,2,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,27,28,29,30,31,32 (Education Start/End Date
# added at 13-14, 2026-08-18 — everything from the old 'Looking For Role' onward shifts +2;
# 'Years Exp' added at 11, directly after Country, 2026-09-04 — everything
# from 'Current Skills' onward shifts +1 again).
_CLIENT_EXPORT_COLUMNS = [
    "Application ID",
    "Received Date",
    "Category",
    "Resume Link",   # moved right after Category / before Full Name at client request, 2026-07-24
    "Full Name",
    "Email",
    "Phone",
    "Location",
    "Country",
    "Years Exp",
    "Current Skills",
    "Education",
    "Education Start Date",
    "Education End Date",
    "Looking For Role",
    "Suggested Role 1",
    "Suggested Role 2",
    "Suggested Role 3",
    "Portfolio 1",
    "Portfolio 2",
    "Portfolio 3",
    "Resume URL",
    "Resume Folder Path",
    "Mail Sent",
]


def _env_truthy(name: str) -> bool:
    return str(os.getenv(name, "") or "").strip().lower() in {"1", "true", "yes", "on"}


def _ensure_schema(client) -> None:
    """Generate the full column schema on BOTH sheets, then set the 'Resume Link' calculated
    column formula. Headers are all created up front here (P2's job — Power Automate can't
    create tables/columns/worksheets); the row writes only fill data. Best-effort: each step
    logs and continues on failure."""
    from sharepoint_client import SharePointError

    logger.info("            Workbook maintenance 1/6: checking Rejected sheet...")
    # If the Rejected sheet exists but is empty and has the wrong column order,
    # delete it so it gets recreated fresh with the correct layout.
    _rej_tbl_early = client._rejected_table_name_if_exists()
    if _rej_tbl_early:
        try:
            cols = client._table_columns_of(_rej_tbl_early)
            if cols and cols[-1] != _DECLINE_COL:
                rows_resp = client._req("GET", f"{client._wb_base()}/tables/{_rej_tbl_early}/rows").json().get("value", [])
                is_empty = True
                for r in rows_resp:
                    vals = r.get("values", [[]])[0]
                    if any(x != "" and x is not None for x in vals):
                        is_empty = False
                        break
                if is_empty:
                    logger.info("   Rejected Candidates sheet has incorrect column order and is empty. Recreating it...")
                    client._req("DELETE", f"{client._wb_base()}/worksheets/Rejected")
                    _rej_tbl_early = None
        except Exception as e:
            logger.warning(f"  WARNING   Could not check/recreate Rejected Candidates sheet: {e}")

    logger.info("            Workbook maintenance 2/6: checking column names...")
    client.rename_column(client.table, "Resume Filename", "Original Filename")
    client.rename_column(client.table, "CV Attempts", "Application Updates")
    client.rename_column(client.table, "Score Attempts", "Retry Count")
    if _rej_tbl_early:
        client.rename_column(_rej_tbl_early, "Resume Filename", "Original Filename")
        client.rename_column(_rej_tbl_early, "CV Attempts", "Application Updates")
        client.rename_column(_rej_tbl_early, "Score Attempts", "Retry Count")
    logger.info("            Workbook maintenance 3/6: checking required columns...")
    try:
        client.ensure_columns(_EXTRA_COLUMNS)
    except SharePointError as e:
        logger.warning(f"  WARNING   Could not ensure candidate-list columns: {e}")
    # Create the Rejected sheet + table with its full header schema up front (not lazily
    # on the first rejection), so both sheets always carry the complete set of columns.
    logger.info("            Workbook maintenance 4/6: checking Rejected table schema...")
    try:
        client._ensure_rejected_table()
        # Repair a Rejected table created before the spacer existed (see
        # ensure_rejected_spacer): no-op unless the table is genuinely empty.
        if client.ensure_rejected_spacer():
            logger.info("            Added the permanent blank spacer row under the "
                        "Rejected header (table had none).")
    except SharePointError as e:
        logger.warning(f"  WARNING   Could not ensure the Rejected sheet: {e}")
    try:
        client.ensure_rejected_columns(["Resume URL", "Resume Link", "Resume Folder Path",
                                        "Retry Count", "Education", "Last Updated Date",
                                        "Education Start Date", "Education End Date",
                                        "Years Exp", _DECLINE_COL])
    except SharePointError as e:
        logger.warning(f"  WARNING   Could not ensure Rejected-sheet columns: {e}")
    # (Re)assert the clickable-link calculated column on both tables, delete the retired
    # 'Candidate Email'/'Possible Duplicate' columns, keep the raw 'Resume URL' helper column
    # VISIBLE-but-narrow (Resume Folder Path stays VISIBLE, plain text), strip any leftover
    # blue+underline hyperlink styling off 'Resume Folder Path' (it's plain text, never a
    # link), autofit both 'Resume Link' and 'Resume Folder Path' to their content, and force
    # 'Phone' to Text so a foreign bare-digit number never renders in scientific notation.
    # Uniform/idempotent — harmless to set every run.
    #
    # 'Resume URL' MUST stay unhidden: a hidden column makes the P1 Excel connector's filtered
    # Get-rows (Get_rows_ref, $filter Email eq sender — the duplicate check) fail with
    # "Unable to match columns in the filtered view (26 vs 27)". We narrow it instead so the
    # long URL doesn't spill visually while the clickable 'Resume Link' remains the nice view.
    logger.info("            Workbook maintenance 5/6: checking workbook formatting...")
    repair_resume_link_formula = _env_truthy("HIRING_REPAIR_RESUME_LINK_FORMULA")
    for _tbl, _sheet in ((client.table, "CandidateList"),
                         (client._rejected_table_name_if_exists(), "Rejected")):
        if _tbl:
            if repair_resume_link_formula:
                try:
                    client.set_calculated_column(_tbl, "Resume Link", _RESUME_LINK_FORMULA)
                except SharePointError as e:
                    logger.warning(f"  WARNING   Could not set Resume Link formula on {_tbl}: {e}")
            for _retired in _RETIRED_COLUMNS:
                client.delete_column(_tbl, _retired)
            client.set_column_hidden(_tbl, _sheet, "Resume URL", False)  # never hide (breaks P1 $filter)
            client.set_column_width(_tbl, _sheet, "Resume URL", 40)      # narrow so it doesn't spill
            client.clear_column_hyperlink_style(_tbl, "Resume Folder Path")
            client.autofit_column(_tbl, _sheet, "Resume Link")
            client.autofit_column(_tbl, _sheet, "Resume Folder Path")
            client.set_column_width(_tbl, _sheet, "Resume Folder Path", 250)
            for _col in _TEXT_FORMAT_COLUMNS:
                try:
                    client.set_column_number_format(_tbl, _col, "@")
                except SharePointError as e:
                    logger.warning(f"  WARNING   Could not set {_col} text format on {_tbl}: {e}")
    logger.info("            Workbook maintenance 6/6: checking row normalization and duplicates...")
    _renormalize_phone_columns(client)
    _renormalize_education_dates(client)
    _reevaluate_missing_scored_fields(client)
    _renormalize_resume_paths(client)
    _renormalize_missing_placeholders(client)
    _merge_result = _merge_duplicate_candidates(client)
    if _merge_result.get("healed"):
        logger.info(f"            Duplicate healing: {_merge_result['healed']} winner row(s) "
                    f"filled from a duplicate this run (see 'Duplicate' lines above).")
    if _merge_result.get("duplicates_found"):
        logger.info(f"            Duplicate rows found this run: "
                    f"{_merge_result['duplicates_found']} - retained pending a reviewed "
                    f"retention procedure (docs/P1_P2_IMPLEMENTATION_STEPS.md Section 1).")
    logger.info("            Workbook maintenance complete.")


def _client_export_path() -> tuple:
    """Return (timestamped_path, latest_path) for the client-facing result workbook."""
    from pathlib import Path
    out_dir = Path(os.getenv("HIRING_CLIENT_EXPORT_DIR") or
                   (_cfg.BASE_DIR / "P2_Final_Results"))
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    return (
        out_dir / f"Candidate_List_Results_{stamp}.xlsx",
        out_dir / "Candidate_List_Results.xlsx",
    )


def _client_export_sort_key(row: dict) -> tuple:
    """Visual order for the export: Received Date descending (current month first), then
    Application ID - reverse=True in export_client_results() applies the descending part.
    Mirrors the CandidateList sheet's own order (resort_candidate_sheets.py, 2026-07-31:
    current month first, newest within a month first) so the client workbook reads the
    same way, per this function's whole reason for existing (see _with_period_separators)."""
    received = _parse_received(row.get("Received Date"))
    app_id = str(row.get("Application ID", "") or "")
    return (received, app_id)


# Must stay in sync with resort_candidate_sheets.py's SEP_LABEL_FMT/YEAR_LABEL_COLUMN - both
# implement the same "current month first, labeled year boundary" convention, just for two
# different artifacts (the live SharePoint sheet vs. this client export workbook).
_YEAR_SEP_LABEL_FMT = "-- {year} --"
_YEAR_SEP_LABEL_COLUMN = "Full Name"


def _client_export_sharepoint_folder(client) -> str:
    """Default upload target: root of Shared Documents library (or override if set)."""
    override = os.getenv("HIRING_CLIENT_EXPORT_SHAREPOINT_FOLDER")
    if override is not None:
        return override.strip().strip("/")
    return ""


def _ensure_period_separator(client, received_raw, *, reader_name: str, adder_name: str,
                             label: str) -> bool:
    """Insert ONE blank row before the first appended row of a new month or year.

    Both live sheets are appended to a row at a time (unlike the client export, which is
    rebuilt whole), so the separator has to be decided per insert: compare the incoming row
    with the last nonblank row and add a spacer when the period changes. Comparing with the
    last row (rather than any historical occurrence) also handles out-of-order repairs that
    reopen an older month at the bottom of the sheet.

    A year change is also a month change, so this deliberately inserts a single blank either
    way - matching P1's rule on CandidateList, where the year check suppresses the month one.
    The very first appended row gets no separator, because the permanent spacer under the
    header already provides that gap.

    Takes the RAW Received Date, not a parsed datetime, on purpose: `_parse_received` returns
    a fallback date rather than None for blank/garbage input, so parsing at the call site
    would make a dateless row look like a brand-new period and insert a spurious blank. The
    gap check therefore has to happen before the parse.

    Best-effort: any failure here is swallowed, since a cosmetic blank row must never stop a
    real rejection - or a real restore - from being recorded. The client's reader/adder are
    resolved by NAME inside the try blocks on purpose, so a stripped-down client that doesn't
    implement them is treated as "can't check" rather than raising at the call site.
    """
    raw = str(received_raw or "").strip()
    if not raw or _is_gap(raw):
        return False
    received = _parse_received(raw)
    if received is None:
        return False
    try:
        rows = getattr(client, reader_name)()
    except Exception as e:
        logger.warning(f"       WARNING  : could not check {label} separators: {e}")
        return False

    last_period = None
    for r in rows:
        vals = r.get("values", r) if isinstance(r, dict) else {}
        raw = str(vals.get("Received Date", "") or "").strip()
        if not raw or _is_gap(raw):
            # Blank spacer / seed row, or a year-label separator (which carries only a
            # Full Name and so survives list_rows' fully-blank filter) - neither is a period.
            continue
        parsed = _parse_received(raw)
        if parsed is not None:
            last_period = (parsed.year, parsed.month)

    if last_period is None:
        return False                      # header spacer already separates the first row
    if (received.year, received.month) == last_period:
        return False                      # same month as the last appended row

    try:
        getattr(client, adder_name)({})
        return True
    except Exception as e:
        logger.warning(f"       WARNING  : could not add {label} separator row: {e}")
        return False


def _ensure_rejected_period_separator(client, received_raw) -> bool:
    """Disabled 2026-09-01 (client instruction): the master sheets must never carry monthly
    blank-row gaps - only the client results export does. Matches P1's flow_config.json,
    which already ships with year_separator.enabled/month_separator.enabled both false. P2
    used to insert its own separator unconditionally here with no off-switch, which is why
    Rejected (P2's primary write path) had drifted to 14 scattered/orphaned blank rows while
    CandidateList (mostly P1-written) had none. Kept as a no-op stub, not deleted outright,
    so every call site below still has a single, well-understood place to re-enable this
    from if that policy ever changes."""
    if getattr(client, "local_execution", False):
        return False
    return False


def _ensure_main_period_separator(client, received_raw) -> bool:
    """Disabled 2026-09-01 (client instruction) - see _ensure_rejected_period_separator."""
    return False


def _with_period_separators(rows: list) -> list:
    """Insert one fully-blank row between calendar months, one after the header, and a
    LABELED row (e.g. '-- 2025 --') at a year boundary instead of a plain blank.

    Mirrors the CandidateList sheet's own convention (resort_candidate_sheets.py,
    2026-07-31) so the client workbook reads the same way - including for rows that were
    NOT re-sorted here (this function only inserts separators; sort order is the caller's
    job, see export_client_results). Blank/labeled rows carry no Application ID, so every
    reader (P1, P2, the audits) skips them.
    """
    blank = {col: "" for col in _CLIENT_EXPORT_COLUMNS}
    if not rows:
        return [dict(blank)]              # permanent spacer directly under the header

    def period(r):
        # A blank/gap Received Date must yield None, not _parse_received's fallback
        # date - otherwise a single dateless row reads as its own period and injects
        # a spurious separator on both sides of itself.
        raw = str(r.get("Received Date") or "").strip()
        if not raw or _is_gap(raw):
            return None
        received = _parse_received(raw)
        return (received.year, received.month) if received is not None else None

    out = [dict(blank)]          # spacer directly under the header
    prev = None
    for r in rows:
        cur = period(r)
        if prev is not None and cur is not None and cur != prev:
            if cur[0] != prev[0]:
                sep = dict(blank)
                sep[_YEAR_SEP_LABEL_COLUMN] = _YEAR_SEP_LABEL_FMT.format(year=cur[0])
                out.append(sep)
            else:
                out.append(dict(blank))
        out.append(r)
        if cur is not None:
            prev = cur
    return out


def audit_client_export_integrity(rows: list, source_candidates: list | None = None) -> list[str]:
    """Audit client export rows to guarantee 100% field completeness and zero drift against source.

    Verifies:
      1. Every candidate row has all 24 _CLIENT_EXPORT_COLUMNS populated (no None, no NaN, no raw gaps).
      2. Required fields (Application ID, Full Name, Email, Category, Suggested Role 1, Location) are non-empty.
      3. Cross-verifies candidate fields against source_candidates to ensure 0% data drift/mismatch.
      4. Auto-heals trivial omissions from source and logs any discrepancies.

    Returns a list of warning/healing messages.
    """
    warns = []
    source_map = {}
    if source_candidates:
        for s in source_candidates:
            if isinstance(s, dict):
                s_vals = s.get("values", s)
            elif hasattr(s, "values") and not callable(s.values):
                s_vals = s.values
            else:
                s_vals = {}
            aid = str(s_vals.get("Application ID", "") or "").strip()
            if aid:
                source_map[aid] = s_vals

    for r in rows:
        aid = str(r.get("Application ID", "") or "").strip()
        if not aid:
            # Separator / spacer row: keep strictly blank
            continue

        # 1. Clean None / NaN values across all 24 client export columns
        for col in _CLIENT_EXPORT_COLUMNS:
            val = r.get(col)
            if val is None or (isinstance(val, float) and pd.isna(val)):
                r[col] = ""

        # 2. Check candidate-owed columns have standardized placeholders rather than bare gaps
        for col in _MISSING_IF_BLANK_COLS:
            if _is_gap(r.get(col, "")):
                r[col] = MISSING_VALUE

        for slot in ("Portfolio 1", "Portfolio 2", "Portfolio 3"):
            if _is_gap(r.get(slot, "")):
                r[slot] = "N/A"

        # 3. Category must never be blank
        if _is_gap(r.get("Category", "")):
            r["Category"] = assign_category(str(r.get("Suggested Role 1", "") or ""),
                                            str(r.get("Current Skills", "") or ""))
            warns.append(f"{aid}: blank Category healed to '{r['Category']}'")

        # 4. Cross-check against source record to guarantee zero mismatch
        if aid in source_map:
            src = source_map[aid]
            for check_col in ("Application ID", "Full Name", "Email", "Category", "Suggested Role 1", "Location"):
                r_val = str(r.get(check_col, "") or "").strip()
                s_val = str(src.get(check_col, "") or "").strip()
                if s_val and r_val != s_val and not _is_gap(s_val):
                    # Align export row directly with source to eliminate any drift
                    warns.append(f"{aid}: mismatch in '{check_col}' ('{r_val}' vs source '{s_val}') -> synced to source")
                    r[check_col] = s_val

    if warns:
        logger.info(f"  Result Sheet Audit: verified {len(rows)} export row(s) with {len(warns)} consistency check(s).")
    return warns


def is_doubt_candidate(fields: dict) -> tuple[bool, str]:
    """Check if candidate has doubts or unaligned fields requiring visual review highlighting.

    Returns (has_doubt: bool, reason: str).
    Triggers on:
      - Status starting with 'Needs Review'
      - Location unconfirmed or ambiguous or 'Missing'
      - Country not matching USA or 'Missing'
      - Phone is weird: masked ('*'), incomplete (<10 digits), or non-US international prefix ('+')
      - Education is 'Missing' or empty
    """
    status = str(fields.get("Status", "") or "").strip()
    if status.startswith("Needs Review"):
        return True, f"Status: {status}"

    loc = str(fields.get("Location", "") or "").strip()
    if not loc or loc.lower() in ("missing", "unknown", "n/a", "not extracted"):
        return True, "Location not aligned/missing"

    country = str(fields.get("Country", "") or "").strip().lower()
    if not country or country in ("missing", "unknown", "n/a", "not extracted") or country not in ("united states", "usa", "us"):
        return True, "Country not matching US"

    phone = str(fields.get("Phone", "") or "").strip()
    digits = "".join(c for c in phone if c.isdigit())
    if not phone or phone.lower() in ("missing", "n/a") or "*" in phone or len(digits) < 10:
        return True, "Phone number is weird/missing/masked"
    if phone.startswith("+") and not (phone.startswith("+1") or phone.startswith("+ 1")):
        return True, "Phone has foreign international prefix"

    edu = str(fields.get("Education", "") or "").strip()
    if not edu or edu.lower() in ("missing", "n/a", "not extracted"):
        return True, "Education is missing/not aligned"

    return False, ""


def prepare_client_export_rows(value_dicts) -> list:
    """Reduce arbitrary candidate value-dicts to the client column set, in display order.

    Shared by every producer of the client-facing sheet (the automatic export below and the
    GUI's Export... button) so there is exactly ONE definition of "the result sheet": the
    minimal column set, the current-month-first order, and the month/year separator rows.
    """
    rows = []
    for vals in value_dicts:
        r = {}
        for col in _CLIENT_EXPORT_COLUMNS:
            v = vals.get(col, "")
            if col in _MISSING_IF_BLANK_COLS:
                v = _missing_if_gap(v)
            elif col in ("Portfolio 1", "Portfolio 2", "Portfolio 3"):
                v = _na_if_gap(v)
            elif col == "Category" and _is_gap(v):
                v = assign_category(str(vals.get("Suggested Role 1", "") or ""),
                                    str(vals.get("Current Skills", "") or ""))
            r[col] = v if v is not None else ""
        rows.append(r)
    rows.sort(key=_client_export_sort_key, reverse=True)
    out = _with_period_separators(rows)
    audit_client_export_integrity(out, value_dicts)
    return out


def write_client_export(rows: list, out_path) -> None:
    """Write already-prepared rows (see prepare_client_export_rows) as the client workbook."""
    audit_client_export_integrity(rows)
    df = pd.DataFrame(rows, columns=_CLIENT_EXPORT_COLUMNS)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Candidates")
        ws = writer.book["Candidates"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for cell in ws[1]:
            cell.fill = cell.fill.copy(fgColor="1F4E78", fill_type="solid")
            cell.font = cell.font.copy(color="FFFFFF", bold=True)
            cell.alignment = cell.alignment.copy(horizontal="left", vertical="center", wrap_text=True)

        url_col = _CLIENT_EXPORT_COLUMNS.index("Resume URL") + 1
        link_col = _CLIENT_EXPORT_COLUMNS.index("Resume Link") + 1
        for r in range(2, ws.max_row + 1):
            url = str(ws.cell(r, url_col).value or "").strip()
            if not url:
                continue
            link_cell = ws.cell(r, link_col)
            display = str(link_cell.value or "").strip()
            # A live 'Resume Link' cell is an Excel TABLE calculated column, so its formula
            # is written in structured-reference form (=IF([@[Resume URL]]="",...)). Those
            # references only resolve INSIDE that table - this export sheet is a plain range
            # with no table at all, so carrying the formula across produces a #NAME? error in
            # the client's copy rather than a link. Graph normally hands back the computed
            # display text, but a not-yet-recalculated or errored cell can still read back as
            # the raw formula, so flatten it to plain text here instead of preserving it.
            if display.startswith("="):
                display = ""
            if not display:
                from urllib.parse import unquote, urlparse, parse_qs
                qs = parse_qs(urlparse(url).query)
                if "file" in qs and qs["file"]:
                    display = qs["file"][0]
                else:
                    display = unquote(url.rstrip("/").rsplit("/", 1)[-1].split("?")[0])
            display = display or "Resume"
            link_cell.value = display
            link_cell.hyperlink = url
            link_cell.style = "Hyperlink"

        for idx, col in enumerate(_CLIENT_EXPORT_COLUMNS, start=1):
            values = [str(ws.cell(r, idx).value or "") for r in range(1, ws.max_row + 1)]
            width = max(len(col), *(min(len(v), 60) for v in values)) + 2
            ws.column_dimensions[ws.cell(1, idx).column_letter].width = min(max(width, 12), 52)
        for name in ("Current Skills", "Education", "Looking For Role"):
            letter = ws.cell(1, _CLIENT_EXPORT_COLUMNS.index(name) + 1).column_letter
            ws.column_dimensions[letter].width = 45
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = cell.alignment.copy(horizontal="left", vertical="top", wrap_text=True)

        # Enforce text format '@' for columns that look numeric/date-like to prevent Excel auto-right-aligning
        for col_name in ("Phone", "Years Exp", "Education Start Date", "Education End Date"):
            if col_name in _CLIENT_EXPORT_COLUMNS:
                col_idx = _CLIENT_EXPORT_COLUMNS.index(col_name) + 1
                for r in range(2, ws.max_row + 1):
                    ws.cell(r, col_idx).number_format = "@"


def export_client_results(client, upload_to_sharepoint: bool = False) -> str:
    """Write a client-facing workbook from the main Candidate List only.

    It excludes the Rejected sheet and emits only _CLIENT_EXPORT_COLUMNS in the
    requested order. By default this is local-only; live P2 runs pass
    upload_to_sharepoint=True so the latest client copy is dropped into SharePoint too.
    """
    if not getattr(client, "local_execution", False) and hasattr(client, "site_id") and storage_backend() == "sqlite":
        from .local_pipeline import run_local_pipeline
        return run_local_pipeline(client, process=False, dry_run=not upload_to_sharepoint)["client_export"]
    if not hasattr(client, "list_rows") and not hasattr(_store(client), "all_rows"):
        logger.info("  Export    : skipped (client does not support full table export).")
        return ""
    try:
        _reevaluate_missing_scored_fields(client)
    except Exception as e:
        logger.debug(f"Missing-field re-evaluation before export skipped: {e}")
    scored = []
    store = _store(client)
    try:
        main_candidates = store.all_rows("main")
    except Exception:
        main_candidates = client.list_rows() if hasattr(client, "list_rows") else []

    scored_candidates = []
    for row in main_candidates:
        if isinstance(row, CandidateRow):
            vals = row.values
        elif isinstance(row, dict):
            vals = row.get("values", row)
        else:
            vals = {}
        if not str(vals.get("Application ID", "") or "").strip():
            continue
        if str(vals.get("Status", "") or "").strip() != STATUS_SCORED:
            continue
        country = str(vals.get("Country", "") or "").strip().lower()
        if country and country not in ("united states", "usa", "us"):
            continue
        phone = str(vals.get("Phone", "") or "").strip()
        if phone.startswith("+") and not (phone.startswith("+1") or phone.startswith("+ 1")):
            continue
        scored_candidates.append(vals)

    # Sort newest-first so deduplication keeps the newest scored application
    scored_candidates.sort(key=_client_export_sort_key, reverse=True)

    seen_applicants = set()
    for vals in scored_candidates:
        email = str(vals.get("Email", "") or "").strip().lower()
        full_name = str(vals.get("Full Name", "") or "").strip().lower()
        key = email if (email and "@" in email) else (full_name if full_name else vals.get("Application ID"))
        if key in seen_applicants:
            continue
        seen_applicants.add(key)
        scored.append(vals)

    out_path, latest_path = _client_export_path()
    write_client_export(prepare_client_export_rows(scored), out_path)

    latest_path.write_bytes(out_path.read_bytes())
    if upload_to_sharepoint:
        sp_name = (os.getenv("HIRING_CLIENT_EXPORT_SHAREPOINT_NAME") or
                   "Candidate_List_Results.xlsx")
        sp_folder = _client_export_sharepoint_folder(client)
        try:
            # Upload to primary target folder (root /)
            client.upload_file(sp_folder, sp_name, latest_path.read_bytes())
            display = "/".join(p for p in (sp_folder, sp_name) if p)
            logger.info(f"  Export    : SharePoint copy updated -> /{display}")
        except Exception as e:
            logger.warning(f"  WARNING   Could not upload client workbook to SharePoint: {e}")
    return str(out_path)


def _renormalize_phone_columns(client) -> None:
    """One-time-safe backfill for Phone text storage and canonical NANP formatting.

    Numeric cells are re-written as text, and any NANP-shaped number (10 digits, or 11
    starting with country code 1 - covers both US and Canada) is normalized to
    '(XXX) XXX-XXXX' regardless of the row's own Country field.

    Deliberately NOT gated on Country == 'United States' (fixed 2026-09-01, client
    instruction): format_phone() already only transforms a 10/11-digit NANP-shaped number
    and leaves every other digit count - a genuine international +NN number - completely
    untouched, so the gate was blocking exactly the rows that most needed it: a live audit
    found 23 rows with a plainly NANP-shaped number ('+1-917-346-8630', '865-356-1921')
    still showing un-canonicalized because Country read 'India' or 'Missing' rather than
    the literal string 'United States'. This only ever changes Phone's punctuation - it
    never reads, infers, or writes Country, so it does not reopen the separate 2026-08-04
    decision not to infer Country from a US-shaped phone.

    Setting a column's numberFormat to Text ('@') only affects values written AFTER the
    format change — Excel doesn't retroactively reformat a cell that already holds a numeric
    value. So existing bare-digit foreign phone numbers (the ones rendering as '9.71546E+11')
    need their exact same digits re-PATCHed back in, now that the column is Text, so Excel
    re-parses the incoming string as literal text instead of a number. No-op (and safe to call
    every run) once a value is text and already canonical."""
    from sharepoint_client import SharePointError
    fixed = 0
    for reader, sheet, label in (
        (client.list_rows, "main", "main"),
        (client.list_rejected_rows, "rejected", "Rejected"),
    ):
        try:
            rows = reader()
        except SharePointError:
            continue
        for r in rows:
            raw = r["values"].get("Phone", "")
            numeric = isinstance(raw, (int, float))
            text_val = (str(int(raw)) if numeric and float(raw).is_integer()
                        else str(raw or ""))
            text_val = format_phone(text_val)
            if not numeric and text_val == str(raw or ""):
                continue
            try:
                if not _save_scan_row(client, r, {"Phone": text_val}, sheet=sheet):
                    continue
                fixed += 1
                logger.info(f"       Fixed     : {label} row {r['index']} Phone "
                            f"{r['values'].get('Phone')!r} -> {text_val!r}")
            except SharePointError as e:
                logger.warning(f"  WARNING   Could not renormalize {label} Phone at row "
                               f"{r['index']}: {e}")
    if fixed:
        logger.info(f"            Phone renormalize: {fixed} cell(s) rewritten this run "
                    f"(drifted since the last pass - see 'Fixed' lines above for which).")


_EDU_DATE_COLUMNS = ("Education Start Date", "Education End Date")
#: Excel's day 0. Serials are days since 1899-12-30 (the 1900 leap-year quirk is already
#: baked into that epoch), so serial 45505 -> 2024-08-01 -> 'Aug 2024'.
_EXCEL_EPOCH = dt.date(1899, 12, 30)


def _excel_serial_to_month_year(raw) -> str:
    """'45505' (however Excel stored it) -> 'Aug 2024'. '' when it isn't a plausible serial."""
    try:
        n = float(raw)
    except (TypeError, ValueError):
        return ""
    if not n.is_integer():
        return ""
    n = int(n)
    # Guard the range: a real education date lands roughly 1950-2100. Anything outside that
    # is not a coerced date and must be left exactly as-is rather than "corrected" into one.
    if not (18264 <= n <= 73415):
        return ""
    try:
        return (_EXCEL_EPOCH + dt.timedelta(days=n)).strftime("%b %Y")
    except (OverflowError, ValueError):
        return ""


def _renormalize_education_dates(client) -> None:
    """One-time-safe backfill: turn any Education Start/End Date cell Excel already converted
    into a date SERIAL back into readable text ('45505' -> 'Aug 2024').

    Same shape and rationale as _renormalize_phone_columns: setting the column's numberFormat
    to Text only governs values written AFTER the change, so cells that were already coerced
    keep rendering as bare integers until they are re-written. Added 2026-08-24 - the two
    columns landed on 2026-08-18 without a Text guard, so every row scored in between shows
    the client two meaningless 5-digit numbers.

    Deliberately conservative: only a whole number inside a plausible calendar range is
    touched. Free-text values the extractor legitimately produces - 'Present', 'Expected May
    2026', or blank - are never numeric and so are never rewritten. Idempotent: once a cell
    holds text, float() fails and it is skipped forever."""
    from sharepoint_client import SharePointError
    for reader, sheet, label in (
        (client.list_rows, "main", "main"),
        (client.list_rejected_rows, "rejected", "Rejected"),
    ):
        try:
            rows = reader()
        except SharePointError:
            continue
        for r in rows:
            patch = {}
            for col in _EDU_DATE_COLUMNS:
                raw = r["values"].get(col, "")
                if isinstance(raw, str) and not raw.strip().isdigit():
                    continue          # already readable text (or blank) - leave it alone
                text_val = _excel_serial_to_month_year(raw)
                if text_val:
                    patch[col] = text_val
            if not patch:
                continue
            try:
                if not _save_scan_row(client, r, patch, sheet=sheet):
                    continue
                logger.info(f"       Fixed     : {label} row {r['index']} education date(s) "
                            f"{ {k: r['values'].get(k) for k in patch} } -> {patch}")
            except SharePointError as e:
                logger.warning(f"  WARNING   Could not renormalize {label} education dates at "
                               f"row {r['index']}: {e}")


#: A Scored row is always a confirmed-US candidate, so its Phone column has exactly two
#: valid states: this canonical shape, or the literal 'Missing'. Used by
#: _reevaluate_missing_scored_fields to catch a non-blank fragment ('245-7289' - 7 digits,
#: no area code) that neither format_phone nor sanitize_phone's <7-digit gap check blanks.
_CANONICAL_US_PHONE_RE = re.compile(r"^\(\d{3}\) \d{3}-\d{4}$")


def _reevaluate_missing_scored_fields(client) -> int:
    """Extra layer (added 2026-09-01, client instruction): re-check every SCORED main-
    sheet row still missing a recoverable field, re-download its resume ONCE, and re-run
    the deterministic extractors against the fresh text. The full field list is the
    `needed` map below: Education Start/End Date, Phone, Location, Full Name, Current
    Skills, Education, Years Exp - in that priority order (Education dates matter most to
    the client, then Phone, then Location).
    Called twice per run: once in the normal workbook-maintenance pass (_ensure_schema,
    before scoring), and again right after the client results export is rebuilt - so a
    field that only just went from 'blank' to genuinely unrecoverable within THIS run
    still gets one more real attempt before conceding.

    _renormalize_education_dates/_renormalize_phone_columns above only clean up a value
    that already exists - they never re-derive one that was never extracted in the first
    place. A live audit found 57/160 scored candidates had no education dates at all
    despite the resume plainly stating them (fixed at the extractor level,
    _extract_education_dates, but that only helps a row scored AFTER the fix) - this is
    the backfill for every row scored before it, generalized to Phone and Location too.
    Deliberately deterministic/regex-only (_extract_phone, _extract_location), never the
    full AI pipeline - this runs in a maintenance sweep, not a scoring pass, and must stay
    fast and Ollama-free. Bounded, shrinking cost: only a row still missing something is
    touched, and a fully-resolved row is never re-downloaded again.

    'Looking For Role' is deliberately NOT in that list, and must never be added back.
    There is no deterministic extractor for it: the only way to produce one is
    infer_looking_for_role, which calls Ollama and is documented to return a non-empty
    sentence no matter how little the resume actually supports. A backfill branch calling
    it lived here from some point after 2026-09-01 until 2026-09-09 and did exactly that -
    a resume stub carrying only a name and a phone number yielded 'Seeking a role in
    customer service or sales, based on relevant skills and experience', which is not
    extraction, it is invention, written onto a Scored row as if a human had read it. It
    also silently broke the Ollama-free guarantee two paragraphs up. A scored row whose
    role never got derived now keeps its blank and _renormalize_missing_placeholders
    labels it 'Missing', which is the honest answer: nobody has read that CV for a role
    preference. The idle recovery pass does NOT pick it up either -
    local_pipeline.RECOVERABLE_FIELDS deliberately omits role-shaped columns as
    "decisions, not transcription". Genuinely re-deriving one needs a real scoring pass
    with the mail body as context: a P1 resume update re-queues the row on its own, and an
    operator can force it with run_local_pipeline(app_ids=...) or rescore_recent_candidates.py.
    See test_resume_availability.py's
    test_missing_resume_does_not_abort_later_scored_field_backfill.

    Scoped to the main sheet only - Rejected rows never reach the client export, so
    there's no result to protect by spending a resume download on them. If a field is
    still unrecoverable after this, _renormalize_missing_placeholders (which runs after
    this in the maintenance pass) is what turns it into the honest 'Missing' label rather
    than a bare blank cell."""
    from sharepoint_client import SharePointError
    store = _store(client)
    try:
        if hasattr(store, "all_rows"):
            rows = store.all_rows("main")
        elif hasattr(client, "list_rows"):
            rows = client.list_rows()
        else:
            rows = []
    except Exception:
        rows = []
    fixed = 0
    for r in rows:
        v = r.values if isinstance(r, CandidateRow) else (r.get("values", r) if isinstance(r, dict) else {})
        if str(v.get("Status", "")).strip() != STATUS_SCORED:
            continue
        start = str(v.get("Education Start Date", "") or "").strip()
        end = str(v.get("Education End Date", "") or "").strip()
        phone = str(v.get("Phone", "") or "").strip()
        location = str(v.get("Location", "") or "").strip()
        name = str(v.get("Full Name", "") or "").strip()
        skills = str(v.get("Current Skills", "") or "").strip()
        education = str(v.get("Education", "") or "").strip()
        years_exp = str(v.get("Years Exp", "") or "").strip()

        # A Scored row is a confirmed-US candidate, so its Phone must read either the
        # canonical (XXX) XXX-XXXX shape or 'Missing' - never a partial/unformattable
        # fragment sitting as-is. Found live 2026-09-01: '245-7289' (7 digits, no area
        # code - too short for format_phone to touch, too long for sanitize_phone's <7-
        # digit gap check to blank) survived untouched on a Scored row.
        needed = {
            "dates": _is_gap(start) or _is_gap(end),
            "phone": _is_gap(phone) or not _CANONICAL_US_PHONE_RE.match(phone),
            "location": _is_gap(location),
            "name": _is_gap(name) or name.lower() in ("not extracted", "missing", "n/a", "none"),
            "skills": _is_gap(skills),
            "education": _is_gap(education),
            "years_exp": _is_gap(years_exp),
        }
        if not any(needed.values()):
            continue          # everything already resolved - nothing to re-check
        app_id = str(v.get("Application ID", "") or "").strip()
        if not _stored_resume_names(
                app_id, v.get("Original Filename", ""),
                full_name=v.get("Full Name"), category=v.get("Category"),
                resume_url=v.get("Resume URL")):
            continue           # no resume on file to re-derive anything from
        subpath = _cfg.dated_subpath(_parse_received(v.get("Received Date")))
        try:
            text, _used, _raw = _download_resume_text(
                client, app_id, v.get("Original Filename", ""), subpath,
                full_name=v.get("Full Name"), category=v.get("Category"),
                resume_url=v.get("Resume URL"))
        except (SharePointError, ResumeUnavailableError) as e:
            logger.warning(f"  WARNING   Skipping field backfill for {app_id}: {e}")
            continue
        if not text:
            continue

        patch = {}
        if needed["dates"]:
            new_start, new_end = _extract_education_dates(text)
            if _is_gap(start) and not _is_gap(new_start):
                patch["Education Start Date"] = new_start
            if _is_gap(end) and not _is_gap(new_end):
                patch["Education End Date"] = new_end
        if needed["phone"]:
            new_phone = format_phone(_extract_phone(text) or "")
            new_phone = sanitize_phone(new_phone)
            if _CANONICAL_US_PHONE_RE.match(new_phone) and new_phone != phone:
                patch["Phone"] = new_phone
            elif not _is_gap(phone) and not _CANONICAL_US_PHONE_RE.match(phone):
                # Non-blank but not canonical, and re-extraction from the resume
                # couldn't produce anything better - never leave a partial/
                # unformattable fragment displayed as if it were a real number.
                patch["Phone"] = MISSING_VALUE
        if needed["location"]:
            new_location = clean_location_text(_extract_location(text) or "")
            if new_location and not _is_gap(new_location) and location_is_plausible(new_location):
                patch["Location"] = new_location
        if needed["name"]:
            new_name = resolve_full_name("", str(v.get("Email", "") or ""), text)
            if new_name and not _is_gap(new_name) and not _looks_like_company_name(new_name):
                patch["Full Name"] = new_name
        if needed["skills"]:
            found_skills = _scan_skill_keywords(text)
            if found_skills:
                patch["Current Skills"] = ", ".join(SKILL_DISPLAY.get(sk, sk.title()) for sk in found_skills[:SCORING_MAX_SKILLS])
        if needed["education"]:
            new_edu = normalize_education(_extract_education(text) or "")
            if new_edu and not _is_gap(new_edu):
                patch["Education"] = new_edu
        if needed["years_exp"]:
            new_exp = _extract_experience(text)
            if new_exp and not _is_gap(new_exp):
                patch["Years Exp"] = new_exp

        if not patch:
            continue
        try:
            hint = r.get("index") if isinstance(r, dict) else None
            _store(client).save_by_id(_app_id_of(v), patch, current_values=v, hint=hint)
            fixed += len(patch)
            logger.info(f"       Fixed     : {app_id} re-derived from resume: {patch}")
        except Exception as e:
            logger.warning(f"  WARNING   Could not backfill field(s) for {app_id}: {e}")
    if fixed:
        logger.info(f"            Missing-field re-evaluation: {fixed} cell(s) filled "
                    f"this run (see 'Fixed' lines above for which).")
    return fixed


def _renormalize_resume_paths(client) -> None:
    """One-time-safe backfill: rebuild any 'Resume Folder Path' cell that doesn't already
    match the expected FOLDER-ONLY shape (e.g. a bare filename or a full path+filename left
    over from an earlier iteration of this column's design) using the row's own
    'Received Date' to reconstruct the dated Year/Month bucket (and the Rejected sub-bucket,
    for a Rejected-sheet row). Safe/idempotent to call every run, same pattern as
    _renormalize_phone_columns."""
    from sharepoint_client import SharePointError
    fixed = 0
    for reader, sheet, label, rejected in (
        (client.list_rows, "main", "main", False),
        (client.list_rejected_rows, "rejected", "Rejected", True),
    ):
        try:
            rows = reader()
        except SharePointError:
            continue
        for r in rows:
            v = r["values"]
            raw = str(v.get("Resume Folder Path", "") or "")
            if not raw:
                continue
            received = _parse_received(v.get("Received Date"))
            subpath = _cfg.dated_subpath(received)
            if rejected:
                subpath = _rejected_subpath(subpath)
            expected = _resume_display_path(client.resumes_folder, subpath)
            if raw == expected:
                continue
            try:
                if not _save_scan_row(client, r, {"Resume Folder Path": expected}, sheet=sheet):
                    continue
                fixed += 1
                logger.info(f"       Fixed     : {label} row {r['index']} Resume Folder Path "
                            f"{raw!r} -> {expected!r}")
            except SharePointError as e:
                logger.warning(f"  WARNING   Could not rebuild {label} Resume Folder Path "
                               f"at row {r['index']}: {e}")
    if fixed:
        logger.info(f"            Resume Folder Path renormalize: {fixed} cell(s) rewritten "
                    f"this run (drifted since the last pass - see 'Fixed' lines above for which).")


def _renormalize_missing_placeholders(client) -> None:
    """One-time-safe backfill: stamp MISSING_VALUE onto any EXISTING row's blank
    candidate-owed cell (Phone, Location, Country, Education, Current Skills, Full Name,
    Looking For Role).

    _apply_missing_placeholders (see above) only ever runs on the `fields` dict for a row
    actively being scored right now - it has no sweep over rows that already exist. A row
    written before the 2026-08-04 'Missing' convention existed, or written by a path that
    doesn't call it, can be left with a genuinely blank cell instead - found live
    2026-09-01: 2 CandidateList rows had a blank Country cell.

    Also catches a cell holding a DIFFERENT gap literal ('Not extracted', 'Not confirmed')
    rather than a bare blank - also found live 2026-09-01, on historical-reconciliation
    rows. MISSING_VALUE itself is deliberately excluded from that check (it's already the
    target state, not a source needing conversion), so this never fires a no-op write.
    Same idempotent-safe-to-call-every-run pattern as the other _renormalize_* functions;
    never overwrites a cell that already holds real content."""
    from sharepoint_client import SharePointError
    fixed = 0
    for reader, sheet, label in (
        (client.list_rows, "main", "main"),
        (client.list_rejected_rows, "rejected", "Rejected"),
    ):
        try:
            rows = reader()
        except SharePointError:
            continue
        for r in rows:
            patch = {}
            for col in _MISSING_IF_BLANK_COLS:
                if col not in r["values"]:
                    continue
                current = str(r["values"].get(col, "") or "").strip()
                if current.lower() != MISSING_VALUE.lower() and _is_gap(current):
                    patch[col] = MISSING_VALUE
            if not patch:
                continue
            try:
                if not _save_scan_row(client, r, patch, sheet=sheet):
                    continue
                fixed += len(patch)
                logger.info(f"       Fixed     : {label} row {r['index']} blank -> "
                            f"'Missing': {', '.join(patch)}")
            except SharePointError as e:
                logger.warning(f"  WARNING   Could not backfill 'Missing' for {label} "
                               f"row {r['index']}: {e}")
    if fixed:
        logger.info(f"            Missing-placeholder backfill: {fixed} cell(s) filled "
                    f"this run (see 'Fixed' lines above for which).")


def _normalize_phone_key(phone: str) -> str:
    """Digits-only phone for cross-row matching, or '' if too short to be meaningful
    (never match on a blank/gap phone or a fragment)."""
    digits = "".join(ch for ch in str(phone or "") if ch.isdigit())
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]   # drop a US country-code prefix so formats line up
    return digits if len(digits) >= 7 else ""


# Columns a merge must never carry from the loser into the winner: P1-owned
# identity/date/mail/resume fields belong to that row's own intake/update
# history. Keep these fixed on the winner even when they are blank-ish in older
# historical rows. The calculated Resume Link formula is also excluded so a
# literal value never replaces the formula.
_MERGE_EXCLUDE_COLS = {
    "Application ID",
    "Received Date",
    "Last Updated Date",
    "Email",
    "Mail Subject",
    "Mail Body",
    "Has Resume",
    "Original Filename",
    "Application Updates",
    "Mail Sent",
    "Resume URL",
    "Resume Link",
    "Resume Folder Path",
}


def _merge_candidate_ready(vals: dict) -> bool:
    """Only de-duplicate rows that have finished intake/scoring.

    A 'New Email Received' row may be a fresh update with a newer resume. Merging an older
    scored/rejected row into it before scoring can copy stale fields and stale resume links
    into the new queue item, so queue rows are deliberately excluded from duplicate merging.
    """
    status = str(vals.get("Status", "") or "").strip()
    return status not in {
        "", "New Email Received", _cfg.STATUS_NEEDS_REVIEW, STATUS_LOCATION_REVIEW,
    }


def _received_timestamp(value):
    """Full timestamp (not just the date) from a 'Received Date' cell, or None if unparsable.

    `_parse_received` deliberately floors to a bare date for folder-bucketing purposes, which
    would make two same-day submissions (e.g. 19:19 and 20:28, the real Akshay Faye case)
    compare as equal - exactly wrong for deciding which one is more recent. This keeps the
    time-of-day so 'most recent wins' is actually accurate."""
    if value in (None, ""):
        return None
    try:
        num = float(value)
    except (TypeError, ValueError):
        num = None
    ts = (pd.to_datetime(num, unit="D", origin="1899-12-30", errors="coerce")
          if num is not None and 20000 <= num <= 80000
          else pd.to_datetime(value, errors="coerce"))
    return ts if pd.notna(ts) else None


def _merge_duplicate_candidates(client, dry_run: bool = False) -> dict:
    """Find rows that represent the SAME candidate, across BOTH sheets (main + Rejected), and
    heal the surviving winner's blank fields from the loser. Matched on **Phone OR Email** —
    never on name (too many false-positive risks: typos, common names; two different people
    can share a name).

    CONTAINMENT CHANGE 2026-09-06 (docs/P1_P2_IMPLEMENTATION_STEPS.md Section 1): this
    function used to also DELETE the loser's row and its saved resume file(s) so exactly one
    row/file remained per candidate. It no longer deletes anything, in either sheet.

    Why: a delete that only half-completes is unrecoverable. The row delete and the file
    delete are two separate Graph calls with no shared transaction, so a run that deletes the
    file and then fails to delete the row - or the reverse - leaves a candidate with a
    dangling reference and no way to tell which half happened after the fact. That is the
    same failure shape as the ghost rows a blocked rejection delete already leaves (see
    hiring_agent/store.py's move_to_rejected), and doing it here too, silently, on every
    ordinary scoring run, multiplies the exposure rather than fixing it. Re-ordering the two
    deletes would not help - see the docstring on that same containment decision - so nothing
    is deleted here until a separate, explicit, reviewed retention procedure exists.

    What is unchanged: identity rule (2026-07-12) - same phone OR same email is the same
    person; same name with a different email AND a different phone is kept as two people.
    Same-email pairs are still found here as P2's safety net for the rare P1 read-after-write
    race (nishantacharekar12@gmail.com, two rows four minutes apart). The row with the more
    recent 'Received Date' is still the WINNER, and any column it is missing is still healed
    in from the loser (never overwriting a real winner value). The loser is simply no longer
    removed: it stays on its sheet, unresolved, visible to a human and to the next run, which
    will find the same group again and skip the healing that already happened (the winner's
    fields are no longer gaps) without re-doing any work.

    dry_run=True computes and logs the same plan without writing anything at all.

    Returns {'healed': int, 'duplicates_found': int, 'errors': int}. 'healed' counts winner
    rows that received at least one field this run. 'duplicates_found' counts every loser row
    identified, healed or not, so a human can locate the accumulating duplicates for the
    future retention procedure - it is not a count of anything removed."""
    from sharepoint_client import SharePointError
    from hiring_agent.config import COLUMNS
    try:
        main_rows = client.list_rows()
    except SharePointError:
        return {"healed": 0, "duplicates_found": 0, "errors": 0}
    rej_rows = client.list_rejected_rows()
    all_rows = [(r, "main") for r in main_rows] + [(r, "rejected") for r in rej_rows]

    by_phone: dict[str, list[int]] = {}
    by_email: dict[str, list[int]] = {}
    for i, (r, _sheet) in enumerate(all_rows):
        if not _merge_candidate_ready(r["values"]):
            continue
        pk = _normalize_phone_key(r["values"].get("Phone", ""))
        if pk:
            by_phone.setdefault(pk, []).append(i)
        ek = _bare_email(r["values"].get("Email", ""))
        if ek:
            by_email.setdefault(ek, []).append(i)

    # Email/phone matches form connected components, not independent pairs. For example,
    # A can share a phone with B while B shares an email with C. Pairwise planning could
    # leave C behind or delete B before A's missing fields reached the final winner.
    # Collapse each full component once, with the newest row as its single winner.
    eligible = {
        i for i, (row, _sheet) in enumerate(all_rows)
        if _merge_candidate_ready(row["values"])
    }
    parent = {i: i for i in eligible}

    def _find(i: int) -> int:
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def _union(a: int, b: int) -> None:
        ra, rb = _find(a), _find(b)
        if ra != rb:
            parent[rb] = ra

    for groups in (by_phone.values(), by_email.values()):
        for members in groups:
            active = [i for i in members if i in eligible]
            for i in active[1:]:
                _union(active[0], i)

    components: dict[int, list[int]] = {}
    for i in eligible:
        components.setdefault(_find(i), []).append(i)

    def _recency_key(pos: int) -> tuple[int, str]:
        ts = _received_timestamp(all_rows[pos][0]["values"].get("Received Date"))
        try:
            tick = int(pd.Timestamp(ts).value) if ts is not None else -1
        except (TypeError, ValueError, OverflowError):
            tick = -1
        # On an exact tie (same Received Date to the second - a genuine near-simultaneous
        # duplicate), break by Application ID instead of physical table position. Position
        # (`-pos`) used to double as "earlier-created row" because rows were always in
        # insertion order, but resort_candidate_sheets.py (2026-07-31) can reorder the table
        # for display, which would silently repoint this tiebreak at an unrelated row. The
        # Application ID is stable regardless of where a row physically sits.
        app_id = str(all_rows[pos][0]["values"].get("Application ID", "") or "")
        return tick, app_id

    plan: list[tuple[int, list[int], dict]] = []
    for members in components.values():
        if len(members) < 2:
            continue
        winner_pos = max(members, key=_recency_key)
        losers = sorted(
            (i for i in members if i != winner_pos),
            key=_recency_key,
            reverse=True,
        )
        winner_values = all_rows[winner_pos][0]["values"]
        heal = {}
        for col in COLUMNS:
            if col in _MERGE_EXCLUDE_COLS or not _is_gap(winner_values.get(col, "")):
                continue
            source = next(
                (all_rows[i][0]["values"].get(col) for i in losers
                 if not _is_gap(all_rows[i][0]["values"].get(col, ""))),
                None,
            )
            if source is not None:
                heal[col] = source
        plan.append((winner_pos, losers, heal))

    healed = duplicates_found = errors = 0
    for winner_pos, loser_positions, heal in plan:
        winner_row, winner_sheet = all_rows[winner_pos]
        wv = winner_row["values"]
        w_id = str(wv.get("Application ID", "")).strip()
        duplicates_found += len(loser_positions)
        for loser_pos in loser_positions:
            lv = all_rows[loser_pos][0]["values"]
            l_id = str(lv.get("Application ID", "")).strip()
            logger.info(
                f"  Duplicate : {l_id} (sender {lv.get('Email','')}, received "
                f"{lv.get('Received Date','')}) is a duplicate of {w_id} (sender "
                f"{wv.get('Email','')}, received {wv.get('Received Date','')}) - "
                f"retained, not deleted"
                + (f" | healed: {', '.join(heal)}" if heal else "")
            )

        if dry_run or not heal:
            continue

        try:
            _save_scan_row(client, winner_row, heal, sheet=winner_sheet)
            healed += 1
        except SharePointError as e:
            errors += 1
            logger.warning(
                f"  WARNING   Could not heal winner {w_id} from its duplicate(s): {e}"
            )

    return {"healed": healed, "duplicates_found": duplicates_found, "errors": errors}


def _bare_email(cell) -> str:
    """Normalize an Email cell ('Jane <j@x.com>' or 'j@x.com') to a lowercase address."""
    s = str(cell or "").strip()
    m = re.search(r'<([^>]+@[^>]+)>', s)
    if m:
        s = m.group(1)
    s = s.strip().lower()
    return s if "@" in s else ""


def _name_needs_rederive(full_name: str, email: str) -> bool:
    """True when the stored name is a gap, a sender-address fallback, or prose.

    Fixed 2026-08-04 - the email-fallback test was INVERTED in practice. It stripped every
    non-letter from both the name and the email local-part and compared them, so anyone
    whose address is simply their name flunked forever:

        'Prerna Saluja'  vs prernasaluja259@   -> 'prernasaluja'  == 'prernasaluja'  -> True
        'Rutuja Shingate' vs rutujashingate2000@ -> 'rutujashingate' == same         -> True

    Five correct names on the live sheet were permanently marked as needing re-derivation,
    while the two genuinely broken ones ('for Scientific Computing.', 'Ashfaq S.') returned
    False and were never repaired. It fired where it shouldn't and stayed silent where it
    mattered.

    P1's fallback writes the bare local-part as the name ('sbhask22'), which is always a
    SINGLE token with no whitespace. A real person's name that happens to match their email
    still contains a space. Requiring the no-space shape separates the two cleanly.

    Also returns True for prose masquerading as a name, so a scraped sentence fragment is
    re-derived instead of being kept (and copied into the resume file name).
    """
    name = str(full_name or "").strip()
    if _is_gap(name) or "@" in name:
        return True
    if _name_is_sentence_fragment(name):
        return True
    # A multi-token name is a real name, even when the address spells it out.
    if len(name.split()) > 1:
        return False
    email_addr = _bare_email(email)
    local = email_addr.split("@", 1)[0] if "@" in email_addr else ""
    if not local:
        return False
    clean_name = re.sub(r"[^a-z]", "", name.lower())
    clean_local = re.sub(r"[^a-z]", "", local.lower())
    return bool(clean_name and clean_local and clean_name == clean_local)


def _retry_count(vals: dict) -> int:
    """Current 'Retry Count' value, tolerant of a blank/gap/non-numeric cell (reads as 0)."""
    try:
        return int(str(vals.get("Retry Count", "") or "0").strip() or "0")
    except (TypeError, ValueError):
        return 0


# Keep old name as a simple alias so any external callers (e.g. bot.py) still work.
_score_attempts_of = _retry_count


def _send_error_alert(client, app_id: str, sender_email: str, reason: str) -> None:
    """Best-effort admin alert when a row is given up on after SCORE_RETRY_MAX failures.
    Never raises - a failed notice must not break a scoring run. Gated on
    ERROR_EMAIL_ENABLED (HIRING_ERROR_EMAIL) and a configured ADMIN_EMAIL; silently does
    nothing if either is off/unset (e.g. during a bulk historical-replay test)."""
    if not _cfg.ERROR_EMAIL_ENABLED or not _cfg.ADMIN_EMAIL:
        return
    try:
        subject = f"[Hiring Agent] Row given up after {_cfg.SCORE_RETRY_MAX} failures ({app_id or 'no ref'})"
        body = (
            f"<p>A candidate row failed processing {_cfg.SCORE_RETRY_MAX} times and has been "
            f"moved to the Rejected sheet under '{_cfg.STATUS_PROCESSING_FAILED}' instead of "
            f"being retried forever.</p>"
            f"<ul><li>Application ID: {app_id or '(none)'}</li>"
            f"<li>Sender: {sender_email or '(unknown)'}</li>"
            f"<li>Reason: {reason}</li></ul>"
            f"<p>Check P2_Logs for the full error history on this row.</p>"
        )
        client.send_mail(_cfg.ADMIN_EMAIL, subject, body, admin=True)
    except Exception as e:
        logger.warning(f"       WARNING  : could not send error alert: {e}")


def _detect_suspicious_content(resume_text: str, mail_body: str) -> list:
    """Return the suspicious phrase(s) found in the resume text or mail body, else [].

    P2's secondary content-safety net (config.yaml content_safety.suspicious_phrases).
    P1 already screens incoming mail for spam/scam/vendor-pitch content before a row is
    ever created (HiringAgent_P1/flow/flow_config.json spam_filters) - this only catches
    the rare row that still slips through. Added 2026-08-01 after a services vendor's
    company-brochure PDF ("Futurism Technologies") was scored as if it were a resume."""
    haystack = f"{resume_text or ''}\n{mail_body or ''}".lower()
    return [p for p in _cfg.SUSPICIOUS_CONTENT_PHRASES if p in haystack]


def _send_suspicious_content_alert(client, app_id: str, sender_email: str, matched: list) -> None:
    """Best-effort admin alert when P2's secondary content-safety net catches a row that
    reached SharePoint despite P1's spam gate. Never raises. Gated on ERROR_EMAIL_ENABLED
    and a configured ADMIN_EMAIL, same as _send_error_alert."""
    if not _cfg.ERROR_EMAIL_ENABLED or not _cfg.ADMIN_EMAIL:
        return
    try:
        subject = f"[Hiring Agent] Row flagged as possible spam/vendor pitch ({app_id or 'no ref'})"
        body = (
            f"<p>P2's secondary content-safety check found suspicious content in a row that "
            f"reached SharePoint despite P1's spam gate. It has been marked "
            f"'{_cfg.STATUS_NEEDS_REVIEW_SPAM}' and otherwise left untouched (no scoring, no "
            f"resume rename/move) - please review it manually.</p>"
            f"<ul><li>Application ID: {app_id or '(none)'}</li>"
            f"<li>Sender: {sender_email or '(unknown)'}</li>"
            f"<li>Matched phrase(s): {', '.join(matched)}</li></ul>"
        )
        client.send_mail(_cfg.ADMIN_EMAIL, subject, body, admin=True)
    except Exception as e:
        logger.warning(f"       WARNING  : could not send suspicious-content alert: {e}")


def _send_missing_workbook_alert(client, reason: str) -> None:
    """Best-effort admin alert when the candidate workbook itself can't be confirmed. P2
    never creates or restores the workbook (same non-destructive design as Phase 1) - a
    genuinely missing/unreachable workbook always needs a human, never a silent auto-create
    that could paper over a real outage, a moved file, or a permissions change."""
    if not _cfg.ERROR_EMAIL_ENABLED or not _cfg.ADMIN_EMAIL:
        return
    try:
        subject = "[Hiring Agent] P2 could not confirm the candidate workbook"
        body = (
            f"<p>P2 could not open the candidate workbook and will not run until this is "
            f"fixed. P2 never auto-creates or overwrites the workbook - if it is genuinely "
            f"missing, restore it once from "
            f"<code>P1_Templates/HiringAgent_P1_CandidateList.xlsx</code> at the configured "
            f"path, then re-run. If it still exists, this may be a transient Graph error or "
            f"a permissions/throttling issue instead.</p>"
            f"<ul><li>Reason: {reason}</li></ul>"
        )
        client.send_mail(_cfg.ADMIN_EMAIL, subject, body, admin=True)
    except Exception as e:
        logger.warning(f"  WARNING   could not send missing-workbook alert: {e}")


def _ensure_workbook_or_alert(client, dry_run: bool) -> bool:
    """Confirm (never create) the candidate workbook before a run touches anything else.
    Shared by every entry point that used to independently downgrade this to a warning-and-
    continue - now a missing/unreachable workbook stops the run and alerts an admin instead
    of proceeding into a pass that would fail on every single row anyway. Returns True if the
    workbook is confirmed present (or this is a dry-run, which never touches it), False if the
    caller should abort."""
    from sharepoint_client import SharePointError
    if dry_run:
        logger.info("  SKIP      Dry-run: not touching the workbook.")
        return True
    try:
        client.ensure_workbook()
        return True
    except SharePointError as e:
        logger.warning(f"  MISSING   Candidate workbook could not be confirmed ({e}) - "
                       f"P2 will NOT create one. Restore it from the template and re-run.")
        _send_missing_workbook_alert(client, str(e))
        return False


def _drop_retry_duplicate_if_completed(client, index: int, app_id: str, vals: dict) -> bool:
    """A retry/no-content row is stale because another copy of the same Application ID
    already completed. Drop only this queue row instead of creating a false review or
    processing-error copy."""
    if not app_id:
        return False
    if not hasattr(client, "list_rows") or not hasattr(client, "list_rejected_rows"):
        return False
    try:
        main_rows = client.list_rows()
        rejected_rows = client.list_rejected_rows()
    except Exception as e:
        logger.warning(f"       WARNING  : could not check duplicate completion before give-up: {e}")
        return False

    completed_main = False
    for row in main_rows:
        row_vals = row.get("values", {})
        if str(row_vals.get("Application ID", "") or "").strip() != app_id:
            continue
        if row.get("index") == index:
            continue
        status = str(row_vals.get("Status", "") or "").strip()
        if status and status not in {"New Email Received", _cfg.STATUS_NEEDS_REVIEW}:
            completed_main = True
            break

    completed_rejected = any(
        str(row.get("values", {}).get("Application ID", "") or "").strip() == app_id
        for row in rejected_rows
    )
    if not (completed_main or completed_rejected):
        return False

    try:
        if not _store(client).delete_by_id(_row_key(app_id, vals), hint=index):
            # Nothing was removed, so the caller must not record a deletion.
            logger.warning(
                f"       WARNING  : stale duplicate retry row for {app_id} could not be "
                f"located to remove; leaving it for the next run."
            )
            return False
        location = "main" if completed_main else "Rejected"
        logger.info(f"       Result   : duplicate retry row for {app_id} already completed on "
                    f"{location}; removed stale main copy, no processing-error alert.")
        return True
    except Exception as e:
        logger.warning(f"       WARNING  : could not delete stale duplicate retry row for {app_id}: {e}")
        return False


def _give_up_and_reject(client, index: int, vals: dict, app_id: str, stored: list,
                        subpath: str, reason: str, dry_run: bool) -> None:
    """A row has failed SCORE_RETRY_MAX times - stop retrying forever and move it to the
    Rejected sheet under STATUS_PROCESSING_FAILED (distinct from the non-USA rejection
    reason, so a human reviewing Rejected can tell the two apart at a glance).

    Moves the resume file(s) to the Rejected sub-bucket if any were saved (some failures,
    like a missing filename, have nothing to move). Pre-stamps 'Decline Sent' so
    _send_pending_declines never sends the non-USA decline text to a row that was
    rejected for processing failure, not geography. Sends a best-effort admin alert."""
    new_attempts = _score_attempts_of(vals) + 1
    if dry_run:
        logger.info(f"       Result   : [DRY-RUN] Would give up after {new_attempts} failures "
                    f"({reason}) - would move to Rejected as '{_cfg.STATUS_PROCESSING_FAILED}'.")
        return
    from sharepoint_client import SharePointError
    if _drop_retry_duplicate_if_completed(client, index, app_id, vals):
        return
    merged = dict(vals)
    merged["Status"] = _cfg.STATUS_PROCESSING_FAILED
    merged["Retry Count"] = new_attempts
    merged["Decline Sent"] = f"N/A - processing error, not a geography decline ({reason})"
    if stored and subpath:
        try:
            _move_resumes(client, stored, subpath, to_rejected=True)
            resume_url, resume_path = _resume_url_and_path(client, stored, _rejected_subpath(subpath))
            if resume_url:
                merged["Resume URL"] = resume_url
                merged["Resume Folder Path"] = resume_path
        except SharePointError as e:
            logger.warning(f"       WARNING  : could not move resume file(s) while giving up: {e}")
    deleted = False
    try:
        merged = _complete_row_before_reject(merged, app_id)
        _ensure_rejected_period_separator(client, merged.get("Received Date"))
        deleted = _store(client).move_to_rejected(_row_key(app_id, vals), merged, hint=index)
        logger.info(f"       Result   : GIVEN UP after {new_attempts} failures ({reason}) - "
                    f"moved to Rejected as '{_cfg.STATUS_PROCESSING_FAILED}'.")
        _send_error_alert(client, app_id, str(vals.get("Email", "")), reason)
    except SharePointError as e:
        logger.warning(f"       WARNING  : could not move given-up row to Rejected: {e}")
    # Whether the MAIN row is gone decides the caller's index bookkeeping, so this must
    # reflect delete_row specifically - not merely "the function ran".
    return deleted


def _mark_needs_review(client, index: int, vals: dict, dry_run: bool,
                       app_id: str = "", stored: list | None = None, subpath: str = "") -> bool:
    """Flag a row whose resume can't be read as 'Needs Review - Unreadable Resume', and
    count it toward the SCORE_RETRY_MAX give-up threshold. A row still unreadable after
    SCORE_RETRY_MAX attempts is moved to Rejected (_give_up_and_reject) instead of being
    retried forever.

    Until then, the row stays IN the scoring queue (list_unscored_rows includes this
    status), so a later parser fix or re-sent resume still gets picked up automatically —
    the status just makes these rows visible in SharePoint instead of looking like fresh
    arrivals.

    Returns True when the MAIN row was DELETED (given up on, or dropped as a stale retry
    duplicate). The caller records that in its deleted_idx bookkeeping, which since the move
    to hiring_agent/store.py only sharpens the POSITION HINT handed to the store - a forgotten
    delete path now costs one recovery scan instead of writing to the wrong row. Fixed
    2026-08-03: this used to return None, so both delete paths here were invisible to
    score_from_sharepoint, which only tracked deletions made by _finish_rejection. Keep
    returning it: the hint is what disambiguates two rows sharing an Application ID."""
    if dry_run:
        return False
    if _drop_retry_duplicate_if_completed(client, index, app_id, vals):
        return True
    new_attempts = _score_attempts_of(vals) + 1
    if new_attempts >= _cfg.SCORE_RETRY_MAX:
        return _give_up_and_reject(client, index, vals, app_id, stored or [], subpath,
                                   "resume unreadable", dry_run)
    try:
        _store(client).save_by_id(_row_key(app_id, vals), {"Status": _cfg.STATUS_NEEDS_REVIEW,
                                  "Retry Count": new_attempts}, current_values=vals, hint=index)
        logger.info(f"       Status    : marked '{_cfg.STATUS_NEEDS_REVIEW}' "
                    f"(attempt {new_attempts}/{_cfg.SCORE_RETRY_MAX}).")
    except Exception as e:
        logger.warning(f"       WARNING  : could not set Needs Review status: {e}")
    return False   # row updated in place, never removed


# P1 owns these columns — P2 reads but never writes them.
_P1_COLUMNS = {
    "Application ID", "Received Date", "Last Updated Date", "Email", "Mail Subject",
    "Mail Body", "Has Resume", "Original Filename", "Application Updates",
    "Mail Sent",   # PA flow's post-reply audit stamp
}

# P2 fills these columns. Mapped to the detail-dict key used during extraction.
_P2_COL_TO_KEY = {
    "Full Name":        "full_name",
    "Phone":            "phone",
    "Location":         "location",
    "Country":          "country",
    "Current Skills":   "skills",
    "Years Exp": "experience",
    "Looking For Role": "looking_for_role",
    "Education":        "education",
    "Education Start Date": "education_start_date",
    "Education End Date":   "education_end_date",
    "Portfolio 1":      "portfolio_1",
    "Portfolio 2":      "portfolio_2",
    "Portfolio 3":      "portfolio_3",
}


def _log_column_diff(fields: dict, before: dict, after_extract: dict) -> None:
    """Step E: diff the Step A snapshot against the final fields dict, per column.

    `before` is every column's value before any extraction ran this candidate;
    `after_extract` is the Tier 1/2 field values right after extract_candidate_details_smart
    returned (before the mail-body merge and the AI recheck pass could still touch them) -
    comparing the two tells "filled by the tiered extraction step itself" apart from
    "filled later, by mail-body merge or AI recheck".
    """
    from hiring_agent.config import COLUMNS
    unchanged, filled, healed_cols = [], [], []
    method_lines = []
    for col in COLUMNS:
        was = str(before.get(col, "") or "").strip()
        now = str(fields.get(col, "") or "").strip()
        was_gap, now_gap = _is_gap(was), _is_gap(now)
        if was_gap and not now_gap:
            filled.append(col)
        elif not was_gap and now != was:
            healed_cols.append(col)
        else:
            unchanged.append(col)

        if col not in _P2_COL_TO_KEY:
            continue
        if was_gap and not now_gap:
            mid = str(after_extract.get(col, "") or "").strip()
            method = ("deterministic-only / AI-confirmed hint" if not _is_gap(mid)
                      else "AI-inferred-from-empty (mail-body merge or AI recheck)")
            method_lines.append(f"{col}: {method}")
        elif not was_gap and now != was:
            # Don't assert "AI" when it might just be a fresh regex re-scan (e.g. this
            # run's resume text differs from what was stored) - after_extract lets us
            # tell "changed during the tiered extraction step" apart from "changed later,
            # by mail-body merge or the AI recheck pass" without overclaiming AI involvement.
            mid = str(after_extract.get(col, "") or "").strip()
            method = ("corrected during extraction (regex re-scan or AI-confirm)" if mid != was
                      else "corrected later (mail-body merge or AI recheck)")
            method_lines.append(f"{col}: {method}")

    logger.info(f"       Diff       : {len(unchanged)} unchanged | {len(filled)} filled | "
                f"{len(healed_cols)} healed")
    if method_lines:
        logger.info("       Method     : " + "; ".join(method_lines))


def _cross_field_check(fields: dict, before: dict | None) -> tuple[dict, list[str]]:
    """Deterministic, enhancement-aware consistency audit — run once, right before the write.

    P2 filling a blank or correcting a P1 value is EXPECTED and is never flagged. This pass
    only reacts to (a) P1's 8 owned columns being altered by P2, (b) internal contradictions
    in P2's final values, and (c) a real value that regressed to a gap. It heals only the
    trivially-safe cases (blank Country derivable from Location; blank Category) and logs
    everything else for human review. It NEVER blocks the write — a flagged row still saves
    so the run continues. Deterministic (no AI), so local and cloud validate identically.

    Returns (possibly-healed fields, list of warning strings).
    """
    from hiring_agent.geo import is_strong_usa, is_foreign_location
    from hiring_agent.extraction import split_location_country
    f = dict(fields)
    warns: list[str] = []

    # 1) P1 pass-through integrity — P2 must never change P1's owned columns. (fields
    #    normally omits them entirely, so this is a cheap guard against a future regression.)
    if before is not None:
        for col in _P1_COLUMNS:
            if col in f:
                was = str(before.get(col, "") or "").strip()
                now = str(f.get(col, "") or "").strip()
                if was and now and was != now:
                    warns.append(f"P1 column '{col}' was changed by P2 ('{was}' -> '{now}')")

    # 2) Country <-> Location agreement (highest-value check). Mirror geo precedence:
    #    a strong US signal (state/ZIP) wins over a foreign-city substring so 'Paris, Texas'
    #    is treated as US, not flagged against France.
    loc = str(f.get("Location", "") or "").strip()
    ctry = str(f.get("Country", "") or "").strip()
    if not _is_gap(loc):
        loc_us = is_strong_usa(loc)
        loc_foreign = (not loc_us) and is_foreign_location(loc)
        ctry_us = (not _is_gap(ctry)) and ctry.lower() in _cfg.US_COUNTRY_TERMS
        ctry_foreign = (not _is_gap(ctry)) and any(fc == ctry.lower() for fc in _cfg.FOREIGN_COUNTRIES)
        if _is_gap(ctry):
            _, derived = split_location_country(loc)   # heal a blank Country from the Location
            derived = normalize_country(derived)       # USA/America → United States
            if derived and not _is_gap(derived):
                f["Country"] = derived
                warns.append(f"Country was blank; derived '{derived}' from Location '{loc}' (healed)")
        elif loc_us and ctry_foreign:
            warns.append(f"Location '{loc}' reads US but Country='{ctry}'")
        elif loc_foreign and ctry_us:
            warns.append(f"Location '{loc}' reads non-US but Country='{ctry}'")

    # 2b) Last Updated must never predate Received. A row cannot have been updated before
    # it arrived; when it reads that way the timestamp is wrong, and anything reasoning
    # about recency (merge tie-breaks, the "did they reply after we asked?" check) silently
    # draws the wrong conclusion. Live case APP-20260716-0617-CB73 (Yashvi Vaghela):
    # Last Updated 2026-07-15T23:04 against Received 2026-07-16T06:17, flagged by the P1
    # audit every run since 2026-07-31. Clamp forward to Received, the one timestamp P1
    # owns and stamps at intake.
    _rec_raw, _upd_raw = f.get("Received Date"), f.get("Last Updated Date")
    if not _is_gap(_rec_raw) and not _is_gap(_upd_raw):
        _rec_dt, _upd_dt = _parse_cell_datetime(_rec_raw), _parse_cell_datetime(_upd_raw)
        if _rec_dt and _upd_dt and _upd_dt < _rec_dt:
            f["Last Updated Date"] = _rec_raw
            warns.append(f"Last Updated Date '{_upd_raw}' predated Received Date "
                         f"'{_rec_raw}' - clamped forward to Received")

    # 3) Phone plausibility — a non-gap phone should carry at least 7 digits.
    phone = str(f.get("Phone", "") or "").strip()
    if not _is_gap(phone) and sum(c.isdigit() for c in phone) < 7:
        warns.append(f"Phone '{phone}' has too few digits to be a real number")

    # 4) Category must never be blank (assign_category guarantees this; belt-and-suspenders).
    if _is_gap(f.get("Category", "")):
        f["Category"] = "General"
        warns.append("Category was blank; set to 'General' (healed)")

    # 5) Status must be a recognized value.
    status = str(f.get("Status", "") or "").strip()
    if status and status not in (
        STATUS_SCORED, STATUS_LOCATION_REVIEW,
        _cfg.STATUS_REJECTED, _cfg.STATUS_LOCATION_UNCONFIRMED,
    ):
        warns.append(f"Status '{status}' is not a recognized value")

    # 6) Portfolio slots — each is either 'N/A' or a URL-shaped string.
    for slot in ("Portfolio 1", "Portfolio 2", "Portfolio 3"):
        v = str(f.get(slot, "") or "").strip()
        if v and v.upper() != "N/A" and "." not in v and "://" not in v:
            warns.append(f"{slot} '{v}' doesn't look like a URL")

    return f, warns


def _validate_all_columns(fields: dict, p1_vals: dict, mail_body: str,
                          before: dict | None = None, after_extract: dict | None = None) -> dict:
    """Single column-by-column validation pass across the configured workbook columns.

    Called exactly once per candidate — after all extraction stages (offline parser,
    mail-body fallback, Ollama AI recheck) and role scoring are complete, right before
    the row is written to SharePoint. Ollama has already run; this pass never calls it
    again and never re-runs the mail-body fallback (those already ran in the main flow).

    What this pass does:
      P1/read-only columns        : read from the stored row, log OK or N/A (no writes).
      P2 scored cols (5)          : Role 1/2/3, Category, Status — accepted as-is.
      P2 extractable text cols (9): log their final state (OK or accepted N/A).
      Full Name last resort       : if still N/A, try the sender display name stored
                                    in P1's Email column (e.g. "Jane Smith <j@x.com>").

    `before`/`after_extract` (Step A's pre-extraction snapshot and the Tier 1/2 values
    right after extraction) are optional — when given, this also logs the Step E
    unchanged/filled/healed diff with a per-field method tag.
    """
    import re as _re
    from hiring_agent.config import COLUMNS
    from hiring_agent.extraction import resolve_full_name

    healed, accepted_na, ok = [], [], []

    # ── Last-resort Full Name: sender display name from P1's Email column ──────
    # P1 stores Email as "Jane Smith <jane@gmail.com>" or just "jane@gmail.com".
    # Stages 1-3 already ran; this is the only new data source available here.
    if _is_gap(fields.get("Full Name", "")):
        _raw_sender = str(p1_vals.get("Email", "") or "").strip()
        _m = _re.match(r'^([^<@]+?)\s*<[^>]+>$', _raw_sender)
        if _m:
            _display = _m.group(1).strip()
            _healed = resolve_full_name("Not extracted", _display, "")
            if not _is_gap(_healed):
                fields["Full Name"] = _healed
                healed.append("Full Name")
                logger.info(f"       Healed     : Full Name from sender display name → {_healed}")

    # ── Column-by-column status log ────────────────────────────────────────────
    for col in COLUMNS:
        if col in _P1_COLUMNS:
            v = str(p1_vals.get(col, "") or "").strip()
            (ok if not _is_gap(v) else accepted_na).append(col)
            continue

        # Scored columns. Category and Status must NEVER be blank - both have a
        # deterministic non-blank fallback, so a gap here is a real defect, not an
        # "accepted N/A". Fixed 2026-08-01: this branch used to mark all five columns
        # 'ok' unconditionally without looking at their values at all, which is exactly
        # how rows reached the Rejected sheet with a blank Category and blank Suggested
        # Roles (live cases: Syyed Nazir Ali / APP-20260727-1431-6F75 and Divy Parmar /
        # APP-20260720-1013-09AC) while the run still logged a clean column check.
        if col == "Category":
            if _is_gap(fields.get(col, "")):
                fields[col] = assign_category(
                    str(fields.get("Suggested Role 1", "") or ""),
                    str(fields.get("Current Skills", "") or ""))
                healed.append(col)
                logger.info(f"       Healed     : Category was blank → {fields[col]}")
            ok.append(col)
            continue

        if col == "Status":
            if _is_gap(fields.get(col, "")):
                logger.warning("       WARNING  : Status is blank after scoring - "
                               "this row would be invisible to the scoring queue.")
                accepted_na.append(col)
            else:
                ok.append(col)
            continue

        # Suggested Role 2/3 are legitimately blank when the scorer found fewer than
        # three matching JDs; Role 1 blank means role matching produced nothing at all.
        if col == "Suggested Role 1":
            if _is_gap(fields.get(col, "")):
                logger.warning("       WARNING  : Suggested Role 1 is blank - role "
                               "matching returned no JD match for this candidate.")
                accepted_na.append(col)
            else:
                ok.append(col)
            continue

        if col in ("Suggested Role 2", "Suggested Role 3"):
            (ok if not _is_gap(fields.get(col, "")) else accepted_na).append(col)
            continue

        # P2 extractable column — final state only; all extraction stages already ran.
        if _is_gap(fields.get(col, "")):
            accepted_na.append(col)
        else:
            ok.append(col)

    logger.info(f"       Col check  : {len(ok)} OK  |  "
                f"{len(healed)} healed  |  {len(accepted_na)} N/A")
    if len(healed) > 1:
        logger.info(f"       Healed     : {', '.join(healed)}")
    if accepted_na:
        logger.info(f"       Accepted NA: {', '.join(accepted_na)}")

    # Deterministic cross-field consistency audit + safe heals (never blocks the write).
    fields, _xwarns = _cross_field_check(fields, before)
    if _xwarns:
        logger.warning(f"       Consistency: {len(_xwarns)} check(s) flagged")
        for _w in _xwarns:
            logger.warning(f"         - {_w}")
    else:
        logger.info("       Consistency: OK (cross-field checks passed)")

    if before is not None:
        _log_column_diff(fields, before, after_extract or {})

    return fields


def _resume_display_path(resumes_folder: str, subpath: str) -> str:
    """Folder-only path for the plain-text 'Resume Folder Path' column, e.g.
    'HiringAgentP1Resume/2026/July' (or with a trailing 'Rejected' segment in `subpath` for a
    rejected candidate) — shows which dated bucket the file lives in, WITHOUT the filename.
    This is separate from the clickable 'Resume Link' column (a calculated column that reads
    P1's 'Application ID' + 'Original Filename' directly for the actual saved filename) — this
    column is plain text, never a hyperlink, purely informational."""
    folder_label = resumes_folder.strip("/").rsplit("/", 1)[-1]
    parts = [p for p in (folder_label, subpath.strip("/")) if p]
    return "/".join(parts)


def _resume_url_and_path(client, stored: list, subpath: str) -> tuple[str, str]:
    """Web URL of the candidate's primary (first) saved resume + the folder-only display
    path, or ('', '') on any failure.

    Stored as the plain 'Resume URL' (hidden) + 'Resume Folder Path' (visible, plain text,
    folder only) values. The 'Resume Link' calculated column is a separate clickable column
    that reads 'Resume URL' + P1's 'Application ID'/'Original Filename' directly — it does not
    use this function's path output at all. A missing URL is left blank (never a fake link).
    Pass `_rejected_subpath(subpath)` instead of `subpath` for a candidate that lives on the
    Rejected sheet."""
    from sharepoint_client import SharePointError
    if client is None or not stored:
        return "", ""
    for name in stored:
        try:
            url = client.resume_web_url(name, subfolder=subpath) or ""
        except SharePointError:
            continue
        if url:
            path = _resume_display_path(client.resumes_folder, subpath)
            return url, path
    return "", ""


def _rejected_subpath(subpath: str) -> str:
    """The dated subpath's YEAR-level 'Rejected' bucket, e.g. '2026/July' -> '2026/Rejected'.

    One 'Rejected' folder per year (changed 2026-07-15; was one per month, nested inside
    each month's own folder) - every declined candidate from that year lands in the same
    place regardless of which month they applied in, rather than scattered across a
    'Rejected' subfolder under every single month."""
    year = subpath.split("/", 1)[0]
    return f"{year}/Rejected"


def _move_resumes(client, stored: list, subpath: str, to_rejected: bool) -> bool:
    """Move a candidate's resume file(s) between a month's main folder and the year-level
    'Rejected' bucket to match which sheet they now belong on. Best-effort per file (one
    failure doesn't block the others) and idempotent (client.move_resume no-ops a file
    already at the destination). Returns True if anything actually moved."""
    if getattr(client, "local_execution", False):
        return True
    from sharepoint_client import SharePointError
    main_folder = f"{client.resumes_folder}/{subpath}"
    rejected_folder = f"{client.resumes_folder}/{_rejected_subpath(subpath)}"
    from_folder, to_folder = ((main_folder, rejected_folder) if to_rejected
                              else (rejected_folder, main_folder))
    moved_any = False
    for name in stored:
        try:
            if client.move_resume(name, from_folder, to_folder):
                moved_any = True
        except SharePointError as e:
            logger.warning(f"       WARNING  : could not move resume '{name}': {e}")
    return moved_any


def _get_cleaned_filename_prefix(full_name: str) -> str:
    """Generate a clean FirstNameLastName string from candidate's full name.
    If no last name, look for middle name, and if none only keep first name.
    Also strips any non-alphanumeric characters.

    Accents are folded to their base letter first (2026-08-03), so an accented Latin name
    keeps its spelling instead of losing letters: 'Jose Garcia' rather than 'JosGarca', and
    'Bjorn Aberg' rather than 'Bjrnberg'. A name in a script with no ASCII form at all
    (CJK, Cyrillic) still falls back to 'Candidate' - the filename stays unique because the
    Application ID tail is appended after the category (see _resume_filename_tail).
    """
    # 'MISSING' added 2026-08-04 alongside _MISSING_IF_BLANK_COLS: a nameless row now stores
    # the literal 'Missing', and without this guard the saved file would be 'Missing_7693.pdf'.
    if not full_name or str(full_name).strip().upper() in ("N/A", "NOT EXTRACTED",
                                                           "NOT PROVIDED", "MISSING"):
        return "Candidate"
    import unicodedata
    # NFKD splits 'é' into 'e' + combining acute; dropping the combining marks leaves 'e'.
    folded = "".join(c for c in unicodedata.normalize("NFKD", str(full_name))
                     if not unicodedata.combining(c))
    parts = [p.strip() for p in re.split(r'\s+', folded) if p.strip()]
    cleaned_parts = [re.sub(r'[^a-zA-Z0-9]', '', p) for p in parts]
    cleaned_parts = [p for p in cleaned_parts if p]
    if not cleaned_parts:
        return "Candidate"
    if len(cleaned_parts) == 1:
        return cleaned_parts[0]
    elif len(cleaned_parts) == 2:
        return cleaned_parts[0] + cleaned_parts[1]
    else:
        return cleaned_parts[0] + cleaned_parts[-1]


def _underscored_name_prefix(full_name: str) -> str:
    """'First_Last' for the 2026-08-04 filename convention, e.g. 'Sadaf_Khan'.

    Same name selection and accent folding as _get_cleaned_filename_prefix (first + last
    token, middle names dropped, accents folded to ASCII, 'Candidate' when nothing usable
    survives) - the ONLY difference is the underscore separator between the two parts, so
    the two functions can never disagree about WHICH names are used. A single-token name
    stays a single token ('Shivam'), which is why the Application ID tail is always appended
    by the caller: that, not the name, is what guarantees uniqueness.
    """
    joined = _get_cleaned_filename_prefix(full_name)
    if joined == "Candidate":
        return joined
    import unicodedata
    folded = "".join(c for c in unicodedata.normalize("NFKD", str(full_name or ""))
                     if not unicodedata.combining(c))
    parts = [re.sub(r'[^a-zA-Z0-9]', '', p) for p in re.split(r'\s+', folded) if p.strip()]
    parts = [p for p in parts if p]
    if len(parts) <= 1:
        return parts[0] if parts else "Candidate"
    return f"{parts[0]}_{parts[-1]}"


def _canonical_resume_name(full_name: str, app_id: str, ext: str, category: str) -> str:
    """The stored resume filename: '<First>_<Last>_<AppIdTail><ext>',
    e.g. 'Yash_Verma_7693.pdf'.

    Category is NOT part of the name (removed 2026-08-11, client instruction: "FN_LN_ID").
    History: dropped 2026-08-04, restored 2026-08-06, removed again here. The reason it keeps
    coming off is structural - Category is a P2 *verdict*, not a property of the document, so
    a re-score that moves a candidate between categories forces a physical file rename plus a
    Resume URL rewrite, and any file whose Category has since changed becomes unfindable by
    name. Dropping it makes the stored name depend only on identity (name + AppID tail), which
    never changes for a given candidate.

    The `category` parameter is KEPT even though the name no longer uses it: every caller
    already has it in scope, and _resume_name_slots still needs it to probe the superseded
    category-bearing shapes so files written before this change stay findable until
    migrate_resume_filenames.py has renamed them.
    """
    return (f"{_underscored_name_prefix(full_name)}"
            f"_{_resume_filename_tail(app_id)}{ext}")


def _clean_category_for_filename(category: str) -> str:
    """PascalCase, filename-safe version of a Category value, e.g. 'Data Analytics' ->
    'DataAnalytics', 'AI/ML/CV (SIN2)' -> 'AIMLCVSIN2', and
    'Mobile Apps (Android IOS)' -> 'MobileAppsAndroidIOS'."""
    if _is_gap(category):
        category = "General"
    cat_str = str(category).strip()
    tokens = [t for t in re.split(r'[^a-zA-Z0-9]+', cat_str) if t]
    cleaned = [t if t.isupper() else t[:1].upper() + t[1:] for t in tokens]
    return "".join(cleaned) or "General"


def _legacy_clean_category_for_filename(category: str) -> str:
    """Pre-2026-07-21 filename token, retained only to find already-saved files."""
    if _is_gap(category):
        category = "General"
    without_parentheses = re.sub(r'\(.*?\)', '', str(category)).strip()
    tokens = [t for t in re.split(r'[^a-zA-Z0-9]+', without_parentheses) if t]
    cleaned = [t if t.isupper() else t[:1].upper() + t[1:] for t in tokens]
    return "".join(cleaned) or "General"


def _resume_filename_tail(app_id: str) -> str:
    """Short unique disambiguator for the saved-resume filename: the last '-'-delimited
    segment of the Application ID (its random hex tail, e.g. 'A5F2' from
    'APP-20260710-2200-A5F2'). Two candidates can share a name AND a category, so the
    filename alone (FirstNameLastName_Category) isn't guaranteed unique - this tail is.
    Reuses AppRef's own uniqueness guarantee instead of inventing a second one, and stays
    in sync automatically if P1's appref.hex_length config ever changes."""
    app_id = str(app_id or "").strip()
    return app_id.rsplit("-", 1)[-1] if "-" in app_id else app_id


def _attachment_slug(original_filename: str) -> str:
    """Mirror of HiringAgent_P1/flow/build_zip.py::_attach_slug - keep the two IDENTICAL.

    P1 appends this to the saved filename ONLY when one email supplied more than one kept
    attachment, so that two files for the same row stop resolving to the same name and
    silently overwriting each other. Definition: lowercase the whole original filename,
    delete space _ - ' , ( ) and . , then take the first 16 characters. The dot is deleted
    rather than split on, so a name like 'my.name.resume.pdf' needs no special handling on
    either side.
    """
    s = str(original_filename or "").lower()
    for ch in (" ", "_", "-", "'", ",", "(", ")", "."):
        s = s.replace(ch, "")
    return s[:16]


def _resume_name_slots(app_id: str, resume_filename_cell: str, full_name: str = None,
                       category: str = None) -> list[list[str]]:
    """Reconstruct the candidate name(s) for each real attachment, one ordered list per
    attachment ('slot' — comma-separated 'Original Filename' entries are genuinely
    different files, e.g. resume + transcript, and must stay independent).

    Within a slot, order matters: P1 saves EVERY resume write — the candidate's first
    submission, a Duplicate resend, or a ref-quoted Update — under the SAME legacy shape
    '<FirstLast>_<FullAppID>.ext' (see HiringAgent_P1/flow/build_zip.py). P2 is the only
    thing that ever renames a file OUT of that shape, into '<First>_<Last>_<Category>_<tail>.ext',
    and only right after scoring it. So a legacy-shape file existing on a row that's ALREADY
    been scored once can only mean one thing: P1 just wrote a fresh update, and the old
    canonical-shape file is now stale leftover content. Checking legacy-shape FIRST means a
    fresh update is always preferred over stale prior content - never combined with it - and
    still correctly falls through to the canonical name for a normal, not-just-updated row
    (legacy 404s, canonical exists). See _download_resume_text, which uses this ordering to
    pick exactly one winner per slot instead of downloading+concatenating every guess."""
    import os
    file_id = app_id.replace("/", "-")
    tail = _resume_filename_tail(app_id)
    have_name = full_name and str(full_name).strip().upper() not in (
        "", "N/A", "NOT EXTRACTED", "NOT PROVIDED", "MISSING")
    slots = []
    _originals = [o.strip() for o in str(resume_filename_cell or "").split(",") if o.strip()]
    # P1 only disambiguates when the email really carried more than one kept attachment; a
    # single-attachment row keeps the plain '<FirstLast>_<AppID>.ext' shape. Mirror that
    # condition exactly, or every ordinary row would waste a probe on a name P1 never writes.
    _multi = len(_originals) > 1
    for original in _originals:
        _, ext = os.path.splitext(original)
        candidates = []
        # Legacy/always-current-write shape first - see docstring above.
        if have_name:
            name_part = _get_cleaned_filename_prefix(full_name)
            if _multi:
                # Written by P1 2026-08-27+ for the 2nd..nth attachment of one email.
                # Probed BEFORE the undisambiguated name: on a multi-attachment row the
                # plain name belongs to at most one of the slots, and letting every slot
                # match it would download the same file N times and concatenate it with
                # itself.
                candidates.append(
                    f"{name_part}_{file_id}_{_attachment_slug(original)}{ext}")
            candidates.append(f"{name_part}_{file_id}{ext}")
        # Canonical post-scoring shape (2026-08-11+): '<First>_<Last>_<tail>'. This is also
        # exactly what the 2026-08-04..08-06 window wrote, so one probe covers both.
        if have_name:
            candidates.append(_canonical_resume_name(full_name, app_id, ext, category))
        # SUPERSEDED post-scoring shape (2026-08-06 .. 2026-08-11): '<First>_<Last>_<Category>_<tail>'.
        # This was the canonical name until Category was removed on 2026-08-11, so EVERY file
        # scored in that window carries it. It used to be produced by _canonical_resume_name
        # itself; now that the canonical form has no Category, it must be probed EXPLICITLY or
        # every one of those files reads as missing and its row gets flagged unreadable.
        # Probed until migrate_resume_filenames.py has renamed them.
        if have_name and not _is_gap(category):
            candidates.append(
                f"{_underscored_name_prefix(full_name)}"
                f"_{_clean_category_for_filename(category)}_{tail}{ext}")
            _old_cat = _legacy_clean_category_for_filename(category)
            if _old_cat != _clean_category_for_filename(category):
                candidates.append(
                    f"{_underscored_name_prefix(full_name)}_{_old_cat}_{tail}{ext}")
        # SUPERSEDED post-scoring shape (2026-07-15 .. 2026-08-04): '<FirstLast>_<Category>_<tail>'.
        # Still probed so a file saved under the old convention - and not yet migrated - is
        # found rather than treated as missing (which would flag the row as unreadable).
        if have_name and not _is_gap(category):
            name_part = _get_cleaned_filename_prefix(full_name)
            cat_part = _clean_category_for_filename(category)
            candidates.append(f"{name_part}_{cat_part}_{tail}{ext}")
            old_cat_part = _legacy_clean_category_for_filename(category)
            if old_cat_part != cat_part:
                candidates.append(f"{name_part}_{old_cat_part}_{tail}{ext}")
        # A historical repair may have stored the real P2 canonical filename in the
        # Original Filename cell. Its _<AppIdTail>.<ext> suffix proves it is already a
        # stored name, so try it verbatim rather than creating an impossible AppID prefix.
        already_stored = bool(
            file_id in original
            or (tail and re.search(rf"_{re.escape(tail)}\.[^.]+$", original, re.I))
        )
        candidates.append(original if already_stored else f"{file_id}_{original}")
        slots.append(list(dict.fromkeys(candidates)))
    return slots


def _resume_name_from_url(resume_url: str) -> str:
    """Return the decoded filename from a stored SharePoint Resume URL.

    Graph hands back a DIFFERENT webUrl shape per file type, and only the PDF one carries
    the filename in the path:
        PDF  -> https://.../Candidate_Resumes/2026/July/Sadaf_Khan_7693.pdf
        DOCX -> https://.../_layouts/15/Doc.aspx?sourcedoc={GUID}&file=Peter_Vishal_6CE3.docx&action=...
    Reading only the path (the behaviour before 2026-08-04) therefore returned the literal
    string 'Doc.aspx' for EVERY Office document - it has a dot, so the old check accepted it
    as a filename. Callers then carried 'Doc.aspx' around as the candidate's resume name:
    _move_resumes tried to move it, the duplicate-merge loser cleanup tried to DELETE it, and
    _finish_rejection treated it as the superseded file to remove. Those all 404'd harmlessly
    against a name that never existed, but it also meant the real .docx was never the linked
    name and was only ever found via the fallback slot guesses. The 'file=' query parameter
    is checked first so Office documents resolve to their true filename.
    """
    from urllib.parse import unquote, urlsplit, parse_qs

    raw = str(resume_url or "").strip()
    if not raw:
        return ""
    try:
        parts = urlsplit(raw)
        qs_name = (parse_qs(parts.query).get("file") or [""])[0].strip()
        if qs_name and "." in qs_name:
            return unquote(qs_name)
        name = unquote(parts.path.rsplit("/", 1)[-1]).strip()
    except Exception:
        return ""
    if not name or "." not in name:
        return ""
    # A bare viewer endpoint ('Doc.aspx', 'WopiFrame.aspx') is the PAGE, not the document.
    if name.lower().endswith(".aspx"):
        return ""
    return name


def _stored_resume_names(app_id: str, resume_filename_cell: str, full_name: str = None,
                         category: str = None, resume_url: str = None) -> list[str]:
    """Flat, deduplicated view of every candidate name across all slots - for callers that
    intentionally want to try/move/delete every guess regardless of which one is 'current'
    (e.g. _move_resumes, the dedup-merge loser cleanup). Extraction should use
    _download_resume_text instead, which picks one winner per slot - see its docstring.

    The filename already present in Resume URL is tried first. It remains valid even
    when Full Name or Category changed after the physical file was last renamed.
    """
    linked_name = _resume_name_from_url(resume_url)
    names = [linked_name] if linked_name else []
    for slot in _resume_name_slots(app_id, resume_filename_cell, full_name, category):
        names.extend(slot)
    return list(dict.fromkeys(names))


def _resume_file_exists(client, name: str, subpath: str = "") -> bool:
    """Read-only existence check used to keep rename logs truthful."""
    from sharepoint_client import SharePointError
    try:
        return bool(client.resume_web_url(name, subfolder=subpath))
    except SharePointError:
        return False


def _rename_scored_resumes(client, stored: list[str], subpath: str, app_id: str,
                           full_name: str, category: str, dry_run: bool) -> list[str]:
    """Rename scored resume files to the canonical P2 shape, with non-noisy logs.

    Multiple attachment slots can legitimately collapse to the same canonical name
    after scoring. When that target is already present, treat it as represented
    instead of warning that a later duplicate rename failed.
    """
    if getattr(client, "local_execution", False):
        return stored
    import os
    if not full_name:
        return stored
    new_stored = []
    for name in stored:
        _, ext = os.path.splitext(name)
        new_name = _canonical_resume_name(full_name, app_id, ext, category)
        if name == new_name:
            new_stored.append(name)
            continue
        if new_name in new_stored:
            logger.info(f"       Rename    : skipped '{name}' because canonical "
                        f"target '{new_name}' is already represented.")
            continue
        if dry_run:
            new_stored.append(new_name)
            continue
        try:
            if client.rename_resume(name, new_name, subfolder=subpath):
                new_stored.append(new_name)
            elif _resume_file_exists(client, new_name, subpath):
                logger.info(f"       Rename    : target '{new_name}' already exists; "
                            f"using it instead of '{name}'.")
                new_stored.append(new_name)
            else:
                logger.warning(f"       WARNING  : Could not rename SharePoint file "
                               f"'{name}' to '{new_name}'")
                new_stored.append(name)
        except Exception as e:
            logger.warning(f"       WARNING  : Error renaming SharePoint file "
                           f"'{name}' to '{new_name}': {e}")
            new_stored.append(name)
    return new_stored


def _rename_resume_for_healed_row(client, heal: dict, vals: dict, app_id: str,
                                  subpath: str, dry_run: bool) -> None:
    """When a heal pass changes Category or Full Name, the resume file's canonical name
    is derived from those same fields (see _rename_scored_resumes) - without this, the
    Category/Full Name cell can drift from the name still baked into the resume filename
    and Resume URL, since a heal-only patch otherwise never touches the physical file.
    Re-renames the stored resume(s) and refreshes Resume URL/Folder Path in `heal` (in
    place) to match. No-op if neither field is part of this heal.
    """
    if "Category" not in heal and "Full Name" not in heal:
        return
    stored = _stored_resume_names(
        app_id, vals.get("Original Filename", ""),
        full_name=vals.get("Full Name"), category=vals.get("Category"),
        resume_url=vals.get("Resume URL"))
    if not stored:
        return
    full_name = heal.get("Full Name", vals.get("Full Name"))
    category = heal.get("Category", vals.get("Category"))
    new_stored = _rename_scored_resumes(client, stored, subpath, app_id, full_name,
                                        category, dry_run)
    if new_stored != stored:
        url, path = _resume_url_and_path(client, new_stored, subpath)
        if url:
            heal["Resume URL"] = url
            heal["Resume Folder Path"] = path


def _text_fingerprint(text: str) -> frozenset:
    """Word set of `text`, lowercased and stripped of everything but alphanumerics.

    Deliberately a SET, not the raw string: the same document exported as .pdf and as .docx
    almost never extracts byte-identically - line breaks, ligatures, bullet glyphs and column
    order all differ - but the vocabulary is the same. Comparing word sets sees through all
    of that, and is O(n) rather than a similarity search over two multi-thousand-word blobs.
    """
    cleaned = "".join(c if c.isalnum() else " " for c in str(text or "").lower())
    return frozenset(w for w in cleaned.split() if len(w) > 2)


def _is_same_document(a: frozenset, b: frozenset, threshold: float = 0.90) -> bool:
    """True when two extracted texts are the same document in different formats.

    Jaccard overlap of the two word sets. 0.90 is deliberately high: a candidate who really
    did attach two DIFFERENT documents (a resume and a genuine portfolio write-up) shares
    their name, their employers and their skills across both, which lands well under this -
    while the same file exported twice lands at 0.97+. Erring high means the failure mode is
    "scored both, as before", never "silently discarded a real second document".
    """
    if not a or not b:
        return False
    union = len(a | b)
    return bool(union) and (len(a & b) / union) >= threshold


class ResumeUnavailableError(RuntimeError):
    """The resume could not be obtained; this is not an unreadable applicant file."""


def _download_resume_text(client, app_id: str, resume_filename_cell: str, subpath: str,
                          full_name: str = None, category: str = None,
                          resume_url: str = None, explicit_names=None
                          ) -> tuple[str, list[str], dict]:
    """Download and extract text for one candidate, choosing exactly one file per real
    attachment slot (the first name in each slot's priority list that actually downloads -
    see _resume_name_slots) instead of blindly downloading+concatenating every guess.

    Returns (combined_text, used_names, raw_by_name) - used_names is which actual filename
    won each slot (for the caller to rename/keep, and raw_by_name lets a caller that wants a
    local copy avoid re-downloading). Any other guess left over for a slot (typically a stale
    pre-update canonical-shape file once a legacy-shape update file wins) is NOT included
    here and should be cleaned up by the caller once the new content is safely scored.

    Raises ResumeUnavailableError if any attachment slot cannot be downloaded.
    An empty text result with downloaded bytes still means a parsing/OCR failure.
    """
    from sharepoint_client import SharePointError
    from requests import RequestException
    used, raw_by_name, texts = [], {}, []
    slots = ([[name] for name in explicit_names] if explicit_names else
             _resume_name_slots(app_id, resume_filename_cell, full_name, category))
    linked_name = _resume_name_from_url(resume_url)
    if linked_name and not explicit_names:
        if slots:
            slots[0] = list(dict.fromkeys([linked_name, *slots[0]]))
        else:
            slots = [[linked_name]]
    # An applicant who attaches the SAME resume as both .pdf and .docx produces two slots
    # whose text is the same document. Concatenating them scored every skill, employer and
    # date twice, which skews extraction and the AI recheck (2026-08-27). The file itself is
    # still recorded in `used`/`raw_by_name` - it genuinely exists on disk and the caller
    # still has to rename, move or clean it up - only its TEXT is left out of the blend.
    _seen: list[frozenset] = []
    for slot in slots:
        last_error = None
        for name in slot:
            try:
                raw = client.download_resume(name, subfolder=subpath)
            except SharePointError as e:
                if e.status_code not in (None, 404):
                    raise ResumeUnavailableError(f"{app_id}: resume download failed: {e}") from e
                last_error = e
                continue
            except RequestException as e:
                raise ResumeUnavailableError(f"{app_id}: resume download failed: {e}") from e
            used.append(name)
            raw_by_name[name] = raw
            body = extract_text_from_bytes(raw, name)
            fp = _text_fingerprint(body)
            if any(_is_same_document(fp, prev) for prev in _seen):
                logger.info(f"       DEDUP    : {name} is the same document as an attachment "
                            f"already scored for {app_id} - kept on disk, text not counted twice")
            else:
                if fp:
                    _seen.append(fp)
                texts.append(body)
            break
        else:
            raise ResumeUnavailableError(
                f"{app_id}: no resume file downloaded in {subpath!r} or the legacy root; "
                f"tried {slot}. Last download error: {last_error}"
            ) from last_error
    text = "\n".join(t for t in texts if t).strip()
    return text, used, raw_by_name


# Identity fields: WHO the candidate is and WHERE they are. The covering email may fill
# these when the CV is silent, but it must never contradict a value the CV actually stated.
_CV_IDENTITY_FIELDS = ("full_name", "location", "country")


def _restore_cv_identity(details: dict, cv_identity: dict, app_id: str = "") -> dict:
    """Re-assert CV-derived identity after the mail-body merge and the AI recheck.

    Both of those stages read the covering EMAIL, which is only the same person as the
    candidate when the applicant sends their own resume. Added 2026-08-01 after a staffing
    agency's submission was scored entirely as the agent: the CV said 'Shivam / Austin,
    Texas' but the row came out 'Vandana Asawa / Plano, Texas' (the agent's name and the
    agency's city, lifted from her signature) because ai_recheck_fields saw the mail body
    and "corrected" the CV's answer to match the sender.

    Agency and vendor submissions are accepted as normal candidates whenever a real resume
    is attached, so the resume - never the covering email - is the authority on identity.
    A field the CV left blank is still fair game for the email to fill: that is the common
    case where the applicant IS the sender and their signature carries the missing detail.
    """
    out = dict(details)
    for field in _CV_IDENTITY_FIELDS:
        cv_val = str(cv_identity.get(field, "") or "").strip()
        if not cv_val or cv_val.lower() in _GAP_LITERALS:
            continue                      # CV never stated it - let the email's value stand
        cur = str(out.get(field, "") or "").strip()
        if cur.lower() != cv_val.lower():
            out[field] = cv_val
            logger.info(f"       CV wins   : {field} '{cur}' -> '{cv_val}' "
                        f"(resume overrides covering email)")
    return out


def _complete_row_before_reject(merged: dict, app_id: str = "") -> dict:
    """Last-chance completeness gate: no row reaches the Rejected sheet half-filled.

    Added 2026-08-01 after two live rows (Syyed Nazir Ali / APP-20260727-1431-6F75 and
    Divy Parmar / APP-20260720-1013-09AC) landed on the Rejected sheet with a blank
    Category and blank Suggested Roles. The Rejected sheet is terminal - nothing re-scores
    a row once it is there - so a gap that slips through here is permanent and silent.

    Every rejection path funnels through _finish_rejection, so this single chokepoint
    covers them all. Deterministic repairs only (no network, no LLM): Category is
    re-derived via assign_category, which is guaranteed non-blank. Anything still missing
    is logged loudly rather than silently accepted, so a real gap is visible in the run
    log instead of only being noticed weeks later on the sheet.
    """
    if _is_gap(merged.get("Category")):
        merged["Category"] = assign_category(
            str(merged.get("Suggested Role 1", "") or ""),
            str(merged.get("Current Skills", "") or ""))
        logger.info(f"       Pre-reject : Category was blank → {merged['Category']}")

    for _num_col in ("Application Updates", "Retry Count"):
        if _is_gap(merged.get(_num_col)):
            merged[_num_col] = 0

    for _na_col in ("Portfolio 1", "Portfolio 2", "Portfolio 3"):
        if _is_gap(merged.get(_na_col)):
            merged[_na_col] = "N/A"

    _still_missing = [c for c in ("Full Name", "Location", "Country", "Status",
                                  "Current Skills", "Suggested Role 1")
                      if _is_gap(merged.get(c))]
    if _still_missing:
        logger.warning(f"       Pre-reject : {app_id or '(no ref)'} moving to Rejected with "
                       f"{len(_still_missing)} unfilled column(s): {', '.join(_still_missing)}")
    return merged


def _finish_rejection(client, index: int, app_id: str, vals: dict, fields: dict,
                      stored: list, subpath: str, dry_run: bool, rej_ids: set,
                      rej_emails: set, rej_phones: set, reason_label: str,
                      rejected_rows: list | None = None) -> bool:
    """Move one row to Rejected or refresh a matching Rejected record in place.

    A repeat applicant who remains non-US keeps one row and one decline marker, while
    newer confirmed details and the newer resume replace the stale stored version.
    Non-USA is currently the only rejection reason; the caller sets ``fields['Status']``.

    Returns True iff a row was actually deleted from the main sheet (i.e. not a
    dry-run) - callers use that to decide whether to bump their own processed/
    rejected/deleted_idx bookkeeping, since dry-run makes no changes worth counting."""
    email_key = _bare_email(vals.get("Email", ""))
    phone_key = _normalize_phone_key(vals.get("Phone", ""))
    already_rejected = ((app_id and app_id in rej_ids)
                        or (email_key and email_key in rej_emails)
                        or (phone_key and phone_key in rej_phones))
    if dry_run:
        # A caller that PRE-STAMPS the decline column has already decided this row gets no
        # email - the age-out pass does exactly that (_AGED_OUT_DECLINE_STAMP), because
        # silence from a candidate is not evidence they are non-US. Reporting "would queue
        # the decline email" for those rows was flatly wrong and alarming: a 2026-08-24
        # dry-run appeared to say 10 real June candidates were about to be sent rejections
        # that the design in fact guarantees are never sent. Read the field instead of
        # assuming, so the preview matches what the live run would actually do.
        _prestamped = str((fields or {}).get(_DECLINE_COL, "") or "").strip()
        if already_rejected:
            _tail = (" — duplicate of an existing Rejected row; would "
                     "refresh that record (no second decline).")
        elif _prestamped:
            _tail = f" — silent close, NO decline email (Decline Sent pre-stamped {_prestamped!r})."
        else:
            _tail = " + queue the decline email."
        logger.info(f"       Result   : [DRY-RUN] Would reject ({reason_label})" + _tail)
        return False
    if already_rejected:
        # A clarification reply can carry a newer resume and better contact/location
        # fields even when the final decision remains non-US. Refresh the existing
        # Rejected row before removing Main, preserving its ID and decline marker.
        rejected_rows = rejected_rows or []
        matching = next(
            (r for r in rejected_rows
             if app_id and str(r["values"].get("Application ID", "")).strip() == app_id),
            None,
        )
        if matching is None and email_key:
            matching = next(
                (r for r in rejected_rows
                 if _bare_email(r["values"].get("Email", "")) == email_key),
                None,
            )
        if matching is None and phone_key:
            matching = next(
                (r for r in rejected_rows
                 if _normalize_phone_key(r["values"].get("Phone", "")) == phone_key),
                None,
            )

        if matching is not None:
            old_vals = matching["values"]
            old_linked_name = _resume_name_from_url(old_vals.get("Resume URL"))
            # Only the NEWER application may replace the stored record. The refresh below
            # rewrites that record from the incoming row, which is right only when the
            # incoming row is the newer one - the docstring's "newer confirmed details and
            # the newer resume replace the stale stored version". Draining a backlog feeds
            # OLDER applications in after newer ones are already filed, and doing this
            # unconditionally inverted the intent: a July record kept its own Application ID
            # but took a June row's Received Date, Original Filename and resume, and its own
            # July resume was then deleted as "superseded". Confirmed live 2026-08-05 on
            # C8CC and 09AC, and it destroyed the newer resume in both cases.
            incoming_ts = _received_timestamp(vals.get("Received Date"))
            existing_ts = _received_timestamp(old_vals.get("Received Date"))
            incoming_is_newer = incoming_ts is not None and (
                existing_ts is None or incoming_ts > existing_ts)

            _move_resumes(client, stored, subpath, to_rejected=True)
            resume_url, resume_path = _resume_url_and_path(
                client, stored, _rejected_subpath(subpath)
            )
            if incoming_is_newer:
                refreshed = dict(vals)
                refreshed.update(fields)
                for col in ("Resume URL", "Resume Folder Path", "Resume Link",
                            _INFO_REQUEST_COL, "Mail Sent"):
                    refreshed.pop(col, None)
                existing_id = str(old_vals.get("Application ID", "") or "").strip()
                if existing_id:
                    refreshed["Application ID"] = existing_id
                if resume_url:
                    refreshed["Resume URL"] = resume_url
                    refreshed["Resume Folder Path"] = resume_path
                refreshed = _complete_row_before_reject(refreshed, app_id)
                _store(client).save_by_id(
                    _app_id_of(old_vals), refreshed, current_values=old_vals,
                    sheet="rejected", hint=matching["index"],
                )

                if old_linked_name and old_linked_name not in set(stored):
                    # CONTAINMENT CHANGE 2026-09-06 (docs/P1_P2_IMPLEMENTATION_STEPS.md
                    # Section 1): this used to delete the superseded file outright. The
                    # "incoming is newer" gate above already fixed one live incident where
                    # this logic ran with its sense inverted and destroyed the NEWER resume
                    # for both C8CC and 09AC (2026-08-05) - the gate makes deletion safer, not
                    # safe, since it is still an unrecoverable action driven by inference
                    # about which file is truly superseded. Preserve the file; a human-
                    # reviewed retention procedure removes it later, not this run.
                    logger.info(
                        f"       Cleanup   : superseded resume '{old_linked_name}' preserved "
                        f"(not deleted) - pending a reviewed retention procedure."
                    )
            else:
                # Older application. Contribute only what the stored record is MISSING, and
                # never the columns that belong to a row's own intake history - the same set
                # the duplicate merge refuses to carry across, plus the decline marker. The
                # stored record keeps its own dates, filename and resume, and nothing of its
                # is deleted; the incoming resume was still moved into the Rejected bucket
                # above, so an older submission is retained rather than lost.
                protected = _MERGE_EXCLUDE_COLS | {_cfg.DECLINE_SENT_COLUMN, "Status"}
                incoming_all = dict(vals)
                incoming_all.update(fields)  # confirmed geo may fill a blank on the record
                heal = {c: v for c, v in incoming_all.items()
                        if c not in protected
                        and not _is_gap(v) and _is_gap(old_vals.get(c, ""))}
                if heal:
                    _store(client).save_by_id(
                        _app_id_of(old_vals), heal, current_values=old_vals,
                        sheet="rejected", hint=matching["index"],
                    )
                logger.info(
                    f"       Note     : incoming {app_id} (received "
                    f"{vals.get('Received Date','?')}) is OLDER than the stored Rejected "
                    f"record {str(old_vals.get('Application ID','')).strip()} (received "
                    f"{old_vals.get('Received Date','?')}) - stored record kept"
                    + (f"; filled its blank {', '.join(sorted(heal))}" if heal else
                       "; nothing to fill")
                )
        _store(client).delete_by_id(_row_key(app_id, vals), hint=index)
        logger.info(
            f"       Result   : REJECTED (repeat applicant, {reason_label}) — "
            + ("existing Rejected record refreshed"
               if matching is not None else "existing Rejected record retained")
            + "; removed Main copy, no second decline email."
        )
        return True
    _move_resumes(client, stored, subpath, to_rejected=True)
    resume_url, resume_path = _resume_url_and_path(client, stored, _rejected_subpath(subpath))
    if resume_url:
        fields["Resume URL"] = resume_url
        fields["Resume Folder Path"] = resume_path
    merged = dict(vals)
    merged.setdefault("Application Updates", 0)
    merged.setdefault("Retry Count", 0)
    merged.update(fields)
    if _is_gap(merged.get("Application Updates")):
        merged["Application Updates"] = 0
    if _is_gap(merged.get("Retry Count")):
        merged["Retry Count"] = 0
    merged = _complete_row_before_reject(merged, app_id)
    _ensure_rejected_period_separator(client, merged.get("Received Date"))
    # The candidate reaching the Rejected sheet and the main copy actually going away are two
    # Graph calls, and the delete is the one that fails (409 InsertDeleteConflict). Returning
    # True regardless reported a completed rejection AND made the caller record a phantom
    # entry in deleted_idx, skewing the position hint for every later row in the batch.
    moved = _store(client).move_to_rejected(_row_key(app_id, vals), merged, hint=index)
    if app_id:
        rej_ids.add(app_id)
    if email_key:
        rej_emails.add(email_key)
    if phone_key:
        rej_phones.add(phone_key)
    if not moved:
        logger.warning(
            f"       PARTIAL  : {app_id} was added to the Rejected sheet but its main-sheet "
            f"row could not be removed - it is on BOTH sheets until that row is cleared."
        )
        return False
    logger.info(f"       Result   : REJECTED ({reason_label}) — moved to Rejected sheet "
                "(decline email queued for the send pass).")
    return True


def _live_row_index(client, row: dict, deleted_idx: list, app_id: str):
    """The CURRENT position of a row read earlier in this batch, or None if it is gone.

    Row identity now lives in `hiring_agent/store.py`; this is the thin shim the batch loop
    still uses because it wants a position to log and to record in `deleted_idx`. Every
    actual write goes through the store by Application ID and never touches an index.

    The offset arithmetic below is now a HINT, not a decision. It used to be load-bearing:
    the batch was read once and written back by position, so a delete path that forgot to
    register itself in `deleted_idx` produced a shift of zero and pasted each candidate's
    results onto the NEXT candidate's row (fixed 2026-08-03). The store re-checks whatever
    this computes and re-locates by Application ID when it is wrong, so a forgotten delete
    path now costs one recovery scan instead of corrupting two rows.

    A good hint still earns its keep: it is the tie-break when two rows share an Application
    ID — a stale `Needs Review` retry beside the completed row for the same candidate — where
    the key alone cannot say which row the caller meant.

    Returns None when the candidate cannot be found at all, so the caller SKIPS it rather
    than writing somewhere unknown. A skipped row keeps its status and is picked up next run.
    """
    if getattr(client, "local_execution", False):
        return _store(client).resolve_position(app_id, "main", row.get("index"))
    cached = row.get("index")
    hint = None
    if cached is not None:
        hint = cached - sum(1 for d in deleted_idx if d < cached)
    if not app_id:
        return hint
    try:
        return _store(client).resolve_position(app_id, "main", hint)
    except RowVanished:
        return None

def _unchanged_completed_retry(retry: dict, completed: dict) -> bool:
    """Only skip a stale retry when identity AND application version still agree.

    A new submission, changed attachment, or clarification must keep its place in
    the queue even when an older copy of the same Application ID is completed.
    Missing version information is not proof of a duplicate.
    """
    if str(retry.get("Status", "")).strip() != _cfg.STATUS_NEEDS_REVIEW:
        return False
    if str(completed.get("Status", "")).strip() not in (
            STATUS_SCORED, "Rejected - Non-USA Location"):
        return False
    app_id = _app_id_of(retry)
    email = _bare_email(retry.get("Email", ""))
    if not app_id or app_id != _app_id_of(completed) or not email:
        return False
    if email != _bare_email(completed.get("Email", "")):
        return False
    for key in ("Received Date", "Last Updated Date", "Original Filename"):
        value = str(retry.get(key, "") or "").strip()
        if not value or value != str(completed.get(key, "") or "").strip():
            return False
    updates = retry.get("Application Updates")
    return (updates is not None and completed.get("Application Updates") is not None
            and str(updates).strip() == str(completed["Application Updates"]).strip())


def _scoring_queue_rows(client) -> list:
    """Return only new/unscored work plus genuinely updated location-review rows.

    P1 normally resets an updated row to ``New Email Received``. The timestamp check
    is a narrow fallback for a clarification reply that updated the row but left its
    location-review status unchanged. Untouched scored/rejected/review rows never enter.
    """
    rows = list(client.list_unscored_rows((_STATUS_NEW, _cfg.STATUS_NEEDS_REVIEW)))
    main_rows = []
    try:
        main_rows = client.list_rows()
        review_updates = [
            row for row in main_rows
            if str(row["values"].get("Status", "")).strip() == STATUS_LOCATION_REVIEW
            and _updated_after_location_request(row["values"])
        ]
    except (AttributeError, TypeError):
        review_updates = []

    seen = {
        (row.get("index"), str(row.get("values", {}).get("Application ID", "")).strip())
        for row in rows
    }
    for row in review_updates:
        key = (
            row.get("index"),
            str(row.get("values", {}).get("Application ID", "")).strip(),
        )
        if key not in seen:
            rows.append(row)
            seen.add(key)

    # A failed Excel deletion can leave a retry beside its completed copy. Exclude
    # that copy before applying the batch limit; this read-only check never deletes
    # a row and never suppresses a new/updated application.
    if any(str(r.get("values", {}).get("Status", "")).strip() == _cfg.STATUS_NEEDS_REVIEW
           for r in rows):
        try:
            completed_rows = main_rows + client.list_rejected_rows()
        except Exception as e:
            logger.warning(f"  Queue     : could not verify completed duplicates ({e}); "
                           "keeping retry rows eligible.")
        else:
            eligible = []
            for row in rows:
                vals = row.get("values", {})
                if any(_unchanged_completed_retry(vals, r.get("values", {}))
                       for r in completed_rows):
                    logger.info(f"  Queue     : skipping unchanged retry {_app_id_of(vals)}; "
                                "the same application already completed. Rows unchanged.")
                else:
                    eligible.append(row)
            rows = eligible
    return rows


def _progress_value(value) -> str:
    """Keep progress fields readable and safe for the app's pipe-delimited parser."""
    text = "" if value is None else str(value)
    return re.sub(r"\s+", " ", text.replace("|", "/")).strip()


def _log_p2_progress(state: str, **fields) -> None:
    """Emit one progress line shared by CMD and the desktop app."""
    parts = [f"state={_progress_value(state)}"]
    parts.extend(f"{key}={_progress_value(value)}" for key, value in fields.items())
    logger.info("  P2 Progress | " + " | ".join(parts))


def score_from_sharepoint(dry_run: bool = False, scorecards: bool = False, *, _local_client=None) -> dict:
    """Score every unscored candidate row in the SharePoint workbook. Returns a summary."""
    from sharepoint_client import SharePointClient, SharePointError, check_graph_reachable

    if _env_truthy("HIRING_P2_DISABLED"):
        logger.warning("  DISABLED  P2 scoring is disabled by HIRING_P2_DISABLED.")
        return {
            "disabled": True, "processed": 0, "rejected": 0,
            "location_review": 0, "errors": 0,
        }

    if _local_client is None and storage_backend() == "sqlite":
        candidate_client = _local_client or SharePointClient()
        if hasattr(candidate_client, "site_id"):
            from .local_pipeline import run_local_pipeline
            return run_local_pipeline(candidate_client, dry_run=dry_run)

    logger.info("  Step 1/4  Checking connectivity...")
    _net_ok, _net_msg = (True, "Local processing") if _local_client else check_graph_reachable()
    if not _net_ok:
        logger.error(f"  FAILED    {_net_msg}")
        return {"error": _net_msg}
    logger.info(f"  OK        {_net_msg}")
    _ok, _brain = ollama_health()
    (logger.info if _ok else logger.warning)(f"  {_brain}")
    if _cfg.REQUIRE_AI and not _ok:
        # require_ai means a row is never published from anything but a real model read.
        # With the brain down EVERY row would degrade, so proceeding would write a whole
        # batch of regex output. Stop instead: nothing is written, the queue keeps its
        # place, and the next run picks it all up unchanged.
        _msg = (f"Ollama is required but unavailable ({_brain}). Nothing was scored and no "
                f"row was changed - the queue is untouched. Start Ollama and re-run, or set "
                f"ai_extraction.ollama.require_ai=false to allow offline-parser fallback.")
        logger.error(f"  FAILED    {_msg}")
        return {"error": _msg, "processed": 0, "rejected": 0,
                "location_review": 0, "errors": 0, "deferred": 0}

    from hiring_agent import extraction as _extraction_mod
    _extraction_mod.STRICT_AI_STAGES = bool(_cfg.REQUIRE_AI)

    try:
        client = _local_client or SharePointClient()
    except SharePointError as e:
        logger.error(f"  FAILED    SharePoint not configured: {e}")
        return {"error": str(e)}
    logger.info(f"  OK        Connected to {client.hostname}")

    logger.info("  Step 2/4  Checking workbook & columns...")
    if dry_run or _local_client is not None:
        # Dry-run must be side-effect-free: never create the workbook or add columns.
        logger.info("  SKIP      Dry-run: not creating/altering workbook or columns.")
    else:
        if not _ensure_workbook_or_alert(client, dry_run):
            return {"error": "candidate workbook could not be confirmed - see log; "
                              "P2 never auto-creates one"}
        _ensure_schema(client)

    logger.info("  Step 3/4  Reading unscored candidates...")
    try:
        queue_rows = _scoring_queue_rows(client)
        queue_total = len(queue_rows)
        batch_total = min(queue_total, _cfg.SCORING_BATCH_LIMIT)
        rows = queue_rows[:batch_total]
        _log_p2_progress(
            "queue",
            waiting=queue_total,
            batch=batch_total,
            after_batch=max(queue_total - batch_total, 0),
        )
        if queue_total > batch_total:
            logger.info(
                f"            Queue has {queue_total} candidates. Processing first "
                f"{batch_total} (batch limit); {queue_total - batch_total} will remain."
            )
    except SharePointError as e:
        logger.error(f"  FAILED    Could not read the candidate table: {e}")
        return {"error": str(e)}

    if not rows and _local_client is not None:
        return {"processed": 0, "errors": 0, "deferred": 0}
    if not rows:
        _log_p2_progress("empty", waiting=0, batch=0, remaining=0)
        logger.info("  OK        No unscored candidates — queue is empty.")
        # Runs even on an idle pass: an unanswered review row ages on the calendar, not on
        # how busy the queue happens to be. Before the mail passes, so a row closed here is
        # never also nudged or declined in the same run.
        aged = _age_out_stale_reviews(client, dry_run=dry_run)
        declines = _send_pending_declines(client, dry_run=dry_run)
        info_reqs = _send_pending_info_requests(client, dry_run=dry_run)
        export_path = ""
        # Extra layer (client instruction, 2026-09-01): even an idle/empty-queue pass
        # still re-checks every scored row for a missing Education date/Phone/Location it
        # can now recover from its resume - counted into _changed below, so an idle poll
        # that recovers something still publishes it instead of waiting for a busy run.
        _reevaluated = _reevaluate_missing_scored_fields(client) if not dry_run else 0
        # Only rebuild the client workbook if this pass actually changed something. In
        # --watch mode an idle poll would otherwise re-read the whole table and re-upload the
        # workbook every interval - 288 pointless uploads a day at 5-minute polling, against
        # the same Excel/Graph connector P1 is polling every minute, for a file whose content
        # is identical each time. Nothing changed means nothing to publish; use
        # `bot.py --export-results` for an unconditional on-demand rebuild.
        _changed = bool(declines or info_reqs or aged.get("closed") or _reevaluated)
        if not dry_run and _changed:
            try:
                export_path = export_client_results(client, upload_to_sharepoint=True)
                if export_path:
                    logger.info(f"  Export    : client workbook updated -> {export_path}")
            except Exception as e:
                logger.warning(f"  WARNING   Client workbook export failed: {e}")
        elif not dry_run:
            logger.info("  Export    : skipped (queue empty and nothing changed this pass).")
        return {
            "processed": 0, "rejected": 0, "location_review": 0, "errors": 0,
            "declines": declines, "info_requests": info_reqs,
            "aged_out": aged, "client_export": export_path,
        }
    logger.info(
        f"  OK        {queue_total} candidate(s) waiting; "
        f"{batch_total} selected for this run."
    )

    # Snapshot who is ALREADY on the Rejected sheet (by Application ID, email, and phone),
    # so a re-ingested candidate (e.g. an old backlog mail moved back to the inbox and
    # picked up again by Phase 1 under a NEW Application ID) is never added twice and
    # never queued for a second decline email.
    try:
        _rej_existing = client.list_rejected_rows()
    except SharePointError:
        _rej_existing = []
    rej_ids = {str(r["values"].get("Application ID", "")).strip()
               for r in _rej_existing} - {""}
    rej_emails = {_bare_email(r["values"].get("Email", "")) for r in _rej_existing} - {""}
    rej_phones = {_normalize_phone_key(r["values"].get("Phone", "")) for r in _rej_existing} - {""}

    logger.info("")
    logger.info("  Step 4/4  Scoring candidates...")
    logger.info("")
    roles = get_active_roles()
    from hiring_agent import extraction as _extraction_mod
    _extraction_mod.STRICT_AI_STAGES = bool(_cfg.REQUIRE_AI)
    processed, errors, rejected, location_review = 0, 0, 0, 0
    degraded_extractions = 0   # rows whose fields came from the regex parser, not the model
    deferred = 0               # rows left in the queue because the model could not read them
    deleted_idx: list[int] = []   # original indexes of rows deleted this run
    consecutive_failures = 0
    consecutive_failure_limit = getattr(_cfg, "CONSECUTIVE_FAILURE_LIMIT", 3)
    circuit_breaker_tripped = False

    for current, row in enumerate(rows, start=1):
        vals = row["values"]
        app_id = str(vals.get("Application ID", "")).strip()
        candidate_name = str(vals.get("Full Name", "") or "").strip()
        if _is_gap(candidate_name):
            candidate_name = "(not extracted yet)"
        original_index = row.get("index")
        row_at_start = original_index + 2 if isinstance(original_index, int) else "?"
        batch_left = batch_total - current
        queue_left = max(queue_total - current, 0)
        candidate_result = "Processing Error"

        def log_candidate_progress(activity: str, state: str = "candidate",
                                   result: str = "") -> None:
            progress_fields = {
                "current": current,
                "batch": batch_total,
                "batch_left": batch_left,
                "queue_left": queue_left,
                "row": row_at_start,
                "ref": app_id or "(no ref)",
                "name": candidate_name,
            }
            if state == "done":
                progress_fields["result"] = result
            else:
                progress_fields["activity"] = activity
            _log_p2_progress(state, **progress_fields)
        logger.info(f"  ── Candidate {current}/{batch_total} ──────────────────────────────")
        logger.info(f"       Ref       : {app_id or '(no ref)'}")
        logger.info(f"       Email     : {vals.get('Email', '?')}")
        # Deleting a row shifts every later position up by one. _live_row_index applies
        # that offset AND verifies it whenever one was needed, so a bookkeeping slip can
        # never silently write this candidate's data onto a different candidate's row.
        index = _live_row_index(client, row, deleted_idx, app_id)
        if index is None:
            errors += 1
            candidate_result = "Skipped - row could not be located"
            log_candidate_progress("Row missing", state="done", result=candidate_result)
            logger.info("")
            continue
        log_candidate_progress("Checking resume")
        stored, subpath = [], ""   # always bound, even if the try block raises before setting these
        try:
            if str(vals.get("Has Resume", "")).strip().lower() != "yes":
                logger.warning(f"       SKIP     : 'Has Resume' is '{vals.get('Has Resume')}' (not 'Yes') — flagging for review.")
                if _mark_needs_review(client, index, vals, dry_run, app_id=app_id):
                    deleted_idx.append(row["index"])
                candidate_result = (
                    "Would need unreadable-resume review"
                    if dry_run else
                    "Rejected - Processing Error"
                    if _score_attempts_of(vals) + 1 >= _cfg.SCORE_RETRY_MAX else
                    "Needs Review - Unreadable Resume"
                )
                continue

            all_candidate_names = vals.get("_Local Resume Names") or _stored_resume_names(
                app_id, vals.get("Original Filename", ""),
                full_name=vals.get("Full Name"), category=vals.get("Category"),
                resume_url=vals.get("Resume URL"))
            if not all_candidate_names:
                logger.warning("       SKIP     : No resume filename on this row — flagging for review.")
                if _mark_needs_review(client, index, vals, dry_run, app_id=app_id):
                    deleted_idx.append(row["index"])
                candidate_result = (
                    "Would need unreadable-resume review"
                    if dry_run else
                    "Rejected - Processing Error"
                    if _score_attempts_of(vals) + 1 >= _cfg.SCORE_RETRY_MAX else
                    "Needs Review - Unreadable Resume"
                )
                continue

            received = _parse_received(vals.get("Received Date"))
            subpath = _cfg.dated_subpath(received)   # e.g. 2026/June
            save_copies = SAVE_LOCAL_COPIES and not dry_run
            dest_dir = _cfg.dated_dir(_cfg.INPUT_DIR, received)
            if save_copies:
                dest_dir.mkdir(parents=True, exist_ok=True)
            _temp_files.clear()
            # Picks exactly one file per real attachment slot (a legacy-shape write always
            # wins over a stale canonical-shape one - see _resume_name_slots) instead of
            # downloading+concatenating every name guess, so a candidate who updates AFTER
            # already being scored once gets re-scored off their new resume, not a garbled
            # blend of old + new text.
            log_candidate_progress("Downloading resume")
            text, stored, _raw_by_name = _download_resume_text(
                client, app_id, vals.get("Original Filename", ""), subpath,
                full_name=vals.get("Full Name"), category=vals.get("Category"),
                resume_url=vals.get("Resume URL"),
                explicit_names=vals.get("_Local Resume Names"))
            if save_copies:
                for name, raw in _raw_by_name.items():
                    fpath = dest_dir / name
                    fpath.write_bytes(raw)
                    _temp_files.append(fpath)
            # Any other guess for a slot we actually used (typically a stale pre-update
            # canonical-shape file, now superseded by a fresh legacy-shape update) is a
            # leftover to remove once the new content is safely scored below - never left
            # behind as an orphan second file for the same candidate.
            # Case-insensitive: SharePoint filenames are, so a probe that differs from a
            # real stored file only in case is the SAME file, never a stale sibling.
            _stored_ci_pre = {str(s).strip().lower() for s in stored}
            stale_siblings = [n for n in all_candidate_names
                              if str(n).strip().lower() not in _stored_ci_pre]
            if not text:
                logger.warning("       SKIP     : No readable text (scanned/image PDF?) — "
                               "flagging for review (still auto-retried every run).")
                if _mark_needs_review(client, index, vals, dry_run, app_id=app_id,
                                      stored=stored, subpath=subpath):
                    deleted_idx.append(row["index"])
                candidate_result = (
                    "Would need unreadable-resume review"
                    if dry_run else
                    "Rejected - Processing Error"
                    if _score_attempts_of(vals) + 1 >= _cfg.SCORE_RETRY_MAX else
                    "Needs Review - Unreadable Resume"
                )
                continue

            # Secondary content-safety net - P1 already screens incoming mail for spam/
            # scam/vendor-pitch content before a row is ever created; this only catches
            # the rare row that still slips through. Checked before any extraction runs
            # so a suspicious row is never scored, renamed, or moved like a real candidate.
            _early_mail_body = html_to_text(str(vals.get("Mail Body", "") or ""))
            _suspicious_hits = _detect_suspicious_content(text, _early_mail_body)
            if _suspicious_hits:
                logger.warning(f"       SKIP     : suspicious content matched "
                               f"{_suspicious_hits} — flagging for review, alerting admin.")
                if not dry_run:
                    _store(client).save_by_id(
                        _row_key(app_id, vals), {"Status": _cfg.STATUS_NEEDS_REVIEW_SPAM},
                        current_values=vals, hint=index)
                    _send_suspicious_content_alert(
                        client, app_id, str(vals.get("Email", "")), _suspicious_hits)
                candidate_result = (
                    "Would flag as possible spam" if dry_run
                    else "Needs Review - Possible Spam"
                )
                processed += 1
                continue

            # Step A - snapshot all configured columns' starting state before any extraction runs,
            # so the Step E diff (below) can show exactly what P1 handed off vs what P2 changed.
            from hiring_agent.config import COLUMNS as _COLUMNS
            before_snapshot = {c: str(vals.get(c, "") or "").strip() for c in _COLUMNS}
            # One line naming which columns P1 handed off pre-filled (the rest arrived blank);
            # the per-column filled/healed picture is the Step E "Diff" line emitted below.
            _prefilled = [_c for _c in _COLUMNS if not _is_gap(before_snapshot[_c])]
            logger.info("       Before     : " + (
                f"{len(_prefilled)} column(s) from P1: {', '.join(_prefilled)}"
                if _prefilled else "row arrived blank"))

            log_candidate_progress("Extracting candidate details")
            details = extract_candidate_details_smart(text)
            # An Ollama timeout falls back to the regex parser and the row still goes out as
            # 'Scored'. Measured 2026-09-06: an ordinary two-page resume takes ~62s against a
            # 60s budget on this hardware, so that fallback is common, not exceptional. Count
            # and name it, or a degraded run is indistinguishable from a good one.
            _extract_src = EXTRACTION_SOURCE["value"]
            if _extract_src != "ollama":
                degraded_extractions += 1
                if _cfg.REQUIRE_AI:
                    # DEFER, do not write. Status stays 'New Email Received' so the next run
                    # retries this candidate untouched. Deliberately NOT a Retry Count bump:
                    # that counter gives up to 'Rejected - Processing Error', and a slow
                    # machine must never reject a real applicant.
                    deferred += 1
                    consecutive_failures += 1
                    logger.warning(
                        f"       DEFERRED : {app_id} - {_extract_src}. Not written; the row "
                        f"keeps its place in the queue and the next run will retry it."
                    )
                    if consecutive_failures >= consecutive_failure_limit:
                        circuit_breaker_tripped = True
                        logger.error(
                            f"  CIRCUIT BREAKER : {consecutive_failures} consecutive candidate failures/timeouts encountered. "
                            f"Halting bot to protect the queue from repeated failures. Remaining candidate(s) in this batch "
                            f"remain untouched in the queue and will be scored next time."
                        )
                        break
                    continue
                logger.warning(f"       DEFERRED : {app_id} extracted by {_extract_src} - "
                               f"regex fallback is forbidden when LLM is dead. Candidate deferred.")
                deferred += 1
                consecutive_failures += 1
                continue

            # Snapshot the Tier 1/2 fields right after extraction, before the mail-body
            # merge or AI recheck can still touch them - Step E uses this to tell "filled
            # during extraction" apart from "filled later" for the per-field method tag.
            after_extract_snapshot = {col: str(details.get(key, "") or "").strip()
                                      for col, key in _P2_COL_TO_KEY.items()}

            # Who the candidate IS comes from the CV, not from the covering email.
            # Captured before the mail-body merge and the AI recheck, both of which read
            # the email and can overwrite these - see _restore_cv_identity below.
            _cv_identity = {k: details.get(k, "") for k in _CV_IDENTITY_FIELDS}

            mail_body_raw = str(vals.get("Mail Body", "") or "")
            mail_body = html_to_text(mail_body_raw) if mail_body_raw else ""
            if mail_body:
                details = merge_mail_body_fallback(details, mail_body)

            from concurrent.futures import ThreadPoolExecutor
            with ThreadPoolExecutor(max_workers=2) as _executor:
                _recheck_fut = _executor.submit(ai_recheck_fields, details, text, mail_body)
                _p1 = details.get("portfolio_1", "N/A")
                _p2 = details.get("portfolio_2", "N/A")
                _p3 = details.get("portfolio_3", "N/A")
                _port_fut = _executor.submit(infer_missing_portfolios, text, _p1, _p2, _p3)

                details = _recheck_fut.result()
                details = _restore_cv_identity(details, _cv_identity, app_id)

                if str(details.get("looking_for_role", "")).strip().lower() in _GAP_LITERALS:
                    details["looking_for_role"] = infer_looking_for_role(
                        text, mail_body, details.get("skills", ""))

                portfolio_1, portfolio_2, portfolio_3 = _port_fut.result()

            full_name = resolve_full_name(details.get("full_name"),
                                          str(vals.get("Full Name", "")), text)
            skills = normalize_skills(details.get("skills", "Not extracted"))
            role_pref = details.get("looking_for_role", "Not extracted")
            education = details.get("education", "Not extracted")
            education_start_date = details.get("education_start_date", "")
            education_end_date = details.get("education_end_date", "")
            experience = details.get("experience", "")
            location = details.get("location", "Not extracted")
            country = normalize_country(details.get("country", ""))   # USA/America → United States
            phone = details.get("phone", "Not extracted")   # US-formatted below, once geo is known

            # Any slot still empty reads 'N/A', never a blank cell.
            portfolio_1, portfolio_2, portfolio_3 = (
                _na_if_gap(portfolio_1), _na_if_gap(portfolio_2), _na_if_gap(portfolio_3))

            log_candidate_progress("Matching roles and checking location")
            res = suggested_roles(skills, role_pref, roles=roles, resume_text=text)
            if res.get("source") != "ollama":
                # Same rule as extraction: a keyword-scored row is not a model-scored row, so
                # do not publish a match percentage that reads as an AI verdict.
                deferred += 1
                consecutive_failures += 1
                logger.warning(
                    f"       DEFERRED : {app_id} - role scoring fell back to the "
                    f"{res.get('source', 'keyword')} scorer. Not written; retried next run."
                )
                if consecutive_failures >= consecutive_failure_limit:
                    circuit_breaker_tripped = True
                    logger.error(
                        f"  CIRCUIT BREAKER : {consecutive_failures} consecutive candidate failures/timeouts encountered. "
                        f"Halting bot to protect the queue from repeated failures. Remaining candidate(s) in this batch "
                        f"remain untouched in the queue and will be scored next time."
                    )
                    break
                continue
            r1, r2, r3 = res["role_1"], res["role_2"], res.get("role_3", "")
            category = assign_category(r1, skills)   # never blank — falls back to "General"

            if GEO_FILTER_USA_ONLY:
                geo_decision, usa_reason = _classify_candidate_geo(
                    location, country, phone, education=education, resume_text=text)
            else:
                geo_decision, usa_reason = GeoDecision.CONFIRMED_US, "geo filter off"

            # Once residence is confirmed, finalize the Country label: a concrete US location
            # signal makes it the canonical 'United States', overriding a US STATE name or a
            # foreign country mis-read off a past overseas degree (see reconcile_us_country).
            if geo_decision != GeoDecision.UNKNOWN:
                country = reconcile_us_country(
                    country, location, geo_decision == GeoDecision.CONFIRMED_US)
            # NOTE (2026-08-04): an earlier revision inferred Country='United States' here
            # from a US phone + US university when the resume carried no address. That was
            # REMOVED on client instruction: a university (e.g. ASU) and a retained phone
            # number are not evidence of where someone lives now, so labelling the cell
            # from them is an assumption, not a fact. A resume with no address leaves
            # Location/Country genuinely unknown - marked MISSING_VALUE and asked for by
            # email - never guessed.

            # Any NANP-shaped number (10 digits, or 11 starting with country code 1) gets
            # the uniform (XXX) XXX-XXXX format regardless of geo decision (fixed
            # 2026-09-01: gating this on CONFIRMED_US left a location-review/unknown row's
            # plainly-US-shaped phone unformatted until the next maintenance sweep). A
            # genuine international +NN number is a different digit count and format_phone
            # leaves it exactly as written - US parens would misrepresent its area code.
            phone = format_phone(phone)
            phone = sanitize_phone(phone)   # never publish a masked/truncated number

            fields = {
                "Full Name": full_name,
                "Phone": phone,
                "Location": location,
                "Country": country,
                "Years Exp": experience,
                "Current Skills": skills,
                "Looking For Role": role_pref,
                "Education": education,
                "Education Start Date": education_start_date,
                "Education End Date": education_end_date,
                "Suggested Role 1": r1,
                "Suggested Role 2": r2,
                "Suggested Role 3": r3,
                "Category": category,
                # Same compaction/junk filter as the derive path, so the two never disagree.
                **dict(zip(("Portfolio 1", "Portfolio 2", "Portfolio 3"),
                           _compact_portfolios(portfolio_1, portfolio_2, portfolio_3))),
                "Status": (
                    STATUS_SCORED
                    if geo_decision == GeoDecision.CONFIRMED_US
                    else STATUS_LOCATION_REVIEW
                    if geo_decision == GeoDecision.UNKNOWN
                    else _cfg.STATUS_REJECTED
                ),
            }
            # Resume URL/Path are computed once the reject-vs-score branch below is known
            # (rejected candidates' files move to a different folder, so the right link needs
            # the right base folder — see the two branches further down).

            # Single column-by-column validation pass across all configured columns.
            # P1 columns are read from the stored row; P2 columns from the fields dict.
            # Any P2 field still N/A gets one mail-body fill attempt, then accepted.
            # This runs exactly once per candidate, right before the write. Step E's
            # diff+method log uses the Step A snapshot captured at the top of this loop.
            fields = _validate_all_columns(fields, vals, mail_body,
                                           before_snapshot, after_extract_snapshot)
            # LAST step before the write: anything the candidate owes us that is still empty
            # reads 'Missing' rather than sitting blank. Runs after _validate_all_columns so
            # the mail-body fill attempt gets its chance first - we only label a cell Missing
            # once every route to a real value has been tried.
            fields = _apply_missing_placeholders(fields)

            logger.info(f"       Name      : {fields['Full Name']}")
            logger.info(f"       Phone     : {fields['Phone']}")
            logger.info(f"       Role want : {fields['Looking For Role']}")
            logger.info(f"       Role #1   : {fields['Suggested Role 1']}")
            logger.info(f"       Role #2   : {fields['Suggested Role 2']}")
            logger.info(f"       Role #3   : {fields['Suggested Role 3']}")
            logger.info(f"       Category  : {fields['Category']}")
            logger.info(f"       Portfolio1: {fields['Portfolio 1']}")
            logger.info(f"       Portfolio2: {fields['Portfolio 2']}")
            logger.info(f"       Portfolio3: {fields['Portfolio 3']}")
            logger.info(f"       Scored by : {res['source']}")
            logger.info(f"       Skills    : {fields['Current Skills'][:80]}")
            logger.info(f"       Location  : {fields['Location']}")
            logger.info(f"       Country   : {fields['Country']}")
            logger.info(f"       Geo check : {geo_decision.value} ({usa_reason})")

            # Rename the resume file(s) in SharePoint if we have a parsed name
            full_name = fields.get("Full Name") or vals.get("Full Name")
            candidate_name = str(full_name or candidate_name).strip()
            log_candidate_progress("Saving result")
            stored = _rename_scored_resumes(client, stored, subpath, app_id, full_name,
                                            fields.get("Category"), dry_run)

            # Clean up any stale sibling file(s) left over from _download_resume_text's
            # slot resolution above (typically a pre-update canonical-shape file, now
            # superseded by the fresh legacy-shape update we just scored and renamed) -
            # so a candidate who updates after already being scored ends up with exactly
            # one resume file, not an orphaned second one under the old name.
            if not dry_run:
                # CASE-INSENSITIVE, and that is the whole point. SharePoint/OneDrive treat
                # filenames case-insensitively, but this guard used a case-SENSITIVE Python
                # `in`, so a stale probe differing from the freshly-renamed file only by case
                # slipped past it and the delete below then removed the real file.
                #
                # It fires on the COMMON path, not an edge case: P1 stores the name off the
                # email header ('yash sik verma') while P2 extracts the corrected one from the
                # resume ('Yash Verma'), so the rename target is 'Yash_Verma_<tail>.pdf' while
                # the probe list still carries 'yash_verma_<tail>.pdf'. Live case 2026-08-24,
                # APP-20260824-2049-642B: the resume was renamed, then immediately deleted as
                # 'stale', leaving the candidate with no file and a blank Resume URL /
                # Resume Folder Path (and therefore a blank calculated Resume Link).
                # CONTAINMENT CHANGE 2026-09-06 (docs/P1_P2_IMPLEMENTATION_STEPS.md
                # Section 1): this used to DELETE each stale guess outright, on every
                # ordinary scoring run, with no test coverage. Its own comment already
                # documents one live incident despite the case-insensitivity guard added
                # after it (2026-08-24, APP-20260824-2049-642B): a stale probe slipped past a
                # case-SENSITIVE check and the delete below removed the candidate's real,
                # freshly-renamed file, leaving a blank Resume URL/Path/Link. The guard
                # narrowed the failure window; it did not remove the risk that this file is
                # not actually stale. Preserve every candidate; log what would have been
                # removed so a future reviewed retention procedure has something to act on.
                _stored_ci = {str(s).strip().lower() for s in stored}
                for stale in stale_siblings:
                    if str(stale).strip().lower() in _stored_ci:
                        continue
                    logger.info(
                        f"       Cleanup   : possible stale resume file '{stale}' preserved "
                        f"(not deleted) - pending a reviewed retention procedure."
                    )

            if geo_decision == GeoDecision.CONFIRMED_NON_US:
                fields["Status"] = _cfg.STATUS_REJECTED
                if _finish_rejection(client, index, app_id, vals, fields, stored, subpath,
                                     dry_run, rej_ids, rej_emails, rej_phones,
                                     "non-USA", rejected_rows=_rej_existing):
                    deleted_idx.append(row["index"])
                    rejected += 1
                consecutive_failures = 0
                processed += 1
                candidate_result = "Would reject - Non-USA" if dry_run else "Rejected - Non-USA"
                logger.info("")
                continue

            if geo_decision == GeoDecision.UNKNOWN:
                fields["Status"] = STATUS_LOCATION_REVIEW
                if dry_run:
                    logger.info("       Result   : [DRY-RUN] Would keep on Main for "
                                "location confirmation (no decline).")
                else:
                    resume_url, resume_path = _resume_url_and_path(client, stored, subpath)
                    if resume_url:
                        fields["Resume URL"] = resume_url
                        fields["Resume Folder Path"] = resume_path
                    if _score_attempts_of(vals) != 0:
                        fields["Retry Count"] = 0
                    _store(client).save_by_id(_row_key(app_id, vals), fields, current_values=vals, hint=index)
                    if hasattr(client, "set_row_fill") and hasattr(client, "table_name"):
                        client.set_row_fill(client.table_name, index, color_hex="#FFF2CC")
                    logger.info("       Result   : LOCATION REVIEW - kept on Main; "
                                "no decline email queued.")
                consecutive_failures = 0
                location_review += 1
                processed += 1
                candidate_result = (
                    "Would need location review" if dry_run
                    else "Needs Review - Location Confirmation"
                )
                logger.info("")
                continue
            # Pre-Scoring Completeness Gate: ensure core fields are strictly present and valid
            _r1 = str(fields.get("Suggested Role 1", "") or "").strip()
            _cat = str(fields.get("Category", "") or "").strip()

            if not _r1 or _is_gap(_r1):
                logger.warning("       PRE-SCORING CHECK: Suggested Role 1 is empty. Setting status to Needs Review.")
                fields["Status"] = _cfg.STATUS_NEEDS_REVIEW
            elif not _cat or _is_gap(_cat):
                fields["Category"] = "General"

            if fields["Status"] == _cfg.STATUS_NEEDS_REVIEW:
                if dry_run:
                    logger.info("       Result   : [DRY-RUN] Would mark Needs Review (missing core fields).")
                else:
                    _store(client).save_by_id(_row_key(app_id, vals), fields, current_values=vals, hint=index)
                    if hasattr(client, "set_row_fill") and hasattr(client, "table_name"):
                        client.set_row_fill(client.table_name, index, color_hex="#FFF2CC")
                    logger.info("       Result   : NEEDS REVIEW — kept on Main for manual review.")
                consecutive_failures = 0
                processed += 1
                candidate_result = "Needs Review - Extraction Incomplete"
                logger.info("")
                continue

            if dry_run:
                logger.info("       Result   : [DRY-RUN] Would score (no write).")
            else:
                resume_url, resume_path = _resume_url_and_path(client, stored, subpath)
                if resume_url:
                    fields["Resume URL"] = resume_url
                    fields["Resume Folder Path"] = resume_path
                # Clean slate on a genuine success - a row that failed once then scored fine
                # on retry shouldn't carry a stale failure count forward.
                if _score_attempts_of(vals) != 0:
                    fields["Retry Count"] = 0
                _store(client).save_by_id(_row_key(app_id, vals), fields, current_values=vals, hint=index)
                has_doubt, doubt_reason = is_doubt_candidate(fields)
                if hasattr(client, "set_row_fill") and hasattr(client, "table_name"):
                    client.set_row_fill(client.table_name, index, color_hex="#FFF2CC" if has_doubt else None)
                if has_doubt:
                    logger.info(f"       HIGHLIGHT : Row highlighted for review ({doubt_reason})")
                if scorecards:
                    card = (
                        f"Application: {app_id}\nName: {fields['Full Name']}\n"
                        f"Email: {vals.get('Email','')}\nPhone: {fields['Phone']}\n"
                        f"Location: {fields['Location']}\nCountry: {fields['Country']}\n"
                        f"Skills: {fields['Current Skills']}\n"
                        f"Looking For: {fields['Looking For Role']}\n"
                        f"Role 1: {fields['Suggested Role 1']}\n"
                        f"Role 2: {fields['Suggested Role 2']}\n"
                        f"Role 3: {fields['Suggested Role 3']}\n"
                        f"Category: {fields['Category']}\n"
                    )
                    try:
                        client.upload_file(f"{client.resumes_folder}/{subpath}",
                                           f"{app_id}_scorecard.txt", card.encode("utf-8"))
                    except SharePointError as e:
                        logger.warning(f"       WARNING  : Scorecard upload failed: {e}")
                logger.info("       Result   : SCORED — row updated in SharePoint.")
            consecutive_failures = 0
            processed += 1
            candidate_result = "Would score" if dry_run else "Scored"
            logger.info("")
        except ResumeUnavailableError as e:
            deferred += 1
            consecutive_failures += 1
            candidate_result = "Deferred - resume unavailable"
            logger.warning(f"       DEFERRED : {e}. Status and Retry Count unchanged; "
                           "repair the resume link/file before retrying.")
            if consecutive_failures >= consecutive_failure_limit:
                circuit_breaker_tripped = True
                logger.error(
                    f"  CIRCUIT BREAKER : {consecutive_failures} consecutive candidate failures/timeouts encountered. "
                    f"Halting bot to protect the queue from repeated failures. Remaining candidate(s) in this batch "
                    f"remain untouched in the queue and will be scored next time."
                )
                break
        except TimeoutError as e:
            deferred += 1
            consecutive_failures += 1
            candidate_result = "Deferred - processing timed out"
            logger.warning(f"       DEFERRED : {app_id} - OCR or model timed out ({e}). "
                           "Status and Retry Count unchanged; candidate retried next run.")
            if consecutive_failures >= consecutive_failure_limit:
                circuit_breaker_tripped = True
                logger.error(
                    f"  CIRCUIT BREAKER : {consecutive_failures} consecutive candidate failures/timeouts encountered. "
                    f"Halting bot to protect the queue from repeated failures. Remaining candidate(s) in this batch "
                    f"remain untouched in the queue and will be scored next time."
                )
                break
        except Exception as e:
            errors += 1
            consecutive_failures += 1
            new_attempts = _score_attempts_of(vals) + 1
            if new_attempts >= _cfg.SCORE_RETRY_MAX:
                candidate_result = (
                    "Would reject - Processing Error"
                    if dry_run else "Rejected - Processing Error"
                )
                logger.error(f"       ERROR    : {app_id} failed: {e} "
                            f"(attempt {new_attempts}/{_cfg.SCORE_RETRY_MAX} - giving up).")
                try:
                    _give_up_and_reject(client, index, vals, app_id, stored, subpath,
                                        f"exception: {e}", dry_run)
                except Exception as give_up_err:
                    logger.warning(f"       WARNING  : could not give up on {app_id}: {give_up_err}")
            else:
                candidate_result = (
                    "Would remain pending after processing error"
                    if dry_run else "Processing Error - Retry Pending"
                )
                logger.error(f"       ERROR    : {app_id} failed: {e} "
                            f"(attempt {new_attempts}/{_cfg.SCORE_RETRY_MAX} - left unscored for retry).")
                if not dry_run:
                    try:
                        _store(client).save_by_id(_row_key(app_id, vals), {"Retry Count": new_attempts}, current_values=vals, hint=index)
                    except Exception as bump_err:
                        logger.warning(f"       WARNING  : could not bump Retry Count "
                                       f"for {app_id}: {bump_err}")
            logger.info("")
            if consecutive_failures >= consecutive_failure_limit:
                circuit_breaker_tripped = True
                logger.error(
                    f"  CIRCUIT BREAKER : {consecutive_failures} consecutive candidate failures/timeouts encountered. "
                    f"Halting bot to protect the queue from repeated failures. Remaining candidate(s) in this batch "
                    f"remain untouched in the queue and will be scored next time."
                )
                break
        finally:
            log_candidate_progress("", state="done", result=candidate_result)
            _cleanup_temp_files()
            # Breathe between candidates so a long unattended client run stays clear of
            # Graph throttling. In `finally` on purpose: a row that errored still made its
            # Graph calls, so it must pace the same as a successful one. Skipped after the
            # LAST row (nothing follows it to protect) and in dry-run (no writes at all).
            if _cfg.ROW_DELAY_SECONDS and not dry_run and current < batch_total:
                _time.sleep(_cfg.ROW_DELAY_SECONDS)

    if _local_client is not None:
        return {"processed": processed, "errors": errors, "deferred": deferred,
                "rejected": rejected, "location_review": location_review}
    _cleanup_empty_dirs()

    # A fresh P1 update is intentionally excluded from the schema-time merge while it is
    # still "New Email Received" so it can be scored from the newest resume first. Once
    # this pass has scored/moved those rows, merge again so same-email/phone duplicates
    # across Main + Rejected do not sit around until the next scheduled run.
    if not dry_run:
        try:
            post_merge = _merge_duplicate_candidates(client)
            if (post_merge.get("healed") or post_merge.get("duplicates_found")
                    or post_merge.get("errors")):
                logger.info(f"  Duplicate : post-score check found "
                            f"{post_merge.get('duplicates_found', 0)} duplicate row(s), "
                            f"healed {post_merge.get('healed', 0)} winner(s) | "
                            f"{post_merge.get('errors', 0)} errors")
        except Exception as e:
            logger.warning(f"  WARNING   Post-score duplicate check failed: {e}")

    try:
        queue_remaining = len(_scoring_queue_rows(client))
    except Exception as e:
        queue_remaining = max(queue_total - processed, 0)
        logger.warning(
            f"  WARNING   Could not refresh the final queue count ({e}); "
            f"showing estimate {queue_remaining}."
        )
    _log_p2_progress(
        "finish",
        attempted=batch_total,
        remaining=queue_remaining,
        errors=errors,
    )

    # Close out review rows nobody answered, BEFORE the mail passes below, so a row retired
    # here is never nudged or declined in the same run (it is pre-stamped as settled anyway).
    aged = _age_out_stale_reviews(client, dry_run=dry_run)

    # Decline emails go out AFTER all sheet moves, one candidate at a time, driven by
    # the blank/stamped 'Decline Sent' column — idempotent, crash-safe, backfill-safe.
    declines = _send_pending_declines(client, dry_run=dry_run)
    # Same idempotent pass, for SCORED rows missing phone/location/portfolio.
    info_reqs = _send_pending_info_requests(client, dry_run=dry_run)

    logger.info("  ── Summary ─────────────────────────────────────────")
    logger.info(f"       Processed : {processed}")
    logger.info(f"       Geo review: {location_review} (kept on Main; no decline)")
    logger.info(f"       Rejected  : {rejected} (non-USA)")
    logger.info(f"       Declines  : {declines.get('sent', 0)} sent | "
                f"{declines.get('failed', 0)} failed (auto-retry next run) | "
                f"{declines.get('skipped', 0)} skipped")
    logger.info(f"       Info reqs : {info_reqs.get('sent', 0)} sent | "
                f"{info_reqs.get('failed', 0)} failed (auto-retry next run) | "
                f"{info_reqs.get('skipped', 0)} skipped")
    logger.info(f"       Aged out  : {aged.get('closed', 0)} closed silently after "
                f"{_cfg.REVIEW_AGE_OUT_DAYS} days | {aged.get('skipped', 0)} still waiting")
    export_path = ""
    if not dry_run:
        try:
            export_path = export_client_results(client, upload_to_sharepoint=True)
            if export_path:
                logger.info(f"  Export    : client workbook updated -> {export_path}")
        except Exception as e:
            logger.warning(f"  WARNING   Client workbook export failed: {e}")
        # Extra layer, run AGAIN now that the results sheet exists (client instruction,
        # 2026-09-01): a field recovered here reaches the JUST-published export
        # immediately, on a re-export, rather than waiting for tomorrow's run to notice.
        try:
            if _reevaluate_missing_scored_fields(client) and export_path:
                export_path = export_client_results(client, upload_to_sharepoint=True)
                logger.info(f"  Export    : re-published after post-export field "
                            f"re-evaluation -> {export_path}")
        except Exception as e:
            logger.warning(f"  WARNING   Post-export field re-evaluation failed: {e}")

    logger.info(f"       Errors    : {errors}")
    if deferred:
        logger.warning(
            f"       DEFERRED  : {deferred} candidate(s) were left in the queue because the "
            f"resume or model was unavailable. Nothing was written for them and no retry "
            f"counter moved; they will be picked up again next run."
        )
    if circuit_breaker_tripped:
        logger.warning(
            f"       CIRCUIT BREAKER : Run stopped early after {consecutive_failures} consecutive failures/timeouts. "
            f"The remaining {queue_remaining} candidate(s) were preserved untouched in the queue and will be scored next time."
        )
    if degraded_extractions:
        logger.warning(
            f"       DEGRADED  : {degraded_extractions} of {len(rows)} candidate(s) were "
            f"read by the OFFLINE parser, not the model - their fields are regex-quality "
            f"even though the rows are marked Scored. Re-score them once Ollama is healthy."
        )
    return {"processed": processed, "rejected": rejected,
            "location_review": location_review, "errors": errors,
            "degraded_extractions": degraded_extractions, "deferred": deferred,
            "circuit_breaker_tripped": circuit_breaker_tripped,
            "declines": declines, "info_requests": info_reqs,
            "aged_out": aged, "client_export": export_path}


# ── Full reconcile sweep (bot.py --recheck-all) ──────────────────────────────

_STATUS_NEW = "New Email Received"
# P2-owned content columns; all of these need the resume to (re)derive.
_P2_CONTENT_COLS = [
    "Full Name", "Phone", "Location", "Country", "Current Skills", "Looking For Role",
    "Years Exp",
    "Education", "Education Start Date", "Education End Date",
    "Suggested Role 1", "Suggested Role 2", "Suggested Role 3",
    "Portfolio 1", "Portfolio 2", "Portfolio 3", "Resume URL", "Resume Folder Path",
]


def _is_gap(v) -> bool:
    """A cell counts as a gap when it's blank/missing or a placeholder (N/A, Not extracted).

    Uses the single canonical placeholder set shared with extraction.py (_GAP_LITERALS) -
    this module used to keep its own separate, slightly narrower copy (_GAP_VALUES).
    """
    return str(v or "").strip().lower() in _GAP_LITERALS


def _foreign_country_only_reject_is_too_weak(fields: dict, usa_reason: str) -> bool:
    """Keep weak country-only rejects for clarification instead of sending a decline.

    Normal non-USA rows still reject: concrete foreign locations reject before this, and
    country-only rows with enough candidate substance still reject. This catches rows
    where the only foreign clue is a fragile Country inference while name/phone/skills/
    education are mostly unreadable.
    """
    if not str(usa_reason or "").lower().startswith("country field"):
        return False
    if not _is_gap(fields.get("Location")):
        return False
    substance_cols = ("Full Name", "Phone", "Current Skills", "Education")
    substantive = sum(0 if _is_gap(fields.get(col)) else 1 for col in substance_cols)
    return substantive < 2


def _derive_row_fields(client, vals: dict, roles, rejected: bool = False) -> dict | None:
    """Re-run the extraction+scoring pipeline for one row from its stored resume.
    Returns the P2-owned content fields, or None if the resume can't be read.

    `rejected` selects which dated bucket the file actually lives under: the month's main
    folder, or its 'Rejected' sub-bucket — pass rejected=True for a row on the Rejected sheet."""
    if client is None:
        return None
    app_id = str(vals.get("Application ID", "")).strip()
    if not _stored_resume_names(
            app_id, vals.get("Original Filename", ""),
            full_name=vals.get("Full Name"), category=vals.get("Category"),
            resume_url=vals.get("Resume URL")):
        return None
    received = _parse_received(vals.get("Received Date"))
    subpath = _cfg.dated_subpath(received)
    fetch_subpath = _rejected_subpath(subpath) if rejected else subpath
    _temp_files.clear()
    # One winner per attachment slot (legacy-shape update beats stale canonical-shape prior
    # content) - same reconciliation as the main scoring loop, see _download_resume_text.
    text, used, _raw = _download_resume_text(
        client, app_id, vals.get("Original Filename", ""), fetch_subpath,
        full_name=vals.get("Full Name"), category=vals.get("Category"),
        resume_url=vals.get("Resume URL"))
    if not text:
        return None

    details = extract_candidate_details_smart(text)
    # Same CV-is-the-authority rule as the main scoring loop - see _restore_cv_identity.
    _cv_identity = {k: details.get(k, "") for k in _CV_IDENTITY_FIELDS}
    mail_body_raw = str(vals.get("Mail Body", "") or "")
    mail_body = html_to_text(mail_body_raw) if mail_body_raw else ""
    if mail_body:
        details = merge_mail_body_fallback(details, mail_body)
    details = ai_recheck_fields(details, text, mail_body)
    details = _restore_cv_identity(details, _cv_identity, app_id)
    if str(details.get("looking_for_role", "")).strip().lower() in _GAP_LITERALS:
        details["looking_for_role"] = infer_looking_for_role(
            text, mail_body, details.get("skills", ""))

    full_name = resolve_full_name(details.get("full_name"), str(vals.get("Full Name", "")), text)
    skills = normalize_skills(details.get("skills", "Not extracted"))
    role_pref = details.get("looking_for_role", "Not extracted")
    res = suggested_roles(skills, role_pref, roles=roles, resume_text=text)
    p1, p2, p3 = infer_missing_portfolios(
        text, details.get("portfolio_1", "N/A"), details.get("portfolio_2", "N/A"),
        details.get("portfolio_3", "N/A"))
    _loc = details.get("location", "Not extracted")
    _ctry = normalize_country(details.get("country", ""))
    # Same Country reconciliation as the scoring path (reconcile_us_country): a concrete US
    # location signal makes the Country cell 'United States'. Gated on main-sheet rows only
    # (is_usa=not rejected) so a Rejected (non-US) row's country is never overwritten.
    _ctry = reconcile_us_country(_ctry, _loc, is_usa=not rejected)
    _phone = details.get("phone", "Not extracted")
    # NANP-shaped (10 digit, or 11 starting with 1) gets US formatting regardless of
    # Country - see _renormalize_phone_columns for why this isn't gated on 'United States'.
    _phone = format_phone(_phone)
    _phone = sanitize_phone(_phone)       # never publish a masked/truncated number
    _cp1, _cp2, _cp3 = _compact_portfolios(p1, p2, p3)
    derived = {
        "Full Name": full_name,
        "Phone": _phone,
        "Location": _loc,
        "Country": _ctry,
        "Years Exp": details.get("experience", ""),
        "Current Skills": skills,
        "Looking For Role": role_pref,
        "Education": details.get("education", "Not extracted"),
        "Education Start Date": details.get("education_start_date", ""),
        "Education End Date": details.get("education_end_date", ""),
        "Suggested Role 1": res["role_1"],
        "Suggested Role 2": res["role_2"],
        "Suggested Role 3": res.get("role_3", ""),
        # Compacted + junk-filtered, so slot 1 always holds a real link when one exists.
        "Portfolio 1": _cp1,
        "Portfolio 2": _cp2,
        "Portfolio 3": _cp3,
    }
    url, path = _resume_url_and_path(client, used, fetch_subpath)
    if url:
        derived["Resume URL"] = url
        derived["Resume Folder Path"] = path
    return derived


def _derive_geo_recovery_fields(client, vals: dict, rejected: bool = True) -> dict:
    """Recover only evidence needed for a historical geo decision.

    This path intentionally avoids Ollama, role matching, portfolio inference, and
    category changes. A university or employer can support a review decision, but it
    must never be copied into Location as though it were the candidate's current
    address. Only missing/broken Phone and missing Education cells are recovered.
    """
    if client is None:
        return {}
    phone_cur = str(vals.get("Phone", "") or "")
    phone_bad = (not _is_gap(phone_cur)) and (
        "*" in phone_cur or sum(ch.isdigit() for ch in phone_cur) < 7)
    needs_phone = _is_gap(phone_cur) or phone_bad
    needs_education = (
        _is_gap(vals.get("Education"))
        or education_needs_repair(vals.get("Education", ""))
    )
    if not needs_phone and not needs_education:
        return {}

    app_id = str(vals.get("Application ID", "") or "").strip()
    if not _stored_resume_names(
            app_id, vals.get("Original Filename", ""),
            full_name=vals.get("Full Name"), category=vals.get("Category"),
            resume_url=vals.get("Resume URL")):
        return {}

    subpath = _cfg.dated_subpath(_parse_received(vals.get("Received Date")))
    fetch_subpath = _rejected_subpath(subpath) if rejected else subpath
    text, _used, _raw = _download_resume_text(
        client, app_id, vals.get("Original Filename", ""), fetch_subpath,
        full_name=vals.get("Full Name"), category=vals.get("Category"),
        resume_url=vals.get("Resume URL"))
    if not text:
        return {}

    details = extract_candidate_details(text)
    mail_body_raw = str(vals.get("Mail Body", "") or "")
    if mail_body_raw:
        details = merge_mail_body_fallback(details, html_to_text(mail_body_raw))

    heal: dict = {}
    if needs_phone:
        phone = sanitize_phone(details.get("phone", "Not extracted"))
        if not _is_gap(phone):
            heal["Phone"] = phone
        elif phone_bad:
            heal["Phone"] = "Not extracted"
    if needs_education:
        education = details.get("education", "Not extracted")
        if not _is_gap(education):
            heal["Education"] = education
    return heal


def _plan_geo_recovery(client, vals: dict,
                       rejected: bool = True) -> tuple[dict, GeoDecision, str]:
    """Plan a fast, conservative historical geo recovery.

    Stored current-location fields remain authoritative. Resume education and phone
    evidence may prevent rejection, but cannot certify current US residence. The
    returned patch is limited to Phone, Education, Country, and Status.
    """
    heal = _derive_geo_recovery_fields(client, vals, rejected=rejected)
    location = str(vals.get("Location", "") or "")
    country = str(vals.get("Country", "") or "")
    phone = heal.get("Phone", vals.get("Phone", ""))
    education = heal.get("Education", vals.get("Education", ""))

    # Deliberately omit resume_text here. Passing it can invoke a slow Ollama context
    # check; deterministic phone/education evidence already gives us everything needed
    # for the conservative US/non-US/unknown decision.
    decision, reason = _classify_candidate_geo(
        location, country, phone, education=education, resume_text="")

    if decision == GeoDecision.CONFIRMED_US:
        reconciled = reconcile_us_country(country, location, True)
        if (not _is_gap(reconciled)
                and reconciled != str(vals.get("Country", "") or "").strip()):
            heal["Country"] = reconciled
        heal["Status"] = (
            STATUS_SCORED
            if not _is_gap(vals.get("Suggested Role 1"))
            else _cfg.STATUS_NEEDS_REVIEW
        )
    elif decision == GeoDecision.UNKNOWN:
        heal["Status"] = STATUS_LOCATION_REVIEW

    return heal, decision, reason


#: Columns re-derived by force_roles. Category is NOT listed: it is a pure function of
#: Suggested Role 1 + Current Skills and is already re-derived further down whenever the
#: role changes, so listing it here would only duplicate that.
_ROLE_COLS = ("Suggested Role 1", "Suggested Role 2", "Suggested Role 3")


def _plan_heal(client, vals: dict, roles,
               rejected: bool = False, force_roles: bool = False) -> tuple[dict, GeoDecision]:
    """Decide how to heal ONE row without churning good data.

    Returns (heal, geo_decision). `heal` holds only the columns that need filling/fixing:
    resume-dependent gaps are re-derived (downloading the resume only when needed);
    Category is re-derived from the role (deterministic) when it's blank or the known
    'General' fallback; already-good values are left untouched. `is_usa` is the fresh
    geo verdict used to reconcile which sheet the row belongs on. `rejected` selects which
    dated bucket the row's file(s) actually live under (main vs Rejected — see
    _derive_row_fields).

    `force_roles` (added 2026-08-04) additionally re-scores Suggested Role 1/2/3 even when
    those cells already hold a value. Normal healing is deliberately gap-only so a good row
    is never churned, but that also means a row scored under an older JD set or an older
    skill vocabulary keeps its stale match forever - there was no route to refresh it, since
    --rematch only processes local workbooks under P2_Output/, not SharePoint. Real case:
    the 2026-08-04 vocabulary gained finance/accounting terms, but APP-20260716-2052-6112
    (MS Finance) kept its pre-vocabulary 'Automation Process Developer (40%)' match because
    Suggested Role 1 was non-blank. OFF by default; opt in with `bot.py --force-rescore`.
    """
    gaps = [c for c in _P2_CONTENT_COLS if _is_gap(vals.get(c))]
    if force_roles:
        for _c in _ROLE_COLS:
            if _c not in gaps:
                gaps.append(_c)
    if education_needs_repair(vals.get("Education", "")) and "Education" not in gaps:
        gaps.append("Education")
    if _name_needs_rederive(vals.get("Full Name", ""), vals.get("Email", "")) and "Full Name" not in gaps:
        gaps.append("Full Name")

    # A stored Phone that isn't blank but can't be a real number — a mangled AI fragment
    # ("-8455"), a mask ("732-4***"), or any value with fewer than 7 digits — is wrong,
    # not missing. Re-derive it from the resume alongside the true gaps; if nothing better
    # comes back, it gets blanked below (a broken-looking number is worse than a clean gap).
    phone_cur = str(vals.get("Phone", "") or "")
    phone_bad = (not _is_gap(phone_cur)) and (
        "*" in phone_cur or sum(ch.isdigit() for ch in phone_cur) < 7)
    if phone_bad and "Phone" not in gaps:
        gaps.append("Phone")

    location = str(vals.get("Location", "") or "")
    country = str(vals.get("Country", "") or "")
    geo_on = GEO_FILTER_USA_ONLY
    need_resume = bool(gaps) or (geo_on and _is_gap(location))

    # Portfolio cells that hold link junk (AI display text, letter-spaced or doubled-
    # domain artifacts) rather than a real URL: fix in place when normalizing alone
    # repairs it, otherwise re-derive from the resume — and if the resume yields
    # nothing, clear the junk to N/A rather than keep publishing garbage.
    junk_ports: list[str] = []
    heal: dict = {}
    for c in ("Portfolio 1", "Portfolio 2", "Portfolio 3"):
        cur = str(vals.get(c, "") or "").strip()
        if _is_gap(cur):
            continue
        norm = _normalize_link_artifacts(cur)
        if not _looks_like_url(norm):
            junk_ports.append(c)
        elif norm != cur:
            heal[c] = _clean_url(norm)

    # A stored Location carrying the resume's own field LABEL or PDF layout artifacts is
    # WRONG, not missing, so gap-only healing never touches it - it just sits there. Cleaning
    # is deterministic (no re-extraction, no resume download), so it is done in place here,
    # before the need_resume branch, and therefore reaches rows with nothing else missing.
    # Live case APP-20260716-0335-0C84: 'Address  -  San  Jose, California' survived a FULL
    # recheck because the cell was non-blank. It also breaks matching - 'San  Jose' with two
    # spaces is not the 'San Jose' in the US-city list, so a real US location can miss.
    if not _is_gap(location):
        _cleaned_loc = clean_location_text(location)
        if _cleaned_loc and _cleaned_loc != location:
            heal["Location"] = _cleaned_loc
            logger.info(f"       Location : cleaned {location!r} -> {_cleaned_loc!r}")
            location = _cleaned_loc   # geo below must judge the CLEANED value

    # SAME REASONING for the other columns whose cleaners only ever ran at EXTRACTION time.
    # Every one of these was still wrong on the live sheet after a full recheck, because the
    # cell was non-blank and gap-only healing therefore skipped it:
    #   Current Skills   'Programming:Python, C++, ...'                      (A455)
    #   Looking For Role 'Mobile Application Developer Email: ashfaq@...'    (8006)
    #   Portfolio 1-3    'https://www.asu.edu/', a Drive file link           (25F9, 0DF7)
    # All three transforms are deterministic and need no resume, so they run here alongside
    # the Location clean and reach rows that have nothing missing at all.
    _stored_skills = str(vals.get("Current Skills", "") or "")
    if not _is_gap(_stored_skills):
        _cleaned_skills = normalize_skills(_stored_skills)
        if _cleaned_skills and not _is_gap(_cleaned_skills) and _cleaned_skills != _stored_skills:
            heal["Current Skills"] = _cleaned_skills
            logger.info(f"       Skills   : cleaned section labels/prose "
                        f"({len(_stored_skills)} -> {len(_cleaned_skills)} chars)")

    _stored_role = str(vals.get("Looking For Role", "") or "")
    if not _is_gap(_stored_role):
        _cleaned_role = clean_role_text(_stored_role)
        if _cleaned_role and _cleaned_role != _stored_role:
            heal["Looking For Role"] = _cleaned_role
            logger.info(f"       Role want: cleaned {_stored_role[:44]!r} -> {_cleaned_role!r}")

    _stored_ports = [str(vals.get(f"Portfolio {i}", "") or "") for i in (1, 2, 3)]
    if any(not _is_gap(p) for p in _stored_ports):
        _cp = _compact_portfolios(*_stored_ports)
        for _i, _new in enumerate(_cp, 1):
            _col = f"Portfolio {_i}"
            if _col in heal or _col in junk_ports:
                continue
            if _new != str(vals.get(_col, "") or "").strip():
                heal[_col] = _new
                logger.info(f"       {_col}: {str(vals.get(_col, ''))[:40]!r} -> {_new!r}")

    # Last Updated must never predate Received - a row cannot have been updated before it
    # arrived. The scoring path already clamps this in _validate_all_columns, but recheck
    # never ran that, so APP-20260716-0617-CB73 kept an inverted pair the P1 audit had been
    # flagging every run since 2026-07-31. Same in-place correction, same reasoning.
    _rec_raw = vals.get("Received Date")
    _upd_raw = vals.get("Last Updated Date")
    if not _is_gap(_rec_raw) and not _is_gap(_upd_raw):
        _rec_dt, _upd_dt = _parse_cell_datetime(_rec_raw), _parse_cell_datetime(_upd_raw)
        if _rec_dt and _upd_dt and _upd_dt < _rec_dt:
            heal["Last Updated Date"] = _rec_raw
            logger.info(f"       Dates    : Last Updated {_upd_raw!r} predated Received "
                        f"{_rec_raw!r} - clamped forward.")

    need_resume = need_resume or bool(junk_ports)
    if need_resume:
        derived = _derive_row_fields(client, vals, roles, rejected=rejected)
        if derived:
            for c in gaps:
                nv = derived.get(c)
                if nv and not _is_gap(nv):
                    # force_roles re-derives cells that already had a value, so skip the
                    # write when the fresh score matches what is already stored - otherwise
                    # every forced run would rewrite identical values and churn the sheet.
                    if str(nv).strip() == str(vals.get(c, "") or "").strip():
                        continue
                    heal[c] = nv
            # A broken phone that the resume couldn't improve on: clear it to a clean gap
            # rather than keep publishing the mask/fragment.
            if phone_bad and "Phone" not in heal:
                heal["Phone"] = MISSING_VALUE

            # An UNSUPPORTED stored Country is wrong, not missing - same class as phone_bad
            # and junk_ports above, so it is corrected rather than left alone. A fresh
            # derivation just read the resume AND the mail body; if that produced no country
            # at all while the sheet holds one, nothing in the candidate's own material
            # supports the stored value. Live case APP-20260716-1037-6CE3 (Peter Vishal):
            # Country said "India" though the resume never contains India/Indian/Hyderabad/
            # Telangana - the model had inferred it from a +91 dial code and a college name,
            # and it then pointed the geo filter at a rejection. Replaced with MISSING_VALUE
            # so we ask him instead of asserting something nobody extracted.
            # Guarded on `derived` being truthy: an unreadable resume must not wipe good data.
            if "Country" not in heal:
                _stored_ctry = str(vals.get("Country", "") or "").strip()
                if (not _is_gap(_stored_ctry)
                        and _is_gap(str(derived.get("Country", "") or ""))):
                    heal["Country"] = MISSING_VALUE
                    # CRITICAL: drop the local too. The geo classification below reads this
                    # variable, and it was still holding the discredited value - so a row
                    # whose Country we had just decided was unsupported was still being
                    # geo-REJECTED on the strength of it. Peter Vishal was rejected as
                    # non-USA off an 'India' that his resume never contained.
                    country = ""
                    logger.info(
                        f"       Country  : stored {_stored_ctry!r} is not supported by the "
                        f"resume or mail body - replaced with '{MISSING_VALUE}'; geo will "
                        f"re-decide without it.")
            for c in junk_ports:
                nv = str(derived.get(c) or "N/A").strip() or "N/A"
                if nv != str(vals.get(c, "") or "").strip():
                    heal[c] = nv
            if _is_gap(location):
                location = str(derived.get("Location", location) or "")
            if _is_gap(country):
                country = str(derived.get("Country", country) or "")

    # Category is a pure function of Suggested Role 1 (+ Current Skills) — safe to
    # re-derive when blank or when it's the 'General' fallback (heals rows scored before
    # a rules update). Also re-derived when the skill-based Mobile Apps override (added
    # 2026-07-24) disagrees with an existing specific category — that override only fires
    # on an unambiguous native-mobile tool stack (Kotlin/Swift/Flutter/...), so it's
    # trusted over a stale category even when the stale value isn't blank/General (e.g. a
    # mobile candidate whose winning JD title was 'Mobile Application Lead Developer',
    # landing them in 'Senior' via the 'lead' keyword instead).
    role1 = heal.get("Suggested Role 1", vals.get("Suggested Role 1", ""))
    if not _is_gap(role1):
        cur_skills = heal.get("Current Skills", vals.get("Current Skills", ""))
        cat = assign_category(role1, cur_skills)
        cur = str(vals.get("Category", "") or "").strip()
        stale_mobile_miscategorization = (
            cat == "Mobile Apps (Android IOS)" and cur and cur != cat)
        # When the role itself was just re-scored (force_roles), the Category derived from
        # the OLD role is stale by definition, so refresh it even if it is a specific value.
        role_was_rescored = "Suggested Role 1" in heal
        if cat and cat != cur and ((_is_gap(cur) or cur == "General")
                                    or stale_mobile_miscategorization
                                    or role_was_rescored):
            heal["Category"] = cat

    geo_decision = GeoDecision.CONFIRMED_US
    if geo_on:
        effective_phone = heal.get("Phone", vals.get("Phone", ""))
        effective_education = heal.get("Education", vals.get("Education", ""))
        geo_decision, _ = _classify_candidate_geo(
            location, country, effective_phone, education=effective_education)
        if geo_decision == GeoDecision.CONFIRMED_US:
            reconciled_country = reconcile_us_country(country, location, True)
            if (not _is_gap(reconciled_country)
                    and reconciled_country != str(vals.get("Country", "") or "").strip()):
                heal["Country"] = reconciled_country
                country = reconciled_country
        elif geo_decision == GeoDecision.UNKNOWN:
            if str(vals.get("Status", "") or "").strip() != STATUS_LOCATION_REVIEW:
                heal["Status"] = STATUS_LOCATION_REVIEW
            # NOTE (2026-08-04): Country is NOT inferred from a US phone + US university
            # here either - see the matching note in the scoring path. A school and a phone
            # number do not establish current residence; the cell stays MISSING_VALUE and
            # the candidate is asked directly.

    stored = _stored_resume_names(
        str(vals.get("Application ID", "") or "").strip(),
        vals.get("Original Filename", ""),
        full_name=vals.get("Full Name"),
        category=vals.get("Category"),
        resume_url=vals.get("Resume URL"),
    )
    if stored:
        subpath = _cfg.dated_subpath(_parse_received(vals.get("Received Date")))
        if rejected:
            subpath = _rejected_subpath(subpath)
        url, folder_path = _resume_url_and_path(client, stored, subpath)
        if url and str(vals.get("Resume URL", "") or "").strip() != url:
            heal["Resume URL"] = url
        if folder_path and str(vals.get("Resume Folder Path", "") or "").strip() != folder_path:
            heal["Resume Folder Path"] = folder_path

    # A blank (not just gap-literal) Portfolio cell that resume re-derivation couldn't
    # fill (no resume found, or the resume simply has no such link) must still end up
    # "N/A", never a raw blank string - the live scoring path already guarantees this
    # unconditionally (_na_if_gap), but a heal pass previously only touched a blank
    # Portfolio cell as a side effect of a successful re-derivation, so a candidate whose
    # resume file could no longer be located stayed blank forever across repeated heals
    # (real case: 'Muhammad Waqas', Portfolio 2/3 blank with an undownloadable resume).
    for c in ("Portfolio 1", "Portfolio 2", "Portfolio 3"):
        if c in heal:
            continue
        cur = str(vals.get(c, "") or "")
        if _is_gap(cur) and cur != "N/A":
            heal[c] = "N/A"

    # Same idea for the candidate-owed columns, but with the OTHER placeholder: anything
    # still empty after every derivation attempt reads 'Missing', so the sheet never shows a
    # bare blank a reader has to guess about. Deliberately last, and skipped for any column
    # already being healed with a real value in this pass. Note this cannot be done by
    # stamping `derived` instead - the heal loop above drops any derived value that is a gap,
    # and MISSING_VALUE is itself a gap, so it would be filtered straight back out.
    for c in _MISSING_IF_BLANK_COLS:
        if c in heal:
            continue
        cur = str(vals.get(c, "") or "")
        if _is_gap(cur) and cur.strip() != MISSING_VALUE:
            heal[c] = MISSING_VALUE
    return heal, geo_decision


def recheck_all_rows(dry_run: bool = False, force_roles: bool = False) -> dict:
    """Reconcile the WHOLE workbook: re-check every row (main table + Rejected sheet)
    one-by-one, healing gaps and re-validating the USA filter. Idempotent and keyed on
    Application ID, so nothing is duplicated: complete, correctly-placed rows are left
    untouched; incomplete rows get their blanks filled; a row on the wrong sheet is moved
    (main→Rejected if now non-USA, Rejected→main if now USA); and a candidate that ended
    up on both sheets is de-duplicated.

    `force_roles` also re-scores already-populated Suggested Role 1/2/3 - see _plan_heal.
    Off by default because a whole-workbook forced re-score is the widest-blast-radius
    operation in P2."""
    if storage_backend() == "sqlite":
        from sharepoint_client import SharePointClient
        from .local_pipeline import run_local_pipeline
        candidate_client = SharePointClient()
        if hasattr(candidate_client, "site_id"):
            return run_local_pipeline(candidate_client, dry_run=dry_run, force=True)
    from sharepoint_client import SharePointClient, SharePointError, check_graph_reachable

    logger.info("  Step 1/4  Checking connectivity...")
    _net_ok, _net_msg = check_graph_reachable()
    if not _net_ok:
        logger.error(f"  FAILED    {_net_msg}")
        return {"error": _net_msg}
    logger.info(f"  OK        {_net_msg}")
    _ok, _brain = ollama_health()
    (logger.info if _ok else logger.warning)(f"  {_brain}")
    if _cfg.REQUIRE_AI and not _ok:
        # require_ai means a row is never published from anything but a real model read.
        # With the brain down EVERY row would degrade, so proceeding would write a whole
        # batch of regex output. Stop instead: nothing is written, the queue keeps its
        # place, and the next run picks it all up unchanged.
        _msg = (f"Ollama is required but unavailable ({_brain}). Nothing was scored and no "
                f"row was changed - the queue is untouched. Start Ollama and re-run, or set "
                f"ai_extraction.ollama.require_ai=false to allow offline-parser fallback.")
        logger.error(f"  FAILED    {_msg}")
        return {"error": _msg, "processed": 0, "rejected": 0,
                "location_review": 0, "errors": 0, "deferred": 0}

    try:
        client = SharePointClient()
    except SharePointError as e:
        logger.error(f"  FAILED    SharePoint not configured: {e}")
        return {"error": str(e)}
    logger.info(f"  OK        Connected to {client.hostname}")

    logger.info("  Step 2/4  Checking workbook & columns...")
    if dry_run:
        logger.info("  SKIP      Dry-run: not creating/altering workbook or columns.")
    else:
        if not _ensure_workbook_or_alert(client, dry_run):
            return {"error": "candidate workbook could not be confirmed - see log; "
                              "P2 never auto-creates one"}
        _ensure_schema(client)

    logger.info("  Step 3/4  Reading ALL rows (main + Rejected)...")
    try:
        main_rows = client.list_rows()
    except SharePointError as e:
        logger.error(f"  FAILED    Could not read the candidate table: {e}")
        return {"error": str(e)}
    rej_rows = client.list_rejected_rows()
    logger.info(f"  OK        {len(main_rows)} main row(s), {len(rej_rows)} rejected row(s).")

    roles = get_active_roles()
    main_ids = {str(r["values"].get("Application ID", "")).strip() for r in main_rows}
    main_emails = {_bare_email(r["values"].get("Email", "")) for r in main_rows} - {""}
    main_phones = {_normalize_phone_key(r["values"].get("Phone", "")) for r in main_rows} - {""}
    rej_ids = {str(r["values"].get("Application ID", "")).strip() for r in rej_rows}
    rej_emails = {_bare_email(r["values"].get("Email", "")) for r in rej_rows} - {""}
    rej_phones = {_normalize_phone_key(r["values"].get("Phone", "")) for r in rej_rows} - {""}

    healed = moved = restored = review_restored = deduped = untouched = errors = 0
    del_main: list[int] = []
    del_rej: list[int] = []

    logger.info("  Step 4/4  Reconciling main table...")
    for row in main_rows:
        vals, index = row["values"], row["index"]
        app_id = str(vals.get("Application ID", "")).strip()
        try:
            heal, geo_decision = _plan_heal(client, vals, roles, force_roles=force_roles)
            email_key = _bare_email(vals.get("Email", ""))
            phone_key = _normalize_phone_key(vals.get("Phone", ""))
            if geo_decision == GeoDecision.CONFIRMED_NON_US:
                if ((app_id and app_id in rej_ids)
                        or (email_key and email_key in rej_emails)
                        or (phone_key and phone_key in rej_phones)):
                    logger.info(f"       {app_id}: already on Rejected — de-duplicating main copy "
                                f"(no second decline email).")
                    if not dry_run:
                        del_main.append(index)
                    deduped += 1
                else:
                    merged = dict(vals); merged.update(heal)
                    merged["Status"] = heal.get("Status", _cfg.STATUS_REJECTED)
                    if not dry_run:
                        stored = _stored_resume_names(
                            app_id, vals.get("Original Filename", ""),
                            full_name=vals.get("Full Name"), category=vals.get("Category"),
                            resume_url=vals.get("Resume URL"))
                        if stored:
                            subpath = _cfg.dated_subpath(_parse_received(vals.get("Received Date")))
                            _move_resumes(client, stored, subpath, to_rejected=True)
                            url, path = _resume_url_and_path(
                                client, stored, _rejected_subpath(subpath))
                            if url:
                                merged["Resume URL"] = url
                                merged["Resume Folder Path"] = path
                        merged = _complete_row_before_reject(merged, app_id)
                        _ensure_rejected_period_separator(
                            client, merged.get("Received Date"))
                        _store(client).add(merged, "rejected")
                        del_main.append(index)
                    if app_id:
                        rej_ids.add(app_id)
                    if email_key:
                        rej_emails.add(email_key)
                    if phone_key:
                        rej_phones.add(phone_key)
                    moved += 1
                    logger.info(f"       {app_id or '(no ref)'}: now non-USA — moved to Rejected "
                                f"(decline email queued for the send pass).")
            elif heal:
                st = str(vals.get("Status", "") or "").strip()
                if (geo_decision == GeoDecision.CONFIRMED_US
                        and (_is_gap(st) or st in (
                            _STATUS_NEW, _cfg.STATUS_NEEDS_REVIEW, STATUS_LOCATION_REVIEW))):
                    heal["Status"] = STATUS_SCORED
                if not dry_run:
                    subpath = _cfg.dated_subpath(_parse_received(vals.get("Received Date")))
                    _rename_resume_for_healed_row(client, heal, vals, app_id, subpath, dry_run)
                    _store(client).save_by_id(_row_key(app_id, vals), heal, current_values=vals, hint=index)
                healed += 1
                logger.info(f"       {app_id or '(no ref)'}: healed {', '.join(heal)}")
            else:
                untouched += 1
        except Exception as e:
            errors += 1
            logger.error(f"       ERROR    : {app_id} failed: {e} (left as-is).")
        finally:
            _cleanup_temp_files()

    for row in rej_rows:
        vals, index = row["values"], row["index"]
        app_id = str(vals.get("Application ID", "")).strip()
        try:
            heal, geo_decision = _plan_heal(client, vals, roles, rejected=True)
            email_key = _bare_email(vals.get("Email", ""))
            phone_key = _normalize_phone_key(vals.get("Phone", ""))
            if geo_decision != GeoDecision.CONFIRMED_NON_US:
                if ((app_id and app_id in main_ids)
                        or (email_key and email_key in main_emails)
                        or (phone_key and phone_key in main_phones)):
                    logger.info(f"       {app_id}: already on main — de-duplicating Rejected copy.")
                    if not dry_run:
                        del_rej.append(index)
                    deduped += 1
                else:
                    merged = dict(vals); merged.update(heal)
                    if geo_decision == GeoDecision.CONFIRMED_US:
                        merged["Status"] = STATUS_SCORED
                    else:
                        merged["Status"] = STATUS_LOCATION_REVIEW
                        merged[_INFO_REQUEST_COL] = (
                            "Manual follow-up required after geo recovery")
                    if not dry_run:
                        stored = _stored_resume_names(
                            app_id, vals.get("Original Filename", ""),
                            full_name=vals.get("Full Name"), category=vals.get("Category"),
                            resume_url=vals.get("Resume URL"))
                        if stored:
                            subpath = _cfg.dated_subpath(_parse_received(vals.get("Received Date")))
                            _move_resumes(client, stored, subpath, to_rejected=False)
                            url, path = _resume_url_and_path(client, stored, subpath)
                            if url:
                                merged["Resume URL"] = url
                                merged["Resume Folder Path"] = path
                        _ensure_main_period_separator(client, merged.get("Received Date"))
                        _store(client).add(merged, "main")
                        del_rej.append(index)
                    if app_id:
                        main_ids.add(app_id)
                    if email_key:
                        main_emails.add(email_key)
                    if phone_key:
                        main_phones.add(phone_key)
                    if geo_decision == GeoDecision.CONFIRMED_US:
                        restored += 1
                        logger.info(f"       {app_id or '(no ref)'}: confirmed USA - "
                                    "restored to main table.")
                    else:
                        review_restored += 1
                        logger.info(f"       {app_id or '(no ref)'}: location unresolved - "
                                    "restored to main for manual confirmation (no decline).")
            else:
                # Still belongs on Rejected — make sure its resume file(s) actually live in that
                # month's 'Rejected' sub-bucket. No-op once already moved; this is what migrates
                # rows rejected before the folder-split existed, the first time recheck-all runs.
                if not dry_run:
                    stored = _stored_resume_names(
                        app_id, vals.get("Original Filename", ""),
                        full_name=vals.get("Full Name"), category=vals.get("Category"),
                        resume_url=vals.get("Resume URL"))
                    if stored:
                        subpath = _cfg.dated_subpath(_parse_received(vals.get("Received Date")))
                        if _move_resumes(client, stored, subpath, to_rejected=True):
                            url, path = _resume_url_and_path(
                                client, stored, _rejected_subpath(subpath))
                            if url:
                                heal["Resume URL"] = url
                                heal["Resume Folder Path"] = path
                if str(vals.get("Status", "") or "").strip() == _cfg.STATUS_PROCESSING_FAILED:
                    heal["Status"] = _cfg.STATUS_REJECTED
                    if str(vals.get(_DECLINE_COL, "") or "").startswith("N/A - processing error"):
                        heal[_DECLINE_COL] = ""
                if heal:
                    if heal.get("Status") not in (_cfg.STATUS_REJECTED, _cfg.STATUS_LOCATION_UNCONFIRMED):
                        heal.pop("Status", None)   # keep it Rejected
                    if not dry_run:
                        rej_subpath = _rejected_subpath(
                            _cfg.dated_subpath(_parse_received(vals.get("Received Date"))))
                        _rename_resume_for_healed_row(client, heal, vals, app_id, rej_subpath, dry_run)
                        _store(client).save_by_id(_row_key(app_id, vals), heal, current_values=vals, sheet="rejected", hint=index)
                    healed += 1
                    logger.info(f"       {app_id or '(no ref)'}: healed in Rejected ({', '.join(heal)})")
                else:
                    untouched += 1
        except Exception as e:
            errors += 1
            logger.error(f"       ERROR    : {app_id} failed: {e} (left as-is).")
        finally:
            _cleanup_temp_files()

    # Deletes last, in descending index order, so positional indices never shift mid-pass.
    if not dry_run:
        for idx in sorted(set(del_main), reverse=True):
            try:
                _store(client).delete_at(idx)
            except SharePointError as e:
                logger.warning(f"  WARNING   Could not delete main row {idx}: {e}")
        for idx in sorted(set(del_rej), reverse=True):
            try:
                _store(client).delete_at(idx, "rejected")
            except SharePointError as e:
                logger.warning(f"  WARNING   Could not delete rejected row {idx}: {e}")

    _cleanup_empty_dirs()

    # After all moves/deletes are settled (indices stable again), send any owed
    # decline emails — same idempotent stamped pass the scoring run uses.
    declines = _send_pending_declines(client, dry_run=dry_run)

    logger.info("  ── Reconcile summary ───────────────────────────────")
    logger.info(f"       Healed     : {healed}")
    logger.info(f"       Moved-Rej  : {moved}")
    logger.info(f"       Restored   : {restored}")
    logger.info(f"       Geo review : {review_restored} restored to Main for confirmation")
    logger.info(f"       De-duped   : {deduped}")
    logger.info(f"       Untouched  : {untouched}")
    logger.info(f"       Declines   : {declines.get('sent', 0)} sent | "
                f"{declines.get('failed', 0)} failed | {declines.get('skipped', 0)} skipped")
    logger.info(f"       Errors     : {errors}")
    return {"healed": healed, "moved": moved, "restored": restored,
            "review_restored": review_restored,
            "deduped": deduped, "untouched": untouched, "errors": errors,
            "declines": declines}


def recover_location_rejections(dry_run: bool = False) -> dict:
    """Recover historical geo rejects without rescoring candidates or sending mail.

    Every row carrying the retired "Location Not Confirmed" rejection is audited.
    Normal non-USA rejections are included only when their stored evidence now resolves
    to UNKNOWN/US under the tri-state policy. Confirmed non-US rows remain untouched.
    """
    if storage_backend() == "sqlite":
        from sharepoint_client import SharePointClient
        from .local_pipeline import run_local_pipeline
        candidate_client = SharePointClient()
        if hasattr(candidate_client, "site_id"):
            return run_local_pipeline(candidate_client, dry_run=dry_run, force=True)
    from sharepoint_client import SharePointClient, SharePointError, check_graph_reachable

    logger.info("  Geo recovery preflight: checking connectivity...")
    net_ok, net_msg = check_graph_reachable()
    if not net_ok:
        logger.error(f"  FAILED    {net_msg}")
        return {"error": net_msg}
    try:
        client = SharePointClient()
        rejected_rows = client.list_rejected_rows()
    except SharePointError as e:
        logger.error(f"  FAILED    Could not read Rejected candidates: {e}")
        return {"error": str(e)}

    targets: list[str] = []
    confirmed_non_us = processing_errors = seed_rows = 0
    for row in rejected_rows:
        vals = row.get("values", {})
        app_id = str(vals.get("Application ID", "") or "").strip()
        status = str(vals.get("Status", "") or "").strip()
        if not app_id:
            seed_rows += 1
            continue
        if status == _cfg.STATUS_PROCESSING_FAILED:
            processing_errors += 1
            continue
        if status not in (_cfg.STATUS_LOCATION_UNCONFIRMED, _cfg.STATUS_REJECTED):
            continue

        # All historical "location not confirmed" rows need the focused resume
        # check. For ordinary geo rejects, avoid reopening rows whose stored current
        # location is still independently corroborated as non-US.
        if status == _cfg.STATUS_REJECTED:
            decision, reason = _classify_candidate_geo(
                vals.get("Location", ""), vals.get("Country", ""),
                vals.get("Phone", ""), education=vals.get("Education", ""),
                resume_text="")
            if decision == GeoDecision.CONFIRMED_NON_US:
                confirmed_non_us += 1
                logger.info(f"       {app_id}: retained candidate preflight "
                            f"({reason})")
                continue
        targets.append(app_id)

    logger.info(f"  Geo recovery scope: {len(targets)} candidate(s); "
                f"{confirmed_non_us} confirmed non-US skipped; "
                f"{processing_errors} processing-error row(s) skipped; "
                f"{seed_rows} seed/blank row(s) skipped.")
    if not targets:
        return {
            "requested": 0, "found": 0, "healed": 0, "moved": 0,
            "restored": 0, "review_restored": 0, "deduped": 0,
            "untouched": confirmed_non_us, "errors": 0,
            "declines": {"sent": 0, "failed": 0, "skipped": 0},
            "client_export": "",
        }

    return recheck_selected_rows(
        targets, dry_run=dry_run, geo_only=True, send_declines=False)


def recheck_selected_rows(app_ids: list[str] | set[str], dry_run: bool = False,
                          geo_only: bool = False,
                          send_declines: bool = True,
                          force_roles: bool = False) -> dict:
    """Targeted version of recheck_all_rows for known Application IDs.

    It uses the same heal/geo/move rules, but only touches rows whose Application ID is
    in `app_ids`, and it only sends/stamps decline emails for those same IDs.

    `force_roles` re-scores Suggested Role 1/2/3 (and the Category derived from them) even
    when those cells already hold a value - see _plan_heal. Ignored under `geo_only`, which
    deliberately does no role scoring at all.

    `geo_only` uses the deterministic historical-recovery planner: no Ollama, role
    scoring, portfolio inference, or category changes. It also forces decline sending
    off, regardless of `send_declines`.
    """
    if storage_backend() == "sqlite":
        from sharepoint_client import SharePointClient
        from .local_pipeline import run_local_pipeline
        candidate_client = SharePointClient()
        if hasattr(candidate_client, "site_id"):
            return run_local_pipeline(candidate_client, dry_run=dry_run, app_ids=app_ids)
    from sharepoint_client import SharePointClient, SharePointError, check_graph_reachable

    targets = {str(x).strip() for x in app_ids if str(x).strip()}
    if not targets:
        logger.error("  FAILED    No Application IDs supplied.")
        return {"error": "no Application IDs supplied"}

    logger.info("  Step 1/4  Checking connectivity...")
    _net_ok, _net_msg = check_graph_reachable()
    if not _net_ok:
        logger.error(f"  FAILED    {_net_msg}")
        return {"error": _net_msg}
    logger.info(f"  OK        {_net_msg}")
    if geo_only:
        logger.info("  Brain     : skipped (deterministic geo-only recovery)")
    else:
        _ok, _brain = ollama_health()
        (logger.info if _ok else logger.warning)(f"  {_brain}")

    try:
        client = SharePointClient()
    except SharePointError as e:
        logger.error(f"  FAILED    SharePoint not configured: {e}")
        return {"error": str(e)}
    logger.info(f"  OK        Connected to {client.hostname}")

    logger.info("  Step 2/4  Checking workbook & columns...")
    if dry_run:
        logger.info("  SKIP      Dry-run: not creating/altering workbook or columns.")
    else:
        if not _ensure_workbook_or_alert(client, dry_run):
            return {"error": "candidate workbook could not be confirmed - see log; "
                              "P2 never auto-creates one"}
        if geo_only:
            logger.info("  SKIP      Geo-only recovery leaves global schema/data "
                        "maintenance untouched.")
        else:
            _ensure_schema(client)

    logger.info("  Step 3/4  Reading rows (main + Rejected)...")
    try:
        main_rows = client.list_rows()
    except SharePointError as e:
        logger.error(f"  FAILED    Could not read the candidate table: {e}")
        return {"error": str(e)}
    rej_rows = client.list_rejected_rows()
    logger.info(f"  OK        {len(main_rows)} main row(s), {len(rej_rows)} rejected row(s).")

    selected_main = [
        r for r in main_rows
        if str(r["values"].get("Application ID", "")).strip() in targets
    ]
    selected_rej = [
        r for r in rej_rows
        if str(r["values"].get("Application ID", "")).strip() in targets
    ]
    found_ids = {
        str(r["values"].get("Application ID", "")).strip()
        for r in selected_main + selected_rej
    }
    for missing_id in sorted(targets - found_ids):
        logger.warning(f"       {missing_id}: not found in main or Rejected.")

    roles = [] if geo_only else get_active_roles()
    main_ids = {str(r["values"].get("Application ID", "")).strip() for r in main_rows}
    main_emails = {_bare_email(r["values"].get("Email", "")) for r in main_rows} - {""}
    main_phones = {_normalize_phone_key(r["values"].get("Phone", "")) for r in main_rows} - {""}
    rej_ids = {str(r["values"].get("Application ID", "")).strip() for r in rej_rows}
    rej_emails = {_bare_email(r["values"].get("Email", "")) for r in rej_rows} - {""}
    rej_phones = {_normalize_phone_key(r["values"].get("Phone", "")) for r in rej_rows} - {""}

    healed = moved = restored = review_restored = deduped = untouched = errors = 0
    del_main: list[int] = []
    del_rej: list[int] = []

    logger.info("  Step 4/4  Reconciling selected row(s)...")
    for row in selected_main:
        vals, index = row["values"], row["index"]
        app_id = str(vals.get("Application ID", "")).strip()
        try:
            if geo_only:
                heal, geo_decision, geo_reason = _plan_geo_recovery(
                    client, vals, rejected=False)
                logger.info(f"       {app_id or '(no ref)'}: geo-only verdict "
                            f"{geo_decision.value} ({geo_reason})")
            else:
                heal, geo_decision = _plan_heal(client, vals, roles, force_roles=force_roles)
            email_key = _bare_email(vals.get("Email", ""))
            phone_key = _normalize_phone_key(vals.get("Phone", ""))
            if geo_decision == GeoDecision.CONFIRMED_NON_US:
                if ((app_id and app_id in rej_ids)
                        or (email_key and email_key in rej_emails)
                        or (phone_key and phone_key in rej_phones)):
                    logger.info(f"       {app_id}: already on Rejected - de-duplicating main copy.")
                    if not dry_run:
                        del_main.append(index)
                    deduped += 1
                else:
                    merged = dict(vals)
                    merged.update(heal)
                    merged["Status"] = heal.get("Status", _cfg.STATUS_REJECTED)
                    if not dry_run:
                        stored = _stored_resume_names(
                            app_id, vals.get("Original Filename", ""),
                            full_name=vals.get("Full Name"), category=vals.get("Category"),
                            resume_url=vals.get("Resume URL"))
                        if stored:
                            subpath = _cfg.dated_subpath(_parse_received(vals.get("Received Date")))
                            _move_resumes(client, stored, subpath, to_rejected=True)
                            url, path = _resume_url_and_path(client, stored, _rejected_subpath(subpath))
                            if url:
                                merged["Resume URL"] = url
                                merged["Resume Folder Path"] = path
                        merged = _complete_row_before_reject(merged, app_id)
                        _ensure_rejected_period_separator(
                            client, merged.get("Received Date"))
                        _store(client).add(merged, "rejected")
                        del_main.append(index)
                    if app_id:
                        rej_ids.add(app_id)
                    if email_key:
                        rej_emails.add(email_key)
                    if phone_key:
                        rej_phones.add(phone_key)
                    moved += 1
                    logger.info(f"       {app_id or '(no ref)'}: now non-USA - moved to Rejected "
                                f"(decline email queued for the targeted send pass).")
            elif heal:
                st = str(vals.get("Status", "") or "").strip()
                if (geo_decision == GeoDecision.CONFIRMED_US
                        and (_is_gap(st) or st in (
                            _STATUS_NEW, _cfg.STATUS_NEEDS_REVIEW, STATUS_LOCATION_REVIEW))):
                    heal["Status"] = STATUS_SCORED
                if not dry_run:
                    subpath = _cfg.dated_subpath(_parse_received(vals.get("Received Date")))
                    if geo_only:
                        # Idempotent repair for a partial/legacy restore: a Main row
                        # can still have its PDF and URL in the Rejected bucket when
                        # an older category remains embedded in the physical filename.
                        stored = _stored_resume_names(
                            app_id, vals.get("Original Filename", ""),
                            full_name=vals.get("Full Name"),
                            category=vals.get("Category"),
                            resume_url=vals.get("Resume URL"))
                        if stored:
                            _move_resumes(
                                client, stored, subpath, to_rejected=False)
                            url, path = _resume_url_and_path(
                                client, stored, subpath)
                            if url:
                                heal["Resume URL"] = url
                                heal["Resume Folder Path"] = path
                    _rename_resume_for_healed_row(client, heal, vals, app_id, subpath, dry_run)
                    _store(client).save_by_id(_row_key(app_id, vals), heal, current_values=vals, hint=index)
                healed += 1
                logger.info(f"       {app_id or '(no ref)'}: healed {', '.join(heal)}")
            else:
                untouched += 1
        except Exception as e:
            errors += 1
            logger.error(f"       ERROR    : {app_id} failed: {e} (left as-is).")
        finally:
            _cleanup_temp_files()

    for row in selected_rej:
        vals, index = row["values"], row["index"]
        app_id = str(vals.get("Application ID", "")).strip()
        try:
            if geo_only:
                heal, geo_decision, geo_reason = _plan_geo_recovery(
                    client, vals, rejected=True)
                logger.info(f"       {app_id or '(no ref)'}: geo-only verdict "
                            f"{geo_decision.value} ({geo_reason})")
            else:
                heal, geo_decision = _plan_heal(client, vals, roles, rejected=True)
            email_key = _bare_email(vals.get("Email", ""))
            phone_key = _normalize_phone_key(vals.get("Phone", ""))
            if geo_decision != GeoDecision.CONFIRMED_NON_US:
                if ((app_id and app_id in main_ids)
                        or (email_key and email_key in main_emails)
                        or (phone_key and phone_key in main_phones)):
                    logger.info(f"       {app_id}: already on main - de-duplicating Rejected copy.")
                    if not dry_run:
                        del_rej.append(index)
                    deduped += 1
                else:
                    merged = dict(vals)
                    merged.update(heal)
                    if geo_decision == GeoDecision.CONFIRMED_US:
                        merged["Status"] = STATUS_SCORED
                    else:
                        merged["Status"] = STATUS_LOCATION_REVIEW
                        merged[_INFO_REQUEST_COL] = (
                            "Manual follow-up required after geo recovery")
                    if not dry_run:
                        stored = _stored_resume_names(
                            app_id, vals.get("Original Filename", ""),
                            full_name=vals.get("Full Name"), category=vals.get("Category"),
                            resume_url=vals.get("Resume URL"))
                        if stored:
                            subpath = _cfg.dated_subpath(_parse_received(vals.get("Received Date")))
                            _move_resumes(client, stored, subpath, to_rejected=False)
                            url, path = _resume_url_and_path(client, stored, subpath)
                            if url:
                                merged["Resume URL"] = url
                                merged["Resume Folder Path"] = path
                        _ensure_main_period_separator(client, merged.get("Received Date"))
                        _store(client).add(merged, "main")
                        del_rej.append(index)
                    if app_id:
                        main_ids.add(app_id)
                    if email_key:
                        main_emails.add(email_key)
                    if phone_key:
                        main_phones.add(phone_key)
                    if geo_decision == GeoDecision.CONFIRMED_US:
                        restored += 1
                        logger.info(f"       {app_id or '(no ref)'}: confirmed USA - "
                                    "restored to main table.")
                    else:
                        review_restored += 1
                        logger.info(f"       {app_id or '(no ref)'}: location unresolved - "
                                    "restored to main for manual confirmation (no decline).")
            else:
                if not dry_run:
                    stored = _stored_resume_names(
                        app_id, vals.get("Original Filename", ""),
                        full_name=vals.get("Full Name"), category=vals.get("Category"),
                        resume_url=vals.get("Resume URL"))
                    if stored:
                        subpath = _cfg.dated_subpath(_parse_received(vals.get("Received Date")))
                        if _move_resumes(client, stored, subpath, to_rejected=True):
                            url, path = _resume_url_and_path(client, stored, _rejected_subpath(subpath))
                            if url:
                                heal["Resume URL"] = url
                                heal["Resume Folder Path"] = path
                if str(vals.get("Status", "") or "").strip() == _cfg.STATUS_PROCESSING_FAILED:
                    heal["Status"] = _cfg.STATUS_REJECTED
                    if str(vals.get(_DECLINE_COL, "") or "").startswith("N/A - processing error"):
                        heal[_DECLINE_COL] = ""
                if heal:
                    if heal.get("Status") not in (_cfg.STATUS_REJECTED, _cfg.STATUS_LOCATION_UNCONFIRMED):
                        heal.pop("Status", None)
                    if not dry_run:
                        rej_subpath = _rejected_subpath(
                            _cfg.dated_subpath(_parse_received(vals.get("Received Date"))))
                        _rename_resume_for_healed_row(client, heal, vals, app_id, rej_subpath, dry_run)
                        _store(client).save_by_id(_row_key(app_id, vals), heal, current_values=vals, sheet="rejected", hint=index)
                    healed += 1
                    logger.info(f"       {app_id or '(no ref)'}: healed in Rejected ({', '.join(heal)})")
                else:
                    untouched += 1
        except Exception as e:
            errors += 1
            logger.error(f"       ERROR    : {app_id} failed: {e} (left as-is).")
        finally:
            _cleanup_temp_files()

    if not dry_run:
        for idx in sorted(set(del_main), reverse=True):
            try:
                _store(client).delete_at(idx)
            except SharePointError as e:
                logger.warning(f"  WARNING   Could not delete main row {idx}: {e}")
        for idx in sorted(set(del_rej), reverse=True):
            try:
                _store(client).delete_at(idx, "rejected")
            except SharePointError as e:
                logger.warning(f"  WARNING   Could not delete rejected row {idx}: {e}")

    _cleanup_empty_dirs()
    if send_declines and not geo_only:
        declines = _send_pending_declines(
            client, dry_run=dry_run, only_app_ids=targets)
    else:
        declines = {"sent": 0, "failed": 0, "skipped": 0}
        logger.info("  Declines  : disabled for geo recovery.")

    export_path = ""
    if not dry_run:
        try:
            export_path = export_client_results(client, upload_to_sharepoint=True)
            if export_path:
                logger.info(f"  Export    : client workbook updated -> {export_path}")
        except Exception as e:
            logger.warning(f"  WARNING   Client workbook export failed: {e}")

    logger.info("  -- Targeted reconcile summary --------------------------------")
    logger.info(f"       Requested : {len(targets)}")
    logger.info(f"       Found     : {len(found_ids)}")
    logger.info(f"       Healed    : {healed}")
    logger.info(f"       Moved-Rej : {moved}")
    logger.info(f"       Restored  : {restored}")
    logger.info(f"       Geo review: {review_restored} restored to Main for confirmation")
    logger.info(f"       De-duped  : {deduped}")
    logger.info(f"       Untouched : {untouched}")
    logger.info(f"       Declines  : {declines.get('sent', 0)} sent | "
                f"{declines.get('failed', 0)} failed | {declines.get('skipped', 0)} skipped")
    logger.info(f"       Errors    : {errors}")
    return {"requested": len(targets), "found": len(found_ids), "healed": healed,
            "moved": moved, "restored": restored,
            "review_restored": review_restored, "deduped": deduped,
            "untouched": untouched, "errors": errors, "declines": declines,
            "client_export": export_path}


def validate_row(app_id: str) -> None:
    """Read-only diagnostic: print every column for one candidate by Application ID.

    Searches the main table first, then the Rejected sheet.
    After printing stored values, re-derives fields from the resume (offline parser
    + Ollama if running) and shows a side-by-side comparison so you can spot
    mismatches or stale extractions. Nothing is written."""
    from sharepoint_client import SharePointClient, SharePointError
    from hiring_agent.config import COLUMNS

    logger.info("  Connecting to SharePoint...")
    try:
        client = SharePointClient()
    except SharePointError as e:
        logger.error(f"  FAILED: {e}")
        return
    logger.info(f"  Connected to {client.hostname}")

    # Find the row (main first, then Rejected).
    logger.info(f"  Searching for Application ID: {app_id}")
    vals = None
    sheet = None
    try:
        for row in client.list_rows():
            if str(row["values"].get("Application ID", "")).strip() == app_id:
                vals = row["values"]
                sheet = "Main"
                break
    except SharePointError as e:
        logger.error(f"  Could not read main table: {e}")
        return

    if vals is None:
        try:
            for row in client.list_rejected_rows():
                if str(row["values"].get("Application ID", "")).strip() == app_id:
                    vals = row["values"]
                    sheet = "Rejected"
                    break
        except SharePointError as e:
            logger.warning(f"  Could not read Rejected sheet: {e}")

    if vals is None:
        logger.error(f"  Application ID '{app_id}' not found in any sheet.")
        return

    # ── Print all configured columns ────────────────────────────────────────
    logger.info("")
    logger.info(f"  Sheet : {sheet}")
    logger.info("  " + "=" * 60)
    logger.info(f"  {'#':<4} {'Column':<24} {'Value':<28} {'Status'}")
    logger.info("  " + "-" * 60)
    for i, col in enumerate(COLUMNS, 1):
        v = str(vals.get(col, "") or "").strip()
        status = "N/A" if _is_gap(v) else "OK"
        display = (v[:26] + "..") if len(v) > 28 else v
        logger.info(f"  {i:<4} {col:<24} {display:<28} {status}")
    logger.info("  " + "=" * 60)

    # ── Re-derive from resume and compare ──────────────────────────────────
    logger.info("")
    logger.info("  Re-deriving from resume (read-only comparison)...")
    roles = get_active_roles()
    derived = _derive_row_fields(client, vals, roles, rejected=(sheet == "Rejected"))

    if derived is None:
        logger.warning("  Could not download or read the resume file — comparison skipped.")
        return

    # Columns that P2 owns and can re-derive — same set as _P2_CONTENT_COLS minus Resume URL/Path
    # (those are file-system artifacts, not derivable fields shown in the comparison table).
    _COL_KEY = [
        "Full Name", "Phone", "Location", "Country", "Current Skills", "Looking For Role",
        "Suggested Role 1", "Suggested Role 2", "Suggested Role 3",
        "Portfolio 1", "Portfolio 2", "Portfolio 3",
    ]

    # Category from the (freshly derived) Suggested Role 1.
    new_role1 = derived.get("Suggested Role 1", "")
    new_category = assign_category(new_role1, derived.get("Current Skills", "")) if new_role1 else ""

    # USA verdict from the (freshly derived) location + country.
    new_location = derived.get("Location", "")
    new_country  = derived.get("Country", "")
    if GEO_FILTER_USA_ONLY:
        geo_decision, usa_reason = _classify_candidate_geo(
            new_location, country=new_country, phone=derived.get("Phone", ""),
            education=derived.get("Education", ""))
        geo_verdict = f"{geo_decision.value} ({usa_reason})"
    else:
        geo_verdict = "geo filter off"

    logger.info("")
    logger.info(f"  {'Column':<24} {'Stored':<28} {'Would derive':<28} {'Match?'}")
    logger.info("  " + "-" * 90)
    any_mismatch = False
    for col in _COL_KEY:
        stored  = str(vals.get(col, "") or "").strip()
        derived_val = str(derived.get(col, "") or "").strip()
        s_disp = (stored[:26] + "..") if len(stored) > 28 else stored
        d_disp = (derived_val[:26] + "..") if len(derived_val) > 28 else derived_val
        match = "=" if stored == derived_val else ("NEW" if _is_gap(stored) and not _is_gap(derived_val) else "DIFF")
        if match != "=":
            any_mismatch = True
        logger.info(f"  {col:<24} {s_disp:<28} {d_disp:<28} {match}")

    # Category
    stored_cat = str(vals.get("Category", "") or "").strip()
    cat_match = "=" if stored_cat == new_category else "DIFF"
    if cat_match != "=":
        any_mismatch = True
    logger.info(f"  {'Category':<24} {stored_cat:<28} {new_category:<28} {cat_match}")

    logger.info("  " + "-" * 90)
    logger.info(f"  Geo verdict (from derived location): {geo_verdict}")
    logger.info("")
    if any_mismatch:
        logger.info("  DIFF rows mean stored value differs from what P2 would extract now.")
        logger.info("  Run --recheck-all to heal gaps, or leave as-is.")
    else:
        logger.info("  All derived fields match stored values.")


def _greeting_for(full_name: str, email: str) -> str:
    """First-name greeting for an outbound email, or 'there' when we don't have one.

    'Full Name' is sometimes just the email's local-part (P1's fallback when the sender had
    no display name, e.g. 'umerta45' from umerta45@gmail.com) rather than a real name -
    fall back to 'there' instead of greeting someone by their email handle."""
    name = (full_name or "").strip()
    local_part = email.split("@")[0].strip().lower()
    if not name or "@" in name or name.lower() == local_part:
        return "there"
    return name.split(" ")[0]


def _send_non_usa_decline(client, email: str, full_name: str = "") -> bool:
    """Email the polite non-USA decline to a rejected candidate (best-effort).

    Off when geo_reject_email is false. Needs Mail.Send (Application) consent on the app;
    if absent, send_mail logs a warning and returns False - the rejection still completes.
    Returns True only when Graph accepted the send.
    """
    from hiring_agent.config import GEO_REJECT_EMAIL, EMAIL_TEMPLATES, COMPANY_NAME
    email = _bare_email(email)
    if not GEO_REJECT_EMAIL or not email:
        return False
    greeting = _greeting_for(full_name, email)
    tmpl = EMAIL_TEMPLATES.get("rejection_non_usa", {})
    subject = tmpl.get("subject", "Update on Your {company} Application").format(company=COMPANY_NAME)
    body = tmpl.get("body", "").format(company=COMPANY_NAME, greeting=greeting)
    if not body:
        return False
    if client.send_mail(email, subject, body):
        logger.info(f"       Email     : non-USA decline sent to {email}")
        return True
    return False


def _send_pending_declines(client, dry_run: bool = False,
                           only_app_ids: set[str] | None = None) -> dict:
    """Send the decline email to every Rejected-sheet row whose 'Decline Sent' is blank.

    This is THE only place decline emails are sent, and it is idempotent:
      - each row's marker is stamped immediately after ITS email goes out, so a crash,
        shutdown, or Ctrl+C mid-pass never re-emails anyone on the next run (daily 9am
        or a manual run from the app both just continue where it stopped);
      - new rejected rows arrive with a blank marker and simply join the queue —
        already-stamped rows are never touched again;
      - de-duped by email address or phone across rows: a candidate present twice (e.g. an old
        backlog mail re-ingested under a new Application ID) gets ONE decline; the
        extra row is stamped 'Duplicate' instead of emailed.
    Sends one at a time with a short pause (graceful, throttle-friendly). A failed
    send (e.g. Mail.Send consent still missing) leaves the marker blank so that row
    is retried automatically on the next run.
    """
    import time as _time
    from datetime import datetime as _dt
    from hiring_agent.config import GEO_REJECT_EMAIL
    none = {"sent": 0, "failed": 0, "skipped": 0}
    if not GEO_REJECT_EMAIL:
        return none
    try:
        rows = client.list_rejected_rows()
    except Exception as e:
        logger.warning(f"  Declines  : could not read the Rejected sheet ({e}); skipping this pass.")
        return none
    if not rows:
        return none
    if _DECLINE_COL not in rows[0]["values"]:
        # Marker column missing (ensure_rejected_columns failed, e.g. throttled) —
        # sending without it could double-email people, so don't.
        logger.warning(f"  Declines  : Rejected table has no '{_DECLINE_COL}' column yet; "
                       f"skipping sends this run to avoid any double-email risk.")
        return none

    already_emails = {_bare_email(r["values"].get("Email", "")) for r in rows
                      if not _awaiting_send(r["values"].get(_DECLINE_COL, ""))} - {""}
    already_phones = {_normalize_phone_key(r["values"].get("Phone", "")) for r in rows
                      if not _awaiting_send(r["values"].get(_DECLINE_COL, ""))} - {""}
    only_app_ids = {str(x).strip() for x in (only_app_ids or set()) if str(x).strip()}
    # A blank Application ID means this is a LAYOUT row, not a candidate: the month spacer
    # added by _ensure_rejected_period_separator, or a '-- 2025 --' year label written by
    # resort_candidate_sheets.py. Both have a blank Decline Sent cell, so without this guard
    # they were selected as "decline still owed", found to have no email address, and stamped
    # 'No valid email on file' - writing text into the separator rows on every single run.
    # Nobody was ever emailed (the send is skipped when the address is empty), but the sheet's
    # blank spacers stopped being blank. Fixed 2026-08-04.
    #
    # Filtering inside list_rejected_rows() would NOT be enough: unlike list_rows(), a year
    # label carries a Full Name, so it is not a fully-blank row. The Application ID is the
    # only reliable "this is a real candidate" test for both separator shapes.
    pending = [r for r in rows
               if _awaiting_send(r["values"].get(_DECLINE_COL, ""))
               and str(r["values"].get("Application ID", "") or "").strip()]
    if only_app_ids:
        pending = [
            r for r in pending
            if str(r["values"].get("Application ID", "")).strip() in only_app_ids
        ]
    if not pending:
        return none

    logger.info("")
    logger.info(f"  Declines  : {len(pending)} rejected candidate(s) awaiting a decline email...")
    sent = failed = skipped = 0
    for r in pending:
        vals = r["values"]
        email = _bare_email(vals.get("Email", ""))
        phone_key = _normalize_phone_key(vals.get("Phone", ""))
        app_id = str(vals.get("Application ID", "")).strip() or "(no ref)"
        stamp = ""
        if not email:
            stamp = "No valid email on file"
            skipped += 1
            logger.info(f"       {app_id}: no valid email — marked, will not retry.")
        elif ((email and email in already_emails)
              or (phone_key and phone_key in already_phones)):
            stamp = "Duplicate - decline already emailed"
            skipped += 1
            logger.info(f"       {app_id}: {email} already received a decline — marked duplicate.")
        elif dry_run:
            logger.info(f"       [DRY-RUN] Would email decline to {email} ({app_id}).")
            continue
        elif _send_non_usa_decline(client, str(vals.get("Email", "")),
                                   str(vals.get("Full Name", ""))):
            stamp = _sent_marker()
            if email:
                already_emails.add(email)
            if phone_key:
                already_phones.add(phone_key)
            sent += 1
            _time.sleep(2)   # one by one, gently
        else:
            failed += 1      # marker stays blank -> retried automatically next run
            continue
        if stamp and not dry_run:
            try:
                _store(client).save_by_id(_app_id_of(vals), {_DECLINE_COL: stamp},
                                          current_values=vals, sheet="rejected",
                                          hint=r["index"])
            except Exception as e:
                logger.warning(f"       WARNING  : could not stamp '{_DECLINE_COL}' "
                               f"for {app_id}: {e}")
    logger.info(f"  Declines  : {sent} sent | {failed} failed (auto-retry next run) | "
                f"{skipped} skipped")
    return {"sent": sent, "failed": failed, "skipped": skipped}


def _missing_fields_list(vals: dict) -> list[str]:
    """Which candidate-supplied details a row is missing, as reader-friendly phrases -
    an empty list means nothing's missing.

    Covers phone, location/country, skills, education, PORTFOLIO and email address.

    Portfolio history: excluded on 2026-07-15 at client request ("only these four"),
    then REINSTATED on 2026-08-04 at client request ("ask for missing details like email,
    portfolio, number, city, location"). Recorded here so the earlier decision is not
    mistaken for the current one and quietly "fixed" back.

    A portfolio slot reading 'N/A' is NOT treated as missing: 'N/A' is the settled "asked,
    and there genuinely is none" label. Only a truly empty slot counts, so a candidate who
    has already told us they have no portfolio is never nagged again.

    If a row was kept as scored but Location/Country are present and still too vague
    to prove US presence or non-US presence, ask for current location/country too.
    Clear non-USA rows should have been moved to Rejected before this pass, so they do
    not receive a missing-info nudge instead of the decline path.
    """
    missing = []
    if _is_gap(vals.get("Phone")):
        missing.append("your phone number")
    if _is_gap(vals.get("Location")):
        missing.append("your city and location")
    elif _needs_location_clarity(vals):
        missing.append("your current location/country")
    if _is_gap(vals.get("Current Skills")):
        missing.append("your listed skills")
    if _is_gap(vals.get("Education")):
        missing.append("your education background")
    # Portfolio: only when every slot is truly empty. An explicit 'N/A' means the candidate
    # has none, which is an answer, not a gap - so it must not re-trigger the request.
    _ports = [str(vals.get(f"Portfolio {i}", "") or "").strip() for i in (1, 2, 3)]
    if all(_is_gap(p) for p in _ports) and not any(p.upper() == "N/A" for p in _ports):
        missing.append("a portfolio or LinkedIn/GitHub link")
    # Email: we normally have it (it is the sender), so this only fires on a genuinely
    # broken/absent address - in which case the request cannot be delivered anyway and the
    # caller stamps 'No valid email on file'. Listed for completeness of the ask.
    if _is_gap(_bare_email(vals.get("Email", ""))):
        missing.append("a contact email address")
    return missing


def _country_is_us(country: str) -> bool:
    c = normalize_country(str(country or "").strip()).lower()
    return c == "united states" or c in _cfg.US_COUNTRY_TERMS


def _country_is_foreign(country: str) -> bool:
    c = str(country or "").strip().lower()
    return bool(c) and any(fc == c for fc in _cfg.FOREIGN_COUNTRIES)


def _looks_like_non_us_international_phone(phone: str) -> bool:
    raw = str(phone or "").strip()
    if not raw or "*" in raw:
        return False
    digits = "".join(ch for ch in raw if ch.isdigit())
    return raw.startswith("+") and digits and not digits.startswith("1")


def _location_clarity_issue(location: str, country: str = "", phone: str = "",
                             education: str = "", resume_text: str = "") -> bool:
    """True when geography is conflicting/ambiguous enough to ask before rejecting.

    A clean foreign current location is now ALSO a clarity issue unless Phone or
    Education corroborates it (see geo._corroborates_non_usa) — a location-text match
    alone, which extraction can pick up from anywhere in the resume (a stale hometown or
    old-degree mention), is no longer enough on its own to skip straight to a hard reject.
    Otherwise this only catches mixed evidence such as foreign-city + US state suffix
    ('Pune, NC'), a US state with a foreign country cell ('NJ' + 'India'), or a US
    location paired with an explicit non-US phone code.
    """
    from hiring_agent.geo import (
        is_usa_location, is_foreign_location, _has_us_admin_signal, _corroborates_non_usa,
    )

    loc = str(location or "").strip()
    if _is_gap(loc):
        return False
    loc_is_foreign = is_foreign_location(loc)
    loc_has_us_admin = _has_us_admin_signal(loc)
    if loc_is_foreign and loc_has_us_admin:
        return True
    if loc_is_foreign:
        return _corroborates_non_usa(phone, education, resume_text) != "corroborates"
    if is_usa_location(loc):
        # A CONCRETE US signal in the location (state / territory / ZIP) BEATS a
        # contradicting Country cell instead of parking the row. geo.classify_location_usa
        # step 1 already states this - "physical-presence evidence must beat that" - because
        # Country is routinely mis-read off a past overseas degree. Leaving the two to
        # disagree is the worst outcome: the sheet then shows 'Los Angeles, CA' + 'India',
        # which is simply wrong, and the row sits in review forever.
        # Fixed 2026-08-04 (client-reported). Live case APP-20260716-0617-CB73: Yashvi
        # Vaghela, location 'Los Angeles, CA', US phone (213), Country mis-read 'India' off
        # her earlier degree. She had ALREADY replied twice confirming she is US-based and
        # was still being asked. Once this returns False the row confirms as US, and
        # reconcile_us_country rewrites the Country cell to 'United States'.
        # A genuinely non-US international PHONE is still a real conflict and still asks.
        if _has_us_admin_signal(loc):
            return _looks_like_non_us_international_phone(phone)
        return _country_is_foreign(country) or _looks_like_non_us_international_phone(phone)
    if _country_is_us(country) or _country_is_foreign(country):
        return False
    return True


def _classify_candidate_geo(location: str, country: str = "", phone: str = "",
                            education: str = "", resume_text: str = "") -> tuple[GeoDecision, str]:
    """Tri-state geo decision with cross-field conflicts downgraded to review.

    A concrete US location is normally confirmed, but a contradictory country or
    international phone still needs clarification. Supporting US context can prevent
    a non-US rejection, never manufacture a confirmed current address.
    """
    decision, reason = classify_location_usa(
        location, country=country, resume_text=resume_text,
        phone=phone, education=education,
    )
    if (decision == GeoDecision.CONFIRMED_US
            and _location_clarity_issue(
                location, country, phone, education=education, resume_text=resume_text)):
        return (
            GeoDecision.UNKNOWN,
            f"{reason}; conflicting country/phone evidence requires current-location confirmation",
        )
    return decision, reason


def _sent_marker() -> str:
    """Timestamp stamped into Mail Sent / Decline Sent / Info Request Sent after a send.

    Under SUPPRESS_EMAILS nothing actually left the mailbox, so the marker says so rather
    than claiming 'Sent' - matching P1's 'TEST-MODE (suppressed)' stamp. Both forms stay
    parseable by _parse_marker_datetime, which reads the date after the word 'Sent'.
    """
    from datetime import datetime as _dt   # imported locally, as its two call sites do
    stamp = f"Sent {_dt.now():%Y-%m-%d %H:%M}"
    return f"TEST-MODE (suppressed) {stamp}" if _cfg.SUPPRESS_EMAILS else stamp


#: A stamp written while SUPPRESS_EMAILS was on records that nothing left the mailbox - see
#: _sent_marker. It must never read as "this candidate has been contacted".
_SUPPRESSED_MARKER_RE = re.compile(r"(?i)TEST-MODE\s*\(suppressed\)")


def _awaiting_send(marker) -> bool:
    """True when a row still owes an email: no marker at all, or one from a suppressed run."""
    return _is_gap(marker) or bool(_SUPPRESSED_MARKER_RE.search(str(marker or "")))


def _parse_marker_datetime(marker: str):
    """When an email actually went out, or None.

    A suppressed stamp carries a parseable 'Sent <ts>' but records the opposite - that the
    send was withheld - so it has no send time. All three callers want a REAL send: the
    age-out clock, the "already contacted" de-dupe set, and the did-they-reply check.
    """
    if _SUPPRESSED_MARKER_RE.search(str(marker or "")):
        return None
    m = re.search(r"\bSent\s+(\d{4}-\d{2}-\d{2})(?:[ T](\d{2}:\d{2}))?", str(marker or ""))
    if not m:
        return None
    text = m.group(1) + (" " + m.group(2) if m.group(2) else "")
    ts = pd.to_datetime(text, errors="coerce")
    if pd.isna(ts):
        return None
    return ts.to_pydatetime().replace(tzinfo=None)


def _parse_cell_datetime(value):
    if value in (None, ""):
        return None
    try:
        num = float(value)
    except (TypeError, ValueError):
        num = None
    if num is not None and 20000 <= num <= 80000:
        ts = pd.to_datetime(num, unit="D", origin="1899-12-30", errors="coerce")
    else:
        ts = pd.to_datetime(value, errors="coerce")
    if pd.isna(ts):
        return None
    return ts.to_pydatetime().replace(tzinfo=None)


def _updated_after_location_request(vals: dict) -> bool:
    sent_at = _parse_marker_datetime(vals.get(_INFO_REQUEST_COL, ""))
    if sent_at is None:
        return False
    updated_at = (_parse_cell_datetime(vals.get("Last Updated Date"))
                  or _parse_cell_datetime(vals.get("Received Date")))
    return bool(updated_at and updated_at > sent_at)


def _needs_location_clarity(vals: dict) -> bool:
    """True when a kept/scored row needs a location clarification mail.

    This intentionally avoids online lookups and mail sends; it is only a cheap,
    deterministic classifier for the info-request pass. Clear foreign signals belong
    on Rejected; conflicting geography stays on Main for one clarification cycle.
    """
    decision, _ = _classify_candidate_geo(
        vals.get("Location", ""), vals.get("Country", ""), vals.get("Phone", ""),
        education=vals.get("Education", ""))
    return decision == GeoDecision.UNKNOWN


def _join_english(items: list[str]) -> str:
    """['a', 'b', 'c'] -> 'a, b, and c'; ['a', 'b'] -> 'a and b'; ['a'] -> 'a'."""
    if len(items) <= 1:
        return items[0] if items else ""
    if len(items) == 2:
        return f"{items[0]} and {items[1]}"
    return ", ".join(items[:-1]) + f", and {items[-1]}"


def _send_missing_info_request(client, email: str, full_name: str, app_id: str,
                               missing: list[str]) -> bool:
    """Email a SCORED candidate asking for the specific fields we don't have on file.

    Needs the same Mail.Send (Application) consent as the decline notice; if absent,
    send_mail logs a warning and returns False - the row is simply retried next run."""
    from hiring_agent.config import EMAIL_TEMPLATES, COMPANY_NAME
    email = _bare_email(email)
    if not email or not missing:
        return False
    tmpl = EMAIL_TEMPLATES.get("missing_info", {})
    subject = tmpl.get("subject", "Quick follow-up on your {company} application (Ref: {ref})").format(
        company=COMPANY_NAME, ref=app_id)
    body = tmpl.get("body", "").format(
        company=COMPANY_NAME, greeting=_greeting_for(full_name, email),
        ref=app_id, missing_list=_join_english(missing))
    if not body:
        return False
    if client.send_mail(email, subject, body):
        logger.info(f"       Email     : missing-info request sent to {email} "
                    f"({_join_english(missing)})")
        return True
    return False


def _send_pending_info_requests(client, dry_run: bool = False) -> dict:
    """Nudge scored or location-review rows missing required info, once each.

    Idempotent the same way _send_pending_declines is: each row's 'Info Request Sent'
    marker is stamped right after its email goes out (or as soon as we know none is
    needed), so a crash/shutdown mid-pass never double-emails on the next run.

    Historical rows stamped 'Nothing missing' can be re-opened only for the newer
    location-clarity rule. A real 'Sent ...' marker stays final and is never resent.
    """
    import time as _time
    from datetime import datetime as _dt
    none = {"sent": 0, "failed": 0, "skipped": 0}
    try:
        rows = client.list_rows()
    except Exception as e:
        logger.warning(f"  Info reqs : could not read the candidate table ({e}); skipping this pass.")
        return none
    pending = []
    for r in rows:
        vals = r["values"]
        if str(vals.get("Status", "")).strip() not in {
            STATUS_SCORED, STATUS_LOCATION_REVIEW,
        }:
            continue
        # Layout rows (month spacer / '-- 2025 --' year label) carry no Application ID.
        # The Status check above already excludes them, but the same explicit guard the
        # decline pass uses is kept here so a future separator that somehow carries a
        # Status can never be treated as a candidate. See _send_pending_declines.
        if not str(vals.get("Application ID", "") or "").strip():
            continue
        marker = str(vals.get(_INFO_REQUEST_COL, "") or "").strip()
        if _awaiting_send(marker):
            pending.append(r)
            continue
        if (marker.lower() == "nothing missing"
                and "your current location/country" in _missing_fields_list(vals)):
            pending.append(r)
    if not pending:
        return none

    # De-dupe by email AND phone, exactly as _send_pending_declines does. P1's duplicate
    # gate is a read-after-write against Excel Online and can miss a same-sender resend,
    # and _merge_candidate_ready deliberately excludes location-review rows from duplicate
    # merging (an ambiguous row may be mid-clarification), so two rows for one person can
    # legitimately coexist here. Without this, both were emailed - the same candidate got
    # two "where are you based?" mails on the same run. Fixed 2026-08-04.
    # "Already contacted" must mean an email ACTUALLY went out, i.e. a parseable
    # 'Sent <timestamp>' stamp - not merely any non-blank marker. The bookkeeping stamps
    # ('Nothing missing', 'No valid email on file', 'Duplicate - ...') record that we looked,
    # not that we wrote to anyone. This matters because, unlike the decline pass, a row can be
    # REOPENED while still carrying a marker: a 'Nothing missing' row that later develops a
    # location-clarity gap is put back into `pending` a few lines above. Keying off "non-blank
    # marker" would put that row's own address into this set and make it match ITSELF, so the
    # one email it was just reopened for would be suppressed as a duplicate of itself.
    _emailed = [r for r in rows
                if _parse_marker_datetime(r["values"].get(_INFO_REQUEST_COL, "")) is not None]
    already_emails = {_bare_email(r["values"].get("Email", "")) for r in _emailed} - {""}
    already_phones = {_normalize_phone_key(r["values"].get("Phone", "")) for r in _emailed} - {""}

    logger.info("")
    logger.info(f"  Info reqs : {len(pending)} scored candidate(s) to check for missing info...")
    sent = failed = skipped = 0
    for r in pending:
        vals = r["values"]
        app_id = str(vals.get("Application ID", "")).strip() or "(no ref)"
        email = _bare_email(vals.get("Email", ""))
        phone_key = _normalize_phone_key(vals.get("Phone", ""))
        missing = _missing_fields_list(vals)
        stamp = ""
        if not missing:
            stamp = "Nothing missing"
            skipped += 1
        elif not email:
            stamp = "No valid email on file"
            skipped += 1
            logger.info(f"       {app_id}: no valid email — marked, will not retry.")
        elif ((email and email in already_emails)
              or (phone_key and phone_key in already_phones)):
            stamp = "Duplicate - info request already sent"
            skipped += 1
            logger.info(f"       {app_id}: {email} already received an info request — "
                        f"marked duplicate.")
        elif dry_run:
            logger.info(f"       [DRY-RUN] Would ask {email} for {_join_english(missing)} ({app_id}).")
            continue
        elif _send_missing_info_request(client, str(vals.get("Email", "")),
                                        str(vals.get("Full Name", "")), app_id, missing):
            stamp = _sent_marker()
            if email:
                already_emails.add(email)
            if phone_key:
                already_phones.add(phone_key)
            sent += 1
            _time.sleep(2)   # one by one, gently
        else:
            failed += 1      # marker stays blank -> retried automatically next run
            continue
        if stamp and not dry_run:
            try:
                _store(client).save_by_id(_app_id_of(vals), {_INFO_REQUEST_COL: stamp},
                                          current_values=vals, hint=r["index"])
            except Exception as e:
                logger.warning(f"       WARNING  : could not stamp '{_INFO_REQUEST_COL}' "
                               f"for {app_id}: {e}")
    logger.info(f"  Info reqs : {sent} sent | {failed} failed (auto-retry next run) | "
                f"{skipped} skipped")
    return {"sent": sent, "failed": failed, "skipped": skipped}


#: Stamped into 'Decline Sent' on an aged-out row so the decline pass never emails it.
#: The 'N/A - ' prefix is the established "deliberately not emailed" convention already
#: recognised by audit_sharepoint_results_health.py, and it is NOT a _GAP_LITERALS value
#: (only a bare 'n/a' is), so _send_pending_declines correctly reads the row as settled.
_AGED_OUT_DECLINE_STAMP = "N/A - aged out, no reply"


def _review_clock_start(vals: dict):
    """When the countdown for an unanswered review row began, or None if undeterminable.

    For a location-review row that is the moment we ASKED (the 'Info Request Sent' stamp) -
    aging someone out before they were ever contacted would be indefensible, so a row with
    no info-request marker yet is not aged at all and simply waits for the next mail pass.
    Spam rows are never emailed, so they age from their last update (or receipt) instead.
    """
    status = str(vals.get("Status", "") or "").strip()
    if status == _cfg.STATUS_NEEDS_REVIEW_SPAM:
        return (_parse_cell_datetime(vals.get("Last Updated Date"))
                or _parse_cell_datetime(vals.get("Received Date")))
    asked_at = _parse_marker_datetime(vals.get(_INFO_REQUEST_COL, ""))
    if asked_at is None:
        return None
    return asked_at


def _age_out_stale_reviews(client, dry_run: bool = False) -> dict:
    """Quietly close review rows nobody answered, so neither bucket grows without bound.

    Added 2026-08-04 on client instruction (14 days, silent close, both statuses). Before
    this, 'Needs Review - Location Confirmation' and 'Needs Review - Possible Spam' had no
    exit at all: one clarification email went out, and a row that was never answered sat on
    the main sheet forever.

    Deliberately SILENT - no decline email is sent, and 'Decline Sent' is pre-stamped so the
    decline pass skips the row permanently. Silence from a candidate is not evidence they
    are non-US; it is only evidence we cannot confirm, so closing them quietly is honest
    where emailing them a rejection would not be.

    Reversible by design: location rows close as 'Rejected - Location Not Confirmed', which
    is precisely the status `recover_location_rejections()` re-audits, so a wrongly-closed
    candidate can be pulled back with `bot.py --recover-locations`. A row is only aged once
    the clarification email has actually gone out (see _review_clock_start).
    """
    from sharepoint_client import SharePointError
    none = {"closed": 0, "skipped": 0, "errors": 0}
    if _cfg.REVIEW_AGE_OUT_DAYS <= 0:
        return none
    try:
        rows = client.list_rows()
    except Exception as e:
        logger.warning(f"  Age-out   : could not read the candidate table ({e}); "
                       f"skipping this pass.")
        return none

    cutoff = dt.datetime.now() - dt.timedelta(days=_cfg.REVIEW_AGE_OUT_DAYS)
    aging = {STATUS_LOCATION_REVIEW, _cfg.STATUS_NEEDS_REVIEW_SPAM}
    stale, waiting = [], 0
    for r in rows:
        vals = r["values"]
        if str(vals.get("Status", "") or "").strip() not in aging:
            continue
        if not str(vals.get("Application ID", "") or "").strip():
            continue          # month spacer / year label, never a candidate
        started = _review_clock_start(vals)
        if started is None:
            waiting += 1      # not yet asked - the mail pass owes them a note first
            continue
        if started > cutoff:
            waiting += 1
            continue
        stale.append(r)

    if not stale:
        return {"closed": 0, "skipped": waiting, "errors": 0}

    logger.info("")
    logger.info(f"  Age-out   : {len(stale)} review row(s) unanswered for "
                f"{_cfg.REVIEW_AGE_OUT_DAYS}+ days; closing quietly (no email). "
                f"{waiting} still within the window.")

    # Snapshot the Rejected sheet ONCE so _finish_rejection's repeat-applicant de-dupe
    # works, exactly as the scoring loop does.
    try:
        rej_existing = client.list_rejected_rows()
    except SharePointError:
        rej_existing = []
    rej_ids = {str(r["values"].get("Application ID", "")).strip()
               for r in rej_existing} - {""}
    rej_emails = {_bare_email(r["values"].get("Email", "")) for r in rej_existing} - {""}
    rej_phones = {_normalize_phone_key(r["values"].get("Phone", "")) for r in rej_existing} - {""}

    closed = errors = 0
    deleted_idx: list[int] = []
    # Descending index order: every delete shifts later rows up, and _live_row_index only
    # corrects for rows deleted from EARLIER positions. Closing bottom-up means no index
    # this pass still holds is ever invalidated by a delete that already happened.
    for r in sorted(stale, key=lambda x: x["index"], reverse=True):
        vals = r["values"]
        app_id = str(vals.get("Application ID", "")).strip()
        was = str(vals.get("Status", "") or "").strip()
        started = _review_clock_start(vals)
        days = (dt.datetime.now() - started).days if started else "?"
        index = _live_row_index(client, r, deleted_idx, app_id)
        if index is None:
            errors += 1
            continue

        # Spam rows keep their own status (the reason they were pulled must survive on the
        # Rejected sheet); location rows take the retired-but-exact 'Location Not Confirmed'.
        closed_status = (_cfg.STATUS_NEEDS_REVIEW_SPAM
                         if was == _cfg.STATUS_NEEDS_REVIEW_SPAM
                         else _cfg.STATUS_LOCATION_UNCONFIRMED)
        fields = {"Status": closed_status, _DECLINE_COL: _AGED_OUT_DECLINE_STAMP}

        subpath = _cfg.dated_subpath(_parse_received(vals.get("Received Date")))
        stored = _stored_resume_names(
            app_id, vals.get("Original Filename", ""), full_name=vals.get("Full Name"),
            category=vals.get("Category"), resume_url=vals.get("Resume URL"))

        logger.info(f"       {app_id}: '{was}' unanswered {days} day(s) -> "
                    f"'{closed_status}' (silent, no email)")
        try:
            if _finish_rejection(client, index, app_id, vals, fields, stored, subpath,
                                 dry_run, rej_ids, rej_emails, rej_phones,
                                 "aged out - no reply", rejected_rows=rej_existing):
                deleted_idx.append(r["index"])
                closed += 1
        except Exception as e:
            errors += 1
            logger.warning(f"       WARNING  : could not close {app_id}: {e}")

    logger.info(f"  Age-out   : {closed} closed | {waiting} still waiting | {errors} error(s)")
    return {"closed": closed, "skipped": waiting, "errors": errors}


def _cleanup_temp_files():
    """Delete locally saved resume files after scoring is done for this candidate."""
    for f in _temp_files:
        try:
            if f.exists():
                f.unlink()
        except OSError:
            pass
    _temp_files.clear()


def _cleanup_empty_dirs():
    """Remove empty dated directories left behind after file cleanup."""
    import os as _os
    for dirpath, dirnames, filenames in _os.walk(str(_cfg.INPUT_DIR), topdown=False):
        if not filenames and not dirnames:
            try:
                _os.rmdir(dirpath)
            except OSError:
                pass
