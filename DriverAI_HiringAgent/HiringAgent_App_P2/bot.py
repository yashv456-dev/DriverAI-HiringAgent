"""DriverAI Hiring Agent - scoring CLI for local tools and SharePoint runs.

One command-line tool that can score local resumes, re-score saved workbooks,
score the Phase 1 SharePoint queue, validate stored rows, and run maintenance
checks. Ollama is optional; the keyword scorer is always available.

Outlook is never used here. Phase 1 owns applicant email; this CLI reads resumes
and reads/writes SharePoint or local workbooks depending on the command.

Examples
--------
  python bot.py --score-folder ./resumes        # score a folder -> Excel
  python bot.py --score-excel ./candidates.xlsx # re-score the rows of a workbook
  python bot.py --score-file ./jane_cv.pdf      # print one resume's breakdown
  python bot.py --rematch                        # re-score every saved workbook
  python bot.py --list-roles                     # show the roles being matched
  python bot.py --test-sharepoint                # verify online credentials
  python bot.py --score-sharepoint               # score the live SharePoint queue
  python bot.py --score-folder ./resumes --no-ai # force the keyword scorer
"""

import argparse
import datetime
import os
import sys
import time
from pathlib import Path

# Make the package importable when run as a script from anywhere.
APP_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(APP_DIR))

_LOCK = APP_DIR / ".agent.lock"

_ONLINE_ENV_KEYS = (
    "TENANT_ID",
    "CLIENT_ID",
    "CLIENT_SECRET",
    "SHAREPOINT_HOSTNAME",
    "SHAREPOINT_SITE_PATH",
    "SHAREPOINT_TABLE",
    "SHAREPOINT_RESUMES_FOLDER",
    "SHAREPOINT_WORKBOOK",
    "SENDER_MAILBOX",
    "HIRING_GEO_REJECT_EMAIL",
    "HIRING_ERROR_EMAIL",
    "HIRING_P2_DISABLED",
)


def _positive_int(value: str) -> int:
    """Argparse type for counts that must be at least one."""
    try:
        parsed = int(value)
    except (TypeError, ValueError) as e:
        raise argparse.ArgumentTypeError("must be a whole number") from e
    if parsed < 1:
        raise argparse.ArgumentTypeError("must be at least 1")
    return parsed


def _local_env_issue() -> str:
    """Return a secret-safe description of a malformed local P2 .env file."""
    env_path = APP_DIR / ".env"
    if not env_path.exists():
        if (APP_DIR / ".env.txt").exists():
            return ".env is missing; .env.txt is only a text/template file"
        return ""
    try:
        raw = env_path.read_text(encoding="utf-8-sig")
    except OSError as e:
        return f".env could not be read ({e})"

    lines = [line.strip() for line in raw.splitlines()
             if line.strip() and not line.lstrip().startswith("#")]
    for line in lines:
        markers = [f"{key}=" for key in _ONLINE_ENV_KEYS if f"{key}=" in line]
        if len(markers) > 1:
            return ".env is malformed: put each KEY=value setting on its own line"

    parsed = {}
    for line in lines:
        if "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        if key in _ONLINE_ENV_KEYS:
            parsed[key] = value.strip()

    contaminated = [
        key for key, value in parsed.items()
        if key != "CLIENT_SECRET"
        and value.lower().endswith(("echo", "echo.echo"))
    ]
    if contaminated:
        return ".env contains pasted 'echo' command text; use plain KEY=value lines"
    return ""


def _acquire_lock() -> bool:
    """Best-effort single-instance lock for --watch (prevents overlapping polls)."""
    def _pid_alive(pid_text: str) -> bool:
        try:
            pid = int((pid_text or "").strip())
            if pid <= 0:
                return False
            os.kill(pid, 0)
            return True
        except (OSError, ValueError):
            return False

    try:
        fd = os.open(str(_LOCK), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        os.write(fd, str(os.getpid()).encode())
        os.close(fd)
        return True
    except FileExistsError:
        if not _pid_alive(_LOCK.read_text(encoding="ascii", errors="ignore")):
            _release_lock()
            try:
                fd = os.open(str(_LOCK), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                os.write(fd, str(os.getpid()).encode())
                os.close(fd)
                return True
            except FileExistsError:
                return False
        return False
    except OSError:
        return True  # if the lock can't be created, don't block the run


def _release_lock() -> None:
    try:
        _LOCK.unlink(missing_ok=True)
    except OSError:
        pass


def _apply_env(args) -> None:
    """Translate CLI toggles into HIRING_* env vars BEFORE the package is imported
    (config.py reads these at import time)."""
    if args.no_ai:
        os.environ["HIRING_OLLAMA_ENABLED"] = "false"
        os.environ["HIRING_OLLAMA_SCORING"] = "false"
    elif args.ai:
        os.environ["HIRING_OLLAMA_ENABLED"] = "true"
        os.environ["HIRING_OLLAMA_SCORING"] = "true"
    if args.model:
        os.environ["HIRING_OLLAMA_MODEL"] = args.model
    if args.host:
        os.environ["HIRING_OLLAMA_HOST"] = args.host
    if args.no_geo:
        os.environ["HIRING_GEO_USA_ONLY"] = "false"
    if args.batch_size is not None:
        os.environ["HIRING_SCORING_BATCH_LIMIT"] = str(args.batch_size)
    if args.excel:
        os.environ["HIRING_EXCEL_FILE"] = str(Path(args.excel).resolve())


def _run_doctor() -> None:
    """Read-only setup health check. Tells the user exactly what is ready and what's
    missing (OCR engine, local AI, credentials) without changing anything."""
    import importlib

    def _have(mod: str) -> bool:
        try:
            importlib.import_module(mod)
            return True
        except Exception:
            return False

    def _line(label: str, value: str) -> None:
        print(f"  {label:<18}: {value}")

    print("\nDriverAI Hiring Agent - setup check")
    print("-" * 56)
    _line("Python", sys.version.split()[0])

    core = [m for m in ("pandas", "openpyxl", "pypdf", "docx", "requests", "yaml") if not _have(m)]
    _line("Core deps", "OK" if not core else "MISSING " + ", ".join(core))
    gui = [m for m in ("customtkinter", "dotenv") if not _have(m)]
    _line("Desktop deps", "OK" if not gui else "MISSING " + ", ".join(gui))

    ocr_missing = [m for m in ("fitz", "pytesseract", "PIL") if not _have(m)]
    _line("OCR libs", "OK (pymupdf/pytesseract/pillow)" if not ocr_missing
          else "MISSING " + ", ".join(ocr_missing))
    from hiring_agent.extraction import find_tesseract
    tess = find_tesseract()
    _line("Tesseract engine", tess if tess else "NOT FOUND - scanned/image PDFs will be skipped")

    from hiring_agent.ollama_scorer import ollama_health
    ok_ai, ai_msg = ollama_health()
    _line("Ollama (local AI)", ai_msg)

    from hiring_agent.config import JD_CACHE_FILE, SHAREPOINT_CONFIGURED, DEFAULT_ROLES
    env_issue = _local_env_issue()
    _line("P2 .env", env_issue or "OK")
    if JD_CACHE_FILE.exists():
        try:
            import json as _json
            d = _json.loads(JD_CACHE_FILE.read_text(encoding="utf-8"))
            _line("JD cache", f"{len(d.get('roles', []))} role(s), built {d.get('built_at', '?')}")
        except Exception:
            _line("JD cache", "present but unreadable (run --refresh-jd)")
    else:
        _line("JD cache", f"not built - using {len(DEFAULT_ROLES)} built-in roles (run --refresh-jd)")
    _line("SharePoint", "configured (online mode)" if SHAREPOINT_CONFIGURED
          else "not configured (local mode only)")

    print("-" * 56)
    ocr_ready = (not ocr_missing) and bool(tess)
    if not ocr_ready:
        print("  ! OCR not fully set up - scanned/image-only PDFs won't be read.")
        print("    Fix: run Launch.bat, or `winget install -e --id UB-Mannheim.TesseractOCR`")
        print("         and `pip install pymupdf pytesseract pillow`.")
    else:
        print("  OCR ready - scanned/image PDFs will be read.")
    if env_issue:
        print(f"  ! {env_issue}.")
    print("")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="DriverAI Hiring Agent - local/SharePoint resume scoring CLI.")
    src = parser.add_mutually_exclusive_group(required=True)
    src.add_argument("--score-folder", metavar="DIR",
                     help="Score every PDF/DOCX in a local folder and upsert into the workbook.")
    src.add_argument("--score-excel", metavar="FILE",
                     help="Re-score the rows of an existing .xlsx (Suggested Roles + Category).")
    src.add_argument("--score-file", metavar="FILE",
                     help="Print the role breakdown for a single PDF/DOCX (nothing written).")
    src.add_argument("--rematch", action="store_true",
                     help="Re-score Suggested Roles for every saved workbook under P2_Output/.")
    src.add_argument("--score-sharepoint", action="store_true",
                     help="ONLINE: read the Phase 1 SharePoint queue (rows left 'New Email "
                          "Received'), score each resume, and write back to the same row "
                          "(needs .env app-only credentials).")
    src.add_argument("--test-sharepoint", action="store_true",
                     help="Verify the SharePoint credentials: authenticate, resolve the site, "
                          "and read the table columns. Prints a clear pass/fail (no writes).")
    src.add_argument("--list-roles", action="store_true",
                     help="Preview the roles candidates are matched against.")
    src.add_argument("--refresh-jd", action="store_true",
                     help="Re-scan every configured JD source ONCE and cache the parsed "
                          "roles to jd_roles_cache.json. Scoring runs read that cache "
                          "instead of re-downloading all JDs each run. Run this after the "
                          "JD folder/URLs change (the app's Refresh button does the same).")
    src.add_argument("--doctor", action="store_true",
                     help="Print a setup health check: Python, core deps, OCR libs + "
                          "Tesseract engine, Ollama, config, JD cache, SharePoint. Read-only.")
    src.add_argument("--validate-row", metavar="APP_ID",
                     help="ONLINE READ-ONLY: print all 22 columns for one candidate by "
                          "Application ID, then re-derive fields from the resume and show "
                          "what P2 would extract now vs what is stored. No writes. "
                          "Example: --validate-row APP-20260630-143052-a3f9c1")
    src.add_argument("--recheck-all", action="store_true",
                     help="ONLINE MANUAL TOOL: sweep the WHOLE workbook - heal blank fields, "
                          "re-validate geo, de-duplicate. Use sparingly: running after new JD "
                          "rules are added WILL re-derive Category on old rows. Normal scoring "
                          "does a per-row self-check automatically - this command is for "
                          "emergency recovery only. Needs .env.")
    src.add_argument("--recheck-row", metavar="APP_ID", action="append",
                     help="ONLINE MANUAL TOOL: re-check only the given Application ID. "
                          "Repeat the flag for multiple rows. Uses the same repair rules as "
                          "--recheck-all, but only touches/sends decline markers for these IDs.")
    src.add_argument("--recover-location-rejections", action="store_true",
                     help="ONLINE RECOVERY: re-audit historical geo rejections with the "
                          "tri-state policy. Uses deterministic phone/education extraction, "
                          "does not rescore roles, and never sends decline email.")

    parser.add_argument("--watch", action="store_true",
                        help="With --score-sharepoint: poll the queue every --interval seconds.")
    parser.add_argument("--interval", type=int, default=300,
                        help="Seconds between polls in --watch mode (default: 300).")
    parser.add_argument("--batch-size", type=_positive_int, metavar="N",
                        help="With --score-sharepoint: process up to N queued candidates "
                             "in this run (default from config: 25).")
    parser.add_argument("--scorecards", action="store_true",
                        help="With --score-sharepoint: also upload a per-candidate scorecard .txt.")
    parser.add_argument("--excel", metavar="FILE",
                        help="Target/output workbook (pins all reads & writes to this file).")
    parser.add_argument("--dry-run", action="store_true",
                        help="Parse and log only - write nothing.")
    parser.add_argument("--ai", action="store_true", help="Force the Ollama brain on.")
    parser.add_argument("--no-ai", action="store_true",
                        help="Force the deterministic keyword scorer (skip Ollama).")
    parser.add_argument("--model", help="Ollama model name (default: llama3.2).")
    parser.add_argument("--host", help="Ollama host URL (default: http://localhost:11434).")
    parser.add_argument("--no-geo", action="store_true",
                        help="Disable the USA-only geo filter (score everyone).")
    # Older launch notes used ``--test-sharepoint run``. Keep accepting that harmless
    # trailing word so the diagnostic works on machines deployed from those notes.
    parser.add_argument("legacy_action", nargs="?", choices=("run",),
                        help=argparse.SUPPRESS)
    args = parser.parse_args()

    _apply_env(args)

    if args.doctor:
        _run_doctor()
        return

    # Imported AFTER env is set so config picks up the toggles.
    from hiring_agent.config import (
        logger, OLLAMA_ENABLED, OLLAMA_SCORING, OLLAMA_MODEL, GEO_FILTER_USA_ONLY,
    )
    from hiring_agent.intake import run_intake, score_existing_excel
    from hiring_agent.scoring import rematch_sheet
    from hiring_agent.jd_sources import get_active_roles, load_jd_sources

    brain = (f"Ollama '{OLLAMA_MODEL}'" if (OLLAMA_ENABLED and OLLAMA_SCORING)
             else "Keyword Scorer (no AI)")
    geo = "USA-only (non-USA rejected)" if GEO_FILTER_USA_ONLY else "OFF (all locations accepted)"
    logger.info("")
    logger.info("=" * 64)
    logger.info("  DriverAI Hiring Agent")
    logger.info("=" * 64)
    logger.info(f"  Brain    : {brain}")
    logger.info(f"  Geo      : {geo}")
    logger.info(f"  Started  : {datetime.datetime.now():%Y-%m-%d %H:%M:%S}")
    logger.info("-" * 64)

    if args.refresh_jd:
        from hiring_agent.jd_sources import build_jd_cache
        logger.info("Refreshing JD cache (scanning all configured JD sources once)...")
        roles = build_jd_cache()
        logger.info(f"Done. {len(roles)} role(s) cached; scoring runs now read the cache "
                    f"instead of re-downloading JDs.")
        return

    if args.list_roles:
        cfg = load_jd_sources()
        logger.info(f"JD matching: {'ON' if cfg['enabled'] else 'OFF'} | "
                    f"{len(cfg['urls'])} URL(s), {len(cfg.get('descriptions', []))} pasted JD(s)")
        for r in get_active_roles():
            logger.info(f"  {r['title']}  <-  {', '.join(r.get('skills', [])[:12])}")
        return

    if args.test_sharepoint:
        from hiring_agent.config import SHAREPOINT_CONFIGURED
        env_issue = _local_env_issue()
        if env_issue:
            logger.error(env_issue)
            raise SystemExit(2)
        if not SHAREPOINT_CONFIGURED:
            logger.error("Not configured. Set TENANT_ID, CLIENT_ID, CLIENT_SECRET, "
                         "SHAREPOINT_HOSTNAME, SHAREPOINT_SITE_PATH, SHAREPOINT_TABLE "
                         "(in .env locally, or app settings in the cloud). See .env.example.")
            raise SystemExit(2)
        try:
            from sharepoint_client import SharePointClient, SharePointError
            client = SharePointClient()
            logger.info("1/5 credentials present; requesting an app-only token...")
            client._get_token()
            logger.info("    token OK.")
            logger.info(f"2/5 resolving site {client.hostname}{client.site_path} ...")
            logger.info(f"    site id: {client.site_id()[:40]}...")
            logger.info(f"3/5 reading table '{client.table}' columns ...")
            cols = client.table_columns()
            logger.info(f"    table has {len(cols)} columns: {', '.join(cols)}")
            logger.info(f"4/5 checking resume folder '{client.resumes_folder}' ...")
            folder_items = client.list_folder_children(client.resumes_folder)
            logger.info(f"    folder OK ({len(folder_items)} immediate item(s)).")
            logger.info(f"5/5 checking Graph Mail.Send for '{client.sender_mailbox}' ...")
            if "Mail.Send" not in client.application_roles():
                raise SharePointError(
                    "Mail.Send application permission is missing or not admin-consented. "
                    "SharePoint scoring can run, but P2 applicant/admin emails will fail."
                )
            logger.info("    Mail.Send OK (read-only check; no message sent).")
            logger.info("PASS - SharePoint files, workbook, table, and P2 mail permission are ready.")
        except Exception as e:
            logger.error(f"FAILED: {e}")
            logger.error("Most common cause: the app's Graph application permissions "
                         "(Sites.ReadWrite.All / Files.ReadWrite.All / Mail.Send) were never "
                         "ADMIN-CONSENTED, or a configured SharePoint path is wrong.")
            raise SystemExit(1)
        return

    if args.score_sharepoint:
        from hiring_agent.config import SHAREPOINT_CONFIGURED
        from hiring_agent.sharepoint_scoring import score_from_sharepoint
        if not SHAREPOINT_CONFIGURED:
            logger.error("SharePoint credentials missing!")
            logger.error("  Create a .env file with: TENANT_ID, CLIENT_ID, CLIENT_SECRET,")
            logger.error("  SHAREPOINT_HOSTNAME, SHAREPOINT_SITE_PATH, SHAREPOINT_TABLE")
            logger.error("  See .env.example for the template.")
            return
        if args.watch:
            if not _acquire_lock():
                logger.warning("Another run holds the lock; exiting.")
                return
            try:
                logger.info(f"[WATCH] Polling SharePoint every {args.interval}s. Ctrl+C to stop.")
                while True:
                    try:
                        score_from_sharepoint(dry_run=args.dry_run, scorecards=args.scorecards)
                    except Exception as e:
                        logger.error(f"[WATCH] Poll error: {e}; retrying after {args.interval}s.")
                    time.sleep(args.interval)
            except KeyboardInterrupt:
                logger.info("[WATCH] Stopped by user.")
            finally:
                _release_lock()
        else:
            if not _acquire_lock():
                logger.warning("Another run holds the lock; exiting.")
                return
            try:
                score_from_sharepoint(dry_run=args.dry_run, scorecards=args.scorecards)
            finally:
                _release_lock()
    elif args.validate_row:
        from hiring_agent.config import SHAREPOINT_CONFIGURED
        from hiring_agent.sharepoint_scoring import validate_row
        if not SHAREPOINT_CONFIGURED:
            logger.error("SharePoint credentials missing! See .env.example.")
            return
        validate_row(args.validate_row)
    elif args.recheck_all:
        from hiring_agent.config import SHAREPOINT_CONFIGURED
        from hiring_agent.sharepoint_scoring import recheck_all_rows
        if not SHAREPOINT_CONFIGURED:
            logger.error("SharePoint credentials missing! See .env.example.")
            return
        recheck_all_rows(dry_run=args.dry_run)
    elif args.recheck_row:
        from hiring_agent.config import SHAREPOINT_CONFIGURED
        from hiring_agent.sharepoint_scoring import recheck_selected_rows
        if not SHAREPOINT_CONFIGURED:
            logger.error("SharePoint credentials missing! See .env.example.")
            return
        recheck_selected_rows(args.recheck_row, dry_run=args.dry_run)
    elif args.recover_location_rejections:
        from hiring_agent.config import SHAREPOINT_CONFIGURED
        from hiring_agent.sharepoint_scoring import recover_location_rejections
        if not SHAREPOINT_CONFIGURED:
            logger.error("SharePoint credentials missing! See .env.example.")
            return
        recover_location_rejections(dry_run=args.dry_run)
    elif args.score_folder:
        run_intake(path=args.score_folder, dry_run=args.dry_run)
    elif args.score_excel:
        score_existing_excel(args.score_excel, dry_run=args.dry_run)
    elif args.score_file:
        from hiring_agent.local_scorer import score_resume_file
        result = score_resume_file(args.score_file)
        if result.get("error"):
            logger.error(result["error"])
            return
        cand = result.get("candidate", {})
        logger.info(f"File: {result['file']}")
        logger.info(f"  Name:     {cand.get('full_name', '?')}")
        logger.info(f"  Location: {cand.get('location', '?')}")
        logger.info(f"  Country:  {cand.get('country', '?')}")
        logger.info(f"  Skills:   {cand.get('skills', '?')}")
        for m in result.get("role_matches", [])[:5]:
            logger.info(f"  {m['score']:3d}%  {m['role_title']}")
            if m.get("matched"):
                logger.info(f"        matched: {', '.join(m['matched'][:12])}")
            if m.get("missing"):
                logger.info(f"        missing: {', '.join(m['missing'][:12])}")
    elif args.rematch:
        rematch_sheet(roles=get_active_roles())

    logger.info("-" * 64)
    logger.info("  Finished.")
    logger.info("=" * 64)
    logger.info("")


if __name__ == "__main__":
    main()
