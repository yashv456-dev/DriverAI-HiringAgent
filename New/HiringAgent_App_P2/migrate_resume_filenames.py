"""One-time migration (2026-07-15), client-requested:
  1. Rename every already-saved resume from the old FirstLast_AppID format to the new
     FirstNameLastName_Category_<tail> format (e.g. 'JaneDoe_APP-20260710-2200-A5F2.pdf'
     -> 'JaneDoe_DataAnalytics_A5F2.pdf').
  2. Consolidate every Rejected candidate's resume out of its old per-month
     '<Year>/<Month>/Rejected/' folder into one '<Year>/Rejected/' folder.

Pure Graph API file moves/renames using data already on each row - no Ollama calls, no
field re-derivation (unlike `bot.py --recheck-all`, which is far too slow for this: it
re-scores every row, including the ~190 still-unscored ones, taking well over an hour).

Usage:
    python migrate_resume_filenames.py            # dry-run: preview only, no changes
    python migrate_resume_filenames.py --live      # actually move/rename + update Resume URL/Path
"""
import argparse
import os

from hiring_agent.config import logger, dated_subpath, STATUS_SCORED
from hiring_agent.excel_output import _parse_received
from hiring_agent.sharepoint_scoring import (
    _stored_resume_names, _get_cleaned_filename_prefix, _clean_category_for_filename,
    _resume_filename_tail, _resume_url_and_path, _rejected_subpath,
)


def _consolidate_rejected_folder(client, stored: list, month_subpath: str, dry_run: bool) -> bool:
    """Move any of `stored`'s candidate names out of the OLD '<Year>/<Month>/Rejected/'
    folder into the NEW '<Year>/Rejected/' folder. Idempotent - a name not found at the
    old location (already moved, or never there under that name) is just skipped.

    dry_run still confirms existence via a real (read-only) download attempt, so the
    preview accurately reports what would move rather than guessing."""
    from sharepoint_client import SharePointError
    old_folder = f"{client.resumes_folder}/{month_subpath}/Rejected"
    new_folder = f"{client.resumes_folder}/{_rejected_subpath(month_subpath)}"
    moved_any = False
    for name in stored:
        if dry_run:
            try:
                client.download_file(old_folder, name)
            except SharePointError:
                continue  # not at the old location under this candidate name
            logger.info(f"  [DRY-RUN]   would move '{name}': {month_subpath}/Rejected -> "
                        f"{_rejected_subpath(month_subpath)}")
            moved_any = True
            continue
        try:
            if client.move_resume(name, old_folder, new_folder):
                logger.info(f"    moved '{name}': {month_subpath}/Rejected -> {_rejected_subpath(month_subpath)}")
                moved_any = True
        except Exception as e:
            logger.warning(f"    WARNING  could not move '{name}': {e}")
    return moved_any


def _migrate_sheet(client, rows, label, rejected, dry_run):
    renamed = skipped = errors = moved = 0
    for r in rows:
        v = r["values"]
        app_id = str(v.get("Application ID", "")).strip()
        if not app_id or str(v.get("Has Resume", "")).strip().lower() != "yes":
            continue
        full_name = v.get("Full Name", "")
        category = v.get("Category", "")
        stored = _stored_resume_names(app_id, v.get("Original Filename", ""),
                                      full_name=full_name, category=category)
        if not stored:
            continue

        received = _parse_received(v.get("Received Date"))
        month_subpath = dated_subpath(received)

        if rejected:
            if _consolidate_rejected_folder(client, stored, month_subpath, dry_run):
                moved += 1
            subpath = _rejected_subpath(month_subpath)
        else:
            subpath = month_subpath

        name_part = _get_cleaned_filename_prefix(full_name)
        cat_part = _clean_category_for_filename(category)
        tail = _resume_filename_tail(app_id)

        # `stored` can hold several fallback candidate names for the SAME single resume
        # (e.g. a stale original-filename fallback that was never actually the live file) -
        # all of them would compute the identical target name, so stop at the first one that
        # actually renames (or, in dry-run, the first whose target differs) rather than
        # attempting every candidate and risking a second real file colliding into the same
        # target name.
        new_names = []
        renamed_this_row = False
        for name in stored:
            _, ext = os.path.splitext(name)
            target = f"{name_part}_{cat_part}_{tail}{ext}"
            if name == target:
                new_names.append(name)
                continue
            if dry_run:
                logger.info(f"  [DRY-RUN] {app_id}: '{name}' -> '{target}'")
                new_names.append(target)
                renamed_this_row = True
                break
            try:
                if client.rename_resume(name, target, subfolder=subpath):
                    logger.info(f"  {app_id}: '{name}' -> '{target}'")
                    new_names.append(target)
                    renamed_this_row = True
                    break
                else:
                    new_names.append(name)
            except Exception as e:
                logger.warning(f"  WARNING  {app_id}: could not rename '{name}': {e}")
                errors += 1
                new_names.append(name)

        if renamed_this_row:
            renamed += 1
            if not dry_run:
                url, path = _resume_url_and_path(client, new_names, subpath)
                if url:
                    fields = {"Resume URL": url, "Resume Folder Path": path}
                    writer = client.update_rejected_row if rejected else client.update_row
                    try:
                        writer(r["index"], fields, current_values=v)
                    except Exception as e:
                        logger.warning(f"  WARNING  {app_id}: could not update Resume URL/Path: {e}")
        else:
            skipped += 1
    return renamed, skipped, errors, moved


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--live", action="store_true",
                    help="Actually move/rename files and update Resume URL/Path (default: dry-run preview only).")
    args = ap.parse_args()
    dry_run = not args.live

    from sharepoint_client import SharePointClient
    client = SharePointClient()

    logger.info("=" * 64)
    logger.info(f"  {'DRY-RUN preview' if dry_run else 'LIVE migration'}: resume filenames + "
                "Rejected folder consolidation")
    logger.info("=" * 64)

    logger.info("")
    logger.info("Main sheet:")
    # Only already-Scored rows have a real Category - an unscored 'New Email Received' row's
    # blank Category would otherwise fall back to 'General', prematurely renaming (and
    # mislabeling) a resume nobody has actually categorized yet.
    scored_rows = [r for r in client.list_rows()
                  if str(r["values"].get("Status", "")).strip() == STATUS_SCORED]
    main_renamed, main_skipped, main_errors, _ = _migrate_sheet(
        client, scored_rows, "main", False, dry_run)
    logger.info(f"  {main_renamed} renamed | {main_skipped} already correct | {main_errors} errors")

    logger.info("")
    logger.info("Rejected sheet:")
    rej_renamed, rej_skipped, rej_errors, rej_moved = _migrate_sheet(
        client, client.list_rejected_rows(), "rejected", True, dry_run)
    logger.info(f"  {rej_renamed} renamed | {rej_skipped} already correct | {rej_errors} errors | "
                f"{rej_moved} candidate(s) had a file moved into the year-level Rejected folder")

    logger.info("")
    logger.info(f"TOTAL: {main_renamed + rej_renamed} renamed | "
                f"{main_skipped + rej_skipped} already correct | {main_errors + rej_errors} errors")
    if dry_run:
        logger.info("")
        logger.info("Dry-run only, nothing changed. Re-run with --live to actually rename/move.")


if __name__ == "__main__":
    main()
