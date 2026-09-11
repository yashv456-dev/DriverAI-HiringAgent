"""Migration, client-requested. Re-run whenever the naming convention changes:
  1. Rename every already-saved resume to the CURRENT canonical format,
     '<First>_<Last>_<Category>_<tail>' (e.g. 'Yash_Verma_DataAnalytics_7693.pdf').
     It reads the target from _canonical_resume_name rather than rebuilding it, so each
     re-run migrates from whatever the previous convention was - 2026-07-15's
     'JaneDoe_APP-20260710-2200-A5F2.pdf', 2026-07-15..08-04's
     'JaneDoe_DataAnalytics_A5F2.pdf', or 2026-08-04..08-06's 'Jane_Doe_A5F2.pdf'.
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
from hiring_agent.store import ExcelCandidateStore


def _save(client, r, fields, *, rejected: bool):
    """Patch a row by Application ID, never by position.

    This tool reached `client.update_row` through a `writer` variable, so the 2026-09-06
    conversion to key-based addressing missed it: a grep for `client.update_row(` cannot see
    an indirect call. Phase 1 appends to the same table every 60 seconds and this tool walks
    the whole sheet, so a positional write here can land on the wrong candidate.

    A row with no Application ID has no key to address by and still falls back to a
    positional write - the same trade `_save_scan_row` documents in sharepoint_scoring.py.
    """
    app_id = str(r["values"].get("Application ID", "") or "").strip()
    if not app_id:
        writer = client.update_rejected_row if rejected else client.update_row
        writer(r["index"], fields, current_values=r["values"])
        return
    ExcelCandidateStore(client).save_by_id(
        app_id, fields, current_values=r["values"],
        sheet="rejected" if rejected else "main", hint=r["index"])


from hiring_agent.sharepoint_scoring import (
    _stored_resume_names, _canonical_resume_name, _resume_file_exists,
    _resume_url_and_path, _rejected_subpath, _is_gap, _STATUS_NEW,
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
    renamed = skipped = errors = moved = resynced = 0
    for r in rows:
        v = r["values"]
        app_id = str(v.get("Application ID", "")).strip()
        if not app_id or str(v.get("Has Resume", "")).strip().lower() != "yes":
            continue
        full_name = v.get("Full Name", "")
        category = v.get("Category", "")
        # resume_url is ESSENTIAL here, not optional. Without it _stored_resume_names can
        # only GUESS filenames from Full Name + Category - and the pre-2026-08-04 naming
        # baked both of those mutable fields into the filename, so once either changed the
        # file became unfindable by name. Two live rows proved it: 6F75's file is
        # 'SyyedAli_MobileAppsAndroidIOS_6F75.pdf' while its row's Category has since moved
        # on, and 09AC's is 'Candidate_General_09AC.pdf' from when its Full Name was still
        # blank. Both sat in the folder the migration was searching and were still reported
        # NOT FOUND, so both kept their stale category-bearing names.
        #
        # The Resume URL records the file that is ACTUALLY there, so _stored_resume_names
        # puts it first and the guesses become a fallback rather than the only hope.
        stored = _stored_resume_names(app_id, v.get("Original Filename", ""),
                                      full_name=full_name, category=category,
                                      resume_url=v.get("Resume URL"))
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
            target = _canonical_resume_name(full_name, app_id, ext, category)
            if name == target:
                new_names.append(name)
                continue
            if dry_run:
                # `stored` is a list of GUESSES, most of which never existed on disk. The
                # live branch below discovers that naturally (rename_resume returns False and
                # it moves to the next guess), but a preview that just printed the first guess
                # would name files that do not exist - useless for reviewing before a live
                # run. Confirm existence read-only, exactly as _consolidate_rejected_folder
                # already does for its own preview.
                if not _resume_file_exists(client, name, subpath):
                    continue
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
                    try:
                        _save(client, r, fields, rejected=rejected)
                    except Exception as e:
                        logger.warning(f"  WARNING  {app_id}: could not update Resume URL/Path: {e}")
        elif dry_run and not any(
                _resume_file_exists(client, n, subpath) for n in stored):
            # No guessed name exists at this location at all. Reporting that as "already
            # correct" (the old behaviour, since both fall into this else) would hide it:
            # the row's resume is genuinely unfindable under every name we know, which is a
            # problem to look at BEFORE a live run, not a no-op to skip past.
            logger.warning(f"  NOT FOUND {app_id}: no resume file at '{subpath}' under any "
                           f"known name ({', '.join(stored[:3])}...)")
            errors += 1
        elif not dry_run:
            # The FILE is already canonically named, but the ROW may still point at the old
            # one. The rename and the Resume URL write are two separate Graph calls, so a
            # throttle or failure between them leaves a correctly-renamed file behind a DEAD
            # link - and a re-run could never repair it, because this branch simply counted
            # the row as "already correct" and moved on. Live case 2026-08-11:
            # APP-20260713-0725-E484 was renamed while a Graph 503 hit its URL write, and two
            # further full migrations both reported it as fine.
            present = [n for n in new_names if _resume_file_exists(client, n, subpath)]
            if present:
                url, path = _resume_url_and_path(client, present, subpath)
                if url and url != str(v.get("Resume URL", "") or ""):
                    try:
                        _save(client, r, {"Resume URL": url, "Resume Folder Path": path},
                              rejected=rejected)
                        logger.info(f"  {app_id}: file already named '{present[0]}' - stale "
                                    f"Resume URL re-synced to it")
                        resynced += 1
                    except Exception as e:
                        logger.warning(f"  WARNING  {app_id}: could not re-sync Resume URL: {e}")
                        errors += 1
            skipped += 1
        else:
            skipped += 1
    return renamed, skipped, errors, moved, resynced


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
    # Two preconditions, both "would this rename immediately be undone by P2?":
    #
    #   Full Name  - a row still at 'New Email Received' has not been extracted, so its name
    #                may be blank and the target would be 'Candidate_...', renamed again the
    #                moment P2 scores it.
    #   Category   - restored as a precondition on 2026-08-06 when Category came back into
    #                the filename. An unscored row's blank Category resolves to 'General'
    #                (see _clean_category_for_filename), which would bake a placeholder into
    #                a real filename and force a second rename after scoring. It had been
    #                dropped on 08-04 only because that convention used no Category at all.
    #
    # Rows failing either check are left alone for P2 to name normally on its next pass.
    scored_rows = [r for r in client.list_rows()
                   if str(r["values"].get("Status", "")).strip() not in ("", _STATUS_NEW)
                   and not _is_gap(r["values"].get("Full Name"))
                   and not _is_gap(r["values"].get("Category"))]
    main_renamed, main_skipped, main_errors, _, main_resynced = _migrate_sheet(
        client, scored_rows, "main", False, dry_run)
    logger.info(f"  {main_renamed} renamed | {main_skipped} already correct | "
                f"{main_resynced} stale URL(s) re-synced | {main_errors} errors")

    logger.info("")
    logger.info("Rejected sheet:")
    rej_renamed, rej_skipped, rej_errors, rej_moved, rej_resynced = _migrate_sheet(
        client, client.list_rejected_rows(), "rejected", True, dry_run)
    logger.info(f"  {rej_renamed} renamed | {rej_skipped} already correct | "
                f"{rej_resynced} stale URL(s) re-synced | {rej_errors} errors | "
                f"{rej_moved} candidate(s) had a file moved into the year-level Rejected folder")

    logger.info("")
    logger.info(f"TOTAL: {main_renamed + rej_renamed} renamed | "
                f"{main_skipped + rej_skipped} already correct | "
                f"{main_resynced + rej_resynced} stale URL(s) re-synced | "
                f"{main_errors + rej_errors} errors")
    if dry_run:
        logger.info("")
        logger.info("Dry-run only, nothing changed. Re-run with --live to actually rename/move.")


if __name__ == "__main__":
    main()
