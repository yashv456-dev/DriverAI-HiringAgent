"""One-shot migration: re-derive Category for rows sitting on a retired/split department.

Background
----------
Three category splits landed on 2026-08-21 (all client requests):
    Senior & Executive -> Senior      (senior, lead, manager, director, head, principal)
                          Executive   (executive, chief, ciso, cmo, cto, ceo, cfo, coo,
                                       cio, officer)
    Business Analytics -> Business Analytics (business analyst/analytics, consultant)
                          Sales             (sales, business development)
                          Marketing         (marketing, marketer)
    Graphics           -> Gaming            (gaming, game)
                          Graphics          (graphics, rendering, design, animation)

Rows scored before a split still carry the old value. 'Senior & Executive' no longer
exists at all, so it shows up as a phantom bucket in the client's Category filter;
'Business Analytics' and 'Graphics' still exist but now cover less than they used to, so
rows that should have moved to the new departments are simply sitting in the wrong one.

Re-deriving is safe and idempotent: a row already in the right category re-derives to the
same value and is skipped, so this can be re-run after any future split.

Why this and not --recheck-all
------------------------------
--recheck-all also re-validates geo, de-duplicates, and CAN SEND DECLINE EMAIL - and
applicant mail is LIVE (config.yaml test_mode.suppress_emails: false since 2026-08-11).
This script writes exactly one cell per row (Category) on rows whose Category is the
retired value. It never emails, never moves a row between sheets, never touches a resume,
and never edits any other column. update_row() merges over the row's existing values, so
unspecified columns - including the 'Resume Link' formula - are preserved.

Usage
-----
    python migrate_senior_executive_split.py            # dry run, prints the plan
    python migrate_senior_executive_split.py --apply    # writes the Category cell
    python migrate_senior_executive_split.py --category "Graphics" --apply   # one bucket
"""
from __future__ import annotations

import argparse

from dotenv import load_dotenv

load_dotenv(".env")

from hiring_agent.config import logger                      # noqa: E402  (after load_dotenv)
from hiring_agent.scoring import assign_category            # noqa: E402
from sharepoint_client import SharePointClient

from hiring_agent.store import ExcelCandidateStore, APP_ID_COL


def _store(client):
    """Row identity — see hiring_agent/store.py. Rows are addressed by Application ID, never
    by position, so a Phase 1 append between this tool's read and its write cannot make it
    patch the wrong candidate."""
    return ExcelCandidateStore(client)


def _app_id_of(values) -> str:
    return str((values or {}).get(APP_ID_COL, "") or "").strip()
              # noqa: E402

# Every department that was split on 2026-08-21. A row sitting in one of these may belong
# in a narrower department now; rows in any OTHER category were never affected and are left
# strictly alone, so this can never reshuffle a category the client did not ask about.
SPLIT_SOURCES = ["Senior & Executive", "Business Analytics", "Graphics"]


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--apply", action="store_true",
                    help="Actually write the re-derived Category (default: dry run).")
    ap.add_argument("--category", action="append", metavar="NAME",
                    help=f"Only migrate rows currently in NAME (repeatable). "
                         f"Default: {', '.join(SPLIT_SOURCES)}")
    args = ap.parse_args()

    sources = args.category or SPLIT_SOURCES

    client = SharePointClient()
    client.ensure_workbook()

    candidates = [r for r in client.list_rows()
                  if str(r["values"].get("Category", "")).strip() in sources]

    # Re-derive first, then keep only the rows that actually CHANGE. A row already in the
    # right department (e.g. a genuine Graphics row that stays Graphics) is not rewritten,
    # so --apply touches the minimum number of cells and is safe to re-run.
    planned = []
    for row in candidates:
        vals = row["values"]
        old_cat = str(vals.get("Category", "")).strip()
        # Same inputs the scorer itself uses, so a migrated row lands exactly where a
        # freshly-scored one would: Suggested Role 1 as the title, Current Skills for the
        # native-mobile override.
        new_cat = assign_category(vals.get("Suggested Role 1", ""),
                                  vals.get("Current Skills", ""))
        if new_cat != old_cat:
            planned.append((row, old_cat, new_cat))

    logger.info(f"Scanned {len(candidates)} row(s) in {sources}; "
                f"{len(planned)} need(s) a new Category.")
    if not planned:
        logger.info("Nothing to do.")
        return

    for row, old_cat, new_cat in planned:
        vals = row["values"]
        logger.info(f"   {vals.get('Application ID', '?'):<28} "
                    f"{old_cat!r} -> {new_cat!r}   "
                    f"[role 1: {str(vals.get('Suggested Role 1', ''))[:52]!r}]")

    if not args.apply:
        from collections import Counter
        logger.info("")
        logger.info("DRY RUN - nothing written. Rows would move to:")
        for cat, n in Counter(c for _, _, c in planned).most_common():
            logger.info(f"   {n:3d}  {cat}")
        logger.info("Re-run with --apply to write these Category values.")
        return

    written = 0
    for row, _old_cat, new_cat in planned:
        try:
            _store(client).save_by_id(_app_id_of(row["values"]), {"Category": new_cat},
                                      current_values=row["values"], hint=row["index"])
            written += 1
        except Exception as e:                              # keep going; report at the end
            logger.warning(f"   FAILED {row['values'].get('Application ID', '?')}: {e}")
    logger.info(f"Done. {written}/{len(planned)} row(s) updated.")


if __name__ == "__main__":
    main()
