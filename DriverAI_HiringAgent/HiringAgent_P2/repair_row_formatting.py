"""Repair row-level formatting (wrap-text, row height) on the live workbook.

Graph's tables/rows/add endpoint writes cell VALUES only - it does not reliably carry
cell format the way typing a row into the Excel UI does (the same class of gap already
found and patched once for the calculated 'Resume Link' formula, see
repair_resume_link_formula.py). A row written this way can come back with wrap text
implicitly on and a stretched row height, which reads as a visibly 'broken' row next to
its plain single-line neighbours.

sharepoint_client.py's add_main_row/add_rejected_row now self-heal this for every NEW row
going forward (see SharePointClient._reapply_row_format). This script is the one-time
backfill for rows that were already written before that fix existed - it walks every
EXISTING physical row on both sheets and re-applies the same plain/auto-height format.

Usage:
    python repair_row_formatting.py --dry-run   # print row counts, write nothing
    python repair_row_formatting.py              # apply for real
"""

import sys

import hiring_agent.config  # noqa: F401 - loads .env and logging config
from sharepoint_client import SharePointClient

DRY_RUN = "--dry-run" in sys.argv


def _repair_table(client: SharePointClient, table_name: str | None, sheet_name: str) -> int:
    if not table_name:
        return 0
    rows = client._rows_paged(f"{client._wb_base()}/tables/{table_name}/rows")
    if DRY_RUN:
        return len(rows)
    for r in rows:
        idx = int(r.get("index", 0) or 0)
        client._reapply_row_format(table_name, sheet_name, idx)
    return len(rows)


def main() -> None:
    client = SharePointClient()
    print("Repairing row formatting (wrap-text off, auto-height)..."
          + (" [DRY RUN]" if DRY_RUN else ""))

    main_count = _repair_table(client, client.table, "CandidateList")
    print(f"  Main table ({client.table}): {main_count} row(s)"
          + ("" if DRY_RUN else " reformatted"))

    rejected_table = client._rejected_table_name_if_exists()
    if rejected_table:
        rejected_count = _repair_table(client, rejected_table, "Rejected")
        print(f"  Rejected table ({rejected_table}): {rejected_count} row(s)"
              + ("" if DRY_RUN else " reformatted"))
    else:
        print("  Rejected table: not present")

    print("Done." + (" (dry run - nothing written)" if DRY_RUN else ""))


if __name__ == "__main__":
    main()
