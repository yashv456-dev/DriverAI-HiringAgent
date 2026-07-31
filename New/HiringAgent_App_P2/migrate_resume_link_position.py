"""One-time live migration: move the 'Resume Link' column to the 5th position (right after
'Category', before 'Full Name') on BOTH the main CandidateList table and the Rejected table
of the live SharePoint master workbook. Client request, 2026-07-24.

WHY THIS IS SAFE (zero data loss): 'Resume Link' is an Excel CALCULATED column - every cell
is derived by a single per-row HYPERLINK formula from that row's 'Resume URL'/'Original
Filename'. So the migration is simply: delete the column, re-add it at the new position, and
re-apply the same formula. No row values are lost because none were ever stored - they are all
recomputed by Excel. Every other read/write in P1 and P2 is keyed by column NAME, never index,
so nothing else is affected by the move.

The blank template, config.yaml schema, client-result export, and both test suites were already
updated in code; this script brings the pre-existing LIVE workbook into line with them.

Default is DRY-RUN (prints the plan, writes nothing). Pass --apply to actually migrate.

    python migrate_resume_link_position.py            # dry-run
    python migrate_resume_link_position.py --apply     # do it
"""
from __future__ import annotations

import argparse
import sys

from dotenv import load_dotenv
load_dotenv(".env")

from sharepoint_client import SharePointClient, SharePointError
from hiring_agent.sharepoint_scoring import _RESUME_LINK_FORMULA
import hiring_agent.config as _cfg

TARGET_INDEX = 4   # 0-based -> physical column 5 (after Application ID, Received Date,
                   # Last Updated Date, Category)
COL = "Resume Link"


def _add_column_at(client: SharePointClient, table_name: str, name: str, index: int) -> None:
    """Add a fresh (empty) table column at an explicit 0-based position."""
    client._req("POST", f"{client._wb_base()}/tables/{table_name}/columns/add",
                json={"name": name, "index": index})
    client._columns_cache.pop(table_name, None)


def _formula_rows(client: SharePointClient, table_name: str) -> int:
    """How many rows of COL currently hold a '='-formula (0 if the column is empty/absent)."""
    from urllib.parse import quote
    colref = f"columns('{quote(COL)}')"
    try:
        data = client._req("GET", f"{client._wb_base()}/tables/{table_name}/{colref}"
                                  f"/dataBodyRange?$select=formulas").json()
    except SharePointError:
        return 0
    return sum(1 for row in data.get("formulas", []) if str(row[0]).startswith("="))


def _migrate_table(client: SharePointClient, table_name: str, apply: bool) -> bool:
    cols = client._table_columns_of(table_name)
    if COL not in cols:
        print(f"  [{table_name}] '{COL}' not present - skipping (nothing to move).")
        return False
    cur = cols.index(COL)
    if cur == TARGET_INDEX:
        print(f"  [{table_name}] '{COL}' already at position {TARGET_INDEX + 1} - no change.")
        return False
    print(f"  [{table_name}] '{COL}' is at position {cur + 1}; moving to position {TARGET_INDEX + 1}.")
    if not apply:
        print(f"           [DRY-RUN] would: delete '{COL}', re-add at index {TARGET_INDEX}, "
              f"re-apply calculated formula.")
        return True
    # Calculated column -> delete + re-add is lossless (values are all formula-derived).
    client.delete_column(table_name, COL)
    _add_column_at(client, table_name, COL, TARGET_INDEX)
    # IMPORTANT: right after a column is re-added, Graph can briefly report the new column's
    # dataBodyRange rowCount as 0, and set_calculated_column early-returns on that - leaving the
    # column BLANK (every clickable link broken). Retry until the formula actually sticks on all
    # rows, forcing a fresh column cache read each attempt.
    import time
    applied = 0
    for attempt in range(6):
        client._columns_cache.pop(table_name, None)
        client.set_calculated_column(table_name, COL, _RESUME_LINK_FORMULA)
        applied = _formula_rows(client, table_name)
        if applied > 0:
            break
        time.sleep(2)
    after = client._table_columns_of(table_name)
    pos_ok = COL in after and after.index(COL) == TARGET_INDEX
    ok = pos_ok and applied > 0
    print(f"           {'OK' if ok else 'WARNING'}: '{COL}' now at position "
          f"{after.index(COL) + 1 if COL in after else 'MISSING'}; "
          f"formula on {applied} row(s).")
    if not ok:
        print(f"           !! Re-run the script, or apply the formula manually - the column "
              f"was moved but the calculated formula did not populate.")
    return True


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true", help="actually migrate (default: dry-run)")
    args = ap.parse_args()

    print(f"{'APPLYING' if args.apply else 'DRY-RUN'} Resume Link -> position 5 migration\n")
    client = SharePointClient()
    print(f"Connected to {client.hostname}\n")

    main_tbl = client.table
    rej_tbl = client._rejected_table_name_if_exists()

    _migrate_table(client, main_tbl, args.apply)
    if rej_tbl:
        _migrate_table(client, rej_tbl, args.apply)
    else:
        print("  [Rejected] table not found on the live workbook - skipping.")

    if not args.apply:
        print("\nDry-run only. Re-run with --apply to perform the migration.")
    else:
        print("\nDone. Verify the live workbook shows 'Resume Link' as the 5th column on both sheets.")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SharePointError as e:
        print(f"FAILED: {e}", file=sys.stderr)
        sys.exit(1)
