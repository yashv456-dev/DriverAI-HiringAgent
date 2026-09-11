"""Re-sort the Main (HiringAgent_P1_Candidates) and Rejected sheets so the CURRENT month's
rows come first, most recent within a month first, older months following below - each
month boundary marked by a blank separator row, and each YEAR boundary marked by a labeled
separator (e.g. '-- 2025 --') so a multi-year sheet is never ambiguous about where a year
changed.

SAFETY DESIGN (data on this workbook has been wiped before by careless bulk-write code -
see workbook self-heal history in HiringAgent_P1/flow/build_zip.py):
  1. Back up every row (main + rejected) to local JSON before touching anything.
  2. Compute the full desired final order in memory; verify the set of Application IDs is
     UNCHANGED (nothing lost, nothing duplicated) before writing a single cell.
  3. ADD all rows in final order to the BOTTOM of the table FIRST (genuine AddRowV2 calls -
     the same primitive P1 already uses for every live candidate, so Excel's calculated
     'Resume Link' column auto-fills exactly as it always does). Verify the new count landed.
  4. Only THEN delete the OLD rows (highest index first, so earlier indices never shift
     under the loop). If step 3 fails partway, the script aborts WITHOUT deleting anything -
     the original rows are still 100% intact, at worst with some harmless extra rows to
     clean up by hand.
  5. Final verification: row count and the full set of Application IDs must match exactly.

Usage:
    python resort_candidate_sheets.py --dry-run   # print the plan, write nothing
    python resort_candidate_sheets.py              # back up, then write for real
"""
import datetime as dt
import json
import sys
import time
from pathlib import Path

from dotenv import load_dotenv

load_dotenv(Path(__file__).with_name(".env"))

from sharepoint_client import SharePointClient  # noqa: E402

DRY_RUN = "--dry-run" in sys.argv
YEAR_LABEL_COLUMN = "Full Name"
SEP_LABEL_FMT = "-- {year} --"


def row_empty(vals: dict) -> bool:
    return not any(str(v or "").strip() for v in vals.values())


def parse_received(vals: dict) -> dt.datetime | None:
    text = str(vals.get("Received Date", "") or "").strip()
    if not text:
        return None
    if text.endswith("Z"):
        text = text[:-1]
    try:
        return dt.datetime.fromisoformat(text)
    except ValueError:
        return None


def build_plan(rows: list[dict]) -> tuple[list[dict], list[tuple[int, int, int]]]:
    """rows: [{'values': {...}}, ...] (real rows only, already filtered).
    Returns (ordered_field_dicts_to_add, group_summary) where group_summary is
    [(year, month, count), ...] in the final (descending) order, for the report."""
    dated = []
    undated = []
    for r in rows:
        received = parse_received(r["values"])
        if received is None:
            undated.append(r)
        else:
            dated.append((received, r))

    groups: dict[tuple[int, int], list] = {}
    for received, r in dated:
        groups.setdefault((received.year, received.month), []).append((received, r))

    group_keys = sorted(groups.keys(), reverse=True)  # newest year/month first

    # Disabled 2026-09-01 (client instruction): master sheets must never carry monthly/
    # yearly blank-row gaps - only the client results export does (see
    # sharepoint_scoring.py's _with_period_separators, which is unaffected). No separator
    # or year-label row is inserted anywhere in the plan below, including under the header.
    plan: list[dict] = []
    summary: list[tuple[int, int, int]] = []
    for key in group_keys:
        year, month = key
        members = sorted(groups[key], key=lambda t: t[0], reverse=True)  # newest first within month
        plan.extend(r["values"] for _received, r in members)
        summary.append((year, month, len(members)))

    if undated:
        # Never silently drop a row with a bad/blank Received Date - surface it, append at
        # the very end (past all real groups) rather than guess where it belongs.
        for r in undated:
            plan.append(r["values"])
        summary.append((0, 0, len(undated)))  # (0,0) sentinel = "undated, appended last"

    return plan, summary


def report(label: str, summary: list[tuple[int, int, int]], total_real: int) -> None:
    print(f"\n=== {label}: planned order ({total_real} real rows) ===")
    for year, month, count in summary:
        if (year, month) == (0, 0):
            print(f"  [UNDATED - Received Date missing/unparseable] {count} row(s) appended at the very end")
        else:
            print(f"  {year}-{month:02d}: {count} row(s)")


def backup(client: SharePointClient, main_rows: list[dict], rej_rows: list[dict]) -> Path:
    out_root = Path(__file__).parent / "P2_Logs" / "backups"
    stamp = dt.datetime.utcnow().strftime("%Y%m%dT%H%M%SZ") + "_pre_resort"
    out_dir = out_root / stamp
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "main_rows.json").write_text(json.dumps(main_rows, indent=2, ensure_ascii=False), encoding="utf-8")
    (out_dir / "rejected_rows.json").write_text(json.dumps(rej_rows, indent=2, ensure_ascii=False), encoding="utf-8")
    return out_dir


def resort_table(client: SharePointClient, *, is_rejected: bool, dry_run: bool) -> None:
    label = "REJECTED" if is_rejected else "MAIN"
    if is_rejected:
        tbl = client._rejected_table_name_if_exists()
        if tbl is None:
            print(f"\n=== {label}: no Rejected table exists yet - nothing to do ===")
            return
        raw_rows = client._rows_paged(f"{client._wb_base()}/tables/{tbl}/rows")
        cols = client._table_columns_of(tbl)
        all_rows = client.list_rejected_rows()
        real_rows = [r for r in all_rows if not row_empty(r["values"])]
    else:
        tbl = client.table
        raw_rows = client._rows_paged(f"{client._wb_base()}/tables/{tbl}/rows")
        real_rows = client.list_rows()  # already filters blanks

    if not real_rows:
        print(f"\n=== {label}: no real rows - nothing to reorder (left untouched) ===")
        return

    original_raw_count = len(raw_rows)
    original_app_ids = {str(r["values"].get("Application ID", "")).strip() for r in real_rows}
    original_app_ids.discard("")

    plan, summary = build_plan(real_rows)
    plan_app_ids = {str(f.get("Application ID", "")).strip() for f in plan}
    plan_app_ids.discard("")

    report(label, summary, len(real_rows))
    print(f"  original physical rows (incl. old separators): {original_raw_count}")
    print(f"  new physical rows to write (incl. new separators): {len(plan)}")

    if plan_app_ids != original_app_ids:
        missing = original_app_ids - plan_app_ids
        extra = plan_app_ids - original_app_ids
        raise SystemExit(
            f"ABORT ({label}): Application ID set mismatch between original and planned rows - "
            f"missing={missing} extra={extra}. No changes made."
        )
    print(f"  [OK] Application ID set verified identical ({len(original_app_ids)} candidates)")

    if dry_run:
        print(f"  [DRY RUN] no changes written for {label}")
        return

    add_fn = client.add_rejected_row if is_rejected else client.add_main_row
    added = 0
    try:
        for fields in plan:
            add_fn(fields)
            added += 1
    except Exception as e:
        print(f"  [ERROR] {label}: add failed after {added}/{len(plan)} rows written: {e}")
        print(f"  Original {original_raw_count} rows are UNTOUCHED (deletion never started). "
              f"{added} extra new-order rows are now sitting at the bottom of the table - "
              f"safe to delete by hand or re-run once the underlying issue is fixed.")
        raise SystemExit(1)

    # Excel Online Business has shown real eventual-consistency lag right after a burst of
    # AddRowV2 calls (same class of staleness documented for this connector in P1's flow -
    # confirmed live 2026-07-31: an immediate count read here under-reported by exactly 1 row
    # that a follow-up read moments later showed was actually there). Retry briefly before
    # treating a short count as a real failure - but still abort (never delete) if it never
    # settles, since a genuinely missing row must never be silently written over.
    expected = original_raw_count + len(plan)
    new_raw_rows = []
    for attempt in range(5):
        new_raw_rows = client._rows_paged(f"{client._wb_base()}/tables/{tbl}/rows")
        if len(new_raw_rows) == expected:
            break
        time.sleep(2)
    if len(new_raw_rows) != expected:
        raise SystemExit(
            f"ABORT ({label}): expected {expected} rows after adding, "
            f"found {len(new_raw_rows)} after retrying. Deletion NOT started - investigate "
            f"before proceeding (the {len(plan)} new rows are still there; nothing was deleted)."
        )
    print(f"  [OK] {len(plan)} new-order rows appended and verified")

    delete_fn = client.delete_rejected_row if is_rejected else client.delete_row
    for idx in range(original_raw_count - 1, -1, -1):
        delete_fn(idx)
    print(f"  [OK] {original_raw_count} old rows deleted")

    final_app_ids = set()
    final_real: list = []
    for attempt in range(5):
        final_real = (client.list_rejected_rows() if is_rejected else client.list_rows())
        if is_rejected:
            final_real = [r for r in final_real if not row_empty(r["values"])]
        final_app_ids = {str(r["values"].get("Application ID", "")).strip() for r in final_real}
        final_app_ids.discard("")
        if final_app_ids == original_app_ids:
            break
        time.sleep(2)
    if final_app_ids != original_app_ids:
        raise SystemExit(
            f"DATA INTEGRITY ALERT ({label}): post-write Application ID set does not match "
            f"the original after retrying ({len(final_app_ids)} vs {len(original_app_ids)}). "
            f"Restore from the pre-resort backup immediately."
        )
    print(f"  [OK] Final verification passed: {len(final_real)} real rows, Application ID set intact")

    # Re-apply the 'Resume Link' calculated-column formula (fixed 2026-07-31, live finding):
    # deleting every row and re-adding them can leave the table's calculated column reset to
    # plain blank values instead of a live per-row HYPERLINK formula - confirmed live on this
    # exact table after the first resort. sharepoint_client.set_calculated_column's own retry
    # comment ("Graph can briefly return the pre-delete rowCount after a table row is moved")
    # shows this class of disruption from row deletion is already anticipated elsewhere in the
    # codebase. P2's normal scoring pass only reapplies this formula when
    # HIRING_REPAIR_RESUME_LINK_FORMULA is explicitly set (off by default), so a resort must
    # not rely on the next scoring run to self-heal it - do it here, unconditionally.
    try:
        from hiring_agent.sharepoint_scoring import _RESUME_LINK_FORMULA
        client.set_calculated_column(tbl, "Resume Link", _RESUME_LINK_FORMULA)
        print(f"  [OK] 'Resume Link' calculated-column formula re-applied")
    except Exception as e:
        print(f"  [WARN] Could not re-apply 'Resume Link' formula: {e} - "
              f"run with HIRING_REPAIR_RESUME_LINK_FORMULA=true on the next P2 scoring pass, "
              f"or re-apply it manually.")


def main() -> None:
    client = SharePointClient()
    print(f"Connected: {client.hostname}")

    main_rows_before = client.list_rows()
    rej_rows_before = client.list_rejected_rows()

    if not DRY_RUN:
        out_dir = backup(client, main_rows_before, rej_rows_before)
        print(f"Backup written to: {out_dir}")

    resort_table(client, is_rejected=False, dry_run=DRY_RUN)
    resort_table(client, is_rejected=True, dry_run=DRY_RUN)

    print("\nDone." + (" (dry run - nothing written)" if DRY_RUN else ""))


if __name__ == "__main__":
    main()
