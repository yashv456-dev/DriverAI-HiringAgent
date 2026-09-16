"""Initialize candidates.db and generate clean starter sheet for Phase 2.

Can initialize an empty database with schemas or seed from the 217-row master snapshot
so that Candidate_List_Results.xlsx is immediately populated with real records.
"""

import os
import sys
from pathlib import Path
import openpyxl

APP_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(APP_DIR))

from hiring_agent.store import SQLiteCandidateStore, get_store
from hiring_agent.sqlite_store import database_path
from hiring_agent.exporter import ExcelReportExporter

REPAIRED_MASTER_PATH = APP_DIR / "P2_Final_Results" / "Sharepoint_Master_File.xlsx"
SNAPSHOT_PATH = REPAIRED_MASTER_PATH if REPAIRED_MASTER_PATH.exists() else (APP_DIR / "P2_Logs" / "audits" / "master_row_audit_20260906" / "master_snapshot.xlsx")
DB_PATH = database_path()


def init_empty_db():
    """Create or migrate the database without deleting existing records."""
    store = SQLiteCandidateStore(DB_PATH)
    print(f"[OK] Initialized empty database at: {DB_PATH}")
    return store


def seed_from_snapshot(snapshot_path: Path = SNAPSHOT_PATH):
    """Seed candidates.db from the master_snapshot.xlsx file."""
    if not snapshot_path.exists():
        print(f"[WARN] Snapshot not found at {snapshot_path}")
        return init_empty_db()

    store = init_empty_db()
    if store.conn.execute("SELECT COUNT(*) FROM candidates").fetchone()[0]:
        print("[OK] Existing candidate database preserved; seed skipped.")
        return store
    wb = openpyxl.load_workbook(snapshot_path, data_only=True)
    
    total_added = 0
    # 'Unfamiliar Role List' is a presentation split of the main sheet, not a third place a
    # candidate lives - they are ordinary main rows whose Status says nothing matched. It has
    # to be listed here or rebuilding the database from a published workbook would silently
    # drop every one of them.
    for sheet_name, target_sheet in [("CandidateList", "main"), ("Rejected", "rejected"),
                                     ("Unfamiliar Role List", "main")]:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        headers = [str(cell.value or "").strip() for cell in ws[1]]
        
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not any(row):
                continue
            row_dict = {}
            for h, v in zip(headers, row):
                if h:
                    row_dict[h] = "" if v is None else str(v).strip()
            
            app_id = str(row_dict.get("Application ID", "")).strip()
            if not app_id:
                continue
            
            store.add(row_dict, sheet=target_sheet)
            total_added += 1

    print(f"[OK] Seeded {total_added} candidates into {DB_PATH}")
    
    # Generate fresh populated results workbook
    exporter = ExcelReportExporter(store)
    out_file = exporter.generate_workbook(APP_DIR / "P2_Final_Results" / "Candidate_List_Results.xlsx")
    print(f"[OK] Generated populated results workbook at: {out_file}")
    return store


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Initialize candidates.db")
    parser.add_argument("--empty", action="store_true", help="Create empty database without seeding")
    args = parser.parse_args()

    if args.empty:
        init_empty_db()
    else:
        seed_from_snapshot()
