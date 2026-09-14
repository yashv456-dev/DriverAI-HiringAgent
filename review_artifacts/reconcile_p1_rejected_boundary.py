"""One-time live repair: restore P2-rejected intake rows to P1 CandidateList."""
from __future__ import annotations

import datetime as dt
import json
import sqlite3
import sys
import time
from pathlib import Path

from dotenv import load_dotenv


ROOT = Path(__file__).resolve().parents[1]
P2 = ROOT / "DriverAI_HiringAgent" / "HiringAgent_P2"
sys.path.insert(0, str(P2))
load_dotenv(P2 / ".env")

from sharepoint_client import SharePointClient


def real_rows(rows):
    return [r for r in rows
            if str(r["values"].get("Application ID", "") or "").strip()]


def main():
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    out = ROOT / "review_artifacts" / f"p1_rejected_boundary_{stamp}"
    out.mkdir(parents=True)

    client = SharePointClient()
    db_path = P2 / "candidates.db"
    db = sqlite3.connect(db_path)
    db.row_factory = sqlite3.Row

    workbook_path = str(client._wb_path).strip("/")
    folder, _, name = workbook_path.rpartition("/")
    (out / "Sharepoint_Master_File_before.xlsx").write_bytes(
        client.download_file(folder, name))
    backup = sqlite3.connect(out / "candidates_before.db")
    db.backup(backup)
    backup.close()

    main_before = real_rows(client.list_rows())
    rejected_before = real_rows(client.list_rejected_rows())
    main_ids = {str(r["values"]["Application ID"]).strip() for r in main_before}
    rejected_ids = {str(r["values"]["Application ID"]).strip() for r in rejected_before}
    if main_ids & rejected_ids:
        raise RuntimeError(f"Cross-sheet duplicates already exist: {sorted(main_ids & rejected_ids)}")

    local = {}
    for row in db.execute("SELECT app_id,values_json FROM candidates WHERE active=1"):
        local[row["app_id"]] = json.loads(row["values_json"])
    if not rejected_ids <= set(local):
        raise RuntimeError(f"Rejected rows missing from P2 store: {sorted(rejected_ids - set(local))}")

    columns = client.table_columns()
    plan = []
    for rejected in rejected_before:
        app_id = str(rejected["values"]["Application ID"]).strip()
        source = []
        for event in db.execute(
                "SELECT payload FROM source_events WHERE app_id=? ORDER BY imported_at",
                (app_id,)):
            sheet, values = json.loads(event["payload"])
            if sheet == "main":
                source.append(values)
        if not source:
            raise RuntimeError(f"No original P1 CandidateList event for {app_id}")
        original = source[-1]
        result = local[app_id]
        restored = {col: original.get(col, "") for col in columns}
        restored["Status"] = result.get("Status", "")
        restored["Resume Link"] = result.get("Resume Link", "")
        plan.append({
            "app_id": app_id,
            "p1_name": original.get("Full Name", ""),
            "p2_name": result.get("Full Name", ""),
            "status": restored["Status"],
            "resume_link": restored["Resume Link"],
            "row": restored,
        })

    (out / "repair_plan.json").write_text(json.dumps(plan, indent=2), encoding="utf-8")

    # Restore the immutable intake record first. If a later delete fails, the data exists on
    # both sheets and remains recoverable; it is never lost between two Graph operations.
    for item in plan:
        client.add_main_row(item["row"])

    for _ in range(5):
        current = real_rows(client.list_rows())
        current_ids = {str(r["values"]["Application ID"]).strip() for r in current}
        if rejected_ids <= current_ids:
            break
        time.sleep(2)
    else:
        raise RuntimeError("Not every restored row appeared on CandidateList; Rejected was left intact")

    # rows/add refreshes the calculated Resume Link column for the whole table. Re-apply the
    # two allowed P2 values once, after every row exists, so all 29 intake rows are consistent.
    for row in current:
        app_id = str(row["values"]["Application ID"]).strip()
        result = local.get(app_id)
        if result is None:
            raise RuntimeError(f"CandidateList row absent from P2 store: {app_id}")
        client.update_row(row["index"], {
            "Status": result.get("Status", ""),
            "Resume Link": result.get("Resume Link", ""),
        }, current_values=row["values"])

    # Delete from the bottom so every remaining Graph row index stays valid.
    for row in sorted(rejected_before, key=lambda r: int(r["index"]), reverse=True):
        client.delete_rejected_row(int(row["index"]))

    main_after = real_rows(client.list_rows())
    rejected_after = real_rows(client.list_rejected_rows())
    ids_after = [str(r["values"]["Application ID"]).strip() for r in main_after]
    expected = main_ids | rejected_ids
    if len(ids_after) != len(set(ids_after)) or set(ids_after) != expected:
        raise RuntimeError("Final CandidateList identity check failed")
    if rejected_after:
        raise RuntimeError("Real rows remain on P1 Rejected after repair")

    p2_columns = {
        "Category", "Phone", "Location", "Country", "Years Exp", "Current Skills",
        "Education", "Education Start Date", "Education End Date", "Looking For Role",
        "Suggested Role 1", "Suggested Role 2", "Suggested Role 3",
        "Portfolio 1", "Portfolio 2", "Portfolio 3", "Resume URL", "Resume Folder Path",
        "Retry Count", "Info Request Sent",
    }
    violations = []
    for row in main_after:
        values = row["values"]
        bad = [col for col in p2_columns if str(values.get(col, "") or "").strip()]
        if bad:
            violations.append({"app_id": values.get("Application ID"), "columns": sorted(bad)})
    if violations:
        raise RuntimeError(f"P2 profile values remain in P1 CandidateList: {violations}")

    (out / "Sharepoint_Master_File_after.xlsx").write_bytes(
        client.download_file(folder, name))
    report = {
        "backup": str(out),
        "candidate_list_before": len(main_before),
        "rejected_before": len(rejected_before),
        "candidate_list_after": len(main_after),
        "rejected_after": len(rejected_after),
        "unique_application_ids_after": len(set(ids_after)),
        "restored": [item["app_id"] for item in plan],
    }
    (out / "result.json").write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(json.dumps(report, indent=2))
    db.close()


if __name__ == "__main__":
    main()
