"""Targeted post-write verification without triggering JD refresh/scoring."""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path

from dotenv import load_dotenv

load_dotenv(Path(__file__).with_name(".env"))

from hiring_agent.extraction import extract_text_from_bytes
from repair_candidate_list_batches import _filename_from_url
from sharepoint_client import SharePointClient


def main(app_id: str, report: Path) -> int:
    client = SharePointClient()
    payload = json.loads(report.read_text(encoding="utf-8"))
    expected = next(
        item for item in payload.get("results", [])
        if str(item.get("application_id", "")) == app_id
    )
    main_row = next(
        (row for row in client.list_rows()
         if str(row["values"].get("Application ID", "")) == app_id), None)
    if main_row is not None:
        sheet, table, row = "CandidateList", client.table, main_row
    else:
        rejected_table = client._rejected_table_name_if_exists()
        rejected_row = next(
            (row for row in client.list_rejected_rows()
             if str(row["values"].get("Application ID", "")) == app_id), None)
        if not rejected_table or rejected_row is None:
            raise RuntimeError(f"{app_id} not found in CandidateList or Rejected")
        sheet, table, row = "Rejected", rejected_table, rejected_row
    vals = row["values"]
    folder = str(vals.get("Resume Folder Path", "") or "").strip().strip("/")
    filename = _filename_from_url(vals.get("Resume URL"))
    raw = client.download_file(folder, filename)
    sha256 = hashlib.sha256(raw).hexdigest()
    text_chars = len(extract_text_from_bytes(raw, filename).strip())
    formula = client._column_formula_at(table, "Resume Link", row["index"])
    expected_hash = str(expected.get("source_sha256") or "")
    checks = {
        "application_id": app_id,
        "sheet": sheet,
        "table_index": row["index"],
        "status": vals.get("Status", ""),
        "filename": filename,
        "folder": folder,
        "bytes": len(raw),
        "sha256_matches": not expected_hash or sha256 == expected_hash,
        "text_chars": text_chars,
        "resume_url_present": bool(vals.get("Resume URL")),
        "resume_link_formula": formula,
        "resume_link_clickable": formula.startswith("=") and "HYPERLINK(" in formula.upper(),
    }
    print(json.dumps(checks, indent=2, ensure_ascii=False))
    okay = (
        checks["sha256_matches"] and text_chars >= 80
        and checks["resume_url_present"] and checks["resume_link_clickable"]
    )
    return 0 if okay else 1


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--app-id", required=True)
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()
    raise SystemExit(main(args.app_id, args.report))
