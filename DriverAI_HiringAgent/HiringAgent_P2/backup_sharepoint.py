"""One-off, read-only backup: dumps every row from both the main CandidateList table and
the Rejected table to timestamped JSON files, before a destructive operation (e.g. the
full historical-replay test wipe). Makes no changes to SharePoint - pure read + local write.

Usage:
    python backup_sharepoint.py
"""
import datetime
import json
from pathlib import Path
from urllib.parse import quote

from dotenv import load_dotenv

load_dotenv(Path(__file__).with_name(".env"))

from sharepoint_client import SharePointClient

OUT_ROOT = Path(__file__).parent / "P2_Logs" / "backups"


def _resume_inventory(client: SharePointClient) -> list[dict]:
    """Read-only recursive inventory of the configured SharePoint resume tree."""
    files: list[dict] = []
    stack = [client.resumes_folder.strip("/")]
    seen: set[str] = set()
    while stack:
        folder = stack.pop()
        if folder.lower() in seen:
            continue
        seen.add(folder.lower())
        for item in client.list_folder_children(folder):
            name = str(item.get("name", "") or "")
            path = f"{folder.rstrip('/')}/{name}".strip("/")
            if item.get("is_folder"):
                stack.append(path)
            else:
                files.append({"folder": folder, "name": name, "path": path})
    return sorted(files, key=lambda item: item["path"].lower())


def main() -> None:
    client = SharePointClient()
    print(f"Connected: {client.hostname}")

    main_rows = client.list_rows()
    rej_rows = client.list_rejected_rows()
    raw_main_rows = client._rows_paged(
        f"{client._wb_base()}/tables/{client.table}/rows")
    resume_link_ref = f"columns('{quote('Resume Link')}')"
    resume_link_formulas = client._req(
        "GET",
        f"{client._wb_base()}/tables/{client.table}/{resume_link_ref}/"
        "dataBodyRange?$select=values,formulas,formulasLocal,text",
    ).json()
    resume_inventory = _resume_inventory(client)
    descending_manifest = [
        {
            "position": position,
            "table_index": row.get("index"),
            "excel_row": (row.get("index") + 2
                          if isinstance(row.get("index"), int) else None),
            "application_id": row["values"].get("Application ID", ""),
        }
        for position, row in enumerate(reversed(main_rows), 1)
    ]
    print(f"  main rows     : {len(main_rows)}")
    print(f"  rejected rows : {len(rej_rows)}")

    stamp = datetime.datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    out_dir = OUT_ROOT / stamp
    out_dir.mkdir(parents=True, exist_ok=True)

    (out_dir / "main_rows.json").write_text(
        json.dumps(main_rows, indent=2, ensure_ascii=False), encoding="utf-8")
    (out_dir / "rejected_rows.json").write_text(
        json.dumps(rej_rows, indent=2, ensure_ascii=False), encoding="utf-8")
    (out_dir / "main_table_rows_raw.json").write_text(
        json.dumps(raw_main_rows, indent=2, ensure_ascii=False), encoding="utf-8")
    (out_dir / "resume_link_range.json").write_text(
        json.dumps(resume_link_formulas, indent=2, ensure_ascii=False), encoding="utf-8")
    (out_dir / "resume_inventory.json").write_text(
        json.dumps(resume_inventory, indent=2, ensure_ascii=False), encoding="utf-8")
    (out_dir / "descending_manifest.json").write_text(
        json.dumps(descending_manifest, indent=2, ensure_ascii=False), encoding="utf-8")

    print(f"\nBackup written to: {out_dir}")
    print(f"  main_rows.json     ({len(main_rows)} rows)")
    print(f"  rejected_rows.json ({len(rej_rows)} rows)")
    print(f"  main_table_rows_raw.json ({len(raw_main_rows)} raw rows)")
    print(f"  resume_link_range.json ({len(resume_link_formulas.get('values', []))} rows)")
    print(f"  resume_inventory.json ({len(resume_inventory)} files)")
    print(f"  descending_manifest.json ({len(descending_manifest)} candidates)")


if __name__ == "__main__":
    main()
