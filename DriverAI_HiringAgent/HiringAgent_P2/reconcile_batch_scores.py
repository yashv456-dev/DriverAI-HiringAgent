"""Reconcile a frozen batch to the final deterministic JD scoring rules."""
from __future__ import annotations

import argparse
import json
import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv(Path(__file__).with_name(".env"))

import hiring_agent.config as cfg
import hiring_agent.scoring as scoring_module
from hiring_agent.excel_output import _parse_received
from hiring_agent.scoring import assign_category, suggested_roles
from hiring_agent.sharepoint_scoring import _RESUME_LINK_FORMULA, _rejected_subpath, _resume_display_path
from repair_candidate_list_batches import (
    REPORT_DIR, _canonical_name, _clean_old_files, _download_current,
    _frozen_cached_roles, _manifest_ids, _real_extension, _verify_uploaded,
)
from sharepoint_client import SharePointClient


def run(batch: int, manifest: Path, apply: bool) -> int:
    cfg.GEO_REJECT_EMAIL = False
    cfg.ERROR_EMAIL_ENABLED = False
    scoring_module.OLLAMA_SCORING = False
    client = SharePointClient()
    roles, built_at = _frozen_cached_roles()
    app_ids = set(_manifest_ids(manifest, batch))
    main = {
        str(row["values"].get("Application ID", "")): ("CandidateList", row)
        for row in client.list_rows()
        if str(row["values"].get("Application ID", "")) in app_ids
    }
    rejected = {
        str(row["values"].get("Application ID", "")): ("Rejected", row)
        for row in client.list_rejected_rows()
        if str(row["values"].get("Application ID", "")) in app_ids
    }
    located = {**rejected, **main}
    results = []
    errors = 0
    print(f"MODE={'APPLY' if apply else 'DRY-RUN'} BATCH={batch} "
          f"JD_ROLES={len(roles)} BUILT_AT={built_at}", flush=True)
    for app_id in _manifest_ids(manifest, batch):
        found = located.get(app_id)
        if not found:
            results.append({"application_id": app_id, "status": "ABSENT"})
            print(f"{app_id}: absent (expected duplicate/removal)", flush=True)
            continue
        sheet, row = found
        vals = row["values"]
        try:
            scored = suggested_roles(
                str(vals.get("Current Skills", "")),
                str(vals.get("Looking For Role", "")), roles=roles)
            desired = {
                "Suggested Role 1": scored["role_1"],
                "Suggested Role 2": scored["role_2"],
                "Suggested Role 3": scored.get("role_3", ""),
                "Category": assign_category(scored["role_1"], str(vals.get("Current Skills", ""))),
            }
            changed = {key: value for key, value in desired.items()
                       if str(vals.get(key, "")) != str(value)}
            result = {
                "application_id": app_id, "sheet": sheet,
                "old_role_1": vals.get("Suggested Role 1", ""),
                "new_role_1": desired["Suggested Role 1"],
                "old_category": vals.get("Category", ""),
                "new_category": desired["Category"],
                "changed_columns": sorted(changed),
                "status": "CHANGE" if changed else "CURRENT",
            }
            print(f"{app_id}: {result['status']} | {result['new_role_1']} | "
                  f"{result['new_category']}", flush=True)
            if apply and changed:
                source_name, raw = _download_current(client, vals)
                extension = _real_extension(raw, source_name)
                month = cfg.dated_subpath(_parse_received(vals.get("Received Date")))
                subpath = month if sheet == "CandidateList" else _rejected_subpath(month)
                folder = f"{client.resumes_folder}/{subpath}".strip("/")
                name_fields = dict(vals)
                name_fields.update(desired)
                canonical = _canonical_name(name_fields, app_id, extension)
                client.upload_file(folder, canonical, raw)
                byte_count, text_chars, web_url = _verify_uploaded(
                    client, folder, canonical, raw)
                updates = dict(desired)
                updates.update({
                    "Resume URL": web_url,
                    "Resume Folder Path": _resume_display_path(
                        client.resumes_folder, subpath),
                })
                if sheet == "CandidateList":
                    client.update_row(row["index"], updates, current_values=vals)
                else:
                    client.update_rejected_row(row["index"], updates, current_values=vals)
                removed = _clean_old_files(client, app_id, folder, canonical, month)
                result.update({
                    "filename": canonical, "verified_bytes": byte_count,
                    "verified_text_chars": text_chars,
                    "removed_stale_files": removed,
                })
            results.append(result)
        except Exception as exc:
            errors += 1
            results.append({"application_id": app_id, "status": "ERROR", "error": str(exc)})
            print(f"{app_id}: ERROR {exc}", flush=True)
    if apply:
        if str(os.getenv("HIRING_REPAIR_RESUME_LINK_FORMULA", "") or "").strip().lower() in {"1", "true", "yes", "on"}:
            client.set_calculated_column(client.table, "Resume Link", _RESUME_LINK_FORMULA)
            rejected_table = client._rejected_table_name_if_exists()
            if rejected_table:
                client.set_calculated_column(rejected_table, "Resume Link", _RESUME_LINK_FORMULA)
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    mode = "apply" if apply else "dry_run"
    report = REPORT_DIR / f"batch_{batch}_score_reconciliation_{mode}.json"
    report.write_text(json.dumps({
        "batch": batch, "mode": mode, "jd_role_count": len(roles),
        "jd_built_at": built_at, "errors": errors, "results": results,
    }, indent=2, ensure_ascii=False), encoding="utf-8")
    changed_count = sum(item.get("status") == "CHANGE" for item in results)
    print(f"REPORT={report}")
    print(f"SUMMARY changed={changed_count} errors={errors}")
    return 1 if errors else 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--batch", type=int, choices=range(1, 6), required=True)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--apply", action="store_true")
    args = parser.parse_args()
    raise SystemExit(run(args.batch, args.manifest, args.apply))
