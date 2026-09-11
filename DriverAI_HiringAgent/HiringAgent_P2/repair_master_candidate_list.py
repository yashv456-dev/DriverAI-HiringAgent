"""Repair narrowly-scoped data-quality defects in the live CandidateList table.

Default is dry-run. Pass --apply to write. This command never sends mail, moves a
candidate between sheets, changes Status/Category/scoring, or deletes a non-empty row.
"""
from __future__ import annotations

import argparse
from dotenv import load_dotenv

load_dotenv(".env")

from sharepoint_client import SharePointClient
from hiring_agent.extraction import (
    education_needs_repair, normalize_education, extract_candidate_details,
    ai_recheck_fields, html_to_text, merge_mail_body_fallback,
)
from hiring_agent.geo import is_foreign_location, normalize_country
import hiring_agent.config as _cfg
from hiring_agent.excel_output import _parse_received
from hiring_agent.sharepoint_scoring import (
    _download_resume_text, _cleanup_temp_files,
)


def _raw_rows(client: SharePointClient) -> tuple[list[str], list[dict]]:
    cols = client.table_columns()
    raw = client._rows_paged(f"{client._wb_base()}/tables/{client.table}/rows")
    rows = []
    for item in raw:
        cells = (item.get("values") or [[]])[0]
        values = {cols[i]: (cells[i] if i < len(cells) else "") for i in range(len(cols))}
        rows.append({"index": item.get("index"), "values": values})
    return cols, rows


def _is_empty(values: dict) -> bool:
    return not any(str(v or "").strip() for v in values.values())


def _geo_conflict(values: dict) -> bool:
    location = str(values.get("Location", "") or "").strip()
    country = normalize_country(values.get("Country", ""))
    return bool(location and is_foreign_location(location) and country == "United States")


def main(apply: bool = False, offset: int = 0, limit: int = 0,
         app_ids: list[str] | None = None) -> int:
    client = SharePointClient()
    _, rows = _raw_rows(client)
    blank_indices = [r["index"] for r in rows if _is_empty(r["values"])]
    real_rows = [r for r in rows if not _is_empty(r["values"])]
    print(f"Mode: {'APPLY' if apply else 'DRY-RUN'}")
    print(f"Rows: {len(real_rows)} candidates, {len(blank_indices)} fully blank")
    planned: list[tuple[int, str, dict]] = []
    errors = 0

    flagged_rows = [r for r in real_rows if (
        education_needs_repair(r["values"].get("Education", ""))
        or _geo_conflict(r["values"])
    )]
    wanted = {str(v).strip() for v in (app_ids or []) if str(v).strip()}
    if wanted:
        selected = [r for r in real_rows
                    if str(r["values"].get("Application ID", "")).strip() in wanted]
    else:
        selected = flagged_rows[offset:offset + limit] if limit else flagged_rows[offset:]
    selected_indices = {r["index"] for r in selected}
    print(f"Flagged: {len(flagged_rows)}; selected batch: offset={offset}, count={len(selected)}")

    for row in real_rows:
        index, values = row["index"], row["values"]
        app_id = str(values.get("Application ID", "") or "").strip()
        force_selected = bool(wanted and index in selected_indices)
        bad_education = education_needs_repair(values.get("Education", "")) or force_selected
        bad_geo = (_geo_conflict(values) or
                   (force_selected and not str(values.get("Location", "") or "").strip()))
        updates: dict = {}
        if index in selected_indices and not str(values.get("Retry Count", "") or "").strip():
            updates["Retry Count"] = 0

        if index in selected_indices and (bad_education or bad_geo):
            try:
                subpath = _cfg.dated_subpath(_parse_received(values.get("Received Date")))
                text, _used, _raw = _download_resume_text(
                    client, app_id, values.get("Original Filename", ""), subpath,
                    full_name=values.get("Full Name"), category=values.get("Category"),
                )
                if not text:
                    raise RuntimeError("stored resume could not be read")
                derived = extract_candidate_details(text)
                mail_body = html_to_text(str(values.get("Mail Body", "") or ""))
                derived = merge_mail_body_fallback(derived, mail_body)
                derived = ai_recheck_fields(derived, text, mail_body)
                if bad_education:
                    education = normalize_education(derived.get("Education", ""))
                    if education and not education_needs_repair(education):
                        if education != str(values.get("Education", "") or "").strip():
                            updates["Education"] = education
                if bad_geo:
                    location = str(derived.get("Location", "") or "").strip()
                    country = normalize_country(derived.get("Country", ""))
                    # Only publish a replacement when the fresh pair is internally consistent.
                    if location and country and not (
                        is_foreign_location(location) and country == "United States"
                    ):
                        if location != str(values.get("Location", "") or "").strip():
                            updates["Location"] = location
                        if country != normalize_country(values.get("Country", "")):
                            updates["Country"] = country
            except Exception as exc:
                errors += 1
                print(f"ERROR {app_id}: {exc}")
            finally:
                _cleanup_temp_files()

        if updates:
            planned.append((index, app_id, updates))
            summary = "; ".join(
                f"{key}: {values.get(key, '')!r} -> {value!r}" for key, value in updates.items()
            )
            print(f"ROW {index} {app_id}: {summary}")

    print(f"Planned candidate updates: {len(planned)}")
    print(f"Planned blank-row deletes: {blank_indices}")
    print(f"Derivation errors: {errors}")

    if apply:
        for index, _app_id, updates in planned:
            original = next(r["values"] for r in real_rows if r["index"] == index)
            client.update_row(index, updates, current_values=original)
        # Positional row indices shift after each delete, so always delete descending.
        # Blank rows are cleaned only on the final/unbounded invocation; batches do not
        # shift positional indices while later batches are still pending.
        if not limit:
            for index in sorted(blank_indices, reverse=True):
                client.delete_row(index)
        print("APPLIED successfully; no messages were sent and no candidate was reclassified.")
    return 1 if errors else 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--apply", action="store_true", help="write the previewed repairs")
    parser.add_argument("--offset", type=int, default=0, help="start within flagged rows")
    parser.add_argument("--limit", type=int, default=0, help="maximum flagged rows; 0 means all")
    parser.add_argument("--app-id", action="append", default=[], help="repair one Application ID")
    args = parser.parse_args()
    raise SystemExit(main(apply=args.apply, offset=args.offset, limit=args.limit,
                          app_ids=args.app_id))
