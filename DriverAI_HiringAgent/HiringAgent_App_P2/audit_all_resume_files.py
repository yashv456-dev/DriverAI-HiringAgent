"""Read-only audit of SharePoint rows vs resume files.

Checks every nonblank row in CandidateList and Rejected against the SharePoint
resume folders. This script never calls update/add/delete/move/rename/upload or
mail functions; it only lists workbook rows and drive folder children.
"""
from __future__ import annotations

import csv
import datetime as dt
import os
import re
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

from hiring_agent.config import (
    STATUS_NEEDS_REVIEW,
    STATUS_PROCESSING_FAILED,
    STATUS_REJECTED,
    STATUS_SCORED,
    dated_subpath,
)
from hiring_agent.sharepoint_scoring import (
    _clean_category_for_filename,
    _get_cleaned_filename_prefix,
    _rejected_subpath,
    _resume_filename_tail,
)
from sharepoint_client import SharePointClient, SharePointError


PENDING_STATUSES = {"", "New Email Received", STATUS_NEEDS_REVIEW}
REJECTED_STATUSES = {STATUS_REJECTED, STATUS_PROCESSING_FAILED}


def load_env() -> None:
    env_path = Path(__file__).with_name(".env")
    if not env_path.exists():
        return
    for raw in env_path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip())


def is_gap(value: object) -> bool:
    return str(value or "").strip().lower() in {
        "",
        "n/a",
        "na",
        "none",
        "not extracted",
        "not found",
        "not provided",
        "-",
    }


def parse_received_date(value: object) -> dt.datetime:
    text = str(value or "").strip()
    if text.endswith("Z"):
        text = text[:-1] + "+00:00"
    try:
        return dt.datetime.fromisoformat(text)
    except ValueError:
        return dt.datetime.today()


def filename_from_url(url: str) -> str:
    if not url:
        return ""
    parsed = urlparse(url)
    query_file = (parse_qs(parsed.query).get("file") or [""])[0]
    if query_file:
        return query_file
    path = unquote(parsed.path or "")
    return path.rsplit("/", 1)[-1] if path else ""


def clean_originals(value: object) -> list[str]:
    out = []
    for item in str(value or "").split(","):
        name = item.strip()
        if name:
            out.append(name)
    return out


def extension_for(vals: dict, original: str) -> str:
    ext = os.path.splitext(original)[1]
    if ext:
        return ext
    url_ext = os.path.splitext(filename_from_url(str(vals.get("Resume URL", "") or "")))[1]
    if url_ext:
        return url_ext
    m = re.search(r"\.(pdf|docx)\b", str(vals.get("Original Filename", "") or ""), re.I)
    return f".{m.group(1)}" if m else ""


def canonical_name(vals: dict, original: str) -> str:
    app_id = str(vals.get("Application ID", "") or "").strip()
    ext = extension_for(vals, original)
    name = _get_cleaned_filename_prefix(str(vals.get("Full Name", "") or ""))
    category = _clean_category_for_filename(str(vals.get("Category", "") or ""))
    tail = _resume_filename_tail(app_id)
    return f"{name}_{category}_{tail}{ext}"


def p1_legacy_name(vals: dict, original: str) -> str:
    app_id = str(vals.get("Application ID", "") or "").strip().replace("/", "-")
    ext = extension_for(vals, original)
    name = _get_cleaned_filename_prefix(str(vals.get("Full Name", "") or ""))
    return f"{name}_{app_id}{ext}"


def old_appid_original_name(vals: dict, original: str) -> str:
    app_id = str(vals.get("Application ID", "") or "").strip().replace("/", "-")
    return original if app_id and app_id in original else f"{app_id}_{original}"


def display_path_to_drive_path(client: SharePointClient, display_path: str, fallback_subpath: str) -> str:
    display = str(display_path or "").strip().strip("/")
    if display:
        parts = display.split("/")
        if len(parts) > 1:
            return f"{client.resumes_folder.rstrip('/')}/{'/'.join(parts[1:])}"
    return f"{client.resumes_folder.rstrip('/')}/{fallback_subpath.strip('/')}"


def folder_children(client: SharePointClient, folder: str, cache: dict) -> tuple[bool, set[str], str]:
    folder = folder.strip("/")
    if folder in cache:
        return cache[folder]
    try:
        names = {item["name"] for item in client.list_folder_children(folder)}
        cache[folder] = (True, names, "")
    except SharePointError as exc:
        cache[folder] = (False, set(), str(exc))
    return cache[folder]


def row_subpath(vals: dict, sheet: str) -> str:
    subpath = dated_subpath(parse_received_date(vals.get("Received Date", "")))
    if sheet == "Rejected":
        return _rejected_subpath(subpath)
    return subpath


def expected_names(vals: dict) -> list[tuple[str, str]]:
    names: list[tuple[str, str]] = []
    url_name = filename_from_url(str(vals.get("Resume URL", "") or ""))
    if url_name:
        names.append(("resume_url", url_name))
    for original in clean_originals(vals.get("Original Filename", "")):
        status = str(vals.get("Status", "") or "").strip()
        if status not in PENDING_STATUSES:
            names.append(("canonical", canonical_name(vals, original)))
        names.append(("p1_current_write", p1_legacy_name(vals, original)))
        names.append(("old_appid_original", old_appid_original_name(vals, original)))
        names.append(("original", original))
    deduped = []
    seen = set()
    for kind, name in names:
        key = name.lower()
        if key not in seen:
            deduped.append((kind, name))
            seen.add(key)
    return deduped


def row_empty(vals: dict) -> bool:
    return not any(str(v or "").strip() for v in vals.values())


def audit_row(client: SharePointClient, sheet: str, row: dict, folder_cache: dict, app_counts: dict) -> dict:
    vals = row["values"]
    app_id = str(vals.get("Application ID", "") or "").strip()
    status = str(vals.get("Status", "") or "").strip()
    subpath = row_subpath(vals, sheet)
    expected_folder = f"{client.resumes_folder.rstrip('/')}/{subpath}"
    checked_folder = display_path_to_drive_path(client, str(vals.get("Resume Folder Path", "")), subpath)
    folder_ok, folder_names, folder_error = folder_children(client, checked_folder, folder_cache)
    lower_to_actual = {name.lower(): name for name in folder_names}

    candidates = expected_names(vals)
    matches = []
    for kind, name in candidates:
        actual = lower_to_actual.get(name.lower())
        if actual:
            matches.append((kind, actual))

    found_file = matches[0][1] if matches else ""
    found_by = matches[0][0] if matches else ""
    all_matched = "; ".join(f"{kind}:{name}" for kind, name in matches)
    stale_or_extra = "; ".join(f"{kind}:{name}" for kind, name in matches[1:])
    canonical_expected = ""
    originals = clean_originals(vals.get("Original Filename", ""))
    if originals:
        canonical_expected = canonical_name(vals, originals[0])

    notes = []
    severity = "PASS"
    if app_counts.get(app_id, 0) > 1:
        severity = "FAIL"
        notes.append(f"duplicate Application ID appears {app_counts.get(app_id)} times")
    if not folder_ok:
        severity = "FAIL"
        notes.append(f"folder read failed: {folder_error}")
    if checked_folder != expected_folder:
        severity = "FAIL"
        notes.append(f"Resume Folder Path points to {checked_folder}, expected {expected_folder}")
    if sheet == "Main" and status in REJECTED_STATUSES:
        severity = "FAIL"
        notes.append("rejected status still on main sheet")
    if sheet == "Rejected" and status == STATUS_SCORED:
        severity = "FAIL"
        notes.append("scored status still on Rejected sheet")
    if sheet == "Rejected" and status not in REJECTED_STATUSES:
        severity = "WARN" if severity == "PASS" else severity
        notes.append(f"unexpected Rejected-sheet status {status!r}")
    found_is_canonical = bool(
        found_file and canonical_expected and found_file.lower() == canonical_expected.lower()
    )
    if sheet == "Main" and status == STATUS_SCORED and not found_is_canonical:
        if found_by == "resume_url":
            notes.append("valid legacy filename found through authoritative Resume URL")
        else:
            severity = "FAIL"
            notes.append(f"scored main row not found by canonical name; found_by={found_by or 'none'}")
    if sheet == "Rejected" and status == STATUS_REJECTED and not found_is_canonical:
        if found_by == "resume_url":
            notes.append("valid legacy filename found through authoritative Resume URL")
        else:
            severity = "FAIL"
            notes.append(f"rejected row not found by canonical/url name; found_by={found_by or 'none'}")
    if not found_file and not is_gap(vals.get("Original Filename")):
        severity = "FAIL"
        notes.append("no expected resume file found")
    if found_file and filename_from_url(str(vals.get("Resume URL", "") or "")):
        url_name = filename_from_url(str(vals.get("Resume URL", "") or ""))
        if url_name.lower() != found_file.lower():
            severity = "FAIL"
            notes.append(f"Resume URL filename {url_name} does not match found file {found_file}")
    if status in PENDING_STATUSES and found_by:
        severity = "PENDING" if severity == "PASS" else severity
        notes.append("pending row; P1-current-write filename is acceptable until P2 scores it")
    elif stale_or_extra and len(matches) > max(1, len(clean_originals(vals.get("Original Filename", "")))):
        severity = "WARN" if severity == "PASS" else severity
        notes.append(f"multiple candidate filenames exist: {stale_or_extra}")

    return {
        "sheet": sheet,
        "row_index": row.get("index", ""),
        "application_id": app_id,
        "status": status,
        "email": vals.get("Email", ""),
        "received_date": vals.get("Received Date", ""),
        "last_updated_date": vals.get("Last Updated Date", ""),
        "full_name": vals.get("Full Name", ""),
        "category": vals.get("Category", ""),
        "country": vals.get("Country", ""),
        "location": vals.get("Location", ""),
        "original_filename": vals.get("Original Filename", ""),
        "resume_url": vals.get("Resume URL", ""),
        "resume_folder_path": vals.get("Resume Folder Path", ""),
        "expected_folder": expected_folder,
        "checked_folder": checked_folder,
        "folder_exists": folder_ok,
        "canonical_expected": canonical_expected,
        "found_file": found_file,
        "found_by": found_by,
        "all_matches": all_matched,
        "candidate_names_checked": "; ".join(name for _kind, name in candidates),
        "severity": severity,
        "notes": "; ".join(notes),
    }


def write_reports(rows: list[dict], out_dir: Path) -> tuple[Path, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = out_dir / f"sharepoint_resume_file_audit_{stamp}.csv"
    md_path = out_dir / f"sharepoint_resume_file_audit_{stamp}.md"
    fields = [
        "sheet", "row_index", "application_id", "status", "email", "received_date",
        "last_updated_date", "full_name", "category", "country", "location",
        "original_filename", "resume_url", "resume_folder_path", "expected_folder",
        "checked_folder", "folder_exists", "canonical_expected", "found_file",
        "found_by", "all_matches", "candidate_names_checked", "severity", "notes",
    ]
    with csv_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for row in rows:
            writer.writerow({k: row.get(k, "") for k in fields})

    counts = {key: sum(1 for r in rows if r["severity"] == key) for key in ("PASS", "PENDING", "WARN", "FAIL")}
    main = sum(1 for r in rows if r["sheet"] == "Main")
    rejected = sum(1 for r in rows if r["sheet"] == "Rejected")
    with md_path.open("w", encoding="utf-8") as f:
        f.write("# SharePoint Resume File Audit\n\n")
        f.write(f"Generated: {dt.datetime.now().isoformat(timespec='seconds')}\n\n")
        f.write("## Summary\n\n")
        f.write(f"- rows checked: {len(rows)}\n")
        f.write(f"- main rows: {main}\n")
        f.write(f"- rejected rows: {rejected}\n")
        for key in ("PASS", "PENDING", "WARN", "FAIL"):
            f.write(f"- {key.lower()}: {counts[key]}\n")
        f.write("\n")
        for title, sev in (("Failures", "FAIL"), ("Warnings", "WARN"), ("Pending Rows", "PENDING")):
            subset = [r for r in rows if r["severity"] == sev]
            if not subset:
                continue
            f.write(f"## {title}\n\n")
            for r in subset:
                f.write(
                    f"- {r['application_id']} [{r['sheet']} row {r['row_index']}] "
                    f"{r['status']} | found={r['found_file'] or '(none)'} | {r['notes']}\n"
                )
            f.write("\n")
        f.write("## Detail\n\n")
        f.write("| Sheet | Row | Application ID | Status | Expected | Found | Severity | Notes |\n")
        f.write("|---|---:|---|---|---|---|---|---|\n")
        for r in rows:
            cells = {
                k: str(r.get(k, "")).replace("|", "\\|")
                for k in ("sheet", "row_index", "application_id", "status", "canonical_expected", "found_file", "severity", "notes")
            }
            f.write(
                f"| {cells['sheet']} | {cells['row_index']} | {cells['application_id']} | "
                f"{cells['status']} | {cells['canonical_expected']} | {cells['found_file']} | "
                f"{cells['severity']} | {cells['notes']} |\n"
            )
    return md_path, csv_path


def main() -> None:
    load_env()
    client = SharePointClient()
    main_rows = [
        r for r in client.list_rows()
        if not row_empty(r["values"]) and str(r["values"].get("Application ID", "") or "").strip()
    ]
    rejected_rows = [
        r for r in client.list_rejected_rows()
        if not row_empty(r["values"]) and str(r["values"].get("Application ID", "") or "").strip()
    ]

    app_counts = {}
    for r in main_rows + rejected_rows:
        app_id = str(r["values"].get("Application ID", "") or "").strip()
        if app_id:
            app_counts[app_id] = app_counts.get(app_id, 0) + 1

    folder_cache = {}
    detail = []
    for row in main_rows:
        detail.append(audit_row(client, "Main", row, folder_cache, app_counts))
    for row in rejected_rows:
        detail.append(audit_row(client, "Rejected", row, folder_cache, app_counts))

    md_path, csv_path = write_reports(detail, Path("P2_Logs") / "audits")
    counts = {key: sum(1 for r in detail if r["severity"] == key) for key in ("PASS", "PENDING", "WARN", "FAIL")}
    print(f"AUDIT_MD={md_path}")
    print(f"AUDIT_CSV={csv_path}")
    print(f"ROWS={len(detail)} MAIN={sum(1 for r in detail if r['sheet'] == 'Main')} REJECTED={sum(1 for r in detail if r['sheet'] == 'Rejected')}")
    print(f"PASS={counts['PASS']} PENDING={counts['PENDING']} WARN={counts['WARN']} FAIL={counts['FAIL']}")
    for row in detail:
        if row["severity"] == "FAIL":
            print(f"FAIL {row['application_id']} [{row['sheet']}:{row['row_index']}] {row['notes']}")


if __name__ == "__main__":
    main()
