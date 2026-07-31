"""Read-only live audit joining SharePoint rows, Sent Items, and resume files.

This script does not update SharePoint, move/rename/delete/upload files, or send mail.
It only reads:
  * CandidateList and Rejected workbook tables
  * Sent Items for the configured SENDER_MAILBOX
  * SharePoint resume folders

Reports are written locally under P2_Logs/audits.
"""
from __future__ import annotations

import csv
import datetime as dt
import os
import re
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

from hiring_agent.config import (
    STATUS_LOCATION_REVIEW,
    STATUS_LOCATION_UNCONFIRMED,
    STATUS_NEEDS_REVIEW,
    STATUS_PROCESSING_FAILED,
    STATUS_REJECTED,
    STATUS_SCORED,
    dated_subpath,
)
from hiring_agent.geo import GeoDecision, classify_location_usa
from hiring_agent.sharepoint_scoring import (
    _bare_email,
    _clean_category_for_filename,
    _get_cleaned_filename_prefix,
    _missing_fields_list,
    _normalize_phone_key,
    _rejected_subpath,
    _resume_filename_tail,
)
from sharepoint_client import GRAPH, SharePointClient, SharePointError


APP_RE = re.compile(r"APP-\d{8}-\d{4}-[A-F0-9]{4}", re.I)
AUDIT_DIR = Path("P2_Logs") / "audits"


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


def row_empty(vals: dict) -> bool:
    return not any(str(v or "").strip() for v in vals.values())


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


def parse_date(value: object) -> dt.datetime:
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
    out: list[str] = []
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


def expected_names(vals: dict) -> list[tuple[str, str]]:
    names: list[tuple[str, str]] = []
    url_name = filename_from_url(str(vals.get("Resume URL", "") or ""))
    if url_name:
        names.append(("resume_url", url_name))
    status = str(vals.get("Status", "") or "").strip()
    for original in clean_originals(vals.get("Original Filename", "")):
        if status not in {"", "New Email Received", "Needs Review - Unreadable Resume"}:
            names.append(("canonical", canonical_name(vals, original)))
        names.append(("p1_legacy", p1_legacy_name(vals, original)))
        names.append(("old_appid_original", old_appid_original_name(vals, original)))
        names.append(("original", original))
    deduped: list[tuple[str, str]] = []
    seen = set()
    for kind, name in names:
        key = name.lower()
        if key not in seen:
            deduped.append((kind, name))
            seen.add(key)
    return deduped


def row_subpath(vals: dict, sheet: str) -> str:
    subpath = dated_subpath(parse_date(vals.get("Received Date", "")))
    if sheet == "Rejected":
        return _rejected_subpath(subpath)
    return subpath


def display_path_to_drive_path(client: SharePointClient, display_path: str, fallback_subpath: str) -> str:
    display = str(display_path or "").strip().strip("/")
    if display:
        parts = display.split("/")
        if len(parts) > 1:
            return f"{client.resumes_folder.rstrip('/')}/{'/'.join(parts[1:])}"
    return f"{client.resumes_folder.rstrip('/')}/{fallback_subpath.strip('/')}"


def list_sent_items(client: SharePointClient) -> list[dict]:
    mailbox = client.sender_mailbox
    url = (
        f"{GRAPH}/users/{mailbox}/mailFolders/sentitems/messages"
        "?$select=subject,sentDateTime,toRecipients,bodyPreview"
        "&$top=100&$orderby=sentDateTime desc"
    )
    out: list[dict] = []
    while url:
        data = client._req("GET", url).json()
        for msg in data.get("value", []):
            recipients = []
            for rec in msg.get("toRecipients", []) or []:
                addr = ((rec.get("emailAddress") or {}).get("address") or "").strip().lower()
                if addr:
                    recipients.append(addr)
            text = f"{msg.get('subject', '')} {msg.get('bodyPreview', '')}"
            refs = {m.group(0).upper() for m in APP_RE.finditer(text)}
            out.append({
                "subject": msg.get("subject", "") or "",
                "sent": msg.get("sentDateTime", "") or "",
                "to": recipients,
                "refs": refs,
            })
        url = data.get("@odata.nextLink")
    return out


def sent_index(messages: list[dict]) -> dict:
    by_ref: dict[str, list[dict]] = {}
    by_recipient_subject: dict[tuple[str, str], list[dict]] = {}
    for msg in messages:
        for ref in msg["refs"]:
            by_ref.setdefault(ref, []).append(msg)
        subject = msg["subject"].strip().lower()
        for to in msg["to"]:
            by_recipient_subject.setdefault((to, subject), []).append(msg)
    return {"by_ref": by_ref, "by_recipient_subject": by_recipient_subject}


def list_all_files(client: SharePointClient, root: str) -> list[dict]:
    files: list[dict] = []
    stack = [root.strip("/")]
    seen = set()
    while stack:
        folder = stack.pop()
        if folder.lower() in seen:
            continue
        seen.add(folder.lower())
        try:
            children = client.list_folder_children(folder)
        except SharePointError as exc:
            files.append({
                "folder": folder,
                "name": "",
                "path": folder,
                "is_folder": False,
                "read_error": str(exc),
            })
            continue
        for child in children:
            name = child.get("name", "")
            path = f"{folder.rstrip('/')}/{name}".strip("/")
            if child.get("is_folder"):
                stack.append(path)
            else:
                files.append({
                    "folder": folder,
                    "name": name,
                    "path": path,
                    "is_folder": False,
                    "read_error": "",
                })
    return files


def read_rows(client: SharePointClient) -> list[dict]:
    records: list[dict] = []
    for sheet, rows in (("Main", client.list_rows()), ("Rejected", client.list_rejected_rows())):
        for row in rows:
            vals = row["values"]
            if row_empty(vals):
                continue
            app_id = str(vals.get("Application ID", "") or "").strip()
            if not app_id:
                continue
            records.append({"sheet": sheet, "row_index": row.get("index", ""), "values": vals})
    return records


def find_resume(client: SharePointClient, rec: dict, folder_files: dict[str, set[str]]) -> tuple[str, str, str, str]:
    vals = rec["values"]
    subpath = row_subpath(vals, rec["sheet"])
    expected_folder = f"{client.resumes_folder.rstrip('/')}/{subpath}"
    folder = display_path_to_drive_path(client, str(vals.get("Resume Folder Path", "")), subpath)
    names = folder_files.get(folder.strip("/").lower(), set())
    lower_to_actual = {n.lower(): n for n in names}
    matches = []
    for kind, name in expected_names(vals):
        actual = lower_to_actual.get(name.lower())
        if actual:
            matches.append((kind, actual))
    if matches:
        return folder, expected_folder, matches[0][0], matches[0][1]
    return folder, expected_folder, "", ""


def classify_row(rec: dict, mail_idx: dict, folder_info: tuple[str, str, str, str]) -> dict:
    vals = rec["values"]
    app_id = str(vals.get("Application ID", "") or "").strip()
    status = str(vals.get("Status", "") or "").strip()
    email = _bare_email(vals.get("Email", "")).lower()
    info_marker = str(vals.get("Info Request Sent", "") or "").strip()
    decline_marker = str(vals.get("Decline Sent", "") or "").strip()
    mail_marker = str(vals.get("Mail Sent", "") or "").strip()
    missing = _missing_fields_list(vals) if rec["sheet"] == "Main" and status == STATUS_SCORED else []
    folder, expected_folder, found_by, found_file = folder_info
    received_at = parse_date(vals.get("Received Date", ""))
    updated_at = parse_date(vals.get("Last Updated Date", ""))

    ref_messages = mail_idx["by_ref"].get(app_id.upper(), [])
    p1_sent_match = [
        m for m in ref_messages
        if email in m["to"] and not m["subject"].startswith("Quick follow-up on your DriverAI application")
    ]
    p2_info_match = [
        m for m in ref_messages
        if email in m["to"] and m["subject"].startswith("Quick follow-up on your DriverAI application")
    ]
    p2_decline_match = mail_idx["by_recipient_subject"].get(
        (email, "update on your driverai application"), []
    )

    notes: list[str] = []
    severity = "PASS"
    if folder.strip("/").lower() != expected_folder.strip("/").lower():
        severity = "FAIL"
        notes.append("Resume Folder Path does not match expected dated/rejected folder")
    if not found_file and not is_gap(vals.get("Original Filename")):
        severity = "FAIL"
        notes.append("no expected resume file found")
    if vals.get("Received Date") and vals.get("Last Updated Date") and updated_at < received_at:
        severity = "FAIL"
        notes.append("Last Updated Date is earlier than Received Date")
    if rec["sheet"] == "Main":
        if status in {"New Email Received", STATUS_NEEDS_REVIEW}:
            severity = "PENDING"
            notes.append(f"pending P1/P2 queue row ({status})")
        elif status == STATUS_LOCATION_REVIEW:
            decision, reason = classify_location_usa(
                vals.get("Location", ""), vals.get("Country", ""),
                phone=vals.get("Phone", ""), education=vals.get("Education", ""))
            if decision == GeoDecision.CONFIRMED_NON_US:
                severity = "FAIL"
                notes.append(f"location-review row now confirms non-US: {reason}")
            elif decision == GeoDecision.CONFIRMED_US:
                severity = "WARN"
                notes.append(f"location-review row now confirms US and can be scored: {reason}")
            if not info_marker:
                severity = "FAIL"
                notes.append("location-review row has no follow-up/manual-review marker")
        elif status != STATUS_SCORED:
            severity = "FAIL"
            notes.append(f"main row is not Scored ({status})")
        decision, reason = classify_location_usa(
            vals.get("Location", ""), vals.get("Country", ""),
            phone=vals.get("Phone", ""), education=vals.get("Education", ""))
        if status == STATUS_SCORED and decision != GeoDecision.CONFIRMED_US:
            severity = "FAIL"
            notes.append(f"main scored row is {decision.value}, not confirmed US: {reason}")
        if missing and info_marker.startswith("Sent ") and not p2_info_match:
            severity = "FAIL"
            notes.append("Info Request Sent marker has no matching Sent Items message by ref+recipient")
        if missing and not (info_marker.startswith("Sent ") or info_marker == "No valid email"):
            severity = "FAIL"
            notes.append("missing fields without a final info-request marker")
        if status == STATUS_SCORED and not missing and info_marker not in {"Nothing missing", ""} and not info_marker.startswith("Sent "):
            severity = "WARN" if severity == "PASS" else severity
            notes.append(f"expected Info Request Sent='Nothing missing', got {info_marker!r}")
    else:
        if status not in {STATUS_REJECTED, STATUS_LOCATION_UNCONFIRMED, STATUS_PROCESSING_FAILED}:
            severity = "FAIL"
            notes.append(f"rejected row has unexpected status {status!r}")
        if status in {STATUS_REJECTED, STATUS_LOCATION_UNCONFIRMED} and decline_marker.startswith("Sent ") and not p2_decline_match:
            severity = "FAIL"
            notes.append("Decline Sent marker has no matching Sent Items decline by recipient")
        if status in {STATUS_REJECTED, STATUS_LOCATION_UNCONFIRMED} and not (
            decline_marker.startswith("Sent ")
            or decline_marker.startswith("No valid email")
            or decline_marker.startswith("N/A -")
            or "Duplicate" in decline_marker
        ):
            severity = "FAIL"
            notes.append(f"bad Decline Sent marker {decline_marker!r}")
    if mail_marker.startswith("Sent ") and not p1_sent_match:
        severity = "WARN" if severity == "PASS" else severity
        notes.append("Mail Sent marker has no matching P1 Sent Items message by ref+recipient")

    return {
        "severity": severity,
        "sheet": rec["sheet"],
        "row_index": rec["row_index"],
        "application_id": app_id,
        "status": status,
        "full_name": vals.get("Full Name", ""),
        "email": vals.get("Email", ""),
        "phone": vals.get("Phone", ""),
        "location": vals.get("Location", ""),
        "country": vals.get("Country", ""),
        "category": vals.get("Category", ""),
        "received_date": vals.get("Received Date", ""),
        "last_updated_date": vals.get("Last Updated Date", ""),
        "mail_sent": mail_marker,
        "decline_sent": decline_marker,
        "info_request_sent": info_marker,
        "missing_fields": ", ".join(missing),
        "p1_sent_matches": len(p1_sent_match),
        "p2_info_matches": len(p2_info_match),
        "p2_decline_matches": len(p2_decline_match),
        "resume_folder": folder,
        "expected_folder": expected_folder,
        "found_by": found_by,
        "found_file": found_file,
        "notes": "; ".join(notes),
    }


def duplicate_issues(records: list[dict]) -> list[dict]:
    buckets = {"app": {}, "email": {}, "phone": {}}
    for rec in records:
        vals = rec["values"]
        app_id = str(vals.get("Application ID", "") or "").strip()
        email = _bare_email(vals.get("Email", "")).lower()
        phone = _normalize_phone_key(vals.get("Phone", ""))
        for key, value in (("app", app_id), ("email", email), ("phone", phone)):
            if value:
                buckets[key].setdefault(value, []).append(rec)
    issues = []
    for key, values in buckets.items():
        for value, rows in values.items():
            if len(rows) <= 1:
                continue
            issues.append({
                "kind": f"duplicate_{key}",
                "value": value,
                "count": len(rows),
                "application_ids": ", ".join(str(r["values"].get("Application ID", "")) for r in rows),
                "sheets": ", ".join(r["sheet"] for r in rows),
                "row_indexes": ", ".join(str(r["row_index"]) for r in rows),
                "emails": ", ".join(str(r["values"].get("Email", "")) for r in rows),
                "names": ", ".join(str(r["values"].get("Full Name", "")) for r in rows),
            })
    return issues


def write_reports(rows: list[dict], duplicates: list[dict], extras: list[dict],
                  sent_count: int, out_dir: Path) -> tuple[Path, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = out_dir / f"live_integrity_audit_{stamp}.csv"
    md_path = out_dir / f"live_integrity_audit_{stamp}.md"
    fields = [
        "severity", "sheet", "row_index", "application_id", "status", "full_name",
        "email", "phone", "location", "country", "category", "received_date",
        "last_updated_date", "mail_sent", "decline_sent", "info_request_sent",
        "missing_fields", "p1_sent_matches", "p2_info_matches", "p2_decline_matches",
        "resume_folder", "expected_folder", "found_by", "found_file", "notes",
    ]
    with csv_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for row in rows:
            writer.writerow({k: row.get(k, "") for k in fields})

    counts = {sev: sum(1 for r in rows if r["severity"] == sev) for sev in ("PASS", "PENDING", "WARN", "FAIL")}
    main = sum(1 for r in rows if r["sheet"] == "Main")
    rejected = sum(1 for r in rows if r["sheet"] == "Rejected")
    with md_path.open("w", encoding="utf-8") as f:
        f.write("# Live SharePoint/Mail/Resume Integrity Audit\n\n")
        f.write(f"Generated: {dt.datetime.now().isoformat(timespec='seconds')}\n\n")
        f.write("## Summary\n\n")
        f.write(f"- rows checked: {len(rows)}\n")
        f.write(f"- main rows: {main}\n")
        f.write(f"- rejected rows: {rejected}\n")
        f.write(f"- sent items read: {sent_count}\n")
        f.write(f"- pass: {counts['PASS']}\n")
        f.write(f"- pending: {counts['PENDING']}\n")
        f.write(f"- warn: {counts['WARN']}\n")
        f.write(f"- fail: {counts['FAIL']}\n")
        f.write(f"- duplicate groups: {len(duplicates)}\n")
        f.write(f"- unreferenced resume files: {len(extras)}\n\n")

        for title, items in (("Failures", [r for r in rows if r["severity"] == "FAIL"]),
                             ("Pending", [r for r in rows if r["severity"] == "PENDING"]),
                             ("Warnings", [r for r in rows if r["severity"] == "WARN"])):
            if not items:
                continue
            f.write(f"## {title}\n\n")
            for r in items:
                f.write(
                    f"- {r['application_id']} [{r['sheet']} row {r['row_index']}] "
                    f"{r['status']} | {r['email']} | {r['notes']}\n"
                )
            f.write("\n")

        if duplicates:
            f.write("## Duplicate Groups\n\n")
            for d in duplicates:
                f.write(
                    f"- {d['kind']} {d['value']} count={d['count']} "
                    f"apps={d['application_ids']} sheets={d['sheets']} rows={d['row_indexes']} "
                    f"emails={d['emails']} names={d['names']}\n"
                )
            f.write("\n")

        if extras:
            f.write("## Unreferenced Resume Files\n\n")
            for extra in extras:
                f.write(f"- {extra['path']}\n")
            f.write("\n")
    return md_path, csv_path


def main() -> None:
    load_env()
    client = SharePointClient()
    records = read_rows(client)
    sent = list_sent_items(client)
    mail_idx = sent_index(sent)

    all_files = list_all_files(client, client.resumes_folder)
    folder_files: dict[str, set[str]] = {}
    for item in all_files:
        if item.get("read_error"):
            continue
        folder_files.setdefault(str(item["folder"]).strip("/").lower(), set()).add(item["name"])

    row_reports = []
    referenced_paths = set()
    for rec in records:
        folder_info = find_resume(client, rec, folder_files)
        folder, _expected_folder, _found_by, found_file = folder_info
        if found_file:
            referenced_paths.add(f"{folder.strip('/')}/{found_file}".lower())
        row_reports.append(classify_row(rec, mail_idx, folder_info))

    ignore_names = {"hiringagent_p1_candidatelist.xlsx"}
    extras = []
    for item in all_files:
        path = str(item.get("path", "")).strip("/")
        name = str(item.get("name", "")).strip()
        if not name or name.lower() in ignore_names:
            continue
        if item.get("read_error"):
            extras.append({"path": f"{path} [READ ERROR: {item['read_error']}]"})
            continue
        if path.lower() not in referenced_paths:
            extras.append({"path": path})

    duplicates = duplicate_issues(records)
    md_path, csv_path = write_reports(row_reports, duplicates, extras, len(sent), AUDIT_DIR)

    counts = {sev: sum(1 for r in row_reports if r["severity"] == sev) for sev in ("PASS", "PENDING", "WARN", "FAIL")}
    print(f"AUDIT_MD={md_path}")
    print(f"AUDIT_CSV={csv_path}")
    print(f"ROWS={len(row_reports)} MAIN={sum(1 for r in row_reports if r['sheet'] == 'Main')} REJECTED={sum(1 for r in row_reports if r['sheet'] == 'Rejected')}")
    print(f"PASS={counts['PASS']} PENDING={counts['PENDING']} WARN={counts['WARN']} FAIL={counts['FAIL']}")
    print(f"SENT_ITEMS_READ={len(sent)}")
    print(f"DUPLICATE_GROUPS={len(duplicates)}")
    print(f"UNREFERENCED_RESUME_FILES={len(extras)}")
    for row in row_reports:
        if row["severity"] in {"FAIL", "WARN", "PENDING"}:
            print(f"{row['severity']} {row['application_id']} [{row['sheet']}:{row['row_index']}] {row['notes']}")
    for d in duplicates:
        print(f"DUP {d['kind']} {d['value']} apps={d['application_ids']} rows={d['row_indexes']}")
    for extra in extras[:25]:
        print(f"EXTRA {extra['path']}")
    if len(extras) > 25:
        print(f"EXTRA_MORE {len(extras) - 25}")


if __name__ == "__main__":
    main()
