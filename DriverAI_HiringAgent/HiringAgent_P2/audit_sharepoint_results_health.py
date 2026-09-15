"""Read-only health audit for live P2 SharePoint results.

Checks:
  * duplicate Application IDs, emails, and normalized phones across main/rejected
  * non-USA rows left on main after scoring
  * USA/scored rows parked on Rejected
  * no-content / unreadable rows and retry state
  * missing-info marker consistency for scored rows

No SharePoint writes and no mail sends.
"""
from __future__ import annotations

import csv
import datetime as dt
import os
import re
from pathlib import Path

from hiring_agent.config import (
    STATUS_LOCATION_REVIEW,
    STATUS_LOCATION_UNCONFIRMED,
    STATUS_NEEDS_REVIEW,
    STATUS_NO_MATCHING_ROLE,
    STATUS_PROCESSING_FAILED,
    STATUS_REJECTED,
    STATUS_SCORED,
)
from hiring_agent.geo import GeoDecision, classify_location_usa
from hiring_agent.sharepoint_scoring import _is_gap, _missing_fields_list, _normalize_phone_key
from sharepoint_client import SharePointClient


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


def norm_email(value: object) -> str:
    text = str(value or "").strip().lower()
    m = re.search(r"<([^>]+)>", text)
    return (m.group(1) if m else text).strip()


def row_empty(vals: dict) -> bool:
    return not any(str(v or "").strip() for v in vals.values())


def _is_sent_marker(marker: str) -> bool:
    """True when this marker records a real send attempt.

    'TEST-MODE (suppressed) Sent <ts>' is stamped when HIRING_SUPPRESS_EMAILS is on: the row
    WAS processed and must never be re-contacted, so for auditing purposes it counts exactly
    like 'Sent <ts>'. Matching only the 'Sent ' prefix (the behaviour before 2026-08-04) made
    every row processed in suppressed mode look unmarked.
    """
    return "Sent " in str(marker or "")


def _is_settled_marker(marker: str) -> bool:
    """True when a marker means 'this row needs no further mail' - sent, or deliberately not
    sent ('No valid email on file', 'N/A - aged out, no reply', 'Duplicate - ...')."""
    m = str(marker or "").strip()
    return bool(m) and (_is_sent_marker(m) or m.startswith("No valid email")
                        or m.startswith("N/A -") or "Duplicate" in m)


def row_record(sheet: str, row: dict) -> dict:
    vals = row["values"]
    rec = {
        "sheet": sheet,
        "row_index": row.get("index", ""),
        "application_id": str(vals.get("Application ID", "") or "").strip(),
        "status": str(vals.get("Status", "") or "").strip(),
        "email": vals.get("Email", ""),
        "email_key": norm_email(vals.get("Email", "")),
        "phone": vals.get("Phone", ""),
        "phone_key": _normalize_phone_key(vals.get("Phone", "")),
        "full_name": vals.get("Full Name", ""),
        "country": vals.get("Country", ""),
        "location": vals.get("Location", ""),
        "category": vals.get("Category", ""),
        "education": vals.get("Education", ""),
        "current_skills": vals.get("Current Skills", ""),
        "info_request_sent": vals.get("Info Request Sent", ""),
        "decline_sent": vals.get("Decline Sent", ""),
        "retry_count": vals.get("Retry Count", ""),
        "received_date": vals.get("Received Date", ""),
        "last_updated_date": vals.get("Last Updated Date", ""),
        "original_filename": vals.get("Original Filename", ""),
    }
    # Keep the exact SharePoint column names too; some shared P2 helpers expect
    # those keys rather than the normalized report keys above.
    # 'Email', the three Portfolio slots and the two date columns were MISSING from this
    # list until 2026-08-04. _missing_fields_list reads them by their exact column name, so
    # for every scored row it saw a blank Email and three blank Portfolio slots and reported
    # "a portfolio or LinkedIn/GitHub link, a contact email address" as missing - on rows that
    # plainly had both. That produced a FAIL on essentially every scored row and made the
    # whole audit unusable as a pre-flight gate.
    for col in (
        "Phone", "Location", "Country", "Current Skills", "Education",
        "Info Request Sent", "Status", "Application ID", "Original Filename",
        "Email", "Portfolio 1", "Portfolio 2", "Portfolio 3",
        "Last Updated Date", "Received Date",
    ):
        rec[col] = vals.get(col, "")
    return rec


def duplicate_notes(records: list[dict], key: str, label: str) -> list[dict]:
    buckets = {}
    for rec in records:
        value = str(rec.get(key, "") or "").strip()
        if value:
            buckets.setdefault(value, []).append(rec)
    issues = []
    for value, rows in buckets.items():
        if len(rows) > 1:
            issues.append({
                "severity": "FAIL",
                "check": f"duplicate_{label}",
                "application_id": ", ".join(r["application_id"] for r in rows),
                "sheet": ", ".join(r["sheet"] for r in rows),
                "row_index": ", ".join(str(r["row_index"]) for r in rows),
                "email": value if label == "email" else "",
                "phone": value if label == "phone" else "",
                "status": ", ".join(r["status"] for r in rows),
                "notes": f"{label} appears {len(rows)} times",
            })
    return issues


def audit_records(records: list[dict]) -> tuple[list[dict], dict]:
    issues: list[dict] = []
    counts = {
        "rows": len(records),
        "main": sum(1 for r in records if r["sheet"] == "Main"),
        "rejected": sum(1 for r in records if r["sheet"] == "Rejected"),
        "main_scored": sum(1 for r in records if r["sheet"] == "Main" and r["status"] == STATUS_SCORED),
        "main_new": sum(1 for r in records if r["sheet"] == "Main" and r["status"] == "New Email Received"),
        "main_needs_review": sum(1 for r in records if r["sheet"] == "Main" and r["status"] == STATUS_NEEDS_REVIEW),
        "main_no_matching_role": sum(
            1 for r in records
            if r["sheet"] == "Main" and r["status"] == STATUS_NO_MATCHING_ROLE),
        "main_location_review": sum(
            1 for r in records
            if r["sheet"] == "Main" and r["status"] == STATUS_LOCATION_REVIEW),
        "rejected_non_usa": sum(1 for r in records if r["sheet"] == "Rejected" and r["status"] == STATUS_REJECTED),
        "rejected_location_unconfirmed": sum(1 for r in records if r["sheet"] == "Rejected" and r["status"] == STATUS_LOCATION_UNCONFIRMED),
        "rejected_processing_failed": sum(1 for r in records if r["sheet"] == "Rejected" and r["status"] == STATUS_PROCESSING_FAILED),
        "info_sent": 0,
        "info_nothing_missing": 0,
        "info_blank_needs_attention": 0,
        "declines_sent": 0,
        "declines_other": 0,
    }
    issues.extend(duplicate_notes(records, "application_id", "application_id"))
    issues.extend(duplicate_notes(records, "email_key", "email"))
    issues.extend(duplicate_notes(records, "phone_key", "phone"))

    for rec in records:
        app_id = rec["application_id"]
        sheet = rec["sheet"]
        status = rec["status"]
        marker = str(rec["info_request_sent"] or "").strip()
        decline = str(rec["decline_sent"] or "").strip()
        if _is_sent_marker(marker):
            counts["info_sent"] += 1
        elif marker == "Nothing missing":
            counts["info_nothing_missing"] += 1
        if _is_sent_marker(decline):
            counts["declines_sent"] += 1
        elif sheet == "Rejected" and decline:
            counts["declines_other"] += 1

        if sheet == "Main" and status == STATUS_SCORED:
            decision, reason = classify_location_usa(
                rec["location"], rec["country"], phone=rec["phone"], education=rec["education"])
            if decision != GeoDecision.CONFIRMED_US:
                issues.append({
                    "severity": "FAIL",
                    "check": "unconfirmed_geo_on_scored",
                    **rec,
                    "notes": f"scored row is {decision.value}, not confirmed US: {reason}",
                })
            missing = _missing_fields_list(rec)
            if missing and not _is_settled_marker(marker):
                counts["info_blank_needs_attention"] += 1
                issues.append({
                    "severity": "FAIL",
                    "check": "missing_info_not_marked",
                    **rec,
                    "notes": "missing/unclear fields without Sent/No valid email marker: " + ", ".join(missing),
                })
            if not missing and marker not in {"", "Nothing missing"} and not _is_settled_marker(marker):
                issues.append({
                    "severity": "WARN",
                    "check": "unexpected_info_marker",
                    **rec,
                    "notes": f"unexpected Info Request Sent marker {marker!r}",
                })

        if sheet == "Main" and status == STATUS_LOCATION_REVIEW:
            decision, reason = classify_location_usa(
                rec["location"], rec["country"], phone=rec["phone"], education=rec["education"])
            if decision == GeoDecision.CONFIRMED_NON_US:
                issues.append({
                    "severity": "FAIL",
                    "check": "confirmed_non_usa_in_location_review",
                    **rec,
                    "notes": f"review row now confirms non-US: {reason}",
                })
            elif decision == GeoDecision.CONFIRMED_US:
                issues.append({
                    "severity": "WARN",
                    "check": "confirmed_usa_still_in_location_review",
                    **rec,
                    "notes": f"review row now confirms US and can be scored: {reason}",
                })
            if not marker:
                issues.append({
                    "severity": "FAIL",
                    "check": "location_review_not_marked",
                    **rec,
                    "notes": "location-review row has no follow-up/manual-review marker",
                })

        if sheet == "Rejected":
            decision, reason = classify_location_usa(
                rec["location"], rec["country"], phone=rec["phone"], education=rec["education"])
            if status == STATUS_REJECTED and decision != GeoDecision.CONFIRMED_NON_US:
                issues.append({
                    "severity": "FAIL",
                    "check": "unconfirmed_geo_on_rejected",
                    **rec,
                    "notes": f"rejected row is {decision.value}, not confirmed non-US: {reason}",
                })
            if status in {STATUS_REJECTED, STATUS_LOCATION_UNCONFIRMED} and not _is_settled_marker(decline):
                issues.append({
                    "severity": "FAIL",
                    "check": "decline_not_marked",
                    **rec,
                    "notes": f"bad/blank Decline Sent marker {decline!r}",
                })

        if status == STATUS_NEEDS_REVIEW:
            if _is_gap(rec["original_filename"]):
                issues.append({
                    "severity": "FAIL",
                    "check": "needs_review_no_original_filename",
                    **rec,
                    "notes": "Needs Review row has no Original Filename",
                })
            else:
                issues.append({
                    "severity": "INFO",
                    "check": "needs_review_unreadable",
                    **rec,
                    "notes": "Unreadable/no-content row remains queued for retry/manual review",
                })
    return issues, counts


def write_reports(issues: list[dict], counts: dict, out_dir: Path) -> tuple[Path, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = out_dir / f"sharepoint_results_health_{stamp}.csv"
    md_path = out_dir / f"sharepoint_results_health_{stamp}.md"
    fields = [
        "severity", "check", "application_id", "sheet", "row_index", "status",
        "email", "phone", "full_name", "location", "country", "category",
        "info_request_sent", "decline_sent", "retry_count", "notes",
    ]
    with csv_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for issue in issues:
            writer.writerow({k: issue.get(k, "") for k in fields})

    grouped = {sev: sum(1 for i in issues if i["severity"] == sev) for sev in ("FAIL", "WARN", "INFO")}
    with md_path.open("w", encoding="utf-8") as f:
        f.write("# SharePoint Results Health Audit\n\n")
        f.write(f"Generated: {dt.datetime.now().isoformat(timespec='seconds')}\n\n")
        f.write("## Counts\n\n")
        for key, value in counts.items():
            f.write(f"- {key}: {value}\n")
        f.write(f"- fail: {grouped['FAIL']}\n")
        f.write(f"- warn: {grouped['WARN']}\n")
        f.write(f"- info: {grouped['INFO']}\n\n")
        for sev in ("FAIL", "WARN", "INFO"):
            subset = [i for i in issues if i["severity"] == sev]
            if not subset:
                continue
            f.write(f"## {sev}\n\n")
            for issue in subset:
                f.write(
                    f"- {issue.get('check')} {issue.get('application_id')} "
                    f"[{issue.get('sheet')}:{issue.get('row_index')}] "
                    f"{issue.get('status')} | {issue.get('email')} | {issue.get('notes')}\n"
                )
            f.write("\n")
    return md_path, csv_path


def main() -> None:
    load_env()
    client = SharePointClient()
    main_rows = [
        row for row in client.list_rows()
        if not row_empty(row["values"]) and str(row["values"].get("Application ID", "") or "").strip()
    ]
    rejected_rows = [
        row for row in client.list_rejected_rows()
        if not row_empty(row["values"]) and str(row["values"].get("Application ID", "") or "").strip()
    ]
    records = [row_record("Main", row) for row in main_rows]
    records.extend(row_record("Rejected", row) for row in rejected_rows)
    issues, counts = audit_records(records)
    md_path, csv_path = write_reports(issues, counts, Path("P2_Logs") / "audits")
    grouped = {sev: sum(1 for i in issues if i["severity"] == sev) for sev in ("FAIL", "WARN", "INFO")}
    print(f"AUDIT_MD={md_path}")
    print(f"AUDIT_CSV={csv_path}")
    for key, value in counts.items():
        print(f"{key.upper()}={value}")
    print(f"FAIL={grouped['FAIL']} WARN={grouped['WARN']} INFO={grouped['INFO']}")
    for issue in issues:
        if issue["severity"] in {"FAIL", "WARN"}:
            print(
                f"{issue['severity']} {issue.get('check')} {issue.get('application_id')} "
                f"[{issue.get('sheet')}:{issue.get('row_index')}] {issue.get('notes')}"
            )


if __name__ == "__main__":
    main()
