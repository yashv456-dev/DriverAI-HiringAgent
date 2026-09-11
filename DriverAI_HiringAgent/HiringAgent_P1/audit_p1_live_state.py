"""Read-only P1 live audit.

Checks the current Power Automate package contract against live SharePoint rows
and the apply@driverai.io mailbox. This script must not update rows, move mail,
mark messages, upload files, rename files, delete files, or send mail.
"""
from __future__ import annotations

import csv
import datetime as dt
import json
import os
import re
import sys
import zipfile
from pathlib import Path
from urllib.parse import quote

ROOT = Path(__file__).resolve().parent
P2 = next((base / "HiringAgent_P2" for base in ROOT.parents
           if (base / "HiringAgent_P2").is_dir()), None)
if P2 is None:
    raise SystemExit("Could not locate HiringAgent_P2 next to the P1 project")
sys.path.insert(0, str(P2))

from hiring_agent.config import COLUMNS, REJECTED_COLUMNS
from sharepoint_client import GRAPH, SharePointClient

APP_RE = re.compile(r"APP-\d{8}-\d{4}-[A-F0-9]{4}", re.I)
EMAIL_RE = re.compile(r"[\w.+%-]+@[\w.-]+\.[A-Za-z]{2,}")
P1_OWNED = [
    "Application ID",
    "Received Date",
    "Last Updated Date",
    "Full Name",
    "Email",
    "Mail Subject",
    "Mail Body",
    "Status",
    "Has Resume",
    "Original Filename",
    "Application Updates",
]
REQUIRED_P1_VALUES = {
    "Application ID",
    "Received Date",
    "Last Updated Date",
    "Email",
    "Status",
    "Has Resume",
    "Original Filename",
}
P1_SUBJECT_PREFIXES = (
    "application received - driverai",
    "your driverai application is already on file",
    "updated resume received - driverai",
    "message received - driverai",
)
AUDIT_DIR = ROOT / "archive" / "audits"


def load_env() -> None:
    env_path = P2 / ".env"
    for raw in env_path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip())


def bare_email(value: object) -> str:
    text = str(value or "").strip().lower()
    if "<" in text and ">" in text:
        text = text.split("<")[-1].split(">")[0]
    m = EMAIL_RE.search(text)
    return m.group(0).lower() if m else text


def parse_dt(value: object) -> dt.datetime | None:
    text = str(value or "").strip()
    if not text:
        return None
    if text.endswith("Z"):
        text = text[:-1] + "+00:00"
    try:
        parsed = dt.datetime.fromisoformat(text)
    except ValueError:
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=dt.timezone.utc)
    return parsed.astimezone(dt.timezone.utc)


def row_empty(vals: dict) -> bool:
    return not any(str(v or "").strip() for v in vals.values())


def list_messages(client: SharePointClient, folder: str, *, order_field: str) -> list[dict]:
    select = ("id,subject,receivedDateTime,sentDateTime,from,toRecipients,isRead,"
              "hasAttachments,bodyPreview,parentFolderId")
    url = (f"{GRAPH}/users/{client.sender_mailbox}/mailFolders/{quote(folder)}"
           f"/messages?$select={select}&$top=100&$orderby={order_field} desc")
    out = []
    while url:
        data = client._req("GET", url).json()
        out.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return out


def list_folders(client: SharePointClient) -> list[dict]:
    url = f"{GRAPH}/users/{client.sender_mailbox}/mailFolders?$top=200"
    folders = []
    while url:
        data = client._req("GET", url).json()
        folders.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return folders


def sent_index(sent: list[dict]) -> dict[tuple[str, str], list[dict]]:
    idx = {}
    for msg in sent:
        subject = str(msg.get("subject", "") or "")
        refs = {m.group(0).upper()
                for m in APP_RE.finditer(subject + " " + str(msg.get("bodyPreview", "") or ""))}
        recipients = []
        for rec in msg.get("toRecipients", []) or []:
            addr = ((rec.get("emailAddress") or {}).get("address") or "").strip().lower()
            if addr:
                recipients.append(addr)
        for ref in refs:
            for to in recipients:
                idx.setdefault((ref, to), []).append(msg)
    return idx


def is_p1_sent(msg: dict) -> bool:
    subject = str(msg.get("subject", "") or "").strip().lower()
    return subject.startswith(P1_SUBJECT_PREFIXES)


def msg_from(msg: dict) -> str:
    return bare_email(((msg.get("from") or {}).get("emailAddress") or {}).get("address", ""))


def app_refs_in_msg(msg: dict) -> set[str]:
    text = f"{msg.get('subject', '')} {msg.get('bodyPreview', '')}"
    return {m.group(0).upper() for m in APP_RE.finditer(text)}


def msg_date(msg: dict) -> dt.datetime | None:
    return parse_dt(msg.get("receivedDateTime") or msg.get("sentDateTime"))


def near(a: dt.datetime | None, b: dt.datetime | None, hours: int = 36) -> bool:
    if a is None or b is None:
        return False
    return abs((a - b).total_seconds()) <= hours * 3600


def package_ok() -> tuple[bool, str]:
    zip_path = ROOT / "flow" / "DriverAI-Hiring-AutoReply-apply.zip"
    def_path = ROOT / "flow" / "definition.json"
    with zipfile.ZipFile(zip_path) as z:
        names = z.namelist()
        definition_member = next((n for n in names if n.endswith("definition.json")), None)
        if not definition_member:
            return False, "zip has no definition.json"
        zipped = json.loads(z.read(definition_member).decode("utf-8"))
    live = json.loads(def_path.read_text(encoding="utf-8"))
    if zipped != live:
        return False, "zip definition differs from flow/definition.json"
    return True, (f"zip synced; LastWriteTime="
                  f"{dt.datetime.fromtimestamp(zip_path.stat().st_mtime).isoformat(timespec='seconds')}")


def main() -> None:
    load_env()
    client = SharePointClient()
    AUDIT_DIR.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    md_path = AUDIT_DIR / f"p1_live_state_audit_{stamp}.md"
    csv_path = AUDIT_DIR / f"p1_live_state_audit_{stamp}.csv"

    package_pass, package_note = package_ok()
    main_cols = client.table_columns()
    rejected_tbl = client._rejected_table_name_if_exists()
    rejected_cols = client._table_columns_of(rejected_tbl) if rejected_tbl else []

    main_rows = [r for r in client.list_rows() if not row_empty(r["values"])]
    rejected_rows = [r for r in client.list_rejected_rows() if not row_empty(r["values"])]
    records = [
        {"sheet": "Main", **r}
        for r in main_rows
        if str(r["values"].get("Application ID", "") or "").strip()
    ] + [
        {"sheet": "Rejected", **r}
        for r in rejected_rows
        if str(r["values"].get("Application ID", "") or "").strip()
    ]

    folders = list_folders(client)
    folder_summary = {
        f.get("displayName", ""): {
            "id": f.get("id", ""),
            "total": f.get("totalItemCount", ""),
            "unread": f.get("unreadItemCount", ""),
            "parent": f.get("parentFolderId", ""),
        }
        for f in folders
    }

    archive = list_messages(client, "archive", order_field="receivedDateTime")
    inbox = list_messages(client, "inbox", order_field="receivedDateTime")
    junk = list_messages(client, "junkemail", order_field="receivedDateTime")
    sent = list_messages(client, "sentitems", order_field="sentDateTime")
    sent_by_ref_to = sent_index([m for m in sent if is_p1_sent(m)])

    row_reports = []
    for rec in records:
        vals = rec["values"]
        app_id = str(vals.get("Application ID", "") or "").strip()
        email = bare_email(vals.get("Email", ""))
        received = parse_dt(vals.get("Received Date"))
        last_updated = parse_dt(vals.get("Last Updated Date"))
        status = str(vals.get("Status", "") or "").strip()
        issues = []
        warnings = []

        if not APP_RE.fullmatch(app_id):
            issues.append("bad/missing Application ID")

        for col in P1_OWNED:
            if col not in vals:
                issues.append(f"missing P1-owned column {col}")
                continue
            if col in REQUIRED_P1_VALUES and str(vals.get(col, "") or "").strip() == "":
                issues.append(f"blank P1-owned value {col}")
        if vals.get("Has Resume") != "Yes":
            issues.append("Has Resume is not Yes")
        if received is None:
            issues.append("Received Date not ISO-parseable")
        if last_updated is None:
            issues.append("Last Updated Date not ISO-parseable")
        if received and last_updated and last_updated < received:
            issues.append("Last Updated Date earlier than Received Date")

        sent_matches = sent_by_ref_to.get((app_id.upper(), email), [])
        if not sent_matches:
            warnings.append("no P1 Sent Items match by Application ID + recipient")

        archive_matches = [
            m for m in archive
            if msg_from(m) == email
            and (near(msg_date(m), received) or near(msg_date(m), last_updated))
        ]
        if not archive_matches:
            warnings.append("no Archive inbound match by sender near Received/Last Updated Date")
        elif any(not bool(m.get("isRead")) for m in archive_matches):
            issues.append("Archive inbound match is unread")

        severity = "PASS"
        if issues:
            severity = "FAIL"
        elif warnings:
            severity = "WARN"
        row_reports.append({
            "severity": severity,
            "sheet": rec["sheet"],
            "row_index": rec.get("index", ""),
            "application_id": app_id,
            "email": email,
            "status": status,
            "received_date": vals.get("Received Date", ""),
            "last_updated_date": vals.get("Last Updated Date", ""),
            "mail_sent": vals.get("Mail Sent", ""),
            "sent_matches": len(sent_matches),
            "archive_matches": len(archive_matches),
            "issues": "; ".join(issues),
            "warnings": "; ".join(warnings),
        })

    p1_sent = [m for m in sent if is_p1_sent(m)]
    no_cv_sent = [m for m in sent
                  if str(m.get("subject", "") or "").strip().lower()
                  .startswith("please attach your resume")]
    wrong_format_sent = [m for m in sent
                         if str(m.get("subject", "") or "").strip().lower()
                         .startswith("please resend your resume")]
    p1_ref_sent = [m for m in p1_sent if app_refs_in_msg(m)]
    p1_refs_in_rows = {str(r["values"].get("Application ID", "") or "").strip().upper()
                       for r in records}
    sent_refs_without_row = sorted(
        {ref for m in p1_ref_sent for ref in app_refs_in_msg(m)} - p1_refs_in_rows
    )

    archive_unread = [m for m in archive if not bool(m.get("isRead"))]
    inbox_unread = [m for m in inbox if not bool(m.get("isRead"))]
    junk_unread = [m for m in junk if not bool(m.get("isRead"))]

    fields = ("severity", "sheet", "row_index", "application_id", "email", "status",
              "received_date", "last_updated_date", "mail_sent", "sent_matches",
              "archive_matches", "issues", "warnings")
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for row in row_reports:
            writer.writerow(row)

    counts = {name: sum(1 for r in row_reports if r["severity"] == name)
              for name in ("PASS", "WARN", "FAIL")}

    with open(md_path, "w", encoding="utf-8") as f:
        f.write("# P1 Live State Audit\n\n")
        f.write(f"Generated: {dt.datetime.now().isoformat(timespec='seconds')}\n\n")
        f.write("## Package\n\n")
        f.write(f"- package synced: {package_pass}\n")
        f.write(f"- package note: {package_note}\n")
        f.write(f"- main column order matches P2 COLUMNS: {main_cols == COLUMNS}\n")
        f.write("- rejected column order matches P2 REJECTED_COLUMNS: "
                f"{rejected_cols == REJECTED_COLUMNS}\n\n")
        f.write("## SharePoint Rows\n\n")
        f.write(f"- main rows: {len(main_rows)}\n")
        f.write(f"- rejected rows: {len(rejected_rows)}\n")
        f.write(f"- total rows checked: {len(row_reports)}\n")
        f.write(f"- pass: {counts['PASS']}\n")
        f.write(f"- warn: {counts['WARN']}\n")
        f.write(f"- fail: {counts['FAIL']}\n\n")
        f.write("## Mailbox Folders\n\n")
        for name in folder_summary:
            if name in {"Archive", "Inbox", "Junk Email", "Sent Items",
                        "Recruiting Review", "Out of US"}:
                fs = folder_summary[name]
                f.write(f"- {name}: total={fs['total']} unread={fs['unread']}\n")
        f.write("\n")
        f.write("## P1 Mail Evidence\n\n")
        f.write(f"- Sent Items messages read: {len(sent)}\n")
        f.write(f"- P1 app-ref sent messages: {len(p1_ref_sent)}\n")
        f.write(f"- P1 no-CV request messages: {len(no_cv_sent)}\n")
        f.write(f"- P1 wrong-format messages: {len(wrong_format_sent)}\n")
        f.write(f"- P1 sent refs without current SharePoint row: {len(sent_refs_without_row)}\n")
        f.write(f"- Archive messages read: {len(archive)}; "
                f"unread in Archive: {len(archive_unread)}\n")
        f.write(f"- Inbox messages read: {len(inbox)}; unread in Inbox: {len(inbox_unread)}\n")
        f.write(f"- Junk messages read: {len(junk)}; unread in Junk: {len(junk_unread)}\n\n")

        for title, items in (("Failures", [r for r in row_reports if r["severity"] == "FAIL"]),
                             ("Warnings", [r for r in row_reports if r["severity"] == "WARN"])):
            f.write(f"## {title}\n\n")
            for row in items:
                note = row["issues"] if title == "Failures" else row["warnings"]
                f.write(f"- {row['application_id']} [{row['sheet']} row {row['row_index']}] "
                        f"{row['email']} | {note}\n")
            f.write("\n")

        f.write("## Sent Refs Without Current Row\n\n")
        for ref in sent_refs_without_row[:100]:
            f.write(f"- {ref}\n")
        if len(sent_refs_without_row) > 100:
            f.write(f"- ... {len(sent_refs_without_row) - 100} more\n")

    print(f"AUDIT_MD={md_path}")
    print(f"AUDIT_CSV={csv_path}")
    print(f"PACKAGE_SYNCED={package_pass}")
    print(f"MAIN_COLUMNS_MATCH={main_cols == COLUMNS}")
    print(f"REJECTED_COLUMNS_MATCH={rejected_cols == REJECTED_COLUMNS}")
    print(f"ROWS={len(row_reports)} MAIN={len(main_rows)} REJECTED={len(rejected_rows)} "
          f"PASS={counts['PASS']} WARN={counts['WARN']} FAIL={counts['FAIL']}")
    print(f"MAIL_SENT_TOTAL={len(sent)} P1_REF_SENT={len(p1_ref_sent)} "
          f"NO_CV_SENT={len(no_cv_sent)} WRONG_FORMAT_SENT={len(wrong_format_sent)}")
    print(f"ARCHIVE_TOTAL={len(archive)} ARCHIVE_UNREAD={len(archive_unread)} "
          f"INBOX_TOTAL={len(inbox)} INBOX_UNREAD={len(inbox_unread)} "
          f"JUNK_TOTAL={len(junk)} JUNK_UNREAD={len(junk_unread)}")
    for name in ("Archive", "Inbox", "Junk Email", "Sent Items",
                 "Recruiting Review", "Out of US"):
        if name in folder_summary:
            fs = folder_summary[name]
            print(f"FOLDER {name}: total={fs['total']} unread={fs['unread']}")
    for row in row_reports:
        if row["severity"] != "PASS":
            print(f"{row['severity']} {row['application_id']} {row['email']} "
                  f"{row['issues']}{row['warnings']}")


if __name__ == "__main__":
    main()
