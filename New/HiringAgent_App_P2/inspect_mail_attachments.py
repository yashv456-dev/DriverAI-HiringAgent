"""Read-only inspection of every archived application attachment for one row."""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import PurePosixPath
from urllib.parse import quote

from hiring_agent.extraction import extract_text_from_bytes
from repair_rejected_month_records import (
    GRAPH,
    SUPPORTED,
    _archive_messages,
    _bare_email,
    _date,
)
from sharepoint_client import SharePointClient

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def main(app_id: str) -> int:
    client = SharePointClient()
    row = next((r for r in client.list_rows()
                if str(r["values"].get("Application ID", "")).strip() == app_id), None)
    if row is None:
        row = next((r for r in client.list_rejected_rows()
                    if str(r["values"].get("Application ID", "")).strip() == app_id), None)
    if row is None:
        raise RuntimeError(f"Application ID not found: {app_id}")
    vals = row["values"]
    email = _bare_email(vals.get("Email"))
    received = _date(vals.get("Received Date"))
    found = 0
    for message in _archive_messages(client):
        sender = _bare_email(
            ((message.get("from") or {}).get("emailAddress") or {}).get("address"))
        when = _date(message.get("receivedDateTime"))
        distance = abs((when - received).total_seconds()) if when and received else 0
        if sender != email or not message.get("hasAttachments") or distance > 7 * 86400:
            continue
        print(f"MESSAGE {message.get('receivedDateTime')} | {message.get('subject')}")
        mid = quote(str(message["id"]), safe="")
        url = f"{GRAPH}/users/{client.sender_mailbox}/messages/{mid}/attachments"
        attachments = client._req("GET", url).json().get("value", [])
        for attachment in attachments:
            name = str(attachment.get("name") or "")
            suffix = PurePosixPath(name).suffix.lower()
            if attachment.get("isInline") or suffix not in SUPPORTED:
                continue
            aid = quote(str(attachment["id"]), safe="")
            raw = client._req(
                "GET", f"{url}/{aid}/$value").content
            text = extract_text_from_bytes(raw, f"resume{suffix}")
            preview = " | ".join(line.strip() for line in text.splitlines()[:8] if line.strip())
            print(f"  ATTACHMENT {name} | {len(raw)} bytes | {len(text.strip())} chars")
            print(f"  PREVIEW {preview[:600]}")
            lines = [line.strip() for line in text.splitlines()]
            hits = [i for i, line in enumerate(lines) if re.search(
                r"\b(education|university|college|bachelor|master|degree|b\.?s\.?|m\.?s\.?)\b",
                line, re.I)]
            if hits:
                indexes = sorted({j for i in hits for j in range(max(0, i - 1), min(len(lines), i + 2))})
                context = " | ".join(lines[j] for j in indexes if lines[j])
                print(f"  EDUCATION_CONTEXT {context[:1600]}")
            found += 1
    print(f"ATTACHMENTS_FOUND={found}")
    return 0 if found else 1


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--app-id", required=True)
    args = parser.parse_args()
    raise SystemExit(main(args.app_id))
