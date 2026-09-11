"""Restore a CandidateList row's resume from the original Archive mailbox attachment.

Dry-run by default; --apply uploads and updates the row. Never sends mail, never rescores,
never touches any column except Resume URL / Resume Folder Path.

Written 2026-08-06 for the two rows the full audit found with no resume file on SharePoint:
  APP-20260603-0641-CC14  Hamza Asif      - lost to the _finish_rejection older-row bug
  APP-20260602-1807-82D3  Vaibhav Pawade  - pre-existing gap, cause unknown

repair_rejected_month_records.py already does this for the REJECTED sheet and is the source
of the mailbox helpers reused here; it selects rows via list_rejected_rows(), so it cannot
reach a main-sheet row.
"""
from __future__ import annotations

import argparse
import os

import hiring_agent.config  # noqa: F401
from hiring_agent.config import dated_subpath
from hiring_agent.excel_output import _parse_received
from hiring_agent.sharepoint_scoring import (
    _canonical_resume_name, _is_gap, _resume_display_path, _stored_resume_names,
)
from repair_rejected_month_records import _archive_messages, _mail_attachment
from sharepoint_client import SharePointClient

from hiring_agent.store import ExcelCandidateStore, APP_ID_COL


def _store(client):
    """Row identity — see hiring_agent/store.py. Rows are addressed by Application ID, never
    by position, so a Phase 1 append between this tool's read and its write cannot make it
    patch the wrong candidate."""
    return ExcelCandidateStore(client)


def _app_id_of(values) -> str:
    return str((values or {}).get(APP_ID_COL, "") or "").strip()



def target_name(vals: dict, app_id: str, source_name: str) -> tuple[str, str]:
    """The filename to restore under, and why.

    A SCORED row gets the canonical '<First>_<Last>_<Category>_<tail>' straight away.
    An UNSCORED row must NOT: its Category is blank, so the canonical name would bake in
    'General' and P2 would rename the file the moment it scores. Restoring under P1's own
    write shape is what P2 probes first (see _resume_name_slots), so the row scores and
    gets named canonically in one pass instead of two.
    """
    ext = os.path.splitext(source_name)[1].lower() or ".pdf"
    name, cat = vals.get("Full Name", ""), vals.get("Category", "")
    status = str(vals.get("Status", "") or "").strip()
    if _is_gap(name) or _is_gap(cat) or status in ("", "New Email Received"):
        guesses = _stored_resume_names(app_id, vals.get("Original Filename", ""),
                                       full_name=name, category=cat, resume_url=None)
        return (guesses[0] if guesses else f"{app_id}_{source_name}",
                "unscored - P1 write shape, P2 will rename on score")
    return _canonical_resume_name(name, app_id, ext, cat), "scored - canonical name"


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--apply", action="store_true",
                    help="Actually upload and update the row (default: preview only).")
    ap.add_argument("--app-id", action="append", default=[], required=True)
    args = ap.parse_args()

    client = SharePointClient()
    wanted = {x.strip() for x in args.app_id if x.strip()}
    rows = [r for r in client.list_rows()
            if str(r["values"].get("Application ID", "")).strip() in wanted]
    found_ids = {str(r["values"].get("Application ID", "")).strip() for r in rows}
    for missing in sorted(wanted - found_ids):
        print(f"  NOT ON CANDIDATELIST : {missing}")

    messages = _archive_messages(client)
    print(f"MODE={'APPLY' if args.apply else 'DRY-RUN'}  "
          f"rows={len(rows)}  archive_messages={len(messages)}\n")

    errors = 0
    for row in rows:
        vals, index = row["values"], row["index"]
        app_id = str(vals.get("Application ID", "")).strip()
        print(f"{app_id}  ({vals.get('Full Name','?')} / {vals.get('Email','?')})")
        try:
            source_name, raw, subject = _mail_attachment(client, messages, vals)
            sub = dated_subpath(_parse_received(vals.get("Received Date")))
            name, why = target_name(vals, app_id, source_name)
            folder = f"{client.resumes_folder}/{sub}"
            print(f"   mail attachment : {source_name} ({len(raw)} bytes) — {subject!r}")
            print(f"   restore as      : {sub}/{name}   [{why}]")
            if not args.apply:
                print("   (dry-run, nothing written)\n")
                continue
            client.upload_file(folder, name, raw)
            url = client.file_web_url(folder, name)
            if not url:
                raise RuntimeError("upload reported success but no webUrl came back")
            _store(client).save_by_id(_app_id_of(vals),
                              {"Resume URL": url,
                               "Resume Folder Path": _resume_display_path(
                                   client.resumes_folder, sub)},
                              current_values=vals, hint=index)
            print(f"   RESTORED        : {url}\n")
        except Exception as exc:
            errors += 1
            print(f"   ERROR           : {exc}\n")

    print(f"done | {len(rows)} row(s) | {errors} error(s)")
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
