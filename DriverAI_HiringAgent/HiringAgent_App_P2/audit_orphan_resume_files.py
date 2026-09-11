"""Read-only audit of resume FILES that no row points at.

`audit_all_resume_files.py` is row-driven: it walks every candidate row and checks that the
file it expects exists. That direction cannot see the opposite failure - a file sitting in
SharePoint that no row references at all. This script walks the FOLDERS instead.

Why the orphans exist. The last 4 characters of an Application ID come from
`toUpper(substring(guid(),0,4))` (HiringAgent_P1/flow/build_zip.py) - random per RUN, not
derived from the email. So re-processing the same message mints a brand-new Application ID
and writes a brand-new resume file, leaving the previous one stranded under the old ref. The
2026-08-24 audit traced replays to 12 Jul, 04 Aug and 15 Aug. Live example, the April folder:

    ajinanpa_APP-20260420-0133-61B2.pdf   the file
    APP-20260420-0133-25DD                the row for that same sender, same minute

A second class is worse: a file whose sender has no row on EITHER sheet - the applicant was
lost outright (`alokshain007_APP-20260421-0345-2AFF.pdf`).

WRITES NOTHING to SharePoint. It only lists folder children and reads workbook rows - there
is no update/add/delete/move/rename/upload or mail call anywhere in this file. Deciding what
to do with an orphan is a separate, deliberate step.

    python audit_orphan_resume_files.py                    # every year found
    python audit_orphan_resume_files.py --year 2026
    python audit_orphan_resume_files.py --months April,March
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import os
import re
from pathlib import Path
from urllib.parse import unquote, urlparse

from hiring_agent.sharepoint_scoring import _resume_name_slots
from sharepoint_client import SharePointClient, SharePointError

#: 'sathishsravanakumar_APP-20260414-0041-9D05.pdf' -> sender hint + the full reference.
_FILE_REF_RE = re.compile(r"^(?P<who>.*?)_?(?P<ref>APP-\d{8}-\d{4,6}-[0-9A-Za-z]{2,8})",
                          re.IGNORECASE)
#: The minute an Application ID encodes: APP-<yyyymmdd>-<HHmm>-<tail>.
_REF_MINUTE_RE = re.compile(r"^APP-(\d{8})-(\d{4})", re.IGNORECASE)

RESUME_EXTS = (".pdf", ".docx", ".doc", ".rtf", ".txt")


def load_env() -> None:
    """Same .env convention as the other P2 tools; never overrides a real environment var."""
    env_path = Path(__file__).with_name(".env")
    if not env_path.exists():
        return
    for line in env_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, val = line.partition("=")
        os.environ.setdefault(key.strip(), val.strip().strip('"').strip("'"))


def filename_from_url(url: str) -> str:
    if not url:
        return ""
    return unquote(urlparse(str(url)).path).rsplit("/", 1)[-1]


def ref_minute(ref: str) -> str:
    """'APP-20260420-0133-61B2' -> '20260420-0133'; '' when the shape does not match.

    This is the join key for the replay case: a replayed email keeps its receivedDateTime,
    so the date and time survive intact and only the random tail differs.
    """
    m = _REF_MINUTE_RE.match(str(ref or "").strip())
    return f"{m.group(1)}-{m.group(2)}" if m else ""


def parse_received(value) -> dt.datetime | None:
    if isinstance(value, dt.datetime):
        return value
    s = str(value or "").strip().replace("Z", "")
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def collect_rows(client: SharePointClient) -> list[dict]:
    """Every nonblank row on both sheets, tagged with which sheet it came from."""
    out = []
    for sheet, getter in (("CandidateList", client.list_rows),
                          ("Rejected", client.list_rejected_rows)):
        try:
            rows = getter()
        except SharePointError as exc:
            print(f"  [WARN] could not read {sheet}: {exc}")
            continue
        for r in rows:
            vals = r.get("values", {})
            if str(vals.get("Application ID", "") or "").strip():
                out.append({"sheet": sheet, "index": r.get("index"), "values": vals})
    return out


def referenced_names(rows: list[dict]) -> dict[str, list[dict]]:
    """filename -> the row(s) that could legitimately be pointing at it.

    Uses P2's own `_resume_name_slots`, so this audit agrees exactly with what the scorer
    would actually look for, plus whatever the row's Resume URL literally names.
    """
    index: dict[str, list[dict]] = {}
    for row in rows:
        v = row["values"]
        names = set()
        url_name = filename_from_url(v.get("Resume URL", ""))
        if url_name:
            names.add(url_name)
        try:
            for slot in _resume_name_slots(
                    str(v.get("Application ID", "")),
                    str(v.get("Original Filename", "") or ""),
                    str(v.get("Full Name", "") or ""),
                    str(v.get("Category", "") or "")):
                names.update(slot)
        except Exception:
            pass  # a malformed row must not abort the sweep
        for n in names:
            if n:
                index.setdefault(n.lower(), []).append(row)
    return index


def walk_resume_tree(client: SharePointClient, root: str, years: list[str] | None,
                     months: list[str] | None) -> list[tuple[str, str]]:
    """[(folder, filename)] for every resume-looking file under the resume root."""
    found: list[tuple[str, str]] = []

    def children(path: str) -> list[dict]:
        try:
            return client.list_folder_children(path)
        except SharePointError as exc:
            print(f"  [WARN] cannot list {path}: {exc}")
            return []

    for year_item in children(root):
        if not year_item["is_folder"]:
            continue
        year = year_item["name"]
        if years and year not in years:
            continue
        for sub in children(f"{root}/{year}"):
            if not sub["is_folder"]:
                continue
            if months and sub["name"] not in months:
                continue
            folder = f"{root}/{year}/{sub['name']}"
            for item in children(folder):
                if not item["is_folder"] and item["name"].lower().endswith(RESUME_EXTS):
                    found.append((folder, item["name"]))
            # One more level: a month folder can hold its own Rejected bucket.
            for deep in children(folder):
                if deep["is_folder"]:
                    for item in children(f"{folder}/{deep['name']}"):
                        if not item["is_folder"] and item["name"].lower().endswith(RESUME_EXTS):
                            found.append((f"{folder}/{deep['name']}", item["name"]))
    return found


def classify(folder: str, name: str, refs: dict, minutes: dict,
             senders: dict) -> dict:
    """Work out what an orphan file most likely belongs to, without changing anything."""
    m = _FILE_REF_RE.match(name)
    file_ref = (m.group("ref").upper() if m else "")
    who = (m.group("who") or "").strip("_ ").lower() if m else ""

    rec = {"folder": folder, "file": name, "ref_in_filename": file_ref,
           "sender_hint": who, "verdict": "", "matched_row": "", "matched_sheet": "",
           "matched_email": "", "detail": ""}

    if file_ref and file_ref in refs:
        rec["verdict"] = "OK"
        return rec

    minute = ref_minute(file_ref)
    same_minute = minutes.get(minute, []) if minute else []
    by_sender = senders.get(who, []) if who else []

    if same_minute:
        r = same_minute[0]
        rec.update(verdict="REPLAY-ORPHAN",
                   matched_row=str(r["values"].get("Application ID", "")),
                   matched_sheet=r["sheet"],
                   matched_email=str(r["values"].get("Email", "")),
                   detail="same received minute, different random ref tail - the row was "
                          "re-minted by a replay and this file was left behind")
    elif by_sender:
        r = by_sender[0]
        rec.update(verdict="SENDER-ORPHAN",
                   matched_row=str(r["values"].get("Application ID", "")),
                   matched_sheet=r["sheet"],
                   matched_email=str(r["values"].get("Email", "")),
                   detail="filename's sender hint matches a row with a DIFFERENT date - "
                          "likely an earlier application whose row no longer exists")
    else:
        rec.update(verdict="NO-ROW",
                   detail="no row on either sheet matches this ref, its minute, or its "
                          "sender hint - this applicant has no record at all")
    return rec


def write_reports(records: list[dict], out_dir: Path) -> tuple[Path, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%dT%H%M%S")
    csv_path = out_dir / f"orphan_resume_audit_{stamp}.csv"
    md_path = out_dir / f"orphan_resume_audit_{stamp}.md"

    fields = ["verdict", "folder", "file", "ref_in_filename", "sender_hint",
              "matched_row", "matched_sheet", "matched_email", "detail"]
    with csv_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in records:
            w.writerow({k: r.get(k, "") for k in fields})

    orphans = [r for r in records if r["verdict"] != "OK"]
    by_verdict: dict[str, list[dict]] = {}
    for r in orphans:
        by_verdict.setdefault(r["verdict"], []).append(r)

    lines = [f"# Orphaned resume files - {stamp}", "",
             f"- files scanned: **{len(records)}**",
             f"- referenced by a row: **{len(records) - len(orphans)}**",
             f"- orphaned: **{len(orphans)}**", ""]
    for verdict, blurb in (
            ("REPLAY-ORPHAN", "Same received minute as an existing row, different random ref "
                              "tail. The email was processed twice; the row kept the newer "
                              "ref and this file was stranded under the older one."),
            ("SENDER-ORPHAN", "The sender matches a row, but at a different date - most "
                              "likely an earlier application whose row no longer exists."),
            ("NO-ROW", "**No record anywhere.** A resume is in SharePoint and this applicant "
                       "appears on neither sheet.")):
        group = by_verdict.get(verdict, [])
        if not group:
            continue
        lines += [f"## {verdict} ({len(group)})", "", blurb, "",
                  "| folder | file | ref in filename | closest row | sheet | email |",
                  "|---|---|---|---|---|---|"]
        for r in sorted(group, key=lambda x: (x["folder"], x["file"])):
            lines.append(f"| {r['folder']} | `{r['file']}` | `{r['ref_in_filename']}` | "
                         f"`{r['matched_row']}` | {r['matched_sheet']} | {r['matched_email']} |")
        lines.append("")
    lines += ["---", "", "This audit changed nothing. Every row above is a finding, not an "
                        "action taken."]
    md_path.write_text("\n".join(lines), encoding="utf-8")
    return md_path, csv_path


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--year", help="only this year folder, e.g. 2026")
    ap.add_argument("--months", help="comma-separated month folders, e.g. April,March")
    args = ap.parse_args()

    load_env()
    client = SharePointClient()
    root = (client.resumes_folder or "/Shared Documents/Candidate_Resumes")
    root = root.replace("/Shared Documents/", "").strip("/")

    print(f"Resume root : {root}")
    rows = collect_rows(client)
    print(f"Rows loaded : {len(rows)} (both sheets)")

    refs = {str(r["values"].get("Application ID", "")).strip().upper(): r for r in rows}
    minutes: dict[str, list[dict]] = {}
    senders: dict[str, list[dict]] = {}
    for r in rows:
        mk = ref_minute(str(r["values"].get("Application ID", "")))
        if mk:
            minutes.setdefault(mk, []).append(r)
        email = str(r["values"].get("Email", "") or "").strip().lower()
        if "@" in email:
            senders.setdefault(email.split("@", 1)[0], []).append(r)

    known = referenced_names(rows)
    files = walk_resume_tree(client, root,
                             [args.year] if args.year else None,
                             [m.strip() for m in args.months.split(",")] if args.months else None)
    print(f"Files found : {len(files)}")

    records = []
    for folder, name in files:
        if name.lower() in known:
            records.append({"folder": folder, "file": name, "verdict": "OK",
                            "ref_in_filename": "", "sender_hint": "", "matched_row": "",
                            "matched_sheet": "", "matched_email": "",
                            "detail": "a row points at this file"})
        else:
            records.append(classify(folder, name, refs, minutes, senders))

    md_path, csv_path = write_reports(records, Path("P2_Logs") / "audits")
    counts: dict[str, int] = {}
    for r in records:
        counts[r["verdict"]] = counts.get(r["verdict"], 0) + 1

    print()
    for verdict in ("OK", "REPLAY-ORPHAN", "SENDER-ORPHAN", "NO-ROW"):
        if verdict in counts:
            print(f"  {verdict:<15} {counts[verdict]}")
    print()
    for r in records:
        if r["verdict"] == "NO-ROW":
            print(f"  NO-ROW  {r['folder']}/{r['file']}")
    print(f"\nAUDIT_MD={md_path}\nAUDIT_CSV={csv_path}")
    print("Nothing was modified.")


if __name__ == "__main__":
    main()
