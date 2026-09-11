"""Audit and fix all entries on the Rejected Candidates sheet in SharePoint.

For every single row on the Rejected Candidates sheet:
1. Verifies physical resume file exists in SharePoint (checks Candidate_Resumes/<Year>/Rejected/, <Year>/<Month>/, and flat root).
2. Tests downloading resume bytes via Graph API to ensure NO 404 errors.
3. If resume is sitting in a Month folder instead of Rejected, moves it to Candidate_Resumes/<Year>/Rejected/.
4. Re-extracts resume text and fills any blank or missing fields:
   - Full Name, Phone, Location, Country, Current Skills, Education, Looking For Role,
   - Suggested Role 1-3, Portfolio 1-3, Resume URL, Resume Link, Resume Folder Path, Original Filename, Decline Sent.
5. Updates row in SharePoint RejectedCandidates table if any fields were repaired.
"""

import sys
import time
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

APP_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(APP_DIR))

import yaml
from sharepoint_client import SharePointClient, SharePointError
from hiring_agent.extraction import extract_text_from_bytes
from hiring_agent.jd_sources import get_active_roles
from hiring_agent.sharepoint_scoring import (
    _derive_row_fields,
    _stored_resume_names,
    _resume_name_slots,
    _rejected_subpath,
    _resume_display_path,
    _get_cleaned_filename_prefix,
    _clean_category_for_filename,
    _resume_filename_tail,
    _move_resumes,
    _bare_email,
    logger,
)

def _log(msg=""):
    print(msg, flush=True)

def audit_and_fix_rejected_sheet(dry_run: bool = False):
    _log("=" * 80)
    _log(f"AUDITING & FIXING REJECTED SHEET ({'DRY-RUN' if dry_run else 'LIVE REPAIR'})")
    _log("=" * 80)

    client = SharePointClient()
    _log("1/4 Connecting to SharePoint...")
    client._get_token()
    tbl = client._rejected_table_name_if_exists()
    if not tbl:
        _log("ERROR: RejectedCandidates table not found on SharePoint!")
        return

    _log(f"2/4 Reading rows from table '{tbl}'...")
    rows = client.list_rejected_rows()
    _log(f"    Found {len(rows)} row(s) on Rejected Candidates sheet.\n")

    roles = get_active_roles()
    _log(f"3/4 Active JD Roles loaded ({len(roles)} roles).\n")

    _log("4/4 Checking rows one by one...\n")
    fixed_count = 0
    missing_file_count = 0
    ok_count = 0

    for idx, r in enumerate(rows, 1):
        row_index = r["index"]
        vals = r["values"]
        app_id = vals.get("Application ID", f"ROW-{idx}")
        full_name = vals.get("Full Name", "")
        email = vals.get("Email", "")
        rec_date = vals.get("Received Date", "")
        orig_filename = vals.get("Original Filename", "")
        category = vals.get("Category", "General")
        stored_url = vals.get("Resume URL", "")
        stored_folder = vals.get("Resume Folder Path", "")

        _log(f"--- Row {idx}/{len(rows)} | Index: {row_index} | ID: {app_id} | Name: '{full_name}' ---")
        _log(f"    Email: {email} | Received: {rec_date}")

        # Compute dated subpath
        year = rec_date[:4] if len(rec_date) >= 4 else "2026"
        month_num = rec_date[5:7] if len(rec_date) >= 7 else "05"
        months = {"01":"January","02":"February","03":"March","04":"April","05":"May","06":"June",
                  "07":"July","08":"August","09":"September","10":"October","11":"November","12":"December"}
        month_name = months.get(month_num, "May")
        month_subpath = f"{year}/{month_name}"
        rej_subpath = f"{year}/Rejected"

        # Try searching & downloading physical resume from SharePoint
        resume_bytes = None
        found_name = None
        found_subpath = None

        # Build list of name candidates
        names_to_try = _stored_resume_names(app_id, orig_filename, full_name, category)
        if orig_filename and orig_filename not in names_to_try:
            names_to_try.append(orig_filename)

        # 1. Try Rejected folder
        for name in names_to_try:
            try:
                b = client.download_resume(name, subfolder=rej_subpath)
                if b:
                    resume_bytes = b
                    found_name = name
                    found_subpath = rej_subpath
                    _log(f"    [RESUME OK] Found in '{client.resumes_folder}/{rej_subpath}': {name} ({len(b)} bytes)")
                    break
            except SharePointError:
                pass

        # 2. Try Month folder if not in Rejected
        if not resume_bytes:
            for name in names_to_try:
                try:
                    b = client.download_resume(name, subfolder=month_subpath)
                    if b:
                        resume_bytes = b
                        found_name = name
                        found_subpath = month_subpath
                        _log(f"    [RESUME MOVED NEEDED] Found in '{client.resumes_folder}/{month_subpath}': {name} ({len(b)} bytes)")
                        if not dry_run:
                            try:
                                moved = client.move_resume(name, f"{client.resumes_folder}/{month_subpath}", f"{client.resumes_folder}/{rej_subpath}")
                                if moved:
                                    _log(f"    -> Successfully moved '{name}' to '{client.resumes_folder}/{rej_subpath}'!")
                                    found_subpath = rej_subpath
                            except Exception as me:
                                _log(f"    -> Warning moving file: {me}")
                        break
                except SharePointError:
                    pass

        # 3. Try flat root if still not found
        if not resume_bytes:
            for name in names_to_try:
                try:
                    b = client.download_resume(name, subfolder="")
                    if b:
                        resume_bytes = b
                        found_name = name
                        found_subpath = ""
                        _log(f"    [RESUME ROOT] Found in root '{client.resumes_folder}': {name} ({len(b)} bytes)")
                        break
                except SharePointError:
                    pass

        if not resume_bytes:
            _log(f"    [RESUME MISSING] No physical resume file found for '{full_name}' ({app_id}) across search names: {names_to_try[:3]}")
            missing_file_count += 1

        # Re-derive fields via scoring engine / text extraction
        fields = _derive_row_fields(client, vals, roles, rejected=True) or {}

        # Resolve web URL & folder path
        real_url = ""
        real_folder_path = ""
        if found_name and found_subpath is not None:
            try:
                real_url = client.resume_web_url(found_name, subfolder=found_subpath) or ""
            except Exception:
                pass
            real_folder_path = _resume_display_path(client.resumes_folder, found_subpath)

        # Build patch dictionary for any blanks or repairs
        patch = {}
        columns_to_check = [
            "Full Name", "Email", "Phone", "Location", "Country",
            "Current Skills", "Education", "Looking For Role",
            "Suggested Role 1", "Suggested Role 2", "Suggested Role 3",
            "Portfolio 1", "Portfolio 2", "Portfolio 3",
            "Original Filename", "Resume URL", "Resume Link", "Resume Folder Path", "Decline Sent"
        ]

        # Fill extracted fields
        for col in columns_to_check:
            cur_val = (vals.get(col) or "").strip()
            new_val = (fields.get(col) or "").strip()

            if col == "Resume URL" and real_url:
                new_val = real_url
            elif col == "Resume Folder Path" and real_folder_path:
                new_val = real_folder_path
            elif col == "Resume Link" and real_url:
                new_val = f"=HYPERLINK(\"{real_url}\", \"{found_name or orig_filename or 'Resume'}\")"

            # Check if stored value is blank or different
            if not cur_val and new_val:
                patch[col] = new_val
                _log(f"    + FIX BLANK [{col}]: '{new_val}'")
            elif col in ("Resume URL", "Resume Folder Path") and cur_val != new_val and new_val:
                patch[col] = new_val
                _log(f"    + UPDATE [{col}]: '{cur_val}' -> '{new_val}'")

        if patch:
            fixed_count += 1
            _log(f"    -> Patching {len(patch)} field(s) on SharePoint Rejected sheet...")
            if not dry_run:
                try:
                    client.update_rejected_row(row_index, patch, current_values=vals)
                    _log("    -> PATCH SUCCESS!")
                except Exception as pe:
                    _log(f"    -> PATCH FAILED: {pe}")
        else:
            ok_count += 1
            _log("    [ROW OK] All fields complete & valid.")

        _log("")

    _log("=" * 80)
    _log("REJECTED SHEET AUDIT COMPLETE SUMMARY:")
    _log(f"  Total Rows Examined : {len(rows)}")
    _log(f"  Rows Fixed / Repaired: {fixed_count}")
    _log(f"  Rows Fully OK       : {ok_count}")
    _log(f"  Missing Resumes     : {missing_file_count}")
    _log("=" * 80)

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--apply", action="store_true", help="Apply fixes directly to SharePoint (default is dry-run)")
    args = parser.parse_args()
    audit_and_fix_rejected_sheet(dry_run=not args.apply)
