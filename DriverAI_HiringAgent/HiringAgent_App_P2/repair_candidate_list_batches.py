"""Deep, silent CandidateList repair in frozen descending batches of 30.

The manifest comes from a pre-write ``backup_sharepoint.py`` checkpoint.  Each
candidate is re-read by Application ID (never by a potentially shifted row index),
the CV is downloaded/recovered, deeply extracted, scored against the current JD
cache, uploaded with overwrite semantics to its canonical SharePoint location, and
downloaded again before the workbook row is updated.

No applicant, decline, information-request, or error-alert mail is sent by this tool.
Dry-run is the default; pass ``--apply`` for live writes.
"""
from __future__ import annotations

import argparse
import hashlib
import io
import json
import os
import re
import zipfile
from pathlib import Path, PurePosixPath
from urllib.parse import parse_qs, unquote, urlparse

from dotenv import load_dotenv

load_dotenv(Path(__file__).with_name(".env"))

import hiring_agent.config as cfg
import hiring_agent.scoring as scoring_module
from hiring_agent.excel_output import _parse_received
from hiring_agent.extraction import extract_text_from_bytes
from hiring_agent.jd_sources import _read_role_cache
from hiring_agent.sharepoint_scoring import (
    _RESUME_LINK_FORMULA,
    _clean_category_for_filename,
    _get_cleaned_filename_prefix,
    _rejected_subpath,
    _resume_display_path,
    _resume_filename_tail,
)
from repair_rejected_month_records import (
    NO_MAIL_MARKER,
    VERIFIED_OVERRIDES,
    _archive_messages,
    _deep_text,
    _extract_fields,
    _mail_attachment,
)
from sharepoint_client import SharePointClient, SharePointError


BATCH_SIZE = 30
REPORT_DIR = Path(__file__).parent / "P2_Logs" / "batch_repairs"

# Where P1 saved a cover letter instead of the actual resume, select the reviewed
# resume attachment from the archived application email and overwrite the SharePoint
# candidate file during the normal canonical upload step.
ARCHIVE_RESUME_OVERRIDES = {
    "APP-20260602-1802-4DBB": "cv-youngbin-ha-driverai-android.pdf",
}

# Corrections are grounded in the contact header/current education evidence from the
# live audit.  They are deliberately specific: a global school->city rule would invent
# current residences for candidates who study remotely or have moved.
AUDIT_OVERRIDES = {
    "APP-20260522-0546-13CB": {
        "Full Name": "Mohit Singh Tevathiya",
        "Phone": "(332) 242-0955",
        "Location": "Arizona",
        "Country": "United States",
        "Education": ("Master of Science in Data Science, Analytics and Engineering, "
                      "Arizona State University; Bachelor of Technology in Electronics "
                      "and Communication, SRM University"),
    },
    "APP-20260508-1540-C732": {
        "Phone": "(617) 778-3940", "Location": "Boston, MA",
        "Country": "United States",
    },
    "APP-20260508-0417-F6B1": {
        "Phone": "(201) 978-0687", "Location": "Jersey City, NJ",
        "Country": "United States",
    },
    "APP-20260508-0042-5286": {
        "Phone": "(682) 556-1848", "Location": "Not extracted",
        "Country": "Not extracted",
        "Education": ("Master of Science in Computer Science, University of Texas "
                      "at Arlington; Bachelor of Technology in Information Technology, "
                      "Netaji Subhas University of Technology"),
    },
    "APP-20260508-1902-EE0A": {"Phone": "Not extracted"},
    "APP-20260508-0004-38B1": {
        "Location": "United States", "Country": "United States",
    },
    "APP-20260508-0010-1042": {
        "Location": "San Francisco, CA", "Country": "United States",
    },
    "APP-20260507-2340-81EC": {
        "Location": "California", "Country": "United States",
    },
    "APP-20260508-0208-38B8": {
        "Location": "Raleigh, NC", "Country": "United States",
        "Education": ("Master of Computer Science, North Carolina State University; "
                      "Bachelor of Engineering in Information Technology, "
                      "Savitribai Phule Pune University"),
    },
    "APP-20260508-0130-F11B": {
        "Location": "Michigan", "Country": "United States",
        "Education": ("Master of Science in Computer Science, Michigan State "
                      "University; B.Tech. in Computer Science, Ramaiah University "
                      "of Applied Sciences"),
    },
    "APP-20260508-0058-DE23": {
        "Location": "Not extracted", "Country": "Not extracted",
    },
    "APP-20260508-0053-3AAE": {
        "Location": "New Jersey", "Country": "United States",
        "Education": "Master of Science in Computer Science, Texas A&M University",
    },
    "APP-20260527-0349-A6ED": {
        "Location": "Arizona", "Country": "United States",
    },
    "APP-20260508-0509-D4E0": {
        "Location": "Cincinnati, OH", "Country": "United States",
    },
    "APP-20260602-0018-EDB5": {
        "Location": "Not extracted", "Country": "Not extracted",
    },
    "APP-20260602-1620-E6F2": {
        "Location": "India", "Country": "India", "Phone": "+91 9154676764",
    },
    "APP-20260602-1657-1FB2": {
        "Education": ("Master of Science in Software Engineering, Arizona State "
                      "University; Bachelor of Engineering in Information Technology, "
                      "University of Mumbai"),
    },
    "APP-20260508-0459-2817": {
        "Education": ("Master of Science in Computer Science, Arizona State University; "
                      "Bachelor of Engineering in Computer Science, Amrita Vishwa "
                      "Vidyapeetham"),
    },
    "APP-20260508-0509-D4E0": {
        # Cincinnati appears only in the education entry and the phone area
        # code; the resume provides no explicit current residence/work location.
        "Location": "Not extracted",
        "Country": "Not extracted",
    },
    "APP-20260508-0507-91E0": {
        "Education": ("Master of Science in Computer Science, Data Science "
                      "concentration, San Jose State University; Bachelor of Science "
                      "in Applied Mathematics, minor in Computer Science, Novosibirsk "
                      "State University"),
    },
    "APP-20260508-0515-CB08": {
        # College Park is present only in the active education entry; the CV
        # does not state a current residence or a city for its USA internship.
        "Location": "Not extracted",
        "Country": "Not extracted",
        "Education": ("Master of Science in Applied Machine Learning, University "
                      "of Maryland; Bachelor of Engineering in Computer Science, "
                      "Chitkara University"),
    },
    "APP-20260508-0627-1DB5": {
        "Full Name": "Harsh Patel",
        # Urbana is only implied by the active university and a local phone
        # number; the resume has no explicit residence/current-work city.
        "Location": "Not extracted",
        "Country": "Not extracted",
        "Education": ("Master of Science in Information Management, University of "
                      "Illinois Urbana-Champaign; Bachelor of Engineering in Computer "
                      "Engineering, Honors in Intelligent Computing, University of Mumbai"),
    },
    "APP-20260508-0531-80CD": {
        # State College belongs to the completed degree; the current position
        # is remote and does not establish a current city/country.
        "Location": "Not extracted",
        "Country": "Not extracted",
    },
    "APP-20260508-0655-274A": {
        "Education": ("Master of Science in Computer Science and Engineering, "
                      "California State University, San Bernardino; Bachelor of "
                      "Engineering in Electronics and Communication Engineering, "
                      "MLR Institute of Technology"),
    },
    "APP-20260508-0759-3576": {
        # Hoboken belongs to a completed degree and past assistantship. The
        # current independent role provides no residence/current-work location.
        "Location": "Not extracted",
        "Country": "Not extracted",
        "Education": ("Master of Science in Business Intelligence and Analytics, "
                      "Stevens Institute of Technology; Master of Philosophy in "
                      "Mathematics, Karakoram International University; Bachelor of "
                      "Science in Economics and Mathematics, Institute of Business "
                      "Administration Karachi"),
    },
    "APP-20260508-1230-75F9": {
        "Education": ("Master of Science in Information Management, University of "
                      "Illinois Urbana-Champaign; Bachelor of Technology in Computer "
                      "Science and Engineering, Data Science, minor in Computational "
                      "Finance, University of Mumbai"),
    },
    "APP-20260508-1310-4C65": {
        "Location": "Illinois",
        "Country": "United States",
        "Education": ("Master's degree, Campbellsville University; Bachelor's degree, "
                      "Jawaharlal Nehru Technological University Kakinada"),
    },
    "APP-20260508-1311-20DF": {
        "Education": ("Master of Science in Computer Science, Illinois Institute of "
                      "Technology; Bachelor of Technology, Jain University"),
    },
    "APP-20260508-1446-04EF": {
        "Education": ("Master of Science in Computer Science, California State "
                      "University, Fullerton"),
    },
    "APP-20260508-1500-3F37": {
        "Education": ("Master of Science in Information Management, University of "
                      "Illinois Urbana-Champaign; Bachelor of Technology in Electronics "
                      "and Telecommunication Engineering, Dwarkadas J. Sanghvi College "
                      "of Engineering, University of Mumbai"),
    },
    "APP-20260508-1540-C732": {
        "Education": ("Master of Professional Studies in Analytics, Applied Machine "
                      "Intelligence, Northeastern University; Bachelor of Engineering "
                      "in Electronics and Telecommunication, Savitribai Phule Pune "
                      "University"),
    },
    "APP-20260508-1550-408B": {
        # The current position explicitly says USA, but the CV does not give a
        # current city. Tampa belongs to the completed USF degree.
        "Location": "United States (city not specified)",
        "Country": "United States",
        "Education": ("Master of Science in Business Analytics and Information "
                      "Systems, University of South Florida; Bachelor of Engineering "
                      "in Electronics, University of Mumbai"),
    },
    "APP-20260602-1737-4CBC": {
        "Education": ("Master of Computer Applications (MCA); Bachelor of Commerce "
                      "and Computer Applications (BCCA)"),
        "Looking For Role": "Senior Mobile Engineer / Mobile Application Lead Developer",
    },
    "APP-20260602-1802-4DBB": {
        "Education": ("Master's Degree, University of Missouri - Columbia "
                      "(2019-2021)"),
        "Looking For Role": "Mobile Application Developer (Android)",
    },
    "APP-20260602-1807-1654": {
        "Education": ("Master of Engineering Management, North Carolina State "
                      "University; Bachelor of Engineering in Computer Science, "
                      "Government College of Engineering, Nagpur"),
    },
    "APP-20260602-1908-4610": {
        "Education": ("Master's in Data Science, Arizona State University; "
                      "Bachelor's in Computer Science, SRM University"),
    },
    "APP-20260602-1915-66DD": {
        "Education": ("Master's in Computer Science, California State University, "
                      "Sacramento; Bachelor's in Computer Engineering, Gujarat "
                      "Technological University - MSCET; Associate's in Computer "
                      "Engineering, Gujarat Technological University - BMEF"),
    },
    "APP-20260602-0348-B279": {
        "Education": ("MS in Information Technology and Management, University of "
                      "Texas at Dallas; BE in Computer Engineering, Pune Institute "
                      "of Computer Technology"),
    },
    "APP-20260603-1231-96B1": {
        "Location": "Not extracted",
        "Country": "Not extracted",
        "Education": ("Data Science and Computer Science, Drexel University; "
                      "Economics and Mathematics, University of Maryland, College "
                      "Park; Allegany College of Maryland"),
    },
    "APP-20260601-1501-FD04": {
        "Education": ("B.S. in Computer Science (Software Engineering), minor in "
                      "Fashion Design, Arizona State University; Associate in Arts, "
                      "Paradise Valley Community College"),
    },
    "APP-20260602-0033-1FAD": {
        "Education": ("M.S. in Software Engineering, Arizona State University; "
                      "B.Tech. in Computer Science with Big Data Analytics, SRM "
                      "Institute of Science and Technology"),
    },
    "APP-20260602-0129-9B18": {
        "Education": ("Master of Science in Data Science, Arizona State University; "
                      "Bachelor of Technology in Computer Science Engineering, SRM "
                      "Institute of Science and Technology"),
    },
    "APP-20260530-0938-D2ED": {
        "Education": ("Master of Science in Business Analytics, Arizona State "
                      "University; Bachelor of Business Administration, SRM University"),
    },
    "APP-20260507-2148-048E": {
        "Education": ("Master of Science in Data Science, Analytics and Engineering, "
                      "Arizona State University; Bachelor of Technology in Electronic "
                      "and Computer Engineering, Vellore Institute of Technology"),
    },
    "APP-20260508-0018-4786": {
        "Location": "Not extracted",
        "Country": "Not extracted",
        "Education": ("Master of Science in Computer Science, University of "
                      "Massachusetts Amherst; Bachelor of Engineering in Computer "
                      "Engineering, University of Pune"),
    },
    "APP-20260508-0029-B72E": {
        "Education": ("Master's in Computer Science, Troy University; Bachelor's "
                      "in Computer Science, Smt. Kashibai Nawale College of Engineering"),
    },
    "APP-20260508-0029-CE79": {
        "Education": ("M.S. in Applied Machine Learning, University of Maryland; "
                      "B.Tech. in Computer Science Engineering with AI and Machine "
                      "Learning, UPES School of Advanced Engineering"),
    },
    "APP-20260508-0040-65B5": {
        "Education": ("Master of Applied Science in Artificial Intelligence, Illinois "
                      "Institute of Technology; Bachelor of Technology in Information "
                      "Technology, SASTRA University"),
    },
    "APP-20260508-0051-9514": {
        "Education": ("Master of Science in Data Science, Northeastern University; "
                      "Bachelor's in Computer Science with Big Data specialization, "
                      "Mody University of Science and Technology"),
    },
    "APP-20260508-0132-DCC5": {
        "Location": "Arizona",
        "Country": "United States",
        "Education": ("B.S. in Computer Science, minor in Mathematics, Arizona State "
                      "University - Barrett Honors College"),
    },
    "APP-20260508-0155-87D2": {
        "Education": ("M.S. in Computer Science, University of Southern California; "
                      "B.E. in Computer Engineering with Honors in Blockchain, "
                      "University of Mumbai"),
    },
    "APP-20260508-0201-56AC": {
        "Education": ("Master of Science in Information Technology, Arizona State "
                      "University; Bachelor of Science in Computer Science, "
                      "Savitribai Phule Pune University"),
    },
    "APP-20260508-0213-F549": {
        "Education": ("Master of Science in Computer Science, University of Southern "
                      "California; Bachelor of Technology in Information Technology, "
                      "University of Mumbai"),
    },
    "APP-20260508-0218-8525": {
        "Education": ("M.S. in Chemical and Biomedical Engineering, University of "
                      "Wyoming; B.S. in Computational Chemistry, University of Colombo"),
    },
    "APP-20260508-0234-C67D": {
        "Full Name": "Rishal Genbhau Gawade",
        "Education": ("M.S.I.S. in AI, Software Engineering and Product, University "
                      "of San Francisco; M.S. in Computer Science / M.Tech. in Software "
                      "Engineering, Birla Institute of Technology; B.E. in Computer "
                      "Science, Pimpri Chinchwad College of Engineering"),
    },
    "APP-20260508-0328-CE6E": {
        "Education": ("Master of Science in Data Science (Computer Science), Stony "
                      "Brook University; Bachelor of Technology in Computer Science, "
                      "Chitkara Institute of Science and Technology"),
    },
    "APP-20260508-0337-00C3": {
        # The resume header has no residence address. College Park belongs to
        # the education entry; the current dated position is in Newport, NH.
        "Location": "Newport, NH",
        "Country": "United States",
        "Education": ("Master of Science in Applied Machine Learning, University "
                      "of Maryland - College of Computer, Mathematical and Natural "
                      "Sciences; Bachelor of Engineering in Computer Engineering, "
                      "Institute of Engineering and Technology, Devi Ahilya University"),
    },
    "APP-20260508-0339-3D4D": {
        # Los Angeles is attached to the completed USC degree. The current
        # dated role is explicitly in Kirkland, USA.
        "Location": "Kirkland, WA",
        "Country": "United States",
        "Education": ("Master of Science in Electrical Engineering (Machine Learning "
                      "and Data Science), University of Southern California; Bachelor "
                      "of Technology in Electronics and Communication Engineering, "
                      "Maharashtra Institute of Technology"),
    },
    "APP-20260508-0347-665C": {
        "Education": ("Master of Science in Computer Science, University of Southern "
                      "California; Bachelor of Technology in Computer Science and "
                      "Engineering, Dr. Vishwanath Karad MIT World Peace University"),
    },
    "APP-20260508-0354-4ADA": {
        "Education": ("Master of Science in Management of Technology, Arizona State "
                      "University; Bachelor of Engineering in Electronics and "
                      "Telecommunication, Savitribai Phule Pune University"),
    },
    "APP-20260508-0414-5D01": {
        "Education": ("Master of International Affairs in Finance, Computational and "
                      "Data Analysis, and Management, Columbia University; Master of "
                      "Business Administration, Indian Institute of Management Raipur"),
    },
    "APP-20260508-0417-F6B1": {
        "Full Name": "Vaibhav Barot",
        "Education": ("Master of Science in Computer Science, Stevens Institute of "
                      "Technology; Bachelor of Engineering in Information Technology, "
                      "University of Mumbai"),
    },
    "APP-20260609-2028-5EE3": {
        "Phone": "(732) 427-3680",
        "Education": ("Master of Computer Applications (MCA); Bachelor of Commerce "
                      "and Computer Applications (BCCA)"),
        "Portfolio 1": "N/A",
        "Portfolio 2": "N/A",
        "Portfolio 3": "N/A",
    },
    "APP-20260706-0752-9507": {
        "Phone": "Not extracted",
        "Location": "Not extracted",
        "Country": "Not extracted",
        "Portfolio 3": "N/A",
    },
    "APP-20260615-2255-6EEB": {
        "Education": ("M.S. in Information Systems (Data Science), University of "
                      "Maryland, Baltimore County; B.Tech. in Computer Engineering, "
                      "Dr. Babasaheb Ambedkar Technological University"),
    },
    "APP-20260619-2212-51EE": {
        "Education": ("Master's in Business Analytics and AI (Data Science "
                      "concentration), University of Texas at Dallas"),
    },
    "APP-20260622-1934-94BA": {
        "Education": ("Bachelor of Fine Arts in Human-Computer Interaction, "
                      "California College of the Arts"),
    },
    "APP-20260602-1438-8177": {
        "Education": ("Master's in Computer Science, Washington University of Science "
                      "and Technology; B.S. in Computer Science and Engineering, "
                      "Sai Tirumala NVR Engineering College"),
    },
    "APP-20260602-1445-4483": {
        "Education": ("Master's in Computer Science, University of Central Missouri; "
                      "Bachelor's in Computer Science, Vel Tech Rangarajan Dr. Sagunthala "
                      "R&D Institute of Science and Technology"),
    },
    "APP-20260602-1502-EA7E": {
        "Education": ("Master's in Applied Computer Science, Southeast Missouri State "
                      "University; Bachelor's in Computer Science and Engineering, "
                      "KL University"),
    },
}

# Same consultant submitted by two agencies.  Keep the later, file-backed application
# and merge its complete phone value; the older row's SharePoint file is already absent.
DUPLICATE_TO_KEEP = {
    "APP-20260609-1919-1CEE": "APP-20260609-2028-5EE3",
}

# Attachments proven by direct text inspection not to be candidate resumes.  Keep the
# source document for audit, but never invent candidate fields or JD scores from it.
INVALID_RESUME_OVERRIDES = {
    "APP-20260717-0559-2AA1": {
        "Full Name": "Parth Sharma",
        "Phone": "Not extracted",
        "Location": "Not extracted",
        "Country": "Not extracted",
        "Current Skills": "Not extracted",
        "Education": "Not extracted",
        "Looking For Role": "Not extracted",
        "Suggested Role 1": "",
        "Suggested Role 2": "",
        "Suggested Role 3": "",
        "Category": "General",
        "Portfolio 1": "N/A",
        "Portfolio 2": "N/A",
        "Portfolio 3": "N/A",
        "reason": ("Attachment is a Futurism Technologies corporate capability deck, "
                   "not a candidate resume"),
    },
}


def _filename_from_url(url: object) -> str:
    parsed = urlparse(str(url or ""))
    query_name = (parse_qs(parsed.query).get("file") or [""])[0]
    return query_name or unquote(parsed.path).rsplit("/", 1)[-1]


def _real_extension(raw: bytes, stored_name: str) -> str:
    if raw.startswith(b"%PDF"):
        return ".pdf"
    if raw.startswith(b"PK"):
        try:
            with zipfile.ZipFile(io.BytesIO(raw)) as archive:
                if "word/document.xml" in archive.namelist():
                    return ".docx"
        except zipfile.BadZipFile:
            pass
    return PurePosixPath(stored_name).suffix.lower()


def _canonical_name(fields: dict, app_id: str, extension: str) -> str:
    return (
        f"{_get_cleaned_filename_prefix(fields.get('Full Name', 'Candidate'))}_"
        f"{_clean_category_for_filename(fields.get('Category', 'General'))}_"
        f"{_resume_filename_tail(app_id)}{extension}"
    )


def _row_by_id(client: SharePointClient, app_id: str) -> dict | None:
    return next(
        (row for row in client.list_rows()
         if str(row["values"].get("Application ID", "")).strip() == app_id),
        None,
    )


def _download_current(client: SharePointClient, vals: dict) -> tuple[str, bytes]:
    folder = str(vals.get("Resume Folder Path", "") or "").strip().strip("/")
    name = _filename_from_url(vals.get("Resume URL"))
    if folder and name:
        try:
            return name, client.download_file(folder, name)
        except SharePointError:
            pass
    app_id = str(vals.get("Application ID", "") or "").strip()
    tail = _resume_filename_tail(app_id).lower()
    month = cfg.dated_subpath(_parse_received(vals.get("Received Date")))
    expected = f"{client.resumes_folder}/{month}".strip("/")
    for candidate_folder in (folder, expected, client.resumes_folder.strip("/")):
        if not candidate_folder:
            continue
        try:
            children = client.list_folder_children(candidate_folder)
        except SharePointError:
            continue
        for item in children:
            candidate_name = str(item.get("name", "") or "")
            if not item.get("is_folder") and tail in candidate_name.lower():
                return candidate_name, client.download_file(candidate_folder, candidate_name)
    raise SharePointError(f"no SharePoint CV found for {app_id}")


def _dedupe_skills(value: object) -> str:
    seen: set[str] = set()
    output: list[str] = []
    for piece in str(value or "").split(","):
        skill = piece.strip()
        if skill and skill.lower() not in seen:
            output.append(skill)
            seen.add(skill.lower())
    return ", ".join(output) or "Not extracted"


def _verify_uploaded(client: SharePointClient, folder: str, name: str,
                     expected: bytes) -> tuple[int, int, str]:
    downloaded = client.download_file(folder, name)
    if hashlib.sha256(downloaded).digest() != hashlib.sha256(expected).digest():
        raise RuntimeError("post-upload bytes/hash mismatch")
    text = extract_text_from_bytes(downloaded, name).strip()
    if len(text) < 80:
        raise RuntimeError(f"post-upload CV has only {len(text)} readable characters")
    web_url = client.file_web_url(folder, name)
    if not web_url:
        raise RuntimeError("SharePoint returned no webUrl for uploaded CV")
    return len(downloaded), len(text), web_url


def _clean_old_files(client: SharePointClient, app_id: str, keep_folder: str,
                     keep_name: str, month: str) -> list[str]:
    tail = _resume_filename_tail(app_id).lower()
    removed: list[str] = []
    for folder in {
        f"{client.resumes_folder}/{month}".strip("/"),
        f"{client.resumes_folder}/{_rejected_subpath(month)}".strip("/"),
    }:
        try:
            children = client.list_folder_children(folder)
        except SharePointError:
            continue
        for item in children:
            name = str(item.get("name", "") or "")
            if item.get("is_folder") or tail not in name.lower():
                continue
            if folder.lower() == keep_folder.lower() and name.lower() == keep_name.lower():
                continue
            try:
                client.delete_file(folder, name)
                removed.append(f"{folder}/{name}")
            except SharePointError:
                pass
    return removed


def _manifest_ids(path: Path, batch: int) -> list[str]:
    manifest = json.loads(path.read_text(encoding="utf-8"))
    start = (batch - 1) * BATCH_SIZE
    stop = start + BATCH_SIZE
    return [str(item.get("application_id", "")).strip()
            for item in manifest[start:stop] if str(item.get("application_id", "")).strip()]


def _frozen_cached_roles() -> tuple[list[dict], str]:
    """Use one already-verified JD snapshot for every batch in this repair.

    ``get_active_roles`` also performs a recursive live folder fingerprint on every
    process start.  That is useful in the interactive app but makes a large, audited
    repair depend on a second moving SharePoint scan.  The preflight/doctor already
    verified today's cache, so fail closed if that snapshot is unavailable.
    """
    cached = _read_role_cache() or {}
    roles = cached.get("roles") or []
    built_at = str(cached.get("built_at") or "")
    if not roles or not built_at:
        raise RuntimeError(
            "verified JD cache is unavailable; run bot.py --refresh-jd before repair")
    return roles, built_at


def _apply_override(app_id: str, fields: dict) -> dict:
    merged = dict(fields)
    for source in (VERIFIED_OVERRIDES.get(app_id, {}), AUDIT_OVERRIDES.get(app_id, {})):
        merged.update({key: value for key, value in source.items() if key != "reason"})
    invalid = INVALID_RESUME_OVERRIDES.get(app_id, {})
    merged.update({key: value for key, value in invalid.items() if key != "reason"})
    merged["Current Skills"] = _dedupe_skills(merged.get("Current Skills"))
    return merged


def _rescore_overridden_fields(fields: dict, roles: list[dict], text: str) -> dict:
    """Keep role results consistent with reviewed skills/role-preference fixes."""
    rescored = dict(fields)
    matches = scoring_module.suggested_roles(
        str(rescored.get("Current Skills", "") or ""),
        str(rescored.get("Looking For Role", "") or ""),
        roles=roles,
        resume_text=text,
    )
    rescored["Suggested Role 1"] = matches["role_1"]
    rescored["Suggested Role 2"] = matches["role_2"]
    rescored["Suggested Role 3"] = matches.get("role_3", "")
    rescored["Category"] = scoring_module.assign_category(
        matches["role_1"], str(rescored.get("Current Skills", "") or ""))
    return rescored


def _process_duplicate(client: SharePointClient, row: dict, keep_id: str,
                       apply: bool) -> dict:
    vals = row["values"]
    keeper = _row_by_id(client, keep_id)
    if keeper is None:
        raise RuntimeError(f"duplicate keeper {keep_id} is missing")
    result = {
        "application_id": vals.get("Application ID", ""),
        "action": "deduplicate",
        "keeper": keep_id,
        "status": "DRY-RUN" if not apply else "APPLIED",
    }
    if apply:
        phone = str(vals.get("Phone", "") or "").strip()
        if phone and phone.lower() not in {"n/a", "not extracted"}:
            client.update_row(
                keeper["index"], {"Phone": phone}, current_values=keeper["values"])
        client.delete_row(row["index"])
    return result


def run(batch: int, manifest: Path, apply: bool = False,
        only_app_id: str = "", prepared_report: Path | None = None) -> dict:
    # Belt-and-suspenders: this script never calls mail methods, and these config values
    # remain false even if a future helper is added accidentally.
    cfg.GEO_REJECT_EMAIL = False
    cfg.ERROR_EMAIL_ENABLED = False
    # Historical bulk repair must be reproducible.  Ollama still performs structured
    # CV extraction/recheck, but role percentages are computed against every frozen JD
    # by the deterministic overlap scorer (no subjective reranking between runs).
    scoring_module.OLLAMA_SCORING = False
    print("STARTUP creating SharePoint client", flush=True)
    client = SharePointClient()
    roles, jd_built_at = _frozen_cached_roles()
    print(f"JD_SNAPSHOT roles={len(roles)} built_at={jd_built_at}", flush=True)
    app_ids = _manifest_ids(manifest, batch)
    if only_app_id:
        if only_app_id not in app_ids:
            raise RuntimeError(f"{only_app_id} is not in frozen batch {batch}")
        app_ids = [only_app_id]
    archive_messages: list[dict] | None = None
    prepared_by_id: dict[str, dict] = {}
    if prepared_report:
        prepared_payload = json.loads(prepared_report.read_text(encoding="utf-8"))
        prepared_by_id = {
            str(item.get("application_id", "")): item
            for item in prepared_payload.get("results", [])
        }
    results: list[dict] = []
    errors = 0
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    suffix = "apply" if apply else "dry_run"
    report_tag = (f"batch_{batch}_{_resume_filename_tail(only_app_id)}"
                  if only_app_id else f"batch_{batch}")
    progress_path = REPORT_DIR / f"{report_tag}_{suffix}_progress.json"

    def checkpoint(current_app_id: str = "") -> None:
        progress_path.write_text(json.dumps({
            "batch": batch,
            "mode": suffix,
            "targets": len(app_ids),
            "completed": len(results),
            "errors": errors,
            "current_application_id": current_app_id,
            "results": results,
        }, indent=2, ensure_ascii=False), encoding="utf-8")

    checkpoint()
    print(f"MODE={'APPLY' if apply else 'DRY-RUN'} BATCH={batch} TARGETS={len(app_ids)}", flush=True)

    for number, app_id in enumerate(app_ids, 1):
        print(f"\n{number:02}/{len(app_ids)} {app_id}", flush=True)
        checkpoint(app_id)
        row = _row_by_id(client, app_id)
        if row is None:
            results.append({"application_id": app_id, "status": "ALREADY_ABSENT"})
            checkpoint(app_id)
            print("  already absent from CandidateList", flush=True)
            continue
        if app_id in DUPLICATE_TO_KEEP:
            try:
                result = _process_duplicate(
                    client, row, DUPLICATE_TO_KEEP[app_id], apply)
                results.append(result)
                checkpoint(app_id)
                print(f"  duplicate -> keep {result['keeper']}", flush=True)
            except Exception as exc:
                errors += 1
                results.append({"application_id": app_id, "status": "ERROR", "error": str(exc)})
                checkpoint(app_id)
                print(f"  ERROR {exc}", flush=True)
            continue

        vals = row["values"]
        try:
            source = "SharePoint"
            preferred_archive_name = ARCHIVE_RESUME_OVERRIDES.get(app_id)
            if preferred_archive_name:
                if archive_messages is None:
                    archive_messages = _archive_messages(client)
                archive_vals = dict(vals)
                archive_vals["Original Filename"] = preferred_archive_name
                source_name, raw, _subject = _mail_attachment(
                    client, archive_messages, archive_vals)
                source = "Outlook Archive (reviewed resume)"
            else:
                try:
                    source_name, raw = _download_current(client, vals)
                except SharePointError:
                    if archive_messages is None:
                        archive_messages = _archive_messages(client)
                    source_name, raw, _subject = _mail_attachment(
                        client, archive_messages, vals)
                    source = "Outlook Archive"
            extension = _real_extension(raw, source_name)
            if extension not in {".pdf", ".docx"}:
                raise RuntimeError(f"unsupported/unknown CV format for {source_name}")
            parse_name = f"resume{extension}"
            text, forced_ocr = _deep_text(raw, parse_name)
            if len(text.strip()) < 80:
                raise RuntimeError(f"CV unreadable after extraction/OCR ({len(text.strip())} chars)")
            prepared = prepared_by_id.get(app_id) if apply else None
            if prepared:
                expected_hash = str(prepared.get("source_sha256") or "")
                actual_hash = hashlib.sha256(raw).hexdigest()
                if expected_hash and expected_hash != actual_hash:
                    raise RuntimeError("CV changed since dry-run; refusing prepared apply")
                fields = dict(prepared.get("prepared_fields") or vals)
                fields.pop("Resume Link", None)
                # Compatibility with the first dry-run generated before full prepared
                # fields were added to the report.
                legacy_map = {
                    "full_name": "Full Name", "phone": "Phone",
                    "location": "Location", "country": "Country",
                    "education": "Education", "role_1": "Suggested Role 1",
                    "category": "Category",
                }
                for result_key, field_key in legacy_map.items():
                    if prepared.get(result_key) not in (None, ""):
                        fields[field_key] = prepared[result_key]
                fields = _apply_override(app_id, fields)
                verdict = prepared.get("geo_verdict")
                reason = str(prepared.get("geo_reason") or "prepared dry-run")
                print(f"  Prepared  : {prepared_report.name} (CV hash/bytes revalidated)", flush=True)
            else:
                fields, verdict, reason = _extract_fields(app_id, vals, text, roles)
                fields = _apply_override(app_id, fields)
            fields = _rescore_overridden_fields(fields, roles, text)
            # Re-evaluate strict geography after every override, including prepared
            # applies.  A reviewed correction must be able to overturn a stale dry-run
            # verdict without rerunning the expensive extraction model.
            from rescore_recent_candidates import _strict_geo
            verdict, reason_after = _strict_geo(
                str(fields.get("Location", "")), str(fields.get("Country", "")),
                str(fields.get("Phone", "")), text)
            reason = reason_after or reason
            invalid_resume = INVALID_RESUME_OVERRIDES.get(app_id)
            if invalid_resume:
                verdict = None
                reason = str(invalid_resume["reason"])
            canonical = _canonical_name(fields, app_id, extension)
            month = cfg.dated_subpath(_parse_received(vals.get("Received Date")))
            target_subpath = month if verdict is True else _rejected_subpath(month)
            target_folder = f"{client.resumes_folder}/{target_subpath}".strip("/")
            status = (cfg.STATUS_PROCESSING_FAILED if invalid_resume else
                      (cfg.STATUS_SCORED if verdict is True else
                       (cfg.STATUS_REJECTED if verdict is False else
                        cfg.STATUS_LOCATION_UNCONFIRMED)))

            result = {
                "application_id": app_id,
                "source": source,
                "source_name": source_name,
                "source_bytes": len(raw),
                "source_sha256": hashlib.sha256(raw).hexdigest(),
                "text_chars": len(text.strip()),
                "forced_ocr": forced_ocr,
                "full_name": fields.get("Full Name", ""),
                "phone": fields.get("Phone", ""),
                "location": fields.get("Location", ""),
                "country": fields.get("Country", ""),
                "education": fields.get("Education", ""),
                "role_1": fields.get("Suggested Role 1", ""),
                "category": fields.get("Category", ""),
                "geo_verdict": verdict,
                "geo_reason": reason,
                "status": status,
                "target_folder": target_folder,
                "target_name": canonical,
                "mode": "APPLY" if apply else "DRY-RUN",
                "prepared_fields": fields,
            }
            print(f"  Source    : {source}/{source_name} ({len(raw)} bytes, {len(text.strip())} chars)", flush=True)
            print(f"  Identity  : {fields.get('Full Name')} | {fields.get('Phone')}", flush=True)
            print(f"  Geo       : {fields.get('Location')} | {fields.get('Country')} -> {status}", flush=True)
            print(f"  Education : {fields.get('Education')}", flush=True)
            print(f"  JD #1     : {fields.get('Suggested Role 1')} | {fields.get('Category')}", flush=True)
            print(f"  CV target : {target_folder}/{canonical}", flush=True)

            if apply:
                client.upload_file(target_folder, canonical, raw)
                byte_count, char_count, web_url = _verify_uploaded(
                    client, target_folder, canonical, raw)
                fields.update({
                    "Status": status,
                    "Resume URL": web_url,
                    "Resume Folder Path": _resume_display_path(
                        client.resumes_folder, target_subpath),
                    "Retry Count": 0,
                })
                if verdict is True:
                    client.update_row(row["index"], fields, current_values=vals)
                else:
                    merged = dict(vals)
                    merged.update(fields)
                    merged[cfg.DECLINE_SENT_COLUMN] = NO_MAIL_MARKER
                    client.add_rejected_row(merged)
                    client.delete_row(row["index"])
                removed = _clean_old_files(
                    client, app_id, target_folder, canonical, month)
                result.update({
                    "verified_bytes": byte_count,
                    "verified_text_chars": char_count,
                    "resume_url": web_url,
                    "removed_stale_files": removed,
                })
                print(f"  VERIFIED  : {byte_count} bytes / {char_count} chars / URL OK", flush=True)
            results.append(result)
            checkpoint(app_id)
        except Exception as exc:
            errors += 1
            results.append({"application_id": app_id, "status": "ERROR", "error": str(exc)})
            checkpoint(app_id)
            print(f"  ERROR     : {exc}", flush=True)

    if apply:
        # Restore a genuine calculated hyperlink column after all row writes/deletes.
        if str(os.getenv("HIRING_REPAIR_RESUME_LINK_FORMULA", "") or "").strip().lower() in {"1", "true", "yes", "on"}:
            client.set_calculated_column(client.table, "Resume Link", _RESUME_LINK_FORMULA)
            try:
                rejected_table = client._ensure_rejected_table()
                client.set_calculated_column(rejected_table, "Resume Link", _RESUME_LINK_FORMULA)
            except SharePointError:
                pass

    report_path = REPORT_DIR / f"{report_tag}_{suffix}.json"
    report_path.write_text(json.dumps({
        "batch": batch,
        "mode": suffix,
        "manifest": str(manifest),
        "jd_snapshot_built_at": jd_built_at,
        "jd_role_count": len(roles),
        "scoring_mode": "deterministic keyword overlap across frozen JD snapshot",
        "targets": len(app_ids),
        "errors": errors,
        "results": results,
    }, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\nREPORT={report_path}", flush=True)
    print(f"SUMMARY targets={len(app_ids)} errors={errors}", flush=True)
    return {"targets": len(app_ids), "errors": errors, "report": str(report_path)}


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--batch", type=int, choices=range(1, 6), required=True)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--app-id", default="",
                        help="Process exactly one Application ID from the selected batch")
    parser.add_argument("--apply", action="store_true")
    parser.add_argument("--prepared-report", type=Path,
                        help="Apply fields from a reviewed dry-run report after revalidating CV bytes")
    args = parser.parse_args()
    summary = run(args.batch, args.manifest, apply=args.apply,
                  only_app_id=args.app_id.strip(), prepared_report=args.prepared_report)
    raise SystemExit(1 if summary["errors"] else 0)
