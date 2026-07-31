"""Deep-rescore Rejected rows whose resume link still points to a month folder.

Dry-run by default. The source CV is recovered from the original Archive mailbox
attachment, so stale/missing SharePoint file links cannot force a geography decision.
Apply mode writes one row at a time, never sends mail, and enforces the storage contract:
USA -> CandidateList + received month; non-USA/unconfirmed -> Rejected + year/Rejected.
"""
from __future__ import annotations

import argparse
import datetime as dt
import re
from pathlib import PurePosixPath
from urllib.parse import quote, unquote, urlparse

from dotenv import load_dotenv

load_dotenv(".env")

import hiring_agent.config as cfg
from sharepoint_client import GRAPH, SharePointClient, SharePointError
from hiring_agent.excel_output import _parse_received
from hiring_agent.extraction import (
    _GAP_LITERALS, _ocr_pdf, ai_recheck_fields, extract_candidate_details_smart,
    extract_text_from_bytes, format_phone, html_to_text, infer_looking_for_role,
    infer_missing_portfolios, merge_mail_body_fallback, normalize_education,
    resolve_full_name, sanitize_phone,
)
from hiring_agent.geo import normalize_country, reconcile_us_country
from hiring_agent.jd_sources import get_active_roles
from hiring_agent.scoring import assign_category, suggested_roles
from hiring_agent.sharepoint_scoring import (
    _clean_category_for_filename, _get_cleaned_filename_prefix, _na_if_gap,
    _rejected_subpath, _resume_display_path, _resume_filename_tail,
)
from rescore_recent_candidates import _strict_geo

SUPPORTED = {".pdf", ".docx"}
NO_MAIL_MARKER = "N/A - USA-only deep audit; applicant email suppressed"

# Human-verified from the original CV's contact header plus current-dated education/work
# lines. These overrides exist only for this bounded historical repair. They prevent the
# small local model from treating a past school, past job, or analytics project geography
# as the candidate's current residence.
VERIFIED_OVERRIDES = {
    "APP-20260604-1210-1CC5": {
        "Full Name": "Amjad Khan", "Phone": "+91 8826711677",
        "Location": "Bengaluru, India", "Country": "India",
        "Education": ("Master of Computer Applications (MCA), Bharati Vidyapeeth Deemed "
                      "University; Bachelor of Computer Applications (BCA), IGNOU"),
        "reason": "CV contact header explicitly says Bengaluru, India",
    },
    "APP-20260602-0447-97F8": {
        "Full Name": "Irfan Ullah", "Phone": "+92 3358277075",
        "Location": "Rawalpindi, Pakistan", "Country": "Pakistan",
        "Education": "B.S. Computer Science, Islamia College University Peshawar",
        "reason": "CV work entries explicitly say Rawalpindi, Pakistan",
    },
    "APP-20260508-1109-BB49": {
        "Location": "Ahmedabad, Gujarat", "Country": "India",
        "Education": ("B.Tech in Artificial Intelligence and Data Science, "
                      "A.D. Patel Institute of Technology"),
        "reason": "CV contact header explicitly says Ahmedabad, Gujarat, India",
    },
    "APP-20260601-1700-A46E": {
        "Location": "Tempe, AZ", "Country": "United States",
        "Education": ("Master of Science in Data Science, Analytics and Engineering, "
                      "Arizona State University; Bachelor of Technology in Electronics and "
                      "Telecommunication Engineering, Dwarkadas Jivanlal Sanghvi College of Engineering"),
        "reason": ("ASU master's and Arizona phone support US presence; Singapore is project "
                   "subject matter and Mumbai entries are older work/education"),
    },
    "APP-20260716-0617-8AA5": {
        "Location": "Los Angeles, CA", "Country": "United States",
        "Education": ("Master of Science in Computer Science, University of Southern California; "
                      "Bachelor of Engineering in Computer Engineering, University of Mumbai"),
        "reason": ("USC education, USC email and Los Angeles phone support US presence; "
                   "Mumbai is the past bachelor's location"),
    },
    "APP-20260603-1133-FCB7": {
        "Location": "Kenya", "Country": "Kenya",
        "Education": "Bachelor of Applied Science in Software Engineering, Kirinyaga University",
        "reason": ("CV contact block has a Kenya phone/postal address and Kenyan university; "
                   "Tel Aviv and India are explicitly remote employer locations"),
    },
    "APP-20260507-2335-ACDD": {
        "Location": "Remote, GA", "Country": "United States",
        "Education": "Master's in Business Analytics, University of Dayton",
        "reason": ("current Apr 2025-present role is Georgia remote with a US phone; "
                   "Hyderabad belongs to older jobs"),
    },
    "APP-20260508-1624-8DFC": {
        "Location": "Portland, OR", "Country": "United States",
        "Education": ("Master's in Computer Science, Portland State University; Bachelor of "
                      "Technology in Information Technology, Anna University"),
        "reason": ("Portland State master's is marked Present, with a US phone and recent "
                   "California internship; Chennai is an older job"),
    },
    "APP-20260603-0135-B93A": {
        "Location": "Izmir, Türkiye", "Country": "Turkey",
        "Education": ("Certificate in Software and Database Programming, "
                      "University College of Applied Sciences (UCAS)"),
        "reason": "CV contact header explicitly says Izmir, Türkiye",
    },
    "APP-20260609-1724-FD48": {
        "Location": "New Haven, CT", "Country": "United States",
        "Education": ("Master of Science in Business Analytics, University of New Haven; "
                      "BBA in Business Administration, Loyola Academy"),
        "reason": ("recent New Haven master's and Connecticut phone support US presence; "
                   "Hyderabad is a past job"),
    },
    "APP-20260604-1545-9DEC": {
        "Location": "Ghaziabad, Uttar Pradesh", "Country": "India",
        "Education": ("B.Tech in Computer Science and Engineering, "
                      "Dr. A.P.J. Abdul Kalam Technical University"),
        "reason": "CV contact, current education and current internship all say Ghaziabad",
    },
    "APP-20260602-0245-038B": {
        "Location": "Pune, Maharashtra", "Country": "India",
        "Education": ("B.E. in Artificial Intelligence and Machine Learning, "
                      "PES Modern College of Engineering"),
        "reason": "CV contact header and education explicitly say Pune, Maharashtra",
    },
}


def _gap(value: object) -> bool:
    return str(value or "").strip().lower() in _GAP_LITERALS


def _bare_email(value: object) -> str:
    text = str(value or "").strip().lower()
    if "<" in text and ">" in text:
        text = text.split("<")[-1].split(">")[0]
    return text


def _date(value: object) -> dt.datetime | None:
    text = str(value or "").strip().replace("Z", "+00:00")
    try:
        parsed = dt.datetime.fromisoformat(text)
    except ValueError:
        return None
    return parsed if parsed.tzinfo else parsed.replace(tzinfo=dt.timezone.utc)


def _archive_messages(client: SharePointClient) -> list[dict]:
    select = "id,subject,receivedDateTime,from,hasAttachments,bodyPreview"
    url = (f"{GRAPH}/users/{client.sender_mailbox}/mailFolders/archive/messages"
           f"?$select={select}&$top=100&$orderby=receivedDateTime desc")
    messages: list[dict] = []
    while url:
        payload = client._req("GET", url).json()
        messages.extend(payload.get("value", []))
        url = payload.get("@odata.nextLink")
    return messages


def _mail_attachment(client: SharePointClient, messages: list[dict], vals: dict
                     ) -> tuple[str, bytes, str]:
    email = _bare_email(vals.get("Email"))
    received = _date(vals.get("Received Date"))
    matches = []
    for message in messages:
        sender = _bare_email(
            ((message.get("from") or {}).get("emailAddress") or {}).get("address"))
        when = _date(message.get("receivedDateTime"))
        if sender != email or not message.get("hasAttachments"):
            continue
        distance = abs((when - received).total_seconds()) if when and received else 0
        if distance <= 7 * 86400:
            matches.append((distance, message))
    matches.sort(key=lambda item: item[0])
    originals = {x.strip().lower() for x in
                 str(vals.get("Original Filename", "")).split(",") if x.strip()}
    for _, message in matches:
        mid = quote(message["id"], safe="")
        url = f"{GRAPH}/users/{client.sender_mailbox}/messages/{mid}/attachments"
        attachments = client._req("GET", url).json().get("value", [])
        files = [a for a in attachments
                 if not a.get("isInline")
                 and PurePosixPath(str(a.get("name", ""))).suffix.lower() in SUPPORTED]
        files.sort(key=lambda a: (
            str(a.get("name", "")).lower() not in originals,
            -int(a.get("size", 0) or 0)))
        if not files:
            continue
        attachment = files[0]
        attachment_id = quote(attachment["id"], safe="")
        raw = client._req(
            "GET",
            f"{GRAPH}/users/{client.sender_mailbox}/messages/{mid}/attachments/"
            f"{attachment_id}/$value",
        ).content
        if raw:
            return (str(attachment.get("name") or "resume.pdf"), raw,
                    str(message.get("subject") or ""))
    raise RuntimeError("no PDF/DOCX attachment found in the matching Archive message")


def _deep_text(raw: bytes, filename: str) -> tuple[str, bool]:
    text = extract_text_from_bytes(raw, filename)
    forced_ocr = False
    single = len(re.findall(r"\b[A-Za-z]\b", text))
    words = len(re.findall(r"\b[A-Za-z]{2,}\b", text))
    low_quality = len(text.strip()) < 250 or single > max(80, words * 2)
    if filename.lower().endswith(".pdf") and low_quality:
        ocr = _ocr_pdf(raw)
        if len(ocr.strip()) > len(text.strip()):
            text = ocr
            forced_ocr = True
    return text, forced_ocr


def _canonical_name(full_name: str, category: str,
                    app_id: str, source_name: str) -> str:
    ext = PurePosixPath(source_name).suffix.lower() or ".pdf"
    return (f"{_get_cleaned_filename_prefix(full_name)}_"
            f"{_clean_category_for_filename(category)}_"
            f"{_resume_filename_tail(app_id)}{ext}")


def _old_url_location(vals: dict) -> tuple[str, str] | None:
    decoded = unquote(urlparse(str(vals.get("Resume URL", "") or "")).path)
    marker = "/Shared Documents/"
    if marker not in decoded:
        return None
    rel = PurePosixPath(decoded.split(marker, 1)[1])
    return "/" + str(rel.parent), rel.name


def _extract_fields(app_id: str, vals: dict, text: str,
                    roles) -> tuple[dict, bool | None, str]:
    details = extract_candidate_details_smart(text)
    mail_body = html_to_text(str(vals.get("Mail Body", "") or ""))
    details = merge_mail_body_fallback(details, mail_body)
    details = ai_recheck_fields(details, text, mail_body)
    if _gap(details.get("looking_for_role")):
        details["looking_for_role"] = infer_looking_for_role(
            text, mail_body, details.get("skills", ""))

    full_name = resolve_full_name(
        details.get("full_name"), str(vals.get("Full Name", "")), text)
    skills = details.get("skills", "Not extracted")
    role_pref = details.get("looking_for_role", "Not extracted")
    matches = suggested_roles(skills, role_pref, roles=roles, resume_text=text)
    p1, p2, p3 = infer_missing_portfolios(
        text, details.get("portfolio_1", "N/A"),
        details.get("portfolio_2", "N/A"), details.get("portfolio_3", "N/A"))
    location = str(details.get("location", "Not extracted") or "Not extracted").strip()
    country = normalize_country(details.get("country", ""))
    phone = sanitize_phone(details.get("phone", "Not extracted"))
    fields = {
        "Full Name": full_name,
        "Phone": phone,
        "Location": location,
        "Country": country,
        "Current Skills": skills,
        "Education": normalize_education(details.get("education", "")),
        "Looking For Role": role_pref,
        "Suggested Role 1": matches["role_1"],
        "Suggested Role 2": matches["role_2"],
        "Suggested Role 3": matches.get("role_3", ""),
        "Category": assign_category(matches["role_1"], skills),
        "Portfolio 1": _na_if_gap(p1),
        "Portfolio 2": _na_if_gap(p2),
        "Portfolio 3": _na_if_gap(p3),
        "Retry Count": 0,
    }
    verified = VERIFIED_OVERRIDES.get(app_id, {})
    fields.update({key: value for key, value in verified.items() if key != "reason"})
    location = str(fields.get("Location", "") or "")
    country = normalize_country(fields.get("Country", ""))
    phone = str(fields.get("Phone", "") or "")
    verdict, geo_reason = _strict_geo(location, country, phone, text)
    reason = str(verified.get("reason") or geo_reason)
    if verdict is True:
        fields["Country"] = reconcile_us_country(country, location, True)
        fields["Phone"] = format_phone(phone)
    return fields, verdict, reason


def run(apply: bool = False, app_ids: list[str] | None = None) -> dict:
    client = SharePointClient()
    rejected = client.list_rejected_rows()
    requested = {str(x).strip() for x in (app_ids or []) if str(x).strip()}
    targets = [r for r in rejected
               if str(r["values"].get("Application ID", "")).startswith("APP-")
               and "/Rejected" not in str(r["values"].get("Resume Folder Path", ""))
               and (not requested or str(
                   r["values"].get("Application ID", "")).strip() in requested)]
    messages = _archive_messages(client)
    roles = get_active_roles()
    main_ids = {str(r["values"].get("Application ID", "")).strip()
                for r in client.list_rows()}
    delete_rejected: list[int] = []
    counts = {"usa": 0, "non_usa": 0, "unconfirmed": 0, "errors": 0}
    print(f"MODE={'APPLY' if apply else 'DRY-RUN'} TARGETS={len(targets)}", flush=True)

    for position, row in enumerate(targets, 1):
        vals, index = row["values"], row["index"]
        app_id = str(vals.get("Application ID", "")).strip()
        print(f"\n{position:02}/{len(targets)} {app_id}", flush=True)
        try:
            source_name, raw, subject = _mail_attachment(client, messages, vals)
            text, forced_ocr = _deep_text(raw, source_name)
            if not text.strip():
                raise RuntimeError("mailbox CV is unreadable even after OCR")
            fields, verdict, reason = _extract_fields(app_id, vals, text, roles)
            canonical = _canonical_name(
                fields["Full Name"], fields["Category"], app_id, source_name)
            month = cfg.dated_subpath(_parse_received(vals.get("Received Date")))
            target_subpath = month if verdict is True else _rejected_subpath(month)
            outcome = "CANDIDATELIST-USA" if verdict is True else (
                "REJECTED-NON-USA" if verdict is False
                else "REJECTED-LOCATION-UNCONFIRMED")
            key = "usa" if verdict is True else (
                "non_usa" if verdict is False else "unconfirmed")
            counts[key] += 1
            print(f"  Source     : Archive / {source_name} ({len(raw)} bytes; "
                  f"{len(text.strip())} chars; OCR={forced_ocr})", flush=True)
            print(f"  Mail       : {subject}", flush=True)
            print(f"  Education  : {fields['Education']}", flush=True)
            print(f"  Location   : {fields['Location']} | Country={fields['Country']} "
                  f"| Phone={fields['Phone']}", flush=True)
            print(f"  JD #1      : {fields['Suggested Role 1']} "
                  f"| Category={fields['Category']}", flush=True)
            print(f"  Decision   : {outcome} | {reason}", flush=True)
            print(f"  File target: {client.resumes_folder}/{target_subpath}/{canonical}",
                  flush=True)

            if not apply:
                continue

            folder = f"{client.resumes_folder}/{target_subpath}"
            # Enumerate any surviving copy by the AppID tail before writing the new
            # canonical file. This catches stale rows whose URL says month while the
            # physical file was already moved to Rejected under an older name.
            tail = _resume_filename_tail(app_id).lower()
            old_files: list[tuple[str, str]] = []
            for possible_subpath in {month, _rejected_subpath(month)}:
                possible_folder = f"{client.resumes_folder}/{possible_subpath}"
                try:
                    for item in client.list_folder_children(possible_folder):
                        if not item["is_folder"] and tail in item["name"].lower():
                            old_files.append((possible_folder, item["name"]))
                except SharePointError:
                    pass
            client.upload_file(folder, canonical, raw)
            fields["Resume URL"] = client.file_web_url(folder, canonical)
            fields["Resume Folder Path"] = _resume_display_path(
                client.resumes_folder, target_subpath)
            merged = dict(vals)
            merged.update(fields)
            old_location = _old_url_location(vals)

            if verdict is True:
                merged["Status"] = cfg.STATUS_SCORED
                if app_id in main_ids:
                    raise RuntimeError(
                        "Application ID already exists on CandidateList; refusing duplicate restore")
                client.add_main_row(merged)
                main_ids.add(app_id)
                delete_rejected.append(index)
            else:
                merged["Status"] = (cfg.STATUS_REJECTED if verdict is False
                                    else cfg.STATUS_LOCATION_UNCONFIRMED)
                if not str(merged.get(cfg.DECLINE_SENT_COLUMN, "") or "").strip():
                    merged[cfg.DECLINE_SENT_COLUMN] = NO_MAIL_MARKER
                client.update_rejected_row(index, merged, current_values=vals)

            if old_location and old_location != (folder, canonical):
                try:
                    client.delete_file(*old_location)
                except SharePointError:
                    pass
            for old_folder, old_name in old_files:
                if (old_folder, old_name) == (folder, canonical):
                    continue
                try:
                    client.delete_file(old_folder, old_name)
                except SharePointError:
                    pass
        except Exception as exc:
            counts["errors"] += 1
            print(f"  ERROR      : {exc}", flush=True)

    if apply:
        for index in sorted(set(delete_rejected), reverse=True):
            client.delete_rejected_row(index)
    print("\nSUMMARY", counts, flush=True)
    return counts


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--apply", action="store_true")
    parser.add_argument("--app-id", action="append", default=[])
    args = parser.parse_args()
    result = run(apply=args.apply, app_ids=args.app_id)
    raise SystemExit(1 if result["errors"] else 0)
