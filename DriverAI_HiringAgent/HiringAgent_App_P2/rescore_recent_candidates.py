"""Freshly extract, JD-score, and geo-audit the latest CandidateList rows.

Dry-run by default. ``--apply`` updates confirmed-USA rows in place and moves confirmed
non-USA or unresolved-location rows to Rejected. It never sends email.
"""
from __future__ import annotations

import argparse
from dotenv import load_dotenv

load_dotenv(".env")

import hiring_agent.config as cfg
from sharepoint_client import SharePointClient, SharePointError
from hiring_agent.excel_output import _parse_received
from hiring_agent.extraction import (
    extract_candidate_details, ai_recheck_fields, merge_mail_body_fallback, html_to_text,
    infer_looking_for_role, infer_missing_portfolios, resolve_full_name, normalize_education,
    format_phone, sanitize_phone, _GAP_LITERALS,
)
from hiring_agent.geo import (
    check_location_usa, is_foreign_location, normalize_country, reconcile_us_country,
)
from hiring_agent.jd_sources import get_active_roles
from hiring_agent.scoring import suggested_roles, assign_category
from hiring_agent.sharepoint_scoring import (
    _download_resume_text, _cleanup_temp_files, _stored_resume_names, _move_resumes,
    _resume_url_and_path, _rejected_subpath, _na_if_gap,
)


def _gap(value) -> bool:
    return str(value or "").strip().lower() in _GAP_LITERALS


def _strict_geo(location: str, country: str, phone: str, resume_text: str) -> tuple[bool | None, str]:
    loc = str(location or "").strip()
    ctry = normalize_country(country)
    if _gap(loc):
        # A stored Country value is not independent evidence: this is the field that
        # the historical geo bug could set incorrectly.  Never certify a candidate as
        # USA-only when the actual current location was not extracted.
        return None, "current location could not be confirmed"
    if loc.lower() == "remote":
        if ctry == "United States":
            return True, "remote location with country confirmed as USA"
        if ctry and not _gap(ctry):
            return False, f"country confirms non-USA: {ctry}"
        return None, "current location and country could not be confirmed"
    if is_foreign_location(loc):
        return False, f"explicit foreign current location: {loc}"
    usa, reason = check_location_usa(
        loc, country=ctry, resume_text=resume_text, phone=phone)
    return (True if usa else False), reason


def run(count: int = 30, apply: bool = False,
        app_ids: list[str] | None = None) -> dict:
    client = SharePointClient()
    all_rows = sorted(client.list_rows(), key=lambda r: r["index"])
    requested = {str(value).strip() for value in (app_ids or []) if str(value).strip()}
    rows = ([row for row in all_rows
             if str(row["values"].get("Application ID", "")).strip() in requested]
            if requested else all_rows[-count:])
    rejected = client.list_rejected_rows()
    rejected_ids = {str(r["values"].get("Application ID", "")).strip() for r in rejected}
    rejected_by_id = {
        str(r["values"].get("Application ID", "")).strip(): r
        for r in rejected
        if str(r["values"].get("Application ID", "")).strip()
    }
    roles = get_active_roles()
    delete_main: list[int] = []
    kept = moved = unresolved = errors = source_missing = 0

    print(f"MODE={'APPLY' if apply else 'DRY-RUN'} COUNT={len(rows)}")
    for position, row in enumerate(rows, 1):
        vals, index = row["values"], row["index"]
        app_id = str(vals.get("Application ID", "") or "").strip()
        try:
            subpath = cfg.dated_subpath(_parse_received(vals.get("Received Date")))
            text, used, _raw = _download_resume_text(
                client, app_id, vals.get("Original Filename", ""), subpath,
                full_name=vals.get("Full Name"), category=vals.get("Category"),
            )
            if not text:
                raise RuntimeError("resume missing or unreadable")

            mail_body = html_to_text(str(vals.get("Mail Body", "") or ""))
            details = extract_candidate_details(text)
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
                text, details.get("portfolio_1", "N/A"), details.get("portfolio_2", "N/A"),
                details.get("portfolio_3", "N/A"))
            location = str(details.get("location", "Not extracted") or "Not extracted").strip()
            country = normalize_country(details.get("country", ""))
            phone = sanitize_phone(details.get("phone", "Not extracted"))
            verdict, reason = _strict_geo(location, country, phone, text)
            if verdict is True:
                country = reconcile_us_country(country, location, True)
                if country == "United States":
                    phone = format_phone(phone)

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

            if verdict is True:
                fields["Status"] = cfg.STATUS_SCORED
                if apply:
                    client.update_row(index, fields, current_values=vals)
                kept += 1
                outcome = "KEEP-USA"
            else:
                fields["Status"] = (cfg.STATUS_REJECTED if verdict is False
                                    else cfg.STATUS_LOCATION_UNCONFIRMED)
                merged = dict(vals); merged.update(fields)
                merged[cfg.DECLINE_SENT_COLUMN] = (
                    "N/A - USA-only audit; applicant email suppressed")
                if apply:
                    stored = _stored_resume_names(
                        app_id, vals.get("Original Filename", ""),
                        full_name=vals.get("Full Name"), category=vals.get("Category"))
                    if stored and _move_resumes(client, stored, subpath, to_rejected=True):
                        url, path = _resume_url_and_path(client, stored, _rejected_subpath(subpath))
                        if url:
                            merged["Resume URL"] = url
                            merged["Resume Folder Path"] = path
                    if app_id in rejected_by_id:
                        existing = rejected_by_id[app_id]
                        client.update_rejected_row(
                            existing["index"], merged,
                            current_values=existing["values"])
                    else:
                        client.add_rejected_row(merged)
                        rejected_ids.add(app_id)
                    delete_main.append(index)
                if verdict is False:
                    moved += 1
                    outcome = "MOVE-NON-USA"
                else:
                    unresolved += 1
                    outcome = "MOVE-REVIEW"

            print(f"{position:02}/{len(rows)} {app_id} {outcome} | "
                  f"{location!r}, {country!r} | {matches['role_1']!r} | {reason}")
        except Exception as exc:
            # A missing source must not bypass the USA-only contract.  Fall back to
            # the already-stored *location* solely to route the row: explicit USA
            # stays, explicit foreign moves, and gaps move to location review.
            source_missing += 1
            location = str(vals.get("Location", "") or "").strip()
            country = normalize_country(vals.get("Country", ""))
            phone = sanitize_phone(vals.get("Phone", ""))
            verdict, reason = _strict_geo(location, country, phone, "")
            if verdict is True:
                kept += 1
                outcome = "KEEP-USA-SOURCE-MISSING"
            else:
                status = (cfg.STATUS_REJECTED if verdict is False
                          else cfg.STATUS_LOCATION_UNCONFIRMED)
                merged = dict(vals)
                merged["Status"] = status
                # Any nonblank marker keeps the normal pending-decline worker from
                # emailing rows moved by this administrative cleanup.
                merged[cfg.DECLINE_SENT_COLUMN] = (
                    "N/A - USA-only audit; applicant email suppressed")
                if apply:
                    if app_id in rejected_by_id:
                        existing = rejected_by_id[app_id]
                        client.update_rejected_row(
                            existing["index"], merged,
                            current_values=existing["values"])
                    else:
                        client.add_rejected_row(merged)
                        rejected_ids.add(app_id)
                    delete_main.append(index)
                if verdict is False:
                    moved += 1
                    outcome = "MOVE-NON-USA-SOURCE-MISSING"
                else:
                    unresolved += 1
                    outcome = "MOVE-REVIEW-SOURCE-MISSING"
            print(f"{position:02}/{len(rows)} {app_id} {outcome} | "
                  f"{location!r}, {country!r} | {reason} | source error: {exc}")
        finally:
            _cleanup_temp_files()

    if apply:
        for index in sorted(set(delete_main), reverse=True):
            client.delete_row(index)
    result = {"kept_usa": kept, "moved_non_usa": moved,
              "moved_review": unresolved, "source_missing": source_missing,
              "errors": errors}
    print("SUMMARY", result)
    return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--count", type=int, default=30)
    parser.add_argument("--app-id", action="append", default=[],
                        help="Audit only this Application ID (repeatable)")
    parser.add_argument("--apply", action="store_true")
    args = parser.parse_args()
    result = run(count=max(1, args.count), apply=args.apply, app_ids=args.app_id)
    raise SystemExit(1 if result["errors"] else 0)
