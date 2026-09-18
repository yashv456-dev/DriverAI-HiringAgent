"""
Candidate routing logic — determines which of the 6 P1 paths applies to a message.

Routing paths (in evaluation order):
  1. SPAM / BAD_SENDER / BAD_SUBJECT  → classify.py handles these before routing
  2. NEW_CANDIDATE     — new sender + PDF/DOCX resume → create row, send Ack
  3. DUPLICATE         — same sender within 90 days, no quoted ref → update row, send Dup Notice
  4. RESUME_UPDATE     — quoted valid APP-Ref + new CV → update row, send Update Ack
  5. FOLLOWUP          — quoted valid APP-Ref, no CV → update row counter, send Noted
  6. CV_REQUEST        — no resume, contains app keywords → send CV Request (no row)
  7. WRONG_FORMAT      — attachment present but not PDF/DOCX → send Format Request (no row)
  8. IGNORED_MAIL      — clean non-application → archive, alert admin (no row)

CRITICAL duplicate window:
  Uses max(received_date) across all rows for the same email — NOT sort order.
  This makes detection order-independent (out-of-order delivery safe).

CRITICAL IsKnownSender guard:
  When a quoted APP-Ref is present, sender email MUST exactly match the record.
  Mismatch → treat as normal intake (prevent cross-candidate record corruption).

Anti-flood caps (from flow_config.json business_rules):
  duplicate_notice_max  = 5
  update_resume_max     = 5
  followup_reply_max    = 5
  reply_cap             = 3  — contacts 1–3 get auto-replies; 4–5 update silently
"""
from __future__ import annotations

import logging
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from enum import Enum
from typing import Any

from .classify import (
    RESUME_EXTENSIONS,
    APP_KEYWORDS,
    Classifier,
    ClassifyResult,
    Classification,
    BLOCKED,
)

logger = logging.getLogger(__name__)


# ─── Route enum ───────────────────────────────────────────────────────────────

class Route(str, Enum):
    """The processing path assigned to a passing message."""
    NEW_CANDIDATE  = "new_candidate"
    DUPLICATE      = "duplicate"
    RESUME_UPDATE  = "resume_update"
    FOLLOWUP       = "followup"
    CV_REQUEST     = "cv_request"
    WRONG_FORMAT   = "wrong_format"
    IGNORED_MAIL   = "ignored_mail"
    # Blocked routes (gates fired)
    SPAM           = "spam"
    BAD_SENDER     = "bad_sender"
    BAD_SUBJECT    = "bad_subject"
    SHORTENER_ONLY = "shortener_only"

CREATES_ROW    = {Route.NEW_CANDIDATE}
UPDATES_ROW    = {Route.DUPLICATE, Route.RESUME_UPDATE, Route.FOLLOWUP}
NO_ROW         = {Route.CV_REQUEST, Route.WRONG_FORMAT, Route.IGNORED_MAIL}
BLOCKED_ROUTES = {Route.SPAM, Route.BAD_SENDER, Route.BAD_SUBJECT, Route.SHORTENER_ONLY}


@dataclass
class RouteResult:
    route: Route
    # If a quoted APP-Ref was found and verified, it's here.
    verified_app_ref: str | None = None
    # The existing DB record for DUPLICATE / UPDATE / FOLLOWUP paths.
    existing_record: dict[str, Any] | None = None
    # If caps are hit, send_email=False (still write to DB).
    send_email: bool = True
    detail: str = ""


# ─── Business rules (mirrors flow_config.json business_rules) ─────────────────

DUPLICATE_CHECK_DAYS  = 90
DUPLICATE_NOTICE_MAX  = 5
UPDATE_RESUME_MAX     = 5
FOLLOWUP_REPLY_MAX    = 5
REPLY_CAP             = 3  # contacts > REPLY_CAP still update DB but suppress email


# ─── Router ───────────────────────────────────────────────────────────────────

class Router:
    """
    Determines the routing path for a message that passed Gates 1–3.

    Requires a callable `lookup` that accepts (sender_email: str) and returns
    a list of dicts representing candidate records, each with at minimum:
        - "Application ID"   : str
        - "Email"            : str (normalized lowercase)
        - "Received Date"    : str (ISO-8601)
        - "Application Updates": int
    """

    def __init__(self, classifier: Classifier | None = None):
        from .classify import classifier as _default
        self._classifier = classifier or _default

    def route(
        self,
        sender_email: str,
        subject: str,
        body: str,
        attach_names: list[str],
        classify_result: ClassifyResult,
        *,
        lookup_by_email: "callable[[str], list[dict]]",
        lookup_by_app_ref: "callable[[str], dict | None]",
        now: datetime | None = None,
    ) -> RouteResult:
        """
        Determine the route for a message that passed the classifier gates.

        Args:
            sender_email: Normalized lowercase sender.
            subject, body, attach_names: Message content.
            classify_result: Output from Classifier.run_gates().
            lookup_by_email: Callable → list of existing records for this email.
            lookup_by_app_ref: Callable → single record for a given APP-Ref, or None.
            now: Override current time (for testing). Defaults to UTC now.
        """
        if classify_result.classification in BLOCKED:
            return RouteResult(
                route=Route(classify_result.classification.value),
                detail=f"Blocked at gate: {classify_result.gate_hit}",
            )

        now = now or datetime.now(timezone.utc)
        c = self._classifier

        has_resume = c.has_resume_attachment(attach_names)
        has_other_attach = c.has_non_resume_attachment(attach_names)
        quoted_ref = c.extract_quoted_ref(subject, body)

        # ── Path A: Quoted APP-Ref present ─────────────────────────────────
        if quoted_ref:
            record = lookup_by_app_ref(quoted_ref)
            if record:
                stored_email = (record.get("Email") or "").strip().lower()
                # IsKnownSender guard — exact equality required.
                if sender_email == stored_email:
                    if has_resume:
                        update_count = int(record.get("Application Updates", 0))
                        if update_count >= UPDATE_RESUME_MAX:
                            logger.info("Update cap hit for %s (count=%d)", quoted_ref, update_count)
                            return RouteResult(
                                route=Route.RESUME_UPDATE,
                                verified_app_ref=quoted_ref,
                                existing_record=record,
                                send_email=False,
                                detail=f"update cap {update_count}/{UPDATE_RESUME_MAX}",
                            )
                        send = update_count < REPLY_CAP
                        return RouteResult(
                            route=Route.RESUME_UPDATE,
                            verified_app_ref=quoted_ref,
                            existing_record=record,
                            send_email=send,
                        )
                    else:
                        # Follow-up note: no CV attached.
                        followup_count = int(record.get("Application Updates", 0))
                        if followup_count >= FOLLOWUP_REPLY_MAX:
                            return RouteResult(
                                route=Route.FOLLOWUP,
                                verified_app_ref=quoted_ref,
                                existing_record=record,
                                send_email=False,
                                detail=f"followup cap {followup_count}/{FOLLOWUP_REPLY_MAX}",
                            )
                        send = followup_count < REPLY_CAP
                        return RouteResult(
                            route=Route.FOLLOWUP,
                            verified_app_ref=quoted_ref,
                            existing_record=record,
                            send_email=send,
                        )
                else:
                    # IsKnownSender failed — drop to normal intake.
                    logger.warning(
                        "IsKnownSender FAILED: quoted %s but sender %r ≠ stored %r — "
                        "dropping to normal intake",
                        quoted_ref, sender_email, stored_email,
                    )
                    # Fall through to Path B (no quoted ref treatment).
            # Ref not found in DB → treat as normal intake.

        # ── Path B: No quoted ref (or unverified ref) ──────────────────────
        if not has_resume:
            if has_other_attach:
                # Has attachments but none are PDF/DOCX.
                return RouteResult(route=Route.WRONG_FORMAT)
            if c.has_app_keywords(body, subject):
                # Intent to apply but forgot the resume.
                return RouteResult(route=Route.CV_REQUEST)
            # No resume, no keywords → ignored non-application mail.
            return RouteResult(route=Route.IGNORED_MAIL)

        # Has a valid resume — check for duplicate within 90-day window.
        existing_records = lookup_by_email(sender_email)
        if existing_records:
            # Use max(received_date) — NOT sort order — for order-independent detection.
            latest = _max_received(existing_records)
            if latest:
                age_days = (now - latest).days
                if age_days <= DUPLICATE_CHECK_DAYS:
                    dup_record = _latest_record(existing_records)
                    dup_count = int(dup_record.get("Application Updates", 0))
                    if dup_count >= DUPLICATE_NOTICE_MAX:
                        return RouteResult(
                            route=Route.DUPLICATE,
                            existing_record=dup_record,
                            send_email=False,
                            detail=f"dup cap {dup_count}/{DUPLICATE_NOTICE_MAX}",
                        )
                    send = dup_count < REPLY_CAP
                    return RouteResult(
                        route=Route.DUPLICATE,
                        existing_record=dup_record,
                        send_email=send,
                    )

        # New candidate — no record within window.
        return RouteResult(route=Route.NEW_CANDIDATE)


# ─── Helpers ──────────────────────────────────────────────────────────────────

def _parse_dt(value: str) -> datetime | None:
    """Parse an ISO-8601 string (with or without timezone) to an aware datetime."""
    if not value:
        return None
    try:
        from datetime import datetime
        dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt
    except (ValueError, TypeError):
        return None


def _max_received(records: list[dict]) -> datetime | None:
    """Return max(received_date) across records — order-independent duplicate detection."""
    dates = [d for r in records if (d := _parse_dt(r.get("Received Date", "")))]
    return max(dates) if dates else None


def _latest_record(records: list[dict]) -> dict:
    """Return the record with the latest Received Date."""
    def key(r):
        d = _parse_dt(r.get("Received Date", ""))
        return d if d else datetime.min.replace(tzinfo=timezone.utc)
    return max(records, key=key)
