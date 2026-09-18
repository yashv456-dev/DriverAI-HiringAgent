"""
Deterministic Application Reference generator.

Mirrors the P1 Power Automate formula exactly:
    APP-{YYYYMMDD}-{HHMM}-{4 chars of message ID hash}

Config (from flow_config.json appref section):
    date_format:  yyyyMMdd → %Y%m%d
    time_format:  HHmm     → %H%M
    hex_length:   4
    detect_pattern: "app-20"
    id_tail_skip: 2   (skip last 2 chars of internet_message_id before hashing)
"""
from __future__ import annotations

import hashlib
import re
from datetime import datetime, timezone

# ── Detection pattern (case-insensitive) ──────────────────────────────────────
# Matches: APP-20YYMMDD-HHMM-XXXX  (4 alphanumeric chars at end)
_DETECT_PATTERN = re.compile(
    r"\bAPP-20\d{6}-\d{4}-[A-Za-z0-9]{4}\b",
    re.IGNORECASE,
)


def mint_app_ref(received_dt: datetime, internet_message_id: str) -> str:
    """
    Generate a deterministic Application Reference ID.

    Args:
        received_dt: The received timestamp of the email (any timezone; will be
                     converted to MST UTC-7 for the date/time portion, matching P1).
        internet_message_id: The immutable internet Message-ID header value.

    Returns:
        e.g. "APP-20260909-1230-8A1F"
    """
    from zoneinfo import ZoneInfo
    mst = ZoneInfo("America/Phoenix")  # UTC-7, no DST — matches P1 timezone rule
    dt_mst = received_dt.astimezone(mst) if received_dt.tzinfo else received_dt

    date_part = dt_mst.strftime("%Y%m%d")
    time_part = dt_mst.strftime("%H%M")

    # Trim id_tail_skip=2 chars from end, then hash.
    msg_id = internet_message_id.rstrip()
    if len(msg_id) > 2:
        msg_id = msg_id[:-2]  # id_tail_skip = 2
    hash_chars = hashlib.sha256(msg_id.encode()).hexdigest()[:4].upper()

    return f"APP-{date_part}-{time_part}-{hash_chars}"


def extract_quoted_ref(subject: str, body: str) -> str | None:
    """
    Scan subject and body for an APP-Ref token.

    Returns the first match normalized to UPPERCASE, or None.
    Checked in subject first, then body, matching P1 QuotedRef action order.
    """
    for text in (subject, body):
        m = _DETECT_PATTERN.search(text)
        if m:
            return m.group(0).upper()
    return None


def is_valid_app_ref(value: str) -> bool:
    """True if *value* looks like a well-formed APP-Ref."""
    return bool(_DETECT_PATTERN.fullmatch(value.strip()))
