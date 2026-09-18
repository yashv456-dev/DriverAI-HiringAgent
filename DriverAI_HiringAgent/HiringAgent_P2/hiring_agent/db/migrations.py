"""
Versioned SQLite schema management for the DriverAI Hiring Agent.

Design invariants:
- Startup NEVER recreates or empties the database.
- Migrations run in order; each is idempotent.
- Foreign keys are ON for every connection.
- WAL mode + FULL synchronous for durability.
- All writes use short transactions; no transaction spans a Graph/OCR/model call.
"""
from __future__ import annotations

import logging
import os
import sqlite3
from pathlib import Path

logger = logging.getLogger(__name__)

# Current schema version. Bump this and add a _migrate_N() function below each time.
SCHEMA_VERSION = 1


def database_path() -> Path:
    configured = os.getenv("HIRING_SQLITE_PATH")
    if configured:
        return Path(configured).resolve()
    base = Path(__file__).resolve().parent.parent.parent
    return (base / "candidates.db").resolve()


def get_connection(db_path: str | Path | None = None) -> sqlite3.Connection:
    """Return a configured connection. Always call migrate() before first use."""
    path = Path(db_path) if db_path else database_path()
    if str(path) != ":memory:":
        path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(path), timeout=30, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=FULL")
    conn.execute("PRAGMA foreign_keys=ON")
    conn.execute("PRAGMA busy_timeout=10000")
    return conn


def migrate(conn: sqlite3.Connection) -> None:
    """Apply any outstanding schema migrations in order."""
    with conn:
        conn.execute(
            "CREATE TABLE IF NOT EXISTS _schema_version "
            "(version INTEGER NOT NULL DEFAULT 0)"
        )
        row = conn.execute("SELECT version FROM _schema_version").fetchone()
        current = row["version"] if row else 0
        if not row:
            conn.execute("INSERT INTO _schema_version(version) VALUES(0)")

    if current < 1:
        _migrate_1(conn)
        with conn:
            conn.execute("UPDATE _schema_version SET version=1")
        logger.info("db: migrated to schema version 1")


# ─── Migration 1 — full initial schema ────────────────────────────────────────

def _migrate_1(conn: sqlite3.Connection) -> None:
    with conn:
        conn.executescript("""
-- ─── Intake events ────────────────────────────────────────────────────────────
-- One immutable row per real email message received at apply@driverai.io.
-- event_key is deterministic: sha256(mailbox + ":" + internet_message_id)
-- so re-delivery of the same message creates no duplicate.
CREATE TABLE IF NOT EXISTS intake_events (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    event_key           TEXT    UNIQUE NOT NULL,
    schema_version      INTEGER NOT NULL DEFAULT 1,
    -- timestamps are ISO-8601 strings in MST (UTC-7, no DST)
    received_at         TEXT,
    captured_at         TEXT    DEFAULT (strftime('%Y-%m-%dT%H:%M:%S', 'now')),
    sender_email        TEXT    NOT NULL DEFAULT '',
    sender_display      TEXT    NOT NULL DEFAULT '',
    subject             TEXT    NOT NULL DEFAULT '',
    body_preview        TEXT    NOT NULL DEFAULT '',
    -- APP-YYYYMMDD-HHMM-xxxx if a valid reference was quoted in subject/body
    referenced_app_id   TEXT,
    -- JSON array: [{filename, size, drive_id, item_id, sharepoint_path, content_hash}]
    attachment_manifest TEXT    NOT NULL DEFAULT '[]',
    -- pending | complete | conflict
    capture_state       TEXT    NOT NULL DEFAULT 'pending',
    -- spam | bad_sender | bad_subject | new | duplicate | resume_update |
    -- followup | cv_request | wrong_format | ignored
    classification      TEXT,
    raw_headers         TEXT    NOT NULL DEFAULT '{}'
);

-- ─── Applications ─────────────────────────────────────────────────────────────
-- One row per unique Application ID (APP-YYYYMMDD-HHMM-xxxx).
-- Internal UUID is the FK target; external app_id is preserved separately.
CREATE TABLE IF NOT EXISTS applications (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    external_app_id         TEXT    UNIQUE NOT NULL,
    current_input_revision  INTEGER NOT NULL DEFAULT 0,
    current_result_id       INTEGER,   -- FK → result_versions.id (filled after scoring)
    -- P1 status values: 'New Email Received', 'Scored', 'Rejected - Non-USA Location', etc.
    business_status         TEXT    NOT NULL DEFAULT 'New Email Received',
    geo_verdict             TEXT,      -- USA | Non-USA | Pending | NULL
    -- Anti-flood counters (mirrors P1 Excel counters)
    application_updates     INTEGER NOT NULL DEFAULT 0,
    duplicate_notices_sent  INTEGER NOT NULL DEFAULT 0,
    update_notices_sent     INTEGER NOT NULL DEFAULT 0,
    followup_notices_sent   INTEGER NOT NULL DEFAULT 0,
    -- Optimistic locking: callers read this, pass it back, fail if changed
    revision                INTEGER NOT NULL DEFAULT 0,
    created_at              TEXT    DEFAULT (strftime('%Y-%m-%dT%H:%M:%S', 'now')),
    updated_at              TEXT    DEFAULT (strftime('%Y-%m-%dT%H:%M:%S', 'now'))
);

-- ─── Application versions ─────────────────────────────────────────────────────
-- Immutable snapshot of the input data for each version of an application.
-- Every resume submission (new, dup, or update) creates a new version row.
CREATE TABLE IF NOT EXISTS application_versions (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    application_id  INTEGER NOT NULL REFERENCES applications(id),
    event_key       TEXT    NOT NULL REFERENCES intake_events(event_key),
    version_seq     INTEGER NOT NULL,   -- 0 = original, 1 = first update, etc.
    sender_email    TEXT    NOT NULL DEFAULT '',
    full_name       TEXT    NOT NULL DEFAULT '',
    mail_subject    TEXT    NOT NULL DEFAULT '',
    mail_body       TEXT    NOT NULL DEFAULT '',
    -- comma-separated original filenames (matches P1 "Original Filename" column)
    original_filenames TEXT NOT NULL DEFAULT '',
    has_resume      INTEGER NOT NULL DEFAULT 1,  -- 1 = yes, 0 = no
    created_at      TEXT    DEFAULT (strftime('%Y-%m-%dT%H:%M:%S', 'now')),
    UNIQUE(application_id, version_seq)
);

-- ─── Documents ────────────────────────────────────────────────────────────────
-- One row per resume file stored in SharePoint.
-- Linked to a specific application_version so history is never lost.
CREATE TABLE IF NOT EXISTS documents (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    application_version_id  INTEGER NOT NULL REFERENCES application_versions(id),
    drive_id                TEXT,
    item_id                 TEXT,
    sharepoint_path         TEXT    NOT NULL DEFAULT '',
    -- cv_{GUID}.{ext} — collision-free filename
    stored_filename         TEXT    NOT NULL DEFAULT '',
    original_filename       TEXT    NOT NULL DEFAULT '',
    -- hex SHA-256 of raw bytes, filled after download+verify
    content_hash            TEXT,
    size_bytes              INTEGER,
    created_at              TEXT    DEFAULT (strftime('%Y-%m-%dT%H:%M:%S', 'now'))
);

-- ─── Jobs / attempts ──────────────────────────────────────────────────────────
-- Processing queue with worker leases and fencing tokens.
-- Stages: intake | extract | score | geo | publish
CREATE TABLE IF NOT EXISTS jobs (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    application_version_id  INTEGER NOT NULL REFERENCES application_versions(id),
    stage                   TEXT    NOT NULL,
    -- pending | leased | done | failed | skipped
    status                  TEXT    NOT NULL DEFAULT 'pending',
    owner_id                TEXT,           -- worker identity holding the lease
    fencing_token           INTEGER NOT NULL DEFAULT 0,
    retry_count             INTEGER NOT NULL DEFAULT 0,
    next_eligible_at        TEXT,           -- ISO-8601; NULL = eligible now
    error_class             TEXT,
    error_detail            TEXT,
    created_at              TEXT    DEFAULT (strftime('%Y-%m-%dT%H:%M:%S', 'now')),
    updated_at              TEXT    DEFAULT (strftime('%Y-%m-%dT%H:%M:%S', 'now'))
);

-- ─── Result versions ──────────────────────────────────────────────────────────
-- Immutable scored outputs. One row per complete scoring run.
CREATE TABLE IF NOT EXISTS result_versions (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    application_id      INTEGER NOT NULL REFERENCES applications(id),
    job_id              INTEGER REFERENCES jobs(id),
    -- Raw JSON blobs (never modified after insert)
    extraction_raw      TEXT    NOT NULL DEFAULT '{}',
    ai_recheck_raw      TEXT    NOT NULL DEFAULT '{}',
    -- JSON of the 22 P2-owned columns
    final_fields        TEXT    NOT NULL DEFAULT '{}',
    -- JSON: {"Senior AI/ML Engineer": 88, ...}
    scores              TEXT    NOT NULL DEFAULT '{}',
    top_role            TEXT,
    top_score           REAL,
    geo_decision        TEXT,   -- USA | Non-USA | Needs Review
    model_id            TEXT,   -- e.g. "qwen3:1.7b"
    jd_snapshot_hash    TEXT,   -- FK → jd_snapshots.hash
    created_at          TEXT    DEFAULT (strftime('%Y-%m-%dT%H:%M:%S', 'now'))
);

-- ─── JD snapshots ─────────────────────────────────────────────────────────────
-- Immutable, content-addressed snapshot of the exact JD set used for a run.
-- Prevents mid-batch role-set drift.
CREATE TABLE IF NOT EXISTS jd_snapshots (
    hash        TEXT    PRIMARY KEY,    -- sha256 of payload
    role_count  INTEGER NOT NULL DEFAULT 0,
    payload     TEXT    NOT NULL,       -- full JSON of all JDs
    created_at  TEXT    DEFAULT (strftime('%Y-%m-%dT%H:%M:%S', 'now'))
);

-- ─── Audit events ─────────────────────────────────────────────────────────────
-- Append-only log for compliance tracing (EU AI Act Art. 12).
CREATE TABLE IF NOT EXISTS audit_events (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_type TEXT    NOT NULL DEFAULT '',   -- application | job | result | intake_event
    entity_id   INTEGER,
    event_type  TEXT    NOT NULL DEFAULT '',   -- created | scored | rejected | exported | …
    actor       TEXT    NOT NULL DEFAULT '',   -- worker_id or 'system'
    detail      TEXT    NOT NULL DEFAULT '{}', -- JSON
    created_at  TEXT    DEFAULT (strftime('%Y-%m-%dT%H:%M:%S', 'now'))
);

-- ─── Export jobs ──────────────────────────────────────────────────────────────
-- Tracks Excel / SharePoint publication runs.
CREATE TABLE IF NOT EXISTS export_jobs (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    trigger         TEXT    NOT NULL DEFAULT 'manual',  -- manual | scheduled | api
    -- pending | running | done | failed
    status          TEXT    NOT NULL DEFAULT 'pending',
    rows_exported   INTEGER,
    error_detail    TEXT,
    started_at      TEXT,
    finished_at     TEXT,
    created_at      TEXT    DEFAULT (strftime('%Y-%m-%dT%H:%M:%S', 'now'))
);

-- ─── Import conflicts ─────────────────────────────────────────────────────────
-- Holds Excel rows that couldn't be cleanly mapped (duplicate IDs, blank rows, etc.)
-- during import_excel.py. These are NEVER silently discarded.
CREATE TABLE IF NOT EXISTS import_conflicts (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    app_id      TEXT    NOT NULL DEFAULT '',
    reason      TEXT    NOT NULL,   -- duplicate_id | blank_row | missing_app_id | identity_mismatch
    payload     TEXT    NOT NULL,   -- JSON of the raw Excel row
    resolved    INTEGER NOT NULL DEFAULT 0,
    created_at  TEXT    DEFAULT (strftime('%Y-%m-%dT%H:%M:%S', 'now'))
);

-- ─── Indexes ──────────────────────────────────────────────────────────────────
CREATE INDEX IF NOT EXISTS idx_intake_events_sender
    ON intake_events(sender_email);

CREATE INDEX IF NOT EXISTS idx_applications_status
    ON applications(business_status);

CREATE INDEX IF NOT EXISTS idx_applications_updated
    ON applications(updated_at DESC);

CREATE INDEX IF NOT EXISTS idx_jobs_status_stage
    ON jobs(status, stage);

CREATE INDEX IF NOT EXISTS idx_jobs_eligible
    ON jobs(status, next_eligible_at)
    WHERE status = 'pending';

CREATE INDEX IF NOT EXISTS idx_result_versions_app
    ON result_versions(application_id, created_at DESC);

CREATE INDEX IF NOT EXISTS idx_audit_entity
    ON audit_events(entity_type, entity_id, created_at DESC);
""")
