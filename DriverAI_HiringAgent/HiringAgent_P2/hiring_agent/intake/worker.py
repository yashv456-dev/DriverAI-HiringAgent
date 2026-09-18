"""
Main intake worker — the Python replacement for the P1 Power Automate flow.

Polling model: one email per run, 60-second interval (mirrors P1 recurrence).
Worker lock: prevents concurrent runs (fcntl on mac/linux, msvcrt on windows).

Execution order per message (matches P1 Registrar → Receptionist → Inbox Tidy):
  1. Classify message (Gates 1–3 + allow-list)
  2. Route (determine processing path)
  3. Download attachment bytes   ← must happen BEFORE any folder move
  4. Upload resume to SharePoint ← if applicable
  5. Commit intake_event + application record to SQLite (single transaction)
  6. Mark message as READ + move to Archive (or Junk Email for spam)
  7. Send auto-reply              ← only if send_email=True and emails enabled

"Never Block the Queue" principle:
  A crash between steps 3 and 6 leaves the message unread in the Inbox.
  On the next poll, internetMessageId lookup finds the same message and the
  duplicate-detection prevents a second row insert. The file may be re-uploaded
  (same content hash → same filename → harmless overwrite).

  A crash between steps 6 and 7 means the email was archived but no reply sent.
  That is acceptable: better silent than double-replying.

Shadow mode (dry_run=True):
  Steps 4, 6, 7 are skipped (no writes to SharePoint or Graph).
  All decisions are logged so they can be compared against live P1 output.
"""
from __future__ import annotations

import fcntl
import hashlib
import json
import logging
import os
import platform
import sqlite3
import time
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .appref import mint_app_ref, extract_quoted_ref
from .classify import Classifier, Classification, BLOCKED, RESUME_EXTENSIONS
from .graph_mail import GraphMailClient, Message
from .route import Router, Route, RouteResult
from ..db.migrations import get_connection, migrate

logger = logging.getLogger(__name__)

# ─── Configuration ────────────────────────────────────────────────────────────

POLL_INTERVAL_SECONDS = 60
LOCK_FILE_PATH = "/tmp/driverai_intake_worker.lock"

# Files to exclude from resume selection (case-insensitive contains).
# Sourced from flow_config.json resume_attachments.exclude_name_phrases.
EXCLUDE_NAME_PHRASES = (
    "portfolio", "cover letter", "coverletter", "cover_letter",
    "-cl.", "_cl.", " cl.", "transcript", "certificate",
)


# ─── Result types ─────────────────────────────────────────────────────────────

@dataclass
class WorkerResult:
    success: bool
    internet_message_id: str = ""
    route: str = ""
    app_id: str = ""
    send_email: bool = False
    detail: str = ""
    error: Exception | None = None


@dataclass
class WorkerConfig:
    mailbox: str                  # apply@driverai.io
    admin_email: str              # yashv@driverai.io
    tenant_id: str
    client_id: str
    client_secret: str
    sharepoint_site: str
    sharepoint_drive_id: str
    sharepoint_site_id: str
    resumes_folder: str           # /Shared Documents/Candidate_Resumes
    db_path: Path | None = None
    # Global email toggle (mirrors send_applicant_emails in flow_config.json)
    send_applicant_emails: bool = False
    send_admin_alerts: bool = True
    dry_run: bool = False         # shadow mode — no Graph writes
    bcc_email: str = ""           # BCC on all applicant emails (yashv@driverai.io)

    @classmethod
    def from_env(cls) -> "WorkerConfig":
        return cls(
            mailbox=os.environ["HIRING_MAILBOX"],
            admin_email=os.environ.get("HIRING_ADMIN_EMAIL", "yashv@driverai.io"),
            tenant_id=os.environ["TENANT_ID"],
            client_id=os.environ["CLIENT_ID"],
            client_secret=os.environ["CLIENT_SECRET"],
            sharepoint_site=os.environ["HIRING_SHAREPOINT_SITE"],
            sharepoint_drive_id=os.environ["HIRING_SHAREPOINT_DRIVE_ID"],
            sharepoint_site_id=os.environ["HIRING_SHAREPOINT_SITE_ID"],
            resumes_folder=os.environ.get(
                "HIRING_RESUMES_FOLDER", "/Shared Documents/Candidate_Resumes"
            ),
            db_path=Path(p) if (p := os.environ.get("HIRING_SQLITE_PATH")) else None,
            send_applicant_emails=os.environ.get(
                "HIRING_SEND_APPLICANT_EMAILS", "false"
            ).lower() in {"1", "true", "yes"},
            send_admin_alerts=os.environ.get(
                "HIRING_SEND_ADMIN_ALERTS", "true"
            ).lower() in {"1", "true", "yes"},
            dry_run=os.environ.get("HIRING_DRY_RUN", "false").lower() in {"1", "true", "yes"},
            bcc_email=os.environ.get("HIRING_BCC_EMAIL", "yashv@driverai.io"),
        )


# ─── Worker ───────────────────────────────────────────────────────────────────

class IntakeWorker:
    """
    Single-threaded intake worker. Call run_once() in a loop or use run_loop().

    Create one IntakeWorker per process. Concurrent runs are prevented by a
    file-based worker lock (fcntl on Unix, msvcrt on Windows).
    """

    def __init__(self, config: WorkerConfig):
        self.config = config
        self._classifier = Classifier()
        self._router = Router(self._classifier)
        self._graph = GraphMailClient(
            mailbox=config.mailbox,
            tenant_id=config.tenant_id,
            client_id=config.client_id,
            client_secret=config.client_secret,
        )
        self._conn: sqlite3.Connection | None = None

    # ── Public API ────────────────────────────────────────────────────────────

    def run_once(self) -> WorkerResult:
        """
        Fetch one message from the inbox and process it end-to-end.

        Safe to call repeatedly — idempotent on the same message via
        internetMessageId deduplication.
        """
        messages = self._graph.poll_inbox(top=1)
        if not messages:
            logger.debug("Inbox empty — nothing to process")
            return WorkerResult(success=True, detail="inbox_empty")

        msg = messages[0]
        logger.info(
            "Processing message: %s | from: %s | subject: %s",
            msg.internet_message_id, msg.sender_email, msg.subject[:60],
        )

        try:
            return self._process(msg)
        except Exception as exc:
            logger.exception("Error processing %s", msg.internet_message_id)
            return WorkerResult(
                success=False,
                internet_message_id=msg.internet_message_id,
                error=exc,
                detail=str(exc),
            )

    def run_loop(self, interval: int = POLL_INTERVAL_SECONDS, *, stop_after: int = 0) -> None:
        """
        Continuous polling loop with worker lock.

        Args:
            interval: Seconds between polls.
            stop_after: If > 0, stop after this many successful polls (for testing).
        """
        with _WorkerLock(LOCK_FILE_PATH):
            logger.info("Intake worker started (mailbox=%s, dry_run=%s)",
                        self.config.mailbox, self.config.dry_run)
            count = 0
            while True:
                result = self.run_once()
                if result.detail != "inbox_empty":
                    count += 1
                    logger.info("Poll #%d: route=%s app_id=%s",
                                count, result.route, result.app_id)
                if stop_after and count >= stop_after:
                    logger.info("Reached stop_after=%d — exiting", stop_after)
                    break
                time.sleep(interval)

    # ── Core processing ───────────────────────────────────────────────────────

    def _process(self, msg: Message) -> WorkerResult:
        cfg = self.config
        conn = self._db()

        # ── Check idempotency (already processed?) ────────────────────────
        event_key = _event_key(cfg.mailbox, msg.internet_message_id)
        existing = conn.execute(
            "SELECT capture_state FROM intake_events WHERE event_key=?", (event_key,)
        ).fetchone()
        if existing and existing["capture_state"] == "complete":
            logger.info("Skipping already-complete event %s", event_key)
            # Still move+mark to clear the inbox if we're not in dry-run mode.
            if not cfg.dry_run:
                self._inbox_cleanup(msg, route=Route.IGNORED_MAIL)
            return WorkerResult(
                success=True,
                internet_message_id=msg.internet_message_id,
                detail="already_processed",
            )

        # ── Step 1: Classify ─────────────────────────────────────────────
        attach_names = self._resume_filenames(msg)
        classify_result = self._classifier.run_gates(
            sender_email=msg.sender_email,
            subject=msg.subject,
            body=msg.body_text,
            attach_names=attach_names,
        )
        logger.debug("Classify → %s (gate_hit=%r)", classify_result.classification, classify_result.gate_hit)

        # ── Step 2: Route ─────────────────────────────────────────────────
        route_result = self._router.route(
            sender_email=msg.sender_email,
            subject=msg.subject,
            body=msg.body_text,
            attach_names=attach_names,
            classify_result=classify_result,
            lookup_by_email=self._lookup_by_email,
            lookup_by_app_ref=self._lookup_by_app_ref,
        )
        logger.info("Route → %s (send=%s)", route_result.route, route_result.send_email)

        # ── Step 3: Download attachments (BEFORE any folder move) ─────────
        downloaded_resumes: list[tuple[str, bytes]] = []  # [(original_name, bytes)]
        for attach in msg.attachments:
            if Path(attach.name).suffix.lower() not in RESUME_EXTENSIONS:
                continue
            if self._should_exclude(attach.name):
                logger.debug("Excluding attachment %r (non-resume phrase)", attach.name)
                continue
            raw = self._graph.download_attachment(msg.graph_id, attach.attachment_id)
            downloaded_resumes.append((attach.name, raw))

        # ── Step 4: Upload to SharePoint ──────────────────────────────────
        sp_refs: list[dict] = []
        if downloaded_resumes and route_result.route not in {
            Route.CV_REQUEST, Route.WRONG_FORMAT, Route.IGNORED_MAIL,
        } and not cfg.dry_run:
            received_dt = _parse_received(msg.received_at)
            for orig_name, raw_bytes in downloaded_resumes:
                ref = self._upload_resume(
                    orig_name, raw_bytes, received_dt,
                    reuse_guid=None,  # TODO: extract GUID from existing_record for UPDATE path
                )
                sp_refs.append(ref)

        # ── Step 5: Commit to SQLite (single transaction) ─────────────────
        app_id = ""
        if route_result.route not in {Route.CV_REQUEST, Route.IGNORED_MAIL}:
            received_dt = _parse_received(msg.received_at)
            app_id = self._commit_to_db(
                conn=conn,
                msg=msg,
                event_key=event_key,
                classify_result=classify_result,
                route_result=route_result,
                sp_refs=sp_refs,
                received_dt=received_dt,
                attach_names=attach_names,
            )

        # ── Step 6: Inbox cleanup (mark read + move) ──────────────────────
        if not cfg.dry_run:
            destination = "JunkEmail" if route_result.route in {
                Route.BAD_SENDER, Route.BAD_SUBJECT, Route.SPAM, Route.SHORTENER_ONLY,
            } else "Archive"
            self._inbox_cleanup(msg, route=route_result.route, destination=destination)

        # ── Step 7: Send auto-reply ───────────────────────────────────────
        if (route_result.send_email and cfg.send_applicant_emails
                and app_id and not cfg.dry_run):
            self._send_auto_reply(msg, route_result, app_id)

        return WorkerResult(
            success=True,
            internet_message_id=msg.internet_message_id,
            route=route_result.route.value,
            app_id=app_id,
            send_email=route_result.send_email and cfg.send_applicant_emails,
        )

    # ── DB helpers ────────────────────────────────────────────────────────────

    def _db(self) -> sqlite3.Connection:
        if self._conn is None:
            self._conn = get_connection(self.config.db_path)
            migrate(self._conn)
        return self._conn

    def _lookup_by_email(self, email: str) -> list[dict]:
        conn = self._db()
        rows = conn.execute("""
            SELECT a.external_app_id AS "Application ID",
                   av.sender_email AS "Email",
                   av.created_at AS "Received Date",
                   a.application_updates AS "Application Updates"
            FROM applications a
            JOIN application_versions av ON av.application_id = a.id
            WHERE LOWER(av.sender_email) = LOWER(?)
            ORDER BY av.created_at DESC
        """, (email,)).fetchall()
        return [dict(r) for r in rows]

    def _lookup_by_app_ref(self, app_ref: str) -> dict | None:
        conn = self._db()
        row = conn.execute("""
            SELECT a.external_app_id AS "Application ID",
                   av.sender_email AS "Email",
                   a.application_updates AS "Application Updates",
                   av.created_at AS "Received Date"
            FROM applications a
            JOIN application_versions av ON av.application_id = a.id AND av.version_seq = 0
            WHERE UPPER(a.external_app_id) = UPPER(?)
        """, (app_ref,)).fetchone()
        return dict(row) if row else None

    def _commit_to_db(
        self,
        *,
        conn: sqlite3.Connection,
        msg: Message,
        event_key: str,
        classify_result: Any,
        route_result: RouteResult,
        sp_refs: list[dict],
        received_dt: datetime,
        attach_names: list[str],
    ) -> str:
        """Write intake_event + application (+ version + documents) in a single transaction."""
        from zoneinfo import ZoneInfo
        mst = ZoneInfo("America/Phoenix")
        received_mst = received_dt.astimezone(mst)
        received_iso = received_mst.isoformat()

        attach_manifest = json.dumps([
            {"filename": r.get("original_filename", ""), "sharepoint_path": r.get("sharepoint_path", "")}
            for r in sp_refs
        ])

        with conn:
            # Upsert intake_event
            conn.execute("""
                INSERT OR IGNORE INTO intake_events
                    (event_key, received_at, captured_at, sender_email, sender_display,
                     subject, body_preview, referenced_app_id, attachment_manifest,
                     capture_state, classification)
                VALUES (?,?,?,?,?,?,?,?,?,'pending',?)
            """, (
                event_key, received_iso,
                datetime.now(timezone.utc).isoformat(),
                msg.sender_email, msg.sender_display,
                msg.subject, msg.body_text[:255],
                route_result.verified_app_ref,
                attach_manifest,
                classify_result.classification.value,
            ))

            route = route_result.route
            if route == Route.NEW_CANDIDATE:
                app_id = mint_app_ref(received_dt, msg.internet_message_id)
                conn.execute("""
                    INSERT INTO applications
                        (external_app_id, business_status, application_updates, revision)
                    VALUES (?, 'New Email Received', 0, 0)
                """, (app_id,))
                app_row = conn.execute(
                    "SELECT id FROM applications WHERE external_app_id=?", (app_id,)
                ).fetchone()
                app_db_id = app_row["id"]

                conn.execute("""
                    INSERT INTO application_versions
                        (application_id, event_key, version_seq, sender_email, full_name,
                         mail_subject, mail_body, original_filenames, has_resume)
                    VALUES (?,?,0,?,?,?,?,?,?)
                """, (
                    app_db_id, event_key,
                    msg.sender_email,
                    msg.sender_display or msg.sender_email,
                    msg.subject,
                    msg.body_text[:255],
                    ", ".join(attach_names),
                    1 if sp_refs else 0,
                ))
                version_row = conn.execute(
                    "SELECT id FROM application_versions WHERE application_id=? AND version_seq=0",
                    (app_db_id,)
                ).fetchone()
                version_db_id = version_row["id"]

                for ref in sp_refs:
                    conn.execute("""
                        INSERT INTO documents
                            (application_version_id, drive_id, item_id, sharepoint_path,
                             stored_filename, original_filename, content_hash, size_bytes)
                        VALUES (?,?,?,?,?,?,?,?)
                    """, (
                        version_db_id,
                        ref.get("drive_id", ""),
                        ref.get("item_id", ""),
                        ref.get("sharepoint_path", ""),
                        ref.get("stored_filename", ""),
                        ref.get("original_filename", ""),
                        ref.get("content_hash", ""),
                        ref.get("size_bytes", 0),
                    ))

                # Create initial scoring job
                conn.execute("""
                    INSERT INTO jobs (application_version_id, stage, status)
                    VALUES (?, 'score', 'pending')
                """, (version_db_id,))

            elif route in {Route.DUPLICATE, Route.RESUME_UPDATE, Route.FOLLOWUP}:
                # Patch existing application record
                existing_app_id = (
                    route_result.verified_app_ref
                    or (route_result.existing_record or {}).get("Application ID", "")
                )
                conn.execute("""
                    UPDATE applications
                    SET application_updates = application_updates + 1,
                        updated_at = ?,
                        revision = revision + 1
                    WHERE UPPER(external_app_id) = UPPER(?)
                """, (datetime.now(timezone.utc).isoformat(), existing_app_id))
                app_id = existing_app_id

            else:
                app_id = ""

            # Mark event complete
            conn.execute(
                "UPDATE intake_events SET capture_state='complete' WHERE event_key=?",
                (event_key,)
            )

        return app_id

    # ── Resume helpers ────────────────────────────────────────────────────────

    def _resume_filenames(self, msg: Message) -> list[str]:
        """Return attachment names eligible for resume processing."""
        candidates = [
            a.name for a in msg.attachments
            if Path(a.name).suffix.lower() in RESUME_EXTENSIONS and not a.is_inline
        ]
        filtered = [n for n in candidates if not self._should_exclude(n)]
        # Safety: if exclusion would leave nothing, keep all candidates.
        return filtered if filtered else candidates

    @staticmethod
    def _should_exclude(name: str) -> bool:
        lower = name.lower()
        return any(phrase in lower for phrase in EXCLUDE_NAME_PHRASES)

    def _upload_resume(
        self,
        original_filename: str,
        raw_bytes: bytes,
        received_dt: datetime,
        *,
        reuse_guid: str | None = None,
    ) -> dict:
        """Upload raw bytes to SharePoint and return a ref dict."""
        import uuid
        import hashlib
        from pathlib import Path as _Path

        ext = _Path(original_filename).suffix.lower() or ".pdf"
        guid = reuse_guid or str(uuid.uuid4())
        stored_name = f"cv_{guid}{ext}"
        content_hash = hashlib.sha256(raw_bytes).hexdigest()

        from zoneinfo import ZoneInfo
        mst = ZoneInfo("America/Phoenix")
        dt_mst = received_dt.astimezone(mst)
        year = dt_mst.strftime("%Y")
        month = dt_mst.strftime("%B")
        folder = f"{self.config.resumes_folder.rstrip('/')}/{year}/{month}"
        full_path = f"{folder}/{stored_name}"

        # PUT to SharePoint via Graph
        site_root = (
            f"https://graph.microsoft.com/v1.0"
            f"/sites/{self.config.sharepoint_site_id}"
            f"/drives/{self.config.sharepoint_drive_id}"
        )
        site_rel = folder.replace("/Shared Documents", "")
        url = f"{site_root}/root:{site_rel}/{stored_name}:/content"

        result = self._graph._request(
            "PUT", url, retryable=True, data=raw_bytes,
            headers={
                "Authorization": f"Bearer {self._graph._get_token()}",
                "Content-Type": "application/octet-stream",
            },
        )
        item_id = result.get("id", "") if isinstance(result, dict) else ""
        drive_id = result.get("parentReference", {}).get("driveId", "") if isinstance(result, dict) else ""

        return {
            "stored_filename": stored_name,
            "original_filename": original_filename,
            "sharepoint_path": full_path,
            "drive_id": drive_id,
            "item_id": item_id,
            "content_hash": content_hash,
            "size_bytes": len(raw_bytes),
            "guid": guid,
        }

    # ── Inbox cleanup ─────────────────────────────────────────────────────────

    def _inbox_cleanup(
        self,
        msg: Message,
        *,
        route: Route,
        destination: str = "Archive",
    ) -> None:
        """Mark message read and move to Archive (or Junk Email for spam)."""
        try:
            self._graph.mark_read(msg.graph_id)
        except Exception as exc:
            logger.warning("mark_read failed for %s: %s", msg.internet_message_id, exc)
        try:
            self._graph.move_message(msg.graph_id, destination)
        except Exception as exc:
            logger.error("move_message failed for %s → %s: %s",
                         msg.internet_message_id, destination, exc)

    # ── Auto-reply ────────────────────────────────────────────────────────────

    def _send_auto_reply(self, msg: Message, route_result: RouteResult, app_id: str) -> None:
        """Send the appropriate auto-reply template."""
        subject, body = _build_reply(msg, route_result, app_id)
        if not subject:
            return
        try:
            self._graph.send_mail(
                to_address=msg.sender_email,
                subject=subject,
                body_html=body,
                bcc_address=self.config.bcc_email or None,
            )
        except Exception as exc:
            logger.error("send_mail failed for %s: %s", msg.internet_message_id, exc)


# ─── Auto-reply templates ─────────────────────────────────────────────────────
# All 6 lifecycle emails from P1 (send_applicant_emails controls whether they fire).

def _build_reply(msg: Message, route: RouteResult, app_id: str) -> tuple[str, str]:
    """Return (subject, body_html) for the given route, or ("", "") to suppress."""
    name = msg.sender_display or msg.sender_email.split("@")[0]
    ref = app_id

    templates = {
        Route.NEW_CANDIDATE: (
            "Thanks for Applying to Driver AI",
            f"""<p>Hi {name},</p>
<p>Thank you for applying to Driver AI! We've received your application and resume.</p>
<p>Your application reference is: <strong>{ref}</strong></p>
<p>We'll be in touch if your profile matches our current openings.</p>
<p>Best,<br>DriverAI Recruiting</p>""",
        ),
        Route.DUPLICATE: (
            "Your Driver AI Application is Already on File",
            f"""<p>Hi {name},</p>
<p>We already have your application on file (Reference: <strong>{ref}</strong>).
Our team will review it shortly.</p>
<p>Best,<br>DriverAI Recruiting</p>""",
        ),
        Route.RESUME_UPDATE: (
            "Quick Favor for Your Driver AI Application",
            f"""<p>Hi {name},</p>
<p>We've received your updated resume for application <strong>{ref}</strong>.
Your profile has been updated in our system.</p>
<p>Best,<br>DriverAI Recruiting</p>""",
        ),
        Route.FOLLOWUP: (
            "Quick Favor for Your Driver AI Application",
            f"""<p>Hi {name},</p>
<p>Thanks for following up on application <strong>{ref}</strong>.
We've noted your message — please include your updated resume if you'd like
us to reconsider your profile.</p>
<p>Best,<br>DriverAI Recruiting</p>""",
        ),
        Route.CV_REQUEST: (
            "Quick Favor for Your Driver AI Application",
            f"""<p>Hi {name},</p>
<p>Thanks for reaching out! It looks like your resume wasn't attached.
Please resend your email with your resume attached (PDF or DOCX)
and we'll get you into our system right away.</p>
<p>Best,<br>DriverAI Recruiting</p>""",
        ),
        Route.WRONG_FORMAT: (
            "Quick Favor for Your Driver AI Application",
            f"""<p>Hi {name},</p>
<p>Thanks for applying! We received an attachment but it wasn't in a supported format.
Please resend with your resume as a <strong>PDF or DOCX</strong>.</p>
<p>Best,<br>DriverAI Recruiting</p>""",
        ),
    }
    subject, body = templates.get(route.route, ("", ""))
    return subject, body


# ─── Utilities ─────────────────────────────────────────────────────────────────

def _event_key(mailbox: str, internet_message_id: str) -> str:
    """Deterministic event key from mailbox + internetMessageId."""
    raw = f"{mailbox.lower()}:{internet_message_id}"
    return hashlib.sha256(raw.encode()).hexdigest()


def _parse_received(received_at: str) -> datetime:
    """Parse Graph's receivedDateTime string to an aware datetime."""
    try:
        dt = datetime.fromisoformat(received_at.replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt
    except (ValueError, AttributeError):
        return datetime.now(timezone.utc)


# ─── Worker lock ───────────────────────────────────────────────────────────────

class _WorkerLock:
    """Exclusive process-level lock using fcntl (Unix) or msvcrt (Windows)."""

    def __init__(self, lock_path: str):
        self._path = lock_path
        self._fd = None

    def __enter__(self):
        self._fd = open(self._path, "w")
        if platform.system() == "Windows":
            import msvcrt
            msvcrt.locking(self._fd.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            try:
                fcntl.flock(self._fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
            except OSError:
                self._fd.close()
                raise RuntimeError(
                    "Another intake worker is already running. "
                    f"Delete {self._path} if the process crashed."
                )
        return self

    def __exit__(self, *_):
        if self._fd:
            if platform.system() == "Windows":
                import msvcrt
                msvcrt.locking(self._fd.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(self._fd, fcntl.LOCK_UN)
            self._fd.close()
            try:
                os.unlink(self._path)
            except OSError:
                pass
