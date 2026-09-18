"""
Tests for the Python P1 intake engine.

Covers:
  - All 3 security gates (exact filter list coverage)
  - Allow-list exact-equality semantics
  - All 6 routing paths
  - APP-Ref minting and detection
  - Anti-flood caps
  - IsKnownSender guard
  - 90-day duplicate window (order-independent)
  - Link-shortener detection
  - Malware extension trailing-space matching
  - DB schema creation and idempotent migration

Run: python -m pytest tests/test_intake.py -v
"""
from __future__ import annotations

import hashlib
import sqlite3
import sys
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path
from unittest.mock import MagicMock, patch

# Ensure project root is on sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from hiring_agent.intake.classify import (
    Classifier, Classification, BLOCKED,
    ALLOW_SENDERS, BAD_SENDERS, BAD_SUBJECTS, SPAM_PHRASES, LINK_SHORTENER_PHRASES,
    RESUME_EXTENSIONS,
)
from hiring_agent.intake.appref import mint_app_ref, extract_quoted_ref, is_valid_app_ref
from hiring_agent.intake.route import Router, Route, RouteResult, DUPLICATE_CHECK_DAYS
from hiring_agent.db.migrations import get_connection, migrate


# ─── Helpers ──────────────────────────────────────────────────────────────────

def _classify(sender="good@example.com", subject="Application for ML Engineer",
              body="Please find my resume attached.", attach_names=("resume.pdf",)):
    return Classifier().run_gates(sender, subject, body, list(attach_names))


def _no_records(_email):
    return []


def _no_app_ref(_ref):
    return None


# One clock for the router and the fixtures. Fixtures build "N days ago" lazily
# inside route(), after it has read the clock; with a live clock 91 days came out
# as 90d 23:59:59.999 and the boundary test flapped.
_NOW = datetime.now(timezone.utc)


def _route(sender="good@example.com", subject="Application for ML Engineer",
           body="Please find my resume attached.", attach_names=("resume.pdf",),
           lookup_by_email=_no_records, lookup_by_app_ref=_no_app_ref, now=_NOW):
    c = Classifier()
    r = Router(c)
    cr = c.run_gates(sender, subject, body, list(attach_names))
    return r.route(
        sender_email=sender,
        subject=subject,
        body=body,
        attach_names=list(attach_names),
        classify_result=cr,
        lookup_by_email=lookup_by_email,
        lookup_by_app_ref=lookup_by_app_ref,
        now=now,
    )


def _in_window_record(email="good@example.com", app_id="APP-20260101-1200-ABCD"):
    recent = (_NOW - timedelta(days=10)).isoformat()
    return [{"Application ID": app_id, "Email": email, "Received Date": recent,
             "Application Updates": 0}]


def _outside_window_record(email="good@example.com"):
    old = (_NOW - timedelta(days=91)).isoformat()
    return [{"Application ID": "APP-20260101-1200-ABCD", "Email": email,
             "Received Date": old, "Application Updates": 0}]


# ─── Gate 1: Bad Senders ──────────────────────────────────────────────────────

class TestGate1Sender(unittest.TestCase):

    def test_clean_sender_passes(self):
        r = _classify(sender="jane.doe@example.com")
        self.assertNotIn(r.classification, BLOCKED)

    def test_noreply_blocked(self):
        r = _classify(sender="noreply@platform.com")
        self.assertEqual(r.classification, Classification.BAD_SENDER)
        self.assertIn("noreply", r.gate_hit)

    def test_no_reply_with_dash_blocked(self):
        r = _classify(sender="no-reply@system.io")
        self.assertEqual(r.classification, Classification.BAD_SENDER)

    def test_self_loop_guard(self):
        r = _classify(sender="apply@driverai.io")
        self.assertEqual(r.classification, Classification.BAD_SENDER)

    def test_mailer_daemon(self):
        r = _classify(sender="mailer-daemon@mail.example.com")
        self.assertEqual(r.classification, Classification.BAD_SENDER)

    def test_linkedin_blocked(self):
        r = _classify(sender="jobs@linkedin.com")
        self.assertEqual(r.classification, Classification.BAD_SENDER)

    def test_all_43_bad_senders_present(self):
        self.assertEqual(len(BAD_SENDERS), 43, "BAD_SENDERS must have exactly 43 entries")

    def test_bacancy_vendor_blocked(self):
        r = _classify(sender="hr@bacancy.com")
        self.assertEqual(r.classification, Classification.BAD_SENDER)


# ─── Allow-list (exact equality) ──────────────────────────────────────────────

class TestAllowList(unittest.TestCase):

    def test_internal_sender_allow_listed(self):
        r = _classify(sender="yashv@driverai.io", subject="test self-send")
        self.assertEqual(r.classification, Classification.ALLOW_LISTED)

    def test_allow_list_exact_equality(self):
        # Substring of allow-listed address must NOT match (security fix 2026-09-08).
        r = _classify(sender="yashv@driverai.io.example.net")
        self.assertNotEqual(r.classification, Classification.ALLOW_LISTED)

    def test_allow_list_bypasses_gate1(self):
        # yashv@ is also a substring of nothing in BAD_SENDERS; test bypass regardless.
        r = _classify(sender="tracys@driverai.io", subject="out of office")
        # Should pass all gates (allow-listed before subject gate).
        self.assertEqual(r.classification, Classification.ALLOW_LISTED)

    def test_apply_not_in_allow_list(self):
        self.assertNotIn("apply@driverai.io", ALLOW_SENDERS)

    def test_exactly_4_allow_senders(self):
        self.assertEqual(len(ALLOW_SENDERS), 4)


# ─── Gate 2: Bad Subjects ─────────────────────────────────────────────────────

class TestGate2Subject(unittest.TestCase):

    def test_clean_subject_passes(self):
        r = _classify(subject="Senior AI Engineer Application")
        self.assertNotIn(r.classification, BLOCKED)

    def test_out_of_office_blocked(self):
        r = _classify(subject="Out of Office: Jane Doe")
        self.assertEqual(r.classification, Classification.BAD_SUBJECT)

    def test_auto_reply_subject_blocked(self):
        r = _classify(subject="Thanks for Applying to Driver AI")
        self.assertEqual(r.classification, Classification.BAD_SUBJECT)

    def test_p1_loop_subject_1(self):
        r = _classify(subject="Your Driver AI Application is Already on File")
        self.assertEqual(r.classification, Classification.BAD_SUBJECT)

    def test_p1_loop_subject_2(self):
        r = _classify(subject="Quick Favor for Your Driver AI Application")
        self.assertEqual(r.classification, Classification.BAD_SUBJECT)

    def test_invoice_attached_blocked(self):
        r = _classify(subject="Invoice Attached - Q3 2026")
        self.assertEqual(r.classification, Classification.BAD_SUBJECT)

    def test_all_27_bad_subjects_present(self):
        self.assertEqual(len(BAD_SUBJECTS), 27)


# ─── Gate 3: Spam Phrases ─────────────────────────────────────────────────────

class TestGate3Spam(unittest.TestCase):

    def test_clean_body_passes(self):
        r = _classify(body="I am applying for the Data Engineer role. Please see my resume.")
        self.assertNotIn(r.classification, BLOCKED)

    def test_phishing_phrase(self):
        r = _classify(body="Please click here to claim your reward.")
        self.assertEqual(r.classification, Classification.SPAM)

    def test_offensive_phrase(self):
        r = _classify(body="I will kill you if you don't reply.")
        self.assertEqual(r.classification, Classification.SPAM)

    def test_malware_extension_exe_blocked(self):
        # Trailing space is critical — ".exe " should match "invoice.exe " (attach+space)
        r = _classify(body="Here is my application.", attach_names=("resume.pdf", "setup.exe"))
        self.assertEqual(r.classification, Classification.SPAM)
        self.assertIn(".exe ", r.gate_hit)

    def test_malware_extension_no_false_positive(self):
        # ".exe" NOT followed by space (inside URL) should NOT match
        c = Classifier()
        result = c.gate3_spam("please visit www.exeter.ac.uk for info", "", [])
        self.assertIsNone(result, "www.exeter.ac.uk should not trigger .exe malware gate")

    def test_vendor_solicitation_blocked(self):
        r = _classify(body="We work as a software development agency. Company profile and corporate deck attached.")
        self.assertEqual(r.classification, Classification.SPAM)

    def test_foreign_scam_spanish(self):
        r = _classify(body="Usted ha ganado la lotería. Transferencia de dinero pendiente.")
        self.assertEqual(r.classification, Classification.SPAM)

    def test_foreign_scam_russian(self):
        r = _classify(body="Вы выиграли крупный приз.")
        self.assertEqual(r.classification, Classification.SPAM)

    def test_123_spam_phrases_present(self):
        self.assertEqual(len(SPAM_PHRASES), 123)


# ─── Link shortener ───────────────────────────────────────────────────────────

class TestLinkShortener(unittest.TestCase):

    def test_shortener_with_resume_passes(self):
        r = _classify(body="My portfolio: bit.ly/myportfolio", attach_names=("resume.pdf",))
        self.assertNotIn(r.classification, BLOCKED)

    def test_shortener_without_resume_blocked(self):
        r = _classify(body="Check my profile: bit.ly/profile123", attach_names=())
        self.assertEqual(r.classification, Classification.SHORTENER_ONLY)

    def test_12_shortener_phrases_present(self):
        self.assertEqual(len(LINK_SHORTENER_PHRASES), 12)


# ─── Routing paths ────────────────────────────────────────────────────────────

class TestRouting(unittest.TestCase):

    def test_new_candidate_route(self):
        result = _route()
        self.assertEqual(result.route, Route.NEW_CANDIDATE)

    def test_wrong_format_no_resume(self):
        result = _route(attach_names=("photo.jpg",))
        self.assertEqual(result.route, Route.WRONG_FORMAT)

    def test_cv_request_no_attach_with_keywords(self):
        result = _route(attach_names=(), body="I am applying for the ML Engineer role.")
        self.assertEqual(result.route, Route.CV_REQUEST)

    def test_ignored_mail_no_keywords(self):
        result = _route(attach_names=(), body="Hello, just checking in.", subject="Hello")
        self.assertEqual(result.route, Route.IGNORED_MAIL)

    def test_duplicate_within_90_days(self):
        result = _route(
            sender="good@example.com",
            lookup_by_email=_in_window_record,
        )
        self.assertEqual(result.route, Route.DUPLICATE)

    def test_not_duplicate_outside_90_days(self):
        result = _route(
            sender="good@example.com",
            lookup_by_email=_outside_window_record,
        )
        self.assertEqual(result.route, Route.NEW_CANDIDATE)

    def test_resume_update_quoted_ref_known_sender(self):
        ref = "APP-20260101-1200-ABCD"
        record = {"Application ID": ref, "Email": "good@example.com",
                  "Application Updates": 0, "Received Date": "2026-01-01T12:00:00+00:00"}
        result = _route(
            sender="good@example.com",
            subject=f"Re: my application {ref}",
            attach_names=("resume_v2.pdf",),
            lookup_by_app_ref=lambda r: record if r == ref else None,
        )
        self.assertEqual(result.route, Route.RESUME_UPDATE)
        self.assertEqual(result.verified_app_ref, ref)

    def test_is_known_sender_guard(self):
        """Quoted APP-Ref from wrong sender → falls through to normal intake."""
        ref = "APP-20260101-1200-ABCD"
        record = {"Application ID": ref, "Email": "real@example.com",
                  "Application Updates": 0, "Received Date": "2026-01-01T12:00:00+00:00"}
        result = _route(
            sender="attacker@evil.com",
            subject=f"Update for {ref}",
            attach_names=("resume.pdf",),
            lookup_by_app_ref=lambda r: record if r == ref else None,
        )
        # Should NOT be RESUME_UPDATE — IsKnownSender blocked cross-candidate corruption.
        self.assertNotEqual(result.route, Route.RESUME_UPDATE)
        self.assertEqual(result.route, Route.NEW_CANDIDATE)

    def test_followup_quoted_ref_no_resume(self):
        ref = "APP-20260101-1200-ABCD"
        record = {"Application ID": ref, "Email": "good@example.com",
                  "Application Updates": 1, "Received Date": "2026-01-01T12:00:00+00:00"}
        result = _route(
            sender="good@example.com",
            subject=f"Following up on {ref}",
            attach_names=(),
            body="Just checking in on my application.",
            lookup_by_app_ref=lambda r: record if r == ref else None,
        )
        self.assertEqual(result.route, Route.FOLLOWUP)

    def test_duplicate_window_uses_max_received_date(self):
        """Duplicate detection uses max(received_date), not sort order."""
        now = datetime.now(timezone.utc)
        # Records in reverse chronological order — out-of-order delivery simulation.
        records = [
            {"Application ID": "APP-1", "Email": "x@y.com",
             "Received Date": (now - timedelta(days=5)).isoformat(),   # ← newer
             "Application Updates": 0},
            {"Application ID": "APP-2", "Email": "x@y.com",
             "Received Date": (now - timedelta(days=200)).isoformat(),  # ← very old
             "Application Updates": 0},
        ]
        result = _route(
            sender="x@y.com",
            lookup_by_email=lambda _: records,
            now=now,
        )
        # The max date (5 days ago) is within 90 days → DUPLICATE.
        self.assertEqual(result.route, Route.DUPLICATE)


# ─── Anti-flood caps ──────────────────────────────────────────────────────────

class TestFloodCaps(unittest.TestCase):

    def _record_with_updates(self, n: int, email="good@example.com") -> list[dict]:
        recent = (datetime.now(timezone.utc) - timedelta(days=1)).isoformat()
        return [{"Application ID": "APP-20260101-1200-ABCD", "Email": email,
                 "Received Date": recent, "Application Updates": n}]

    def test_reply_cap_at_3(self):
        # Contact 3 (update_count=3, i.e. 4th contact) → no email.
        result = _route(
            sender="good@example.com",
            lookup_by_email=lambda _: self._record_with_updates(3),
        )
        self.assertEqual(result.route, Route.DUPLICATE)
        self.assertFalse(result.send_email)

    def test_under_reply_cap_sends_email(self):
        result = _route(
            sender="good@example.com",
            lookup_by_email=lambda _: self._record_with_updates(1),
        )
        self.assertEqual(result.route, Route.DUPLICATE)
        self.assertTrue(result.send_email)

    def test_dup_cap_at_5_silences(self):
        result = _route(
            sender="good@example.com",
            lookup_by_email=lambda _: self._record_with_updates(5),
        )
        self.assertEqual(result.route, Route.DUPLICATE)
        self.assertFalse(result.send_email)


# ─── APP-Ref minting and detection ────────────────────────────────────────────

class TestAppRef(unittest.TestCase):

    def test_mint_format(self):
        dt = datetime(2026, 9, 13, 12, 30, 0, tzinfo=timezone.utc)
        ref = mint_app_ref(dt, "<msg123@mail.example.com>")
        self.assertTrue(ref.startswith("APP-20260913-"))
        self.assertRegex(ref, r"^APP-\d{8}-\d{4}-[A-Z0-9]{4}$")

    def test_mint_deterministic(self):
        dt = datetime(2026, 9, 13, 12, 30, 0, tzinfo=timezone.utc)
        r1 = mint_app_ref(dt, "<same-id@mail.com>")
        r2 = mint_app_ref(dt, "<same-id@mail.com>")
        self.assertEqual(r1, r2)

    def test_mint_different_ids_different_refs(self):
        dt = datetime(2026, 9, 13, 12, 30, 0, tzinfo=timezone.utc)
        r1 = mint_app_ref(dt, "<id-a@mail.com>")
        r2 = mint_app_ref(dt, "<id-b@mail.com>")
        self.assertNotEqual(r1, r2)

    def test_extract_from_subject(self):
        ref = extract_quoted_ref("Re: APP-20260913-1230-ABCD update", "")
        self.assertEqual(ref, "APP-20260913-1230-ABCD")

    def test_extract_from_body(self):
        ref = extract_quoted_ref("", "My application app-20260913-1230-abcd")
        self.assertEqual(ref, "APP-20260913-1230-ABCD")  # normalized to upper

    def test_extract_none_when_absent(self):
        ref = extract_quoted_ref("Hello", "No reference here")
        self.assertIsNone(ref)

    def test_is_valid(self):
        self.assertTrue(is_valid_app_ref("APP-20260913-1230-ABCD"))
        self.assertFalse(is_valid_app_ref("not-a-ref"))

    def test_mst_timezone_used(self):
        # UTC midnight on May 1 → April 30 in MST (UTC-7) → date part should be 20260430
        dt = datetime(2026, 5, 1, 3, 0, 0, tzinfo=timezone.utc)  # 20:00 Arizona local Apr 30
        ref = mint_app_ref(dt, "<x@y.com>")
        self.assertIn("20260430", ref)


# ─── Database schema ──────────────────────────────────────────────────────────

class TestDbSchema(unittest.TestCase):

    def setUp(self):
        self.conn = get_connection(":memory:")
        migrate(self.conn)

    def tearDown(self):
        self.conn.close()

    def test_all_tables_created(self):
        tables = {r[0] for r in self.conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table'"
        ).fetchall()}
        required = {
            "intake_events", "applications", "application_versions",
            "documents", "jobs", "result_versions", "jd_snapshots",
            "audit_events", "export_jobs", "import_conflicts", "_schema_version",
        }
        missing = required - tables
        self.assertFalse(missing, f"Missing tables: {missing}")

    def test_schema_version_is_1(self):
        row = self.conn.execute("SELECT version FROM _schema_version").fetchone()
        self.assertEqual(row["version"], 1)

    def test_migration_idempotent(self):
        """Running migrate() twice must not raise or corrupt state."""
        migrate(self.conn)
        row = self.conn.execute("SELECT version FROM _schema_version").fetchone()
        self.assertEqual(row["version"], 1)

    def test_foreign_keys_enabled(self):
        pragma = self.conn.execute("PRAGMA foreign_keys").fetchone()
        self.assertEqual(pragma[0], 1)

    def test_intake_event_unique_key(self):
        key = "abc123"
        self.conn.execute(
            "INSERT INTO intake_events(event_key, sender_email, classification) "
            "VALUES (?, 'a@b.com', 'new')", (key,)
        )
        self.conn.commit()
        with self.assertRaises(sqlite3.IntegrityError):
            self.conn.execute(
                "INSERT INTO intake_events(event_key, sender_email, classification) "
                "VALUES (?, 'c@d.com', 'new')", (key,)
            )

    def test_insert_full_candidate_flow(self):
        """Insert intake_event → application → version → document → job in order."""
        with self.conn:
            self.conn.execute(
                "INSERT INTO intake_events(event_key, sender_email, classification) "
                "VALUES ('ev1', 'jane@example.com', 'new')"
            )
            self.conn.execute(
                "INSERT INTO applications(external_app_id) VALUES ('APP-20260913-1200-ABCD')"
            )
            app_id = self.conn.execute(
                "SELECT id FROM applications WHERE external_app_id='APP-20260913-1200-ABCD'"
            ).fetchone()["id"]
            self.conn.execute(
                "INSERT INTO application_versions"
                "(application_id, event_key, version_seq, sender_email) "
                "VALUES (?,?,0,'jane@example.com')", (app_id, "ev1")
            )
            ver_id = self.conn.execute(
                "SELECT id FROM application_versions WHERE application_id=?", (app_id,)
            ).fetchone()["id"]
            self.conn.execute(
                "INSERT INTO documents(application_version_id, stored_filename) "
                "VALUES (?, 'cv_guid.pdf')", (ver_id,)
            )
            self.conn.execute(
                "INSERT INTO jobs(application_version_id, stage) VALUES (?, 'score')", (ver_id,)
            )

        jobs = self.conn.execute("SELECT * FROM jobs WHERE stage='score'").fetchall()
        self.assertEqual(len(jobs), 1)
        self.assertEqual(jobs[0]["status"], "pending")


if __name__ == "__main__":
    unittest.main()


# ── resume-vs-envelope email (added 2026-09-18) ───────────────────────────────
# The CV address used to be regexed out and then dropped on the floor by
# extract_candidate_details_smart. These pin the fix: it is carried through, and
# it never displaces the envelope address, which is the system's identity key.

def test_resolve_email_keeps_envelope_as_contact():
    from hiring_agent.extraction import resolve_email
    cv = "Jane Doe\njane@gmail.com | 555-1234 | Austin, TX\nSUMMARY\n"
    contact, resume = resolve_email("jane@gmail.com", "agency@staffing.com", cv)
    assert contact == "agency@staffing.com", "identity key must not move"
    assert resume == "jane@gmail.com"


def test_resolve_email_silent_when_addresses_agree():
    from hiring_agent.extraction import resolve_email
    cv = "Jane Doe\njane@gmail.com | Austin, TX\n"
    contact, resume = resolve_email("jane@gmail.com", "jane@gmail.com", cv)
    assert (contact, resume) == ("jane@gmail.com", "")


def test_resolve_email_ignores_addresses_below_the_header():
    from hiring_agent.extraction import resolve_email
    # A referee's address at the foot of the CV is not the candidate's.
    cv = "Jane Doe\njane@gmail.com | Austin, TX\n" + "EXPERIENCE\n" * 25 + "ref: bob@oldjob.com\n"
    _, resume = resolve_email("bob@oldjob.com", "agency@staffing.com", cv)
    assert resume == ""


def test_resolve_email_handles_gap_literals():
    from hiring_agent.extraction import resolve_email
    for miss in ("Not extracted", "N/A", "", "unknown"):
        assert resolve_email(miss, "someone@corp.com", "CV\n")[1] == ""


def test_smart_extraction_carries_email_through():
    from hiring_agent.extraction import extract_candidate_details_smart
    cv = "Jane Doe\njane@gmail.com | 555-1234 | Austin, TX\n\nSKILLS\nPython\n"
    assert "email" in extract_candidate_details_smart(cv), "email key was being dropped"
