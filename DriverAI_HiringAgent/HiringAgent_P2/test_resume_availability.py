"""Offline regressions for stale queue copies and unavailable SharePoint resumes."""
from contextlib import ExitStack
from copy import deepcopy
import os
import unittest
from unittest.mock import Mock, patch

import requests

from hiring_agent import sharepoint_scoring as scoring
from sharepoint_client import SharePointClient, SharePointError


def candidate(status="Needs Review - Unreadable Resume", **changes):
    values = {
        "Application ID": "APP-20260819-0140-WPOA", "Email": "jane@example.com",
        "Status": status, "Full Name": "Jane Doe", "Has Resume": "Yes",
        "Received Date": "2026-08-19T01:40:01",
        "Last Updated Date": "2026-08-19T01:40:01",
        "Original Filename": "Jane_CV.pdf", "Application Updates": 0, "Retry Count": 2,
    }
    values.update(changes)
    return values


class QueueClient:
    hostname = "example.sharepoint.com"
    resumes_folder = "/Candidate_Resumes"

    def __init__(self, main, rejected=()):
        self.main = [{"index": i, "values": v} for i, v in enumerate(main)]
        self.rejected = [{"index": i, "values": v} for i, v in enumerate(rejected)]

    def list_rows(self):
        return self.main

    def list_rejected_rows(self):
        return self.rejected

    def list_unscored_rows(self, statuses):
        return [r for r in self.main if r["values"]["Status"] in statuses]


class ResumeAvailabilityTests(unittest.TestCase):
    def test_missing_resume_does_not_abort_later_scored_field_backfill(self):
        first = candidate("Scored", **{"Phone": "Missing"})
        second = candidate("Scored", **{"Application ID": "APP-NEXT", "Phone": "Missing"})
        client = QueueClient([first, second])
        store = Mock(spec=["save_by_id"])  # no all_rows -> exercises the client.list_rows() fallback
        with patch.object(scoring, "_download_resume_text", side_effect=[
            scoring.ResumeUnavailableError("404 missing first resume"),
            ("Jane Doe\nPhone: (623) 297-5055\n", [], {}),
        ]) as download, patch.object(scoring, "_store", return_value=store):
            fixed = scoring._reevaluate_missing_scored_fields(client)
        self.assertEqual(download.call_count, 2)
        self.assertEqual(fixed, 1)
        store.save_by_id.assert_called_once_with(
            "APP-NEXT", {"Phone": "(623) 297-5055"}, current_values=second, hint=1)

    def test_scored_field_backfill_never_calls_the_model(self):
        """The sweep documents itself as deterministic and Ollama-free; hold it to that.

        A role backfill calling infer_looking_for_role lived here until 2026-09-09 and
        invented "Seeking a role in customer service or sales, based on relevant skills
        and experience" for a resume stating neither - from a stub carrying only a name
        and a phone number. Every extractor this sweep may use has to be regex-only.
        """
        resume = """Jane Doe
Phone: (623) 297-5055
"""
        row = candidate("Scored", **{"Phone": "Missing", "Looking For Role": ""})
        client = QueueClient([row])
        store = Mock(spec=["save_by_id"])
        with (
            patch.object(scoring, "_download_resume_text", return_value=(resume, [], {})),
            patch.object(scoring, "_store", return_value=store),
            patch.object(scoring, "infer_looking_for_role") as role_model,
        ):
            fixed = scoring._reevaluate_missing_scored_fields(client)
        role_model.assert_not_called()
        self.assertEqual(fixed, 1)
        patch_written = store.save_by_id.call_args.args[1]
        self.assertEqual(patch_written, {"Phone": "(623) 297-5055"})
        self.assertNotIn("Looking For Role", patch_written)

    def test_completed_main_or_rejected_copy_does_not_consume_batch(self):
        for sheet in ("main", "rejected"):
            retry = candidate()
            completed = candidate("Scored" if sheet == "main" else "Rejected - Non-USA Location")
            new = candidate("New Email Received", **{"Application ID": "APP-NEW"})
            client = QueueClient([retry, new] + ([completed] if sheet == "main" else []),
                                 [completed] if sheet == "rejected" else [])
            before = deepcopy((client.main, client.rejected))
            with self.subTest(sheet=sheet):
                self.assertEqual([r["values"]["Application ID"]
                                  for r in scoring._scoring_queue_rows(client)[:1]], ["APP-NEW"])
                self.assertEqual((client.main, client.rejected), before)

    def test_new_updated_or_unverified_application_remains_eligible(self):
        completed = candidate("Rejected - Non-USA Location")
        variants = [
            {"Status": "New Email Received"},
            {"Last Updated Date": "2026-09-06T12:00:00"},
            {"Received Date": "2026-09-06T12:00:00"},
            {"Original Filename": "Jane_Updated_CV.pdf"},
            {"Application Updates": 1},
            {"Application ID": "APP-DIFFERENT"},
            {"Email": "someone-else@example.com"},
            {"Email": ""}, {"Last Updated Date": ""}, {"Application Updates": None},
        ]
        for change in variants:
            retry = candidate(**change)
            with self.subTest(change=change):
                self.assertFalse(scoring._unchanged_completed_retry(retry, completed))
                self.assertEqual(len(scoring._scoring_queue_rows(QueueClient([retry], [completed]))), 1)

    def test_pending_or_processing_error_copy_does_not_suppress_retry(self):
        for status in ("", "New Email Received", "Needs Review - Location Confirmation",
                       "Rejected - Processing Error", "Needs Review - Unreadable Resume"):
            with self.subTest(status=status):
                self.assertFalse(scoring._unchanged_completed_retry(candidate(), candidate(status)))

    def test_failed_completion_lookup_keeps_queue_row(self):
        client = QueueClient([candidate()])
        client.list_rejected_rows = Mock(side_effect=SharePointError("403", status_code=403))
        self.assertEqual(scoring._scoring_queue_rows(client), client.main)

    def test_missing_dated_file_still_uses_legacy_root(self):
        client = object.__new__(SharePointClient)
        client.resumes_folder = "/Candidate_Resumes"
        client.download_file = Mock(side_effect=[SharePointError("404", status_code=404), b"pdf"])
        self.assertEqual(client.download_resume("cv.pdf", "2026/August"), b"pdf")
        self.assertEqual(client.download_file.call_args_list[0].args,
                         ("/Candidate_Resumes/2026/August", "cv.pdf"))
        self.assertEqual(client.download_file.call_args_list[1].args,
                         ("/Candidate_Resumes", "cv.pdf"))

    def test_access_and_server_errors_are_not_masked_by_root_fallback(self):
        for status in (401, 403, 429, 503, None):
            client = object.__new__(SharePointClient)
            client.resumes_folder = "/Candidate_Resumes"
            error = SharePointError("unavailable", status_code=status)
            client.download_file = Mock(side_effect=error)
            with self.subTest(status=status), self.assertRaises(SharePointError) as raised:
                client.download_resume("cv.pdf", "2026/August")
            self.assertIs(raised.exception, error)
            client.download_file.assert_called_once()

    def test_all_missing_names_raise_unavailable_with_actual_error(self):
        client = Mock()
        client.download_resume.side_effect = SharePointError("404 itemNotFound", status_code=404)
        with self.assertRaisesRegex(scoring.ResumeUnavailableError, "404 itemNotFound"):
            scoring._download_resume_text(client, "APP-1", "cv.pdf", "2026/August", "Jane Doe")

    def test_downloaded_but_unreadable_file_keeps_parser_failure_path(self):
        client = Mock()
        client.download_resume.return_value = b"scanned PDF"
        with patch.object(scoring, "extract_text_from_bytes", return_value=""):
            text, used, raw = scoring._download_resume_text(
                client, "APP-1", "cv.pdf", "2026/August", "Jane Doe")
        self.assertEqual(text, "")
        self.assertEqual(len(used), 1)
        self.assertEqual(raw[used[0]], b"scanned PDF")

    def test_missing_second_attachment_does_not_score_partial_resume(self):
        client = Mock()
        def download(name, **kwargs):
            if name.endswith(".pdf"):
                return b"First attachment"
            raise SharePointError("404 missing second attachment", status_code=404)
        client.download_resume.side_effect = download
        with patch.object(scoring, "extract_text_from_bytes", return_value="First attachment"):
            with self.assertRaises(scoring.ResumeUnavailableError):
                scoring._download_resume_text(client, "APP-1", "cv.pdf, history.docx", "2026/August")

    def test_pipeline_defers_download_failure_without_applicant_retry_or_rejection(self):
        for dry_run in (False, True):
            for error in (SharePointError("404 itemNotFound", status_code=404),
                          SharePointError("403 Access denied", status_code=403),
                          requests.ConnectionError("connection reset")):
                client = QueueClient([candidate()])
                client.download_resume = Mock(side_effect=error)
                store = Mock()
                before = deepcopy(client.main)
                with self.subTest(dry_run=dry_run, error=str(error)), ExitStack() as stack:
                    stack.enter_context(patch.dict(os.environ, {"HIRING_P2_DISABLED": "false"}))
                    stack.enter_context(patch("sharepoint_client.SharePointClient", return_value=client))
                    stack.enter_context(patch("sharepoint_client.check_graph_reachable", return_value=(True, "OK")))
                    patches = {
                        "ollama_health": (True, "OK"), "_ensure_workbook_or_alert": True,
                        "_ensure_schema": None, "get_active_roles": [], "_live_row_index": 0,
                        "_store": store, "_merge_duplicate_candidates": {},
                        "_age_out_stale_reviews": {}, "_send_pending_declines": {},
                        "_send_pending_info_requests": {}, "export_client_results": "",
                        "_reevaluate_missing_scored_fields": 0,
                    }
                    for name, value in patches.items():
                        stack.enter_context(patch.object(scoring, name, return_value=value))
                    review = stack.enter_context(patch.object(scoring, "_mark_needs_review"))
                    reject = stack.enter_context(patch.object(scoring, "_give_up_and_reject"))
                    model = stack.enter_context(patch.object(scoring, "extract_candidate_details_smart"))
                    result = scoring.score_from_sharepoint(dry_run=dry_run)
                    self.assertEqual(result["deferred"], 1)
                    self.assertEqual(result["processed"], 0)
                    self.assertEqual(result["errors"], 0)
                    self.assertEqual(result["rejected"], 0)
                    self.assertEqual(client.main, before)
                    self.assertEqual(store.mock_calls, [])
                    review.assert_not_called()
                    reject.assert_not_called()
                    model.assert_not_called()


if __name__ == "__main__":
    unittest.main()
