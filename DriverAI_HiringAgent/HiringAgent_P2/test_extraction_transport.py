"""Offline transport regressions; no Graph or live Ollama calls.

Run separately to preserve the existing acceptance suites and their counts:
    .venv/Scripts/python test_extraction_transport.py
"""
import json
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

import requests

sys.path.insert(0, str(Path(__file__).resolve().parent))
import test_isolation  # noqa: F401 - must precede every hiring_agent import

from hiring_agent import extraction as ex


TEXT = "Alex Morgan\nAustin, Texas, United States\n+1 (512) 555-0142\nPython\n"
FIELDS = {
    "full_name": "Alex Morgan", "phone": "+1 (512) 555-0142",
    "location": "Austin, Texas", "country": "United States", "skills": "Python",
    "looking_for_role": "Software Engineer", "education": "",
}
PAYLOADS = {
    "extract": FIELDS,
    "recheck": FIELDS,
    "role": {"summary": "Seeking a software engineering role."},
    "portfolios": {"linkedin": "", "github": "", "other": ""},
}


def response(kind, *, content=None, status=200):
    result = requests.Response()
    result.status_code = status
    result._content = json.dumps({"message": {
        "content": json.dumps(PAYLOADS[kind]) if content is None else content,
    }}).encode()
    return result


class ExtractionTransportTests(unittest.TestCase):
    def setUp(self):
        # Keep fallback tests independent of the machine's production settings.
        # Strict behavior is exercised explicitly below.
        for name, value in (("OLLAMA_ENABLED", True), ("REQUIRE_AI", False),
                            ("STRICT_AI_STAGES", False)):
            setting = patch.object(ex, name, value, create=True)
            setting.start()
            self.addCleanup(setting.stop)
        provenance = patch.dict(ex.EXTRACTION_SOURCE)
        provenance.start()
        self.addCleanup(provenance.stop)

    def invoke(self, kind):
        if kind == "extract":
            return ex.extract_with_ollama(TEXT)
        if kind == "recheck":
            return ex.ai_recheck_fields(FIELDS, TEXT, "")
        if kind == "role":
            return ex.infer_looking_for_role(TEXT, "", "Python")
        return ex.infer_missing_portfolios(TEXT, "N/A", "N/A", "N/A")

    def fallback(self, kind):
        with patch.object(ex, "OLLAMA_ENABLED", False):
            return self.invoke(kind)

    def test_request_contract_and_single_successful_attempt(self):
        for kind, payload in PAYLOADS.items():
            with self.subTest(kind=kind), patch.object(
                    requests, "post", return_value=response(kind)) as post:
                self.invoke(kind)
                post.assert_called_once()
                request = post.call_args.kwargs
                self.assertEqual(request["timeout"], ex.OLLAMA_TIMEOUT)
                body = request["json"]
                self.assertIs(body["think"], False)
                self.assertIs(body["stream"], False)
                self.assertEqual(body["options"], {"temperature": 0,
                                                   "num_predict": ex._EXTRACTION_NUM_PREDICT,
                                                   "seed": 42})
                schema = body["format"]
                self.assertEqual(schema["type"], "object")
                self.assertFalse(schema["additionalProperties"])
                self.assertEqual(set(schema["required"]), set(payload))
                self.assertEqual(set(schema["properties"]), set(payload))
                self.assertTrue(all(schema["properties"][key]["type"] == "string"
                                    for key in payload))
                if kind in ("extract", "recheck"):
                    # Grammar-enforced bounds, so a runaway enumeration cannot run the reply
                    # past the token cap and leave invalid JSON (Saad Ullah, 2026-09-16).
                    self.assertTrue(all(schema["properties"][key].get("maxLength", 0) > 0
                                        for key in payload))

    def test_timeout_then_success_retries_exactly_once(self):
        for kind in PAYLOADS:
            with patch.object(requests, "post", return_value=response(kind)):
                expected = self.invoke(kind)
            for error in (requests.Timeout, requests.ReadTimeout, requests.ConnectTimeout):
                with self.subTest(kind=kind, error=error.__name__), patch.object(
                        requests, "post", side_effect=[error("slow"), response(kind)]) as post:
                    self.assertEqual(self.invoke(kind), expected)
                    self.assertEqual(post.call_count, 2)
                    self.assertEqual(post.call_args_list[0], post.call_args_list[1])

    def test_two_timeouts_exhaust_retry_and_keep_fallback(self):
        for kind in PAYLOADS:
            expected = self.fallback(kind)
            with self.subTest(kind=kind), patch.object(
                    requests, "post", side_effect=requests.ReadTimeout("slow")) as post:
                self.assertEqual(self.invoke(kind), expected)
                self.assertEqual(post.call_count, 2)
                if kind == "extract":
                    self.assertIn("timed out", ex.EXTRACTION_SOURCE["value"])

    def test_connection_error_is_not_retried(self):
        for kind in PAYLOADS:
            expected = self.fallback(kind)
            with self.subTest(kind=kind), patch.object(
                    requests, "post", side_effect=requests.ConnectionError("refused")) as post:
                self.assertEqual(self.invoke(kind), expected)
                post.assert_called_once()
                if kind == "extract":
                    self.assertEqual(ex.EXTRACTION_SOURCE["value"], "offline (ConnectionError)")

    def test_timeout_then_connection_error_does_not_get_a_third_attempt(self):
        for kind in PAYLOADS:
            expected = self.fallback(kind)
            with self.subTest(kind=kind), patch.object(requests, "post", side_effect=[
                    requests.Timeout("slow"), requests.ConnectionError("refused"),
                    response(kind)]) as post:
                self.assertEqual(self.invoke(kind), expected)
                self.assertEqual(post.call_count, 2)
                if kind == "extract":
                    self.assertEqual(ex.EXTRACTION_SOURCE["value"], "offline (ConnectionError)")

    def test_http_error_is_not_retried(self):
        for kind in PAYLOADS:
            expected = self.fallback(kind)
            with self.subTest(kind=kind), patch.object(
                    requests, "post", return_value=response(kind, status=503)) as post:
                self.assertEqual(self.invoke(kind), expected)
                post.assert_called_once()

    def test_malformed_json_retries_only_initial_extraction(self):
        for kind in PAYLOADS:
            expected = self.fallback(kind)
            with self.subTest(kind=kind), patch.object(
                    requests, "post", return_value=response(kind, content='{ "broken":')) as post:
                self.assertEqual(self.invoke(kind), expected)
                self.assertEqual(post.call_count, 2 if kind == "extract" else 1)

    def test_timeout_then_bad_json_retries_initial_extraction_again(self):
        for kind in PAYLOADS:
            if kind == "extract":
                with patch.object(requests, "post", return_value=response(kind)):
                    expected = self.invoke(kind)
            else:
                expected = self.fallback(kind)
            with self.subTest(kind=kind), patch.object(requests, "post", side_effect=[
                    requests.Timeout("slow"), response(kind, content="{"),
                    response(kind)]) as post:
                self.assertEqual(self.invoke(kind), expected)
                self.assertEqual(post.call_count, 3 if kind == "extract" else 2)

    def test_malformed_json_then_success_records_model_provenance(self):
        with patch.object(requests, "post", return_value=response("extract")):
            expected = self.invoke("extract")
        ex.EXTRACTION_SOURCE["value"] = "offline (ConnectionError)"
        with patch.object(requests, "post", side_effect=[
                response("extract", content="{"), response("extract")]) as post:
            self.assertEqual(self.invoke("extract"), expected)
            self.assertEqual(post.call_count, 2)
            # The JSON retry is the same request with double the output room: an identical
            # request at temperature 0 reproduces an identical truncation.
            first, retry = (c.kwargs["json"] for c in post.call_args_list)
            self.assertEqual(retry["options"]["num_predict"],
                             2 * first["options"]["num_predict"])
            self.assertEqual(dict(retry, options=None), dict(first, options=None))
        self.assertEqual(ex.EXTRACTION_SOURCE["value"], "ollama")

    def test_combined_timeout_and_json_retries_are_bounded_at_four_requests(self):
        with patch.object(requests, "post", return_value=response("extract")):
            successful = self.invoke("extract")
        for ending in ("success", "bad_json", "timeout"):
            last = {"success": response("extract"),
                    "bad_json": response("extract", content="{"),
                    "timeout": requests.ReadTimeout("slow")}[ending]
            with self.subTest(ending=ending), patch.object(
                    requests, "post", side_effect=[
                        requests.Timeout("slow"), response("extract", content="{"),
                        requests.Timeout("slow"), last, response("extract")]) as post:
                self.assertEqual(self.invoke("extract"),
                                 successful if ending == "success" else None)
                self.assertEqual(post.call_count, 4)
                # Calls 1-2 are the first attempt and its transport retry (identical); calls
                # 3-4 are the JSON retry and ITS transport retry, both with double the room.
                calls = post.call_args_list
                self.assertEqual(calls[0], calls[1])
                self.assertEqual(calls[2], calls[3])
                self.assertEqual(calls[2].kwargs["json"]["options"]["num_predict"],
                                 2 * calls[0].kwargs["json"]["options"]["num_predict"])
                if ending == "success":
                    self.assertEqual(ex.EXTRACTION_SOURCE["value"], "ollama")
                elif ending == "bad_json":
                    self.assertEqual(ex.EXTRACTION_SOURCE["value"],
                                     "offline (JSONDecodeError)")
                else:
                    self.assertIn("timed out", ex.EXTRACTION_SOURCE["value"])

    def test_recheck_failure_raises_when_either_strict_setting_is_enabled(self):
        failures = (
            ("timeout", {"side_effect": requests.ReadTimeout("slow")},
             requests.ReadTimeout, 2),
            ("connection", {"side_effect": requests.ConnectionError("refused")},
             requests.ConnectionError, 1),
            ("http", {"return_value": response("recheck", status=503)},
             requests.HTTPError, 1),
            ("bad_json", {"return_value": response("recheck", content="{")},
             json.JSONDecodeError, 1),
        )
        for require_ai, strict_stages in ((True, False), (False, True), (True, True)):
            with patch.object(ex, "REQUIRE_AI", require_ai), patch.object(
                    ex, "STRICT_AI_STAGES", strict_stages):
                for name, behavior, cause, attempts in failures:
                    with self.subTest(require_ai=require_ai, strict_stages=strict_stages,
                                      failure=name), patch.object(
                            requests, "post", **behavior) as post:
                        with self.assertRaisesRegex(
                                RuntimeError, "Independent AI recheck failed; result deferred") as caught:
                            self.invoke("recheck")
                        self.assertIsInstance(caught.exception.__cause__, cause)
                        self.assertEqual(post.call_count, attempts)

    def test_recovered_timeout_records_model_provenance(self):
        ex.EXTRACTION_SOURCE["value"] = "offline (ConnectionError)"
        with patch.object(requests, "post", side_effect=[
                requests.ReadTimeout("slow"), response("extract")]):
            self.assertIsNotNone(self.invoke("extract"))
        self.assertEqual(ex.EXTRACTION_SOURCE["value"], "ollama")


if __name__ == "__main__":
    unittest.main()
