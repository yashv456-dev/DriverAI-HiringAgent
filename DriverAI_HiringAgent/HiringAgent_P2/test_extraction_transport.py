"""Offline transport regressions; no Graph or live Ollama calls.

Run separately to preserve the existing acceptance suites and their counts:
    .venv/Scripts/python test_extraction_transport.py
"""
import json
import unittest
from unittest.mock import patch

import requests

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
        enabled = patch.object(ex, "OLLAMA_ENABLED", True)
        enabled.start()
        self.addCleanup(enabled.stop)

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
                self.assertEqual(body["options"], {"temperature": 0, "num_predict": 512, "seed": 42})
                schema = body["format"]
                self.assertEqual(schema["type"], "object")
                self.assertFalse(schema["additionalProperties"])
                self.assertEqual(set(schema["required"]), set(payload))
                self.assertEqual(schema["properties"],
                                 {key: {"type": "string"} for key in payload})

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

    def test_malformed_json_keeps_failure_handling_without_retry(self):
        for kind in PAYLOADS:
            expected = self.fallback(kind)
            with self.subTest(kind=kind), patch.object(
                    requests, "post", return_value=response(kind, content='{ "broken":')) as post:
                self.assertEqual(self.invoke(kind), expected)
                post.assert_called_once()

    def test_timeout_then_bad_json_does_not_get_a_third_attempt(self):
        for kind in PAYLOADS:
            expected = self.fallback(kind)
            with self.subTest(kind=kind), patch.object(requests, "post", side_effect=[
                    requests.Timeout("slow"), response(kind, content="{"),
                    response(kind)]) as post:
                self.assertEqual(self.invoke(kind), expected)
                self.assertEqual(post.call_count, 2)

    def test_recovered_timeout_records_model_provenance(self):
        ex.EXTRACTION_SOURCE["value"] = "offline (ConnectionError)"
        with patch.object(requests, "post", side_effect=[
                requests.ReadTimeout("slow"), response("extract")]):
            self.assertIsNotNone(self.invoke("extract"))
        self.assertEqual(ex.EXTRACTION_SOURCE["value"], "ollama")


if __name__ == "__main__":
    unittest.main()
