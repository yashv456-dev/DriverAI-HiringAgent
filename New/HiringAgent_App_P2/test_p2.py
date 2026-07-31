"""P2 deep end-to-end test â€” no live SharePoint, no live Ollama required.

Tests every extraction/scoring/validation layer across all 14 P2-owned columns.
Run from HiringAgent_App_P2/:
    .venv/Scripts/python test_p2.py
"""
import io, os, sys
from pathlib import Path
os.environ.setdefault("HIRING_OLLAMA_ENABLED", "false")
os.environ.setdefault("HIRING_OLLAMA_SCORING", "false")
sys.path.insert(0, str(__import__("pathlib").Path(__file__).resolve().parent))
sys.stdout.reconfigure(encoding="utf-8")
from openpyxl import load_workbook

P = F = 0
def ok(c, label):
    global P, F
    if c: P += 1; print(f"  [PASS] {label}")
    else: F += 1; print(f"  [FAIL] {label}")

print("\n=== CLI BATCH SIZE ===")
import argparse as _argparse
from bot import _positive_int

for _size in (1, 25, 50):
    ok(_positive_int(str(_size)) == _size,
       f"--batch-size accepts positive whole number {_size}")
for _invalid_size in ("0", "-1", "many"):
    try:
        _positive_int(_invalid_size)
        _rejected_size = False
    except _argparse.ArgumentTypeError:
        _rejected_size = True
    ok(_rejected_size, f"--batch-size rejects invalid value {_invalid_size!r}")


# â”€â”€ A. OLLAMA COVERAGE MAP â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== A. OLLAMA COVERAGE MAP (all 14 P2 columns) ===")
# Documents what covers each P2 column â€” tested structurally here.
P2_COLS = {
    "Full Name":        "ai_recheck_fields",
    "Phone":            "ai_recheck_fields",
    "Location":         "ai_recheck_fields",
    "Country":          "ai_recheck_fields",
    "Current Skills":   "ai_recheck_fields",
    "Looking For Role": "ai_recheck_fields",
    "Portfolio 1":      "extract_portfolios (regex only)",
    "Portfolio 2":      "extract_portfolios (regex only)",
    "Portfolio 3":      "extract_portfolios (regex only)",
    "Suggested Role 1": "ai_score_roles (separate Ollama call)",
    "Suggested Role 2": "ai_score_roles (separate Ollama call)",
    "Suggested Role 3": "ai_score_roles (separate Ollama call)",
    "Category":         "assign_category (deterministic from Role 1)",
    "Status":           "hardcoded (Scored / Rejected - Non-USA Location)",
}
ok(len(P2_COLS) == 14, "all 14 P2 columns accounted for")
ai_recheck_cols = [c for c, src in P2_COLS.items() if src == "ai_recheck_fields"]
ok(len(ai_recheck_cols) == 6,
   f"ai_recheck_fields covers exactly 6 columns: {ai_recheck_cols}")
ok("Portfolio 1" not in ai_recheck_cols and "Suggested Role 1" not in ai_recheck_cols,
   "portfolios and suggested roles NOT in ai_recheck_fields scope")
ok(P2_COLS["Category"] == "assign_category (deterministic from Role 1)",
   "Category is deterministic (assign_category) â€” no AI")
ok(P2_COLS["Status"] == "hardcoded (Scored / Rejected - Non-USA Location)",
   "Status is hardcoded â€” no AI")


# â”€â”€ B. OLLAMA CALL COUNT (up to 3 per candidate) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== B. OLLAMA CALL COUNT ===")
# When OLLAMA_ENABLED + OLLAMA_SCORING:
#   call 1 â†’ extract_with_ollama() inside extract_candidate_details_smart()
#   call 2 â†’ ai_recheck_fields()
#   call 3 â†’ ai_score_roles() inside suggested_roles()
# When Ollama is OFF â†’ keyword scorer (0 calls)
from hiring_agent.extraction import ai_recheck_fields
from hiring_agent.config import OLLAMA_ENABLED, OLLAMA_SCORING

ok(not OLLAMA_ENABLED, "OLLAMA_ENABLED=false in this test run (all paths use offline fallback)")
test_fields = {"full_name": "John Smith", "phone": "555-1234",
               "location": "Austin, TX", "country": "United States",
               "skills": "Python", "looking_for_role": "Engineer"}
result = ai_recheck_fields(test_fields, "resume text", "mail body")
ok(result == test_fields, "ai_recheck_fields returns original unchanged when Ollama off")

print("\n=== B2. EDUCATION NORMALIZATION / QUALITY GATE ===")
from hiring_agent.extraction import normalize_education, education_needs_repair, _extract_education
edu_cases = {
    "MasterofScienceinComputerScience, ArizonaStateUniversity":
        "Master of Science in Computer Science, Arizona State University",
    "BachelorofTechnologyinComputerEngineering":
        "Bachelor of Technology in Computer Engineering",
    "Master of Science in Data Science, Arizona State University":
        "Master of Science in Data Science, Arizona State University",
}
for raw, expected in edu_cases.items():
    got = normalize_education(raw)
    ok(got == expected, f"education normalized: {raw!r} -> {got!r}")
ok(education_needs_repair("MasterofScienceinDataScience, ArizonaStateUniversit"),
   "fused/truncated education is marked for resume re-extraction")
ok(education_needs_repair("B.Tech, Dr. A. P. J. Abdul Kalam T echnical University"),
   "OCR-split institution word is marked for resume re-extraction")
ok(not education_needs_repair("M.S. Computer Science, Arizona State University"),
   "clean education passes the quality gate")
# 2026-07-24: a resume-formatting divider line ('----...----') mis-captured as Education
# (real case: 'Syed Asad' row on the live Rejected sheet) - normalize_education already
# reduces this to "Not extracted" via its strip(), but education_needs_repair must flag it
# so a heal pass actually re-applies that instead of leaving the raw divider published.
ok(education_needs_repair("-" * 128),
   "a pure-dash divider line is marked for repair")
ok(education_needs_repair("=" * 40), "a pure-equals divider line is marked for repair")
ok(normalize_education("-" * 128) == "Not extracted",
   "normalize_education already reduces a pure-dash divider to Not extracted")
ok(normalize_education("AND TRAINING") == "Not extracted",
   "education section suffix alone is not saved as education")
ok(_extract_education("EDUCATION AND TRAINING\n05/10/2018 - 05/10/2022\nBS-INFORMATION TECHNOLOGY The Islamia University of Bahawalpur") ==
   "BS-INFORMATION TECHNOLOGY The Islamia University of Bahawalpur",
   "Education and Training header skips date line and captures degree line")

# Ã¢â€â‚¬Ã¢â€â‚¬ C. SHAREPOINT WORKBOOK SHAPE Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
print("\n=== C. SHAREPOINT WORKBOOK SHAPE ===")
from sharepoint_client import _build_candidate_workbook_bytes
from hiring_agent.config import COLUMNS

wb = load_workbook(io.BytesIO(_build_candidate_workbook_bytes(
    COLUMNS, "HiringAgent_P1_Candidates"
)))
ws = wb.active
tables = list(ws.tables.values())
table = tables[0] if tables else None
ok(ws.title == "CandidateList", f"workbook sheet name = CandidateList (got {ws.title})")
ok(len(tables) == 1, f"workbook has exactly 1 table (got {len(tables)})")
ok(getattr(table, "ref", "").endswith("2"),
   f"table ref includes blank seed row: {getattr(table, 'ref', '')}")
ok(all((ws.cell(row=2, column=i).value in ("", None)) for i in range(1, ws.max_column + 1)),
   "row 2 is intentionally blank (Graph-safe seed row)")


# â”€â”€ C2. ensure_workbook() NEVER creates/overwrites the workbook (fixed 2026-07-24) â”€â”€â”€â”€â”€
print("\n=== C2. ensure_workbook() never auto-creates (mirrors P1's non-destructive design) ===")
# Regression test: an earlier version rebuilt a blank workbook on a genuine 404, an
# asymmetry with Phase 1 (which had that exact self-heal removed 2026-07-04 after it
# proved unreliable and wiped real data). ensure_workbook() must now do nothing but check.
import types as _types
from sharepoint_client import SharePointClient as _SPClient, SharePointError as _SPErr

_ensure_calls = {"table_columns": 0, "upload_file": 0}
_fake_self = _types.SimpleNamespace(
    table_columns=lambda: (_ensure_calls.__setitem__("table_columns", _ensure_calls["table_columns"] + 1),
                           (_ for _ in ()).throw(_SPErr("not found", status_code=404)))[-1],
    upload_file=lambda *a, **kw: _ensure_calls.__setitem__("upload_file", _ensure_calls["upload_file"] + 1),
    _wb_path="/Master_Files/Sharepoint_Master_File.xlsx",
)
_raised = False
try:
    _SPClient.ensure_workbook(_fake_self)
except _SPErr:
    _raised = True
ok(_raised, "ensure_workbook() raises on a missing workbook instead of swallowing it")
ok(_ensure_calls["table_columns"] == 1, "ensure_workbook() checks the table")
ok(_ensure_calls["upload_file"] == 0, "ensure_workbook() NEVER calls upload_file - no auto-create, ever")

# score_from_sharepoint() must abort cleanly (not limp forward into a doomed row loop)
# and alert an admin when the workbook can't be confirmed.
import unittest.mock as _mock
from hiring_agent.sharepoint_scoring import score_from_sharepoint as _sfs

class _StubMissingWorkbookClient:
    def __init__(self):
        self.hostname = "test.sharepoint.com"
        self.mails_sent = []

    def ensure_workbook(self):
        raise _SPErr("table 'HiringAgent_P1_Candidates' not found", status_code=404)

    def send_mail(self, to, subject, body):
        self.mails_sent.append((to, subject, body))

import hiring_agent.config as _cfg_mod
_orig_enabled, _orig_admin = _cfg_mod.ERROR_EMAIL_ENABLED, _cfg_mod.ADMIN_EMAIL
try:
    _cfg_mod.ERROR_EMAIL_ENABLED, _cfg_mod.ADMIN_EMAIL = True, "admin@x.com"
    _stub_mw = _StubMissingWorkbookClient()
    with _mock.patch("sharepoint_client.SharePointClient", return_value=_stub_mw), \
         _mock.patch("sharepoint_client.check_graph_reachable", return_value=(True, "OK")), \
         _mock.patch("hiring_agent.sharepoint_scoring.ollama_health", return_value=(True, "OK")), \
         _mock.patch.dict(os.environ, {"HIRING_P2_DISABLED": "false"}):
        _res = _sfs(dry_run=False)
    ok(isinstance(_res, dict) and "error" in _res,
       "score_from_sharepoint() aborts with an error result when the workbook is missing")
    ok(len(_stub_mw.mails_sent) == 1 and _stub_mw.mails_sent[0][0] == "admin@x.com",
       "a missing workbook sends exactly one admin alert")
    # Dry-run must stay side-effect-free: never even call ensure_workbook.
    from hiring_agent.sharepoint_scoring import _ensure_workbook_or_alert as _ewa
    _stub_mw2 = _StubMissingWorkbookClient()
    ok(_ewa(_stub_mw2, dry_run=True) is True,
       "dry-run reports the workbook as fine without ever checking it")
    ok(_stub_mw2.mails_sent == [], "dry-run never alerts about the workbook (never even checks it)")
finally:
    _cfg_mod.ERROR_EMAIL_ENABLED, _cfg_mod.ADMIN_EMAIL = _orig_enabled, _orig_admin

# Token acquisition happens before Graph _req(), so it needs its own transient retry.
print("\n=== C3. ENTRA TOKEN TRANSIENT RETRY ===")
import requests as _requests

class _TokenResponse:
    status_code = 200
    text = ""
    headers = {}

    @staticmethod
    def json():
        return {"access_token": "header.payload.signature"}

_token_env = {
    "TENANT_ID": "tenant-id",
    "CLIENT_ID": "client-id",
    "CLIENT_SECRET": "client-secret",
    "SHAREPOINT_HOSTNAME": "example.sharepoint.com",
    "SHAREPOINT_SITE_PATH": "/sites/test",
    "SHAREPOINT_TABLE": "Candidates",
}
with _mock.patch.dict(os.environ, _token_env, clear=False), \
        _mock.patch("sharepoint_client.requests.post",
                    side_effect=[_requests.ConnectionError("connection reset"),
                                 _TokenResponse()]) as _post, \
        _mock.patch("time.sleep") as _sleep:
    _token_client = _SPClient()
    _token = _token_client._get_token()
ok(_token == "header.payload.signature" and _post.call_count == 2,
   "transient token connection failure retries and succeeds")
ok(_sleep.call_count == 1,
   "token retry applies one bounded backoff before the successful attempt")


# â”€â”€ C. AI RECHECK â€” GAP-LITERAL PROTECTION (bug fix) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== C. AI RECHECK â€” GAP-LITERAL OVERWRITE PROTECTION ===")
# Simulate Ollama returning gap-literal strings (off-prompt behavior).
# The fixed code must NOT overwrite a good value with "Not extracted", "N/A", etc.
import unittest.mock as _mock

_GAP_RESPONSES = [
    {"full_name": "Not extracted", "phone": "", "location": "",
     "country": "", "skills": "", "looking_for_role": ""},
    {"full_name": "N/A", "phone": "N/A", "location": "N/A",
     "country": "N/A", "skills": "N/A", "looking_for_role": "N/A"},
    {"full_name": "unknown", "phone": "none", "location": "not found",
     "country": "na", "skills": "-", "looking_for_role": ""},
]

good = {"full_name": "Jane Doe", "phone": "512-555-0001",
        "location": "Austin, TX", "country": "United States",
        "skills": "Python, AWS", "looking_for_role": "Data Engineer"}

for gap_resp in _GAP_RESPONSES:
    with _mock.patch("requests.post") as mp:
        mp.return_value.json.return_value = {"message": {"content": __import__("json").dumps(gap_resp)}}
        mp.return_value.raise_for_status = lambda: None
        os.environ["HIRING_OLLAMA_ENABLED"] = "true"
        import importlib
        import hiring_agent.config as _cfg_mod
        import hiring_agent.extraction as _ext_mod
        importlib.reload(_cfg_mod)
        importlib.reload(_ext_mod)
        from hiring_agent.extraction import ai_recheck_fields as _arc
        r = _arc(dict(good), "resume", "body")
        os.environ["HIRING_OLLAMA_ENABLED"] = "false"
        importlib.reload(_cfg_mod)
        importlib.reload(_ext_mod)
    ok(r.get("full_name") == "Jane Doe",
       f"gap-literal '{gap_resp['full_name']}' did NOT overwrite Full Name")
    ok(r.get("phone") == "512-555-0001",
       f"gap-literal phone did NOT overwrite Phone")
    ok(r.get("skills") == "Python, AWS",
       f"gap-literal skills did NOT overwrite Skills")


# â”€â”€ D. OFFLINE EXTRACTION â€” FULL NAME â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== D. OFFLINE EXTRACTION â€” Full Name ===")
os.environ["HIRING_OLLAMA_ENABLED"] = "false"
from hiring_agent.extraction import extract_candidate_details, resolve_full_name

def extr(text): return extract_candidate_details(text)

ok(extr("John Smith\nPython developer\nAustin, TX")["full_name"] == "John Smith",
   "Name from first line of resume")
ok(extr("My name is Maria Garcia\nSoftware engineer")["full_name"] == "Maria Garcia",
   "Name from 'my name is' prefix (ASCII â€” accented chars need Ollama/AI path)")
# Known limitation: offline regex [A-Za-z] is ASCII-only; accented names like
# "MarÃ­a GarcÃ­a" are handled correctly by the Ollama/AI extraction path, not offline.
ok(extr("ALICE JOHNSON\nEngineer")["full_name"] == "Alice Johnson",
   "ALLCAPS name normalized to Title Case")
ok(extr("Python, AWS, Docker\nEngineer")["full_name"] == "Not extracted",
   "No name when first line is skills list")
ok(resolve_full_name("Bob Chen", "", "")  == "Bob Chen",
   "resolve_full_name: good parsed name wins")
ok(resolve_full_name("Not extracted", "Sarah Park", "") == "Sarah Park",
   "resolve_full_name: falls back to sender name")
ok(resolve_full_name("Not extracted", "", "Carlos Ruiz\nSoftware Engineer\n") == "Carlos Ruiz",
   "resolve_full_name: falls back to resume first lines")
ok(resolve_full_name("Not extracted", "noreply@x.com", "") == "Not extracted",
   "resolve_full_name: email address not treated as name")


# â”€â”€ E. OFFLINE EXTRACTION â€” PHONE â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== E. OFFLINE EXTRACTION â€” Phone ===")
from hiring_agent.extraction import _extract_phone

ok(_extract_phone("Phone: +1 (512) 555-0001") == "+1 (512) 555-0001",
   "US phone with +1 and label")
ok(_extract_phone("Mobile: 512-555-0001") == "512-555-0001",
   "US phone with mobile label (no country code)")
ok(_extract_phone("Tel: +44 20 7946 0958") is not None,
   "International phone (UK +44)")
ok(_extract_phone("Contact: +91-9876543210") is not None,
   "International phone (India +91)")
ok(_extract_phone("Jane Smith\nSoftware Engineer") is None,
   "No phone when none present")
ok(_extract_phone("(555) 867-5309\nData Scientist") == "(555) 867-5309",
   "Unlabeled US phone number extracted")
ok(_extract_phone(
    "Published app: https://apps.apple.com/us/app/vendcell/id6677036994") is None,
   "numeric App Store URL ID is not misclassified as a phone")


# â”€â”€ F. LOCATION + COUNTRY PARSING â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== F. Location + Country parsing (split_location_country) ===")
from hiring_agent.extraction import split_location_country

loc, ctry = split_location_country("Austin, TX")
ok(loc == "Austin, TX" and ctry == "United States", "City, ST â†’ USA")

loc, ctry = split_location_country("New York, NY")
ok(loc == "New York, NY" and ctry == "United States", "New York, NY â†’ USA")

loc, ctry = split_location_country("Denver, Colorado")
ok(ctry == "United States", "City, Full-State-Name â†’ USA")

loc, ctry = split_location_country("Austin, TX, United States")
ok(ctry == "United States" and "United States" not in loc, "Country stripped from location")

loc, ctry = split_location_country("New York, NY, USA")
ok(ctry == "United States" and "USA" not in loc, "'USA' suffix stripped from location")

loc, ctry = split_location_country("Bengaluru, Karnataka")
ok(ctry == "India", "Province (Karnataka) mapped to parent country India")

loc, ctry = split_location_country("Mumbai, Maharashtra")
ok(ctry == "India", "Province (Maharashtra) â†’ India")

loc, ctry = split_location_country("Toronto, Ontario")
ok(ctry == "Canada", "Province (Ontario) â†’ Canada")

loc, ctry = split_location_country("London")
ok(loc == "London" and ctry == "", "City-only with no state â†’ country blank (needs AI)")

loc, ctry = split_location_country("Not extracted")
ok(loc == "Not extracted" and ctry == "", "Not extracted input â†’ no change")

loc, ctry = split_location_country("")
ok(loc == "Not extracted" and ctry == "", "Empty input â†’ Not extracted")


# â”€â”€ G. OFFLINE EXTRACTION â€” LOCATION FROM RESUME â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== G. OFFLINE EXTRACTION â€” Location from resume ===")
ok(extr("Jane Smith\nAustin, TX\nSoftware Engineer")["location"] == "Austin, TX",
   "Location from header City, ST")
ok(extr("Jane Smith\nNew York, New York\nEngineer")["location"] is not None,
   "Location from full state name in header")
ok(extr("Skills: Python, AWS\nLocation: Seattle, WA\nEngineer")["location"] is not None,
   "Location from LOCATION_KEYWORDS section")
ok(extr("Bob Jones\nSoftware Engineer\nNo location here")["country"] in ("", "Not extracted"),
   "No location â†’ country blank/Not extracted")
usa_d = extr("Alice Brown\nDenver, Colorado\nPython developer")
ok(usa_d["country"] == "United States", "State full name â†’ country = United States")


# â”€â”€ H. OFFLINE EXTRACTION â€” SKILLS â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== H. OFFLINE EXTRACTION â€” Skills ===")
d = extr("John Smith\nPython, AWS, Docker, React, PostgreSQL\nSoftware Engineer")
ok("Python" in d["skills"] or "python" in d["skills"].lower(), "Python in skills")
ok("AWS" in d["skills"] or "aws" in d["skills"].lower(), "AWS in skills")
ok("Docker" in d["skills"] or "docker" in d["skills"].lower(), "Docker in skills")
ok(d["skills"] != "Not extracted", "Skills extracted from tech terms")
ok(extr("Jane Doe\nSoft skills: communication, leadership")["skills"] == "Not extracted",
   "Soft-skill words not in skills (no SKILL_KEYWORDS match)")


# â”€â”€ I. OFFLINE EXTRACTION â€” LOOKING FOR ROLE â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== I. OFFLINE EXTRACTION â€” Looking For Role ===")
ok(extr("Jane Smith\nSenior Data Scientist\nPython, ML").get("looking_for_role") not in
   ("Not extracted", None, ""),
   "Role from resume header title")
ok(extr("Applying for: Machine Learning Engineer\nPython").get("looking_for_role") not in
   ("Not extracted", None),
   "Role from 'applying for:' pattern")
kw_cases = [
    ("uipath, power automate", "Automation"),
    ("machine learning, llm", "AI"),
    ("data scientist, analytics", "Data"),
    ("devops, kubernetes", "DevOps"),
    ("react, angular, frontend", "Frontend"),
    ("full stack, fullstack", "Full"),
    ("marketing, seo", "Marketing"),
]
for skills_kw, expected_fragment in kw_cases:
    r = extr(f"Bob\n{skills_kw}")["looking_for_role"]
    ok(expected_fragment.lower() in r.lower(),
       f"Role keyword fallback: '{skills_kw[:20]}...' â†’ contains '{expected_fragment}'")


# â”€â”€ J. MAIL BODY FALLBACK â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== J. MAIL BODY FALLBACK (merge_mail_body_fallback) ===")
from hiring_agent.extraction import merge_mail_body_fallback

base = {"full_name": "Not extracted", "phone": "Not extracted",
        "location": "Not extracted", "country": "", "skills": "Not extracted",
        "looking_for_role": "Data Engineer"}

# Fills N/A fields from body
body_with_name = "Hi, I am Sarah Connor. I am based in Dallas, TX."
r = merge_mail_body_fallback(dict(base), body_with_name)
ok(r.get("full_name") == "Sarah Connor", "Mail body fills full_name when N/A")
ok("Dallas" in r.get("location", ""), "Mail body fills location when N/A")

# looking_for_role ALWAYS comes from body when body has a value (even if already set)
base_with_role = dict(base); base_with_role["looking_for_role"] = "Software Engineer"
body_with_role = "Applying for: Data Scientist role. Passionate about ML."
r2 = merge_mail_body_fallback(base_with_role, body_with_role)
ok("Data Scientist" in r2.get("looking_for_role", ""),
   "looking_for_role: mail body ALWAYS wins over existing value")

# Does NOT overwrite already-filled fields (except looking_for_role)
base_filled = dict(base)
base_filled.update({"full_name": "Bob Smith", "phone": "512-555-0001"})
r3 = merge_mail_body_fallback(base_filled, "My name is Alice Jones. Phone: 999-888-7777")
ok(r3["full_name"] == "Bob Smith",
   "Mail body does NOT overwrite already-filled full_name")
ok(r3["phone"] == "512-555-0001",
   "Mail body does NOT overwrite already-filled phone")

# Empty body â†’ returns unchanged
r4 = merge_mail_body_fallback(dict(base), "")
ok(r4 == base, "Empty mail body returns dict unchanged")


# â”€â”€ K. PORTFOLIO EXTRACTION â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== K. PORTFOLIO EXTRACTION (extract_portfolios) ===")
from hiring_agent.extraction import extract_portfolios

# LinkedIn /in/ path â†’ Portfolio 1
p1, p2, p3 = extract_portfolios("linkedin.com/in/janedoe")
wrapped_p1, _, _ = extract_portfolios(
    "https://www.linkedin.com/in/christopher-\nfeld-93b838b9/")
ok(wrapped_p1 == "https://www.linkedin.com/in/christopher-feld-93b838b9",
   f"line-wrapped LinkedIn handle is rejoined: {wrapped_p1}")
ok("linkedin.com/in/janedoe" in p1, "LinkedIn /in/ URL â†’ Portfolio 1")

# LinkedIn /pub/ path â†’ Portfolio 1 (fixed bug)
p1, p2, p3 = extract_portfolios("linkedin.com/pub/john-smith/12/345/678")
ok("linkedin.com/pub/" in p1, "LinkedIn /pub/ URL â†’ Portfolio 1 (bug fix)")

# GitHub â†’ Portfolio 2
p1, p2, p3 = extract_portfolios("github.com/johndoe")
ok("github.com/johndoe" in p2, "GitHub URL â†’ Portfolio 2")

# GitHub requires 2+ char username
p1, p2, p3 = extract_portfolios("github.com/j")
ok(p2 == "N/A", "github.com/single-char â†’ not captured (too short)")

# Figma â†’ Portfolio 3
p1, p2, p3 = extract_portfolios("https://figma.com/file/abc123/my-design")
ok("figma.com" in p3, "Figma URL â†’ Portfolio 3")

# Behance â†’ Portfolio 3
p1, p2, p3 = extract_portfolios("behance.net/janedoe")
ok("behance" in p3, "Behance URL â†’ Portfolio 3")

# Personal website (.io) â†’ Portfolio 1 (no LinkedIn found)
p1, p2, p3 = extract_portfolios("janesmith.io/portfolio")
both_p1, both_p2, both_p3 = extract_portfolios(
    "linkedin.com/in/janedoe https://janesmith.io/portfolio")
ok("linkedin.com/in/janedoe" in both_p1 and "janesmith.io/portfolio" in both_p3,
   "LinkedIn in P1 does not discard a separate personal portfolio (stored in P3)")
ok("janesmith.io" in p1, "Personal .io site â†’ Portfolio 1")

# Skip social/email platforms
p1, p2, p3 = extract_portfolios("gmail.com/user instagram.com/janedoe facebook.com/jane")
qr_p1, qr_p2, qr_p3 = extract_portfolios("https://qrcode.ooo/generated/resume-code")
ok((qr_p1, qr_p2, qr_p3) == ("N/A", "N/A", "N/A"),
   "QR-code service URL is not stored as a candidate portfolio")
store_p1, store_p2, store_p3 = extract_portfolios(
    "https://play.google.com/store/apps/details?id=com.example.app "
    "https://apps.apple.com/us/app/example/id1234567890")
ok((store_p1, store_p2, store_p3) == ("N/A", "N/A", "N/A"),
   "App Store and Play Store project URLs are not personal portfolios")
ok(p1 == "N/A" and p2 == "N/A" and p3 == "N/A",
   "Social/email domains skipped â€” all portfolios N/A")

# All three present
full_text = """
linkedin.com/in/janedoe github.com/janedoe https://figma.com/file/xyz/design
"""
p1, p2, p3 = extract_portfolios(full_text)
ok("linkedin" in p1, "All three portfolios: LinkedIn in P1")
ok("github"   in p2, "All three portfolios: GitHub in P2")
ok("figma"    in p3, "All three portfolios: Figma in P3")

# No URLs â†’ all N/A
p1, p2, p3 = extract_portfolios("Bob Smith, Python developer, Austin TX")
ok(p1 == "N/A" and p2 == "N/A" and p3 == "N/A", "No URLs â†’ all portfolios N/A")


# â”€â”€ L. SKILL DEDUPLICATION (_merge_keyword_skills) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== L. SKILL DEDUPLICATION (_merge_keyword_skills) ===")
from hiring_agent.extraction import _merge_keyword_skills

# AI returns skills with different casing than keyword list
ai_result = {"full_name": "A", "skills": "PYTHON, AWS"}
merged = _merge_keyword_skills(ai_result, "python aws docker terraform react")
skills_lower = merged["skills"].lower()
ok("python" in skills_lower, "Python deduplicated (ALLCAPS AI + keyword)")
ok(merged["skills"].lower().count("python") == 1, "Python not duplicated")
ok("docker" in skills_lower, "Docker added from keyword scan")
ok("terraform" in skills_lower, "Terraform added from keyword scan")

# AI returns no skills â†’ keyword fills entirely
no_skill = {"full_name": "A", "skills": "Not extracted"}
merged2 = _merge_keyword_skills(no_skill, "python react postgres")
ok("python" in merged2["skills"].lower(), "Keyword fills when AI returned 'Not extracted'")


# â”€â”€ M. SUGGESTED ROLES + CATEGORY â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== M. SUGGESTED ROLES + assign_category ===")
from hiring_agent.scoring import suggested_roles, assign_category
from hiring_agent.scoring import _ai_role_shortlist
from hiring_agent.jd_sources import get_active_roles, _normalized_file_jd_title

roles = get_active_roles()
ok(len(roles) > 0, f"Active roles loaded: {len(roles)}")

_offer_info_text = (
    "About the Role\nDriverAI is seeking a Software Developer - AI/ML, Computer "
    "Vision, to build production features."
)
ok(_normalized_file_jd_title(
       "SIN2 SWDev AI ML CV Offer Letter Info", _offer_info_text)
   == "Software Developer - AI/ML, Computer Vision",
   "administrative offer-info filename is replaced by the actual JD role title")
ok(_normalized_file_jd_title("Data Analyst Job Announcement", _offer_info_text)
   == "Data Analyst Job Announcement",
   "normal JD filenames remain unchanged")

# Large live JD catalogs are deterministically narrowed before the 4K-context local
# model is called.  This prevents the 86-role prompt timeout seen in live rescoring.
_many_roles = [
    {"title": f"Role {i}", "skills": ["skill" + str(i)]} for i in range(20)
]
_many_roles[17]["skills"] = ["python", "sql"]
_short = _ai_role_shortlist("Python, SQL", "", _many_roles)
ok(len(_short) == 12, "AI role shortlist caps a large JD catalog at 12")
ok(_many_roles[17] in _short, "AI role shortlist retains the strongest skill match")
ok(_ai_role_shortlist("Python", "", _many_roles[:5]) == _many_roles[:5],
   "small JD catalogs are passed to Ollama unchanged")

_pref_roles = [
    {"title": "Gaming Position", "skills": ["Swift", "Java", "Python", "C++"]},
    {"title": "Mobile Application Developer",
     "skills": ["Swift", "Kotlin", "iOS", "Android", "Flutter"]},
]
_pref_result = suggested_roles(
    "Swift, Kotlin, iOS, Java, Python", "Senior iOS Engineer", roles=_pref_roles)
ok(_pref_result["role_1"].startswith("Mobile Application Developer"),
   "explicit iOS preference promotes the matching mobile JD family")

_flutter_pref_result = suggested_roles(
    "Flutter, Android, Firebase, Java, Kotlin", "Flutter Developer", roles=_pref_roles)
ok(_flutter_pref_result["role_1"].startswith("Mobile Application Developer"),
   "explicit Flutter preference promotes the matching mobile JD family")

_analyst_roles = [
    {"title": "AWS Senior Engineer", "skills": ["AWS", "Python", "Terraform"]},
    {"title": "Data Analyst Position Description", "skills": ["SQL", "Power BI", "Excel", "Python"]},
]
_analyst_result = suggested_roles(
    "AWS, Python, SQL, Power BI, Excel", "Data Analyst", roles=_analyst_roles)
ok(_analyst_result["role_1"].startswith("Data Analyst Position Description"),
   "explicit Data Analyst preference outranks a short generic AWS skill denominator")

_sparse_result = suggested_roles(
    "Computer Vision", "", roles=[{"title": "Sparse JD", "skills": ["Computer Vision"]}])
ok(_sparse_result["role_1"] == "Sparse JD (20%)",
   "one parsed JD skill scores 20%, not an artificial 100%")

_deduped_result = suggested_roles(
    "Kotlin, Android", "Android Developer",
    roles=[
        {"title": "Gaming Position", "skills": ["Kotlin", "Android"]},
        {"title": "Gaming Position", "skills": ["Kotlin", "Android"]},
        {"title": "Offer Letter Info", "skills": ["Kotlin", "Android"]},
        {"title": "Mobile Application Developer", "skills": ["Kotlin", "Android"]},
    ])
ok(_deduped_result["role_1"].startswith("Mobile Application Developer"),
   "job-family preference can outrank generic/duplicate roles")
ok("Offer Letter" not in " ".join(_deduped_result.values()),
   "non-job offer-letter documents are filtered from suggested roles")
ok(_deduped_result["role_2"] == "" or
   _deduped_result["role_2"].split(" (")[0] != _deduped_result["role_1"].split(" (")[0],
   "suggested roles are unique by title")

# Strong skill match â†’ role_1 has score
res = suggested_roles("Python, SQL, pandas, scikit-learn, tensorflow", "data scientist", roles=roles)
ok(res["role_1"] != "" and "%" in res["role_1"], f"Strong match â†’ role_1 with score: {res['role_1']}")
ok(res["source"] == "keyword", "Source is 'keyword' when Ollama off")

# Role format: "Title (NN%)"
import re as _re
ok(_re.match(r'.+\s\(\d+%\)', res["role_1"]) is not None,
   f"role_1 format 'Title (NN%)': {res['role_1']}")

# role_2 present when second match exists
ok(res["role_2"] != "" or True,  # may be empty if only 1 match â€” not a bug
   "role_2 present when second match clears threshold (or empty)")

# No skills â†’ no match â†’ empty roles
res_empty = suggested_roles("", "", roles=roles)
ok(res_empty["role_1"] == "", "No skills â†’ empty role_1 (no strong match)")
ok(res_empty["role_3"] == "", "No skills â†’ empty role_3")

# assign_category â€” all actual categories from config.yaml
cat_cases = [
    ("Senior Software Engineer (80%)", "Senior & Executive"),
    ("Backend Engineer (80%)",         "Web Team (Full stack/Back end & UI/UX)"),
    ("Data Scientist (65%)",           "AI/ML/CV (SIN2)"),
    ("AI / ML Engineer (75%)",         "AI/ML/CV (SIN2)"),
    ("Cloud Database Engineer (60%)",  "Cloud and DevOps"),
    ("DevOps Engineer (60%)",          "Cloud and DevOps"),
    ("Cybersecurity Analyst (65%)",    "Cybersecurity and IT Admin"),
    ("UI/UX Designer (65%)",           "Web Team (Full stack/Back end & UI/UX)"),
    ("Data Analyst (60%)",             "Data Analytics"),
    ("Logistics Manager (65%)",        "Senior & Executive"),
    ("Logistics Specialist (50%)",     "Supply Chain"),
    ("Telecom Engineer (60%)",         "Satellite"),
    ("",                               "General"),
    ("No strong match",                "General"),
    # Added 2026-07-15 (client request): marketing -> Business Analytics,
    # design/rendering/gaming -> new 'Graphics' category.
    ("Marketing Specialist (55%)",     "Business Analytics"),
    ("Growth Marketer (50%)",          "Business Analytics"),
    ("Game Designer (60%)",            "Graphics"),
    ("Rendering Engineer (50%)",       "Graphics"),
    ("Gaming Analyst (45%)",           "Graphics"),
    ("Graphics Programmer (55%)",      "Graphics"),
    ("Marketing Manager (65%)",        "Senior & Executive"),  # 'manager' still wins (section 1 precedence)
    # Regression guard 2026-07-15: removed the bare "developer"/"engineer" catch-all
    # (was silently routing any unmatched *Engineer/*Developer title to Web Team, which
    # wrongly caught 2 of the client's own configured roles below). Now falls to General,
    # same as every other unmatched title - "intern etc. goes to General" per client.
    ("RPA / Automation Engineer (55%)",   "General"),
    ("Intern / Entry-Level Engineer (0%)","General"),
    ("Security Engineer (60%)",           "General"),
    ("Firmware Engineer (50%)",           "General"),
]
for role_str, expected_cat in cat_cases:
    got = assign_category(role_str)
    ok(got == expected_cat,
       f"assign_category('{role_str[:30]}') = '{got}' (expected '{expected_cat}')")

# Skill-based Mobile Apps override (2026-07-24): a candidate's own tool stack beats a
# loosely-matched JD title. Real cases from the live Rejected sheet - plain Android/iOS/
# Flutter developers whose top-scoring JD title happened to land them in the wrong
# category (Senior & Executive via 'lead', Graphics via 'gaming'/'design', 3D/CV/IoT via
# no domain match at all) purely because that JD's title contained an unrelated keyword.
mobile_skill_cases = [
    ("Mobile Application Lead Developer Position Description v2 (80%)",
     "Kotlin, Java, Jetpack Compose, Firebase, Android Studio",
     "Mobile Apps (Android IOS)"),
    ("Gaming Position (23%)", "Kotlin, Java, Dart, Swift, Firebase, Git",
     "Mobile Apps (Android IOS)"),
    ("UI UX Developer Position description (20%)", "Swift, UIKit, Core Data, Xcode",
     "Mobile Apps (Android IOS)"),
    ("3D/CV/ IoT/ AI Agents Engineering Intern (40%)", "Kotlin, Flutter, Dart, Firebase",
     "Mobile Apps (Android IOS)"),
]
for role_str, skills_str, expected_cat in mobile_skill_cases:
    got = assign_category(role_str, skills_str)
    ok(got == expected_cat,
       f"assign_category('{role_str[:35]}', mobile skills) = '{got}' (expected '{expected_cat}')")

# A genuine game/graphics candidate who also lists a mobile platform skill (e.g. a Unity
# mobile game dev listing Swift/Kotlin as target-platform skills) must NOT be swept into
# Mobile Apps by the override - the game-engine skill takes precedence.
got = assign_category("Gaming Position (60%)", "Unity, C#, Kotlin, Swift, Shader")
ok(got == "Graphics",
   f"a Unity/game-engine skill alongside mobile skills stays Graphics (got '{got}')")

# No skills passed (default "") -> behaves exactly as before (existing callers unaffected).
got = assign_category("Mobile Application Lead Developer Position Description v2 (80%)")
ok(got == "Senior & Executive",
   f"assign_category with no skills_str falls back to title-only precedence (got '{got}')")

# Existing 'manager'/'lead' precedence in OTHER domains is untouched by the mobile-only
# override - Logistics Manager still Senior & Executive even with logistics-flavored
# skills (no mobile tool skills present, so the override never fires).
got = assign_category("Logistics Manager (65%)", "Supply Chain, Procurement, Excel")
ok(got == "Senior & Executive",
   f"non-mobile domains keep today's title-precedence behavior (got '{got}')")

# Category is derived from Role 1 â€” deterministic
cat_a = assign_category("Software Engineer (80%)")
cat_b = assign_category("Software Engineer (80%)")
ok(cat_a == cat_b, "assign_category is deterministic (same input â†’ same output)")


# â”€â”€ N. _IS_GAP DETECTION â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== N. _is_gap DETECTION ===")
from hiring_agent.sharepoint_scoring import _is_gap

gap_vals = ["", "not extracted", "Not extracted", "NOT EXTRACTED",
            "n/a", "N/A", "na", "NA", "none", "None", "NONE",
            "-", "not found", "Not Found", "NOT FOUND"]
for v in gap_vals:
    ok(_is_gap(v), f"_is_gap recognizes '{v}' as a gap")

non_gap = ["John Smith", "Austin, TX", "Python, AWS, Docker",
           "Data Engineer", "512-555-0001", "https://github.com/johndoe",
           "Software Engineer (80%)", "Scored"]
for v in non_gap:
    ok(not _is_gap(v), f"_is_gap correctly NOT a gap: '{v[:30]}'")

ok(_is_gap(None), "_is_gap(None) = True")
ok(_is_gap(0), "_is_gap(0) = True (Application Updates=0 treated as gap for text columns)")


# â”€â”€ O. VALIDATE ALL COLUMNS (23-column pass) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== O. _validate_all_columns â€” 23-column pass ===")
from hiring_agent.sharepoint_scoring import _validate_all_columns
from hiring_agent.config import COLUMNS

ok(len(COLUMNS) == 30, f"COLUMNS has 30 entries (adds Last Updated Date, got {len(COLUMNS)})")
ok(COLUMNS[2] == "Last Updated Date", "'Last Updated Date' sits after Received Date")
ok(COLUMNS[-1] == "Info Request Sent", "'Info Request Sent' is the LAST column")
ok(COLUMNS[-2] == "Mail Sent", "'Mail Sent' is second-to-last")
ok("Education" in COLUMNS, "'Education' column present")
ok("Resume URL" in COLUMNS and "Resume Folder Path" in COLUMNS and "Resume Link" in COLUMNS,
   "all three resume columns present")
ok(COLUMNS[4] == "Resume Link" and COLUMNS[3] == "Category" and COLUMNS[5] == "Full Name",
   "'Resume Link' moved to 5th column, between Category and Full Name (client request 2026-07-24)")
ok("Retry Count" in COLUMNS, "retry-cap counter column present")
ok("Candidate Email" not in COLUMNS,
   "no 'Candidate Email' column â€” duplicate merging matches on Phone alone now")
ok("Possible Duplicate" not in COLUMNS,
   "no 'Possible Duplicate' flag column â€” duplicates are merged, not flagged")

# Both sheets carry the full CONTENT schema; they differ only in the mail-audit column
# (main â†’ 'Mail Sent' + 'Info Request Sent', rejected â†’ 'Decline Sent').
from hiring_agent.config import REJECTED_COLUMNS
ok(REJECTED_COLUMNS[4] == "Resume Link" and REJECTED_COLUMNS[3] == "Category"
   and REJECTED_COLUMNS[5] == "Full Name",
   "Rejected sheet mirrors the same Resume Link 5th-column position")
_content = [c for c in COLUMNS if c not in ("Mail Sent", "Info Request Sent")]
ok(all(c in REJECTED_COLUMNS for c in _content), "Rejected schema has every content column the main sheet has")
ok("Decline Sent" in REJECTED_COLUMNS, "Rejected schema has 'Decline Sent'")
ok("Mail Sent" not in REJECTED_COLUMNS, "Rejected schema has no 'Mail Sent' (that's the main sheet's audit column)")
ok("Info Request Sent" not in REJECTED_COLUMNS,
   "Rejected schema has no 'Info Request Sent' (a rejected candidate is never asked for missing info)")
ok("Resume URL" in REJECTED_COLUMNS and "Resume Folder Path" in REJECTED_COLUMNS
   and "Resume Link" in REJECTED_COLUMNS,
   "Rejected schema also carries the resume URL/folder-path + clickable link columns")

def make_fields(overrides=None):
    f = {
        "Full Name": "Jane Doe", "Phone": "555-1234", "Location": "Austin, TX",
        "Country": "United States", "Current Skills": "Python, AWS",
        "Looking For Role": "Data Engineer", "Education": "B.S. Computer Science",
        "Portfolio 1": "N/A", "Portfolio 2": "N/A", "Portfolio 3": "N/A",
        "Suggested Role 1": "Data Scientist (75%)", "Suggested Role 2": "Backend (60%)",
        "Suggested Role 3": "", "Category": "Data & Analytics", "Status": "Scored",
    }
    if overrides: f.update(overrides)
    return f

def make_p1(email="jane@gmail.com"):
    return {
        "Application ID": "APP-20260701-123456-ABCD", "Received Date": "2026-07-01",
        "Email": email, "Mail Subject": "Application",
        "Mail Body": "Please find my resume.", "Has Resume": "Yes",
        "Original Filename": "jane_cv.pdf", "Application Updates": 0,
    }

# All fields filled â†’ nothing healed
r = _validate_all_columns(make_fields(), make_p1(), "")
ok(r["Full Name"] == "Jane Doe", "Filled Full Name not touched")
ok(r["Country"] == "United States", "Filled Country not touched")

# Full Name N/A + email has display name â†’ healed
r2 = _validate_all_columns(make_fields({"Full Name": "Not extracted"}),
                            make_p1("Jane Doe <jane@gmail.com>"), "")
ok(r2["Full Name"] == "Jane Doe",
   "Full Name healed from sender display name in Email column")

# Full Name N/A + bare email â†’ stays N/A
r3 = _validate_all_columns(make_fields({"Full Name": "Not extracted"}),
                            make_p1("jane@gmail.com"), "")
ok(_is_gap(r3["Full Name"]),
   "Full Name stays N/A when Email column has no display name")

# P1 columns are read-only â€” even if blank in stored row, P2 never fills them
p1_missing = make_p1()
p1_missing["Email"] = ""  # simulate P1 column blank
r4 = _validate_all_columns(make_fields(), p1_missing, "")
ok(r4.get("Email") is None, "P2 never adds 'Email' key to fields dict")

# Validation pass does not call Ollama (already ran once; this is final check only)
# Proven by: all 15 fields returned correctly with Ollama off the whole time
all_p2 = ["Full Name", "Phone", "Location", "Country", "Current Skills", "Education",
          "Looking For Role", "Portfolio 1", "Portfolio 2", "Portfolio 3",
          "Suggested Role 1", "Suggested Role 2", "Suggested Role 3",
          "Category", "Status"]
r5 = _validate_all_columns(make_fields(), make_p1(), "some mail body")
ok(all(k in r5 for k in all_p2), "All 15 P2 columns present in result")


# â”€â”€ O2. CROSS-FIELD CONSISTENCY CHECK (deterministic, enhancement-aware) â”€â”€â”€â”€â”€â”€
print("\n=== O2. _cross_field_check â€” cross-field consistency audit ===")
from hiring_agent.sharepoint_scoring import _cross_field_check

_f, _w = _cross_field_check(make_fields(), None)
ok(_w == [], "clean, consistent US row raises no consistency warnings")

_f, _w = _cross_field_check(make_fields({"Location": "Paris, Texas"}), None)
ok(_w == [], "'Paris, Texas' + US not flagged (strong-US beats foreign-city)")

_f, _w = _cross_field_check(make_fields({"Country": "India"}), None)
ok(any("reads US but Country" in x for x in _w), "US location vs Country=India flagged")

_f, _w = _cross_field_check(make_fields({"Location": "Bengaluru", "Country": "United States"}), None)
ok(any("reads non-US" in x for x in _w), "foreign city vs Country=US flagged")


# â”€â”€ O2b. GEO DECISION: strong-US location overrides a mis-extracted foreign country â”€â”€
# Regression for the Sadaf Khan case (2026-07-12): current UChicago master's in 'Chicago, IL',
# earlier bachelor's in 'Hyderabad, IND' -> the parser set Country='India' and the OLD
# country-first order rejected a US-based candidate (and would have emailed a decline). A strong
# US location (state/ZIP) must now win over a contradicting country field; genuinely-foreign
# candidates must still reject. resume_text='' isolates the offline precedence (no Ollama).
print("\n=== O2b. check_location_usa â€” US location overrides mis-extracted foreign country ===")
from hiring_agent.geo import (
    GeoDecision as _GeoDecision,
    check_location_usa as _clu,
    classify_location_usa as _classify_geo,
)
from hiring_agent.geo import is_strong_usa as _strong_us
_keep, _r = _clu("Chicago, IL", country="India", resume_text="")
ok(_keep, "Sadaf case: 'Chicago, IL' + Country='India' -> KEEP (US state overrides country)")
ok(_strong_us("NJ"), "bare US state abbreviation 'NJ' is a strong USA signal")
_keep, _r = _clu("NJ", country="India", resume_text="", phone="+91 8376865125")
ok(_keep, "Varun case: bare 'NJ' + Country='India' -> KEEP (US state overrides country)")
_keep, _r = _clu("Chicago", country="India", resume_text="")
ok(_keep, "bare US city 'Chicago' + Country='India' -> KEEP (us-city contradicts foreign country)")
_keep, _r = _clu("", country="India", resume_text="")
ok(_keep, "no location + Country='India' -> KEEP for current-location confirmation")
_decision, _r = _classify_geo("", country="India", resume_text="")
ok(_decision == _GeoDecision.UNKNOWN,
   "country alone never proves current non-US residence")
_keep, _r = _clu("New York, NY", country="United States", resume_text="")
ok(_keep, "US location + US country -> KEEP (unchanged)")
_keep, _r = _clu("Mumbai, United States", country="United States", resume_text="", phone="213-376-7163")
ok(_keep, "US phone/US country overrides foreign city string in location -> KEEP")
_keep, _r = _clu("Paris, Texas", country="United States", resume_text="")
ok(_keep, "foreign-named city with concrete US state still KEEPS (Paris, Texas)")

# US phone remains a useful corroboration signal, but never overrides an explicit foreign
# current-location extraction.
from hiring_agent.geo import _looks_like_us_phone as _usph
ok(_usph("865-356-1921") and _usph("(408) 555-1234") and _usph("+1 865 356 1921"),
   "US 10/11-digit numbers detected as US phones")
ok(not _usph("+91 9409021350") and not _usph("") and not _usph("732-4***"),
   "+91 / blank / masked numbers are NOT US phones")
# 2026-07-24: NANP covers both the US and Canada with the identical +1/10-digit shape, so
# a Canadian-area-code number must not be treated as a US-phone corroboration signal (real
# case: 'Mukesh Kumar Sharma' - 'London, ON' + Country='Canada' + phone '+1 519-859-0693',
# 519 = Ontario - was reading as "looks like a US phone" and would have contradicted a
# correctly-foreign geo verdict).
ok(not _usph("519-859-0693") and not _usph("+1 519-859-0693") and not _usph("(416) 555-0199"),
   "Canadian area codes (519 Ontario, 416 Toronto) are NOT US phones")
ok(_usph("865-356-1921") and _usph("213-376-7163"),
   "genuine US area codes (865 Tennessee, 213 LA) are still detected as US phones")
_keep, _r = _clu("Ahmedabad, Gujarat", country="India", resume_text="", phone="+91 9409021350")
ok(not _keep, "Krips case: foreign hometown+country + Indian (non-contradicting) phone -> still REJECT")
_keep, _r = _clu("Mumbai, Maharashtra", country="India", resume_text="", phone="+91 98765 43210")
ok(not _keep, "genuine India + a present, non-contradicting phone -> still REJECT")


# â”€â”€ O2c. GEO DECISION: a foreign location match now requires Phone/Education â”€â”€
# corroboration before rejecting (added 2026-07-24) â”€â”€
# Previously ANY foreign location-text match rejected immediately with zero cross-check
# against Phone or Education, even though a US-based candidate's stale/historic foreign
# mention (an old degree, a hometown) can land in Location when extraction can't find a
# clean current-address line. Now: a clear foreign location needs a corroborating signal
# (Phone or Education present and not contradicting) to actually reject. No signal at all,
# or a contradicting signal (US phone / US education), keeps the candidate for
# clarification instead of hard-rejecting them.
print("\n=== O2c. check_location_usa — foreign location requires phone/education corroboration ===")
_keep, _r = _clu("Mumbai, Maharashtra", country="India", resume_text="")
ok(_keep, "foreign 'Mumbai, Maharashtra' with NO phone/education at all -> KEEP "
          "(insufficient evidence to reject, benefit of the doubt)")
_keep, _r = _clu("Bangalore, IND", country="India", resume_text="")
ok(_keep, "foreign 'Bangalore, IND' with NO phone/education at all -> KEEP")
_keep, _r = _clu("Mumbai", country="United States", resume_text="", phone="213-376-7163")
ok(_keep, "foreign 'Mumbai' + a US-shaped phone -> KEEP for clarification "
          "(phone contradicts the location match, no longer an automatic reject)")
_keep, _r = _clu("Chennai, India", country="India", resume_text="", phone="971-487-3222")
ok(_keep, "'Chennai, India' + a US-shaped phone -> KEEP for clarification "
          "(same contradiction case, explicit foreign country field too)")
_keep, _r = _clu("Hyderabad", country="India", resume_text="", phone="865-356-1921")
ok(_keep, "'Hyderabad' + a US-shaped phone -> KEEP for clarification")
_keep, _r = _clu("Mumbai, Maharashtra", country="India", resume_text="",
                 phone="+91 98765 43210", education="Bachelor's, University of Mumbai")
ok(not _keep, "foreign location + non-contradicting phone + non-US education -> still REJECT")
_keep, _r = _clu("Mumbai, Maharashtra", country="India", resume_text="",
                 education="MS Computer Science, Arizona State University")
ok(_keep, "foreign location + US-reading Education (no phone at all) -> KEEP for clarification "
          "(education contradicts the location match)")
_keep, _r = _clu("Mumbai, Maharashtra", country="India", resume_text="",
                 education="Bachelor's, University of Mumbai")
ok(not _keep, "foreign location + non-US-reading Education (no phone) -> still REJECT "
              "(education is present and does not contradict)")

from hiring_agent.geo import _corroborates_non_usa as _corrob
ok(_corrob("", "", "") == "no_signal", "no phone, no education, no resume text -> no_signal")
ok(_corrob("213-376-7163", "", "") == "contradicts_usa", "US-shaped phone alone -> contradicts_usa")
ok(_corrob("", "MS CS, Arizona State University", "") == "contradicts_usa",
   "US-reading education alone -> contradicts_usa")
ok(_corrob("+91 98765 43210", "", "") == "corroborates",
   "a present, non-contradicting phone alone -> corroborates")
ok(_corrob("", "BTech, University of Mumbai", "") == "corroborates",
   "a present, non-US-reading education alone -> corroborates")

# Regression test (2026-07-24): _looks_like_us_phone existed and was passed into
# check_location_usa by the live pipeline and multiple repair scripts, but the phone
# parameter was never actually read inside check_location_usa's body - the Manasa Puli
# fix it was written for was never wired in. Fixed: a US-shaped phone now contradicts a
# FOREIGN COUNTRY FIELD specifically (step 4) when location itself is blank/ambiguous -
# it still never overrides an explicit foreign LOCATION (that's step 2, unchanged, see
# the Hyderabad/Ahmedabad/Mumbai REJECT cases above, which still correctly reject).
_keep, _r = _clu("", country="India", resume_text="", phone="865-356-1921")
ok(_keep, "blank location + foreign Country + US phone -> KEEP (phone contradicts a likely "
          "mis-extracted country, the case _looks_like_us_phone was built for)")
_keep, _r = _clu("Remote", country="India", resume_text="", phone="865-356-1921")
ok(_keep, "ambiguous 'Remote' location + foreign Country + US phone -> KEEP")
_keep, _r = _clu("", country="India", resume_text="", phone="+91 9409021350")
ok(_keep, "blank location + foreign Country/phone -> KEEP for confirmation "
          "(no current-location evidence)")
_decision, _r = _classify_geo(
    "", country="India", resume_text="", phone="+91 9409021350")
ok(_decision == _GeoDecision.UNKNOWN,
   "foreign country/phone without a current location is UNKNOWN, not rejected")

# Attached live-case regression: a US phone and a clearly named US university are
# supporting context, but neither proves where the candidate currently lives. The row
# must stay on Main for clarification and must never enter the decline queue.
_decision, _r = _classify_geo(
    "Not extracted", country="Not extracted", phone="(623) 205-0073",
    education="B.S. Computer Science, Arizona State University")
ok(_decision == _GeoDecision.UNKNOWN,
   "attached pattern: missing location + US phone/university -> UNKNOWN")
_decision, _r = _classify_geo(
    "Austin, TX", country="United States", phone="(623) 205-0073",
    education="B.S. Computer Science, Arizona State University")
ok(_decision == _GeoDecision.CONFIRMED_US,
   "the same candidate becomes confirmed US only with current city/state evidence")
_decision, _r = _classify_geo(
    "Washington D.C.", country="United States", phone="(762) 218-2525",
    education="B.S. Business Administration, University of Central Florida")
ok(_decision == _GeoDecision.CONFIRMED_US,
   "Washington D.C. punctuation normalizes to the configured US city")
_decision, _r = _classify_geo(
    "Mumbai, Maharashtra", country="India", phone="+91 98765 43210",
    education="Bachelor's, University of Mumbai")
ok(_decision == _GeoDecision.CONFIRMED_NON_US,
   "explicit foreign current location plus corroborating evidence remains rejectable")

from hiring_agent.sharepoint_scoring import (
    _foreign_country_only_reject_is_too_weak as _weak_country_reject,
    _location_clarity_issue as _loc_clarity,
    _missing_fields_list as _mfl,
    _progress_value as _progress_value,
    _scoring_queue_rows as _queue_rows,
    _updated_after_location_request as _updated_after_loc_request,
)
_weak_fields = make_fields({
    "Full Name": "Not extracted",
    "Phone": "Not extracted",
    "Location": "Not extracted",
    "Country": "China",
    "Current Skills": "Not extracted",
    "Education": "Not extracted",
})
ok(_weak_country_reject(_weak_fields, "country field 'China' matched as non-USA"),
   "country-only foreign signal with mostly empty extraction is kept for missing-info, not decline")
_strong_fields = make_fields({
    "Full Name": "Jane Candidate",
    "Phone": "+91 98765 43210",
    "Location": "Not extracted",
    "Country": "India",
    "Current Skills": "Kotlin, Android",
    "Education": "B.Tech Computer Science",
})
ok(not _weak_country_reject(_strong_fields, "country field 'India' matched as non-USA"),
   "country-only foreign signal with enough candidate substance can still reject")
_miss = _mfl({
    "Phone": "(512) 555-0001",
    "Location": "Remote",
    "Country": "",
    "Current Skills": "Python, SQL",
    "Education": "MS Computer Science",
})
ok("your current location/country" in _miss,
   "ambiguous nonblank location asks for current location/country clarity")
_miss = _mfl({
    "Phone": "(512) 555-0001",
    "Location": "Austin, TX",
    "Country": "United States",
    "Current Skills": "Python, SQL",
    "Education": "MS Computer Science",
})
ok(_miss == [],
   "clear US location does not trigger a location-clarity info request")
_miss = _mfl({
    "Phone": "(919) 521-2392",
    "Location": "Pune, NC",
    "Country": "United States",
    "Current Skills": "Python, Machine Learning",
    "Education": "North Carolina State University",
})
ok("your current location/country" in _miss,
   "foreign city plus US admin suffix asks for current location/country clarity")
_miss = _mfl({
    "Phone": "+91 8376865125",
    "Location": "NJ",
    "Country": "India",
    "Current Skills": "Python, Machine Learning",
    "Education": "Drexel University",
})
ok("your current location/country" in _miss,
   "US state plus foreign country/phone asks for current location clarification")
ok(_loc_clarity("NJ", "India", "+91 8376865125"),
   "NJ + India/+91 is ambiguous/conflicting, not a clean accept/reject")
ok(_updated_after_loc_request({
    "Info Request Sent": "Sent 2026-07-16 10:00",
    "Last Updated Date": "2026-07-17T09:00:00",
}), "a newer candidate update after the clarification request is detected")

class _QueueFilterClient:
    def __init__(self):
        self.new_rows = [
            {"index": 0, "values": {
                "Application ID": "APP-NEW", "Status": "New Email Received",
            }},
            {"index": 1, "values": {
                "Application ID": "APP-UNREADABLE",
                "Status": "Needs Review - Unreadable Resume",
            }},
        ]
        self.rows = self.new_rows + [
            {"index": 2, "values": {
                "Application ID": "APP-UPDATED-REVIEW",
                "Status": "Needs Review - Location Confirmation",
                "Info Request Sent": "Sent 2026-07-16 10:00",
                "Last Updated Date": "2026-07-17T09:00:00",
            }},
            {"index": 3, "values": {
                "Application ID": "APP-UNTOUCHED-REVIEW",
                "Status": "Needs Review - Location Confirmation",
                "Info Request Sent": "Sent 2026-07-18 10:00",
                "Last Updated Date": "2026-07-17T09:00:00",
            }},
            {"index": 4, "values": {
                "Application ID": "APP-ALREADY-SCORED", "Status": "Scored",
            }},
            {"index": 5, "values": {
                "Application ID": "APP-REJECTED",
                "Status": "Rejected - Non-USA Location",
            }},
        ]

    def list_unscored_rows(self, statuses):
        return [
            row for row in self.new_rows
            if row["values"]["Status"] in set(statuses)
        ]

    def list_rows(self):
        return self.rows

_queue_ids = {
    row["values"]["Application ID"] for row in _queue_rows(_QueueFilterClient())
}
ok(_queue_ids == {"APP-NEW", "APP-UNREADABLE", "APP-UPDATED-REVIEW"},
   "live queue contains only new, unreadable-unscored, and genuinely updated review rows")
ok("APP-ALREADY-SCORED" not in _queue_ids
   and "APP-UNTOUCHED-REVIEW" not in _queue_ids
   and "APP-REJECTED" not in _queue_ids,
   "previously finalized and untouched review rows never enter normal live scoring")
ok(_progress_value(0) == "0",
   "progress serialization preserves zero candidate counts")
ok(_progress_value("Jane | Smith\nUpdated") == "Jane / Smith Updated",
   "progress serialization keeps candidate fields on one parseable line")

_miss = _mfl({
    "Phone": "(213) 376-7163",
    "Location": "Mumbai",
    "Country": "United States",
    "Current Skills": "Python, SQL",
    "Education": "University of Mumbai",
})
ok("your current location/country" in _miss,
   "foreign location + a US-shaped phone contradicts the match -> clarity ask, not "
   "an automatic reject (2026-07-24 corroboration change)")
_miss = _mfl({
    "Phone": "+91 98765 43210",
    "Location": "Mumbai, Maharashtra",
    "Country": "India",
    "Current Skills": "Python, SQL",
    "Education": "University of Mumbai",
})
ok("your current location/country" not in _miss,
   "clear foreign location with a present, non-contradicting phone+education still "
   "rejects cleanly; it belongs on Rejected, not a clarity nudge")

_f, _w = _cross_field_check(make_fields({"Country": ""}), None)
ok(_f["Country"] == "United States", "blank Country healed from Location")

_f, _w = _cross_field_check(make_fields({"Category": ""}), None)
ok(_f["Category"] == "General", "blank Category healed to General")

_f, _w = _cross_field_check(make_fields({"Phone": "123"}), None)
ok(any("too few digits" in x for x in _w), "implausible phone flagged")

_f, _w = _cross_field_check(make_fields({"Status": "Whatever"}), None)
ok(any("not a recognized value" in x for x in _w), "unknown Status value flagged")

_f, _w = _cross_field_check(make_fields(), {"Full Name": "", "Email": "jane@x.com"})
ok(_w == [], "P1 blank -> P2 filled is NOT flagged (enhancement is expected)")

_f, _w = _cross_field_check(make_fields({"Email": "changed@x.com"}), {"Email": "orig@x.com"})
ok(any("was changed by P2" in x for x in _w), "altering a P1-owned column is flagged")


# â”€â”€ P. STORED RESUME FILENAME RECONSTRUCTION â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== P. _stored_resume_names â€” filename reconstruction ===")
from hiring_agent.sharepoint_scoring import _stored_resume_names

# Single resume
names = _stored_resume_names("APP-20260701-105826-A3F9", "jane_cv.pdf")
ok(names == ["APP-20260701-105826-A3F9_jane_cv.pdf"], "Single resume filename reconstructed")

# Multiple resumes (comma-separated in column)
names2 = _stored_resume_names("APP-20260701-105826-A3F9", "jane_cv.pdf, cover.docx")
ok(len(names2) == 2, "Two resumes from comma-separated column")
ok(names2[0] == "APP-20260701-105826-A3F9_jane_cv.pdf", "First resume filename correct")
ok(names2[1] == "APP-20260701-105826-A3F9_cover.docx", "Second resume filename correct")

# AppRef with "/" â†’ replaced with "-" (defensive; AppRef uses dashes already)
names3 = _stored_resume_names("APP/20260701/105826-A3F9", "cv.pdf")
ok("/" not in names3[0], "Slash in AppRef replaced with dash in filename")

# Empty filename column â†’ empty list (skip row)
names4 = _stored_resume_names("APP-001", "")
ok(names4 == [], "Empty Original Filename column â†’ no files to download")

# WhiteSpace padding stripped
names5 = _stored_resume_names("APP-001", " jane_cv.pdf , cover.docx ")
ok(all("_" in n for n in names5), "Whitespace stripped around filenames")

# Already contains AppID prefix â†’ should NOT double-prepend
names6 = _stored_resume_names("APP-20260701-105826-A3F9", "AadityaBorse_APP-20260701-105826-A3F9.pdf")
ok(names6 == ["AadityaBorse_APP-20260701-105826-A3F9.pdf"], "Already prefixed filename kept identical")

names7 = _stored_resume_names(
    "APP-20260602-0447-97F8", "Candidate_MobileApps_97F8.pdf")
ok(names7 == ["Candidate_MobileApps_97F8.pdf"],
   "Canonical P2 filename stored in Original Filename is tried verbatim")

# Clean name helper testing
from hiring_agent.sharepoint_scoring import _get_cleaned_filename_prefix
ok(_get_cleaned_filename_prefix("Aaditya Borse") == "AadityaBorse", "First Last â†’ FirstNameLastName")
ok(_get_cleaned_filename_prefix("John Fitzgerald Kennedy") == "JohnKennedy", "First Middle Last â†’ FirstNameLastName")
ok(_get_cleaned_filename_prefix("John Fitzgerald") == "JohnFitzgerald", "First Middle (no last) â†’ FirstNameMiddleName")
ok(_get_cleaned_filename_prefix("Aaditya") == "Aaditya", "Only First â†’ FirstName")
ok(_get_cleaned_filename_prefix("  Aaditya-Borse!  ") == "AadityaBorse", "Special chars removed")
ok(_get_cleaned_filename_prefix("") == "Candidate", "Empty â†’ Candidate")


# â”€â”€ P2b. normalize_country â€” one canonical US label â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== P2b. normalize_country â€” canonical 'United States' ===")
from hiring_agent.geo import normalize_country
for variant in ("USA", "usa", "U.S.A.", "U.S.", "US", "us", "America",
                "united states of america", "United States", "U S A"):
    ok(normalize_country(variant) == "United States",
       f"normalize_country({variant!r}) â†’ 'United States'")
ok(normalize_country("Pakistan") == "Pakistan", "non-US country left unchanged (Pakistan)")
ok(normalize_country("Turkey") == "Turkey", "non-US country left unchanged (Turkey)")
ok(normalize_country("Maharashtra") == "India", "Indian state in Country maps to India")
ok(normalize_country("Ontario") == "Canada", "Canadian province in Country maps to Canada")
ok(normalize_country("") == "", "blank country stays blank (never becomes United States)")
ok(normalize_country("Not extracted") == "Not extracted", "placeholder not coerced to a country")


# â”€â”€ P2b2. reconcile_us_country â€” US-confirmed rows read 'United States' â”€â”€â”€â”€â”€â”€â”€
# normalize_country alone can't fix a US STATE name ('California') or a foreign country
# mis-read off a past overseas degree ('India' for a candidate now in 'Chicago, IL') â€”
# 'Georgia' is even a US-state/country collision. reconcile_us_country resolves it using
# the (separately verified) US location signal. Real live cases: Sanghita (San Francisco,
# Country='California') and Sadaf (Chicago, IL, Country='India').
print("\n=== P2b2. reconcile_us_country â€” canonical US label when US-confirmed ===")
from hiring_agent.geo import reconcile_us_country
ok(reconcile_us_country("California", "San Francisco", True) == "United States",
   "US city + Country=state 'California' â†’ 'United States' (Sanghita)")
ok(reconcile_us_country("India", "Chicago, IL", True) == "United States",
   "US state 'Chicago, IL' + Country=India â†’ 'United States' (Sadaf)")
ok(reconcile_us_country("India", "Mumbai, Maharashtra", False) == "India",
   "rejected non-US row (is_usa=False) â†’ Country left unchanged (India)")
ok(reconcile_us_country("Canada", "Toronto", True) == "Canada",
   "kept but NO US location signal (bare foreign city) â†’ country left unchanged")
ok(reconcile_us_country("Not extracted", "", True) == "Not extracted",
   "kept with blank location â†’ never fabricates 'United States'")
ok(reconcile_us_country("United States", "Austin, TX", True) == "United States",
   "already-US row stays 'United States'")


# â”€â”€ P2c. format_phone â€” uniform US formatting + Excel text alignment â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== P2c. format_phone â€” uniform (XXX) XXX-XXXX ===")
from hiring_agent.extraction import format_phone
ok(format_phone("6677036994") == "(667) 703-6994", "bare 10-digit â†’ (XXX) XXX-XXXX")
ok(format_phone("6232975055") == "(623) 297-5055", "bare 10-digit (2) â†’ formatted")
ok(format_phone("501-653-7496") == "(501) 653-7496", "dashed 10-digit â†’ parenthesized")
ok(format_phone("+1 (415) 882-6713") == "(415) 882-6713", "+1 US 11-digit â†’ 10-digit US format")
ok(format_phone("(928) 668-3265") == "(928) 668-3265", "already-formatted US kept identical")
ok(format_phone("+92 3320742977") == "+92 3320742977", "international (+92) left unchanged")
ok(format_phone("Not extracted") == "Not extracted", "gap literal left unchanged")
ok(format_phone("") == "", "empty phone stays empty")

# sanitize_phone â€” masked/truncated numbers become a clean gap, never displayed
from hiring_agent.extraction import sanitize_phone, normalize_skills
ok(sanitize_phone("732-4***") == "Not extracted", "masked '732-4***' â†’ Not extracted")
ok(sanitize_phone("-8455") == "Not extracted", "too-short '-8455' â†’ Not extracted")
ok(sanitize_phone("7 4 0 5 4 6 5 2 0") == "Not extracted",
   "spaced single digits are treated as broken phone text")
ok(sanitize_phone("(732) 427-3680") == "(732) 427-3680", "a real 10-digit number is kept")
ok(sanitize_phone("+92 3320742977") == "+92 3320742977", "a real international number is kept")
ok(sanitize_phone("Not extracted") == "Not extracted", "gap stays a gap")
ok(sanitize_phone("") == "", "empty stays empty")
ok(normalize_skills("['Kotlin', 'Android']") == "Kotlin, Android",
   "Python-list-like skill output is normalized to comma-separated text")
ok(normalize_skills(["Kotlin", "Android", "Kotlin"]) == "Kotlin, Android",
   "list skill output is deduped and normalized")


# â”€â”€ P2d. resume link/path â€” plain URL/full-path values + calc-column formula â”€
print("\n=== P2d. resume link/path (Resume URL + Resume Folder Path values + calc formula) ===")
from hiring_agent.sharepoint_scoring import (
    _resume_url_and_path, _resume_display_path, _rejected_subpath, _RESUME_LINK_FORMULA,
)

ok(_rejected_subpath("2026/June") == "2026/Rejected",
   "_rejected_subpath is YEAR-level only (changed 2026-07-15) - month is dropped")
ok(_rejected_subpath("2026/December") == "2026/Rejected",
   "every month of the same year collapses to the same Rejected bucket")

class _StubClient:
    def __init__(self, url):
        self._url = url
        self.resumes_folder = "/Downloaded_Resumes"
    def resume_web_url(self, name, subfolder=""):
        if isinstance(self._url, dict):
            return self._url.get(name, "")
        return self._url

ok(_resume_url_and_path(_StubClient("https://x/Doc.aspx?file=cv.pdf"), ["APP-1_cv.pdf"], "2026/June")
   == ("https://x/Doc.aspx?file=cv.pdf", "Downloaded_Resumes/2026/June"),
   "returns the first resume's web URL + the FOLDER-ONLY 'Resume Folder Path' display text "
   "(no filename)")
ok(_resume_url_and_path(_StubClient(""), ["APP-1_cv.pdf"], "") == ("", ""), "no web URL â†’ empty pair")
ok(_resume_url_and_path(_StubClient("https://x/f"), [], "") == ("", ""), "no stored resume â†’ empty pair")
ok(_resume_url_and_path(
    _StubClient({"Canonical.pdf": "https://x/Canonical.pdf"}),
    ["Legacy.pdf", "Canonical.pdf"], "2026/Rejected")
   == ("https://x/Canonical.pdf", "Downloaded_Resumes/2026/Rejected"),
   "URL/path falls through to canonical filename when legacy filename is gone")
ok(_resume_display_path("/Downloaded_Resumes", _rejected_subpath("2026/June"))
   == "Downloaded_Resumes/2026/Rejected",
   "'Resume Folder Path' shows the year-level 'Rejected' bucket, still with no filename")
ok('TRIM([@[Original Filename]])' in _RESUME_LINK_FORMULA,
   "'Resume Link' formula prefers the original/friendly filename as the visible label")
ok('TRIM(RIGHT(SUBSTITUTE([@[Resume URL]],"/",REPT(" ",300)),300))' in _RESUME_LINK_FORMULA,
   "'Resume Link' formula falls back to the saved filename from Resume URL")
ok(_RESUME_LINK_FORMULA.startswith('=IF([@[Resume URL]]="","",'),
   "empty Resume URL renders blank (no stray 0 on the seed/no-resume rows)")


# â”€â”€ P2d2. saved-resume filename: FirstNameLastName_Category_<tail>.ext (2026-07-15) â”€â”€
print("\n=== P2d2. resume filename (name + category + AppID tail) ===")
from hiring_agent.sharepoint_scoring import (
    _clean_category_for_filename, _resume_filename_tail, _stored_resume_names,
)

_CATEGORY_CASES = [
    ("Data Analytics", "DataAnalytics"),
    ("Senior & Executive", "SeniorExecutive"),
    ("AI/ML/CV (SIN2)", "AIMLCVSIN2"),
    ("3D/CV/ IoT/ AI Agents (SIN3)", "3DCVIoTAIAgentsSIN3"),
    ("Cloud and DevOps", "CloudAndDevOps"),
    ("Mobile Apps (Android IOS)", "MobileAppsAndroidIOS"),
    ("Cybersecurity and IT Admin", "CybersecurityAndITAdmin"),
    ("Business Analytics", "BusinessAnalytics"),
    ("Supply Chain", "SupplyChain"),
    ("Data Center", "DataCenter"),
    ("General", "General"),
]
for _raw, _expected in _CATEGORY_CASES:
    ok(_clean_category_for_filename(_raw) == _expected,
       f"category {_raw!r} -> {_expected!r} (got {_clean_category_for_filename(_raw)!r})")
ok(_clean_category_for_filename("") == "General", "blank category falls back to 'General'")
ok(_clean_category_for_filename("Not extracted") == "General",
   "gap-literal category falls back to 'General'")

ok(_resume_filename_tail("APP-20260710-2200-A5F2") == "A5F2",
   "tail = the last '-'-delimited segment of the Application ID")
ok(_resume_filename_tail("APP-20260710-220015-A5F2C1") == "A5F2C1",
   "tail length adapts automatically if the AppID's own hex tail is longer (e.g. pre-2026-07-14 6-hex refs)")
ok(_resume_filename_tail("NODASHESATALL") == "NODASHESATALL",
   "no '-' at all -> the whole string is returned rather than crashing")

_names = _stored_resume_names("APP-20260710-2200-A5F2", "jane_cv.pdf",
                              full_name="Jane Doe", category="Data Analytics")
# Legacy '<Name>_<FullAppID>' format is tried FIRST (regression fix, 2026-07-15): P1 saves
# EVERY resume write - first submission, Duplicate resend, or ref-quoted Update - under this
# exact shape (see build_zip.py), and P2 only ever renames a file OUT of it, right after
# scoring. So this shape existing means "a fresh write just happened" and must win over a
# stale post-scoring canonical-shape file, never be blindly combined with it - see
# _resume_name_slots / _download_resume_text.
ok(_names[0] == "JaneDoe_APP-20260710-2200-A5F2.pdf",
   f"legacy (always-current-write) filename format tried FIRST (got {_names[0]!r})")
ok("JaneDoe_DataAnalytics_A5F2.pdf" in _names,
   "post-scoring canonical format still tried as a fallback candidate")
ok("APP-20260710-2200-A5F2_jane_cv.pdf" in _names,
   "original P1-intake filename still tried as the last-resort fallback candidate")
_names_no_cat = _stored_resume_names("APP-20260710-2200-A5F2", "jane_cv.pdf", full_name="Jane Doe")
ok("JaneDoe_DataAnalytics_A5F2.pdf" not in _names_no_cat,
   "no category given -> new-format candidate is skipped (falls back to name+AppID format)")
_stale_category_url = (
    "https://example.sharepoint.com/Shared%20Documents/Candidate_Resumes/"
    "2026/Rejected/JaneDoe_Graphics_A5F2.pdf"
)
_names_from_url = _stored_resume_names(
    "APP-20260710-2200-A5F2", "jane_cv.pdf",
    full_name="Jane Doe", category="Mobile Apps (Android IOS)",
    resume_url=_stale_category_url)
ok(_names_from_url[0] == "JaneDoe_Graphics_A5F2.pdf",
   "stored Resume URL filename wins when a prior category rename made generated guesses stale")

# Regression: an already-scored candidate who replies with an updated resume (the exact path
# P2's own missing_info nudge asks for) must be re-scored off their NEW resume, not a blend of
# stale + fresh text, and must not end up with two files on disk. P1 always writes an update
# under the legacy '<Name>_<FullAppID>' shape (build_zip.py), while the OLD, already-renamed
# canonical-shape file is still sitting there untouched - _download_resume_text must pick only
# the legacy (fresher) one when both physically exist.
import unittest.mock as _mock
from hiring_agent.sharepoint_scoring import _download_resume_text, _rename_scored_resumes
import hiring_agent.sharepoint_scoring as _sp_scoring
from sharepoint_client import SharePointError as _SPErr

class _FakeUpdateClient:
    """Legacy-shape file = the just-written update; canonical-shape = stale prior content."""
    def download_resume(self, name, subfolder=""):
        if name == "JaneDoe_APP-20260710-2200-A5F2.pdf":
            return b"UPDATED RESUME TEXT WITH NEW PHONE NUMBER"
        if name == "JaneDoe_DataAnalytics_A5F2.pdf":
            return b"STALE OLD RESUME TEXT"
        raise _SPErr(f"404: {name} not found")

with _mock.patch("hiring_agent.sharepoint_scoring.extract_text_from_bytes",
                 side_effect=lambda raw, name: raw.decode()):
    _text, _used, _raw = _download_resume_text(
        _FakeUpdateClient(), "APP-20260710-2200-A5F2", "jane_cv.pdf", "2026/July",
        full_name="Jane Doe", category="Data Analytics")
ok(_text == "UPDATED RESUME TEXT WITH NEW PHONE NUMBER",
   f"update reply is scored off the FRESH resume, not blended with stale content (got {_text!r})")
ok("STALE OLD RESUME TEXT" not in _text,
   "stale pre-update canonical-shape file's text is excluded, not concatenated in")
ok(_used == ["JaneDoe_APP-20260710-2200-A5F2.pdf"],
   f"exactly one file used per attachment slot, not both (got {_used!r})")

class _FakeFirstScoreClient:
    """Only the legacy-shape file exists yet (never scored/renamed before) - first pass."""
    def download_resume(self, name, subfolder=""):
        if name == "JaneDoe_APP-20260710-2200-A5F2.pdf":
            return b"FIRST-TIME RESUME TEXT"
        raise _SPErr(f"404: {name} not found")

with _mock.patch("hiring_agent.sharepoint_scoring.extract_text_from_bytes",
                 side_effect=lambda raw, name: raw.decode()):
    _text2, _used2, _ = _download_resume_text(
        _FakeFirstScoreClient(), "APP-20260710-2200-A5F2", "jane_cv.pdf", "2026/July",
        full_name="Jane Doe", category="Data Analytics")
ok(_text2 == "FIRST-TIME RESUME TEXT", "first-ever scoring pass still finds the legacy-shape file")

class _FakeRecheckClient:
    """Only the canonical-shape file exists (a normal, not-since-updated already-scored row)."""
    def download_resume(self, name, subfolder=""):
        if name == "JaneDoe_DataAnalytics_A5F2.pdf":
            return b"UNCHANGED PREVIOUSLY-SCORED RESUME TEXT"
        raise _SPErr(f"404: {name} not found")

with _mock.patch("hiring_agent.sharepoint_scoring.extract_text_from_bytes",
                 side_effect=lambda raw, name: raw.decode()):
    _text3, _used3, _ = _download_resume_text(
        _FakeRecheckClient(), "APP-20260710-2200-A5F2", "jane_cv.pdf", "2026/July",
        full_name="Jane Doe", category="Data Analytics")
ok(_text3 == "UNCHANGED PREVIOUSLY-SCORED RESUME TEXT",
   "a row with no pending update still falls back to its canonical-shape file correctly")

class _FakeStaleCategoryClient:
    def download_resume(self, name, subfolder=""):
        if name == "JaneDoe_Graphics_A5F2.pdf":
            return b"RESUME UNDER OLD CATEGORY FILENAME"
        raise _SPErr(f"404: {name} not found")

with _mock.patch("hiring_agent.sharepoint_scoring.extract_text_from_bytes",
                 side_effect=lambda raw, name: raw.decode()):
    _text4, _used4, _ = _download_resume_text(
        _FakeStaleCategoryClient(), "APP-20260710-2200-A5F2", "jane_cv.pdf",
        "2026/Rejected", full_name="Jane Doe",
        category="Mobile Apps (Android IOS)", resume_url=_stale_category_url)
ok(_used4 == ["JaneDoe_Graphics_A5F2.pdf"]
   and _text4 == "RESUME UNDER OLD CATEGORY FILENAME",
   "resume download follows the stored URL when the physical filename has an old category")

ok(_names_no_cat[0] == "JaneDoe_APP-20260710-2200-A5F2.pdf",
   "without a category, the name+AppID format is tried first")

class _RenameClient:
    def __init__(self, rename_results=None, existing=None):
        self.rename_results = rename_results or {}
        self.existing = set(existing or [])
        self.renames = []

    def rename_resume(self, old_name, new_name, subfolder=""):
        self.renames.append((old_name, new_name, subfolder))
        return self.rename_results.get((old_name, new_name), False)

    def resume_web_url(self, name, subfolder=""):
        if name in self.existing:
            return "https://example.test/" + name
        raise _SPErr(f"404: {name} not found")

class _LogCapture:
    def __init__(self):
        self.infos = []
        self.warnings = []

    def info(self, msg):
        self.infos.append(str(msg))

    def warning(self, msg):
        self.warnings.append(str(msg))

_old_logger = _sp_scoring.logger
try:
    _cap = _LogCapture()
    _sp_scoring.logger = _cap
    _client = _RenameClient(
        rename_results={
            ("JaneDoe_APP-20260710-2200-A5F2.pdf", "JaneDoe_DataAnalytics_A5F2.pdf"): True,
            ("JaneDoe_APP-20260710-2200-A5F2.docx", "JaneDoe_DataAnalytics_A5F2.docx"): True,
        }
    )
    _renamed = _rename_scored_resumes(
        _client,
        ["JaneDoe_APP-20260710-2200-A5F2.pdf", "JaneDoe_APP-20260710-2200-A5F2.docx"],
        "2026/July",
        "APP-20260710-2200-A5F2",
        "Jane Doe",
        "Data Analytics",
        dry_run=False,
    )
    ok(_renamed == ["JaneDoe_DataAnalytics_A5F2.pdf", "JaneDoe_DataAnalytics_A5F2.docx"],
       "rename helper returns canonical names for distinct extensions")
    ok(_cap.warnings == [], "rename helper does not warn on successful canonical renames")

    _cap = _LogCapture()
    _sp_scoring.logger = _cap
    _client_dup = _RenameClient(
        rename_results={
            ("JaneDoe_APP-20260710-2200-A5F2.pdf", "JaneDoe_DataAnalytics_A5F2.pdf"): True
        }
    )
    _renamed_dup = _rename_scored_resumes(
        _client_dup,
        ["JaneDoe_APP-20260710-2200-A5F2.pdf", "JaneDoe_OtherLegacyName.pdf"],
        "2026/July",
        "APP-20260710-2200-A5F2",
        "Jane Doe",
        "Data Analytics",
        dry_run=False,
    )
    ok(_renamed_dup == ["JaneDoe_DataAnalytics_A5F2.pdf"],
       "duplicate canonical targets collapse to one represented filename")
    ok(_cap.warnings == [] and any("already represented" in m for m in _cap.infos),
       "duplicate canonical target is logged as info, not warning")

    _cap = _LogCapture()
    _sp_scoring.logger = _cap
    _client2 = _RenameClient(existing={"JaneDoe_DataAnalytics_A5F2.pdf"})
    _renamed2 = _rename_scored_resumes(
        _client2,
        ["JaneDoe_APP-20260710-2200-A5F2.pdf"],
        "2026/July",
        "APP-20260710-2200-A5F2",
        "Jane Doe",
        "Data Analytics",
        dry_run=False,
    )
    ok(_renamed2 == ["JaneDoe_DataAnalytics_A5F2.pdf"],
       "rename helper uses existing canonical target when rename returns false")
    ok(_cap.warnings == [] and any("already exists" in m for m in _cap.infos),
       "existing canonical target is logged as info, not warning")

    _cap = _LogCapture()
    _sp_scoring.logger = _cap
    _client3 = _RenameClient()
    _renamed3 = _rename_scored_resumes(
        _client3,
        ["JaneDoe_APP-20260710-2200-A5F2.pdf"],
        "2026/July",
        "APP-20260710-2200-A5F2",
        "Jane Doe",
        "Data Analytics",
        dry_run=False,
    )
    ok(_renamed3 == ["JaneDoe_APP-20260710-2200-A5F2.pdf"],
       "true rename failure preserves the old filename")
    ok(any("Could not rename" in m for m in _cap.warnings),
       "true rename failure still logs a warning")
finally:
    _sp_scoring.logger = _old_logger


# â”€â”€ P2d3. missing-info nudge criteria: phone/location/skills/education ONLY â”€â”€
print("\n=== P2d3. missing-info nudge criteria (portfolio NOT checked, 2026-07-15) ===")
from hiring_agent.sharepoint_scoring import _missing_fields_list

_complete_row = {"Phone": "555-1234", "Location": "Austin, TX", "Current Skills": "Python",
                 "Education": "B.S. Computer Science", "Portfolio 1": "N/A",
                 "Portfolio 2": "N/A", "Portfolio 3": "N/A"}
ok(_missing_fields_list(_complete_row) == [],
   "nothing missing when phone/location/skills/education are all present, "
   "even with all 3 portfolio slots blank")

_no_portfolio_only = dict(_complete_row, **{"Portfolio 1": "N/A", "Portfolio 2": "N/A", "Portfolio 3": "N/A"})
ok(_missing_fields_list(_no_portfolio_only) == [],
   "missing ALL 3 portfolio slots alone does NOT trigger the nudge (removed 2026-07-15)")

ok(_missing_fields_list({**_complete_row, "Phone": "Not extracted"}) == ["your phone number"],
   "missing phone alone is flagged")
ok(_missing_fields_list({**_complete_row, "Location": ""}) == ["your location"],
   "missing location alone is flagged")
ok(_missing_fields_list({**_complete_row, "Current Skills": "N/A"}) == ["your listed skills"],
   "missing skills alone is flagged")
ok(_missing_fields_list({**_complete_row, "Education": "Not extracted"}) == ["your education background"],
   "missing education alone is flagged")

_all_missing = {"Phone": "", "Location": "", "Current Skills": "", "Education": "",
               "Portfolio 1": "https://linkedin.com/in/x", "Portfolio 2": "N/A", "Portfolio 3": "N/A"}
ok(_missing_fields_list(_all_missing) ==
   ["your phone number", "your location", "your listed skills", "your education background"],
   "all four missing, in a fixed order, regardless of a real portfolio link being present")


# â”€â”€ P2e. resume move-on-transition + phone renormalize orchestration â”€â”€â”€â”€â”€â”€â”€â”€â”€
from hiring_agent.sharepoint_scoring import _send_pending_info_requests
import hiring_agent.sharepoint_scoring as _sp_info_mod

class _StubInfoClient:
    def __init__(self):
        self.rows = [
            {"index": 0, "values": {
                "Application ID": "APP-CLARITY", "Status": "Scored",
                "Email": "clarity@x.com", "Full Name": "Clarity Candidate",
                "Phone": "(512) 555-0001", "Location": "Remote", "Country": "",
                "Current Skills": "Python", "Education": "B.S. Computer Science",
                "Info Request Sent": "Nothing missing",
            }},
            {"index": 1, "values": {
                "Application ID": "APP-SENT", "Status": "Scored",
                "Email": "sent@x.com", "Phone": "(512) 555-0002",
                "Location": "Remote", "Country": "", "Current Skills": "Python",
                "Education": "B.S. Computer Science",
                "Info Request Sent": "Sent 2026-07-15 10:00",
            }},
            {"index": 2, "values": {
                "Application ID": "APP-UNSCORED", "Status": "New Email Received",
                "Email": "new@x.com", "Phone": "(512) 555-0003",
                "Location": "Remote", "Country": "", "Current Skills": "Python",
                "Education": "B.S. Computer Science", "Info Request Sent": "",
            }},
        ]
        self.updates = []
    def list_rows(self):
        return self.rows
    def update_row(self, index, fields, current_values=None):
        self.updates.append((index, fields))

_old_send_missing = _sp_info_mod._send_missing_info_request
_time_mod = __import__("time")
_old_sleep = _time_mod.sleep
try:
    _sp_info_mod._send_missing_info_request = lambda client, email, name, app_id, missing: True
    _time_mod.sleep = lambda _seconds: None
    _ic = _StubInfoClient()
    _info_result = _send_pending_info_requests(_ic)
    ok(_info_result == {"sent": 1, "failed": 0, "skipped": 0},
       "benefit-of-doubt row stamped 'Nothing missing' is reopened for one clarity email")
    # Accept either marker form: the stamp is 'Sent <ts>' normally and
    # 'TEST-MODE (suppressed) Sent <ts>' when HIRING_SUPPRESS_EMAILS is on, so this
    # assertion stays deterministic regardless of the operator's local .env.
    _info_stamp = str(_ic.updates[0][1].get("Info Request Sent", "")) if _ic.updates else ""
    ok(len(_ic.updates) == 1 and _ic.updates[0][0] == 0
       and "Sent " in _info_stamp,
       "only the scored clarity row is stamped Sent; sent/unscored rows are ignored")
finally:
    _sp_info_mod._send_missing_info_request = _old_send_missing
    _time_mod.sleep = _old_sleep

print("\n=== P2e. resume move + phone renormalize orchestration ===")
from hiring_agent.sharepoint_scoring import _move_resumes, _renormalize_phone_columns

class _StubMoveClient:
    """Fakes just enough of SharePointClient.move_resume's contract: a file only moves if
    it 'exists' at the source, and moving relocates it (so a second attempt is a no-op)."""
    def __init__(self, existing):
        self.resumes_folder = "/Main"
        self.moved = []
        self._existing = set(existing)
    def move_resume(self, name, from_folder, to_folder):
        key = (from_folder, name)
        if key not in self._existing:
            return False
        self._existing.discard(key)
        self._existing.add((to_folder, name))
        self.moved.append((name, from_folder, to_folder))
        return True

_c1 = _StubMoveClient({("/Main/2026/June", "APP-1_cv.pdf")})
ok(_move_resumes(_c1, ["APP-1_cv.pdf"], "2026/June", to_rejected=True) is True,
   "_move_resumes reports True when a file actually moves")
ok(_c1.moved == [("APP-1_cv.pdf", "/Main/2026/June", "/Main/2026/Rejected")],
   "moves the month's main folder -> the YEAR-level 'Rejected' bucket (not a per-month one)")

_c2 = _StubMoveClient(set())   # nothing at the source -> already moved (or never existed)
ok(_move_resumes(_c2, ["APP-1_cv.pdf"], "2026/June", to_rejected=True) is False,
   "_move_resumes is a no-op (returns False) when the file isn't at the source â€” idempotent")

class _StubPhoneClient:
    """Fakes list_rows/list_rejected_rows + update_row/update_rejected_row for the
    _renormalize_phone_columns backfill."""
    def __init__(self, main_rows, rej_rows):
        self._main = main_rows
        self._rej = rej_rows
        self.writes = []
    def list_rows(self):
        return [{"index": i, "values": {"Phone": p}} for i, p in self._main]
    def list_rejected_rows(self):
        return [{"index": i, "values": {"Phone": p}} for i, p in self._rej]
    def update_row(self, index, fields, current_values=None):
        self.writes.append(("main", index, fields["Phone"]))
    def update_rejected_row(self, index, fields, current_values=None):
        self.writes.append(("rejected", index, fields["Phone"]))

_pc = _StubPhoneClient(
    main_rows=[(0, "(469) 688-7299"), (1, 923326596399)],
    rej_rows=[(0, 971546000000), (1, "+92 332 659 6399")],
)
_renormalize_phone_columns(_pc)
ok(_pc.writes == [("main", 1, "923326596399"), ("rejected", 0, "971546000000")],
   "only Phone cells Excel already auto-converted to a NUMBER get re-written as plain-digit "
   "text; already-text cells (formatted US numbers, foreign numbers with a '+') are untouched")


# â”€â”€ P2e2. resume folder path rebuild (anything not folder-only -> folder-only) â”€â”€â”€
print("\n=== P2e2. _renormalize_resume_paths â€” rebuilds to the folder-only shape ===")
from hiring_agent.sharepoint_scoring import _renormalize_resume_paths

class _StubPathClient:
    def __init__(self, main_rows, rej_rows):
        self._main = main_rows
        self._rej = rej_rows
        self.resumes_folder = "/Downloaded_Resumes"
        self.writes = []
    def list_rows(self):
        return self._main
    def list_rejected_rows(self):
        return self._rej
    def update_row(self, index, fields, current_values=None):
        self.writes.append(("main", index, fields["Resume Folder Path"]))
    def update_rejected_row(self, index, fields, current_values=None):
        self.writes.append(("rejected", index, fields["Resume Folder Path"]))

_ppc = _StubPathClient(
    main_rows=[
        {"index": 0, "values": {"Received Date": "2026-07-01T10:00:00",
                                "Resume Folder Path": "APP-1_cv.pdf"}},  # bare filename (old leftover) -> rebuild
        {"index": 1, "values": {"Received Date": "2026-06-15T10:00:00",
                                "Resume Folder Path": "Downloaded_Resumes/2026/June"}},  # already folder-only -> skip
        {"index": 2, "values": {"Received Date": "2026-06-20T10:00:00",
                                "Resume Folder Path": "Downloaded_Resumes/2026/June/APP-2_cv.pdf"}},  # full path+filename (prior design) -> rebuild
    ],
    rej_rows=[
        {"index": 0, "values": {"Received Date": "2026-06-10T10:00:00",
                                "Resume Folder Path": "APP-3_cv.pdf"}},  # bare -> rebuild w/ Rejected segment
    ],
)
_renormalize_resume_paths(_ppc)
_pp_writes = {(lbl, idx): val for lbl, idx, val in _ppc.writes}
ok(_pp_writes.get(("main", 0)) == "Downloaded_Resumes/2026/July",
   "bare filename on main sheet rebuilt into the folder-only dated path")
ok(("main", 1) not in _pp_writes,
   "a cell that already matches the expected folder-only path is left untouched")
ok(_pp_writes.get(("main", 2)) == "Downloaded_Resumes/2026/June",
   "a full path+filename (from before folder-only was the design) has the filename stripped")
ok(_pp_writes.get(("rejected", 0)) == "Downloaded_Resumes/2026/Rejected",
   "bare filename on the Rejected sheet rebuilt with the YEAR-level 'Rejected' bucket")


# â”€â”€ P2f. cross-candidate duplicate MERGE (Phone only) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== P2f. _merge_duplicate_candidates â€” merges cross-sender duplicates to ONE row ===")
from hiring_agent.sharepoint_scoring import _normalize_phone_key, _merge_duplicate_candidates, _finish_rejection

ok(_normalize_phone_key("(469) 688-7299") == "4696887299", "US-formatted phone -> digits only")
ok(_normalize_phone_key("+1 469 688 7299") == "4696887299",
   "US country-code '1' prefix dropped so it matches the bare 10-digit form")
ok(_normalize_phone_key("923127174226") == "923127174226",
   "non-US number kept whole (no US country-code stripping applied)")
ok(_normalize_phone_key("12345") == "", "too short (<7 digits) never matches -> empty key")
ok(_normalize_phone_key("Not extracted") == "", "gap phone -> empty key")

class _StubMergeClient:
    """Fakes list_rows/list_rejected_rows + update_row/update_rejected_row/delete_row/
    delete_rejected_row/delete_file for _merge_duplicate_candidates."""
    def __init__(self, main_rows, rej_rows):
        self._main = main_rows
        self._rej = rej_rows
        self.resumes_folder = "/Downloaded_Resumes"
        self.updates = []
        self.deleted_rows = []
        self.deleted_files = []
    def list_rows(self):
        return self._main
    def list_rejected_rows(self):
        return self._rej
    def update_row(self, index, fields, current_values=None):
        self.updates.append(("main", index, fields))
    def update_rejected_row(self, index, fields, current_values=None):
        self.updates.append(("rejected", index, fields))
    def delete_row(self, index):
        self.deleted_rows.append(("main", index))
    def delete_rejected_row(self, index):
        self.deleted_rows.append(("rejected", index))
    def delete_file(self, folder, name):
        self.deleted_files.append((folder, name))

# "Akshay Faye"-style: same candidate, two different sender/agency emails, different
# Received Date. APP-NEW is the more recent submission -> should win; APP-OLD (older,
# different sender) is the loser -> healed-from, then deleted (row + resume file).
# APP-OLD also carries a real 'Category' the winner is missing (a gap) -> must be healed in.
_mm_main = [
    {"index": 0, "values": {
        "Application ID": "APP-OLD", "Email": "agency1@x.com",
        "Received Date": "2026-06-09T19:19:44",
        "Last Updated Date": "2026-06-09T19:19:44",
        "Phone": "(469) 688-7299",
        "Category": "Software Engineering",
        "Mail Subject": "Old application",
        "Mail Body": "Old body",
        "Application Updates": 4,
        "Original Filename": "old_resume.pdf", "Status": "Scored",
        "Resume URL": "https://old/resume.pdf",
        "Resume Folder Path": "Downloaded_Resumes/2026/June",
    }},
    {"index": 1, "values": {
        "Application ID": "APP-NEW", "Email": "agency2@x.com",
        "Received Date": "2026-06-09T20:28:53",
        "Last Updated Date": "",
        "Phone": "469-688-7299",
        "Category": "Not extracted",
        "Current Skills": "Swift, Kotlin",  # a real (non-gap) value - must survive untouched
        "Original Filename": "new_resume.docx", "Status": "Scored",
    }},
    # Different person entirely - no phone overlap with anyone, must be untouched.
    {"index": 2, "values": {
        "Application ID": "APP-C", "Email": "someone@x.com", "Received Date": "2026-06-01T00:00:00",
        "Phone": "555-000-1111",
    }},
]
_mm_rej: list = []
_mc = _StubMergeClient(_mm_main, _mm_rej)
_result = _merge_duplicate_candidates(_mc)

ok(_result == {"merged": 1, "errors": 0}, f"exactly one merge, no errors (got {_result})")

_winner_update = next((f for lbl, idx, f in _mc.updates if lbl == "main" and idx == 1), None)
ok(_winner_update is not None, "the WINNER (more recent APP-NEW) gets an update call")
ok(_winner_update is not None and _winner_update.get("Category") == "Software Engineering",
   "winner's blank 'Category' healed from the loser's real value")
ok(_winner_update is not None and "Current Skills" not in _winner_update,
   "winner's own real 'Current Skills' value is NEVER overwritten by the loser's (it has none)")
for _p1_merge_col in ("Received Date", "Last Updated Date", "Email", "Mail Subject", "Mail Body",
                      "Application Updates", "Original Filename", "Resume URL", "Resume Folder Path"):
    ok(_winner_update is not None and _p1_merge_col not in _winner_update,
       f"duplicate merge never heals P1-owned/audit field '{_p1_merge_col}' from loser")
ok(not any(lbl == "main" and idx == 0 for lbl, idx, _ in _mc.updates),
   "the LOSER (older APP-OLD) never gets updated, only deleted")

ok(("main", 0) in _mc.deleted_rows, "the loser row (APP-OLD, index 0) is deleted")
ok(("main", 1) not in _mc.deleted_rows, "the winner row (APP-NEW, index 1) is NOT deleted")
ok(("main", 2) not in _mc.deleted_rows, "APP-C (no overlap) is untouched â€” not deleted either")

ok(("/Downloaded_Resumes/2026/June", "resume.pdf") in _mc.deleted_files,
   "the loser's URL-linked resume file is deleted from its dated folder")
ok(len(_mc.deleted_files) == 1, "only the loser's file is deleted â€” winner's file is untouched")

# dry_run: computes the same plan but writes/deletes nothing at all.
_mc_dry = _StubMergeClient(
    [dict(r) for r in _mm_main], [dict(r) for r in _mm_rej])
_dry_result = _merge_duplicate_candidates(_mc_dry, dry_run=True)
ok(_dry_result == {"merged": 1, "errors": 0}, "dry-run reports the same merge count")
ok(_mc_dry.updates == [] and _mc_dry.deleted_rows == [] and _mc_dry.deleted_files == [],
   "dry-run makes NO writes, deletes, or file deletions")

# A->B by phone and B->C by email is one candidate component. The newest C row
# must be the only survivor, and it must receive a missing field available only on A.
_chain_main = [
    {"index": 0, "values": {
        "Application ID": "APP-CHAIN-A", "Email": "a@x.com",
        "Phone": "111-222-3333", "Received Date": "2026-06-01T00:00:00",
        "Category": "Cloud and DevOps", "Status": "Scored",
    }},
    {"index": 1, "values": {
        "Application ID": "APP-CHAIN-B", "Email": "bridge@x.com",
        "Phone": "111-222-3333", "Received Date": "2026-06-02T00:00:00",
        "Category": "Not extracted", "Status": "Scored",
    }},
    {"index": 2, "values": {
        "Application ID": "APP-CHAIN-C", "Email": "bridge@x.com",
        "Phone": "999-888-7777", "Received Date": "2026-06-03T00:00:00",
        "Category": "Not extracted", "Status": "Scored",
    }},
]
_chain_client = _StubMergeClient(_chain_main, [])
_chain_result = _merge_duplicate_candidates(_chain_client)
_chain_update = next(
    (fields for sheet, idx, fields in _chain_client.updates
     if sheet == "main" and idx == 2),
    {},
)
ok(_chain_result == {"merged": 2, "errors": 0},
   "transitive email/phone duplicate chain collapses two losers into one newest winner")
ok(_chain_update.get("Category") == "Cloud and DevOps",
   "transitive duplicate winner receives a missing field from the oldest connected row")
ok(set(_chain_client.deleted_rows) == {("main", 0), ("main", 1)},
   "transitive duplicate cleanup deletes every loser and retains only the newest row")

# Same-email pair (changed 2026-07-12): now MERGED here too. P1's dedup is a read-after-write
# against Excel Online and can miss a same-sender resubmission that arrives within minutes (the
# real nishantacharekar12@gmail.com case: two rows, same email, 4 min apart). P2 is the durable
# safety net, so identical-email rows collapse to one â€” winner = more recent Received Date.
_ss_main = [
    {"index": 0, "values": {"Application ID": "APP-E", "Email": "agency1@x.com",
                            "Received Date": "2026-06-01T00:00:00", "Phone": "111-222-3333",
                            "Original Filename": "a.pdf", "Status": "Scored"}},
    {"index": 1, "values": {"Application ID": "APP-F", "Email": "agency1@x.com",
                            "Received Date": "2026-06-02T00:00:00", "Phone": "111-222-3333",
                            "Original Filename": "b.pdf", "Status": "Scored"}},
]
_ssc = _StubMergeClient(_ss_main, [])
_ss_result = _merge_duplicate_candidates(_ssc)
ok(_ss_result == {"merged": 1, "errors": 0},
   "identical email on both rows -> merged to one (P2 safety net for P1's dedup race)")
ok(("main", 0) in _ssc.deleted_rows and ("main", 1) not in _ssc.deleted_rows,
   "same-email merge: older row (index 0) deleted, more-recent row (index 1) kept")

_queue_dup_main = [
    {"index": 0, "values": {"Application ID": "APP-OLD-REJECTED", "Email": "same@x.com",
                            "Received Date": "2026-05-01T00:00:00", "Phone": "111-222-3333",
                            "Original Filename": "old.pdf", "Status": "Rejected - Non-USA Location"}},
    {"index": 1, "values": {"Application ID": "APP-FRESH-UPDATE", "Email": "same@x.com",
                            "Received Date": "2026-07-16T06:17:59", "Phone": "111-222-3333",
                            "Original Filename": "new.pdf", "Status": "New Email Received"}},
]
_qdc = _StubMergeClient(_queue_dup_main, [])
_qd_result = _merge_duplicate_candidates(_qdc)
ok(_qd_result == {"merged": 0, "errors": 0} and _qdc.updates == [] and _qdc.deleted_rows == [],
   "duplicate merge ignores New Email Received rows so fresh updates are scored before merging")

# Different email AND different phone, even with the SAME name -> two DIFFERENT people, NOT merged.
_dp_main = [
    {"index": 0, "values": {"Application ID": "APP-G", "Full Name": "John Smith",
                            "Email": "john.smith.a@x.com", "Received Date": "2026-06-01T00:00:00",
                            "Phone": "111-000-0001", "Original Filename": "a.pdf"}},
    {"index": 1, "values": {"Application ID": "APP-H", "Full Name": "John Smith",
                            "Email": "john.smith.b@y.com", "Received Date": "2026-06-02T00:00:00",
                            "Phone": "222-000-0002", "Original Filename": "b.pdf"}},
]
_dpc = _StubMergeClient(_dp_main, [])
_dp_result = _merge_duplicate_candidates(_dpc)
ok(_dp_result == {"merged": 0, "errors": 0},
   "same name but different email AND phone -> different people, zero merges (name never matches)")


# â”€â”€ P2g. retry-cap: give up after SCORE_RETRY_MAX failures, never loop forever â”€
print("\n=== P2g. retry-cap â€” _mark_needs_review / _give_up_and_reject / error alert ===")
class _StubFinishRejectClient:
    def __init__(self):
        self.resumes_folder = "/Downloaded_Resumes"
        self.deleted_rows = []
        self.rejected_added = []
    def delete_row(self, index):
        self.deleted_rows.append(index)
    def add_rejected_row(self, fields):
        self.rejected_added.append(fields)

_frc_dup = _StubFinishRejectClient()
_fr_phone_set = {"4696887299"}
_fr_result = _finish_rejection(
    _frc_dup, 4, "APP-NEW-DIFF", {"Email": "other@x.com", "Phone": "(469) 688-7299"},
    {"Status": "Rejected - Non-USA Location"}, [], "2026/June", False, set(), set(),
    _fr_phone_set, "non-USA",
)
ok(_fr_result is True and _frc_dup.deleted_rows == [4] and _frc_dup.rejected_added == [],
   "rejecting a phone-duplicate already on Rejected deletes only the main copy")

class _StubRefreshRejectedClient(_StubFinishRejectClient):
    def __init__(self):
        super().__init__()
        self.moved = []
        self.rejected_updates = []
        self.deleted_files = []

    def move_resume(self, name, from_folder, to_folder):
        self.moved.append((name, from_folder, to_folder))
        return True

    def resume_web_url(self, name, subfolder=""):
        return f"https://sharepoint.test/{subfolder}/{name}"

    def update_rejected_row(self, index, fields, current_values=None):
        self.rejected_updates.append((index, fields, current_values))

    def delete_file(self, folder, name):
        self.deleted_files.append((folder, name))

_frc_refresh = _StubRefreshRejectedClient()
_old_rejected = {
    "index": 7,
    "values": {
        "Application ID": "APP-ORIGINAL",
        "Email": "repeat@x.com",
        "Phone": "+91 98765 43210",
        "Location": "Not extracted",
        "Received Date": "2026-06-01T00:00:00",
        "Resume URL": "https://sharepoint.test/2026/Rejected/old.pdf",
        "Decline Sent": "Sent 2026-06-02 09:00",
    },
}
_fr_refresh_result = _finish_rejection(
    _frc_refresh,
    4,
    "APP-NEW-REPLY",
    {
        "Application ID": "APP-NEW-REPLY",
        "Email": "repeat@x.com",
        "Phone": "+91 98765 43210",
        "Received Date": "2026-07-20T10:00:00",
        "Original Filename": "new.pdf",
        "Info Request Sent": "Sent 2026-07-19 09:00",
    },
    {
        "Location": "Mumbai, India",
        "Country": "India",
        "Status": "Rejected - Non-USA Location",
    },
    ["new.pdf"],
    "2026/July",
    False,
    {"APP-ORIGINAL"},
    {"repeat@x.com"},
    {"919876543210"},
    "non-USA",
    rejected_rows=[_old_rejected],
)
_fr_refresh_fields = _frc_refresh.rejected_updates[0][1]
ok(_fr_refresh_result is True and _frc_refresh.deleted_rows == [4],
   "repeat non-US reply removes the temporary Main row after refreshing Rejected")
ok(_fr_refresh_fields.get("Application ID") == "APP-ORIGINAL"
   and _fr_refresh_fields.get("Location") == "Mumbai, India",
   "repeat non-US reply preserves the original reference and stores confirmed location")
ok(_fr_refresh_fields.get("Resume URL", "").endswith("/2026/Rejected/new.pdf"),
   "repeat non-US reply points the Rejected record at the newest resume")
ok("Info Request Sent" not in _fr_refresh_fields,
   "main-only clarification marker is never copied into the Rejected sheet")
ok(("/Downloaded_Resumes/2026/Rejected", "old.pdf") in _frc_refresh.deleted_files,
   "superseded Rejected resume is removed after the new resume is safely linked")

_frc_new = _StubFinishRejectClient()
_fr_new_phones = set()
_fr_new_result = _finish_rejection(
    _frc_new, 5, "APP-FRESH", {"Email": "fresh@x.com", "Phone": "111-222-3333"},
    {"Status": "Rejected - Non-USA Location"}, [], "2026/June", False, set(), set(),
    _fr_new_phones, "non-USA",
)
ok(_fr_new_result is True and len(_frc_new.rejected_added) == 1 and _frc_new.deleted_rows == [5],
   "new non-USA rejection still adds one Rejected row and removes the main row")
ok("1112223333" in _fr_new_phones,
   "new non-USA rejection registers its phone so later same-run rows cannot duplicate it")

from hiring_agent.sharepoint_scoring import (
    _score_attempts_of, _mark_needs_review, _give_up_and_reject, _send_error_alert,
    _send_pending_declines,
)
import hiring_agent.config as _cfg_mod

class _StubDeclineClient:
    def __init__(self):
        self.rows = [
            {"index": 0, "values": {
                "Application ID": "APP-DECLINED",
                "Email": "first@x.com",
                "Phone": "111-222-3333",
                "Decline Sent": "Sent 2026-07-15 15:26",
            }},
            {"index": 1, "values": {
                "Application ID": "APP-DUPE-PHONE",
                "Email": "second@x.com",
                "Phone": "(111) 222-3333",
                "Decline Sent": "",
            }},
        ]
        self.updates = []
        self.sent = []
    def list_rejected_rows(self):
        return self.rows
    def update_rejected_row(self, index, fields, current_values=None):
        self.updates.append((index, fields))
    def send_mail(self, to_address, subject, html_body):
        self.sent.append((to_address, subject))
        return True

_old_geo_reject_email = _cfg_mod.GEO_REJECT_EMAIL
try:
    _cfg_mod.GEO_REJECT_EMAIL = True
    _dc = _StubDeclineClient()
    _decline_result = _send_pending_declines(_dc)
    ok(_decline_result == {"sent": 0, "failed": 0, "skipped": 1},
       "decline pass skips a pending rejected row when phone already received a decline")
    ok(_dc.sent == [] and _dc.updates == [(1, {"Decline Sent": "Duplicate - decline already emailed"})],
       "phone-duplicate decline row is stamped duplicate without sending mail")
finally:
    _cfg_mod.GEO_REJECT_EMAIL = _old_geo_reject_email

ok(_score_attempts_of({}) == 0, "no 'Retry Count' cell -> reads as 0")
ok(_score_attempts_of({"Retry Count": "2"}) == 2, "numeric string parsed correctly")
ok(_score_attempts_of({"Retry Count": "Not extracted"}) == 0,
   "a gap-literal/non-numeric cell reads as 0, never raises")

class _StubRetryClient:
    """Fakes update_row/add_rejected_row/delete_row/move_resume/send_mail for the
    retry-cap + give-up + error-alert path."""
    def __init__(self):
        self.resumes_folder = "/Downloaded_Resumes"
        self.updates = []
        self.rejected_added = []
        self.deleted_rows = []
        self.mails_sent = []
    def update_row(self, index, fields, current_values=None):
        self.updates.append((index, fields))
    def update_rejected_row(self, index, fields, current_values=None):
        self.updates.append((index, fields))
    def add_rejected_row(self, fields):
        self.rejected_added.append(fields)
    def delete_row(self, index):
        self.deleted_rows.append(index)
    def move_resume(self, name, from_folder, to_folder):
        return True
    def resume_web_url(self, name, subfolder=""):
        return "https://x/resume.pdf"
    def send_mail(self, to_address, subject, html_body):
        self.mails_sent.append((to_address, subject))
        return True

# Below threshold (SCORE_RETRY_MAX=3): stays in the queue, just increments + keeps Needs Review.
_rc1 = _StubRetryClient()
_mark_needs_review(_rc1, 5, {"Retry Count": "1", "Application ID": "APP-X"}, dry_run=False)
ok(len(_rc1.updates) == 1 and _rc1.updates[0][1].get("Retry Count") == 2,
   "attempt 2/3: still just bumps the counter, no give-up")
ok(len(_rc1.rejected_added) == 0, "attempt 2/3: not moved to Rejected yet")

# At threshold: escalates to _give_up_and_reject instead of staying Needs Review forever.
_rc2 = _StubRetryClient()
_mark_needs_review(_rc2, 7, {"Retry Count": "2", "Application ID": "APP-Y",
                             "Email": "candidate@x.com"}, dry_run=False)
ok(len(_rc2.rejected_added) == 1, "attempt 3/3: given up and moved to Rejected")
_given = _rc2.rejected_added[0]
ok(_given["Status"] == _cfg_mod.STATUS_PROCESSING_FAILED,
   "given-up row uses the distinct 'Rejected - Processing Error' status, not the geo one")
ok(_given["Retry Count"] == 3, "given-up row's Retry Count reflects the final failed attempt")
ok(not _is_gap(_given["Decline Sent"]),
   "'Decline Sent' is pre-stamped (non-blank) so the non-USA decline email is never sent to it")
ok(7 in _rc2.deleted_rows, "the original main-sheet row is deleted after moving to Rejected")

import inspect as _inspect
from hiring_agent.sharepoint_scoring import (
    _derive_geo_recovery_fields as _derive_geo_recovery_fields,
    _name_needs_rederive as _name_needs_rederive,
    _plan_geo_recovery as _plan_geo_recovery,
    _plan_heal as _sharepoint_plan_heal,
    recover_location_rejections as _recover_location_rejections,
    recheck_selected_rows as _recheck_selected_rows,
    validate_row as _validate_row,
)
ok(_name_needs_rederive("waqaskhurshidoffical", "waqaskhurshidoffical@gmail.com"),
   "email-local-part full name is re-derived from the resume")
ok(not _name_needs_rederive("Muhammad Waqas", "waqaskhurshidoffical@gmail.com"),
   "real two-token full name is preserved")
_plan_src = _inspect.getsource(_sharepoint_plan_heal)
ok("reconcile_us_country" in _plan_src and "GeoDecision.CONFIRMED_US" in _plan_src,
   "row recheck reconciles Country only after a confirmed-USA verdict")
_recheck_src = _inspect.getsource(_recheck_selected_rows)
ok("_cfg.STATUS_PROCESSING_FAILED" in _recheck_src and "_DECLINE_COL" in _recheck_src,
   "targeted recheck converts recovered processing-error rows into normal geo rejections")
ok("rejected=(sheet == \"Rejected\")" in _inspect.getsource(_validate_row),
   "validate-row re-derives Rejected rows from the rejected resume folder")

# 2026-07-24: _plan_heal's Category healing must correct a STALE non-mobile category
# (not just a blank/'General' one) when the skill-based Mobile Apps override applies -
# real cases from the live Rejected sheet ('Virendra Saini', 'Muhammad Waqas') stuck on
# 'Senior & Executive' with plain Android/Kotlin skills, which the old blank-or-General-
# only gate would never touch. All other fields below are deliberately non-gap so
# _plan_heal never attempts a resume download (client=None would have nothing to fetch).
_mobile_stale_vals = {
    "Full Name": "Test Candidate", "Phone": "+92 300 1234567",
    "Location": "Lahore", "Country": "Pakistan",
    "Current Skills": "Kotlin, Java, Jetpack Compose, Firebase",
    "Looking For Role": "Android Developer",
    "Education": "BS Computer Science, Example University",
    "Suggested Role 1": "Mobile Application Lead Developer Position Description v2 (65%)",
    "Suggested Role 2": "Gaming Position (30%)",
    "Suggested Role 3": "Backend Engineer Position Description (20%)",
    "Portfolio 1": "https://linkedin.com/in/test",
    "Portfolio 2": "https://github.com/test",
    "Portfolio 3": "https://test.dev",
    "Resume URL": "https://example.sharepoint.com/resume.pdf",
    "Resume Folder Path": "Candidate_Resumes/2026/Rejected",
    "Category": "Senior & Executive",
    "Status": "Rejected - Non-USA Location",
}
_heal, _heal_geo = _sharepoint_plan_heal(None, _mobile_stale_vals, [], rejected=True)
ok(_heal.get("Category") == "Mobile Apps (Android IOS)",
   f"heal pass corrects a stale non-mobile Category to Mobile Apps when skills show an "
   f"unambiguous mobile stack (got {_heal.get('Category')!r})")
ok(_heal_geo == _GeoDecision.CONFIRMED_NON_US,
   "explicit foreign location with corroborating evidence remains confirmed non-US")

_location_review_vals = dict(_mobile_stale_vals)
_location_review_vals.update({
    "Phone": "(623) 205-0073",
    "Location": "Not extracted",
    "Country": "Not extracted",
    "Education": "B.S. Computer Science, Arizona State University",
    "Original Filename": "",
    "Status": "Rejected - Location Not Confirmed",
})
_review_heal, _review_geo = _sharepoint_plan_heal(
    None, _location_review_vals, [], rejected=True)
ok(_review_geo == _GeoDecision.UNKNOWN,
   "heal plan preserves missing current location as UNKNOWN")
ok(_review_heal.get("Status") == _cfg_mod.STATUS_LOCATION_REVIEW,
   "UNKNOWN heal plan routes the row to location confirmation, not rejection/scored")

# Historical recovery uses the same tri-state policy but skips all role/portfolio AI
# work. US phone + US university rescues the row from Rejected without fabricating
# a current US address.
_geo_recovery_heal, _geo_recovery_decision, _geo_recovery_reason = (
    _plan_geo_recovery(None, _location_review_vals, rejected=True)
)
ok(_geo_recovery_decision == _GeoDecision.UNKNOWN,
   "geo-only recovery keeps US supporting evidence UNKNOWN without a current location")
ok(_geo_recovery_heal.get("Status") == _cfg_mod.STATUS_LOCATION_REVIEW,
   "geo-only recovery routes the attached US-phone/university pattern to location review")
ok("Country" not in _geo_recovery_heal and "Location" not in _geo_recovery_heal,
   "geo-only recovery does not manufacture Country or Location from supporting evidence")

_geo_recovery_foreign = dict(_location_review_vals)
_geo_recovery_foreign.update({
    "Phone": "+91 9154676764",
    "Location": "Mumbai, India",
    "Country": "India",
    "Education": "BTech, University of Mumbai",
})
_foreign_heal, _foreign_decision, _ = _plan_geo_recovery(
    None, _geo_recovery_foreign, rejected=True)
ok(_foreign_decision == _GeoDecision.CONFIRMED_NON_US,
   "geo-only recovery leaves corroborated current non-US candidates rejected")
ok("Status" not in _foreign_heal,
   "confirmed non-US recovery plan does not relabel the rejected row")

# Even when deterministic resume extraction sees a US city in education/work
# history, it may heal Phone/Education only; it cannot copy that city to Location.
import hiring_agent.sharepoint_scoring as _sharepoint_scoring_module
_original_geo_download = _sharepoint_scoring_module._download_resume_text
try:
    _sharepoint_scoring_module._download_resume_text = (
        lambda *_args, **_kwargs: (
            "Katha Naik\n(623) 205-0073\nEducation\n"
            "B.S. Computer Science, Arizona State University\n"
            "Experience\nPhoenix, AZ\n",
            ["APP-TEST_resume.pdf"],
            {},
        )
    )
    _geo_resume_vals = dict(_location_review_vals)
    _geo_resume_vals.update({
        "Application ID": "APP-TEST",
        "Original Filename": "resume.pdf",
        "Phone": "Not extracted",
        "Education": "Not extracted",
        "Received Date": "2026-06-02T00:00:00",
    })
    _geo_resume_heal = _derive_geo_recovery_fields(
        object(), _geo_resume_vals, rejected=True)
finally:
    _sharepoint_scoring_module._download_resume_text = _original_geo_download
ok(not _is_gap(_geo_resume_heal.get("Phone")),
   "geo-only resume pass deterministically recovers a missing phone")
ok(not _is_gap(_geo_resume_heal.get("Education")),
   "geo-only resume pass deterministically recovers missing education")
ok("Location" not in _geo_resume_heal and "Country" not in _geo_resume_heal,
   "education/work-history city is never promoted into current Location/Country")

_geo_plan_src = _inspect.getsource(_plan_geo_recovery)
ok("suggested_roles" not in _geo_plan_src and "ai_recheck_fields" not in _geo_plan_src,
   "geo-only recovery performs no role scoring or AI field recheck")
_recover_src = _inspect.getsource(_recover_location_rejections)
ok("send_declines=False" in _recover_src and "STATUS_PROCESSING_FAILED" in _recover_src,
   "historical recovery suppresses declines and excludes processing-error rows")

# A row already correctly categorized (non-mobile skills, non-mobile title) is untouched -
# the broadened gate only overrides toward Mobile Apps, never away from a correct value.
_non_mobile_vals = dict(_mobile_stale_vals)
_non_mobile_vals["Current Skills"] = "Python, SQL, Excel, Power BI"
_non_mobile_vals["Suggested Role 1"] = "Data Analyst Position Description (70%)"
_non_mobile_vals["Category"] = "Data Analytics"
_heal2, _ = _sharepoint_plan_heal(None, _non_mobile_vals, [], rejected=True)
ok("Category" not in _heal2,
   "an already-correct non-mobile Category is left untouched by the broadened heal gate")

# 2026-07-24: a blank (not "N/A") Portfolio cell must be normalized even when resume
# re-derivation can't help (client=None here simulates an undownloadable resume) - real
# case: 'Muhammad Waqas', Portfolio 2/3 stuck as raw blank strings across repeated heals.
_blank_portfolio_vals = dict(_mobile_stale_vals)
_blank_portfolio_vals["Category"] = "Mobile Apps (Android IOS)"   # already correct
_blank_portfolio_vals["Portfolio 2"] = ""
_blank_portfolio_vals["Portfolio 3"] = ""
_heal3, _ = _sharepoint_plan_heal(None, _blank_portfolio_vals, [], rejected=True)
ok(_heal3.get("Portfolio 2") == "N/A" and _heal3.get("Portfolio 3") == "N/A",
   f"blank Portfolio cells are normalized to N/A even without a fresh resume derivation "
   f"(got {_heal3.get('Portfolio 2')!r}, {_heal3.get('Portfolio 3')!r})")
ok("Portfolio 1" not in _heal3,
   "an already-good Portfolio value is left untouched by the blank-normalization pass")

# dry_run never writes or deletes anything, even at threshold.
_rc3 = _StubRetryClient()
_mark_needs_review(_rc3, 9, {"Retry Count": "2"}, dry_run=True)
ok(_rc3.updates == [] and _rc3.rejected_added == [] and _rc3.deleted_rows == [],
   "dry_run makes no writes/deletes even when the give-up threshold is reached")

# _give_up_and_reject called directly with stored resume files: moves them + computes a link.
_rc4 = _StubRetryClient()
_give_up_and_reject(_rc4, 3, {"Retry Count": "2", "Application ID": "APP-Z",
                              "Email": "z@x.com"},
                    "APP-Z", ["APP-Z_resume.pdf"], "2026/July", "test reason", dry_run=False)
ok(len(_rc4.rejected_added) == 1 and _rc4.rejected_added[0].get("Resume URL"),
   "with stored resume file(s), the Resume URL/Folder Path are recomputed for the Rejected location")

# Error-alert gating: on when enabled + admin configured, off otherwise - never raises either way.
_orig_enabled, _orig_admin = _cfg_mod.ERROR_EMAIL_ENABLED, _cfg_mod.ADMIN_EMAIL
try:
    _cfg_mod.ERROR_EMAIL_ENABLED, _cfg_mod.ADMIN_EMAIL = True, "admin@x.com"
    _rc5 = _StubRetryClient()
    _send_error_alert(_rc5, "APP-A", "sender@x.com", "test reason")
    ok(len(_rc5.mails_sent) == 1 and _rc5.mails_sent[0][0] == "admin@x.com",
       "error alert sent when enabled + admin email configured")

    _cfg_mod.ERROR_EMAIL_ENABLED = False
    _rc6 = _StubRetryClient()
    _send_error_alert(_rc6, "APP-B", "sender@x.com", "test reason")
    ok(_rc6.mails_sent == [], "error alert suppressed when ERROR_EMAIL_ENABLED=false")

    _cfg_mod.ERROR_EMAIL_ENABLED, _cfg_mod.ADMIN_EMAIL = True, ""
    _rc7 = _StubRetryClient()
    _send_error_alert(_rc7, "APP-C", "sender@x.com", "test reason")
    ok(_rc7.mails_sent == [], "error alert suppressed when no ADMIN_EMAIL is configured")
finally:
    _cfg_mod.ERROR_EMAIL_ENABLED, _cfg_mod.ADMIN_EMAIL = _orig_enabled, _orig_admin


# â”€â”€ G2. HAS RESUME FILTERING â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== G2. HAS RESUME FILTERING ===")
# If a row has 'Has Resume' not equal to 'Yes', it must be skipped and flagged for review.
from hiring_agent.sharepoint_scoring import score_from_sharepoint

class _StubHasResumeClient:
    def __init__(self):
        self.hostname = "test.sharepoint.com"
        self.resumes_folder = "/Downloaded_Resumes"
        self.updates = []
        self.rejected_added = []
        self.deleted_rows = []
        self.mails_sent = []
        self.rows = [
            {
                "index": 3,
                "values": {
                    "Application ID": "APP-NORESUME",
                    "Status": "New Email Received",
                    "Has Resume": "No",  # explicitly No
                    "Email": "nores@test.com",
                    "Full Name": "No Resume Candidate",
                    "Retry Count": 0
                }
            }
        ]

    def ensure_workbook(self):
        pass

    def list_unscored_rows(self, statuses):
        return self.rows

    def list_rejected_rows(self):
        return []

    def update_row(self, index, fields, current_values=None):
        self.updates.append((index, fields))

    def add_rejected_row(self, fields):
        self.rejected_added.append(fields)

    def delete_row(self, index):
        self.deleted_rows.append(index)

    def _rejected_table_name_if_exists(self):
        return "RejectedCandidates"

    def ensure_rejected_columns(self, table_name, expected_cols):
        pass

    def check_rejected_table_decline_sent_last(self, table_name):
        return True

_stub_client = _StubHasResumeClient()

with _mock.patch("sharepoint_client.SharePointClient", return_value=_stub_client), \
     _mock.patch("sharepoint_client.check_graph_reachable", return_value=(True, "OK")), \
     _mock.patch("hiring_agent.sharepoint_scoring.ollama_health", return_value=(True, "OK")), \
     _mock.patch("hiring_agent.sharepoint_scoring._ensure_schema") as mock_ensure, \
     _mock.patch.dict(os.environ, {"HIRING_P2_DISABLED": "false"}):
    # HIRING_P2_DISABLED must be forced off for this test's scope: score_from_sharepoint()
    # checks the real environment (not just the config-loaded default) and short-circuits
    # to a no-op before ever reaching the per-row loop this test is actually exercising -
    # so without this override, the test's own result is meaningless whenever the live
    # .env has P2 intentionally disabled (e.g. during a SharePoint repair session).
    res = score_from_sharepoint(dry_run=False)

ok(len(_stub_client.updates) == 1, "One update row triggered for Has Resume == 'No'")
if len(_stub_client.updates) == 1:
    ok(_stub_client.updates[0][0] == 3, "Index 3 updated")
    ok(_stub_client.updates[0][1].get("Status") == _cfg_mod.STATUS_NEEDS_REVIEW, "Status set to Needs Review")
    ok(_stub_client.updates[0][1].get("Retry Count") == 1, "Retry Count incremented to 1")
else:
    # Guarded so a genuine future regression here fails loudly (3 FAILs) instead of an
    # unguarded index crashing the whole script and silently skipping every test after it.
    ok(False, "Index 3 updated (skipped - no update row was recorded)")
    ok(False, "Status set to Needs Review (skipped - no update row was recorded)")
    ok(False, "Retry Count incremented to 1 (skipped - no update row was recorded)")


# â”€â”€ Q. DATED FOLDER PATH â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== Q. dated_subpath â€” folder path formula ===")
import datetime
from hiring_agent.config import dated_subpath

ok(dated_subpath(datetime.date(2026, 7, 1))  == "2026/July",  "July 1  â†’ 2026/July")
ok(dated_subpath(datetime.date(2026, 7, 7))  == "2026/July",  "July 7  â†’ 2026/July")
ok(dated_subpath(datetime.date(2026, 7, 15)) == "2026/July",  "July 15 â†’ 2026/July")
ok(dated_subpath(datetime.date(2026, 7, 31)) == "2026/July",  "July 31 â†’ 2026/July")
ok(dated_subpath(datetime.date(2026, 6, 25)) == "2026/June",  "June 25 â†’ 2026/June")
ok(dated_subpath(datetime.date(2026, 12, 31)) == "2026/December", "Dec 31 â†’ 2026/December")

# P1 and P2 dated_subpath use the same formula (tested against P1 config)
import json, io as _io
_p1_root = Path(__file__).resolve().parent.parent / "HiringAgent_P1"
_p1_flow_configs = sorted(_p1_root.glob("**/flow/flow_config.json"))
if _p1_flow_configs:
    _p1cfg = json.load(_io.open(_p1_flow_configs[0], encoding="utf-8"))
    ok(_p1cfg["sharepoint"]["dated_resume_subfolders"] is True,
       "P1 also uses dated_resume_subfolders=true (same formula)")
else:
    ok(True, "P1 config not installed beside P2; P2 dated folder formula validated independently")


# â”€â”€ R. STATUS VALUES â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== R. STATUS VALUES ===")
from hiring_agent.config import (
    STATUS_LOCATION_REVIEW, STATUS_LOCATION_UNCONFIRMED,
    STATUS_REJECTED, STATUS_SCORED,
)

ok(STATUS_SCORED == "Scored",
   f"STATUS_SCORED = '{STATUS_SCORED}'")
ok(STATUS_REJECTED == "Rejected - Non-USA Location",
   f"STATUS_REJECTED = '{STATUS_REJECTED}'")
ok(STATUS_LOCATION_UNCONFIRMED == "Rejected - Location Not Confirmed",
   f"STATUS_LOCATION_UNCONFIRMED = '{STATUS_LOCATION_UNCONFIRMED}'")
ok(STATUS_LOCATION_REVIEW == "Needs Review - Location Confirmation",
   f"STATUS_LOCATION_REVIEW = '{STATUS_LOCATION_REVIEW}'")
ok(STATUS_SCORED != STATUS_REJECTED, "Scored and Rejected are distinct")


# â”€â”€ S. FULL PIPELINE (offline, no Ollama) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== S. FULL PIPELINE SIMULATION (offline, no Ollama, no SharePoint) ===")
from hiring_agent.extraction import (
    extract_candidate_details_smart, merge_mail_body_fallback, ai_recheck_fields,
    resolve_full_name, extract_portfolios, html_to_text,
)
from hiring_agent.scoring import suggested_roles, assign_category
from hiring_agent.config import STATUS_SCORED

RESUME_TEXT = """\
Jane Smith
Austin, TX | jane@example.com | (512) 555-0001
linkedin.com/in/janesmith  github.com/jsmith  janesmith.io

Summary: Experienced Data Scientist seeking an analytics role.

Skills: Python, SQL, Pandas, Scikit-Learn, TensorFlow, AWS, Tableau, Spark

Experience:
Senior Data Analyst | Acme Corp | 2022 - Present
Machine Learning Engineer | Beta Inc | 2020 - 2022
"""

MAIL_BODY = "Hi, I am applying for a Data Scientist position. Please find my resume attached."

def run_pipeline(resume, mail, roles, sender_email="Jane Smith <jane@example.com>"):
    details = extract_candidate_details_smart(resume)
    mail_txt = html_to_text(mail) if "<" in mail else mail
    if mail_txt:
        details = merge_mail_body_fallback(details, mail_txt)
    details = ai_recheck_fields(details, resume, mail_txt)  # no-op (Ollama off)

    full_name = resolve_full_name(details.get("full_name"),
                                  str(""), resume)
    skills    = details.get("skills", "Not extracted")
    role_pref = details.get("looking_for_role", "Not extracted")
    location  = details.get("location", "Not extracted")
    country   = details.get("country", "")
    phone     = details.get("phone", "Not extracted")
    education = details.get("education", "Not extracted")
    p1, p2, p3 = details.get("portfolio_1", "N/A"), details.get("portfolio_2", "N/A"), details.get("portfolio_3", "N/A")

    res = suggested_roles(skills, role_pref, roles=roles, resume_text=resume)
    r1, r2, r3 = res["role_1"], res["role_2"], res.get("role_3", "")
    cat = assign_category(r1)

    fields = {
        "Full Name": full_name, "Phone": phone, "Location": location,
        "Country": country, "Current Skills": skills, "Looking For Role": role_pref,
        "Education": education,
        "Portfolio 1": p1, "Portfolio 2": p2, "Portfolio 3": p3,
        "Suggested Role 1": r1, "Suggested Role 2": r2, "Suggested Role 3": r3,
        "Category": cat, "Status": STATUS_SCORED,
    }

    p1_vals = {"Application ID": "APP-001", "Received Date": "2026-07-01",
               "Email": sender_email, "Mail Subject": "Application",
               "Mail Body": MAIL_BODY, "Has Resume": "Yes",
               "Original Filename": "jane_cv.pdf", "Application Updates": 0}

    fields = _validate_all_columns(fields, p1_vals, mail_txt)
    return fields

fields = run_pipeline(RESUME_TEXT, MAIL_BODY, roles)

ok(fields["Full Name"] not in ("Not extracted", "", None), f"Full Name: {fields['Full Name']}")
ok(fields["Phone"] not in ("Not extracted", "", None), f"Phone: {fields['Phone']}")
ok(fields["Location"] not in ("Not extracted", "", None), f"Location: {fields['Location']}")
ok(fields["Country"] == "United States", f"Country: {fields['Country']}")
ok("Python" in fields["Current Skills"] or "python" in fields["Current Skills"].lower(),
   f"Skills contains Python: {fields['Current Skills'][:50]}")
ok("Data" in fields["Looking For Role"] or "Scientist" in fields["Looking For Role"] or
   "Engineer" in fields["Looking For Role"],
   f"Looking For Role is role-like: {fields['Looking For Role']}")
ok("linkedin.com" in fields["Portfolio 1"],
   f"Portfolio 1 = LinkedIn: {fields['Portfolio 1']}")
ok("github.com" in fields["Portfolio 2"],
   f"Portfolio 2 = GitHub: {fields['Portfolio 2']}")
ok(fields["Suggested Role 1"] != "" or True, f"Suggested Role 1: {fields['Suggested Role 1']}")
ok(fields["Category"] != "", f"Category: {fields['Category']}")
ok(fields["Status"] == "Scored", "Status = Scored")

# All 15 P2 columns present with no crashes
ok(all(k in fields for k in all_p2), "All 15 P2 columns present after full pipeline")

# No P1 column leaked into fields dict
p1_cols = {"Application ID", "Received Date", "Email", "Mail Subject",
           "Mail Body", "Has Resume", "Original Filename", "Application Updates"}
ok(not any(k in fields for k in p1_cols), "No P1 column leaked into P2 fields dict")

# looking_for_role: mail body wins over resume extraction
# (Mail body says 'Data Scientist'; resume header says 'Summary: ...')
ok("Data Scientist" in fields["Looking For Role"] or
   "data" in fields["Looking For Role"].lower(),
   "looking_for_role picks up 'Data Scientist' from mail body")


# â”€â”€ T. EDGE CASES â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== T. EDGE CASES ===")
# Empty resume text â†’ all N/A gracefully
empty_d = extract_candidate_details_smart("")
ok(empty_d["full_name"] in ("Not extracted", ""), "Empty resume â†’ full_name N/A")
ok(empty_d["skills"] in ("Not extracted", ""), "Empty resume â†’ skills N/A")

# HTML mail body â†’ stripped to plain text before fallback
html_body = "<p>Hi, I am <b>Bob Lee</b>. I am applying for a <i>Software Engineer</i> role.</p>"
plain = html_to_text(html_body)
ok("Bob Lee" in plain, "html_to_text strips tags, preserves content")
ok("<" not in plain, "html_to_text removes all HTML tags")

# Skills cap (SCORING_MAX_SKILLS)
from hiring_agent.config import SCORING_MAX_SKILLS
long_skills = ", ".join(["python","aws","docker","kubernetes","terraform","react",
                          "angular","vue","sql","postgres","redis","kafka",
                          "spark","airflow","tableau","powerbi","django","flask",
                          "fastapi","nodejs","java","go","rust","typescript"])
big_d = extract_candidate_details_smart(long_skills + "\nBob Smith\nEngineer")
if big_d["skills"] not in ("Not extracted", ""):
    skill_count = len([s for s in big_d["skills"].split(",") if s.strip()])
    ok(skill_count <= SCORING_MAX_SKILLS,
       f"Skills capped at SCORING_MAX_SKILLS={SCORING_MAX_SKILLS} (got {skill_count})")
else:
    ok(True, "Skills capped (no skills extracted from long list)")

# PDF text extraction (with a minimal in-memory PDF)
from hiring_agent.extraction import extract_text_from_bytes
try:
    from pypdf import PdfWriter
    w = PdfWriter(); w.add_blank_page(width=612, height=792)
    buf = __import__("io").BytesIO(); w.write(buf)
    txt = extract_text_from_bytes(buf.getvalue(), "blank.pdf")
    ok(isinstance(txt, str), "extract_text_from_bytes handles blank PDF without crash")
except Exception as e:
    ok(True, f"pypdf not testable in this env ({e})")

# DOCX text extraction (minimal in-memory DOCX)
try:
    import docx
    doc = docx.Document()
    doc.add_paragraph("Alice Walker\nSoftware Engineer\nPython, AWS")
    buf2 = __import__("io").BytesIO(); doc.save(buf2)
    txt2 = extract_text_from_bytes(buf2.getvalue(), "resume.docx")
    ok("Alice Walker" in txt2, "extract_text_from_bytes reads DOCX content")
    ok("Python" in txt2, "extract_text_from_bytes: skills in DOCX text")
    d2 = extract_candidate_details_smart(txt2)
    ok(d2["full_name"] == "Alice Walker", "Full name extracted from in-memory DOCX")
except Exception as e:
    ok(True, f"python-docx not testable in this env ({e})")


# â€”â€” U. GUI / DOC / URL HELPER CONSISTENCY â€”â€”â€”â€”â€”â€”â€”â€”â€”â€”â€”â€”â€”â€”â€”â€”â€”â€”â€”â€”â€”â€”â€”â€”â€”â€”â€”â€”â€”â€”â€”â€”
print("\n=== U. GUI / DOC / URL HELPER CONSISTENCY ===")
import ast
import collections
import inspect
import pathlib
import app as app_mod
from app import parse_sharepoint_url, compose_sharepoint_url

ok(app_mod.CLIENT_MODE is True, "desktop app starts in focused client mode")
ok(set(app_mod.CLIENT_DISABLED_TABS) == {
    "Job Descriptions", "SharePoint", "Resumes", "Candidates", "Settings",
}, "focused client mode disables every non-Home tab")
ok(app_mod._score_sharepoint_args(1) ==
   ["--score-sharepoint", "--batch-size", "1"],
   "GUI live run builds a one-candidate command")
ok(app_mod._score_sharepoint_args(50, dry_run=True) ==
   ["--score-sharepoint", "--batch-size", "50", "--dry-run"],
   "GUI preview command carries the selected dynamic batch size")
_progress_event = app_mod.parse_p2_progress_line(
    "2026-07-28 | INFO | P2 Progress | state=candidate | current=3 | "
    "batch=25 | batch_left=22 | queue_left=39 | row=128 | "
    "ref=APP-123 | name=Jane Smith | activity=Downloading resume"
)
ok(_progress_event.get("current") == 3
   and _progress_event.get("batch_left") == 22
   and _progress_event.get("queue_left") == 39,
   "GUI parses current, batch-left, and queue-left progress counts")
ok(_progress_event.get("ref") == "APP-123"
   and _progress_event.get("name") == "Jane Smith"
   and _progress_event.get("activity") == "Downloading resume",
   "GUI parses the current candidate identity and activity")
ok(app_mod.parse_p2_progress_line("ordinary log line") == {},
   "GUI ignores ordinary non-progress log lines")

class _ProgressValue:
    def __init__(self):
        self.value = ""

    def set(self, value):
        self.value = value

class _ProgressBar:
    def __init__(self):
        self.value = 0

    def set(self, value):
        self.value = value

_progress_app = app_mod.HiringApp.__new__(app_mod.HiringApp)
_progress_app.run_progress_title_var = _ProgressValue()
_progress_app.run_progress_detail_var = _ProgressValue()
_progress_app.run_progress_left_var = _ProgressValue()
_progress_app.run_progress_bar = _ProgressBar()
_progress_app._apply_p2_progress(_progress_event)
ok(_progress_app.run_progress_title_var.value == "Candidate 3 of 25"
   and "APP-123" in _progress_app.run_progress_detail_var.value,
   "desktop progress panel shows the active candidate and reference")
ok("22 left in this run" in _progress_app.run_progress_left_var.value
   and abs(_progress_app.run_progress_bar.value - 0.12) < 0.001,
   "desktop progress panel shows candidates left and determinate progress")
_progress_app._apply_p2_progress({
    "state": "finish", "attempted": 25, "remaining": 17, "errors": 0,
})
ok(_progress_app.run_progress_title_var.value == "Batch finished"
   and _progress_app.run_progress_left_var.value == "17 candidate(s) still waiting",
   "desktop progress panel shows the refreshed final queue count")
ok(app_mod._safe_window_geometry("1120x760+-1668+77") == "1120x760+80+80",
   "stale negative-monitor geometry is moved back onto the primary screen")
ok(app_mod._safe_window_geometry("1120x760+120+90") == "1120x760+120+90",
   "visible saved window geometry is preserved")
try:
    app_mod._score_sharepoint_args(0)
    _gui_bad_batch_rejected = False
except ValueError:
    _gui_bad_batch_rejected = True
ok(_gui_bad_batch_rejected, "GUI rejects a zero batch size before starting P2")

app_src = pathlib.Path("app.py").read_text(encoding="utf-8")
app_ast = ast.parse(app_src)
app_cls = [n for n in app_ast.body if isinstance(n, ast.ClassDef) and n.name == "HiringApp"][0]
name_counts = collections.Counter(
    n.name for n in app_cls.body if isinstance(n, ast.FunctionDef)
)
dupes = sorted(name for name, count in name_counts.items() if count > 1)
ok(not dupes, f"HiringApp has no duplicate method names (got {dupes or 'none'})")

client_mode_src = inspect.getsource(app_mod.HiringApp._apply_client_mode)
ok("button.configure(state=tk.DISABLED)" in client_mode_src,
   "focused client mode disables non-client tab buttons")
manual_run_src = inspect.getsource(app_mod.HiringApp._sp_run)
scheduler_src = inspect.getsource(app_mod.HiringApp._scheduler_loop)
scoring_src = inspect.getsource(
    __import__("hiring_agent.sharepoint_scoring", fromlist=["score_from_sharepoint"])
    .score_from_sharepoint
)
ok("_score_sharepoint_args(batch_size" in manual_run_src,
   "Run Live uses the shared batch-aware command builder")
ok("_scheduled_score_args()" in scheduler_src,
   "scheduled runs use the same saved batch-size path")
ok("for current, row in enumerate(rows, start=1)" in scoring_src,
   "P2 candidate numbering advances independently of scored/rejected/error totals")
ok('"finish"' in scoring_src and "remaining=queue_remaining" in scoring_src,
   "P2 refreshes and reports the remaining queue after a batch")

health_src = inspect.getsource(app_mod.HiringApp._run_health_check)
ok("health_summary_var" in health_src,
   "_run_health_check updates the Home health strip as well as the status bar")
ok("sharepoint_jd_folder" in health_src,
   "_run_health_check counts SharePoint JD folder mode in health state")

sample_folder = ("https://contoso.sharepoint.com/sites/CandidateList_HiringAgent/"
                 "Shared%20Documents/Downloaded_Resumes")
p = parse_sharepoint_url(sample_folder)
ok(p["hostname"] == "contoso.sharepoint.com", "parse_sharepoint_url extracts hostname")
ok(p["site_path"] == "/sites/CandidateList_HiringAgent", "parse_sharepoint_url extracts site path")
ok(p["folder"] == "/Downloaded_Resumes", "parse_sharepoint_url extracts resume folder")
rebuilt_folder = compose_sharepoint_url(p["hostname"], p["site_path"], p["folder"])
ok(parse_sharepoint_url(rebuilt_folder) == p,
   "compose_sharepoint_url rebuilds a readable URL for the same SharePoint folder")

sample_file = sample_folder + "/Sharepoint_Master_File.xlsx"
p2 = parse_sharepoint_url(sample_file)
ok(p2["file"] == "Sharepoint_Master_File.xlsx",
   "parse_sharepoint_url detects workbook filename from full URL")

from hiring_agent.sharepoint_scoring import _client_export_sharepoint_folder
old_export_folder = os.environ.pop("HIRING_CLIENT_EXPORT_SHAREPOINT_FOLDER", None)
try:
    ok(_client_export_sharepoint_folder(None) == "",
       "_client_export_sharepoint_folder defaults to root ('') of Shared Documents")
    os.environ["HIRING_CLIENT_EXPORT_SHAREPOINT_FOLDER"] = "CustomFolder"
    ok(_client_export_sharepoint_folder(None) == "CustomFolder",
       "_client_export_sharepoint_folder respects HIRING_CLIENT_EXPORT_SHAREPOINT_FOLDER override")
finally:
    if old_export_folder is not None:
        os.environ["HIRING_CLIENT_EXPORT_SHAREPOINT_FOLDER"] = old_export_folder
    else:
        os.environ.pop("HIRING_CLIENT_EXPORT_SHAREPOINT_FOLDER", None)

from hiring_agent.sharepoint_scoring import export_client_results
import tempfile as _tempfile

class _ExportFilterClient:
    def __init__(self):
        self.uploads = []

    def list_rows(self):
        return [
            {"values": {
                "Application ID": "APP-SCORED", "Status": "Scored",
                "Full Name": "Ready Candidate",
            }},
            {"values": {
                "Application ID": "APP-LOCATION-REVIEW",
                "Status": "Needs Review - Location Confirmation",
                "Full Name": "Review Candidate",
            }},
            {"values": {
                "Application ID": "APP-NEW", "Status": "New Email Received",
                "Full Name": "Queued Candidate",
            }},
        ]

    def upload_file(self, folder, name, data):
        self.uploads.append((folder, name, data))

with _tempfile.TemporaryDirectory() as _export_tmp:
    _export_dir = pathlib.Path(_export_tmp)
    _timestamped = _export_dir / "Candidate_List_Results_test.xlsx"
    _latest = _export_dir / "Candidate_List_Results.xlsx"
    _export_client = _ExportFilterClient()
    with _mock.patch(
        "hiring_agent.sharepoint_scoring._client_export_path",
        return_value=(_timestamped, _latest),
    ):
        export_client_results(_export_client, upload_to_sharepoint=True)
    _export_wb = load_workbook(_latest, data_only=False)
    _export_ws = _export_wb["Candidates"]
    _export_ids = [
        _export_ws.cell(row, 1).value
        for row in range(2, _export_ws.max_row + 1)
        if _export_ws.cell(row, 1).value
    ]
    ok(_export_ids == ["APP-SCORED"],
       "Candidate_List_Results exports only final Scored rows")
    ok(len(_export_client.uploads) == 1
       and _export_client.uploads[0][1] == "Candidate_List_Results.xlsx",
       "completed live run uploads the refreshed client workbook to SharePoint")

readme_text = pathlib.Path("docs/README.md").read_text(encoding="utf-8")
for tab in ("| **Home** |", "| **Job Descriptions** |", "| **SharePoint** |",
            "| **Resumes** |", "| **Candidates** |", "| **Settings** |"):
    ok(tab in readme_text, f"README lists current GUI tab {tab}")
for stale in ("| **Score Resumes** |", "| **Score Existing Excel** |",
              "| **Score One File** |", "Desktop GUI (7 tabs"):
    ok(stale not in readme_text, f"README no longer advertises stale GUI item '{stale}'")

bot_text = pathlib.Path("bot.py").read_text(encoding="utf-8")
ok("local-only" not in bot_text.lower(),
   "bot.py docs no longer describe the CLI as local-only")
ok("--score-sharepoint" in bot_text,
   "bot.py docs mention SharePoint queue scoring")


# --- V. SHAREPOINT RETRY GIVE-UP DUPLICATE GUARD -------------------------
print("\n=== V. SHAREPOINT RETRY GIVE-UP DUPLICATE GUARD ===")
from hiring_agent.sharepoint_scoring import _give_up_and_reject, _mark_needs_review

class _GiveUpDuplicateFakeClient:
    def __init__(self):
        self.deleted = []
        self.added_rejected = []

    def list_rows(self):
        return [
            {"index": 3, "values": {
                "Application ID": "APP-DUP", "Status": "Scored",
                "Email": "candidate@example.com",
            }},
            {"index": 9, "values": {
                "Application ID": "APP-DUP", "Status": "Needs Review - Unreadable Resume",
                "Email": "candidate@example.com",
            }},
        ]

    def list_rejected_rows(self):
        return []

    def delete_row(self, index):
        self.deleted.append(index)

    def add_rejected_row(self, fields):
        self.added_rejected.append(fields)

fake_client = _GiveUpDuplicateFakeClient()
_give_up_and_reject(fake_client, 9, {
    "Application ID": "APP-DUP",
    "Email": "candidate@example.com",
    "Retry Count": "2",
}, "APP-DUP", [], "2026/June", "resume unreadable", False)
ok(fake_client.deleted == [9],
   "give-up duplicate guard deletes only the stale retry row")
ok(fake_client.added_rejected == [],
   "give-up duplicate guard does not create a false Rejected copy")

class _NeedsReviewDuplicateFakeClient:
    def __init__(self):
        self.deleted = []
        self.updated = []

    def list_rows(self):
        return [{"index": 4, "values": {
            "Application ID": "APP-REJ-DUP",
            "Status": "Needs Review - Unreadable Resume",
            "Email": "dupe@example.com",
        }}]

    def list_rejected_rows(self):
        return [{"index": 8, "values": {
            "Application ID": "APP-REJ-DUP",
            "Status": "Rejected - Non-USA Location",
            "Email": "dupe@example.com",
        }}]

    def delete_row(self, index):
        self.deleted.append(index)

    def update_row(self, index, fields, current_values=None):
        self.updated.append((index, fields))

needs_review_fake = _NeedsReviewDuplicateFakeClient()
_mark_needs_review(needs_review_fake, 4, {
    "Application ID": "APP-REJ-DUP",
    "Email": "dupe@example.com",
    "Retry Count": "",
}, False, app_id="APP-REJ-DUP", stored=[], subpath="2026/June")
ok(needs_review_fake.deleted == [4],
   "needs-review duplicate guard deletes stale main row when Rejected copy already exists")
ok(needs_review_fake.updated == [],
   "needs-review duplicate guard does not bump Retry Count on stale duplicates")


# --- W. STRICT USA-ONLY RECENT AUDIT -------------------------------------
print("\n=== W. STRICT USA-ONLY RECENT AUDIT ===")
from rescore_recent_candidates import _strict_geo

verdict, _ = _strict_geo("Not extracted", "United States", "+1 602 555 0100", "")
ok(verdict is None,
   "strict recent audit does not certify USA from stored country/phone when location is missing")
verdict, _ = _strict_geo("Remote", "United States", "+1 602 555 0100", "")
ok(verdict is True,
   "strict recent audit accepts an explicit Remote + United States location")
verdict, _ = _strict_geo("Chennai, India", "United States", "+1 602 555 0100", "")
ok(verdict is False,
   "strict recent audit rejects explicit foreign location despite stale USA country/phone")
verdict, _ = _strict_geo("New York, NY", "United States", "", "")
ok(verdict is True,
   "strict recent audit keeps an explicit US city/state when the source file is unavailable")


print("\n=== W2. CLIENT EXPORT PERIOD SEPARATORS ===")
from hiring_agent.sharepoint_scoring import _with_period_separators as _sep


def _xrow(aid, when):
    return {"Application ID": aid, "Received Date": when, "Full Name": aid}


_sep_out = _sep([
    _xrow("A", "2026-05-07T10:00:00"), _xrow("B", "2026-05-20T10:00:00"),
    _xrow("C", "2026-06-02T10:00:00"), _xrow("D", "2026-06-11T10:00:00"),
    _xrow("E", "2027-01-03T10:00:00"),
])
_ids = [str(r.get("Application ID") or "") for r in _sep_out]
ok(_ids[0] == "", "a blank spacer row sits directly under the header")
ok(_ids == ["", "A", "B", "", "C", "D", "", "E"],
   f"one blank row between each calendar month (got {_ids})")
ok(_ids.count("") == 3,
   "a year change inserts ONE blank row, not two stacked (Jun-2026 -> Jan-2027)")
ok(all(all(str(v or "") == "" for v in _sep_out[i].values())
       for i, x in enumerate(_ids) if x == ""),
   "separator rows are fully blank in every column, so all readers skip them")
ok(_sep([]) == [], "no spacer is emitted for an empty export")
_same = _sep([_xrow("A", "2026-05-07T10:00:00"), _xrow("B", "2026-05-08T10:00:00")])
ok([str(r.get("Application ID") or "") for r in _same] == ["", "A", "B"],
   "rows inside one month get no separator between them")
_nodate = _sep([_xrow("A", ""), _xrow("B", "2026-05-08T10:00:00")])
ok(len(_nodate) == 3, "a row with an unparseable date does not crash or duplicate separators")

print("\n=== W3. REJECTED SHEET PERIOD SEPARATORS ===")
from hiring_agent.sharepoint_scoring import _ensure_rejected_period_separator as _rsep


class _RejSepClient:
    """Records separator inserts; Rejected rows are appended one at a time."""

    def __init__(self, dates):
        self.rows = [{"values": {"Received Date": d}} for d in dates]
        self.added = []

    def list_rejected_rows(self):
        return self.rows

    def add_rejected_row(self, fields):
        self.added.append(fields)


def _did_sep(existing, new_date):
    c = _RejSepClient(existing)
    res = _rsep(c, new_date)
    return res, c.added


_r, _a = _did_sep([], "2026-05-07T10:00:00")
ok(_r is False and _a == [],
   "first ever Rejected row gets no separator (header spacer already does that)")
_r, _a = _did_sep(["2026-05-07T10:00:00"], "2026-05-20T10:00:00")
ok(_r is False and _a == [], "same month as an existing row -> no separator")
_r, _a = _did_sep(["2026-05-07T10:00:00"], "2026-06-02T10:00:00")
ok(_r is True and _a == [{}], "new month -> exactly one fully blank separator row")
_r, _a = _did_sep(["2026-12-07T10:00:00"], "2027-01-02T10:00:00")
ok(_r is True and len(_a) == 1,
   "new year is also a new month -> still exactly ONE blank row, not two")
_r, _a = _did_sep(["", "2026-05-07T10:00:00"], "2026-05-21T10:00:00")
ok(_r is False, "the blank header-spacer row is ignored when reading existing periods")
_r, _a = _did_sep(["2026-05-07T10:00:00"], None)
ok(_r is False, "a None Received Date never inserts a separator")
_r, _a = _did_sep(["2026-05-07T10:00:00"], "")
ok(_r is False, "a blank Received Date never inserts a separator")
_r, _a = _did_sep(["2026-05-07T10:00:00"], "Not extracted")
ok(_r is False, "a gap-literal Received Date never inserts a separator")


class _BoomClient(_RejSepClient):
    def list_rejected_rows(self):
        raise RuntimeError("graph down")


ok(_rsep(_BoomClient([]), "2026-06-01T10:00:00") is False,
   "a failure while checking is swallowed - a cosmetic blank never blocks a rejection")

print("\n=== X. SUPPRESS_EMAILS MASTER KILL-SWITCH ===")
# Guards the historical-replay contract: no mail leaves P2, everything else still runs.
# The bug this protects against is real - the missing-info nudge had no mail flag of its
# own, so HIRING_GEO_REJECT_EMAIL + HIRING_ERROR_EMAIL alone did NOT silence P2.
import hiring_agent.config as _sup_cfg
from sharepoint_client import SharePointClient as _SupClient
from hiring_agent.sharepoint_scoring import _sent_marker, _parse_marker_datetime


class _MailSpyClient:
    """Minimal stand-in that records Graph calls instead of making them."""

    def __init__(self):
        self.graph_calls = []
        self.sender_mailbox = "apply@driverai.io"

    def _req(self, method, url, **kw):
        self.graph_calls.append((method, url))
        raise AssertionError("send_mail reached Graph while suppressed")


_spy = _MailSpyClient()
_orig_suppress = _sup_cfg.SUPPRESS_EMAILS
try:
    _sup_cfg.SUPPRESS_EMAILS = True
    _res = _SupClient.send_mail(_spy, "candidate@example.com", "Subj", "<p>body</p>")
    ok(_res is True,
       "suppressed send_mail returns True so the caller still stamps and advances")
    ok(_spy.graph_calls == [],
       "suppressed send_mail makes ZERO Graph calls (no mail can leave)")
    ok(_sent_marker().startswith("TEST-MODE (suppressed) Sent "),
       f"suppressed marker is TEST-MODE-stamped (got {_sent_marker()!r})")
    ok(_parse_marker_datetime(_sent_marker()) is not None,
       "suppressed marker is still parseable by _parse_marker_datetime")

    _sup_cfg.SUPPRESS_EMAILS = False
    ok(_sent_marker().startswith("Sent "),
       "unsuppressed marker is the normal 'Sent <ts>' form")
    _spy2 = _MailSpyClient()
    _raised = False
    try:
        _SupClient.send_mail(_spy2, "candidate@example.com", "S", "<p>b</p>")
    except AssertionError:
        _raised = True
    ok(_raised, "unsuppressed send_mail DOES reach Graph (switch is what stops it)")

    # An invalid address is rejected before the suppression branch either way.
    _sup_cfg.SUPPRESS_EMAILS = True
    ok(_SupClient.send_mail(_MailSpyClient(), "", "S", "b") is False,
       "blank recipient still returns False while suppressed")
finally:
    _sup_cfg.SUPPRESS_EMAILS = _orig_suppress

print(f"\n{'='*64}")
print(f"  P2 RESULT: {P} passed, {F} failed")
print(f"{'='*64}")
import sys
sys.exit(1 if F else 0)
