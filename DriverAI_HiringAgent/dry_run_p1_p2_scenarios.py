"""Offline dry-run harness for the P1 mail gates and the P2 scoring decisions.

Writes NOTHING to SharePoint and sends NO mail. Everything runs against dummy
in-memory candidates, so this is safe to run at any time - in particular before
re-importing the P1 zip into Power Automate.

P1 side: the four gates are plain `contains()` checks over composed strings, so
they are re-implemented here EXACTLY as definition.json expresses them, reading
the real phrase lists straight out of flow_config.json (never a copy):

    LowerFrom     = toLower(from)
    LowerSubject  = toLower(subject)
    LowerBody     = toLower(body + ' ' + join(attachmentNames,' ') + ' ')
    MatchSender   = any(bad_senders   in LowerFrom)
    MatchSubject  = any(bad_subjects  in LowerSubject)
    MatchSpam     = any(SpamPhrases   in LowerSubject or LowerBody)
    MatchLinkShort= any(shorteners    in LowerSubject or LowerBody)
    HasValidResume= any attachment ending .pdf/.docx with content

    IsSpam = MatchSpam>0  OR  (MatchLinkShort>0 AND not HasValidResume)

P2 side: calls the REAL production functions (no reimplementation) with dummy
inputs - extraction, geo classification, the content-safety net, and the
pre-reject completeness gate.

Usage:  python dry_run_p1_p2_scenarios.py
"""
import io
import json
import os
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
P1_FLOW = HERE / "HiringAgent_P1" / "flow"
P2_DIR = HERE / "HiringAgent_App_P2"

# P2 must not pick up live credentials or try to reach Ollama during a dry run.
os.environ["HIRING_OLLAMA_ENABLED"] = "false"
os.environ["HIRING_OLLAMA_SCORING"] = "false"
os.environ["HIRING_SUPPRESS_EMAILS"] = "true"
sys.path.insert(0, str(P2_DIR))

_PASS = _FAIL = 0


def check(cond, label, detail=""):
    global _PASS, _FAIL
    if cond:
        _PASS += 1
        print(f"  [PASS] {label}")
    else:
        _FAIL += 1
        print(f"  [FAIL] {label}" + (f"  -> {detail}" if detail else ""))


# ── P1: load the real lists and re-implement the gates verbatim ───────────────
_cfg = json.load(io.open(P1_FLOW / "flow_config.json", encoding="utf-8"))
_sf = _cfg["spam_filters"]

BAD_SENDERS = [s.lower() for s in _sf["bad_senders"]]
BAD_SUBJECTS = [s.lower() for s in _sf["bad_subjects"]]
SPAM_PHRASES = [s.lower() for s in (
    list(_sf["spam_phrases"])
    + list(_sf.get("offensive_phrases", []))
    + list(_sf.get("malware_phrases", []))
    + list(_sf.get("foreign_scam_phrases", []))
    + list(_sf.get("vendor_solicitation_phrases", []))
)]
SHORTENERS = [s.lower() for s in _sf.get("link_shortener_phrases", [])]


def p1_route(sender, subject, body, attachments=()):
    """Return (destination, unread, reason) exactly as the flow decides."""
    lower_from = (sender or "").lower()
    lower_subject = (subject or "").lower()
    # LowerBody appends attachment names + a trailing space (malware extensions
    # are stored space-terminated on purpose, e.g. '.exe ').
    lower_body = (str(body or "") + " " + " ".join(a[0] for a in attachments) + " ").lower()

    hit = next((p for p in BAD_SENDERS if p in lower_from), None)
    if hit:
        return "Junk Email", True, f"bad sender ~ {hit!r}"

    hit = next((p for p in BAD_SUBJECTS if p in lower_subject), None)
    if hit:
        return "Junk Email", True, f"bad subject ~ {hit!r}"

    spam = [p for p in SPAM_PHRASES if p in lower_subject or p in lower_body]
    shorts = [p for p in SHORTENERS if p in lower_subject or p in lower_body]
    has_resume = any(
        n.lower().endswith((".pdf", ".docx")) and content for n, content in attachments
    )
    if spam:
        return "Junk Email", True, f"spam phrase ~ {spam[0]!r}"
    if shorts and not has_resume:
        return "Junk Email", True, f"link shortener ~ {shorts[0]!r} (no resume)"

    if not has_resume:
        return "Archive", False, "no valid resume -> CV request, NO row created"
    return "Archive", False, "accepted -> row created + acknowledgment"


# ── Dummy scenarios ──────────────────────────────────────────────────────────
PDF = ("resume.pdf", b"%PDF-1.4 fake")

SCENARIOS = [
    # (name, sender, subject, body, attachments, expect_folder, expect_row)
    ("1  genuine applicant + CV", "jane.doe@gmail.com", "Application for Backend Engineer",
     "Hi team, please find my resume attached. Thanks, Jane", [PDF], "Archive", True),

    ("2  loop guard (our own mail)", "apply@driverai.io", "Application Received - DriverAI",
     "Thanks for applying to DriverAI.", [PDF], "Junk Email", False),

    ("3  scam phrase in body", "someone@gmail.com", "Great opportunity",
     "URGENT ACTION REQUIRED: verify your account to claim your prize.", [PDF],
     "Junk Email", False),

    ("4  offensive content", "troll@gmail.com", "hey",
     "you are a fucking idiot", [], "Junk Email", False),

    ("5  malware attachment", "bad@gmail.com", "My resume",
     "See attached.", [("payload.exe", b"MZ")], "Junk Email", False),

    ("6  vendor pitch (Futurism-style)", "parths@vendorco.com",
     "Company Profile & Capability Overview",
     "As discussed, I'm sharing our Company Profile and Corporate Deck for your review. "
     "Happy to share relevant profiles, case studies, and engagement models.", [PDF],
     "Junk Email", False),

    ("7  link shortener, NO resume", "x@gmail.com", "check this",
     "see bit.ly/abc123 for my profile", [], "Junk Email", False),

    ("8  link shortener WITH resume (bypass)", "y@gmail.com", "Application",
     "portfolio at bit.ly/abc123, resume attached", [PDF], "Archive", True),

    ("9  non-English scam", "z@gmail.com", "Oportunidad",
     "Has ganado la loteria nacional, reclama tu premio.", [], "Junk Email", False),

    ("10 applicant, no CV attached", "nocv@gmail.com", "Interested in the role",
     "Hello, I would like to apply. Please let me know next steps.", [],
     "Archive", False),

    ("11 corporate-domain REAL candidate", "engineer@realcompany-example.com",
     "Full-Stack Engineer | 5 Years",
     "I am writing to express my interest in a Senior React Native role. Resume attached.",
     [PDF], "Archive", True),

    ("12 .docx resume accepted", "docx@gmail.com", "Application",
     "Resume attached.", [("cv.docx", b"PK")], "Archive", True),
]

print("=" * 74)
print("  P1 MAIL GATES - DRY RUN (no mailbox touched)")
print("=" * 74)
print(f"  lists loaded from flow_config.json: {len(BAD_SENDERS)} senders, "
      f"{len(BAD_SUBJECTS)} subjects, {len(SPAM_PHRASES)} spam, {len(SHORTENERS)} shorteners")
print()

for name, snd, subj, body, atts, want_folder, want_row in SCENARIOS:
    folder, unread, reason = p1_route(snd, subj, body, atts)
    got_row = folder == "Archive" and "NO row" not in reason
    okf = folder == want_folder
    okr = got_row == want_row
    status = "PASS" if (okf and okr) else "FAIL"
    if okf and okr:
        _PASS += 1
    else:
        _FAIL += 1
    print(f"  [{status}] {name}")
    print(f"         -> {folder} | unread={unread} | row={got_row} | {reason}")
    if not okf:
        print(f"         !! expected folder {want_folder}")
    if not okr:
        print(f"         !! expected row={want_row}")

# Junk-is-unread / Archive-is-read invariant
print()
junk_all_unread = all(
    p1_route(s, su, b, a)[1]
    for _, s, su, b, a, wf, _ in SCENARIOS if wf == "Junk Email"
)
arch_all_read = all(
    not p1_route(s, su, b, a)[1]
    for _, s, su, b, a, wf, _ in SCENARIOS if wf == "Archive"
)
check(junk_all_unread, "every Junk-routed mail is left UNREAD")
check(arch_all_read, "every Archive-routed mail is marked READ")


# ── P2: real production functions, dummy data ────────────────────────────────
print()
print("=" * 74)
print("  P2 SCORING DECISIONS - DRY RUN (no SharePoint writes)")
print("=" * 74)

from hiring_agent.geo import classify_location_usa, GeoDecision           # noqa: E402
from hiring_agent.extraction import (                                      # noqa: E402
    extract_candidate_details, _looks_like_company_name, extract_portfolios,
)
from hiring_agent.sharepoint_scoring import (                              # noqa: E402
    _detect_suspicious_content, _complete_row_before_reject, _is_gap,
)

GEO_CASES = [
    ("USA city/state",        "Austin, TX",        "United States", "512-555-0100", "", GeoDecision.CONFIRMED_US),
    ("USA, foreign country field", "Chicago, IL",  "India",         "312-555-0100", "", GeoDecision.CONFIRMED_US),
    ("clear foreign + corroboration", "Hyderabad", "India",         "+91-9848999100", "B.Tech, JNTU Hyderabad", GeoDecision.CONFIRMED_NON_US),
    ("blank location",        "",                  "",              "",             "", GeoDecision.UNKNOWN),
    ("foreign, no corroboration", "Mumbai",        "",              "",             "", GeoDecision.UNKNOWN),
]
print("\n-- geo routing (Main vs Rejected) --")
for label, loc, ctry, phone, edu, want in GEO_CASES:
    got, reason = classify_location_usa(location=loc, country=ctry, phone=phone,
                                        education=edu, resume_text="")
    sheet = {"confirmed_us": "MAIN/Scored",
             "confirmed_non_us": "REJECTED",
             "unknown": "MAIN/Needs Review"}[got.value]
    check(got == want, f"{label:32s} -> {sheet}", f"got {got.value}, want {want.value}")

print("\n-- content-safety net (P2 leftover catcher) --")
check(_detect_suspicious_content("", "sharing our Company Profile and Corporate Deck") != [],
      "vendor pitch in mail body is flagged")
check(_detect_suspicious_content("Senior React Native Developer, 5 years", "") == [],
      "genuine resume text is NOT flagged")

print("\n-- identity guards --")
check(_looks_like_company_name("Transforming Business Models"), "company tagline rejected as a person name")
check(not _looks_like_company_name("Sai Krishna Yallapu"), "real name accepted")
check(extract_portfolios("we use Socket.io for realtime") == ("N/A", "N/A", "N/A"),
      "library mention not stored as a portfolio")

print("\n-- pre-reject completeness gate --")
gap = {"Application ID": "APP-DRY-1", "Full Name": "Test User", "Location": "Mumbai",
       "Country": "India", "Status": "Rejected - Non-USA Location",
       "Current Skills": "Kotlin", "Suggested Role 1": "Backend Engineer (40%)",
       "Category": "", "Portfolio 1": "", "Portfolio 2": "", "Portfolio 3": "",
       "Application Updates": "", "Retry Count": ""}
fixed = _complete_row_before_reject(dict(gap), "APP-DRY-1")
check(not _is_gap(fixed["Category"]), "blank Category repaired before reaching Rejected")
check(fixed["Portfolio 1"] == "N/A" and fixed["Retry Count"] == 0,
      "blank portfolios/counters normalized before reaching Rejected")

print("\n-- offline extraction on a dummy resume --")
dummy = ("Jane Q Public\nAustin, TX 78701 | jane@example.com | 512-555-0100\n\n"
         "EDUCATION\nB.S. Computer Science, University of Texas\n\n"
         "SKILLS\nPython, Docker, Kubernetes\n")
det = extract_candidate_details(dummy)
check(det.get("full_name") == "Jane Q Public", f"name  -> {det.get('full_name')!r}")
check("Austin" in str(det.get("location")), f"location -> {det.get('location')!r}")
check("512" in str(det.get("phone")), f"phone -> {det.get('phone')!r}")

print()
print("=" * 74)
print(f"  DRY-RUN RESULT: {_PASS} passed, {_FAIL} failed")
print("=" * 74)
print("  No SharePoint row was created or modified. No mail was sent or moved.")
sys.exit(1 if _FAIL else 0)
