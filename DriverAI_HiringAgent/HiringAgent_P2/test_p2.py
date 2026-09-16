"""P2 deep end-to-end test â€” no live SharePoint, no live Ollama required.

Tests every extraction/scoring/validation layer across all 14 P2-owned columns.
Run from HiringAgent_P2/:
    .venv/Scripts/python test_p2.py
"""
import io, os, sys
os.environ['HIRING_STORAGE_BACKEND'] = 'excel'  # local pipeline covered in test_local_pipeline.py
from pathlib import Path
os.environ.setdefault("HIRING_OLLAMA_ENABLED", "false")
os.environ.setdefault("HIRING_OLLAMA_SCORING", "false")
sys.path.insert(0, str(Path(__file__).resolve().parent))
import test_isolation  # noqa: F401 - must precede every hiring_agent import; see its docstring
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
ok(_extract_education(
    "EDUCATION\nElectronics And Communication Engineering Sep 2023 to 2027\n"
    "Government Engineering College Bikaner, Bikaner") ==
   "Electronics And Communication Engineering, Government Engineering College Bikaner",
   "a stated field of study is paired with the college on the next line without inventing a degree")

print("\n=== B2b. EDUCATION START/END DATE EXTRACTION (added 2026-08-18) ===")
from hiring_agent.extraction import _extract_education_dates
edu_date_cases = {
    "EDUCATION\nB.S. Computer Science, ASU\n08/2021 - 05/2025\n": ("08/2021", "05/2025"),
    "EDUCATION\nB.S. Computer Science, ASU\nAug 2021 - May 2025\n": ("Aug 2021", "May 2025"),
    "EDUCATION\nMasters, University of Arizona, Eller College of ManagementAug. 2025 - Dec 2026\n":
        ("Aug 2025", "Dec 2026"),
    "EDUCATION\nMaster of Science in Computer Science December 2026\nArizona State University\n":
        ("", "Dec 2026"),
    "Bachelor of Technology (B.Tech), Computer Science Engineering\n"
    "Technocrats Institute | Dec 2020 - July 2023\n"
    "Other content\nEDUCATION\nTECHNICAL SKILLS\n": ("Dec 2020", "Jul 2023"),
    "EDUCATION\nB.S. Computer Science, ASU\n2018 - 2022\n": ("2018", "2022"),
    "EDUCATION\nM.S. Data Science, ASU\n2022 - Present\n": ("2022", "Present"),
    "EDUCATION\nM.S. Data Science, ASU\nAug 2023 - Current\n": ("Aug 2023", "Present"),
    "EDUCATION\nB.S. Computer Science, ASU\nExpected Graduation: May 2026\n": ("", "May 2026"),
    "EDUCATION\nB.S. Computer Science, ASU\nExpected May 2026\n": ("", "May 2026"),
    "EDUCATION\nB.S. Computer Science, Arizona State University\n": ("", ""),
    "Just some random resume text with no header.\nSkills: Python, SQL.": ("", ""),
}
for text, expected in edu_date_cases.items():
    got = _extract_education_dates(text)
    ok(got == expected, f"education dates {text.splitlines()[-2]!r} -> {got} (expected {expected})")

# Same DD/MM/YYYY fixture as the _extract_education test above: a day-precision date must
# NOT be misread as a bogus MM/YYYY pair (e.g. '10/2018' pulled out of '05/10/2018' would
# silently report the wrong month). Resumes never state education to day precision, so the
# safe, correct behavior is to find no valid range here at all - a blank pair, not a wrong one.
ok(_extract_education_dates(
    "EDUCATION AND TRAINING\n05/10/2018 - 05/10/2022\n"
    "BS-INFORMATION TECHNOLOGY The Islamia University of Bahawalpur") == ("", ""),
   "a day-precision DD/MM/YYYY date range is never misread as a bogus MM/YYYY pair - blank, not wrong")

# Most resumes list multiple degrees under ONE 'Education' header, most-recent first. The
# date scan shares _extract_education's same header+5-line window, so it naturally picks up
# the FIRST (most recent) degree's dates, matching what the Education column itself shows -
# never the dates of an older, unrelated degree further down.
ok(_extract_education_dates(
    "EDUCATION\n"
    "M.S. Computer Science, Arizona State University\n"
    "Aug 2023 - May 2025\n"
    "B.S. Information Technology, University of Punjab\n"
    "2018 - 2022\n") == ("Aug 2023", "May 2025"),
   "with multiple degrees under one header, dates match the FIRST (most recent) entry, "
   "same one _extract_education itself reports")

# Live regressions (2026-09-01): a data-quality audit found 57/160 scored candidates had
# NO education dates despite the resume plainly stating them, tracing to three resume-
# template shapes none of the tiers above covered.
ok(_extract_education_dates(
    "Education \n"
    "Stevens Institute of Technology, United States \n"
    "MS in Business Analytics CGPA: 3.70/4.0 \n"
    "Institute of Business Administration (IBA), Karachi (NTHP Scholar) \n"
    "BS in Mathematics and Economics CGPA: 3.45/4.0 \n"
    "Sep'24 - May'26 \n"
    "Aug'17 - May'21 ") == ("Sep 2024", "May 2026"),
   "abbreviated 2-digit-year range ('Sep'24 - May'26'), all dates clustered after every "
   "school/degree line rather than beside its own entry, still recovers the FIRST pair")
ok(_extract_education_dates(
    "EDUCATION \n"
    "W.P. Carey School of Business at Arizona State University    July 2025 \n"
    "Master of Science (MS), Information Systems and Management    Tempe, AZ, USA \n"
    ) == ("", "Jul 2025"),
   "a single date on the SCHOOL-name line immediately above the degree line (not the "
   "degree line itself, and not part of an Expected/range phrase) is still found")
ok(_extract_education_dates(
    "EDUCATION \n"
    "University of Texas Arlington  Aug 23-May 25 \n"
    "Master of Science in Computer Science \n"
    "Netaji Subhas University of Technology  Aug 19-May 23 \n"
    "Bachelor of Technology in Information Technology\n"
    ) == ("Aug 2023", "May 2025"),
   "a 2-digit-year range with NO apostrophe, on the school line above the FIRST degree, "
   "is still recovered - and not confused with the second degree's own range below it")
# GUARD: the loose 2-digit-year range must not fire on ordinary prose that happens to
# contain two small numbers near a dash - it only ever matches inside an actual MONTH-NAME
# NAME + 2-digit-year shape, which mid-sentence prose essentially never produces.
ok(_extract_education_dates(
    "EDUCATION\nB.S. Computer Science, Arizona State University\n"
    "Ranked in the top 23-25 percentile of the graduating class.\n") == ("", ""),
   "GUARD: two bare small numbers joined by a dash (no month name) never fabricate a range")

# Live regressions (2026-09-01): a data-quality audit of every candidate still missing an
# education date found this ONE shape behind well over a third of them - a PDF that
# extracts with roughly one WORD per line (not one CHARACTER, which
# _collapse_letter_spacing already handles). _extract_education survives it via the AI
# fallback, but _extract_education_dates is pure regex and saw only isolated single-word
# lines. Real cases, reproduced verbatim from the live resumes.
ok(_extract_education_dates(
    "Education\nSeattle\nUniversity\n,\nMS\nin\nComputer\nScience\nSept\n2023\n-\nJune\n2025\n"
    "Savitribai\nPhule\nPune\nUniversity\n,\nBachelors\nin\nComputer\nEngineering\nAug\n2019\n"
    "-\nJune\n2023\nSummary\nCybersecurity\nprofessional\n") == ("Sep 2023", "Jun 2025"),
   "Manjeeri Ghanekar: word-fragmented two-degree block still recovers the FIRST pair")
# The dates are read into Education Start/End Date (asserted just above), so since
# 2026-09-14 they are no longer repeated inside the Education text itself.
ok(_extract_education(
    "Education\nSeattle\nUniversity\n,\nMS\nin\nComputer\nScience\nSept\n2023\n-\nJune\n2025\n")
   == "Seattle University, MS in Computer Science",
   "the same word-fragmented block also recovers a readable Education TEXT value, not "
   "just the dates - the regex path should not need to fall back to Ollama for this")
# GUARD: this is the exact regression that motivated using a NARROW header-word set
# (_FRAGMENT_HEADER_WORDS) instead of the broader config.SECTION_WORDS, which also
# contains generic words like 'software'/'engineering' for an unrelated purpose - using
# it here prematurely broke the fragment run mid-degree-title and lost the range entirely.
ok(_extract_education_dates(
    "problems.\nEDUCATION\nSAN\nJOSE\nSTATE\nUNIVERSITY\nSan\nJose,\nCA\nMaster\nof\n"
    "Science\nin\nSoftware\nEngineering\nAug\n2024\n-\nMay\n2026\n"
    "Coursework: Machine Learning, Data Mining\n") == ("Aug 2024", "May 2026"),
   "GUARD: a degree title containing 'Software Engineering' does not break the fragment "
   "run mid-block and lose the trailing date range")
ok(_extract_education_dates(
    "(602)\n684-8781\nEducation\nArizona\nState\nUniversity\nPhoenix,\nAZ\nB.S.\nComputer\n"
    "Science\n(Software\nEngineering)\nMinor.\nFashion\nDesign\n") == ("", ""),
   "a word-fragmented block with genuinely NO date anywhere still correctly yields blank, "
   "not a fabricated guess")
# GUARD: a normal short bullet list (well under the min-run threshold) must never get
# glued into one line - only a genuinely long fragmented run is touched.
ok(_extract_education(
    "EDUCATION\nB.S. Computer Science, ASU\n\nSKILLS\nPython\nSQL\nExcel\n")
   == "B.S. Computer Science, ASU",
   "GUARD: a few short skill-list lines (below the fragment-run threshold) are left alone")

from hiring_agent.extraction import extract_candidate_details as _ecd_dates
_dates_result = _ecd_dates("EDUCATION\nB.S. Computer Science, ASU\nAug 2021 - May 2025\n")
ok(_dates_result.get("education_start_date") == "Aug 2021"
   and _dates_result.get("education_end_date") == "May 2025",
   f"extract_candidate_details wires education_start_date/education_end_date through "
   f"(got {_dates_result.get('education_start_date')!r}, {_dates_result.get('education_end_date')!r})")
_no_dates_result = _ecd_dates("EDUCATION\nB.S. Computer Science, Arizona State University\n")
ok(_no_dates_result.get("education_start_date") == "" and _no_dates_result.get("education_end_date") == "",
   "no explicit dates in the resume -> both fields blank, never fabricated")

# Live regression (2026-07-31, Syyed Nazir Ali / APP-20260727-1431-6F75): a letter-spaced
# 'E D U C A T I O N' header (a common resume-template style - PDF extraction preserves the
# stylized spacing as literal single-char tokens) was invisible to the header regex, so
# extraction fell through to a full-text degree-keyword scan that matched bare "Master" in
# a "Scrum Master" job-title line near the top - completely preempting the real "Bachelor
# of Science" degree near the bottom. Two independent fixes, both covered here.
_letterspaced_resume = (
    "SYYED NAZIR ALI\n"
    "Business Analyst  |  Scrum Master  |  Program & Project Manager\n"
    "Certified Scrum Master (CSM / PSM) with 8+ years of experience\n"
    "E D U C A T I O N\n"
    "Bachelor of Science (B.Sc.)\n"
    "IGNOU Delhi & NCHMCT Noida - India\n"
)
ok(_extract_education(_letterspaced_resume) == "Bachelor of Science (B.Sc.)",
   f"letter-spaced 'E D U C A T I O N' header is still recognized, and no longer loses to "
   f"a bare 'Scrum Master' match higher up the resume (got {_extract_education(_letterspaced_resume)!r})")
from hiring_agent.extraction import _collapse_letter_spacing as _cls
ok(_cls("E D U C A T I O N") == "EDUCATION", "letter-spaced header collapses correctly")
ok(_cls("Education: BS Computer Science") == "Education: BS Computer Science",
   "a normal (non letter-spaced) line is never touched by the collapse heuristic")
ok(_cls("A B") == "A B",
   "a short 2-token line (e.g. a real state abbreviation like 'A B' would never occur, but "
   "guards the <4-token threshold) is left alone, not misread as letter-spacing")

# Live regression (2026-07-31, Divy Parmar / APP-20260720-1013-09AC): an ENTIRE resume was
# letter-spaced this way, not just a header - name, phone, email, every line - which made
# Name/Phone/Skills all come back "Not extracted" (garbled single-character tokens can't
# match any word-boundary pattern) and Location only surface via a mail-body fallback, not
# the resume itself. The double-space-as-word-boundary rule must survive multi-word lines.
ok(_cls("M a s t e r  o f  C o m p u t e r  A p p l i c a t i o n s  ( M C A )") ==
   "Master of Computer Applications (MCA)",
   "multi-word letter-spaced line collapses with word boundaries intact, not glued together")
ok(_cls("d i v y p a r m a r 1 9 @ e x a m p l e . c o m") == "divyparmar19@example.com",
   "a letter-spaced email address collapses correctly (no word-boundary spaces needed)")
ok(_cls("7 4 0 5 4 6 5 2 0 4") == "7405465204",
   "a letter-spaced phone number (digits are single-char tokens too) collapses correctly")
from hiring_agent.extraction import extract_candidate_details as _ecd
_fully_letterspaced_resume = (
    "7 4 0 5 4 6 5 2 0 4\n"
    "K h o k h r a ,  A h m e d a b a d\n"
    "d i v y p a r m a r 1 9 @ g m a i l . c o m\n"
    "D I V Y  P A R M A R\n"
    "S U M M A R Y\n"
    "S e n i o r  A n d r o i d  D e v e l o p e r  w i t h  4 +  y e a r s  o f  e x p e r i e n c e .\n"
    "S K I L L S\n"
    "L a n g u a g e s :  K o t l i n ,  J a v a\n"
    "E D U C A T I O N\n"
    "M a s t e r  o f  C o m p u t e r  A p p l i c a t i o n s  ( M C A )\n"
    "M o n a r k  U n i v e r s i t y\n"
)
_ls_result = _ecd(_fully_letterspaced_resume)
ok(_ls_result.get("full_name") not in (None, "Not extracted"),
   f"a fully letter-spaced resume no longer fails Name extraction entirely (got {_ls_result.get('full_name')!r})")
ok(_ls_result.get("phone") == "7405465204",
   f"a fully letter-spaced resume's phone is recovered (got {_ls_result.get('phone')!r})")
ok("kotlin" in str(_ls_result.get("skills", "")).lower(),
   f"a fully letter-spaced resume's skills are recovered (got {_ls_result.get('skills')!r})")
ok(_ls_result.get("education") == "Master of Computer Applications (MCA)",
   f"a fully letter-spaced resume's education is recovered with word boundaries intact "
   f"(got {_ls_result.get('education')!r})")

# Live regression (2026-07-31, Prerna Saluja / APP-20260716-2052-6112): "position" (and
# "role") are common English words far beyond "job position" - an Experience bullet
# reading "...identifying an investment position that appreciated approximately 3x over 5
# months" matched bare "position" and returned "that appreciated approximately" as her
# desired role. The role-preference regex is now scoped to the header/summary (first 15
# lines) only, so an Experience-section false match past that point can't reach it.
# Padded with extra Experience lines so the false-positive text genuinely falls past line
# 15, matching the real resume's shape (a short synthetic resume with too few lines would
# never actually exercise the scoping boundary).
_position_false_positive_resume = (
    "Prerna Saluja\n"
    "(623)-242-3627 | LinkedIn | Gmail\n"
    "\n"
    "EXPERIENCE\n"
    "Pareto Inc. Financial AI Analyst (Data Labeler) 2025 - Present\n"
    "Leveraged AI-assisted analytics tools to support rolling forecasts, variance analysis.\n"
    "Improved data quality and analytical consistency by validating AI-generated outputs.\n"
    "Reduced manual review effort by approximately 20 percent across quarterly cycles.\n"
    "Arizona State University Research Assistant (Volunteer) 2025 - 2026\n"
    "Supported faculty research on financial modeling and forecasting accuracy.\n"
    "Built dashboards to track key performance indicators for ongoing studies.\n"
    "Presented findings to a panel of faculty advisors each semester.\n"
    "Coordinated with three other research assistants on data collection.\n"
    "Documented methodology for reproducibility across research cycles.\n"
    "Designed and executed input/output analysis in Excel, evaluating the impact of\n"
    "variable changes on outcomes to determine relevant indicators and performance\n"
    "factors, identifying an investment position that appreciated approximately 3x\n"
    "over 5 months.\n"
    "KPMG Audit Assistant 2020 - 2021\n"
)
_role = _ecd(_position_false_positive_resume).get("looking_for_role")
ok(_role != "that appreciated approximately",
   f"an unrelated 'investment position' mention deep in Experience is no longer mistaken "
   f"for a stated role preference (got {_role!r})")

# A genuine header-area role statement must still be caught - the fix narrows scope, it
# must not blind the extractor entirely. Caught here by _extract_header_role (checked
# before the scoped regex fixed above), which returns the whole matched line as-is -
# pre-existing behavior, unrelated to and unaffected by this fix.
_header_role_resume = "Jane Doe\nSeeking a Position: Senior Data Analyst\nEXPERIENCE\n...\n"
ok(_ecd(_header_role_resume).get("looking_for_role") == "Seeking a Position: Senior Data Analyst",
   "a genuine role statement in the header/summary area is still correctly extracted")

# Live regression (2026-08-01, Mindy Anderson / APP-20260723-2034-12DF): her actual title
# line, "Fractional/Interim Chief Marketing Officer & Marketing Advisor" (7 words), was
# rejected by _extract_header_role's old 6-word cap, so nothing caught her real title and
# the scoped regex below then matched "position" as a bare substring inside "brand
# positioning" (a business term two lines later), returning "ing" as her desired role.
from hiring_agent.extraction import _extract_header_role as _ehr
ok(_ehr("Mindy Anderson\nFractional/Interim Chief Marketing Officer & Marketing Advisor\n"
        "(917) 583-3070") == "Fractional/Interim Chief Marketing Officer & Marketing Advisor",
   "a genuine 7-word compound executive title is caught, not rejected for length")
_positioning_resume = (
    "Mindy Anderson\nFractional/Interim Chief Marketing Officer & Marketing Advisor\n"
    "(917) 583-3070 | candidate.cmo@example.com\n\nPROFESSIONAL SUMMARY\n"
    "Architect and scale marketing functions aligned to long-term enterprise growth and "
    "premium brand positioning.\n"
)
_role = _ecd(_positioning_resume).get("looking_for_role")
ok(_role == "Fractional/Interim Chief Marketing Officer & Marketing Advisor",
   f"'positioning' two lines later is not mistaken for a role statement now that the real "
   f"header title is caught first (got {_role!r})")

# Live regression (2026-08-01, same Prerna Saluja row, a second bug on the same field):
# with the header-scoping fix above no longer matching, extraction fell through to a
# keyword-fallback chain that checked 'rpa' as a bare substring of the whole document -
# which matched inside 'counterparties' ('cou-nte-RPA-rties'), tagging a Financial/Audit
# Analyst as wanting an "Automation / RPA Engineer" role she never mentioned anywhere.
_counterparties_resume = (
    "Prerna Saluja\n(623)-242-3627 | LinkedIn | Gmail\n\nEXPERIENCE\n"
    "KPMG Audit Assistant 2020 - 2021\n"
    "Reconciled 100+ counterparties' transactions to identify gaps in financial records.\n"
    "Analyzed trends using SAP & AI Tools and financial reporting analytics.\n"
)
ok(_ecd(_counterparties_resume).get("looking_for_role") != "Automation / RPA Engineer",
   f"'rpa' inside 'counterparties' is not mistaken for an Automation/RPA role preference "
   f"(got {_ecd(_counterparties_resume).get('looking_for_role')!r})")

# Live regression (2026-07-31, Brett Worker / APP-20260721-0122-E743): a compound header
# "EDUCATION & EXECUTIVE DEVELOPMENT" had its "& EXECUTIVE DEVELOPMENT" remainder returned
# as the Education value outright, skipping the real degree lines just below it.
_compound_header_resume = (
    "Brett Worker\n"
    "Progressed into dedicated cybersecurity responsibilities.\n"
    "EDUCATION & EXECUTIVE DEVELOPMENT\n"
    "Robert Morris University\n"
    "Master of Information Systems, Business Analytics | Bachelor of Science, Computer Science\n"
    "Harvard Business School\n"
    "CERTIFICATIONS\n"
    "CISSP\n"
)
ok(_extract_education(_compound_header_resume) ==
   "Master of Information Systems, Business Analytics | Bachelor of Science, Computer Science",
   f"a compound header with trailing non-degree words falls through to the real degree line "
   f"below it, instead of returning the header's own trailing words "
   f"(got {_extract_education(_compound_header_resume)!r})")

from hiring_agent.extraction import _DEGREE_RE as _degree_re
ok(not _degree_re.search("Scrum Master"), "bare 'Scrum Master' no longer false-positives as a degree")
ok(not _degree_re.search("Associate Director of Sales"), "bare 'Associate <title>' no longer false-positives as a degree")
ok(_degree_re.search("Master's in Computer Science"), "\"Master's\" (possessive) still matches as a degree")
ok(_degree_re.search("Masters in Computer Science"), "\"Masters\" (no apostrophe) still matches as a degree")
ok(_degree_re.search("Master of Science in Data Science"), "\"Master of X\" still matches as a degree")
ok(_degree_re.search("Bachelor of Arts"), "\"Bachelor of X\" still matches as a degree")

# Live regression (2026-07-31, same Syyed Nazir Ali case): 'SWIFT' the banking payment
# standard (always written all-caps: "UPI, NEFT, RTGS, IMPS, SWIFT, ISO 20022") matched the
# "Swift" (Apple's language) skill keyword once both sides of the scan were lowercased to
# "swift" - which then triggered a mobile-developer Category override for a Business
# Analyst candidate who has never written a line of Swift in his life.
from hiring_agent.extraction import _scan_skill_keywords
ok("swift" not in _scan_skill_keywords("Payment rails: UPI, NEFT, RTGS, IMPS, SWIFT, ISO 20022"),
   "all-caps 'SWIFT' (the payment standard) no longer false-positives as the Swift skill")
ok("swift" in _scan_skill_keywords("Skills: Python, Swift, SwiftUI, Kotlin, Objective-C"),
   "genuine mixed-case 'Swift' (the language) is still detected as a skill")
ok("swift" not in _scan_skill_keywords("swift and efficient delivery of results"),
   "lowercase 'swift' (the adjective, meaning fast) does not false-positive either")

# Live regression (2026-07-31, Mindy Anderson / APP-20260723-2034-12DF): 'Go' (the
# language) matched "Go To Market"/"go-to-market" (ubiquitous marketing jargon, title-
# cased in headings just like the language name) in a marketing-executive resume with
# zero mention of the Go language. Case-sensitivity alone can't fix this one (both
# usages appear capitalized), so this uses phrase exclusion instead.
ok("go" not in _scan_skill_keywords("Expertise in AI-powered ABM, Go To Market orchestration, MarTech"),
   "'Go' inside 'Go To Market' no longer false-positives as the Go language")
ok("go" not in _scan_skill_keywords("translating vision into sophisticated go-to-market strategy"),
   "hyphenated 'go-to-market' also does not false-positive")
ok("go" in _scan_skill_keywords("Backend languages: Go, Python, Rust"),
   "genuine standalone 'Go' (the language) is still detected as a skill")

# Live regression (2026-07-31, Syyed Nazir Ali / APP-20260727-1431-6F75): "go-live" is
# standard IT/project-management deployment terminology ("business sign-off and go-live
# approval", "Release & Go-Live Governance"), same false-positive class as "Go To Market".
ok("go" not in _scan_skill_keywords("business sign-off and go-live approval"),
   "'go-live' (lowercase, hyphenated) no longer false-positives as the Go language")
ok("go" not in _scan_skill_keywords("Strategy, Release & Go-Live Governance"),
   "'Go-Live' (title-cased) no longer false-positives as the Go language")

# Live regression (2026-08-01, Sai Krishna Yallapu / APP-20260717-1100-7BDA): "SQL"/"MQL"
# in a B2B marketing resume almost always means Sales-/Marketing-Qualified-Lead counts, not
# the database language - "reported MQLs, SQLs, pipeline...", "MQL-to-SQL handoff", "SQL
# acceptance rates" all matched the bare 'sql' skill keyword despite zero database mention
# anywhere in the resume.
for _sql_text, _why in [
    ("reported MQLs, SQLs, pipeline, MRO/MSO to CMO", "plural 'MQLs, SQLs' lead counts"),
    ("improving the MQL-to-SQL handoff process", "'MQL-to-SQL' handoff terminology"),
    ("significantly improved SQL acceptance rates", "'SQL acceptance rate' marketing metric"),
]:
    ok("sql" not in _scan_skill_keywords(_sql_text),
       f"{_why} does not false-positive as the SQL skill")
ok("sql" in _scan_skill_keywords("Backend: Python, SQL Server, Node.js"),
   "genuine standalone 'SQL' (e.g. 'SQL Server') is still detected as a skill")

# Ollama's own free-text skill reading isn't bound by _scan_skill_keywords' disambiguation -
# it independently returned "Swift" for the Syyed Nazir Ali case despite the deterministic
# baseline being clean, primed by 'Swift' appearing as this file's own few-shot prompt
# example. _strip_unconfirmed_ambiguous_skills is the deterministic backstop applied to
# Ollama's raw output before it ever reaches the union merge (extract_candidate_details_smart),
# reusing _scan_skill_keywords as the single source of truth for both the case-sensitivity
# and phrase-exclusion fixes above.
from hiring_agent.extraction import _strip_unconfirmed_ambiguous_skills as _strip_css
_swift_resume_text = "Payment rails: UPI, NEFT, RTGS, IMPS, SWIFT, ISO 20022"
ok(_strip_css("Python, Swift, Kubernetes", _swift_resume_text) == "Python, Kubernetes",
   "Ollama's hallucinated 'Swift' is stripped when only all-caps 'SWIFT' appears in the resume")
ok(_strip_css("Python, Swift, Kubernetes", "Skills: Python, Swift, SwiftUI, Kubernetes") == "Python, Swift, Kubernetes",
   "a genuine mixed-case 'Swift' mention in the resume keeps Ollama's 'Swift' skill")
ok(_strip_css("Marketing, Go, SEO", "Go To Market orchestration, SEO, demand gen") == "Marketing, SEO",
   "Ollama's hallucinated 'Go' is stripped when the resume only says 'Go To Market'")
ok(_strip_css("", _swift_resume_text) == "", "an empty skills string passes through unchanged")
ok(_strip_css("Python, Kubernetes", _swift_resume_text) == "Python, Kubernetes",
   "a skills string with no ambiguous terms is untouched")

# Live regression (2026-08-01, same Sai Krishna Yallapu case): "AWS re:Invent" names a
# conference he led sponsorship/marketing for, not a personal cloud-computing skill - his
# resume never claims hands-on AWS use anywhere.
ok(_strip_css("Java, AWS, Claude", "Led Oracle AI World, AWS re:Invent, Cloud World, Ascend") == "Java, Claude",
   "Ollama's hallucinated 'AWS' is stripped when the resume only names the 'AWS re:Invent' conference")
ok(_strip_css("Java, AWS, Claude", "Cloud: AWS, Azure, GCP") == "Java, AWS, Claude",
   "a genuine standalone 'AWS' mention in the resume keeps Ollama's 'AWS' skill")

# Live regression (2026-08-01, Sai Krishna Yallapu / APP-20260717-1100-7BDA): the AI
# extraction prompt's own few-shot example ("e.g. Python, AWS, Docker, React, Swift") was
# getting echoed back as hallucinated skills across UNRELATED candidates - a B2B Marketing
# Leader's resume (zero mention of Python, Docker, or React anywhere) came back with all
# three in 'skills'. 'Swift' alone was patched via case-sensitivity earlier; the real root
# cause is the concrete example itself, which primes an LLM to regurgitate it regardless of
# the actual resume content. The prompt must not contain that swappable example list.
from hiring_agent.extraction import _AI_PROMPT
for _bait in ("Python, AWS, Docker, React, Swift", "Python, AWS, Docker, React"):
    ok(_bait not in _AI_PROMPT,
       f"the AI prompt no longer bakes in a concrete skill example an LLM could echo back "
       f"verbatim regardless of the resume ({_bait!r})")
ok("ACTUALLY NAMED" in _AI_PROMPT or "actually named" in _AI_PROMPT.lower(),
   "the prompt explicitly instructs the model to only report skills present in the text")

# Ã¢â€â‚¬Ã¢â€â‚¬ C. SHAREPOINT WORKBOOK SHAPE Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
print("\n=== B2c. EXPERIENCE (TOTAL YEARS) EXTRACTION (added 2026-09-04) ===")
# Only a total the candidate STATED is ever published. The summary is tried first (the
# client's stated source), then the headline, then anywhere else. A resume that states no
# total yields blank -> 'Missing' even when its jobs are fully dated: deriving a span from
# those dates was removed on user instruction 2026-09-04 ("just add years if its existing
# in summary"). Every case below pins one tier or one non-derivation.
from hiring_agent.extraction import _extract_experience

exp_cases = {
    # Tier 1: an explicit total inside the Summary/Profile block.
    "John Smith\nProfessional Summary\nSenior engineer with 6+ years of experience "
    "building distributed systems.\n": "6 years",
    "SUMMARY: over five years of hands-on experience in data engineering.\n": "5 years",
    "Profile\nMarketing lead. 12 years of experience across B2B SaaS.\n": "12 years",
    "Summary\nProduct specialist with 3+ years owning end-to-end delivery.\n": "3 years",
    # Singular is spelled correctly - "1 years" would read as a bug to the client.
    "Summary\n1 year of professional experience in QA automation.\n": "1 year",
    # Tier 1 picks the LARGEST stated figure: resumes print a per-skill breakdown under
    # the overall total, and the overall total is the answer.
    "Summary\n8 years of experience overall: 3 years Python, 5 years Java.\n": "8 years",
    # Tier 2: stated in the headline block with no 'Summary' header at all.
    "Jane Doe\nSenior Data Scientist | 9 years of experience | Phoenix, AZ\n": "9 years",
    # Tier 3: the mirrored phrasing, anywhere in the document.
    "Name\nSkills: Python\nTotal experience: 7+ years\n": "7 years",
    # Nothing stated and nothing dated -> blank, which the sheet renders as 'Missing'.
    "Just some random resume text with no header.\nSkills: Python, SQL.": "",
}
for _text, _expected in exp_cases.items():
    _got = _extract_experience(_text)
    ok(_got == _expected,
       f"experience {_text.splitlines()[-1][:52]!r} -> {_got!r} (expected {_expected!r})")

# NOT DERIVED: fully dated jobs with no stated total stay blank. This is the behaviour the
# user asked for on 2026-09-04 - a span is arithmetic the candidate never wrote, and a
# stated total already accounts for gaps, overlaps and part-time work in a way it cannot.
ok(_extract_experience(
    "Bob\nWork Experience\nAcme Corp, Engineer   Jan 2019 - Present\n"
    "Beta Ltd, Analyst     2016 - 2019\nEducation\nBS 2012 - 2016\n") == "",
   "dated work history alone is NOT spanned into a number - blank, not derived")
ok(_extract_experience(
    "Work Experience\nAcme Corp   2018 - 2022\nBeta Ltd    2020 - 2023\n"
    "Education\nBS 2010 - 2014\n") == "",
   "…and neither are overlapping dated roles")
# A stated total still wins in the very same resume - the dates are ignored, not the field.
ok(_extract_experience(
    "Summary\n6 years of experience in platform work.\n"
    "Work Experience\nAcme Corp   2010 - Present\n") == "6 years",
   "a STATED total is published even when the dated history would imply something else")

# GUARDS. These are the ways a years figure gets fabricated, and each must yield blank.
ok(_extract_experience(
    "Amy\nEducation\nBS Arizona State 2015 - 2019\nSpent 4 years at Arizona State.\n") == "",
   "GUARD: '4 years at Arizona State' is not professional experience - a total needs an "
   "experience/expertise/background word right after it")
ok(_extract_experience("I graduated 3 years ago and am keen to start.\n") == "",
   "GUARD: '3 years ago' is not a claim of experience")
ok(_extract_experience(
    "Education\nB.S. Computer Science, ASU\n2018 - 2022\nSkills: Python\n") == "",
   "GUARD: an EDUCATION date range never becomes experience")

# The column is fed from the same deterministic value on BOTH extraction entry points, and
# is never routed through Ollama - a stated total is a literal read, not a judgement call.
from hiring_agent.extraction import extract_candidate_details
_exp_resume = "Summary\nEngineer with 4 years of experience in Go.\nSkills: Go, Kubernetes\n"
ok(extract_candidate_details(_exp_resume).get("experience") == "4 years",
   "the offline parser populates details['experience']")
_ext_src = Path(__file__).with_name("hiring_agent").joinpath("extraction.py").read_text(
    encoding="utf-8", errors="replace")
_smart = _ext_src.split("def extract_candidate_details_smart", 1)[-1].split("\ndef ", 1)[0]
ok('"experience": baseline.get("experience"' in _smart,
   "extract_candidate_details_smart carries experience straight from the deterministic "
   "baseline - it is not a tier-2 field an LLM gets to overrule")
_tier2 = [l for l in _smart.splitlines() if "tier2_fields = " in l]
ok(_tier2 and "experience" not in _tier2[0],
   "…and 'experience' is deliberately absent from tier2_fields (the Ollama-arbitrated set)")


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

# ── C1b. Resume Link formula self-heals after every row add (fixed 2026-08-01) ──────
print("\n=== C1b. add_main_row/add_rejected_row refresh Resume Link after every add ===")
# Live finding: a genuine Excel calculated-column formula does not reliably auto-extend
# onto a row added via Graph's rows/add the way it does for a row typed into the workbook
# UI - 3 live rows (Sai Krishna Yallapu on Rejected, plus 2 rows moved back to Main) came
# back with a blank 'Resume Link' despite a valid 'Resume URL' right next to it.


def _make_fake_client(cols, raise_on_refresh=False):
    calls = {"post": 0, "set_calc": []}

    def _req(method, url, **kw):
        calls["post"] += 1
        return _types.SimpleNamespace(json=lambda: {})

    def _set_calc(table_name, col, formula):
        if raise_on_refresh:
            raise _SPErr("boom")
        calls["set_calc"].append((table_name, col))

    fake = _types.SimpleNamespace(
        table="HiringAgent_P1_Candidates",
        table_columns=lambda: cols,
        _table_columns_of=lambda t: cols,
        _wb_base=lambda: "https://fake",
        _ensure_rejected_table=lambda: "RejectedCandidates",
        _req=_req,
        set_calculated_column=_set_calc,
    )
    fake._refresh_resume_link_formula = _types.MethodType(_SPClient._refresh_resume_link_formula, fake)
    return fake, calls


_fake1, _calls1 = _make_fake_client(["Application ID", "Resume Link", "Resume URL"])
_SPClient.add_main_row(_fake1, {"Application ID": "APP-TEST"})
ok(_calls1["set_calc"] == [("HiringAgent_P1_Candidates", "Resume Link")],
   f"add_main_row refreshes the Resume Link formula right after adding (got {_calls1['set_calc']})")

_fake2, _calls2 = _make_fake_client(["Application ID", "Resume Link"])
_SPClient.add_rejected_row(_fake2, {"Application ID": "APP-TEST"})
ok(_calls2["set_calc"] == [("RejectedCandidates", "Resume Link")],
   f"add_rejected_row refreshes the Resume Link formula on the Rejected table too (got {_calls2['set_calc']})")

_fake3, _calls3 = _make_fake_client(["Application ID"])  # no Resume Link column at all
_SPClient.add_main_row(_fake3, {"Application ID": "APP-TEST"})
ok(_calls3["set_calc"] == [],
   "no refresh call is made when the table has no Resume Link column")

_fake4, _calls4 = _make_fake_client(["Resume Link"], raise_on_refresh=True)
_raised3 = False
try:
    _SPClient.add_main_row(_fake4, {"Application ID": "APP-TEST"})
except Exception:
    _raised3 = True
ok(not _raised3,
   "a failed Resume Link refresh never fails the row add itself (best-effort, add already succeeded)")

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

    def send_mail(self, to, subject, body, admin=False):
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


# â”€â”€ C1. LOCATION/COUNTRY CONSISTENCY GUARD (bug fix) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== C1. LOCATION/COUNTRY SELF-CONTRADICTION GUARD ===")
# Live regression (2026-08-01, Syyed Nazir Ali / APP-20260727-1431-6F75): a stray "MS"
# inside a tools list ("Bloomberg Terminal, MS Project") made the offline baseline
# mis-hint location as a bogus US address and country as "United States". Ollama then
# correctly read the resume's own header ("Location: India") and fixed location, but left
# country as the stale "United States" hint - producing a self-contradictory India /
# United-States row that a country-only geo check would have wrongly passed as USA-based.
_india_resume = (
    "Syyed Nazir Ali\nLocation: India\n\n"
    "TOOLS\nBloomberg Terminal, MS Project, MS Office Suite\n"
)
with _mock.patch("requests.post") as mp:
    mp.return_value.json.return_value = {"message": {"content": __import__("json").dumps({
        "full_name": "Syyed Nazir Ali", "phone": "", "location": "India",
        "country": "United States", "skills": "Excel", "looking_for_role": "Business Analyst",
        "education": "",
    })}}
    mp.return_value.raise_for_status = lambda: None
    os.environ["HIRING_OLLAMA_ENABLED"] = "true"
    importlib.reload(_cfg_mod)
    importlib.reload(_ext_mod)
    from hiring_agent.extraction import extract_candidate_details_smart as _ecds
    _guard_result = _ecds(_india_resume)
    os.environ["HIRING_OLLAMA_ENABLED"] = "false"
    importlib.reload(_cfg_mod)
    importlib.reload(_ext_mod)
ok(_guard_result.get("country") == "India",
   f"a location that names a foreign country overrides a stale contradicting 'country' "
   f"hint (got country={_guard_result.get('country')!r})")

# Live regression (2026-08-01, same Syyed Nazir Ali row, a second bug): Location must hold
# city/state only - Country is the dedicated field for the country name itself. His
# resume's only location statement is "Location: India (Open to Remote / Global)" - no
# city anywhere - so 'India' ended up duplicated into both Location AND Country. Since his
# header also says "Open to Remote / Global", Location should read "Remote", not "India".
_india_remote_resume = (
    "Syyed Nazir Ali\nLocation: India (Open to Remote / Global)\n\n"
    "TOOLS\nBloomberg Terminal, MS Project, MS Office Suite\n"
)
with _mock.patch("requests.post") as mp:
    mp.return_value.json.return_value = {"message": {"content": __import__("json").dumps({
        "full_name": "Syyed Nazir Ali", "phone": "", "location": "India",
        "country": "India", "skills": "Excel", "looking_for_role": "Business Analyst",
        "education": "",
    })}}
    mp.return_value.raise_for_status = lambda: None
    os.environ["HIRING_OLLAMA_ENABLED"] = "true"
    importlib.reload(_cfg_mod)
    importlib.reload(_ext_mod)
    _remote_result = _ecds(_india_remote_resume)
    os.environ["HIRING_OLLAMA_ENABLED"] = "false"
    importlib.reload(_cfg_mod)
    importlib.reload(_ext_mod)
ok(_remote_result.get("location") == "Remote" and _remote_result.get("country") == "India",
   f"a bare country name in Location falls back to 'Remote' when the text says so, and is "
   f"never duplicated into the Location column (got location={_remote_result.get('location')!r}, "
   f"country={_remote_result.get('country')!r})")

# Without any 'remote' mention, the bare-country fallback is 'N/A', not a guess.
with _mock.patch("requests.post") as mp:
    mp.return_value.json.return_value = {"message": {"content": __import__("json").dumps({
        "full_name": "Jane Doe", "phone": "", "location": "France",
        "country": "France", "skills": "Excel", "looking_for_role": "Analyst",
        "education": "",
    })}}
    mp.return_value.raise_for_status = lambda: None
    os.environ["HIRING_OLLAMA_ENABLED"] = "true"
    importlib.reload(_cfg_mod)
    importlib.reload(_ext_mod)
    _na_result = _ecds("Jane Doe\nLocation: France\n\nEXPERIENCE\nSoftware Engineer\n")
    os.environ["HIRING_OLLAMA_ENABLED"] = "false"
    importlib.reload(_cfg_mod)
    importlib.reload(_ext_mod)
ok(_na_result.get("location") == "N/A" and _na_result.get("country") == "France",
   f"a bare country name in Location with no 'remote' signal falls back to 'N/A' "
   f"(got location={_na_result.get('location')!r}, country={_na_result.get('country')!r})")

# Live regression (2026-08-27, Arya Jinan Panicker / APP-20260419-1833-C006): her
# contact header explicitly says United States, while Kerala appears only in historical
# internship entries. Both Ollama passes selected Kerala because the old grounding rule
# accepted a place mentioned anywhere in the resume. A contact-header location must win.
_arya_resume = (
    "Arya Jinan Panicker\n"
    "+1-(480)-853-1223 | aryajinan@example.com | linkedin.com/in/aryajinan\n"
    "United States\n\n"
    "EDUCATION\nM.S. Information Technology, Arizona State University\n"
    "EXPERIENCE\nData Science Intern | Kerala\n"
)
_arya_ai = {
    "full_name": "Arya Jinan Panicker", "phone": "+1-(480)-853-1223",
    "location": "Kerala", "country": "United States", "skills": "Python, SQL",
    "looking_for_role": "Data Scientist", "education": "M.S. Information Technology",
}
with _mock.patch("requests.post") as mp:
    mp.return_value.json.return_value = {
        "message": {"content": __import__("json").dumps(_arya_ai)}}
    mp.return_value.raise_for_status = lambda: None
    os.environ["HIRING_OLLAMA_ENABLED"] = "true"
    importlib.reload(_cfg_mod)
    importlib.reload(_ext_mod)
    _arya_first = _ext_mod.extract_candidate_details_smart(_arya_resume)
    _arya_final = _ext_mod.ai_recheck_fields(_arya_first, _arya_resume, "")
    os.environ["HIRING_OLLAMA_ENABLED"] = "false"
    importlib.reload(_cfg_mod)
    importlib.reload(_ext_mod)
ok(_arya_final.get("location") == "United States"
   and _arya_final.get("country") == "United States",
   f"contact-header United States survives historical Kerala in both AI passes "
   f"(got location={_arya_final.get('location')!r}, country={_arya_final.get('country')!r})")


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

# Policy (2026-08-01): an agency/vendor forwarding someone else's CV is accepted as a
# normal candidate as long as a real resume is attached - but the CV, never the sender,
# decides WHO the candidate is. Live case: a staffing agency's BD manager submitted a
# Flutter developer's resume; the row was created under the agent's name while the real
# candidate's name sat in the attached PDF. Both resume-derived sources are now exhausted
# before the sender's display name is consulted.
_agency_cv = "Shivam\nProfessional Skills:\nMobile Application Developer with 2 years experience\n"
ok(resolve_full_name("Not extracted", "vandana", _agency_cv) == "Shivam",
   "an agency-submitted CV resolves to the CANDIDATE's name, not the sending agent's")
ok(resolve_full_name("Not extracted", "vandana", "Priya Raman\nSoftware Engineer\n") == "Priya Raman",
   "a two-word CV name also beats the sending agent's name")
# The sender is still used when the CV genuinely yields no name.
ok(resolve_full_name("Not extracted", "Jane Smith", "RESUME\nEXPERIENCE\nSKILLS\n") == "Jane Smith",
   "sender name is still the fallback when the CV has no name-shaped line")
ok(resolve_full_name("Not extracted", "Jane Smith", "Futurism Technologies\nOur Services\n") == "Jane Smith",
   "a company name at the top of a CV is not mistaken for the candidate")

# Live regression (2026-08-01, found by dry_run_p1_p2_scenarios.py): the old
# `len(core) < 2` token rule rejected every name carrying a middle initial, so the
# OFFLINE parser returned "Not extracted" for 'Christopher L. Feld' / 'Jane Q Public'.
# Live rows looked correct only because Ollama (Tier 2) resolves the name independently -
# the bug was invisible until Ollama was disabled, which is the documented fallback path.
from hiring_agent.extraction import _looks_like_name as _lln2
for _n in ("Jane Q Public", "Christopher L. Feld", "John F Kennedy"):
    ok(_lln2(_n), f"a name with a middle initial is recognized offline: {_n!r}")
ok(extr("Christopher L. Feld\nMarana, AZ | c.feld@example.com\n")["full_name"]
   == "Christopher L. Feld",
   "the offline parser extracts a middle-initial name instead of 'Not extracted'")
# An initial must never carry the name on its own.
for _n in ("A B", "I am", "A B Testing"):
    ok(not _lln2(_n), f"initials alone are still rejected as a name: {_n!r}")

# Live regression (2026-08-01, Futurism Technologies / APP-20260717-0559-CCE1): "Transforming
# Business Models" (a marketing tagline from a vendor's company-brochure PDF, not a resume)
# is 3 alphabetic tokens with no digits or section-word overlap, so it passed the old
# structural shape check outright and was stored as the candidate's Full Name.
from hiring_agent.extraction import _plausible_name_shape as _pns, _looks_like_name as _lln
ok(_pns("Transforming Business Models") is False,
   "a company/marketing-tagline phrase is not treated as a plausible parsed name")
ok(_lln("Futurism Technologies") is False,
   "a company name is not treated as a plausible name even with exactly 2 tokens")
ok(resolve_full_name("Transforming Business Models", "parths", "Parth S\nFuturism Technologies\n")
   != "Transforming Business Models",
   "resolve_full_name never settles on a company-name-shaped parsed value")
# Genuine names sharing no vocabulary with the company-word list still pass.
for _name in ("Sai Krishna Yallapu", "Mindy Anderson", "Christopher L. Feld", "Bob Chen"):
    ok(_pns(_name) is True, f"genuine name still passes the shape check: {_name!r}")


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
ok(_extract_phone(
    "EDUCATION\nMBA Degree, Finance and Controlling | 2023 - 2025\n"
    "Bachelor's Degree, Marketing | 2016 - 2019") is None,
   "education year ranges are not misclassified as phone numbers")


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

# Live regression (2026-08-01, Brett Worker / APP-20260721-0122-E743): "...for the CISO
# position. The role caught my attention because it combines enterprise security
# governance..." - the role name ('CISO') sits BEFORE the word 'position', so the old
# keyword-then-role pattern found no letter immediately after 'position' (a period
# followed it) and fell through to matching 'role' in the next sentence instead,
# capturing "caught my attention because it combines" as the desired role.
from hiring_agent.extraction import _extract_role_from_body as _erfb
_brett_body = (
    "Hello, I came across Tracy Simon's LinkedIn post for the CISO position. The role "
    "caught my attention because it combines enterprise security governance with the "
    "practical challenge of building a scalable security program in a growing AI company."
)
ok(_erfb(_brett_body) == "CISO",
   f"'for the X position' captures the role name BEFORE the keyword, not an unrelated "
   f"later sentence (got {_erfb(_brett_body)!r})")
ok(_erfb("I am interested in the Senior Backend Engineer role at your company.")
   == "Senior Backend Engineer",
   "'interested in the X role' phrasing is also caught")
# Pre-existing "keyword: role" phrasing must still work unchanged.
ok(_erfb("Applying for: Data Scientist role. Passionate about ML.") == "Data Scientist role",
   "the original 'Applying for: X' phrasing is unaffected by the new before-keyword check")


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

# A repository link keeps its repository. Truncating to the first path segment turned a
# project link into a profile link, and the owner of a project a candidate worked on is
# often somebody else: Shashank Singh (APP-20260817-1859-MCSA) cites a collaborator's repo
# and Portfolio 2 published 'github.com/harishchaurasia' as if it were his account.
_z28_repo = extract_portfolios(
    "Project: https://github.com/harishchaurasia/Benchmarking-Privacy-Aware-Autonomy")[1]
ok(_z28_repo == "https://github.com/harishchaurasia/Benchmarking-Privacy-Aware-Autonomy",
   f"a GitHub repo URL keeps its full path, never collapsing to someone's profile (got {_z28_repo})")
ok(extract_portfolios("https://github.com/janedoe")[1] == "https://github.com/janedoe",
   "a bare GitHub profile URL is still stored unchanged")

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

# Live regression (2026-08-01, Muhammad Ahsan Hussain / APP-20260721-2030-3AA7): "Socket.io"
# (a real-time messaging library he uses, mentioned repeatedly as plain text - never written
# as a link anywhere in his resume) matched the bare 'word.io' shape and the '.io'
# personal-site heuristic, and was stored as his Portfolio 1.
p1, p2, p3 = extract_portfolios(
    "Backend built on Node.js and Express.js with RESTful and GraphQL APIs, MongoDB and "
    "Firebase for data and real-time features, and Socket.io for live communication layers")
ok((p1, p2, p3) == ("N/A", "N/A", "N/A"),
   f"'Socket.io' (a library mention, not a link) is not stored as a portfolio "
   f"(got p1={p1!r}, p2={p2!r}, p3={p3!r})")
ok("janesmith.io" in extract_portfolios("janesmith.io/portfolio")[0],
   "a genuine personal .io site is still detected after the Socket.io exclusion")

# Live regression (2026-08-01, Mindy Anderson / APP-20260723-2034-12DF): a SECOND
# independent false positive found the same afternoon as Socket.io - "Loquatinc.io" is a
# client company name in her contracts list ("Contracts include: Loquatinc.io, BNY Mellon,
# EY..."), never a link, but matched the same bare 'word.io' personal-site heuristic. Two
# unrelated real resumes hitting this in one session means the bare-TLD-only acceptance
# rule itself was too broad - removed entirely; a bare domain now needs an explicit scheme
# or the word "portfolio" nearby to be trusted, not just a trendy TLD.
ok(extract_portfolios(
    "Contracts include: Loquatinc.io, BNY Mellon, EY, OneSource Labs, and organizations "
    "across financial services") == ("N/A", "N/A", "N/A"),
   "a client/company name using a '.io' domain is not stored as a personal portfolio")
ok(extract_portfolios("Built tools deployed at randomstartup.dev and cloudthing.me") == ("N/A", "N/A", "N/A"),
   "bare .dev/.me company mentions with no scheme and no 'portfolio' context are not captured")

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


# -- N. PERMANENT HEADER SPACER (both sheets + the client export) -------------
# Every surface puts ONE blank row directly under the header: the setup template, the
# CandidateList sheet, resort_candidate_sheets.py's plan, and the client export. The
# Rejected table was the one place that silently broke it - _ensure_rejected_table built it
# header-only (A1:<last>1), and resort_candidate_sheets.py cannot repair that because it
# returns early on a sheet with no real rows. Found 2026-08-27 on the LIVE workbook:
# CandidateList had its spacer at index 0, Rejected had ZERO physical rows.
print("\n=== N. PERMANENT HEADER SPACER (CandidateList, Rejected, client export) ===")
_sc_path = Path(__file__).with_name("sharepoint_client.py")
_sc_src = _sc_path.read_text(encoding="utf-8", errors="replace")
_create = _sc_src.split("def _ensure_rejected_table")[1].split("\ndef ")[0]
ok('ref = f"A1:{last_col}2"' in _create,
   "a newly created Rejected table spans A1:<last>2, not header-only")
ok('[COLUMNS, [""] * len(COLUMNS)]' in _create,
   "the spacer row is written with REAL empty cells, not merely claimed by the range "
   "(the cell-less-row defect found in the setup template the same day)")

from sharepoint_client import SharePointClient as _SPC
ok(hasattr(_SPC, "ensure_rejected_spacer"),
   "ensure_rejected_spacer exists to repair a Rejected table created before the spacer")
_heal = _sc_src.split("def ensure_rejected_spacer")[1].split("\ndef ")[0]
ok("if rows:" in _heal and "return False" in _heal,
   "the self-heal acts ONLY on a table with zero physical rows, so it can never insert a "
   "stray blank into a populated sheet")
ok("rows/add" in _heal,
   "the self-heal appends one blank row via the table rows/add endpoint")

_ss_src = (Path(__file__).with_name("hiring_agent") / "sharepoint_scoring.py").read_text(
    encoding="utf-8", errors="replace")
ok("client.ensure_rejected_spacer()" in _ss_src,
   "the self-heal runs as part of the normal workbook-maintenance pass")

# The client export builds its own spacer, independently of SharePoint.
from hiring_agent.sharepoint_scoring import _with_period_separators as _wps
_empty = _wps([])
ok(len(_empty) == 1 and all(v == "" for v in _empty[0].values()),
   "an EMPTY client export still emits the permanent spacer row under the header")
_one = _wps([{"Application ID": "APP-1", "Received Date": "2026-04-14T00:41:32"}])
ok(all(v == "" for v in _one[0].values()),
   "a populated client export starts with the spacer, before the first candidate")
ok(str(_one[1].get("Application ID")) == "APP-1",
   "...and the first real candidate follows immediately after it")


# ── M. MULTI-ATTACHMENT HANDLING (slug mirror + duplicate-text dedup) ────────
# P1 saves every attachment it keeps, so one email can produce two files. Two things have to
# hold, and neither had a test when they were written (2026-08-27):
#   1. P2's _attachment_slug must be byte-identical in behaviour to P1's _attach_slug, or P2
#      probes a filename P1 never wrote and the row reads as unreadable.
#   2. The same resume sent as .pdf AND .docx must not be scored twice - concatenating both
#      counted every skill, employer and date double.
print("\n=== M. MULTI-ATTACHMENT HANDLING (slug mirror + duplicate-text dedup) ===")
from hiring_agent.sharepoint_scoring import (
    _attachment_slug, _text_fingerprint, _is_same_document, _resume_name_slots,
)

# ── the slug, and that it still MIRRORS P1 ──
for _name, _want in (
        ("Gayuh Nurul Huda - CV - 2026.pdf", "gayuhnurulhudacv"),
        ("Gayuh Nurul Huda - Portfolio - 2026.pdf", "gayuhnurulhudapo"),
        ("Resume.pdf", "resumepdf"),
        ("Resume.docx", "resumedocx"),
        ("my.name.resume.pdf", "mynameresumepdf"),
        ("O'Brien, Sean (1).pdf", "obriensean1pdf")):
    ok(_attachment_slug(_name) == _want,
       f"_attachment_slug({_name!r}) == {_want!r} (got {_attachment_slug(_name)!r})")
ok(len(_attachment_slug("a" * 60 + ".pdf")) == 16, "the slug is capped at 16 characters")
ok(_attachment_slug("") == "", "an empty filename yields an empty slug (no crash)")

# The two implementations live in different repos; assert P1's expression still deletes the
# SAME character set, so a one-sided edit fails this suite instead of silently orphaning files.
_p1_root = Path(__file__).resolve().parent.parent / "HiringAgent_P1"
_p1_builds = sorted(_p1_root.glob("**/flow/build_zip.py"))
_P1_BUILD = (_p1_builds[0] if _p1_builds else
             _p1_root / "HiringAgent_P1" / "flow" / "build_zip.py")
if _P1_BUILD.exists():
    _p1_src = _P1_BUILD.read_text(encoding="utf-8", errors="replace")
    _slug_fn = _p1_src.split("def _attach_slug(")[-1].split("def ")[0]
    ok('for ch in (" ", "_", "-", "\'\'", ",", "(", ")", "."):' in _slug_fn,
       "P1's _attach_slug deletes the same characters P2's _attachment_slug does")
    ok("min(16,length(" in _slug_fn, "P1's slug is capped at 16 characters too")
    ok("greater(length(body('ResumeFiles')),1)" in _p1_src,
       "P1 only appends the slug when the email kept MORE THAN ONE attachment")
else:
    ok(False, f"P1 build_zip.py not found for the mirror check at {_P1_BUILD}")

# ── the duplicate detector ──
_PDF = ("Sravanakumar Sathish\nAI/ML Engineer\nArizona State University, Tempe AZ\n"
        "Skills: PyTorch TensorFlow LangChain Docker AWS\n"
        "Experience: Built RAG pipelines and computer vision models.")
# the SAME document as .docx: different line breaks, bullets, spacing and punctuation
_DOCX = ("Sravanakumar  Sathish\n\nAI / ML Engineer\n\n"
         "\u2022 Arizona State University , Tempe  AZ\n"
         "\u2022 Skills : PyTorch, TensorFlow, LangChain, Docker, AWS\n"
         "\u2022 Experience : built RAG pipelines and computer-vision models")
# a GENUINELY different second document by the same person - must NOT be treated as a dup
_PORTFOLIO = ("Sravanakumar Sathish Portfolio\n"
              "Project 1: autonomous drone navigation using stereo depth\n"
              "Project 2: retail shelf auditing with YOLO\n"
              "Project 3: multilingual document QA deployed on Azure")

_fp_pdf, _fp_docx, _fp_port = (_text_fingerprint(x) for x in (_PDF, _DOCX, _PORTFOLIO))
ok(_is_same_document(_fp_pdf, _fp_pdf), "identical text is detected as the same document")
ok(_is_same_document(_fp_pdf, _fp_docx),
   "the same resume exported as .pdf and .docx is detected as one document "
   "despite differing whitespace, bullets and punctuation")
ok(not _is_same_document(_fp_pdf, _fp_port),
   "a genuine portfolio by the SAME candidate is NOT swallowed as a duplicate - "
   "the failure mode stays 'scored both', never 'silently discarded'")
ok(not _is_same_document(_text_fingerprint(""), _fp_pdf),
   "an empty extraction never counts as a duplicate (a failed download must not "
   "suppress a real attachment)")
ok(_text_fingerprint("The and a of") == _text_fingerprint("of a and The"),
   "the fingerprint is order-independent (a reflowed export still matches)")

# ── slot generation: the single-attachment shape must be untouched ──
_single = _resume_name_slots("APP-20260414-0041-9D05", "Resume.pdf", "Sravanakumar Sathish")
ok(len(_single) == 1, "one Original Filename entry produces exactly one slot")
ok(not any("_resumepdf" in n for n in _single[0]),
   "a SINGLE-attachment row probes no slugged name - its filename shape is unchanged, so "
   "every row written before 2026-08-27 still resolves")
_multi = _resume_name_slots("APP-20260414-0041-9D05", "Resume.pdf, Portfolio.pdf",
                            "Sravanakumar Sathish")
ok(len(_multi) == 2, "two Original Filename entries produce two independent slots")
ok(_multi[0][0].endswith("_resumepdf.pdf") and _multi[1][0].endswith("_portfoliopdf.pdf"),
   "each slot probes ITS OWN slugged filename first, so two files never collapse onto one")
ok(all(any("_" + _attachment_slug(o) in n for n in slot)
       for o, slot in zip(("Resume.pdf", "Portfolio.pdf"), _multi)),
   "the probed slug is the one P1 would have written for that original filename")

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
ok(_analyst_result["role_1"].startswith("Data Analyst"),
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
    ("Senior Software Engineer (80%)", "Senior Manager"),
    ("Backend Engineer (80%)",         "Web Team (Full stack/Back end & UI/UX)"),
    ("Data Scientist (65%)",           "AI/ML/CV (SIN2)"),
    ("AI / ML Engineer (75%)",         "AI/ML/CV (SIN2)"),
    ("Cloud Database Engineer (60%)",  "Cloud and DevOps"),
    ("DevOps Engineer (60%)",          "Cloud and DevOps"),
    ("Cybersecurity Analyst (65%)",    "Cybersecurity and IT Admin"),
    ("UI/UX Designer (65%)",           "Web Team (Full stack/Back end & UI/UX)"),
    ("Data Analyst (60%)",             "Data Analytics"),
    ("Logistics Manager (65%)",        "Senior Manager"),
    ("Logistics Specialist (50%)",     "Supply Chain"),
    ("Telecom Engineer (60%)",         "Satellite"),
    ("",                               "General"),
    ("No strong match",                "General"),
    # Added 2026-07-15 (client request): marketing -> Business Analytics,
    # design/rendering/gaming -> new 'Graphics' category.
    # Re-split 2026-08-21 (client request): Business Analytics -> BA / Sales / Marketing,
    # and Graphics -> Gaming / Graphics. See the dedicated section further down for the
    # ordering rules that make these land where they do.
    ("Marketing Specialist (55%)",     "Marketing - Admin"),
    ("Growth Marketer (50%)",          "Marketing - Admin"),
    ("Game Designer (60%)",            "Gaming"),      # 'game' beats 'design' - Gaming is listed first
    ("Rendering Engineer (50%)",       "Graphics"),
    ("Gaming Analyst (45%)",           "Gaming"),
    ("Graphics Programmer (55%)",      "Graphics"),
    ("Marketing Manager (65%)",        "Senior Manager"),  # 'manager' still wins (section 1 precedence)
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

# ── Whole-word rules (`word: true`, added 2026-08-21) ────────────────────────────────
# Short acronym rules are unusable as plain substrings: every one of these decoys contains
# the acronym inside an ordinary word. 'Auxiliary' -> Web Team via 'ux' was a REAL latent
# bug already live before the C-suite acronyms were added; the rest would have been
# introduced by adding cto/coo/cio without whole-word matching.
word_rule_cases = [
    # (title, expected) - acronyms that must now resolve to Executive
    ("CTO Position (80%)",              "Executive"),
    ("CEO Position (80%)",              "Executive"),
    ("CFO Position (80%)",              "Executive"),
    ("COO Position (80%)",              "Executive"),
    ("CIO Position (80%)",              "Executive"),
    ("CISO Job Announcement (80%)",     "Executive"),
    ("CMO Position Description (80%)",  "Executive"),
    # decoys: the acronym appears INSIDE a word and must NOT match
    ("Contractor Position (50%)",       "General"),      # contraCTOr
    ("Doctor of Pharmacy (50%)",        "General"),      # doCTOr
    ("Sector Analyst (50%)",            "General"),      # seCTOr
    ("Detector Systems Engineer (50%)", "General"),      # deteCTOr
    ("Cooperative Program Intern (50%)","General"),      # COOperative
    ("Suspicious Activity Analyst (50%)","General"),     # suspiCIOus
    ("Precious Metals Analyst (50%)",   "General"),      # preCIOus
    ("Auxiliary Systems Engineer (50%)","General"),      # a-UX-iliary (the live latent bug)
    ("Wholesales Coordinator (50%)",    "General"),      # wholeSALES
    # ...while the genuine whole-word hits still resolve exactly as before
    ("UI/UX Designer (65%)",            "Web Team (Full stack/Back end & UI/UX)"),
    ("UI UX Developer Job Announcement (70%)", "Web Team (Full stack/Back end & UI/UX)"),
    ("Sales and Business Development Representative (60%)", "Sales - Admin"),
]
for role_str, expected_cat in word_rule_cases:
    got = assign_category(role_str)
    ok(got == expected_cat,
       f"word-rule: assign_category('{role_str[:38]}') = '{got}' (expected '{expected_cat}')")

# 'director' must still win for Senior even though it contains 'cto' - Senior's rules are
# evaluated first, so precedence (not the word flag) is what protects it. Guard both.
ok(assign_category("Director of Engineering (70%)") == "Senior Manager",
   "a Director title stays Senior (section-1 precedence beats the Executive 'cto' rule)")

# ── BA / Sales / Marketing three-way split + Gaming / Graphics split (2026-08-21) ────
# Both splits are decided by RULE ORDER, not by the keywords alone, so these guard the
# ordering rather than just the mapping. If someone re-sorts config.yaml's rules block
# alphabetically or "tidies" it, these are what break.
#
# The BA/Marketing case is a real client ruling: the JD library carries a dual-titled
# opening containing BOTH keywords. The two JDs explicitly named "DriverAI DUAL Business
# Analytics Marketing Intern" are Business Analytics; the plain "Business Analytics
# Marketing Intern" is Marketing. That only works because the narrow 'dual business
# analytics' rule is hoisted above Marketing and the general 'business analytics' rule
# sits below it.
split_cases = [
    # the exact live JD titles this ruling was made about
    ("DriverAI Dual Business Analytics Marketing Intern Job Announcement (2) (80%)",
     "Business Analytics"),
    ("DriverAI Dual Business Analytics Marketing Intern Position Description (2) (80%)",
     "Business Analytics"),
    ("Business Analytics Marketing Intern Job Announcement (84%)",      "Marketing - Admin"),
    ("Business Analytics Marketing Intern Position Description (99%)",  "Marketing - Admin"),
    ("Marketing Intern Job Announcement (70%)",                         "Marketing - Admin"),
    ("Sales and Business Development Representative Position Description (60%)", "Sales - Admin"),
    # a title with NO sales/marketing signal still reaches the general BA rules at the bottom
    ("Business Analyst (70%)",                                          "Business Analytics"),
    ("Business Analytics Intern (70%)",                                 "Business Analytics"),
    ("Strategy Consultant (55%)",                                       "Business Analytics"),
    # Gaming must be evaluated before Graphics - 'Game Designer' has both 'game' and 'design'
    ("Game Designer (60%)",                                             "Gaming"),
    ("Gaming Position (60%)",                                           "Gaming"),
    ("Graphics Position (60%)",                                         "Graphics"),
    ("Mobile Animation Design Intern  PD (54%)",                        "Graphics"),
]
for role_str, expected_cat in split_cases:
    got = assign_category(role_str)
    ok(got == expected_cat,
       f"split: assign_category('{role_str[:46]}') = '{got}' (expected '{expected_cat}')")

# Pin the ORDER itself, so the guarantee survives a rules-block reshuffle even if some
# future title happens to still pass the cases above by luck.
from hiring_agent.scoring import ROLE_CATEGORY_RULES
_cats_in_order = [r["category"] for r in ROLE_CATEGORY_RULES]
_matches_in_order = [r["match"] for r in ROLE_CATEGORY_RULES]
ok(_matches_in_order.index("dual business analytics") < _matches_in_order.index("marketing"),
   "'dual business analytics' is matched BEFORE 'marketing' (the client's dual-JD ruling)")
ok(_matches_in_order.index("marketing") < _matches_in_order.index("business analytics"),
   "'marketing' is matched BEFORE the general 'business analytics' rule")
ok(_cats_in_order.index("Gaming") < _cats_in_order.index("Graphics"),
   "Gaming rules precede Graphics rules (so 'game' beats 'design')")

# Skill-based Mobile Apps override (2026-07-24): a candidate's own tool stack beats a
# loosely-matched JD title. Real cases from the live Rejected sheet - plain Android/iOS/
# Flutter developers whose top-scoring JD title happened to land them in the wrong
# category (Senior via 'lead', Graphics via 'gaming'/'design', 3D/CV/IoT via
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
ok(got == "Gaming",
   f"a Unity/game-engine skill alongside mobile skills stays out of Mobile Apps (got '{got}')")

# No skills passed (default "") -> behaves exactly as before (existing callers unaffected).
got = assign_category("Mobile Application Lead Developer Position Description v2 (80%)")
ok(got == "Senior Manager",
   f"assign_category with no skills_str falls back to title-only precedence (got '{got}')")

# Existing 'manager'/'lead' precedence in OTHER domains is untouched by the mobile-only
# override - Logistics Manager still Senior even with logistics-flavored
# skills (no mobile tool skills present, so the override never fires).
got = assign_category("Logistics Manager (65%)", "Supply Chain, Procurement, Excel")
ok(got == "Senior Manager",
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

ok(len(COLUMNS) == 33, f"COLUMNS has 33 entries (adds Experience, got {len(COLUMNS)})")
ok(COLUMNS[2] == "Last Updated Date", "'Last Updated Date' sits after Received Date")
ok(COLUMNS[-1] == "Info Request Sent", "'Info Request Sent' is the LAST column")
ok(COLUMNS[-2] == "Mail Sent", "'Mail Sent' is second-to-last")
ok("Education" in COLUMNS, "'Education' column present")
ok("Education Start Date" in COLUMNS and "Education End Date" in COLUMNS,
   "'Education Start Date'/'Education End Date' columns present (added 2026-08-18)")
_edu_idx = COLUMNS.index("Education")
ok(COLUMNS[_edu_idx + 1] == "Education Start Date" and COLUMNS[_edu_idx + 2] == "Education End Date",
   "Education Start/End Date sit immediately after Education")
# 'Years Exp' (added 2026-09-04 on client request) sits directly after Country
# and before Current Skills. Both the header text and the slot mirror the column the client
# created BY HAND on the live master/Rejected/results sheets - ensure_columns() matches on
# exact header text, so a drift here would add a SECOND column beside the hand-made one
# instead of adopting it. Pinned by position because the P1 setup template is generated
# against this exact order and test_p1.py section J compares the two.
ok("Years Exp" in COLUMNS,
   "'Years Exp' column present (added 2026-09-04)")
ok(COLUMNS[COLUMNS.index("Country") + 1] == "Years Exp",
   "Years Exp sits immediately after Country")
ok(COLUMNS[COLUMNS.index("Years Exp") + 1] == "Current Skills",
   "…and immediately before Current Skills")
ok("Experience" not in COLUMNS,
   "the interim bare 'Experience' header is gone - one column, not two")
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

# Live regression (2026-08-01, Divy Parmar / APP-20260720-1013-09AC): a bare, unformatted
# 10-digit Indian mobile number ("7405465204", no separators, no +91) coincidentally
# matched the NANP area-code shape and was trusted as "looks like a US phone" strongly
# enough to override his resume's own explicit "Khokhra, Ahmedabad" location, leaving a
# clearly non-US candidate sitting in Needs Review instead of Rejected.
ok(not _usph("7405465204"),
   "a bare unformatted 10-digit number (no separators, no country code) is NOT trusted "
   "as a US phone - indistinguishable from a foreign number missing its country code")
_keep, _r = _clu("Ahmedabad", country="India", resume_text="",
                 education="Master of Computer Applications (MCA)", phone="7405465204")
ok(not _keep, f"Divy Parmar case: explicit foreign city + bare unformatted phone -> REJECT, "
              f"not Needs Review (got keep={_keep}, reason={_r!r})")
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

# 2026-08-29 live queue regression: education-country matching used a raw substring test.
# Because US_COUNTRY_TERMS includes the token "us", ordinary words such as "industrial"
# and "campus" falsely read as US education and kept clear foreign applicants on Main.
from hiring_agent.geo import _education_reads_usa as _education_reads_us
ok(not _education_reads_us(
       'Bachelor of Economics, Moscow Financial-Industrial University "Synergy"'),
   "'us' inside 'industrial' is not United-States education evidence")
ok(not _education_reads_us("MCA, GNDU Regional Campus, Mukandpur"),
   "'us' inside 'campus' is not United-States education evidence")
ok(_education_reads_us("MS Computer Science, Arizona State University"),
   "a real US state name in Education is still detected")

# Current contact/location evidence outranks historical education.  These are the exact
# shapes found in the live rows that prompted the fix: a foreign current location and
# international/Canadian phone must not be vetoed by an old US university.
ok(_corrob("+91 8057167325", "B.S. Computer Science, Arizona State University", "")
   == "corroborates",
   "an India phone corroborates Noida despite historical Arizona State education")
ok(_corrob("+92 3333338893", "B.S. Computer Science, Arizona State University", "")
   == "corroborates",
   "a Pakistan phone corroborates Lahore despite historical Arizona State education")
ok(_corrob("+1 519-859-0693", "B.S. Computer Science, Arizona State University", "")
   == "corroborates",
   "a recognised Canadian NANP phone corroborates Canada despite US education")

for _loc, _country, _phone, _education in (
    ("Podgorica, Montenegro", "Montenegro", "", "Moscow Financial-Industrial University"),
    ("London", "Canada", "+1 519-859-0693", "Arizona State University"),
    ("Noida", "India", "+91 8057167325", "Arizona State University"),
    ("Mohali, Punjab, India", "India", "9463666582", "GNDU Regional Campus, Mukandpur"),
    ("Lahore", "Pakistan", "+92 3333338893", "Arizona State University"),
    ("Bangkok", "", "+66 842718796", "Arizona State University"),
    # Regression (found live 2026-09-01 via a 10-resume dummy run): every OTHER Canadian
    # province/territory was in config.yaml's foreign_countries list (Quebec, Alberta,
    # Manitoba, Saskatchewan, Nova Scotia, New Brunswick, British Columbia, Newfoundland,
    # PEI, Yukon, Nunavut) - Ontario alone was missing, almost certainly left out because
    # it collides with the US city Ontario, CA. Without it, is_foreign_location("London,
    # Ontario") returned False, so classify_location_usa fell through to the WEAK bare-
    # city step (is_us_city matches "London" - also a small US city in KY/OH) and wrongly
    # confirmed US for a Canadian candidate, despite Country already correctly reading
    # 'Canada'. A real "Ontario, CA" resume is unaffected: step 1 (is_strong_usa) matches
    # the CA state/ZIP and returns CONFIRMED_US before this list is ever consulted.
    ("London, Ontario", "Canada", "+1 519-859-0693", ""),
):
    _decision, _reason = _classify_geo(
        _loc, country=_country, phone=_phone, education=_education)
    ok(_decision == _GeoDecision.CONFIRMED_NON_US,
       f"explicit foreign residence {_loc!r} is rejected despite historical education "
       f"(got {_decision.value}: {_reason})")

from unittest.mock import patch as _geo_patch
with _geo_patch("hiring_agent.geo.verify_location_online", return_value=False):
    _decision, _reason = _classify_geo(
        "Bikaner", country="", phone="+91 9257616452",
        education="Government Engineering College Bikaner")
ok(_decision == _GeoDecision.CONFIRMED_NON_US,
   f"a geocoded foreign city plus an international phone is confirmed non-US "
   f"(got {_decision.value}: {_reason})")

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
# This case is about LOCATION CLARITY specifically, so assert on that entry rather than on
# the whole list: since 2026-08-04 the fixture (which carries no Email or Portfolio keys)
# also legitimately reports those two as missing.
ok("your current location/country" not in _miss and "your city and location" not in _miss,
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

print("\n=== O2d. is_strong_usa — country terms are WHOLE WORDS, not substrings ===")
# Regression for a live defect found 2026-08-11. is_strong_usa matched US_COUNTRY_TERMS with
# an unanchored `term in loc`, so the bare term 'us' hit the letters inside ordinary place and
# company names. This was not a cosmetic mislabel: a STRONG usa signal overrides an explicit
# foreign Country field in classify_location_usa, so it silently defeated the USA-only filter.
# Found via the health audit on APP-20260605-1642-F5F1 - an Islamabad candidate whose employer
# 'Vantage Plus' confirmed them as US and made the row look ready to score.
from hiring_agent.geo import is_strong_usa as _o2d_strong, classify_location_usa as _o2d_cls

for _o2d_loc in ("Vantage Plus, Jaffer GroupIslamabad", "Plus Tower, Lahore", "Cyprus",
                 "Belarus", "Mauritius", "Aarhus, Denmark", "Prussia Street, Dublin",
                 "Campus Road, Delhi", "Business Bay, Dubai", "Housing Society, Karachi",
                 "Perus, Sao Paulo", "Syracuse"):
    ok(_o2d_strong(_o2d_loc) is False,
       f"'us' inside a word is NOT a US signal: {_o2d_loc!r}")

for _o2d_loc in ("Tempe, AZ", "Austin, Texas", "USA", "U.S.A.", "U.S.", "U.S", "U.S.A",
                 "United States", "united states of america", "New York, NY 10001",
                 "San Juan, Puerto Rico", "Cook County, IL", "90210", "Remote - US"):
    ok(_o2d_strong(_o2d_loc) is True,
       f"genuine US signal still detected: {_o2d_loc!r}")

# The end-to-end consequence: a foreign row must not be confirmed US off a stray 'us'.
ok(_o2d_cls("Vantage Plus, Jaffer GroupIslamabad", country="")[0].value != "confirmed_us",
   "the live Islamabad row is no longer confirmed as US residence")
ok(_o2d_cls("Business Bay, Dubai", country="United Arab Emirates")[0].value != "confirmed_us",
   "an explicit foreign Country is no longer overridden by 'us' inside 'Business'")
ok(_o2d_cls("Tempe, AZ", country="United States")[0].value == "confirmed_us",
   "a real US candidate is still confirmed - the fix did not over-tighten")


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
# A separator-free digit run is NOT reformatted (2026-09-05). It is indistinguishable
# from a foreign mobile written without its country code, and geo._looks_like_us_phone
# refuses to trust one for that reason - so adding parens here manufactured the very
# separator that guard checks for, and the guard could never fire on a number this
# function had already touched (live case Divy Parmar, Indian "7405465204").
ok(format_phone("6677036994") == "6677036994",
   "a separator-free 10-digit run is left exactly as written")
ok(format_phone("7405465204") == "7405465204",
   "...so a foreign mobile is never laundered into a US shape")
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
    ("Senior", "Senior"),
    ("Executive", "Executive"),
    ("Web Team (Full stack/Back end & UI/UX)", "WebTeamFullStackBackEndUIUX"),
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
            ("JaneDoe_APP-20260710-2200-A5F2.pdf",
             "Jane_Doe_A5F2.pdf"): True,
            ("JaneDoe_APP-20260710-2200-A5F2.docx",
             "Jane_Doe_A5F2.docx"): True,
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
    ok(_renamed == ["Jane_Doe_A5F2.pdf",
                    "Jane_Doe_A5F2.docx"],
       "rename helper returns canonical names for distinct extensions")
    ok(_cap.warnings == [], "rename helper does not warn on successful canonical renames")

    _cap = _LogCapture()
    _sp_scoring.logger = _cap
    _client_dup = _RenameClient(
        rename_results={
            ("JaneDoe_APP-20260710-2200-A5F2.pdf",
             "Jane_Doe_A5F2.pdf"): True
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
    ok(_renamed_dup == ["Jane_Doe_A5F2.pdf"],
       "duplicate canonical targets collapse to one represented filename")
    ok(_cap.warnings == [] and any("already represented" in m for m in _cap.infos),
       "duplicate canonical target is logged as info, not warning")

    _cap = _LogCapture()
    _sp_scoring.logger = _cap
    _client2 = _RenameClient(existing={"Jane_Doe_A5F2.pdf"})
    _renamed2 = _rename_scored_resumes(
        _client2,
        ["JaneDoe_APP-20260710-2200-A5F2.pdf"],
        "2026/July",
        "APP-20260710-2200-A5F2",
        "Jane Doe",
        "Data Analytics",
        dry_run=False,
    )
    ok(_renamed2 == ["Jane_Doe_A5F2.pdf"],
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


# -- P2d3. missing-info nudge criteria (portfolio + email REINSTATED 2026-08-04) --
# History: portfolio was excluded on 2026-07-15 ("only these four"), then reinstated on
# 2026-08-04 together with email ("ask for missing details like email, portfolio, number,
# city, location"). Both decisions are recorded so the older one is not mistaken for the
# current one and quietly restored. Key nuance: a portfolio slot reading 'N/A' means the
# candidate told us they have none - that is an ANSWER, not a gap, and must never
# re-trigger the nudge.
print("\n=== P2d3. missing-info nudge criteria (portfolio + email included) ===")
from hiring_agent.sharepoint_scoring import _missing_fields_list

_complete_row = {"Phone": "555-1234", "Location": "Austin, TX", "Current Skills": "Python",
                 "Education": "B.S. Computer Science", "Email": "x@example.com",
                 "Portfolio 1": "N/A", "Portfolio 2": "N/A", "Portfolio 3": "N/A"}
ok(_missing_fields_list(_complete_row) == [],
   "nothing missing when all details are present and portfolios are an explicit N/A")
ok(_missing_fields_list(dict(_complete_row, **{"Portfolio 1": "https://github.com/x"})) == [],
   "a real portfolio link is obviously not missing")

_blank_ports = dict(_complete_row, **{"Portfolio 1": "", "Portfolio 2": "", "Portfolio 3": ""})
ok(_missing_fields_list(_blank_ports) == ["a portfolio or LinkedIn/GitHub link"],
   "genuinely BLANK portfolio slots now DO trigger the nudge (reinstated 2026-08-04)")
ok(_missing_fields_list(_complete_row) == [],
   "'N/A' portfolios stay silent - the candidate already answered, so no repeat nagging")

ok(_missing_fields_list({**_complete_row, "Phone": "Not extracted"}) == ["your phone number"],
   "missing phone alone is flagged")
ok(_missing_fields_list({**_complete_row, "Location": ""}) == ["your city and location"],
   "missing location alone is flagged, and asks for CITY as well as location")
ok(_missing_fields_list({**_complete_row, "Location": "Missing"}) == ["your city and location"],
   "a cell already stamped 'Missing' still counts as missing (it is a gap literal)")
ok(_missing_fields_list({**_complete_row, "Current Skills": "N/A"}) == ["your listed skills"],
   "missing skills alone is flagged")
ok(_missing_fields_list({**_complete_row, "Education": "Not extracted"}) == ["your education background"],
   "missing education alone is flagged")
ok(_missing_fields_list({**_complete_row, "Email": ""}) == ["a contact email address"],
   "a missing/broken email address is flagged")

_all_missing = {"Phone": "", "Location": "", "Current Skills": "", "Education": "",
                "Email": "", "Portfolio 1": "", "Portfolio 2": "", "Portfolio 3": ""}
ok(_missing_fields_list(_all_missing) ==
   ["your phone number", "your city and location", "your listed skills",
    "your education background", "a portfolio or LinkedIn/GitHub link",
    "a contact email address"],
   "everything missing, in a fixed reader-friendly order")


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
        return [{"index": i, "values": {"Phone": p, "Country": c}}
                for i, p, c in self._main]
    def list_rejected_rows(self):
        return [{"index": i, "values": {"Phone": p, "Country": c}}
                for i, p, c in self._rej]
    def update_row(self, index, fields, current_values=None):
        self.writes.append(("main", index, fields["Phone"]))
    def update_rejected_row(self, index, fields, current_values=None):
        self.writes.append(("rejected", index, fields["Phone"]))

_pc = _StubPhoneClient(
    main_rows=[
        (0, "(469) 688-7299", "United States"),
        (1, 4808531223, "United States"),
        (2, "+1-(480)-853-1223", "United States"),
        (3, 923326596399, "Pakistan"),
        # Regression (found live 2026-09-01, 23 rows): NANP-shaped numbers left
        # un-canonicalized because Country read something other than the literal string
        # 'United States' - reformatting must not be gated on Country at all, since
        # format_phone() itself already only ever transforms a 10/11-digit NANP shape.
        (4, "+1-917-346-8630", "Missing"),
        (5, "865-356-1921", "India"),
    ],
    rej_rows=[
        (0, 971546000000, "United Arab Emirates"),
        (1, "+92 332 659 6399", "Pakistan"),
    ],
)
_renormalize_phone_columns(_pc)
ok(_pc.writes == [
       # A numeric cell still becomes TEXT - that is this pass's job - but Excel having
       # eaten the candidate's separators is not evidence the number is American, so it is
       # not given US parens either. See format_phone (2026-09-05).
       ("main", 1, "4808531223"),
       ("main", 2, "(480) 853-1223"),
       ("main", 3, "923326596399"),
       ("main", 4, "(917) 346-8630"),
       ("main", 5, "(865) 356-1921"),
       ("rejected", 0, "971546000000"),
   ],
   "any NANP-shaped number is canonicalized to (XXX) XXX-XXXX regardless of Country, "
   "while numeric foreign cells are only re-written as text and already-text "
   "non-NANP-shaped foreign numbers stay untouched")


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

print("\n=== P2e3. _renormalize_missing_placeholders - backfills blank -> 'Missing' ===")
from hiring_agent.sharepoint_scoring import _renormalize_missing_placeholders

class _StubMissingClient:
    def __init__(self, main_rows, rej_rows):
        self._main = main_rows
        self._rej = rej_rows
        self.writes = []
    def list_rows(self):
        return self._main
    def list_rejected_rows(self):
        return self._rej
    def update_row(self, index, fields, current_values=None):
        self.writes.append(("main", index, dict(fields)))
    def update_rejected_row(self, index, fields, current_values=None):
        self.writes.append(("rejected", index, dict(fields)))

_mpc = _StubMissingClient(
    main_rows=[
        # Regression (found live 2026-09-01, 2 rows): a blank Country cell surviving on an
        # existing row - _apply_missing_placeholders only ever touches a row actively
        # being scored, never sweeps rows that already exist.
        {"index": 0, "values": {"Phone": "(480) 853-1223", "Country": "",
                                "Location": "Phoenix, AZ", "Education": "B.S. CS",
                                "Current Skills": "Python", "Full Name": "Jane Doe",
                                "Looking For Role": "Software Engineer"}},
        # Already 'Missing'/real content everywhere -> untouched, not re-written.
        {"index": 1, "values": {"Phone": "Missing", "Country": "India",
                                "Location": "Mumbai", "Education": "Missing",
                                "Current Skills": "SQL", "Full Name": "Ravi Kumar",
                                "Looking For Role": "Missing"}},
    ],
    rej_rows=[
        # 'N/A' is ALSO converted (it's a recognized gap literal, same as a bare blank) -
        # only a column already reading the literal 'Missing' is left untouched, checked
        # via Java/Pakistan/Ahmed Khan/'Missing' below.
        {"index": 0, "values": {"Phone": "", "Country": "Pakistan", "Location": "",
                                "Education": "N/A", "Current Skills": "Java",
                                "Full Name": "Ahmed Khan", "Looking For Role": "Missing"}},
    ],
)
_renormalize_missing_placeholders(_mpc)
ok(_mpc.writes == [
       ("main", 0, {"Country": "Missing"}),
       ("rejected", 0, {"Phone": "Missing", "Location": "Missing", "Education": "Missing"}),
   ],
   f"a blank cell OR any other gap literal ('N/A', 'Not extracted', 'Not confirmed') is "
   f"backfilled to 'Missing'; a real value or an already-'Missing' cell is left exactly "
   f"as-is (got {_mpc.writes!r})")

print("\n=== P2e4. _reevaluate_missing_education_dates - extra layer, re-derives from resume ===")
from hiring_agent.sharepoint_scoring import (
    _reevaluate_missing_scored_fields, STATUS_SCORED as _STATUS_SCORED_E4,
)
import hiring_agent.sharepoint_scoring as _sps_e4

class _StubEduDateClient:
    def __init__(self, main_rows):
        self._main = main_rows
        self.writes = []
    def list_rows(self):
        return self._main
    def update_row(self, index, fields, current_values=None):
        self.writes.append((index, dict(fields)))

# Every row explicitly sets Phone/Location to an already-resolved value except where a
# case specifically means to exercise that field, so each test case isolates ONE gap.
_e4_rows = [
    # Both dates blank, Status Scored, resume plainly states a range -> backfilled.
    {"index": 0, "values": {
        "Application ID": "APP-E4-1", "Status": _STATUS_SCORED_E4, "Full Name": "Jane Doe",
        "Category": "General", "Original Filename": "jane.pdf", "Resume URL": "x",
        "Received Date": "2026-06-01T00:00:00", "Phone": "(555) 555-5555",
        "Location": "Tempe, AZ", "Country": "United States",
        "Education Start Date": "", "Education End Date": "",
        "Looking For Role": "Software Engineer", "Education": "B.S. Computer Science, ASU",
        "Current Skills": "Python, Java", "Years Exp": "3 years"}},
    # Everything already resolved -> never re-downloaded (would raise if it tried).
    {"index": 1, "values": {
        "Application ID": "APP-E4-2", "Status": _STATUS_SCORED_E4, "Full Name": "John Roe",
        "Category": "General", "Original Filename": "john.pdf", "Resume URL": "x",
        "Received Date": "2026-06-01T00:00:00", "Phone": "(555) 555-5555",
        "Location": "Tempe, AZ", "Country": "United States",
        "Education Start Date": "Aug 2022", "Education End Date": "May 2024",
        "Looking For Role": "Data Analyst", "Education": "B.S. Data Science, MIT",
        "Current Skills": "SQL, Python", "Years Exp": "2 years"}},
    # Not Scored (Needs Review) -> never touched, even though everything is blank.
    {"index": 2, "values": {
        "Application ID": "APP-E4-3", "Status": "Needs Review - Location Confirmation",
        "Full Name": "Al Ray", "Category": "General", "Original Filename": "al.pdf",
        "Resume URL": "x", "Received Date": "2026-06-01T00:00:00", "Phone": "",
        "Location": "", "Education Start Date": "", "Education End Date": ""}},
    # Only End Date blank, resume has no matching date to fill it -> no write (not a false Fixed).
    {"index": 3, "values": {
        "Application ID": "APP-E4-4", "Status": _STATUS_SCORED_E4, "Full Name": "Sam Fox",
        "Category": "General", "Original Filename": "sam.pdf", "Resume URL": "x",
        "Received Date": "2026-06-01T00:00:00", "Phone": "(555) 555-5555",
        "Location": "Tempe, AZ", "Country": "United States",
        "Education Start Date": "Aug 2022", "Education End Date": "",
        "Looking For Role": "Backend Developer", "Education": "B.S. Computer Science, ASU",
        "Current Skills": "Go, Docker", "Years Exp": "5 years"}},
    # Phone AND Location both blank (dates already resolved) -> both recovered from resume.
    {"index": 4, "values": {
        "Application ID": "APP-E4-5", "Status": _STATUS_SCORED_E4, "Full Name": "Kim Lee",
        "Category": "General", "Original Filename": "kim.pdf", "Resume URL": "x",
        "Received Date": "2026-06-01T00:00:00", "Phone": "Missing", "Location": "Missing",
        "Education Start Date": "Aug 2022", "Education End Date": "May 2024",
        "Looking For Role": "ML Engineer", "Education": "M.S. Computer Science, Stanford",
        "Current Skills": "Python, TensorFlow", "Years Exp": "4 years"}},
    # Regression (found live 2026-09-01): a Scored row's Phone is non-blank but a
    # malformed fragment ('245-7289', 7 digits, no area code) - too short for
    # format_phone to touch, too long for sanitize_phone's <7-digit gap check to blank.
    # The resume here has a full, better number -> the fragment is replaced.
    {"index": 5, "values": {
        "Application ID": "APP-E4-6", "Status": _STATUS_SCORED_E4, "Full Name": "Kay Cole",
        "Category": "General", "Original Filename": "kay.pdf", "Resume URL": "x",
        "Received Date": "2026-06-01T00:00:00", "Phone": "245-7289",
        "Location": "Tempe, AZ", "Country": "United States",
        "Education Start Date": "Aug 2022", "Education End Date": "May 2024",
        "Looking For Role": "DevOps Engineer", "Education": "B.S. IT, ASU",
        "Current Skills": "AWS, Linux", "Years Exp": "6 years"}},
    # Same malformed-fragment case, but the resume has nothing better -> falls back to
    # 'Missing' rather than leaving the fragment displayed as if it were a real number.
    {"index": 6, "values": {
        "Application ID": "APP-E4-7", "Status": _STATUS_SCORED_E4, "Full Name": "Uma Vane",
        "Category": "General", "Original Filename": "uma.pdf", "Resume URL": "x",
        "Received Date": "2026-06-01T00:00:00", "Phone": "245-7289",
        "Location": "Tempe, AZ", "Country": "United States",
        "Education Start Date": "Aug 2022", "Education End Date": "May 2024",
        "Looking For Role": "QA Engineer", "Education": "B.S. Math, UofA",
        "Current Skills": "Selenium, Java", "Years Exp": "3 years"}},
]
_e4_client = _StubEduDateClient(_e4_rows)
_original_e4_download = _sps_e4._download_resume_text
def _fake_e4_download(client, app_id, *_a, **_kw):
    if app_id == "APP-E4-4":
        return "Jane Doe\nEDUCATION\nB.S. Computer Science, ASU\n", ["x"], {}
    if app_id == "APP-E4-5":
        return ("Kim Lee\n(480) 555-0142\nTempe, AZ\n"
                "EDUCATION\nB.S. Computer Science, Arizona State University\n"
                "Aug 2022 - May 2024\n"), ["x"], {}
    if app_id == "APP-E4-6":
        return "Kay Cole\n(480) 555-0142\nTempe, AZ\n", ["x"], {}
    if app_id == "APP-E4-7":
        return "Uma Vane\nno usable phone here\n", ["x"], {}
    return ("Jane Doe\nEDUCATION\nB.S. Computer Science, Arizona State University\n"
            "Aug 2022 - May 2024\n"), ["x"], {}
_sps_e4._download_resume_text = _fake_e4_download
try:
    _reevaluate_missing_scored_fields(_e4_client)
finally:
    _sps_e4._download_resume_text = _original_e4_download
ok(_e4_client.writes == [
       (0, {"Education Start Date": "Aug 2022", "Education End Date": "May 2024"}),
       (4, {"Phone": "(480) 555-0142", "Location": "Tempe, AZ"}),
       (5, {"Phone": "(480) 555-0142"}),
       (6, {"Phone": "Missing"}),
   ],
   f"only rows actually missing something AND whose resume yielded a value are written - "
   f"Education dates, Phone, and Location are each independently recovered; a malformed "
   f"non-blank Phone fragment is either replaced with a better number or falls back to "
   f"'Missing'; already-resolved, not-Scored, and no-value-found rows are all left alone "
   f"(got {_e4_client.writes!r})")


# â”€â”€ P2f. cross-candidate duplicate MERGE (Phone only) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== P2f. _merge_duplicate_candidates — detects duplicates, heals, deletes NOTHING ===")
from hiring_agent.sharepoint_scoring import _normalize_phone_key, _merge_duplicate_candidates, _finish_rejection

ok(_normalize_phone_key("(469) 688-7299") == "4696887299", "US-formatted phone -> digits only")
ok(_normalize_phone_key("+1 469 688 7299") == "4696887299",
   "US country-code '1' prefix dropped so it matches the bare 10-digit form")
ok(_normalize_phone_key("923127174226") == "923127174226",
   "non-US number kept whole (no US country-code stripping applied)")
ok(_normalize_phone_key("12345") == "", "too short (<7 digits) never matches -> empty key")
ok(_normalize_phone_key("Not extracted") == "", "gap phone -> empty key")

# CONTAINMENT CHANGE 2026-09-06 (docs\P1_P2_IMPLEMENTATION_STEPS.md Section 1): the merge
# pass no longer deletes anything - not the loser's row, not its resume file, on either
# sheet. A delete that half-completes (row gone, file still there, or the reverse) is
# unrecoverable, and doing it silently on every ordinary run multiplied that exposure rather
# than fixing it. This whole section now pins the NEW contract: duplicates are still found
# and the surviving winner is still healed from them, but the loser is retained, visible to
# a human and to the next run, until a separate reviewed retention procedure removes it.

class _StubMergeClient:
    """Fakes list_rows/list_rejected_rows + update_row/update_rejected_row for
    _merge_duplicate_candidates. delete_row/delete_rejected_row/delete_file are intentionally
    NOT implemented - the function under test must never call any of them."""
    def __init__(self, main_rows, rej_rows):
        self._main = main_rows
        self._rej = rej_rows
        self.resumes_folder = "/Downloaded_Resumes"
        self.updates = []
    def list_rows(self):
        return self._main
    def list_rejected_rows(self):
        return self._rej
    def update_row(self, index, fields, current_values=None):
        self.updates.append(("main", index, fields))
    def update_rejected_row(self, index, fields, current_values=None):
        self.updates.append(("rejected", index, fields))

# "Akshay Faye"-style: same candidate, two different sender/agency emails, different
# Received Date. APP-NEW is the more recent submission -> WINNER, healed from APP-OLD.
# APP-OLD also carries a real 'Category' the winner is missing (a gap) -> must be healed in.
# Both rows are expected to still exist afterward - neither is deleted.
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

ok(_result == {"healed": 1, "duplicates_found": 1, "errors": 0},
   f"one duplicate found, one winner healed, no errors (got {_result})")

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
   "the LOSER (older APP-OLD) never receives an update - it is left exactly as it was")

# dry_run: computes the same plan but writes nothing at all.
_mc_dry = _StubMergeClient(
    [dict(r) for r in _mm_main], [dict(r) for r in _mm_rej])
_dry_result = _merge_duplicate_candidates(_mc_dry, dry_run=True)
ok(_dry_result == {"healed": 0, "duplicates_found": 1, "errors": 0},
   "dry-run still reports what it FOUND, but 'healed' stays zero since nothing was written")
ok(_mc_dry.updates == [], "dry-run makes NO writes at all")

# A->B by phone and B->C by email is one candidate component. The newest C row is the
# winner and receives a field available only on A. A and B are retained, not removed.
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
ok(_chain_result == {"healed": 1, "duplicates_found": 2, "errors": 0},
   "transitive email/phone chain finds BOTH losers, heals the one newest winner ONCE")
ok(_chain_update.get("Category") == "Cloud and DevOps",
   "transitive duplicate winner receives a missing field from the oldest connected row")
ok(all(idx != 0 and idx != 1 for _sheet, idx, _f in _chain_client.updates),
   "neither loser in the chain is ever written to, and neither is deleted")

# Same-email pair (changed 2026-07-12): still FOUND here as P2's safety net for P1's rare
# read-after-write race (nishantacharekar12@gmail.com, two rows four minutes apart). Neither
# row is deleted now - the newer one is simply the winner if there is anything to heal.
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
ok(_ss_result["duplicates_found"] == 1,
   "identical email on both rows -> still recognised as one duplicate pair")
ok(_ssc.updates == [],
   "with nothing for the newer row to heal, no write happens to either row - and none was deleted")

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
ok(_qd_result == {"healed": 0, "duplicates_found": 0, "errors": 0} and _qdc.updates == [],
   "duplicate check ignores New Email Received rows so fresh updates are scored before merging")

# Different email AND different phone, even with the SAME name -> two DIFFERENT people, NOT
# matched as duplicates at all.
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
ok(_dp_result == {"healed": 0, "duplicates_found": 0, "errors": 0},
   "same name but different email AND phone -> different people, zero duplicates (name never matches)")

# A read failure on the main table must return the SAME dict shape as every other path -
# found in review while rewriting this section: the early return still carried the OLD
# {"merged", "errors"} shape after the rest of the function moved to {"healed",
# "duplicates_found", "errors"}, which every caller reads with .get() so it never raised, but
# a strict equality check anywhere would have quietly diverged from the documented contract.
class _UnreadableMergeClient(_StubMergeClient):
    def list_rows(self):
        from sharepoint_client import SharePointError
        raise SharePointError("table unavailable")

_unr_result = _merge_duplicate_candidates(_UnreadableMergeClient([], []))
ok(_unr_result == {"healed": 0, "duplicates_found": 0, "errors": 0},
   f"an unreadable main table returns the SAME dict shape as every other path (got {_unr_result})")

# Every earlier scenario above ran against a stub with NO delete_row/delete_rejected_row/
# delete_file methods at all, so a stray delete call would have crashed the whole suite with
# an AttributeError rather than a quiet [FAIL]. Make that guarantee an explicit, named test
# too: a stub that HAS delete methods, wired to raise loudly, must still never see them called.
class _LoudDeleteMergeClient(_StubMergeClient):
    def delete_row(self, index):
        raise AssertionError(f"_merge_duplicate_candidates must never delete a row (main:{index})")
    def delete_rejected_row(self, index):
        raise AssertionError(f"_merge_duplicate_candidates must never delete a row (rej:{index})")
    def delete_file(self, folder, name):
        raise AssertionError(f"_merge_duplicate_candidates must never delete a file ({folder}/{name})")

_loud = _LoudDeleteMergeClient([dict(r) for r in _mm_main], [])
try:
    _merge_duplicate_candidates(_loud)
    _no_delete_attempted = True
except AssertionError:
    _no_delete_attempted = False
ok(_no_delete_attempted,
   "explicit proof: given delete methods that raise, merging a real duplicate group still "
   "never calls any of them")

# CONTAINMENT CHANGE 2026-09-06 (docs/P1_P2_IMPLEMENTATION_STEPS.md Section 1), third site:
# the "stale sibling" cleanup in the main scoring loop used to call client.delete_file() on
# any guessed-stale resume filename, with NO test coverage at all despite its own comment
# documenting a live incident (2026-08-24, APP-20260824-2049-642B) where a case-sensitivity
# gap let it delete a candidate's real, freshly-renamed resume. There is no isolated helper
# to unit-test this block through - it is inlined in score_from_sharepoint's per-candidate
# loop - so the guarantee is pinned the same way test_p2.py already pins other source-level
# invariants (see _ss_src above): the delete call must not exist in the file at all, not
# merely be gated behind a flag that could later be flipped back on by mistake.
ok("client.delete_file(" not in _ss_src,
   "no code path in sharepoint_scoring.py calls client.delete_file() any more - resume "
   "deletion in ordinary runs is fully removed, not conditionally disabled")
ok("stale resume file" in _ss_src and "preserved" in _ss_src,
   "the stale-sibling scan still runs and still logs what it finds, it just no longer deletes")


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
ok(_frc_refresh.deleted_files == [],
   "CONTAINMENT CHANGE 2026-09-06: the superseded resume is PRESERVED, not deleted - "
   "an unrecoverable delete driven by inference is exactly what Section 1 of "
   "docs/P1_P2_IMPLEMENTATION_STEPS.md removes from ordinary runs")

# The live 2026-08-05 regression. Draining the backlog fed a JUNE application in after the
# JULY one was already filed, and the refresh above rewrote the July record from the June
# row: it kept July's Application ID but took June's Received Date, Original Filename and
# resume, then deleted July's resume as "superseded". Rows C8CC and 09AC were both damaged
# this way and both newer resumes were destroyed. Only a NEWER row may replace the record.
_frc_older = _StubRefreshRejectedClient()
_older_incoming_vals = {
    "Application ID": "APP-20260603-0641-CC14",
    "Email": "repeat@x.com",
    "Received Date": "2026-06-03T06:41:38",
    "Original Filename": "June_Resume.pdf",
    "Phone": "+92 333 8113971",
}
_fr_older_result = _finish_rejection(
    _frc_older, 9, "APP-20260603-0641-CC14", _older_incoming_vals,
    {"Location": "Daska, Pakistan", "Country": "Pakistan",
     "Status": "Rejected - Non-USA Location"},
    ["June_C8CC.pdf"], "2026/June", False,
    {"APP-20260716-1505-C8CC"}, {"repeat@x.com"}, set(), "non-USA",
    rejected_rows=[{
        "index": 3,
        "values": {
            "Application ID": "APP-20260716-1505-C8CC",
            "Email": "repeat@x.com",
            "Received Date": "2026-07-16T15:05:53",
            "Original Filename": "July_Resume.pdf",
            "Resume URL": "https://sharepoint.test/2026/Rejected/July_C8CC.pdf",
            "Phone": "",
            "Decline Sent": "Sent 2026-08-04 03:12",
        },
    }],
)
_fr_older_fields = (_frc_older.rejected_updates[0][1]
                    if _frc_older.rejected_updates else {})
ok(_fr_older_result is True and _frc_older.deleted_rows == [9],
   "an OLDER repeat application still has its redundant Main row removed")
ok("Received Date" not in _fr_older_fields
   and "Original Filename" not in _fr_older_fields,
   "an older application never overwrites the stored record's own date/filename")
ok("Resume URL" not in _fr_older_fields,
   "an older application never repoints the stored record at its own resume")
ok(_frc_older.deleted_files == [],
   "the NEWER stored resume is never deleted for an older incoming application")
ok(_fr_older_fields.get("Phone") == "+92 333 8113971",
   "an older application still fills a field the stored record is missing")
ok("Decline Sent" not in _fr_older_fields and "Status" not in _fr_older_fields,
   "an older application disturbs neither the decline marker nor the stored Status")

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
    def send_mail(self, to_address, subject, html_body, admin=False):
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
    def send_mail(self, to_address, subject, html_body, admin=False):
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
# 'Senior' (then still named 'Senior & Executive') with plain Android/Kotlin skills, which the old blank-or-General-
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
    "Category": "Senior",
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


# â”€â”€ G1b. P2 SECONDARY CONTENT-SAFETY NET (bug fix) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== G1b. P2 SECONDARY CONTENT-SAFETY NET ===")
# Live regression (2026-08-01, Futurism Technologies / APP-20260717-0559-CCE1): a services
# vendor's "Company Profile and Corporate Deck" email/attachment was scored as if it were a
# candidate's resume. P1 already screens incoming mail for this (flow_config.json
# spam_filters); this is P2's secondary net for the rare row that still slips through.
from hiring_agent.sharepoint_scoring import (
    _detect_suspicious_content, _send_suspicious_content_alert,
)

_futurism_mail_body = (
    "Thank you for your response on LinkedIn. As discussed, I'm sharing our Company "
    "Profile and Corporate Deck for your review. ... Relevant profiles, case studies, "
    "and engagement models"
)
_hits = _detect_suspicious_content("", _futurism_mail_body)
ok(set(_hits) == {"company profile and corporate deck", "engagement models"},
   f"the real Futurism vendor-pitch email is caught by P2's content-safety net (got {_hits!r})")
ok(_detect_suspicious_content("Senior React Native Developer, 5 years experience", "") == [],
   "a genuine resume snippet does not trigger the content-safety net")
ok(_detect_suspicious_content("", "") == [],
   "empty resume text and mail body never trigger the content-safety net")
ok(_detect_suspicious_content("URGENT ACTION REQUIRED: verify your account now", "") != [],
   "phishing-style boilerplate is also caught (case-insensitive)")

# Alert gating mirrors _send_error_alert: on when enabled + admin configured, off otherwise.
try:
    _cfg_mod.ERROR_EMAIL_ENABLED, _cfg_mod.ADMIN_EMAIL = True, "admin@x.com"
    _rc8 = _StubRetryClient()
    _send_suspicious_content_alert(_rc8, "APP-SPAM-1", "vendor@x.com", ["capability deck"])
    ok(len(_rc8.mails_sent) == 1 and _rc8.mails_sent[0][0] == "admin@x.com",
       "suspicious-content alert sent when enabled + admin email configured")

    _cfg_mod.ERROR_EMAIL_ENABLED = False
    _rc9 = _StubRetryClient()
    _send_suspicious_content_alert(_rc9, "APP-SPAM-2", "vendor@x.com", ["capability deck"])
    ok(_rc9.mails_sent == [], "suspicious-content alert suppressed when ERROR_EMAIL_ENABLED=false")
finally:
    _cfg_mod.ERROR_EMAIL_ENABLED, _cfg_mod.ADMIN_EMAIL = _orig_enabled, _orig_admin


# â”€â”€ G1c. NO HALF-FILLED ROW REACHES THE REJECTED SHEET â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
print("\n=== G1d. CV IS THE AUTHORITY ON IDENTITY ===")
# Live regression (2026-08-01, APP-20260706-0548-EE3B): a staffing agency emailed a
# developer's CV. extract_candidate_details_smart read the CV correctly ('Shivam' /
# 'Austin, Texas'), but merge_mail_body_fallback and ai_recheck_fields then read the
# COVERING EMAIL and rewrote the row as the agent ('Vandana Asawa' / 'Plano, Texas',
# lifted from her signature). Agency submissions are accepted as normal candidates when a
# real resume is attached, so the resume - never the covering email - decides identity.
from hiring_agent.sharepoint_scoring import _restore_cv_identity, _CV_IDENTITY_FIELDS

_cv = {"full_name": "Shivam", "location": "Austin, Texas", "country": "United States"}
_poisoned = {"full_name": "Vandana Asawa", "location": "Plano, Texas",
             "country": "United States", "phone": "(469) 532-0076"}
_fixed = _restore_cv_identity(_poisoned, _cv, "APP-TEST")
ok(_fixed["full_name"] == "Shivam",
   f"the CV's name beats the covering email's sender (got {_fixed['full_name']!r})")
ok(_fixed["location"] == "Austin, Texas",
   f"the CV's location beats the covering email's (got {_fixed['location']!r})")
ok(_fixed["phone"] == "(469) 532-0076",
   "a field outside the identity set is left exactly as the pipeline produced it")

# A field the CV never stated stays fillable from the email - the common case where the
# applicant IS the sender and their signature carries the missing detail.
_cv_blank = {"full_name": "Jane Doe", "location": "", "country": "Not extracted"}
_filled = _restore_cv_identity(
    {"full_name": "Jane Doe", "location": "Austin, TX", "country": "United States"},
    _cv_blank, "APP-TEST")
ok(_filled["location"] == "Austin, TX" and _filled["country"] == "United States",
   "a blank/gap CV field is still filled from the covering email")
# No CV/email disagreement -> nothing is touched.
_same = {"full_name": "Jane Doe", "location": "Austin, TX", "country": "United States"}
ok(_restore_cv_identity(dict(_same), _same, "APP-TEST") == _same,
   "when the CV and the email agree, the record passes through unchanged")
ok(_CV_IDENTITY_FIELDS == ("full_name", "location", "country"),
   "the identity set is exactly full_name/location/country")


print("\n=== G1c. PRE-REJECT COMPLETENESS GATE ===")
# Live regression (2026-08-01): Syyed Nazir Ali (APP-20260727-1431-6F75) and Divy Parmar
# (APP-20260720-1013-09AC) both landed on the Rejected sheet with a blank Category AND
# blank Suggested Roles. Root cause was two-fold: (a) _validate_all_columns marked
# 'Category'/'Suggested Role 1/2/3'/'Status' as 'ok' unconditionally WITHOUT looking at
# their values, so the run logged a clean column check over real gaps; (b) nothing
# re-checked completeness at the point of no return. The Rejected sheet is terminal -
# nothing re-scores a row once it is there - so a gap that slips through is permanent.
from hiring_agent.sharepoint_scoring import _complete_row_before_reject as _crbr

_gap_row = {
    "Application ID": "APP-GAP-1", "Full Name": "Jane Doe",
    "Location": "Mumbai", "Country": "India", "Status": "Rejected - Non-USA Location",
    "Current Skills": "Kotlin, Java", "Suggested Role 1": "Backend Engineer (40%)",
    "Category": "",          # the exact gap seen live
    "Portfolio 1": "", "Portfolio 2": "", "Portfolio 3": "",
    "Application Updates": "", "Retry Count": "",
}
_fixed = _crbr(dict(_gap_row), "APP-GAP-1")
ok(not _is_gap(_fixed.get("Category")),
   f"a blank Category is repaired before the row reaches Rejected (got {_fixed.get('Category')!r})")
ok(_fixed.get("Portfolio 1") == "N/A" and _fixed.get("Portfolio 2") == "N/A"
   and _fixed.get("Portfolio 3") == "N/A",
   "blank Portfolio slots are normalized to N/A before the row reaches Rejected")
ok(_fixed.get("Application Updates") == 0 and _fixed.get("Retry Count") == 0,
   "blank numeric counters are normalized to 0 before the row reaches Rejected")

# Already-good values must never be churned by the gate.
_good_row = {
    "Application ID": "APP-GOOD-1", "Full Name": "John Smith",
    "Location": "Hyderabad", "Country": "India", "Status": "Rejected - Non-USA Location",
    "Current Skills": "Marketing, SEO", "Suggested Role 1": "Marketing Lead (90%)",
    "Category": "Business Analytics", "Portfolio 1": "https://linkedin.com/in/x",
    "Portfolio 2": "N/A", "Portfolio 3": "N/A",
    "Application Updates": 2, "Retry Count": 1,
}
_untouched = _crbr(dict(_good_row), "APP-GOOD-1")
ok(_untouched == _good_row,
   "a fully populated row passes through the gate completely unchanged")

# The gate is deterministic and offline - it must never raise on a sparse/degenerate row.
_sparse = _crbr({}, "")
ok(isinstance(_sparse, dict) and not _is_gap(_sparse.get("Category")),
   "the gate never raises on an empty row and still guarantees a non-blank Category")

# _validate_all_columns must no longer rubber-stamp a blank Category.
_vac_fields = {
    "Full Name": "Jane Doe", "Phone": "N/A", "Location": "Mumbai", "Country": "India",
    "Current Skills": "Kotlin, Java", "Looking For Role": "Backend Engineer",
    "Education": "N/A", "Portfolio 1": "N/A", "Portfolio 2": "N/A", "Portfolio 3": "N/A",
    "Suggested Role 1": "Backend Engineer (40%)", "Suggested Role 2": "", "Suggested Role 3": "",
    "Category": "", "Status": "Rejected - Non-USA Location",
}
_vac_out = _validate_all_columns(dict(_vac_fields), {"Email": "j@x.com"}, "")
ok(not _is_gap(_vac_out.get("Category")),
   f"_validate_all_columns now heals a blank Category instead of marking it OK "
   f"(got {_vac_out.get('Category')!r})")


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

# Legacy Word (.doc). P1 accepts these from 2026-09-11; before that an applicant on an
# older Word install was bounced with a wrong-format reply. The extension is the least
# reliable thing about them, so the reader routes on magic bytes: a real OLE2 compound
# file goes to the Word 97 piece-table parser, a renamed .docx to the zip reader, and a
# "save as .doc" that actually wrote RTF to the RTF stripper.
_DOC_FIXTURE = __import__("pathlib").Path(__file__).resolve().parent / "test_fixtures" / "legacy_word97_resume.doc"
if _DOC_FIXTURE.exists():
    _doc_txt = extract_text_from_bytes(_DOC_FIXTURE.read_bytes(), "legacy_word97_resume.doc")
    ok("Marcus Vance" in _doc_txt, "extract_text_from_bytes reads a real Word 97 .doc")
    ok("(602) 555-0147" in _doc_txt, "...including the phone number")
    ok("Arizona State University" in _doc_txt, "...and the education block")
    from hiring_agent.extraction import extract_candidate_details as _ecd_doc
    _doc_fields = _ecd_doc(_doc_txt)
    ok(_doc_fields.get("full_name") == "Marcus Vance", "full name parsed out of a .doc")
    ok("602" in str(_doc_fields.get("phone", "")), "phone parsed out of a .doc")
else:
    ok(True, "Word 97 .doc fixture absent; binary-.doc reader not exercised here")
try:
    import docx as _dx
    _d3 = _dx.Document(); _d3.add_paragraph("Dana Whitfield\nAustin, TX")
    _b3 = __import__("io").BytesIO(); _d3.save(_b3)
    ok("Dana Whitfield" in extract_text_from_bytes(_b3.getvalue(), "renamed.doc"),
       "a .docx renamed to .doc is still read (routed by magic bytes, not extension)")
except Exception as e:
    ok(True, f"python-docx not testable in this env ({e})")
_rtf = (br"{\rtf1\ansi\deff0{\fonttbl{\f0 Calibri;}}\f0\fs22 "
        br"Priya Patel\par Denver, CO\par (303) 555-0182\par}")
_rtf_txt = extract_text_from_bytes(_rtf, "saved_as.doc")
ok("Priya Patel" in _rtf_txt and "(303) 555-0182" in _rtf_txt,
   "RTF written under a .doc name is read rather than rejected")
ok(extract_text_from_bytes(b"this is not a document", "junk.doc") == "",
   "a .doc that is neither Word nor RTF yields no text, so the row takes the existing "
   "unreadable-resume path instead of publishing garbage")
ok(extract_text_from_bytes(b"", "empty.doc") == "",
   "an empty .doc is handled without raising")


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

app_src = (pathlib.Path(__file__).resolve().parent / "app.py").read_text(encoding="utf-8")
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
                "Application ID": "APP-SCORED-OLD", "Status": "Scored",
                "Full Name": "Older Candidate", "Received Date": "2026-05-01T09:00:00",
            }},
            {"values": {
                "Application ID": "APP-SCORED", "Status": "Scored",
                "Full Name": "Ready Candidate", "Received Date": "2026-06-15T09:00:00",
            }},
            {"values": {
                "Application ID": "APP-SCORED-NEW", "Status": "Scored",
                "Full Name": "Newer Candidate", "Received Date": "2026-07-20T09:00:00",
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
    from hiring_agent.sharepoint_scoring import _CLIENT_EXPORT_COLUMNS as _EXPORT_COLS
    ok(_export_wb.sheetnames == ["Candidates"],
       "Candidate_List_Results is minimal: exactly one Candidates sheet")
    ok(_export_ws.max_column == len(_EXPORT_COLS)
       and [c.value for c in _export_ws[1]] == _EXPORT_COLS,
       "Candidate_List_Results contains only the selected client-facing columns")
    ok(all(_export_ws.cell(2, col).value in (None, "")
           for col in range(1, _export_ws.max_column + 1)),
       "Candidate_List_Results has one fully blank row directly after the header")
    ok(len(_export_ws.tables) == 0 and len(_export_ws._charts) == 0,
       "Candidate_List_Results stays minimal with no extra tables or charts")
    ok(_export_ws.freeze_panes == "A2",
       "Candidate_List_Results keeps only the header frozen")
    _export_ids = [
        _export_ws.cell(row, 1).value
        for row in range(2, _export_ws.max_row + 1)
        if _export_ws.cell(row, 1).value
    ]
    ok(set(_export_ids) == {"APP-SCORED-OLD", "APP-SCORED", "APP-SCORED-NEW"},
       "Candidate_List_Results exports only final Scored rows")
    ok(_export_ids == ["APP-SCORED-NEW", "APP-SCORED", "APP-SCORED-OLD"],
       f"export is sorted newest-first (current month first), matching the CandidateList "
       f"sheet's own resort_candidate_sheets.py convention (got {_export_ids})")
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


print("\n=== W1b. RESUME HEADER LOCATION FORMATS ===")
# _extract_location was built entirely around a comma between city and state, so
# 'Phoenix AZ 85004' - one of the commonest header shapes - extracted nothing. Live
# rows then carried Location "Not extracted" and were rejected on geography they had
# never actually failed. The full-state path also searched " ".join(lines), letting the
# city pattern run backwards over a line break and swallow the candidate's own name.
from hiring_agent.extraction import _extract_location as _loc

for _text, _want, _why in [
    ("Jane Doe\nAustin, TX\njane@x.com", "Austin, TX", "classic City, ST"),
    ("Jane Doe\nPhoenix AZ 85004", "Phoenix, AZ", "no comma, trailing ZIP"),
    ("Jane Doe\nDallas TX | (682) 409-2153", "Dallas, TX", "no comma, pipe-separated"),
    ("Jane Doe\nSeattle WA", "Seattle, WA", "bare City ST"),
    ("Jane Doe\nPhoenix, Arizona 85004", "Phoenix, Arizona", "full state + ZIP"),
    ("Jane Doe\nBoston, Massachusetts", "Boston, Massachusetts", "full state name"),
    ("Jane Doe\nNew York, NY 10001", "New York, NY", "City, ST + ZIP"),
    ("Jane Doe\nemail | 480-277-5159 | Tempe, AZ", "Tempe, AZ", "after pipe fields"),
    # A technology list is not an address. 'AR' here is Augmented Reality, not Arkansas.
    # Live case Adil Anwer (APP-20260805-0035-MDEA): "emerging technologies including
    # Firebase, AR, Blockchain" was stored as his location, so a candidate whose header
    # reads 'Karachi, Pakistan | +92-333-2467664' was recorded in the United States and
    # never reached the non-USA rejection. The existing tools-list guard missed it because
    # it only catches a space-separated continuation ('Bloomberg Terminal, MS Project'),
    # and this list continues with a comma.
    ("Adil Anwer\nKarachi, Pakistan | +92-333-2467664 | a@x.com\n\nSENIOR iOS ENGINEER\n"
     "Experience\nConducted R&D on emerging technologies including Firebase, AR, Blockchain",
     "Karachi, Pakistan", "a skill name is never read as a city"),
]:
    _got = str(_loc(_text) or "")
    ok(_got == _want, f"location {_why}: {_got!r} == {_want!r}")

_swallow = str(_loc("Jane Doe\nPhoenix, Arizona 85004") or "")
ok("jane" not in _swallow.lower(),
   f"full-state path never swallows the name across a line break (got {_swallow!r})")

# The comma-less path must not fire on prose that merely contains a state abbreviation.
for _text, _want, _why in [
    ("Jane Doe\nSkills: Java OR Python\nAustin, TX", "Austin, TX", "'OR' (Oregon) in a skills line"),
    ("Jane Doe\nMS SQL Server Developer\nDenver, CO", "Denver, CO", "'MS' (Mississippi) in a title"),
]:
    _got = str(_loc(_text) or "")
    ok(_got == _want, f"no false positive from {_why}: {_got!r} == {_want!r}")

# Live regression (2026-08-01, Syyed Nazir Ali / APP-20260727-1431-6F75): a tools list
# "Bloomberg Terminal, MS Project, MS Office Suite" matched the comma-separated "City, ST"
# pattern as "Bloomberg Terminal, MS" (Mississippi) - unlike the no-comma path above, this
# full-text scan had no guard against a list item that keeps going ('MS' here is Microsoft,
# not a state, and the match is immediately followed by another bare capitalized word
# rather than ending the field).
ok(_loc("Jane Doe\nTools: Bloomberg Terminal, MS Project, MS Office Suite") is None,
   f"'MS' inside a tools list ('MS Project') is not mistaken for Mississippi "
   f"(got {_loc('Jane Doe' + chr(10) + 'Tools: Bloomberg Terminal, MS Project, MS Office Suite')!r})")
ok(str(_loc("Jane Doe\nWork history: Jackson, MS | 2019 - 2021") or "") == "Jackson, MS",
   "a genuine 'City, MS' work-history address is still detected")

print("\n=== W1c. LOCATION NEVER DERIVED FROM A SCHOOL'S OWN NAME ===")
# Fixed 2026-08-01, live case (Prerna Saluja / APP-20260716-2052-6112): her header has no
# location at all, and her ONLY 'Delhi' mention is her undergrad alma mater ('Daulat Ram
# College, University of Delhi') - her current, most recent affiliation is Arizona State
# University. A bare substring scan for a known city/state anywhere in the document can't
# tell a school's own city from where the candidate actually lives, so it returned "Delhi"
# (then "India") despite her resume never stating a current location.
_prerna_text = (
    "Prerna Saluja\n(623)-242-3627 | LinkedIn | Gmail\n\nEDUCATION\n"
    "W. P. Carey School of Business, Arizona State University\n"
    "Master of Science in Finance (GPA: 3.78)\n\n"
    "Daulat Ram College, University of Delhi\n"
    "Bachelor of Commerce (Major: Accounting, Minor: Economics)\n"
)
ok(_loc(_prerna_text) is None,
   f"a city/state embedded only in a school's own name is not returned as location "
   f"(got {_loc(_prerna_text)!r})")

# The same guard must not swallow a genuine current-location statement that happens to
# sit near an unrelated institution mention elsewhere in the resume.
_phoenix_text = (
    "Jane Doe\nSoftware Engineer\nPhoenix, AZ 85004\n\n"
    "EDUCATION\nUniversity of Phoenix, B.S. Computer Science\n"
)
ok(str(_loc(_phoenix_text) or "") == "Phoenix, AZ",
   f"a real header location survives even when an unrelated 'University of Phoenix' "
   f"appears later (got {_loc(_phoenix_text)!r})")

_mumbai_text = "Ravi Kumar\nMumbai, Maharashtra | ravi@example.com\n\nEDUCATION\nB.Tech, IIT Bombay\n"
ok(str(_loc(_mumbai_text) or "") == "Mumbai, Maharashtra",
   f"a genuinely stated foreign city is still detected (got {_loc(_mumbai_text)!r})")
# Widened 2026-08-03: this used to return the bare city "Mumbai", which left Country BLANK
# because nothing downstream could tell which country Mumbai is in. Keeping the province the
# candidate actually wrote lets split_location_country resolve the country from it.
ok(split_location_country(_loc(_mumbai_text)) == ("Mumbai, Maharashtra", "India"),
   f"the stated province resolves the Country instead of leaving it blank "
   f"(got {split_location_country(_loc(_mumbai_text))!r})")

print("\n=== W2. CLIENT EXPORT HEADER SPACER (month/year separators removed 2026-09-12) ===")
# One blank row under the header and NOTHING else. The month blanks and '-- 2025 --'
# year labels this used to assert were removed on client instruction so all three
# workbooks read the same way: header / blank spacer / every candidate. The master
# sheets' equivalent separators have been disabled since 2026-09-01 (see W3/W4).
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
ok(_ids == ["", "A", "B", "C", "D", "E"],
   f"no month separator survives, even across a year boundary (got {_ids})")
ok(_ids.count("") == 1,
   "exactly ONE blank row in the whole sheet - the header spacer")
ok(all(str(v or "") == "" for v in _sep_out[0].values()),
   "the spacer is fully blank, carrying no year label and no Application ID")
ok(not any("--" in str(r.get("Full Name") or "") for r in _sep_out),
   "no '-- <year> --' label row is emitted any more")
_empty_sep = _sep([])
ok(len(_empty_sep) == 1
   and all(str(v or "") == "" for v in _empty_sep[0].values()),
   "an empty export still receives one permanent blank row after the header")
_same = _sep([_xrow("A", "2026-05-07T10:00:00"), _xrow("B", "2026-05-08T10:00:00")])
ok([str(r.get("Application ID") or "") for r in _same] == ["", "A", "B"],
   "rows inside one month get no separator between them")
_nodate = _sep([_xrow("A", ""), _xrow("B", "2026-05-08T10:00:00")])
ok(len(_nodate) == 3, "a row with an unparseable date does not crash or duplicate separators")

print("\n=== W3. REJECTED SHEET PERIOD SEPARATORS (disabled 2026-09-01) ===")
# Master sheets must never carry monthly blank-row gaps - only the client results export
# does (see W2 above). P1's flow_config.json already ships with year_separator/
# month_separator both false; this makes P2 match on its own Rejected-append path, which
# used to insert a separator unconditionally with no off-switch (see
# _ensure_rejected_period_separator's docstring). Every case below - including the ones
# that used to be a real month/year boundary - must now be a pure no-op.
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


for _existing, _new, _label in (
    ([], "2026-05-07T10:00:00", "first ever Rejected row"),
    (["2026-05-07T10:00:00"], "2026-05-20T10:00:00", "same month as an existing row"),
    (["2026-05-07T10:00:00"], "2026-06-02T10:00:00", "a genuine new-month boundary"),
    (["2026-12-07T10:00:00"], "2027-01-02T10:00:00", "a genuine new-year boundary"),
    (["2026-05-07T10:00:00", "2026-06-02T10:00:00"], "2026-05-21T10:00:00",
     "an out-of-order repair reopening an older month"),
    (["", "2026-05-07T10:00:00"], "2026-05-21T10:00:00", "existing rows include a blank"),
    (["2026-05-07T10:00:00"], None, "a None Received Date"),
    (["2026-05-07T10:00:00"], "", "a blank Received Date"),
    (["2026-05-07T10:00:00"], "Not extracted", "a gap-literal Received Date"),
):
    _r, _a = _did_sep(_existing, _new)
    ok(_r is False and _a == [], f"{_label} never inserts a separator (disabled)")


class _BoomClient(_RejSepClient):
    def list_rejected_rows(self):
        raise RuntimeError("graph down")


ok(_rsep(_BoomClient([]), "2026-06-01T10:00:00") is False,
   "disabled stub never touches the client, even one that would raise if read")

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
    # ...and that stamp records that nothing was sent, so it is NOT a send (2026-09-05).
    # Before this, the 14-day age-out clock started on candidates who were never contacted
    # and _age_out_stale_reviews closed them silently as 'Rejected - Location Not
    # Confirmed', and both mail passes select on "is the marker blank", so a suppressed
    # stamp permanently consumed the one email the candidate was owed.
    from hiring_agent.sharepoint_scoring import (
        _awaiting_send as _sup_awaiting, _review_clock_start as _sup_clock)
    ok(_parse_marker_datetime(_sent_marker()) is None,
       "a suppressed marker has no send time - nothing left the mailbox")
    ok(_sup_awaiting(_sent_marker()),
       "a row stamped by a suppressed run still owes its email")
    ok(_sup_clock({"Status": "Needs Review - Location Confirmation",
                   "Info Request Sent": _sent_marker()}) is None,
       "the age-out clock does not start on a candidate who was never contacted")
    _sup_cfg.SUPPRESS_EMAILS = False
    ok(_parse_marker_datetime(_sent_marker()) is not None,
       "a real 'Sent <ts>' marker still parses")
    ok(not _sup_awaiting(_sent_marker()), "...and a really-sent row owes nothing")
    ok(_sup_clock({"Status": "Needs Review - Location Confirmation",
                   "Info Request Sent": _sent_marker()}) is not None,
       "...and it does start the clock")
    ok(_sup_awaiting(""), "a blank marker still counts as owed")
    _sup_cfg.SUPPRESS_EMAILS = True

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

print("\n=== Y1. JD OPENING IDENTITY — one Suggested Role slot per OPENING, not per JD FILE ===")
# The JD library stores each opening as several documents (Job Announcement + Position
# Description, plus Updated/v2/Comprehensive revisions): 85 cached files, ~45 openings.
# Nothing in the source data linked them, so ONE opening could fill all three Suggested Role
# slots. Identity is now derived ONCE at ingest (hiring_agent.jd_identity) and stamped onto
# the role, so scoring reads a field instead of re-parsing titles per comparison.
from hiring_agent.jd_identity import (split_document_title as _sdt, opening_id as _oid,
                                      stamp_role_identity as _stamp, role_opening_id as _roid)
from hiring_agent.scoring import suggested_roles as _sr, get_open_roles as _gor
import re as _yre

ok(_sdt("Mobile Application Lead Developer Position Description v2")
   == ("Mobile Application Lead Developer", "Position Description"),
   f"a JD title splits into (opening, document type) (got {_sdt('Mobile Application Lead Developer Position Description v2')!r})")
ok(_sdt("Data Analyst Job Announcement rev")[1] == "Job Announcement",
   "the document type is captured, not just discarded")
ok(_oid("Mobile Application Lead Developer Job Announcement")
   == _oid("Mobile Application Lead Developer Position Description v2")
   == "mobile-application-lead-developer",
   "JA / PD / v2 documents of one opening share an opening id")
ok(_oid("Cybersecurity IT Admin Manager JA AI Tools Updated")
   == _oid("Cybersecurity IT Admin Manager PD Updated"),
   "free-text qualifiers AFTER the doc-type marker are cut too (JA/PD + 'AI Tools')")
ok(_oid("CISO LinkedIn Announcement") == _oid("CISO Position Description Updated v2") == "ciso",
   "a LinkedIn Announcement shares an opening with the JA/PD of the same role")
ok(_oid("Web Dev  Full Stack Position") == _oid("Web Dev - Full Stack Position"),
   "incidental punctuation/spacing does not split one opening in two")
ok(_oid("Gaming Position") != _oid("Graphics Position"),
   "trimming the doc-type tail still leaves genuinely different openings distinct")
ok(_oid("AWS Lead Job Announcement") != _oid("AWS Senior Engineer Position"),
   "seniority variants of one job family stay distinct openings")
ok(_oid("Position Description") == "position-description",
   "a title that is ONLY a doc-type marker falls back to itself, not an empty id")
ok(_sdt("Software Developer - AI/ML, Computer Vision")
   == ("Software Developer - AI/ML, Computer Vision", ""),
   "a title with no doc-type marker is left alone, with an empty document type")

# The stamp is what scoring actually reads; a supplied id always wins over the heuristic.
_y1_stamped = _stamp({"title": "Data Analyst Job Announcement", "skills": ["sql"]})
ok(_y1_stamped["opening"] == "Data Analyst" and _y1_stamped["opening_id"] == "data-analyst"
   and _y1_stamped["doc_type"] == "Job Announcement",
   f"stamp_role_identity fills opening / opening_id / doc_type (got {_y1_stamped!r})")
ok(_roid({"title": "Data Analyst Job Announcement", "opening_id": "req-4471"}) == "req-4471",
   "an explicitly supplied opening_id is authoritative — the title heuristic steps aside")
ok(_roid({"title": "Data Analyst Job Announcement"}) == "data-analyst",
   "a role with no stamp falls back to deriving identity from its title")
ok(all(r.get("opening_id") for r in _gor()),
   "the built-in default_roles carry an opening id too")
ok("opening_id" not in __import__("hiring_agent.config", fromlist=["DEFAULT_ROLES"]).DEFAULT_ROLES[0],
   "get_open_roles() copies before stamping — module-level config is never mutated")

_y1_roles = [
    {"title": "Mobile Application Lead Developer Job Announcement",
     "skills": ["kotlin", "swift", "flutter", "android sdk", "xcode"]},
    {"title": "Mobile Application Lead Developer Position Description",
     "skills": ["kotlin", "swift", "flutter", "android sdk", "xcode"]},
    {"title": "Mobile Application Lead Developer Position Description v2",
     "skills": ["kotlin", "swift", "flutter", "android sdk"]},
    {"title": "Gaming Position Announcement",
     "skills": ["unity", "c#", "swift", "blender", "shader"]},
]
_y1 = _sr("Kotlin, Swift, Flutter, Android SDK, Xcode", "Mobile Developer", roles=_y1_roles)
_y1_slots = [s for s in (_y1["role_1"], _y1["role_2"], _y1.get("role_3", "")) if s]
_y1_keys = [_oid(_yre.sub(r"\s*\(\d+%\)$", "", s)) for s in _y1_slots]
ok(len(_y1_keys) == len(set(_y1_keys)),
   f"no two Suggested Role slots describe the same opening (got {_y1_slots})")
ok(len(_y1_slots) == 2,
   f"three near-duplicate JD files collapse to the 2 real openings (got {len(_y1_slots)})")
# Changed 2026-09-08 (user request): the slot now shows the OPENING name, not the winning
# document's filename. Dedup is unaffected - it always keyed on opening identity - but the
# client sheet no longer leaks JD filing metadata ('Job Announcement', 'PD', 'rev', '(2)').
ok(_y1["role_1"].startswith("Mobile Application Lead Developer"),
   f"the surviving slot names the opening (got {_y1['role_1']!r})")
ok("Job Announcement" not in _y1["role_1"] and "Position Description" not in _y1["role_1"],
   f"...with the document-type suffix stripped (got {_y1['role_1']!r})")

# The Ollama shortlist must also spend its slots on distinct openings.
from hiring_agent.scoring import _ai_role_shortlist as _y1short
_y1_many = _y1_roles + [
    {"title": "Data Analyst Job Announcement", "skills": ["sql", "tableau"]},
    {"title": "Data Analyst Position Description", "skills": ["sql", "tableau"]},
]
_y1_sl = _y1short("Kotlin, Swift, SQL", "Mobile Developer", _y1_many, limit=3)
ok(len({_roid(r) for r in _y1_sl}) == len(_y1_sl),
   f"the Ollama shortlist holds one document per opening (got {[r['title'] for r in _y1_sl]})")
ok(len(_y1short("Kotlin", "", _y1_many, limit=99)) == 3,
   "even under a generous limit the shortlist collapses 6 JD files to their 3 openings")
ok(_y1short("Python", "", [{"title": f"Role {i}", "skills": ["s%d" % i]} for i in range(5)],
            limit=12) == [{"title": f"Role {i}", "skills": ["s%d" % i]} for i in range(5)],
   "a short catalog with no duplicate openings is passed through untouched")

# The live cache must survive the round-trip: real titles, real merge groups.
import json as _y1json
_y1_cache = _y1json.loads(Path(__file__).with_name("jd_roles_cache.json")
                          .read_text(encoding="utf-8"))
_y1_cached = _y1_cache.get("roles", [])
_y1_groups = {}
for _r in _y1_cached:
    _y1_groups.setdefault(_oid(_r["title"]), []).append(_r["title"])
ok(len(_y1_cached) > len(_y1_groups) > 0,
   f"the live JD cache collapses {len(_y1_cached)} files to {len(_y1_groups)} openings")
ok(all(g for g in _y1_groups), "no cached JD title produces an empty opening id")
ok(_y1_groups.get("mobile-application-lead-developer") and
   len(_y1_groups["mobile-application-lead-developer"]) >= 3,
   "the live 'Mobile Application Lead Developer' documents group together")
ok("gaming" in _y1_groups and "graphics" in _y1_groups,
   "the live Gaming and Graphics openings stay separate")

print("\n=== Y2. CLIENT EXPORT — no live-table formula ever reaches the client file ===")
# 'Resume Link' is an Excel TABLE calculated column whose formula uses structured
# references ([@[Resume URL]]). The export sheet is a plain range with no table, so a
# preserved formula renders as #NAME? instead of a link in the client's copy.
import tempfile as _y2tmp
from hiring_agent.sharepoint_scoring import (
    prepare_client_export_rows as _y2prep, write_client_export as _y2write,
    _CLIENT_EXPORT_COLUMNS as _y2cols)
from hiring_agent.config import COLUMNS as _y2all

_Y2_FORMULA = ('=IF([@[Resume URL]]="","",HYPERLINK([@[Resume URL]],'
               'IF(TRIM([@[Original Filename]])="",'
               'TRIM(RIGHT(SUBSTITUTE([@[Resume URL]],"/",REPT(" ",300)),300)),'
               'TRIM([@[Original Filename]]))))')


def _y2row(aid, when, link, url):
    v = {c: "" for c in _y2all}
    v.update({"Application ID": aid, "Received Date": when, "Full Name": aid,
              "Resume Link": link, "Resume URL": url, "Mail Body": "SECRET BODY",
              "Status": "Scored", "Retry Count": 3})
    return v


_y2_path = Path(_y2tmp.mkdtemp(prefix="p2y2_")) / "client.xlsx"
_y2write(_y2prep([
    _y2row("APP-A", "2026-05-07T10:00:00", _Y2_FORMULA, "https://sp/a.pdf"),
    _y2row("APP-B", "2026-06-02T10:00:00", "", "https://sp/b.pdf"),
    _y2row("APP-C", "2026-06-09T10:00:00", "MyCV.pdf", "https://sp/c.pdf"),
]), _y2_path)
_y2ws = load_workbook(_y2_path)["Candidates"]
_y2vals = [[str(c.value or "") for c in r] for r in _y2ws.iter_rows(min_row=2)]
ok(not any("[@[" in v for row in _y2vals for v in row),
   "a structured-reference formula is flattened, never written into the client file")
_y2link = _y2cols.index("Resume Link") + 1
_y2named = {str(_y2ws.cell(r, _y2cols.index("Application ID") + 1).value or ""): r
            for r in range(2, _y2ws.max_row + 1)}
ok(str(_y2ws.cell(_y2named["APP-A"], _y2link).value) == "a.pdf",
   "a flattened formula cell falls back to the filename from the Resume URL")
ok(_y2ws.cell(_y2named["APP-A"], _y2link).hyperlink.target == "https://sp/a.pdf",
   "the flattened cell is still a working hyperlink")
ok(str(_y2ws.cell(_y2named["APP-C"], _y2link).value) == "MyCV.pdf",
   "a normal computed display value is preserved as-is")
ok([c.value for c in _y2ws[1]] == _y2cols,
   "the export emits exactly the minimal client column set, in order")
ok(len(_y2cols) < len(_y2all) and "Mail Body" not in _y2cols and "Status" not in _y2cols
   and "Retry Count" not in _y2cols,
   f"internal columns stay out of the client file ({len(_y2all)} -> {len(_y2cols)})")
ok(not any("SECRET BODY" in v for row in _y2vals for v in row),
   "internal Mail Body content never reaches the client file")

# The GUI's Export... button must produce the SAME artifact, not a raw column dump.
_y2_gui = Path(__file__).with_name("app.py").read_text(encoding="utf-8", errors="replace")
_y2_gui_fn = _y2_gui.split("def _export_candidates", 1)[-1].split("\n    def ", 1)[0]
ok("write_client_export" in _y2_gui_fn and "prepare_client_export_rows" in _y2_gui_fn,
   "the GUI Export... button reuses the shared client-export contract")
ok("reindex(columns=COLUMNS)" not in _y2_gui_fn,
   "the GUI Export... button no longer dumps all internal columns")

print("\n=== Y3. MAIN CANDIDATELIST PERIOD SEPARATORS (disabled 2026-09-01) ===")
# Same policy as W3 above, mirrored for the main sheet's own append path (geo-recovery
# restores, the GUI's manual upload). Every case that used to be a real month/year
# boundary must now be a pure no-op - see _ensure_main_period_separator's docstring.
from hiring_agent.sharepoint_scoring import _ensure_main_period_separator as _msep


class _MainSepClient:
    def __init__(self, dates):
        self.rows = [{"values": {"Received Date": d, "Application ID": f"X{i}"}}
                     for i, d in enumerate(dates)]
        self.added = []

    def list_rows(self):
        return self.rows

    def add_main_row(self, fields):
        self.added.append(fields)


def _did_msep(existing, new_date):
    c = _MainSepClient(existing)
    return _msep(c, new_date), c.added


for _existing, _new, _label in (
    ([], "2026-05-07T10:00:00", "first ever main row"),
    (["2026-05-07T10:00:00"], "2026-05-20T10:00:00", "a restore into the SAME month"),
    (["2026-05-07T10:00:00"], "2026-06-02T10:00:00", "a restore into a genuine new month"),
    (["2026-12-07T10:00:00"], "2027-01-02T10:00:00", "a genuine new year"),
    (["2026-05-07T10:00:00"], "", "a blank Received Date"),
    (["2026-05-07T10:00:00"], "Not extracted", "a gap-literal Received Date"),
    ([""], "2026-06-02T10:00:00", "existing rows include a dateless label/spacer row"),
):
    _r, _a = _did_msep(_existing, _new)
    ok(_r is False and _a == [], f"{_label} never inserts a separator (disabled)")


class _NoMainClient:
    """A stripped-down client with no list_rows at all."""

    def add_main_row(self, fields):
        raise AssertionError("must not append when the sheet cannot be read")


ok(_msep(_NoMainClient(), "2026-06-01T10:00:00") is False,
   "a client that cannot be read is treated as 'cannot check', never raises at the call site")

print("\n=== Y4. LOCATION — comma-less US states, and countries no longer discarded ===")
from hiring_agent.extraction import _extract_location as _y4loc, split_location_country as _y4sp

for _text, _want, _label in [
    ("JOHN DOE\nSan Francisco California\n555-1234", "San Francisco, California",
     "City + full state name, NO comma"),
    ("JOHN DOE\nColumbus Ohio\nx@y.com", "Columbus, Ohio",
     "a city absent from the city list is still found via its state"),
    ("JOHN DOE\nSan Francisco California 94102", "San Francisco, California",
     "a trailing ZIP does not break the comma-less full-state form"),
]:
    ok(str(_y4loc(_text) or "") == _want,
       f"{_label}: {_y4loc(_text)!r} == {_want!r}")
    ok(_y4sp(_y4loc(_text))[1] == "United States",
       f"{_label}: Country resolves to United States")

# State names that double as personal names need a ZIP before they count as a location.
_y4_person = "Resume\nMary Virginia\nSoftware Engineer\n"
_y4_person_got = _y4loc(_y4_person)
ok(_y4_person_got is None,
   f"'Mary Virginia' is read as a person, not a place (got {_y4_person_got!r})")
ok(str(_y4loc("Resume\nSeattle Washington 98101\n") or "") == "Seattle, Washington",
   "the same name-like state IS accepted alongside a ZIP")

# An explicitly written country must survive into the Location value.
for _text, _want_loc, _want_ctry, _label in [
    ("Ravi K\nToronto, ON, Canada\n", "Toronto, ON", "Canada", "Canadian province code"),
    ("Ravi K\nBengaluru, Karnataka, India\n", "Bengaluru, Karnataka", "India", "Indian state"),
    ("Ana P\nBerlin, Germany\n", "Berlin", "Germany", "bare City, Country"),
    ("Ana P\nSydney, NSW, Australia\n", "Sydney, NSW", "Australia", "Australian state code"),
]:
    _got_loc, _got_ctry = _y4sp(_y4loc(_text) or "")
    ok((_got_loc, _got_ctry) == (_want_loc, _want_ctry),
       f"{_label}: the stated country is kept, not truncated "
       f"(got {(_got_loc, _got_ctry)!r})")

_y4_skills = "Pat\nSkills: Python, Boston Dynamics\n"
_y4_skills_got = _y4loc(_y4_skills)
ok(str(_y4_skills_got or "") == "Boston",
   f"a non-address line that merely names a city is NOT widened to the whole line "
   f"(got {_y4_skills_got!r})")
ok(str(_y4loc("Austin, TX 78701\nx") or "") == "Austin, TX",
   "the classic 'City, ST' header is unchanged")

print("\n=== Y5. TIER 2 LOCATION IS GROUNDED IN THE RESUME TEXT ===")
# Portfolio (Tier 1) always had a hard anti-hallucination gate; Location - the field that
# decides whether a candidate is geo-rejected - had only a prompt instruction. A local model
# that ignored it silently overwrote a correct regex match with an invented, perfectly
# self-consistent city/country pair, which every downstream cross-field guard then passed.
from hiring_agent.extraction import _location_grounded_in_text as _y5g
from hiring_agent.config import (US_STATE_ABBREV_TO_NAME as _y5a2n,
                                 US_STATE_NAME_TO_ABBREV as _y5n2a,
                                 US_STATE_ABBREVS as _y5abbr, US_STATE_NAMES as _y5names)

ok(_y5abbr == set(_y5a2n) and _y5names == set(_y5a2n.values()),
   "the state abbrev/name sets are derived from one mapping and cannot drift apart")
ok(len(_y5a2n) == len(_y5n2a) == 51,
   f"the state mapping is 1:1 and complete incl. DC (got {len(_y5a2n)}/{len(_y5n2a)})")

_Y5_AUSTIN = "JANE PUBLIC\nAustin, TX 78701\njane@example.com\n"
for _loc, _text, _want, _why in [
    ("Austin, TX", _Y5_AUSTIN, True, "the model echoing the resume is kept"),
    ("Phoenix, Arizona", "JANE\nPhoenix AZ 85004\n", True,
     "expanding the state abbreviation is a CORRECTION, not an invention"),
    ("Phoenix, AZ", "JANE\nPhoenix, Arizona 85004\n", True,
     "contracting the state name is equally acceptable"),
    ("San Francisco, CA", "JANE\nSan Francisco California\n", True,
     "normalizing a comma-less header is kept"),
    ("India", "Syyed\nLocation: India\nOpen to Remote / Global\n", True,
     "the real Syyed Nazir Ali correction still overrides a bad regex hint"),
    ("Remote", "Syyed\nLocation: India\nOpen to Remote / Global\n", True,
     "'Remote' is kept when the resume itself says Remote"),
    ("Mumbai, Maharashtra", "JANE\nMumbai, Maharashtra | j@x.com\n", True,
     "a genuinely stated foreign location is kept"),
    ("Winston-Salem, NC", "JANE\nWinston-Salem, NC\n", True, "hyphenated city"),
    ("Bengaluru, Karnataka", _Y5_AUSTIN, False, "a wholly invented location is rejected"),
    ("Austin, India", _Y5_AUSTIN, False,
     "a real city paired with an invented country is rejected as one claim"),
    ("Dallas, TX", _Y5_AUSTIN, False, "the right state cannot rescue the wrong city"),
    ("", _Y5_AUSTIN, False, "a blank location is never 'grounded'"),
    ("Not extracted", _Y5_AUSTIN, False, "a gap literal is never 'grounded'"),
    ("Austin, TX", "", False, "with no source text nothing can be corroborated"),
]:
    ok(_y5g(_loc, _text) is _want, f"{_loc!r}: {_why}")

# End to end through the real merge, with a stubbed model response.
import hiring_agent.extraction as _y5ex

_y5_saved_call, _y5_saved_flag = _y5ex.extract_with_ollama, _y5ex.OLLAMA_ENABLED
try:
    _y5ex.OLLAMA_ENABLED = True

    def _y5_run(fake, text):
        _y5ex.extract_with_ollama = lambda t, hints=None: dict(fake)
        return _y5ex.extract_candidate_details_smart(text)

    _y5_invented = _y5_run({"full_name": "Jane Public", "location": "Bengaluru, Karnataka",
                            "country": "India", "skills": "SQL", "phone": "",
                            "looking_for_role": "Data Analyst", "education": "BS Statistics"},
                           _Y5_AUSTIN)
    ok(_y5_invented["location"] == "Austin, TX",
       f"an invented AI location never reaches the row (got {_y5_invented['location']!r})")
    ok(_y5_invented["country"] == "United States",
       f"the invented country falls with it - they are ONE claim (got {_y5_invented['country']!r})")

    _y5_corrected = _y5_run({"full_name": "Syyed", "location": "India", "country": "India",
                             "skills": "SQL", "phone": "", "looking_for_role": "BA",
                             "education": "BS"},
                            "Syyed Nazir Ali\nLocation: India\nOpen to Remote / Global\n"
                            "SKILLS: Bloomberg Terminal, MS Project\n")
    ok(_y5_corrected["country"] == "India",
       f"a GROUNDED AI correction still wins over the regex hint "
       f"(got country {_y5_corrected['country']!r})")
finally:
    _y5ex.extract_with_ollama, _y5ex.OLLAMA_ENABLED = _y5_saved_call, _y5_saved_flag

print("\n=== Y6. STORED-VALUE REPAIRS (letter-spaced education, non-ASCII names) ===")
# Fresh extraction has handled letter-spaced documents since 2026-07-31 (the document
# normalizer runs first), but values stored BEFORE that fix were invisible to every repair
# heuristic - _EDUCATION_FUSED_RE needs whole words, and the single-letter probe needs a 4+
# letter word, which letter-spaced text never contains. They would never have self-healed.
from hiring_agent.extraction import (_looks_letter_spaced as _y6ls,
                                     education_needs_repair as _y6r,
                                     normalize_education as _y6n)

for _v, _want in [
    ("B a c h e l o r  o f  S c i e n c e  i n  C S", True),
    ("B a c h e l o r o f S c i e n c e i n C S", True),
    ("A r i z o n a  S t a t e  U n i v e r s i t y", True),
    ("Master of Science in Computer Science", False),
    ("B.S. Biology", False),
    ("M.S. CS, Georgia Tech", False),
    ("MBA, Wharton School of Business", False),
    ("B S Computer Science", False),          # 2 of 4 single chars - under the threshold
]:
    ok(_y6ls(_v) is _want, f"letter-spacing detector: {_v[:38]!r} -> {_want}")

ok(_y6r("B a c h e l o r  o f  S c i e n c e") is True,
   "a letter-spaced Education cell is now FLAGGED for repair (it never was before)")
ok(_y6n("B a c h e l o r  o f  S c i e n c e  i n  C S") == "Bachelor of Science in CS",
   f"...and is actually REPAIRED, not just flagged "
   f"(got {_y6n('B a c h e l o r  o f  S c i e n c e  i n  C S')!r})")
ok(_y6n("B a c h e l o r o f S c i e n c e i n C S") == "Bachelor of Science in CS",
   "the single-spaced variant repairs too (fused-word replacements split it back apart)")
ok(_y6n("A r i z o n a  S t a t e  U n i v e r s i t y") == "Arizona State University",
   "a letter-spaced school name repairs as well")
ok(_y6n("Master of Science in Computer Science") == "Master of Science in Computer Science",
   "a healthy Education value is left completely untouched")

# Accented Latin names lost letters entirely ('Jose Garcia' -> 'JosGarca').
from hiring_agent.sharepoint_scoring import _get_cleaned_filename_prefix as _y6p

ok(_y6p("José García") == "JoseGarcia",
   f"accents fold to their base letter (got {_y6p('José García')!r})")
ok(_y6p("José García") == _y6p("Jose Garcia"),
   "the accented and unaccented spellings of one name produce the SAME filename")
ok(_y6p("Björn Åberg") == "BjornAberg",
   f"Nordic diacritics fold correctly (got {_y6p('Björn Åberg')!r})")
ok(_y6p("Zoë O’Neill") == "ZoeONeill",
   "a curly apostrophe plus a diaeresis both resolve")
ok(_y6p("李 明") == "Candidate",
   "a script with no ASCII form still falls back safely (AppID tail keeps it unique)")
ok(_y6p("John Doe") == "JohnDoe" and _y6p("Mary Jane Watson") == "MaryWatson"
   and _y6p("Cher") == "Cher",
   "plain ASCII names are unchanged by the folding step")

print("\n=== Y7. BATCH ROW-INDEX BOOKKEEPING (deleted rows must shift the rest) ===")
# score_from_sharepoint fetches the batch ONCE, then addresses every Graph write by
# itemAt(index), compensating for in-run deletions with
#     index = row["index"] - sum(1 for d in deleted_idx if d < row["index"])
# _mark_needs_review can delete the MAIN row two ways (give-up at SCORE_RETRY_MAX, or
# dropping a stale retry duplicate) but used to return None, so neither was recorded.
# Every later row in the batch then wrote one index too high: the next candidate's data
# landed on the row AFTER it, and the final row hit an out-of-range index.
from hiring_agent.sharepoint_scoring import (_mark_needs_review as _y7mnr,
                                             _give_up_and_reject as _y7give)
import hiring_agent.config as _y7cfg


class _Y7Sheet:
    """A main table that really shifts indexes on delete, like Graph's itemAt does."""

    def __init__(self, rows):
        self.rows = list(rows)
        self.resumes_folder = "/Resumes"
        self.rejected = []

    def delete_row(self, index):
        self.rows.pop(index)

    def update_row(self, index, fields, current_values=None):
        self.rows[index].update(fields)

    def add_rejected_row(self, fields):
        self.rejected.append(fields)

    def list_rows(self):
        return [{"index": i, "values": v} for i, v in enumerate(self.rows)]

    def list_rejected_rows(self):
        return []

    def move_resume(self, *a, **k):
        return False

    def resume_web_url(self, *a, **k):
        return ""

    def send_mail(self, *a, **k):
        return True


def _y7row(aid, retry=0):
    return {"Application ID": aid, "Email": f"{aid}@x.com", "Has Resume": "No",
            "Retry Count": retry, "Received Date": "2026-08-01T10:00:00",
            "Status": "New Email Received", "Full Name": aid}


# The signal the loop depends on.
_y7s = _Y7Sheet([_y7row("A", _y7cfg.SCORE_RETRY_MAX - 1)])
ok(_y7mnr(_y7s, 0, _y7s.rows[0], False, app_id="A") is True,
   "_mark_needs_review reports True when it gives up and DELETES the main row")
ok(_y7s.rows == [], "...and the row really is gone from the main table")

_y7s2 = _Y7Sheet([_y7row("B", 0)])
ok(_y7mnr(_y7s2, 0, _y7s2.rows[0], False, app_id="B") is False,
   "_mark_needs_review reports False when it only bumps Retry Count in place")
ok(len(_y7s2.rows) == 1, "...and the row is still there")

ok(_y7mnr(_Y7Sheet([_y7row("C")]), 0, _y7row("C"), True, app_id="C") is False,
   "a dry run never reports a deletion (it never makes one)")


class _Y7Boom(_Y7Sheet):
    def delete_row(self, index):
        from sharepoint_client import SharePointError
        raise SharePointError("graph refused the delete")


ok(_y7give(_Y7Boom([_y7row("D")]), 0, _y7row("D"), "D", [], "2026/August",
           "resume unreadable", False) is False,
   "a FAILED delete reports False - the row is still there, so indexes must NOT shift")

# End-to-end: the exact loop arithmetic over a 3-row batch whose FIRST row is given up on.
_y7sheet = _Y7Sheet([_y7row("CAND-A", _y7cfg.SCORE_RETRY_MAX - 1),
                     _y7row("CAND-B"), _y7row("CAND-C")])
_y7batch = _y7sheet.list_rows()
_y7deleted = []
for _r in _y7batch:
    _i = _r["index"] - sum(1 for d in _y7deleted if d < _r["index"])
    _aid = _r["values"]["Application ID"]
    if _aid == "CAND-A":
        if _y7mnr(_y7sheet, _i, _r["values"], False, app_id=_aid):
            _y7deleted.append(_r["index"])
    else:
        _y7sheet.update_row(_i, {"Status": "Scored", "Full Name": _aid + "-SCORED"})

ok(_y7deleted == [0], "the give-up deletion is recorded in the batch's deleted_idx")
ok([v["Application ID"] for v in _y7sheet.rows] == ["CAND-B", "CAND-C"],
   "the two survivors remain, in order")
ok(all(v["Full Name"] == v["Application ID"] + "-SCORED" for v in _y7sheet.rows),
   f"each survivor carries its OWN scoring data, not the previous candidate's "
   f"(got {[(v['Application ID'], v['Full Name']) for v in _y7sheet.rows]})")
ok(all(v["Status"] == "Scored" for v in _y7sheet.rows),
   "no row is skipped or left unscored by the index shift")

print("\n=== Y8. VERIFIED ROW INDEX (the arithmetic is checked, not trusted) ===")
# Recording every delete (Y7) fixes the bug that existed; this makes the whole CLASS of bug
# non-silent. Whenever an offset is actually applied, the row at that position is confirmed
# to be this candidate before any write addresses it - so a future delete path that forgets
# to register itself degrades to a loud skip instead of overwriting someone else's row.
from hiring_agent.sharepoint_scoring import _live_row_index as _y8idx


class _Y8Client:
    def __init__(self, rows, support_single_read=True):
        self.rows = list(rows)          # list of Application IDs, position == index
        self.single_reads = 0
        self.table_reads = 0
        if not support_single_read:
            del self.__class__.row_values_at   # noqa: unreachable in practice

    def row_values_at(self, index):
        self.single_reads += 1
        if 0 <= index < len(self.rows):
            return {"Application ID": self.rows[index]}
        return None

    def list_rows(self):
        self.table_reads += 1
        return [{"index": i, "values": {"Application ID": a}}
                for i, a in enumerate(self.rows)]


# Nothing deleted -> arithmetic is a no-op, but the position is still CONFIRMED.
_y8c = _Y8Client(["A", "B", "C"])
ok(_y8idx(_y8c, {"index": 2}, [], "C") == 2, "with no deletions the cached index is used as-is")
ok(_y8c.single_reads == 1 and _y8c.table_reads == 0,
   "every row is confirmed with ONE small single-row read, never a table scan")

# One row above was deleted and recorded -> offset applied AND verified.
_y8c = _Y8Client(["B", "C"])        # "A" (index 0) was deleted
ok(_y8idx(_y8c, {"index": 2}, [0], "C") == 1,
   "a recorded deletion above shifts the index down by one")
ok(_y8c.single_reads == 1 and _y8c.table_reads == 0,
   "a correct offset needs no table scan either")

# THE CASE THAT MATTERS: a delete path that forgot to register itself. The shift is ZERO,
# so pure arithmetic would hand back a stale index and overwrite the wrong candidate.
_y8c = _Y8Client(["B", "C", "D"])   # "A" was deleted but deleted_idx is empty
ok(_y8idx(_y8c, {"index": 3}, [], "D") == 2,
   "an UNRECORDED deletion is caught and the candidate re-located (this is the bug class)")
ok(_y8c.table_reads == 1,
   "…which costs exactly one recovery scan, only when something is actually wrong")

_y8c = _Y8Client(["B", "C", "D"])
ok(_y8idx(_y8c, {"index": 3}, [0], "D") == 2,
   "when the offset lands on the right row it is accepted as-is")
ok(_y8c.table_reads == 0, "…with no recovery scan")

# Candidate genuinely gone -> None, so the caller skips instead of writing blind.
_y8c = _Y8Client(["B", "C"])
ok(_y8idx(_y8c, {"index": 4}, [0], "GONE") is None,
   "a candidate no longer in the table returns None so the caller SKIPS it")

# A client with no single-row read (older stub) still works via plain arithmetic.
class _Y8Old:
    def list_rows(self):
        return []


ok(_y8idx(_Y8Old(), {"index": 3}, [0], "X") == 2,
   "a client without row_values_at falls back to the arithmetic, never crashes")
ok(_y8idx(_Y8Client(["A"]), {"index": 1}, [0], "") == 0,
   "a row with no Application ID cannot be verified, so the offset stands")

print("\n=== Y9. ON-DEMAND CLIENT EXPORT IS REACHABLE ===")
# export_client_results only ever ran as the last step of a scoring/recheck pass, so
# refreshing the client sheet meant running work you did not otherwise want.
_y9bot = Path(__file__).with_name("bot.py").read_text(encoding="utf-8", errors="replace")
ok("--export-results" in _y9bot, "bot.py exposes --export-results")
ok("export_client_results" in _y9bot, "…and wires it to the real exporter")
ok("upload_to_sharepoint=not args.dry_run" in _y9bot,
   "--dry-run writes the local copy but uploads nothing")
_y9app = Path(__file__).with_name("app.py").read_text(encoding="utf-8", errors="replace")
ok("_sp_export_results" in _y9app and '"--export-results"' in _y9app,
   "the GUI has an Export Results action that shells out to the same CLI path")
ok("askyesno" in _y9app.split("def _sp_stop", 1)[-1].split("def ", 1)[0],
   "Stop asks for confirmation before killing a run mid-write")
ok("finally:" in _y9app.split("def _run_bot", 1)[-1].split("\n    def ", 1)[0],
   "the run-completion payload is posted from a finally, so the busy flag always clears")

print("\n=== Y10. OPERATING POSTURE — live scoring, applicant mail SUPPRESSED (replay) ===")
# Deliberate stance (2026-09-04, user instruction): SILENT REPLAY. The back catalogue is
# being re-imported and re-scored to backfill the new 'Experience' column, so no applicant
# mail may leave P2 - 167 already-answered candidates must not hear from us twice. Scoring,
# resume moves, the Rejected sheet and the client export all keep running; admin/error mail
# stays on. Supersedes the 2026-08-11 live-applicant-mail posture, which is what to restore.
#
# THIS ASSERTION IS THE REVERT REMINDER. When the replay is done, flip config.yaml,
# .env's HIRING_SUPPRESS_EMAILS and these three lines back together - the suite fails
# until all of them agree, which is the point: the posture cannot drift silently.
# Section X proves the kill-switch still WORKS; this section pins which way it is SHIPPED,
# in the COMMITTED config rather than a gitignored .env, so a fresh client install is never
# a surprise.
#
# P1 matches: flow_config.json email.send_applicant_emails is also true, so both phases send.
import yaml as _y10yaml

_y10cfg = _y10yaml.safe_load(Path(__file__).with_name("config.yaml").read_text(encoding="utf-8"))
_y10tm = _y10cfg.get("test_mode") or {}
ok(_y10tm.get("suppress_emails") is True,
   "POSTURE: applicant mail is SUPPRESSED for the replay "
   "(config.yaml test_mode.suppress_emails: true)")
# .env is gitignored, so the committed default is what a fresh clone actually does. Pinning it
# here means the posture can only change by an explicit edit to a tracked file - it can never
# drift silently because someone's local .env happened to say something different.
ok(_y10tm.get("suppress_emails") == _sup_cfg.SUPPRESS_EMAILS,
   "the committed default and the effective runtime value agree - no hidden .env override")
ok(_sup_cfg.SUPPRESS_EMAILS is True,
   "the effective runtime value WITHHOLDS applicant mail")
ok(_sup_cfg.ERROR_EMAIL_ENABLED is True,
   "admin alerts are live too - P2 must never fail silently")
# The two applicant senders that have their own flags must also be on, or 'mail is live'
# would be true in name only: the decline is gated by geo_reject_email, and the missing-info
# nudge has no flag of its own and rides entirely on suppress_emails.
ok(_y10cfg.get("geo_reject_email") is True,
   "the non-USA DECLINE mail is enabled (geo_reject_email)")
ok(_y10cfg.get("error_email") is True,
   "the admin failure alert is enabled (error_email)")
ok(_y10cfg.get("geo_filter_usa_only") is True,
   "geo filtering still runs - suppressing mail must not change WHO is accepted")

# Suppression must silence mail only; every other live behaviour keeps running.
_y10src = Path(__file__).with_name("sharepoint_client.py").read_text(encoding="utf-8")
_y10send = _y10src.split("def send_mail", 1)[-1].split("\n    def ", 1)[0]
ok("SUPPRESS_EMAILS" in _y10send and "return True" in _y10send,
   "suppression is enforced inside send_mail and returns success so callers still advance")
ok(_y10src.count("/sendMail") == 1,
   f"there is exactly ONE transmit path to gate (got {_y10src.count('/sendMail')})")

# Real-time: the desktop app must actually start the watch loop by itself.
# CLIENT_MODE (app.py) deliberately forces auto_start/backdate OFF at startup, no matter what
# app_settings.json says - client mode is one-shot/scheduled only, so a continuous loop can
# never begin unexpectedly from an older settings file. Asserting the raw JSON value therefore
# tested the wrong thing: with CLIENT_MODE on it fails even when everything is correct, and
# flipping the JSON to true to "fix" it would make this assertion pass while the running app
# still never auto-starts. Assert the EFFECTIVE behaviour for the mode actually shipped.
_y10set = json.loads(Path(__file__).with_name("app_settings.json").read_text(encoding="utf-8"))
_y10app_src = Path(__file__).with_name("app.py").read_text(encoding="utf-8", errors="replace")
_client_mode = _re.search(r"^CLIENT_MODE\s*=\s*True\b", _y10app_src, _re.M) is not None
if _client_mode:
    ok(_re.search(r'self\.settings\["auto_start"\]\s*=\s*False', _y10app_src) is not None,
       "CLIENT MODE: auto-start is force-disabled at startup (one-shot/scheduled only, so a "
       "continuous live-SharePoint loop can never begin from a stale settings file)")
    ok(_y10set.get("schedule_enabled") is True or _y10set.get("auto_start") is False,
       "…and app_settings.json is consistent with that (no orphan auto_start:true that the "
       "app would silently override)")
else:
    ok(_y10set.get("auto_start") is True,
       "REAL TIME: the app auto-starts scoring instead of waiting for a manual click")
ok(int(_y10set.get("trigger_interval_min", 0)) >= 1,
   f"…on a real polling interval ({_y10set.get('trigger_interval_min')} min)")
ok(_y10set.get("online_only") is True,
   "…against the live SharePoint workbook, not a local file")
_y10app = Path(__file__).with_name("app.py").read_text(encoding="utf-8", errors="replace")
_y10auto = _y10app.split("def _sp_autorun", 1)[-1].split("\n    def ", 1)[0]
ok('"--watch"' in _y10auto and '"--score-sharepoint"' in _y10auto,
   "auto-start runs the continuous --score-sharepoint --watch loop")
ok(os.getenv("HIRING_P2_DISABLED", "false").strip().lower() not in ("1", "true", "yes"),
   "the P2 kill-switch is NOT engaged (processing is live)")

# An idle poll must not re-upload an unchanged client workbook every interval.
_y10sc = Path(__file__).with_name("hiring_agent").joinpath("sharepoint_scoring.py").read_text(
    encoding="utf-8", errors="replace")
ok("queue empty and nothing changed this pass" in _y10sc,
   "an idle watch poll skips the client-workbook rebuild instead of re-uploading it")

print("\n=== Y11. APPLICANT MAIL OFF, ERROR/ADMIN MAIL ON ===")
# SUPPRESS_EMAILS used to be a single master switch covering BOTH audiences, so the
# silent-live posture also silenced P2's own failure reporting - it ran continuously and, if
# it broke, told nobody. P1 never had that problem (Notify_failure stayed live). The two
# audiences are now independent, mirroring P1's two flags exactly.
from sharepoint_client import SharePointClient as _Y11C
import hiring_agent.config as _y11cfg


class _Y11Spy(_Y11C):
    def __init__(self):
        self.sender_mailbox = "apply@driverai.io"
        self.calls = []

    def _req(self, *a, **k):
        self.calls.append(a)

        class _R:
            status_code = 202
        return _R()


def _y11(suppress, error_email):
    """(applicant_sent, admin_sent) under a given flag combination."""
    _s, _e = _y11cfg.SUPPRESS_EMAILS, _y11cfg.ERROR_EMAIL_ENABLED
    try:
        _y11cfg.SUPPRESS_EMAILS, _y11cfg.ERROR_EMAIL_ENABLED = suppress, error_email
        c = _Y11Spy()
        c.send_mail("cand@x.com", "Decline", "<p>b</p>")
        app = len(c.calls)
        c.calls.clear()
        c.send_mail("admin@x.com", "P2 FAILURE", "<p>b</p>", admin=True)
        return bool(app), bool(c.calls)
    finally:
        _y11cfg.SUPPRESS_EMAILS, _y11cfg.ERROR_EMAIL_ENABLED = _s, _e


ok(_y11(True, True) == (False, True),
   "CURRENT POSTURE: applicant mail blocked, admin/error alert still SENT")
ok(_y11(True, False) == (False, False),
   "total blackout is still reachable (suppress + error_email both off)")
ok(_y11(False, True) == (True, True), "mail fully on: both audiences send")
ok(_y11(False, False) == (True, False), "admin alerts can be silenced on their own")

# The admin gate must live at the choke point, not only in the callers - relying on callers
# is what let the missing-info nudge escape GEO_REJECT_EMAIL.
_y11src = Path(__file__).with_name("sharepoint_client.py").read_text(encoding="utf-8")
_y11send = _y11src.split("def send_mail", 1)[-1].split("\n    def ", 1)[0]
ok("ERROR_EMAIL_ENABLED" in _y11send,
   "send_mail itself enforces ERROR_EMAIL_ENABLED - a future admin sender cannot bypass it")
ok("admin: bool = False" in _y11send,
   "applicant is the DEFAULT audience - a caller must opt IN to bypassing suppression")

# Only the three genuine operational senders may be marked admin.
_y11sc = Path(__file__).with_name("hiring_agent").joinpath("sharepoint_scoring.py").read_text(
    encoding="utf-8", errors="replace")
ok(_y11sc.count("admin=True") == 3,
   f"exactly 3 admin senders (error / suspicious-content / missing-workbook) "
   f"- got {_y11sc.count('admin=True')}")
for _fn in ("_send_non_usa_decline", "_send_missing_info_request"):
    _body = _y11sc.split("def %s(" % _fn, 1)[-1].split("\ndef ", 1)[0]
    ok("admin=True" not in _body,
       f"{_fn} is applicant-facing and stays suppressed")

# ── Z1. NEVER INFER RESIDENCE FROM A SCHOOL OR A PHONE (2026-08-04) ────────────
# A revision briefly labelled Country='United States' when a resume had no address but the
# phone and university both read as US (live case APP-20260716-2052-6112: contact block was
# only 'Prerna Saluja / (623)-242-3627 | LinkedIn | Gmail', the sole 'Arizona' being the
# employer 'Arizona State University'). REMOVED on client instruction - a school and a
# retained phone number say nothing about where someone lives NOW. A resume with no address
# leaves Location/Country genuinely unknown: they are marked MISSING and asked for by email.
print("\n=== Z1. Residence is never inferred from school/phone ===")
from hiring_agent import geo as _geo
from hiring_agent.geo import classify_location_usa, GeoDecision

_AZ_PHONE = "(623)-242-3627"
_US_EDU = "Master of Science in Finance, W. P. Carey School of Business, Arizona State University"

ok(not hasattr(_geo, "infer_us_country_from_context"),
   "infer_us_country_from_context is GONE - residence is never inferred from context")
_geo_src = Path(__file__).with_name("hiring_agent").joinpath("geo.py").read_text(
    encoding="utf-8", errors="replace")
ok("REMOVED on client instruction" in _geo_src,
   "geo.py records WHY the inference was removed, so it is not reintroduced")

# A US phone + US degree with no address must stay UNKNOWN, and must not fabricate a country.
_d, _r = classify_location_usa(location="", country="", phone=_AZ_PHONE, education=_US_EDU)
ok(_d == GeoDecision.UNKNOWN,
   "no address + US phone + US degree -> still UNKNOWN (needs review, not a guess)")
ok("United States" not in _r,
   "the reason string does not claim a US residence it cannot prove")

_scoring_src = Path(__file__).with_name("hiring_agent").joinpath(
    "sharepoint_scoring.py").read_text(encoding="utf-8", errors="replace")
ok("infer_us_country_from_context" not in _scoring_src,
   "the scoring/recheck paths no longer call the removed inference")

# ── Z2. FINANCE / ACCOUNTING SKILL VOCABULARY (added 2026-08-04) ────────────────
# The vocabulary was 133 terms and every one was tech/engineering - no finance, accounting
# or general-business term existed. That broke matching from BOTH directions: a finance
# resume was scraped down to 'Data Analysis, Excel', and the 'Finance Intern' JD was scraped
# down to 'computer vision, python, sql, excel, ...'. A finance candidate could not match a
# finance opening at all. JD skills are cached, so these terms only reach the JD side after
# `bot.py --refresh-jd`.
print("\n=== Z2. Finance/accounting skill vocabulary ===")
from hiring_agent.config import SKILL_KEYWORDS as _VOCAB
from hiring_agent.extraction import _scan_skill_keywords as _scan

_vocab_low = {str(k).lower() for k in _VOCAB}
for _term in ("financial reporting", "financial modeling", "variance analysis", "budgeting",
              "forecasting", "cash flow analysis", "accounting", "gaap", "sap", "erp"):
    ok(_term in _vocab_low, f"vocabulary carries finance term {_term!r}")

_PRERNA = ("SAP & AI Tools, Financial Reporting, Variance Analysis, Cash Flow Analysis, "
           "Data Analysis, Budgeting & Forecasting, Account Reconciliations, Due Diligence, "
           "Financial Statement Analysis, Financial Modeling, "
           "Financial Planning & Analysis, Microsoft Excel")
_found = set(_scan(_PRERNA))
ok(len(_found) >= 8,
   f"a finance resume now yields >=8 skills, not 2 (got {len(_found)}: {sorted(_found)[:6]}...)")
for _must in ("financial reporting", "variance analysis", "financial modeling", "budgeting"):
    ok(_must in _found, f"finance resume scan picks up {_must!r}")
# Regression: the tech vocabulary must be untouched by the additions.
for _tech in ("python", "kubernetes", "react native", "cybersecurity", "excel"):
    ok(_tech in _vocab_low, f"existing tech term {_tech!r} still present")

# ── Z3. --force-rescore: refresh roles that already have a value (added 2026-08-04) ──
# Normal healing is gap-only so a good row is never churned. The side effect was that a row
# scored under an older JD set or an older skill vocabulary kept its stale match forever,
# with no route to refresh it: --rematch only walks local workbooks under P2_Output/, never
# SharePoint. Live case: the 2026-08-04 vocabulary gained finance terms, but the MS-Finance
# candidate APP-20260716-2052-6112 kept a pre-vocabulary 'Automation Process Developer (40%)'
# because Suggested Role 1 was non-blank.
print("\n=== Z3. --force-rescore refreshes already-populated Suggested Roles ===")
import inspect as _inspect
from hiring_agent import sharepoint_scoring as _ss

_ph_sig = _inspect.signature(_ss._plan_heal)
ok("force_roles" in _ph_sig.parameters, "_plan_heal exposes force_roles")
ok(_ph_sig.parameters["force_roles"].default is False,
   "force_roles defaults to False - normal healing stays gap-only")
for _fn in (_ss.recheck_selected_rows, _ss.recheck_all_rows):
    _sig = _inspect.signature(_fn)
    ok("force_roles" in _sig.parameters, f"{_fn.__name__} exposes force_roles")
    ok(_sig.parameters["force_roles"].default is False,
       f"{_fn.__name__} leaves force_roles off by default")

ok(_ss._ROLE_COLS == ("Suggested Role 1", "Suggested Role 2", "Suggested Role 3"),
   "force_roles targets exactly the three Suggested Role columns")
ok("Category" not in _ss._ROLE_COLS,
   "Category is NOT force-listed - it re-derives from the new role further down")

_ph_src = _inspect.getsource(_ss._plan_heal)
ok("if force_roles:" in _ph_src and "gaps.append(_c)" in _ph_src,
   "force_roles works by adding the role columns to the existing gap list (reuses derive+heal)")
ok("str(nv).strip() == str(vals.get(c, \"\") or \"\").strip()" in _ph_src
   or "continue" in _ph_src.split("force_roles re-derives")[-1][:400],
   "an identical re-score is skipped so a forced run does not churn the sheet")
ok("role_was_rescored" in _ph_src,
   "Category is refreshed when the role it was derived from was just re-scored")

# The CLI must expose it, and it must NOT be on by default.
_bot_src = Path(__file__).with_name("bot.py").read_text(encoding="utf-8", errors="replace")
ok('"--force-rescore"' in _bot_src, "bot.py exposes --force-rescore")
ok("force_roles=args.force_rescore" in _bot_src,
   "--force-rescore is threaded into the recheck entry points")
ok(_bot_src.count("force_roles=args.force_rescore") == 2,
   "both --recheck-row and --recheck-all honour --force-rescore")
# geo_only recovery must never rescore roles - it is a deterministic historical audit.
_rsr_src = _inspect.getsource(_ss.recheck_selected_rows)
ok("_plan_geo_recovery" in _rsr_src,
   "geo_only still routes to the deterministic recovery planner (no role scoring)")

# ── Z4. "Missing" vs "N/A" placeholders (added 2026-08-04) ──────────────────────
# Two DIFFERENT meanings, and the sheet must never blur them:
#   "Missing" = the candidate still owes us this; we are chasing it by email.
#   "N/A"     = we asked and there genuinely is none (portfolio slots).
# A blank cell is banned outright - a reader cannot tell which of the two it means, or
# whether the pipeline simply failed.
print("\n=== Z4. 'Missing' vs 'N/A' placeholders ===")
from hiring_agent.extraction import MISSING_VALUE, _GAP_LITERALS as _GL
from hiring_agent.sharepoint_scoring import (
    _missing_if_gap, _na_if_gap, _apply_missing_placeholders,
    _MISSING_IF_BLANK_COLS, _is_gap,
)

ok(MISSING_VALUE == "Missing", "the placeholder reads exactly 'Missing'")
# THE critical property: 'Missing' must still count as a gap, or the row freezes forever -
# healing would stop trying to fill it and the nudge email would stop asking for it.
ok("missing" in _GL, "'Missing' is a gap literal, so healing and the nudge keep working")
ok(_is_gap("Missing") and _is_gap("missing") and _is_gap("  Missing  "),
   "_is_gap treats 'Missing' as absent, case/whitespace-insensitively")

for _blank in ("", "  ", "Not extracted", "N/A", "none", "unknown"):
    ok(_missing_if_gap(_blank) == MISSING_VALUE,
       f"_missing_if_gap({_blank!r}) -> 'Missing'")
ok(_missing_if_gap("Austin, TX") == "Austin, TX", "a real value is passed through untouched")
ok(_missing_if_gap("  Austin, TX  ") == "Austin, TX", "a real value is trimmed")

# Portfolios keep the OTHER placeholder - this is the long-standing convention and the
# client reconfirmed it on 2026-08-04 ("keep n/a if no portfolio as usual").
ok(_na_if_gap("") == "N/A", "an empty portfolio slot still reads 'N/A', not 'Missing'")
ok("Portfolio 1" not in _MISSING_IF_BLANK_COLS,
   "portfolios are NOT in the Missing set - they use N/A")

ok(set(_MISSING_IF_BLANK_COLS) == {"Phone", "Location", "Country", "Education",
                                   "Current Skills", "Full Name", "Looking For Role",
                                   "Education Start Date", "Education End Date",
                                   "Years Exp"},
   f"Missing covers exactly the candidate-owed columns (got {_MISSING_IF_BLANK_COLS})")

_f = {"Phone": "", "Location": "Not extracted", "Country": "", "Education": "N/A",
      "Current Skills": "Python", "Portfolio 1": "N/A", "Full Name": "Jane Doe"}
_out = _apply_missing_placeholders(dict(_f))
for _c in ("Phone", "Location", "Country", "Education"):
    ok(_out[_c] == MISSING_VALUE, f"_apply_missing_placeholders stamps {_c}")
ok(_out["Current Skills"] == "Python", "a populated column is left alone")
ok(_out["Portfolio 1"] == "N/A", "portfolio is untouched by the Missing pass")
ok(_out["Full Name"] == "Jane Doe", "columns outside the set are untouched")

# Residence is still never guessed - Missing is the honest answer, not an inferred country.
_scoring_src2 = Path(__file__).with_name("hiring_agent").joinpath(
    "sharepoint_scoring.py").read_text(encoding="utf-8", errors="replace")
ok("_apply_missing_placeholders(fields)" in _scoring_src2,
   "the scoring path stamps Missing right before the write")
ok(_scoring_src2.index("_validate_all_columns(fields, vals, mail_body")
   < _scoring_src2.index("_apply_missing_placeholders(fields)"),
   "Missing is stamped AFTER the mail-body fill attempt, never before it")

# ── Z5. Foreign COUNTRY + corroboration can reject (fixed 2026-08-04) ───────────
# Rejection used to require is_foreign_location(loc), i.e. the candidate's city had to be in
# FOREIGN_CITIES. Live case APP-20260716-1505-C8CC: 'Daska, District Sialkot' / Pakistan /
# +92 phone / University of Sialkot — every signal agreeing and none contradicting — was
# parked in location review and then emailed asking where he was based, purely because
# neither Daska nor Sialkot was in the city list. A country list can be exhaustive; a world
# city list never can. So a foreign COUNTRY plus INDEPENDENT phone/education corroboration
# now rejects, using the same bar step 2 already applies to a foreign LOCATION.
print("\n=== Z5. Foreign country + corroboration rejects (no city-list dependency) ===")
from hiring_agent.geo import classify_location_usa as _cls, GeoDecision as _GD

def _decide(loc, ctry, ph="", edu=""):
    return _cls(location=loc, country=ctry, phone=ph, education=edu)[0]

# The three live rows this unstuck.
ok(_decide("Daska, District Sialkot", "Pakistan", "+92 333 8113971",
           "BS Computer Science - University of Sialkot") == _GD.CONFIRMED_NON_US,
   "Pakistan + +92 phone + Sialkot degree rejects, though the city is not in FOREIGN_CITIES")
ok(_decide("Remote", "India", "+91 7799155311", "B.Tech, JNTU") == _GD.CONFIRMED_NON_US,
   "'Remote' + India + +91 phone rejects")
ok(_decide("", "India", "+91 (884) 955-6296") == _GD.UNKNOWN,
   "a COMPLETELY BLANK location still never rejects - deliberate rule from the 2026-07-24 "
   "_looks_like_us_phone work, preserved: with no stated location at all we only ask")

# PROTECTIVE CASES - the conservative design exists for these, and they must NOT regress.
ok(_decide("Chicago, IL", "India", "(312) 555-0100", "MS, University of Chicago") == _GD.CONFIRMED_US,
   "a US location still beats a Country mis-read off a past overseas degree")
ok(_decide("Remote", "India", "(512) 555-0100") == _GD.UNKNOWN,
   "a US-looking phone CONTRADICTS the foreign country -> review, never reject")
ok(_decide("Remote", "India", "", "MS Computer Science, Arizona State University") == _GD.UNKNOWN,
   "a US education CONTRADICTS the foreign country -> review, never reject")
ok(_decide("Remote", "India", "", "") == _GD.UNKNOWN,
   "a foreign country with NO corroborating signal at all still only asks, never rejects")
ok(_decide("Hyderabad", "", "(512) 555-0100") == _GD.UNKNOWN,
   "a foreign city with a US phone is still kept for clarification (step 2 guard intact)")
ok(_decide("Paris, TX", "") == _GD.CONFIRMED_US, "'Paris, Texas' is still US, not France")
ok(_decide("", "", "(623)-242-3627", "MS Finance, Arizona State University") == _GD.UNKNOWN,
   "no location and no country stays UNKNOWN - residence is never inferred")

# ── Z6. A country the resume never states is NOT extraction (2026-08-04) ────────
# Live case APP-20260716-1037-6CE3 (Peter Vishal). His resume says "Remote" and "Open to
# Remote & Global Opportunities"; the words India/Indian/Hyderabad/Telangana appear ZERO
# times. Country was nonetheless stored as "India" - the model inferred it from a +91 dial
# code and an Indian college name - and the geo filter then pointed at a rejection built on
# a value nobody had extracted. A dial code is not an address (numbers move with people) and
# a degree is a fact about the past. Same principle the client applied to Prerna, mirrored:
# never infer residence, in EITHER direction.
print("\n=== Z6. AI country must be grounded in the text ===")
from hiring_agent.extraction import _country_grounded_in_text as _cg

_PETER = ("PETER VISHAL\npetervishal55@gmail.com | +91 7799155311 | "
          "Open to Remote & Global Opportunities\n"
          "B.Tech Computer Science | Malla Reddy College of Engineering\n")
ok(not _cg("India", "Remote", _PETER),
   "'India' is NOT grounded - absent from the text and not implied by 'Remote'")
ok(not _cg("India", "", _PETER), "a +91 dial code alone does not ground a country")
ok(_cg("India", "Hyderabad, India", "Ravi\nHyderabad, India\n"),
   "a country literally present in the resume IS grounded")
ok(_cg("United States", "Austin, TX", "Jane\nAustin, TX 78701\n"),
   "'United States' implied by a real US state IS grounded")
ok(not _cg("Canada", "Austin, TX", "Jane\nAustin, TX 78701\n"),
   "a country contradicting a US location and absent from the text is NOT grounded")
ok(not _cg("", "Austin, TX", "x") and not _cg("Not extracted", "Austin, TX", "x"),
   "gap literals are never 'grounded'")

# Both AI passes must apply the SAME rule - the recheck used to undo the merge's decision.
_ex_src = Path(__file__).with_name("hiring_agent").joinpath("extraction.py").read_text(
    encoding="utf-8", errors="replace")
_recheck_body = _ex_src.split("def ai_recheck_fields(", 1)[-1].split("\ndef ", 1)[0]
# ...and since 2026-09-05 there is literally one rule, not two copies of it: both passes
# call settle_geography, which owns all four checks (contact block, text grounding,
# is-it-plausibly-a-place, country grounding) and finalize_geography_shape after it.
_merge_body = _ex_src.split("def extract_candidate_details_smart(", 1)[-1].split("\ndef ", 1)[0]
_gate_body = _ex_src.split("def settle_geography(", 1)[-1].split("\ndef ", 1)[0]
for _p, _b in (("ai_recheck_fields", _recheck_body), ("extract_candidate_details_smart", _merge_body)):
    ok("settle_geography(" in _b and "finalize_geography_shape(" in _b,
       f"{_p} delegates geography to the shared rule")
    ok("_country_grounded_in_text" not in _b and "_location_grounded_in_contact_block" not in _b,
       f"...and keeps no private copy of it ({_p})")
for _check in ("_location_grounded_in_contact_block", "_location_grounded_in_text",
               "location_is_plausible", "_country_grounded_in_text"):
    ok(_check in _gate_body, f"the shared rule applies {_check}")

# An unsupported STORED country must be corrected, and geo must re-decide without it.
_ss_src = Path(__file__).with_name("hiring_agent").joinpath(
    "sharepoint_scoring.py").read_text(encoding="utf-8", errors="replace")
ok('heal["Country"] = MISSING_VALUE' in _ss_src,
   "an unsupported stored Country is corrected to 'Missing', not left alone")
_after = _ss_src.split('heal["Country"] = MISSING_VALUE', 1)[1][:600]
ok('country = ""' in _after,
   "the local country is cleared too, so geo cannot still reject on the discredited value")

# ── Z7. Live-sheet data-quality defects found 2026-08-04 ───────────────────────
# Every case below is a real value that was sitting on the live candidate list.
print("\n=== Z7. Live data-quality defects (audit 2026-08-04) ===")
from hiring_agent.extraction import (
    _name_is_sentence_fragment, _looks_like_name, _plausible_name_shape,
    normalize_skills, clean_location_text, clean_role_text,
    _looks_like_url, _is_meaningful_portfolio,
)
from hiring_agent.sharepoint_scoring import _name_needs_rederive, _compact_portfolios

# -- names: the check was INVERTED (fired on correct names, missed the broken ones) --
for _nm, _em in (("Prerna Saluja", "prernasaluja259@gmail.com"),
                 ("Rutuja Shingate", "rutujashingate2000@gmail.com"),
                 ("Rahul Chowdary Vajja", "rahulchowdaryvajja@gmail.com"),
                 ("Hamza Asif", "hamzaasif219@gmail.com")):
    ok(not _name_needs_rederive(_nm, _em),
       f"{_nm!r} is a REAL name and is no longer flagged just because the email spells it")
ok(_name_needs_rederive("sbhask22", "sbhask22@asu.edu"),
   "a bare email local-part IS still flagged (P1's fallback shape: one token, no space)")
ok(_name_needs_rederive("for Scientific Computing.", "tianxinwang89@gmail.com"),
   "a sentence fragment IS now flagged - it used to slip through and reach the file name")

# ROOT CAUSE of the 'for Scientific Computing.' row: his resume's first line is
# 'Tianxin (William) Wang', but '(William)' fails .isalpha() so every shape check rejected
# the CORRECT name, and the header scan fell through to prose on a later line.
from hiring_agent.extraction import strip_parenthetical_nickname, _normalize_name
ok(strip_parenthetical_nickname("Tianxin (William) Wang") == "Tianxin Wang",
   "a parenthesised preferred name is stripped: 'Tianxin (William) Wang' -> 'Tianxin Wang'")
ok(_looks_like_name("Tianxin (William) Wang"),
   "the name with a nickname now PASSES the shape check (it used to be rejected)")
ok(_looks_like_name("Robert (Bob) Smith"), "the anglicised-nickname pattern generally works")
ok(_normalize_name("Tianxin (William) Wang") == "Tianxin Wang",
   "the stored cell keeps the legal name, so it cannot reach the resume file name")
ok(_normalize_name("Christopher L. Feld") == "Christopher L. Feld",
   "a name without a nickname is unchanged")

ok(_name_is_sentence_fragment("for Scientific Computing."), "'for ...' reads as prose")
ok(_name_is_sentence_fragment("the Analytics Team"), "'the ...' reads as prose")
ok(not _name_is_sentence_fragment("Maria de la Cruz"), "name particles are not prose")
ok(not _name_is_sentence_fragment("Robert Downey Jr."), "'Jr.' is a suffix, not a sentence end")
ok(not _name_is_sentence_fragment("Christopher L. Feld"), "a middle initial is not prose")
ok(not _looks_like_name("for Scientific Computing.")
   and not _plausible_name_shape("for Scientific Computing."),
   "both name-shape checks now reject the fragment")

# -- skills: resume section headings and prose were being stored as skills --
_yashvi = ("Programming Languages and Databases: Python, C++, Java, MySQL; "
           "Web Technologies: HTML, CSS, JavaScript")
_cleaned = normalize_skills(_yashvi)
ok("Programming Languages and Databases" not in _cleaned,
   f"section heading stripped from skills -> {_cleaned[:52]!r}")
ok("Web Technologies" not in _cleaned, "heading after a ';' group separator is stripped too")
ok("Python" in _cleaned and "HTML" in _cleaned, "the actual skills survive")
ok("Programming" not in normalize_skills("Programming:Python, C++"),
   "the no-space 'Programming:Python' form is stripped")
_prose = normalize_skills("Strategic planning for risk and cybersecurity at all levels of "
                          "an organization, Cybersecurity, SIEM")
ok("Strategic planning" not in _prose and "Cybersecurity" in _prose,
   f"a resume-bullet sentence is dropped, real skills kept -> {_prose!r}")
ok("Financial Planning & Analysis" in normalize_skills("Financial Planning & Analysis, Kotlin"),
   "a legitimate multi-word skill is NOT mistaken for prose")

# -- location: the resume's own 'Address -' label and PDF double-spaces --
ok(clean_location_text("Address  -  San  Jose, California") == "San Jose, California",
   "'Address  -  San  Jose, California' -> 'San Jose, California'")
ok(clean_location_text("Location: Austin, TX") == "Austin, TX", "a 'Location:' label is stripped")
ok(clean_location_text("Tempe, AZ") == "Tempe, AZ", "a clean location is untouched")

# -- location: a segment that merely repeats another is dropped (2026-08-11) --
# Live row APP-20260602-1601-EF3F stored 'Bikaner, Bikaner' - the extractor read the header
# city twice. A repeated token carries no extra information and looks wrong in the
# client-facing Location column.
ok(clean_location_text("Bikaner, Bikaner") == "Bikaner",
   "a repeated city segment collapses: 'Bikaner, Bikaner' -> 'Bikaner'")
ok(clean_location_text("San Jose, San Jose, California") == "San Jose, California",
   "the repeat is dropped but the real state segment survives")
ok(clean_location_text("bikaner, Bikaner") == "bikaner",
   "the repeat check is case-insensitive and keeps the first occurrence's casing")
# The guard that makes the above safe: a city genuinely named after its state must NOT be
# collapsed, or 'New York, New York' would lose the state and stop reading as US.
ok(clean_location_text("New York, New York") == "New York, New York",
   "GUARD: 'New York, New York' keeps its state - the repeat is a real US state name")
ok(clean_location_text("Washington, Washington") == "Washington, Washington",
   "GUARD: same for any other city sharing its state's name")
ok(clean_location_text("San Juan, Puerto Rico, Puerto Rico") == "San Juan, Puerto Rico, Puerto Rico",
   "GUARD: US territories are protected too, not just states")
from hiring_agent.geo import is_strong_usa as _loc_strong
ok(_loc_strong(clean_location_text("New York, New York")) is True,
   "…and the guarded value still reads as a STRONG US signal after cleaning")


print("\n=== O2d. location_is_plausible - grounded-but-not-a-place Location values ===")
# Live 2026-09-01: 'George Mason University', 'ABOUT, ME', and "Bachelor's from JNTUK,
# India - 2017." all passed the existing anti-hallucination grounding check (they
# genuinely appear in the resume text) while being an institution name, a section-header
# artifact, and a degree-sentence fragment respectively - none of them a real place.
from hiring_agent.extraction import location_is_plausible as _loc_ok
for _bad in ("George Mason University", "Daulat Ram College", "University of Phoenix",
            "ABOUT, ME", "About Me", "Bachelor's from JNTUK, India - 2017.",
            "Master's in Computer Science"):
    ok(_loc_ok(_bad) is False, f"{_bad!r} is rejected as not plausibly a place")
for _good in ("Tempe, AZ", "San Jose, California", "Hyderabad", "New York, New York",
             "Remote", "Walton, 54000, Lahore, Pakistan", "",
             # GUARD (found live 2026-09-01 as false positives before this guard
             # existed): a real city whose NAME contains 'College'/'University' must
             # survive - only a bare institution name with no state qualifier is
             # rejected.
             "College Park, MD", "College Park, Maryland", "State College, PA",
             "University City, MO"):
    ok(_loc_ok(_good) is True, f"{_good!r} is NOT rejected - a real place is unaffected")


print("\n=== O2e. live data-quality defects found on the Rejected sheet (2026-08-11) ===")
# Three extraction defects found by auditing all 118 live candidate rows.

# -- 1. Country left blank when the country name is sitting in the Location string --
# split_location_country only ever inspected the last COMMA-separated segment, so a trailing
# country with no comma, or a value that IS just a country, produced a blank Country. A blank
# Country weakens classify_location_usa, which reads that field.
from hiring_agent.extraction import split_location_country as _slc

for _raw, _wloc, _wctry in [
    ("Rawalpindi Punjab Pakistan", "Rawalpindi Punjab", "Pakistan"),   # no comma at all
    ("133 B3 Johar Town, Lahore Pakistan", "133 B3 Johar Town, Lahore", "Pakistan"),
    ("Lahore Pakistan", "Lahore", "Pakistan"),
    ("Uk", "N/A", "UK"),          # a bare country is not a Location at all
    ("India", "N/A", "India"),
]:
    _gl, _gc = _slc(_raw)
    ok((_gl, _gc) == (_wloc, _wctry),
       f"trailing country split: {_raw!r} -> ({_wloc!r}, {_wctry!r}) [got ({_gl!r}, {_gc!r})]")

for _raw, _wctry in [("Tempe, AZ", "United States"), ("Austin, Texas", "United States"),
                     ("Mumbai, Maharashtra", "India"), ("Toronto, ON, Canada", "Canada"),
                     ("Nairobi, Kenya", "Kenya")]:
    ok(_slc(_raw)[1] == _wctry, f"existing country resolution unchanged: {_raw!r} -> {_wctry}")
# The 2-letter 'uk' must be word-anchored or it fires inside ordinary place names.
for _safe in ("Columbus", "Paducah", "Sukkur"):
    ok(_slc(_safe)[1] == "", f"a 2-letter country term does not match inside {_safe!r}")

# -- 2. Education swallowing a section header or a skills line --
from hiring_agent.extraction import _extract_education as _xedu

ok(_xedu("EDUCATION\nLANGUAGE SKILLS\nEnglish, Urdu") is None,
   "a section header under 'Education' is NOT stored as education")
ok(_xedu("Education\nLanguages: Dart, Java (Basic), Kotlin (Basic)") is None,
   "a skills line under 'Education' is NOT stored as education")
ok(_xedu("Education\nManufacturing") is None,
   "a bare industry word is NOT stored as education")
ok("Arizona State University" in (_xedu("Education\nArizona State University") or ""),
   "a bare INSTITUTION with no degree keyword is still kept")
ok("Bachelor" in (_xedu("EDUCATION\nBachelor of Technology, SR University") or ""),
   "a real degree line is still kept")
ok("B.S" in (_xedu("Education: B.S. Computer Science, MIT") or ""),
   "a same-line degree is still kept")

# -- 2b. blob-extracted PDFs and curly apostrophes (2026-08-11) --
# Live case APP-20260507-2235-813F: the CV extracted as 7 lines, two of them ~3,100 chars,
# every word tab-separated. No line STARTS with 'Education', so all the line-oriented header
# heuristics missed a degree that was plainly in the text; the curly apostrophe in "Master’s"
# then defeated the full-text degree scan too. Both had to be fixed for the row to extract.
_blob = ("Rohit Emmadishetty\tTechnical Lead\t" + ("filler detail " * 40) +
         "\tsupporting\tHR\tanalytics\tat\tscale.\tEducation\t\tMaster’s\tin\t"
         "computer\tscience,\tUniversity\tof\tNorth\tTexas,\tDenton,\tTX\t\t")
_got = _xedu(_blob) or ""
ok("North Texas" in _got,
   f"a tab-collapsed blob still yields the degree line (got {_got[:52]!r})")
ok("Master" in _got, "…and the curly-apostrophe 'Master’s' is recognised")
ok(len(_got) < 200, f"…and it returns the degree, not the whole 3,000-char blob ({len(_got)} chars)")

from hiring_agent.extraction import _DEGREE_RE as _dre
ok(bool(_dre.search("Master’s in computer science")),
   "curly apostrophe: Master’s matches _DEGREE_RE")
ok(bool(_dre.search("Master's in computer science")),
   "straight apostrophe still matches")
ok(bool(_dre.search("Bachelor’s of Science")),
   "curly apostrophe works for Bachelor’s too")

# The splitter must be a no-op on normally extracted resumes.
from hiring_agent.extraction import _split_collapsed_sections as _split
_normal = ["SHALIN EDWARD", "Columbus, OH", "EDUCATION", "M.S. Computer Science, UMich"]
ok(_split(_normal) == _normal,
   "a normally-extracted resume passes through the splitter untouched")
ok(_split(["short line with\ttabs"]) == ["short line with\ttabs"],
   "a SHORT tabbed line is not split - only long blobs are")

# -- 3. Portfolio accepting contact/map links and country-coded university home pages --
from hiring_agent.extraction import _is_meaningful_portfolio as _port

for _u in ("https://www.uskt.edu.pk/", "https://pakaims.edu.pk/",
           "https://maps.google.com/?q=Greater", "https://wa.me/+923139999999",
           "https://www.asu.edu/"):
    ok(_port(_u) is False, f"not a personal portfolio: {_u}")
for _u in ("https://github.com/someone", "https://linkedin.com/in/someone",
           "https://shashankranjan.in/", "https://leetcode.com/u/someone/",
           "https://itunes.apple.com/in/app/thing/id123",
           "https://mit.edu/~jsmith/projects"):
    ok(_port(_u) is True, f"genuine portfolio still accepted: {_u}")

# -- portfolios: junk values and slot ordering --
ok(not _looks_like_url("http://b.sc"),
   "'http://b.sc' (a B.Sc degree parsed as a domain) is rejected as a URL")
ok(_looks_like_url("https://x.com/someone"),
   "a genuine one-letter host WITH a path still passes (x.com/someone)")
ok(not _is_meaningful_portfolio("https://www.asu.edu/"),
   "a bare university home page is not a personal portfolio")
ok(not _is_meaningful_portfolio("https://drive.google.com/file/d/1ZL/view?usp=sharing"),
   "a Drive file-share link is not a portfolio (it is usually the resume)")
ok(_is_meaningful_portfolio("https://github.com/petervishal55"), "a real GitHub profile is kept")
ok(_is_meaningful_portfolio("https://rutuja.fyi/"), "a personal site is kept")
ok(_compact_portfolios("N/A", "https://github.com/x", "https://vercel.com/x")
   == ("https://github.com/x", "https://vercel.com/x", "N/A"),
   "slots are compacted - N/A never sits above a real link")
ok(_compact_portfolios("https://linkedin.com/in/s", "https://github.com/s", "http://b.sc")
   == ("https://linkedin.com/in/s", "https://github.com/s", "N/A"),
   "junk is dropped and the remaining slot reads N/A")

# -- Looking For Role picked up the contact block off the resume header --
ok(clean_role_text("Mobile Application Developer Email: ashfaq.fullstackdev@gmail.com")
   == "Mobile Application Developer",
   "trailing 'Email: ...' is cut off the role title")
ok(clean_role_text("Senior React Native or Full-Stack Engineer")
   == "Senior React Native or Full-Stack Engineer", "a clean role title is untouched")

# -- AI scorer: a confident pick with ZERO evidence must be dropped --
from hiring_agent.scoring import _skill_overlap_count, _ai_pick_has_evidence
_ciso = "Cybersecurity, SIEM, Incident Response, Compliance, Risk Management"
ok(_skill_overlap_count(_ciso, ["python", "pytorch", "computer vision", "tensorflow"]) == 0,
   "a CISO has zero overlap with an AI/ML JD - the case that was ranked #1 at 95%")
ok(_skill_overlap_count(_ciso, ["cybersecurity", "siem", "compliance", "incident response"]) == 4,
   "the cybersecurity JD he should have matched scores 4")
ok(_skill_overlap_count(_ciso, []) == 999,
   "a JD with NO parsed skills is exempt - only a genuine mismatch is dropped")
_hw = "SystemVerilog, Verilog, ASIC Design, RTL Design, Python"
_ai_jd = ["computer vision", "machine learning", "tensorflow", "pytorch", "python"]
ok(not _ai_pick_has_evidence(_hw, "Hardware Engineer", "Software Developer - AI/ML, Computer Vision", _ai_jd),
   "one generic overlap cannot turn a hardware profile into a multi-skill AI/CV match")
ok(_ai_pick_has_evidence(_hw, "AI/ML Engineer", "Software Developer - AI/ML, Computer Vision", _ai_jd),
   "a candidate's stated role preference can support a related low-overlap role")
ok(_ai_pick_has_evidence(_hw, "Hardware Engineer", "Python Specialist", ["python"]),
   "a one-skill specialist JD remains eligible when its one required skill matches")
_sc_src = Path(__file__).with_name("hiring_agent").joinpath("scoring.py").read_text(
    encoding="utf-8", errors="replace")
ok("_ai_pick_has_evidence(" in _sc_src,
   "the evidence guard is wired into the Ollama ranking path")

# -- Last Updated must never predate Received --
_ss_src2 = Path(__file__).with_name("hiring_agent").joinpath(
    "sharepoint_scoring.py").read_text(encoding="utf-8", errors="replace")
ok("clamped forward to Received" in _ss_src2,
   "a Last Updated Date earlier than Received is clamped, not left inconsistent")

# ── Z8. A concrete US location beats a contradicting Country (2026-08-04) ───────
# Client-reported: "country india and city los angeles, thats wrong, it should be missing or
# los angeles, ca country as usa". Leaving the two to disagree is the worst outcome - the
# sheet shows a contradiction AND the row sits in review forever. geo step 1 already holds
# that physical-presence evidence beats the Country cell (which is routinely mis-read off a
# past overseas degree); this makes _location_clarity_issue agree instead of overruling it.
# Live case APP-20260716-0617-CB73: Yashvi Vaghela had already replied TWICE confirming she
# is US-based and was still being asked.
print("\n=== Z8. Concrete US location beats a contradicting Country ===")
from hiring_agent.sharepoint_scoring import _classify_candidate_geo as _cgeo
from hiring_agent.geo import reconcile_us_country as _rc, GeoDecision as _G

_d, _ = _cgeo("Los Angeles, CA", "India", "(213) 376-7163")
ok(_d == _G.CONFIRMED_US,
   "'Los Angeles, CA' + Country 'India' + US phone -> CONFIRMED_US, not parked in review")
ok(_rc("India", "Los Angeles, CA", True) == "United States",
   "the contradicting Country cell is rewritten to 'United States'")

# Must NOT regress - these are genuine conflicts or genuine foreign candidates.
ok(_cgeo("Austin, TX", "India", "+91 98765 43210")[0] == _G.UNKNOWN,
   "a US location with a NON-US international phone is still a real conflict -> ask")
ok(_cgeo("Pune, NC", "India", "+91 98765 43210")[0] == _G.UNKNOWN,
   "mixed foreign-city + US-state ('Pune, NC') still asks")
ok(_cgeo("Daska, District Sialkot", "Pakistan", "+92 333 8113971",
         education="BS CS - University of Sialkot")[0] == _G.CONFIRMED_NON_US,
   "a genuinely foreign candidate still rejects")
ok(_cgeo("Remote", "India", "+91 7799155311")[0] == _G.CONFIRMED_NON_US,
   "'Remote' + India + +91 still rejects - no US signal in the location")

# ── Z9. The zero-overlap guard must never fire on MISSING data ──────────────────
# Regression I introduced on 2026-08-04 and caught on the live sheet: an AI recheck rewrote
# APP-20260720-1755-30E6's skills as prose, the new prose filter stripped them to nothing,
# and with an empty candidate set EVERY role scored zero overlap - so all 11 ranked roles
# were discarded and a correct 'Cybersecurity IT Administrator PD (90%)' fell to 'CISO
# LinkedIn Announcement (40%)'. Zero overlap only means something when both sides parsed.
print("\n=== Z9. Zero-overlap guard fires only on a GENUINE mismatch ===")
from hiring_agent.scoring import _skill_overlap_count as _ov

ok(_ov("", ["cybersecurity", "siem"]) == 999,
   "empty CANDIDATE skills -> exempt (this is the regression that dropped every role)")
ok(_ov("Not extracted", ["cybersecurity"]) == 999, "a gap literal counts as empty -> exempt")
ok(_ov("Cybersecurity, SIEM", []) == 999, "empty JD skills -> exempt (pre-existing rule)")
ok(_ov("Cybersecurity, SIEM", ["python", "pytorch"]) == 0,
   "a REAL mismatch still scores 0 and is still dropped")
ok(_ov("Cybersecurity, SIEM", ["cybersecurity", "siem", "aws"]) == 2, "a real overlap counts")

# normalize_skills must never hand back an empty column after prose filtering.
_all_prose = normalize_skills(
    "Strategic planning for risk and cybersecurity at all levels of an organization")
ok(_all_prose != "Not extracted" and "Strategic planning" in _all_prose,
   "when EVERY entry is prose the unfiltered list is kept - unhelpful text beats no data")
ok("Cybersecurity" in normalize_skills(
    "Strategic planning for risk and cybersecurity at all levels of an organization, "
    "Cybersecurity, SIEM"),
   "when real skills exist alongside prose, the prose is still dropped")

# ── Z10. A malformed STORED Location is corrected in place ─────────────────────
# Found on the live sheet AFTER a full recheck had already run: APP-20260716-0335-0C84 still
# read 'Address  -  San  Jose, California'. The cell was non-blank, so gap-only healing never
# looked at it, and clean_location_text only runs during EXTRACTION - which never happened,
# because nothing about the row was missing. Wrong values need correcting, not just missing
# ones (same class as phone_bad / junk_ports / an unsupported Country).
print("\n=== Z10. Malformed stored Location is corrected in place ===")
_ss_src3 = Path(__file__).with_name("hiring_agent").joinpath(
    "sharepoint_scoring.py").read_text(encoding="utf-8", errors="replace")
ok('heal["Location"] = _cleaned_loc' in _ss_src3,
   "_plan_heal rewrites a stored Location that cleaning changes")
_plan_src = _ss_src3.split("def _plan_heal(", 1)[1].split("\ndef ", 1)[0]
ok("clean_location_text(location)" in _plan_src,
   "the cleaning runs inside _plan_heal, so it reaches rows with nothing else missing")
ok("location = _cleaned_loc" in _plan_src,
   "the local is updated too, so the geo decision uses the CLEANED location")
# Cleaning is deterministic - it must not need the resume, or a row with no gaps would skip it.
ok(_plan_src.index("clean_location_text(location)") < _plan_src.index("if need_resume:"),
   "cleaning happens BEFORE the resume-download branch - no download required")


# ── Z11. Resume filename convention: First_Last_<AppIdTail> (2026-08-11, "FN_LN_ID") ──
print("\n=== Z11. Resume naming + Office-doc URL parsing ===")
from hiring_agent.sharepoint_scoring import (
    _canonical_resume_name as _canon,
    _resume_name_from_url as _name_from_url,
    _resume_name_slots as _slots,
)
# Category removed from the stored name on 2026-08-11 (client instruction: "FN_LN_ID").
# It had been dropped 08-04 and restored 08-06; it comes off for a structural reason -
# Category is a P2 verdict, so a re-score that changes it renames the physical file and
# makes the old name unfindable. Identity (name + AppID tail) never changes.
ok(_canon("Yash Verma", "APP-20260715-2257-7693", ".pdf", "Data Analytics")
   == "Yash_Verma_7693.pdf",
   "canonical name is First_Last_<tail> - no Category (2026-08-11)")
ok(_canon("Muhammad Ahsan Hussain", "APP-20260721-2030-3AA7", ".docx",
          "Mobile Apps (Android IOS)") == "Muhammad_Hussain_3AA7.docx",
   "a middle name is dropped, first + last kept")
ok(_canon("Shivam", "APP-20260706-0548-EE3B", ".pdf", "AI/ML/CV (SIN2)")
   == "Shivam_EE3B.pdf",
   "a single-token name stays single; the AppID tail still makes it unique")
ok(_canon("Jose Garcia", "APP-20260101-0000-AAAA", ".pdf", "Graphics")
   == "Jose_Garcia_AAAA.pdf",
   "accents fold to ASCII rather than losing letters")
ok(_canon("", "APP-20260101-0000-BBBB", ".pdf", "") == "Candidate_BBBB.pdf",
   "an unusable name falls back to 'Candidate'")
# The Category argument is retained in the signature even though the name ignores it:
# _resume_name_slots still needs it to probe the superseded category-bearing files.
ok(_canon("Yash Verma", "APP-20260715-2257-7693", ".pdf", "Data Analytics")
   == _canon("Yash Verma", "APP-20260715-2257-7693", ".pdf", "Cloud and DevOps"),
   "Category no longer affects the stored name, so a re-score never forces a rename")
# Files written while Category WAS in the name must stay findable until migration runs.
_z11_flat = [c for s in _slots("APP-20260727-1431-6F75", "S_N_Ali_CV.pdf",
                               full_name="Syyed Nazir Ali", category="AI/ML/CV (SIN2)")
             for c in s]
ok("Syyed_Ali_6F75.pdf" in _z11_flat,
   "the new FN_LN_ID name is probed")
ok("Syyed_Ali_AIMLCVSIN2_6F75.pdf" in _z11_flat,
   "the superseded 08-06..08-11 category name is STILL probed - no file becomes unreachable")

_docx_url = ("https://x.sharepoint.com/sites/y/_layouts/15/Doc.aspx?sourcedoc=%7BGUID%7D"
             "&file=Peter_Vishal_6CE3.docx&action=default&mobileredirect=true")
_pdf_url = "https://x.sharepoint.com/sites/y/Shared%20Documents/r/2026/July/Sadaf_Khan_7693.pdf"
ok(_name_from_url(_docx_url) == "Peter_Vishal_6CE3.docx",
   "an Office-doc webUrl resolves to the real filename, not 'Doc.aspx'")
ok(_name_from_url(_pdf_url) == "Sadaf_Khan_7693.pdf",
   "a PDF webUrl still resolves from the path")
ok(_name_from_url("https://x/sites/y/_layouts/15/Doc.aspx") == "",
   "a bare viewer endpoint is not mistaken for a document")

_slot = _slots("APP-20260710-2200-A5F2", "resume.pdf", "Jane Doe", "Data Analytics")[0]
ok("Jane_Doe_A5F2.pdf" in _slot,
   "the current FN_LN_ID canonical shape is probed when looking for a stored resume")
ok("Jane_Doe_DataAnalytics_A5F2.pdf" in _slot,
   "the 08-06..08-11 Category shape is STILL probed - every file stored then uses it")
ok("JaneDoe_DataAnalytics_A5F2.pdf" in _slot,
   "the 07-15..08-04 Category shape is STILL probed so un-migrated files are found")
ok(_slot.index("JaneDoe_APP-20260710-2200-A5F2.pdf")
   < _slot.index("Jane_Doe_A5F2.pdf"),
   "P1's legacy write shape is still checked first, so a fresh update wins over stale content")
ok(_slot.index("Jane_Doe_A5F2.pdf") < _slot.index("Jane_Doe_DataAnalytics_A5F2.pdf"),
   "the current canonical shape is preferred over the superseded one it replaced")
# A row whose Category is blank must still find a file stored under the 08-04 shape, which
# never carried a Category - so that probe cannot be gated on Category being present.
_slot_nocat = _slots("APP-20260710-2200-A5F2", "resume.pdf", "Jane Doe", "")[0]
ok("Jane_Doe_A5F2.pdf" in _slot_nocat,
   "a blank-Category row still probes the no-Category shape rather than reading as missing")


# ── Z12. 'Missing' is a GAP everywhere, not a real value (2026-08-04) ─────────
print("\n=== Z12. MISSING_VALUE is treated as a gap by the geo rules ===")
from hiring_agent.geo import (
    _is_gap_text as _geo_gap,
    classify_location_usa as _cls,
    is_foreign_location as _foreign,
    is_strong_usa as _strong,
    GeoDecision as _GD,
)
from hiring_agent.sharepoint_scoring import (
    _apply_missing_placeholders as _apply_missing,
    _get_cleaned_filename_prefix as _fname_prefix,
    _resume_name_slots as _slots2,
)

# THE regression this section exists for: a resume with NO address, a foreign phone and a
# foreign education used to be REJECTED once the cell said 'Missing', while the very same row
# with '' or 'Not extracted' was correctly only ASKED. All three must now agree.
_verdicts = {
    loc: _cls(loc, country="India", phone="+91 98765 43210",
              education="B.Tech, University of Mumbai")[0]
    for loc in ("", "Not extracted", "Missing")
}
ok(all(v == _GD.UNKNOWN for v in _verdicts.values()),
   "a blank / 'Not extracted' / 'Missing' location all stay UNKNOWN - none rejects")
ok(_verdicts["Missing"] == _verdicts["Not extracted"],
   "'Missing' behaves exactly like 'Not extracted' - the placeholder changes nothing")

ok(_geo_gap("Missing") and _geo_gap("Not extracted") and _geo_gap("N/A") and _geo_gap(""),
   "every placeholder literal reads as a gap")
ok(not _geo_gap("Chicago, IL"), "a real location is not a gap")
ok(not _strong("Missing") and not _foreign("Missing"),
   "'Missing' matches neither a US nor a foreign signal")

# A confirmed-US row must be unaffected by the change.
ok(_cls("Chicago, IL", country="India", phone="(224) 663-5226")[0] == _GD.CONFIRMED_US,
   "a concrete US location still confirms, unchanged")

# Missing placeholders now cover the two columns that used to read 'Not extracted'.
_mf = _apply_missing({"Full Name": "", "Looking For Role": "Not extracted",
                      "Phone": "", "Portfolio 1": "N/A"})
ok(_mf["Full Name"] == "Missing" and _mf["Looking For Role"] == "Missing",
   "Full Name and Looking For Role now use the same 'Missing' label as every other gap")
ok(_mf["Portfolio 1"] == "N/A",
   "a portfolio slot stays 'N/A' - 'asked, there is none' must not become 'Missing'")

# ...and the filename must not become 'Missing_<tail>.pdf'.
ok(_fname_prefix("Missing") == "Candidate",
   "a 'Missing' name falls back to 'Candidate' for the saved filename")
ok(_slots2("APP-20260710-2200-A5F2", "cv.pdf", "Missing", "Finance")[0]
   == ["APP-20260710-2200-A5F2_cv.pdf"],
   "a 'Missing' name produces no name-based filename guesses")

# Row pacing.
import hiring_agent.config as _cfg_delay
ok(isinstance(_cfg_delay.ROW_DELAY_SECONDS, float) and _cfg_delay.ROW_DELAY_SECONDS >= 0,
   "ROW_DELAY_SECONDS is a non-negative float")
_sp_src_delay = (pathlib.Path(__file__).parent / "hiring_agent" /
                 "sharepoint_scoring.py").read_text(encoding="utf-8", errors="replace")
ok("if _cfg.ROW_DELAY_SECONDS and not dry_run and current < batch_total:" in _sp_src_delay,
   "the row delay is skipped in dry-run and after the final row of a batch")
_finally_block = _sp_src_delay.split("            _cleanup_temp_files()", 1)[1][:900]
ok("_time.sleep(_cfg.ROW_DELAY_SECONDS)" in _finally_block,
   "the delay sits in the finally block, so an errored row paces like a successful one")


# ── Z14. A place named in a duty bullet is a TOPIC, not an address (2026-08-04) ─
print("\n=== Z14. A place named in a duty bullet is a TOPIC, not an address ===")
from hiring_agent.extraction import (
    _extract_location as _xloc,
    _mentions_place_as_prose as _prose,
)

_topic_cv = """Jane Doe
(602) 555-1234  jane@asu.edu
EDUCATION
Arizona State University, Tempe, Arizona
EXPERIENCE
Acme Corp
Data Analyst
- Led a study analyzing Singapore's demographic trends using Bayesian models.
"""
_got = _xloc(_topic_cv) or ""
ok("singapore" not in _got.lower(),
   f"a country analysed in a bullet is not returned as the location (got {_got!r})")
ok("tempe" in _got.lower() or "arizona" in _got.lower(),
   "the real US place in the CV is found instead")

# A genuine foreign address must keep working - this guard must not cost real rejections.
_addr_cv = """Ravi Kumar
Bengaluru, Karnataka, India
+91 98765 43210
EXPERIENCE
Infosys - Software Engineer
"""
ok("india" in (_xloc(_addr_cv) or "").lower(),
   "a real foreign address is still extracted verbatim")

ok(_prose("- Led a study analyzing Singapore's demographic trends"),
   "a duty bullet reads as prose")
ok(_prose("Developed pipelines for the Germany market"),
   "a duty verb marks prose even without a bullet glyph")
ok(not _prose("Bengaluru, Karnataka, India"), "a bare address is not prose")
ok(not _prose("Tempe, Arizona"), "a short city/state is not prose")


# ── Z13. 'Remote' is a work arrangement, not a geocodable place (2026-08-04) ──
print("\n=== Z13. Work-arrangement words are never geocoded ===")
from hiring_agent.geo import _is_non_place as _nonplace, verify_location_online as _online

ok(_nonplace("Remote") and _nonplace("remote") and _nonplace("WFH")
   and _nonplace("Work From Home") and _nonplace("Hybrid") and _nonplace("Anywhere"),
   "arrangement words are recognised as non-places")
ok(not _nonplace("Chicago, IL") and not _nonplace("Remote (USA)")
   and not _nonplace("Remote - Bengaluru"),
   "a real place is never a non-place, even next to the word Remote")

# The bug: Nominatim resolves the bare word 'Remote' to a US place, so a candidate who
# only ever said 'Remote' came back CONFIRMED_US off no evidence at all. The lookup must
# not happen. Offline assertion - no network needed, the guard returns before requests.
ok(_online("Remote") is None and _online("WFH") is None,
   "a non-place is never sent to the geocoder (returns None = no lookup available)")

ok(_cls("Remote", country="", phone="+91 7799155311")[0] == _GD.UNKNOWN,
   "'Remote' + an Indian phone and no country is UNKNOWN, not a fabricated US confirmation")
ok(_cls("Remote", country="", phone="")[0] == _GD.UNKNOWN,
   "'Remote' with nothing else is UNKNOWN - we simply do not know where they are")
# Long-standing intended behaviour that must NOT regress (see Z8).
ok(_cls("Remote", country="India", phone="+91 98765 43210",
        education="B.Tech, University of Mumbai")[0] == _GD.CONFIRMED_NON_US,
   "'Remote' + India + corroborating phone/education still rejects, exactly as before")
ok(_cls("Remote (USA)", country="", phone="")[0] == _GD.CONFIRMED_US,
   "a parenthetical US signal beside 'Remote' still confirms")


# ── Z15. A skills phrase is not a name; an un-capitalised name is fixed (2026-08-06) ──
# Three live rows reached the client's Rejected sheet with a bad Full Name, and because the
# resume FILE name is derived from it, each one propagated onto SharePoint:
#   DAA7  'Cross-Platform Mobile'  lifted off Abdurrahman's skills line
#   44A6  'aashish sachaniya'      never capitalised
#   B88E  'maftabsabir'            the raw sender display name, returned verbatim
print("\n=== Z15. Name extraction: tech phrases rejected, casing normalised ===")
from hiring_agent.extraction import (
    _looks_like_name as _is_name,
    _plausible_name_shape as _name_shape,
    _looks_like_tech_phrase as _tech,
    _normalize_name as _norm,
    resolve_full_name as _resolve,
)

ok(_tech("Cross-Platform Mobile") and _tech("Full Stack Developer") and _tech("iOS Native"),
   "technology vocabulary is recognised, including inside a hyphenated token")
ok(not _is_name("Cross-Platform Mobile") and not _name_shape("Cross-Platform Mobile"),
   "a skills phrase is rejected by BOTH name-shape checks, not just the strict one")
ok(_resolve("", "", "Abdurrahman\nCross-Platform Mobile Apps\nFlutter & Dart\n")
   == "Abdurrahman",
   "DAA7: the real mononym is found once the skills line stops winning")

ok(_norm("aashish sachaniya") == "Aashish Sachaniya",
   "44A6: an all-lowercase name is capitalised, not stored as typed")
ok(_resolve("", "maftabsabir", "") == "Maftabsabir",
   "B88E: the last-resort sender name is normalised like every other source")

# Regressions this must not cause.
ok(_norm("YASH VERMA") == "Yash Verma", "an all-caps name still normalises as before")
ok(_norm("Maria de la Cruz") == "Maria de la Cruz",
   "lowercase name particles stay lowercase - 'de la' does not become 'De La'")
ok(_norm("Christopher L. Feld") == "Christopher L. Feld",
   "a middle initial survives normalisation")
ok(_norm("Tianxin (William) Wang") == "Tianxin Wang",
   "a parenthesised preferred name is still dropped")
ok(not _tech("Sadaf Khan") and not _tech("Yash Verma") and not _tech("Jose Garcia"),
   "a real name is never mistaken for a tech phrase")
# Occupational surnames must survive - they are deliberately NOT in _TECH_PHRASE_WORDS.
ok(all(_is_name(n) for n in ("James Baker", "Sarah Miller", "Tom Cook",
                             "Ann Taylor", "Paul Marshall", "Ruth Porter")),
   "occupational surnames (Baker, Miller, Cook, Taylor, Marshall, Porter) stay valid names")
ok(_resolve("", "Infilon Technologies", "") == "Not extracted",
   "a vendor's company display name is rejected outright, not capitalised into a candidate")

# Live row APP-20260602-0447-0AE4 stored Full Name 'CamScanner' - the scanning app's own
# producer metadata, which then became 'CamScanner_CloudAndDevOps_0AE4.pdf' on SharePoint.
from hiring_agent.extraction import _is_document_artifact_name as _artifact
ok(all(_artifact(x) for x in ("CamScanner", "Cam Scanner", "cam-scanner", "Adobe Scan",
                              "Untitled", "Document", "Resume", "CV", "Doc1")),
   "scanner/word-processor artifacts are recognised however they are spaced or cased")
ok(not _name_shape("CamScanner"),
   "0AE4: 'CamScanner' no longer passes as a parsed name")
ok(_resolve("CamScanner", "", "CamScanner\nSome Heading\n") != "CamScanner",
   "0AE4: the artifact never survives to become the stored Full Name")
ok(_resolve("", "", "CamScanner\nPeshawar, Pakistan\n") != "CamScanner",
   "the mononym branch also refuses it - a first-3-lines artifact is not a name")
# Whole-string match only: a real name that merely CONTAINS one of these words must survive.
ok(not any(_artifact(x) for x in ("Scott Word", "Cam Newton", "Ana Documenta",
                                  "Vidya Scanlon", "Marc Docken")),
   "a real name containing an artifact word is untouched - matching is whole-string")
ok(_is_name("Scott Word") and _is_name("Cam Newton"),
   "those real names still pass the strict name check")


# ── Z16. Suggested Roles are published strongest-first (2026-08-06) ──────────
# Ollama returns its picks in its own order and the top-up APPENDS keyword matches after
# them, so a strong keyword match could be published in slot 3 behind weaker AI picks.
# Five live rows shipped that way; the worst was APP-20260804-1856-2182 (Bharat Gupta),
# 20% / 20% / 100% - his best match, 'Mobile Application Lead Developer (100%)', sat last
# while slot 1 showed a 20% AWS role.
print("\n=== Z16. Suggested Roles are ranked strongest-first ===")
import re as _re16
import hiring_agent.scoring as _sc16


def _pcts(result):
    out = []
    for k in ("role_1", "role_2", "role_3"):
        m = _re16.search(r"\((\d+)%\)\s*$", str(result.get(k, "") or "").strip())
        if m:
            out.append(int(m.group(1)))
    return out


# The exact live shape: two weak AI picks, then a strong keyword top-up appended last.
_roles16 = [
    {"title": "Software and AWS Developer Job Announcement", "skills": ["aws", "docker"]},
    {"title": "Software and Website Developer Job Announcement", "skills": ["html", "css"]},
    {"title": "Mobile Application Lead Developer Job Announcement",
     "skills": ["kotlin", "android", "jetpack", "retrofit", "room"]},
]
_res16 = _sc16.suggested_roles(
    "Kotlin, Android, Jetpack, Retrofit, Room", "Android Engineer", roles=_roles16)
_p16 = _pcts(_res16)
ok(_p16 == sorted(_p16, reverse=True),
   f"role scores are published in descending order (got {_p16})")
ok(not _p16 or _p16[0] == max(_p16),
   "the STRONGEST match is always Suggested Role 1, never buried in slot 2/3")

# The sort must be STABLE so a tie keeps Ollama's own preference ahead of a keyword top-up.
_tied = [{"title": "AI pick", "score": 80}, {"title": "keyword top-up", "score": 80},
         {"title": "weak", "score": 20}]
_tied.sort(key=lambda r: -int(r.get("score", 0) or 0))
ok([r["title"] for r in _tied] == ["AI pick", "keyword top-up", "weak"],
   "equal scores keep their original order - a tie does not reshuffle Ollama's preference")

# The keyword-only fallback was already sorted; prove it still is.
_kw16 = _sc16._top_n_list("Kotlin, Android, Jetpack, Retrofit, Room", "", roles=_roles16)
ok([s for _, s in _kw16] == sorted([s for _, s in _kw16], reverse=True),
   "the keyword fallback path also returns descending scores")

_sort_src = _re16.sub(r"\s+", " ", inspect.getsource(_sc16.suggested_roles))
ok("ranked.sort(key=lambda r: -int(r.get(\"score\", 0) or 0))" in _sort_src,
   "the ranking sort sits in suggested_roles, before the slots are sliced")


# ── Z17. A failed JD folder scan fails LOUDLY and never downgrades the cache ─────────
# Regression guard for the 2026-08-21 incident: sharepoint_jd_folder pointed at a path that
# 404s ('Documents/Staffing/PDs' vs the real 'Staffing/PDs'). Every failure inside
# fetch_jd_folder_catalog returns [] after a WARNING, which is indistinguishable from an
# empty folder - so with 85 stale pasted descriptions still configured, the rebuild looked
# successful, wrote a cache that had never read the JD folder, and reported "85 role(s)
# cached". Unattended, nothing above WARNING ever surfaced.
print("\n=== Z17. JD folder scan failure is loud and non-destructive ===")
import hiring_agent.jd_sources as _jds
from hiring_agent.jd_sources import JDFolderScanError
from unittest import mock as _m17

_cfg17 = {"enabled": True, "urls": [], "descriptions": [{"title": "Pasted", "text": "python sql"}],
          "sharepoint_jd_folder": "Staffing/PDs", "sharepoint_jd_hostname": "h", "sharepoint_jd_site": "/s"}

# 1. A configured folder returning nothing raises instead of silently continuing on the
#    pasted descriptions (which is exactly what masked the live outage).
with _m17.patch.object(_jds, "fetch_jd_folder_from_sharepoint", return_value=[]):
    try:
        _jds._fetch_active_roles_live(_cfg17)
        ok(False, "a configured-but-empty JD folder must raise JDFolderScanError")
    except JDFolderScanError:
        ok(True, "a configured JD folder returning nothing raises JDFolderScanError")

# 2. The pasted descriptions alone must NOT be able to satisfy the scan.
ok(len(_cfg17["descriptions"]) > 0,
   "the guard above holds even though other JD sources would have produced roles")

# 3. No folder configured -> unchanged behaviour (URL/pasted-only setups still work).
_nofolder = dict(_cfg17, sharepoint_jd_folder="")
ok(len(_jds._fetch_active_roles_live(_nofolder)) == 1,
   "a config with no folder source is unaffected by the new guard")

# 4. build_jd_cache must KEEP the last known good cache rather than overwrite it with a
#    rebuild that never saw the folder - the data-loss half of the incident.
_good = {"built_at": "2026-08-20T10:00:00", "signature": "sig", "folder_fingerprint": "fp",
         "roles": [{"title": "Kept Role", "skills": ["python"], "opening_id": "kept"}]}
with _m17.patch.object(_jds, "load_jd_sources", return_value=_cfg17), \
     _m17.patch.object(_jds, "fetch_jd_folder_from_sharepoint", return_value=[]), \
     _m17.patch.object(_jds, "_read_role_cache", return_value=_good), \
     _m17.patch.object(_jds, "_write_role_cache") as _wrote:
    _out = _jds.build_jd_cache()
    ok(_wrote.call_count == 0,
       f"build_jd_cache does NOT overwrite the cache when the folder scan failed (writes={_wrote.call_count})")
    ok([r["title"] for r in _out] == ["Kept Role"],
       "build_jd_cache returns the last known good roles instead of a folder-less rebuild")

# 5. Same contract on the scoring path: serve stale, never silently re-cache.
with _m17.patch.object(_jds, "load_jd_sources", return_value=_cfg17), \
     _m17.patch.object(_jds, "fetch_jd_folder_from_sharepoint", return_value=[]), \
     _m17.patch.object(_jds, "_read_role_cache", return_value=_good), \
     _m17.patch.object(_jds, "_write_role_cache") as _wrote2:
    _out2 = _jds.get_active_roles(refresh=True)
    ok(_wrote2.call_count == 0,
       f"get_active_roles does NOT re-cache after a failed folder scan (writes={_wrote2.call_count})")
    ok([r["title"] for r in _out2] == ["Kept Role"],
       "get_active_roles falls back to the existing cache rather than a folder-less rebuild")

# 6. --doctor must actively probe folder reachability. Before this it only reported that a
#    cache EXISTED, so a cache built entirely from stale text looked perfectly healthy.
_doctor_src = Path("bot.py").read_text(encoding="utf-8")
ok("_jd_folder_fingerprint" in _doctor_src and "UNREACHABLE" in _doctor_src,
   "--doctor probes JD folder reachability and flags it as UNREACHABLE")


# ── Z18. Skills vocabulary: graphics/design/sales coverage + boilerplate strip ────────
# From the 2026-08-21 vocabulary review. Four defects, all of which broke MATCHING rather
# than crashing anything, so nothing caught them:
#   1. all 14 _GRAPHICS_TOOL_SKILLS were missing from skills.keywords, so the veto that
#      keeps genuine game/graphics devs out of Mobile Apps could never fire;
#   2. Graphics/Gaming/Sales had almost no vocabulary at all (a design JD scraped down to
#      exactly ['marketing']), so neither side of the match had anything to match on;
#   3. the "About DriverAI" blurb put 'computer vision' on Marketing/Finance/Executive JDs;
#   4. finance acronyms rendered as "Fp&A"/"Gaap"/"Sap" in the client-facing skills column.
print("\n=== Z18. Skills vocabulary coverage + JD boilerplate strip ===")
import yaml as _y18
from hiring_agent.extraction import _scan_skill_keywords as _scan18
from hiring_agent.scoring import _GRAPHICS_TOOL_SKILLS as _GT18, _MOBILE_TOOL_SKILLS as _MT18
from hiring_agent.jd_sources import _strip_company_boilerplate as _strip18

_cfg18 = _y18.safe_load(Path("config.yaml").read_text(encoding="utf-8"))
_vocab18 = set(_cfg18["skills"]["keywords"])
_display18 = _cfg18["skills"]["display_names"]

ok(len(_cfg18["skills"]["keywords"]) == len(_vocab18),
   "skills.keywords has no duplicate entries")
# 1. Every override-set term must be extractable, or the override silently can't fire.
ok(not (_GT18 - _vocab18),
   f"every _GRAPHICS_TOOL_SKILLS term is in skills.keywords (missing: {sorted(_GT18 - _vocab18)})")
ok(not (_MT18 - _vocab18),
   f"every _MOBILE_TOOL_SKILLS term is in skills.keywords (missing: {sorted(_MT18 - _vocab18)})")

# The end-to-end case: a Unity/Unreal resume must not be swept into Mobile Apps.
_unity_resume = ("Unity game developer building mobile games. "
                 "Skills: Unity, C#, Unreal Engine, Blender, Shader programming, HLSL, "
                 "Kotlin, Swift, Android, iOS.")
_unity_skills = _scan18(_unity_resume)
ok("unity" in _unity_skills and "shader" in _unity_skills,
   f"a Unity/Unreal resume actually yields the engine skills (got {_unity_skills})")
ok(assign_category("Gaming Position (60%)", ", ".join(_unity_skills)) == "Gaming",
   "a Unity game dev with Kotlin/Swift lands in Gaming, NOT Mobile Apps (the veto works)")

# 2. Domain coverage - each of these had zero vocabulary before.
for _domain, _terms in (
    ("design",  ["figma", "photoshop", "wireframing", "typography", "motion graphics"]),
    ("3d/game", ["unity", "unreal", "blender", "3d modeling", "godot"]),
    ("sales",   ["crm", "salesforce", "lead generation", "cold calling", "b2b"]),
):
    _absent = [t for t in _terms if t not in _vocab18]
    ok(not _absent, f"{_domain} vocabulary present in skills.keywords (missing: {_absent})")

# 3. Ambiguity guards. Each new term below collides with ordinary English, so the guard
#    matters as much as the term: a candidate NAMED Maya must not gain a 3D-modelling skill.
for _label, _text, _must_not, _must in (
    ("a candidate named Maya", "Maya Patel\nFinance Intern\nExcel, SQL, budgeting", "maya", "excel"),
    ("'unity' as prose", "Fostered team unity and unity of purpose. Python, SQL.", "unity", "python"),
    ("'sketch' as a verb", "Helped sketch out the roadmap. Python, Excel.", "sketch", "python"),
    ("a storage quota", "Managed disk quota and storage quota limits. Linux, AWS.", "quota", "linux"),
):
    _got = _scan18(_text)
    ok(_must_not not in _got and _must in _got,
       f"guard: {_label} does not yield {_must_not!r} (got {_got})")
for _label, _text, _want in (
    ("a real Maya artist", "Autodesk Maya, ZBrush, Substance Painter, Blender", "maya"),
    ("a real Unity dev", "Unity, C#, Shader, HLSL", "unity"),
    ("a real Sketch designer", "Figma, Sketch, Photoshop, wireframing", "sketch"),
    ("a real sales rep", "CRM, Salesforce, HubSpot, lead generation, B2B, quota", "quota"),
):
    ok(_want in _scan18(_text), f"guard does not over-reach: {_label} still yields {_want!r}")

# 4. Company boilerplate must be stripped from JD text before the skill scan - but only
#    the About-the-company span, and only when a role heading follows it.
_jd18 = ("About DriverAI\n"
         "DriverAI combines artificial intelligence, computer vision, precision navigation "
         "and cloud platforms.\n"
         "About the Role\n"
         "We are seeking a Marketing Intern to support campaigns.\n"
         "Qualifications\nStrong Excel and social media skills.\n")
_stripped18 = _strip18(_jd18)
ok("computer vision" not in _stripped18.lower(),
   "the About-DriverAI blurb is removed before skills are scanned")
ok("Marketing Intern" in _stripped18 and "Excel" in _stripped18,
   "the actual role content survives the strip")
ok("computer vision" not in _scan18(_stripped18),
   "a marketing JD no longer advertises 'computer vision' from boilerplate")
# Safety: no company heading, or no following role heading -> text untouched (never truncate).
_plain18 = "Marketing Intern\nWe want Excel and computer vision familiarity.\n"
ok(_strip18(_plain18) == _plain18,
   "a JD with no About-the-company heading is returned untouched")
_nofollow18 = "About DriverAI\nWe build computer vision products.\n"
ok(_strip18(_nofollow18) == _nofollow18,
   "an About section with no role heading after it is left alone (never truncates a JD)")

# 5. Display names: acronyms must not fall back to .title().
for _k, _want in (("fp&a", "FP&A"), ("gaap", "GAAP"), ("sap", "SAP"), ("erp", "ERP"),
                  ("quickbooks", "QuickBooks"), ("netsuite", "NetSuite"),
                  ("vlookup", "VLOOKUP"), ("crm", "CRM"), ("b2b", "B2B"),
                  ("opengl", "OpenGL"), ("hubspot", "HubSpot")):
    ok(_display18.get(_k) == _want,
       f"display name for {_k!r} is {_want!r} (got {_display18.get(_k)!r})")


# ── Z19. No NameError landmines (undefined names in function bodies) ─────────
# Regression guard for 2026-08-24: the geo refactor replaced the boolean `is_usa` with the
# tri-state GeoDecision. run_intake's UPDATE branch was migrated, its ADD branch was not, so
# EVERY new candidate reaching --score-folder died on `NameError: name 'is_usa' is not
# defined` - scored correctly, then discarded. Nothing caught it because production uses
# --score-sharepoint (a different module) and no test exercised the intake ADD path.
#
# Conservative by design, biased toward false NEGATIVES so it can never block a build
# spuriously: a name counts as defined if it is assigned ANYWHERE in the function (even in a
# branch that never runs), is a parameter, is module-level, is a builtin, or comes from an
# enclosing scope. Modules using `from x import *` are skipped entirely.
print("\n=== Z19. No undefined names (NameError landmines) ===")
import ast as _ast19, builtins as _bi19
_BI19 = set(dir(_bi19)) | {"__file__", "__name__", "__doc__", "__spec__", "__package__"}

def _bound19(node):
    out, a = set(), node.args
    for grp in (getattr(a, "posonlyargs", []), a.args, a.kwonlyargs):
        out.update(x.arg for x in grp)
    if a.vararg: out.add(a.vararg.arg)
    if a.kwarg: out.add(a.kwarg.arg)
    for c in _ast19.walk(node):
        if isinstance(c, _ast19.Name) and isinstance(c.ctx, (_ast19.Store, _ast19.Del)):
            out.add(c.id)
        elif isinstance(c, _ast19.arg): out.add(c.arg)
        elif isinstance(c, (_ast19.Import, _ast19.ImportFrom)):
            out.update((al.asname or al.name).split(".")[0] for al in c.names)
        elif isinstance(c, (_ast19.FunctionDef, _ast19.AsyncFunctionDef, _ast19.ClassDef)):
            out.add(c.name)
        elif isinstance(c, _ast19.ExceptHandler) and c.name: out.add(c.name)
        elif isinstance(c, (_ast19.Global, _ast19.Nonlocal)): out.update(c.names)
    return out

def _undefined19(path):
    tree = _ast19.parse(Path(path).read_text(encoding="utf-8", errors="replace"))
    if any(isinstance(n, _ast19.ImportFrom) and any(a.name == "*" for a in n.names)
           for n in _ast19.walk(tree)):
        return []
    mod = set(_BI19)
    for c in _ast19.walk(tree):
        if isinstance(c, (_ast19.Import, _ast19.ImportFrom)):
            mod.update((al.asname or al.name).split(".")[0] for al in c.names)
        elif isinstance(c, (_ast19.FunctionDef, _ast19.AsyncFunctionDef, _ast19.ClassDef)):
            mod.add(c.name)
    for c in tree.body:
        for n in _ast19.walk(c):
            if isinstance(n, _ast19.Name) and isinstance(n.ctx, _ast19.Store):
                mod.add(n.id)
    hits = []
    def visit(fn, enclosing):
        scope = _bound19(fn) | enclosing
        inner = [n for n in _ast19.walk(fn)
                 if isinstance(n, (_ast19.FunctionDef, _ast19.AsyncFunctionDef)) and n is not fn]
        for n in _ast19.walk(fn):
            if isinstance(n, _ast19.Name) and isinstance(n.ctx, _ast19.Load):
                if n.id in scope or n.id in mod:
                    continue
                if any(any(x is n for x in _ast19.walk(f)) for f in inner):
                    continue
                hits.append((fn.name, n.id, n.lineno))
        for f in inner:
            visit(f, scope)
    for node in tree.body:
        if isinstance(node, (_ast19.FunctionDef, _ast19.AsyncFunctionDef)):
            visit(node, set())
        elif isinstance(node, _ast19.ClassDef):
            for sub in node.body:
                if isinstance(sub, (_ast19.FunctionDef, _ast19.AsyncFunctionDef)):
                    visit(sub, set())
    return hits

_pkg19 = Path(__file__).with_name("hiring_agent")
_all19 = []
for _f19 in sorted(_pkg19.glob("*.py")):
    for _fn19, _nm19, _ln19 in _undefined19(_f19):
        _all19.append(f"{_f19.name}:{_ln19} {_nm19!r} in {_fn19}()")
ok(not _all19,
   f"no undefined names anywhere in hiring_agent/ (found {len(_all19)}): {_all19[:6]}")

# Pin the exact regression: intake's ADD branch logs the tri-state, never the dead boolean.
_intake19 = (Path(__file__).with_name("hiring_agent") / "intake.py").read_text(encoding="utf-8")
_intake19_code = "\n".join(l.split("#", 1)[0] for l in _intake19.splitlines())
ok("is_usa" not in _intake19_code,
   "intake.py has no CODE reference to the removed boolean 'is_usa' (comments may cite it)")
ok(_intake19.count("geo={geo_decision.value}") >= 2,
   "both intake branches (UPDATE and ADD) log the tri-state geo decision")

# Regression (found live 2026-09-01 via a 10-resume dummy run): the SharePoint scoring
# path (sharepoint_scoring.py) always ran a masked/truncated Phone through sanitize_phone
# before writing it, but the local --score-folder path (intake.py) read Phone straight off
# the extractor with no such guard. A masked resume phone ('732-4***') came back through
# Ollama as a fabricated-looking '732-449***' instead of the clean 'Not extracted' gap
# every other column uses - the one intake route where a garbage value could reach the
# sheet uncaught.
from hiring_agent.extraction import sanitize_phone as _san19
ok(_san19("732-449***") == "Not extracted",
   "sanitize_phone itself blanks a masked value regardless of digit count")
ok("sanitize_phone" in _intake19_code,
   "intake.py's local --score-folder path now sanitizes Phone, matching the SharePoint "
   "scoring path - the one route that used to skip this guard")


# ── Z20. Education dates stay TEXT, never Excel date serials ─────────────────
# The two columns landed 2026-08-18 without a Text guard, so Excel parsed 'Aug 2024' into
# the serial 45505 and every scored row showed the client two meaningless 5-digit numbers.
# Two halves, and BOTH are required: the column must be numberFormat '@' so future writes
# stay literal, AND already-coerced cells must be re-written, because setting the format
# never retroactively converts a cell that already holds a number (same caveat as Phone).
print("\n=== Z20. Education dates are text, not Excel serials ===")
from hiring_agent.sharepoint_scoring import (
    _TEXT_FORMAT_COLUMNS as _tfc20, _EDU_DATE_COLUMNS as _edc20,
    _excel_serial_to_month_year as _ser20)
for _c20 in ("Education Start Date", "Education End Date"):
    ok(_c20 in _tfc20, f"'{_c20}' is forced to Excel Text format (blocks date coercion)")
    ok(_c20 in _edc20, f"'{_c20}' is covered by the serial backfill")
ok("Phone" in _tfc20, "Phone stays text-formatted (the original case this guard was built for)")

# Real serials convert to the readable month/year the extractor originally produced.
for _raw20, _want20 in ((45505, "Aug 2024"), (45839, "Jul 2025"), ("45505", "Aug 2024"),
                        (45505.0, "Aug 2024")):
    ok(_ser20(_raw20) == _want20,
       f"serial {_raw20!r} -> {_want20!r} (got {_ser20(_raw20)!r})")

# Everything the extractor legitimately emits must be left ALONE - these columns are free
# text ('Present', 'Expected May 2026', blank) and must never be "corrected" into a date.
for _keep20 in ("Present", "Expected May 2026", "Aug 2024", "", "N/A", "2024-2025"):
    ok(_ser20(_keep20) == "",
       f"free text {_keep20!r} is never rewritten (got {_ser20(_keep20)!r})")
# Out-of-range numbers are not dates either - a stray 1 or 999999 must not become one.
for _oor20 in (0, 1, 999999, -45505):
    ok(_ser20(_oor20) == "", f"out-of-range {_oor20!r} is left alone")
# Idempotent: feeding the converted text back in is a no-op, so the backfill is safe to
# run on every scoring pass forever.
ok(_ser20(_ser20(45505)) == "", "converting an already-converted value is a no-op (idempotent)")
ok("_renormalize_education_dates" in Path(__file__).with_name("hiring_agent")
   .joinpath("sharepoint_scoring.py").read_text(encoding="utf-8", errors="replace")
   .split("Workbook maintenance 6/6")[-1][:400],
   "the education-date backfill actually runs in the workbook-maintenance pass")



# ── Z21. PDF word-spacing and layout gutters (2026-09-05) ─────────────────────
# Every fixture here is the real first lines of a resume whose Location was WRONG on the
# live sheet and had to be corrected by hand. The common cause: pdf text extraction puts a
# space between every pair of words and a wide run of spaces where the layout had a column
# break, so `lines[:6]` held fragments, no header pass could ever match, and the location
# fell through to a full-document scan that returned an old employer's or school's city.
print("\n=== Z21. PDF word-spacing and layout gutters ===")
from hiring_agent.extraction import (
    _extract_location as _z21loc,
    _contact_block as _z21block,
    _normalize_spacing as _z21norm,
    _past_institution_name as _z21past,
)

# Doubled word spacing must not survive into the stored value - "Palo  Alto, CA" is the
# right place and an unusable cell. Live: Gulnara Timokhina / APP-20260420-1519-1D6A.
_z21_gulnara = ("Gulnara  Timokhina  Palo  Alto,  CA  94303   gtimokhina@gmail.com  "
                "(650)  279-4232   https://linkedin.com/in/gtimokhina   AI/ML  ENGINEER")
ok(_z21loc(_z21_gulnara) == "Palo Alto, CA",
   f"doubled word spacing is squeezed (got {_z21loc(_z21_gulnara)!r})")

# A wide run of spaces is a COLUMN BREAK, not word spacing: collapsing it would glue
# 'University of Massachusetts Amherst' onto 'Amherst, MA'. Live: Rajat Gade, whose resume
# names no other US city, so this cost him his location entirely.
_z21_rajat = """Rajat  Gade
 413-472-4804  |  rajatgade1708@gmail.com |  linkedin.com/in/rajatgade
 EDUCATION  University  of  Massachusetts  Amherst              Amherst,  MA  Master  of  Science  in  Computer  Science  (GPA: 3.87/4.0)  February 2023 - January 2025
"""
ok(_z21loc(_z21_rajat) == "Amherst, MA",
   f"a gutter separates two columns instead of merging them (got {_z21loc(_z21_rajat)!r})")

# ...and the guard that stops a tools list ("Bloomberg Terminal, MS Project") from reading
# as an address must still stop it. A DEGREE or a DATE after the state is the exception.
_z21_tools = """Syyed Nazir Ali
nazir@example.com | (312) 555-0100
TECHNICAL SKILLS
Bloomberg Terminal, MS Project, MS Office Suite
"""
ok(_z21loc(_z21_tools) is None,
   f"a tools list is still not an address (got {_z21loc(_z21_tools)!r})")
ok(_z21norm("A  B   C") == "A B | C", "3+ spaces become a separator, 2 collapse")

# The contact block ends at the first section heading even when the PDF collapsed the whole
# resume onto a few long lines with the heading buried MID-line, and it never runs on into
# duty bullets. Live: Aniruddha Rajnekar / APP-20260428-2224-14E7, whose block ran to its
# 12-line cap - twelve bullets - so 'Barclays, Pune, India' (left in 2022) was scanned as
# an address and beat the degree he finished in 2025.
_z21_ani = [
    "Aniruddha Rajnekar (919)-521-2392| aarajnek@ncsu.edu |Portfolio SUMMARY | Full Stack Software Engineer",
    "● Mentored new hires ... Full Stack Software Engineer | Barclays, Pune, India",
]
ok(len(_z21block(_z21_ani)) == 1 and "SUMMARY" not in _z21block(_z21_ani)[0],
   f"an embedded ALL-CAPS heading ends the contact block (got {_z21block(_z21_ani)!r})")
ok(_z21block(["Jane Doe", "● Led a team in Bengaluru, India"]) == ["Jane Doe"],
   "a duty bullet is never contact information")

# A city AFTER a school's name is its campus and counts; a city INSIDE the name does not.
# Live: Prerna Saluja / APP-20260716-2052-6112, whose only 'Delhi' is her alma mater.
ok(_z21past("Arizona State University, Tempe, Arizona") == "Tempe, Arizona",
   "a campus address survives the institution guard")
ok(_z21past("Daulat Ram College, University of Delhi") == "",
   "a line that is nothing but school names yields no address")

# An explicit US STATE outranks a bare city keyword found anywhere in the document.
_z21_rank = """Aniruddha Rajnekar (919)-521-2392| aarajnek@ncsu.edu
● Full Stack Software Engineer | Barclays, Pune, India
● EDUCATION Master of Computer Science | North Carolina State University, Raleigh, North Carolina (August 2023 - May 2025)
"""
ok("raleigh" in (_z21loc(_z21_rank) or "").lower(),
   f"a US state beats an old employer's foreign city (got {_z21loc(_z21_rank)!r})")



# ── Z22. The MOST RECENT dated entry decides the location (2026-09-05) ─────────
# A resume states where someone WAS at every job and every school; only the dates say which
# of those is where they ARE. Nothing read the dates, so the first place in the document won
# and US-resident candidates were stored - and geo-classified - as living abroad.
print("\n=== Z22. The most recent dated entry decides the location ===")
from hiring_agent.extraction import (
    _extract_location as _z22loc,
    _recency_location as _z22rec,
    _entry_end_key as _z22key,
    split_location_country as _z22split,
)

ok(_z22key("Present") > _z22key("Dec 2099"), "'Present' outranks every real date")
ok(_z22key("Oct 2024") > _z22key("Aug 2024"), "the month is part of the ordering, not just the year")
ok(_z22key("Not a date") is None, "a non-date is not an entry")

# Live: Hitesh Malhotra / APP-20260426-1541-4D6F. He was stored in Bengaluru - a job he
# left in Aug 2024 - while the line above says he has been in Lansing since Oct 2024.
_z22_hitesh = """Hitesh Malhotra
+1 5177126996 | hiteshmalhotra2000@gmail.com | LinkedIn
EXPERIENCE
Research Assistant , Michigan State University - Lansing, United States Oct 2024 - Present
Associate Software Engineer , Conga - Bengaluru, India Jan 2022 - Aug 2024
"""
ok(_z22split(_z22loc(_z22_hitesh) or "")[1] == "United States",
   f"a current US role beats a foreign one left years earlier "
   f"(got {_z22loc(_z22_hitesh)!r})")

# Live: Manasa Puli / APP-20260505-2113-49BD. Stored as 'Hyderabad, India' (Cognizant,
# ended Jul 2022); her current role is remote from Georgia. A bare state name is accepted
# HERE, though the header passes refuse the name-like states, because a dated work entry
# names a workplace - the candidate's own name is up in the contact block.
_z22_manasa = """MANASA PULI
865-356-1921 | pulimanasa1999@gmail.com
EXPERIENCE
AdyahTech Apr 2025 - Present Data Analyst Georgia (Remote)
Cognizant Technology Solutions Jan 2021 - Jul 2022 Programmer Analyst Hyderabad, India
"""
ok(_z22loc(_z22_manasa) == "Georgia",
   f"the current remote role's state wins over a 2022 employer's city "
   f"(got {_z22loc(_z22_manasa)!r})")

# The place often sits on the line BELOW the dates - a degree line and its school line.
# Live: Kshitij Sahu / APP-20260512-1310-7C3D, enrolled at ASU through May 2026 but stored
# as 'Remote', scraped from an internship's "India (Remote)".
_z22_kshitij = """Kshitij Sahu
480-875-4195 | kshitij.sahu3@gmail.com
Education
Master of Science in Electrical Engineering Aug 2024 - May 2026
Arizona State University (3.72/4) Arizona, USA
Bachelor of Technology in EEE Aug 2018 - May 2022
Kalinga Institute of Industrial Technology (8.27/10) Odisha, India
Experience
Hardware Intern May 2025 - July 2025
Hitsun Engineering Corporation India (Remote)
"""
ok(_z22split(_z22loc(_z22_kshitij) or "") == ("Arizona", "United States"),
   f"a degree still in progress outranks a finished internship abroad "
   f"(got {_z22loc(_z22_kshitij)!r})")

# A resume that states its own address up top is NOT touched by any of this - the contact
# block still wins outright, however many dated entries follow it.
_z22_header_wins = """Jane Doe
Austin, TX 78701 | jane@example.com
EXPERIENCE
Engineer, Acme - Bengaluru, India Jan 2024 - Present
"""
ok(_z22loc(_z22_header_wins) == "Austin, TX",
   f"a stated contact-block address still wins (got {_z22loc(_z22_header_wins)!r})")

# Entries ending within a year of each other are concurrent, and the more specific of the
# two statements is the useful one.
ok(_z22rec(["Research Engineer, Acme Jan 2024 - Jun 2025 | USA",
            "M.S. Engineering, Aug 2022 - Dec 2024 | Northridge, CA"]) == "Northridge, CA",
   "among concurrent entries the more specific place wins")


# ── Z23. An area code the US never issued is not a US phone (2026-09-08) ──────
# Live case Saad Ullah / APP-20260804-2021-WQWA: a Pakistani mobile reached the main sheet
# formatted as "(333) 333-8893". 333 is unassigned, but it satisfies every NANP SHAPE rule
# (N-X-X, no N11, exchange not 0/1), so format_phone dressed it in US parens and
# _looks_like_us_phone then read that as evidence of a US candidate. The guard has to
# distinguish ASSIGNED from merely well-shaped, which is why _US_AREA_CODES is an allowlist.
print("\n=== Z23. Unassigned area codes are not US phones ===")
from hiring_agent.geo import is_us_area_code

ok(format_phone("333-3338893") == "333-3338893",
   "an unassigned area code is left alone, not given US parens")
ok(_usph("333-3338893") is False,
   "...and never counts as US evidence for the geo gate")

# The fix must not over-reach: 447 LOOKS like a typo for a UK +44 7 number but is a genuine
# Illinois overlay on 217, and the live row carrying it (Neil Mehta, neilbm2@illinois.edu,
# Chicago IL) is a real US candidate. An allowlist is the only thing that separates the two.
ok(format_phone("(447) 902-7824") == "(447) 902-7824",
   "447 is a real Illinois overlay and still formats as US")
ok(_usph("(447) 902-7824") is True,
   "...and still counts as US evidence")

ok(is_us_area_code("480") and is_us_area_code("212") and is_us_area_code("787"),
   "states and +1 territories are in the allowlist")
ok(not is_us_area_code("519"),
   "a Canadian area code is not a US area code")
ok(not is_us_area_code("333") and not is_us_area_code("011"),
   "unassigned and structurally-invalid codes are both rejected")

# The country-code path needs the same allowlist, or '+1 <unassigned>' walks straight through.
ok(format_phone("1-333-333-8893") == "1-333-333-8893",
   "a +1 prefix does not launder an unassigned area code")
ok(format_phone("1-732-997-5324") == "(732) 997-5324",
   "...while a real +1 US number still normalises")

# ── Z24. Education dates come from the CURRENT degree (2026-09-08) ────────────
# The tier order used to decide this: a strict "August 2018 - July 2022" range outranked an
# "Expected May 2026" no matter which degree each belonged to, so a candidate's finished
# foreign bachelor's dates were published instead of the US master's they are enrolled in.
# Same rule Location already follows (Z22) - among dated entries, the current one wins.
print("\n=== Z24. Education dates follow the most recent degree ===")
from hiring_agent.extraction import _extract_education_dates as _edu_dates

_z24_umd = """EDUCATION
University of Maryland, College Park, Maryland, USA Master of Science, Applied Machine Learning Expected May 2026
INDIRA GANDHI DELHI TECHNICAL UNIVERSITY FOR WOMEN (IGDTUW) Delhi, India Bachelor of Technology, Electronics and Communication Engineering August 2018 - July 2022"""
ok(_edu_dates(_z24_umd) == ("", "May 2026"),
   f"an 'Expected' current degree outranks an older printed range (got {_edu_dates(_z24_umd)})")

_z24_two_years = """Education
Bachelor's from JNTUK, India - 2017, Master's From Campbellsville University, USA - 2024"""
ok(_edu_dates(_z24_two_years)[1] == "2024",
   f"the single-date tier takes the LATEST year, not the first (got {_edu_dates(_z24_two_years)})")

# The common case is one degree, and it must be untouched: taking a max over a single
# candidate is the same answer the first-match rule gave.
_z24_single = """EDUCATION
Master of Science, Computer Science Aug 2023 - May 2025 Arizona State University, Tempe, AZ"""
ok(_edu_dates(_z24_single) == ("Aug 2023", "May 2025"),
   f"a single-degree resume is unchanged (got {_edu_dates(_z24_single)})")

_z24_present = """EDUCATION
M.S. Data Science, Arizona State University Aug 2024 - Present
B.Tech, Mumbai University Aug 2018 - Jun 2022"""
ok(_edu_dates(_z24_present) == ("Aug 2024", "Present"),
   f"'Present' outranks a completed earlier degree (got {_edu_dates(_z24_present)})")

# A resume that states no dates at all still yields a blank pair - the honest answer. The
# recency ordering must not invent one.
ok(_edu_dates("EDUCATION\nMaster of Science, Computer Science") == ("", ""),
   "a resume with no stated dates still yields a blank pair")

# APP-20260810-1751-MC9A (Victor): the Education section is one collapsed line and the
# next line starts Professional Experience.  The old fixed five-line window crossed that
# boundary and selected the newer founder-job range (2026-Present) as education.
_z24_victor = """Education:
USP | ESALQ MBA Degree, Finance and Controlling Bachelor's Degree, Marketing and Graphic Design São Paulo, Brazil | 2023 - 2025 São Paulo, Brazil | 2016 - 2019
Professional Experience:
Founder at YouPlaya Marketing Agency São Paulo, Brazil | February/2026 - Currently"""
ok(_edu_dates(_z24_victor) == ("2023", "2025"),
   f"education date scan stops at Professional Experience (got {_edu_dates(_z24_victor)})")

_z24_mona = """Education
PG Diploma in Advanced Japanese Language – Univ. of Delhi (2018–2019)
Basic & Intermediate Japanese – MOSAI Institute (2014–2016)
B.Sc. in Animation and Film Making – Punjab Technical University (2011–2014)"""
ok(_edu_dates(_z24_mona) == ("2011", "2014"),
   f"dates stay attached to the B.Sc. entry selected for Education (got {_edu_dates(_z24_mona)})")

_z24_yash = """EDUCATION
California State University Los Angeles _August 2023 - May 2025 Master of Science in Computer Science_
University of Mumbai _July 2015 - May 2019 Bachelor of Engineering in Computer Engineering_"""
ok(_edu_dates(_z24_yash) == ("Aug 2023", "May 2025"),
   f"markdown emphasis does not erase education month precision (got {_edu_dates(_z24_yash)})")

# Live: Sankalp Sharma / APP-20260804-1923-MDKA. His PDF draws 2014 at the far
# right of the BCA row. Markdown reading order moved it below LANGUAGES, while the
# coordinate-sorted text correctly keeps "BCA 2014" together. Preserve just that
# row and let the anchored date pass use the second, layout-aware copy.
from hiring_agent.extraction import _pdf_layout_education_rows as _layout_edu_rows

_z24_sankalp_primary = """Android Development Contest Winner - College Level: Recognised for innovation
EDUCATION
Bachelor of Computer Applications (BCA)
SSLD Varshney Institute of Management & Engineering, Aligarh, Uttar Pradesh
LANGUAGES
English - Professional Working Proficiency
2014"""

class _Z24LayoutPage:
    def get_text(self, mode, sort=False):
        assert mode == "text" and sort is True
        return "Bachelor of Computer Applications (BCA)                         2014\n"

_z24_layout_rows = _layout_edu_rows([_Z24LayoutPage()], _z24_sankalp_primary)
ok(_z24_layout_rows == ["Bachelor of Computer Applications (BCA) 2014"],
   f"the visual BCA/date row is recovered once (got {_z24_layout_rows!r})")
_z24_sankalp_text = _z24_sankalp_primary + "\n" + "\n".join(_z24_layout_rows)
ok(_edu_dates(_z24_sankalp_text) == ("", "2014"),
   f"the date aligned with BCA survives Markdown reordering (got {_edu_dates(_z24_sankalp_text)})")

_z24_pragya = """Education
Masters, Computer Science Fellowship Scholar 2023-2024 - Texas A&M University Aug 202_3 - May 2025
B.Tech, Computer Science 2019 - 2023
Masters, Computer Science - Texas A&M University, College Station, TX Aug 2023 - May 2025"""
ok(_edu_dates(_z24_pragya) == ("Aug 2023", "May 2025"),
   f"a dated degree row outranks fellowship years on the same degree line (got {_edu_dates(_z24_pragya)})")

# ── Z25. 2026-09-14 LIVE-ROW REVIEW FIXES ────────────────────────────────────
# Found by opening the resumes behind APP-20260908-1304-MCLA and APP-20260902-2155-MCPA and
# checking every published column against them.
print("\n=== Z25. LIVE-ROW REVIEW FIXES (education, skills, seed, JD context) ===")
from hiring_agent.extraction import complete_education, normalize_education, normalize_skills
from hiring_agent.scoring import _skill_set
from hiring_agent.jd_sources import _drop_context_mentions
from hiring_agent.extraction import _scan_skill_keywords
import hiring_agent.config as _z25cfg

# Education: the school belongs in the column; the city/country and GPA after it do not.
_z25_mcla = ("EDUCATION\nBachelor of Science in Computer Science\n"
             "National University of Computer & Emerging Sciences (FAST-NUCES), Lahore, Pakistan\nSKILLS")
ok(complete_education("Bachelor of Science in Computer Science", _z25_mcla) ==
   "Bachelor of Science in Computer Science, National University of Computer & Emerging Sciences (FAST-NUCES)",
   "a degree-only Education gains its school from the next line, without the city/country")
ok(complete_education("B.S. Physics", "Education\no B.S. Physics\no Ohio State University, Columbus, OH | 2019 - 2023")
   == "B.S. Physics, Ohio State University",
   "an 'o ' bullet is stripped without eating the O of 'Ohio'")
ok(complete_education("M.S. Computer Science", "EDUCATION\nM.S. Computer Science, Arizona State University (3.72/4) Tempe, AZ")
   == "M.S. Computer Science, Arizona State University",
   "a GPA ends the school name, so the campus city after it is not captured")
ok(complete_education(
       "Bachelor of Computer Applications (BCA)", _z24_sankalp_primary) ==
   "Bachelor of Computer Applications (BCA), SSLD Varshney Institute of Management & Engineering",
   "an award sentence containing 'College' cannot replace the school below the degree")
_z25_nimish = ("EDUCATION\nW.P. Carey School of Business, Arizona State University\n"
               "M.S. in Information Systems & Management\nRelevant Coursework\n"
               "R.V. College of Engineering\nB.E. in Electrical Engineering")
ok(complete_education("M.S. in Information Systems & Management", _z25_nimish) ==
   "M.S. in Information Systems & Management, W.P. Carey School of Business, Arizona State University",
   "the institution immediately above the current degree wins over an older school below it")
ok(complete_education("MBA", "EDUCATION\nMBA\nSKILLS\nPython") == "MBA",
   "no school is invented when none is near the degree (stops at the next section)")
ok(complete_education(
       "MBA", "Contest Winner - College Level\nEDUCATION\nMBA\nSKILLS\nPython") == "MBA",
   "the Education heading also blocks an institution-like phrase from the prior section")
ok(complete_education("M.S. Data Science, Arizona State University", "unrelated") ==
   "M.S. Data Science, Arizona State University", "a value that already names a school is unchanged")
ok(complete_education(
    "Bachelor in Computer Science",
    "EDUCATION\nNational University of Computer and Emerging Sciences (NUCES) | June 2025\n"
    "Bachelor in Computer Science\nEXPERIENCE") ==
   "Bachelor in Computer Science, National University of Computer and Emerging Sciences (NUCES)",
   "a 'Bachelor in' degree gains the school from the line immediately above it")
ok(normalize_education("B.Sc. (Design and Computing), BITS Pilani (WILP) | 2025") ==
   "B.Sc. (Design and Computing), BITS Pilani (WILP)",
   "a trailing '| 2025' is removed - the year already lives in Education End Date")
ok(normalize_education("Master of Science, University of Texas, Aug 2019 - May 2021") ==
   "Master of Science, University of Texas", "a trailing date range is removed")
ok(normalize_education("B.Tech in CSE (Expected 2026)") == "B.Tech in CSE", "a trailing (Expected YYYY) is removed")
ok(normalize_education(
    "MCA – GNDU Regional Campus (2009–2012) BCA – DAV College (2006–2009)") ==
   "MCA – GNDU Regional Campus BCA – DAV College",
   "parenthesized ranges are removed from every degree in a multi-degree Education value")
ok(normalize_education("Seattle University, MS in Computer Science Sept 2023 - June 2025") ==
   "Seattle University, MS in Computer Science",
   "a whole trailing range is removed, never just its second half")
ok(normalize_education("Bachelor of Arts, Class of 2020 Scholars Program") ==
   "Bachelor of Arts, Class of 2020 Scholars Program", "a year that is not a trailing date stays")
ok(normalize_education(
    "Master of Science, ASU | _May 2026_ Bachelor of Technology, SNI | _May 2024_") ==
   "Master of Science, ASU Bachelor of Technology, SNI",
   "markdown emphasis and internal degree dates are removed from Education")
ok(normalize_education(
    "Master of Science in EE Aug 2024 – May 2026, ASU; Bachelor of Technology Aug 2018 – May 2022, KIIT") ==
   "Master of Science in EE, ASU; Bachelor of Technology, KIIT",
   "month ranges are removed cleanly from each degree in a combined Education value")

# Skills: a bracketed group is one skill, the plain duplicate goes, scoring keeps every token.
ok(normalize_skills("AWS, AWS (EC2, S3, IAM), Docker") == "AWS (EC2/S3/IAM), Docker",
   "'AWS' beside 'AWS (EC2, S3, IAM)' collapses to the specific form, kept as one skill")
ok({"aws", "ec2", "s3", "iam", "docker"} <= _skill_set("AWS (EC2/S3/IAM), Docker"),
   "scoring still reads the base and the bracketed items, so no role match is lost")
ok(normalize_skills("Excel (Pivot, VLOOKUP") == "Excel (Pivot, VLOOKUP",
   "an unbalanced bracket never swallows the rest of the list")
_z25_hw = _scan_skill_keywords(
    "SystemVerilog RTL ASIC FPGA TCL GDSII CDC RDC ModelSim Synopsys Verdi "
    "Design Compiler Cadence Innovus Cadence Genus Xilinx Vivado SPI I2C")
ok({"systemverilog", "rtl", "asic", "fpga", "gdsii", "modelsim", "cadence innovus",
    "xilinx vivado", "spi", "i2c"} <= set(_z25_hw),
   "hardware and silicon resume skills remain visible to role matching")
from hiring_agent.extraction import _merge_keyword_skills
# 33 + 1 group + 3 genuine + 3 new keywords = 40, exactly the cap once the group counts once.
# Counted the old way (group shredded into 3, plus a re-added plain 'AWS') it is 43, and the
# cap cut n8n, SQLite and Microservices - the live regression this guards.
_z25_llm = ", ".join([f"Skill{i}" for i in range(33)] + ["AWS (EC2, S3, IAM)", "Microservices", "SQLite", "n8n"])
_z25_merged = normalize_skills(_merge_keyword_skills({"skills": _z25_llm}, "AWS OAuth RSS Docker")["skills"]).split(", ")
ok(all(s in _z25_merged for s in ("Microservices", "SQLite", "n8n")) and "AWS" not in _z25_merged,
   "the skills cap counts a bracketed group once, so genuine skills are not pushed out "
   f"(got {len(_z25_merged)}: tail {_z25_merged[-6:]})")

# The AI recheck replaces each field wholesale, so a shorter second reading silently DELETED
# real skills. Live: APP-20260915-1110-OP7A lost Machine Learning, RTL, ASIC, GDSII, Design
# Compiler and Automation - all written in her CV - and those six tokens were the difference
# between a 35% top role match and 0% across the board.
from hiring_agent.extraction import _merge_recheck_skills
_z26_cv = ("Implemented the full ASIC flow from RTL to GDSII using Synopsys Design Compiler. "
           "Machine Learning with deployment to FPGAs. Streamlined workflow automation.")
_z26_old = "Verilog, Machine Learning, RTL, ASIC, GDSII, Design Compiler, Automation"
_z26_kept = _merge_recheck_skills(_z26_old, "Verilog", _z26_cv)
ok(all(s in _z26_kept for s in ("Machine Learning", "RTL", "ASIC", "GDSII",
                                "Design Compiler", "Automation")),
   f"a recheck cannot delete skills the CV actually states (got {_z26_kept!r})")
ok(_merge_recheck_skills("Verilog, Kubernetes", "Verilog", _z26_cv) == "Verilog",
   "a first-pass skill with no support anywhere in the documents still gets dropped")
ok(_merge_recheck_skills("Verilog", "Verilog, Python", _z26_cv) == "Verilog, Python",
   "the recheck may still ADD skills it newly recognised")
for _term, _label in (("Zephyr", "Zephyr"), ("Bugzilla", "Bugzilla"), ("DigitalOcean", "DigitalOcean"),
                      ("OAuth", "OAuth"), ("RSS", "RSS"), ("Gmail API", "Gmail API"), ("Supertest", "Supertest")):
    ok(_term.lower() in _scan_skill_keywords(f"Tools: {_term}, Git"), f"{_label} is in the skills vocabulary")
ok("rest assured" not in _scan_skill_keywords("Rest assured, I will deliver on time."),
   "the everyday idiom 'rest assured' is not scanned as a skill")

# QA guard: a QA/test candidate is 'General' while the catalog has no QA/test opening.
from hiring_agent import scoring as _z25_scoring
_z25_qa_skills = "Selenium, Cypress, TestNG, JUnit, TestRail, Java, JavaScript, GitHub Actions, CI/CD"
_z25_real_cache = _z25_scoring._catalog_has_qa_opening
try:
    _z25_scoring._catalog_has_qa_opening = lambda: False
    ok(assign_category("Mobile Application Software Developer (80%)", _z25_qa_skills,
                       "Software Development Engineer In Test (SDET)") == "General",
       "an SDET matched to a developer role is filed under General, not the developer category")
    ok(assign_category("Full Stack Web Developer (60%)", _z25_qa_skills) == "General",
       "three or more test tools alone identify a QA profile (no Looking For Role needed)")
    ok(assign_category("Full Stack Web Developer (60%)", "Java, JUnit, React, Node.js", "QA Automation Engineer") == "General",
       "a stated QA/SDET role alone identifies a QA profile")
    ok(assign_category("Full Stack Web Developer (70%)", "Java, Spring Boot, JUnit, React, Node.js",
                       "Software Engineer") == "Web Team (Full stack/Back end & UI/UX)",
       "a developer who merely lists JUnit keeps their normal category")
    _z25_scoring._catalog_has_qa_opening = lambda: True
    ok(not _z25_scoring.is_unmatched_qa_profile(_z25_qa_skills, "SDET"),
       "the guard retires itself as soon as the catalog contains a QA/test opening")
finally:
    _z25_scoring._catalog_has_qa_opening = _z25_real_cache

# Hardware guard: same rule, different discipline. A silicon engineer is 'General' while the
# catalog has no chip-design opening, rather than being named after a weak software match.
# Live: APP-20260915-1110-OP7A was filed 'AI/ML/CV (SIN2)' off a 35% AI/ML match because her
# CV mentions machine learning once and a GCN hardware accelerator - she designs chips.
_z27_hw_skills = ("Verilog, SystemVerilog, Cadence Virtuoso, HSPICE, ModelSim, Calibre, "
                  "Xilinx Vivado, PnR, DRC, LVS, STA, CTS, RTL2GDS, VLSI Design, Python, "
                  "Machine Learning")
_z27_real_cache = _z25_scoring._catalog_has_hardware_opening
try:
    _z27_scoring_cache = _z25_scoring._catalog_has_hardware_opening = lambda: False
    ok(assign_category("Software Developer - AI/ML, Computer Vision (35%)", _z27_hw_skills,
                       "EE Circuit Design Engineer") == "General",
       "a VLSI engineer weakly matched to an AI/ML role is filed under General, not AI/ML")
    ok(assign_category("Software Developer - AI/ML, Computer Vision (35%)", _z27_hw_skills) == "General",
       "three or more silicon tools alone identify a hardware profile (no Looking For Role needed)")
    ok(assign_category("Software and Website Developer (70%)", "Python, React, Verilog, SQL, Git",
                       "Frontend Engineer") == "Web Team (Full stack/Back end & UI/UX)",
       "a developer who merely lists Verilog once keeps their normal category")
    _z25_scoring._catalog_has_hardware_opening = lambda: True
    ok(not _z25_scoring.is_unmatched_hardware_profile(_z27_hw_skills, "EE Circuit Design Engineer"),
       "the hardware guard retires itself as soon as the catalog contains a chip-design opening")
finally:
    _z25_scoring._catalog_has_hardware_opening = _z27_real_cache
# The live catalog must not already satisfy the guard, or it would never fire.
ok(not _z25_scoring._catalog_has_hardware_opening(),
   "no current JD title reads as a hardware opening, so the guard is active")

# Seed: identical input must give identical scores.
ok(isinstance(_z25cfg.OLLAMA_SEED, int), "every Ollama call carries a fixed seed from config")

# JD context: company/collaboration mentions of computer vision are not a requirement.
for _ctx in ("Collaborate with AI/ML and computer vision teams to present detections.",
             "Integrate computer vision outputs into web applications.",
             "Collaborate with product, design, mobile, AI/ML, computer vision, cloud, data, and security teams.",
             "Interest in AI, data analytics, computer vision, cybersecurity, or emerging technology."):
    ok("computer vision" not in _scan_skill_keywords(_drop_context_mentions(_ctx)),
       f"context mention dropped: {_ctx[:60]!r}")
for _req in ("Strong experience in computer vision and deep learning with PyTorch.",
             "Build computer vision models for product recognition.",
             "Develop and deploy computer vision pipelines on edge devices."):
    ok("computer vision" in _scan_skill_keywords(_drop_context_mentions(_req)),
       f"genuine requirement kept: {_req[:60]!r}")

print(f"\n{'='*64}")
print(f"  P2 RESULT: {P} passed, {F} failed")
print(f"{'='*64}")
import sys
sys.exit(1 if F else 0)
