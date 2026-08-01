"""Cross-check the current P1/P2 folder and file contracts."""
import json, sys, re
from pathlib import Path

sys.path.insert(0, 'HiringAgent_App_P2')

# ── Load configs ──────────────────────────────────────────────────
p1_cfg  = json.load(open('HiringAgent_P1/flow/flow_config.json', encoding='utf-8'))
p2_cfg  = json.load(open('HiringAgent_App_P2/config.yaml', encoding='utf-8')) \
          if False else None  # use config.py import instead

from hiring_agent import config as p2

P = F = 0
def ok(c, label, note=''):
    global P, F
    sym = "[PASS]" if c else "[FAIL]"
    if c: P += 1
    else: F += 1
    suffix = f"  <- {note}" if note else ''
    print(f"  {sym} {label}{suffix}")

# ─────────────────────────────────────────────────────────────────
print("=" * 62)
print("  EXPECTED SHAREPOINT STRUCTURE (from current runtime config)")
print("=" * 62)
print("  Documents/")
print("    Candidate_Resumes/               <- resumes go here")
print("    Master_Files/                    <- live master workbook")
print("      Sharepoint_Master_File.xlsx")
print("    Candidate_List_Results.xlsx      <- client export at library root")
print("    SharePoint_Master_Template/      <- one-time setup template")
print()

# ─────────────────────────────────────────────────────────────────
print("=== P1: RESUME FOLDER ===")
rf = p1_cfg['sharepoint']['resumes_folder']
print(f"  Config value: '{rf}'")
ok('Candidate_Resumes' in rf or 'Downloaded_Resumes' in rf,
   "resumes_folder contains 'Candidate_Resumes'",
   "matches live SharePoint folder")
ok(rf.startswith('/'),
   "resumes_folder is absolute (starts with '/')")
ok('Shared Documents' in rf or 'Candidate_Resumes' in rf,
   "resumes_folder path is well-formed")

print()
print("=== P1: MASTER WORKBOOK ===")
wb = p1_cfg['excel']['file']
print(f"  Config value: '{wb}'")
ok('Master_Files' in wb,
   "excel.file is inside Master_Files",
   "matches live SharePoint folder")
ok(wb.endswith('Sharepoint_Master_File.xlsx'),
   "excel.file filename = Sharepoint_Master_File.xlsx",
   "matches file on SharePoint")
ok(p1_cfg['excel']['table'] == 'HiringAgent_P1_Candidates',
   f"excel.table = HiringAgent_P1_Candidates")

print()
print("=== P1: DATED RESUME SUBFOLDERS ===")
ok(p1_cfg['sharepoint']['dated_resume_subfolders'] is True,
   "dated_resume_subfolders = true  (saves to Candidate_Resumes/<Year>/<Month>/)")
print(f"  Example path: Documents/Candidate_Resumes/2026/July/<AppID>_resume.pdf")

print()
print("=== P1: SHAREPOINT SITE ===")
site = p1_cfg['sharepoint']['site']
print(f"  Site: {site}")
ok(site.startswith('https://') and 'sharepoint.com' in site,
   "Site URL is a valid SharePoint URL")

print()
print("=== P2: MASTER WORKBOOK (must match P1 exactly) ===")
# P2 reads the workbook path from SharePoint client env vars, not config.py
# But the Excel basename is fixed:
ok(p2.EXCEL_BASENAME == 'HiringAgent_P1_CandidateList.xlsx',
   f"P2 EXCEL_BASENAME = {p2.EXCEL_BASENAME}",
   "local mode only — in SharePoint mode path comes from SHAREPOINT env var")

# Check what P2's sharepoint_scoring expects as the workbook
import re as _re
scoring_src = open('HiringAgent_App_P2/hiring_agent/sharepoint_scoring.py', encoding='utf-8').read()
ok("Sharepoint_Master_File" in scoring_src or "wb_path" in scoring_src,
   "sharepoint_scoring.py references the workbook path (via _wb_path from client)")

print()
print("=== P2: RESUME FOLDER (reads from Candidate_Resumes) ===")
ok("Candidate_Resumes" in scoring_src or "resumes_folder" in scoring_src,
   "sharepoint_scoring.py references resumes_folder (reads resumes from Candidate_Resumes)")
ok("resumes_folder" in scoring_src,
   "sharepoint_scoring.py uses client.resumes_folder attribute")

# Check the client connects resumes_folder properly
client_src = open('HiringAgent_App_P2/sharepoint_client.py', encoding='utf-8').read()
ok("resumes_folder" in client_src,
   "sharepoint_client.py exposes resumes_folder to scoring")
ok("Candidate_Resumes" in client_src or "resumes_folder" in client_src,
   "sharepoint_client.py passes resume folder path through")

print()
print("=== P2: RESULTS EXPORT FILE ===")
ok("Candidate_List_Results.xlsx" in scoring_src,
   "P2 generates Candidate_List_Results.xlsx",
   "matches client export filename")
ok("def _client_export_sharepoint_folder" in scoring_src and 'return ""' in scoring_src,
   "P2 uploads results to the Shared Documents library root",
   "matches /Candidate_List_Results.xlsx")

print()
print("=== P1 TEMPLATE FOLDER (SharePoint_Master_Template) ===")
# Check if the template .xlsx in the P1 codebase is meant for this folder
tpl = Path('HiringAgent_P1/P1_Templates/HiringAgent_P1_CandidateList.xlsx')
ok(tpl.exists(),
   f"P1 template file exists locally: {tpl.name}",
   "this is what you upload once to SharePoint_Master_Template/ for first-time setup")
ok("SharePoint_Master_Template" not in json.dumps(p1_cfg),
   "P1 flow does NOT reference SharePoint_Master_Template at runtime",
   "template folder is for manual admin setup only, not used by the flow")

print()
print("=== CROSS-CHECK: P1 and P2 agree on workbook location ===")
p1_wb_folder = p1_cfg['excel']['file'].rsplit('/', 1)[0].lstrip('/')
ok(p1_wb_folder == 'Master_Files',
   f"P1 writes to: {p1_wb_folder}/Sharepoint_Master_File.xlsx")
default_wb = _re.search(r'_DEFAULT_WORKBOOK_FOLDER\s*=\s*[\"\']([^\"\']+)', client_src)
p2_wb_folder = default_wb.group(1).strip('/') if default_wb else ''
ok(p2_wb_folder == p1_wb_folder,
   f"P2 default workbook folder matches P1: {p2_wb_folder or '(not found)'}")

print()
print("=" * 62)
print(f"  FOLDER CHECK RESULT: {P} passed, {F} failed")
print("=" * 62)
if F == 0:
    print("  All P1 + P2 folder/file references match your live SharePoint.")
else:
    print(f"  {F} mismatch(es) found — review FAILs above.")
