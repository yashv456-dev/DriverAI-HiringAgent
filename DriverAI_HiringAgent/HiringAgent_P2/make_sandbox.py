"""
Build a throwaway sandbox so P2 can be exercised without touching the live
SharePoint master or emailing a real candidate.

Everything it writes lands in ./sandbox/ and nothing in here talks to Graph.

    python make_sandbox.py

Produces:
    sandbox/Sandbox_Master.xlsx   CandidateList + Rejected, real table objects,
                                  same column contract as the live master, but
                                  seeded only with the operator's own addresses
    sandbox/resumes/*.pdf|.docx   two resumes matching those two rows
    sandbox/.env.sandbox          local-only env with mail hard-off

The schema is read from P2_Preview/generation_39/P2-MasterFile.xlsx so it stays
correct by construction -- HEADERS below are only the fallback for a fresh clone
that has no preview generation on disk. Candidate data is never copied from it.
"""
from __future__ import annotations

import json
import shutil
from pathlib import Path

import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo

P2 = Path(__file__).resolve().parent
SANDBOX = P2 / "sandbox"
SCHEMA_SOURCE = P2 / "P2_Preview" / "generation_39" / "P2-MasterFile.xlsx"

# Fallback column contract, used only when SCHEMA_SOURCE is absent.
HEADERS = {
    "CandidateList": (
        "HiringAgent_P1_Candidates",
        ["Application ID", "Received Date", "Last Updated Date", "Category", "Resume Link",
         "Full Name", "Email", "Phone", "Location", "Country", "Years Exp", "Current Skills",
         "Education", "Education Start Date", "Education End Date", "Looking For Role",
         "Suggested Role 1", "Suggested Role 2", "Suggested Role 3", "Mail Subject",
         "Mail Body", "Status", "Has Resume", "Original Filename", "Application Updates",
         "Retry Count", "Portfolio 1", "Portfolio 2", "Portfolio 3", "Resume URL",
         "Resume Folder Path", "Mail Sent", "Info Request Sent"],
    ),
    "Rejected": (
        "RejectedCandidates",
        ["Application ID", "Received Date", "Last Updated Date", "Category", "Resume Link",
         "Full Name", "Email", "Phone", "Location", "Country", "Years Exp", "Current Skills",
         "Education", "Education Start Date", "Education End Date", "Looking For Role",
         "Suggested Role 1", "Suggested Role 2", "Suggested Role 3", "Mail Subject",
         "Mail Body", "Status", "Has Resume", "Original Filename", "Application Updates",
         "Retry Count", "Portfolio 1", "Portfolio 2", "Portfolio 3", "Resume URL",
         "Resume Folder Path", "Decline Sent"],
    ),
}

# The demo job description, injected into jd_roles_cache.json as an extra role so
# the scorer has something deliberate to compare the two candidates against. Same
# shape as the 105 real roles: lowercase skill tokens, which is what the matcher
# compares on.
DEMO_JD = {
    "title": "DEMO - Senior AI/ML Engineer, Agentic Systems",
    "skills": [
        "python", "machine learning", "pytorch", "llm", "rag",
        "langchain", "mcp", "fastapi", "docker", "kubernetes",
        "aws", "ci/cd", "vector database", "prompt engineering", "mlops",
    ],
    "opening": "DEMO - Senior AI/ML Engineer, Agentic Systems",
    "opening_id": "demo-senior-aiml-agentic",
    "doc_type": "Position Description",
}

# Both addresses belong to the operator, so a stray send lands in their own inbox.
#
# BOTH ARE US-BASED ON PURPOSE. An earlier version made one of them India-based to
# exercise the geo gate, but a non-US candidate is Rejected BEFORE role scoring runs,
# so there would be nothing to compare -- which is the whole point of the demo. The
# contrast is in the SKILLS instead: one is a close match to DEMO_JD, the other is a
# data engineer who overlaps only on the generic terms.
#
# Only app_id / name / email / filename reach the workbook -- see build_workbook.
# Everything else here exists solely to write the resume TEXT, which is where P2
# must read it back from.
CANDIDATES = [
    {
        "app_id": "APP-20260918-0901-T001",
        "name": "Akhil Testcase",
        "email": "akhilme008@gmail.com",
        "phone": "(470) 940-8044",
        "location": "Austin, TX",
        "country": "United States",
        "years": "5",
        # Strong match: hits 12 of the 15 DEMO_JD skills.
        "skills": ("Python, PyTorch, LLM fine-tuning, RAG pipelines, LangChain, MCP servers, "
                   "FastAPI, Docker, Kubernetes, AWS, CI/CD, Pinecone vector database, "
                   "prompt engineering, MLOps, machine learning"),
        "education": "M.S. Artificial Intelligence, University of North Texas",
        "looking_for": "Senior AI/ML Engineer",
        "filename": "akhil_testcase_resume.pdf",
    },
    {
        "app_id": "APP-20260918-0902-T002",
        "name": "Akhil Dataside",
        "email": "akhilme0090@gmail.com",
        "phone": "(312) 555-0184",
        "location": "Chicago, IL",
        "country": "United States",
        "years": "4",
        # Weak match: overlaps on python/machine learning/docker only.
        "skills": ("Python, SQL, Apache Spark, Airflow, dbt, Snowflake, Tableau, "
                   "Docker, ETL pipelines, dimensional modeling, machine learning"),
        "education": "B.S. Information Systems, University of Illinois Chicago",
        "looking_for": "Data Engineer",
        "filename": "akhil_dataside_resume.docx",
    },
]

RESUME_BODY = """{name}
{email} | {phone} | {location}

SUMMARY
{years} years building and deploying production machine learning systems.
Synthetic record generated by make_sandbox.py for pipeline testing only.

SKILLS
{skills}

EXPERIENCE
Test Employer | {looking_for} | {location} | January 2024 - Present
- Built retrieval-augmented pipelines over internal document stores.
- Deployed model serving with autoscaling and request-level tracing.
- Wrote evaluation harnesses with golden datasets and regression gates.

EDUCATION
{education}
"""


def _schema():
    """Column contract from the preview master when present, else HEADERS."""
    if not SCHEMA_SOURCE.exists():
        print(f"  ! {SCHEMA_SOURCE.name} not found, using built-in headers")
        return HEADERS
    wb = openpyxl.load_workbook(SCHEMA_SOURCE)
    out = {}
    for ws in wb.worksheets:
        table = next(iter(ws.tables), HEADERS.get(ws.title, ("", []))[0])
        # Header row only. No candidate rows are read from the live workbook.
        out[ws.title] = (table, [c.value for c in ws[1]])
    wb.close()
    return out


def build_workbook(path: Path) -> None:
    schema = _schema()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for sheet_name, (table_name, headers) in schema.items():
        ws = wb.create_sheet(sheet_name)
        ws.append(headers)

        if sheet_name == "CandidateList":
            col = {h: i for i, h in enumerate(headers)}
            for c in CANDIDATES:
                # Only the columns P1 actually writes. Verified against
                # HiringAgent_P1/flow/definition.json: P1 reads the email
                # envelope and never opens the attachment, so every
                # resume-derived field below stays empty and P2 has to derive
                # it from the PDF/DOCX. Pre-filling them would let a scoring
                # run read back its own seed data and look like it worked.
                row = [""] * len(headers)
                row[col["Application ID"]] = c["app_id"]
                row[col["Received Date"]] = "2026-09-18T09:01:00"
                row[col["Last Updated Date"]] = "2026-09-18T09:01:00"
                # Display-name half of the From header, not the resume header.
                row[col["Full Name"]] = c["name"]
                # Address half of the From header, lowercased, as P1 does.
                row[col["Email"]] = c["email"].lower()
                row[col["Mail Subject"]] = f"Application for {c['looking_for']}"
                row[col["Mail Body"]] = "Please find my resume attached."
                row[col["Has Resume"]] = "Yes"
                row[col["Original Filename"]] = c["filename"]
                row[col["Application Updates"]] = ""
                # The handoff value P2's queue selects on.
                row[col["Status"]] = "New Email Received"
                ws.append(row)

        # A real table object, because the P1 contract is a table and the
        # Excel store addresses rows through it.
        last_col = openpyxl.utils.get_column_letter(len(headers))
        ref = f"A1:{last_col}{max(ws.max_row, 2)}"
        table = Table(displayName=table_name, ref=ref)
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(table)

        for i, h in enumerate(headers, start=1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = max(
                12, min(len(str(h)) + 4, 40))

    wb.save(path)


def build_resumes(folder: Path) -> None:
    import docx
    import fitz

    for c in CANDIDATES:
        text = RESUME_BODY.format(**c)
        target = folder / c["filename"]

        if target.suffix == ".docx":
            d = docx.Document()
            for line in text.splitlines():
                d.add_paragraph(line)
            d.save(target)
        else:
            doc = fitz.open()
            page = doc.new_page()
            page.insert_text((60, 70), text, fontsize=10, fontname="helv")
            doc.save(target)
            doc.close()


ENV_SANDBOX = """# Sandbox env -- load this BEFORE running anything you don't want going live.
#   set -a; source sandbox/.env.sandbox; set +a
#
# The credentials below are blanked ON PURPOSE, and it has to be done this way.
# config.py calls load_dotenv(BASE_DIR/".env"), which backfills any variable that
# is not already present in the environment -- so simply *omitting* TENANT_ID here
# would let the real .env supply it and SHAREPOINT_CONFIGURED would come back True.
# python-dotenv skips a key that is already set even when its value is empty, so
# exporting these as empty is what actually blocks the live tenant.
# Verified: with this file sourced, SHAREPOINT_CONFIGURED is False.
TENANT_ID=
CLIENT_ID=
CLIENT_SECRET=
SHAREPOINT_HOSTNAME=

# Mail: every outbound path off.
HIRING_SUPPRESS_EMAILS=true
HIRING_GEO_REJECT_EMAIL=false
HIRING_ERROR_EMAIL=false
HIRING_DRY_RUN=true

# Storage: a sandbox database, never the P2-root candidates.db.
HIRING_STORAGE_BACKEND=sqlite
HIRING_SQLITE_PATH=sandbox/sandbox.db

# Keep the geo gate on -- the two seeded rows exist to exercise both branches.
HIRING_GEO_USA_ONLY=true

# Gemini stays ON here. GEMINI_API_KEY is left to backfill from the real .env.
# The two sandbox resumes are synthetic, so nothing real leaves the machine, but
# this IS a live call to Google and it does draw on the free-tier quota.
HIRING_GEMINI_ENABLED=true
"""


def inject_demo_jd() -> str:
    """Add DEMO_JD to the role cache, idempotently. Returns a status line.

    The cache path is not configurable (config.py:376 pins it to BASE_DIR), so the
    demo role has to go into the real jd_roles_cache.json. It is prefixed 'DEMO -'
    so it is obvious in any scoring output, and re-running this never duplicates it.

    To remove it later: delete the entry, or let jd_sources rebuild the cache from
    the SharePoint JD folder, which overwrites the whole roles list.
    """
    cache_path = P2 / "jd_roles_cache.json"
    cache = json.loads(cache_path.read_text())
    roles = cache["roles"]
    for i, r in enumerate(roles):
        if r.get("opening_id") == DEMO_JD["opening_id"]:
            roles[i] = DEMO_JD
            cache_path.write_text(json.dumps(cache, indent=2))
            return f"  jd_roles_cache.json  demo role refreshed ({len(roles)} roles)"
    roles.append(DEMO_JD)
    cache_path.write_text(json.dumps(cache, indent=2))
    return f"  jd_roles_cache.json  demo role added ({len(roles)} roles)"


def main() -> None:
    SANDBOX.mkdir(exist_ok=True)
    resumes = SANDBOX / "resumes"
    if resumes.exists():
        shutil.rmtree(resumes)
    resumes.mkdir()

    wb_path = SANDBOX / "Sandbox_Master.xlsx"
    build_workbook(wb_path)
    build_resumes(resumes)
    (SANDBOX / ".env.sandbox").write_text(ENV_SANDBOX)
    jd_line = inject_demo_jd()

    print(f"  {wb_path.relative_to(P2)}")
    for c in CANDIDATES:
        print(f"  sandbox/resumes/{c['filename']}  <- {c['email']} ({c['country']})")
    print("  sandbox/.env.sandbox")
    print(jd_line)
    print("\nsandbox built. Load it with:  set -a; source sandbox/.env.sandbox; set +a")


if __name__ == "__main__":
    main()
