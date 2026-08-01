# HiringAgent P2 â€” Complete Reference

**Engine:** `hiring_agent/` package â€” one codebase for all three run modes.  
**Single business-logic config:** `config.yaml` â€” roles, skills, scoring thresholds, Ollama settings, geo filter.  
**Credentials config:** `.env` â€” only needed for online (SharePoint) mode. See `.env.example`.

Phase 2 reads resumes, extracts candidate fields, scores role fit, geo-filters non-USA, and writes results â€” either to a local Excel workbook or back to the Phase 1 SharePoint row. Phase 1 (`../HiringAgent_P1/`, the Power Automate email flow) owns intake replies; Phase 2 only sends its two scoring-stage Graph emails: non-USA declines and missing-info nudges.

The desktop app is now a SharePoint-first operations console. Local workbook flows still exist, but they live in `bot.py` rather than the main GUI.

For client laptop deployment, share/keep guidance, schedule changes, app-vs-bot usage, columns, folder creation, and resume naming, use the step-by-step checklist in [`P2_CLIENT_DEPLOYMENT_RUNBOOK.md`](P2_CLIENT_DEPLOYMENT_RUNBOOK.md).

> **Lineage:** Replaces the old Phase 2 (`../archive/HiringAgent_P2/`, Outlook/SharePoint engine) and old Phase 3 (`../archive/HiringAgent_P3/`, GUI over it). Their Outlook code was deliberately dropped.

---

## Folder structure

```
HiringAgent_App_P2/
â”œâ”€â”€ config.yaml               â† EDIT THIS: roles, skills, scoring, geo, Ollama, templates
â”œâ”€â”€ .env.example              â† copy to .env for online (SharePoint) mode
â”œâ”€â”€ .env                      â† live credentials (never commit â€” already in .gitignore)
â”œâ”€â”€ jd_sources.json           â† runtime: JDs built by the GUI (optional)
â”œâ”€â”€ app_settings.json         â† runtime: GUI state (Ollama model/host, toggles)
â”‚
â”œâ”€â”€ app.py                    â† Desktop operations console (6 tabs, SharePoint-first)
â”œâ”€â”€ bot.py                    â† CLI worker + local utilities + SharePoint diagnostics
â”œâ”€â”€ Launch.bat                â† Windows launcher (auto-creates venv + installs deps)
â”‚
â”œâ”€â”€ function_app.py           â† Azure Functions host (timer + HTTP triggers)
â”œâ”€â”€ server.py                 â† HTTP server for Docker containers (ThreadingHTTPServer)
â”œâ”€â”€ host.json                 â† Azure Functions host config (v2, extension bundle 4.x)
â”œâ”€â”€ sharepoint_client.py      â† Graph API client: app-only auth, workbook/file I/O
â”‚
â”œâ”€â”€ requirements.txt          â† Core + cloud deps (no GUI; what Azure / Docker installs)
â”œâ”€â”€ requirements-desktop.txt  â† GUI extras (customtkinter, python-dotenv)
â”‚
â”œâ”€â”€ .dockerignore             â† excludes .env, P2_Output/, P2_Logs/, etc.
â”œâ”€â”€ Dockerfile                â† Container image (defaults to server.py HTTP server)
â”œâ”€â”€ .github/workflows/score.yml  â† GitHub Actions scheduled worker (free, no card)
â”‚
â”œâ”€â”€ hiring_agent/             â† Engine package
â”‚   â”œâ”€â”€ config.py             â† Loads config.yaml + env vars; paths, logger, constants
â”‚   â”œâ”€â”€ extraction.py         â† Text extraction, field parsing, AI/Ollama/offline
â”‚   â”œâ”€â”€ scoring.py            â† Role matching, category rules, Ollama dispatcher
â”‚   â”œâ”€â”€ ollama_scorer.py      â† Ollama JSON API integration
â”‚   â”œâ”€â”€ intake.py             â† Local folder scoring, Excel upsert
â”‚   â”œâ”€â”€ local_scorer.py       â† Single-file scoring, candidate browser
â”‚   â”œâ”€â”€ geo.py                â† USA-only filter (offline rules + Nominatim)
â”‚   â”œâ”€â”€ excel_output.py       â† Excel write helper, locked-file rescue
â”‚   â”œâ”€â”€ jd_sources.py         â† JD config load/save, URL fetch, text parse
â”‚   â””â”€â”€ sharepoint_scoring.py â† SharePoint queue worker
â”‚
â”œâ”€â”€ P2_Output/                â† runtime-created local CLI workbooks: <Year>/HiringAgent_P1_CandidateList.xlsx
â”œâ”€â”€ P2_Final_Results/         â† runtime-created client-facing exports: Candidates sheet only
â”œâ”€â”€ P2_Input/                 â† runtime-created optional local resume copies: <Year>/<Month>/
â””â”€â”€ P2_Logs/                  â† runtime-created P2.log (rotating, 1MB x 5 backups)
```

---

## Setup (local)

1. **Python 3.10+** required.
2. **Optional â€” Ollama** (free local brain, no API key):
   ```
   ollama pull llama3.2
   ```
   Without Ollama the app runs the keyword scorer â€” all modes still work.
3. **Install deps:**
   ```
   pip install -r requirements.txt          # CLI / cloud / engine
   pip install -r requirements-desktop.txt  # add this for the GUI
   ```
   On Windows you can instead double-click **`Launch.bat`** â€” it creates a venv, installs everything, and opens the app.
4. **Desktop app / online mode** â€” copy `.env.example` â†’ `.env` and fill in the app-only credentials (Entra app registration with Graph `Sites.ReadWrite.All` + `Files.ReadWrite.All`, admin-consented). The refreshed desktop app expects this SharePoint connection. Pure local CLI runs need none of it.

---

## Run it â€” GUI

```
python app.py
```

| Tab | What it does |
|---|---|
| **Home** | Minimal operations cockpit: `Run Now`, `Stop`, queue count, scheduler launcher, recent runs, and compact analytics from run history + SharePoint. |
| **Job Descriptions** | Source manager for URL JDs, SharePoint/OneDrive JD folders, and pasted JD text. Shows full URLs, folder/file details, source status, and SharePoint folder preview results. |
| **SharePoint** | Readable data map for the full resume input URL, workbook URL, workbook file name, table name, sheet names, and live workbook metadata, with raw `.env` fields collapsed below. |
| **Resumes** | Upload local PDF/DOCX files into the online SharePoint queue, or preview how one local resume is read before uploading it. |
| **Candidates** | Browse the live SharePoint candidate table with search, filters, detail view, and export. |
| **Settings** | Ollama model/host, AI + geo toggles, and `Test Ollama`. Saved to `app_settings.json`. |

---

## Run it â€” CLI

```bash
python bot.py --score-folder   ./resumes          # score folder â†’ workbook
python bot.py --score-excel    ./candidates.xlsx  # re-score rows of a workbook
python bot.py --score-file     ./jane_cv.pdf      # print one resume's breakdown
python bot.py --rematch                            # re-score every saved workbook
python bot.py --list-roles                         # show roles being matched
python bot.py --test-sharepoint                    # verify credentials (no writes)
python bot.py --score-sharepoint                   # online: score Phase 1 queue + write back
python bot.py --score-sharepoint --watch --interval 300  # poll every 5 min
python bot.py --validate-row APP-20260630-143052-a3f9c1  # inspect one stored candidate
python bot.py --recheck-row APP-20260630-143052-a3f9c1   # targeted live repair for one row
python bot.py --recheck-all                         # maintenance sweep of the full workbook
python bot.py --recover-location-rejections --dry-run  # safe historical geo audit
python bot.py --recover-location-rejections         # apply geo recovery; never sends declines
python audit_all_resume_files.py                    # read-only: validate Resume URL/folder filenames
python audit_sharepoint_results_health.py           # read-only: duplicates, geo placement, info/decline markers
```

Flags: `--dry-run` (read + log, no writes) Â· `--no-ai` (force keyword scorer) Â· `--ai` (force Ollama on) Â· `--model NAME` Â· `--host URL` Â· `--no-geo` (score everyone) Â· `--excel FILE` (pin workbook).

---

## Pipeline (per resume)

```
PDF/DOCX
  â†“ extract text   (pypdf / python-docx; optional OCR if pymupdf + pytesseract present)
  â†“ extract fields â†’ Ollama local brain (if running)
                   â†’ offline parser (always available â€” regex + keyword scan)
  â†“ merge mail body fallback (fill "Not extracted" fields from the email body text)
  â†“ AI recheck pass (Ollama validates ambiguous fields; no-op if Ollama off)
  â†“ score role fit â†’ Ollama brain   â†’ role 1, role 2, one-line reason
                   â†’ keyword scorer (fallback) â†’ same output
  â†“ geo filter (USA-only), precedence (reordered 2026-07-12, corroboration added 2026-07-24):
      strong US location (state / ZIP / territory) wins first â€” overrides a
        mis-extracted foreign country (e.g. Country parsed 'India' off a PAST
        foreign degree while the candidate is a current US master's student)
      â†’ foreign location match now also requires Phone or Education corroboration
        before it actually rejects (see `_corroborates_non_usa()` in `geo.py`): a
        location-text match alone (which extraction can pick up from anywhere in the
        resume, including a stale hometown/old-degree mention) is no longer treated as
        proof of current whereabouts. Phone reading as a US number, or Education/resume
        context reading as US-based, CONTRADICTS the match â†’ kept for clarification
        instead of rejected. Neither Phone nor Education carrying any signal at all â†’
        kept (benefit of the doubt, insufficient evidence to reject). Only rejected when
        the foreign location match has at least one non-contradicting corroborating signal.
      â†’ a country field without a confirmed current location remains UNKNOWN.
        Country, phone, education, and work history are supporting evidence; they
        prevent an unsafe rejection but do not manufacture a current residence.
      â†’ bare US city â†’ Nominatim online lookup â†’ keep (benefit of doubt)
    A US phone rescues a US-based applicant whose hometown/country were mis-read from
    a past foreign degree (Manasa Puli â€” current job 'Georgia (Remote)', US phone, but
    parser read 'Hyderabad'/'India'). The `_looks_like_us_phone()` check was added
    2026-07-12 but not actually wired into the country-field branch until 2026-07-24 â€”
    it was accepted as a parameter and passed in by the live pipeline the whole time,
    but silently never read inside `check_location_usa()`. Fixed and covered by a
    regression test (`test_p2.py` Â§O2b) so it can't silently drop out again. The same
    NANP check now also excludes Canadian area codes (e.g. Ontario's 519) so a genuinely
    Canadian candidate's phone doesn't misread as "looks like a US phone" (Â§O2b/Â§O2c).
  â†“ write result
      Local mode  â†’ upsert row in P2_Output/<Year>/HiringAgent_P1_CandidateList.xlsx
      Online mode â†’ PATCH the same SharePoint row; confirmed non-USA â†’ Rejected sheet;
                    unknown/conflicting geo â†’ Main as Needs Review - Location Confirmation;
                    scored/review rows missing required details â†’ missing-info nudge
```

**Brain fallback order:** every step falls back gracefully. If Ollama isn't running, the keyword scorer takes over â€” the run never fails because of a missing AI.

---

**Current geo clarification policy (updated 2026-07-28):** P2 makes one of three decisions: `confirmed_us`, `confirmed_non_us`, or `unknown`. Only `confirmed_non_us` moves to `Rejected - Non-USA Location`. Missing or conflicting current-location evidence stays on CandidateList as `Needs Review - Location Confirmation`; it is excluded from the client results workbook until confirmed. A US phone, university, or employer is useful supporting evidence and prevents an unsafe rejection, but it does not prove current residence or cause P2 to write `United States` into Location/Country. P2 no longer converts unresolved candidates to `Rejected - Location Not Confirmed` after a follow-up attempt.

## Two run styles

### Local CLI utilities

Input is a folder of resumes or an `.xlsx` you pick through `bot.py`. Output is a local workbook under `P2_Output/`. No credentials required. Resume copies go to `P2_Input/<Year>/<Month>/` if `HIRING_SAVE_LOCAL_COPIES=true`.

### Online â€” SharePoint (desktop app default, also available in CLI)

Input is the Phase 1 SharePoint queue: rows with `Status = "New Email Received"` and their associated resume files. Output is **written back to the same SharePoint row** — Full Name, Phone, Location, Country, Current Skills, Education, Looking For Role, Suggested Role 1/2/3, Category, Portfolio 1/2/3, Resume URL, Resume Link, Resume Folder Path, Info Request Sent. Confirmed US rows become `Scored`; unknown geo rows become `Needs Review - Location Confirmation`; confirmed non-USA rows move to the workbook's `Rejected` sheet.

`Resume Link` is a SharePoint calculated/display column. The visible label is the P1-owned `Original Filename` when present, with the stored SharePoint URL as the target. P2 may rename/move the physical resume file for storage, but it must not replace `Original Filename` with the canonical storage filename.

After each live SharePoint scoring run, P2 also writes a clean client-facing workbook under `P2_Final_Results/` and uploads `Candidate_List_Results.xlsx` directly to the root `Shared Documents` library (`/Candidate_List_Results.xlsx`), outside the master file folder and resume tree. That export has one sheet (`Candidates`), excludes the Rejected sheet, includes only scored main-sheet candidates, and keeps only the selected presentation columns.

**Period separator rows (added 2026-07-31).** `_with_period_separators()` always inserts one fully blank row directly under the header, including when the export has no candidates, and one between each calendar month, so the client workbook reads the same way as P1's CandidateList sheet. A year change is also a month change, so a year boundary produces **one** blank row, never two stacked. Separator rows are blank in every column and carry no `Application ID`, so P1, P2 and every audit script skip them the same way they skip P1's own separators. A row whose `Received Date` is blank or unparseable is treated as having no period at all, so it can't inject a spurious separator on either side of itself. Covered by `test_p2.py` Â§W2.

Because Power Automate runs in the cloud and can't push to your machine, polling is used (Home dashboard / scheduler window / `--watch` / a scheduled job). Row writes are **surgical** (one PATCH per row) â€” Power Automate can keep appending at the same time.

**Duplicate rows collapse to one (changed 2026-07-12; guarded 2026-07-16).** Every scoring pass (and `--recheck-all`) reconciles duplicate candidate rows: two rows are treated as the **same person and merged into one** if they share the **same email OR the same phone number** â€” never on name alone (two different people can share a name; a name match with a *different* email **and** a *different* phone is kept as two separate rows). The row with the more recent `Received Date` wins; the older row and its resume file are deleted, leaving one row + one resume per person. Active queue rows are deliberately excluded from this merge while they are still `New Email Received`, blank, or `Needs Review - Unreadable Resume`; P2 scores the fresh update first, then later duplicate cleanup may collapse only rows that already have final processing state. This prevents an older scored row from swallowing a newly arrived resume update before the new CV is read.

Duplicate merge healing is limited to scored/profile fields. P1-owned identity, date, mail, original filename, and resume URL/path fields are never copied from the older duplicate into the newer winner.

---

## Happy / sad paths

| Situation | What happens |
|---|---|
| New resume folder run | Each PDF/DOCX extracted, scored, upserted into workbook; live log streams |
| Re-score an Excel | Brain re-scores Suggested Role 1/2/3 for every row in place |
| Score one file | Instant per-role skill breakdown (matched / missing); nothing written |
| SharePoint run | Reads queue, scores each resume, writes results back to the same row |
| Browse candidates | The live SharePoint candidate table is shown in a sortable table in the GUI |
| Ollama not installed / not running | Falls back to the keyword scorer (warning logged) |
| Broken or scanned PDF text layer | Tries pypdf page-by-page, then `pdfplumber` fallback when pypdf returns no readable text, then optional OCR if PyMuPDF + pytesseract + Tesseract are present |
| Workbook open in Excel during a run | Writes a `*_locked_YYYYMMDD-HHMMSS.xlsx` rescue file; path shown in log |
| Non-USA location (geo filter on) | Row -> Rejected sheet as `Rejected - Non-USA Location`; optional polite decline email sent via Graph |
| Conflicting/unclear current location | Row stays on CandidateList as `Needs Review - Location Confirmation`; it is not exported as Scored and no decline is queued |
| Still unclear after an updated resume reply | Row remains in location review for manual confirmation; uncertainty alone never becomes a rejection |
| Scored row missing phone/location/skills/education | One-time missing-info request sent via Graph; tracked by `Info Request Sent` |
| SharePoint row has no readable resume text | Marked `Needs Review - Unreadable Resume`, retried up to `score_retry_max`, then moved to Rejected as `Rejected - Processing Error` without sending the normal non-USA decline |
| Spam/vendor content reaches a row despite P1's gate | Marked `Needs Review - Possible Spam` **before any extraction runs** (no scoring, no resume rename/move) and an alert is emailed to `admin_email`. See "Secondary content-safety net" below |
| Flow action / SharePoint API error | Error logged, that row skipped, run continues; row stays unscored for retry |

### Secondary content-safety net (added 2026-08-01)

P1 screens incoming mail against 114 phrases before a row is ever created ([`../../HiringAgent_P1/docs/README.md`](../../HiringAgent_P1/docs/README.md), Gate 3). This is P2's backstop for the rare row that still slips through — **not** a second spam filter, and deliberately much smaller: 12 phrases in `config.yaml` under `content_safety.suspicious_phrases` (vendor-pitch + unambiguous phishing only), matched against the resume text **and** the mail body.

On a match P2 sets `Needs Review - Possible Spam`, emails `admin_email`, and skips the row entirely — it is never scored, its resume is never renamed or moved, and it never reaches the Rejected sheet.

**What it deliberately does not do:** P2 has no mailbox write access, so it cannot move the offending mail to Junk or mark it unread — P1 already moved it to `Archive` (read) at intake. Cleanup of the mail itself is manual today. Automating it would require storing the source message ID on the row at intake, since matching a row back to its message by subject/sender/date is guesswork and a wrong match would junk a real applicant's mail.

**Keep the phrase list short.** Every phrase runs against every genuine resume; a loose phrase silently withholds a real candidate from scoring. Validate any addition against the live resumes before enabling it.

---

## Config reference

### `config.yaml` â€” business logic (the only config file to edit)

| Section | Key fields |
|---|---|
| `company_name` | Display name used in email templates |
| `default_roles` | List of `{title, skills[]}` â€” the roles candidates are scored against |
| `skills.keywords` | 180+ skill terms for extraction and scoring |
| `skills.display_names` | Proper casing map (e.g. `postgresql`, `c++`) |
| `scoring.min_match_percent` | Minimum % to include a role in output (default: 20) |
| `scoring.top_n` | Number of roles to return (default: 3) |
| `scoring.title_boost` | Extra % when candidate's stated role matches (configured: 40; code fallback is 10 only if this key is absent from `config.yaml`) |
| `scoring.batch_limit` | Maximum candidates scored per run to prevent timeouts (default: 25) |
| `role_categories.rules` | First-match-wins title substring rules mapping `Suggested Role 1` to a department Category (see Pipeline Step 4 / Category Mapping). **Exception (2026-07-24):** an unambiguous native-mobile skill stack (Kotlin/Swift/Flutter/etc.) in `Current Skills` overrides these title rules straight to `Mobile Apps (Android IOS)` — not config-driven, see `assign_category()` in `hiring_agent/scoring.py` and the comment above `role_categories:` in `config.yaml` |
| `ai_extraction.ollama.enabled` | `true` by default locally, `false` in serverless |
| `ai_extraction.ollama.model` | e.g. `llama3.2` |
| `ai_extraction.ollama.host` | e.g. `http://localhost:11434` |
| `ai_extraction.ollama.scoring` | Use Ollama for role scoring too (not just extraction) |
| `geo_filter_usa_only` | `true` = keep USA, reject non-USA |
| `geo_reject_email` | `true` = send polite decline email to rejected candidates |
| `error_email` / `score_retry_max` | Admin alert + retry cap for rows that repeatedly fail processing |
| `email_templates` | Subject + body for `rejection_non_usa` and `missing_info`; `acknowledgment` / `request_cv` are legacy/dead config for P2 |
| `filters.*` | 1,400+ geo terms: `foreign_countries`, `foreign_cities`, `foreign_regions`, `foreign_demonyms`, `known_us_cities` |


### Environment variables (override any `config.yaml` value)

| Variable | Effect |
|---|---|
| `HIRING_OLLAMA_ENABLED` | `true`/`false` â€” force Ollama on or off |
| `HIRING_OLLAMA_SCORING` | `true`/`false` â€” Ollama for scoring (not just extraction) |
| `HIRING_OLLAMA_MODEL` | Model name override |
| `HIRING_OLLAMA_HOST` | Host URL override |
| `HIRING_GEO_USA_ONLY` | `true`/`false` â€” geo filter toggle |
| `HIRING_GEO_REJECT_EMAIL` | `true`/`false` â€” send decline emails (this flag alone does **not** silence P2 â€” see below) |
| `HIRING_ERROR_EMAIL` | `true`/`false` â€” send admin alert after processing retry cap |
| `HIRING_SUPPRESS_EMAILS` | `true` â€” **master kill-switch: no P2 email of any kind leaves the mailbox**, everything else runs normally |
| `HIRING_ADMIN_EMAIL` | Admin alert recipient override |
| `HIRING_SCORE_RETRY_MAX` | Retry cap before a row is moved to Rejected as processing error |
| `HIRING_EXCEL_FILE` | Pin all reads/writes to one specific workbook |
| `HIRING_SAVE_LOCAL_COPIES` | `true` â€” download + save resume copies locally (off by default) |
| `HIRING_SERVERLESS` | `true` â€” console-only logging, temp scratch dir, Ollama off |
| `HIRING_SCORING_BATCH_LIMIT` | Max candidates scored per run to prevent cloud execution timeouts (e.g. 25) |


### `.env` â€” SharePoint credentials (online mode only)

```
TENANT_ID               real tenant GUID (NOT 'common')
CLIENT_ID               app registration Application ID
CLIENT_SECRET           client secret value
SHAREPOINT_HOSTNAME     e.g. led1234567.sharepoint.com
SHAREPOINT_SITE_PATH    e.g. /sites/CandidateList_HiringAgent
SHAREPOINT_TABLE        e.g. HiringAgent_P1_Candidates
SHAREPOINT_RESUMES_FOLDER  e.g. /Candidate_Resumes
# Optional:
SHAREPOINT_WORKBOOK     explicit workbook path; defaults to /P1P2_SharePoint_Master_Files/Sharepoint_Master_File.xlsx (single workbook, no <Year> subfolder)
SENDER_MAILBOX          P2 decline + missing-info emails sent FROM here (default: apply@driverai.io); needs Mail.Send Application permission
HIRING_P2_DISABLED      true = all P2 scoring entrypoints no-op; use while repairing SharePoint data
HIRING_ENABLE_P2_TIMER  true = Azure timer trigger is allowed to run; without it timer invocations no-op
```

### JD sources â€” one-time scan + cache (`jd_sources.json` â†’ `jd_roles_cache.json`)

- **Scanned once, then cached â€” not re-downloaded every run.** Downloading + parsing every JD in the folder is a **one-time** step. Trigger it from the app's **Job Descriptions** tab (scanning the SharePoint folder builds the cache) or headless with `python bot.py --refresh-jd`. That writes `jd_roles_cache.json` (role titles + skills). Every scoring run then reads that cache **instantly** (~1s) instead of re-downloading (~45s).
- **When the cache refreshes:** only when you re-scan (app / `--refresh-jd`), or when the JD *configuration* changes (folder path, URLs, pasted JDs â€” tracked by a signature). **Adding a new JD file to the same folder does _not_ auto-refresh** â€” re-scan to pick it up. On a cache miss the next run fetches live once and re-warms the cache automatically.
- **Cloud:** commit `jd_roles_cache.json` so fresh GitHub Actions / Docker checkouts start warm (it holds only role titles + skills â€” no candidate data, safe to commit). Without it, a cloud run with candidates fetches live once and caches to ephemeral disk.
- **Environment-specific source.** The committed `jd_sources.json` points at one tenant's personal OneDrive (`*-my.sharepoint.com/personal/...`). A fork or different tenant **silently falls back to `config.yaml`'s `default_roles`** if that folder isn't reachable or the app lacks **`Files.Read.All` (Application, admin-consented)** â€” the run never fails. Watch the log for `Matching against N JD role(s)` vs. the fallback warning.
- To score against built-in roles only (fast, no network, no cache), set `"enabled": false` in `jd_sources.json`.

---

## Cloud deploy (serverless, machineless)

All three serverless options use the same credentials (`.env` values) and the same engine. Pick one.

### Option A â€” GitHub Actions (free, no credit card)

The scheduled workflow at `.github/workflows/score.yml` runs `bot.py --score-sharepoint` on GitHub's runners.

1. Make `HiringAgent_App_P2/` the repo root:
   ```bash
   cd HiringAgent_App_P2
   git init && git add . && git commit -m "Hiring agent worker"
   gh repo create hiring-agent --private --source=. --push
   ```
2. Add credentials as **repo Secrets** (Settings â†’ Secrets and variables â†’ Actions â†’ New repository secret): `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`, `SHAREPOINT_HOSTNAME`, `SHAREPOINT_SITE_PATH`, `SHAREPOINT_TABLE`, `SHAREPOINT_RESUMES_FOLDER`. Optional: `HIRING_GEO_USA_ONLY` variable.
3. **Change the schedule:** edit the `cron:` line in `.github/workflows/score.yml` â€” e.g. `'*/15 * * * *'` for 15 min, `'0 9,18 * * *'` for 9 AM + 6 PM UTC. Push the change; GitHub picks it up immediately. (GitHub Actions does not support reading the cron value from a variable.)
4. **Test:** Actions tab â†’ *Score SharePoint queue* â†’ **Run workflow** (leave Dry run = true). The `Verify credentials` step prints PASS/FAIL.
5. Once it passes, the schedule runs automatically â€” every ~30 min, $0, no machine.

**Free-tier notes:** public repos have unlimited Actions minutes; private repos get 2,000 min/month. With pip caching each run stays under 1 min, so the 30-min schedule is well inside the limit. GitHub disables scheduled workflows after **60 days of no repo activity** â€” push any commit to re-arm. Resumes live in SharePoint, not the repo, so a public repo is safe.

### Option B â€” Azure Functions (timer + HTTP, needs a credit card on file)

```bash
# 1) Create the Function App (Python 3.11, Consumption plan â€” free grant).
az group create -n HiringAgent-rg -l eastus
az storage account create -n hiringagentsa$RANDOM -g HiringAgent-rg -l eastus --sku Standard_LRS
az functionapp create -g HiringAgent-rg --consumption-plan-location eastus \
   --runtime python --runtime-version 3.11 --functions-version 4 \
   --name hiring-agent-fn --storage-account <storage-account-name>

# 2) Set Application Settings (env vars). SCORE_TIMER_CRON must be set â€” it drives
#    the timer trigger binding. Default: every 15 min (NCRONTAB with seconds field).
az functionapp config appsettings set -g HiringAgent-rg -n hiring-agent-fn --settings \
   TENANT_ID=<guid> CLIENT_ID=<guid> CLIENT_SECRET=<secret> \
   SHAREPOINT_HOSTNAME=led1234567.sharepoint.com \
   SHAREPOINT_SITE_PATH=/sites/CandidateList_HiringAgent \
   SHAREPOINT_TABLE=HiringAgent_P1_Candidates \
   SHAREPOINT_RESUMES_FOLDER=/Candidate_Resumes \
   HIRING_GEO_USA_ONLY=true \
   SCORE_TIMER_CRON="0 */15 * * * *"

# 3) Deploy.
func azure functionapp publish hiring-agent-fn --python
```

**Triggers:**
- **Timer** (`score_timer`): fires on `SCORE_TIMER_CRON` only when `HIRING_ENABLE_P2_TIMER=true` and `HIRING_P2_DISABLED` is not true. Change it: `az functionapp config appsettings set ... --settings SCORE_TIMER_CRON="0 */30 * * * *"` — no redeploy needed.
- **HTTP** (`score_http`): `POST https://hiring-agent-fn.azurewebsites.net/api/score?code=<function-key>`. Add `?dry_run=1` for a test run. Call this from the Phase 1 Power Automate flow (HTTP action â€” premium connector) to score the moment a resume lands instead of waiting for the timer.

**Logs:** Application Insights captures all `logger.*` output.

### Option C â€” Docker (AWS, GCP, Fly.io, Railway, Render, any container host)

```bash
# Build and test locally.
cd HiringAgent_App_P2
docker build -t hiring-agent .
docker run --rm \
  -e TENANT_ID=<guid> -e CLIENT_ID=<guid> -e CLIENT_SECRET=<secret> \
  -e SHAREPOINT_HOSTNAME=led1234567.sharepoint.com \
  -e SHAREPOINT_SITE_PATH=/sites/CandidateList_HiringAgent \
  -e SHAREPOINT_TABLE=HiringAgent_P1_Candidates \
  -e SHAREPOINT_RESUMES_FOLDER=/Candidate_Resumes \
  -e HIRING_GEO_USA_ONLY=true \
  hiring-agent python bot.py --score-sharepoint --dry-run

# Run as HTTP server (default â€” Fly.io, Railway, Render, Cloud Run, App Runner):
docker run -d -p 8080:8080 --env-file .env hiring-agent
# POST http://localhost:8080/score      trigger a scoring run
# GET  http://localhost:8080/health     health check (returns {"status":"ok"})
# POST http://localhost:8080/score?dry=1  dry run

# Or run as a polling worker (AWS ECS scheduled task, plain VPS):
docker run -d --env-file .env hiring-agent \
  python bot.py --score-sharepoint --watch --interval 900
```

**Platform quick starts:**

*AWS ECS / Fargate (scheduled task):*
```bash
aws ecr create-repository --repository-name hiring-agent
docker tag hiring-agent <account>.dkr.ecr.<region>.amazonaws.com/hiring-agent
docker push <account>.dkr.ecr.<region>.amazonaws.com/hiring-agent
# Create task definition + EventBridge rule; override CMD to:
# ["python", "bot.py", "--score-sharepoint"]
# Pass creds as task env vars or Secrets Manager references.
```

*Fly.io (free tier â€” 3 shared VMs):*
```bash
fly launch --name hiring-agent --no-deploy
fly secrets set TENANT_ID=<guid> CLIENT_ID=<guid> CLIENT_SECRET=<secret> \
  SHAREPOINT_HOSTNAME=... SHAREPOINT_SITE_PATH=... SHAREPOINT_TABLE=... SHAREPOINT_RESUMES_FOLDER=...
fly deploy
```

*Railway / Render:* push the repo, set environment variables in the dashboard, deploy. Both detect `Dockerfile` automatically and expose port 8080. Use their built-in cron to `POST /score` every 15â€“30 min.

**Scoring brain in Docker:** the offline regex parser (free, no key, no AI). Ollama is off in the cloud (`HIRING_SERVERLESS=1` is set in the Dockerfile), so extraction and scoring are fully deterministic there.

---

## Testing before you deploy

```bash
python bot.py --list-roles                              # confirm roles load
python bot.py --score-file ./any_resume.pdf             # confirm extraction works
python bot.py --test-sharepoint                         # verify Graph credentials
python bot.py --score-sharepoint --dry-run              # end-to-end, no writes

# Azure Functions local test:
# copy local.settings.json.example â†’ local.settings.json, fill in creds
func start
curl -X POST http://localhost:7071/api/score?dry_run=1

# Docker local test:
docker run --rm --env-file .env hiring-agent python bot.py --test-sharepoint
```

---

## What Phase 2 reads and writes

| | Local mode | Online mode |
|---|---|---|
| **Reads** | Folder of PDF/DOCX you pick; `config.yaml`; `jd_sources.json` | Phase 1 SharePoint workbook rows (Status = "New Email Received") + resume files in SharePoint |
| **Writes** | `P2_Output/<Year>/HiringAgent_P1_CandidateList.xlsx`; optional resume copies in `P2_Input/`; `P2_Logs/P2.log` | Same SharePoint row (PATCH); Rejected sheet (non-USA); optional scorecard `.txt` in SharePoint |
| **Network** | Ollama localhost (optional); Nominatim OSM (1 req/sec, geo verify) | All of the above + Microsoft Graph API (SharePoint); optional Graph `sendMail` for declines and missing-info nudges |

**Outlook mailbox polling is never used in P2.** Phase 1 owns intake replies. Phase 2 reads/writes SharePoint and can send two Graph `sendMail` templates when enabled/needed: non-USA declines and missing-info nudges.

### Silencing P2 for a historical replay

`HIRING_SUPPRESS_EMAILS=true` (or `test_mode.suppress_emails: true` in `config.yaml`) stops **every** outbound P2 email while leaving the rest of the pipeline fully live â€” rows are scored and patched, resumes renamed and moved, the Rejected sheet maintained, the client workbook exported. P1 uses its separate `flow_config.json` setting `email.send_applicant_emails`; the two phases are controlled independently.

It is enforced inside `SharePointClient.send_mail()` â€” the single function every sender calls â€” rather than at each call site. That matters: **`HIRING_GEO_REJECT_EMAIL` and `HIRING_ERROR_EMAIL` are not sufficient on their own.** They cover the non-USA decline and the admin alert, but the missing-info / current-location nudge has never had a flag of its own; it fires whenever a scored row has a gap. Before this switch existed, turning both flags off still let nudges reach real candidates during a replay.

Rows touched while suppressed are stamped `TEST-MODE (suppressed) Sent <ts>` in `Mail Sent` / `Decline Sent` / `Info Request Sent` instead of `Sent <ts>`, so a replayed row can never be mistaken for a candidate who was genuinely contacted. Both forms remain parseable by `_parse_marker_datetime`, so the idempotent decline/nudge passes still treat the row as handled and will not re-send it.

**Set it back to `false` when the replay finishes** â€” otherwise genuine new applicants receive nothing at all. Covered by `test_p2.py` Â§X, which asserts a suppressed send makes zero Graph calls.

---

## Hand-off with Phase 1

Phase 1 (Power Automate) writes rows with `Status = "New Email Received"` and saves resume files to SharePoint. Phase 2 treats that `Status` value as a work queue signal â€” it polls for those rows, scores them, and sets `Status = "Scored"`. The two phases never call each other; the SharePoint workbook is the only contract.

**Config values that must match between P1 and P2:**
- `SHAREPOINT_HOSTNAME` + `SHAREPOINT_SITE_PATH` â†’ must point to the same SharePoint site
- `SHAREPOINT_TABLE` â†’ `HiringAgent_P1_Candidates`
- `SHAREPOINT_RESUMES_FOLDER` â†’ `/Candidate_Resumes`
- Resume display label: P1 owns `Original Filename`; P2 owns the storage URL/path and physical rename. The `Resume Link` formula must display `Original Filename` while linking to `Resume URL`, so client-facing exports show the candidate's submitted filename even after P2's canonical SharePoint storage rename.
- Extraction/scoring guardrails: P2 normalizes list-like skill output, rejects broken phone fragments, maps state/province names in `Country`, falls back to `pdfplumber` when pypdf cannot read a PDF text layer, skips education header fragments such as `AND TRAINING`, filters non-job/admin documents such as offer-letter info out of role matching, and dedupes Suggested Role 1/2/3 by visible title.
- Resume storage path: P1 currently saves resumes under dated `<Year>/<Month>/` subfolders; P2 reconstructs that path from the row's `Received Date` and also falls back to the flat root folder for older files
- Resume file naming (changed 2026-07-15, client request): P2 renames every resume, right after scoring, to `<FirstNameLastName>_<Category>_<tail>.<ext>` (e.g. `JaneSmith_DataAnalytics_A3F9.pdf`) â€” `Category` is PascalCase/alphanumeric-only (`_clean_category_for_filename()`, e.g. `AI/ML/CV (SIN2)` â†’ `AIMLCVSIN2`), and `tail` is the AppID's own last hyphen-segment (`_resume_filename_tail()`, e.g. `A3F9` from `APP-20260710-2200-A3F9`) â€” kept specifically so two same-name-same-category candidates never collide, since the AppID itself is no longer in the visible part of the name. P1 still saves each resume pre-named as `<FirstLast>_<AppID>.<ext>` at intake (Category doesn't exist yet at that point â€” it's a P2-only field), and P2 immediately renames it once scoring assigns a Category. `_stored_resume_names()` tries the current format first, then the prior `<FirstLast>_<AppID>` shape, then the original P1-intake shape, so a file at any stage of this history is still found and re-renamed forward â€” nothing needs manual fixing. A one-time bulk migration (`migrate_resume_filenames.py`) renamed every already-saved resume to the new format on 2026-07-15.
- Rejected candidates (P2-only, no P1 counterpart): when P2 rejects a candidate it physically moves that candidate's resume file(s) into **one `Rejected/` folder per year** (e.g. `.../2026/Rejected/JaneSmith_DataAnalytics_A3F9.pdf`) â€” changed 2026-07-15 from a separate `Rejected/` subfolder nested under every month; now every decline from a given year lands in the same place regardless of which month they applied in. Restoring a candidate to the main sheet moves the file back into its month bucket.
- Sheet display order (added 2026-07-31): `resort_candidate_sheets.py` (project root, standalone/on-demand, `--dry-run` supported) re-sorts the live CandidateList and Rejected SharePoint tables to current-month-first / newest-within-month-first, with a labeled `-- YYYY --` row at year boundaries. Not wired into the scoring pipeline - new rows still append at the bottom between runs, so the order drifts until it's re-run. See `p2_detailed_summary.md` Â§6c for the full mechanism and safety design.
