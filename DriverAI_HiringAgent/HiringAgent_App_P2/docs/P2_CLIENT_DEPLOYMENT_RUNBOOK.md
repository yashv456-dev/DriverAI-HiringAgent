# P2 Client Deployment Runbook

Use this checklist when installing or validating Phase 2 on a client Windows laptop.

Last updated: 2026-07-24

## 0. Supported Systems

P2's scoring engine is cross-platform. The desktop GUI launcher and Windows Task
Scheduler instructions are Windows-specific.

| System | Recommended operation |
|---|---|
| Windows 10/11 client laptop | `Launch.bat` for setup/GUI; `RunDaily.bat` for Task Scheduler |
| macOS | Python virtual environment + `bot.py`; use `launchd` for scheduling |
| Linux | Python virtual environment + `bot.py`; use `systemd`/cron for scheduling |
| Docker/cloud | Use `Dockerfile` and `docs/CLOUD_DEPLOY.md`; Ollama is disabled by default |

Minimum practical requirements: Python 3.11-3.12, internet access to Microsoft Graph,
an admin-consented Entra app registration, and access to the P1 SharePoint workbook and
resume library. Python 3.14 currently works in the validated development environment, but
3.12 is the recommended client version because third-party wheel availability is broader.

## 1. What To Bring

- Latest P2 code zip from:
  `archive\backups\HiringAgent_App_P2_current_backup_*.zip`
- The real `.env` file, copied separately and handled as a secret.
- Client laptop with Windows, internet access, and permission to install Python/Ollama if needed.
- Microsoft Graph app registration already created and admin-consented.
- SharePoint workbook and resume folders already created by P1.

Do not put `.env` in email, chat, Git, or a shared public folder.

## 1A. What To Share And What To Keep Local

Share with the client/CEO:

- `HiringAgent_App_P2_current_backup_*.zip`
- `docs\README.md`
- `docs\P2_CLIENT_DEPLOYMENT_RUNBOOK.md`
- `docs\p2_detailed_summary.md`
- `docs\CLOUD_DEPLOY.md` only if cloud deployment is being discussed
- `.env.example` and `local.settings.json.example`

Keep local/private:

- Real `.env`
- `.venv\`
- `P2_Logs\`
- `P2_Input\`
- `P2_Output\`
- `archive\`
- Any downloaded resumes, audit CSVs, screenshots, or live SharePoint exports

The shareable backup zip intentionally excludes `.env`, `.venv`, logs, local input/output folders, and old backups. The real `.env` should be transferred only through the client's approved secret channel, or recreated on the client laptop from `.env.example`.

## 2. Required Graph Permissions

The app identity in `.env` must have admin consent for the permissions used by P2:

- Read/write SharePoint workbook rows.
- Read/write resume files/folders.
- Send mail from the configured sender mailbox.

Typical application permissions:

- `Sites.ReadWrite.All`
- `Files.ReadWrite.All`
- `Mail.Send`

The sender mailbox in `.env` must match the mailbox allowed to send P2 emails.

## 3. Install Folder

On the client laptop, create:

```powershell
C:\HiringAgent\HiringAgent_App_P2
```

Extract the P2 zip into that folder.

Expected files after extraction:

- `Launch.bat`
- `Start.bat`
- `RunDaily.bat`
- `bot.py`
- `app.py`
- `config.yaml`
- `sharepoint_client.py`
- `hiring_agent\`
- `docs\`

Copy the real `.env` into:

```powershell
C:\HiringAgent\HiringAgent_App_P2\.env
```

## 3A. P2 App vs P2 Bot

Use `app.py` / `Launch.bat` when the client wants a desktop operations screen:

- Run Now / Stop / health checks
- SharePoint connection view
- JD source/cache management
- Candidate browsing
- Local resume preview/upload tools

Use `bot.py` / `RunDaily.bat` when the client wants automation:

- `bot.py --test-sharepoint` checks credentials without writing candidate data.
- `bot.py --score-sharepoint` processes the SharePoint queue once.
- `bot.py --score-sharepoint --watch --interval 300` keeps polling every 5 minutes.
- `RunDaily.bat` is the Windows Task Scheduler entrypoint.

In focused client mode, `Start.bat` opens a Home-only operations console. The
Job Descriptions, SharePoint, Resumes, Candidates, and Settings tabs remain
visible but disabled. Home provides a dynamic batch-size selector, one-click
live run, and daily in-app schedule editor. In-app schedules require the app to
remain open; the optional startup checkbox launches it after Windows sign-in.

Both use the same engine and `config.yaml`; the app is the operator UI, and the bot is the headless worker.

## 3B. Terminal Choice - Command Prompt / PowerShell vs VS Code

There is no functional difference. Every command in this runbook is a plain `python`/`.venv\Scripts\python.exe` invocation, so it works identically no matter what window you type it into:

- **Command Prompt or PowerShell** - already on every Windows machine, no install needed. Right-click the P2 folder in File Explorer and choose "Open in Terminal" (or `cd` there manually), then run the commands as written in this runbook.
- **VS Code** - not required, but a reasonable choice if the client also wants to browse/edit `config.yaml` or `.env` in a proper editor instead of Notepad. Open the P2 folder (`File > Open Folder`), then use its built-in terminal (`` Ctrl+` ``) - that terminal is still just PowerShell/cmd underneath, so the same commands apply unchanged.

Pick whichever the client is already comfortable with. `Launch.bat` and `RunDaily.bat` both work by double-click from File Explorer too, with no terminal at all.

## 4. First-Time Setup

Open PowerShell in the P2 folder:

```powershell
cd C:\HiringAgent\HiringAgent_App_P2
```

Run:

```powershell
.\Launch.bat
```

If `Launch.bat` is blocked:

```powershell
Unblock-File .\Launch.bat
Unblock-File .\Start.bat
Unblock-File .\RunDaily.bat
.\Launch.bat
```

Expected result:

- `.venv\` is created.
- Python dependencies install.
- GUI opens, or setup completes without dependency errors.

## 5. Optional Local AI Setup

P2 can run without Ollama, but extraction and scoring are better with it.

Install Ollama, then run:

```powershell
ollama pull llama3.2
ollama serve
```

Verify:

```powershell
Invoke-WebRequest -UseBasicParsing http://localhost:11434/api/tags
```

Expected result: HTTP status `200`.

## 6. Preflight Checks

Run from:

```powershell
cd C:\HiringAgent\HiringAgent_App_P2
```

Check local setup:

```powershell
.venv\Scripts\python.exe bot.py --doctor
```

Check SharePoint and Graph credentials:

```powershell
.venv\Scripts\python.exe bot.py --test-sharepoint
```

Expected:

- SharePoint connection passes.
- Workbook columns can be read.
- No missing `.env` fields.
- Ollama either shows reachable or clearly says it will fall back.

## 7. Verify Mail Is Enabled

Open `config.yaml` and confirm:

```yaml
geo_reject_email: true
```

Confirm `.env` contains the sender mailbox setting:

```text
SENDER_MAILBOX=...
```

P2 sends only scoring-stage emails:

- Non-USA decline emails from Rejected rows.
- Missing-info requests for scored rows missing phone, location, skills, or education.
- Admin error alerts for rows that fail processing 3 times, if configured.

P1 still owns intake replies such as acknowledgment, no-resume, unsupported format, scam/spam/noise handling.

### Silencing P2 completely (historical replay / test runs)

To reprocess old mail without contacting anyone, set in `.env`:

```text
HIRING_SUPPRESS_EMAILS=true
```

No P2 email of any kind then leaves the mailbox, while scoring, row patches, resume
renames/moves, Rejected-sheet maintenance, and the client workbook export all run normally.
Affected rows are stamped `TEST-MODE (suppressed) Sent <timestamp>` instead of
`Sent <timestamp>`, so replayed rows stay distinguishable from real contacts.

`HIRING_GEO_REJECT_EMAIL=false` and `HIRING_ERROR_EMAIL=false` are **not** enough on their
own: they cover the decline and the admin alert, but the missing-info nudge has no flag of
its own and would still be sent. Use `HIRING_SUPPRESS_EMAILS`.

P1 has its own separate switch (`flow/flow_config.json` -> `email.send_applicant_emails`,
then rebuild and re-import the zip). Silencing one does not silence the other.

**Set it back to `false` when the replay is done**, or genuine new applicants get no email at all.

## 8. Safe Test Before Live Run

Run a read-only validation against one known row:

```powershell
.venv\Scripts\python.exe bot.py --validate-row APP-REPLACE-WITH-REAL-ID
```

This should not write to SharePoint or send mail.

Optional local regression test:

```powershell
.venv\Scripts\python.exe test_p2.py
```

Expected current result:

```text
P2 RESULT: 501 passed, 0 failed
```

## 8A. Manual Setup On macOS Or Linux

From the extracted P2 directory:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
cp .env.example .env
```

Fill `.env` through the client's approved secret channel, then validate:

```bash
python bot.py --doctor
python bot.py --test-sharepoint
python test_p2.py
python bot.py --score-sharepoint --dry-run --no-ai
```

For OCR on image-only PDFs, install the Tesseract binary with the operating system's
package manager and install the desktop extras:

```bash
python -m pip install -r requirements-desktop.txt
```

Do not copy a Windows `.venv` to macOS/Linux (or between client machines). Always create
the virtual environment locally from the requirements file.

## 9. Live Run

Run once:

```powershell
.venv\Scripts\python.exe bot.py --score-sharepoint
```

What to watch in `P2_Logs\YYYY\Month\P2.log`:

- `Step 1/4 Checking connectivity`
- `OK Internet/Graph reachable`
- `OK Connected to ...sharepoint.com`
- `Queue has N candidates`
- Candidate-by-candidate `SCORED` or `REJECTED`
- `Declines : ... sent | 0 failed`
- `Info reqs : ... sent | 0 failed`
- `Errors : 0`
- `Finished.`

## 9A. Batch Size And Running The Full Queue

Every `--score-sharepoint` run caps how many candidates it scores in that one run - this is `scoring.batch_limit` in `config.yaml` (default **25**), overridable with the `HIRING_SCORING_BATCH_LIMIT` environment variable. It exists to keep a single run fast and to avoid cloud execution timeouts (Azure Functions has a hard 10-minute limit); it is not a special CLI flag, just an ordinary config value.

If the queue has more candidates than the limit, the log says so and only scores the first batch:

```text
Queue has 47 candidates. Processing first 25 (batch limit).
```

The other 22 stay queued (`Status = "New Email Received"`) untouched, ready for the next run.

**To score in smaller batches (e.g. 20 rows at a time):**

```powershell
# one-off override for a single run, PowerShell:
$env:HIRING_SCORING_BATCH_LIMIT="20"
.venv\Scripts\python.exe bot.py --score-sharepoint
```

or set it permanently by adding a line to `.env`:

```text
HIRING_SCORING_BATCH_LIMIT=20
```

or edit `config.yaml` directly:

```yaml
scoring:
  batch_limit: 20
```

`.env` wins if both are set.

**To process everything currently queued in one go ("full run"):** there is no dedicated "no limit" flag - pick whichever of these fits the situation:

- Run `bot.py --score-sharepoint` repeatedly until the log no longer shows the "Processing first N (batch limit)" line (i.e. the queue is fully drained) - safe and simple for a small backlog.
- Temporarily raise the limit above the queue size for one run (e.g. `$env:HIRING_SCORING_BATCH_LIMIT="500"`) - fine for a local desktop run with no timeout constraint; avoid this on Azure Functions if the batch is large enough to risk the 10-minute cap.
- For a manual CMD run, choose the limit directly for that invocation:
  `.venv\Scripts\python.exe bot.py --score-sharepoint --batch-size 50`.
- Use continuous polling (`--watch --interval 300`, see below) - it naturally drains the queue a batch at a time across successive polls without any manual re-running.

## 10. Continue Automatically

Continuous polling every 5 minutes:

```powershell
.venv\Scripts\python.exe bot.py --score-sharepoint --watch --interval 300
```

For a client laptop, Windows Task Scheduler is usually better than leaving a terminal open.

## 11. Windows Task Scheduler

Create a task:

- Program/script:
  `C:\HiringAgent\HiringAgent_App_P2\RunDaily.bat`
- Start in:
  `C:\HiringAgent\HiringAgent_App_P2`
- Run whether user is logged on or not.
- Run with highest privileges if the client environment requires it.
- Configure power settings so the laptop does not sleep while the task should run.

Important:

- Locking the laptop is okay.
- Sleep or hibernate can pause or stop P2.
- To intentionally stop all P2 processing while repairing SharePoint data, set `HIRING_P2_DISABLED=true` in `.env` and disable the Windows Task Scheduler task.
- For Azure Functions, the timer also requires `HIRING_ENABLE_P2_TIMER=true`; without that flag the timer no-ops even if the function is deployed.

To change the schedule later:

1. Open **Task Scheduler**.
2. Find the P2 task.
3. Open **Properties > Triggers**.
4. Edit the trigger time/frequency.
5. Confirm **Actions** still points to:
   `C:\HiringAgent\HiringAgent_App_P2\RunDaily.bat`
6. Confirm **Start in** is:
   `C:\HiringAgent\HiringAgent_App_P2`

Do not edit `RunDaily.bat` just to change time. Use the task trigger for schedule changes.

## 11A. Columns P2 Uses

Main `CandidateList` columns, in order:

1. Application ID
2. Received Date
3. Last Updated Date
4. Category
5. Resume Link  *(moved here from position 27 at client request, 2026-07-24)*
6. Full Name
7. Email
8. Phone
9. Location
10. Country
11. Current Skills
12. Education
13. Looking For Role
14. Suggested Role 1
15. Suggested Role 2
16. Suggested Role 3
17. Mail Subject
18. Mail Body
19. Status
20. Has Resume
21. Original Filename
22. Application Updates
23. Retry Count
24. Portfolio 1
25. Portfolio 2
26. Portfolio 3
27. Resume URL
28. Resume Folder Path
29. Mail Sent
30. Info Request Sent

Rejected sheet columns are the same except:

- `Mail Sent` is removed.
- `Info Request Sent` is removed.
- `Decline Sent` is added as the final column.

P1 owns intake fields such as Application ID, Received Date, Last Updated Date, Email, Mail Subject, Mail Body, Has Resume, Original Filename, Application Updates, and Mail Sent. P2 preserves those and fills scoring fields, status, retry count, resume links/paths, and P2 mail markers.

`Resume Link` should display the P1-owned `Original Filename` while pointing at the P2-owned `Resume URL`. P2 may rename the stored SharePoint file, but it should not overwrite `Original Filename` with the canonical storage filename.

## 11B. Folder Creation And Resume Naming

P1 creates the base SharePoint resume area:

```text
Candidate_Resumes/<Year>/<Month>/
```

P1 saves new resumes into the received-date month folder. P2 reconstructs that folder from `Received Date`, downloads the resume for scoring, then updates the same SharePoint row.

When P2 scores a candidate, it renames the resume to:

```text
<FirstNameLastName>_<Category>_<AppIdTail>.<ext>
```

Example:

```text
JaneSmith_DataAnalytics_A3F9.pdf
```

Rules:

- `FirstNameLastName` comes from the extracted full name. Middle names are dropped when a clear first/last name exists.
- `Category` is alphanumeric PascalCase, for example `AI/ML/CV (SIN2)` becomes `AIMLCVSIN2`.
- `AppIdTail` is the final hyphen segment of the Application ID.
- The extension stays the same as the original file.
- P2 tries current names first, then older P1 legacy names, so old files can still be found and renamed forward.
- `Resume Link` is repaired only when explicitly requested by maintenance tooling; normal P2 runs do not keep rewriting the calculated-column formula.

Rejected candidates are moved into one year-level rejected folder:

```text
Candidate_Resumes/<Year>/Rejected/
```

Example:

```text
Candidate_Resumes/2026/Rejected/JaneSmith_DataAnalytics_A3F9.pdf
```

Local-only runs may create:

- `P2_Output\<Year>\HiringAgent_P1_CandidateList.xlsx`
- `P2_Input\<Year>\<Month>\` only when local copy saving is enabled
- `P2_Logs\YYYY\Month\P2.log`

Every live SharePoint scoring run also creates a client-facing workbook:

- `P2_Final_Results\Candidate_List_Results_<timestamp>.xlsx`
- `P2_Final_Results\Candidate_List_Results.xlsx`
- SharePoint root `Shared Documents/` folder: `Candidate_List_Results.xlsx`

That workbook contains one `Candidates` sheet only. It excludes the Rejected sheet and includes only scored candidates from the main sheet.

## 11C. Practical Pros And Cons

Desktop app on client laptop:

- Pros: easy for a non-developer to run, visible health/status, simple JD refresh, browse candidates.
- Cons: laptop must be awake, Windows/permissions can interrupt runs, `.env` must be protected locally.

Headless bot on client laptop:

- Pros: easiest scheduled setup with Task Scheduler, same code as app, logs every run.
- Cons: less visual, schedule depends on laptop power/sleep settings.

GitHub Actions / cloud:

- Pros: no laptop dependency, schedule keeps running, secrets stay in platform secret storage.
- Cons: requires repo/tenant setup, schedule changes happen in workflow/cloud settings, debugging is through logs.

Offline fallback without Ollama:

- Pros: deterministic and no AI install/API key required.
- Cons: less flexible for messy resumes than Ollama-assisted extraction/recheck.

## 12. After-Run Validation

Run read-only audits:

```powershell
.venv\Scripts\python.exe audit_sharepoint_results_health.py
.venv\Scripts\python.exe audit_all_resume_files.py
```

Manual SharePoint checks:

- Main sheet has no `New Email Received` rows left, unless new mail arrived during the run.
- Main scored rows have `Status = Scored`.
- Main unresolved geo rows have `Status = Needs Review - Location Confirmation`; they are not part of the client results export.
- Rejected sheet has non-USA rows with `Status = Rejected - Non-USA Location`.
- `Rejected - Location Not Confirmed` is historical only; current P2 keeps unresolved candidates on Main for manual confirmation.
- Processing-error rows have `Status = Rejected - Processing Error`.
- `Decline Sent` is stamped for non-USA rejected rows.
- `Info Request Sent` is stamped or set to `Nothing missing` / `No valid email` for scored rows.
- No duplicate Application IDs across Main + Rejected.

## 13. Common Problems

### Another Run Holds The Lock

If no P2 is actually running but `.agent.lock` remains:

```powershell
Get-Process python -ErrorAction SilentlyContinue
Get-Content .agent.lock
```

Only remove the lock if no matching P2 process is active:

```powershell
Remove-Item .agent.lock
```

### Ollama Not Reachable

P2 will continue with keyword fallback. To restore AI:

```powershell
ollama serve
ollama pull llama3.2
```

Then rerun or let the next scheduled run use Ollama.

### Graph 503 Or Throttling

P2 retries Graph `429`, `503`, and `504` automatically. If it continues failing:

- Wait a few minutes.
- Confirm internet.
- Confirm SharePoint is reachable in browser.
- Rerun `bot.py --test-sharepoint`.

### Mail Not Sending

Check:

- `.env` sender mailbox.
- `Mail.Send` application permission admin consent.
- `geo_reject_email: true`.
- `P2_Logs` for `Declines` or `Info reqs` failed counts.

### Resume Unreadable

P2 retries unreadable/no-content rows up to the configured retry max. After 3 failures, it moves the row to:

```text
Rejected - Processing Error
```

This is not a non-USA rejection. P2 does not send the normal candidate decline for processing-error rows.

## 14. Rollback

If the new install has issues:

1. Stop P2.
2. Keep `.env` safe.
3. Rename the current folder:
   ```powershell
   Rename-Item C:\HiringAgent\HiringAgent_App_P2 C:\HiringAgent\HiringAgent_App_P2_broken
   ```
4. Extract the previous known-good backup zip.
5. Copy `.env` back into the restored folder.
6. Run:
   ```powershell
   .venv\Scripts\python.exe bot.py --test-sharepoint
   ```

## 15. Deployment Sign-Off Checklist

- [ ] P2 folder extracted on client laptop.
- [ ] `.env` present and protected.
- [ ] `Launch.bat` completed.
- [ ] `bot.py --doctor` checked.
- [ ] `bot.py --test-sharepoint` passed.
- [ ] Ollama reachable or keyword fallback accepted.
- [ ] `geo_reject_email: true` confirmed.
- [ ] One live `bot.py --score-sharepoint` run completed.
- [ ] Log shows `Errors : 0`.
- [ ] Log shows decline/info mail pass with `0 failed`.
- [ ] SharePoint Main has no stale pending rows.
- [ ] Rejected sheet has stamped `Decline Sent` values.
- [ ] No duplicate Application IDs across Main + Rejected.
- [ ] Task Scheduler configured if ongoing laptop automation is needed.
