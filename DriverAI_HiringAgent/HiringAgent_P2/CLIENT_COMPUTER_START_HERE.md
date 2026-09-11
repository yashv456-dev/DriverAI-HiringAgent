# Client computer handoff — P1 and P2, current state

Prepared September 5, 2026; status corrections added September 7. Start with the [current architecture](../CURRENT_ARCHITECTURE.md). Installation steps remain useful; source-computer test counts, defects, and tenant observations below are historical evidence, not a fresh validation. SQLite integration is partial and is not the default handoff path.

## 1. What was delivered

The owner requested an exact folder transfer, with notes added and defects investigated/fixed on the client computer if needed. No runtime code, configuration, mail switch, schedule, or existing P1 import package was changed for this handoff. No new ZIP/archive was created. The additional portable model definition is a setup resource; it does not switch or modify the currently installed model.

The five-candidate dry run completed, but it found real extraction defects. This handoff is installation guidance, not a production-quality sign-off. Preserve the evidence while setting up; do not assume a successful run means every field is accurate.

Keep this layout (the doubled P1 directory is intentional):

```text
DriverAI_HiringAgent/
  HiringAgent_P1/
    HiringAgent_P1/
      CLIENT_COMPUTER_START_HERE.md
      flow/DriverAI-Hiring-AutoReply-apply.zip
      flow/flow_config.json
      docs/P1_RUNBOOK.md
      test_p1.py
  HiringAgent_P2/
    CLIENT_COMPUTER_START_HERE.md
    Launch.bat
    Start.bat
    RunDaily.bat
    bot.py
    config.yaml
    app_settings.json
    .env
    jd_sources.json
    jd_roles_cache.json
    docs/client_handoff/qwen3-1.7b-p2.Modelfile
```

Transfer these folders directly, without compressing or rebuilding them. P1's ZIP already exists and is the Power Automate import artifact, not a newly compressed handoff. The copied folders contain credentials and candidate data; use the client's restricted transfer/storage location. Do not paste `.env`, resumes or evidence files into a public issue or repository.

## 2. Record the current behavior

| Setting | Delivered state / effect |
|---|---|
| P1 runtime | Microsoft cloud / Power Automate; moving laptops does not move or stop it |
| P1 applicant email | `flow/flow_config.json`: `email.send_applicant_emails: false` (silent intake since 2026-09-05; `true` restores the six lifecycle emails after a rebuild + re-import) |
| P1 admin failure alerts | `email.send_admin_failure_alerts: true` |
| P2 model | `qwen3-1.7b-p2:latest`, extraction and scoring enabled |
| P2 timeouts | Extraction 60 seconds; scoring 150 seconds; several calls per candidate |
| P2 applicant email | Suppressed in the tested environment; YAML `test_mode.suppress_emails: true` |
| P2 admin alerts | Enabled independently; suppression does not silence these |
| P2 USA filter | Enabled; unknown location normally stays on Main for review |
| P2 CLI batch default | 500; explicitly use `--batch-size 5` for the first client run |
| P2 GUI batch default | 25 in `app_settings.json`; check the Home selector |
| P2 GUI scheduling | Auto-start, scheduling and launch-on-startup disabled in delivered settings |

Process environment variables override `.env`/YAML values. Verify effective settings on the client. The old `.env.example` sets `HIRING_SUPPRESS_EMAILS=false`; copying it unchanged can enable applicant emails even with YAML suppression true. Its claim that suppression also stops admin alerts is stale. Older runbooks also contain stale llama3.2 and mail-state instructions; use this note for the current settings.

## 3. Choose the client location and preserve the old worker state

Example PowerShell paths used below; change only the first path if installed elsewhere:

```powershell
$handoffRoot = 'C:\HiringAgent\DriverAI_HiringAgent'
$p2Path = Join-Path $handoffRoot 'HiringAgent_P2'
$p1Path = Join-Path $handoffRoot 'HiringAgent_P1\HiringAgent_P1'
Set-Location -LiteralPath $p2Path
```

Record Windows version, installed RAM, GPU/VRAM, free space, and the Windows account that will run P2. Source-machine timing is not a client performance guarantee: the development machine had 32 GB RAM and an RTX 3070 laptop GPU with 8 GB VRAM.

Before the first client LIVE write run, stop P2 on the old computer and disable its scheduler/watch loop. Use one P2 worker for this workbook. Do not start both the GUI schedule and Task Scheduler. A laptop move alone does not require re-importing P1.

## 4. Install Python and recreate the local environment

Run in PowerShell on the CLIENT computer:

```powershell
winget install -e --id Python.Python.3.12 --accept-source-agreements --accept-package-agreements
winget install -e --id Ollama.Ollama --accept-source-agreements --accept-package-agreements
```

Open a new PowerShell window afterward, set the paths from step 3 again, and confirm `py -3.12 --version` and `ollama --version` work. Python 3.14 passed tests on the source computer; 3.12 is the existing client runbook's default.

Recreate `.venv` on the client even if one arrived with the folder. Preserve the transferred copy by renaming it; if `.venv.transferred` already exists, choose an unused backup name before running the rename. Run this only inside the client's P2 folder:

```powershell
Set-Location -LiteralPath $p2Path
if (Test-Path -LiteralPath '.\.venv') {
    Rename-Item -LiteralPath '.\.venv' -NewName '.venv.transferred'
}
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements-desktop.txt
```

Inspect errors before continuing. These steps change the client environment, not the transferred application logic.

## 5. Recreate the EXACT custom Qwen setup before launching P2

The model weights normally live outside the project folder under the Ollama user's profile. Copying P1/P2 does not copy them. `qwen3-1.7b-p2:latest` is a custom local name; do not try to install it with `ollama pull qwen3-1.7b-p2`.

Start the Ollama application. If its server is not running, run `ollama serve` in another terminal and leave that terminal open. Then, in the P2 folder:

```powershell
ollama pull qwen3:1.7b
ollama create qwen3-1.7b-p2:latest -f .\docs\client_handoff\qwen3-1.7b-p2.Modelfile
ollama list
ollama show qwen3-1.7b-p2:latest --parameters
```

The Modelfile was exported on the source computer. The later [September 6 benchmark](docs/benchmarks/stage2a/REPORT.md) observed a different live template, so this file alone does not prove exact reproduction of that benchmark environment. Only `FROM` changed from the source computer's absolute blob path to `qwen3:1.7b`; template, parameters and license were preserved. It uses the custom non-thinking prompt template, temperature 0, context 8192, top_k 20, top_p 0.95, repeat_penalty 1 and the original stop tokens. The source base and custom model were verified to use the same weights blob.

Source identifiers: base `8f68893c685c`; custom `338d99fd07ec`. Full weights SHA-256 and Modelfile hash are in `docs/client_handoff/MODEL_BASELINE.json`. A tag can later point to different weights. If the pulled base ID differs, investigate before calling it an exact reproduction; use the matching source weights as a separate folder/file transfer if necessary. Do not silently substitute another Qwen or llama model. Confirm both `config.yaml` and `app_settings.json` still select `qwen3-1.7b-p2:latest`.

Now finish launcher setup without opening the GUI or starting scoring:

```powershell
.\Launch.bat --setup-only
```

This can install missing OCR dependencies and Tesseract. Check failures: the launcher treats Ollama/OCR installation as best-effort and can still finish when they are unavailable. It otherwise attempts to pull the custom name when it cannot find it, which is why the explicit model creation comes first. For Task Scheduler, use the same Windows account that installed Ollama/model files, and ensure the Ollama server is running for that account.

## 6. Connect to the intended Microsoft 365 data

If this is the same tenant/site/mailbox, preserve the existing `.env` values. For a different client tenant, an admin must supply matching credentials and targets locally; an exact copy still points to the original tenant until those settings change. Never display the secret in a diagnostic transcript.

Verify `.env` entries locally: `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET` (secret VALUE, not secret ID), `SHAREPOINT_HOSTNAME`, `SHAREPOINT_SITE_PATH`, `SHAREPOINT_TABLE`, `SHAREPOINT_WORKBOOK`, `SHAREPOINT_RESUMES_FOLDER`, and `SENDER_MAILBOX`. Record the secret expiry. Use the [P1 operations runbook](../HiringAgent_P1/docs/P1_RUNBOOK.md) and `.env.example` for connection setup; do not change permissions simply because the laptop changed.

Keep `HIRING_SUPPRESS_EMAILS=true` for the initial client validation. Leave the delivered admin-alert behavior in place for normal live operation. If you intentionally need a completely silent P2 live-write test, `HIRING_ERROR_EMAIL=false` is a separate setting. Changing either is a client-side operational change; neither was changed for this handoff. Restart P2 after any `.env` change, and check Windows/process environment overrides too.

Also check `jd_sources.json`: it currently points to the existing DriverAI OneDrive/SharePoint `Staffing/PDs` folder. The older handoff tests used 100 cached roles; September 6 measurements/design refer to 106. Verify the actual cache with --list-roles rather than treating either historical count as fixed. Preserve `jd_roles_cache.json` for like-for-like testing; use `--refresh-jd` only when a deliberate JD refresh is needed. That command updates the local role cache.

```powershell
.\.venv\Scripts\python.exe bot.py --doctor
.\.venv\Scripts\python.exe bot.py --test-sharepoint
.\.venv\Scripts\python.exe bot.py --list-roles
.\.venv\Scripts\python.exe -c "from hiring_agent import config as c; print({'model':c.OLLAMA_MODEL,'ai_extraction':c.OLLAMA_ENABLED,'ai_scoring':c.OLLAMA_SCORING,'applicant_mail_suppressed':c.SUPPRESS_EMAILS,'admin_alerts':c.ERROR_EMAIL_ENABLED,'usa_filter':c.GEO_FILTER_USA_ONLY,'cli_batch':c.SCORING_BATCH_LIMIT})"
```

These checks must resolve the intended workbook, model and roles. A model health probe does not prove inference will meet the timeout on the client hardware.

## 7. P1: retain an existing flow, or import the already supplied package

If P1 is already running against the intended inbox/workbook, leave it running; a P2 computer migration does not require a P1 import. Verify the flow owner/connections and recent run history.

For a first deployment, follow `HiringAgent_P1/docs/P1_RUNBOOK.md`, installation steps 1–11. Use its tenant setup details, but the mail-state corrections in this note take precedence.

1. Verify the intended mailbox, SharePoint site, existing master workbook/table, resume folders, and connection account. Do not overwrite an existing populated workbook with a template. Create the template workbook only for a genuinely new empty deployment.
2. Import the EXISTING `flow/DriverAI-Hiring-AutoReply-apply.zip` using the package import workflow. Do not rebuild, recompress, or substitute an archive backup for this handoff.
3. First install: create one flow. Updating an existing flow: choose Update and verify the target flow identity. Do not leave two flows watching the same inbox. Confirm import actually changed the intended flow rather than assuming a success banner proves it.
4. Bind Outlook, SharePoint and Excel Online connections. Check mailbox, site, workbook and table references inside the imported flow, especially for a different tenant. Verify Send As rights for the sending account.
5. Save and enable the intended flow when its inbox contents are ready to be processed. The supplied local configuration suppresses applicant emails and enables admin alerts; verify the imported flow matches it. Its inbox poll is not restricted to unread messages (`fetch_only_unread: false`), so merely reading a test email does not stop processing.
6. On the client, submit a clearly labeled synthetic application from a controlled test mailbox and verify the saved resume, one Application ID/row, suppressed applicant-mail behavior, and mailbox disposition. This is a real P1 transaction and may send mail. No such message or flow activation was performed during this handoff.

Changing P1's local JSON alone does not change the cloud flow or existing package. If P1 behavior needs changing later, treat that as a separate rebuild/import task on the client, outside this unchanged transfer.

## 8. Run tests and preview P2 on the client

From the P2 directory:

```powershell
.\.venv\Scripts\python.exe test_p2.py
.\.venv\Scripts\python.exe ..\HiringAgent_P1\HiringAgent_P1\test_p1.py
.\.venv\Scripts\python.exe bot.py --score-sharepoint --dry-run --batch-size 5
```

Source baseline: P1 **715 passed / 0 failed**, P2 **1,393 passed / 0 failed**. Existing suites are mostly offline/structural; they do not exercise a deployed P1 transaction or guarantee field accuracy.

The normal dry-run command only examines the current queue. If it reports zero, that is not a five-candidate test. The source computer used a dedicated read-only five-candidate replay harness for that; it and its one-off companions were removed from this package in the 2026-09-06 cleanup and are not part of the handoff. To exercise a comparable batch here, stage controlled test rows or raise `--batch-size` on the normal queue dry run, and record the count actually tested rather than assuming five.

Expected evidence: zero write attempts; zero email sends; scorer source recorded; every intended eligible ID appears once in the local workbook with correct resume links. Dry-run proposed Scored statuses are not evidence of factual accuracy. Do not treat `--dry-run` as a universal safety switch on unrelated commands: for example `--export-results` uploads a workbook.

Repair and replay tools should only be run for their explicit purpose, not as an installation step.

## 9. Known defects to check and fix HERE if needed

Do not silently erase this list after tests pass. Inspect the resume/email evidence, reproduce the issue, make a focused client-side fix, and re-run the relevant test plus a dry run before changing live rows.

| Area | Observed defect / reproduction | Repair target and acceptance check |
|---|---|---|
| Location | Doreen `90BB`: contact header says San Francisco Bay Area; output became Pacific Stockton, CA from a university address. Saniya `987A`: University Boston, MA from an institution/work line. | Inspect `hiring_agent/extraction.py` and `hiring_agent/sharepoint_scoring.py`; residence-header evidence must beat school/employer locations. Preserve explicit client-confirmed values with provenance. Remote alone does not establish US residence. |
| Education | Nefthali `E2B9`: work dates Jan 2025–Present used for education whose end is May 2026. Saniya: master's paired with bachelor's May 2024 instead of master's Dec 2026. | Bind dates to the chosen degree/education entry and respect section boundaries; leave absent start dates Missing. |
| Forwarded contacts | Saniya's stored Email/Phone belong to the internal forwarder; P2 also proposed the forwarder's Process Automation Engineer title. | Distinguish applicant from forwarder/signature, ground phone/email/role in candidate evidence, and check P1 identity plus P2 write fields. Main P2 fields do not currently repair Email. |
| Other extraction | Doreen's explicit 4+ years became Missing; Qianqi's degree name was omitted; plausible unsupported AI phones can pass sanitization. | Verify values against source spans; formatting validity is not identity or source evidence. |
| Fallback / JSON | Historical reproduction: extraction/scoring fallbacks and wrong-shaped JSON were observed. Later transport/schema guards and require_ai=true changed this path; verify current failure handling rather than assuming the earlier behavior still applies. | Validate response schema, record failures and scorer source persistently. Do not label keyword percentages as AI scores. |
| Location-review queue | Directly editing Location with a suppressed clarification marker may not requeue the row; suppressed markers do not start age-out. | Verify status/timestamp/queue behavior after a manual correction; preserve corrections rather than blindly re-extracting them. |
| Publication | Export trusts Scored without fresh geo validation; upload failure can return a local path without a failure status. | Gate publication on evidence/status and verify the remote workbook IDs and links after upload. A local file path is not proof of upload success. |
| Rejection recovery | Failed resume move can still allow row transfer; send and marker update are not atomic. | Verify destination file/row before deleting source; surface retryable failures and reconcile send markers before retrying. |

The scripts and JSON artifacts that recorded these reproductions, and the review reports formerly in the shared `DriverAI_HiringAgent` parent, were removed in the 2026-09-06 cleanup. The table above is now the record: each row is a defect reproduced on the source computer that still needs a fix plus an acceptance check here - not a behavior that passed a test.

Source five-person replay took 622.2 seconds with no unhandled processing error; Doreen's first extraction timed out. All five scored using Qwen, but incorrect fields still passed as Scored. Source scores were Qianqi 90%, Sneha 95%, Doreen 85%, Nefthali 75%, Saniya 100%; these are observations, not target scores to force on the client.

The prior source audit also found published-result drift (18 currently Scored IDs absent, one review ID present) and four rejected resumes unresolved in expected folders. Re-audit current state before repair; these snapshots do not prove candidate loss or that files are absent everywhere.

## 10. Run live with the current mail behavior

After connectivity, model, source checks and the client dry run are reviewed, take a recoverable copy of the current master workbook and resume folders and stop the old P2 worker. Close Excel desktop sessions that may lock the workbook. Keep P2 applicant suppression true for the initial live write run; P1 replies remain enabled and P2 admin alerts remain independent.

```powershell
Set-Location -LiteralPath $p2Path
.\.venv\Scripts\python.exe bot.py --score-sharepoint --batch-size 5
```

This command DOES write live data: row updates, file rename/moves, routing/reconciliation and results export can occur. The batch limit caps queued scoring, not every maintenance/email/reconciliation pass across existing rows. Mail suppression is not a dry run. Unknown location normally remains Main/Needs Review; confirmed non-US routes to Rejected. Do not assert all Main rows are confirmed USA or auto-approve every Scored row while the documented defects remain.

For each touched Application ID, read back: candidate contact/location/degree dates, status, top roles, resume file opening from its URL, and presence in the correct Main/Rejected sheet. Compare the remote results workbook with current eligible Scored IDs, not just its modified timestamp. Confirm no unintended applicant mail and review any admin alert. For a rejected candidate, verify the destination file before treating the transfer as complete. For failure/review cases, verify the record still exists with an actionable state.

If the results export is stale after rows are checked, `bot.py --export-results` rebuilds and UPLOADS it; it is a write to the published workbook and does not fix bad fields. Verify the uploaded workbook again because upload failure handling is a known gap.

If the client later wants P2 applicant emails enabled, make that a deliberate client-side switch (`HIRING_SUPPRESS_EMAILS=false`) and restart/check effective settings. First inspect pending suppressed markers: they can become eligible for real sends on the next live pass, including historical rows. This handoff does not enable those emails or clear any markers.

## 11. Enable one scheduler and monitor the first cycles

Use either the GUI (`Start.bat`, Home schedule controls; app must remain open) OR Windows Task Scheduler calling the full client path to `RunDaily.bat`. Do not enable both, and do not retain an old-machine watcher. Configure a Windows task to avoid overlapping instances, use the P2 folder as its working directory, and run under the configured account with access to Ollama and `.env`. Ensure the machine stays on/awake for the scheduled interval.

`RunDaily.bat` uses the CLI default batch (currently 500), not the GUI's 25. Set a deliberate client-side batch limit if required before scheduling. Review `P2_Logs`, including `task_console.log`, the queue/review states, remote export membership, and admin messages after the first scheduled cycles. The script's comment promising no duplicates is stronger than the current implementation guarantees.

If extraction or scoring times out, inspect whether source is offline/keyword before changing hardware/model/timeouts. More time alone does not fix field provenance. If a run fails, stop repeated automatic runs while checking the affected IDs and remote state; do not delete/requeue entire sheets or replay the archive as a generic recovery step.

## 12. Handoff completion record

Record client install path, Windows account/specs, Python/Ollama versions, model/base IDs, workbook/site, P1 flow identity, effective mail switches, queue count, test outputs, five checked IDs, remote export verification, and chosen schedule. Record which defects remain or were fixed, with evidence. Never record secret values.

`docs/client_handoff/SOURCE_SHA256_BEFORE_HANDOFF.json` records the existing runtime/config/import files before these notes were added. It can distinguish subsequent client-side fixes from the delivered source. No deployment, live run, mail send, existing-file edit, or compression was performed while preparing this note.
