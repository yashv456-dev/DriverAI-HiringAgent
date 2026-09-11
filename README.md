# DriverAI Hiring Agent

Two-phase hiring pipeline.

**P1** watches the `apply@` mailbox, screens what arrives, saves resumes to SharePoint
and writes one row per applicant. **P2** reads those rows, extracts and scores each
resume locally, and publishes a client-facing results workbook.

They are independent. P1 runs in Microsoft's cloud whether or not P2 is running, and P2
can be re-run over the queue at any time without touching the mailbox.

---

## How it works

### P1 — intake

A scheduled flow polls the Inbox once a minute and takes one message per run, so writes
stay sequential. Each message passes four screens in order:

1. **Sender**, 43 terms. Bots, marketing platforms, and the self-loop guard so the flow never answers itself.
2. **Subject**, 27 terms. Invoices, password resets, out-of-office.
3. **Content**, 123 terms over subject, body and attachment names. Scam, abuse, malware, non-English scam, vendor pitches.
4. **Link shorteners**, 12 domains, but only when no real resume is attached.

Blocked mail goes to Junk and is left unread. Mail that passes is routed:

| Situation | Result |
|---|---|
| New application with a resume | Row created, resume saved, acknowledgment sent |
| Same person again within 90 days | Existing row's counter incremented, no second row |
| Reply quoting a reference, with a new resume | That row updated, resume replaced |
| Reply quoting a reference, no resume | Noted on the row |
| Application keywords but no attachment | Reply asking for the resume, no row |
| Attachment that is not PDF or Word | Reply asking for the right format, no row |
| Clean mail that is not an application | Archived, admin alerted, no row |

Handled mail is marked read and moved to Archive. Failures move unread so they stay
visibly distinct, and only after an alert was actually attempted.

Resumes land in `Candidate_Resumes/<year>/<month>/` named `<original>_<reference>.pdf`.
A repeat submission overwrites that person's file rather than adding a second one.

### P2 — scoring

P2 reads rows marked new or needs-review. For each it downloads the resume, extracts the
text, extracts candidate fields, runs an independent recheck, scores against the cached
job descriptions, applies the geography rule, and commits the result.

Text extraction is a ladder and OCR is its last rung: multi-column reading order first,
then the raw text layer, then a second parser. Only when all three come back empty, which
means a scan or an image-only PDF, does OCR run.

P1 fills 12 of the 33 columns, the identity and mail envelope. P2 fills the rest.

---

## Getting started

**Docker is the recommended path**, and the only one that behaves identically on macOS and
Windows. The native path exists because the Windows launchers and the desktop GUI predate it.

| | Docker | Native |
|---|---|---|
| macOS | yes | CLI only, no `.bat` launchers |
| Windows | yes | yes, including the GUI |
| Gets you | `bot.py`, the whole scorer | plus the desktop window |

---

## P2 with Docker (macOS and Windows)

### 1. Prerequisites

- **Docker Desktop**
- **[Ollama](https://ollama.com) on your host machine**, not in the container. The model is
  multi-gigabyte and on a Mac needs Metal acceleration a Linux container cannot reach.

```bash
ollama pull qwen3:1.7b
```

### 2. Configure

```bash
cd DriverAI_HiringAgent/HiringAgent_P2
cp .env.example .env
cp app_settings.example.json app_settings.json
```

Fill `.env` with `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`, the `SHAREPOINT_*` values and
`SENDER_MAILBOX`. Ask whoever owns the Azure app registration.

**`.env` is gitignored and must never be committed.** It holds a secret with write access
to the live SharePoint site.

`app_settings.json` holds run defaults. Copy it even if you only use the CLI: the test
suite reads it directly and fails without it.

### 3. Build and check

```bash
docker compose build
docker compose run --rm p2 --doctor            # environment and dependencies
docker compose run --rm p2 --test-sharepoint   # credentials reach Graph
```

> The container definition was written on a Windows machine without Docker installed,
> so the image has not been built and run end to end yet. The Python it wraps is the
> same `bot.py` that runs natively, and nothing in the scorer is Windows-bound — the
> worker lock already branches to `fcntl` off Windows. If the build or first run trips
> on something, raise it and it gets fixed rather than worked around.


### 4. Run

```bash
# reads and logs, writes nothing — always start here
docker compose run --rm p2 --score-sharepoint --dry-run

# score the live queue
docker compose run --rm p2 --score-sharepoint --batch-size 500

# publish the client results workbook
docker compose run --rm p2 --export-results
```

Everything after `p2` is passed straight to `bot.py`. The database, results, logs and
cached resumes are mounted from your working copy, so they survive between runs.

---

## P2 natively

Use this if you want the desktop GUI, which the container does not carry.

```bash
cd DriverAI_HiringAgent/HiringAgent_P2
python -m venv .venv

.venv/Scripts/activate          # Windows
source .venv/bin/activate       # macOS

pip install -r requirements-desktop.txt
```

Configure exactly as in step 2 above, then:

```bash
python bot.py --doctor
python bot.py --score-sharepoint --dry-run
python bot.py --score-sharepoint
```

On Windows the batch files wrap the same commands: `Launch.bat` sets up the environment and
opens the GUI, `RunP2.bat` scores the queue, `RunDaily.bat` is the scheduled form.
**There is no macOS equivalent**; use `bot.py` or Docker.

### Which requirements file

| File | Use |
|---|---|
| `requirements-docker.txt` | Container. CLI and OCR, no GUI. |
| `requirements-desktop.txt` | Native local runs. Adds the GUI. |
| `requirements.txt` | Azure Functions only. **No OCR** — scanned PDFs will not be read. |

### Tests

```bash
python test_p2.py                    # main suite
python test_store.py                 # storage layer
python test_local_pipeline.py        # local-first pipeline
python test_resume_availability.py   # resume fetch paths
```

---

## P1 — the mail flow

P1 is **not** a program you run. It is a Power Automate flow. The Python here builds and
validates the package you import. Any machine with Python 3.11+ can build it, macOS included.

### 1. Configure

Everything lives in `HiringAgent_P1/flow/flow_config.json` — mailbox, admin address, spam
lists, caps, timezone, SharePoint targets. **Never hand-edit `definition.json`**, it is
generated.

### 2. Build

```bash
cd DriverAI_HiringAgent/HiringAgent_P1/flow
python build_zip.py
```

Writes `DriverAI-Hiring-AutoReply-apply.zip` and syncs `definition.json`. The banner reports
what it built: poll interval, caps, spam-term counts, whether applicant mail is on, and
nesting depth.

### 3. Test

```bash
cd ..
python test_p1.py
```

Structural validation of the built package plus a routing simulation over ~55 real
scenarios. Expect `797 passed, 0 failed`. **Always rebuild and re-run this before
importing** — a committed package can silently disagree with the config.

### 4. Import

Power Automate → My flows → Import → Import Package (Legacy) → upload the zip, then
authorize Office 365 Outlook, SharePoint and Excel Online.

It imports as an **update** to an existing flow ID, replacing that flow rather than
creating a second copy.

### Settings worth knowing

| Setting | Meaning |
|---|---|
| `email.send_applicant_emails` | Master switch for all six applicant replies |
| `email.send_admin_failure_alerts` | Keeps the 14 admin alerts live |
| `trigger.interval_min` | Poll frequency |
| `business_rules.*` | Duplicate window and per-applicant reply caps |
| `intake_sidecars.enabled` | Off. Read the config comment before re-enabling. |

---

## Things that will catch you out

**P1 has no regex.** Power Automate offers only case-insensitive substring matching. Every
spam term must be safe as a bare substring — `marketing@` matches inside a real applicant's
address, and once did. Prefer multi-word phrases.

**Applicant mail is a switch, not a code change.** When off, the six replies compile to
no-op steps and rows are stamped as suppressed rather than falsely recording a send.

**P1's duplicate check reads the master workbook.** If that workbook is empty, every arrival
looks like a new person and gets a new row. The behaviour is correct, it just has nothing
to match against.

**A fresh clone has no database.** The first run starts empty. That is expected.

**`requirements.txt` has no OCR.** It is the serverless file. A scanned PDF comes back with
no text rather than an error.

**P1 sits at the platform nesting ceiling**, depth 8 of 8. Adding a condition or loop in the
update or follow-up branches needs something flattened first.

---

## Documentation

| Document | Covers |
|---|---|
| [`VERSION_STACK.md`](DriverAI_HiringAgent/VERSION_STACK.md) | Verified runtime, OCR, LLM, SQL and regex versions |
| [P1 Technical Design](DriverAI_HiringAgent/HiringAgent_P1/P1_Technical_Design_Document_Consolidated.md) | Full P1 specification |
| [P1 Runbook](DriverAI_HiringAgent/HiringAgent_P1/docs/P1_RUNBOOK.md) | Day-to-day P1 operations |
| [P2 Manual Run Steps](DriverAI_HiringAgent/HiringAgent_P2/docs/P2_MANUAL_RUN_STEPS.md) | P2 operator guide |
| [P2 Deployment Runbook](DriverAI_HiringAgent/HiringAgent_P2/docs/P2_CLIENT_DEPLOYMENT_RUNBOOK.md) | Installing P2 on a new machine |
| [Cloud Deploy](DriverAI_HiringAgent/HiringAgent_P2/docs/CLOUD_DEPLOY.md) | Serverless options. Its Docker section predates the Dockerfile here. |
| [Architecture Review](DriverAI_HiringAgent/HiringAgent_P2/docs/ARCHITECTURE_REVIEW_20260908.md) | Known gaps and priorities |

---

## What is not in this repository

Credentials, the candidate database, logs, result workbooks and cached resumes are all
gitignored. They are secrets or personal data. Use `.env.example` and
`app_settings.example.json` as your templates.
