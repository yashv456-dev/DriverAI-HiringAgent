# DriverAI Hiring Agent

Two-phase hiring pipeline. **P1** captures and screens inbound applications from a
shared mailbox and writes a queue. **P2** reads that queue, extracts and scores each
resume locally, and publishes a client-facing results workbook.

```
apply@ mailbox ──► P1 (Power Automate) ──► SharePoint resumes + Excel queue
                                                      │
                                                      ▼
                                          P2 (local Python) ──► scored results
```

| | P1 | P2 |
|---|---|---|
| Runs on | Power Automate cloud | Your machine |
| Language | Flow definition, built by Python | Python 3.11+ |
| Owns | Screening, intake, applicant mail | Extraction, scoring, publishing |
| Data | Excel table over Microsoft Graph | SQLite, publishes to Excel |

---

## Prerequisites

Both phases talk to the same Microsoft 365 tenant.

- **Python 3.11 or newer.** P2's virtualenv is built on 3.14; CI and the Azure
  Functions path pin 3.11. Avoid syntax newer than 3.11 if you touch shared code.
- **An Azure app registration** with `Sites.ReadWrite.All`, `Files.ReadWrite.All`
  and `Mail.Send`, and its tenant ID, client ID and client secret.
- **Power Automate access** to import the P1 flow (P1 only).
- **[Ollama](https://ollama.com)** with a local model (P2 only, for AI scoring).
- **Tesseract** is optional; P2's primary OCR engine needs no external binary.

---

## P2 — resume scoring (start here)

P2 is self-contained and the easier of the two to run.

### 1. Install

```bash
cd DriverAI_HiringAgent/HiringAgent_P2
python -m venv .venv
.venv/Scripts/activate          # Windows
# source .venv/bin/activate     # macOS / Linux
pip install -r requirements-desktop.txt
```

Two requirements files exist and they are not interchangeable:

| File | Use |
|---|---|
| `requirements-desktop.txt` | Local runs. Includes OCR and the GUI. |
| `requirements.txt` | Serverless only. No OCR, no GUI. |

### 2. Configure

```bash
cp .env.example .env
```

Fill in `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`, the `SHAREPOINT_*` values and
`SENDER_MAILBOX`. **`.env` is gitignored and must never be committed** — it holds a
secret with write access to the live SharePoint site.

### 3. Pull the model

```bash
ollama pull qwen3:1.7b
```

Set `HIRING_OLLAMA_MODEL` in `.env` if you use a different one. Scoring falls back to
an offline parser when Ollama is unreachable, so the pipeline does not break without it.

### 4. Check the setup before touching live data

```bash
.venv/Scripts/python bot.py --doctor            # environment + dependency check
.venv/Scripts/python bot.py --test-sharepoint   # credentials and Graph reachability
```

### 5. Run

```bash
# safest first run: reads and logs, writes nothing
.venv/Scripts/python bot.py --score-sharepoint --dry-run

# score the live queue
.venv/Scripts/python bot.py --score-sharepoint --batch-size 500

# score local files instead of the queue
.venv/Scripts/python bot.py --score-folder ./some_resumes

# export the client results workbook
.venv/Scripts/python bot.py --export-results
```

On Windows the batch files wrap the same thing: `Launch.bat` sets up the environment
and opens the GUI, `RunP2.bat` scores the queue, `RunDaily.bat` is the scheduled form.
`Launch.bat` rebuilds the virtualenv on first run or if it is unusable.

### 6. Tests

```bash
.venv/Scripts/python test_p2.py                    # main suite
.venv/Scripts/python test_store.py                 # storage layer
.venv/Scripts/python test_local_pipeline.py        # local-first pipeline
.venv/Scripts/python test_resume_availability.py   # resume fetch paths
```

---

## P1 — mail intake flow

P1 is **not** a program you run. It is a Power Automate flow. The Python here builds
and validates the package you import.

### 1. Configure

Everything lives in `HiringAgent_P1/flow/flow_config.json` — mailbox, admin address,
spam lists, caps, timezone, SharePoint targets. **Never hand-edit `definition.json`**;
it is generated.

### 2. Build

```bash
cd DriverAI_HiringAgent/HiringAgent_P1/flow
python build_zip.py
```

Writes `DriverAI-Hiring-AutoReply-apply.zip` and syncs `definition.json`. The banner
reports the posture it built: poll interval, caps, spam-term counts, whether applicant
mail is on, and nesting depth.

### 3. Test

```bash
cd ..
python test_p1.py
```

Structural validation of the built package plus a routing simulation over ~55
scenarios. Expect `797 passed, 0 failed`. **Always rebuild and re-run this before
importing** — a committed package can silently disagree with the config.

### 4. Import

Power Automate → My flows → Import → Import Package (Legacy) → upload the zip.
Authorize the three connections: Office 365 Outlook, SharePoint, Excel Online.

It imports as an **update** to an existing flow ID, replacing it rather than creating
a second copy.

### Key settings

| Setting | Meaning |
|---|---|
| `email.send_applicant_emails` | Master switch for all six applicant replies |
| `email.send_admin_failure_alerts` | Keeps the 14 admin alerts live |
| `trigger.interval_min` | Poll frequency |
| `business_rules.*` | Duplicate window and per-applicant reply caps |
| `intake_sidecars.enabled` | Off. See the config comment before re-enabling. |

---

## Operational notes

**P1 has no regex.** Power Automate offers only case-insensitive substring matching.
Every spam term must be safe as a bare substring — a term like `marketing@` will match
inside a real applicant's address. Prefer multi-word phrases.

**Applicant mail is a deliberate switch.** When off, the six replies compile to no-op
steps and rows are stamped as suppressed rather than falsely recording a send.

**P2 never deletes a resume.** Rejected candidates are moved, not removed.

**Nesting is at the platform ceiling.** P1 sits at depth 8 of 8. Adding a condition or
loop in the update or follow-up branches requires flattening something first.

---

## Documentation

| Document | Covers |
|---|---|
| [`VERSION_STACK.md`](DriverAI_HiringAgent/VERSION_STACK.md) | Verified runtime, OCR, LLM, SQL and regex versions |
| [P1 Technical Design](DriverAI_HiringAgent/HiringAgent_P1/P1_Technical_Design_Document_Consolidated.md) | Full P1 specification |
| [P1 Runbook](DriverAI_HiringAgent/HiringAgent_P1/docs/P1_RUNBOOK.md) | Day-to-day P1 operations |
| [P2 Manual Run Steps](DriverAI_HiringAgent/HiringAgent_P2/docs/P2_MANUAL_RUN_STEPS.md) | P2 operator guide |
| [P2 Deployment Runbook](DriverAI_HiringAgent/HiringAgent_P2/docs/P2_CLIENT_DEPLOYMENT_RUNBOOK.md) | Installing P2 on a new machine |
| [Architecture Review](DriverAI_HiringAgent/HiringAgent_P2/docs/ARCHITECTURE_REVIEW_20260908.md) | Known gaps and priorities |

---

## What is not in this repository

Credentials (`.env`), the candidate database, logs, result workbooks and cached
resumes are all gitignored. They are personal data or secrets. Use `.env.example` as
your template and expect an empty database on a fresh clone.
