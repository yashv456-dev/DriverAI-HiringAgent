# Current P1 / P2 architecture

Source review: September 8, 2026. This describes the files and local configuration in this workspace; it does not certify the imported Power Automate flow, active scheduler, installed model, or live SharePoint contents.

P2 is configured to run Ollama locally and select SQLite (`HIRING_STORAGE_BACKEND=sqlite` in the local `.env`). The intended boundary is local processing and state, followed by publication of the master and client results. That boundary is **not yet implemented end to end**: intake/maintenance still use Excel, database import is one-time, and regular scoring publishes only the separate client results workbook. See the [September 8 reliability review](HiringAgent_P2/docs/ARCHITECTURE_REVIEW_20260908.md) for reproduced failures and completion priorities.

## Current flow

P1's scheduled Power Automate flow reads the mailbox, saves resume files to SharePoint and writes rows to the Excel master. P2 reads the eligible queue rows from that master, downloads the matching resume files and extracts their text. It extracts candidate fields, runs an independent AI recheck and validates the result. Cached job descriptions feed role scoring and category assignment, which is followed by the geo and processing outcome. CandidateStore writes that outcome to whichever backend is configured, and the client results export is built from the stored rows.

The two store backends are alternatives rather than successive stages. Geo checks can exit early; maintenance, mail, and export passes also run outside the per-candidate path. Selecting SQLite does not move the Excel queue or all maintenance reads into the database.

## P1: capture and queue ownership

Sources: [flow configuration](HiringAgent_P1/flow/flow_config.json), [builder](HiringAgent_P1/flow/build_zip.py), and [generated definition](HiringAgent_P1/flow/definition.json).

- The configured recurrence is one minute, fetching one Inbox message per run. `fetch_only_unread` is false: opening a message does not remove it from the queue.
- Mail filters and attachment selection run before candidate intake. P1 handles new applications, duplicates within the 90-day window, reference replies, resume updates, and follow-ups with configured limits.
- P1 saves attachments under year/month folders and adds or patches the existing Excel candidate table. `Application ID` identifies applications; `Status = New Email Received` signals P2 work. P1 does not perform resume scoring.
- P1 has no explicit CV-delete action, but genuine updates can overwrite an existing upload path. It does not preserve immutable versions of every CV. P2's `move_resume` also deletes the source after uploading a destination copy; removal of direct stale-file cleanup is not a system-wide no-deletion guarantee.
- Terminal handling moves normal messages to Archive, spam to Junk Email, and failures to Archive with failure handling. Year/month separator insertion is disabled.
- Local applicant mail configuration is off; admin alerts are on. Changing local configuration requires rebuilding and importing the package to change the cloud flow.

The workbook remains P1's handoff. There is no implemented immutable intake-event publisher/consumer in the current P1/P2 path.

## P2: processing and outputs

Entry points are [bot.py](HiringAgent_P2/bot.py) and [app.py](HiringAgent_P2/app.py). [sharepoint_scoring.py](HiringAgent_P2/hiring_agent/sharepoint_scoring.py) orchestrates the online workflow; [sharepoint_client.py](HiringAgent_P2/sharepoint_client.py) handles Graph operations.

1. Check Graph/model availability and the existing workbook. Live runs check schema; scoring dry runs skip workbook/schema mutations.
2. Read new and unreadable-review rows, plus qualifying updated location-review rows. Exclude unchanged retry copies with verified completed counterparts before applying the batch limit.
3. Download the required attachments. Missing/inaccessible attachments defer the candidate without consuming the processing retry count; downloaded but unreadable content follows the review/error path.
4. Extract text and candidate fields with [extraction.py](HiringAgent_P2/hiring_agent/extraction.py), run the independent `ai_recheck_fields` pass, and validate/merge fields. [jd_sources.py](HiringAgent_P2/hiring_agent/jd_sources.py) supplies cached roles; [scoring.py](HiringAgent_P2/hiring_agent/scoring.py) and the scorer adapters produce role matches and category.
5. Apply the three-way geo policy: confirmed US can become `Scored`; unknown/conflicting residence stays in location review; confirmed non-US goes to Rejected. Other review/error states cover spam and processing failures. Separate age-out logic can close eligible old reviews.
6. Persist fields through [store.py](HiringAgent_P2/hiring_agent/store.py), maintain resume names/locations, and export scored main-sheet rows through `export_client_results`. An idle live pass can still repair fields, process mail markers, age reviews, or publish changes.

The regular client export excludes Rejected and uses its own column layout. The separate [exporter.py](HiringAgent_P2/hiring_agent/exporter.py) now also generates the client layout: one `Candidates` sheet without the master tables. Initialization/repair utilities call its local generation method; normal scoring does not call its `export_and_upload` method. That upload method still targets the master path and is unsafe for this generated layout.

Configuration in [config.yaml](HiringAgent_P2/config.yaml) currently uses `qwen3-1.7b-p2:latest`, extraction timeout 180 seconds, scoring timeout 150 seconds, `require_ai: true`, processing retry limit 3, and a two-second row delay. The independent AI recheck remains. Role counts belong to a particular cache snapshot; the September 6 benchmarks/design refer to 106 roles.

Applicant mail is suppressed in YAML and the local `.env`; admin alerts are enabled. Suppressed markers do not count as actual contact and do not start the clarification age-out clock. Environment/GUI overrides must be considered when checking an active process.

## SQLite: implemented pieces and remaining gaps

`SQLiteCandidateStore` provides a WAL-backed `candidates` table, JSON row values, indexes, reads/writes, and a transactional move between main and rejected classifications. `bot.py` exposes `--sqlite` and `--backend sqlite|excel`. The factory fallback is Excel, but the local `.env` selects SQLite. The normal batch launchers change into the P2 directory, where `candidates.db` exists; the current constructor therefore selects that root database unless overridden.

| Finding in current source | Consequence |
|---|---|
| `_scoring_queue_rows`, schema checks, and several maintenance passes still call the SharePoint client directly | Selecting SQLite does not disconnect P2 from Excel or supply database-native intake |
| `sync_from_client` imports only when the database is empty unless explicitly forced | New P1 rows are not incrementally imported into a populated SQLite store |
| The constructor uses `HIRING_SQLITE_PATH`, otherwise an existing working-directory `candidates.db`, otherwise `data/candidates.db`; `init_db.py` uses P2-root `candidates.db` | Normal launchers currently align, but other working directories can silently select a different database |
| `init_empty_db` unlinks the root database; the default seed path calls it too | Initialization is a reset utility, not a safe startup migration |
| Excel queue indexes are passed as hints, while SQLite writes interpret hints as database IDs; `get` and writes also choose different duplicate-ID ordering | Duplicate IDs and mixed addressing need coverage before a cutover |
| The schema is a single current-row table | Durable events, input/result history, job leases, fencing, import conflict reconciliation, and recovery described in the recent design are not implemented |
| Regular scoring exports client results without publishing SQLite state to the master | Excel can retain `New Email Received` after local completion, causing repeat queue selection; master and client results can disagree |
| `ExcelReportExporter.export_and_upload` targets `client._wb_path` but generates only the client `Candidates` sheet | Calling it can replace the master with a workbook missing CandidateList/Rejected and their table contracts |

These are source-review findings, not changes made during documentation cleanup. The latest [SQLite design](HiringAgent_P2/docs/P2_SQLITE_ARCHITECTURE_PLAN.md) and [implementation checklist](HiringAgent_P2/docs/P1_P2_IMPLEMENTATION_STEPS.md) remain as the target specification, with status notes distinguishing the partial code from completed migration work.

## Recent changes and retained evidence

- September 6: key-based Excel store access and duplicate-ID handling; extraction source reporting; thinking disabled, bounded/schema-constrained extraction, and one timeout retry. See the [extraction benchmark](HiringAgent_P2/docs/benchmarks/stage2a/REPORT.md).
- September 6: verified completed-retry filtering and unavailable-resume deferral. See the [availability report](HiringAgent_P2/docs/benchmarks/resume_availability/REPORT.md).
- September 6: normal duplicate merging preserves loser rows/files and reports healing/detection counters. Ordinary scoring no longer calls `client.delete_file`; rejection transfers and specific retry-row cleanup still exist. See section 1 of the implementation checklist.
- September 7 source additions include the SQLite store/CLI switch, separate Excel exporter, initialization/reconciliation utilities, OCR benchmark utility, and benchmark verification script. Their presence does not establish live cutover or measured success.
- The [frozen row audit](HiringAgent_P2/P2_Logs/audits/master_row_audit_20260906/ROW_BY_ROW_REPORT.md), other audit artifacts, benchmark snapshots, and [client handoff](HiringAgent_P2/CLIENT_COMPUTER_START_HERE.md) retain historical evidence and unresolved field-quality issues. Snapshot counts are not current production counts.

This review used local source/configuration and documentation checks. It did not run scoring, reset/seed a database, rebuild/import P1, publish a workbook, contact applicants, or verify cloud state.
