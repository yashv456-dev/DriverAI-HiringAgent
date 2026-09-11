# P1/P2 reliability review — September 8, 2026

P2's intended architecture is local OCR, LLM processing, candidate state and report construction, followed by publication of completed master and client results. The local configuration selects SQLite and `http://localhost:11434`. The current orchestration still mixes local state with a live Excel queue, so selecting SQLite does not complete this architecture.

Scope: local source/configuration review, existing regression suites, and synthetic in-memory/temporary-file probes. No production scoring, database initialization, flow import, workbook publication or applicant messaging was performed. Live deployment and tenant state were not verified.

## Flow and ownership

P1's generated Power Automate definition polls once per minute with one concurrent run and fetches one Inbox message. It screens mail and attachments, applies application-reference/duplicate/update/follow-up rules, saves resumes to SharePoint, and writes the master CandidateList table. Its Excel operations are two reads, one append and seven patches. `New Email Received` hands processing to P2; a resume update can requeue an existing application.

P2's desktop GUI invokes `bot.py`; the batch launchers run that CLI from the P2 directory. The SharePoint scoring path checks Graph/model availability, checks and may modify workbook schema, reads eligible Excel rows, downloads resumes, extracts text locally, extracts candidate fields, performs the independent AI recheck, scores cached roles and applies the existing geo policy. It then saves through the configured store. Resume rename/move operations and several maintenance reads still call SharePoint directly.

There is no normal scoring step that publishes the completed SQLite state to the master. The separate repair utility is not a scheduled master publisher.

## Findings, in completion order

1. **Queue and state disagree.** `_scoring_queue_rows` reads Excel directly, while `sync_from_client` does nothing once SQLite contains any rows. A new P1 application can be absent locally and skipped when `_live_row_index` cannot resolve it. A locally completed application can remain `New Email Received` in Excel and be selected again. Changes to existing applications also lack a durable incremental-import/version contract. Sources: [queue](../hiring_agent/sharepoint_scoring.py), [store/import](../hiring_agent/store.py).

2. **Master publication is missing, and the alternate upload helper targets the wrong workbook.** Normal `export_client_results` publishes scored Main records in the client layout. `ExcelReportExporter.generate_workbook` now produces the same single `Candidates` sheet; its `export_and_upload` still uploads to `client._wb_path`, the master. A synthetic probe confirmed a payload with no Excel tables targeting `Master_Files/Sharepoint_Master_File.xlsx`. Calling that method could destroy P1's required workbook structure. No normal-scoring call site was found for this helper. Sources: [regular export](../hiring_agent/sharepoint_scoring.py), [alternate exporter](../hiring_agent/exporter.py).

3. **SQLite addressing still mixes row offsets and database IDs.** `resolve_position` returns a zero-based position and returns the first matching Application ID even when a later duplicate was hinted. Writes interpret the hint as a database primary key; `get` independently chooses the newest duplicate. A synthetic completed/retry pair reproduced a write to the completed record when the retry position was supplied. The numeric schema `version` is incremented but not used to reject stale writes. Source: [store](../hiring_agent/store.py).

4. **Remote mutations remain inside processing.** `_ensure_schema` runs before work; `_rename_scored_resumes` and `_move_resumes` mutate files; rejection and maintenance still consult Excel. In the Excel backend, rejection is append-then-delete, so a failed delete leaves both rows. In SQLite mode, the local classification update cannot make the preceding SharePoint file move transactional. Direct stale/duplicate-file cleanup has been removed, which addresses a real earlier failure. However, `SharePointClient.move_resume` still implements relocation as download, upload, then `delete_file` at the source; this is not a system-wide no-deletion guarantee. Moves/renames can still leave row references out of step with files. P1 also reuses upload paths for genuine updates, so it does not guarantee preservation of every prior CV version. Sources: [orchestrator](../hiring_agent/sharepoint_scoring.py), [Excel store](../hiring_agent/store.py), [file move](../sharepoint_client.py), [P1 builder](../../HiringAgent_P1/flow/build_zip.py).

5. **Local execution does not bound all processing time.** RapidOCR processes every rendered PDF page synchronously; the Tesseract fallback also has no explicit timeout. Extraction-family calls allow 180 seconds per attempt and retry a timeout once; extraction and the independent recheck have separate budgets, and optional role/portfolio calls can add more. Role scoring has a separate 150-second timeout. There is no overall candidate deadline. Graph requests have a 60-second timeout and retry selected HTTP responses; request timeout exceptions are not caught by that retry loop. Sources: [OCR/extraction](../hiring_agent/extraction.py), [role scorer](../hiring_agent/ollama_scorer.py), [Graph transport](../sharepoint_client.py).

6. **Completion and recovery are not durably separated.** SQLite currently stores one mutable row, without durable input versions, jobs/attempts or pending export generations. A publication failure is logged, and a subsequent idle pass does not have a durable export job to retry. Export itself calls missing-field re-evaluation, so it may perform further extraction and candidate writes. Recheck failure keeps original fields; role-scoring failure can fall back to keywords. `require_ai` protects initial extraction, but does not prove every LLM stage succeeded. Sources: [export/orchestration](../hiring_agent/sharepoint_scoring.py), [recheck](../hiring_agent/extraction.py), [fallback scoring](../hiring_agent/scoring.py).

## Architecture assessment

The existing [SQLite design](P2_SQLITE_ARCHITECTURE_PLAN.md) addresses the central ownership problem: durable P1 intake, local versioned processing, and replaceable Excel outputs. It remains the target specification. The local store switch is only one implemented component.

The completed boundary should be:

1. P1 captures source messages/attachments and publishes durable, deduplicated intake events. Its existing stateful rules must be preserved in the intake consumer; it cannot continue to use a generated master as mutable application state.
2. P2 imports new events/versions transactionally, uses local database jobs as its queue, and caches the exact resume bytes and JD snapshot for each attempt.
3. OCR, extraction, recheck, scoring, review and rejection update local state. Rejection changes classification; source files retain stable references. Bounded processing and infrastructure failures are recorded separately from business outcomes.
4. A successful result commit also records pending publication. An exporter reads a consistent database snapshot and builds the master (`CandidateList` and `Rejected`, with their existing contracts) and the separate client results workbook.
5. Publication verifies each output and tracks its generation. A failed upload retries publication without another OCR/LLM run. Both outputs should identify the same generation so a partially published pair is detectable.

If P1 must keep its current Excel contract temporarily, that is an explicit bridge requiring incremental reconciliation, input-version checks and coordinated writes. Blindly uploading a locally rebuilt master while P1 writes it can discard applications received after the local snapshot. A PC-local worker alone does not resolve this ownership conflict.

Additional operational gaps: use one explicit absolute database path across entry points; replace destructive `init_db.py` startup/reset behavior with migrations; retain and restore-test database backups; enforce one worker/claim owner; add an OCR deadline and distinguish per-call timeout from total candidate time. Preserve the current role/geo behavior while making these storage changes.

## Reproduced checks

Synthetic checks used only `:memory:` SQLite, fake clients and a temporary workbook:

- Initial import copied one record; adding a second fake P1 row and importing again copied zero records and left the new application absent locally.
- Marking the imported row `Scored` locally left the fake Excel queue row at `New Email Received`.
- Supplying the second duplicate's position resolved to the first record and patched the completed record instead of the retry record.
- The alternate exporter generated only `Candidates`, with no Excel tables, and passed the master filename to a fake upload method. No real upload occurred.

Fresh regression results: `test_p1.py` **706 passed, 0 failed**; `test_p2.py` **1,392 passed, 0 failed**; `test_store.py` **68 passed, 0 failed**; `test_extraction_transport.py` **9 tests, OK**. P2 also passed with socket connections blocked to verify the suite can finish without live services. Passing legacy/fake-client scenarios does not establish a complete SQLite intake-to-publication cutover.

## Update — same day, after implementation

Findings 1–6 above were implemented in this codebase, not just designed. Re-verified by re-running every suite from disk (not re-derived from memory of the change):

1. **Fixed.** `SQLiteCandidateStore.sync_from_client` (`hiring_agent/sqlite_store.py`) runs on every `run_local_pipeline` call, not once. Ambiguous imports land in `import_conflicts` and are excluded from the queue instead of being guessed.
2. **Fixed.** `hiring_agent/publisher.py` builds both the master (`CandidateList`+`Rejected`) and client-results workbook per generation and advances the "current" pointer only after both uploads succeed. `exporter.py`'s `export_and_upload` now calls `publisher.safe_target()` before any upload, so it can no longer write a results-only workbook over the master.
3. **Fixed.** Row identity now goes through `begin_attempt`/`finish_attempt` with a version check (optimistic concurrency), not a worksheet-position hint. `test_store.py` section K covers this directly.
4. **Fixed.** `sharepoint_client.py`'s `move_resume` (line ~776) no longer deletes the source — it verifies the copy and returns, source and destination both intact. The local scoring path (`local_pipeline.cache_documents`) never calls it at all; it only downloads and hash-verifies. P1's `flow/reliability_flow.py` (wired into `build_zip.py` via `apply_reliability`) now saves every CV under a GUID-based unique filename plus a companion `intake_<name>.json` sidecar with the exact attachment references — a resend or reply can no longer overwrite an earlier upload. **Not yet live:** the rebuilt zip (`DriverAI-Hiring-AutoReply-apply.zip`) has not been imported into Power Automate; the currently-running flow still has the old overwrite-on-pre-score-update behavior until that import happens.
5. **Fixed.** `local_pipeline.isolated_score` runs OCR + extraction + recheck + scoring for one candidate in a separate process with a single overall deadline (`HIRING_CANDIDATE_TIMEOUT`, default 900s) and kills it on timeout without losing other candidates' results. `worker_lock` enforces one worker at a time.
6. **Fixed.** `attempts` rows track `running`/`interrupted`/`completed`/`failed` state separately from the business outcome; a crashed run is marked `interrupted` and retried on the next pass rather than left ambiguous. `init_db.py` no longer resets an existing database on startup (`seed_from_snapshot` skips seeding when `candidates` already has rows).

All of this is wired into the real entry points, gated by `HIRING_STORAGE_BACKEND=sqlite` (currently set in `.env`): `score_from_sharepoint`, `heal_workbook`, geo-recovery and targeted-rescoring in `sharepoint_scoring.py` all branch to `run_local_pipeline` when that flag is set — this is not orphaned code sitting next to the old path unused.

**Outstanding:**
- Re-import the rebuilt P1 zip into Power Automate to activate the unique-filename/no-overwrite CV fix live (finding 4).
- No live-tenant/live-Ollama run has exercised this path yet — all verification above is against the regression suites and synthetic fakes.
- The old Excel-backend code (`store.py`, the Excel branch of `sharepoint_scoring.py`) still exists as a fallback if `HIRING_STORAGE_BACKEND` is ever unset; it still has the pre-fix positional-addressing and delete-then-add rejection behavior. That's fine as long as the env var stays `sqlite`, but it means the fix is a posture, not a deletion of the old bug.
