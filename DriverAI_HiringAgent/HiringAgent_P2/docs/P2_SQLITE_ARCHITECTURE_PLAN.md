# P2 reliability architecture and implementation plan

Design recorded September 6, 2026. Status reviewed September 7: a basic SQLite store, CLI backend selection, and export/initialization utilities now exist locally. The durable intake, versioned state, worker recovery, and complete P1/P2 cutover specified below are still unimplemented. See [current architecture](../../CURRENT_ARCHITECTURE.md) for verified source-level gaps. The original design preparation did not change tenant resources.

## Decision and scope

Design for one P2 host using **SQLite for application state and results, immutable JD JSON snapshots for scoring inputs, SharePoint for source documents and durable intake, and Excel for generated reports**. This assumes the single-host option discussed with the user; it is not a decision to deploy SQLite without further implementation review. If several independent worker hosts must write concurrently, use PostgreSQL through the same service interfaces instead.

Read alongside the retained [Stage 2A measurements](benchmarks/stage2a/REPORT.md) and [current architecture](../../CURRENT_ARCHITECTURE.md). This design supersedes the retired SharePoint List roadmap. Swapping ExcelCandidateStore for SQLite is insufficient: P1 and queue/maintenance reads still use the workbook. The September 6 containment change removed ordinary scoring file deletions, while rejection transfers and specific row cleanup remain.

Preserve the existing 106-role scoring prompt, role ordering and ranking behaviour; ai_recheck_fields as a second opinion; the geo three-way decision; score_retry_max=3; row_delay_seconds=2; extraction timeout=180 and scoring_timeout=150. Preserve existing report columns: CandidateList's 33-column contract, Rejected's existing layout, and the separate client export layout. Both applicant and admin mail remain suppressed throughout implementation and testing. Plan the extraction fixes below as separate work; do not start grounded-extraction rework as part of the storage replacement.

## Evidence driving the design

The [frozen row audit](../P2_Logs/audits/master_row_audit_20260906/ROW_BY_ROW_REPORT.md) provides the worksheet rows, stored values and source evidence behind this design.

The snapshot contains 215 CandidateList records, 79 Rejected records and one empty table row. Seventeen Application IDs occur across both sheets. Among 182 Scored records, 28 lack Years Exp and 10 lack Phone; among 79 Rejected records, 68 lack Years Exp. Twenty-seven path-based resume URLs point to files absent from the current inventory. Blank and document-ID URLs were counted separately, not assumed broken.

Confirmed failures include file deletion before failed Excel row deletion, misleading merge-success counters, backfill excluding Years Exp and Rejected, and an unavailable-resume exception aborting a repair sweep. That exception regression was corrected locally before this plan. The audit also found decimal experience misparsing and a conflict between preserving bare phone numbers and a backfill routine requiring US formatting. The database cannot correct these extraction defects by itself.

## Target flow

Mail arrives and P1 applies its filters and durable capture, producing SharePoint resume files and immutable intake events. P2 imports those events under the existing intake rules into the SQLite applications and jobs tables. It then claims one application version, downloads every resume attachment by file ID, extracts text and fields, runs the independent AI recheck and validation, scores the existing 106 roles against a validated JD JSON snapshot, and applies the existing geo verdict. Result and outcome commit together back to SQLite. An unavailable or unreadable source, and any extraction, recheck or scoring failure, is recorded against the same record instead of being dropped. Excel reports are built and validated from the committed state, and a completed report generation is published.

That ordering states processing dependencies, not permission to change geo classification order or early-exit behaviour. Retain existing tested business ordering during the storage phase; move it only with equivalent scenario tests.

## Ownership rules

| Component | Owns | Must not do |
|---|---|---|
| P1 | Mail selection, durable source capture, intake delivery | Write into the generated P2 report or contact anyone during the migration |
| P2 intake adapter | Event deduplication, application/update identity under existing intake rules | Guess identity from a filename or blindly collapse applications sharing an email |
| SQLite | Canonical application versions, processing state, results, history, export jobs | Depend on an Excel row number or be cleared at startup |
| Resume repository | Original bytes, file identifiers, hashes and attachment manifests | Delete or move a resume as a side effect of rejection or duplicate cleanup |
| JD loader | Validate and freeze the exact role set for a run | Change the role set halfway through a batch |
| Exporter | Render database snapshots into existing Excel layouts | Feed report edits back into SQLite implicitly or clear the published workbook first |

## P1 handoff: required work, not an assumed connection

Power Automate cannot open a local SQLite file directly. For this single-host design, use a SharePoint document-library inbox rather than exposing the PC through a public API. The SharePoint connector supports file creation; the durable event protocol below is application logic we must implement and test. [Microsoft connector reference](https://learn.microsoft.com/en-us/connectors/sharepoint/)

1. Capture one immutable intake event for each relevant source message, including updates and follow-ups. Assign an event key derived deterministically from the source mailbox/message identity, not a random key per retry. Store a schema version, source timestamps, sender, original message reference, body, any referenced Application ID, and the expected attachment manifest. Verify that the selected message identifier remains usable after archive/move; do not assume the existing connector ID has that property.
2. Upload attachments into an event/version-specific folder. Record drive ID and item ID, original filename, size and the available version metadata. Do not overwrite a previous resume version. P2 records a content hash after downloading.
3. Publish a completion manifest only after required source capture succeeds. P2 accepts only complete manifests. A partial upload remains pending capture; it must not masquerade as a completed update or erase the old resume reference. A message with no attachment can be a valid no-resume event under the current rules; an attachment expected but failed is a different case.
4. P2 polls and imports complete manifests. A unique event key makes re-delivery harmless. Import the event, application revision and job in one local transaction. On restart, re-read an overlap and periodically reconcile the full manifest inventory; never rely solely on a last-seen timestamp. Do not delete manifests when consumed.
5. P1 marks capture handled only after durable handoff is verified. Failed capture stays explicitly recoverable and is reconciled against source messages. If the host is offline, manifests accumulate in SharePoint and processing resumes when it returns.

Out-of-order delivery must not replace a newer application version with an older message. Preserve the original received timestamp and a deterministic tie-breaker; a late event can attach to history or update a pending revision only under the existing update rules. If the same event key arrives with different payload bytes, quarantine the conflict instead of overwriting it. After an uncertain upload response, verify the existing event/file before repeating the operation; a duplicate filename alone is not proof of a successful handoff.

**P1's stateful decisions are a real migration dependency.** Its current two lookups, append and seven patches cannot keep reading/writing the generated report. Inventory and port all ten Excel operations and the duplicate/update/follow-up branches. Preserve reference-ID rules, the 90-day window (including max submitted timestamp), application-update limits, and mail-filter scenarios. Put stateful application decisions in the SQLite intake consumer; P1 remains the durable event producer. Any acknowledgement dependent on those decisions must wait for the outcome, with mail still suppressed. This changes ownership, not the intended business rules, and requires P1 scenario parity tests.

For shadow evaluation only, P2 may import a frozen workbook or read the existing P1 workbook into a separate local database and write separate local reports. A changing Excel table is not an immutable event queue: this bridge cannot guarantee capture of intermediate updates. Do not call that arrangement the final cutover or let both systems write the same report. If P1 cannot be changed, keep this as a limited shadow design and prefer the previously proposed SharePoint List route for a separately reviewed migration.

## Database model and invariants

Use an internal application UUID and preserve the external Application ID. Reconcile legacy collisions before enforcing external-ID uniqueness; conflicting identities go into an import-conflict queue. An email address or phone match is supporting evidence, not a database primary key.

| Entity | Essential data and constraint |
|---|---|
| intake_events | Unique event key, original payload/hash, source references, capture and import state; replay cannot create another application/update |
| applications | Internal ID, external Application ID, current input revision, current accepted result, existing business status, row-level retry count, optimistic revision number |
| application_versions | Immutable input snapshot, parent application, event key, update sequence; references all attachments, mail-body version and approved corrections |
| documents / version_documents | Drive ID + item ID + fetched version/hash, filename and ordered attachment slot; preserve separate attachments even if names collide |
| jobs / attempts | Input version, stage, owner lease and fencing token, attempt timings, next eligible time, exception class and source provenance |
| result_versions / result_fields | Immutable extraction, second-opinion output, final accepted values, scores, geo verdict, per-field quality, model/prompt/parser versions |
| jd_snapshots | Hash and local immutable JSON reference containing the exact scoring payload and order used |
| audit_events / export_jobs | State-change history and durable pending report generations; unique generation ID and database snapshot revision |

Every result references one exact input version and one JD snapshot. A completed job key includes application input version, JD hash, model identity, parser/prompt versions and relevant configuration hash. Replaying the same job is a no-op unless an explicit new attempt is scheduled. Distinct versions are never combined merely because they belong to the same person.

SQLite setup: local persistent disk, foreign keys enabled on every connection, WAL mode, durability settings appropriate to the host (proposed synchronous=FULL), bounded busy handling, and short write transactions. SQLite permits one writer at a time; WAL allows concurrent reading of committed snapshots. Do not keep a write transaction open during Graph requests, OCR or model inference. [SQLite transactions](https://www.sqlite.org/lang_transaction.html), [WAL documentation](https://www.sqlite.org/wal.html)

## Worker lifecycle and crash recovery

1. Start by checking configuration, schema version, disk space, database access, mail suppression and model availability. A second scheduler invocation must not run a second copy of the same job.
2. In a short transaction, claim one eligible job and record a lease plus fencing token. Commit immediately. Renew the lease during long OCR/model calls. A stale worker cannot commit after another worker has claimed the job.
3. Resolve the full attachment manifest. Download using drive/item IDs, verify source version before and after download, hash bytes and extract text per attachment. A changed file creates a new input version or invalidates this attempt; do not attach new bytes to an old version silently.
4. Run existing extraction and the independent AI recheck. Persist local stage artifacts against this attempt and input hash; these are not yet accepted candidate results.
5. Score against the frozen JD set using existing ranking behaviour, and apply the existing geo decision. Preserve all role candidates in the scorer, not a shortlist.
6. In one transaction, verify the lease, application input revision and job key, then insert the accepted result, update the application's current result/status, finish the attempt, and queue export work. If a new resume arrived meanwhile, retain this attempt as superseded; it cannot replace the current result.
7. Release the job and continue the batch. Export failure cannot undo a scored result or trigger another model run. A crash after commit replays pending export work; a crash before commit leaves the previous result intact and the expired job recoverable.

There are three separate dimensions: processing state (pending/running/deferred/review/completed), extraction quality (assessed, incomplete, conflicting, failed, legacy-unverified), and the existing business/geo status. One generic Status column must not carry all three meanings internally.

An updated resume makes its previous score historical, not current. Keep that result in the database, but do not show an old score as if it belongs to newly extracted fields. While reprocessing, the current report shows the existing pending/review status and does not publish a misleading mixture of versions.

## Failure policy

| Failure | Action | Row-level retry / hiring effect |
|---|---|---|
| Expected file not available / 404 | Mark source unavailable, record exact attachment, queue reconciliation | Do not count as unreadable or reject; preserve prior data |
| 401/403 | Record access failure and pause affected work until credentials/permissions are corrected | No candidate penalty |
| Graph throttling / temporary server failure | Bounded transport handling, respect Retry-After, defer if still unavailable | Separate infrastructure attempt record |
| Extraction Timeout | Exactly one immediate retry; after the second timeout retain fallback only as diagnostic and defer under require_ai | No row-level give-up increment |
| Extraction ConnectionError | No immediate retry; record service unavailable and resume through a later operational recovery/run | No row-level give-up increment |
| Invalid/truncated model JSON | Keep json.loads failure handling; record failure and do not publish a model-completed result | No extra extraction retry under the timeout-only policy |
| Downloaded bytes yield no readable text | Existing unreadable-resume review policy | Preserve score_retry_max=3 and current give-up semantics |
| Database busy/disk full/save failure | Roll back incomplete save and retain job/prior result; surface operational failure | No success counter, no candidate rejection |
| Workbook upload/lock failure | Keep last published report; retry export only | No rescore and no candidate penalty |
| New input while processing | Reject stale commit and queue the newer version | No rejection or duplicate application |

Keep the current think=false, observed 512-token output cap and JSON schemas; record token usage and truncation. Revisit the cap only from measurements on representative real resumes. Preserve require_ai=true and the distinction between deliberate deterministic field extraction and an LLM-failure fallback. Infrastructure deferral counters are separate from score_retry_max=3. No new retry loops may accidentally multiply the timeout-only extraction retry.

## Missing data and field repair

The storage phase preserves extraction behaviour and records its limitations. A separately tested repair phase addresses the audit findings without changing the 106-role scorer or the geo decision.

For each field retain raw value, normalized value, extraction source and assessment state: present, not stated after review, unresolved, conflicting, or extraction failed. An empty regex result is **unresolved**, not proof the resume omitted the field. Legacy imports start unverified. Store existing provenance now; do not claim source-grounded validation until it has actually been implemented and evaluated.

Repair jobs must query both business outcomes and all relevant missing fields, not only Scored Main phones/dates. Check full_name, phone, email, location, country, experience (the actual key mapped to Years Exp), education and education dates, plus the existing skills/role fields. Each repair reads the exact input version, records before/after and the reason, and proceeds past another candidate's unavailable source.

Specific work items: recover explicitly stated totals such as Sanchit's two years; prevent 2.8+ becoming 8 and 1.5+ becoming 5; preserve a supplied bare phone as contact data without manufacturing evidence of US residence; distinguish sender email from CV email; avoid overwriting a populated value with an unresolved blank. Do not infer total experience from employment dates or invent education dates. If a correction affects scoring or geography, it must go through the unchanged decision logic in a new result version rather than editing the displayed field alone.

Valid results can contain genuinely unstated optional data. Technical extraction failure cannot become ordinary Scored under require_ai. Do not add a rule that rejects people for lacking optional resume fields. Keep quality reasons in an accompanying processing report or application view, without expanding the 33-column candidate contract.

## Resume and JD lifecycle

Resumes stay in a stable event/application-version folder after capture. Rejection changes the application record only. Store drive ID and item ID as the retrieval key; filenames and URLs are presentation metadata. Item IDs survive rename/move within their drive, while path URLs change. A cross-drive copy or deleted/recreated file needs a new verified reference; an expired download URL is not a permanent file identifier. [Microsoft ID-based addressing](https://learn.microsoft.com/en-us/graph/onedrive-addressing-driveitems)

No automatic destructive duplicate cleanup runs in the scoring path. Link duplicate source records to a canonical application, preserve originals/history, and separately review retention/deletion. A same-email file discovered elsewhere is a recovery candidate, never an automatically accepted replacement version.

Keep jd_sources.json as source configuration and jd_roles_cache.json as the current cache. They are different roles. Freeze the full validated scoring payload into a content-addressed snapshot for each run, preserving the existing 106 roles, ordering and prompt. Include existing opening_id plus a stable per-document/role identifier where multiple documents share an opening. Record source document references, payload hash and build time. A refresh builds and validates a new snapshot before making it current; in-flight jobs keep the old snapshot. A failed refresh retains a known-good snapshot and records staleness. With no approved usable role set, defer rather than silently score against a different built-in set. This last guard is a planned explicit failure-path correction to today's fallback, tested separately from normal scoring parity.

## Excel refresh, visibility and backups

Generate reports from one consistent SQLite read snapshot at batch completion and on an explicit refresh. Record when intake was last polled and when the report snapshot was taken. A report remains available while new candidates are being processed; it is not continuously cleared and rebuilt in place.

CandidateList and Rejected are mutually exclusive projections of one application table using existing business status rules. Preserve the separate client export's existing field selection. The current export_client_results function exports Main only, so it is not already a replacement master-workbook publisher; build and test that publisher explicitly.

Build a new local workbook, reopen it, and verify schemas, unique application IDs, expected row counts, disjoint Main/Rejected sets, absence of empty data rows, stable source references, and literal text handling for untrusted resume/email values. Counts reconcile to publishable canonical applications; import conflicts appear in the processing report and must never silently disappear. A pending application can have missing fields, but cannot be labelled newly scored.

Publish a versioned, complete workbook, verify its remote metadata, then update a latest-generation pointer using a version check. Readers retain access to the prior good generation if upload/promotion fails. Remote publication and SQLite commit are not one atomic transaction; the durable export job reconciles uncertain uploads by generation ID before retrying. A fixed Sharepoint_Master_File.xlsx convenience copy may be maintained, but a file lock leaves that copy stale and visibly reported; never delete it before replacement or claim it is current without verification.

The database is never emptied automatically. Back it up with SQLite's backup API before migrations and on a scheduled basis, with a restore drill. Copy completed backup artifacts off the host; do not place the live database in a OneDrive sync/network directory or copy only its main file while WAL writes are active. Resume files, JD snapshots and original intake manifests need their own retained copies/version policy. [SQLite online backup](https://www.sqlite.org/backup.html)

## Operational visibility and mail

One run summary must distinguish imported events, duplicate replays, attempted jobs, committed results, superseded attempts, unavailable sources, unreadable sources, degraded extractions, unresolved field counts, pending exports and published generations. Count success only after the relevant operation commits. Show the oldest pending job and exact failure reason per application; repeated infrastructure failures remain visible instead of silently looping or penalizing applicants.

Both mail gates remain off. Suppressed actions must be counted as suppressed, never sent. When mail is separately authorized later, use durable notification intents tied to result versions and recipients. A database transaction cannot guarantee exactly-once email delivery; a crash after remote acceptance requires reconciliation rather than blindly resending. Neither storage migration nor report refresh releases accumulated suppressed messages.

## Implementation sequence and acceptance gates

| Phase | Deliverable | Exit evidence |
|---|---|---|
| 0. Contain current failures | Reviewable local changes removing destructive cleanup from ordinary scoring, truthful counters and independent backfill exception handling | Missing-file-first test continues to next candidate; simulated failed deletes do not remove resumes or claim successful merge |
| 1. Reconcile import | Frozen workbook + files + source references, every legacy row mapped to canonical application/version or explicit conflict | All 294 records accounted for, blank row documented, 17 duplicate-ID pairs adjudicated; no name/email-only automatic relinking |
| 2. SQLite store and worker commits | Schema migrations, backup/restore, short claims, immutable results and export outbox behind the existing store seam | Crash, replay, stale-worker, disk-full and updated-resume tests; no tenant writes |
| 3. Offline/shadow P2 | Existing extraction/recheck/scorer over local fixtures, SQLite outputs, separate Excel reports | Same input yields expected field/role/geo parity; all report rows reconcile to database |
| 4. P1 durable intake | Draft flow and intake consumer; port all ten Excel operations and existing stateful scenarios | Duplicate delivery, interrupted uploads, follow-ups, 90-day window, host-offline recovery and out-of-order events tested |
| 5. Field repair | Separate fixes for explicit experience/decimals, phone preservation, field coverage and both outcomes | Reviewed before/after evidence for audited examples; no regression in geo or invented dates/totals |
| 6. Reviewed cutover | Frozen import, durable intake catch-up and one authoritative P2 database; report publication enabled separately | Two-candidate dry-run plus representative shadow batch; restore tested; P1 no longer writes generated reports; mail off |

Phases are separate reviewable changes, not one module rewrite. Keep sharepoint_client.py for document transport, extend store.py for database persistence, and move only orchestration needed for claims, versioned commits and report delivery. Reuse extraction.py, jd_sources.py and the scorer rather than reimplementing their business logic in the storage layer. No deployment dates are promised before the P1 mapping and import conflicts are resolved.

Validation preserves existing P2 1393/0 and P1 706/0 behaviour assertions. The previously observed store suite is 59/0, despite an earlier requested count of 53; record that discrepancy rather than deleting tests to force a number. Keep new architecture/failure tests separate so baseline counts remain interpretable. Re-run the appropriate suites after implementation; this planning-only change does not claim a fresh suite run.

Required end-to-end evidence (read-only tenant, --dry-run for any E2E): same two selected candidates with traceable source versions; baseline/after values for full_name, phone, email, location, country, experience, education and skills; all per-call timings and provenance; no tenant mutation methods or mail calls. Use representative known failures in local tests as well, not only two clean resumes. The storage change must not claim to make model inference faster or extrapolate a synthetic timing result to all real CVs.

Rollback is not just restoring an old database: pause consumers/publishers, record the last imported event and report generation, restore a verified backup and replay retained manifests. Prevent new and old workers from running simultaneously. If returning to the legacy intake/Excel worker, reconcile all post-cutover events and restore a compatible workbook/flow together, with destructive cleanup and mail still disabled. Never overwrite newer applications with a pre-cutover snapshot.

## Done means

Every intake event is accounted for; every accepted result belongs to the current source/JD version; rejection does not move/delete records or resumes; a crash cannot publish partial database results; reruns do not duplicate applications; one bad source does not stop a batch; unresolved fields have an honest explanation; Excel is a verified, replaceable report; and a tested restore can resume processing without losing newer applications. Existing historical mistakes are corrected only by the reviewed import/repair process, not by changing storage alone.
