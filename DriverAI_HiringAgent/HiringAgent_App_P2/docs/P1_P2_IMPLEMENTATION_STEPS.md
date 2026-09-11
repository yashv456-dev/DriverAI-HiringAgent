# P1 and P2 implementation steps

Status reviewed September 7, 2026. This remains the acceptance checklist for the full reliability design. Section 1's code containment shipped September 6; its fresh evidence snapshot remains outstanding. A basic SQLite store, CLI backend selection, and initialization/export utilities now exist, but they do not satisfy the event-import, versioning, recovery, or cutover requirements below. Do not interpret the existence of a candidates table as completion of sections 2-13.

Target: one P2 host with SQLite; SharePoint retains documents and durable intake events; versioned JD JSON supplies the scorer; Excel is generated output. See [the design](P2_SQLITE_ARCHITECTURE_PLAN.md), [current implementation and gaps](../../CURRENT_ARCHITECTURE.md), and [retained extraction measurements](benchmarks/stage2a/REPORT.md). This checklist records requirements; it is not evidence of a production cutover.

## 1. Preserve evidence and stop the known destructive behaviour

Before implementation, retain a new consistent workbook snapshot, source-file inventory, existing flow package/configuration, JD cache, code/configuration snapshots, and relevant logs. Record hashes and snapshot times. The prior audit is evidence, not a substitute for a fresh cutover snapshot.

Prepare a small containment change that removes resume deletion and destructive duplicate cleanup from ordinary P2 runs. Do not merely reverse the delete order: database state and SharePoint file deletion cannot be made one local transaction. Preserve files until a separate retention procedure explicitly authorizes deletion. Fix counters so attempted merges and failed saves are not reported as completed operations. Preserve the already-corrected per-candidate backfill exception handling.

Keep both applicant and administrator mail suppressed. Leave score_retry_max=3, row_delay_seconds=2, extraction timeout=180, scoring_timeout=150 and require_ai=true. Do not change the 106-role prompt, independent AI recheck or geo rules.

Exit check: a simulated unavailable file and failed row deletion leave resume bytes and existing results intact; the following candidate still runs; the summary reports the failure truthfully. No production test is required to prove this.

> **Build status, 2026-09-06 - the containment change shipped, the evidence snapshot did
> not.** The first paragraph above (workbook/source-file/flow-package/JD-cache/log snapshot
> before implementation) needs live-tenant access and was not done. The second paragraph -
> the actual code change - is done:
>
> - `hiring_agent/sharepoint_scoring.py` no longer contains a single `client.delete_file(`
>   call. All three sites that used to remove a resume during an ordinary run now preserve
>   the file and log what would have been removed: the duplicate-merge loser's file, the
>   superseded file on a repeat-applicant Rejected refresh, and the stale-sibling guess after
>   a rename (the site with the documented 2026-08-24 near-miss, APP-20260824-2049-642B, and
>   previously zero test coverage).
> - `_merge_duplicate_candidates` no longer deletes the loser ROW either, on either sheet -
>   same rationale: a row delete and a file delete are two ungrouped Graph calls, and a run
>   that completes one but not the other used to leave a dangling reference with no way to
>   tell which half happened. It still detects duplicate groups and heals the winner's blank
>   fields from the loser; the loser is left in place, unresolved, for the retention
>   procedure this exit check assumes exists later. Return shape changed from `{'merged',
>   'errors'}` to `{'healed', 'duplicates_found', 'errors'}` - `duplicates_found` is a count
>   of rows located, not a count of anything removed.
> - Counter honesty: a review pass on this same change found the merge function's early
>   error return still carried the pre-change dict shape (`{'merged': 0, 'errors': 0}`),
>   which every caller reads with `.get()` so it never raised, but silently diverged from
>   the documented contract. Fixed, and pinned with a dedicated test.
> - Proof, not just claim: `test_p2.py` includes a stub whose `delete_row`/`delete_rejected_row`/
>   `delete_file` methods raise `AssertionError` if called at all, and asserts a real
>   duplicate-merge run never touches them; a static source check asserts
>   `"client.delete_file("` does not appear anywhere in the file; and the repeat-applicant
>   refresh test asserts its recorded deletions list is empty where it previously asserted a
>   deletion happened.
> - Not touched: `move_resume` (a relocate, not a delete - still needed), the 33-column
>   contract, `config.yaml`, extraction/scoring/geo logic, and the already-correct backfill
>   exception handling this section says to preserve.
>
> Verified: `test_store.py` 59/0, `test_p2.py` 1392/0 (net -1 from the 1393 baseline - several
> old assertions checked deletion OUTCOMES that no longer exist and were replaced by fewer,
> stronger ones proving non-deletion, including the counter-shape fix above),
> `test_p1.py` 706/0 unchanged.
> The pre-implementation evidence snapshot in this section's first paragraph is still
> outstanding and needs a live-tenant session to complete.

## 2. Define the contract before changing either worker

Write down the ownership of every existing column and workflow decision. P1 supplies captured message/attachment facts; the intake consumer owns application identity and updates; P2 owns extraction/scoring; report generation owns only presentation. Distinguish sender email from resume email and original input values from normalized values.

Preserve CandidateList's 33 columns, Rejected's existing layout and the separate client export layout. Add internal database metadata without adding candidate-sheet columns. Put processing reasons, data-quality assessments and import conflicts in a separate operational report.

Define three independent states: processing state, field-assessment state and the existing business status/geo verdict. A successful score need not imply every optional field was stated. Conversely, model failure cannot produce an ordinary model-completed result. A new resume version invalidates the old result as the current score while retaining its history.

Exit check: every existing field, status and P1 scenario has an owner and a defined mapping; unresolved decisions are listed explicitly, not filled with invented defaults.

## 3. Reconcile legacy records into a reviewable import manifest

Read all source rows by Application ID plus original sheet/row reference. Capture every row as an immutable import record before merging anything. The previous snapshot had 294 candidate records, 17 duplicate-ID pairs and one blank row; use these as prior evidence, not fixed counts for a later snapshot.

For each duplicate-ID group compare identity, message timestamps, original attachments, update history and completed results. Map it to one canonical application with retained versions, or flag an identity/version conflict. Do not choose the newest physical row automatically. Never enforce unique Application IDs by discarding records that violate the constraint.

Resolve files using verified metadata and content where available. A matching name or email is a recovery lead; it does not prove that a different file belongs to that application version. Preserve unresolved links as explicit import issues. Do not bulk-copy the audit's parser suggestions: it includes demonstrated incorrect decimal-year results.

Exit check: every nonblank legacy row maps to a canonical record/version or a visible conflict record, every excluded blank row is documented, and no applicant silently disappears.

## 4. Build SQLite persistence and backup recovery

Create a versioned schema on persistent local disk with entities for intake events, applications, input versions, documents/attachment slots, jobs/attempts, result versions/fields, JD snapshots, audit events and export jobs. Use internal IDs with foreign keys; preserve external Application IDs separately and enforce uniqueness only after collision reconciliation. An email address must not be a primary key.

Use short write transactions, optimistic application revisions, unique event/job keys, and worker leases with fencing tokens. Configure foreign keys on every connection, bounded database-busy handling and WAL/durability settings as described in the architecture. Do not hold a transaction across Graph, OCR or model calls.

Implement an explicit import command and separate schema migrations. Startup must never recreate or empty the database. Back up with SQLite's backup mechanism before schema changes and on a schedule; copy completed backups off-host. Test restoration into a separate location and reconcile retained intake events after restore.

Exit check: duplicate import is harmless; rollback preserves previous values; restart resumes pending work; backup restoration reproduces counts and result references.

## 5. Replace P1's workbook-dependent handoff

Keep existing mail selection and classification behaviour. Create immutable intake events in a designated SharePoint folder so incoming work remains durable while the P2 machine is offline. This folder is document storage, not a SharePoint List migration.

Each event contains a deterministic event key, schema version, source mailbox/message reference, received timestamp, sender, subject/body, any referenced Application ID, event details, and an attachment manifest. Select and test source message identity across mail moves; do not assume a move-sensitive connector ID is permanent.

Upload each new attachment version separately. Retain original filename, drive/item identifiers and available version metadata. P2 adds verified hashes when reading the bytes. Publish an event completion manifest only after all expected capture operations succeed. If no attachment was supplied, explicitly record that fact; if attachment upload failed, keep capture incomplete. Never patch the application as updated while still pointing to the previous resume.

P1's existing Excel operations all need an explicit replacement:

| Existing operation | Replacement responsibility |
|---|---|
| Add_row | Event import creates the application/version exactly once |
| Get_rows and Get_rows_ref | SQLite intake consumer resolves application/update identity using the preserved rules |
| Patch_dup_attempts and Patch_dup_attempts_noreply | Record and apply the existing duplicate-attempt decision exactly once |
| Patch_mail_sent | Record the actual notification outcome separately; suppression is not a send |
| Patch_update_attempts and Patch_update_attempts_noreply | Apply accepted update/version transitions only after durable capture |
| Patch_followup_attempts and Patch_followup_attempts_noreply | Apply the existing follow-up semantics against canonical state |

This ports ten operations. It is a deliberate P1 change, not a filename/configuration swap. Preserve the reference-ID behaviour, 90-day duplicate window, max submitted timestamp logic, application-update limits and every existing scenario. State-dependent acknowledgements must wait for the intake decision; mail remains disabled in this work.

Exit check: repeated delivery, interrupted uploads, updated resumes, reference replies, no-resume messages, duplicate-window boundaries and host-offline recovery all produce the intended application state without duplicate writes or mail.

## 6. Make event import repeatable and ordered

P2 reads only complete manifests and imports their payload, application transition and job in one transaction. A repeated event key with identical content is already handled. The same key with different content is a conflict, not an update to overwrite.

Handle late messages using source timestamps and a deterministic tie-breaker under the existing intake rules. An older event cannot silently replace a newer current resume. Retain it in history where appropriate. Track import progress locally; use an overlap when polling plus periodic full reconciliation. Do not delete cloud manifests simply because one worker consumed them.

If an upload response was uncertain, check the existing event/file before retrying. Import recovery must work after the worker stops between any two stages, without relying on an in-memory cursor.

Exit check: replay a captured event batch twice and in different delivery orders; current application versions and update counts remain correct.

## 7. Replace path guessing with versioned document retrieval

Download by drive ID and item ID, then verify the expected document version and calculate a content hash. Maintain an ordered attachment manifest and extract text from each required slot. Multiple attachments with the same filename are separate slots, not interchangeable files.

An expected attachment that cannot be downloaded blocks acceptance of a partial extraction. Distinguish unavailable source, permission failure and downloaded-but-unreadable bytes. A resume changed during download/processing creates a newer input revision or invalidates the attempt; it does not silently become the old version.

Do not move resumes when rejecting applications. The database status controls which report shows the candidate. Preserve source history and avoid using short-lived download URLs as permanent references. Cross-drive copies and deleted/recreated files require verified new references.

Exit check: rename within a drive, missing file, access denied, second attachment missing and file replacement all follow the correct path without scoring the wrong bytes.

## 8. Preserve scoring while fixing the known extraction defects separately

First run the existing extraction/recheck/scoring behaviour against database-backed fixtures. Keep the independent ai_recheck_fields call and the full 106-role prompt. Retain thinking disabled where currently supported, JSON-schema output, json.loads failure handling and the observed output cap. Do not enlarge the cap or timeout without measurement.

Retry an extraction Timeout exactly once. Do not immediately retry ConnectionError. If the model remains unavailable, record an infrastructure deferral under require_ai without incrementing the candidate row-level give-up counter. Keep the existing unreadable-resume give-up threshold of 3 for its intended failure class.

Then implement separate, targeted field-repair changes: include both Main and Rejected outcomes; cover name, email, phone, location, country, experience, education/dates and skills as appropriate. Correct the demonstrated 2.8-to-8 and 1.5-to-5 errors. Preserve a bare supplied phone as contact data without converting it into evidence of US residence. Distinguish unresolved values from confirmed-not-stated fields; preserve source and before/after values. Do not calculate experience from job dates or invent education dates.

Any correction that affects scoring or geography creates a new result version through the existing rules. Do not patch displayed skills/experience while retaining a score derived from different data.

Exit check: manually reviewed before/after values on the audited failures, no new fabricated fields, unchanged geo/ranking rules, and per-call latency/provenance recorded on the same resumes.

## 9. Freeze the JD set and commit a candidate result safely

Validate the configured JD JSON and save the exact ordered scorer payload as an immutable hashed snapshot. Record its ID with each job/result. Keep source configuration separate from the cached role payload. Refresh into a new snapshot; in-flight work keeps its original snapshot. Never silently replace the full approved role set with a fallback subset after a refresh error.

Claim one pending application version using a short transaction and a worker lease. Run inference outside the transaction. Before accepting the result, check the lease/fencing token and current input revision. An expired or superseded worker may save diagnostic history, but cannot replace the current result.

Commit the final fields, scoring output, geo/business status, completed attempt and pending export job together. If saving fails, report failure and retain the prior accepted result. Infrastructure attempts and candidate retry count remain separate. A new resume during scoring must produce a new job, not a stale result overwrite.

Exit check: stop the process before and after commit, start overlapping schedulers and inject a newer resume mid-run. At most one current result exists for the intended version, and a failed commit is never counted as success.

## 10. Generate Excel instead of maintaining rows in place

Read one consistent committed database snapshot and generate CandidateList and Rejected from mutually exclusive status selections. Preserve the existing schemas and the separate client export. Do not add blank separator records or treat an Excel row number as application identity.

Build a new local workbook and reopen it before publication. Check headers/order, record counts, duplicate IDs, Main/Rejected overlap, missing links, empty rows, and text/formula handling. Report unresolved import conflicts separately so count reconciliation exposes them. Fields legitimately absent may remain missing, but an operational failure must not be labelled newly scored.

Upload a complete versioned report, verify the uploaded generation, and only then promote it as latest. Retain the previous good generation. A fixed-name master copy is a convenience output; if it is locked or upload fails, show that it is stale and retry export. Never clear/delete the published master before replacement. Database commits and SharePoint uploads are separate operations, joined by a durable export job and reconciliation.

Do not re-run inference because an export failed. Excel edits must not silently flow back into SQLite; an explicit reviewed correction/import path is required.

Exit check: simulate file lock, upload timeout, lost acknowledgement and process restart. Candidate scores stay saved, the previous report stays usable, and a retry publishes the intended generation once.

## 11. Show accurate progress and preserve mail suppression

Report imported events, replayed events, attempted candidates, committed results, source-unavailable deferrals, unreadable resumes, degraded extractions, conflicts, missing-field assessments and pending/published exports separately. Include the oldest pending work and a per-application error reason.

A failed merge is not a successful merge; an attempted model call is not a committed score; a suppressed email is not a sent email. Keep applicant and admin transport disabled. Storage migration and workbook refresh must not release the existing backlog of suppressed notifications. A future mail rollout requires its own reviewed delivery/reconciliation design.

Exit check: counters reconcile to recorded attempts, commits and exports, with zero tenant mutation or mail calls during read-only validation.

## 12. Validate, cut over and retain a working rollback

Run the existing P2 1393/0 and P1 706/0 behavioural suites after implementation. The observed store suite is 59/0 despite an earlier requested 53; record the discrepancy instead of deleting tests to match a number. Keep added architecture tests separate. Also test real data quality: green unit-test counts alone did not detect the damaged workbook.

Run a two-candidate end-to-end dry-run using the same verified source versions and all required fields, plus a representative local/shadow batch containing the known failure cases. Measure per-call timing and compare values/roles/geo outcomes. Any live-tenant E2E during preparation uses --dry-run and a read-only mutation guard; local scratch databases and reports are separate from production state.

Prepare the concrete cutover artifacts: versioned code/configuration, flow package, reviewed import mapping/conflicts, populated and backed-up database, parity results, generated workbook and restore instructions. Production cutover is a later authorized action, not part of writing this checklist.

During cutover, pause intake/consumers long enough to establish a final snapshot and message boundary; keep incoming mail retained. Import final changes, enable the durable P1 handoff, reconcile messages/events around that boundary, start only the new worker, then publish the verified report. P1 and P2 must no longer share a writable report workbook. Keep both mail gates off.

For rollback, stop the new consumer/publisher, restore a tested compatible backup, and replay retained post-backup events. If returning to the legacy system, restore a compatible flow/workbook/code set and reconcile every post-cutover application first. Do not overwrite newer work with an old snapshot or run both writers concurrently.

Exit check: no missing intake events, no duplicate canonical applications, no destructive resume cleanup, all source/result versions traceable, report counts reconciled and restoration proven.

## 13. Evaluate the model upgrade independently

After the pipeline is measurable, benchmark the selected local candidate Qwen3.5 9B Q4_K_M against the current model before promotion. The host has 8 GB GPU memory; verify actual runtime memory and the full 106-role prompt rather than inferring fit from download size. Include the smaller 4B candidate only if the measured resource/performance trade-off warrants it.

Use manually checked representative resumes with explicit and unstated experience, decimal totals, bare phone numbers, missing contacts, multi-column layouts and scanned pages. Compare incorrect/invented values, missing-but-present values, JSON failures, geo/ranking regressions, warm/cold latency and timeout rate. Retain the independent second opinion and fix parser defects regardless of model choice.

Exit check: choose from measured P2 evidence. Do not label the recommended model a proven winner or combine its rollout with the first database cutover.
