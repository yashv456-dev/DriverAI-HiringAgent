**P1/P2 end-to-end review — September 14, 2026**

Reviewed `main` at `71c7c75f3ac05a173cfad5643f336800a00faeb9`. This was a review, not a repair or deployment. Production source, configuration, database results, reports, and applicant mail were not changed by this review. Tests used synthetic data; live checks used reads and isolated local scoring.

The normal path works for the two applications currently stored, and all six regression suites pass. Several important failure and update paths are still incomplete. The highest risks are processing an old CV after an update, deleting a newer CV during rename recovery, and losing failure-handling behavior at the boundary between the isolated worker and its parent.

**What was verified live**

| Item | Observed state |
|---|---|
| Backend | SQLite; P2-root `candidates.db` is the effective database. The separate untracked `hiring_agent.db` is not the selected database. |
| Active applications | Two: one `Scored`, one `Needs Review - Location Confirmation`. Two other stored rows are inactive. |
| Intake workbook | The same two active applications and statuses; no rejected applications. |
| P2 master report | Generation 19, two active applications. Its data values matched SQLite; the generated Resume Link display was excluded from the literal comparison. |
| Client report | Generation 19; one scored application. The location-review application is correctly excluded. |
| Publication state | Revision 19 / published revision 19, with no recorded publication error. Unchanged after the review. |
| Current resume files | Both active resume files were downloadable. |
| Model | `qwen3-1.7b-p2:latest` is installed and reachable. Extraction and scoring are enabled; `require_ai=true`. |
| Dependencies | Setup doctor reports core/desktop dependencies, layout extraction, and RapidOCR available. This verifies setup, not every possible document format. |
| JD catalog | 105 cached roles, built September 14 at 01:28:28; configured SharePoint JD folder reachable. Both latest accepted candidate artifacts reference the same JD hash as this cache. |
| Mail switches | P1 package applicant mail off / admin mail on. Effective P2 applicant suppression on / admin-alert setting on. P2's SQLite alert delivery has a gap described below. |
| Timing | Extraction request timeout 180 seconds; scoring 150 seconds; default whole-candidate limit 900 seconds; OCR limit 120 seconds. |
| Launch state | GUI settings have auto-start and scheduling disabled. No P2 Python worker or Excel process was present at the process snapshot. Ollama was running. |
| Mailbox activity | Initial counts: Inbox 37, Archive 342. A later read found Inbox 34, Archive 345. The 34 remaining Inbox messages in that read were from August 2–17. This is evidence of mailbox processing during the review. |
| Power Automate deployment | Could not inspect the installed flow definition, connections, or run history: no browser session was available. Mailbox movement does not establish that the newly built package was imported. |

The later Archive read contained 345 messages: February 1, March 2, April 6, May 175, June 111, July 41, August 4, September 5. These are message counts, not counts of unprocessed applications. Archive contains both handled and failed mail, so its unread count is not a reliable replay-work count.

Two live resumes were then processed in isolated workers. Their remote client was restricted to GET/HEAD, and the worker's mail adapter cannot send. Neither result was committed or published:

- `…MCLA`: completed without errors or deferral; all 13 checked fields matched the stored result.
- `…MCPA`: completed without errors or deferral and correctly remained in location review. Suggested Role 1/2/3 differed from the stored result. The model name and JD hash match the latest accepted context, but the review did not establish byte-identical inference requests or diagnose the variation. A fixed seed should not be described as a guarantee of identical end-to-end results.

**Confirmed findings, in priority order**

1. **High — A replacement CV can be saved successfully while P2 reads the old CV.**

   P1's update action saves into the month of the incoming update email. Its row patch changes the update counter, timestamp, status, and optional mail marker, but does not publish the replacement attachment's filename, folder, or URL. P2 resolves attachments from the original `Received Date` and gives an existing Resume URL filename priority. With sidecars disabled, there is no new attachment manifest to resolve this discrepancy.

   Reproduction: an August application retained its August CV and received a September replacement. `cache_documents` selected the old August bytes even though the new September bytes existed. This is also a risk within one month after P2 has renamed the earlier file: the existing link can still resolve successfully, so fallback discovery is never reached.

   Fix direction: pass an explicit, versioned attachment reference from P1 to P2 on every accepted update, and resolve that input version before accepting the result. Preserve the original submission metadata separately.

   Sources: [P1 replacement upload](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P1/flow/definition.json:2100), [P1 update patch](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P1/flow/definition.json:2139), [P2 attachment resolution](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P2/hiring_agent/local_pipeline.py:189).

2. **High — Rename recovery can delete a newer CV and link the result to older content.**

   When rename fails, `_rename_to_canonical` checks whether the desired target has a web URL. If it does, it selects that target and deletes the incoming source. It does not compare the two files' content hashes. The returned document metadata can therefore retain the new CV's hash while pointing to the old CV's URL.

   Reproduction: target held `OLD CV`, incoming source held `NEW CV`, and rename returned false. The function deleted the new remote source and returned the old target's link with the new source's hash. The local content-hash cache may preserve recovery bytes, but the SharePoint source is deleted and the published link can be wrong.

   Fix direction: verify equal content before treating a target as a duplicate; retain both versions or fail the rename safely when they differ. Avoid remote cleanup until attachment identity and result linkage are verified.

   Source: [rename recovery and delete](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P2/hiring_agent/local_pipeline.py:360).

3. **High — Failed intake write-back is not durably retried.**

   The parent commits the result to SQLite, then attempts to sync Status/Resume Link to P1's workbook. Failure only produces a warning. On the next run, unchanged input preserves the completed SQLite result and the candidate is no longer ordinary queue work; there is no pending-sync record to retry. Rejected-row moves can also leave the remote row on both sheets because the remote operation is add-then-delete, and a false return is not handled as pending work.

   Reproduction: forced a 423 write-back failure. The first run reported one processed candidate and zero errors. The second run processed none and made no second sync attempt. The remote status remained `New Email Received`.

   Fix direction: track remote row synchronization separately from local scoring, retry by application/input version, and report unsynchronized outcomes explicitly.

   Sources: [commit followed by best-effort sync](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P2/hiring_agent/local_pipeline.py:540), [remote sync](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P2/hiring_agent/local_pipeline.py:576), [Excel rejection move](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P2/hiring_agent/store.py:352).

4. **High — Generic processing errors lose their retry-count changes, and the batch circuit breaker does not span workers.**

   A child can return `errors > 0` with an incremented Retry Count in its private store. The parent discards those returned rows and marks only the attempt deferred. Consequently a recurring generic processing exception can start from the same Retry Count on every run. The child sees one candidate, so its three-consecutive-failure breaker cannot protect the parent batch.

   Reproduction: four synthetic candidates each returned a processing error. All four were attempted despite the configured breaker of three. After a second run there were eight deferred attempts, but every persisted Retry Count remained zero.

   This does **not** mean every retry is broken: unreadable-text review outcomes without `errors` are accepted and can retain their counters. Missing-file and infrastructure timeouts are intentionally deferred without penalizing the applicant. The problem is the absent parent-level distinction and counter/circuit handling for generic failures.

   Fix direction: return typed outcomes, persist intended failure-state changes, preserve infrastructure deferral, and implement the consecutive-failure limit in the parent batch.

   Sources: [parent outcome handling](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P2/hiring_agent/local_pipeline.py:505), [child exception handling](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P2/hiring_agent/sharepoint_scoring.py:3591), [attempt persistence](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P2/hiring_agent/sqlite_store.py:304).

5. **High — A valid “no suitable match” answer is classified as an AI outage.**

   `suggested_roles` filters scores below 20% and drops unsupported matches. If nothing remains, the strict-mode branch returns `source='unavailable'` and “AI role scoring failed or timed out,” even when Ollama answered successfully. The orchestrator then defers the applicant before reaching the final geography/status decision.

   Reproduction: a valid model result rating the only role at 0% returned the outage result. An applicant with no suitable opening can therefore be repeatedly deferred instead of receiving a deliberate no-match outcome.

   Fix direction: distinguish a successful ranking with no eligible matches from failed inference. Define the intended status/category and publication behavior for a legitimate no-match applicant.

   Sources: [ranking filters](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P2/hiring_agent/scoring.py:466), [strict fallback result](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P2/hiring_agent/scoring.py:552), [orchestrator source gate](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P2/hiring_agent/sharepoint_scoring.py:3316).

6. **Medium — SQLite skips the normal mail and age-out passes, including real admin delivery from children.**

   The isolated adapter's `send_mail` always returns false. The child returns before the normal post-batch decline, information-request, age-out, and cross-candidate maintenance passes. The parent supplies publication and gap recovery but does not replay these mail/age-out operations through the real client.

   The synthetic pipeline check recorded zero calls to all three post-batch mail/age-out functions. Source inspection also confirms that child suspicious-content and give-up alerts reach the non-sending adapter. Thus `ERROR_EMAIL_ENABLED=true` is not sufficient to deliver these alerts in the current SQLite route. Changing applicant suppression to false alone will not enable the omitted post-batch mail passes.

   Applicant suppression is intentional now, so the absence of applicant emails is expected. The missing admin delivery and backend differences still need an explicit contract. The configured 14-day age-out also does not run through this route; suppressed contacts should not start its clock in any case.

   Fix direction: persist notification/maintenance work in the parent with separate applicant/admin policies and idempotent delivery state.

   Sources: [local mail adapter](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P2/hiring_agent/local_pipeline.py:82), [early child return](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P2/hiring_agent/sharepoint_scoring.py:3639), [admin alert implementation](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P2/hiring_agent/sharepoint_scoring.py:1508).

7. **Medium — Publication can partially succeed while the command still appears successful; verification checks shape rather than content.**

   The two reports are uploaded sequentially. If the master upload succeeds and the client upload fails, they can temporarily expose different generations. The database correctly retains a pending publication error, and a later invocation can retry without inference. However, `publish_results` returns a local filename and the run summary can still show zero errors. The CLI does not convert that state into a nonzero exit, so the launcher can print success.

   Separately, `_verify_report` accepts different bytes when worksheet names, row counts, and column counts match. It does not verify cell values or the generation marker.

   Reproductions: a locked second upload left only the master delivered, `published_revision=-1`, and `error='workbook locked'`, while the run reported one processed candidate and zero errors. A workbook containing different cell text but the same shape passed `_verify_report`.

   Fix direction: expose publication status in the result/exit code, verify a semantic content fingerprint and generation, and make a mixed-generation state visible to operators/readers. Keep the existing retry-without-rescoring behavior.

   Sources: [verification](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P2/hiring_agent/publisher.py:40), [publication](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P2/hiring_agent/publisher.py:136), [CLI scoring branch](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P2/bot.py:431).

8. **Medium — A follow-up email's new information is not handed to P2.**

   P1's follow-up patches record the counter, timestamp, and optional mail marker. They do not write the new Mail Body or a separate reply payload. SQLite notices the changed counter/timestamp and can requeue the application, but receives the original mail text and original resume again. A reply such as “I now live in Phoenix” is therefore not a reliable automatic location-correction path.

   With applicant emails suppressed, manual follow-up remains necessary for `…MCPA`, and the answer must also be persisted into an input path P2 actually consumes; sending or receiving an email alone does not establish that.

   Fix direction: capture follow-up content as a separate dated input event and make current-location clarification consumption explicit.

   Sources: [P1 follow-up patch](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P1/flow/definition.json:2736), [SQLite input fields](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P2/hiring_agent/sqlite_store.py:11).

9. **Medium — SQLite maintenance commands have broader effects than their help/docstrings promise.**

   `recover_location_rejections` delegates to `run_local_pipeline(..., force=True)`, which performs full scoring of active main/rejected rows up to the batch limit. It does not restrict itself to historical geography rejections. `recheck_all_rows` also forces full scoring regardless of `force_roles`; selected-row recheck does not honor the documented `geo_only` distinction in its SQLite branch.

   This is established by the routing code; no live recovery command was executed. An operator following the “no role rescoring” description can unintentionally change good rows and their roles/categories.

   Fix direction: retain the requested operation and flags across backend dispatch, or expose separate commands whose scope matches what they do.

   Sources: [all-row recheck dispatch](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P2/hiring_agent/sharepoint_scoring.py:4240), [location-recovery dispatch](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P2/hiring_agent/sharepoint_scoring.py:4505), [selected-row dispatch](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P2/hiring_agent/sharepoint_scoring.py:4589).

10. **Medium — P1 duplicate-lookup failure still permits a new application row.**

    `Filter_submitted` runs after a failed lookup and coalesces its absent output to an empty list. That makes ordinary intake take the new-applicant branch. The new alert makes this visible, but does not prevent a duplicate. The package's alert text explicitly acknowledges this behavior. A failed reference lookup can also fall back to ordinary intake.

    P2's current isolated path does not run the old batch-wide email/phone duplicate merge. SQLite protects unique active Application IDs, which does not eliminate two different IDs for one person.

    Fix direction: distinguish lookup failure from a successful “not found”; defer and retain the message for retry when identity lookup is unavailable. Treat replay idempotency and cross-ID identity reconciliation as separate concerns.

    Sources: [lookup failure fallback](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P1/flow/definition.json:1455), [reference lookup](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P1/flow/definition.json:1992), [child maintenance bypass](C:/Users/yash9/Downloads/HiringAgent2/DriverAI_HiringAgent/HiringAgent_P2/hiring_agent/sharepoint_scoring.py:3639).

**P1 path map: what the committed package does**

The generated definition contains 153 actions, including 17 conditions and three attachment loops. There are 25 actual admin SendEmailV2 actions with the current switches; the six applicant replies are Compose no-ops. Older comments/readmes saying nine or fourteen admin alerts are stale. The ZIP/loose-definition contract passes the P1 suite; deployment equivalence remains unverified.

| Trigger or branch | Normal behavior | Failure/fallback and limits |
|---|---|---|
| Scheduled intake | Recurrence every minute, one run at a time, fetch one Inbox message with attachments included. Both read and unread messages qualify. | Poll failure has an admin alert. Empty result terminates successfully. This is an Inbox queue, not an unread-only queue. |
| Sender allowlist / system screen | Allowlist uses exact normalized sender equality. Other sender screens use configured substrings. | Sender allowlisting bypasses spam screens. Review these addresses as privileged routing configuration. |
| Subject screen | Blocks configured subjects before intake. | Routed to Junk; move failure alerts. |
| Body/attachment-name spam screen | Checks subject plus body and attachment names. | Routed to Junk; move failure alerts. Substring rules can false-positive and do not inspect attachment text. |
| Shortened links | Suspicious shortener domains are blocked when there is no qualifying resume attachment. | A qualifying resume bypasses this particular shortener screen, not every other screen. |
| Resume eligibility | PDF, DOCX, DOC with nonempty attachment content. Non-resume filename exclusions are applied. | If exclusions would remove every eligible file, all eligible files are restored. Thus a sole portfolio/cover letter may still be treated as a resume. Multiple retained files can be processed; this is not a guaranteed single-CV semantic classifier. |
| Reference reply | Normal senders resolve their latest matching row; allowlisted routing can use the quoted reference. | This is not always an exact quoted-application lookup for ordinary senders. Lookup failure can fall back to ordinary intake. |
| New application | Save resume, add row, then acknowledgment/no-op; stamp mail state. | Save failure blocks downstream row/ack work and alerts. Add-row failure can leave an orphan file and alerts. Stamp failure also alerts. Recovery is manual replay. |
| Duplicate with resume | Match submitted rows using sender/reference and the 90-day expression, update the existing application under its counter cap. | Counter below 3 permits an applicant reply when enabled; contacts under 5 still accept updates. At cap, processing goes silent. Replacement attachment handoff has finding 1. |
| Referenced resume update | Save replacement, increment counter/timestamp, requeue existing row. | Same reply and overall caps; save and patch failures alert. Replacement file identity/folder is not handed off adequately. |
| Follow-up without resume | Record contact count/timestamp and optional reply. | No reply text stored for extraction; finding 8. At cap, silently archived. |
| Application keywords, no attachment | Resume-request reply/no-op; no candidate row. | No persistent row means no per-row cap for repeated no-attachment requests. With mail off, this produces neither candidate row nor applicant contact. |
| Attachments, wrong format | Format-request reply/no-op; no row. | Inline/signature attachments can influence attachment-presence branches. |
| Clean nonapplication | Admin ignored-mail notice; no row. | Can be archived with no applicant contact, as intended by current mail switches. |
| Terminal cleanup | Normal handling marks read then archives; spam moves to Junk. | Move/mark retries are disabled; move failures can leave the message eligible for the next poll. Several processing failures are deliberately finalized as Succeeded after alert/cleanup, so success in run history is not proof of successful intake. |
| Failed-message recovery | Save/core failure paths move to Archive without the normal mark-read step. | “Unread failure” means the flag is preserved, not necessarily forced to unread. A message already opened before processing can remain read. Some patch/add failures are deliberately archived read; their alert is the principal signal. |
| Workbook/folders | P1 will not recreate/overwrite a missing candidate workbook. Dated folders can be created. | Human restoration is required for a genuinely missing workbook. |
| Replay | Moving archived mail back to Inbox makes it eligible regardless of read flag. | Can reapply counters, create files, or update a current application. The duplicate date comparison has a lower bound but no corresponding upper bound; replay order should not be treated as protection against stale updates. No replay was performed in this review. |

**P2 path map: current SQLite route and its fallbacks**

| Stage | Current behavior | Failure/fallback and limits |
|---|---|---|
| Entry/configuration | CLI, GUI-launched CLI, and PowerShell launcher reach `score_from_sharepoint`; SQLite dispatch occurs before the legacy Excel workflow. Environment overrides config. | Invalid backend names are rejected. Unset backend selects Excel, so a fresh deployment can differ from this machine. |
| Worker ownership | OS file lock keyed to the resolved database path serializes local pipeline owners. | Protects workers sharing that lock/path; it is not a tenant-wide lease across separate machines/databases. |
| Import | Read main/rejected remote rows before the transaction; keep source snapshots, compare input hashes, preserve historical results on new input. | Changed inputs requeue. Identity/same-time conflicts are retained and excluded from processing. Missing remote application IDs deactivate local rows; the remote snapshot is authoritative for membership. |
| Sidecar import | Code exists but P1 sidecars are off. | Reader checks `folder` while listing returns `is_folder`; dated subdirectories are not traversed. This remains broken and should stay disabled until repaired. |
| Queue | New rows, unreadable review, and qualifying location updates; batch default 500. | No explicit chronological sort in the parent; ordering is database row order. Persistent generic failures can keep occupying a bounded batch because parent backoff/breaker handling is missing. |
| Resume resolution | Manifest if supplied; otherwise guessed names plus URL filename, original-date folder, opposite main/rejected bucket, legacy flat folder, then application-reference search. | Missing required attachment defers before OCR. Ambiguous reference-tail matches may be selected by filename preference or first match, rather than failing closed. See stale-update finding. |
| Local cache | Content-addressed copies; worker verifies cache bytes against SHA-256. | Protects cached-file integrity. It does not prove the selected remote CV was the newest/intended version. |
| Isolation | One spawned process per candidate, default 15-minute deadline. | Timeout terminates the candidate process and retains prior durable state. Per-request bounds are not the total candidate bound. |
| PDF text | PyMuPDF4LLM reading-order extraction, then pypdf if empty/failing. | Low-quality or short text triggers OCR; if still empty, pdfplumber is attempted. Actual OCR order differs from the README's “all text parsers before OCR” description. |
| OCR | RapidOCR primary; Tesseract fallback; separate default 120-second process budget. | Timeout propagates as infrastructure deferral. No useful text follows unreadable-review processing. The later pdfplumber fallback is not subjected to the same text-quality test again. |
| Word | DOCX parser; `.doc` routes by actual ZIP/RTF/OLE magic bytes. | Unsupported/corrupt contents return no text. Legacy Word fixture is covered by regression checks. |
| Content safety | Secondary suspicious-phrase check before candidate extraction/scoring. | Marks possible spam. Intended admin alert cannot send through the isolated adapter. |
| Candidate fields | Deterministic baseline, then model-assisted extraction/recheck. Phone, links, education dates, experience and skill normalization have deterministic paths. | Extraction failure is not permitted to become a published regex-only row in the online orchestrator, even when `require_ai` is false. Helpers and orchestrator therefore have different fallback contracts. |
| Extraction retries | Request timeout gets one retry; initial field extraction separately retries malformed JSON once. | Mixed timeout/JSON sequences can make four requests. Recheck/role-note/portfolio helpers do not all share the initial extraction's JSON retry. Twelve transport tests now cover the distinction. |
| Independent recheck | Rechecks fields and merges corrections; portfolio inference can run concurrently. | With REQUIRE_AI or STRICT_AI_STAGES, recheck failure raises. Parent currently discards the child's generic retry-counter changes. |
| Role catalog | Current cached set is frozen for the worker's main scoring calls. Missing/invalid/stale cache can trigger a live scan. | Failed scans retain usable stale cache; with no usable cache, built-in roles can be used. “AI available” does not establish that the openings are current. |
| Role ranking | Shortlist up to 12 roles, model ranking, minimum 20%, overlap guard, opening deduplication, keyword top-up, descending ordering. Category includes domain/QA guards. | Some published percentages can be keyword top-ups even when overall source is `ollama`; they are match heuristics. Zero surviving matches are currently misclassified as unavailable. |
| Geography | Confirmed US may score; unknown/conflicting location stays in review; corroborated non-US can reject. | Country/phone/university alone do not establish current US residence. Ambiguous places can use Nominatim; uncertainty is retained. Current `…MCPA` review behavior reproduced. |
| Commit | Version-checked attempt completion and result history in SQLite; active Application ID uniqueness; sheet move is local transactional state. | The local version check does not itself provide a distributed lease over P1 changes. New remote input is reconciled on the next import. |
| Remote intake sync | Main rows get Status/Resume Link only. Rejection moves copy a full row while restoring P1-owned inputs. | Failed sync is only logged, not queued durably. Remote rejected moves are not atomic. |
| Publication | Generate master + scored-only client workbook from committed rows, local generation directories, guarded targets that do not replace P1 intake. | Remote retries can occur on later runs without rescoring. Partial delivery and weak content verification are finding 7. |
| Idle recovery | With no queued work, reread up to five scored rows with recoverable blanks; track input hash to avoid repeating unfillable gaps forever. | Recovery is stamped even when the attempt fails, so a transient failure can suppress another automatic attempt until input changes. |
| Mail/age-out/dedup | Legacy Excel path contains those batch-wide passes. | Current SQLite route omits their equivalent parent phase. Application-ID uniqueness does not replace cross-ID email/phone reconciliation. |
| Dry run | Clone SQLite state into memory; separate preview output; no intended remote mutation. | It still opens/initializes the persistent database, takes a lock, can fill the local resume cache, and writes local preview files. It is remote-safe preview behavior, not literally zero filesystem writes. |

**Other entry points and operational boundaries**

- `--export-results` on SQLite imports current intake state and publishes reports without candidate inference. It can therefore change local imported state; its old “read-only” wording is too broad.
- `--recheck-all`, `--recheck-row`, and location recovery must be interpreted using finding 9, not their legacy descriptions.
- Folder/file/Excel scoring helpers still exist, with different persistence and fallback rules. They are not the active P1-to-P2 queue on this machine and were not used on live data in this review.
- Azure Functions has timer and HTTP wrappers. Timer execution is gated by settings; its serverless/regex documentation conflicts with the online orchestrator's unconditional model-provenance gates. No deployed Function App or Docker runtime was certified here.
- Graph reads retry transient transport errors; mutating transport exceptions are not automatically retried. HTTP 429/503/504 responses are retried, including mutating requests. An uncertain append therefore still needs reconciliation rather than a blanket exactly-once assumption.
- The existing safe-copy `move_resume` helper compares destination bytes and preserves the source. That safety is undermined by the separate rename-recovery deletion in finding 2; comments claiming system-wide no deletion are inaccurate.
- Backup, replay, initialization, workbook repair, and migration utilities were not executed. They can mutate or reset live/local state and are separate from a scoring review.

**Validation evidence and limits**

| Suite/check | Result |
|---|---|
| `test_p1.py` | 947 passed, 0 failed; package structure/contract and routing simulation. |
| `test_p2.py` | 1,450 passed, 0 failed in this session; same production source as reviewed. |
| `test_store.py` | 68 passed, 0 failed. |
| `test_local_pipeline.py` | 20 unittest methods passed. |
| `test_resume_availability.py` | 12 unittest methods passed. |
| `test_extraction_transport.py` | 12 unittest methods passed after the earlier authorized test update. |
| Setup doctor | Core/desktop/layout/OCR/model/config/JD-folder checks healthy. |
| Live reads | Intake/report/database comparison, current resume access, model presence, mailbox counts. |
| Isolated real-model checks | Two applications completed with expected statuses; one reproduced all checked values, one changed three role fields. |
| Synthetic fault cases | Reproduced stale attachment selection, unsafe rename collision, missing sync retry, lost generic retry counters, absent parent breaker, invalid no-match classification, partial publication reporting, weak report verification, and omitted post-batch hooks. |

Passing suites establish their covered assertions; they do not certify every Power Automate connector outcome, concurrent P1/P2 update, attachment ambiguity, model answer, or tenant failure. The synthetic faults demonstrate specific uncovered behavior without altering production.

**Recommended repair order**

1. Fix attachment version handoff and collision handling before broad archive replay or new-CV update testing.
2. Add durable remote synchronization and typed parent-level failure handling, including retry counters and the circuit breaker.
3. Separate valid no-match outcomes from inference failure; make publication failures visible to commands/operators.
4. Restore the intended admin/mail/age-out behavior and narrow maintenance commands to their advertised scope.
5. Close P1 lookup-failure and follow-up handoff gaps, then validate the imported flow with explicit new/update/failure scenarios.

No fixes, flow import, replay, applicant contact, or report publication were performed as part of this review.
