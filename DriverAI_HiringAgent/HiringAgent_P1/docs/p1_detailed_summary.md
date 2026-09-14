# Phase 1 (Intake Flow) — Detailed Architectural Summary

The front door of the hiring pipeline is **Phase 1** (`HiringAgent_P1`), a serverless Power Automate flow (`DriverAI_HiringAgent_P1_AutoReply`). It watches the shared inbox, filters out spam/system mail, saves incoming resumes, and creates candidate rows in SharePoint. Applicant-facing mail is configuration-controlled and is currently disabled for silent live intake. It never reads a resume's content and never runs a model — that's entirely Phase 2's job.

---

## 1. Scheduled Unread-Inbox Poll

* **Trigger action:** `Recurrence`, followed by Outlook `Get emails (V3)`.
* **Inbox monitored:** `apply@driverai.io` (root `Inbox` folder only).
* **Poll interval:** **1 minute** (`trigger.interval_min` in `flow_config.json`; changes take effect after rebuilding and re-importing the zip). History: reduced to 2 minutes on 2026-08-04 for quota headroom, then returned to 1 minute on 2026-08-14 on owner instruction. A 1-minute poll is 1,440 runs/day against the owner's plan quota (2 minutes was 720/day); on a free/seeded Power Automate licence the 24 h Power Platform request budget is the binding constraint, so watch for throttling under load.
* **Selection:** every message in the Inbox, **read or unread**, attachments included, top **1** (`trigger.unread_per_run`). The **Inbox itself is the queue** — every terminal path moves the message out of it (Archive on success, Junk Email on spam, Archive-unread on a hard failure), so "still in Inbox" is the accurate definition of "not yet handled". Changed 2026-08-04 (`trigger.fetch_only_unread: false`): the old unread-only filter meant that simply **opening** a message in `apply@driverai.io` removed it from the queue permanently, with no error and no alert, while the flow kept reporting Succeeded — a real case (a real applicant, `redacted-applicant@example.edu`: clean application, PDF attached, zero spam matches, no reply and no row). This handles new mail and historical mail moved back to Inbox; at the 1-minute interval, 41 messages require about 45 minutes plus runtime overhead.
* **Concurrency:** strictly **1**. Every branch reads-then-writes the same Excel row (`Get_rows` → `Add_row`/`Patch_*`); raising this risks duplicate rows under burst load.
* **Completion:** legitimate mail is marked read and archived, spam moves unread to Junk Email, and a failed item triggers an admin alert then moves unread to `Archive` so the next unread item can proceed. Success and failure moves are mutually exclusive: the failure move runs only after an actual `Notify_failure` attempt and never when that action is skipped on success. This removes the prior double-move race against one mutable Outlook message ID. Terminal `MoveV2` retries are disabled and cleanup results are explicitly finalized. The well-known Archive folder is intentional: Outlook `MoveV2` rejected the valid nested Recruiting Review Graph ID after import.
* **Import-safe hierarchy:** `Has_unread_email`, `CurrentEmail`, and the legacy processing graph are flat top-level siblings. The no-mail branch terminates the scheduled run successfully. This preserves the original deepest actions at Power Automate level 8; wrapping the graph would create invalid level-9 actions and fail package import. Builder and package tests enforce the level-8 maximum.

---

## 1A. Year-Boundary Separator Row & Test Mode

* **Year-boundary separator row** (`flow_config.json: year_separator.enabled`, default `true`): before writing the first candidate row of a new calendar year, `build_zip.py` performs one paginated `Get_rows_for_separators` scan (up to 5000 rows), then `Filter_rows_this_year` compares the returned ISO `Received Date` values in memory before `IsFirstOfNewYear` optionally inserts one blank row. This avoids Excel OData entirely: Microsoft documents Filter Query as supporting only alphanumeric column names, and the prior `Received Date ge ...` query failed live at the space. The prior-applicant guard reuses the template's permanent opening spacer on a fresh workbook.

* **Month-boundary separator row** (`flow_config.json: month_separator.enabled`, default `true`, added 2026-07-31): `Filter_rows_this_month` reuses the same paginated scan and checks the `yyyy-MM` prefix in memory. The prior-applicant guard keeps the permanent header spacer as the only blank before the first-ever candidate. At a January boundary the month condition also requires at least one row already in the incoming year, so the year separator wins and exactly one blank row lands. `Add_row` proceeds after both checks even if either is skipped or fails, and PATCH-only paths never reference this feature.
* **Outbound mail controls:** `email.send_applicant_emails=false` replaces all 6 applicant-facing sends with no-op `Compose` actions while preserving their wiring, and writes `"Suppressed (applicant email disabled) <timestamp>"` instead of a false sent stamp. `email.send_admin_failure_alerts=true` keeps all **11** admin alerts live — `Notify_failure`, `Notify_poll_failure`, the 7 patch-failure alerts, and the 2 intake-write alerts (`Add_row`, `Create_file`) — so operational failures remain visible. Triggering, filtering, SharePoint/Excel writes, and inbox tidy are unaffected.

---

## 2. What P1 Fills In — The 11-Column Contract

The live table has **33 columns** (27 at the 2026-07-04 baseline; grew to 29 on 2026-07-14 when P2 added `Education` and `Info Request Sent`; grew to 30 on 2026-07-15 when P1 added `Last Updated Date`; grew to 33 on 2026-08-04 when P2 added `Years Exp`, `Education Start Date`, and `Education End Date`). P1 writes **only 11** of them at row creation; the other 22 land as blank cells for P2 scoring/resume metadata plus later audit stamps.

| Column | Value/Expression | Purpose |
|:---|:---|:---|
| **Application ID** | `@outputs('AppRef')` | Unique reference key, e.g. `APP-20260710-2200-A1B2` |
| **Received Date** | `formatDateTime(receivedDateTime, 'yyyy-MM-ddTHH:mm:ss')` | ISO-8601, no timezone offset |
| **Last Updated Date** | `formatDateTime(receivedDateTime, 'yyyy-MM-ddTHH:mm:ss')` | Same as Received Date on first intake; refreshed on every patched contact (resume update/resend AND text-only follow-up, fixed 2026-07-31 - follow-up used to be silently skipped). Guarded to never move backward: `max(existing value, this email's own time)`, since backlog replay can process an older queued reply after a newer one already patched the row. |
| **Full Name** | Heuristic guess | Sender display name if present, else cleaned email local-part |
| **Email** | `from` | Raw sender address (see §9 for a known matching caveat) |
| **Mail Subject** | `subject` | — |
| **Mail Body** | `body/bodyPreview` | Outlook's plain-text preview, ~255 chars (not the full HTML body, no 30KB cap) |
| **Status** | `"New Email Received"` | Signals P2 the row is ready to score |
| **Has Resume** | `"Yes"` | Invariant — a row only exists if a real CV was attached |
| **Original Filename** | `item()?['name']` | As sent by the candidate, joined with `, ` if multiple |
| **Application Updates** | `0` | Shared counter driving every cap (see §7) |

Deliberately **not** listed in the intake write: Category, Phone, Location, Country, Years Exp, Current Skills, Education, Education Start Date, Education End Date, Looking For Role, Suggested Role 1-3, Portfolio 1-3, Retry Count, Resume URL, Resume Link, Resume Folder Path, Info Request Sent, Mail Sent. Writing only the 11 real values (rather than all 33 with 22 empty strings) minimizes the import-time schema-binding surface — listing every column forces Power Automate to resolve that many column-name matches against the live table at import time, and any one mismatch can silently drop a field in the designer.

---

## 3. Application Reference (AppRef) — now config-driven

Format: **`APP-<yyyyMMdd>-<HHmm>-<4hex>`** (22 chars, minute precision), e.g. `APP-20260710-2200-A5F2`.

As of this session, the shape is no longer hardcoded in `build_zip.py` — `flow_config.json`'s `appref` block carries it directly:

```json
"appref": {
  "date_format": "yyyyMMdd",
  "time_format": "HHmm",
  "hex_length": 4,
  "detect_pattern": "app-20"
}
```

`_ref_len` (used to extract a quoted ref out of a reply email) is computed from these same two values at build time, so minting and extraction can never drift apart even if the shape changes later — change `time_format`/`hex_length`, rebuild, re-import, and both sides move together automatically. Rebuilt and reverified this session: `definition.json` came out **byte-identical** to the pre-change file, confirming this was a pure config-plumbing change with zero behavior drift (defaults match what was previously hardcoded).

`detect_pattern` (`"app-20"`) is the case-insensitive substring that flags a reply as quoting a reference — decade-proof through 2099, matched anywhere in subject or body, deliberately **not** added to `bad_subjects` (that would drop an applicant's own `Re: <our subject>` reply before the ref gate ever sees it).

---

## 4. Resume Storage

* **Root:** `Candidate_Resumes/`
* **Dated path:** `Candidate_Resumes/<Year>/<Month>/` (e.g. `2026/July/`) — weekly subfolders were deprecated and removed.
* **Auto-creation:** `CreateNewFolder` (`Ensure_*_resume_folder`) runs immediately before each save; safe no-op if the folder already exists. No artificial delays — testing showed folder→file save works reliably without one.
* **Saved filename at P1 intake:** `<FirstLast>_<AppRef>.<ext>` (e.g. `JaneDoe_APP-20260710-2200-A5F2.pdf`) — changed 2026-07-12 from the older `<AppRef>_<original>` shape. `FirstLast` strips apostrophe/hyphen/period/comma from the candidate's name and keeps first+last word (or `Candidate` if the name is empty), mirroring P2's own `_get_cleaned_filename_prefix()` exactly.
* **Renamed again by P2 once scored** (2026-08-04, client request): `<First>_<Last>_<tail>.<ext>` (e.g. `Jane_Doe_A5F2.pdf`) — P1 cannot produce this name itself because the candidate's real name comes from the CV, which only P2 reads. `<tail>` is the AppRef's own last hyphen-segment, kept so two same-name candidates never collide now that the full AppRef is no longer visible in the name. **Category used to be part of this name** (`<FirstLast>_<Category>_<tail>`, 2026-07-15) and was removed on 2026-08-04: Category is a P2 verdict, so a re-score that moved a candidate between categories forced a physical rename, and any file whose Category had since changed became unfindable by name. A one-time bulk migration (`HiringAgent_P2/migrate_resume_filenames.py`) renamed every already-scored resume to the current format.
* **Workbook:** single file in `Master_Files/`, no `<Year>` segment - one workbook covers every year.
* **Self-heal:** **off, on purpose** (removed 2026-07-04). The old metadata probe produced false "not found" results and re-uploaded a blank template over real data, wiping rows. Now: a genuinely missing workbook fails the run and pages the admin via `Notify_failure`, and a human restores it once from `P1_Templates/HiringAgent_P1_CandidateList.xlsx` — never silently replaced.

---

## 5. Email Routing Logic & Spam Gates

Three consecutive, case-insensitive substring gates run before anything else:

* **Gate 1 — Bad Sender (31):** notification/system bots, marketing platforms (LinkedIn, Slack, Zoom, Calendly, GitHub, DocuSign, Mailchimp, SendGrid, Salesforce, Facebook, Dropbox, Notion, Asana...), and the self-loop guard (`apply@driverai.io` itself, plus mailer-daemon/postmaster/bounce). Sender-based, not subject-based — deliberately, so an applicant's `Re:` reply is never caught by the loop guard. Matching is an unanchored `contains` over the raw From header, so every term must be safe as a substring: `marketing@` was removed 2026-07-31 after it matched `redactedmarketing@example.com`, a real CMO applicant junked three times.
* **Gate 2 — Bad Subject (24):** invoices, password resets, out-of-office, delivery failures, and three retired outbound subjects (so a bounce of P1's *own* old emails doesn't loop back in). Narrowed 2026-07-31: the bare tokens `alert`/`invoice`/`receipt`/`payment`/`survey` also matched real job titles ("Alert Systems Engineer", "Payments Platform", "Land Surveyor") and are now specific phrases.
* **Gate 3 — Content (123 = 30 scam + 36 offensive + 20 malware-extension/phrase + 25 non-English scam + 12 vendor-pitch):** scans subject + body + attachment names. Offensive-phrase list deliberately avoids bare high-false-positive tokens (no standalone `sex`/`ass`/`hell` — only full phrases like `sex chat`). Executable extensions are stored **space-terminated** (`".exe "`) and `LowerBody` appends a trailing space, so a real attachment matches while `www.exeter.ac.uk` and `www.iso.org` no longer read as malware.
* **Gate 3b — Link shorteners (12, conditional):** only blocks when there's *no* valid PDF/DOCX attached — a resume with a shortened LinkedIn icon in the signature is common and legitimate, so a real attachment overrides this check.

All three gates route the same way on a match: move to **Junk Email**, leave **unread**, terminate silently — no reply, no row, no trace in Archive.

---

## 6. Work Branch Scenarios

| # | Incoming email | Row written? | Resume saved? | Reply |
|---|---|---|---|---|
| A | New applicant, PDF/DOCX, not a dup | **Yes** — the only row-writer in the flow | Yes, into `<Year>/<Month>/` | **Email 1** (acknowledgment) |
| B | Attachment present, wrong format (`.jpg`, `.zip`, ...) | No | No | **Email 4** (please resend as PDF/Word) |
| C | No attachment, has application keywords (`resume`, `cv`, `apply`, ...) | **No — CV-only table invariant** | No | **Email 3** (please attach CV) |
| D | No attachment, no keywords | No | No | none |

All four tidy the same way on completion: mark **read**, move to **Archive**.

---

## 7. Repeated Applicants — Capping & Overwriting

Matched against existing rows by sender email. Two entry points, sharing **one** `Application Updates` counter and a two-tier cap:

* **`duplicate_check_days = 90`** — the resubmit window (inclusive at exactly 90 days).
* **`*_max = 5`** — the hard ceiling per path (`duplicate_notice_max`, `update_resume_max`, `followup_reply_max`); the row goes fully silent forever (within the window) once hit.
* **`reply_cap = 3`** — a *smaller*, separate threshold. Contacts 1-3 still get an actual reply email (the 3rd is a final-notice variant). Contacts 4-5 still update the row/resume and re-queue it for P2 — but get **no email at all**. This two-tier design means the file and the SharePoint row both stay current even after the applicant stops hearing back, so P2 always scores the latest resume regardless of email noise.

| Scenario | No ref quoted (resend) | Ref quoted (`APP-20...` in reply) |
|---|---|---|
| Resume attached, <3 contacts | Email 2 (duplicate notice), resume saved under P1's standard name in the original folder¹ | Email 5 (update ack), same save¹ |
| Resume attached, 4-5 contacts | Silent — file/row still updated | Silent — file/row still updated |
| No resume, ref quoted, <3 contacts | n/a | Email 6 (noted reply) |
| Any path, ≥5 contacts | Fully silent, nothing saved | Fully silent, nothing saved |

Gate ordering matters: `HasAppRefGate` runs **before** `HasResume`, so a ref-reply with a resume attached is always handled as an update (never double-counted as a brand-new applicant).

¹ "Save" not "overwrite" is deliberate wording (corrected 2026-07-15): P1 always computes the same `<FirstLast>_<FullAppID>` target, which is a true in-place overwrite *only* if P2 hasn't scored the row yet. Once P2 has scored it, the file's been renamed to `<First>_<Last>_<tail>` — a shape P1 never writes (it is derived from the CV name, which only P2 reads) — so this save lands as a fresh file alongside the stale renamed one. P2's next scoring pass (`_download_resume_text`, fixed 2026-07-15) always prefers that fresh file over the stale one and deletes the leftover, so the candidate still ends up with exactly one resume — just resolved a run later by P2, not atomically by this save. See `HiringAgent_P2/docs/p2_detailed_summary.md` §2 Step 1.

---

## 8. Error Handling and Recovery Tooling

* **`Notify_failure`** and **`Notify_poll_failure`** keep admin alerts live for core candidate-processing and Inbox-read failures. Applicant-facing sends remain disabled. `Notify_failure` watches `HasResume` specifically — not every action in the flow.
* **7 patch-failure alerts (added 2026-08-03)** close the gap that left. Inbox cleanup is unconditional by design (every action after a candidate-row `PatchItem` lists `Failed` in its `runAfter`, so a poisoned message can never block the unread queue), which meant a failed patch was completely silent: the enclosing `If` reported `Succeeded`, the branch ended on a `Terminate` with `runStatus: Succeeded`, and **the run went green** while the candidate's `Application Updates` counter had not incremented — and on the dup/update paths, while the row had not been re-queued for Phase 2. One sibling alert now watches each of `Patch_mail_sent`, `Patch_dup_attempts(_noreply)`, `Patch_update_attempts(_noreply)`, and `Patch_followup_attempts(_noreply)`, each wired on `[Failed, TimedOut]` only. They are **siblings, not children**: `Patch_update_attempts` and `Patch_followup_attempts` already sit at depth 8, the Power Automate nesting ceiling, so a child action there would fail import. Nothing depends on them, so they cannot alter any existing path. **Six of the seven are genuinely silent failures where the alert is the only signal; `Patch_dup_attempts_noreply` is the exception** — it is the last action in its scope, so its failure propagates to `HasResume`, turns the run red, and fires `Notify_failure` as well. That alert is a deliberate duplicate and its body says so; it is retained because the reply-path twin *is* swallowed, so appending any sibling to that scope would start hiding the failure. `test_p1.py` asserts the swallow/propagate map against the live wiring so the alert wording cannot drift out of truth.

* **2 intake-write alerts (added 2026-08-24)** close the remaining hole. The seven above watch every PATCH of an *existing* row; the two writes that **create** a candidate — `Add_row` and `Create_file` — had no watcher at all, and their failure is strictly worse (a failed patch leaves a stale row; a failed `Add_row` leaves **no row anywhere**). It was silent because `Mark_as_read`/`Move_to_processed` hang off `HasResume`, **not** off `Add_row`, so cleanup sits on a parallel branch and files the mail as handled regardless, while `Send_acknowledgment` — correctly gated on `Add_row` succeeding — is skipped. Net effect: no row, no reply, mail archived, run green, silence. The 2026-08-24 mailbox audit found **11 applicants lost exactly this way**, clustered 2–6 June 2026. The cleanup is deliberately *not* re-gated (that would re-introduce the poisoned-message queue block, and re-marking unread races the mutable Outlook message ID), so the alert is the signal — and its body carries the recovery step: move the message from Archive back to the Inbox as unread and the next poll re-runs intake. `test_p1.py` enumerates every `Notify_*` watcher and asserts `Add_row` and `Create_file` are covered, so a rename or rewire fails the suite.
* **Admin alerts retry (changed 2026-08-03).** All nine now use `exponential ×2` instead of `retryPolicy: none`. The no-retry rule protects `MoveV2`/`MarkAsRead_V3`, where a replay races the mutable Outlook message ID; a send has no such hazard, and with applicant mail suppressed these are the only live sends in the flow — so a transient blip must not lose the only signal.
* **Failure quarantine:** after a message-processing failure and an actual admin-alert attempt, P1 makes a best-effort move of the message unread to `Archive`. The route cannot activate merely because `Notify_failure` was skipped on a successful run, which removes the former success-path double move. A successful quarantine prevents queue starvation; if Outlook rejects the cleanup itself, the admin alert remains the recovery signal. Unread state distinguishes failed items from successfully processed Archive mail. After correcting the cause, move it back to Inbox to retry.
* **`trigger_reset.py`** replays historical Archive mail by marking it unread and moving it to Inbox. The scheduled poll then processes it one message per poll. Eligibility is only *in Inbox + unread* — `Get_unread_emails` applies **no date filter**, so age is irrelevant and any archived message becomes live again the moment it is moved back unread. (Its docstring was corrected 2026-08-03; it previously described the retired "When a new email arrives (V2)" delta-query trigger.)
* **`bulk_move_tool.py`** can restore larger Archive batches to Inbox; no special trigger reset is required after the updated ZIP is imported and enabled.

---

## 9. Confirmed-Good, Known Limitations, and What "Maximum Potential" Would Still Need

The 2026-07-12 audit is the authoritative source for the original findings — **PASS, deployable as-is**, no correctness bug in the deployed logic. What follows is current state, split by whether each item is applied or still open. Current test count is tracked by `test_p1.py`; the latest run validates the upload ZIP, the loose generated definition, config, docs, and workbook template together.

**Fixed this session (verified, tests green):**
- AppRef time-format/hex-length exposed as real `flow_config.json` keys instead of hardcoded Python constants — the file's "single source of truth" claim is now actually true for this value. Verified byte-identical rebuild (zero drift).
- Every stale `HHmmss`/6-hex AppRef example across README.md, this file, and the audit prompt corrected to the real `HHmm`/4-hex shape.
- Stale historical column-count labels (comments and doc prose only — the executable assertions were already correct) corrected to the current 30-column main / 29-column Rejected contract throughout `test_p1.py`, `trigger_reset.py`, and this file.
- README's claim that the Rejected sheet is "auto-created by P2" corrected — it's already shipped inside `P1_Templates/HiringAgent_P1_CandidateList.xlsx`, so a first-time operator doesn't wonder why P2-looking content exists on day one.
- Row-schema table updated for the current 30-column contract, including `Education`, `Info Request Sent`, and P1-owned `Last Updated Date`. P1 now writes 11 intake fields on creation.
- **Sender-address matching normalized (audit finding 2, closed).** Both the `Get_rows`/`Get_rows_ref` `$filter` and the stored `Email` column now run through the same shared expression: lowercase, and if the address has a `Name <addr>` wrapper, only the `<addr>` part is kept. `jane@x.com`, `Jane@X.com`, and `"Jane Doe" <jane@x.com>` now all match the same row going forward. **Scope, by design:** this is forward-only — rows already written before this change keep whatever raw value they have; a resend from someone whose *existing* row has display-name cruft may still miss once, but their new email's own row (or any row written after this change) is stored normalized, so it matches cleanly from then on. No backfill of historical rows was done or requested. This is a real behavior change to live dedup logic — rebuilt and reverified via `test_p1.py`'s new §D2 (5 new assertions checking both sides normalize identically), **not yet deployed** — still needs the zip re-imported into Power Automate to take effect live (see the workflow note in `flow_config.json`'s own header).
- **Resume-update reconciliation fixed (2026-07-15).** Root cause traced to P1, but the actual fix landed entirely on the P2 side — see the ¹ footnote above and `p2_detailed_summary.md` §2 Step 1. Also corrected several comments in `build_zip.py` that had over-claimed a guaranteed "byte-identical overwrite" that stopped being universally true once P2's 2026-07-15 filename convention shipped.
- **Category catch-all bug fixed (P2-side, but affects roles P1's own `flow_config.json` open-role list can produce)** — a bare `"developer"`/`"engineer"` substring rule was silently routing non-web roles (e.g. "Intern / Entry-Level Engineer") into Web Team instead of General. Removed; see `p2_detailed_summary.md` §2 Step 4.
- **Setup template (`P1_Templates/HiringAgent_P1_CandidateList.xlsx`) was stale** — regenerated to match the current 30-column main table and 29-column Rejected table. `test_p1.py` §J cross-checks the template headers directly against P2's canonical column lists, so this can't silently drift again.

**Fixed later the same day, 2026-07-31 (separate pass, after the P1 flow was already live-processing the backlog):**
- **`Last Updated Date` regression fixed.** All 6 `Patch_*_attempts`/`_noreply` actions used to stamp this column blindly from the incoming email's own `receivedDateTime`, with no guard - so replaying an older queued reply after a newer one had already patched the same row could move the timestamp *backward*. Now every patch uses `max(existing value, this email's own time)` (`_last_updated_expr` in `build_zip.py`). `Patch_followup_attempts` (a text-only "noted" reply, no resume) also used to skip this column entirely - now stamps it like every other path. This mattered downstream: P2's `_updated_after_location_request()` (`sharepoint_scoring.py`) trusts this column to detect whether a candidate replied after being asked for missing info: a corrupted/frozen timestamp made P2 think they never responded.
- **Resume-file mismatch fixed: `AppRef_current` replaces `AppRef_from_subject` for the saved filename and the applicant-facing Ref display.** Root cause: a sender still replying to an OLD email thread (their original application predates a since-fixed flow bug and never got a row) keeps quoting that dead reference number forever. `Get_rows_ref` correctly finds their real, CURRENT row by email regardless, and `Patch_update_attempts`/`Patch_followup_attempts` correctly patch it - but the saved update-resume filename and Emails 5/6's displayed "Ref:" used to be built from the text-quoted `AppRef_from_subject` instead, so the resume landed under a stale Application ID that P2's `_resume_name_slots()` (keyed strictly to the row's own current Application ID, no fallback search) could never find. New `AppRef_current` Compose action reads the matched row's real Application ID from `Get_rows_ref` and is now what `FileRef_from_subject` and Emails 5/6 use; `AppRef_from_subject` is kept computed but is display/audit-only now. One-time live repair (`resort`-adjacent script, not part of the flow) found 11 already-orphaned resume files from this bug on live SharePoint: 10 were byte-identical duplicates of the row's already-correct file (deleted); 1 (`redactedmarketing`) had genuinely newer content that wasn't reachable under the row's real Application ID (promoted to canonical). Verified after: 0 orphans, 0 missing, 0 duplicates across all rows.
- **Sheet display order.** P2's `resort_candidate_sheets.py` (new, standalone, on-demand - not wired into this flow or run automatically) can now re-sort the live CandidateList and Rejected sheets to current-month-first / newest-within-month-first, with a labeled `-- YYYY --` row at year boundaries instead of a plain blank. P1 itself is unaffected and unchanged - it still only ever appends new rows/separators at the physical bottom of the table in arrival order, same as always. See `HiringAgent_P2/docs/p2_detailed_summary.md` for the full mechanism and its safety design.

**Still open, deliberately not touched (each needs an explicit decision, not a silent fix):**
- **Malware detection is extension/phrase-substring only** — no magic-byte, MIME, or content scan, no size limit. A `resume.pdf` that's actually an executable, or a well-crafted malicious PDF, reaches SharePoint; O365/SharePoint's own AV is the only backstop. Hardening this is a real flow change (new actions, likely a new connector/Azure Function), not a config tweak — flagged as its own, separate design discussion, not started.
- **No sender allow-list**, **no rate-limit/429 backoff** beyond Power Automate's built-in per-action retry — both accepted gaps, not defects.
- **P1 is entirely untracked in git** (`git ls-files HiringAgent_P1` → empty) — nothing here is diffable against history. Worth deciding whether to start tracking it, independent of anything else in this doc.
- **`trigger_reset.py` can't touch Inbox-stuck mail** (§8) — the acknowledged gap behind the separate backfill-script effort.

**Bottom line on "maximum potential":** the sender-matching gap is now closed (pending live re-import). The extension-only malware check is addressed as of 2026-07-31 — not by a separate attachment gate, but by space-anchoring each extension against a space-terminated scan field, which removes the prose false positives at no structural risk. Remaining: two operational/tooling gaps (git tracking, Inbox-stuck backlog recovery) that aren't flow bugs at all.

**Spam-gate false-positive audit (2026-07-31).** All 167 terms were replayed against every message in the live mailbox (272 Archive + 6 Junk). Exactly one term had ever fired: `marketing@`, on three emails, all from the same genuine applicant. The other 166 had never matched anything. Seven terms were then proven unsafe by probe (`marketing@`, `alert`, `invoice`, `receipt`, `payment`, `survey`, plus `.exe`/`.iso` against `exeter.ac.uk`/`iso.org`) and fixed. Post-fix the same replay yields **zero hits across all 278 messages**. Note that three of the six Junk messages matched **no** P1 gate — those were junked by Exchange Online before the flow ever ran, which P1 cannot fix; they need a safe-sender or anti-spam policy change on the mailbox.

---

## 10. Coexistence Contract with Phase 2

Phase 1 writes `Status = "New Email Received"` and stops. Phase 2 (`../HiringAgent_P2/`) polls for that status, reads the resume, and keeps the full extraction/scoring result in its local database and `P2-MasterFile.xlsx`. It writes only `Status` and `Resume Link` back to the same P1 `CandidateList` row. Clear foreign current locations appear as `Rejected - Non-USA Location` on that intake row and under P2 master's `Rejected` sheet; P2 does not move or delete rows inside the P1 workbook.

---

## 11. Changes — 2026-08-03

**Duplicate window no longer depends on sheet row order.** `Is_duplicate` read
`last(Filter_submitted).Received Date`, which assumed the sheet is in P1's own append order
(oldest → newest). P2's `resort_candidate_sheets.py` re-sorts the live sheet to
**newest-first**, so `last()` began returning the sender's *oldest* row: a returning applicant
whose first contact was over 90 days ago read as brand new — new row, fresh acknowledgment,
cap system bypassed — while all 12 other row expressions read `first()` and stayed correct.
The two halves disagreed by construction, and under either ordering one had to be wrong.

Fixed by removing the ordering assumption rather than swapping ends: a new `Submitted_ticks`
Select projects every matching row to its tick value and the window compares
`max(union(Submitted_ticks, [0])) >= ticks(now - 90 days)`. The `[0]` floor matters — Logic
Apps evaluates **both** operands of an `and` (no short-circuit), so a brand-new applicant's
empty array would otherwise hit `max([])` and fail the run. `test_p1.py` §M3 now fails the
build if any expression reads `last()` off a candidate-row array.

**No acknowledgment without a recorded row.** `Send_acknowledgment` ran after
`Add_row: [Succeeded, Failed, Skipped]`, so an Excel write failure still mailed the applicant
an Application ID with no row behind it — and left nothing for P2 to score. Now pinned to
`{"Add_row": ["Succeeded"]}` in `build_zip.py` (not the zip base, which would drift on the
next rebuild). `Skipped` is excluded too: it only occurs when an upstream save timed out,
which is equally not a recorded application.

**Operating posture: silent live intake.** `email.send_applicant_emails: false` — all six
applicant sends are no-op `Compose` stubs; `Notify_failure` / `Notify_poll_failure` stay live.
The trigger still polls on its schedule and every SharePoint/Excel write still happens. Pinned by
`test_p1.py` §O2 so flipping it back on is a conscious edit. P2 holds the matching stance in
its `config.yaml` (`test_mode.suppress_emails: true`).

> **Known ceiling:** `Get_rows_for_separators` reads the whole table (`$top: 5000`,
> pagination 5000) on every new-applicant run, because Excel OData cannot `$filter` on a
> column whose name contains a space (`Received Date`). Past 5,000 rows the month/year
> separator check silently degrades. Closing it needs a no-space helper column.
