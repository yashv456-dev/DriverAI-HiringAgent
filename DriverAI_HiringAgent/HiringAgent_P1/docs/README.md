# HiringAgent P1 — Complete Reference

**Package:** `DriverAI_HiringAgent_P1_AutoReply`  
**Single source of truth:** `flow/flow_config.json` — edit this, run `python flow/build_zip.py`, re-import.  
**Generated artifacts (never hand-edit):** `flow/definition.json` · `flow/DriverAI-Hiring-AutoReply-apply.zip`

Power Automate watches `apply@driverai.io`, filters spam and non-applications, and saves each resume + candidate row to SharePoint. Applicant replies are controlled by `email.send_applicant_emails`; the current silent-live setting is `false`, so intake runs without contacting applicants. Phase 2 (`../HiringAgent_P2/`) scores the queued rows.

---

## Deploy

**One-time setup (~5 min):**

1. In the SharePoint site's **Documents** library, create folder **`Candidate_Resumes`** for resumes, **`Master_Files`** for the live workbook/client export, and **`SharePoint_Master_Template`** for a reference copy of the blank setup template.
2. Upload `P1_Templates/HiringAgent_P1_CandidateList.xlsx` into `Master_Files/` and name the SharePoint copy `Sharepoint_Master_File.xlsx`. There is **one workbook for all years** (no `<Year>` subfolder); this is what makes duplicate detection and Phase-2 scanning work across a year boundary. After upload, do not rename the SharePoint file, table (`HiringAgent_P1_Candidates`), worksheet (`CandidateList`), or Rejected worksheet/table. The template includes one intentionally blank seed row because Excel Online / Graph reject header-only tables; the automation ignores that blank row. **This upload is a one-time setup step - the flow no longer auto-creates the workbook** (see next note).
3. Grant the Office 365 connection owner **Send As / Send on Behalf** on `apply@driverai.io` — Full Access alone is not enough; all applicant replies are sent *from* that mailbox.

> **Workbook is set up once, never auto-overwritten (changed 2026-07-04).** Earlier versions auto-created the workbook via a metadata probe + template upload. That probe proved unreliable (false "NotFound"), so the self-heal kept re-uploading a blank template over real data and wiping candidate rows. The workbook is now **never auto-created or overwritten** — you upload it once (step 2). If it is ever genuinely missing or deleted, the run fails and `Notify_failure` emails the admin (`yashv@driverai.io`) so a human restores it from `P1_Templates/HiringAgent_P1_CandidateList.xlsx`; your data is never silently replaced. Dated **resume folders** (`<Year>/<Month>`) are still auto-created, which is safe (creating an existing folder is a harmless no-op).

> **Migrating from a per-year workbook?** If you previously ran the yearly layout, move the existing workbook into the master folder: from `Candidate_Resumes/<Year>/HiringAgent_P1_CandidateList.xlsx` to `Master_Files/Sharepoint_Master_File.xlsx` (keeps all rows + the table). Also update Phase 2's `SHAREPOINT_WORKBOOK` in `.env` to the new path (or remove it to use the new default).

**Auth and permissions runbook:**

Live P1 does **not** need an Entra client secret to run the imported flow. The production intake flow runs through the Power Automate connections selected during import:

1. **Office 365 Outlook connection** — must be able to read the shared mailbox trigger and send applicant replies.
2. **SharePoint connection** — must be able to create folders/files under `Documents/Candidate_Resumes`.
3. **Excel Online (Business) connection** - must be able to read/update `Sharepoint_Master_File.xlsx` and table `HiringAgent_P1_Candidates`.
4. **Shared mailbox permission** — the Office 365 connection owner must have **Send As** or **Send on Behalf** for `apply@driverai.io`. Full Access alone is not enough for the applicant-facing send actions.

The included Python helper/audit tools (`audit_p1_live_state.py`, `trigger_reset.py`, `bulk_move_tool.py`) are separate from the live Power Automate flow. They use the sibling P2 `.env` and the same Entra app registration used by P2:

1. Entra admin creates/uses an **App registration**.
2. Copy these values into `../HiringAgent_P2/.env`: `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`, plus the SharePoint settings used by P2.
3. In **Certificates & secrets**, create a **Client secret** and store only its value in `.env`; never place the secret in P1 docs, config, or the Power Automate zip.
4. In **API permissions**, use Microsoft Graph **Application permissions** and grant admin consent:
   - `Sites.ReadWrite.All` / `Files.ReadWrite.All` for SharePoint workbook/resume access used by P2 and the audits.
   - `Mail.ReadWrite` for mailbox audit/replay/move helpers (`audit_p1_live_state.py`, `trigger_reset.py`, `bulk_move_tool.py`).
   - `Mail.Send` is needed by P2 scoring-time emails, not by live P1; P1 sends through the Power Automate Outlook connector.
5. If mailbox Graph calls return `403 ErrorAccessDenied` even after admin consent, check Exchange **Application Access Policy** for the app and mailbox. The policy must allow that app's `CLIENT_ID` to access `apply@driverai.io`.

**Import / update:**

1. `make.powerautomate.com` → **My flows → Import → Import Package (Legacy)** → upload `flow/DriverAI-Hiring-AutoReply-apply.zip`.
2. On the review screen: first time → **Create as new**; re-deploy → **Update** and pick the existing flow. Do not "Create as new" on an update or you get two flows firing on the same inbox.
3. For each connection — **Office 365 Outlook**, **SharePoint**, **Excel Online (Business)** — click **Select during import** and pick your connection.
4. Click **Import**. After "successfully imported", open the flow → **Edit** and confirm the four Excel row actions (`Get_rows`, `Get_rows_ref`, `Add_row`, all `Patch_*`) show the correct workbook and table. Re-pick from dropdowns if any show red.
5. **Save → toggle On.**

The package reuses the same GUID (`36d96cc9-25b1-4e7c-952c-378ba087f142`) so Update imports preserve connections.

**Import troubleshooting (`DynamicParameterInputInvalid` / `GetTable` missing `table`):**

If Power Automate fails during import with:

```text
Flow save failed with code 'DynamicParameterInputInvalid'
The request to API 'sharepointonline' operation 'GetTable' is missing required property 'table'
```

the package usually imported with a workbook binding Power Automate could not resolve. Check these in order:

1. Confirm the workbook exists at the exact expected path: `Documents/Master_Files/Sharepoint_Master_File.xlsx`.
2. Confirm the file was **not renamed** during upload and still contains table **`HiringAgent_P1_Candidates`**.
3. Confirm the **SharePoint** and **Excel Online (Business)** connections selected during import are the same account and both can open that workbook.
4. If you are redeploying from the old yearly layout, move the workbook from `Candidate_Resumes/<Year>/HiringAgent_P1_CandidateList.xlsx` into `Master_Files/` as `Sharepoint_Master_File.xlsx` before importing.
5. If the SharePoint site or document library changed, update `flow/flow_config.json` (`sharepoint.site`, `excel.source`, `excel.drive`, `excel.file`, `excel.table`), run `python flow/build_zip.py`, then re-import the rebuilt zip.
6. If **Update** still fails, import once as **Create as new** / **Save as a new flow**, open the flow, re-pick any red Excel actions, save, then turn the old flow off so both do not process the same mailbox.

**After any config change:** edit `flow/flow_config.json` → `python flow/build_zip.py` → re-import (Update, re-pick the 3 connections). The build script prints a summary — use it to verify what was baked in:

```
bad_senders         : 42
bad_subjects        : 24 (config; self-loop guard is sender-based)
spam/safety gate    : 123 (30 scam + 36 offensive + 20 malware/exe + 25 non-EN + 12 vendor-pitch); scans subject+body+attachment names
app_keywords        : 21 (OR clauses: 42; E3 gated on no-attachment)
row extras          : Subject + body preview (~255 chars, bodyPreview) columns + improved Full Name guess
year separator      : ON - one blank row before the first new-applicant row of each year; the permanent header spacer is reused for the first-ever applicant
month separator     : ON - one blank row before the first new-applicant row of each month (suppressed at a year boundary so only ONE blank row lands)
```

(This is a trimmed excerpt of the real console output — run `python flow/build_zip.py` yourself to see every line, including trigger/admin/caps/appref/excel/resume-folder/inbox-tidy/workbook-self-heal settings.)

> **Coexistence rule:** P1/Power Automate owns intake replies only: acknowledgment, duplicate/update, request-CV, and spam/noise handling. Phase 2 with `--score-sharepoint` owns scoring-time mail only: non-USA/location-not-confirmed declines and missing-info/current-location requests. Do not add a separate second auto-reply flow on the same mailbox or the applicant can receive duplicate intake replies.

---

## Trigger

| Setting | Value |
|---|---|
| Trigger | Power Automate `Recurrence` |
| Mail action | Office 365 Outlook — Get emails (V3), shared mailbox |
| Mailbox | `apply@driverai.io` |
| Folder | Inbox |
| Read filter | Unread only (`fetchOnlyUnread=true`) |
| Include Attachments | Yes — content bytes are included; no second download needed |
| Poll interval | **1 min** (`trigger.interval_min`) — only takes effect after re-import |
| Messages per run | **1** (`trigger.unread_per_run`) |
| Concurrency | **1** — one scheduled run at a time; prevents double-row writes |

Every scheduled run fetches one unread message from the shared Inbox, so both newly delivered and older mail moved back to Inbox can be processed. One message per run means the backlog drain rate is one message per poll: at the 1-minute interval, 41 unread messages take about 45 minutes plus connector/runtime overhead. Successful legitimate mail is marked read and archived; spam is moved unread to Junk Email. A permanently failing item is moved **unread** to `Archive` after its admin alert so it cannot block the rest of the unread queue. Success and failure cleanup are mutually exclusive: the failure move cannot run when `Notify_failure` is skipped on a successful message. This prevents two `MoveV2` actions racing for the same Outlook message ID. Automatic retries are disabled for terminal moves, and their outcomes are explicitly finalized. The nested `Recruiting Review` folder is not used by `MoveV2`: its Graph ID was visible but the imported connector rejected it as `NotFound` at runtime.

The unread guard and the existing processing graph are top-level siblings. This is deliberate: Power Automate permits action nesting only through level 8, and wrapping the legacy graph would push its deepest update/follow-up actions to invalid level 9. The build and ZIP audits now reject any definition deeper than level 8.

---

## Full flow

```
TRIGGER — recurrence every 1 minute
│
└─ Get_unread_emails — first unread Inbox message from apply@driverai.io
│
├─ SETUP (all run in parallel after trigger)
│   CONFIG       Compose  { sharepoint_site, resumes_folder }
│   LowerFrom    Compose  toLower(from)
│   LowerSubject Compose  toLower(subject)
│   AttachNames  Select   [attachment file names only — merged into spam scan]
│   LowerBody    Compose  toLower(body + " " + join(AttachNames))
│   BadSenders   Compose  [31 sender fragments]
│   BadSubjects  Compose  [24 subject strings]
│   SpamPhrases  Compose  [123 always-on = 30 scam + 36 offensive + 20 malware/exe + 25 non-EN + 12 vendor-pitch]
│   LinkShortenerPhrases Compose [12 URL shorteners — checked separately, bypassed if
│                                 HasValidResumeEarly is true (a real resume was attached)]
│
├─ GATE 1 — IsSystem
│   MatchSender: filter BadSenders where fragment is contained in LowerFrom
│   IF length(MatchSender) > 0:
│     Move_to_processed_spam_sender → Terminate [silent; remains unread in Junk]
│   else:
│
├─ GATE 2 — IsSubject
│   MatchSubject: filter BadSubjects where string is contained in LowerSubject
│   IF length(MatchSubject) > 0:
│     Move_to_processed_spam_subject → Terminate [silent; remains unread in Junk]
│   else:
│
├─ GATE 3 — IsSpam
│   ResumeFilesEarly Query [attachments filtered to real .pdf/.docx w/ content, nested here, not top-level parallel setup; only runs after Gates 1 & 2 pass]
│   HasValidResumeEarly Compose [true if ResumeFilesEarly is non-empty]
│   MatchSpam: filter SpamPhrases where phrase in LowerSubject OR LowerBody
│   IF length(MatchSpam) > 0:
│     Move_to_processed_spam_body → Terminate [silent; remains unread in Junk]
│   else: ────────────────────────── WORK BRANCH ─────────────────────────────────
│
├─ MINT REFERENCE
│   AppRef     Compose  "APP-" + formatDateTime(receivedDateTime,"yyyyMMdd")
│                            + "-" + formatDateTime(receivedDateTime,"HHmm")
│                            + "-" + toUpper(substring(guid(),0,4))
│                       e.g. APP-20260630-1430-A3F9  (unique even if two
│                       emails arrive in the same minute — random 4-char hex tail;
│                       time_format/hex_length are flow_config.json keys, see appref.*)
│   FileRef    Compose  replace(AppRef, "/", "-")   (filename-safe)
│   ResumeFiles  Filter  attachments where name endsWith ".pdf" or ".docx"
│   ResumeNames  Select  item()?['name']  (original filenames — no prefix here;
│                         prefix goes on the saved file, not this column)
│
├─ BRANCH A — HasAppRefGate  [runs BEFORE HasResume — reply-first ordering]
│   IF LowerSubject OR LowerBody contains "app-20":
│   │
│   ├─ Get_rows_ref   Excel GetItems  $filter: Email eq '<normalized from>'
│   │                 (from lowercased + '<addr>' unwrapped from a display-name
│   │                  header - changed 2026-07-15, see docs/p1_detailed_summary.md §9)
│   └─ IsKnownSender  IF length(Get_rows_ref.value) > 0:
│       │
│       ├─ TRUE — known applicant:
│       │   RefSrc              subject if "app-20" in subject, else body
│       │   AppRef_from_subject extract "APP-YYYYMMDD-HHMM-XXXX" token from RefSrc
│       │                       (display/audit only since 2026-07-31 - see AppRef_current)
│       │   AppRef_current      first(Get_rows_ref.value)?['Application ID'] - the MATCHED
│       │                       row's real, live Application ID (fixed 2026-07-31: a sender
│       │                       replying to an old dead thread keeps quoting a stale ref
│       │                       forever; Get_rows_ref still finds their real current row by
│       │                       email, but the saved file/displayed Ref used to trust the
│       │                       quoted text instead, orphaning the resume from that row)
│       │   FileRef_from_subject replace("/", "-") on AppRef_current
│       │   │
│       │   └─ Has_resume_in_update  IF length(ResumeFiles) > 0:
│       │       │
│       │       ├─ TRUE — resume attached:
│       │       │   IsUnderUpdateCap  IF Application Updates (first row) < 5 (overall cap):
│       │       │     TRUE:
│       │       │       Save_update_resume   ForEach → Create_update_file
│       │       │         folderPath = <ORIGINAL row's Received Date> /<yyyy>/<MMMM>
│       │       │         name       = <FirstLast>_<FileRef_from_subject>.<ext>  (changed 2026-07-12; was <FileRef_from_subject>_<original-filename>)
│       │       │       IsUnderReplyCap_update  IF Application Updates (first row) < 3 (reply_cap):
│       │       │         TRUE:
│       │       │           Send_update_ack      → [Email 5]
│       │       │           Patch_update_attempts  Application Updates + 1, Mail Sent stamped,
│       │       │             Status → "New Email Received", Original Filename untouched
│       │       │             (re-queues the row so P2 re-scores with the updated CV)
│       │       │           Mark_as_read_update → Move_to_processed_update → Terminate
│       │       │         FALSE — reply_cap reached, overall cap not yet reached:
│       │       │           Patch_update_attempts_noreply  Application Updates + 1, Status re-queued,
│       │       │             Mail Sent left untouched (no email sent — "no other reminders")
│       │       │           Mark_as_read_update_noreply → Move_to_processed_update_noreply → Terminate
│       │       │     FALSE — overall cap reached:
│       │       │       Mark_as_read_update_cap → Move_to_processed_update_cap → Terminate [silent]
│       │       │
│       │       └─ FALSE — no resume:
│       │           IsUnderFollowupCap  IF Application Updates (first row) < 5 (overall cap):
│       │             TRUE:
│       │               IsUnderReplyCap_followup  IF Application Updates (first row) < 3 (reply_cap):
│       │                 TRUE:
│       │                   Send_noted_reply     → [Email 6]
│       │                   Patch_followup_attempts  Application Updates + 1, Mail Sent stamped
│       │                   Mark_as_read_followup → Move_to_processed_followup → Terminate
│       │                 FALSE — reply_cap reached, overall cap not yet reached:
│       │                   Patch_followup_attempts_noreply  Application Updates + 1 only
│       │                   Mark_as_read_followup_noreply → Move_to_processed_followup_noreply → Terminate
│       │             FALSE — overall cap reached:
│       │               Mark_as_read_followup_cap → Move_to_processed_followup_cap → Terminate [silent]
│       │
│       └─ FALSE — sender not in workbook: no-op (falls through to Branch B)
│
│   else — no ref quoted: no-op (falls through to Branch B)
│
├─ BRANCH B — HasResume  [runs after HasAppRefGate Succeeded OR Skipped]
│   IF length(ResumeFiles) > 0:
│   │
│   ├─ TRUE — PDF/DOCX attached:
│   │   Get_rows          Excel GetItems  $filter: Email eq '<normalized from>'
│   │                     (same lowercase + display-name-stripped value as Get_rows_ref)
│   │   Filter_submitted  filter Get_rows where Has Resume == 'Yes'
│   │   Is_duplicate      IF length(Filter_submitted) > 0
│   │                     AND max(Submitted_ticks) >= ticks(now - 90 days):   [order-independent]
│   │   │
│   │   ├─ TRUE — prior submission within 90 days:
│   │   │   IsUnderDupCap  IF Application Updates (first row) < 5 (overall cap):
│   │   │     TRUE:
│   │   │       Ensure_dup_resume_folder → Save_dup_update_resume (resend saved regardless of reply_cap)
│   │   │       IsUnderReplyCap_dup  IF Application Updates (first row) < 3 (reply_cap):
│   │   │         TRUE:
│   │   │           Send_duplicate_notice  → [Email 2]
│   │   │           Patch_dup_attempts     Application Updates + 1, Mail Sent stamped
│   │   │           Dup_patch_done         no-op Compose — absorbs PatchItem failure (best-effort)
│   │   │         FALSE — reply_cap reached, overall cap not yet reached:
│   │   │           Patch_dup_attempts_noreply  Application Updates + 1 only (no email — "no other reminders")
│   │   │     FALSE — overall cap reached: no-op [silent]
│   │   │
│   │   └─ FALSE — new applicant:
│   │       Save_resumes_to_SharePoint  ForEach → Create_file
│   │         folderPath = resumes_folder/<yyyy>/<MMMM>
│   │         name       = <FirstLast>_<FileRef>.<ext>  (changed 2026-07-12; was <FileRef>_<original-filename>)
│   │       Get_rows_for_separators → Filter_rows_this_year  → IsFirstOfNewYear
│   │                               └→ Filter_rows_this_month → IsFirstOfNewMonth
│   │         (one paginated Excel scan; date filtering is done in memory because Excel
│   │          OData Filter Query rejects the spaced `Received Date` column name)
│   │         (month check also requires rows already exist this year, so a January
│   │          arrival inserts ONE blank row, not two stacked)
│   │       Add_row   Status="New Email Received", Application Updates=0  (11 P1 fields)
│   │         waits on BOTH separator checks (Succeeded/Failed/Skipped) so a blank
│   │         row always lands before the real row, and never blocks it
│   │       Send_acknowledgment  → [Email 1]
│   │
│   └─ FALSE — no PDF/DOCX:
│       Has_application_keyword
│         IF length(all attachments) == 0
│         AND LowerSubject or LowerBody contains any of 21 app_keywords:
│           TRUE:
│             Send_CV_request  → [Email 3]   (no row — table is CV-only)
│           FALSE: no-op [silent]
│       HasWrongFormat  [always runs after Has_application_keyword]
│         IF length(all attachments) > 0 AND length(ResumeFiles) == 0:
│           TRUE:
│             Send_wrong_format  → [Email 4]
│           FALSE: no-op [silent]
│
│   Has_application_keyword and HasWrongFormat are mutually exclusive:
│   the keyword check requires zero attachments; wrong-format requires at least one.
│
├─ INBOX TIDY (three outcome categories, two well-known folders)
│   Legit mail: Mark_as_read → Move_to_processed (Archive, read)
│   Spam/junk:  Move_to_processed only (Junk Email, left UNREAD — no mark)
│   20 tidy actions total: 7 mark-read + 11 move + 2 terminal finalizers
│     3 spam-gate moves → Junk Email, unread (no mark)
│     6 ref-reply pairs → Archive, read (update, update-noreply, update-cap,
│                                        followup, followup-noreply, followup-cap)
│     1 top-level pair → Archive, read (new-applicant / duplicate / CV-request / wrong-format / ignored)
│     1 failure move → Archive, unread; only after an actual Notify_failure attempt
│     2 finalizers handle terminal move outcomes without cross-activating skipped paths
│
├─ ERROR HANDLER — Notify_failure
│   Runs after HasResume Failed or TimedOut
│   Sends high-priority admin alert to yashv@driverai.io
│   Best-effort move of the failed message unread to Archive; admin alert remains authoritative
│
└─ ERROR HANDLER — 7 patch-failure alerts (2026-08-03) + 2 intake-write alerts (2026-08-24)
    One SIBLING alert per candidate-row PatchItem, each on [Failed, TimedOut] only
      Patch_mail_sent · Patch_dup_attempts(+_noreply)
      Patch_update_attempts(+_noreply) · Patch_followup_attempts(+_noreply)
    Covers the gap Notify_failure never did: cleanup runs on Failed by design, so a
    failed patch left the run GREEN and the counter silently un-incremented
    Sibling, not child — the update/follow-up patches are at depth 8 (the PA ceiling)
    Nothing depends on them, so they cannot alter any existing path
```

**Resume folder path**: `<Year>/<Month>` from the email's `receivedDateTime` — no week-level
subfolder (dropped 2026-07-08; one folder per month). Phase 2 reconstructs this identical path
from the row's `Received Date` to download the resume; for a candidate it rejects, the file
instead moves to a single `<Year>/Rejected/` folder (changed 2026-07-15 from one `Rejected/`
subfolder per month — see the P2 README).

---

## Spam / phishing / malware / duplicate system

### Gate 1 — Sender (31 fragments, substring match on lowercased From)

Blocks system senders, notification platforms, SaaS mailers, and the flow's own address:

`noreply` · `no-reply` · `donotreply` · `notifications` · `mailer-daemon` · `postmaster` · `bounce` · `newsletter` · `billing@` · `teams.mail.microsoft.com` · `notify.microsoft.com` · `email.microsoftonline.com` · `linkedin.com` · `accounts.google.com` · `slack.com` · `slack-mail.com` · `zoom.us` · `zoom.com` · `calendly.com` · `atlassian.net` · `github.com` · `docusign.net` · `mailchimp` · `sendgrid.net` · `salesforce.com` · `amazonses.com` · `facebookmail.com` · `dropbox.com` · `notion.so` · `asana.com` · **`apply@driverai.io`**

`apply@driverai.io` is in this list so the flow never answers its own sent mail or any bounce. Self-loop protection is **sender-based, not subject-based** — outbound subjects are intentionally NOT in Gate 2. An applicant's `Re: <our subject>` reply must pass Gate 2 to reach the reference branch. Per-applicant caps (5 max) stop any runaway loop.

> **Every term here must be safe as an unanchored substring.** The gate is `contains(LowerFrom, item())` over the raw From header, so a bare role local-part also matches any personal address ending in it. `marketing@` was removed on 2026-07-31 for exactly this: it matched `redactedmarketing@example.com`, a real Chief Marketing Officer applicant whose three emails were all silently junked. `billing@` carries the same latent risk (no evidence of a hit yet). Prefer a distinctive domain (`mailchimp`, `sendgrid.net`) over a short local-part. `test_p1.py` guards the specific terms already found unsafe.

### Gate 2 — Subject (24 strings, substring match on lowercased Subject)

`sent a message` · `notification` · `security alert` · `digest` · `invoice attached` · `your receipt` · `unsubscribe` · `password reset` · `confirm your email` · `verify your` · `out of office` · `delivery failed` · `returned mail` · `webinar` · `meeting recording` · `your order` · `payment declined` · `security code` · `verification code` · `free trial` · `take our survey` · `thanks for applying to driver ai` · `your driver ai application is already` · `quick favor for your driver ai application`

The last 3 are legacy outbound subjects from a prior flow version — kept to block re-replies to stale templates.

> **Narrowed 2026-07-31.** Five entries were bare words that also appear in ordinary job titles, and each would have silently junked a real applicant: `alert` ("Alert Systems Engineer"), `invoice` ("Invoice Processing Specialist"), `receipt` ("Receipt Reconciliation Analyst"), `payment` ("Senior Engineer | Payments Platform"), `survey` ("Land Surveyor"). They are now the specific spam phrases shown above. Keep new entries multi-word for the same reason.

### Gate 3 — Content (123 always-on phrases + 12 conditional, against lowercased Subject + Body + all attachment filenames)

Attachment filenames are merged into the scan field via `AttachNames` → `join`, so `invoice.exe` is caught even with an innocent subject. The join is followed by **one trailing space** (`build_zip.py`, `LowerBody`), which space-terminates the last filename — see the malware row below.

| Group | Count | Always blocks? | Examples |
|---|---|---|---|
| Scam / phishing | 30 | Yes | `you have won`, `wire transfer`, `lottery`, `click here to claim`, `unusual sign-in activity`, `update your payment`, `your package is waiting`, `confirm your identity`, `login immediately` |
| Offensive / harassment | 36 | Yes | Sexual content phrases, racial slurs, explicit threats (`kill yourself`, `i will kill you`, `rape you`, `i will find you`, `i will rape`) |
| Malware / executables | 20 | Yes | `.exe ` `.scr ` `.bat ` `.cmd ` `.vbs ` `.ps1 ` `.lnk ` `.iso ` `.dll ` `.hta ` `.cpl ` `.pif ` (**trailing space is part of the term**) · `enable macros` `disable antivirus` `download and run` |
| Non-English scams | 25 | Yes | Spanish · French · German · Portuguese · Russian · Chinese · Japanese · Hindi · Arabic — lottery, wire-transfer, and inheritance phrases |
| Vendor / agency solicitation | 3 | Yes | `company profile and corporate deck`, `capability deck`, `engagement models` |
| **URL shorteners** | **12** | **No — bypassed if a real PDF/DOCX is attached** (`HasValidResumeEarly`) | `bit.ly/` `tinyurl.com/` `t.me/` `ow.ly/` `rb.gy/` `is.gd/` `rebrand.ly/` `cutt.ly/` `shorturl.at/` `v.gd/` `s.id/` `urlzs.com/` |

All matching is lowercase substring. No reply is sent on a Gate 3 hit — replying confirms a live mailbox to the sender.

> **Vendor / agency solicitation phrases (added 2026-08-01).** B2B vendors pitching their own services to the recruiting mailbox are not applicants, but nothing in their wording is malicious, so none of the other four groups touched them. Live case: a services vendor emailed `apply@driverai.io` three times sharing a "Company Profile and Corporate Deck" and offering "engagement models"; P1 created a candidate row each time and sent real applicant lifecycle replies ("Please Resend Your Resume", "we don't have your location on file") to a sales rep. The three phrases are deliberately narrow B2B sales-deck language that essentially never appears in a genuine application — verified against all live candidate resumes with zero false positives. **Known limit:** this catches wording, not intent. A staffing agency submitting someone else's CV ("please find the attached resume of my developer") uses ordinary polite English and still passes — those are removed manually. A stronger structural signal (corporate sender domain **plus** a business-title signature such as *Business Development Manager* / *Team Lead*) was validated offline at 5/5 vendors caught, 0/9 real candidates flagged, but is **not implemented**; it must live in P1 because the `Mail Body` column stores only ~256 characters and email signatures fall past that cutoff.

> **Executable extensions are space-anchored (2026-07-31).** Each is stored with a trailing space (`".exe "`, not `".exe"`) because the scan field is prose *plus* filenames. Unanchored, `.exe` matched `www.exeter.ac.uk` and `.iso` matched `www.iso.org` — so a University of Exeter graduate, or any ISO 27001 / CISSP candidate citing the standards body, was classified as malware and silently junked. With the trailing space, `invoice.exe` still matches (`...invoice.exe `) while `.exet` and `.iso.` do not. If you add an extension, include the trailing space; if you change `LowerBody`, keep its final `' '`. The shortener list is separate on purpose: email signature generators commonly wrap a LinkedIn/social icon in one of these redirect domains, so a real resume attachment is treated as a strong enough good-faith signal to let it through — the other four groups have no such exception, since there's no legitimate reason a real applicant's email would contain that content.

### Duplicate protection

`Filter_submitted` filters `Get_rows` output where **`Has Resume == 'Yes'`** (not by Status — every table row has a real CV; this is the invariant). `Is_duplicate` fires when at least one prior row exists AND `max(Submitted_ticks) >= ticks(now - 90 days)`, where `Submitted_ticks` is a `Select` projecting every matching row to its `Received Date` tick value.

> **The date check does not depend on row order (fixed 2026-08-03).** It previously read `last(Filter_submitted).Received Date`, documented here as "the most recent row" — true only while P1's own append order (oldest → newest) was the sheet's order. P2's `resort_candidate_sheets.py` re-sorts the live sheet to **newest-first**, which inverted it: `last()` began returning the sender's **oldest** row, so a returning applicant whose first contact was over 90 days ago was read as brand new — new row, fresh acknowledgment, cap system bypassed — while all 12 other row expressions read `first()` and stayed correct. Taking the **max over the whole array** removes the assumption instead of swapping one end for the other, so any future re-sort is harmless. `test_p1.py` M3 now fails the build if any expression reads `last()` off a candidate-row array. `first()` is still used for reading/writing `Application Updates` and the other row fields — correct on the live newest-first sheet.

**Cross-year duplicates (fixed):** the candidate workbook is now a **single file** for all years (`Master_Files/Sharepoint_Master_File.xlsx`, no `<Year>` subfolder), so a December applicant re-applying in January is correctly detected as a duplicate, and rows written near year-end never orphan when Phase 2 rolls to a new year. Resumes are still filed under `<Year>/<Month>/`.

**Cap system — Emails 2, 5, and 6 share one `Application Updates` counter per row, with TWO thresholds:**

- `business_rules.reply_cap` (**3**) — how many contacts actually get an EMAIL. The 3rd is the final-notice; nothing is ever sent twice.
- `business_rules.duplicate_notice_max` / `update_resume_max` / `followup_reply_max` (**5** each) — the OVERALL cap. `Application Updates` keeps counting past `reply_cap` up to this ceiling — the row/resume is still saved and re-queued for Phase 2 on every contact under this cap (dup/update paths only; follow-up has no resume to re-queue) — but contacts 4 and 5 get **no reply at all** ("no other reminders"). At the ceiling, everything (including the counter) stops.

| Contact | Application Updates before reply | Email sent? | Closing | After patch |
|---|---|---|---|---|
| 1st | 0 | Yes | Standard | 1 |
| 2nd | 1 | Yes | Standard | 2 |
| **3rd** | **2** | **Yes** | **Final-notice** | 3 |
| 4th | 3 | **No — silent** | — | 4 (row/resume still updated) |
| 5th | 4 | **No — silent** | — | 5 (row/resume still updated) |
| **6th+** | **5** | **No — fully silent, no patch either** | — | — |

Final-notice threshold = `reply_cap − 1` = 2. Controlled by `_closing()` in `build_zip.py`; the reply/no-reply split is a nested `IsUnderReplyCap_*` gate inside each of `IsUnderDupCap`/`IsUnderUpdateCap`/`IsUnderFollowupCap` (the overall-cap gates, unchanged at 5).  
Email 3 (no-CV reply) is **not capped** — no row means no counter. Each distinct no-CV email from the same sender gets one reply, indefinitely.

### Silent / no-reply scenarios

| Scenario | Where | Row? |
|---|---|---|
| Sender in `bad_senders` | Gate 1 | No |
| Subject in `bad_subjects` | Gate 2 | No |
| Spam / offensive / malware phrase in subject, body, or attachment name | Gate 3 | No |
| Quotes `APP-20` ref but sender not in workbook | `IsKnownSender = false` | No |
| Known sender, new resume, `Application Updates ≥ 5` | `IsUnderUpdateCap = false` | No |
| Known sender, no resume, `Application Updates ≥ 5` | `IsUnderFollowupCap = false` | No |
| Duplicate resubmit, `Application Updates ≥ 5` | `IsUnderDupCap = false` | No |
| Contact 4th/5th (`reply_cap ≤ Application Updates < 5`), any of the 3 capped scenarios | `IsUnderReplyCap_* = false` | No new row — the existing row IS still updated (resume re-saved + re-queued for dup/update); only the email is skipped |
| No attachment AND no hiring keyword | Falls through all branches | No |
| Core candidate-processing branch errors | Admin alert → move message unread to `Archive` | No |

**Inbox tidy — three outcome categories, two folders:**
- **Legitimate application mail** (new CV, non-CV request, duplicate, wrong-format, update, follow-up, and the "ignored / not an application" fall-through) → **marked read + moved to `Archive`**.
- **Spam / junk** caught by the 3 gates (bad sender, bad subject, scam / phishing / malware / virus / offensive / foreign-scam) → **moved to `Junk Email`, left UNREAD** (just moved, not marked read) so it stands out as junk. No reply is ever sent to these.
- **Error-path emails** trigger an admin alert and make a best-effort move **unread** to **`Archive`**. On a successful move, this quarantines the failed item and lets the next unread message run on the next minute. If Outlook itself rejects that cleanup, the admin alert remains the authoritative recovery signal. Unread state distinguishes failed Archive items from successfully processed mail.

Config: `inbox_tidy.destination` (`Archive`) for legit, `inbox_tidy.spam_destination` (`Junk Email`) for spam, and `inbox_tidy.failure_destination` for failures. Moving an email out of Inbox prevents it from being polled again.

---

## Email replies

All 6 applicant emails are sent **from `apply@driverai.io`** via `SharedMailboxSendEmailV2`.  
All bodies use `font-family:Arial,Helvetica,sans-serif; font-size:14px; color:#222; line-height:1.7`.  
All carry this shared footer:

```
────────────────────────────────────────
This is an auto-generated email and this mailbox is not monitored.
```

---

### Email 1 — Application Received (`Send_acknowledgment`)

**Fires:** New sender, PDF/DOCX attached, not a duplicate within 90 days.  
**Writes row:** Yes — `Status = "New Email Received"`, `Application Updates = 0`. No cap.

Subject: `Application Received - DriverAI (Ref: APP-20260630-1430-A3F9)`

```
Hello,

Thanks for applying to DriverAI. Your application has been received
and is under review.

Reference number: APP-20260630-1430-A3F9

If you are selected, a member of our team will contact you to discuss
next steps. You do not need to reply unless you are updating your
resume. To do that, send a new or reply email to apply@driverai.io
with the subject "Update - APP-20260630-1430-A3F9" and attach the new file.

Good luck on your next journey.
Warm regards,
The DriverAI Recruiting Team

────────────────────────────────────────
This is an auto-generated email and this mailbox is not monitored.
```

---

### Email 2 — Already On File (`Send_duplicate_notice`)

**Fires:** Same sender re-applies with a resume within 90 days, `Application Updates < 5`.  
**Writes row:** No *new* row — but the resume IS saved, under the same `<FirstLast>_<FullAppID>`
name every P1 save uses (fixed 2026-07-04). That's a true in-place overwrite only if the row hasn't
been scored yet; once P2 has scored and renamed the file to its `<First>_<Last>_<tail>`
shape (2026-07-15 filename convention — see the "Resume filenames" section below), this instead
writes a fresh file under the old name, and P2's next pass reconciles the two (see below) so the
candidate still ends up with exactly one current resume file. The *existing* row is re-queued
for Phase 2 rescoring (Status reset to "New Email Received"), same as the ref-quoting update path
(Email 5). **Fixed 2026-07-03:** a resend that never quoted the reference number used to be silently
discarded — only `Application Updates` incremented, the file was never saved and the row never rescored.
Cap: 5 (shared across duplicate/update/follow-up contacts on this row).

Subject: `Your DriverAI Application Is Already On File (Ref: APP-20260630-1430-A3F9)`

**Standard closing (Application Updates 0–1 before this reply — the 1st or 2nd reply after the original application):**

```
Hello,

Your application is already on file with DriverAI and under review,
so there's no need to resubmit.

If you are selected, a member of our team will contact you to discuss
next steps. You do not need to reply unless you are updating your
resume. To do that, send a new or reply email to apply@driverai.io
with the subject "Update - APP-20260630-1430-A3F9" and attach the new file.

Good luck on your next journey.
Warm regards,
The DriverAI Recruiting Team

────────────────────────────────────────
This is an auto-generated email and this mailbox is not monitored.
```

**Final-notice closing (Application Updates = 2 before this reply — the 3rd and LAST reply that will ever send; fires exactly once, never twice. Contacts 4 and 5 still update the row/resume but get NO reply at all — "no other reminders" — and 6th+ is fully silent):**

```
Hello,

Your application is already on file with DriverAI and under review,
so there's no need to resubmit.

This is our final automated reply regarding this application. If there
is a match, a member of our team will contact you directly. You do not
need to reply, and you are welcome to apply again after 90 days.

Good luck on your next journey.
Warm regards,
The DriverAI Recruiting Team

────────────────────────────────────────
This is an auto-generated email and this mailbox is not monitored.
```

---

### Email 3 — Please Attach Your Resume (`Send_CV_request`)

**Fires:** Zero attachments AND subject or body contains a hiring keyword. No row written. No cap.

Subject: `Please Attach Your Resume - DriverAI`

```
Hello,

Thanks for your interest in DriverAI. We didn't see a resume
attached, so your application isn't complete yet.

Please reply with your resume attached in PDF or Word (.docx) format.
Once we have it, your application goes under review, and if you are
selected, a member of our team will contact you to discuss next steps.

Good luck on your next journey.
Warm regards,
The DriverAI Recruiting Team

────────────────────────────────────────
This is an auto-generated email and this mailbox is not monitored.
```

---

### Email 4 — Wrong File Format (`Send_wrong_format`)

**Fires:** Attachment present but not `.pdf` or `.docx` (e.g. `.doc`, `.jpg`, `.zip`). No row. No cap.

Subject: `Please Resend Your Resume as PDF or Word - DriverAI`

```
Hello,

Thanks for your interest in DriverAI. We received your submission,
but we couldn't open the attached file.

Please reply with your resume as a PDF (.pdf) or Word (.docx) file
and we'll process it right away. If you are selected, a member of
our team will contact you to discuss next steps.

Good luck on your next journey.
Warm regards,
The DriverAI Recruiting Team

────────────────────────────────────────
This is an auto-generated email and this mailbox is not monitored.
```

---

### Email 5 — Updated Resume Received (`Send_update_ack`)

**Fires:** Email quotes `APP-20...` ref + sender in workbook + new PDF/DOCX attached + `Application Updates < 5`.  
**Ref shown/used for the saved filename (fixed 2026-07-31):** always the matched row's real, current Application ID (`AppRef_current`, sourced from `Get_rows_ref` by sender email) - never the `APP-20...` text the sender's email happened to quote. A sender still replying to an old dead thread otherwise gets told (and the file gets saved under) a stale reference that no longer corresponds to any live row.  
**Writes row:** No new row, but **re-queues the existing row** — saves the new resume into the original application's dated folder under the same `<FirstLast>_<FullAppID>` name P1 always uses (fixed 2026-07-04), increments `Application Updates`, and flips `Status` back to `New Email Received` so **P2 re-scores with the updated CV**. `Original Filename` is deliberately left unchanged. On a row that's never been scored, this is a true in-place overwrite (one CV, never a second file). On a row that's **already been scored once**, P2 has since renamed the file to its `<First>_<Last>_<tail>` shape (2026-07-15 filename convention) — so this write lands as a *new* file under the old name, not an overwrite of the renamed one. **P2's re-score pass is what reconciles this** (`_download_resume_text` in `sharepoint_scoring.py`, fixed 2026-07-15): it always prefers a `<FirstLast>_<FullAppID>`-shaped file over a `<First>_<Last>_<tail>`-shaped one for the same Application ID — since that shape can only exist post-scoring if P1 just wrote a fresh update — scores off it alone (never blended with the stale prior content), then deletes the stale sibling once the rename lands. Net effect for the candidate is still exactly one current resume file, just resolved a run later by P2 rather than by P1's write itself. Cap: 5.

Subject: `Updated Resume Received - DriverAI (Ref: APP-20260630-1430-A3F9)`

**Standard closing (Application Updates 0–1 before this reply — the 1st or 2nd reply after the original application):**

```
Hello,

Thanks for sending your updated resume. Your application now reflects
the latest version and our team will review it.

If you are selected, a member of our team will contact you to discuss
next steps. You do not need to reply unless you are updating your
resume again. To do that, send a new or reply email to
apply@driverai.io with the subject "Update - APP-20260630-1430-A3F9"
and attach the new file.

Good luck on your next journey.
Warm regards,
The DriverAI Recruiting Team

────────────────────────────────────────
This is an auto-generated email and this mailbox is not monitored.
```

**Final-notice closing (Application Updates = 2 before this reply — the 3rd and LAST reply that will ever send; fires exactly once, never twice. Contacts 4 and 5 still update the row/resume but get NO reply at all — "no other reminders" — and 6th+ is fully silent):**

```
Hello,

Thanks for sending your updated resume. Your application now reflects
the latest version and our team will review it.

This is our final automated reply regarding this application. If there
is a match, a member of our team will contact you directly. You do not
need to reply, and you are welcome to apply again after 90 days.

Good luck on your next journey.
Warm regards,
The DriverAI Recruiting Team

────────────────────────────────────────
This is an auto-generated email and this mailbox is not monitored.
```

---

### Email 6 — Message Received (`Send_noted_reply`)

**Fires:** Email quotes `APP-20...` ref + sender in workbook + no new resume + `Application Updates < 5`.  
**Writes row:** No, but now also stamps `Last Updated Date` (fixed 2026-07-31 - previously the one patch path that silently skipped it). Increments `Application Updates`. Cap: 5.  
**Ref shown (fixed 2026-07-31):** same `AppRef_current` fix as Email 5 - always the matched row's real Application ID, never the sender's possibly-stale quoted text.

Subject: `Message Received - DriverAI (Ref: APP-20260630-1430-A3F9)`

**Standard closing (Application Updates 0–1 before this reply — the 1st or 2nd reply after the original application):**

```
Hello,

Thanks for following up. Your message has been noted alongside your
application.

If you are selected, a member of our team will contact you to discuss
next steps. You do not need to reply unless you are updating your
resume. To send one, reply with the subject "Update - APP-20260630-1430-A3F9"
and attach the file (PDF or Word).

Good luck on your next journey.
Warm regards,
The DriverAI Recruiting Team

────────────────────────────────────────
This is an auto-generated email and this mailbox is not monitored.
```

**Final-notice closing (Application Updates = 2 before this reply — the 3rd and LAST reply that will ever send; fires exactly once, never twice. Contacts 4 and 5 still update the row/resume but get NO reply at all — "no other reminders" — and 6th+ is fully silent):**

```
Hello,

Thanks for following up. Your message has been noted alongside your
application.

This is our final automated reply regarding this application. If there
is a match, a member of our team will contact you directly. You do not
need to reply, and you are welcome to apply again after 90 days.

Good luck on your next journey.
Warm regards,
The DriverAI Recruiting Team

────────────────────────────────────────
This is an auto-generated email and this mailbox is not monitored.
```

---

### Admin alert — `Notify_failure` (not applicant-facing)

**Fires:** `HasResume` fails or times out — i.e. the main candidate-processing branch. *Not* every action in the flow (corrected 2026-08-03; see the patch-failure alerts below for the paths this one never covered).  
**Sent from:** Connection owner account — not `apply@`; works regardless of Send As rights.  
**To:** `yashv@driverai.io` (`email.admin_email`). No shared footer.  
**Retries:** exponential ×2 (changed 2026-08-03). The blanket "no retries" rule exists for `MoveV2`/`MarkAsRead_V3`, where a replay races the mutable Outlook message ID — a send carries no such hazard, and while applicant mail is suppressed these alerts are the *only* live sends, so a single transient blip must not lose the only signal.  
**Effect:** The failed message is moved unread to `Archive` only after `Notify_failure` actually runs (or fails/times out), preventing one bad message from becoming the first unread item forever. A skipped alert is the normal success path and never activates this failure move. After correcting the cause, move it back to Inbox (still unread) to retry. `flowFailureAlertSubscribed: true` remains a backup PA-level alert.

Subject: `[Hiring Auto-Reply] A run needs attention (Ref APP-20260630-1430-A3F9)`

```
Hiring Auto-Reply: a run did not complete.

Something went wrong while handling an incoming email, so the
applicant may not have received a reply and the candidate row or
resume file may not have been saved.

  Reference: APP-20260630-1430-A3F9
  From: applicant@example.com
  Subject: Applying for driver role
  Received: 2026-06-30T14:30:52Z

Open the flow run history for the exact step and error.
The email was moved unread to Archive so the remaining unread queue can continue.
```

---

### Admin alert — patch-failure alerts ×7 (added 2026-08-03)

**Fires:** one of the seven candidate-row `PatchItem` actions fails or times out.

> See also **intake-write alerts ×2** below — those cover `Add_row` and `Create_file`, the two writes that CREATE a candidate rather than update one.

| Alert | Watches | Branch | On failure |
|---|---|---|---|
| `Notify_mail_sent_failed` | `Patch_mail_sent` | new applicant — acknowledgment stamp | silent → **alert is the only signal** |
| `Notify_dup_attempts_failed` | `Patch_dup_attempts` | duplicate re-application | silent → **alert is the only signal** |
| `Notify_dup_attempts_noreply_failed` | `Patch_dup_attempts_noreply` | duplicate, past reply cap | **propagates** → run goes red, `Notify_failure` fires too |
| `Notify_update_attempts_failed` | `Patch_update_attempts` | resume update | silent → **alert is the only signal** |
| `Notify_update_attempts_noreply_failed` | `Patch_update_attempts_noreply` | resume update, past reply cap | silent → **alert is the only signal** |
| `Notify_followup_attempts_failed` | `Patch_followup_attempts` | follow-up reply | silent → **alert is the only signal** |
| `Notify_followup_attempts_noreply_failed` | `Patch_followup_attempts_noreply` | follow-up, past reply cap | silent → **alert is the only signal** |

**One exception, on purpose.** `Patch_dup_attempts_noreply` is the *last* action in its scope, so nothing absorbs its failure — it propagates up to `HasResume`, the run reports **Failed**, and `Notify_failure` fires as well. Its alert is therefore a deliberate duplicate (you get two emails), and its body says so rather than falsely claiming to be the only signal. It is kept because the reply-path twin `Patch_dup_attempts` *is* swallowed by `Dup_patch_done`: if a sibling is ever appended to the noreply scope for symmetry, that failure would start going silent and the coverage is already in place. `test_p1.py` asserts this swallow/propagate map against the real wiring, so the alert text can never quietly start lying.

**Why these exist.** Inbox cleanup is deliberately unconditional — every action after a patch lists `Failed` in its `runAfter` so a poisoned message can never block the unread queue. The side effect was that a failed patch was *fully silent*: the enclosing `If` reported `Succeeded`, the branch ended on a `Terminate` whose `runStatus` is `Succeeded`, and **the run went green**. The candidate's `Application Updates` counter had silently not incremented, and on the dup/update paths the row was never re-queued for Phase 2 — with no signal anywhere. `Notify_failure` never covered these paths; it only watches `HasResume`.

**Shape.** Each alert is a plain **sibling** of the patch it watches, wired on `[Failed, TimedOut]` only (never `Succeeded`, so no false alarms). Sibling placement is not stylistic: `Patch_update_attempts` and `Patch_followup_attempts` already sit at **depth 8**, the Power Automate nesting ceiling, so a *child* action there would fail import. Nothing runs after an alert and nothing depends on one, so they cannot alter any existing path and a failing alert is inert. Retries: exponential ×2, same reasoning as `Notify_failure`.

**Effect:** the message is still filed normally — only the row is wrong. The alert names the exact patch action so you can correct the row by hand.

---

### Admin alert — intake-write alerts ×2 (added 2026-08-24)

**Fires:** `Add_row` or `Create_file` fails or times out — the two writes that *create* a candidate rather than update an existing one.

| Alert | Watches | On failure |
|---|---|---|
| `Notify_add_row_failed` | `Add_row` | **no row on either sheet, and no acknowledgment** — `Send_acknowledgment` is gated on `Add_row` succeeding, so it is skipped |
| `Notify_create_file_failed` | `Create_file` | the résumé was never saved, but `Add_row` still runs and always writes `Has Resume = 'Yes'` — so the row claims a file that does not exist |

**Why these exist.** The ×7 alerts above cover every PATCH of an *existing* row. The two writes that create a candidate had no watcher at all, and their failure is strictly worse: a failed patch leaves a row that is merely stale, a failed `Add_row` leaves **no row anywhere**.

It was silent because `Mark_as_read` and `Move_to_processed` hang off **`HasResume`, not off `Add_row`** — inbox cleanup is on a *parallel* branch and files the mail as handled either way. Combined with the correctly-gated acknowledgment, the outcome was: no row, no reply, mail archived as done, run reports **Succeeded**, and nothing said otherwise. The 2026-08-24 mailbox audit found **11 applicants lost this way**, clustered 2–6 June 2026.

**Deliberately not re-gating the cleanup.** Making `Move_to_processed` conditional on `Add_row` would re-introduce the poisoned-message queue block that `fetch_only_unread: false` exists to prevent, and re-marking the message unread afterwards races the mutable Outlook message ID — the same hazard that disables retries on `MoveV2`/`MarkAsRead`. So the alert is the signal, exactly as it is for the seven swallowed patch failures.

**Recovery is in the alert body**, because the message is still sitting in Archive and is recoverable: move it back to the Inbox as unread and the next poll re-runs the full intake. `test_p1.py` asserts that recovery wording is present, and — more importantly — enumerates every `Notify_*` watcher to assert that `Add_row` and `Create_file` are among the watched actions, so a rename or rewire still fails the suite.

Subject: `[Hiring Auto-Reply] APPLICANT LOST - the candidate row was never created (Ref APP-20260604-1545-47D9)`

Subject: `[Hiring Auto-Reply] Candidate row NOT updated - resume update (Ref APP-20260630-1430-A3F9)`

```
Hiring Auto-Reply: a candidate row failed to update.

The incoming email was handled and filed normally, but
Patch_update_attempts (resume update) did not commit. The
Application Updates counter was NOT incremented, and the row
was NOT re-queued for Phase 2 scoring.

  Reference: APP-20260630-1430-A3F9
  From: applicant@example.com
  Subject: Update - APP-20260630-1430-A3F9
  Received: 2026-06-30T14:30:52Z

The run itself reports Succeeded because inbox cleanup completed
as designed — this alert is the only signal. Open run history and
inspect Patch_update_attempts, then correct the row by hand.
```

All seven are suppressed to no-op `Compose` actions along with `Notify_failure` when `email.send_admin_failure_alerts` is `false`.

---

## Row schema

Phase 1 writes intake fields; Phase 2 fills scored fields and flips Status.

| # | Column | Writer | Value |
|---|---|---|---|
| 1 | Application ID | P1 | `APP-YYYYMMDD-HHMM-XXXX` — from `receivedDateTime` plus a 4-char random hex tail so it's unique even if two emails arrive in the same minute. Format is `flow_config.json`'s `appref.time_format`/`appref.hex_length`, not hardcoded. |
| 2 | Received Date | P1 | `body/receivedDateTime` |
| 3 | Last Updated Date | P1 | Same as Received Date on first intake; refreshed on every patched contact (resume update/resend AND text-only follow-up). Guarded (fixed 2026-07-31) to never move backward - `max(existing value, this email's own time)` - since backlog replay can process an older queued reply after a newer one already patched the row. |
| 4 | Category | P2 | Business **department** derived from Suggested Role 1 — one of: `Senior & Executive`, `AI/ML/CV (SIN2)`, `Data Analytics`, `3D/CV/ IoT/ AI Agents (SIN3)`, `Cloud and DevOps`, `Web Team (Full stack/Back end & UI/UX)`, `Graphics`, `Mobile Apps (Android IOS)`, `Business Analytics`, `Supply Chain`, `Finance`, `Cybersecurity and IT Admin`, `Data Center`, `Satellite`, `General`. `Graphics` added 2026-07-15 (design/rendering/gaming roles, split out of Web Team); marketing roles route to `Business Analytics` (also 2026-07-15). **Every scored candidate always gets a department** (an unmatched / role-less candidate falls back to `General`) so the client's Category filter never hides anyone. Blank until P2 scores the row. Config-driven in `role_categories.rules` (config.yaml) — add a JD → no code change needed. |
| 5 | Resume Link | P2 | Calculated column formula displaying the candidate's clickable resume. **Moved to 5th position (right after Category, before Full Name) at client request, 2026-07-24.** |
| 6 | Full Name | P1 (P2 may refine) | Display name before `<`; else email local-part with `.`/`_` → spaces |
| 7 | Email | P1 | Normalized `body/from` — lowercased, and a `"Name" <addr>` wrapper is stripped to just `addr` (changed 2026-07-15; was the raw header value). Applies going forward only — rows written before this change keep whatever raw value they already have. |
| 8 | Phone | P2 | Filled at scoring |
| 9 | Location | P2 | City and state only (e.g. `Tempe, AZ`) — country excluded |
| 10 | Country | P2 | Country name (e.g. `United States`, `India`) — separate from Location |
| 11 | Current Skills | P2 | Comma-separated skills extracted from resume |
| 12 | Education | P2 | Highest degree + field of study, plus school if given (e.g. `B.S. Computer Science, Arizona State University`); blank if the resume doesn't mention any degree/schooling |
| 13 | Looking For Role | P2 | Role preference inferred from resume / email |
| 14 | Suggested Role 1 | P2 | Top role match with score (e.g. `Backend Engineer (72%)`) |
| 15 | Suggested Role 2 | P2 | Second-best role match with score |
| 16 | Suggested Role 3 | P2 | Third-best role match with score; blank if fewer than 3 roles qualify |
| 17 | Mail Subject | P1 | `coalesce(body/subject, '')` |
| 18 | Mail Body | P1 | `bodyPreview` (~255 chars, O365 plain-text; newlines → ` \| `) |
| 19 | Status | P1 -> P2 | `New Email Received` -> `Scored`, `Rejected - Non-USA Location`, or `Rejected - Location Not Confirmed` |
| 20 | Has Resume | P1 | `Yes` — every row has a real CV (invariant) |
| 21 | Original Filename | P1 | Original filename(s) joined with `, ` (no APP-... prefix in this cell) |
| 22 | Application Updates | P1 | `0` on add; `Patch_*` increments on duplicate / update / follow-up |
| 23 | Retry Count | P2 | Bumps when scoring fails. Giving up after max retries moves row to Rejected. |
| 24 | Portfolio 1 | P2 | LinkedIn URL or personal website; `N/A` if not found |
| 25 | Portfolio 2 | P2 | GitHub URL; `N/A` if not found |
| 26 | Portfolio 3 | P2 | Figma / Behance / Dribbble or other creative portfolio; `N/A` if not found |
| 27 | Resume URL | P2 | Raw web URL of the saved resume on SharePoint; automatically hidden in Excel. |
| 28 | Resume Folder Path | P2 | Info-only folder path bucket (e.g. `Candidate_Resumes/2026/July`). |
| 29 | Mail Sent | P1 | Audit timestamp stamped only after a reply email successfully sends. |
| 30 | Info Request Sent | P2 | Audit timestamp for the missing-info nudge (phone/location/skills/education); blank until that email sends, `Nothing missing` if there was nothing to ask for. Main sheet only — never appears on Rejected. |

**Status values:** `New Email Received` - P2 picks these up. `Scored` - P2 done. `Rejected - Non-USA Location` - P2 moved a clear non-USA row to Rejected. `Rejected - Location Not Confirmed` - P2 moved a still-conflicting location row to Rejected only after the candidate already received a current-location request and replied with an updated resume that still did not resolve it.

**Removed columns:** `Is Categorized`, `ATS Score`, `Score Reason`, and `Sentiment` were dropped. `Is Categorized` was redundant with `Status` (every scored row is categorized); the `Suggested Role N (NN%)` percentage already conveys match quality, so ATS/Reason/Sentiment were redundant too. If your live workbook still has these columns, delete them by hand (the flow/worker only stop *writing* them; they can't delete existing columns from the live file).

The **table has 30 columns** (from the workbook template plus schema additions since), but `Add_row` only **writes the 11 P1-owned intake fields** — it does not list the 19 remaining columns at row creation, so they land as blank cells for P2 scoring/resume metadata plus later audit stamps. This matches the genuine tenant PA export pattern and is deliberate: listing every column in the write forces Power Automate to bind that many column-name matches against the live table schema at import time, and any incomplete match dropped the per-column values. Writing only the 11 that carry real values leaves the same blank cells with far less import-binding fragility.

No-CV applications get Email 3 but **no row**. The table only ever holds rows with a real CV. P2 never sees a CV-less candidate.

The saved resume filename in SharePoint is `<FirstLast>_<AppRef>.<ext>` (e.g. `JaneDoe_APP-20260630-1430-A3F9.pdf`) **at P1 intake time** — changed 2026-07-12; was `<AppRef>_<original-filename>`. `FirstLast` is derived from the candidate's name (same cleaning rule as P2's `_get_cleaned_filename_prefix()`: strip apostrophe/hyphen/period/comma, first+last word, or `Candidate` if empty) so the file is readable in the SharePoint folder without opening it. The `Original Filename` column stores only the original filename as sent by the candidate — it is never the saved file's actual name. **P2 renames it again once it has read the CV** (changed 2026-08-04, client request): the final saved name is `<First>_<Last>_<tail>.<ext>` (e.g. `Jane_Doe_A3F9.pdf`) — `<tail>` is the AppRef's own last hyphen-segment, kept specifically so two same-name candidates never collide now that the full AppRef is no longer in the visible name. Category was part of this name until 2026-08-04 and was dropped: it is a P2 verdict, so a re-score that changed it forced a pointless rename and made the file unfindable by name. P1 itself is unaffected by this — it still saves under the intake-time name shown above; only P2 renames it further.

---

## Logging and observability

| What | Where | Retention |
|---|---|---|
| Per-run record | Power Automate built-in run history | 30 days (PA default) |
| New candidate | SharePoint workbook row + saved resume file | Permanent |
| Audit trail per candidate | `Application ID`, `Received Date`, `Status`, `Mail Sent`, `Application Updates` columns | Permanent |
| Main-branch error | `Notify_failure` admin email to `yashv@driverai.io` | Email inbox |
| Candidate-row PATCH failure | 7 patch-failure alerts → same admin address | Email inbox |
| Intake write failure (no row / no resume file) | `Notify_add_row_failed`, `Notify_create_file_failed` → same admin address | Email inbox |
| Inbox poll failure | `Notify_poll_failure` admin email | Email inbox |
| PA-level failure backup | `flowFailureAlertSubscribed: true` | PA alert email |
| Metrics (emails/day, spam blocked, reply types) | **None** | — |
| Success notification to admin | **None** | — |

There is no custom log table and no counters. To count applicants for a period, filter the workbook by `Received Date`. To see spam volume, count the flow's terminated runs in PA run history.

**Two blind spots to be aware of.**

1. **Junked mail leaves no record anywhere but the Junk Email folder.** A message caught by any of the three spam gates creates no row, increments no counter, and raises no alert — by design, since the whole point is to not process it. The consequence is that a *false positive* is invisible unless someone opens Junk. The mitigation is process, not code: periodically replay the phrase lists over recent mail (as was done on 2026-07-31 across 278 messages, which is how `marketing@` and five bare `bad_subjects` tokens were found to be junking real applicants). Re-run that replay after **any** change to the phrase lists.

2. **Run history is the only step-level log, and it is not permanent.** Once it ages out, a failed run's exact error is gone. The admin alert emails are the durable record, which is why they are worth keeping enabled and now retry on a transient send failure.

---

## Config reference

Edit `flow/flow_config.json`, run `python flow/build_zip.py`, re-import.

| Key | Current value | Notes |
|---|---|---|
| `flow_name` | `DriverAI_HiringAgent_P1_AutoReply` | Package + flow display name |
| `email.trigger_mailbox` | `apply@driverai.io` | Watched inbox and reply "from" |
| `email.admin_email` | `yashv@driverai.io` | Failure alert recipient |
| `trigger.interval_min` | `1` | Poll frequency — only takes effect after re-import. `1` is 1,440 runs/day against the owner plan; `2` was 720/day. Set to `2` on 2026-08-04 for quota headroom, returned to `1` on 2026-08-14 on owner instruction — watch the daily Power Platform request budget |
| `trigger.unread_per_run` | `1` | One unread Inbox message per sequential run |
| `business_rules.duplicate_check_days` | `90` | Re-application window in days |
| `business_rules.duplicate_notice_max` | `5` | Overall cap for Email 2 (counter ceiling, not all 5 get an email) |
| `business_rules.update_resume_max` | `5` | Overall cap for Email 5 |
| `business_rules.followup_reply_max` | `5` | Overall cap for Email 6 |
| `business_rules.reply_cap` | `3` | How many contacts actually get an email (3rd = final-notice); contacts between this and the overall cap still update the row/resume but get NO reply |
| `sharepoint.site` | `https://led1234567.sharepoint.com/sites/CandidateList_HiringAgent` | SharePoint site URL |
| `sharepoint.resumes_folder` | `/Shared Documents/Candidate_Resumes` | Root folder for saved resumes |
| `sharepoint.documents_library_id` | `3600a6b3-142f-45f1-bce0-cfcb285c69e7` | Documents library GUID — required by `CreateNewFolder`'s `table` param (workbook + dated resume folders both use this) |
| `sharepoint.dated_resume_subfolders` | `true` | `true` = file into `<Year>/<Month>/`; `false` = flat. Every dated folder is explicitly guaranteed to exist via `CreateNewFolder` before saving — never assumed |
| `excel.source` | `led1234567.sharepoint.com,...` | Site ID triple for Excel connector |
| `excel.drive` | `b!Mba...` | Document library drive ID |
| `excel.file` | `/Master_Files/Sharepoint_Master_File.xlsx` | Single workbook path (no `<Year>` subfolder) — one queue for all years |
| `excel.table` | `HiringAgent_P1_Candidates` | Excel table name |
| `inbox_tidy.enabled` | `true` | Move each processed email out of the Inbox by category |
| `inbox_tidy.destination` | `Archive` | Legit application mail → marked read + moved here. Well-known folder name (capitalized) or a mail-folder ID |
| `inbox_tidy.spam_destination` | `Junk Email` | Spam/junk → moved here, left **unread**. Well-known folder name (capitalized) or a mail-folder ID |
| `inbox_tidy.failure_destination` | `Archive` | Failed messages move here unread; avoids the nested custom-folder `MoveV2` runtime failure |
| `appref.date_format` | `yyyyMMdd` | Date part of the reference token |
| `appref.detect_pattern` | `app-20` | Substring that flags a quoted ref in a reply; decade-proof (2000–2099) |
| `appref.time_format` | `HHmm` | Time part of the reference token |
| `appref.hex_length` | `4` | Length of the random hex tail (`toUpper(substring(guid(),0,4))`) |
| `year_separator.enabled` | `true` | Insert one truly blank row before the first candidate row of each new calendar year. The first-ever row uses the template's permanent blank row below the header, so a fresh workbook never gets two blanks (see §1A in `p1_detailed_summary.md`) |
| `month_separator.enabled` | `true` | Same, per calendar **month**. At a January boundary the year separator wins, so exactly **one** blank row lands, never two stacked (see §1A) |
| `email.send_applicant_emails` | `false` | `false` suppresses all six applicant-facing sends while trigger/filter/save/row/tidy processing remains active; suppression is recorded in `Mail Sent` |
| `email.send_admin_failure_alerts` | `true` | Keeps all 11 admin alerts enabled even while applicant mail is disabled — `Notify_failure`, `Notify_poll_failure`, the 7 patch-failure alerts, and the 2 intake-write alerts |
| `spam_filters.bad_senders` | 32 fragments | Gate 1 list |
| `spam_filters.bad_subjects` | 24 strings | Gate 2 list |
| `spam_filters.spam_phrases` | 30 scam phrases | Gate 3 group 1 |
| `spam_filters.offensive_phrases` | 36 phrases | Gate 3 group 2 |
| `spam_filters.malware_phrases` | 20 extensions/macro phrases | Gate 3 group 3 — always blocks |
| `spam_filters.link_shortener_phrases` | 12 URL shorteners | Gate 3 group 5 — blocks only if NO valid resume is attached (`HasValidResumeEarly`) |
| `spam_filters.foreign_scam_phrases` | 25 phrases (9 languages) | Gate 3 group 4 |
| `spam_filters.app_keywords` | 21 keywords | Triggers Email 3 when no attachment present |

**To use a custom archive folder instead of the built-in Archive:**
1. Graph Explorer: `GET /v1.0/users/apply@driverai.io/mailFolders?$filter=displayName eq 'Processed'&$select=id`
2. Copy the `id` value.
3. Set `inbox_tidy.destination` to that id.
4. Rebuild and re-import.

**What is config-driven vs in code:** all values above are config-driven (edit `flow_config.json`, no code change). The 6 email subjects and HTML bodies (see `docs/MAIL_REPLIES.md`), the row field map (`_ROW_ITEM`), and the flow structure (gates, branches, run-after wiring) are hardcoded Python in `build_zip.py`; `_pristine_send_actions.json` only restores those 7 actions' connector shape as a self-heal, it does not supply their text.

---

## Test / reset the PA trigger

`trigger_reset.py` marks archived emails as unread and moves them back to Inbox.
The scheduled unread poll then processes them one at a time on successive 1-minute runs.

```
# Preview — no changes:
python HiringAgent_P1/trigger_reset.py --dry-run

# Re-trigger last 31 days (default):
python HiringAgent_P1/trigger_reset.py

# Custom window:
python HiringAgent_P1/trigger_reset.py --days 7
python HiringAgent_P1/trigger_reset.py --days 90

# ALL emails in Archive — no date limit (re-run June mails, old batches, everything):
python HiringAgent_P1/trigger_reset.py --all
python HiringAgent_P1/trigger_reset.py --all --dry-run
```

Credentials are loaded from `HiringAgent_P2/.env` (the same Entra app as the P2 worker).  
**Requires `Mail.ReadWrite` (Application) permission admin-consented** on that app registration —
in addition to the `Sites.ReadWrite.All` / `Files.ReadWrite.All` the P2 worker already uses.

> **Resolved as of 2026-07-15** (originally found 2026-07-04, documented below for history). The 403
> is gone: `python trigger_reset.py --dry-run` ran clean this session, returned 26 real Archive emails
> with no auth error — the Exchange Application Access Policy fix (or equivalent) has evidently landed
> since. Original note, kept for context: every Graph call to `apply@driverai.io`'s mail folders was
> failing with `403 ErrorAccessDenied` — the signature of an Exchange Online **Application Access
> Policy** not scoped to this app, not a missing Graph API permission (the same token read/wrote
> SharePoint fine throughout, so app-only auth itself was never the issue). This never affected the
> live P1 flow either way — Power Automate uses its own Office 365 connection, a separate auth path.
> If it ever reappears, a tenant admin needs Exchange Online PowerShell: `Get-ApplicationAccessPolicy`
> to check current scope, `New-ApplicationAccessPolicy` / `Set-ApplicationAccessPolicy` to grant this
> app's `CLIENT_ID` access to `apply@driverai.io`.

| Flag | Default | Effect |
|---|---|---|
| `--days N` | `31` | Look back N days from today |
| `--all` | off | Reset ALL emails in Archive regardless of age (overrides `--days`) |
| `--mailbox addr` | `apply@driverai.io` | Mailbox to reset |
| `--dry-run` | off | List emails without making any changes |

**What happens after the reset:** the PA flow processes each email exactly as if it just arrived — spam gates, duplicate check (within 90 days), reply emails, SharePoint row. Emails from the same sender re-submitting within 90 days will get the duplicate notice, not a new row.

---

## Hand-off to Phase 2

Phase 1 writes rows with `Status = "New Email Received"`. Phase 2 (`../HiringAgent_P2/`) polls for those rows, reads each resume, fills the P2-owned scoring fields, applies USA-only geo-filter, and flips `Status` to `Scored` or moves the row to Rejected as `Rejected - Non-USA Location` / `Rejected - Location Not Confirmed`. The two phases never call each other - the SharePoint workbook is the only contract.

**The table has 30 columns; `Add_row` writes only the 11 P1-owned ones.** The other 19 are simply not listed in the write, so they land as blank cells — identical outcome to writing them as empty strings, but it matches the genuine tenant export format and avoids the import-time schema-binding fragility that was dropping fields in the designer.

**P1 intake values (11 fields — the ones Add_row writes):** Application ID, Received Date, Last Updated Date, Full Name, Email, Mail Subject, Mail Body, Status (`New Email Received`), Has Resume (`Yes`), Original Filename, Application Updates (`0`).

**Blank/calculated/audit cells left after Add_row (19 fields — NOT written by Add_row):** Category, Phone, Location, Country, Current Skills, Education, Looking For Role, Suggested Role 1/2/3, Portfolio 1/2/3, Retry Count, Resume URL, Resume Link, Resume Folder Path, Info Request Sent, Mail Sent. P2 fills the scoring/profile, retry, resume-link, and info-request fields; P1 stamps `Mail Sent` only after a reply send path succeeds or is test-suppressed.

**P2 fills its blanks:** scores the resume, writes the scoring/profile and resume metadata columns, handles missing-info/current-location requests, moves clear non-USA candidates to Rejected, and flips Status to `Scored` when scoring completes. If a candidate already received a current-location request and then sends an updated resume that is still conflicting or unclear, P2 moves that row to Rejected as `Rejected - Location Not Confirmed`. P2's duplicate cleanup deliberately skips rows that are still `New Email Received`, blank, or `Needs Review - Unreadable Resume`, so a newly arrived resume update is scored before any older duplicate row can be merged away.

Phase 2's SharePoint settings must match Phase 1's names exactly: `SHAREPOINT_HOSTNAME` +
`SHAREPOINT_SITE_PATH` (same site), `SHAREPOINT_TABLE` (`HiringAgent_P1_Candidates`), and
`SHAREPOINT_RESUMES_FOLDER` (`/Candidate_Resumes`). Phase 2 reconstructs the same dated
`<Year>/<Month>/` resume path from `Received Date` and also falls back to the flat root
folder for older files. For a candidate P2 rejects, it files/moves that resume into a single
`<Year>/Rejected/` folder (changed 2026-07-15 from one `Rejected/` subfolder per month — see
the P2 README).

---

## SharePoint setup & structure — complete reference

**This is the single source of truth for the SharePoint side.** Everything the flow reads/writes and every exact name lives here.

### Exact names & locations (must match `flow_config.json`)

| What | Value | Notes |
|---|---|---|
| Site URL | `https://led1234567.sharepoint.com/sites/CandidateList_HiringAgent` | `sharepoint.site` |
| Document library | **Documents** (shows as `Shared Documents` in the URL) | The default library |
| Documents library GUID | `3600a6b3-142f-45f1-bce0-cfcb285c69e7` | `sharepoint.documents_library_id` — used by `CreateNewFolder` |
| Resumes root folder | `/Shared Documents/Candidate_Resumes` | `sharepoint.resumes_folder` |
| Workbook file | `Sharepoint_Master_File.xlsx` | sits in `Master_Files` (one file, all years) |
| Full workbook path | `/Shared Documents/Master_Files/Sharepoint_Master_File.xlsx` | this is the only workbook P1 uses |
| Worksheet (tab) name | **`CandidateList`** | do not rename |
| Table name | **`HiringAgent_P1_Candidates`** | `excel.table` — 30 columns; do not rename |
| Rejected worksheet | **`Rejected`** (table `RejectedCandidates`) | **already shipped in `P1_Templates/HiringAgent_P1_CandidateList.xlsx`** — uploading that template gives you both sheets on day one, P2 only ever writes rows into it |
| Resume files | `Candidate_Resumes/<Year>/<Month>/<FirstLast>_APP-YYYYMMDD-HHMM-XXXX.pdf` at intake | dated subfolders auto-created by the flow; P2 renames to `<First>_<Last>_<tail>.ext` once scored, and moves a rejected candidate's file into `<Year>/Rejected/` (one folder per year, not per month) |
| Trigger mailbox | `apply@driverai.io` | `email.trigger_mailbox` |
| Applicant reply "from" | `apply@driverai.io` | needs Send-As on the connection owner |
| Admin alert "to" | `yashv@driverai.io` | `email.admin_email` |
| Legit mail moves to | **Archive** folder (marked read) | `inbox_tidy.destination` |
| Spam/junk moves to | **Junk Email** folder (left unread) | `inbox_tidy.spam_destination` |
| Failed mail moves to | **Archive** (left unread) | `inbox_tidy.failure_destination` |

### SharePoint runtime tree

```
CandidateList_HiringAgent  (site)
└── Shared Documents  (Documents library)
    ├── Master_Files/
    │   ├── Sharepoint_Master_File.xlsx                 ← THE workbook (one for all years)
    │   │     ├── sheet "CandidateList"  → table "HiringAgent_P1_Candidates" (30 cols)
    │   │     └── sheet "Rejected"       → table "RejectedCandidates", shipped in the template
    │   └── Candidate_List_Results.xlsx                 ← P2 final client-facing result
    └── Candidate_Resumes/                          ← the resumes root (resumes_folder)
        ├── 2026/June/Jane_Doe_A1B2.pdf    ← saved resumes, dated Year/Month
        ├── 2026/July/JohnSmith_CloudAndDevOps_D4E5.docx
        ├── 2026/Rejected/JaneRejected_General_701C.pdf ← ONE Rejected folder per year
        └── <Year>/<Month>/... , <Year>/Rejected/...
```
> The `SharePoint_Master_Template` folder (if you created one) is a **master-copy holder only** — the flow never reads it. The live workbook must be the one inside `Master_Files/`.

### What YOU set up once (manual) vs what the automation creates

| Created how | Item |
|---|---|
| **You, once** | The `Candidate_Resumes`, `Master_Files`, and `SharePoint_Master_Template` folders. Put the live `Sharepoint_Master_File.xlsx` inside `Master_Files/`, and keep a blank reference copy in `SharePoint_Master_Template/`. Grant Send-As on `apply@driverai.io`. |
| **P1, automatically** | The dated resume subfolders (`<Year>/<Month>`) + the saved resume files + candidate rows in the table. |
| **P2, automatically** | The `Category`/scored columns; renames each resume to `<First>_<Last>_<tail>.ext` once scored; writes rows into the pre-created `Rejected` worksheet/table; and creates/uses the year-level `Rejected/` resume folder (moves a rejected candidate's file(s) into it, one folder per year). |
| **NOT auto** | The workbook file itself — P1 never re-creates or overwrites it (removed 2026-07-04). If it's ever missing, `Notify_failure` alerts the admin to restore it from the template. |

### Category dropdown

The committed template (`P1_Templates/HiringAgent_P1_CandidateList.xlsx`) has a data-validation dropdown on the **Category** column listing the DriverAI departments (e.g., `Senior & Executive`, `AI/ML/CV (SIN2)`, `Data Analytics`, `3D/CV/ IoT/ AI Agents (SIN3)`, `Cloud and DevOps`, `Web Team (Full stack/Back end & UI/UX)`, `Graphics`, `Mobile Apps (Android IOS)`, `Business Analytics`, `Supply Chain`, `Finance`, `Cybersecurity and IT Admin`, `Data Center`, `Satellite`, `General`). P1 leaves Category blank; **P2 fills it** when it scores. The column's *filter* dropdown fills with real values only as candidates get scored (Excel filters only list values that exist). **Note:** the template's dropdown *list itself* (the manual-entry validation, distinct from the auto-filled filter) was not re-generated to add `Graphics` as of 2026-07-15 — P2 writes it fine via Graph (API writes bypass UI dropdown validation), but a human manually editing the Category cell in Excel won't see `Graphics` as a pickable option in the existing template/live workbook until that dropdown list is regenerated.

### Local repo structure

```
HiringAgent_P1/
├── flow/
│   ├── flow_config.json               ← EDIT THIS (single source of truth for CODE values)
│   ├── build_zip.py                   ← run after any config edit
│   ├── _pristine_send_actions.json    ← real (non-suppressed) shapes of the 7 mail actions
│   ├── definition.json                ← GENERATED — do not edit
│   ├── manifest.json / apisMap.json   ← GENERATED
│   └── DriverAI-Hiring-AutoReply-apply.zip  ← import into Power Automate (Update mode)
├── docs/
│   ├── README.md                 ← this file (all human setup + reference lives here)
│   ├── MAIL_REPLIES.md           ← verbatim email copy for all 6 applicant replies
│   └── p1_detailed_summary.md    ← architectural deep-dive
├── P1_Templates/
│   └── HiringAgent_P1_CandidateList.xlsx   ← upload to SharePoint once (has the category dropdown)
├── archive/                      ← historical records, NOT part of the deployable zip
│   ├── audits/                   ← audit_p1_live_state.py output (.md + .csv per run)
│   └── backups/                  ← full-folder snapshot zips
├── audit_p1_live_state.py        ← read-only: live SharePoint + mailbox audit
├── test_p1.py                    ← structural test suite (run after any build)
├── trigger_reset.py              ← test utility: move Archive→Inbox to re-fire PA
└── bulk_move_tool.py             ← bulk mailbox move utility
```



### Where each kind of information lives (doc map)



| Doc | Holds | Edit it for |

|---|---|---|

| **docs/README.md** (this file) | All human setup, deployment, SharePoint structure, email copy, config reference | Setup steps, structure, how-to |

| **low/flow_config.json** | Code-driven values only (mailbox, interval, paths, spam lists, caps, folder names) | Any value the flow uses — then run uild_zip.py |

| **docs/MAIL_REPLIES.md** | Verbatim email copy for all 6 applicant replies | Email content changes |

| **docs/p1_detailed_summary.md** | Architectural deep-dive | Structural or design changes |



Rule of thumb: **humans read the README; the build reads `flow_config.json`.** Anything a person needs to know goes in the README; anything the code consumes goes in `flow_config.json`.
