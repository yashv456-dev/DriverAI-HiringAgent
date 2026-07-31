# P1 — Mail Replies Reference

All 6 applicant-facing emails P1 can send, verbatim, with exact trigger conditions. **All 6 — including Acknowledgment and Request CV — are hardcoded HTML Python string literals in `flow/build_zip.py` (`_set_email()`, ~lines 1218-1341).** `flow/_pristine_send_actions.json` is not the source of truth — `build_zip.py` reads it only to restore each of these 7 actions' connector *shape* (host/apiId/operationId/authentication) as a self-heal against test-mode corruption; the actual Subject/Body text is written by `_set_email()`'s hardcoded strings, which unconditionally overwrite whatever the pristine JSON or base zip held. (As of 2026-07-24 the two happen to hold matching text, verified byte-for-byte, so there's no visible drift today — but `build_zip.py` is what actually governs deployed content.) `build_zip.py` never reads P2's `config.yaml` (confirmed: zero references to it anywhere in the file) — none of P1's email content is config-driven the way P2's is. All 6 send from `apply@driverai.io` via `SharedMailboxSendEmailV2`, need Send-As granted on that mailbox, and share one visual style (Arial/14px, a `<hr>` + 12px grey "auto-generated, mailbox not monitored" footer).

> **Correction (2026-07-15):** an earlier pass in this doc's history showed different Acknowledgment/Request CV content, sourced from P2's `config.yaml → email_templates.acknowledgment`/`request_cv`. Those two config.yaml entries look plausible (same visual style, similar purpose) but are **dead config** — confirmed unread by both `build_zip.py` (P1) and every P2 Python file (grepped `EMAIL_TEMPLATES["acknowledgment"]`/`["request_cv"]` — zero hits anywhere). They were never the live content. The versions below match what's hardcoded in `build_zip.py` (see correction above for exact sourcing).

State as of **2026-07-24**.

> ⚠️ **Currently suppressed (set 2026-07-31).** `flow/flow_config.json` has `test_mode.suppress_emails: true`, so the build replaces all 6 applicant emails below — plus `Notify_failure` — with no-op `Compose` actions. **None of this content is being delivered right now.** The rest of the flow is untouched: resumes still save, rows are still added and patched, inbox tidy still runs. Rows handled while suppressed are stamped `TEST-MODE (suppressed) <timestamp>` instead of `Sent <timestamp>`. This is for a historical replay of the existing mailbox; set `suppress_emails` back to `false`, run `python flow/build_zip.py`, and re-import the zip to restore delivery. P2's equivalent switch is separate (`HIRING_SUPPRESS_EMAILS` in its `.env`).

---

## The cap system (shared by Duplicate Notice, Update Ack, Noted Reply)

One counter — `Application Updates` — is shared across all three re-contact paths. Two independent thresholds:

- **`reply_cap: 3`** — contacts 1-3 still get an actual reply email (the 3rd swaps to a final-notice closing paragraph, still sent).
- **`duplicate_notice_max` / `update_resume_max` / `followup_reply_max: 5`** (per-path) — contacts 4-5 still update the resume file and row silently, but get **no email at all**. Contact 6+ (past the overall max) is fully silent: no email, no file update, no row change.

The final-notice swap is a single PA `if(...)` expression evaluated against the row's current `Application Updates` value (`>= 2`, i.e. this is the 3rd contact) — both branches are complete HTML, never a partial/broken string.

---

## 1. Acknowledgment

**Fires when:** new applicant, PDF/DOCX attached (or a CV-less applicant using body keywords who *then* attaches — same email either way once a valid resume exists).
**Row:** created, `Status: New Email Received`.
**Tidy:** Archive, marked read.
**Source:** hardcoded, `Send_acknowledgment` (`_pristine_send_actions.json`).

> **Subject:** Application Received - DriverAI (Ref: `{AppRef}`)
>
> Hello,
> Thanks for applying to **DriverAI**. Your application has been received and is under review.
> Reference number: **`{AppRef}`**
> If you are selected, a member of our team will contact you to discuss next steps. You do not need to reply unless you are updating your resume. To do that, send a new or reply email to *apply@driverai.io* with the subject "Update - `{AppRef}`" and attach the new file.
> Good luck on your next journey.
> Warm regards, **The DriverAI Recruiting Team**
> *This is an auto-generated email and this mailbox is not monitored.*

---

## 2. Request CV

**Fires when:** no attachment, but subject/body contains an application keyword (resume, cv, apply, hiring, ...).
**Row:** **none created** — the table only ever holds rows with a real CV (CV-only invariant).
**Tidy:** Archive, marked read.
**Source:** hardcoded, `Send_CV_request` (`_pristine_send_actions.json`). Fixed subject every time — no ref/original-subject variant (unlike the other ref-bearing emails, since no application exists yet to have a ref).

> **Subject:** Please Attach Your Resume - DriverAI
>
> Hello,
> Thanks for your interest in **DriverAI**. We didn't see a resume attached, so your application isn't complete yet.
> Please reply with your resume attached in **PDF or Word (.docx)** format. Once we have it, your application goes under review, and if you are selected, a member of our team will contact you to discuss next steps.
> Good luck on your next journey.
> Warm regards, **The DriverAI Recruiting Team**
> *This is an auto-generated email and this mailbox is not monitored.*

---

## 3. Duplicate Notice

**Fires when:** resend within the 90-day window, **no** reference quoted, under `reply_cap`.
**Row:** existing row **re-queued** (`Status → New Email Received`), resume saved under the same `<FirstLast>_<FullAppID>` name P1 always uses — a true overwrite pre-scoring; post-scoring, P2's next pass reconciles it against the already-renamed file (see the `_download_resume_text` note in Email 5 below / the P1 README's resume-filenames section).
**Tidy:** Archive, marked read.
**Source:** hardcoded, `Send_duplicate_notice`.

> **Subject:** Your DriverAI Application Is Already On File (Ref: `{app_id}`)
>
> Hello,
> Your application is already on file with **DriverAI** and under review, so there's no need to resubmit.
> *(under reply_cap)* If you are selected, a member of our team will contact you to discuss next steps. You do not need to reply unless you are updating your resume. To do that, send a new or reply email to *apply@driverai.io* with the subject "Update - `{app_id}`" and attach the new file.
> *(3rd contact — final notice instead)* This is our final automated reply regarding this application. If there is a match, a member of our team will contact you directly. You do not need to reply, and you are welcome to apply again after 90 days.
> Good luck on your next journey.
> Warm regards, **The DriverAI Recruiting Team**
> *This is an auto-generated email and this mailbox is not monitored.*

---

## 4. Wrong Format

**Fires when:** an attachment exists but isn't `.pdf`/`.docx` (image, `.txt`, `.zip`, ...).
**Row:** none created.
**Tidy:** Archive, marked read.
**Source:** hardcoded, `Send_wrong_format`. No caps, no dynamic ref — fires the same every time.

> **Subject:** Please Resend Your Resume as PDF or Word - DriverAI
>
> Hello,
> Thanks for your interest in **DriverAI**. We received your submission, but we couldn't open the attached file.
> Please reply with your resume as a **PDF (.pdf)** or **Word (.docx)** file and we'll process it right away. If you are selected, a member of our team will contact you to discuss next steps.
> Good luck on your next journey.
> Warm regards, **The DriverAI Recruiting Team**
> *This is an auto-generated email and this mailbox is not monitored.*

---

## 5. Update Ack

**Fires when:** reply quotes the reference AND attaches a new resume, under `reply_cap`.
**Row:** re-queued (`Status → New Email Received`), resume saved under the same `<FirstLast>_<FullAppID>` name P1 always uses. **This is the only reply path P2 will actually re-score** — the new file gets re-read on the next scoring pass. If the row's never been scored before, this is a true overwrite. If it's already been scored once, P2 has since renamed the file to `<FirstLast>_<Category>_<tail>`, so this write lands as a fresh file under the old name rather than replacing it directly — P2's re-score pass (`_download_resume_text`, fixed 2026-07-15) is what reconciles the two, always preferring the fresh `<FirstLast>_<FullAppID>` write over the stale renamed one and cleaning up the leftover afterward, so the candidate still ends up with exactly one resume file on disk.
**Tidy:** Archive, marked read.
**Source:** hardcoded, `Send_update_ack`.

> **Subject:** Updated Resume Received - DriverAI (Ref: `{app_id}`)
>
> Hello,
> Thanks for sending your updated resume. Your application now reflects the latest version and our team will review it.
> *(under reply_cap)* If you are selected, a member of our team will contact you to discuss next steps. You do not need to reply unless you are updating your resume again. To do that, send a new or reply email to *apply@driverai.io* with the subject "Update - `{app_id}`" and attach the new file.
> *(3rd contact — final notice instead, same swap as Duplicate Notice)*
> Good luck on your next journey.
> Warm regards, **The DriverAI Recruiting Team**
> *This is an auto-generated email and this mailbox is not monitored.*

---

## 6. Noted Reply

**Fires when:** reply quotes the reference, **no** attachment, under `reply_cap`.
**Row:** `Application Updates` incremented, `Mail Sent` stamped — **`Status` and `Mail Body` are never touched.** Confirmed by tracing `_make_patch_item` in `build_zip.py`: its Excel patch item writes only those two keys. Whatever the candidate actually typed in this reply is acknowledged but **never stored or acted on anywhere** — it does not get re-scored, and the text itself is not captured.
**Tidy:** Archive, marked read.
**Source:** hardcoded, `Send_noted_reply`.

> **Subject:** Message Received - DriverAI (Ref: `{app_id}`)
>
> Hello,
> Thanks for following up. Your message has been noted alongside your application.
> *(under reply_cap)* If you are selected, a member of our team will contact you to discuss next steps. You do not need to reply unless you are updating your resume. To send one, reply with the subject "Update - `{app_id}`" and attach the file (PDF or Word).
> *(3rd contact — final notice instead, same swap as above)*
> Good luck on your next journey.
> Warm regards, **The DriverAI Recruiting Team**
> *This is an auto-generated email and this mailbox is not monitored.*

**Practical consequence:** if you ever want a candidate's reply to actually change their row (e.g. answering P2's missing-info nudge with a phone number), it must arrive as an **Update** (ref quoted + resume attached) — a **Noted Reply** (ref quoted, text only) is a dead end by design. This is exactly why P2's `missing_info` template (see the P2 mail doc / Mail Routing Manifest artifact) asks for a resume re-attachment rather than a typed answer.

---

## Quick reference

| # | Name | Trigger | New/re-queued row? | Caps apply? |
|---|---|---|---|---|
| 1 | Acknowledgment | new applicant + valid resume | New row | No |
| 2 | Duplicate Notice | resend, no ref, <90d | Re-queued | Yes (reply_cap 3 / max 5) |
| 3 | Request CV | no attachment, has keywords | No row | No |
| 4 | Wrong Format | attachment, wrong type | No row | No |
| 5 | Update Ack | ref quoted + new resume | Re-queued, **will re-score** | Yes (reply_cap 3 / max 5) |
| 6 | Noted Reply | ref quoted, no attachment | Untouched (counter only) | Yes (reply_cap 3 / max 5) |

Everything that clears the 3 spam gates and isn't one of the 6 above (a casual no-attachment/no-keyword email) gets **no reply, no row**, and still moves to Archive/read — it's simply not treated as an application.
