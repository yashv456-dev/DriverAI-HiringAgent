# P1 Reply Content - Verified September 8, 2026

## 6. Complete applicant reply content

Source: the six enabled actions in the supplied ZIP, checked against section 15 of flow/build_zip.py. P2 config.yaml does not supply P1's reply copy. Text below is the rendered wording with {AppRef} substituted for the dynamic reference; HTML styling is omitted. These are existing templates, not newly deployed messages.

All six use SharedMailboxSendEmailV2: From apply@driverai.io; To the incoming CurrentEmail.from; BCC yashv@driverai.io. They are new outbound messages with reference-bearing subjects, not connector reply-in-thread operations. No explicit Reply-To or original-message thread identifier is set in these actions. The imported connection's actual permissions and sender behavior require tenant verification.

The local suppressed build replaces these actions with Compose and writes a distinct suppression marker. The supplied ZIP instead enables all six sends. Its BCC does not redirect the primary recipient to the tester.

### 6.1 Application acknowledgment

Action: `Send_acknowledgment`.

Trigger and data effect: New applicant with an accepted resume; send follows successful Add_row. A new intake row is created.

**Subject:** Application Received - DriverAI (Ref: {AppRef})

> Hello,
>
> Thanks for applying to DriverAI. Your application has been received and is under review.
>
> Reference number: {AppRef}
>
> If you are selected, a member of our team will contact you to discuss next steps. You do not need to reply unless you are updating your resume. To do that, send a new or reply email to apply@driverai.io with the subject "Update - {AppRef}" and attach the new file.
>
> Good luck on your next journey.
>
> Warm regards,
> The DriverAI Recruiting Team
>
> This is an auto-generated email and this mailbox is not monitored.

### 6.2 Duplicate notice

Action: `Send_duplicate_notice`.

Trigger and data effect: Existing sender inside the 90-day window, no reference route, accepted resume, and shared counter below reply_cap. Row is requeued and counter advances.

**Subject:** Your DriverAI Application Is Already On File (Ref: {AppRef})

> Hello,
>
> Your application is already on file with DriverAI and under review, so there's no need to resubmit.
>
> If you are selected, a member of our team will contact you to discuss next steps. You do not need to reply unless you are updating your resume. To do that, send a new or reply email to apply@driverai.io with the subject "Update - {AppRef}" and attach the new file.
>
> Good luck on your next journey.
>
> Warm regards,
> The DriverAI Recruiting Team
>
> This is an auto-generated email and this mailbox is not monitored.

### 6.3 Missing resume

Action: `Send_CV_request`.

Trigger and data effect: Application keywords but no accepted attachment in the missing-CV path. No candidate row is created.

**Subject:** Please Attach Your Resume - DriverAI

> Hello,
>
> Thanks for your interest in DriverAI. We didn't see a resume attached, so your application isn't complete yet.
>
> Please reply with your resume attached in PDF or Word (.docx) format. Once we have it, your application goes under review, and if you are selected, a member of our team will contact you to discuss next steps.
>
> Good luck on your next journey.
>
> Warm regards,
> The DriverAI Recruiting Team
>
> This is an auto-generated email and this mailbox is not monitored.

### 6.4 Unsupported attachment

Action: `Send_wrong_format`.

Trigger and data effect: Attachments fail the accepted-resume selection path. No candidate row is created. This is an extension/selection check, not proof the file was opened.

**Subject:** Please Resend Your Resume as PDF or Word - DriverAI

> Hello,
>
> Thanks for your interest in DriverAI. We received your submission, but we couldn't open the attached file.
>
> Please reply with your resume as a PDF (.pdf) or Word (.docx) file and we'll process it right away. If you are selected, a member of our team will contact you to discuss next steps.
>
> Good luck on your next journey.
>
> Warm regards,
> The DriverAI Recruiting Team
>
> This is an auto-generated email and this mailbox is not monitored.

### 6.5 Updated resume acknowledgment

Action: `Send_update_ack`.

Trigger and data effect: Reference-bearing mail resolves to an existing sender/application and includes an accepted resume; shared counter below reply_cap. See the upload-failure finding below.

**Subject:** Updated Resume Received - DriverAI (Ref: {AppRef})

> Hello,
>
> Thanks for sending your updated resume. Your application now reflects the latest version and our team will review it.
>
> If you are selected, a member of our team will contact you to discuss next steps. You do not need to reply unless you are updating your resume again. To do that, send a new or reply email to apply@driverai.io with the subject "Update - {AppRef}" and attach the new file.
>
> Good luck on your next journey.
>
> Warm regards,
> The DriverAI Recruiting Team
>
> This is an auto-generated email and this mailbox is not monitored.

### 6.6 Text-only follow-up

Action: `Send_noted_reply`.

Trigger and data effect: Reference-bearing mail resolves to an existing sender/application without an accepted resume; shared counter below reply_cap. Counter and Last Updated Date change; reply body is not added to the candidate row.

**Subject:** Message Received - DriverAI (Ref: {AppRef})

> Hello,
>
> Thanks for following up. Your message has been noted alongside your application.
>
> If you are selected, a member of our team will contact you to discuss next steps. You do not need to reply unless you are updating your resume. To send one, reply with the subject "Update - {AppRef}" and attach the file (PDF or Word).
>
> Good luck on your next journey.
>
> Warm regards,
> The DriverAI Recruiting Team
>
> This is an auto-generated email and this mailbox is not monitored.

### 6.7 Final notice, caps and copy limitations

For Duplicate notice, Updated resume acknowledgment and Text-only follow-up, the normal paragraph beginning "If you are selected" is replaced on the third counted follow-up by:

> This is our final automated reply regarding this application. If there is a match, a member of our team will contact you directly. You do not need to reply, and you are welcome to apply again after 90 days.

All three paths share Application Updates. Existing count 0 or 1 uses normal copy; count 2 uses final notice; counts 3 and 4 process silently; count 5 or greater stops the capped update/follow-up path. Suppression and unsuccessful sends can still consume this shared contact counter. It is not a count of successfully delivered emails. The initial acknowledgment, missing-CV request and wrong-format request do not share this limiter.

The current footer says the mailbox is not monitored while the body invites replies. Proposed replacement, not deployed: "This is an automated confirmation. Replies are processed by our hiring system; a recruiter will contact you if further information is needed."

The follow-up wording says a message is noted alongside the application, but the candidate record does not retain its body. Either persist the reply and make it available to review, or narrow the copy to "We received your follow-up email." Do not promise a review time: none is implemented or stated in the current templates.

### 6.8 Admin mail and delivery audit contract

The package has 13 Notify_* failure alerts: polling, general intake failure, four intake-write failures and seven row-patch failures. It also contains Notify_ignored_mail, a low-importance informational alert for non-application mail, for 14 Notify_* actions total. A success-colored flow run alone is not proof of successful intake or contact.

Mail Sent is a last-status cell, not a delivery ledger. Successful connector execution is not proof of inbox delivery; bounced mail and ambiguous timeouts need reconciliation. Required future ledger fields: message key, application ID, template version, intended recipient, actual recipient, test/live mode, queued/sent/failed/unknown status, provider message ID, attempt time and error. No such P1 outbox is implemented in the reviewed flow.

