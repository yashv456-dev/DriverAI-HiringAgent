# HiringAgent P1 — Installation & Operations Runbook

**Audience:** whoever installs, runs, or supports Phase 1.
**Companion document:** [`P1_Technical_Design_Document_Consolidated.md`](../P1_Technical_Design_Document_Consolidated.md) — the technical design. **Change how Phase 1 behaves there, not here.**

---

## Phase 1 in one page

| | |
|---|---|
| **What it is** | An automation in the Microsoft cloud that watches the job inbox and files every application |
| **Setup time** | About 90 minutes, once. Eleven steps |
| **How it runs** | On a timer. It checks the inbox every minute, day and night |
| **Needs a computer?** | Only to set it up. After that it runs on its own, with no laptop switched on |
| **What it does not do** | It never reads or scores a résumé. That is Phase 2, covered by its own guide |

Start at Step 1 and work down. Each step depends on the one before it.

| Reader | Where to go |
|---|---|
| Understanding what it does | **How Phase 1 works**, below — no technical detail needed |
| Installing Phase 1 | Steps 1–11, in order |
| Running it afterwards | Part C |
| Fixing something | Part D — Troubleshooting |
| The exact wording applicants receive | Appendix — The six applicant emails |
| Looking up a technical detail | `P1_Technical_Design_Document_Consolidated.md` |
| Changing how it behaves | `P1_Technical_Design_Document_Consolidated.md` §16, then Part C.4 here |

---

## About Phase 1

Phase 1 watches the job inbox around the clock. It reads every incoming email, filters out spam and anything that is not an application, saves the résumé to SharePoint, creates a candidate row, and replies to the applicant.

| It handles this | In practice |
|---|---|
| Watches the inbox | Checks `apply@driverai.io` every minute, all day, every day |
| Blocks spam and malware | Three screens, applied before anything else happens |
| Saves résumés | Into `Candidate_Resumes / Year / Month`, created automatically |
| Creates the candidate row | One row per applicant in the master workbook |
| Replies to the applicant | One of six emails, matched to their situation |
| Spots repeat applicants | Within 90 days, so nobody gets two rows |
| Tidies the mailbox | Every message is filed, so the inbox stays empty |

**What Phase 1 does *not* do:** open a résumé, score a candidate, decide anything about them, or send any decline. All of that is Phase 2.

> ### Live mail warning — read before any test
>
> `email.send_applicant_emails` is **`true`**. Real applicants receive real emails the moment the flow processes their message. Phase 2 is live too (`config.yaml → test_mode.suppress_emails: false`).
>
> **Both switches must be off for a genuinely silent replay.** Turning either off requires a rebuild *and* a re-import before it takes effect.

---

# How Phase 1 works

*Read this once before installing. It is the whole system in plain English — no technical detail needed.*

## The dividing line

**Phase 1 fills the queue. Phase 2 empties it.** They never talk to each other directly. Phase 1 writes a row marked `New Email Received`, and Phase 2 looks for exactly that.

| | Phase 1 | Phase 2 |
|---|---|---|
| Where it runs | Microsoft cloud, automatic, 24/7 | A laptop or container, when scheduled |
| What it does | Files applications and replies to applicants | Reads and scores résumés, matches roles |
| Does it open the résumé? | **Never** | Yes — that is its whole job |
| Emails it sends | The 6 intake replies | Declines and missing-info requests |

If a candidate was matched to the wrong role, or a job description is out of date, **that is never a Phase 1 problem.**

## Is it always running?

Phase 1 is not started by a person and it is not a program that sits running. **It is a timer.**

Every minute it wakes up, looks in the inbox, handles one email, and goes back to sleep. If the inbox is empty it stops immediately. This happens 1,440 times a day whether or not anyone is at a desk.

| Setting | Value | What it means |
|---|---|---|
| How often it wakes | Every **1 minute** | There is no button to press |
| Mailbox watched | `apply@driverai.io` | The Inbox folder only |
| Which mail it picks up | Read **or** unread | The Inbox itself is the queue |
| Emails per wake-up | **1** | One message at a time |
| Overlapping runs | **Never** | So no candidate ever gets two rows |

> ### The Inbox is the queue, not the unread flag
>
> An email is picked up because it is **sitting in the Inbox**, not because it is unread. To re-run an email, move it back to the Inbox. You do not need to mark it unread first.

**How long a backlog takes** — one message per minute:

| Emails waiting | Time to clear |
|---|---|
| 10 | 10 minutes |
| 40 | 40 minutes |
| 100 | 1 hour 40 minutes |
| 500 | About 8 hours 20 minutes |

## The journey of one email

```
EMAIL ARRIVES in the Inbox
  |
  +-- SCREEN 1: sender a machine or junk source?  -> Junk. Silent.
  +-- SCREEN 2: subject looks like spam?          -> Junk. Silent.
  +-- SCREEN 3: scam / offensive / malware words? -> Junk. Silent.
  |
  +-- Is a PDF or Word resume attached?
  |     +-- NO, but it mentions a job   -> EMAIL 3: Please attach your resume
  |     +-- YES, but wrong file type    -> EMAIL 4: Please resend as PDF or Word
  |     +-- NO, and no mention of a job -> Filed away quietly. Admin is notified.
  |
  +-- Applied in the last 90 days?
        +-- NO  -> Save resume, CREATE ROW -> EMAIL 1: Application received
        +-- YES -> Save resume, UPDATE ROW -> EMAIL 2: Already on file

  Replying and quoting their APP-... reference:
        +-- with a new resume -> EMAIL 5: Updated resume received
        +-- text only         -> EMAIL 6: Message received

  IF ANYTHING FAILS -> administrator alert, message set aside in Archive
  FINALLY: mark read, move to Archive.  (Spam stays unread in Junk Email.)
```

Every email ends in exactly one of those outcomes. Nothing is ever left in the Inbox once handled — which is why "still in the Inbox" means "not yet dealt with".

## What Phase 1 filters out

Three screens run before anything else happens. A message caught by any of them is moved to **Junk Email**, left unread, and **no reply is ever sent** — replying would confirm to a spammer that the mailbox is real.

| Screen | Looks at | Blocks |
|---|---|---|
| 1 — Sender | Who it is from | Automated senders (`noreply`, `notifications`, bounces), SaaS mailers (LinkedIn, Slack, Zoom, DocuSign…), known outsourcing/bench-sales domains, and the mailbox's own address so it never answers itself |
| 2 — Subject | The subject line | Newsletters, digests, receipts, password resets, delivery failures, webinars, and old outbound subjects from a previous version |
| 3 — Content | Subject + body + **attachment names** | Scam and phishing wording, offensive and threatening language, executable file types (`.exe`, `.bat`, `.iso`…), foreign-language scams in 9 languages, and B2B sales-deck language |

A fourth, softer check looks for **link-shortener URLs** (`bit.ly`, `tinyurl`…). These are *not* blocked if a real PDF or Word résumé is attached, because signature generators often wrap a LinkedIn icon in one.

> ### What is not caught
>
> **A staffing agency sending someone else's CV** uses ordinary polite English and passes every screen, so it becomes a normal candidate row. The known agency *domains* are blocked, but a new one, or an individual writing politely, will get through. **Those have to be spotted by hand.**
>
> The reason is structural: the only reliable signal is whose name is on the résumé versus who sent it, and the automation cannot open the résumé to compare.

> ### The one blind spot to schedule time for
>
> Blocked mail leaves **no trace except the Junk Email folder** — no row, no counter, no alert. So if a screen ever blocks a real applicant, nobody finds out unless somebody looks. **Review Junk Email occasionally.** This is not theoretical: a real Chief Marketing Officer had three applications silently junked because the word `marketing@` was on the sender list.

## How it spots repeat applicants

**Matched on email address.** Phase 1 looks up the sender's address in the candidate workbook.

- If they already have a row and applied **within the last 90 days**, they are a repeat. They get **Email 2** ("Already On File"), and **no second row is created**. Their newer résumé is still saved, and their row is re-queued so Phase 2 re-scores it.
- **Beyond 90 days** they count as a fresh application and do get a new row.
- **If they quote their `APP-…` reference number**, the 90-day rule is skipped entirely — naming the exact application counts for more than the date. They get **Email 5** (new résumé) or **Email 6** (text only).

**How many times will it reply?** All the repeat emails share one counter per candidate:

| Their contact | Do they get an email? |
|---|---|
| 1st and 2nd | Yes, a normal reply |
| 3rd | Yes — and it says it is the last automated reply |
| 4th and 5th | **No.** Their résumé and row are still updated, silently |
| 6th onward | **No.** Nothing happens at all |

This is what stops a persistent applicant, or a mail loop, from generating endless replies.

**Duplicate résumé files: replaced, not piled up.** The file is named from the candidate and their reference, so a repeat writes to the same name and you end up with exactly **one current résumé per candidate**. If Phase 2 has already scored and renamed that candidate's file, Phase 1's new copy lands alongside it briefly, and Phase 2 tidies it up on its next run — so it can take one Phase 2 cycle before the folder looks clean.

**One email, several attachments.** If someone attaches a portfolio, cover letter or transcript alongside their résumé, Phase 1 keeps the résumé and skips the others, so Phase 2 does not score a portfolio as though it were a CV. If *every* attachment looks like one of those, it keeps them all rather than treating a real application as having no résumé.

## What Phase 1 fills in, and what it leaves blank

The candidate workbook has **32 columns**. Phase 1 writes **11** of them and leaves the rest blank for Phase 2.

| Phase 1 writes these 11 | |
|---|---|
| Application ID | The `APP-20260630-1430-A3F9` reference |
| Received Date | When they applied (Arizona local time) |
| Last Updated Date | Refreshed every time they contact again |
| Full Name | Taken from the email header |
| Email | The sender's address |
| Mail Subject | Their subject line |
| Mail Body | The first ~255 characters of their message |
| Status | `New Email Received` — Phase 2's cue to pick it up |
| Has Resume | Always `Yes` |
| Original Filename | What their file was called when they sent it |
| Application Updates | The repeat-contact counter, starting at 0 |

**Phase 2 fills the rest:** Category, Phone, Location, Country, Current Skills, Education (and its dates), Looking For Role, Suggested Roles 1–3, the three Portfolio links, Resume Link and URL, Retry Count, and the Info Request stamp.

> **A row only ever exists if there is a real résumé behind it.** Someone who writes in asking about a job without attaching anything gets Email 3, but **no row** — so Phase 2 never sees an empty candidate.

## Where things are stored

| What | Where |
|---|---|
| The candidate workbook | `Documents / Master_Files / Sharepoint_Master_File.xlsx` — one file, all years |
| Résumés | `Documents / Candidate_Resumes / <Year> / <Month> /` — folders created automatically |
| Résumé file name | `JaneDoe_APP-20260630-1430-A3F9.pdf` at intake; Phase 2 renames it to `Jane_Doe_A3F9.pdf` once scored |
| Rejected candidates | The `Rejected` sheet in the same workbook; their résumés move to `Candidate_Resumes / <Year> / Rejected /` |
| The client-facing report | `Documents / Master_Files / Candidate_List_Results.xlsx` — written by Phase 2 |
| **Job descriptions** | **Not used by Phase 1 at all.** They live on the Phase 2 side: SharePoint `led1234567-my.sharepoint.com` → `/personal/tracys_driverai_io` → folder **`Staffing/PDs`** |

> **Do not move résumé files between folders.** Phase 2 works out where a file is from the candidate's row. A file moved by hand becomes invisible to it.

---

# Part A — Before you start

## Step 1 — Check what you need

*5 min · whoever is setting up*

| What you need | Essential? |
|---|---|
| Microsoft 365 account with Power Automate | Yes |
| Global administrator access, for about an hour | Yes |
| The shared mailbox `apply@driverai.io` | Yes |
| **Send As** permission on that mailbox | Yes |
| Permission to create a SharePoint site | Yes |
| Excel Online (Business) | Yes |
| The delivered files | Yes |
| A computer with Python | No — only for Step 4 |

> **Phase 1 itself needs no computer.** Once imported it runs entirely in Microsoft's cloud, 24 hours a day. A computer is only needed to set it up, to change settings later, or to push old mail back through it.

### Licensing

Phase 1 uses only standard connectors. No premium connector, no Dataverse, no Power Apps licence, and no separate per-user Power Automate plan.

| Item | What is needed |
|---|---|
| Power Automate | The rights included with a Microsoft 365 business or enterprise plan |
| The shared mailbox | A shared mailbox. Under 50 GB it needs no licence of its own |
| SharePoint | Any plan that allows creating a Team site |
| Environment | The default Power Automate environment is fine |

> ### Check your daily allowance before going live
>
> Microsoft counts every action a flow performs against a daily allowance. Checking the mailbox **once a minute uses roughly 5,700 actions a day before a single application is processed**, and the allowance included with a Microsoft 365 plan is commonly **6,000**.
>
> If yours is tight, set `trigger.interval_min` to `2`. It halves the usage and everything still works — the queue simply drains at half speed. See Part C.4.

### Storage, backups and dates

| Item | What to expect |
|---|---|
| Résumé storage | About 200 KB–2 MB each. 1,000 applicants is under 2 GB |
| Run history | Power Automate keeps roughly 28 days. After that the alert emails are the only record |
| Backups | **Nobody backs up the candidate workbook automatically.** Decide who does, and how often |
| Dates and times | Recorded in **Arizona local time** (UTC−7, no DST) since 2026-08-27. Rows written before that date hold UTC values |
| Mailbox growth | Nothing is ever deleted, so review the mailbox once or twice a year |

### Who does what

| Steps | Who | How long |
|---|---|---|
| 1–4 Files and tools | Whoever is setting up | 15 minutes |
| 5–6 Site, folders, workbook | SharePoint administrator | 20 minutes |
| 7 Mailbox permission | Microsoft 365 administrator | 5 minutes |
| 8–9 App registration | Global administrator | 20 minutes |
| 10–11 Import and test | Whoever owns the flow | 20 minutes |

---

## Step 2 — Understand the three connections

*5 min · whoever is setting up*

During import you pick a connection for each of these. They can all be the same account.

| Connection | Must be able to |
|---|---|
| **Office 365 Outlook** | Read the shared mailbox, and **send as** `apply@driverai.io` |
| **SharePoint** | Create folders and files under `Documents / Candidate_Resumes` |
| **Excel Online (Business)** | Read and update `Sharepoint_Master_File.xlsx` |

> ### Use an account that will not leave the company
>
> The flow runs under whoever's connections were chosen at import. If that person's account is disabled, **Phase 1 stops dead**, and the failure is confusing because nothing was changed. A service or shared account is far safer than a personal one.

Phase 1 holds **no secret** inside the package. It authenticates entirely through these three connections.

---

## Step 3 — Copy the delivered files

*5 min · whoever is setting up*

1. In File Explorer, create a folder called `HiringAgent` on the `C:` drive.
2. Copy the delivered `DriverAI_HiringAgent` folder inside it.
3. If the delivered files came as ZIPs, right-click each one and choose **Extract All**.
4. Right-click the folder, choose **Properties**, and if there is an **Unblock** check box, select it.

```
C:\HiringAgent\DriverAI_HiringAgent\HiringAgent_P1\      <- Phase 1 (this runbook)
C:\HiringAgent\DriverAI_HiringAgent\HiringAgent_App_P2\  <- Phase 2 (separate guide)
```

> **Watch for a doubled folder.** The delivered ZIP often extracts into a folder called `HiringAgent_P1` containing *another* `HiringAgent_P1`. If you see that, move the inner folder up one level.

The two delivered files that matter:

| Delivered file | What you do with it |
|---|---|
| `flow\DriverAI-Hiring-AutoReply-apply.zip` | Import into Power Automate (Step 10) |
| `P1_Templates\HiringAgent_P1_CandidateList.xlsx` | Upload to SharePoint (Step 6) |

Everything else in the folder is for changing settings later or replaying old mail.

---

## Step 4 — Install Python (only if you will change settings)

*10 min · optional*

Phase 1 does not need Python to run. You only need it to change a setting, or to push a batch of old mail back through the system. If you will never do either, skip this step.

1. Go to <https://www.python.org/downloads/windows/>
2. Download **Python 3.12**, the Windows installer (64-bit). 3.11 also works; avoid 3.13 or newer.
3. Run the installer.
4. On the first screen, select the check box **"Add python.exe to PATH"** at the bottom.
5. Choose **Install Now**, then **Close**.

> **The PATH check box is easy to miss.** If you do not select "Add python.exe to PATH" on the first screen, nothing will work afterwards and the error will not explain why. Re-run the installer if you missed it.

---

# Part B — Build it

## Step 5 — Create the SharePoint site

*10 min · SharePoint administrator*

Skip this step if the site already exists.

1. Go to <https://www.office.com> and open **SharePoint**.
2. Choose **+ Create site** at the top left.
3. Choose **Team site**, which gives you a Documents library automatically.
4. Name it `CandidateList_HiringAgent`, exactly.
5. Check the address ends with `/sites/CandidateList_HiringAgent`
6. Set privacy to **Private**, then choose **Next**.
7. Add the people who need access as members, then choose **Finish**.

> **If you use a different site name,** that is fine — but the name must then be updated in `flow/flow_config.json` (`sharepoint.site`, `excel.source`, `excel.drive`), the whole package rebuilt and re-imported, plus Phase 2's settings. Using the exact name above avoids all of that.

---

## Step 6 — Create the folders and upload the workbook

*10 min · SharePoint administrator*

> ### Everything here is matched by exact name
>
> A rename, an extra space, or different capitalisation breaks the automation, and the failure usually appears somewhere else entirely. Set these up once and leave them.

Open the site, choose **Documents** in the left menu, and create three folders:

| Folder | What it holds |
|---|---|
| `Candidate_Resumes` | Every résumé, in Year / Month subfolders |
| `Master_Files` | The live candidate workbook and the Phase 2 report |
| `SharePoint_Master_Template` | A spare blank copy of the template |

Then:

1. Take `HiringAgent_P1\P1_Templates\HiringAgent_P1_CandidateList.xlsx` from the delivered files.
2. Upload it into `Master_Files`.
3. Rename the uploaded copy to exactly `Sharepoint_Master_File.xlsx`
4. Upload a second copy into `SharePoint_Master_Template` and never touch it again.

The template already contains **both** sheets (`CandidateList` and `Rejected`) and the Category dropdown, so uploading it gives you everything on day one. It also ships one intentionally blank seed row, because Excel Online and Graph reject a header-only table — the automation ignores that row.

> ### Do not open the master workbook, and do not reorganise the folders
>
> Once Phase 1 is live, leave `Sharepoint_Master_File.xlsx` **closed**. Opening it in the Excel *desktop* app locks the file, and while it is locked the automation cannot write. You get administrator alerts and candidate rows stop updating even though mail keeps arriving.
>
> Do not rename, move or delete the three folders either, and do not move résumé files between the Year and Month subfolders — Phase 2 rebuilds each file's path from its row and will no longer find it. If you need to look at the data, open it in the **browser** and close it again, or use the Phase 2 report.

### The names that must be exact

| What | Exact value |
|---|---|
| Résumés folder | `Candidate_Resumes` |
| Workbook folder | `Master_Files` |
| Template folder | `SharePoint_Master_Template` |
| The workbook file | `Sharepoint_Master_File.xlsx` |
| Main worksheet | `CandidateList` (do not rename) |
| Main table | `HiringAgent_P1_Candidates` (do not rename) |
| Second worksheet | `Rejected` |
| Second table | `RejectedCandidates` |
| Processed mail goes to | `Archive` |
| Junk goes to | `Junk Email` |

### What the finished structure looks like

```
Documents
 ├── Master_Files/
 │    ├── Sharepoint_Master_File.xlsx      <- THE candidate workbook
 │    │      ├── sheet "CandidateList"
 │    │      └── sheet "Rejected"
 │    └── Candidate_List_Results.xlsx      <- Phase 2 writes this for you
 ├── Candidate_Resumes/
 │    ├── 2026/June/Jane_Doe_A1B2.pdf
 │    └── 2026/July/John_Smith_D4E5.docx
 └── SharePoint_Master_Template/           <- spare copy, never read
```

| Created by | What |
|---|---|
| **You, once** | The site, the three folders, the workbook, and Send As |
| **Phase 1** | The Year / Month subfolders, the résumé files, and the candidate rows |
| **Phase 2** | The scored columns, renamed résumés, the Rejected sheet, and the report |
| **Nobody** | The workbook itself. If deleted, **a person** restores it from the template — Phase 1 never re-creates it |

> **Migrating from a per-year workbook?** Move the existing workbook from `Candidate_Resumes/<Year>/HiringAgent_P1_CandidateList.xlsx` to `Master_Files/Sharepoint_Master_File.xlsx` (this keeps all rows and the table). Also update Phase 2's `SHAREPOINT_WORKBOOK` in `.env` to the new path, or remove it to use the new default.

---

## Step 7 — Grant Send As on the mailbox

*5 min · Microsoft 365 administrator*

1. In the Microsoft 365 admin centre, go to **Teams & groups → Shared mailboxes**.
2. Select `apply@driverai.io`.
3. Choose **Send as**, then **Add permissions**.
4. Add the account that will own the Outlook connection.

> ### Full Access is not enough
>
> All six applicant replies are sent *from* the shared mailbox. With Full Access alone the flow imports fine, processes mail fine, and then **every single reply fails**. This is the most common Phase 1 setup mistake.

---

## Step 8 — Create the app registration

*10 min · Global administrator*

> **The live flow does not use this.** Phase 1 authenticates through the three Power Automate connections and holds no secret. This registration is for **Phase 2** and for the tools that replay old mail (`trigger_reset.py`, `bulk_move_tool.py`, `audit_p1_live_state.py`).

1. Go to <https://entra.microsoft.com>
2. Choose **Applications → App registrations → + New registration**.
3. Name it something recognisable, for example "DriverAI Hiring Agent".
4. Choose **Accounts in this organizational directory only** (single tenant).
5. Leave Redirect URI empty, then choose **Register**.
6. Copy the **Directory (tenant) ID** and the **Application (client) ID** somewhere safe.
7. Choose **Certificates & secrets → + New client secret**.
8. Give it a description and an expiry, then choose **Add**.

> ### Copy the Value, not the Secret ID
>
> The two columns sit next to each other and look almost identical. The **Value** is shown once. Refresh the page and it is masked forever.
>
> **Put the expiry date in a calendar now** — when it expires, Phase 2 simply stops connecting.

---

## Step 9 — Add and consent the permissions

*10 min · Global administrator*

1. In the registration, choose **API permissions**.
2. Choose **+ Add a permission → Microsoft Graph**.
3. Choose **Application permissions**, not Delegated permissions.
4. Select the four permissions below.
5. Choose **Add permissions**.
6. Choose **Grant admin consent for your organisation**, and confirm.
7. Check every row shows a green check mark under **Status**.

> ### Application permissions, not Delegated
>
> Delegated permissions act on behalf of a signed-in person. There is no signed-in person here — the system runs unattended. Choose the wrong one and nothing authenticates, and the error message will not tell you why.

| Permission | What it is for |
|---|---|
| `Sites.ReadWrite.All` | Read and update the candidate workbook |
| `Files.ReadWrite.All` | Read, rename and move résumé files |
| `Mail.Send` | Send Phase 2's emails |
| `Mail.ReadWrite` | Move mail back to the Inbox to re-run Phase 1 |

The values go into Phase 2's settings file, `HiringAgent_App_P2\.env`:

```
TENANT_ID=<from Step 8>
CLIENT_ID=<from Step 8>
CLIENT_SECRET=<the Value from Step 8>
SENDER_MAILBOX=apply@driverai.io
```

> **Treat that file like a password.** Never email it, put it in a chat, or place it in a shared folder. The secret grants read and write access to your SharePoint and the ability to send mail as the mailbox.

> **If mailbox Graph calls return `403 ErrorAccessDenied` even after admin consent,** the cause is an Exchange **Application Access Policy**, not a missing Graph permission. A tenant admin needs Exchange Online PowerShell: `Get-ApplicationAccessPolicy` to check the current scope, then `New-ApplicationAccessPolicy` / `Set-ApplicationAccessPolicy` to allow this app's `CLIENT_ID` to access `apply@driverai.io`.

---

## Step 10 — Import the flow and switch it on

*10 min · whoever owns the flow*

1. Go to <https://make.powerautomate.com>
2. Choose **My flows → Import → Import Package (Legacy)**.
3. Upload `flow\DriverAI-Hiring-AutoReply-apply.zip`. **Do not unzip it first.**
4. On the review screen:
   - **First install** → choose **Create as new**.
   - **Re-deploying a change** → choose **Update** and pick the existing flow.
5. For each of the three connections — **Office 365 Outlook**, **SharePoint**, **Excel Online (Business)** — click **Select during import** and pick your connection.
6. Choose **Import**.
7. Open the flow and choose **Edit**.
8. Look for anything marked in red. If the four Excel row actions (`Get_rows`, `Get_rows_ref`, `Add_row`, the `Patch_*` actions) are red, pick the workbook and table again from the dropdowns.
9. Choose **Save**, then switch the flow **On**.

> ### This is the only supported way to install Phase 1
>
> Import Package (Legacy) → Create as new → three connections → save → switch On. Do not use the plain **Import** option, do not unzip and re-zip the package, and do not rebuild the flow by hand.

> ### After a first install, record the flow's GUID
>
> An **Update** import only lands on an existing flow if the package's id matches that flow's real id. When they differ, Power Automate silently offers "Create as new" instead, and **the existing flow keeps running the old definition** — the import appears to succeed and changes nothing.
>
> This really happened on 2026-08-27: the package carried a stale GUID and **three consecutive imports did nothing at all**. The current value is `22f1cacc-4e05-41ca-806e-f4baf41de15b`, stored as `flow_id` in `flow/flow_config.json`.
>
> **If you create the flow fresh in a new tenant** (or ever delete and recreate it), export the flow once, read the id out of the export — it is both the manifest resource id and the `Microsoft.Flow/flows/<id>/` folder name — and put that value into `flow_id` before your next rebuild. Otherwise every future update silently no-ops.

> **Never "Create as new" on an update** once the flow exists, or you get two flows firing on the same inbox and applicants receive duplicate replies. If an Update genuinely will not work, import once as new, re-pick any red actions, save, and then **switch the old flow Off**.

### If the import fails

The error usually mentions a missing `table` property:

```
Flow save failed with code 'DynamicParameterInputInvalid'
The request to API 'sharepointonline' operation 'GetTable' is missing required property 'table'
```

This means the package imported with a workbook binding Power Automate could not resolve. Check in this order:

1. The workbook is at exactly `Documents/Master_Files/Sharepoint_Master_File.xlsx`.
2. It was **not renamed** during upload and still contains the table `HiringAgent_P1_Candidates`.
3. The **SharePoint** and **Excel Online (Business)** connections use the same account, and it can open that workbook.
4. Nobody has the workbook open in the Excel **desktop** app.
5. If you are redeploying from the old yearly layout, move the workbook into `Master_Files/` first (Step 6).
6. If the site or library changed, update `flow/flow_config.json` (`sharepoint.site`, `excel.source`, `excel.drive`, `excel.file`, `excel.table`), run `python flow/build_zip.py`, then re-import the rebuilt zip.

---

## Step 11 — Test it: the five signs of success

*10 min · whoever owns the flow*

1. From an **outside** email address, send a message to `apply@driverai.io` with a PDF or Word résumé attached.
2. Use a normal subject such as "Application for Data Analyst".
3. Wait about two minutes.
4. Check all five signs:

| # | What to check | Expected |
|---|---|---|
| 1 | The candidate row exists | A new row with `Status = New Email Received` |
| 2 | The intake fields are filled | Application ID, Received Date, Full Name, Email, `Has Resume = Yes` |
| 3 | The résumé is stored | A file in `Candidate_Resumes / Year / Month` |
| 4 | The applicant was contacted | The acknowledgment arrived, and `Mail Sent` holds a timestamp |
| 5 | The mailbox was tidied | The message is now read and in `Archive` |

> **If nothing happens,** open the flow in Power Automate and look at **Run history**. A failed run shows exactly which step broke — far faster than guessing.

Remember that applicant mail is **live**: your test address really receives the acknowledgment, and the test row is a real row. Delete it afterwards if it would pollute reporting.

---

# Part C — Running it day to day

## C.1 What it sends

Six applicant emails, matched to the situation. **The full text of each is in the appendix at the back of this document.**

| # | Subject line | When it fires |
|---|---|---|
| 1 | `Application Received - DriverAI (Ref: …)` | New applicant with a valid résumé |
| 2 | `Your DriverAI Application Is Already On File (Ref: …)` | Same person re-applies within 90 days |
| 3 | `Please Attach Your Resume - DriverAI` | Hiring keywords but no attachment at all |
| 4 | `Please Resend Your Resume as PDF or Word - DriverAI` | An attachment, but not a PDF or Word file |
| 5 | `Updated Resume Received - DriverAI (Ref: …)` | They quoted their reference and attached a new résumé |
| 6 | `Message Received - DriverAI (Ref: …)` | They quoted their reference with no attachment |

**Repeat contacts are capped.** Emails 2, 5 and 6 share one counter per candidate: contacts 1–3 get a reply (the 3rd says it is the last one), contacts 4–5 still update the row and résumé but send **nothing**, and from the 6th everything stops. Emails 1, 3 and 4 are not capped.

**Admin alerts** go to `yashv@driverai.io` — 14 alert actions across five kinds (core failure, poll failure, seven row-patch failures, four intake-write failures, and one for mail that was not an application). You do not need to know them individually; what matters is that **an alert email means something needs a human**, and the alert names the exact step.

## C.2 Checking it is healthy

Open the flow in Power Automate and choose **Run history**.

- A **failed** run shows exactly which step broke, with the error message.
- A run that shows **Succeeded but did nothing** means the inbox was empty. That is normal, and happens most of the time.
- **Unread mail sitting in Archive** means something failed and was set aside. Worth checking occasionally.

> ### Two blind spots
>
> 1. **Mail caught by the spam screens leaves no trace except the Junk Email folder** — no row, no counter, no alert. A wrongly-blocked real applicant is invisible unless somebody opens Junk. Review it occasionally.
> 2. **A row update that failed will have sent an alert, but the run itself still reports Succeeded.** Green runs are not proof that every row is correct — the alert mailbox is.

**A weekly five-minute check:**

1. Run history — any red runs?
2. Archive — any **unread** messages?
3. The admin alert mailbox — anything new?
4. Junk Email — anything that is obviously a real applicant?
5. The workbook — are new rows appearing with today's date?

## C.3 Putting an email back through

Phase 1 picks up whatever is sitting in the **Inbox**. To re-run an email, put it back there.

1. Open the `apply@driverai.io` mailbox.
2. Find the message in `Archive` or `Junk Email`.
3. Move it back to the **Inbox**.
4. Within a minute or two Phase 1 picks it up again.

That is the whole recovery procedure for a single message, and it is what the alert emails tell you to do.

> ### Before you replay a large batch
>
> **Replayed applicants get emailed again.** A batch drains at one message per minute, so 500 messages take about eight hours. Consider setting `email.send_applicant_emails` to `false` first (rebuild and re-import), and remember Phase 2 has its own separate switch.

For bulk replays, use the tools in Part C.5 rather than moving mail by hand.

## C.4 Changing a setting

Phase 1's single source of truth is `flow/flow_config.json`. **Never hand-edit `definition.json` or the zip** — they are regenerated on every build and your change is silently overwritten.

```bat
cd /d C:\HiringAgent\DriverAI_HiringAgent\HiringAgent_P1
python flow\build_zip.py
```

Then **re-import as Update**, re-picking the three connections (Step 10).

The build script prints a summary of what was baked in — read it to confirm your change actually landed:

```
bad_senders         : 43
bad_subjects        : 27 (config; self-loop guard is sender-based)
spam/safety gate    : 123 (30 scam + 36 offensive + 20 malware/exe + 25 non-EN + 12 vendor-pitch)
app_keywords        : 21
year separator      : OFF
month separator     : OFF
admin alerts        : ON
```

The settings you are most likely to touch:

| Key | Current | Effect |
|---|---|---|
| `email.send_applicant_emails` | `true` | All 6 applicant emails. `false` = silent intake, everything else still runs |
| `email.send_admin_failure_alerts` | `true` | All 14 admin alerts |
| `trigger.interval_min` | `1` | Poll interval in minutes. `2` halves the daily action usage |
| `business_rules.reply_cap` | `3` | How many contacts actually get an email |
| `business_rules.duplicate_check_days` | `90` | The re-application window |
| `spam_filters.*` | — | The spam lists. **Read `P1_Technical_Design_Document_Consolidated.md` §6 before editing any of them** |

> ### Two traps when changing the spam lists
>
> 1. **Every term is an unanchored substring match.** A bare word like `alert` blocks "Alert Systems Engineer"; `marketing@` blocked a real Chief Marketing Officer. Prefer multi-word phrases or distinctive domains, never a bare common noun.
> 2. **After any list change, replay the lists over recent mail** — a false positive is otherwise invisible (Part C.2). `P1_Technical_Design_Document_Consolidated.md` §6 records which terms have already been proven unsafe; do not re-add them.

> **A config change does nothing until the rebuilt zip is re-imported.** Editing the JSON alone changes no behaviour in the cloud.

## C.5 Recovery and audit tools

Three Python helpers, run from `HiringAgent_P1/`. They are separate from the live flow and use Phase 2's `.env` credentials, which need `Mail.ReadWrite` admin-consented (Step 9).

**`trigger_reset.py`** — moves Archive mail back to the Inbox as unread, so the poll processes it again.

```
python trigger_reset.py --dry-run          # preview, no changes
python trigger_reset.py                    # re-trigger last 31 days (default)
python trigger_reset.py --days 7           # custom window
python trigger_reset.py --all              # ALL Archive mail, no date limit
python trigger_reset.py --all --dry-run
```

| Flag | Default | Effect |
|---|---|---|
| `--days N` | `31` | Look back N days |
| `--all` | off | Reset ALL Archive mail regardless of age (overrides `--days`) |
| `--mailbox addr` | `apply@driverai.io` | Mailbox to reset |
| `--dry-run` | off | List emails without making changes |

**Always run `--dry-run` first.** After a reset each email is processed exactly as if it had just arrived — spam gates, 90-day duplicate check, reply emails, SharePoint row.

**`bulk_move_tool.py`** — bulk mailbox move utility for larger Archive batches. No re-import needed.

**`audit_p1_live_state.py`** — read-only audit of live SharePoint and mailbox state. Writes a dated `.md`/`.csv` snapshot pair to `archive/audits/`. Run it to sanity-check without changing anything.

> **Known gap:** `trigger_reset.py` only moves *Archive* → Inbox. It cannot recover mail that is stuck in the Inbox already.

## C.6 Testing after a change

```bat
cd /d C:\HiringAgent\DriverAI_HiringAgent\HiringAgent_P1
python test_p1.py
```

The suite validates the upload ZIP, the loose generated definition, the config, the docs and the workbook template together. **Run it after every build, before importing.** A failure here is far cheaper than a failure in production.

---

# Part D — Troubleshooting

| Symptom | Most likely cause | What to do |
|---|---|---|
| Nothing happens at all | The flow is switched off | Open the flow, confirm it shows **On**, check Run history |
| Every reply fails | **Send As** was never granted | Grant Send As on the mailbox (Step 7). Full Access is not enough |
| Import fails on a `table` error | The workbook or table was renamed, or is open in Excel desktop | Confirm the exact names in Step 6; close the workbook |
| An import "succeeded" but nothing changed | The package GUID does not match the live flow | See the GUID box in Step 10 — check `flow_id` |
| Rows appear, no résumé saved | SharePoint connection cannot write | Re-pick it with an account that can |
| Applicant got two identical emails | Two flows are running on the same mailbox | Switch the older flow **Off** |
| A real applicant never got a reply | Caught by a spam screen | Open **Junk Email** and look. See `P1_Technical_Design_Document_Consolidated.md` §6 |
| Blank `Mail Sent` on a full row | The reply failed to send | Check Run history; the row and file are fine |
| Row stuck on `New Email Received` | Phase 2 has not run yet | Not a Phase 1 fault. Check the Phase 2 schedule |
| Alerts about rows not updating | The workbook is open and locked | Close it in the Excel **desktop** app |
| A candidate row was never created | `Add_row` failed — you will have an `APPLICANT LOST` alert | Move the message from Archive back to the Inbox; the next poll re-runs intake |
| Résumé on disk is the old version | A `CreateFile` failed — you will have an alert | Same recovery: move the message back to the Inbox |
| A vendor/agency pitch got a candidate row | It used ordinary polite English and passed the gates | Delete the row by hand. See `P1_Technical_Design_Document_Consolidated.md` §6 for why this is not fully solvable in P1 |
| Résumé filed under the wrong month | A row written **before** 2026-08-27 (UTC dates) | Expected. Historical rows keep UTC; see `P1_Technical_Design_Document_Consolidated.md` §12 |
| Flow stopped after a staff change | The connection owner's account was disabled | Re-point the three connections to a service account (Step 2) |
| Phase 2 stopped connecting | The client secret expired | Create a new secret (Step 8) and update `.env` |

**The general recovery move.** For almost any single message that went wrong: find it in `Archive`, move it back to the **Inbox**, and the next poll re-runs the whole intake for it. This is safe and repeatable — the duplicate check and the caps prevent runaway replies.

**When you need the exact error:** Power Automate keeps run history roughly 28 days. After that the admin alert emails are the only record, which is why they are worth keeping enabled and worth reading.

---

# Part E — Backups and rollback

**Nothing backs up the candidate workbook automatically.** Decide who does this and how often — it is the one piece of data that cannot be reconstructed.

| What | How to protect it |
|---|---|
| The candidate workbook | Periodic copy of `Master_Files/Sharepoint_Master_File.xlsx`. A row-level export alone is **not** enough — it cannot rebuild the calculated `Resume Link` column |
| The résumé files | They live in SharePoint and are covered by whatever retention the tenant has |
| The flow itself | Export the flow from Power Automate before any risky change, and keep the delivered zip |
| The settings | `flow/flow_config.json` — keep a copy before editing |

`_backups\` in the delivered tree holds dated restore points, each containing P1 + P2 code zips. Two of them also contain a binary copy of the live workbook plus a JSON dump of both sheets and the résumé inventory — **that pair is the true restore set**.

**To roll back the flow:** re-import an earlier flow zip as an **Update**. Because the package carries a fixed GUID, this replaces the definition in place and keeps the connections.

**If the workbook is ever deleted:** Phase 1 will **not** re-create it — that safety was removed deliberately in 2026-07-04 after an auto-heal wiped real candidate rows. Instead the run fails and an admin alert fires. Restore the workbook by hand from `SharePoint_Master_Template/` or from `P1_Templates/HiringAgent_P1_CandidateList.xlsx`, keeping the exact name `Sharepoint_Master_File.xlsx`.

---

# Check-off sheet

**Files and tools**

- [ ] Delivered files copied to `C:\HiringAgent`
- [ ] Files unblocked in Properties
- [ ] No doubled `HiringAgent_P1` folder
- [ ] Python installed with PATH selected (only if changing settings)

**SharePoint**

- [ ] Site created at `/sites/CandidateList_HiringAgent`
- [ ] `Candidate_Resumes`, `Master_Files` and `SharePoint_Master_Template` created
- [ ] Workbook uploaded and renamed to `Sharepoint_Master_File.xlsx`
- [ ] Spare copy placed in `SharePoint_Master_Template`
- [ ] Both sheets present: `CandidateList` and `Rejected`

**Permissions**

- [ ] **Send As** granted on `apply@driverai.io`
- [ ] App registration created
- [ ] Tenant ID, client ID and secret **Value** copied and stored safely
- [ ] Secret expiry date put in a calendar
- [ ] Four Application permissions added and consented

**Import and test**

- [ ] Package imported with **Import Package (Legacy)**, Create as new
- [ ] All three connections selected
- [ ] No steps left showing red
- [ ] Flow saved and switched **On**
- [ ] `flow_id` in `flow_config.json` matches the live flow's real GUID
- [ ] Test application sent and all five signs confirmed

**Ongoing**

- [ ] `Sharepoint_Master_File.xlsx` is kept closed
- [ ] The SharePoint folders have not been renamed or reorganised
- [ ] Somebody reads the administrator alert mailbox
- [ ] `Junk Email` is reviewed occasionally for wrongly blocked applicants
- [ ] `Archive` is checked occasionally for unread (failed) messages
- [ ] Somebody backs up the candidate workbook
- [ ] The client secret's expiry date is in a calendar

---

# Appendix — The six applicant emails

Every email below is sent **from `apply@driverai.io`** and carries the same footer:

```
────────────────────────────────────────
This is an auto-generated email and this mailbox is not monitored.
```

`APP-20260630-1430-A3F9` is an example reference — each applicant gets their own.

The three repeat emails (2, 5 and 6) have a **second version** used on the third and final reply, which tells the applicant it is the last automated message they will get. Those variants are in `P1_Technical_Design_Document_Consolidated.md` §13.

---

**1 · Application Received** — a new applicant with a valid résumé. *Creates their row.*

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
```

---

**2 · Already On File** — they applied again within 90 days. *No second row; their résumé is still updated.*

Subject: `Your DriverAI Application Is Already On File (Ref: APP-20260630-1430-A3F9)`

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
```

---

**3 · Please Attach Your Resume** — they mentioned a job but attached nothing. *No row is created.*

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
```

---

**4 · Wrong File Format** — they attached something that is not a PDF or Word file. *No row is created.*

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
```

---

**5 · Updated Resume Received** — they quoted their reference and attached a new résumé. *Their row is re-queued so Phase 2 re-scores the new version.*

Subject: `Updated Resume Received - DriverAI (Ref: APP-20260630-1430-A3F9)`

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
```

---

**6 · Message Received** — they quoted their reference but sent no résumé. *Counter only; nothing else changes.*

Subject: `Message Received - DriverAI (Ref: APP-20260630-1430-A3F9)`

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
```

---

> **To change any of this wording** you need a developer: the text lives in Python code (`flow/build_zip.py`), not in a settings file, and the package has to be rebuilt and re-imported. See `P1_Technical_Design_Document_Consolidated.md` §13.

> **To stop applicant email entirely** — for a test, or to replay a backlog quietly — set `email.send_applicant_emails` to `false`, rebuild, and re-import. Phase 2 has its own separate switch, and **both** must be off for a genuinely silent run.

---

# Appendix — Command cheat sheet

All commands run from `C:\HiringAgent\DriverAI_HiringAgent\HiringAgent_P1`.

```bat
cd /d C:\HiringAgent\DriverAI_HiringAgent\HiringAgent_P1

REM Rebuild the package after editing flow\flow_config.json
python flow\build_zip.py

REM Validate everything before importing
python test_p1.py

REM Preview which archived mail would be replayed
python trigger_reset.py --dry-run

REM Replay the last 31 days of archived mail
python trigger_reset.py

REM Replay everything in Archive, no date limit
python trigger_reset.py --all

REM Read-only audit of live SharePoint + mailbox state
python audit_p1_live_state.py
```

Phase 2 side, for the sheet layout:

```bat
cd /d C:\HiringAgent\DriverAI_HiringAgent\HiringAgent_App_P2
python resort_candidate_sheets.py
```

---

# Appendix — Things that will bite you

1. **Full Access is not Send As.** The single most common setup failure (Step 7).
2. **Never "Create as new" when updating** an existing flow — you end up with two flows on one mailbox.
3. **An Update import with a mismatched GUID silently does nothing.** Three imports were lost to this on 2026-08-27 (Step 10).
4. **Editing `flow_config.json` changes nothing until you rebuild *and* re-import.**
5. **Never hand-edit `definition.json`, `manifest.json` or `apisMap.json`** — every build overwrites them.
6. **Opening the workbook in Excel desktop locks it** and stops the automation writing.
7. **Spam terms are unanchored substrings.** A bare common noun will block real applicants, silently.
8. **Replaying mail re-emails applicants** while `send_applicant_emails` is `true`.
9. **P1 and P2 have separate mail switches.** Both must be off for a silent run.
10. **The workbook is never auto-restored.** If it is deleted, a human must put it back.
11. **Do not move résumé files between folders** — Phase 2 rebuilds each path from the row and will not find a moved file.
12. **A green run is not proof the row is correct** — read the alert mailbox.
13. **Phase 1 is not in git.** Keep your own copies before editing.

---

**Document map**

| Doc | Holds |
|---|---|
| **`P1_RUNBOOK.md`** (this file) | Installation, day-to-day operation, monitoring, troubleshooting, backups |
| **`P1_Technical_Design_Document_Consolidated.md`** | Full technical design — flow trace, spam gates, email text, row schema, config reference, change history |
| `flow/flow_config.json` | Every value the flow actually consumes |

*Last consolidated 2026-08-28. Values verified against the built `flow/definition.json` on that date.*
