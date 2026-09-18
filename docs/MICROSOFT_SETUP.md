# Microsoft 365, Azure Entra ID & SharePoint Setup Guide

This guide walks you through configuring Microsoft 365, Azure Entra ID (Azure AD) App Registrations, SharePoint Document Libraries, and Phase 1 Power Automate flows for the DriverAI Hiring Agent.

---

## 🏢 System Architecture: Two-Phase Microsoft Integration

The DriverAI Hiring Agent runs a dual-phase Microsoft integration:

```text
┌─────────────────────────────────────────────────────────────────────────────┐
│                       Phase 1: Cloud Ingestion (P1)                         │
│                                                                             │
│  Candidate Email ➔ apply@driverai.io ➔ Power Automate Cloud Flow           │
│                                         │                                   │
│                     ┌───────────────────┴───────────────────┐               │
│                     ▼                                       ▼               │
│      Save Attachment to SharePoint           Create row in Master Excel     │
│      /Candidate_Resumes/YYYY/MM/             /Master_Files/Sharepoint_      │
│                                              Master_File.xlsx               │
└─────────────────────────────────────┬───────────────────────────────────────┘
                                      │
                                      ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│                      Phase 2: Evaluation & Scoring (P2)                     │
│                                                                             │
│  P2 Engine (Web UI / Docker / CLI) ➔ Microsoft Graph API (App-Only Auth)    │
│                                         │                                   │
│                     ┌───────────────────┴───────────────────┐               │
│                     ▼                                       ▼               │
│         Fetch unread candidates                 Upload scored results:      │
│         Download raw resumes                    - P2-MasterFile.xlsx        │
│                                                 - Candidate_List_Results    │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## 🔑 Phase 2: Microsoft Entra ID (Azure AD) App Registration

Phase 2 connects headlessly to Microsoft Graph using **Client Credentials (App-Only Authentication)**. There is no user prompt or browser popup during background processing.

### Step 1: Create an App Registration
1. Sign in to the **[Microsoft Entra Admin Center](https://entra.microsoft.com/)** or the **[Azure Portal](https://portal.azure.com/)**.
2. In the left navigation, go to **Identity** ➔ **Applications** ➔ **App registrations**.
3. Click **+ New registration**.
4. Configure the application:
   * **Name**: `DriverAI-HiringAgent-P2`
   * **Supported account types**: Select *Accounts in this organizational directory only (Single tenant)*.
   * **Redirect URI**: Leave blank (not required for client credentials flow).
5. Click **Register**.
6. On the **Overview** page, copy the following values:
   * **Application (client) ID** ➔ Save as `CLIENT_ID`
   * **Directory (tenant) ID** ➔ Save as `TENANT_ID`

---

### Step 2: Generate a Client Secret
1. In the app registration sidebar, click **Certificates & secrets**.
2. Select the **Client secrets** tab and click **+ New client secret**.
3. Provide a description (e.g. `DriverAI Production Secret`) and choose an expiration period (e.g. 180 days).
4. Click **Add**.
5. **⚠️ CRITICAL**: Immediately copy the string in the **Value** column (NOT the *Secret ID*). Azure only shows this value once upon creation!
   * Save as `CLIENT_SECRET`.

---

### Step 3: Configure Microsoft Graph Permissions
1. In the app sidebar, click **API permissions**.
2. Click **+ Add a permission** ➔ select **Microsoft Graph**.
3. Click **Application permissions** *(Do NOT select Delegated permissions; headless backend daemons do not have a signed-in user)*.
4. Add the following permissions:

| Permission | Type | Reason |
|---|---|---|
| **`Sites.ReadWrite.All`** | Application | Required to read and update candidate rows in SharePoint Excel tables and create folders |
| **`Files.ReadWrite.All`** | Application | Required to upload scored result workbooks (`P2-MasterFile.xlsx`) and candidate scorecards |
| **`Mail.Send`** | Application | *(Optional)* Required only if automated applicant emails or admin failure alerts are active |

5. Click **Add permissions**.
6. **⚠️ MANDATORY STEP**: Click **Grant admin consent for [Your Organization]** and confirm with "Yes".
   * Ensure all rows display a green checkmark in the **Status** column (*"Granted for..."*).

---

## 📁 Phase 2: SharePoint Addressing & Environment Variables

Add your credentials and SharePoint paths to `.env`:

```ini
# ── Microsoft Entra ID (Azure AD) Credentials ──────────────────────────────
TENANT_ID=00000000-0000-0000-0000-000000000000
CLIENT_ID=00000000-0000-0000-0000-000000000000
CLIENT_SECRET=your_client_secret_value_here

# ── SharePoint Addressing ──────────────────────────────────────────────────
# The tenant hostname (do not include https://)
SHAREPOINT_HOSTNAME=driverai.sharepoint.com

# URL path to the SharePoint site
SHAREPOINT_SITE_PATH=/sites/CandidateList_HiringAgent

# Name of the table inside the master Excel workbook
SHAREPOINT_TABLE=HiringAgent_P1_Candidates

# Location of the master Excel workbook on the SharePoint site
SHAREPOINT_WORKBOOK=/Master_Files/Sharepoint_Master_File.xlsx

# Root folder where P1 deposits incoming resumes
SHAREPOINT_RESUMES_FOLDER=/Candidate_Resumes

# Shared mailbox address for sending notifications (if enabled)
SENDER_MAILBOX=apply@driverai.io

# ── Safety Master Switch ───────────────────────────────────────────────────
# Keep true during testing to prevent live emails from being sent to applicants
HIRING_SUPPRESS_EMAILS=true

# Filter applicants to US geographic eligibility
HIRING_GEO_USA_ONLY=true
```

---

## 🧪 Testing Microsoft Connection

Test that your Entra ID credentials, permissions, SharePoint site, Excel table, and resume directories resolve correctly:

```bash
cd DriverAI_HiringAgent/HiringAgent_P2
python bot.py --test-sharepoint
```

Or via Docker:
```bash
docker compose run --rm cli --test-sharepoint
```

**Expected Successful Output**:
```text
✓ Microsoft Graph token acquired successfully
✓ SharePoint Site resolved: /sites/CandidateList_HiringAgent
✓ Candidate Excel table found: HiringAgent_P1_Candidates
✓ Resumes folder verified: /Candidate_Resumes
✓ Connection test PASSED: Ready for scoring
```

---

## 🚀 Phase 1: Power Automate Flow Deployment

To deploy the Phase 1 email intake flow to your Microsoft 365 tenant:

1. Sign in to **[Power Automate](https://make.powerautomate.com/)**.
2. In the left navigation, click **My flows**.
3. In the top bar, select **Import** ➔ **Import Package (Legacy)**.
4. Upload the pre-packaged zip:
   `DriverAI_HiringAgent/HiringAgent_P1/flow/DriverAI-Hiring-AutoReply-apply.zip`
5. During import, map the required connectors:
   * **Office 365 Outlook**: Authorize access to `apply@driverai.io`.
   * **SharePoint**: Select your target SharePoint site.
   * **Excel Online (Business)**: Connect to `Sharepoint_Master_File.xlsx`.
6. Click **Import**.
7. Once imported, open the flow, verify step configurations, and turn the flow **On**.
8. For deeper operational guidance, see the full [`P1_RUNBOOK.md`](../DriverAI_HiringAgent/HiringAgent_P1/docs/P1_RUNBOOK.md).
