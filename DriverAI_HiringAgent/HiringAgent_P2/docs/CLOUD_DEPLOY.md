# Cloud Deploy â€” DriverAI Hiring Agent (Phase 2)

> **Three serverless options** â€” no machine, no GUI, fully automated scoring.
> All three use the same `.env` credentials and the same engine. Pick one.

---

## Option A â€” GitHub Actions (free, no credit card)

The scheduled workflow at `.github/workflows/score.yml` runs `bot.py --score-sharepoint`
on GitHub's runners.

**Setup:**

1. Make `HiringAgent_P2/` the repo root:
   ```bash
   cd HiringAgent_P2
   git init && git add . && git commit -m "Hiring agent worker"
   gh repo create hiring-agent --private --source=. --push
   ```

2. Add credentials as **repo Secrets** (Settings â†’ Secrets and variables â†’ Actions â†’
   New repository secret):
   `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`, `SHAREPOINT_HOSTNAME`,
   `SHAREPOINT_SITE_PATH`, `SHAREPOINT_TABLE`, `SHAREPOINT_RESUMES_FOLDER`.
   Optional: `HIRING_GEO_USA_ONLY` variable.

3. **Change the schedule:** edit the `cron:` line in `.github/workflows/score.yml` â€”
   e.g. `'*/15 * * * *'` for 15 min, `'0 9,18 * * *'` for 9 AM + 6 PM UTC.
   Push the change; GitHub picks it up immediately.

4. **Test:** Actions tab â†’ *Score SharePoint queue* â†’ **Run workflow** (leave Dry run = true).
   The `Verify credentials` step prints PASS/FAIL.

5. Once it passes the schedule runs automatically â€” every ~30 min, $0, no machine.

**Free-tier notes:**
- Public repos: unlimited Actions minutes.
- Private repos: 2,000 min/month. With pip caching each run stays under 1 min.
- GitHub **disables scheduled workflows after 60 days of no repo activity** â€” push any commit
  to re-arm. Resumes live in SharePoint (not the repo), so a public repo is safe.

---

## Option B â€” Azure Functions (timer + HTTP, needs a credit card on file)

```bash
# 1) Create the Function App (Python 3.11, Consumption plan â€” free grant).
az group create -n HiringAgent-rg -l eastus
az storage account create -n hiringagentsa$RANDOM -g HiringAgent-rg -l eastus --sku Standard_LRS
az functionapp create -g HiringAgent-rg --consumption-plan-location eastus \
   --runtime python --runtime-version 3.11 --functions-version 4 \
   --name hiring-agent-fn --storage-account <storage-account-name>

# 2) Set Application Settings (env vars). SCORE_TIMER_CRON must be set.
#    Default: every 15 min (NCRONTAB with seconds field).
az functionapp config appsettings set -g HiringAgent-rg -n hiring-agent-fn --settings \
   TENANT_ID=<guid> CLIENT_ID=<guid> CLIENT_SECRET=<secret> \
   SHAREPOINT_HOSTNAME=led1234567.sharepoint.com \
   SHAREPOINT_SITE_PATH=/sites/CandidateList_HiringAgent \
   SHAREPOINT_TABLE=HiringAgent_P1_Candidates \
   SHAREPOINT_RESUMES_FOLDER=/Downloaded_Resumes \
   HIRING_GEO_USA_ONLY=true \
   SCORE_TIMER_CRON="0 */15 * * * *"

# 3) Deploy.
func azure functionapp publish hiring-agent-fn --python
```

**Triggers:**
- **Timer** (`score_timer`): fires on `SCORE_TIMER_CRON` only when `HIRING_ENABLE_P2_TIMER=true` and `HIRING_P2_DISABLED` is not true.
  Change it without redeployment:
  `az functionapp config appsettings set ... --settings SCORE_TIMER_CRON="0 */30 * * * *"`
- **HTTP** (`score_http`):
  `POST https://hiring-agent-fn.azurewebsites.net/api/score?code=<function-key>`
  Add `?dry_run=1` for a test run.
  Call this from the Phase 1 Power Automate flow (HTTP action) to score the moment a resume
  lands instead of waiting for the timer.

**Logs:** Application Insights captures all `logger.*` output.

---

## Option C â€” Docker (AWS, GCP, Fly.io, Railway, Render, any container host)

```bash
# Build and test locally.
cd HiringAgent_P2
docker build -t hiring-agent .
docker run --rm \
  -e TENANT_ID=<guid> -e CLIENT_ID=<guid> -e CLIENT_SECRET=<secret> \
  -e SHAREPOINT_HOSTNAME=led1234567.sharepoint.com \
  -e SHAREPOINT_SITE_PATH=/sites/CandidateList_HiringAgent \
  -e SHAREPOINT_TABLE=HiringAgent_P1_Candidates \
  -e SHAREPOINT_RESUMES_FOLDER=/Downloaded_Resumes \
  -e HIRING_GEO_USA_ONLY=true \
  hiring-agent python bot.py --score-sharepoint --dry-run

# Run as HTTP server (default â€” Fly.io, Railway, Render, Cloud Run, App Runner):
docker run -d -p 8080:8080 --env-file .env hiring-agent
# POST http://localhost:8080/score        trigger a scoring run
# POST http://localhost:8080/score?dry=1  dry run
# GET  http://localhost:8080/health       health check (returns {"status":"ok"})

# Or run as a polling worker (AWS ECS scheduled task, plain VPS):
docker run -d --env-file .env hiring-agent \
  python bot.py --score-sharepoint --watch --interval 900
```

**Platform quick starts:**

*AWS ECS / Fargate (scheduled task):*
```bash
aws ecr create-repository --repository-name hiring-agent
docker tag hiring-agent <account>.dkr.ecr.<region>.amazonaws.com/hiring-agent
docker push <account>.dkr.ecr.<region>.amazonaws.com/hiring-agent
# Create task definition + EventBridge rule; override CMD to:
# ["python", "bot.py", "--score-sharepoint"]
# Pass creds as task env vars or Secrets Manager references.
```

*Fly.io (free tier â€” 3 shared VMs):*
```bash
fly launch --name hiring-agent --no-deploy
fly secrets set TENANT_ID=<guid> CLIENT_ID=<guid> CLIENT_SECRET=<secret> \
  SHAREPOINT_HOSTNAME=... SHAREPOINT_SITE_PATH=... SHAREPOINT_TABLE=... SHAREPOINT_RESUMES_FOLDER=...
fly deploy
```

*Railway / Render:* push the repo, set environment variables in the dashboard, deploy.
Both detect `Dockerfile` automatically and expose port 8080.
Use their built-in cron to `POST /score` every 15â€“30 min.

**Scoring brain in Docker:** the offline regex parser (free, no key, no AI).
Ollama is off in Docker (`HIRING_SERVERLESS=1` set in Dockerfile), so extraction and
scoring are fully deterministic in the cloud.

---

## Shared environment variables

| Variable | Description |
|---|---|
| `TENANT_ID` | Azure AD tenant GUID (must be real tenant, NOT `common`) |
| `CLIENT_ID` | App registration Application ID |
| `CLIENT_SECRET` | Client secret value |
| `SHAREPOINT_HOSTNAME` | e.g. `led1234567.sharepoint.com` |
| `SHAREPOINT_SITE_PATH` | e.g. `/sites/CandidateList_HiringAgent` |
| `SHAREPOINT_TABLE` | e.g. `HiringAgent_P1_Candidates` |
| `SHAREPOINT_RESUMES_FOLDER` | e.g. `/Downloaded_Resumes` |
| `HIRING_GEO_USA_ONLY` | `true` / `false` â€” geo filter toggle (default `true`) |
| `SCORE_TIMER_CRON` | Azure Functions only: NCRONTAB schedule string |

---

## Pre-deploy verification

```bash
# Confirm config + roles load correctly.
python bot.py --list-roles

# Confirm extraction works on a sample resume.
python bot.py --score-file ./any_resume.pdf

# Verify Graph credentials and SharePoint access.
python bot.py --test-sharepoint

# Full end-to-end dry run (no writes to SharePoint).
python bot.py --score-sharepoint --dry-run

# Azure Functions local test:
# copy local.settings.json.example -> local.settings.json, fill in creds.
func start
curl -X POST http://localhost:7071/api/score?dry_run=1

# Docker local test:
docker run --rm --env-file .env hiring-agent python bot.py --test-sharepoint
```
