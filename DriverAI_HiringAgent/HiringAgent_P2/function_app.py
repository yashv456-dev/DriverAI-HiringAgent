"""Azure Functions entry point — the serverless 'enhance SharePoint' worker.

Two triggers, both run the SAME job (score the Phase 1 SharePoint queue and write back):
  * Timer  — polls on a schedule (default every 15 min; configurable via
             SCORE_TIMER_CRON app setting, e.g. "0 */30 * * * *" for 30 min).
  * HTTP   — POST /api/score, so Power Automate (Phase 1) can call it the moment a resume
             lands -> event-driven, "right after P1", no waiting for the timer.

Config comes from the Function App's Application settings (env vars): TENANT_ID, CLIENT_ID,
CLIENT_SECRET, SHAREPOINT_HOSTNAME, SHAREPOINT_SITE_PATH, SHAREPOINT_TABLE,
SHAREPOINT_RESUMES_FOLDER. The scoring brain is always the offline regex parser here — no
machine, no Ollama, no GUI, and no cloud-LLM hook (removed; it was never configured with a
key in this deployment). Regex extraction needs no AI, so serverless mode is unaffected.
"""

import json
import logging
import os

import azure.functions as func

from hiring_agent.config import SHAREPOINT_CONFIGURED, logger
from hiring_agent.sharepoint_scoring import score_from_sharepoint

app = func.FunctionApp()

# SCORE_TIMER_CRON must exist as an Azure Function App Application Setting.
# The %...% below is an Azure binding-expression resolved at startup, not a Python default.
# If the setting is absent the function app will fail to start with a binding error.
# Default to add in the portal: SCORE_TIMER_CRON = 0 */15 * * * *
SCORE_TIMER_CRON = os.getenv("SCORE_TIMER_CRON", "0 */15 * * * *")


def _truthy_env(name: str) -> bool:
    return str(os.getenv(name, "") or "").strip().lower() in {"1", "true", "yes", "on"}


def _run(dry_run: bool = False) -> dict:
    if _truthy_env("HIRING_P2_DISABLED"):
        msg = "P2 scoring is disabled by HIRING_P2_DISABLED."
        logger.warning(msg)
        return {"ok": True, "disabled": True, "message": msg}
    if not SHAREPOINT_CONFIGURED:
        msg = ("SharePoint not configured: set TENANT_ID, CLIENT_ID, CLIENT_SECRET, "
               "SHAREPOINT_HOSTNAME, SHAREPOINT_SITE_PATH, SHAREPOINT_TABLE in the Function "
               "App's Application settings.")
        logger.error(msg)
        return {"ok": False, "error": msg}
    try:
        result = score_from_sharepoint(dry_run=dry_run)
        return {"ok": "error" not in result, **result}
    except Exception as e:                     # never let one bad run kill the host
        logger.exception("score_from_sharepoint failed")
        return {"ok": False, "error": str(e)}


@app.timer_trigger(schedule="%SCORE_TIMER_CRON%", arg_name="timer",
                   run_on_startup=False, use_monitor=True)
def score_timer(timer: func.TimerRequest) -> None:
    """Poll the SharePoint queue on a schedule (default: every 15 min)."""
    if not _truthy_env("HIRING_ENABLE_P2_TIMER"):
        logging.info("Timer trigger skipped: set HIRING_ENABLE_P2_TIMER=true to enable P2 scheduling.")
        return
    logging.info("Timer trigger: scoring SharePoint queue.")
    _run(dry_run=False)


@app.route(route="score", auth_level=func.AuthLevel.FUNCTION)
def score_http(req: func.HttpRequest) -> func.HttpResponse:
    """POST /api/score?dry_run=0 — run on demand (e.g. called by Power Automate)."""
    dry = (req.params.get("dry_run", "") or "").strip().lower() in {"1", "true", "yes"}
    result = _run(dry_run=dry)
    status = 200 if result.get("ok") else 500
    return func.HttpResponse(json.dumps(result), status_code=status,
                             mimetype="application/json")
