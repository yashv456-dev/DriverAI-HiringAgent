"""
build_zip.py  —  Generate the DriverAI Hiring Auto-Reply PA flow from flow_config.json.

flow_config.json is the SINGLE SOURCE OF TRUTH. This script reads it, overlays every
tunable value onto the deployable flow definition, and repackages the importable .zip.
Edit flow_config.json, run this, then re-import the zip in Power Automate. You should
never hand-edit definition.json - it is regenerated here (and re-synced for readability).

What is config-driven (from flow_config.json):
  - trigger mailbox + poll interval
  - applicant reply "from" mailbox + admin failure-alert recipient
  - SharePoint site / library / workbook / resumes folder
  - Excel Online addressing (source / drive / file / table) on all 4 row actions
  - duplicate-block window (days) and CV-request cap
  - application-reference date format + reply-detection pattern
  - bad senders / bad subjects / spam phrases / application keywords

Loop protection: the flow can never answer itself because every send is FROM the trigger
mailbox, which is listed in bad_senders (alongside mailer-daemon / postmaster / bounce).
Outbound subjects are NOT added to bad_subjects - doing so would drop applicants' "Re: ..."
replies before they reach the reference gate. Per-applicant caps stop any runaway loop.

Idempotent: every value is OVERWRITTEN (not appended), so re-running is safe.

Run:  python build_zip.py     (from HiringAgent_P1/flow/)
"""
import copy
import json
import zipfile
from pathlib import Path

HERE = Path(__file__).parent
ZIP = HERE / "DriverAI-Hiring-AutoReply-apply.zip"
CONFIG_FILE = HERE / "flow_config.json"
# ── load config (single source of truth) ─────────────────────────────────────
cfg = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
email, trig_cfg = cfg["email"], cfg["trigger"]
send_applicant_emails = bool(email.get("send_applicant_emails", True))
send_admin_failure_alerts = bool(email.get("send_admin_failure_alerts", True))
rules, sp, ex = cfg["business_rules"], cfg["sharepoint"], cfg["excel"]
appref, filt = cfg["appref"], cfg["spam_filters"]


# ── local time (2026-08-27) ──────────────────────────────────────────────────
# Outlook's receivedDateTime and utcNow() are both UTC. Every date P1 derives from them
# is user-facing - the Application ID, the two date columns, the Year/Month resume folder,
# the year/month separator rows - so all of them belong in local wall-clock time. Arizona
# is UTC-7, so a mail received before 07:00 UTC belongs to the PREVIOUS local day; on the
# 1st of a month that put the resume in the wrong month folder (a real applicant, 2026-05-01
# 06:41 UTC = 2026-04-30 23:41 local, filed under May instead of April).
#
# convertFromUtc() is used rather than convertTimeZone() because the source is always UTC.
# The format string is fixed at the ISO shape the Excel columns already use, so a converted
# value stays parseable by ticks(), by formatDateTime() downstream, and by P2's date parser.
#
# id='UTC' skips the wrapper entirely instead of emitting a no-op conversion, so the old
# behaviour is restorable byte-for-byte.
_TZ_ID = (cfg.get("timezone") or {}).get("id", "UTC")
_TZ_OFFSET_H = int((cfg.get("timezone") or {}).get("utc_offset_hours", 0) or 0)


def _local(utc_expr: str) -> str:
    """Wrap a UTC timestamp expression so it reads as local wall-clock time.

    Uses addHours(), NOT convertFromUtc(). convertFromUtc was tried and fails at runtime on
    this connector's values: it demands a first parameter that declares itself UTC, and
    rejected the Outlook timestamp outright ("expects its first parameter to be a string
    that contains a UTC date time. The value was '2026-05-01T06:41:25.0000000'"). Wrapping
    it in formatDateTime() to normalise the shape made it worse - that strips the Z/offset
    convertFromUtc relies on to recognise the value as UTC in the first place. addHours()
    parses whatever the connector emits, offset or not, and cannot fail this way.

    EXACT ONLY FOR A NO-DST ZONE. Arizona is UTC-7 every day of the year, so a fixed offset
    is not an approximation here. A DST zone would need a real timezone function.

    ONLY for genuinely-UTC expressions - the trigger's receivedDateTime and utcNow(). A value
    read back OUT of the workbook was written local by this same build and must NOT be passed
    through here, or it shifts a second time.
    """
    if _TZ_OFFSET_H == 0:
        return utc_expr
    return "addHours(%s,%d,'yyyy-MM-ddTHH:mm:ss')" % (utc_expr, _TZ_OFFSET_H)


# The trigger's receivedDateTime, in local time. Section 19 rewrites the inner
# triggerOutputs()?['body/...'] to outputs('CurrentEmail')?['...'] by plain substring
# replacement, which reaches inside the convertFromUtc() wrapper unchanged.
_RECEIVED_UTC = "triggerOutputs()?['body/receivedDateTime']"
_RECEIVED_LOCAL = _local(_RECEIVED_UTC)
# The separator actions are authored directly against outputs('CurrentEmail') rather than
# triggerOutputs(), so they need their own already-rewritten twin.
_RECEIVED_LOCAL_CE = _local("outputs('CurrentEmail')?['receivedDateTime']")
_NOW_LOCAL = _local("utcNow()")
year_sep = cfg.get("year_separator", {"enabled": True})
month_sep = cfg.get("month_separator", {"enabled": True})
flow_name = cfg.get("flow_name", "DriverAI - Hiring Auto-Reply P1")

# ── read the current package (structural base; keeps GUIDs/connections intact) ─
# The base zip's per-flow folder may be under a DIFFERENT GUID than the one we are writing
# (that is the whole point of flow_id: retarget the package at the live flow). Detect the
# folder that is actually in the archive rather than assuming it matches, otherwise a
# retarget makes the very first read fail.
with zipfile.ZipFile(ZIP, "r") as zf:
    _existing = [n for n in zf.namelist()
                 if n.startswith("Microsoft.Flow/flows/") and n.endswith("/definition.json")]
    if not _existing:
        raise SystemExit(f"{ZIP} has no Microsoft.Flow/flows/<guid>/definition.json")
    READ_BASE = _existing[0][: -len("definition.json")]
    _BASE_GUID = READ_BASE.split("/")[-2]        # the GUID actually present in the base zip
    defn = json.loads(zf.read(READ_BASE + "definition.json").decode("utf-8"))
    apis_map = zf.read(READ_BASE + "apisMap.json").decode("utf-8")
    conns_map = zf.read(READ_BASE + "connectionsMap.json").decode("utf-8")
    flow_mf = zf.read("Microsoft.Flow/flows/manifest.json").decode("utf-8")
    root_mf = zf.read("manifest.json").decode("utf-8")

# The LIVE flow id (flow_config.json: flow_id). Power Automate binds an "Update" import by
# this GUID; a stale one leaves the update nothing to land on, so the import appears to
# succeed while the running flow keeps its OLD definition - three imports on 2026-08-27
# changed nothing because of exactly that.
#
# Computed HERE, not earlier: the documented fallback is "whatever GUID the base zip carries",
# and _BASE_GUID only exists once the archive has been opened above. Assigning this before the
# read left the fallback referencing an unbound name - harmless only while flow_id is set,
# because `or` short-circuits, and a NameError the moment anyone blanked it.
GUID = str(cfg.get("flow_id") or _BASE_GUID).strip()
BASE = f"Microsoft.Flow/flows/{GUID}/"

# Enforce package identity: the definition's name/id MUST equal the asset-folder GUID
# referenced by the manifests. A definition re-exported from the tenant carries the
# tenant-assigned flow id instead, and that mismatch fails import with
# MalformedFlowAssetFlowDefinition. Overwrite both every build so the zip base can
# never drift again.
defn["name"] = GUID
defn["id"] = f"/providers/Microsoft.Flow/flows/{GUID}"

wd = defn["properties"]["definition"]
actions = wd["actions"]

# The ZIP is also the next build's structural base. The first unread-poll build
# wrapped the whole processing graph inside Has_unread_email; Power Automate rejected
# that package because the extra control level pushed existing level-8 actions to
# level 9. Unwrap that one historical shape, or strip the newer flat poll plumbing,
# before reapplying overlays. Repeated builds therefore stay idempotent.
if "Has_unread_email" in actions and "CONFIG" in actions["Has_unread_email"].get("actions", {}):
    _saved_branch = actions["Has_unread_email"]["actions"]
    _saved_branch.pop("CurrentEmail", None)
    actions = _saved_branch
else:
    for _poll_action in (
        "Get_unread_emails",
        "Has_unread_email",
        "CurrentEmail",
        "Notify_poll_failure",
        "Notify_processing_scope_failure",
        "Move_unhandled_failure_to_recruiting_review",
    ):
        actions.pop(_poll_action, None)
if "CONFIG" in actions:
    actions["CONFIG"]["runAfter"] = {}
wd["actions"] = actions


# ── helpers ──────────────────────────────────────────────────────────────────
def set_email_mailbox(node) -> None:
    """Recursively point every SharedMailbox send at the configured trigger mailbox."""
    if isinstance(node, dict):
        params = node.get("inputs", {}).get("parameters") if isinstance(node.get("inputs"), dict) else None
        if isinstance(params, dict) and "emailMessage/MailboxAddress" in params:
            params["emailMessage/MailboxAddress"] = email["trigger_mailbox"]
        for v in node.values():
            set_email_mailbox(v)
    elif isinstance(node, list):
        for v in node:
            set_email_mailbox(v)


def set_excel(node) -> None:
    """Point one Excel row action at the configured workbook/table.

    Import-time behaviour (learned the hard way, 2026-07-02):
    Excel AddRowV2/PatchItem/GetItems carry a dynamic column schema on 'item'/'$filter'.
    When the legacy package import saves the flow, PA resolves that schema by calling
    GetTable with the action's source+drive+file+table values. There is NO "skip the
    check" path: if any of the four is an expression, PA cannot build the GetTable
    request and the import dies with DynamicParameterInputInvalid / "The API operation
    'GetTable' is missing required property 'table'". Every expression variant failed:
      - @outputs('CONFIG')[...]                     (the all-CONFIG version)
      - concat('<path>', substring(utcNow(),0,0))   (failed import 2026-07-02)
      - if(greater(ticks(utcNow()),0),'<path>','<decoy>')  (failed import 2026-07-02)
    The earlier belief that a utcNow() in 'file' defers the check came from ne3's
    tenant EXPORT (runtime accepts expressions fine) - it was never proven on import.

    So: ALL FOUR params must be static literals. The import-time GetTable then really
    executes against the tenant and succeeds as long as the workbook + table exist at
    exactly this path (import the package AFTER the workbook is in place; the
    EnsureWb_Probe self-heal only covers runtime, not import). Community failures of
    this check surface as GetTable 'NotFound'/'Unauthorized' - if you see those, the
    static path/table here doesn't match the tenant or the chosen connection lacks
    access.
    """
    p = node["inputs"]["parameters"]
    p["source"] = ex["source"]
    p["drive"] = ex["drive"]
    p["table"] = ex["table"]
    p["file"] = ex["file"]


def resume_folder_path(rd: str = None) -> str:
    """folderPath expression for resume Create-file actions.

    Flat = the configured resumes folder. Dated = .../<yyyy>/<MMMM> computed from the date
    expression `rd` (default: the email's receivedDateTime) — mirrors hiring_agent.config.
    dated_subpath so Phase 2 reconstructs the identical path. Pass a row's 'Received Date' to
    file an updated CV into the ORIGINAL application's folder so Phase 2 (which rebuilds the
    path from that same date) finds it. No week-level subfolder (dropped 2026-07-08 — one
    folder per month is enough at this volume and made the tree harder to browse).
    """
    rd = _RECEIVED_LOCAL if rd is None else rd
    if not sp.get("dated_resume_subfolders", False):
        return "@outputs('CONFIG')['resumes_folder']"
    return (
        "@{concat(outputs('CONFIG')['resumes_folder'], "
        "'/', formatDateTime(%s,'yyyy'), "
        "'/', formatDateTime(%s,'MMMM'))}" % (rd, rd)
    )


def resume_folder_path_lib_relative(rd: str = None) -> str:
    """Library-relative twin of resume_folder_path() (no '/Shared Documents' prefix) - the shape
    CreateNewFolder's 'parameters/path' needs (same fix as Create_workbook_folder). CreateFile does
    NOT auto-create missing folders (confirmed: it fails with NotFound) - so the very first resume
    of a new year/month needs its dated folder created explicitly before Create_file runs."""
    rd = _RECEIVED_LOCAL if rd is None else rd
    if not sp.get("dated_resume_subfolders", False):
        return "@{'%s'}" % _sf_folder_lib_relative
    return (
        "@{concat('%s', "
        "'/', formatDateTime(%s,'yyyy'), "
        "'/', formatDateTime(%s,'MMMM'))}" % (_sf_folder_lib_relative, rd, rd)
    )


def _make_ensure_resume_folder(rd: str) -> dict:
    """CreateNewFolder action guaranteeing the dated resume subfolder exists before a Create_file
    call. Succeeds silently if the folder already exists (same behavior as Create_workbook_folder)."""
    return {
        "type": "OpenApiConnection",
        "runAfter": {},
        "inputs": {
            "parameters": {
                "dataset": sp["site"],
                "table": sp["documents_library_id"],
                "parameters/path": resume_folder_path_lib_relative(rd),
            },
            "host": {
                "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
                "connectionName": "shared_sharepointonline",
                "operationId": "CreateNewFolder",
            },
            "authentication": "@parameters('$authentication')",
        },
    }


def _make_folder_wait(ensure_action_name: str) -> dict:
    """Short pause after creating a dated resume subfolder, before saving a file into it - same
    propagation-lag reasoning as Wait_after_workbook_heal, just a shorter wait since a folder create
    is a much lighter write than uploading a whole workbook."""
    return {
        "type": "Wait",
        "runAfter": {ensure_action_name: ["Succeeded", "Failed"]},
        "inputs": {"interval": {"count": int(sp.get("folder_wait_seconds", 2)), "unit": "Second"}},
    }


def _deep_find(tree, name):
    """Find an action dict by name anywhere in a nested action tree."""
    if isinstance(tree, dict):
        if name in tree:
            return tree[name]
        for v in tree.values():
            r = _deep_find(v, name)
            if r is not None:
                return r
    return None


def _find_owner_actions(tree, name):
    """Return the `actions` dict that DIRECTLY contains `name`, or None.

    _deep_find hands back the action itself; this hands back its parent map, which is what
    you need to insert a true SIBLING. runAfter may only reference a literal sibling at the
    same nesting level, so anything wired to an existing action must be added here.
    """
    if isinstance(tree, dict):
        for key, value in tree.items():
            if key == "actions" and isinstance(value, dict):
                if name in value:
                    return value
            # The top-level `actions` map is passed in bare, so check it directly too.
            if name in tree and not isinstance(tree.get(name), str):
                return tree
            r = _find_owner_actions(value, name)
            if r is not None:
                return r
    elif isinstance(tree, list):
        for v in tree:
            r = _find_owner_actions(v, name)
            if r is not None:
                return r
    return None


# ── self-healing pristine restore for email-sending actions ──────────────────
# The zip is BOTH the deployable output AND the next run's structural base, and nothing
# else in this script reconstructs Notify_failure/Send_* from scratch - _set_email() and
# Notify_failure's setup only overlay Subject/Body onto an ASSUMED-real base. If a prior
# build suppressed outbound actions (replacing them with no-op Compose actions), every
# subsequent build - suppressed or not - would find a corrupted
# base and crash trying to set dynamic-content fields on a Compose action. Force these 7
# back to their real connector shape (host/apiId/operationId/authentication) from a
# permanent pristine snapshot at the START of every single build, before anything else
# touches them. Makes suppression fully reversible and every build idempotent regardless
# of what the previous run left behind - discovered the hard way when a first pass at
# this feature corrupted the live zip on its second build.
_PRISTINE_FILE = HERE / "_pristine_send_actions.json"
_pristine_shapes = json.loads(_PRISTINE_FILE.read_text(encoding="utf-8"))
for _name, _real_shape in _pristine_shapes.items():
    _act = _deep_find(wd["actions"], _name)
    if _act is not None:
        _act.clear()
        _act.update(copy.deepcopy(_real_shape))


# ── 1. trigger: mailbox + poll interval ──────────────────────────────────────
interval = int(trig_cfg["interval_min"])
unread_per_run = int(trig_cfg.get("unread_per_run", 1))
# False (the default since 2026-08-04) makes the INBOX the queue rather than the read
# flag: every terminal path already moves the message out of the Inbox, whereas merely
# OPENING a message used to drop it from the queue silently and permanently.
# See _comment_fetch_only_unread in flow_config.json.
fetch_only_unread = bool(trig_cfg.get("fetch_only_unread", False))
if unread_per_run != 1:
    raise ValueError(
        "trigger.unread_per_run must be 1: P1 intentionally processes one unread email per run"
    )
wd["triggers"] = {
    "Poll_unread_shared_mailbox": {
        "recurrence": {"frequency": "Minute", "interval": interval},
        "evaluatedRecurrence": {"frequency": "Minute", "interval": interval},
        "type": "Recurrence",
        "runtimeConfiguration": {"concurrency": {"runs": 1}},
    }
}

# ── 1b. flow + package display name (config-driven; fixes the "new2" package name) ──
defn["properties"]["displayName"] = flow_name
_root = json.loads(root_mf)
if isinstance(_root.get("details"), dict):
    _root["details"]["displayName"] = flow_name          # package name (import dialog / list)
# Re-key the flow resource from whatever GUID the base zip carried onto the LIVE flow id.
# The package folder name, defn["id"] and this manifest key must all agree, or the import
# dialog cannot offer "Update" against the real flow and silently falls back to creating a
# new one - three imports on 2026-08-27 changed nothing because of exactly this.
_resources = _root.setdefault("resources", {})
if GUID != _BASE_GUID and _BASE_GUID in _resources:
    _resources[GUID] = _resources.pop(_BASE_GUID)
_res = _resources.get(GUID)
if isinstance(_res, dict) and isinstance(_res.get("details"), dict):
    _res["details"]["displayName"] = flow_name           # flow resource name
root_mf = json.dumps(_root, ensure_ascii=False)

# ── 2. CONFIG compose (SharePoint paths used by Create_file) ──────────────────
# Only sharepoint_site and resumes_folder are referenced in flow expressions. Excel
# addressing is written directly onto each Excel action (see set_excel), NOT via CONFIG:
# PA constant-folds a CONFIG literal at import time, which re-triggers the GetTable check.
actions["CONFIG"]["inputs"] = {
    "sharepoint_site": sp["site"],
    "resumes_folder": sp["resumes_folder"],
}

# ── 3. full-body + attachment-name keyword/spam scan ─────────────────────────
# AttachNames = just the attachment FILE NAMES (never contentBytes). Folding them into the
# lowercased scan field lets the spam gate catch a dangerous attachment (e.g. "invoice.exe")
# via the malware_phrases extensions, without adding a separate gate. Empty attachments -> ''.
actions["AttachNames"] = {
    "runAfter": {"LowerSubject": ["Succeeded"]},
    "type": "Select",
    "inputs": {
        "from": "@coalesce(triggerOutputs()?['body/attachments'], json('[]'))",
        "select": "@item()?['name']",
    },
}
actions["LowerBody"]["runAfter"] = {"AttachNames": ["Succeeded"]}
# The TRAILING space after the joined attachment names is load-bearing: malware_phrases'
# executable extensions are stored space-terminated ('.exe ') so they match a real attachment
# ('invoice.exe' -> '...invoice.exe ') but not prose that merely contains the letters
# ('www.exeter.ac.uk' -> '.exet', 'www.iso.org' -> '.iso.'). Without this final space a lone
# attachment at the end of the field would not be space-terminated and would stop matching.
actions["LowerBody"]["inputs"] = (
    "@toLower(concat(coalesce(triggerOutputs()?['body/body'], "
    "triggerOutputs()?['body/bodyPreview'], ''), ' ', join(body('AttachNames'), ' '), ' '))"
)

# ── 4. spam/keyword lists ────────────────────────────────────────────────────
actions["BadSenders"]["inputs"] = list(filt["bad_senders"])
# Loop protection is SENDER-based: apply@driverai.io (plus mailer-daemon/postmaster/bounce)
# is in bad_senders so the flow never answers itself or a bounce. We do NOT add our own
# outbound subjects to bad_subjects — an applicant's "Re: <our subject>" reply must reach
# the reference gate. Per-applicant caps stop any runaway loop.
bad_subjects = list(filt["bad_subjects"])
actions["BadSubjects"]["inputs"] = bad_subjects
# Spam gate also enforces content safety: scam + offensive/harassment + malware/exe links +
# non-English scam + vendor/agency solicitation, all checked (lowercased substring) against
# subject, body, AND attachment names. Any match => Terminate_Spam (silent, no reply). Lists
# live in flow_config.json. vendor_solicitation_phrases added 2026-08-01 (Futurism
# Technologies case): a services vendor pitching itself to us, not applying for a job,
# triggered none of the other lists since nothing in the email was malicious.
actions["SpamPhrases"]["inputs"] = (
    list(filt["spam_phrases"])
    + list(filt.get("offensive_phrases", []))
    + list(filt.get("malware_phrases", []))
    + list(filt.get("foreign_scam_phrases", []))
    + list(filt.get("vendor_solicitation_phrases", []))
)
# link_shortener_phrases are a SEPARATE gate (see HasValidResumeEarly above): checked always, but
# bypassed when a real resume is attached, since a shortened social/portfolio link in a normal
# email signature was silently blocking genuine applicants alongside actual malware/phishing links.
actions["LinkShortenerPhrases"] = {
    "runAfter": {"SpamPhrases": ["Succeeded"]},
    "type": "Compose",
    "inputs": list(filt.get("link_shortener_phrases", [])),
}
# Keywords Compose is dead — its output is never referenced; keyword matching is done
# inline in the Has_application_keyword expression. Remove it and re-anchor MatchSender.
actions.pop("Keywords", None)

# ── allow-list: internal testers must never be silently junked ───────────────
# Mail from these addresses skips ALL THREE spam screens and goes straight into normal
# processing. Without it, a live test from an internal account can be moved to Junk with no
# row, no reply and no alert - the documented blind spot - and the tester has no way to tell
# that from a genuine processing failure.
#
# These two Compose/Query actions sit at the TOP level. IsSubject and IsSpam are nested and
# therefore cannot list MatchAllowSender in their runAfter (PA rejects cross-level runAfter,
# see the ResumeFilesEarly note below), but they CAN read its output in an expression - the
# same thing MatchLinkShortener already does with the top-level LowerSubject/LowerBody. The
# ordering is safe regardless: MatchAllowSender gates IsSystem, which precedes both.
allow_senders = [str(x).strip().lower() for x in filt.get("allow_senders", []) if str(x).strip()]
if "apply@driverai.io" in allow_senders or email["trigger_mailbox"].lower() in allow_senders:
    raise ValueError(
        "spam_filters.allow_senders must not contain the trigger mailbox: allow-listing it "
        "defeats the self-loop guard in bad_senders and the flow would answer its own mail."
    )
actions["AllowSenders"] = {
    "runAfter": {"LinkShortenerPhrases": ["Succeeded"]},
    "type": "Compose",
    "inputs": allow_senders,
}
actions["MatchAllowSender"] = {
    "runAfter": {"AllowSenders": ["Succeeded"]},
    "type": "Query",
    "inputs": {
        "from": "@outputs('AllowSenders')",
        "where": "@contains(outputs('LowerFrom'), item())",
    },
}
# QuotedRef: the APP-... reference the sender actually typed, pulled out of the subject (or the
# body if it is not in the subject). An Application ID is exactly 22 characters - APP- + 8 date
# + - + 4 time + - + 4 hex - so a fixed-length substring from the 'app-20' marker is enough; the
# length guard stops substring() overrunning when the marker sits at the very end of a truncated
# body. Empty string when nothing usable is found, which the Get_rows_ref filter treats as absent.
_SUBJ = "outputs('LowerSubject')"
_BODY = "outputs('LowerBody')"


def _extract_ref(src):
    return ("if(and(contains({s},'app-20'), greaterOrEquals(length({s}), "
            "add(indexOf({s},'app-20'),22))), substring({s}, indexOf({s},'app-20'), 22), '')"
            ).format(s=src)


actions["QuotedRef"] = {
    "runAfter": {"MatchAllowSender": ["Succeeded"]},
    "type": "Compose",
    "inputs": "@toUpper(if(not(equals(%s,'')), %s, %s))" % (
        _extract_ref(_SUBJ), _extract_ref(_SUBJ), _extract_ref(_BODY)),
}
actions["MatchSender"]["runAfter"] = {"QuotedRef": ["Succeeded"]}

# An empty allow list makes this clause always true, so the gates behave exactly as before.
_NOT_ALLOWLISTED = {"equals": ["@length(body('MatchAllowSender'))", 0]}

actions["IsSystem"]["expression"] = {"and": [
    {"greater": ["@length(body('MatchSender'))", 0]},
    _NOT_ALLOWLISTED,
]}
actions["IsSystem"]["else"]["actions"]["IsSubject"]["expression"] = {"and": [
    {"greater": ["@length(body('MatchSubject'))", 0]},
    _NOT_ALLOWLISTED,
]}
# ResumeFilesEarly/HasValidResumeEarly are dead at the top level - an earlier build put them here,
# but they had to move into spam_gate (below) as same-level siblings of IsSpam, since PA rejects a
# runAfter reference from a nested action to anything not at its exact same level, even a top-level
# action. build_zip.py's base zip carries stale top-level entries forward otherwise (same class of
# leftover as the old "Keywords" action above) - pop them so they don't linger as dead duplicates.
actions.pop("ResumeFilesEarly", None)
actions.pop("HasValidResumeEarly", None)

# ── navigate to the spam-gate scope and the post-spam work branch ────────────
# NOTE (fixed 2026-07-03, 2nd attempt): runAfter can ONLY reference actions at the exact same
# nesting level - PA's validator rejected referencing the top-level HasValidResumeEarly from the
# deeply-nested IsSpam ("must belong to same level as action 'IsSpam'"), even though it was
# always guaranteed to have completed first. So ResumeFilesEarly/HasValidResumeEarly must live
# HERE, as direct siblings of IsSpam/MatchSpam/MatchLinkShortener, not at the top level.
spam_gate = actions["IsSystem"]["else"]["actions"]["IsSubject"]["else"]["actions"]
spam_gate["ResumeFilesEarly"] = {
    "runAfter": {},
    "type": "Query",
    "inputs": {
        "from": "@coalesce(triggerOutputs()?['body/attachments'], json('[]'))",
        "where": "@and(or(endsWith(toLower(item()?['name']), '.pdf'), "
                 "endsWith(toLower(item()?['name']), '.docx')), "
                 "not(empty(item()?['contentBytes'])))",
    },
}
spam_gate["HasValidResumeEarly"] = {
    "runAfter": {"ResumeFilesEarly": ["Succeeded"]},
    "type": "Compose",
    "inputs": "@greater(length(body('ResumeFilesEarly')), 0)",
}
spam_gate["MatchLinkShortener"] = {
    "runAfter": {},
    "type": "Query",
    "inputs": {
        "from": "@outputs('LinkShortenerPhrases')",
        "where": "@or(contains(outputs('LowerSubject'), item()), contains(outputs('LowerBody'), item()))",
    },
}
spam_gate["IsSpam"]["runAfter"] = {
    "MatchSpam": ["Succeeded"],
    "MatchLinkShortener": ["Succeeded"],
    "HasValidResumeEarly": ["Succeeded"],
}
spam_gate["IsSpam"]["expression"] = {"and": [
    {"or": [
        {"greater": ["@length(body('MatchSpam'))", 0]},
        {"and": [
            {"greater": ["@length(body('MatchLinkShortener'))", 0]},
            {"equals": ["@outputs('HasValidResumeEarly')", False]},
        ]},
    ]},
    # allow-listed internal testers bypass the content screen too (see AllowSenders above)
    _NOT_ALLOWLISTED,
]}

spam = spam_gate["IsSpam"]["else"]["actions"]
hr = spam["HasResume"]
dup = hr["actions"]["Is_duplicate"]
kw = hr["else"]["actions"]["Has_application_keyword"]
gate = spam["HasAppRefGate"]

# ── 5. application reference: UNIQUE per email = APP-<date>-<time>-<hex> ───────
# Built from the email's received timestamp plus a short random hex tail. The timestamp keeps
# refs readable and roughly sortable; the guid() tail guarantees UNIQUENESS even for two emails
# that arrive in the exact same second (concurrency=1 runs them sequentially, but their
# receivedDateTime can be identical — which would otherwise collide the Application ID row key
# and the saved resume filename). toUpper() so the ref round-trips through the reply-extraction's
# toUpper() and matches the stored Application ID and the saved <ref>_<file> name exactly.
_time_fmt = appref.get("time_format", "HHmm")   # PA formatDateTime token, e.g. "HHmm" or "HHmmss"
_suffix_len = int(appref.get("hex_length", 4))   # length of the tail sliced off the message id
_tail_skip = int(appref.get("id_tail_skip", 2))  # low-entropy trailing chars to step over
#   Ref shape: APP-20260624-2238-A5F2 (~22 chars at the defaults). The minute timestamp + 4-hex
#   tail still makes a same-minute collision astronomically unlikely at this volume, and the '-'
#   delimiters keep it readable. _ref_len below recomputes automatically from _time_fmt/_suffix_len,
#   so minting and reply-extraction can never drift apart even if these config values change.
_rd = _local("coalesce(triggerOutputs()?['body/receivedDateTime'], utcNow())")
# The reference must be DETERMINISTIC: the same email has to mint the same ref every time
# it is processed, or a retry saves the resume under one ref and keys the row to another.
# guid() cannot do that. The message's own Graph id can - see AppRefSeed.
#
# Section 19 rewrites triggerOutputs()?['body/id'] to outputs('CurrentEmail')?['id'], the
# same way it does for receivedDateTime.
spam["AppRefSeed"] = {
    "runAfter": {},
    "type": "Compose",
    # Alphanumerics only (a base64url id also carries '-', '_' and '=' padding), left-padded
    # so the slice below can never run off the front of a short or missing value.
    "inputs": ("@{concat('%s', replace(replace(replace("
               "coalesce(triggerOutputs()?['body/id'],''),'-',''),'_',''),'=',''))}"
               % ("0" * (_suffix_len + _tail_skip))),
}
spam["AppRef"]["runAfter"] = {"AppRefSeed": ["Succeeded"]}
# The tail sits _tail_skip characters from the END - measured across 241 real messages in
# this mailbox: the first, middle and last 4 characters of an Outlook id are IDENTICAL on
# every message (they are the mailbox guid and the trailing 'AAA' padding), while the four
# characters at offset -6 are the per-message counter and gave 185 distinct values. Two
# emails in the same minute are consecutive, so they always differ there.
spam["AppRef"]["inputs"] = (
    "@{concat('APP-', formatDateTime(%s,'%s'), '-', formatDateTime(%s,'%s'), '-', "
    "toUpper(substring(outputs('AppRefSeed'), "
    "sub(length(outputs('AppRefSeed')), %d), %d)))}"
    % (_rd, appref["date_format"], _rd, _time_fmt, _suffix_len + _tail_skip, _suffix_len)
)
# The old row-counting actions are gone; drop them if the base zip still carries them.
for _dead in ("TodayPrefix", "Get_rows_count", "Filter_today"):
    spam.pop(_dead, None)

# FileRef: filename-safe variant (/ replaced with -)
spam["FileRef"] = {
    "runAfter": {"AppRef": ["Succeeded"]},
    "type": "Compose",
    "inputs": "@{replace(outputs('AppRef'), '/', '-')}",
}

# ResumeFiles now waits for FileRef
# ── attachment triage: resume yes, portfolio/cover letter no (2026-08-27) ──────────
# Three Query actions rather than one, because a Query cannot fall back to its own input:
#   ResumeFilesAll   every .pdf/.docx with real bytes            (what P1 used to use)
#   ResumeFilesKept  the same list minus the excluded filenames
#   ResumeFiles      whichever of those two is non-empty         (unchanged name, so the
#                    three Foreach loops and ResumeNames need no edits at all)
# The fallback is the safety property: if the exclusion would leave nothing, the unfiltered
# list is used, so an applicant whose ONLY attachment is called portfolio.pdf is still a
# normal application rather than a "please attach your resume" reply.
_ATTACH_EXCLUDE = [str(p).strip().lower()
                   for p in (cfg.get("resume_attachments") or {}).get("exclude_name_phrases", [])
                   if str(p).strip()]

_IS_RESUME_FILE = ("@and(or(endsWith(toLower(item()?['name']), '.pdf'), "
                   "endsWith(toLower(item()?['name']), '.docx')), "
                   "not(empty(item()?['contentBytes'])))")

spam["ResumeFilesAll"] = {
    "type": "Query",
    "runAfter": {"FileRef": ["Succeeded"]},
    "inputs": {
        "from": "@coalesce(triggerOutputs()?['body/attachments'], json('[]'))",
        "where": _IS_RESUME_FILE,
    },
}

if _ATTACH_EXCLUDE:
    # Logic Apps' or() is BINARY - or(a, b), exactly two arguments. A flat or(a, b, c, ...)
    # is rejected at RUNTIME, not at import: the flow imports clean, then every run dies
    # inside this Query. That took out a live run on 2026-08-27 (a real applicant: no row, no
    # file, no acknowledgment, "Reference: n/a" because AppRef never resolved, and the mail
    # quarantined unread to Archive). Every pre-existing multi-term test in this flow is
    # either a Query over the term list (MatchAllowSender, MatchSpam) or a binary or() -
    # there was no flat multi-arg or() anywhere before this one. Fold right-associatively
    # into nested binary or() calls instead: or(a, or(b, or(c, d))).
    def _or_chain(terms: list) -> str:
        """Right-fold a list of boolean expressions into nested BINARY or() calls."""
        if len(terms) == 1:
            return terms[0]
        return "or(%s, %s)" % (terms[0], _or_chain(terms[1:]))

    _excl_terms = ["contains(toLower(item()?['name']), '%s')" % p.replace("'", "''")
                   for p in _ATTACH_EXCLUDE]
    spam["ResumeFilesKept"] = {
        "type": "Query",
        "runAfter": {"ResumeFilesAll": ["Succeeded"]},
        "inputs": {
            "from": "@body('ResumeFilesAll')",
            "where": "@not(%s)" % _or_chain(_excl_terms),
        },
    }
    _chosen = ("if(greater(length(body('ResumeFilesKept')), 0), "
               "body('ResumeFilesKept'), body('ResumeFilesAll'))")
    _resume_files_after = {"ResumeFilesKept": ["Succeeded"]}
else:
    spam.pop("ResumeFilesKept", None)
    _chosen = "body('ResumeFilesAll')"
    _resume_files_after = {"ResumeFilesAll": ["Succeeded"]}

# Pass-through Query: same array, but body('ResumeFiles') keeps its meaning everywhere.
spam["ResumeFiles"] = {
    "type": "Query",
    "runAfter": _resume_files_after,
    "inputs": {"from": "@%s" % _chosen, "where": "@not(empty(item()?['name']))"},
}

# ── 6. duplicate-block window (days) ─────────────────────────────────────────
# The window is measured from the sender's MOST RECENT application. This used to be written
# as last(Filter_submitted), which silently assumed the sheet is in P1's own append order
# (oldest -> newest). P2's resort_candidate_sheets.py (2026-07-31) re-sorts the live sheet to
# newest-FIRST, which inverted that assumption: last() started returning the sender's OLDEST
# row, so a returning applicant whose first contact was >90 days ago was read as brand new -
# new row, fresh acknowledgment, cap system bypassed - while every OTHER expression in the
# flow reads first() and was therefore correct. The two halves disagreed by construction, and
# under either ordering one of them had to be wrong.
#
# Fixed 2026-08-03 by removing the ordering assumption entirely rather than swapping one end
# for the other: project every matching row to its tick value and take the max. Correct under
# append order, newest-first, or any future re-sort, so this can never silently invert again.
#
# The union with [0] guarantees a non-empty collection: Logic Apps evaluates BOTH operands of
# an `and` (no short-circuit), so a brand-new applicant - whose Filter_submitted is empty -
# would otherwise hit max([]) and fail the run. 0 is below any real tick value, so the window
# check is simply false for them, which is what the length() operand already means.
days = int(rules["duplicate_check_days"])
hr["actions"]["Submitted_ticks"] = {
    "type": "Select",
    "runAfter": {"Filter_submitted": ["Succeeded"]},
    "inputs": {
        "from": "@body('Filter_submitted')",
        "select": "@ticks(coalesce(item()?['Received Date'], '2000-01-01T00:00:00Z'))",
    },
}
dup.setdefault("runAfter", {})["Submitted_ticks"] = ["Succeeded"]
# The 90-day window exists so a candidate who applied long ago and applies again is treated as
# a genuinely NEW application rather than an edit of a stale record. That is right when the row
# was found by ADDRESS - the window is the only evidence we have about intent.
#
# It is wrong when the row was found by an EXPLICITLY QUOTED reference. There the sender has
# named the exact application this belongs to, which is stronger evidence than any date
# heuristic, and applying the window anyway re-introduces the very bug the ref lookup fixes:
# Saniya's row is dated 2026-02-04, so a 2026-08-25 forward quoting her ref is ~200 days old,
# fails the window, and falls through to Add_row - a new row again, just found by a different
# key. So: when the lookup was BY REFERENCE, skip the window; otherwise keep it exactly as-is.
_looked_up_by_ref = {"and": [
    {"greater": ["@length(body('MatchAllowSender'))", 0]},
    {"not": {"equals": ["@outputs('QuotedRef')", ""]}},
]}
dup["expression"] = {
    "and": [
        {"greater": ["@length(body('Filter_submitted'))", 0]},
        {"or": [
            _looked_up_by_ref,
            {"greaterOrEquals": [
                "@max(union(body('Submitted_ticks'), json('[0]')))",
                # Submitted_ticks reads 'Received Date', which is stored in LOCAL time, so
                # the boundary is converted too. Left as UTC this drifts by the offset -
                # immaterial against a 90-day window, but the two sides should be one clock.
                #
                # The boundary is measured from THIS EMAIL'S OWN received date, not utcNow()
                # (changed 2026-08-28). Anchoring it to "now" silently disables duplicate
                # protection for any REPLAY of mail older than the window: the gate finds the
                # prior row, then rejects it as too old and falls through to Add_row. Live case:
                # redacted-applicant2@example.edu had two physical copies of ONE message in the mailbox (same
                # internetMessageId, received 2026-05-08). Replayed on 2026-08-28 - 112 days
                # later - each copy cleared the 90-day test independently and minted its own
                # row (APP-20260508-1642-20A1 and -E766). Against the email's own date the two
                # copies are 0 days apart, so the second is correctly seen as a duplicate.
                # For live intake received ~= now, so ordinary behaviour is unchanged.
                "@ticks(addDays(%s, -%d))" % (_RECEIVED_LOCAL_CE, days),
            ]},
        ]},
    ]
}

# ── 7. No-CV application: send the "please attach your resume" reply ONLY ─────
# An application email with NO PDF/DOCX attachment gets the reminder but is NOT written
# to SharePoint - the candidate table only ever holds rows that have a real CV. A row is
# created later, when the applicant actually sends the resume (the happy path). We drop the
# old "Awaiting CV" row + its row-based reminder cap (Get_rows_cv / Filter_awaiting /
# IsUnderCap / Add_row_awaiting); the trigger watermark still prevents reprocessing the same
# email. The keyword branch is flattened to just Send_CV_request.
_send_cv = _deep_find(kw["actions"], "Send_CV_request")
_send_cv["runAfter"] = {}
kw["actions"] = {"Send_CV_request": _send_cv}

# ── 7b. IGNORED mail: tell the admin instead of archiving it silently (2026-08-24) ───────
# This else branch is everything that cleared ALL THREE spam gates, quoted no reference,
# attached no resume and matched no application keyword. It used to be completely empty: the
# message was marked read, moved to Archive, and left no row, no reply and no signal anywhere.
# A press enquiry from a national financial daily sat unnoticed for ten weeks because of it -
# the mail was handled perfectly by design and nobody could tell it had arrived.
#
# Volume is inherently low: real spam is already gone by this point and real applications carry
# a keyword or an attachment, so what lands here is genuine human mail that simply is not an
# application. That is exactly the mail worth a human glance, so the alert is informational
# (Importance Low) rather than a failure.
#
# NON-BLOCKING, and that is load-bearing: the top-level Mark_as_read/Move_to_processed pair
# runs on `HasResume: Succeeded`, and an If only succeeds when its nested actions do. A bare
# failing send here would flip HasResume to Failed, skip the tidy, strand the message in the
# Inbox, and - because the Inbox IS the queue - have it re-processed and re-alerted every
# minute. The trailing Compose absorbs Failed/TimedOut so the branch always closes clean,
# the same catch pattern already used by Ack_stamp_done and Dup_patch_done.
if send_admin_failure_alerts:
    _notify_ignored = {
        "type": "OpenApiConnection",
        "runAfter": {},
        "runtimeConfiguration": {"retryPolicy": {"type": "exponential", "count": 2,
                                                 "interval": "PT10S"}},
        "inputs": {
            "parameters": {
                "emailMessage/To": email["admin_email"],
                "emailMessage/Subject": (
                    "[Hiring Auto-Reply] Non-application mail archived: "
                    "@{coalesce(outputs('CurrentEmail')?['subject'],'(no subject)')}"
                ),
                "emailMessage/Body": (
                    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;">'
                    '<p><strong>An email reached the hiring mailbox that is not an '
                    'application.</strong> It passed every spam screen, but carried no resume, '
                    'no application keyword and no application reference, so P1 created no row '
                    'and sent no reply.</p>'
                    '<p><strong>From:</strong> '
                    "@{coalesce(outputs('CurrentEmail')?['from'],'(unknown)')}<br>"
                    '<strong>Subject:</strong> '
                    "@{coalesce(outputs('CurrentEmail')?['subject'],'(no subject)')}<br>"
                    '<strong>Received:</strong> '
                    "@{coalesce(outputs('CurrentEmail')?['receivedDateTime'],'(unknown)')}</p>"
                    '<p><strong>Preview:</strong><br>'
                    "@{coalesce(outputs('CurrentEmail')?['bodyPreview'],'(empty)')}</p>"
                    '<p>The message has been marked read and moved to <strong>Archive</strong> '
                    '- nothing was lost. Reply from the mailbox directly if it needs a human '
                    'answer. This alert is informational, not a failure.</p></div>'
                ),
                "emailMessage/Importance": "Low",
            },
            "host": {
                "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
                "connectionName": "shared_office365",
                "operationId": "SendEmailV2",
            },
            "authentication": "@parameters('$authentication')",
        },
    }
else:
    _notify_ignored = {
        "type": "Compose",
        "runAfter": {},
        "inputs": "Admin alerts disabled: non-application mail archived without notification",
    }
kw["else"]["actions"] = {
    "Notify_ignored_mail": _notify_ignored,
    # CATCH: consume a failed alert so the enclosing If still succeeds (see the note above).
    "Ignored_alert_done": {
        "type": "Compose",
        "runAfter": {"Notify_ignored_mail": ["Succeeded", "Failed", "Skipped", "TimedOut"]},
        "inputs": "non-application mail alert best-effort",
    },
}

# ── 8. reply-detection pattern (update/reply branch) ─────────────────────────
pat = appref["detect_pattern"]
gate["expression"] = {"or": [
    {"contains": ["@outputs('LowerSubject')", pat]},
    {"contains": ["@outputs('LowerBody')", pat]},
]}

# ── 8b. reply-first ordering: evaluate the quoted-reference reply BEFORE the new-application
#       branch. A reply quoting APP-... (its True branch ends in Terminate) is then handled
#       ONCE - it can no longer also be added as a brand-new candidate. For normal emails the
#       gate's else is empty (a no-op) so HasResume runs exactly as before.
gate["runAfter"] = {"ResumeNames": ["Succeeded"]}
hr["runAfter"] = {"HasAppRefGate": ["Succeeded", "Skipped"]}

# FileRef_from_subject: filename-safe form of the MATCHED ROW's real Application ID (fixed
# 2026-07-31: was built from the text-quoted AppRef_from_subject - see AppRef_current's comment
# in section 14b for why that let an updated resume get saved under a stale/dead reference).
# (use _deep_find because after the first build these live inside IsKnownSender)
_frs = _deep_find(gate["actions"], "FileRef_from_subject")
if _frs is None:
    _frs = {"type": "Compose", "inputs": ""}
_frs["runAfter"] = {"AppRef_current": ["Succeeded"]}
_frs["inputs"] = "@{replace(outputs('AppRef_current'), '/', '-')}"
_hru = _deep_find(gate["actions"], "Has_resume_in_update")
if _hru is not None:
    _hru["runAfter"] = {"FileRef_from_subject": ["Succeeded"]}

# ── 9. application-keyword OR (regenerated from app_keywords) ─────────────────
# Gated on NO attachment so it can't double-fire with HasWrongFormat (its sibling). E3
# ("please attach your CV") means "you forgot the file"; E4 ("resend as PDF/Word") means
# "we got a file but can't read it". Without this guard, an email like subject "my resume"
# + a photo.jpg would match BOTH and send two contradictory replies.
kws = list(filt["app_keywords"])
kw["expression"] = {"and": [
    {"equals": ["@length(coalesce(triggerOutputs()?['body/attachments'], json('[]')))", 0]},
    {"or": [
        {"contains": ["@outputs('%s')" % field, k]}
        for field in ("LowerSubject", "LowerBody")
        for k in kws
    ]},
]}

# ── 10. Excel row actions (all 4) + admin failure recipient ──────────────────
# Normalized sender address: lowercased, and stripped of a 'Display Name <addr>' wrapper if
# present (same '<'-split shape the Full Name guess already relies on, section 12/13 below -
# this just takes the OTHER side of the split and drops the trailing '>'). Applied BOTH to the
# $filter's comparison value and to the stored 'Email' cell on every new row, so sender matching
# is symmetric with Gate 1's own LowerFrom check and immune to display-name/casing differences
# between two emails from the same person - going forward only; rows written before this change
# keep whatever raw value they already have (2026-07-12 audit finding 2).
_email_norm_expr = (
    "toLower(trim(if("
    "contains(triggerOutputs()?['body/from'],'<'),"
    "replace(last(split(triggerOutputs()?['body/from'],'<')),'>',''),"
    "triggerOutputs()?['body/from']"
    ")))"
)
# OData string values must escape single quotes by DOUBLING them, or a sender address like
# o'brien@x.com breaks the Excel `$filter`. Used by Get_rows and Get_rows_ref.
_EMAIL_FILTER = "@concat('Email eq ''', replace(%s, '''', ''''''), '''')" % _email_norm_expr
set_excel(hr["actions"]["Get_rows"])
# Get_rows feeds Filter_submitted -> Is_duplicate, i.e. the NEW-vs-DUPLICATE decision. Keying
# it purely on the sender's address means an internal forward is never recognised as belonging
# to an existing candidate: the row is looked up under the FORWARDER's address, nothing matches,
# and Add_row mints a brand-new row in the forwarder's name with the real candidate's resume
# attached to it. Live case 2026-08-25: yashv@driverai.io forwarded Saniya Sekhon's updated
# resume quoting APP-20260204-1433-3B16 in the subject; P1 created APP-20260825-1505-F08E under
# yashv@driverai.io, sent HIM the acknowledgment, and left Saniya's real row untouched at
# Application Updates 0.
#
# Same conditional shape as Get_rows_ref below, and deliberately the SAME safety restriction:
# by-reference lookup only for ALLOW-LISTED senders. Without that, any external sender could
# quote someone else's Application ID and overwrite that candidate's row and resume file.
_REF_OR_EMAIL_FILTER = (
    "@if(and(greater(length(body('MatchAllowSender')), 0), "
    "not(equals(outputs('QuotedRef'), ''))), "
    "concat('Application ID eq ''', outputs('QuotedRef'), ''''), "
    "%s)" % _EMAIL_FILTER[1:]
)
hr["actions"]["Get_rows"]["inputs"]["parameters"]["$filter"] = _REF_OR_EMAIL_FILTER
set_excel(dup["else"]["actions"]["Add_row"])

# ── 10a. Workbook is NON-DESTRUCTIVE now — NO auto-create/overwrite (changed 2026-07-04) ──
# The old design probed the workbook with GetFileMetadataByPath and, on a "NotFound", uploaded a
# blank template over it. That probe proved UNRELIABLE (it kept returning NotFound for a workbook
# that genuinely existed), so self-heal fired on nearly every run and RE-UPLOADED A BLANK TEMPLATE
# OVER THE REAL DATA - silently wiping candidate rows down to just the header. Root cause: a false
# NotFound from the metadata probe triggered a destructive overwrite.
#
# New design (user decision): never auto-create or overwrite the workbook. Get_rows just runs
# directly. The Excel connector's own source/drive/file/table addressing finds the real workbook
# regardless (that addressing has always worked - rows read/write fine). If the workbook is GENUINELY
# missing (first-ever setup, or accidental deletion), Get_rows fails, HasResume fails, and the
# existing Notify_failure alerts the admin (yashv@driverai.io) so a human can set it up once from the
# committed P1_Templates/HiringAgent_P1_CandidateList.xlsx - it is NOT silently replaced. Zero
# overwrite risk, no unreliable probe, no self-heal delays.
#
# The dated RESUME folders are still auto-created (Ensure_*_resume_folder) - that's safe: creating a
# folder that already exists is a harmless no-op, no data can be lost.

# These path helpers are still needed by the dated resume-folder creation (Ensure_*_resume_folder
# via resume_folder_path_lib_relative). CreateNewFolder takes a library-relative path (no
# "/Shared Documents" prefix - that's implied by the 'table' library GUID selector).
_resumes_rel   = sp["resumes_folder"]          # /Shared Documents/Candidate_Resumes
_sp_lib_prefix = "/Shared Documents/"
_sf_folder_lib_relative = (
    _resumes_rel[len(_sp_lib_prefix):] if _resumes_rel.startswith(_sp_lib_prefix)
    else _resumes_rel.lstrip("/")
)

# Strip any leftover workbook self-heal actions the base zip still carries forward (same
# base-zip-reuse concern as the old Keywords / top-level HasValidResumeEarly cleanup). None of
# these exist in the new non-destructive design, so pop them so they can't linger as dead actions.
for _dead_wb in ("EnsureWb_Probe", "WorkbookExists_NoOp", "WorkbookTemplate_B64",
                 "Create_workbook_folder", "Create_workbook",
                 "Wait_after_workbook_folder", "Wait_after_workbook_heal"):
    hr["actions"].pop(_dead_wb, None)

# Get_rows is now the first action in the HasResume-true branch (no probe/create ahead of it).
hr["actions"]["Get_rows"]["runAfter"] = {}
set_email_mailbox(wd)                                   # all SharedMailbox sends
spam["Notify_failure"]["inputs"]["parameters"]["emailMessage/To"] = email["admin_email"]

# ── 10b/10c. Year/month separator rows ──────────────────────────────────────
# Excel Online's List rows action only supports alphanumeric column names in OData
# Filter Query. "Received Date ge ..." therefore fails at runtime with a syntax error
# at the space (confirmed live 2026-07-31). Do one paginated, unfiltered row scan and
# use built-in Query actions to filter the returned ISO timestamps in memory. This is
# one Excel request path shared by both checks, supports the visible header unchanged,
# and avoids every OData reference to the spaced column name.
_year_sep_enabled = year_sep.get("enabled", True)
_month_sep_enabled = month_sep.get("enabled", True)

_sep_actions = dup["else"]["actions"]
for _old_separator_action in (
    "Get_rows_any_applicant",
    "Get_rows_this_year",
    "Get_rows_this_month",
    "Get_rows_for_separators",
    "Filter_rows_this_year",
    "Filter_rows_this_month",
    "IsFirstOfNewYear",
    "IsFirstOfNewMonth",
):
    _sep_actions.pop(_old_separator_action, None)
for _old_separator_dep in ("IsFirstOfNewYear", "IsFirstOfNewMonth"):
    _sep_actions["Add_row"]["runAfter"].pop(_old_separator_dep, None)

# The workbook ships with one permanent blank table row directly under each header
# because Excel/Graph reject header-only tables. A separator is only needed after at
# least one real applicant exists; otherwise the first applicant would get two stacked
# blanks (the seed row plus a new period separator).
if _year_sep_enabled or _month_sep_enabled:
    _get_rows_any_applicant = {
        "type": "OpenApiConnection",
        "runAfter": {},
        "inputs": {
            "parameters": {"source": "", "drive": "", "file": "", "table": "",
                           "$filter": "Email ne ''", "$top": 1},
            "host": {
                "apiId": "/providers/Microsoft.PowerApps/apis/shared_excelonlinebusiness",
                "connectionName": "shared_excelonlinebusiness",
                "operationId": "GetItems",
            },
            "authentication": "@parameters('$authentication')",
        },
    }
    set_excel(_get_rows_any_applicant)
    _sep_actions["Get_rows_any_applicant"] = _get_rows_any_applicant

    _get_rows_for_separators = {
        "type": "OpenApiConnection",
        "runAfter": {},
        "runtimeConfiguration": {
            "paginationPolicy": {"minimumItemCount": 5000}
        },
        "inputs": {
            "parameters": {
                "source": "", "drive": "", "file": "", "table": "",
                "$top": 5000,
                "dateTimeFormat": "ISO 8601",
            },
            "host": {
                "apiId": "/providers/Microsoft.PowerApps/apis/shared_excelonlinebusiness",
                "connectionName": "shared_excelonlinebusiness",
                "operationId": "GetItems",
            },
            "authentication": "@parameters('$authentication')",
        },
    }
    set_excel(_get_rows_for_separators)
    _sep_actions["Get_rows_for_separators"] = _get_rows_for_separators

if _year_sep_enabled:
    _sep_actions["Filter_rows_this_year"] = {
        "type": "Query",
        "runAfter": {"Get_rows_for_separators": ["Succeeded"]},
        "inputs": {
            "from": "@coalesce(body('Get_rows_for_separators')?['value'], json('[]'))",
            "where": (
                "@startsWith(string(item()?['Received Date']), "
                "concat(formatDateTime(%s,'yyyy'), '-'))" % _RECEIVED_LOCAL_CE
            ),
        },
    }

    _add_row_separator = {
        "type": "OpenApiConnection",
        "runAfter": {},
        "inputs": {
            "parameters": {"source": "", "drive": "", "file": "", "table": "", "item": {}},
            "host": {
                "apiId": "/providers/Microsoft.PowerApps/apis/shared_excelonlinebusiness",
                "connectionName": "shared_excelonlinebusiness",
                "operationId": "AddRowV2",
            },
            "authentication": "@parameters('$authentication')",
        },
    }
    set_excel(_add_row_separator)

    _sep_actions["IsFirstOfNewYear"] = {
        "type": "If",
        "runAfter": {"Filter_rows_this_year": ["Succeeded"],
                     "Get_rows_any_applicant": ["Succeeded"]},
        "expression": {"and": [
            {"equals": ["@length(body('Filter_rows_this_year'))", 0]},
            {"greater": ["@length(body('Get_rows_any_applicant')?['value'])", 0]},
        ]},
        "actions": {"Add_row_year_separator": _add_row_separator},
        "else": {"actions": {}},
    }

    # Add_row (the real candidate row) always waits for the separator check to finish -
    # regardless of outcome - so a just-inserted blank row is guaranteed to land BEFORE
    # the real row, never after or concurrently with it.
    _sep_actions["Add_row"]["runAfter"] = {
        **_sep_actions["Add_row"]["runAfter"],
        "IsFirstOfNewYear": ["Succeeded", "Failed", "Skipped"],
    }

# At a January boundary the year separator already fires. Requiring at least one row
# in the incoming year makes the year separator win, so exactly one blank row lands.
if _month_sep_enabled:
    _sep_actions["Filter_rows_this_month"] = {
        "type": "Query",
        "runAfter": {"Get_rows_for_separators": ["Succeeded"]},
        "inputs": {
            "from": "@coalesce(body('Get_rows_for_separators')?['value'], json('[]'))",
            "where": (
                "@startsWith(string(item()?['Received Date']), "
                "formatDateTime(%s,'yyyy-MM'))" % _RECEIVED_LOCAL_CE
            ),
        },
    }

    _add_row_month_separator = {
        "type": "OpenApiConnection",
        "runAfter": {},
        "inputs": {
            "parameters": {"source": "", "drive": "", "file": "", "table": "", "item": {}},
            "host": {
                "apiId": "/providers/Microsoft.PowerApps/apis/shared_excelonlinebusiness",
                "connectionName": "shared_excelonlinebusiness",
                "operationId": "AddRowV2",
            },
            "authentication": "@parameters('$authentication')",
        },
    }
    set_excel(_add_row_month_separator)

    _none_this_month = {"equals": ["@length(body('Filter_rows_this_month'))", 0]}
    _has_any_applicant = {
        "greater": ["@length(body('Get_rows_any_applicant')?['value'])", 0]}
    _month_expr = ({"and": [_none_this_month, _has_any_applicant,
                            {"greater": ["@length(body('Filter_rows_this_year'))", 0]}]}
                   if _year_sep_enabled else
                   {"and": [_none_this_month, _has_any_applicant]})

    _month_run_after = {
        "Filter_rows_this_month": ["Succeeded"],
        "Get_rows_any_applicant": ["Succeeded"],
    }
    if _year_sep_enabled:
        _month_run_after["Filter_rows_this_year"] = ["Succeeded"]

    _sep_actions["IsFirstOfNewMonth"] = {
        "type": "If",
        "runAfter": _month_run_after,
        "expression": _month_expr,
        "actions": {"Add_row_month_separator": _add_row_month_separator},
        "else": {"actions": {}},
    }

    _sep_actions["Add_row"]["runAfter"] = {
        **_sep_actions["Add_row"]["runAfter"],
        "IsFirstOfNewMonth": ["Succeeded", "Failed", "Skipped"],
    }

# ── 11. resume save location (flat, or dated Year/Month subfolders) ──────────
# CreateFile does NOT auto-create missing folders (fails with NotFound) - so the dated folder must
# be explicitly ensured before every Create_file call, not just assumed to exist.
_folder = resume_folder_path()
dup["else"]["actions"]["Ensure_resume_folder"] = _make_ensure_resume_folder(_RECEIVED_LOCAL)
# Strip any stale folder-wait actions the base zip still carries (all delays removed 2026-07-04).
dup["else"]["actions"].pop("Wait_after_resume_folder", None)
# No wait between folder-create and file-save: CreateNewFolder returns after the folder exists, and
# the immediate CreateFile into it works (confirmed live 2026-07-04 - resumes saved fine even when
# the old 2s wait was itself failing). Removed per user request; PA's Wait minimum is 5s anyway.
dup["else"]["actions"]["Save_resumes_to_SharePoint"]["runAfter"] = {
    "Ensure_resume_folder": ["Succeeded", "Failed"]}
(dup["else"]["actions"]["Save_resumes_to_SharePoint"]["actions"]["Create_file"]
 ["inputs"]["parameters"]["folderPath"]) = _folder
# An updated CV is filed into the ORIGINAL application's dated folder (rebuilt from the row's
# 'Received Date'), NOT the update email's date — so Phase 2, which reconstructs the path from
# that same Received Date, downloads the new file. Same original filename => the old CV is
# replaced (the only intentional overwrite: a genuine resume update).
_deep_find(gate["actions"], "Create_update_file")[
    "inputs"]["parameters"]["folderPath"] = resume_folder_path(
        "first(body('Get_rows_ref')?['value'])?['Received Date']")

# ── 11b. literal 'dataset' on every SharePoint action ─────────────────────────
# The legacy base left dataset = @outputs('CONFIG')[...] on the two Create-file
# actions. Runtime resolves that fine, but Create-as-new import validation resolves
# connector metadata from 'dataset', and an expression there is the same class of
# avoidable risk that broke the Excel source/drive/file/table bindings. Overwrite
# ALL SP datasets with the literal site URL (idempotent, matches hand-built flows).
def _set_sp_dataset(node) -> None:
    if isinstance(node, dict):
        if node.get("type") == "OpenApiConnection":
            host = node.get("inputs", {}).get("host", {})
            if str(host.get("apiId", "")).endswith("shared_sharepointonline"):
                node["inputs"]["parameters"]["dataset"] = sp["site"]
        for v in node.values():
            _set_sp_dataset(v)
    elif isinstance(node, list):
        for v in node:
            _set_sp_dataset(v)


_set_sp_dataset(wd)

# ── 12. filename prefix: <FirstLast>_<AppID>.<ext> (readable in the SharePoint folder) ──
# Changed 2026-07-12 (user request): was <AppID>_<original-filename>. This '<FirstLast>_
# <FullAppID>' shape is what P1 ALWAYS saves under - arrival, Duplicate resend, and
# ref-quoted Update all target this exact same name (sections 12/13a/14a), never anything
# else. It's deliberately readable immediately (P2's lookup, _stored_resume_names, tries it
# as a candidate on its first pass) but it is NOT the final shape: once P2 scores a row, it
# renames the file OUT of this legacy shape into '<FirstLast>_<Category>_<tail>' (the
# 2026-07-15 filename convention, sharepoint_scoring.py) - Category is P2-only knowledge, so
# P1 can never compute or target that shape itself. P2's _download_resume_text() is what
# reconciles a fresh legacy-shape write (from a later Update) against a stale, already-
# renamed canonical-shape file - see the 2026-07-15 correction note on the Update path below.
# _name_part_expr() mirrors P2's _get_cleaned_filename_prefix() as closely as PA's expression
# language allows: strip apostrophe/hyphen/period/comma, split on space, first+last word (or
# the only word), 'Candidate' if empty - MUST stay in sync with the Python function if either
# changes.
def _name_part_expr(raw_name_expr: str) -> str:
    cleaned = ("trim(replace(replace(replace(replace(%s,'''',''),'-',''),'.',''),',',''))"
               % raw_name_expr)
    words = "split(%s,' ')" % cleaned
    return ("if(equals(length(%s),0),'Candidate',"
            "if(equals(length(%s),1),first(%s),concat(first(%s),last(%s))))"
            % (cleaned, words, words, words, words))


# Raw (unwrapped, no leading @) From-header name guess - shared with the row's 'Full Name'
# column below (section 13) so the two stay byte-identical without duplicating the expression.
_name_expr_body = (
    "trim(if("
    "and(contains(triggerOutputs()?['body/from'],'<'),"
    "greater(length(trim(first(split(triggerOutputs()?['body/from'],'<')))),0)),"
    "first(split(triggerOutputs()?['body/from'],'<')),"
    "replace(replace(first(split(coalesce(triggerOutputs()?['body/from'],''),'@')),'.',' '),'_',' ')"
    "))"
)

def _attach_slug(item_expr: str) -> str:
    """A short, deterministic tag distinguishing two attachments saved for the same row.

    ONLY appended when the email really carried more than one kept attachment - the single
    attachment case (essentially all of them, especially now that portfolios are excluded)
    keeps the exact '<FirstLast>_<AppRef>.<ext>' shape P2's _resume_name_slots() reconstructs,
    so nothing downstream changes for it.

    Definition, which hiring_agent/sharepoint_scoring.py::_attachment_slug MIRRORS EXACTLY -
    change one and you must change the other: lowercase the WHOLE original filename, delete
    space _ - ' , ( ) and . , then take the first 16 characters. The dot is deleted rather
    than split on, so 'my.name.resume.pdf' needs no multi-dot handling on either side.

    Two attachments with the IDENTICAL filename still collide; Outlook allows that and there
    is no loop index in Power Automate to break the tie. It is rare enough to accept, and it
    is strictly better than today, where ANY two same-extension files collide.
    """
    clean = "toLower(%s)" % item_expr
    for ch in (" ", "_", "-", "''", ",", "(", ")", "."):
        clean = "replace(%s,'%s','')" % (clean, ch)
    return "substring(%s,0,min(16,length(%s)))" % (clean, clean)


def _multi_suffix(loop_name: str) -> str:
    """'' for a single attachment, '_<slug>' when this email kept more than one."""
    item_expr = "items('%s')?['name']" % loop_name
    return ("@{if(greater(length(body('ResumeFiles')),1),concat('_',%s),'')}"
            % _attach_slug(item_expr))


# New applicant: NamePart derived from THIS email's From header (becomes the row's Full Name too).
(dup["else"]["actions"]["Save_resumes_to_SharePoint"]["actions"]["Create_file"]
 ["inputs"]["parameters"]["name"]) = (
    "@{%s}_@{outputs('FileRef')}%s.@{last(split(items('Save_resumes_to_SharePoint')?['name'],'.'))}"
    % (_name_part_expr(_name_expr_body), _multi_suffix("Save_resumes_to_SharePoint"))
)
# Update (ref-reply): OVERWRITE fix (2026-07-04) - save under the ORIGINAL row's identity, not
# the update attachment's name. NamePart is recomputed from the ORIGINAL row's stored 'Full
# Name' (Get_rows_ref), not this reply's From header, so this always targets the SAME
# '<FirstLast>_<FullAppID>.<ext>' legacy shape Create_file (section 12) used on arrival -
# P1 never varies this, on first submission, a Duplicate resend, OR a ref-quoted Update.
#
# CORRECTION (2026-07-15): the "byte-identical path -> CreateFile overwrites, one file per
# candidate" claim this comment used to make is only true BEFORE the row has ever been
# scored. Once P2 scores it, P2 renames the file OUT of this legacy shape into
# '<FirstLast>_<Category>_<tail>.<ext>' (sharepoint_scoring.py, 2026-07-15 filename
# convention) - a shape this flow has no visibility into (Category is P2-only knowledge).
# So an Update arriving AFTER the candidate has already been scored once does NOT overwrite
# anything here: it creates a genuinely NEW file under this legacy name, alongside the
# now-stale renamed one. That's by design, not a bug to fix on this side - P2's
# _download_resume_text() is what reconciles it: it always prefers a legacy-shape file over
# a canonical-shape one for the same Application ID (a legacy-shape file existing on an
# already-scored row can ONLY mean "P1 just wrote a fresh update"), scores off that alone
# instead of blending it with the stale content, and deletes the stale sibling once the
# rename lands. See HiringAgent_App_P2/hiring_agent/sharepoint_scoring.py
# (_resume_name_slots / _download_resume_text) for the actual reconciliation logic.
_deep_find(gate["actions"], "Create_update_file")[
    "inputs"]["parameters"]["name"] = (
    "@{%s}_@{outputs('FileRef_from_subject')}%s.@{last(split("
    "first(body('Get_rows_ref')?['value'])?['Original Filename'],'.'))}"
    % (_name_part_expr("coalesce(first(body('Get_rows_ref')?['value'])?['Full Name'],'')"),
       _multi_suffix("Save_update_resume"))
)

# ── 13. row fields: Subject, Email Body, improved Full Name ──────────────
# Better name: if From has a display name before '<' use it; else derive a
# rough name from the email local part (replace . and _ with spaces).
# P2's resolve_full_name overwrites with the resume-parsed name later.
# (raw body shared with section 12's NamePart via _name_expr_body, defined above)
_name_expr = "@{%s}" % _name_expr_body

# IMPORTANT (2026-07-04): Add_row writes ONLY the 10 P1-owned columns - it does NOT list the 12
# P2-owned columns at all. This exactly matches the genuine tenant PA exports (archive/ne3, ne4),
# which are the ground-truth format PA itself produces and accepts. The old build listed all 22
# (the 10 real ones + 12 P2 columns with empty-string ""). That deviation mattered on IMPORT: the
# Excel connector binds each item key against the live table's dynamic schema fetched at import
# time, and listing 22 keys forced PA to resolve 22 column-name matches (any one mismatch/incomplete
# fetch can drop the per-column binding, which is what shows as EMPTY fields in the designer). The
# 12 P2 columns already exist in the table (from the 22-col template) and P2 fills them later, so
# omitting them from the write leaves them as blank cells - identical outcome, far less binding
# surface, and byte-for-byte the format proven to work in the tenant.
_ROW_ITEM = {
    "Application ID": "@outputs('AppRef')",
    # ISO 8601 without the timezone offset — drops the confusing '+00:00' and the sub-minute
    # noise while staying a valid ISO timestamp so the 90-day dup check's ticks() and P2's
    # date parser both still read it. e.g. '2026-06-24T22:38:50'.
    "Received Date": "@{formatDateTime(%s,'yyyy-MM-ddTHH:mm:ss')}" % _RECEIVED_LOCAL,
    # Latest resume-bearing email timestamp. On first intake it equals Received Date; update
    # and duplicate-resend paths patch this while leaving Received Date fixed for folder paths.
    "Last Updated Date": "@{formatDateTime(%s,'yyyy-MM-ddTHH:mm:ss')}" % _RECEIVED_LOCAL,
    "Full Name": _name_expr,
    "Email": "@%s" % _email_norm_expr,
    "Mail Subject": "@coalesce(triggerOutputs()?['body/subject'],'')",
    # bodyPreview = O365 plain-text summary (~255 chars). P2 uses this as a
    # fallback extraction source; the resume is primary. Newlines → ' | '.
    "Mail Body": ("@replace(replace(coalesce(triggerOutputs()?['body/bodyPreview'],''),decodeUriComponent('%0A'),' | '),decodeUriComponent('%0D'),'')"),
    "Status": "New Email Received",
    "Has Resume": "Yes",
    "Original Filename": "@join(body('ResumeNames'), ', ')",
    "Application Updates": 0,
}

# Happy-path Add_row (the ONLY row writer: every candidate row has a real CV)
_item = dup["else"]["actions"]["Add_row"]["inputs"]["parameters"]["item"]
_item.clear()
_item.update(_ROW_ITEM)
# (No "Awaiting CV" row: a no-CV application is replied to but never written to SharePoint.)

# ── 13b. "Mail Sent" audit column ─────────────────────────────────────────────
# Stamped on the row only AFTER a reply actually went out (Add_row runs BEFORE
# Send_acknowledgment, so stamping at row creation would lie). The 'Mail Sent'
# column must exist in the live table BEFORE this package is imported (the Excel
# connector binds every item key against the table schema at import time); P2's
# ensure_columns creates it on the live workbook.
# Stamp carries a date/time (not just 'Yes') so it reads consistently with the
# Rejected sheet's 'Decline Sent' timestamp ("Sent 2026-07-07 09:15").
_MAIL_SENT_STAMP = "@{concat('Sent ', formatDateTime(%s,'yyyy-MM-dd HH:mm'))}" % _NOW_LOCAL

# Ack path: stamp right after the send succeeds (Send_acknowledgment itself runs
# between Add_row and this patch, which is settle time enough - and PA Consumption
# rejects Wait intervals under 5s, so no explicit Wait here; the flow is deliberately
# zero-Wait). The trailing Compose consumes any failure so the run (and the
# read+archive tidy pair, which only runs when HasResume SUCCEEDS) is never lost
# over a missing stamp - the blank cell itself flags the row for a look.
_ack_acts = dup["else"]["actions"]
# A 2s Wait shipped in one earlier build lives on in the zip base (the structural
# template) - pop it every build so it can never resurface (same pattern as
# Wait_after_resume_folder above).
_ack_acts.pop("Wait_before_mail_stamp", None)

# Never acknowledge an application that was not actually recorded. The zip base shipped
# Send_acknowledgment with runAfter Add_row [Succeeded, Failed, Skipped], so an Excel write
# failure (locked workbook, schema mismatch, throttling) still sent the applicant a mail
# quoting an Application ID that has no row - telling a candidate they are in a system they
# are not in, and leaving nothing for P2 to score. Notify_failure does alert the admin, but
# by then the mail has gone. Pinned here (not left to the base template) so it is re-asserted
# on every build. Add_row Skipped is deliberately excluded too: it is skipped only when an
# upstream save timed out, which is equally not a recorded application.
_deep_find(wd["actions"], "Send_acknowledgment")["runAfter"] = {"Add_row": ["Succeeded"]}
_patch_mail_sent = {
    "type": "OpenApiConnection",
    "runAfter": {"Send_acknowledgment": ["Succeeded"]},
    "inputs": {
        "parameters": {
            "source": "", "drive": "", "file": "", "table": "",
            "idColumn": "Application ID",
            "id": "@{outputs('AppRef')}",
            "item": {"Mail Sent": _MAIL_SENT_STAMP},
        },
        "host": {
            "apiId": "/providers/Microsoft.PowerApps/apis/shared_excelonlinebusiness",
            "connectionName": "shared_excelonlinebusiness",
            "operationId": "PatchItem",
        },
        "authentication": "@parameters('$authentication')",
    },
}
set_excel(_patch_mail_sent)
_ack_acts["Patch_mail_sent"] = _patch_mail_sent
_ack_acts["Ack_stamp_done"] = {
    "type": "Compose",
    "runAfter": {"Patch_mail_sent": ["Succeeded", "Failed", "Skipped"]},
    "inputs": "ack mail-sent stamp best-effort",
}

# ── 14. same-row capping (CV Attempts counter on the candidate's own row) ──
# Each scenario applies its OWN threshold from flow_config.business_rules. NOTE: all four
# scenarios share ONE counter column ("CV Attempts") - truly independent counters would need
# extra Excel columns (e.g. "Update Attempts"), which we don't add to the external table here.
dup_cap = int(rules["duplicate_notice_max"])
update_cap = int(rules["update_resume_max"])
followup_cap = int(rules["followup_reply_max"])
# reply_cap is a SEPARATE, smaller threshold than the three *_max caps above: those still
# govern the overall silence point (CV Attempts keeps counting up to *_max, then the row goes
# fully silent forever within the 90-day window). reply_cap decides whether a given contact
# actually gets an EMAIL at all - contacts under reply_cap get a reply (the last one, at
# reply_cap - 1, is the final-notice); contacts at/after reply_cap but still under *_max still
# update the row/resume (re-queued for P2) but get NO email - "no other reminders".
reply_cap = int(rules["reply_cap"])


def _last_updated_expr(rows: str) -> str:
    """max(row's existing Last Updated Date, this email's own receivedDateTime) - never move
    the column BACKWARD. Backlog replay can process an older queued email after a newer one
    already patched the same row (e.g. two stale unread replies from the same sender, whichever
    order the trigger's Inbox happens to drain them in) - a blind overwrite would then stamp an
    earlier timestamp than the row already has. That corrupts P2's _updated_after_location_request()
    (sharepoint_scoring.py), which trusts this column to mean "most recent candidate contact".
    Falls back to Received Date (then a fixed epoch) when Last Updated Date is blank, matching the
    existing coalesce shape already proven safe by the 90-day dup-window ticks() check above."""
    return (
        # Both sides of the comparison are LOCAL: the stored column was written local by
        # _ROW_ITEM above, so the incoming timestamp is converted before it is compared -
        # otherwise every fresh email would look 7 hours newer than it is.
        "@{if(greater(ticks(%s),"
        "ticks(coalesce(first(%s)?['Last Updated Date'], first(%s)?['Received Date'], '2000-01-01T00:00:00'))),"
        "formatDateTime(%s,'yyyy-MM-ddTHH:mm:ss'),"
        "coalesce(first(%s)?['Last Updated Date'], first(%s)?['Received Date']))}"
        % (_RECEIVED_LOCAL, rows, rows, _RECEIVED_LOCAL, rows, rows)
    )


def _make_patch_item(source_action: str, run_after: dict, excel: bool = False,
                     stamp_last_updated: bool = False) -> dict:
    """PatchItem action: increment CV Attempts on an existing row.

    excel=True when source_action is an Excel GetItems whose body() is the object
    {"value": [...rows]} (needs ?['value']); excel=False for a Query/Filter-array body
    that is already a row array.
    """
    rows = ("body('%s')?['value']" if excel else "body('%s')") % source_action
    item = {
        "Application Updates": "@add(int(coalesce(first(%s)?['Application Updates'], '0')), 1)" % rows,
        # This patch also runs on Failed/Skipped sends (the attempt counter
        # must always move), so only stamp 'Mail Sent' when the reply really
        # went out - and otherwise KEEP the row's existing stamp instead of
        # blanking it. run_after's single key is that send action.
        "Mail Sent": (
            "@{if(equals(actions('%s')?['status'],'Succeeded'),"
            "concat('Sent ', formatDateTime(%s,'yyyy-MM-dd HH:mm')),"
            "coalesce(first(%s)?['Mail Sent'],''))}" % (next(iter(run_after)), _NOW_LOCAL, rows)),
    }
    if stamp_last_updated:
        item["Last Updated Date"] = _last_updated_expr(rows)
    return {
        "runAfter": run_after,
        "type": "OpenApiConnection",
        "inputs": {
            "parameters": {
                "source": "", "drive": "", "file": "", "table": "",
                "idColumn": "Application ID",
                "id": "@{first(%s)?['Application ID']}" % rows,
                "item": item,
            },
            "host": {
                "apiId": "/providers/Microsoft.PowerApps/apis/shared_excelonlinebusiness",
                "connectionName": "shared_excelonlinebusiness",
                "operationId": "PatchItem",
            },
            "authentication": "@parameters('$authentication')",
        },
    }


def _make_patch_item_noreply(source_action: str, run_after: dict, excel: bool = False,
                              requeue: bool = False, stamp_last_updated: bool = False) -> dict:
    """PatchItem action for a contact PAST reply_cap but still under the overall cap: increments
    CV Attempts same as _make_patch_item, but 'Mail Sent' is simply OMITTED (no email went out,
    so the column is left completely untouched - PatchItem only writes the keys listed here).
    requeue=True also re-queues Status (duplicate/update paths still update the resume/row even
    though no reply is sent; follow-up has no new resume, so it never requeues)."""
    rows = ("body('%s')?['value']" if excel else "body('%s')") % source_action
    item = {"Application Updates": "@add(int(coalesce(first(%s)?['Application Updates'], '0')), 1)" % rows}
    if stamp_last_updated:
        item["Last Updated Date"] = _last_updated_expr(rows)
    if requeue:
        item["Status"] = "New Email Received"
    return {
        "runAfter": run_after,
        "type": "OpenApiConnection",
        "inputs": {
            "parameters": {
                "source": "", "drive": "", "file": "", "table": "",
                "idColumn": "Application ID",
                "id": "@{first(%s)?['Application ID']}" % rows,
                "item": item,
            },
            "host": {
                "apiId": "/providers/Microsoft.PowerApps/apis/shared_excelonlinebusiness",
                "connectionName": "shared_excelonlinebusiness",
                "operationId": "PatchItem",
            },
            "authentication": "@parameters('$authentication')",
        },
    }


def _cv_check_expr(source_action: str, cap_value: int, excel: bool = False) -> dict:
    """Expression: CV Attempts on the first matched row < the given cap.

    excel=True unwraps an Excel GetItems body (object {"value": [...]}); excel=False
    treats the body as a Query/Filter-array (already a row array).
    """
    rows = ("body('%s')?['value']" if excel else "body('%s')") % source_action
    return {"less": [
        "@int(coalesce(first(%s)?['Application Updates'], '0'))" % rows,
        cap_value,
    ]}


def _make_get_items(run_after: dict) -> dict:
    """Skeleton GetItems action filtered by sender email."""
    return {
        "type": "OpenApiConnection",
        "runAfter": run_after,
        "inputs": {
            "parameters": {
                "source": "", "drive": "", "file": "", "table": "",
                "$filter": _EMAIL_FILTER,
            },
            "host": {
                "apiId": "/providers/Microsoft.PowerApps/apis/shared_excelonlinebusiness",
                "connectionName": "shared_excelonlinebusiness",
                "operationId": "GetItems",
            },
            "authentication": "@parameters('$authentication')",
        },
    }


# ── 14a. Duplicate cap (same-row: SAVE the resend as a real update, then send + PATCH or silent)
# A resend within 90 days from a known sender is still a genuine new resume even when the reply
# never quotes "APP-..." (most people don't retype a reference into a fresh email). Previously this
# branch only sent the "already on file" notice and incremented CV Attempts - the new attachment was
# never saved and Status/Resume Filename were never re-queued, so an applicant who didn't know the
# "Update - <ref>" convention could resend their updated resume indefinitely and it would never
# actually reach Phase 2. Now it saves the file into the ORIGINAL row's dated folder and re-queues
# Status/Resume Filename exactly like the ref-quoting update path (Patch_update_attempts) does.
_send_dup = _deep_find(dup["actions"], "Send_duplicate_notice")
_send_dup["runAfter"] = {}   # now the first action inside the nested IsUnderReplyCap_dup branch

_fileref_dup = {
    "type": "Compose",
    "runAfter": {},
    "inputs": "@{replace(first(body('Filter_submitted'))?['Application ID'], '/', '-')}",
}

_ensure_dup_folder = _make_ensure_resume_folder(
    "first(body('Filter_submitted'))?['Received Date']")

# OVERWRITE fix (2026-07-04): save the resend under the ORIGINAL row's identity, not the new
# attachment's name - same legacy '<FirstLast>_<FullAppID>.<ext>' shape as section 12/the
# ref-quoted Update path above (see the 2026-07-15 correction note there: this is a genuine
# overwrite only pre-scoring; post-scoring it creates a fresh legacy-shape file that P2's
# _download_resume_text() reconciles against the stale, already-renamed canonical-shape one).
# The 'Original Filename'/'Full Name' columns are left UNCHANGED, so P2 still finds it.
_dup_orig_ext = "@{last(split(first(body('Filter_submitted'))?['Original Filename'],'.'))}"
_dup_name_part = _name_part_expr("coalesce(first(body('Filter_submitted'))?['Full Name'],'')")
_save_dup_resume = {
    "type": "Foreach",
    "runAfter": {"Ensure_dup_resume_folder": ["Succeeded", "Failed"]},
    "foreach": "@body('ResumeFiles')",
    "actions": {
        "Create_dup_update_file": {
            "type": "OpenApiConnection",
            "runAfter": {},
            "inputs": {
                "host": {
                    "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
                    "connectionName": "shared_sharepointonline",
                    "operationId": "CreateFile",
                },
                "parameters": {
                    "dataset": sp["site"],
                    "folderPath": resume_folder_path(
                        "first(body('Filter_submitted'))?['Received Date']"),
                    "name": "@{%s}_@{outputs('FileRef_from_dup')}%s.%s"
                            % (_dup_name_part, _multi_suffix("Save_dup_update_resume"),
                               _dup_orig_ext),
                    "body": "@base64ToBinary(items('Save_dup_update_resume')?['contentBytes'])",
                },
                "authentication": "@parameters('$authentication')",
            },
        },
    },
}

_patch_dup = _make_patch_item("Filter_submitted",
                               {"Send_duplicate_notice": ["Succeeded", "Failed", "Skipped"]},
                               stamp_last_updated=True)
set_excel(_patch_dup)
_patch_dup["inputs"]["parameters"]["item"]["Status"] = "New Email Received"
# Resume Filename is deliberately NOT changed - the file was overwritten in place under its original
# name, so the column still points at the right file (see OVERWRITE fix above).

# Contact reply_cap and beyond (but still under the overall dup_cap): the resend is still saved
# and the row still re-queued for P2 - only the EMAIL stops ("no other reminders").
_patch_dup_noreply = _make_patch_item_noreply("Filter_submitted", {}, requeue=True,
                                              stamp_last_updated=True)
set_excel(_patch_dup_noreply)

dup["actions"] = {
    "FileRef_from_dup": _fileref_dup,
    "IsUnderDupCap": {
        "type": "If", "runAfter": {"FileRef_from_dup": ["Succeeded"]},
        "expression": _cv_check_expr("Filter_submitted", dup_cap),
        "actions": {
            "Ensure_dup_resume_folder": _ensure_dup_folder,
            "Save_dup_update_resume": _save_dup_resume,
            "IsUnderReplyCap_dup": {
                "type": "If",
                "runAfter": {"Save_dup_update_resume": ["Succeeded", "Failed", "Skipped"]},
                "expression": _cv_check_expr("Filter_submitted", reply_cap),
                "actions": {
                    "Send_duplicate_notice": _send_dup,
                    "Patch_dup_attempts": _patch_dup,
                    # CATCH: the dup-cap increment is best-effort. If the PatchItem fails (e.g.
                    # legacy rows that share an Application ID, so 'update by key' can't find a
                    # unique row), this no-op consumes the failure so the RUN still succeeds - the
                    # duplicate notice was sent, and the email still gets marked read + archived
                    # instead of sitting unread indefinitely (an unread email is NOT auto-retried
                    # by the trigger - see Notify_failure's own comment).
                    "Dup_patch_done": {
                        "type": "Compose",
                        "runAfter": {"Patch_dup_attempts": ["Succeeded", "Failed", "Skipped"]},
                        "inputs": "dup-cap patch best-effort",
                    },
                },
                "else": {"actions": {
                    "Patch_dup_attempts_noreply": _patch_dup_noreply,
                    # CATCH (2026-08-27). This patch used to be the LAST action in its scope,
                    # which was deliberate: with nothing after it, a failure propagated up to
                    # HasResume, turned the run RED and fired Notify_failure as well, so the
                    # one swallowed-by-design path still had a loud signal.
                    #
                    # That trade is now wrong. Rows and resumes get deleted deliberately while
                    # the flow is polling every 60 seconds - resetting a candidate to re-run
                    # them is normal operating procedure here. Get_rows reads a row at second
                    # 2, the row is deleted, and the patch at second 15 has nothing to write
                    # to. A routine data reset should not surface as a failed flow run.
                    #
                    # The failure is NOT hidden: Notify_dup_attempts_noreply_failed still
                    # fires on [Failed, TimedOut] exactly as before - it is a sibling and is
                    # unaffected by this Compose. All this changes is that the RUN stays
                    # green, matching the other six patch paths (Ack_stamp_done,
                    # Dup_patch_done and the four Mark_as_read_* absorbers). Same depth as the
                    # patch, so it adds no nesting - which matters, because the update and
                    # follow-up patches already sit at the depth-8 ceiling.
                    "Dup_noreply_patch_done": {
                        "type": "Compose",
                        "runAfter": {"Patch_dup_attempts_noreply":
                                     ["Succeeded", "Failed", "Skipped", "TimedOut"]},
                        "inputs": "dup no-reply patch best-effort "
                                  "(a row deleted mid-run is expected, not a failure)",
                    },
                }},
            },
        },
        "else": {"actions": {}},
    },
}

# ── 14b. Ref-reply caps + fake-ref guard (inside HasAppRefGate TRUE) ────────
gate_acts = gate["actions"]

gate_acts["Get_rows_ref"] = _make_get_items({})
set_excel(gate_acts["Get_rows_ref"])

# ── ref lookup: by Application ID for allow-listed testers, by sender for everyone else ──────
# Default behaviour is unchanged and deliberate: a reply is matched to a row by WHO SENT IT, never
# by the APP-... text it quotes, so one applicant can never patch another applicant's row by
# quoting their reference (references appear in every acknowledgment email, so they are not
# secret). That rule also makes internal testing impossible - yashv@ quoting a candidate's
# reference has no row of their own, falls through to intake, and lands as a brand-new row.
#
# So for allow-listed senders ONLY, and only when a reference was actually quoted, look the row up
# by Application ID instead. Everything downstream already derives from whatever row Get_rows_ref
# returns - AppRef_current, the resume filename, and every Patch_* (which addresses rows by
# idColumn 'Application ID') - so redirecting this one filter redirects the whole update path.
gate_acts["Get_rows_ref"]["inputs"]["parameters"]["$filter"] = (
    "@if(and(greater(length(body('MatchAllowSender')), 0), "
    "not(equals(outputs('QuotedRef'), ''))), "
    "concat('Application ID eq ''', outputs('QuotedRef'), ''''), "
    "%s)" % _EMAIL_FILTER[1:]
)

# AppRef_current: the MATCHED row's real, live Application ID (Get_rows_ref matches by sender
# email, not by whatever ref the email happens to quote). Fixed 2026-07-31: the saved
# update-resume filename and the "Ref:" shown back to the applicant used to be built from
# AppRef_from_subject instead - the reference literally quoted in the email text. That's fine
# for a normal thread, but a sender who is still replying to an OLD email thread (their original
# application never got a row due to a since-fixed flow bug, so every reply since then keeps
# quoting that dead reference) would have their real, current row's Application ID silently
# ignored: Get_rows_ref finds the right row and Patch_update_attempts/Patch_followup_attempts
# correctly patch it (by Application ID from the matched row), but the resume got saved under
# the OLD quoted ref instead - a file P2's _resume_name_slots() (sharepoint_scoring.py), which
# only ever looks for the row's OWN current Application ID, can never find. AppRef_current is
# always the row's own truth. AppRef_from_subject (the ref the sender literally quoted) became
# unused at that point and was deleted 2026-08-24 along with RefSrc, its only input - see the
# removal note further down for why the pristine snapshot had to be rewritten in the same pass.
gate_acts["AppRef_current"] = {
    "type": "Compose",
    "runAfter": {},
    "inputs": "@{first(body('Get_rows_ref')?['value'])?['Application ID']}",
}

_fileref_subj = _deep_find(gate_acts, "FileRef_from_subject")
_has_resume_up = _deep_find(gate_acts, "Has_resume_in_update")
_save_up = _deep_find(gate_acts, "Save_update_resume")
_send_up = _deep_find(gate_acts, "Send_update_ack")
_term_up = _deep_find(gate_acts, "Terminate_update")
_send_fu = _deep_find(gate_acts, "Send_noted_reply")
_term_fu = _deep_find(gate_acts, "Terminate_noted")

# RefSrc / AppRef_from_subject were REMOVED 2026-08-24 (dead code, see AppRef_current above).
# They extracted the APP-... token the sender literally quoted, which every consumer stopped
# using on 2026-07-31 when the ack subject, the in-body "Update - ..." example and the saved
# update-resume filename all switched to AppRef_current (the MATCHED ROW's own live Application
# ID). After that they computed a value nothing read: two Compose actions per ref-reply, plus a
# comment block claiming three consumers that no longer existed. RefSrc had no other consumer
# than AppRef_from_subject, so both went together, along with the pristine snapshot's last two
# references to them (_pristine_send_actions.json now names AppRef_current directly, matching
# what every build already overlaid). Deleting the action while that snapshot still referenced
# it would have left a dangling reference the moment an overlay was skipped - which is why they
# had to be removed as a pair, not one at a time.
_fileref_subj["runAfter"] = {"AppRef_current": ["Succeeded"]}
_has_resume_up["runAfter"] = {"FileRef_from_subject": ["Succeeded"]}

# --- Resume-update: Ensure folder + Save + Send + PATCH CV Attempts + Terminate ---
_ensure_update_folder = _make_ensure_resume_folder(
    "first(body('Get_rows_ref')?['value'])?['Received Date']")
_save_up["runAfter"] = {"Ensure_update_resume_folder": ["Succeeded", "Failed"]}
_send_up["runAfter"] = {}   # now the first action inside the nested IsUnderReplyCap_update branch
_patch_up = _make_patch_item("Get_rows_ref",
                              {"Send_update_ack": ["Succeeded", "Failed", "Skipped"]},
                              excel=True, stamp_last_updated=True)
set_excel(_patch_up)
# Updated CV: increment CV Attempts and re-queue the row for Phase 2 (Status → "New Email Received")
# so the new resume gets rescored. Resume Filename is deliberately NOT changed - the update resume is
# saved OVER the original file under its original name (see OVERWRITE fix below), so the column still
# points at the correct (now-updated) file and P2 downloads it fine.
_patch_up["inputs"]["parameters"]["item"]["Status"] = "New Email Received"
_term_up["runAfter"] = {"Patch_update_attempts": ["Succeeded", "Failed"]}

# Contact reply_cap and beyond (but still under update_cap): resume still saved + row still
# re-queued for P2 - only the email stops.
_patch_up_noreply = _make_patch_item_noreply("Get_rows_ref", {}, excel=True, requeue=True,
                                             stamp_last_updated=True)
set_excel(_patch_up_noreply)
_term_up_noreply = {"type": "Terminate", "runAfter": {"Patch_update_attempts_noreply": ["Succeeded", "Failed"]},
                     "inputs": {"runStatus": "Succeeded"}}

_has_resume_up["actions"] = {
    "IsUnderUpdateCap": {
        "type": "If", "runAfter": {},
        "expression": _cv_check_expr("Get_rows_ref", update_cap, excel=True),
        "actions": {
            "Ensure_update_resume_folder": _ensure_update_folder,
            "Save_update_resume": _save_up,
            "IsUnderReplyCap_update": {
                "type": "If",
                "runAfter": {"Save_update_resume": ["Succeeded", "Failed", "Skipped"]},
                "expression": _cv_check_expr("Get_rows_ref", reply_cap, excel=True),
                "actions": {
                    "Send_update_ack": _send_up,
                    "Patch_update_attempts": _patch_up, "Terminate_update": _term_up,
                },
                "else": {"actions": {
                    "Patch_update_attempts_noreply": _patch_up_noreply,
                    "Terminate_update_noreply": _term_up_noreply,
                }},
            },
        },
        "else": {"actions": {
            "Terminate_update_silent": {"type": "Terminate", "runAfter": {},
                                        "inputs": {"runStatus": "Succeeded"}},
        }},
    },
}

# --- Follow-up: Send + PATCH CV Attempts + Terminate ---
_send_fu["runAfter"] = {}
# stamp_last_updated=True (fixed 2026-07-31): a text-only follow-up is still a genuine candidate
# contact and must move Last Updated Date same as the dup/update paths - this was the one patch
# site silently omitting it, which P2's _updated_after_location_request() depends on to detect
# a reply.
_patch_fu = _make_patch_item("Get_rows_ref",
                              {"Send_noted_reply": ["Succeeded", "Failed", "Skipped"]},
                              excel=True, stamp_last_updated=True)
set_excel(_patch_fu)
_term_fu["runAfter"] = {"Patch_followup_attempts": ["Succeeded", "Failed"]}

# Contact reply_cap and beyond (but still under followup_cap): no resume in a follow-up, so
# nothing to re-queue - just the increment, no email.
_patch_fu_noreply = _make_patch_item_noreply("Get_rows_ref", {}, excel=True, stamp_last_updated=True)
set_excel(_patch_fu_noreply)
_term_fu_noreply = {"type": "Terminate", "runAfter": {"Patch_followup_attempts_noreply": ["Succeeded", "Failed"]},
                     "inputs": {"runStatus": "Succeeded"}}

_has_resume_up["else"]["actions"] = {
    "IsUnderFollowupCap": {
        "type": "If", "runAfter": {},
        "expression": _cv_check_expr("Get_rows_ref", followup_cap, excel=True),
        "actions": {
            "IsUnderReplyCap_followup": {
                "type": "If", "runAfter": {},
                "expression": _cv_check_expr("Get_rows_ref", reply_cap, excel=True),
                "actions": {
                    "Send_noted_reply": _send_fu, "Patch_followup_attempts": _patch_fu,
                    "Terminate_noted": _term_fu,
                },
                "else": {"actions": {
                    "Patch_followup_attempts_noreply": _patch_fu_noreply,
                    "Terminate_noted_noreply": _term_fu_noreply,
                }},
            },
        },
        "else": {"actions": {
            "Terminate_noted_silent": {"type": "Terminate", "runAfter": {},
                                       "inputs": {"runStatus": "Succeeded"}},
        }},
    },
}

# --- Wrap in IsKnownSender (fake-ref guard) ---
gate["actions"] = {
    "Get_rows_ref": gate_acts["Get_rows_ref"],
    "IsKnownSender": {
        "type": "If",
        # Failed/Skipped are handled here for the same reason Filter_submitted handles them
        # on Get_rows: an unguarded read makes HasAppRefGate fail, and NOTHING watches that
        # gate for Failed (HasResume only runs on Succeeded/Skipped, and Notify_failure only
        # watches HasResume). A locked or missing workbook on a reply carrying 'app-20' would
        # otherwise skip the alert AND the quarantine move, leaving the message in the Inbox
        # to be re-polled and re-failed silently. The expression already coalesces a missing
        # body to [], so a failed read evaluates false and falls through to normal intake.
        "runAfter": {"Get_rows_ref": ["Succeeded", "Failed", "Skipped"]},
        "expression": {"greater": [
            "@length(coalesce(outputs('Get_rows_ref')?['body/value'], json('[]')))", 0
        ]},
        "actions": {
            "AppRef_current": gate_acts["AppRef_current"],
            "FileRef_from_subject": _fileref_subj,
            "Has_resume_in_update": _has_resume_up,
        },
        "else": {"actions": {}},
    },
}

# ── 15. email reply content (all 6 applicant-facing emails) ──────────────────

_FOOTER = (
    '<hr style="border:none;border-top:1px solid #ddd;margin:24px 0 12px 0;">'
    '<p style="font-size:12px;color:#888;line-height:1.5;">'
    'This is an auto-generated email and this mailbox is not monitored.</p>'
)
_STYLE = 'style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222;line-height:1.7;"'


def _set_email(action, subject: str, body_html: str) -> None:
    """Overwrite an email action's subject and body.

    Also disables PA's automatic action retry (2026-07-12, real-world finding): the live
    2026-07 backfill run sent 2 applicants the SAME acknowledgment 4 times each - 1 row/AppRef
    each (Add_row ran once), but Send_acknowledgment itself fired repeatedly, matching PA's
    default retry count exactly. Root cause: the connector call likely succeeded but PA's
    orchestrator perceived a transient failure (slow/dropped response) and retried a send that
    wasn't actually failing - a real send has no idempotency guard, so each retry is a genuine
    duplicate email to a real applicant. Trade-off of disabling retry: a genuine one-off
    transient failure no longer self-heals - it just fails once, same as any other failure path
    (Notify_failure alerts the admin, email stays unread for manual reprocessing). Silently
    duplicate-emailing applicants is worse than that, so retry is OFF for every applicant-facing
    send."""
    p = action["inputs"]["parameters"]
    p["emailMessage/Subject"] = subject
    p["emailMessage/Body"] = f'<div {_STYLE}>{body_html}{_FOOTER}</div>'
    action["runtimeConfiguration"] = {"retryPolicy": {"type": "none"}}


# Shared final paragraph — sent on the LAST reply before the cap silences the sender (CV Attempts
# == cap-1), and only that one. Apostrophe-free so it is safe inside the single-quoted if() literal.
# No "update your resume" invitation - this is genuinely the final auto-reply.
_FINAL_CLOSING = (
    "<p>This is our final automated reply regarding this application. If there is a match, "
    "a member of our team will contact you directly. You do not need to reply, and you are "
    "welcome to apply again after 90 days.</p>"
)


def _closing(cv_src: str, threshold: int, normal_a: str, ref_expr: str, normal_b: str) -> str:
    """One closing-paragraph expression that SWAPS (not appends): the short final paragraph once
    CV Attempts >= threshold, otherwise the normal closing rebuilt via concat() so its dynamic
    application-reference still renders. Callers pass threshold = cap-1 so this fires on exactly
    ONE reply - the last one that will ever send (CV Attempts one below the cap) - never twice. All
    literals are apostrophe-free (safe inside the single-quoted if()/concat() strings)."""
    return ("@{if(greaterOrEquals(int(coalesce(%s,'0')),%d),'%s',concat('%s',%s,'%s'))}"
            % (cv_src, threshold, _FINAL_CLOSING, normal_a, ref_expr, normal_b))


_CV_DUP = "first(body('Filter_submitted'))?['Application Updates']"
_CV_REF = "first(body('Get_rows_ref')?['value'])?['Application Updates']"


# ── Email 1: Application acknowledged (happy path — new applicant with resume)
_set_email(
    _deep_find(dup["else"]["actions"], "Send_acknowledgment"),
    "Application Received - DriverAI (Ref: @{outputs('AppRef')})",

    '<p>Hello,</p>'

    '<p>Thanks for applying to <strong>DriverAI</strong>. '
    'Your application has been received and is under review.</p>'

    '<p>Reference number: <strong>@{outputs(\'AppRef\')}</strong></p>'

    '<p>If you are selected, a member of our team will contact you to discuss next steps. '
    'You do not need to reply unless you are updating your resume. To do that, send a new or reply '
    'email to <em>apply@driverai.io</em> with the subject '
    '<strong>"Update - @{outputs(\'AppRef\')}"</strong> and attach the new file.</p>'

    '<p>Good luck on your next journey.</p>'
    '<p>Warm regards,<br><strong>The DriverAI Recruiting Team</strong></p>',
)

# ── Email 2: Already on file (duplicate resume within 90 days)
_set_email(
    _deep_find(dup["actions"], "Send_duplicate_notice"),
    "Your DriverAI Application Is Already On File (Ref: @{first(body('Filter_submitted'))?['Application ID']})",

    '<p>Hello,</p>'

    '<p>Your application is already on file with <strong>DriverAI</strong> and under '
    'review, so there\'s no need to resubmit.</p>'

    + _closing(
        _CV_DUP, reply_cap - 1,
        '<p>If you are selected, a member of our team will contact you to discuss next steps. '
        'You do not need to reply unless you are updating your resume. To do that, send a new or '
        'reply email to <em>apply@driverai.io</em> with the subject <strong>"Update - ',
        "first(body('Filter_submitted'))?['Application ID']",
        '"</strong> and attach the new file.</p>',
    ) +

    '<p>Good luck on your next journey.</p>'
    '<p>Warm regards,<br><strong>The DriverAI Recruiting Team</strong></p>',
)

# ── Email 3: Missing resume (application keywords but no CV attached)
_set_email(
    _deep_find(kw["actions"], "Send_CV_request"),
    "Please Attach Your Resume - DriverAI",

    '<p>Hello,</p>'

    '<p>Thanks for your interest in <strong>DriverAI</strong>. We didn\'t see a resume '
    'attached, so your application isn\'t complete yet.</p>'

    '<p>Please reply with your resume attached in <strong>PDF or Word (.docx)</strong> '
    'format. Once we have it, your application goes under review, and if you are selected, '
    'a member of our team will contact you to discuss next steps.</p>'

    '<p>Good luck on your next journey.</p>'
    '<p>Warm regards,<br><strong>The DriverAI Recruiting Team</strong></p>',
)

# ── Email 4: Wrong file format (attachment is not PDF/DOCX)
_set_email(
    _deep_find(hr["else"]["actions"], "Send_wrong_format"),
    "Please Resend Your Resume as PDF or Word - DriverAI",

    '<p>Hello,</p>'

    '<p>Thanks for your interest in <strong>DriverAI</strong>. We received your '
    'submission, but we couldn\'t open the attached file.</p>'

    '<p>Please reply with your resume as a <strong>PDF (.pdf)</strong> or '
    '<strong>Word (.docx)</strong> file and we\'ll process it right away. If you are '
    'selected, a member of our team will contact you to discuss next steps.</p>'

    '<p>Good luck on your next journey.</p>'
    '<p>Warm regards,<br><strong>The DriverAI Recruiting Team</strong></p>',
)

# ── Email 5: Resume update acknowledged (ref-reply with a new resume)
_set_email(
    _deep_find(gate_acts, "Send_update_ack"),
    "Updated Resume Received - DriverAI (Ref: @{outputs('AppRef_current')})",

    '<p>Hello,</p>'

    '<p>Thanks for sending your updated resume. Your application now reflects the latest '
    'version and our team will review it.</p>'

    + _closing(
        _CV_REF, reply_cap - 1,
        '<p>If you are selected, a member of our team will contact you to discuss next steps. '
        'You do not need to reply unless you are updating your resume again. To do that, send a new '
        'or reply email to <em>apply@driverai.io</em> with the subject <strong>"Update - ',
        "outputs('AppRef_current')",
        '"</strong> and attach the new file.</p>',
    ) +

    '<p>Good luck on your next journey.</p>'
    '<p>Warm regards,<br><strong>The DriverAI Recruiting Team</strong></p>',
)

# ── Email 6: Follow-up noted (ref-reply without a resume)
_set_email(
    _deep_find(gate_acts, "Send_noted_reply"),
    "Message Received - DriverAI (Ref: @{outputs('AppRef_current')})",

    '<p>Hello,</p>'

    '<p>Thanks for following up. Your message has been noted alongside your application.</p>'

    + _closing(
        _CV_REF, reply_cap - 1,
        '<p>If you are selected, a member of our team will contact you to discuss next steps. '
        'You do not need to reply unless you are updating your resume. To send one, reply with the '
        'subject <strong>"Update - ',
        "outputs('AppRef_current')",
        '"</strong> and attach the file (<strong>PDF or Word</strong>).</p>',
    ) +

    '<p>Good luck on your next journey.</p>'
    '<p>Warm regards,<br><strong>The DriverAI Recruiting Team</strong></p>',
)

# ── Internal alert: a run did not complete (admin-only, NO applicant footer) ──
# Dynamic on purpose - one alert covers ANY failure. Pulls whatever context exists: the
# reference from the new-application branch, plus sender / subject / received time, all
# coalesced so a missing value never breaks the message.
# NOTE: cannot coalesce anything computed inside HasAppRefGate's conditional true-branch here
# (AppRef_current, and formerly AppRef_from_subject before its 2026-08-24 removal) - those only
# run if the email quoted a ref, so they are not on Notify_failure's guaranteed runAfter path and
# referencing one fails PA's template validation ("InvalidTemplate" -
# the same bug class fixed on IsSpam/HasValidResumeEarly above). AppRef IS safe to reference here:
# it's an unconditional sibling action in this same scope. The raw Subject line below still shows
# the quoted reference text for ref-reply-branch failures, so no information is actually lost.
# Retries ARE wanted here (changed 2026-08-03) - see Notify_poll_failure for the reasoning:
# the no-retry rule protects MoveV2/MarkAsRead from racing the mutable Outlook message ID,
# which does not apply to a send, and this alert is the primary failure signal.
spam["Notify_failure"]["runtimeConfiguration"] = {"retryPolicy": {"type": "exponential",
                                                                 "count": 2,
                                                                 "interval": "PT10S"}}
_notify = spam["Notify_failure"]["inputs"]["parameters"]
_ref_any = "@{coalesce(outputs('AppRef'),'n/a')}"
_notify["emailMessage/Subject"] = "[Hiring Auto-Reply] A run needs attention (Ref %s)" % _ref_any
_notify["emailMessage/Body"] = (
    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222;line-height:1.6;">'
    '<p><strong>Hiring Auto-Reply: a run did not complete.</strong></p>'
    '<p>Something went wrong while handling an incoming email, so the applicant may not '
    'have received a reply and the candidate row or resume file may not have been saved.</p>'
    '<ul>'
    '<li>Reference: <strong>' + _ref_any + '</strong></li>'
    '<li>From: @{coalesce(triggerOutputs()?[\'body/from\'],\'unknown\')}</li>'
    '<li>Subject: @{coalesce(triggerOutputs()?[\'body/subject\'],\'(none)\')}</li>'
    '<li>Received: @{coalesce(triggerOutputs()?[\'body/receivedDateTime\'],\'(unknown)\')}</li>'
    '</ul>'
    '<p>Open the flow run history for the exact step and error. '
    'The failed email is moved unread to Archive so it cannot block the Inbox queue.</p>'
    '</div>'
)

# ── 16. inbox tidy: move mail out of the Inbox by category ──────────────────
# TWO destinations (changed 2026-07-04):
#   LEGITIMATE application mail -> mark READ (MarkAsRead_V3) + move to `destination` (Archive).
#   SPAM/JUNK (the 3 gates)     -> just MOVE to `spam_destination` (Junk Email), left UNREAD.
# Archive holds successful legitimate mail (read) plus admin-alerted failures (unread);
# junk lands in Junk Email standing out as unread.
# Uses MoveV2 (+ optional MarkAsRead_V3) on the shared mailbox. Non-blocking: the move runs even if
# a preceding mark fails, and every terminal move has a success/failure/timeout continuation. This
# matters because Outlook IDs are mutable and MoveV2 can return ErrorItemNotFound after another
# actor has already moved the message. An email moved out of the Inbox is not re-polled (the trigger
# reads the Inbox only), so 'unread in Junk' is safe.
# Skipped entirely when inbox_tidy.enabled is false.
tidy_cfg = cfg.get("inbox_tidy", {})
if tidy_cfg.get("enabled", False):
    _tidy_mailbox = email["trigger_mailbox"]
    _tidy_dest = tidy_cfg.get("destination", "Archive")
    _tidy_spam_dest = tidy_cfg.get("spam_destination", "Junk Email")
    _tidy_failure_dest = tidy_cfg.get("failure_destination", "")
    _tidy_marks = 0
    _tidy_moves = 0
    _tidy_finalizers = 0

    def _tidy_mark(run_after: dict) -> dict:
        return {
            "type": "OpenApiConnection",
            "runAfter": run_after,
            "inputs": {
                "parameters": {
                    "messageId": "@triggerOutputs()?['body/id']",
                    "mailboxAddress": _tidy_mailbox,
                    "body/isRead": True,  # V3 is "mark read OR unread"; absent flag = unread
                },
                "host": {
                    "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
                    "connectionName": "shared_office365",
                    "operationId": "MarkAsRead_V3",
                },
                "authentication": "@parameters('$authentication')",
            },
        }

    def _tidy_move(run_after: dict, destination: str) -> dict:
        return {
            "type": "OpenApiConnection",
            "runAfter": run_after,
            "inputs": {
                # Move is a non-idempotent operation. Do not let the connector replay a
                # request whose first attempt may already have moved the item and changed
                # its ordinary Outlook ID. Cleanup failures are handled explicitly below.
                "retryPolicy": {"type": "none"},
                "parameters": {
                    "messageId": "@triggerOutputs()?['body/id']",
                    "mailboxAddress": _tidy_mailbox,
                    "folderPath": destination,
                },
                "host": {
                    "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
                    "connectionName": "shared_office365",
                    "operationId": "MoveV2",
                },
                "authentication": "@parameters('$authentication')",
            },
        }

    def _inject_tidy_before(acts: dict, term_name: str, suffix: str,
                            mark_read: bool, destination: str) -> None:
        """Insert tidy actions BEFORE an existing Terminate.

        mark_read=True  -> Mark_as_read_<suffix> then Move_to_processed_<suffix> (legit -> Archive, read).
        mark_read=False -> Move_to_processed_<suffix> only (spam -> Junk Email, left unread).

        Idempotent: on re-builds the base zip may still carry stale tidy actions (possibly with a
        circular dependency, and possibly a mark that this build no longer wants). Strip BOTH names
        first, recover the Terminate's real predecessor, then re-inject cleanly for the new mode.
        """
        global _tidy_marks, _tidy_moves
        term = acts[term_name]
        mark_n = "Mark_as_read_%s" % suffix
        move_n = "Move_to_processed_%s" % suffix
        # Recover the Terminate's original predecessor (whatever isn't one of our tidy actions).
        stale = {mark_n, move_n}
        if mark_n in acts or move_n in acts:
            src = acts.get(mark_n, acts.get(move_n, {}))
            real_ra = {k: v for k, v in src.get("runAfter", {}).items() if k not in stale}
            acts.pop(mark_n, None)
            acts.pop(move_n, None)
            term["runAfter"] = real_ra if real_ra else {}
        orig_ra = copy.deepcopy(term.get("runAfter", {}))
        if mark_read:
            acts[mark_n] = _tidy_mark(orig_ra)
            acts[move_n] = _tidy_move(
                {mark_n: ["Succeeded", "Failed", "TimedOut"]}, destination
            )
            _tidy_marks += 1
        else:
            acts[move_n] = _tidy_move(orig_ra, destination)
        term["runAfter"] = {move_n: ["Succeeded", "Failed", "Skipped", "TimedOut"]}
        _tidy_moves += 1

    # 16a. spam gates (3 Terminates) -> Junk Email, UNREAD (move only, no mark-read)
    _inject_tidy_before(
        actions["IsSystem"]["actions"], "Terminate_System", "spam_sender",
        mark_read=False, destination=_tidy_spam_dest)
    _inject_tidy_before(
        actions["IsSystem"]["else"]["actions"]["IsSubject"]["actions"],
        "Terminate_Subject", "spam_subject", mark_read=False, destination=_tidy_spam_dest)
    _inject_tidy_before(
        actions["IsSystem"]["else"]["actions"]["IsSubject"]["else"]["actions"]["IsSpam"]["actions"],
        "Terminate_Spam", "spam_body", mark_read=False, destination=_tidy_spam_dest)

    # 16b. ref-reply Terminates (6) -> Archive, READ (legit applicants who quoted a ref).
    # Each of update/followup now has an extra reply_cap-reached-but-not-yet-cap-reached
    # Terminate (a real contact that gets no email, but is still legit application mail).
    _uc = _has_resume_up["actions"]["IsUnderUpdateCap"]
    _ucr = _uc["actions"]["IsUnderReplyCap_update"]
    _inject_tidy_before(_ucr["actions"], "Terminate_update", "update",
                        mark_read=True, destination=_tidy_dest)
    _inject_tidy_before(_ucr["else"]["actions"], "Terminate_update_noreply", "update_noreply",
                        mark_read=True, destination=_tidy_dest)
    _inject_tidy_before(_uc["else"]["actions"], "Terminate_update_silent", "update_cap",
                        mark_read=True, destination=_tidy_dest)
    _fc = _has_resume_up["else"]["actions"]["IsUnderFollowupCap"]
    _fcr = _fc["actions"]["IsUnderReplyCap_followup"]
    _inject_tidy_before(_fcr["actions"], "Terminate_noted", "followup",
                        mark_read=True, destination=_tidy_dest)
    _inject_tidy_before(_fcr["else"]["actions"], "Terminate_noted_noreply", "followup_noreply",
                        mark_read=True, destination=_tidy_dest)
    _inject_tidy_before(_fc["else"]["actions"], "Terminate_noted_silent", "followup_cap",
                        mark_read=True, destination=_tidy_dest)

    # 16c. top-level pair -> Archive, READ. Runs after HasResume succeeds (new-applicant, duplicate,
    # CV-request, wrong-format, ignored). Does NOT run on HasResume failure (Notify_failure handles
    # that - the email stays unread as an admin flag; the trigger watermark is time-based so this is
    # not an auto-retry, the alert is the recovery path). Base-zip leftovers are popped first.
    spam.pop("Mark_as_read", None)
    spam.pop("Move_to_processed", None)
    spam.pop("Finalize_processed_cleanup", None)
    spam["Mark_as_read"] = _tidy_mark({"HasResume": ["Succeeded"]})
    spam["Move_to_processed"] = _tidy_move(
        {"Mark_as_read": ["Succeeded", "Failed", "TimedOut"]}, _tidy_dest
    )
    spam["Finalize_processed_cleanup"] = {
        "type": "Compose",
        "runAfter": {
            "Move_to_processed": ["Succeeded", "Failed", "TimedOut"],
        },
        "inputs": (
            "Processed-message terminal cleanup handled. Candidate processing is complete; "
            "inspect Move_to_processed in run history if Outlook rejected the cleanup."
        ),
    }
    _tidy_marks += 1
    _tidy_moves += 1
    _tidy_finalizers += 1

    # A scheduled unread poll would otherwise select the same permanently failing
    # message every minute and starve the rest of the queue. After the admin alert,
    # move that message unread to the configured failure destination. Production uses
    # Archive: MoveV2 accepts that well-known shared-mailbox folder consistently, while
    # it rejected the nested Recruiting Review Graph ID at runtime. The move is attempted
    # even if the alert itself fails; unread status distinguishes failures in Archive.
    spam.pop("Move_failed_to_recruiting_review", None)
    spam.pop("Move_failed_to_archive_unread", None)
    spam.pop("Finalize_failed_cleanup", None)
    if _tidy_failure_dest:
        spam["Move_failed_to_archive_unread"] = _tidy_move(
            # CRITICAL: never include Skipped here. Notify_failure is intentionally
            # skipped when HasResume succeeds. Including Skipped made this failure move
            # race Move_to_processed on every successful email; one move won and the
            # other then failed with ErrorItemNotFound against the now-stale message ID.
            {"Notify_failure": ["Succeeded", "Failed", "TimedOut"]},
            _tidy_failure_dest,
        )
        spam["Finalize_failed_cleanup"] = {
            "type": "Compose",
            "runAfter": {
                "Move_failed_to_archive_unread": ["Succeeded", "Failed", "TimedOut"],
            },
            "inputs": (
                "Failed-message terminal cleanup handled. The admin alert is authoritative; "
                "inspect the failed move if the message remains in Inbox."
            ),
        }
        _tidy_moves += 1
        _tidy_finalizers += 1

    _tidy_total = _tidy_marks + _tidy_moves + _tidy_finalizers
else:
    _tidy_total = 0

# ── 16d. PATCH-FAILURE ALERTS (added 2026-08-03) ────────────────────────────
# Notify_failure only watches HasResume. Every PatchItem below is "handled" by whatever
# runs after it (the runAfter lists include Failed so inbox cleanup is unconditional and a
# poisoned message can never block the queue). That resilience has a cost: the enclosing
# If therefore reports Succeeded, the branch ends on a Terminate whose runStatus is
# Succeeded, and the run goes GREEN. A failed patch means the candidate's Application
# Updates counter never incremented - and on the dup/update paths, that the row was never
# re-queued for Phase 2 - with no signal anywhere. That was silent until now.
#
# Each alert is a plain SIBLING of the patch it watches, wired on [Failed, TimedOut]. It
# adds NO nesting: Patch_update_attempts / Patch_followup_attempts already sit at depth 8,
# the Power Automate ceiling, so a child action there would fail import - a sibling is at
# the same depth and is fine. Nothing runs after these alerts and nothing depends on them,
# so they cannot alter any existing path, and a failing alert is inert.
# value = (human label, is_swallowed). is_swallowed records whether a failure of that patch is
# absorbed inside its own branch - i.e. some sibling runs after it with Failed in runAfter - which
# is what makes the enclosing If report Succeeded and the run go GREEN with no other signal.
# Patch_dup_attempts_noreply is the one exception: it is the LAST action in its scope, so its
# failure propagates all the way up to HasResume, turning the run RED and firing Notify_failure
# too. Its alert is therefore belt-and-braces rather than the sole signal, and says so - keeping
# it means that if a sibling is ever appended to that scope (which would start swallowing the
# failure, exactly as on the reply-path twin), the coverage is already in place. test_p1.py locks
# this table against the real wiring so the two can never drift silently.
_PATCH_ALERTS = {
    "Patch_mail_sent":                 ("new applicant - acknowledgment stamp",     True),
    "Patch_dup_attempts":              ("duplicate re-application",                 True),
    # Swallowed as of 2026-08-27 (was the one PROPAGATING path). Dup_noreply_patch_done
    # now absorbs it so a row deleted mid-run cannot turn a routine run red - see the
    # catch comment in the IsUnderReplyCap_dup else-branch.
    "Patch_dup_attempts_noreply":      ("duplicate re-application (past reply cap)", True),
    "Patch_update_attempts":           ("resume update",                            True),
    "Patch_update_attempts_noreply":   ("resume update (past reply cap)",           True),
    "Patch_followup_attempts":         ("follow-up reply",                          True),
    "Patch_followup_attempts_noreply": ("follow-up reply (past reply cap)",         True),
}


def _make_patch_alert(patch_name: str, what: str, swallowed: bool = True) -> dict:
    """Admin alert fired when a candidate-row PATCH fails. Sibling of `patch_name`."""
    return {
        "type": "OpenApiConnection",
        "runAfter": {patch_name: ["Failed", "TimedOut"]},
        # Unlike the MoveV2/MarkAsRead actions (retries disabled there because a retry
        # races the mutable Outlook message ID), a send has no such hazard and this is
        # the only signal that survives a silent patch failure - so let it retry.
        "runtimeConfiguration": {"retryPolicy": {"type": "exponential", "count": 2,
                                                 "interval": "PT10S"}},
        "inputs": {
            "parameters": {
                "emailMessage/To": email["admin_email"],
                "emailMessage/Subject":
                    "[Hiring Auto-Reply] Candidate row NOT updated - %s (Ref %s)"
                    % (what, _ref_any),
                "emailMessage/Body": (
                    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;'
                    'color:#222;line-height:1.6;">'
                    '<p><strong>Hiring Auto-Reply: a candidate row failed to update.</strong></p>'
                    '<p>The incoming email was handled and filed normally, but <strong>'
                    + patch_name + '</strong> (' + what + ') did not commit. The '
                    'Application Updates counter was NOT incremented'
                    + (', and the row was NOT re-queued for Phase 2 scoring'
                       if "dup" in patch_name or "update" in patch_name else '')
                    + '.</p>'
                    '<ul>'
                    '<li>Reference: <strong>' + _ref_any + '</strong></li>'
                    '<li>From: @{coalesce(outputs(\'CurrentEmail\')?[\'from\'],\'unknown\')}</li>'
                    '<li>Subject: @{coalesce(outputs(\'CurrentEmail\')?[\'subject\'],\'(none)\')}</li>'
                    '<li>Received: @{coalesce(outputs(\'CurrentEmail\')?[\'receivedDateTime\'],\'(unknown)\')}</li>'
                    '</ul>'
                    + ('<p>The run itself reports Succeeded because inbox cleanup completed '
                       'as designed - <strong>this alert is the only signal</strong>. Open run '
                       'history and inspect <strong>' + patch_name + '</strong>, then correct '
                       'the row by hand.</p>'
                       if swallowed else
                       '<p>This patch is the last action in its branch, so the failure also '
                       'propagates up: the run reports <strong>Failed</strong> and the standard '
                       'Notify_failure alert fires as well. Expect two emails for this one - '
                       'this is the belt-and-braces copy. Open run history, inspect '
                       '<strong>' + patch_name + '</strong>, then correct the row by hand.</p>')
                    + '</div>'
                ),
                "emailMessage/Importance": "High",
            },
            "host": {"apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
                     "connectionName": "shared_office365",
                     "operationId": "SendEmailV2"},
            "authentication": "@parameters('$authentication')",
        },
    }


_patch_alert_names = []
for _patch_name, (_what, _swallowed) in _PATCH_ALERTS.items():
    _owner = _find_owner_actions(actions, _patch_name)
    if _owner is None:
        continue
    _alert_name = "Notify_" + _patch_name.replace("Patch_", "") + "_failed"
    # POP before re-adding, same as the inbox-tidy section does. The zip is both the output
    # AND the next build's structural base, and the year/month-separator section re-appends
    # its own actions to this same map on every run. Plain assignment would keep an existing
    # alert at its old position while the separators moved past it, so build N and build N+1
    # produced byte-different (though semantically identical) definitions. Popping first
    # makes the alert always land last, so repeated builds are byte-stable.
    _owner.pop(_alert_name, None)
    _owner[_alert_name] = _make_patch_alert(_patch_name, _what, _swallowed)
    _patch_alert_names.append(_alert_name)

# ── 16b. INTAKE-WRITE FAILURE ALERTS (Add_row / Create_file) ─────────────────
# The 7 alerts above cover every PATCH of an EXISTING row. The two writes that CREATE a
# candidate had no watcher at all, which is a strictly worse failure: a patch that fails
# leaves a row that is merely stale, but an Add_row that fails leaves NO ROW ANYWHERE.
#
# Why it was silent: Mark_as_read/Move_to_processed hang off HasResume, not off Add_row,
# so inbox cleanup is on a PARALLEL branch and files the mail as handled regardless.
# Send_acknowledgment is correctly gated on Add_row=Succeeded (test B: a failed Excel write
# must never acknowledge), so the applicant gets nothing either. Net effect before this
# section existed: no row, no reply, mail archived as done, run reports Succeeded, silence.
# The 2026-08-24 audit found 11 candidates lost this way, clustered 2-6 June 2026.
#
# Deliberately NOT re-gating the cleanup on Add_row: making Move_to_processed conditional
# would re-introduce the poisoned-message queue block that fetch_only_unread=false exists
# to prevent, and re-marking the message unread afterwards races the mutable Outlook
# message ID (the same hazard that disables retries on MoveV2/MarkAsRead). So the alert is
# the signal, exactly as it is for the 7 swallowed patch failures.
_WRITE_ALERTS = {
    "Add_row": (
        "the candidate row was never created",
        "This applicant has <strong>no row on either sheet</strong> and received "
        "<strong>no acknowledgment</strong> - Send_acknowledgment is gated on Add_row "
        "succeeding, so it was skipped. Their resume may still have been saved to "
        "SharePoint, which leaves an orphan file with nothing pointing at it.",
    ),
    "Create_file": (
        "the resume file was never saved",
        "Add_row still runs when the save fails and always writes Has Resume = 'Yes', so "
        "the row now <strong>claims a resume that does not exist</strong> in SharePoint. "
        "Phase 2 cannot score it and will not be able to read the file.",
    ),
    # The other TWO resume writes. Create_file above only covers the NEW-applicant path;
    # a resend (duplicate) and a ref-quoted update each have their own CreateFile, and both
    # were silent for exactly the same reason the patch failures were: nothing runs after
    # them inside their Foreach, so the Foreach fails - and the sibling that follows the
    # Foreach (IsUnderReplyCap_dup / IsUnderReplyCap_update) lists Failed in its runAfter,
    # which ABSORBS it. The run then continues normally: the reply is sent, Application
    # Updates increments, Last Updated Date bumps, Status resets to 'New Email Received',
    # the mail is marked read and archived, and the run reports Succeeded. The candidate is
    # told their new resume arrived while the old file is still the one on disk - and P2
    # re-scores the STALE document. Found 2026-08-26 by the pre-import branch audit.
    "Create_dup_update_file": (
        "a repeat applicant's NEW resume was never saved",
        "This was a duplicate/resend, so the row was patched as though the new resume "
        "landed - <strong>Application Updates was incremented and Status reset to 'New "
        "Email Received'</strong> - but the file on disk is still the PREVIOUS version. "
        "Phase 2 will re-score the stale document, and the applicant has already been told "
        "their update was received.",
    ),
    "Create_update_file": (
        "a ref-quoted update's NEW resume was never saved",
        "The applicant quoted their Application ID, so the row was patched as an update - "
        "<strong>Application Updates was incremented and Status reset to 'New Email "
        "Received'</strong> - but the file on disk is still the PREVIOUS version. Phase 2 "
        "will re-score the stale document, and the applicant has already been told their "
        "update was received.",
    ),
}


def _make_write_alert(action_name: str, what: str, consequence: str) -> dict:
    """Admin alert for a failed INTAKE write. Sibling of `action_name`."""
    return {
        "type": "OpenApiConnection",
        "runAfter": {action_name: ["Failed", "TimedOut"]},
        "runtimeConfiguration": {"retryPolicy": {"type": "exponential", "count": 2,
                                                 "interval": "PT10S"}},
        "inputs": {
            "parameters": {
                "emailMessage/To": email["admin_email"],
                "emailMessage/Subject":
                    "[Hiring Auto-Reply] APPLICANT LOST - %s (Ref %s)" % (what, _ref_any),
                "emailMessage/Body": (
                    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;'
                    'color:#222;line-height:1.6;">'
                    '<p><strong>Hiring Auto-Reply: an applicant was not recorded.</strong></p>'
                    '<p><strong>' + action_name + '</strong> failed, so ' + what + '.</p>'
                    '<p>' + consequence + '</p>'
                    '<ul>'
                    '<li>Reference: <strong>' + _ref_any + '</strong></li>'
                    '<li>From: @{coalesce(outputs(\'CurrentEmail\')?[\'from\'],\'unknown\')}</li>'
                    '<li>Subject: @{coalesce(outputs(\'CurrentEmail\')?[\'subject\'],\'(none)\')}</li>'
                    '<li>Received: @{coalesce(outputs(\'CurrentEmail\')?[\'receivedDateTime\'],\'(unknown)\')}</li>'
                    '</ul>'
                    '<p>The run reports <strong>Succeeded</strong> because inbox cleanup is '
                    'deliberately unconditional, and the message has been marked read and '
                    'moved to Archive as though it were handled - <strong>this alert is the '
                    'only signal</strong>.</p>'
                    '<p><strong>To recover:</strong> find the message in Archive by the From '
                    'and Received values above and move it back to the Inbox as unread, so '
                    'the next poll re-runs the full intake for it. If applicant mail is '
                    'currently ON, that will also re-send their acknowledgment - which is '
                    'correct here, because they never received one.</p>'
                    '</div>'
                ),
                "emailMessage/Importance": "High",
            },
            "host": {"apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
                     "connectionName": "shared_office365",
                     "operationId": "SendEmailV2"},
            "authentication": "@parameters('$authentication')",
        },
    }


_write_alert_names = []
for _write_name, (_what, _consequence) in _WRITE_ALERTS.items():
    _owner = _find_owner_actions(actions, _write_name)
    if _owner is None:
        continue
    _alert_name = "Notify_" + _write_name.lower() + "_failed"
    # Pop before re-adding, for the same byte-stability reason as the patch alerts.
    _owner.pop(_alert_name, None)
    _owner[_alert_name] = _make_write_alert(_write_name, _what, _consequence)
    _write_alert_names.append(_alert_name)

# ── 17. enable PA-level failure alert as a backup (Notify_failure is primary) ─
defn["properties"]["flowFailureAlertSubscribed"] = True

# ── 18. OUTBOUND MAIL CONTROLS ──────────────────────────────────────────────
# Replaces the 6 applicant-facing Send_* actions with no-op Compose steps when applicant
# mail is disabled. Notify_failure is controlled separately so a silent live intake can
# still alert the admin on real failures. The no-op steps retain the same runAfter wiring,
# so downstream tidy/patch actions behave exactly as they do with applicant mail enabled.
# Must run LAST, after every other step has finished building/wiring the flow.
_SUPPRESSED_SENDS = ["Send_acknowledgment", "Send_duplicate_notice", "Send_CV_request",
                    "Send_wrong_format", "Send_update_ack", "Send_noted_reply"]
_SUPPRESSED_STAMP = ("@{concat('Suppressed (applicant email disabled) ', "
                     "formatDateTime(%s,'yyyy-MM-dd HH:mm'))}" % _NOW_LOCAL)
if not send_applicant_emails:
    for _name in _SUPPRESSED_SENDS:
        _act = _deep_find(actions, _name)
        if _act is None:
            continue
        _run_after = _act.get("runAfter", {})
        _act.clear()
        _act["type"] = "Compose"
        _act["runAfter"] = _run_after
        _act["inputs"] = f"Applicant email disabled: suppressed {_name}; intake processing continues"
    # A Compose action always reports status=Succeeded (it can't fail the way a real send
    # can), so the existing conditional 'Mail Sent' stamps would otherwise write a
    # real-looking 'Sent <timestamp>' even though nothing was sent. Force a visibly-distinct
    # marker instead, on every patch action that stamps one. ('Decline Sent' is P2's own
    # column, stamped by the Python worker, not by this flow - suppressed separately via
    # HIRING_GEO_REJECT_EMAIL=false on that command.)
    for _patch_name in ("Patch_mail_sent", "Patch_dup_attempts", "Patch_update_attempts",
                        "Patch_followup_attempts"):
        _patch = _deep_find(actions, _patch_name)
        if _patch is not None and "Mail Sent" in _patch.get("inputs", {}).get("parameters", {}).get("item", {}):
            _patch["inputs"]["parameters"]["item"]["Mail Sent"] = _SUPPRESSED_STAMP
    print("  [SILENT INTAKE] Applicant emails SUPPRESSED; processing remains active.")

if not send_admin_failure_alerts:
    # Notify_failure plus every patch-failure alert from section 16d. Same treatment as the
    # applicant sends: swapped for a no-op Compose that keeps the runAfter wiring intact, so
    # the surrounding graph behaves identically whether alerts are on or off.
    for _alert in ["Notify_failure"] + _patch_alert_names + _write_alert_names:
        _act = _deep_find(actions, _alert)
        if _act is None:
            continue
        _run_after = _act.get("runAfter", {})
        _act.clear()
        _act["type"] = "Compose"
        _act["runAfter"] = _run_after
        _act["inputs"] = f"Admin failure alert disabled: {_alert} suppressed"
    print(f"  [WARNING] Admin failure alerts SUPPRESSED "
          f"({1 + len(_patch_alert_names) + len(_write_alert_names)} alert actions).")

# -- 19. unread-Inbox production wrapper ------------------------------------
# Every expression in the historical processing graph reads one message directly from
# triggerOutputs(). The scheduled poller puts that message in CurrentEmail instead. A
# mechanical source rewrite is safer than maintaining dozens of hand-edited expressions
# and is asserted by the tests (zero triggerOutputs message references remain).
def _use_current_email(node):
    if isinstance(node, dict):
        return {k: _use_current_email(v) for k, v in node.items()}
    if isinstance(node, list):
        return [_use_current_email(v) for v in node]
    if isinstance(node, str):
        return node.replace(
            "triggerOutputs()?['body/", "outputs('CurrentEmail')?['"
        )
    return node


actions = _use_current_email(actions)
actions["CONFIG"]["runAfter"] = {"CurrentEmail": ["Succeeded"]}

_get_unread = {
    "type": "OpenApiConnection",
    "runAfter": {},
    "inputs": {
        "parameters": {
            "folderPath": "Inbox",
            "fetchOnlyWithAttachment": False,
            "fetchOnlyUnread": fetch_only_unread,
            "mailboxAddress": email["trigger_mailbox"],
            "includeAttachments": True,
            "top": unread_per_run,
        },
        "host": {
            "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
            "connectionName": "shared_office365",
            "operationId": "GetEmailsV3",
        },
        "authentication": "@parameters('$authentication')",
    },
}

_current_email = {
    "type": "Compose",
    "runAfter": {"Has_unread_email": ["Succeeded"]},
    "inputs": "@first(body('Get_unread_emails')?['value'])",
}
_has_unread = {
    "type": "If",
    "runAfter": {"Get_unread_emails": ["Succeeded"]},
    "expression": {
        "and": [
            {
                "greater": [
                    "@length(coalesce(body('Get_unread_emails')?['value'], json('[]')))",
                    0,
                ]
            }
        ]
    },
    "actions": {
        "Unread_email_available": {
            "type": "Compose",
            "runAfter": {},
            "inputs": "Unread Inbox message found; continue with the flat processing graph",
        }
    },
    "else": {
        "actions": {
            "Stop_no_unread_email": {
                "type": "Terminate",
                "runAfter": {},
                "inputs": {"runStatus": "Succeeded"},
            }
        }
    },
}

if send_admin_failure_alerts:
    _poll_failure = {
        "type": "OpenApiConnection",
        "runAfter": {"Get_unread_emails": ["Failed", "TimedOut"]},
        # Retries ARE wanted here (changed 2026-08-03). The blanket 'no retries' rule exists
        # for MoveV2/MarkAsRead, where a retry races the mutable Outlook message ID; a send
        # has no such hazard. With applicant mail suppressed, the admin alerts are the only
        # live sends in the flow, so a single transient failure would otherwise lose the
        # only signal that the poll broke. Worst case is a duplicate admin email.
        "runtimeConfiguration": {"retryPolicy": {"type": "exponential", "count": 2,
                                                 "interval": "PT10S"}},
        "inputs": {
            "parameters": {
                "emailMessage/To": email["admin_email"],
                "emailMessage/Subject": "[Hiring Auto-Reply] Unread Inbox poll failed",
                "emailMessage/Body": (
                    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;">'
                    '<p><strong>P1 could not read the shared Inbox.</strong></p>'
                    '<p>No applicant email was sent. The unread messages remain in Inbox. '
                    'Open the flow run history and inspect Get_unread_emails.</p></div>'
                ),
                "emailMessage/Importance": "High",
            },
            "host": {
                "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
                "connectionName": "shared_office365",
                "operationId": "SendEmailV2",
            },
            "authentication": "@parameters('$authentication')",
        },
    }
else:
    _poll_failure = {
        "type": "Compose",
        "runAfter": {"Get_unread_emails": ["Failed", "TimedOut"]},
        "inputs": "Admin failure alert disabled: unread Inbox poll failure alert suppressed",
    }

wd["actions"] = {
    "Get_unread_emails": _get_unread,
    "Has_unread_email": _has_unread,
    "CurrentEmail": _current_email,
    **actions,
    "Notify_poll_failure": _poll_failure,
}

# ── validate + repackage (same GUID/maps so it Updates in place) ─────────────
def _action_depths(action_map: dict, level: int = 0):
    """Yield (name, level) using Power Automate's control-action nesting count."""
    for action_name, action in action_map.items():
        yield action_name, level
        child_actions = action.get("actions")
        if isinstance(child_actions, dict):
            yield from _action_depths(child_actions, level + 1)
        else_actions = action.get("else", {}).get("actions")
        if isinstance(else_actions, dict):
            yield from _action_depths(else_actions, level + 1)
        for case in action.get("cases", {}).values():
            case_actions = case.get("actions")
            if isinstance(case_actions, dict):
                yield from _action_depths(case_actions, level + 1)
        default_actions = action.get("default", {}).get("actions")
        if isinstance(default_actions, dict):
            yield from _action_depths(default_actions, level + 1)


from reliability_flow import apply_reliability
apply_reliability(wd, cfg)

# Correctness fixes for the four findings of the 2026-09-08 package review. Applied last,
# over the fully assembled flow, so the whole change set lives in one reviewable file and
# the depth check below validates the result rather than the pre-fix shape.
from integrity_fixes import apply_integrity_fixes
_integrity = apply_integrity_fixes(wd, tidy_cfg.get("failure_destination", "Archive"),
                                   email["trigger_mailbox"])

_depths = list(_action_depths(wd["actions"]))
_max_depth = max(level for _, level in _depths)
if _max_depth > 8:
    _too_deep = [name for name, level in _depths if level == _max_depth]
    raise ValueError(
        f"Power Automate action nesting depth {_max_depth} exceeds limit 8: {_too_deep}"
    )
# The flow already sits AT the limit, so the next control action added inside the deepest
# branch fails the build rather than degrading gracefully. Say so at build time, naming the
# branch, instead of leaving the next person to decode a raise from a seemingly unrelated
# edit. Warn only - being at 8 is legal and is the current, shipped state.
if _max_depth == 8:
    _at_limit = sorted({name for name, level in _depths if level == 8})
    print(f"  [!] NESTING AT LIMIT: depth 8/8 (PA max). Zero headroom in: "
          f"{', '.join(_at_limit[:4])}{' ...' if len(_at_limit) > 4 else ''}")
    print("      Adding another If/Foreach/Scope inside that branch WILL fail this build. "
          "Flatten a level there first, or attach new work to a shallower branch.")

final = json.dumps(defn, indent=2, ensure_ascii=False)
json.loads(final)  # parse check

with zipfile.ZipFile(ZIP, "w", compression=zipfile.ZIP_DEFLATED) as z:
    z.writestr("manifest.json", root_mf)
    # assetPaths names the per-flow folder; it must follow the live GUID as well.
    z.writestr("Microsoft.Flow/flows/manifest.json",
               flow_mf.replace(_BASE_GUID, GUID))
    z.writestr(BASE + "definition.json", final)
    z.writestr(BASE + "apisMap.json", apis_map)
    z.writestr(BASE + "connectionsMap.json", conns_map)

# Sync the readable copies so they stop drifting from the deployable artifact.
(HERE / "definition.json").write_text(final, encoding="utf-8")
(HERE / "manifest.json").write_text(
    json.dumps(json.loads(root_mf), indent=2, ensure_ascii=False), encoding="utf-8")
(HERE / "apisMap.json").write_text(
    json.dumps(json.loads(apis_map), indent=2, ensure_ascii=False), encoding="utf-8")

print("[OK] flow rebuilt from flow_config.json -> ZIP + definition.json synced")
print(f"  flow / package name : {flow_name}")
print(f"  trigger mailbox     : {email['trigger_mailbox']}")
print(f"  poll interval (min) : {interval}")
print(f"  trigger mode        : scheduled Inbox poll ({unread_per_run} message per run, "
      f"{'UNREAD ONLY' if fetch_only_unread else 'read OR unread - Inbox is the queue'})")
print(f"  max action nesting  : {_max_depth} (Power Automate limit: 8)")
print(f"  admin alert -> to   : {email['admin_email']}")
print(f"  duplicate window    : {days} day(s)")
print(f"  caps (shared cnt)   : duplicate {dup_cap} | update {update_cap} | follow-up {followup_cap} "
      f"(no-CV: reply only, no row) | reply_cap {reply_cap} (contacts past this get no email, just counted)")
print(f"  timezone            : {_TZ_ID} - all derived dates are LOCAL wall-clock" if _TZ_ID != 'UTC' else " - dates are RAW UTC (pre-2026-08-27 behaviour)")
print(f"  appref format       : APP-{appref['date_format']}-{_time_fmt}-<{_suffix_len} chars of the "
      f"message id, offset -{_suffix_len + _tail_skip}> - DETERMINISTIC: a retry of the same email mints the same ref")
print(f"  bad_senders         : {len(actions['BadSenders']['inputs'])}")
print(f"  bad_subjects        : {len(bad_subjects)} (config; self-loop guard is sender-based)")
print(f"  spam/safety gate    : {len(actions['SpamPhrases']['inputs'])} "
      f"({len(filt['spam_phrases'])} scam + {len(filt.get('offensive_phrases', []))} offensive "
      f"+ {len(filt.get('malware_phrases', []))} malware/exe + {len(filt.get('foreign_scam_phrases', []))} non-EN "
      f"+ {len(filt.get('vendor_solicitation_phrases', []))} vendor-pitch); "
      f"scans subject+body+attachment names")
print(f"  app_keywords        : {len(kws)} (OR clauses: {len(kw['expression']['and'][1]['or'])}; E3 gated on no-attachment)")
print(f"  excel table         : {ex['table']}")
print(f"  resume subfolders   : {'dated Year/Month' if sp.get('dated_resume_subfolders') else 'flat'}")
print(f"  inbox tidy          : {'ON (%d actions - legit->%s read, spam->%s unread, failures->%s unread)' % (_tidy_total, tidy_cfg.get('destination','n/a'), tidy_cfg.get('spam_destination','n/a'), tidy_cfg.get('failure_destination','n/a')) if _tidy_total else 'OFF (mail stays unread in Inbox)'}")
print(f"  row extras          : Subject + body preview (~255 chars, bodyPreview) columns + improved Full Name guess")
print(f"  workbook self-heal  : OFF (non-destructive) - never auto-creates/overwrites; missing workbook -> Notify_failure alert")
print(f"  admin alerts        : {'ON' if send_admin_failure_alerts else 'OFF'} - "
      f"Notify_failure + Notify_poll_failure + {len(_patch_alert_names)} patch-failure alerts "
      f"+ {len(_write_alert_names)} intake-write alerts ({', '.join(_WRITE_ALERTS)}) "
      f"(retry: exponential x2)")
print(f"  resume folder create: ON (Ensure_*_resume_folder, safe no-op if exists) - no delays")
print(f"  resume overwrite    : update/resend always saves under the legacy <Name>_<AppID> shape "
      f"(true in-place overwrite pre-scoring; post-scoring, P2 reconciles it against the "
      f"renamed file - see sharepoint_scoring.py _download_resume_text)")
print(f"  integrity fixes     : allow-list exact match | retryPolicy also under inputs on "
      f"{len(_integrity['retry_policy'])} actions | success copy gated on "
      f"{len(_integrity['save_gated'])} resume saves | latest-row pick via "
      f"{', '.join(_integrity['row_pick'])}")
print(f"  year separator      : {'ON - one blank row before each new year; first-ever applicant uses the header spacer' if year_sep.get('enabled', True) else 'OFF'}")
print(f"  month separator     : {'ON - one blank row before the first new-applicant row of each month (suppressed at a year boundary so only ONE blank row lands)' if month_sep.get('enabled', True) else 'OFF'}")
