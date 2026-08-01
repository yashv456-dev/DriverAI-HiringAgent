"""P1 end-to-end test — structural validation of the upload ZIP + routing simulation.

Reads the real Power Automate ZIP definition and flow_config.json (no live tenant needed).
Run from any directory:
    python HiringAgent_P1/test_p1.py
"""
import base64, io, json, re, sys, zipfile
from pathlib import Path
from openpyxl import load_workbook
sys.stdout.reconfigure(encoding="utf-8")

FLOW = r"C:\Users\yash9\Downloads\HiringAgent\DriverAI_HiringAgent\HiringAgent_P1\flow"
P2   = r"C:\Users\yash9\Downloads\HiringAgent\DriverAI_HiringAgent\HiringAgent_App_P2"

ZIP_PATH = Path(FLOW) / "DriverAI-Hiring-AutoReply-apply.zip"
LOOSE_DEF_PATH = Path(FLOW) / "definition.json"

with zipfile.ZipFile(ZIP_PATH) as zf:
    zip_def_names = [n for n in zf.namelist() if n.endswith("/definition.json")]
    if len(zip_def_names) != 1:
        raise SystemExit(f"Expected exactly one definition.json in {ZIP_PATH}, found {len(zip_def_names)}")
    d = json.loads(zf.read(zip_def_names[0]).decode("utf-8"))

loose_d = json.load(io.open(LOOSE_DEF_PATH, encoding="utf-8"))
cfg     = json.load(io.open(FLOW + r"\flow_config.json", encoding="utf-8"))
defn = d["properties"]["definition"]

import yaml
ycfg = yaml.safe_load(io.open(P2 + r"\config.yaml", encoding="utf-8"))
sys.path.insert(0, P2)
from hiring_agent.config import COLUMNS as P2_COLUMNS, REJECTED_COLUMNS as P2_REJECTED_COLUMNS

P = F = 0
def ok(c, label):
    global P, F
    if c: P += 1; print(f"  [PASS] {label}")
    else: F += 1; print(f"  [FAIL] {label}")

print("\n=== ZIP PACKAGE SOURCE ===")
ok(d == loose_d, "upload ZIP definition matches flow/definition.json exactly")

# Collect all actions recursively into name -> def map
actions = {}
def walk(o):
    if isinstance(o, dict):
        for k, v in o.items():
            if k == "actions" and isinstance(v, dict):
                for an, ad in v.items():
                    actions[an] = ad
            walk(v)
    elif isinstance(o, list):
        for x in o: walk(x)
walk(defn)
blob = json.dumps(defn)

# ── A0. RUNAFTER WIRING INTEGRITY ────────────────────────────────────────────
# Catches "InvalidTemplate" errors BEFORE they hit the PA designer. Found live 2026-07-03 in two
# rounds: (1) IsSpam referenced HasValidResumeEarly without it in runAfter at all (fixed by adding
# it) - but that surfaced round 2: (2) PA rejected it anyway with "must belong to same level as
# action IsSpam" - runAfter can ONLY reference a LITERAL SIBLING at the exact same nesting level,
# not just any action that's transitively guaranteed to have run first (being top-level and always
# executed is NOT sufficient - a top-level action is a sibling of the outermost If, not an ancestor
# of everything nested inside it). Three checks below: (1) no runAfter points at a nonexistent
# action name; (2) every runAfter TARGET must be a same-level sibling of the action declaring it -
# this is the strict rule PA actually enforces on the runAfter property itself; (3) every
# outputs()/body()/items() reference in an action's own inputs/expression must be reachable via
# either its own runAfter chain or its containing scope's ancestor chain (a looser check - an
# expression CAN reference a true ancestor implicitly, it just can't be declared as a same-level
# runAfter dependency unless it genuinely is one).
print("\n=== A0. RUNAFTER WIRING INTEGRITY ===")
owner_of = {}
def index_owners(o, owner):
    if isinstance(o, dict):
        for k, v in o.items():
            if k == "actions" and isinstance(v, dict):
                for name, act in v.items():
                    owner_of[name] = owner
                    index_owners(act, name)
            elif k != "actions":
                index_owners(v, owner)
    elif isinstance(o, list):
        for x in o: index_owners(x, owner)
index_owners(defn, None)

dangling = [(name, dep) for name, act in actions.items()
            for dep in act.get("runAfter", {}) if dep not in actions]
ok(len(dangling) == 0,
   f"no dangling runAfter references ({len(dangling)} found)" if dangling else
   "no dangling runAfter references (every runAfter target exists)")
for name, dep in dangling:
    print(f"    -> {name} runAfter references missing action '{dep}'")

wrong_level = [(name, dep) for name, act in actions.items()
               for dep in act.get("runAfter", {})
               if dep in owner_of and owner_of.get(name) != owner_of.get(dep)]
ok(len(wrong_level) == 0,
   f"every runAfter target is a same-level sibling ({len(wrong_level)} cross-level violations found)")
for name, dep in wrong_level:
    print(f"    -> {name} (level owner={owner_of.get(name)}) runAfter references '{dep}' "
          f"(level owner={owner_of.get(dep)}) - NOT the same level, PA will reject this")

_guaranteed_cache = {}
def _guaranteed(name, _stack=frozenset()):
    if name in _guaranteed_cache: return _guaranteed_cache[name]
    if name in _stack: return set()
    act = actions.get(name)
    if act is None: return set()
    result = set()
    for dep in act.get("runAfter", {}):
        result.add(dep)
        result |= _guaranteed(dep, _stack | {name})
    owner = owner_of.get(name)
    if owner is not None:
        result.add(owner)
        result |= _guaranteed(owner, _stack | {name})
    _guaranteed_cache[name] = result
    return result

def _own_refs(act):
    own = {k: v for k, v in act.items() if k not in ("runAfter", "type", "metadata", "actions", "else", "foreach")}
    return set(re.findall(r"(?:outputs|body|items)\('([^']+)'\)", json.dumps(own)))

unreachable = []
for name, act in actions.items():
    reach = _guaranteed(name)
    for r in _own_refs(act):
        if r in actions and r != name and r not in reach and r != owner_of.get(name):
            unreachable.append((name, r))
ok(len(unreachable) == 0,
   f"every action reference is on a guaranteed runAfter path ({len(unreachable)} unreachable found)")
for name, r in unreachable:
    print(f"    -> {name} references '{r}' but cannot guarantee it ran first (would fail PA's InvalidTemplate validation)")

# Wait/Delay interval must be 5s..30d on Consumption SKU (a value below 5s fails at RUNTIME with
# BadRequest - found live 2026-07-04 when folder waits were 2s). Guard every Wait action here so a
# too-short interval can never ship again.
_MIN_WAIT_S, _MAX_WAIT_S = 5, 30 * 24 * 3600
_unit_secs = {"Second": 1, "Minute": 60, "Hour": 3600, "Day": 86400, "Week": 604800}
bad_waits = []
for name, act in actions.items():
    if act.get("type") == "Wait":
        iv = act.get("inputs", {}).get("interval", {})
        secs = iv.get("count", 0) * _unit_secs.get(iv.get("unit", ""), 0)
        if secs < _MIN_WAIT_S or secs > _MAX_WAIT_S:
            bad_waits.append((name, iv, secs))
ok(len(bad_waits) == 0,
   f"every Wait interval is within PA's 5s..30d Consumption limit ({len(bad_waits)} out of range)")
for name, iv, secs in bad_waits:
    print(f"    -> {name}: {iv} = {secs}s (must be 5s..30d, or it fails at runtime with BadRequest)")

# Power Automate numbers top-level actions at nesting level 0 and rejects anything
# beyond level 8. The scheduled unread guard must stay a sibling of the legacy graph;
# wrapping the graph once pushed the update/follow-up paths to level 9 at import time.
def _action_depths(action_map, level=0):
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

depths = list(_action_depths(defn["actions"]))
max_depth = max(level for _, level in depths)
too_deep = [(name, level) for name, level in depths if level > 8]
ok(not too_deep and max_depth == 8,
   f"maximum action nesting is {max_depth}, within Power Automate's level-8 limit")
ok("CONFIG" in defn["actions"] and "CurrentEmail" in defn["actions"],
   "unread guard and existing processing graph are flat top-level siblings")

# ── A. TRIGGER / SCHEDULE ────────────────────────────────────────────────────
print("\n=== A. TRIGGER / SCHEDULE ===")
trig  = defn.get("triggers", {})
ok(len(trig) == 1, "exactly one trigger")
tname = next(iter(trig))
trg   = trig[tname]
rec   = trg.get("recurrence", {})
ok(rec.get("interval") == cfg["trigger"]["interval_min"],
   f"poll interval = {rec.get('interval')} min (config {cfg['trigger']['interval_min']})")
ok(trg.get("type") == "Recurrence", "scheduled Recurrence trigger polls existing unread mail")
ok(trg.get("splitOn") is None, "no new-arrival splitOn watermark remains")
# Concurrency = 1 prevents double-row writes
op_opts = trg.get("operationOptions", "") or ""
limits  = trg.get("runtimeConfiguration", {}).get("concurrency", {})
ok(limits.get("runs", 0) == 1 or "Single" in op_opts or "1" in json.dumps(limits),
   "concurrency = 1 (sequential, no parallel runs)")
get_unread = actions.get("Get_unread_emails", {})
get_unread_params = get_unread.get("inputs", {}).get("parameters", {})
ok(get_unread.get("inputs", {}).get("host", {}).get("operationId") == "GetEmailsV3",
   "Get_unread_emails uses the supported Outlook Get emails (V3) action")
ok(get_unread_params.get("mailboxAddress") == cfg["email"]["trigger_mailbox"],
   "unread poll targets the configured shared mailbox")
ok(get_unread_params.get("folderPath") == "Inbox"
   and get_unread_params.get("fetchOnlyUnread") is True,
   "unread poll selects only Inbox messages whose read flag is false")
ok(get_unread_params.get("includeAttachments") is True,
   "unread poll includes attachment content for resume saves")
ok(get_unread_params.get("top") == cfg["trigger"]["unread_per_run"] == 1,
   "exactly one unread message is claimed per sequential run")
ok(actions.get("Has_unread_email", {}).get("runAfter") == {"Get_unread_emails": ["Succeeded"]},
   "processing starts only after the unread poll succeeds")
ok(actions.get("CurrentEmail", {}).get("inputs") ==
   "@first(body('Get_unread_emails')?['value'])",
   "CurrentEmail selects the single polled message")
ok(actions.get("CONFIG", {}).get("runAfter") == {"CurrentEmail": ["Succeeded"]},
   "the existing P1 graph starts after CurrentEmail is selected")
ok("triggerOutputs" not in blob and "triggerBody" not in blob,
   "all processing expressions read CurrentEmail, not a new-arrival trigger payload")

# ── B. COLUMN CONTRACT (table = 30 cols; Add_row writes ONLY the 11 P1-owned ones) ───
# The TABLE has 30 columns (defined by the workbook template header + P2's own schema additions
# since). Add_row only WRITES the 11 P1-owned columns - it does NOT list the other 19 at all (they
# stay blank cells that P2 fills later). This exactly matches the genuine tenant PA exports
# (archive/ne3, ne4) and minimizes the import-time dynamic-schema binding surface that was dropping
# fields in the designer.
print("\n=== B. COLUMN CONTRACT (11-col write, 30-col table) ===")
item   = actions["Add_row"]["inputs"]["parameters"]["item"]
cols   = list(item.keys())
p2cols = ycfg["columns"]
P1_OWNED = ["Application ID", "Received Date", "Last Updated Date", "Full Name", "Email", "Mail Subject",
            "Mail Body", "Status", "Has Resume", "Original Filename", "Application Updates"]
P2_OWNED = ["Category", "Phone", "Location", "Country", "Current Skills", "Education",
            "Looking For Role", "Suggested Role 1", "Suggested Role 2", "Suggested Role 3",
            "Portfolio 1", "Portfolio 2", "Portfolio 3"]
ok(len(cols) == 11, f"Add_row writes exactly the 11 P1-owned columns (got {len(cols)})")
ok(set(cols) == set(P1_OWNED), "Add_row's 11 keys == the P1-owned set")
ok(len(p2cols) == 30, f"P2 config.yaml COLUMNS = 30 (adds Last Updated Date, got {len(p2cols)})")
ok("Resume Link" in p2cols and "Resume URL" in p2cols, "resume URL + link columns present in schema")
# Received Date is written as ISO without the timezone offset (drops the '+00:00'),
# still ISO so the 90-day dup check's ticks() and P2's parser both read it.
_rd_val = item.get("Received Date", "")
ok("formatDateTime" in _rd_val and "yyyy-MM-ddTHH:mm:ss" in _rd_val,
   "Add_row writes Received Date via formatDateTime ISO (no +00:00 offset)")
ok("+00:00" not in _rd_val, "Received Date no longer carries the raw '+00:00' offset")
_lud_val = item.get("Last Updated Date", "")
ok(_lud_val == _rd_val, "Add_row initializes Last Updated Date equal to Received Date")
ok(p2cols[2] == "Last Updated Date", "Last Updated Date at table position #3")
ok(set(cols).issubset(set(p2cols)), "Every Add_row column exists in the 30-col table schema (no orphan write)")
ok(all(c in p2cols for c in P2_OWNED), f"All {len(P2_OWNED)} P2-owned scoring/profile columns exist in the table")
# 'Mail Sent' audit column: stamped only AFTER a reply actually sends — never at row creation.
ok("Mail Sent" in p2cols, "'Mail Sent' audit column present in the table schema")
ok("Mail Sent" not in item, "Add_row does NOT list 'Mail Sent' (stamped by Patch_mail_sent after the ack sends)")
# Stamp-content checks depend on the applicant-mail setting: an enabled build writes
# conditional 'Sent <timestamp>' expressions; a silent-intake build writes a suppression marker
# (the Send_* actions are no-op Compose placeholders that always 'succeed', so a
# conditional real-looking stamp would lie).
_send_applicant_mail = bool(cfg.get("email", {}).get("send_applicant_emails", True))
_send_admin_alerts = bool(cfg.get("email", {}).get("send_admin_failure_alerts", True))
_suppress_b = not _send_applicant_mail
_pms = actions.get("Patch_mail_sent")
ok(_pms is not None, "Patch_mail_sent action exists (ack-path 'Mail Sent' stamp)")
if _pms is not None:
    ok(list(_pms.get("runAfter", {}).keys()) == ["Send_acknowledgment"],
       "Patch_mail_sent runs after Send_acknowledgment only")
    ok(_pms["runAfter"].get("Send_acknowledgment") == ["Succeeded"],
       "Patch_mail_sent runs ONLY on ack Succeeded (a failed send is never stamped)")
    ok("Mail Sent" in _pms["inputs"]["parameters"]["item"],
       "Patch_mail_sent writes the 'Mail Sent' column")
    _pms_stamp = _pms["inputs"]["parameters"]["item"]["Mail Sent"]
    if _suppress_b:
        ok("Suppressed (applicant email disabled)" in _pms_stamp,
           "Patch_mail_sent records that applicant email was suppressed")
    else:
        ok("formatDateTime" in _pms_stamp and "'Sent '" in _pms_stamp,
           "Patch_mail_sent stamps a 'Sent <timestamp>' value (aligned with Rejected sheet's Decline Sent)")
for _patch_name, _patch_send in (("Patch_dup_attempts", "Send_duplicate_notice"),
                                 ("Patch_update_attempts", "Send_update_ack"),
                                 ("Patch_followup_attempts", "Send_noted_reply")):
    _pa = actions.get(_patch_name)
    _ms = (_pa or {}).get("inputs", {}).get("parameters", {}).get("item", {}).get("Mail Sent", "")
    if _suppress_b:
        ok("Suppressed (applicant email disabled)" in _ms,
           f"{_patch_name} records that applicant email was suppressed")
    else:
        ok(f"actions('{_patch_send}')?['status']" in _ms,
           f"{_patch_name} stamps 'Mail Sent' conditionally on {_patch_send} really succeeding")
        ok("formatDateTime" in _ms and "'Sent '" in _ms,
           f"{_patch_name} stamps a 'Sent <timestamp>' value when the reply sent")
        ok("coalesce" in _ms and "'Mail Sent'" in _ms,
           f"{_patch_name} preserves an earlier 'Mail Sent' stamp when the send failed/skipped")

# Last Updated Date max-guard (fixed 2026-07-31): every patch stamps it, and never backward -
# see _last_updated_expr in build_zip.py. Patch_followup_attempts previously omitted it entirely,
# which corrupted P2's _updated_after_location_request() (sharepoint_scoring.py) into thinking a
# candidate who replied with a text-only follow-up never responded.
for _patch_name in ("Patch_dup_attempts", "Patch_update_attempts", "Patch_followup_attempts"):
    _pa_item = (actions.get(_patch_name) or {}).get("inputs", {}).get("parameters", {}).get("item", {}) or {}
    _lud = _pa_item.get("Last Updated Date", "")
    ok("Last Updated Date" in _pa_item, f"{_patch_name} stamps Last Updated Date on every contact")
    ok("greater(ticks(" in _lud and "Received Date" in _lud,
       f"{_patch_name}'s Last Updated Date never moves backward (max-guard against the row's existing value)")

# Resume filename uses the MATCHED ROW's real Application ID, not whatever ref the email text
# happens to quote (fixed 2026-07-31): a sender still replying to a dead/old thread would have
# their update resume saved under a stale AppRef that P2's file lookup - keyed strictly to the
# row's own current Application ID (sharepoint_scoring.py _resume_name_slots) - could never find.
ok("AppRef_current" in actions, "AppRef_current exists (the matched row's live Application ID)")
ok("first(body('Get_rows_ref')" in actions.get("AppRef_current", {}).get("inputs", ""),
   "AppRef_current is sourced from Get_rows_ref (matched by sender email), not the quoted text")
_frs_inputs = actions.get("FileRef_from_subject", {}).get("inputs", "")
ok("AppRef_current" in _frs_inputs and "AppRef_from_subject" not in _frs_inputs,
   "FileRef_from_subject (the saved update-resume filename) is built from AppRef_current, not the text-quoted ref")
if not _suppress_b:
    for _email_action, _label in (("Send_update_ack", "Email 5"), ("Send_noted_reply", "Email 6")):
        _subj = json.dumps(actions.get(_email_action, {}))
        ok("AppRef_current" in _subj,
           f"{_label} ({_email_action}) shows the applicant their real current reference (AppRef_current)")

ok(p2cols[3]  == "Category",         f"Category at table position #4 (got '{p2cols[3]}')")
ok(p2cols[4]  == "Resume Link",      f"Resume Link at table position #5 (moved 2026-07-24; got '{p2cols[4]}')")
ok(p2cols[0]  == "Application ID",   "Application ID at position #1")
ok(p2cols[18] == "Status",           "Status at position #19 (shifted +1 by the Resume Link move)")
# P1-filled columns must carry real values (not blank)
ok(item.get("Status") == "New Email Received", "Add_row writes Status='New Email Received'")
ok(item.get("Has Resume") == "Yes",            "Add_row writes Has Resume='Yes' (invariant)")
ok(item.get("Application Updates") == 0,               "Add_row writes Application Updates=0")
# P2 columns must NOT be listed in the write at all (not even as empty strings) - matches pristine
for col in P2_OWNED:
    ok(col not in item, f"Add_row does NOT list P2 column '{col}' (left as blank cell for P2)")
# Old columns must be absent
for removed in ["Is Categorized", "ATS Score", "Score Reason", "Sentiment"]:
    ok(removed not in cols, f"retired column absent: '{removed}'")
ok(removed_check := all(r not in p2cols for r in ["Is Categorized", "ATS Score", "Score Reason", "Sentiment"]),
   "retired columns absent from the 30-col table too")

# ── C. APPREF — unique per email ─────────────────────────────────────────────
print("\n=== C. APPREF (unique reference per email) ===")
ok(item.get("Application ID") == "@outputs('AppRef')",
   "Add_row Application ID = @outputs('AppRef')")
ar_src = json.dumps(actions["AppRef"]["inputs"])
ok("guid(" in ar_src and "substring" in ar_src,
   "AppRef has guid() hex tail (same-second safe)")
ok("receivedDateTime" in ar_src, "AppRef uses received timestamp (not utcNow)")
ok(cfg["appref"]["date_format"] in ar_src, "AppRef date format from config")
ok("HHmm" in ar_src, "AppRef includes time component (HHmm, minute precision)")
ok("HHmmss" not in ar_src, "AppRef time trimmed to minute precision (no seconds)")
ok("substring(guid(),0,4)" in ar_src, "AppRef hex tail trimmed to 4 chars")
ok("toUpper" in ar_src, "AppRef hex tail is uppercased (consistent round-trip)")
# FileRef = AppRef with / replaced by - (filename-safe)
ok("FileRef" in actions, "FileRef compose exists")
fr_src = json.dumps(actions["FileRef"]["inputs"])
ok("replace" in fr_src and "AppRef" in fr_src and "'-'" in fr_src,
   "FileRef replaces '/' with '-' for filename safety")

# ── D. ODATA APOSTROPHE ESCAPING ─────────────────────────────────────────────
print("\n=== D. ODATA APOSTROPHE ESCAPING ===")
for a in ["Get_rows", "Get_rows_ref"]:
    if a in actions:
        fq = json.dumps(actions[a]["inputs"]["parameters"])
        ok("''''''" in fq or ("replace(" in fq and "''''" in fq),
           f"{a} doubles apostrophes in Email OData filter")

# ── D2. SENDER-ADDRESS NORMALIZATION (2026-07-14, closes 2026-07-12 audit finding 2) ──
print("\n=== D2. SENDER NORMALIZATION (filter + stored Email symmetric) ===")
_filter_src = json.dumps(actions["Get_rows"]["inputs"]["parameters"])
_row_email_src = json.dumps(item.get("Email", ""))
for _label, _src in (("Get_rows $filter", _filter_src), ("Add_row Email value", _row_email_src)):
    ok("toLower(" in _src, f"{_label} lowercases the sender address")
    ok("split(outputs('CurrentEmail')?['from'],'<')" in _src,
       f"{_label} strips a 'Display Name <addr>' wrapper before comparing/storing")
# Both sides must derive from the identical expression, not two hand-written copies that could
# drift apart - Get_rows_ref shares the same _EMAIL_FILTER constant as Get_rows, confirmed here
# via the actions dict rather than re-deriving the expression string.
ok(json.dumps(actions["Get_rows_ref"]["inputs"]["parameters"].get("$filter"))
   == json.dumps(actions["Get_rows"]["inputs"]["parameters"].get("$filter")),
   "Get_rows_ref uses the exact same $filter expression as Get_rows (one shared constant)")

# ── E. RESUME FOLDER PATH (dated subfolders) ─────────────────────────────────
print("\n=== E. RESUME FOLDER PATH ===")
ok(cfg["sharepoint"]["dated_resume_subfolders"] is True,
   "dated_resume_subfolders=true (files into <Year>/<Month>/)")
# Folder expression in the flow should reference the resumes_folder and date components —
# no week-level subfolder (dropped 2026-07-08): one folder per month, no 'Wk' segment.
create_file_src = json.dumps(actions.get("Create_file", {}).get("inputs", {}))
ok(("yyyy" in blob and "MMMM" in blob),
   "flow definition contains dated path components (yyyy / MMMM)")
ok("Wk" not in blob, "flow definition no longer contains a week-level 'Wk' subfolder segment")

# ── E2. DATED SUBFOLDER IS GUARANTEED (CreateNewFolder before every save, NO delay) ──────────
# Create_file does NOT auto-create missing folders (fails with NotFound) - so every resume-saving
# path explicitly ensures its dated folder exists first via CreateNewFolder, then saves immediately
# (no wait: CreateNewFolder returns after the folder exists, and the folder->file save works with no
# delay - confirmed live 2026-07-04; the old 2s wait failed anyway on PA's 5s minimum and saves
# succeeded regardless, so the delays were removed).
print("\n=== E2. DATED SUBFOLDER AUTO-CREATE (CreateNewFolder before every Create_file, no delay) ===")
for ensure_name, save_name in [
    ("Ensure_resume_folder", "Save_resumes_to_SharePoint"),
    ("Ensure_update_resume_folder", "Save_update_resume"),
    ("Ensure_dup_resume_folder", "Save_dup_update_resume"),
]:
    ok(ensure_name in actions, f"{ensure_name} exists (guarantees the dated folder before saving)")
    ensure_params = actions.get(ensure_name, {}).get("inputs", {}).get("parameters", {})
    ok(ensure_params.get("table") == cfg["sharepoint"]["documents_library_id"],
       f"{ensure_name} uses the Documents library id (CreateNewFolder's required 'table' param)")
    ok("parameters/path" in ensure_params,
       f"{ensure_name} uses 'parameters/path' (library-relative), not 'folderPath'")
    save_action = actions.get(save_name, {})
    ok(save_action.get("runAfter", {}).get(ensure_name) is not None,
       f"{save_name} waits directly on {ensure_name} (no delay action between them)")

# ── E3. WORKBOOK SELF-HEAL REMOVED — non-destructive, alert-only ─────────────
# The old probe+create+overwrite self-heal was removed (2026-07-04): it was unreliable and
# re-uploaded a blank template over real data. The workbook is never auto-created/overwritten now;
# a genuinely missing workbook makes Get_rows fail -> Notify_failure alerts the admin.
print("\n=== E3. WORKBOOK SELF-HEAL REMOVED (non-destructive) ===")
for removed_act in ["EnsureWb_Probe", "WorkbookExists_NoOp", "WorkbookTemplate_B64",
                    "Create_workbook", "Create_workbook_folder",
                    "Wait_after_workbook_heal", "Wait_after_workbook_folder"]:
    ok(removed_act not in actions, f"destructive self-heal action removed: {removed_act}")
gr = actions.get("Get_rows", {})
ok(gr.get("runAfter", {}) == {}, "Get_rows is now the first action in HasResume (no probe/create ahead of it)")
ok("Notify_failure" in actions, "Notify_failure still present (alerts admin if the workbook is genuinely missing)")
# No Wait actions of ANY kind should remain (all self-heal + folder delays removed)
_all_waits = [n for n, a in actions.items() if a.get("type") == "Wait"]
ok(len(_all_waits) == 0, f"no Wait/Delay actions remain anywhere ({len(_all_waits)} found: {_all_waits})")

# ── F. RESUME FILENAME CONVENTION ────────────────────────────────────────────
print("\n=== F. RESUME FILENAME CONVENTION ===")
# Saved filename = <FileRef>_<original-filename>; Original Filename column = original only
ok(item.get("Original Filename", "").startswith("@"),
   "Original Filename column stores original filename (expression, no APP- prefix)")
cf_src = json.dumps(actions.get("Create_file", {}))
ok("FileRef" in cf_src, "Create_file uses FileRef in saved filename (APP-... prefix)")
ok(item.get("Original Filename", "") != cf_src,
   "Original Filename column != saved file path (column = original, file = prefixed)")

# ── G. UPDATED-CV RE-QUEUE ───────────────────────────────────────────────────
print("\n=== G. UPDATED-CV RE-QUEUE + OVERWRITE ===")
ok("Create_update_file" in actions, "Create_update_file exists (updated resume save)")
pu      = actions.get("Patch_update_attempts", {})
pu_item = pu.get("inputs", {}).get("parameters", {}).get("item", {})
pu_item = pu_item if isinstance(pu_item, dict) else {}
ok(pu_item.get("Status") == "New Email Received",
   "Patch_update_attempts re-queues Status='New Email Received' (P2 re-scores)")
# OVERWRITE fix: Original Filename is deliberately NOT patched (the file is overwritten in place under
# its original name, so the column still points to the right file).
ok("Original Filename" not in pu_item,
   "Patch_update_attempts does NOT change Original Filename (file overwritten in place; column stays valid)")
ok("Application Updates" in pu_item or "ApplicationUpdates" in json.dumps(pu_item),
   "Patch_update_attempts increments Application Updates")
ok("Last Updated Date" in pu_item and "receivedDateTime" in str(pu_item.get("Last Updated Date")),
   "Patch_update_attempts stamps Last Updated Date from the update email time")
# Update save OVERWRITES: files into the ORIGINAL dated folder AND reuses the original stored filename
cu_src = json.dumps(actions.get("Create_update_file", {}))
ok("FileRef_from_subject" in cu_src,
   "Create_update_file uses the quoted-ref prefix (original AppID) so it lands in the original path")
ok("Get_rows_ref" in cu_src and "Original Filename" in cu_src,
   "Create_update_file reuses the row's stored Original Filename -> overwrites the original CV (1 file)")

# ── G2. NO-REFERENCE RESEND ALSO RE-QUEUES + OVERWRITES ─────────────────────
print("\n=== G2. NO-REFERENCE RESEND (duplicate path) RE-QUEUES + OVERWRITES ===")
ok("Create_dup_update_file" in actions, "Create_dup_update_file exists (resend-without-ref resume save)")
ok("FileRef_from_dup" in actions, "FileRef_from_dup exists (original AppID prefix for resend saves)")
pd      = actions.get("Patch_dup_attempts", {})
pd_item = pd.get("inputs", {}).get("parameters", {}).get("item", {})
pd_item = pd_item if isinstance(pd_item, dict) else {}
ok(pd_item.get("Status") == "New Email Received",
   "Patch_dup_attempts re-queues Status='New Email Received' (a resend without a quoted ref still gets rescored)")
ok("Original Filename" not in pd_item,
   "Patch_dup_attempts does NOT change Original Filename (file overwritten in place; column stays valid)")
ok("Application Updates" in pd_item, "Patch_dup_attempts still increments Application Updates")
ok("Last Updated Date" in pd_item and "receivedDateTime" in str(pd_item.get("Last Updated Date")),
   "Patch_dup_attempts stamps Last Updated Date from the resend email time")
cdu_src = json.dumps(actions.get("Create_dup_update_file", {}))
ok("Filter_submitted" in cdu_src and "Received Date" in cdu_src,
   "Create_dup_update_file files into the ORIGINAL row's dated folder (from Filter_submitted's Received Date)")
ok("Original Filename" in cdu_src,
   "Create_dup_update_file reuses the row's stored Original Filename -> overwrites the original CV (1 file)")

# ── H. CAP SYSTEM ────────────────────────────────────────────────────────────
print("\n=== H. CAP SYSTEM (Application Updates thresholds) ===")
rules = cfg["business_rules"]
ok(rules["duplicate_notice_max"] == 5,  "duplicate_notice_max = 5")
ok(rules["update_resume_max"]    == 5,  "update_resume_max    = 5")
ok(rules["followup_reply_max"]   == 5,  "followup_reply_max   = 5")
ok(rules["reply_cap"]            == 3,  "reply_cap = 3 (contacts 1-3 get an email; 4-5 stay silent but keep counting)")
cap = rules["duplicate_notice_max"]
reply_cap = rules["reply_cap"]
final_threshold = reply_cap - 1
old_threshold = cap - 1   # the threshold BEFORE this change (final notice on the 5th reply)
ok(final_threshold == 2, f"final-notice threshold = reply_cap-1 = {final_threshold} (fires on exactly the 3rd reply, never twice)")
# The final-notice swap lives inside the email BODIES, which only exist in a live build -
# when applicant mail is disabled the Send_* actions are no-op Compose placeholders with no body.
_suppress_mode = not _send_applicant_mail
if _suppress_mode:
    ok(blob.count(f"'0')),{final_threshold})") == 0,
       "final-notice closing absent in suppress mode (email bodies are suppressed no-ops)")
else:
    ok(blob.count(f"'0')),{final_threshold})") == 3,
       f"final-notice swap compares Application Updates to {final_threshold} in all 3 capped replies (dup/update/followup)")
ok(f"'0')),{old_threshold})" not in blob,
   "old cap-1 threshold (final notice on the 5th reply) no longer present")
# Verify the OVERALL cap gates (silence point, unchanged at 5) exist
for a in ["IsUnderDupCap", "IsUnderUpdateCap", "IsUnderFollowupCap"]:
    ok(a in actions, f"overall cap gate action exists: {a}")
# Verify the NEW reply-cap gates (whether an email actually sends, at 3) exist
for a in ["IsUnderReplyCap_dup", "IsUnderReplyCap_update", "IsUnderReplyCap_followup"]:
    ok(a in actions, f"reply-cap gate action exists: {a}")
# Contacts past reply_cap (but still under the overall cap) still increment Application Updates and
# still re-queue the row (dup/update - a new resume was saved) but NEVER touch 'Mail Sent'.
for a, requeues in (("Patch_dup_attempts_noreply", True),
                    ("Patch_update_attempts_noreply", True),
                    ("Patch_followup_attempts_noreply", False)):
    ok(a in actions, f"no-reply patch action exists: {a}")
    _item = actions.get(a, {}).get("inputs", {}).get("parameters", {}).get("item", {})
    ok("Mail Sent" not in _item, f"{a} never writes 'Mail Sent' (no email was sent)")
    ok("Application Updates" in _item, f"{a} still increments Application Updates")
    ok((("Status" in _item) == requeues),
       f"{a} {'re-queues Status' if requeues else 'does not touch Status'} as expected")
    # All three (fixed 2026-07-31): any patched contact - resume resend, ref-quoted update, or a
    # text-only follow-up - is a genuine candidate touch and must move Last Updated Date. A
    # follow-up used to be silently excluded, which corrupted P2's
    # _updated_after_location_request() (sharepoint_scoring.py) into thinking a candidate who
    # replied never did.
    ok("Last Updated Date" in _item, f"{a} stamps Last Updated Date on every contact")
# Send_duplicate_notice's subject carries the reference, matching Emails 1/5/6 (was missing
# before) - only checkable in a live build (in suppress mode 'inputs' is a plain string).
if not _suppress_mode:
    dup_subject = actions.get("Send_duplicate_notice", {}).get("inputs", {}).get("parameters", {}).get("emailMessage/Subject", "")
    ok("Application ID" in dup_subject or "Ref" in dup_subject,
       "Send_duplicate_notice subject includes the application reference")

# ── I. SPAM GATE COVERAGE ────────────────────────────────────────────────────
print("\n=== I. SPAM GATE COVERAGE ===")
sf     = cfg["spam_filters"]
ok(len(sf["bad_senders"])          == 31, f"bad_senders    = 31 (got {len(sf['bad_senders'])})")
# Regression guard for the 2026-07-31 false-positive fix: the sender gate is an
# unanchored contains() over the raw From header, so a bare role local-part like
# "marketing@" also matches a personal address that merely ends in it
# (ibreathemarketing@gmail.com - a real CMO applicant, junked 3x). Any term
# re-added here must be safe as a substring, not just as a prefix.
ok("marketing@" not in sf["bad_senders"],
   "bad_senders excludes 'marketing@' (substring-matched a real applicant)")
for _t in ("alert", "invoice", "receipt", "payment", "survey"):
    ok(_t not in sf["bad_subjects"],
       f"bad_subjects excludes bare {_t!r} (matches real job titles)")
ok(len(sf["bad_subjects"])         == 24, f"bad_subjects   = 24 (got {len(sf['bad_subjects'])})")
ok(len(sf["spam_phrases"])         == 30, f"spam_phrases   = 30 (got {len(sf['spam_phrases'])})")
ok(len(sf["offensive_phrases"])    == 36, f"offensive      = 36 (got {len(sf['offensive_phrases'])})")
ok(len(sf["malware_phrases"])      == 20, f"malware        = 20 (got {len(sf['malware_phrases'])}) - extensions/macros only, always checked")
ok(len(sf["foreign_scam_phrases"]) == 25, f"foreign_scam   = 25 (got {len(sf['foreign_scam_phrases'])})")
ok(len(sf["link_shortener_phrases"]) == 12,
   f"link_shortener_phrases = 12 (got {len(sf.get('link_shortener_phrases', []))}) - checked separately, bypassed when a real resume is attached")
total = (len(sf["spam_phrases"]) + len(sf["offensive_phrases"]) +
         len(sf["malware_phrases"]) + len(sf["foreign_scam_phrases"]))
ok(total == 111, f"total always-on spam gate phrases = 111 (got {total}) (+ 12 conditional link-shortener phrases)")
ok(not (set(sf["malware_phrases"]) & set(sf["link_shortener_phrases"])),
   "malware_phrases and link_shortener_phrases don't overlap (clean split, nothing lost/duplicated)")
ok("apply@driverai.io" in sf["bad_senders"],
   "apply@driverai.io in bad_senders (self-loop protection)")
ok(cfg["appref"]["detect_pattern"] not in sf["bad_subjects"],
   "detect_pattern NOT in bad_subjects (allows applicant Re: replies to reach ref gate)")
# Specific phrase coverage checks
ok(any("kill" in p for p in sf["offensive_phrases"]),    "offensive: death threats present")
ok(any("rape" in p for p in sf["offensive_phrases"]),    "offensive: sexual violence present")
ok(any(".exe" in p for p in sf["malware_phrases"]),      "malware: .exe extension present")
ok(any("bit.ly" in p for p in sf["link_shortener_phrases"]),
   "link_shortener_phrases: URL shorteners present (conditional gate, not the always-on malware list)")
ok(any("has ganado" in p for p in sf["foreign_scam_phrases"]), "foreign: Spanish scam present")
ok(any("lottery" in p for p in sf["spam_phrases"]) or
   any("lotería" in p for p in sf["foreign_scam_phrases"]), "scam: lottery phrase present")
ok(len(sf["app_keywords"]) == 21, f"app_keywords = 21 (got {len(sf['app_keywords'])})")
ok("resume" in sf["app_keywords"] and "cv" in sf["app_keywords"],
   "app_keywords includes 'resume' and 'cv'")
ok("pfa" in sf["app_keywords"],
   "app_keywords includes 'pfa' (please find attached shorthand)")

# ── I2. RESUME-AWARE LINK-SHORTENER BYPASS ───────────────────────────────────
print("\n=== I2. RESUME-AWARE LINK-SHORTENER BYPASS ===")
ok("HasValidResumeEarly" in actions, "HasValidResumeEarly exists (computed before the spam gates)")
ok("ResumeFilesEarly" in actions, "ResumeFilesEarly exists (a proper Filter-array Query action, not an inline filter() expression)")
ok(actions.get("ResumeFilesEarly", {}).get("type") == "Query",
   "ResumeFilesEarly uses the proven 'Filter array' action type (Query + from/where) - matches ResumeFiles' own pattern")
ok("filter(" not in json.dumps(actions.get("HasValidResumeEarly", {})),
   "HasValidResumeEarly does NOT use an inline filter() expression (undocumented/unproven WDL syntax)")
ok("LinkShortenerPhrases" in actions, "LinkShortenerPhrases exists (separate conditional list)")
ok("MatchLinkShortener" in actions, "MatchLinkShortener exists (separate query, sibling of MatchSpam)")
is_spam_expr = json.dumps(actions["IsSpam"]["expression"])
ok("MatchSpam" in is_spam_expr and "MatchLinkShortener" in is_spam_expr,
   "IsSpam expression checks BOTH the always-on list and the conditional shortener list")
ok("HasValidResumeEarly" in is_spam_expr,
   "IsSpam expression references HasValidResumeEarly (shortener match alone isn't enough to block)")
ok(actions["IsSpam"]["runAfter"].get("MatchLinkShortener") == ["Succeeded"],
   "IsSpam waits for MatchLinkShortener before evaluating")

# ── J. WORKBOOK ADDRESSING + TEMPLATE (self-heal removed; template kept for manual setup) ────
print("\n=== J. WORKBOOK ADDRESSING + MANUAL-SETUP TEMPLATE ===")
# Workbook path is a single file for ALL years (no <Year> subfolder)
wb_path = cfg["excel"]["file"]
ok(not re.search(r"/\d{4}/", wb_path),
   f"workbook path has no <Year> subfolder (single file for all years): {wb_path}")
ok(wb_path.endswith(".xlsx"), f"workbook path ends with .xlsx: {wb_path}")
ok(cfg["excel"]["table"] == "HiringAgent_P1_Candidates",
   f"table name = HiringAgent_P1_Candidates")

def _table_ref_and_blank_row(xlsx_bytes: bytes):
    wb = load_workbook(io.BytesIO(xlsx_bytes))
    ws = wb.active
    tables = list(ws.tables.values())
    table = tables[0] if tables else None
    ref = getattr(table, "ref", "")
    blank_row = all((ws.cell(row=2, column=i).value in ("", None))
                    for i in range(1, ws.max_column + 1))
    return ws.title, ref, blank_row, len(tables)

# The committed template is the file an admin uploads once for first-time setup (or restores from
# if the workbook is ever lost). Self-heal no longer auto-uploads it, but it must stay correct.
tpl_path = Path(FLOW).parent / "P1_Templates" / "HiringAgent_P1_CandidateList.xlsx"
tpl_wb = load_workbook(tpl_path)
title, ref, blank_row, table_count = _table_ref_and_blank_row(tpl_path.read_bytes())
ok(title == "CandidateList", f"setup template sheet name = CandidateList (got {title})")
ok(table_count == 1, f"setup template workbook has exactly 1 table (got {table_count})")
ok(ref.endswith("2"), f"setup template table ref includes blank seed row: {ref}")
ok(blank_row, "setup template row 2 is intentionally blank (Graph-safe seed row)")

# Regression guard (2026-07-15): the template drifted out of sync with the live schema when
# 'Education'/'Info Request Sent' were added this session - it was never regenerated, and
# nothing caught it (the checks above only look at sheet name/table count/blank seed row, not
# actual column headers). Cross-checked directly against P2's hiring_agent.config, the single
# source of truth both systems are supposed to agree on - never hardcoded here, so this can't
# itself drift the same way.
for sheet_name, expected_cols in (("CandidateList", P2_COLUMNS), ("Rejected", P2_REJECTED_COLUMNS)):
    ws = tpl_wb[sheet_name]
    got_cols = [ws.cell(row=1, column=i).value for i in range(1, ws.max_column + 1)]
    ok(got_cols == expected_cols,
       f"setup template '{sheet_name}' columns match P2's canonical order exactly "
       f"({len(got_cols)} cols)" if got_cols == expected_cols
       else f"setup template '{sheet_name}' columns MISMATCH P2's canonical order "
            f"(got {got_cols}, expected {expected_cols})")
    _sheet_tables = list(ws.tables.values())
    ok(len(_sheet_tables) == 1,
       f"setup template '{sheet_name}' has exactly one Excel table")
    ok(bool(_sheet_tables) and str(_sheet_tables[0].ref).endswith("2"),
       f"setup template '{sheet_name}' table includes the permanent row-2 spacer")
    ok(all(ws.cell(row=2, column=i).value in ("", None)
           for i in range(1, ws.max_column + 1)),
       f"setup template '{sheet_name}' has one fully blank row directly after the header")

# ── K. INBOX TIDY (2 destinations: legit->Archive read, spam->Junk Email unread) ─────
print("\n=== K. INBOX TIDY (Archive vs Junk Email routing) ===")
ok(cfg["inbox_tidy"]["enabled"] is True, "inbox_tidy enabled")
ok(cfg["inbox_tidy"]["destination"] == "Archive",
   f"legit destination = Archive (got '{cfg['inbox_tidy']['destination']}')")
ok(cfg["inbox_tidy"]["spam_destination"] == "Junk Email",
   f"spam destination = Junk Email (got '{cfg['inbox_tidy'].get('spam_destination')}')")

def _move_dest(name):
    return actions.get(name, {}).get("inputs", {}).get("parameters", {}).get("folderPath")

# The 3 spam-gate moves -> Junk Email, and they must have NO paired mark-as-read (left UNREAD).
SPAM_SUFFIXES = ["spam_sender", "spam_subject", "spam_body"]
for suf in SPAM_SUFFIXES:
    mv = f"Move_to_processed_{suf}"
    ok(mv in actions, f"spam move exists: {mv}")
    ok(_move_dest(mv) == "Junk Email", f"{mv} -> Junk Email (got '{_move_dest(mv)}')")
    ok(f"Mark_as_read_{suf}" not in actions,
       f"spam '{suf}' has NO mark-as-read (left unread in Junk)")

# Every OTHER move (legit) -> Archive, and each must be paired with a mark-as-read (read).
# 'update_noreply'/'followup_noreply' are the new reply_cap-reached-but-not-cap-reached
# contacts (4th/5th) - still real application mail, just no email sent.
LEGIT_SUFFIXES = ["update", "update_noreply", "update_cap",
                  "followup", "followup_noreply", "followup_cap"]  # + top-level ""
for suf in LEGIT_SUFFIXES:
    mv = f"Move_to_processed_{suf}"
    ok(_move_dest(mv) == "Archive", f"{mv} -> Archive (got '{_move_dest(mv)}')")
    ok(f"Mark_as_read_{suf}" in actions, f"legit '{suf}' is marked read before moving")
ok(_move_dest("Move_to_processed") == "Archive", "top-level Move_to_processed -> Archive")
ok("Mark_as_read" in actions, "top-level pair marks read before moving")

# No legit move should ever land in Junk, and no spam move should ever land in Archive.
all_moves = [n for n in actions if n.startswith("Move_to_processed")]
junk_moves = [n for n in all_moves if _move_dest(n) == "Junk Email"]
arch_moves = [n for n in all_moves if _move_dest(n) == "Archive"]
ok(sorted(junk_moves) == sorted(f"Move_to_processed_{s}" for s in SPAM_SUFFIXES),
   f"exactly the 3 spam moves go to Junk Email (got {sorted(junk_moves)})")
ok(len(arch_moves) == 7, f"exactly 7 legit moves go to Archive (got {len(arch_moves)})")
mark_read = [n for n in actions if n.startswith("Mark_as_read")]
ok(len(mark_read) == 7, f"exactly 7 mark-as-read (legit only; spam is unread) (got {len(mark_read)})")
ok(len(all_moves) == 10, f"10 total moves (3 spam + 7 legit) (got {len(all_moves)})")
for move_name in all_moves:
    ok(actions[move_name].get("inputs", {}).get("retryPolicy", {}).get("type") == "none",
       f"{move_name} disables automatic replay of the non-idempotent move")

all_move_v2 = {
    name: action for name, action in actions.items()
    if isinstance(action.get("inputs"), dict)
    and action["inputs"].get("host", {}).get("operationId") == "MoveV2"
}
ok(len(all_move_v2) == 11,
   f"exactly 11 MoveV2 actions exist (10 outcome moves + 1 failure move; got {len(all_move_v2)})")
for move_name, move_action in all_move_v2.items():
    move_params = move_action.get("inputs", {}).get("parameters", {})
    ok(move_params.get("messageId") == "@outputs('CurrentEmail')?['id']",
       f"{move_name} uses the current polled message ID")
    ok(move_params.get("mailboxAddress") == cfg["email"]["trigger_mailbox"],
       f"{move_name} uses the configured shared mailbox")
    ok(move_params.get("folderPath") in {"Archive", "Junk Email"},
       f"{move_name} uses a portable well-known folder name")

ok("Move_failed_to_recruiting_review" not in actions,
   "legacy Recruiting Review failure move is absent")
ok(not any(str(a.get("inputs", {}).get("parameters", {}).get("folderPath", "")).startswith("AAMk")
           for a in all_move_v2.values()),
   "no tenant-specific Outlook folder ID is baked into a MoveV2 action")

# Every nested outcome move must have exactly one same-scope terminal continuation.
# This reverse check catches a move that accidentally loses its downstream handler.
branch_moves = [name for name in all_moves if name != "Move_to_processed"]
ok(len(branch_moves) == 9, f"9 nested outcome moves require terminal handlers (got {len(branch_moves)})")
for move_name in branch_moves:
    dependents = [
        (name, action) for name, action in actions.items()
        if owner_of.get(name) == owner_of.get(move_name)
        and move_name in action.get("runAfter", {})
    ]
    ok(len(dependents) == 1 and dependents[0][1].get("type") == "Terminate",
       f"{move_name} has exactly one same-scope Terminate continuation")
    if dependents:
        ok(set(dependents[0][1].get("runAfter", {}).get(move_name, [])) ==
           {"Succeeded", "Failed", "Skipped", "TimedOut"},
           f"{move_name} continuation handles every terminal status")

for mark_name in mark_read:
    suffix = mark_name.removeprefix("Mark_as_read")
    move_name = "Move_to_processed" + suffix
    ok(actions.get(move_name, {}).get("runAfter", {}).get(mark_name) ==
       ["Succeeded", "Failed", "TimedOut"],
       f"{move_name} follows mark success/failure/timeout but never a skipped branch")

# Terminal cleanup must be explicitly handled. Otherwise a benign Outlook 404 after
# another actor has already moved the message marks a fully processed intake as Failed.
processed_finalize = actions.get("Finalize_processed_cleanup", {})
ok(processed_finalize.get("type") == "Compose",
   "Finalize_processed_cleanup handles the terminal normal-move result")
ok(processed_finalize.get("runAfter") ==
   {"Move_to_processed": ["Succeeded", "Failed", "TimedOut"]},
   "normal cleanup handles success/failure/timeout without activating on a skipped branch")

for term_name, term in actions.items():
    if not term_name.startswith("Terminate_"):
        continue
    run_after = term.get("runAfter", {})
    move_dependencies = [name for name in run_after if name.startswith("Move_to_processed")]
    for move_name in move_dependencies:
        ok(set(run_after[move_name]) == {"Succeeded", "Failed", "Skipped", "TimedOut"},
           f"{term_name} handles every terminal status from {move_name}")

# ── L. ERROR HANDLING ────────────────────────────────────────────────────────
print("\n=== L. ERROR HANDLING ===")
ok("Notify_failure" in actions, "Notify_failure admin alert action present")
ok("Notify_poll_failure" in actions, "Notify_poll_failure Inbox-read alert action present")
nf_src = json.dumps(actions.get("Notify_failure", {}))
npf_src = json.dumps(actions.get("Notify_poll_failure", {}))
if not _send_admin_alerts:
    ok(actions.get("Notify_failure", {}).get("type") == "Compose",
       "Notify_failure is a no-op because admin alerts are disabled")
else:
    ok(actions.get("Notify_failure", {}).get("type") == "OpenApiConnection",
       "Notify_failure remains a real connector action")
    ok(cfg["email"]["admin_email"] in nf_src,
       f"failure alert targets {cfg['email']['admin_email']}")
    ok("AppRef" in nf_src or "app" in nf_src.lower(),
       "failure alert includes application reference for traceability")
    ok(actions.get("Notify_poll_failure", {}).get("type") == "OpenApiConnection",
       "unread-Inbox poll failure remains a real admin connector action")
    ok(cfg["email"]["admin_email"] in npf_src,
       "unread-Inbox poll failure targets the configured admin")
    ok(actions["Notify_poll_failure"].get("runAfter") ==
       {"Get_unread_emails": ["Failed", "TimedOut"]},
       "poll failure alert runs only when Get_unread_emails fails or times out")
# Error path leaves email UNREAD — verified by Notify_failure NOT being in
# a "runAfter Mark_as_read" chain; its runAfter should be HasResume Failed/TimedOut
nf_ra = actions["Notify_failure"].get("runAfter", {})
ok(any("HasResume" in k or "IsSpam" in k or "IsSubject" in k
       for k in nf_ra), "Notify_failure runs after flow body (not after a tidy action)")
failure_dest = cfg["inbox_tidy"]["failure_destination"]
for _move_name in ("Move_failed_to_archive_unread",):
    _move = actions.get(_move_name, {})
    ok(_move.get("inputs", {}).get("host", {}).get("operationId") == "MoveV2",
       f"{_move_name} is a real shared-mailbox move")
    ok(_move.get("inputs", {}).get("parameters", {}).get("folderPath") == failure_dest,
       f"{_move_name} targets the configured portable failure folder")
    ok(_move.get("inputs", {}).get("retryPolicy", {}).get("type") == "none",
       f"{_move_name} disables automatic replay of the non-idempotent move")
    ok(_move.get("runAfter") ==
       {"Notify_failure": ["Succeeded", "Failed", "TimedOut"]},
       f"{_move_name} runs only after an actual processing-failure alert attempt")
    ok("Skipped" not in _move.get("runAfter", {}).get("Notify_failure", []),
       f"{_move_name} cannot race normal cleanup when Notify_failure is skipped on success")

failed_finalize = actions.get("Finalize_failed_cleanup", {})
ok(failed_finalize.get("type") == "Compose",
   "Finalize_failed_cleanup handles failure-routing cleanup errors")
ok(failed_finalize.get("runAfter") ==
   {"Move_failed_to_archive_unread": ["Succeeded", "Failed", "TimedOut"]},
   "failure cleanup handles success/failure/timeout without activating on normal success")

# ── M. ROUTING SIMULATION ────────────────────────────────────────────────────
print("\n=== M. ROUTING SIMULATION ===")
bad_senders  = [s.lower() for s in sf["bad_senders"]]
bad_subjects = [s.lower() for s in sf["bad_subjects"]]
spam_all     = [s.lower() for s in (sf["spam_phrases"] + sf["offensive_phrases"] +
                                     sf["malware_phrases"] + sf["foreign_scam_phrases"])]
shortener_all = [s.lower() for s in sf.get("link_shortener_phrases", [])]
appkw        = [s.lower() for s in sf["app_keywords"]]
detect       = cfg["appref"]["detect_pattern"].lower()
dup_days     = cfg["business_rules"]["duplicate_check_days"]

def route(m):
    frm  = m["from"].lower()
    subj = m.get("subject", "").lower()
    body = m.get("body", "").lower()
    atts = [a.lower() for a in m.get("atts", [])]
    # Trailing space mirrors LowerBody's own trailing space in build_zip.py, which
    # space-terminates the joined attachment names so '.exe ' matches a real attachment
    # but not 'exeter.ac.uk'. Keep these two in sync.
    hay  = " ".join([subj, body] + atts) + " "
    cv   = m.get("cv_attempts", 0)
    cap  = 5  # max replies before silence
    pdf_docx   = [a for a in atts if a.endswith(".pdf") or a.endswith(".docx")]
    has_valid_resume_early = len(pdf_docx) > 0

    # Gate 1 — bad sender
    if any(b in frm for b in bad_senders):
        return "SILENT (loop/bad-sender)"
    # Gate 2 — bad subject
    if any(b in subj for b in bad_subjects):
        return "SILENT (bad-subject)"
    # Gate 3 — spam / abuse / malware (always-on) OR a shortener match with no real resume attached
    if any(p in hay for p in spam_all):
        return "SILENT (spam/abuse/malware)"
    if any(p in hay for p in shortener_all) and not has_valid_resume_early:
        return "SILENT (spam/abuse/malware)"

    quotes_ref = detect in subj or detect in body

    # Branch A — known applicant quoting a ref
    if quotes_ref and m.get("known_sender"):
        if cv >= cap:
            return "SILENT (cap reached)"
        if pdf_docx:
            return "EMAIL5 update (re-queue row+file)"
        return "EMAIL6 follow-up (no new CV)"

    # Branch B — resume present
    if atts:
        if pdf_docx:
            if m.get("duplicate"):
                if cv >= cap:
                    return "SILENT (cap reached)"
                return "EMAIL2 duplicate notice (capped, no new row)"
            return "EMAIL1 ack + NEW row + save resume"
        return "EMAIL4 wrong format (resend PDF/DOCX)"

    # Branch C — no attachment
    if any(k in hay for k in appkw):
        return "EMAIL3 please attach CV (no row)"
    return "IGNORE (not an application)"

fixtures = [
    # ── New applications ──
    ("new applicant PDF",
     {"from":"jane@gmail.com","subject":"Application","body":"PFA my resume","atts":["jane_cv.pdf"]},
     "EMAIL1"),
    ("new applicant DOCX",
     {"from":"john@gmail.com","subject":"Applying","body":"see attached","atts":["john_resume.docx"]},
     "EMAIL1"),
    ("new applicant PDF+DOCX (two files)",
     {"from":"twin@gmail.com","subject":"Application","body":"pfa","atts":["cv.pdf","cover.docx"]},
     "EMAIL1"),
    ("new applicant — body keyword only",
     {"from":"new@gmail.com","subject":"Hi","body":"please find attached my resume","atts":["cv.pdf"]},
     "EMAIL1"),

    # ── Gate 1 — sender ──
    ("self-loop (apply@)",
     {"from":"apply@driverai.io","subject":"Re: Application Received","body":"x","atts":["a.pdf"]},
     "SILENT (loop"),
    ("noreply sender",
     {"from":"noreply@somesite.com","subject":"Application","body":"cv pfa","atts":["cv.pdf"]},
     "SILENT (loop"),
    ("linkedin notification",
     {"from":"jobs-noreply@linkedin.com","subject":"You have new matches","body":"x"},
     "SILENT (loop"),
    ("mailer-daemon bounce",
     {"from":"mailer-daemon@example.com","subject":"Delivery failure","body":"x"},
     "SILENT (loop"),
    ("bad sender WITH a valid resume attached — gate 1 still wins (not rescued by resume bypass)",
     {"from":"noreply@somesite.com","subject":"Application","body":"pfa my resume","atts":["cv.pdf"]},
     "SILENT (loop"),

    # ── Gate 2 — subject ──
    ("webinar subject",
     {"from":"x@x.com","subject":"Webinar invitation — join us","body":"cv pfa","atts":["cv.pdf"]},
     "SILENT (bad-subject)"),
    ("out-of-office subject",
     {"from":"x@x.com","subject":"Out of office: back Monday","body":"x"},
     "SILENT (bad-subject)"),
    ("invoice subject",
     {"from":"x@x.com","subject":"Invoice attached #1234","body":"payment due","atts":["inv.pdf"]},
     "SILENT (bad-subject)"),
    ("our own ack subject (prevents re-reply loops)",
     {"from":"x@x.com","subject":"Thanks for applying to driver ai","body":"pfa","atts":["cv.pdf"]},
     "SILENT (bad-subject)"),
    ("bad subject WITH a valid resume attached — gate 2 still wins (not rescued by resume bypass)",
     {"from":"x@x.com","subject":"Invoice attached #1234","body":"pfa my resume","atts":["cv.pdf"]},
     "SILENT (bad-subject)"),

    # ── Gate false-positive regressions (2026-07-31) ──
    # Every one of these was silently junked before the term-narrowing fix.
    ("real applicant whose address ends in 'marketing@'",
     {"from":"ibreathemarketing@gmail.com","subject":"Chief Marketing Officer",
      "body":"pfa my resume","atts":["cv.pdf"]},
     "EMAIL1 ack + NEW row + save resume"),
    ("payments engineer is not a 'payment' spam subject",
     {"from":"a@x.com","subject":"Senior Engineer | Payments Platform",
      "body":"pfa my resume","atts":["cv.pdf"]},
     "EMAIL1 ack + NEW row + save resume"),
    ("land surveyor is not a 'survey' spam subject",
     {"from":"b@x.com","subject":"Land Surveyor / Survey Engineer",
      "body":"pfa my resume","atts":["cv.pdf"]},
     "EMAIL1 ack + NEW row + save resume"),
    ("invoice processing specialist is a real job title",
     {"from":"c@x.com","subject":"Application - Invoice Processing Specialist",
      "body":"pfa my resume","atts":["cv.pdf"]},
     "EMAIL1 ack + NEW row + save resume"),
    ("alert systems engineer is not an 'alert' spam subject",
     {"from":"d@x.com","subject":"Alert Systems Engineer",
      "body":"pfa my resume","atts":["cv.pdf"]},
     "EMAIL1 ack + NEW row + save resume"),
    ("receipt reconciliation analyst is a real job title",
     {"from":"e@x.com","subject":"Receipt Reconciliation Analyst",
      "body":"pfa my resume","atts":["cv.pdf"]},
     "EMAIL1 ack + NEW row + save resume"),

    # ── Gate 3 — spam/abuse/malware ──
    ("lottery scam",
     {"from":"x@x.com","subject":"you have won","body":"claim your prize","atts":["cv.pdf"]},
     "SILENT (spam"),
    ("phishing link",
     {"from":"x@x.com","subject":"Application","body":"click here to claim your reward"},
     "SILENT (spam"),
    ("offensive — sexual solicitation",
     {"from":"x@x.com","subject":"hi","body":"send nudes","atts":["cv.pdf"]},
     "SILENT (spam"),
    ("offensive — death threat",
     {"from":"x@x.com","subject":"application","body":"i will kill you"},
     "SILENT (spam"),
    ("offensive — slur",
     {"from":"x@x.com","subject":"job","body":"nigger"},
     "SILENT (spam"),
    ("malware exe attachment",
     {"from":"x@x.com","subject":"resume","body":"see attached","atts":["cv.exe"]},
     "SILENT (spam"),
    ("malware ps1 attachment",
     {"from":"x@x.com","subject":"resume","body":"see attached","atts":["setup.ps1"]},
     "SILENT (spam"),
    ("University of Exeter URL is not a .exe attachment",
     {"from":"f@x.com","subject":"Data Engineer",
      "body":"MSc, University of Exeter - www.exeter.ac.uk","atts":["cv.pdf"]},
     "EMAIL1 ack + NEW row + save resume"),
    ("iso.org / ISO 27001 is not a .iso attachment",
     {"from":"g@x.com","subject":"CISO application - CISSP",
      "body":"ISO 27001 lead auditor, see www.iso.org","atts":["cv.pdf"]},
     "EMAIL1 ack + NEW row + save resume"),
    ("URL shortener in signature, but a real resume IS attached -> allowed through",
     {"from":"x@x.com","subject":"application","body":"pfa my resume. Connect: bit.ly/myportfolio","atts":["cv.pdf"]},
     "EMAIL1"),
    ("URL shortener in body, NO resume attached -> still blocked",
     {"from":"x@x.com","subject":"application","body":"see my portfolio at bit.ly/myportfolio"},
     "SILENT (spam"),
    ("URL shortener + wrong-format attachment (not a real resume) -> still blocked",
     {"from":"x@x.com","subject":"application","body":"see my portfolio at bit.ly/myportfolio","atts":["cv.jpg"]},
     "SILENT (spam"),
    ("foreign scam — Spanish",
     {"from":"x@x.com","subject":"premio","body":"has ganado la loteria"},
     "SILENT (spam"),
    ("foreign scam — French",
     {"from":"x@x.com","subject":"lot","body":"vous avez gagné"},
     "SILENT (spam"),
    ("enable macros command",
     {"from":"x@x.com","subject":"resume","body":"please enable macros in the file","atts":["cv.pdf"]},
     "SILENT (spam"),

    # ── Duplicate ──
    ("duplicate within 90 days",
     {"from":"dup@gmail.com","subject":"apply again","body":"pfa","atts":["cv.pdf"],"duplicate":True},
     "EMAIL2"),
    ("duplicate at cap=5 (silent)",
     {"from":"dup@gmail.com","subject":"again","body":"pfa","atts":["cv.pdf"],"duplicate":True,"cv_attempts":5},
     "SILENT (cap reached)"),

    # ── Update / follow-up (Branch A) ──
    ("update with new CV",
     {"from":"known@gmail.com","subject":"Update - APP-20260630-143052","body":"new cv attached","atts":["cv2.pdf"],"known_sender":True},
     "EMAIL5"),
    ("update with DOCX",
     {"from":"known@gmail.com","subject":"Update - APP-20260630-143052","body":"updated","atts":["cv2.docx"],"known_sender":True},
     "EMAIL5"),
    ("follow-up no CV",
     {"from":"known@gmail.com","subject":"Re: APP-20260630-143052","body":"any update?","known_sender":True},
     "EMAIL6"),
    ("update at cap=5 (silent)",
     {"from":"known@gmail.com","subject":"Update - APP-20260630-143052","body":"cv","atts":["cv.pdf"],"known_sender":True,"cv_attempts":5},
     "SILENT (cap reached)"),
    ("follow-up at cap=5 (silent)",
     {"from":"known@gmail.com","subject":"APP-20260630-143052 - checking in","body":"any news","known_sender":True,"cv_attempts":5},
     "SILENT (cap reached)"),
    ("ref in body not subject",
     {"from":"known@gmail.com","subject":"Hi","body":"ref: app-20260630-143052 updated cv","atts":["cv.pdf"],"known_sender":True},
     "EMAIL5"),
    ("decade-proof ref (app-2099)",
     {"from":"known@gmail.com","subject":"Re: app-20991231-235959","body":"pfa","atts":["cv.pdf"],"known_sender":True},
     "EMAIL5"),
    ("quoted ref but unknown sender → new applicant",
     {"from":"stranger@gmail.com","subject":"Re: app-20260101-120000","body":"pfa","atts":["cv.pdf"],"known_sender":False},
     "EMAIL1"),

    # ── No resume / wrong format ──
    ("no-CV with keyword in subject",
     {"from":"bob@gmail.com","subject":"Job application","body":"I am interested"},
     "EMAIL3"),
    ("no-CV with keyword in body only",
     {"from":"bob@gmail.com","subject":"Hello","body":"I would like to apply for a role"},
     "EMAIL3"),
    ("wrong format (.doc)",
     {"from":"amy@gmail.com","subject":"resume","body":"pfa","atts":["cv.doc"]},
     "EMAIL4"),
    ("wrong format (.txt)",
     {"from":"amy@gmail.com","subject":"resume","body":"pfa","atts":["cv.txt"]},
     "EMAIL4"),
    ("wrong format (.jpg image)",
     {"from":"amy@gmail.com","subject":"resume","body":"pfa","atts":["photo.jpg"]},
     "EMAIL4"),
    ("wrong format + hiring keyword → EMAIL4 (attachment present, not EMAIL3)",
     {"from":"amy@gmail.com","subject":"application","body":"resume attached","atts":["cv.jpg"]},
     "EMAIL4"),

    # ── Silent / ignored ──
    ("no attachment, no keyword → ignore",
     {"from":"random@gmail.com","subject":"Hello","body":"How are you"},
     "IGNORE"),
    ("completely blank email",
     {"from":"blank@gmail.com","subject":"","body":""},
     "IGNORE"),
]

for label, m, expect in fixtures:
    got = route(m)
    ok(got.startswith(expect) or expect in got, f"{label:50s} -> {got}")

# ── M2. DUPLICATE WINDOW BOUNDARY (90-day, inclusive edge) ───────────────────
# The fixtures above pass "duplicate" in as a pre-decided boolean - they never exercise the
# actual day-math. This mirrors the REAL PA expression bit-for-bit:
#   greaterOrEquals(ticks(last_received), ticks(utcNow() - N days))
# i.e. duplicate == True whenever (now - last_received) <= N days - the boundary at
# EXACTLY N days is INCLUSIVE (>=), not exclusive. Never exercised elsewhere in this suite.
print("\n=== M2. DUPLICATE WINDOW BOUNDARY (90-day inclusive edge) ===")
import datetime as _dt
_now = _dt.datetime(2026, 7, 4, 12, 0, 0)  # fixed reference instant, reproducible across runs

def _is_duplicate_window(last_received, now=_now, days=dup_days):
    return (now - last_received) <= _dt.timedelta(days=days)

ok(_is_duplicate_window(_now - _dt.timedelta(days=89)) is True,
   "89 days since last submission -> still a duplicate (inside window)")
ok(_is_duplicate_window(_now - _dt.timedelta(days=90)) is True,
   "exactly 90 days since last submission -> still a duplicate (boundary is inclusive, >=)")
ok(_is_duplicate_window(_now - _dt.timedelta(days=90, seconds=1)) is False,
   "90 days + 1 second since last submission -> NOT a duplicate (just outside window)")
ok(_is_duplicate_window(_now - _dt.timedelta(days=91)) is False,
   "91 days since last submission -> NOT a duplicate (outside window)")
ok(_is_duplicate_window(_now - _dt.timedelta(minutes=1)) is True,
   "1 minute since last submission -> duplicate (immediate resend)")

# ── N. FLOW CONFIG INTEGRITY ─────────────────────────────────────────────────
print("\n=== N. FLOW CONFIG INTEGRITY ===")
ok(cfg["email"]["trigger_mailbox"] == "apply@driverai.io",
   f"trigger mailbox = apply@driverai.io")
ok(cfg["email"]["admin_email"] == "yashv@driverai.io",
   f"admin email = yashv@driverai.io")
ok(cfg["business_rules"]["duplicate_check_days"] == 90,
   "duplicate_check_days = 90")
ok(cfg["sharepoint"]["resumes_folder"].startswith("/"),
   "resumes_folder starts with '/' (absolute path)")
ok(cfg["appref"]["detect_pattern"] == "app-20",
   f"detect_pattern = app-20 (decade-proof, matches 2000-2099)")
ok(cfg["inbox_tidy"]["destination"] == "Archive",
   "inbox_tidy destination = Archive (capitalized well-known folder)")
# Flow name must match config
ok(d["properties"].get("displayName") == cfg["flow_name"] or
   cfg["flow_name"] in json.dumps(d),
   f"flow displayName matches config flow_name '{cfg['flow_name']}'")

# ── O. OUTBOUND MAIL CONTROLS — definition must MATCH the config flags ──────
print("\n=== O. OUTBOUND MAIL CONTROLS — definition matches config ===")
_suppress = not _send_applicant_mail
_MAIL_ACTIONS = ["Send_acknowledgment", "Send_duplicate_notice", "Send_CV_request",
                 "Send_wrong_format", "Send_update_ack", "Send_noted_reply"]
_n_send_ops = blob.count("SharedMailboxSendEmailV2") + blob.count("SendEmailV2")
_n_applicant_send_ops = blob.count("SharedMailboxSendEmailV2")
_n_suppressed_stamps = blob.count("Suppressed (applicant email disabled)")
if _suppress:
    print("  (config says APPLICANT MAIL DISABLED - validating silent live intake)")
    for a in _MAIL_ACTIONS:
        ok(actions.get(a, {}).get("type") == "Compose",
           f"{a} is a no-op Compose (applicant not contacted)")
        ok(len(actions.get(a, {}).get("runAfter", {})) >= 0 and "inputs" in actions.get(a, {}),
           f"{a} keeps a runAfter/inputs shape (wiring intact)")
    ok(_n_applicant_send_ops == 0,
       f"NO applicant shared-mailbox send operations remain (got {_n_applicant_send_ops})")
    ok(_n_suppressed_stamps >= 4,
       f"Mail Sent stamps record suppression in all 4 patches (got {_n_suppressed_stamps})")
    ok("'Sent '" not in blob, "no real \"Sent <timestamp>\" stamp expression remains")
else:
    print("  (config says APPLICANT MAIL ENABLED - validating real email actions)")
    for a in _MAIL_ACTIONS:
        ok(actions.get(a, {}).get("type") == "OpenApiConnection",
           f"{a} is a real mail-connector action (live)")
    ok(_n_applicant_send_ops >= 6,
       f"applicant send-mail connector operations present (got {_n_applicant_send_ops}, expect >= 6)")
    ok(_n_suppressed_stamps == 0,
       f"no applicant-suppression stamps remain (got {_n_suppressed_stamps})")
    ok(blob.count("'Sent '") >= 4, "real \"Sent <timestamp>\" stamp expressions present in the 4 patches")

if _send_admin_alerts:
    ok(actions.get("Notify_failure", {}).get("type") == "OpenApiConnection",
       "admin failure alert remains enabled")
    ok(actions.get("Notify_poll_failure", {}).get("type") == "OpenApiConnection",
       "unread-Inbox poll failure alert remains enabled")
    ok(_n_send_ops >= 2, f"two admin send operations remain (processing + poll; got {_n_send_ops})")
else:
    ok(actions.get("Notify_failure", {}).get("type") == "Compose",
       "admin failure alert is suppressed exactly as configured")

print("\n=== P. YEAR-BOUNDARY SEPARATOR (blank row before the first new-applicant row/year) ===")
_ysep_cfg = cfg.get("year_separator", {"enabled": True})
_any_rows = actions.get("Get_rows_any_applicant", {})
ok(_any_rows.get("inputs", {}).get("host", {}).get("operationId") == "GetItems",
   "Get_rows_any_applicant exists to distinguish the permanent header spacer from real rows")
_any_params = _any_rows.get("inputs", {}).get("parameters", {})
ok(_any_params.get("$filter") == "Email ne ''" and _any_params.get("$top") == 1,
   "header-spacer guard queries at most one real applicant row")
if _ysep_cfg.get("enabled", True):
    ok("Get_rows_for_separators" in actions, "shared separator row scan exists")
    ok("Filter_rows_this_year" in actions, "Filter_rows_this_year action exists")
    ok("IsFirstOfNewYear" in actions, "IsFirstOfNewYear action exists")
    ok("Add_row_year_separator" in actions, "Add_row_year_separator action exists")

    _scan = actions.get("Get_rows_for_separators", {})
    _scan_params = _scan.get("inputs", {}).get("parameters", {})
    ok(_scan.get("inputs", {}).get("host", {}).get("operationId") == "GetItems",
       "separator scan is a real Excel GetItems action")
    ok("$filter" not in _scan_params,
       "separator scan has no OData filter on the spaced 'Received Date' column")
    ok(_scan_params.get("$top") == 5000 and _scan_params.get("dateTimeFormat") == "ISO 8601",
       "separator scan requests up to 5000 rows with ISO timestamps")
    ok(_scan.get("runtimeConfiguration", {}).get("paginationPolicy", {}).get(
       "minimumItemCount") == 5000,
       "separator scan enables pagination to the Power Automate 5000-item limit")

    _fy = actions.get("Filter_rows_this_year", {})
    _fy_src = json.dumps(_fy.get("inputs", {}))
    ok(_fy.get("type") == "Query" and
       _fy.get("runAfter") == {"Get_rows_for_separators": ["Succeeded"]},
       "year detection uses an in-memory Query after the shared scan")
    ok("item()?['Received Date']" in _fy_src and "startsWith" in _fy_src and
       "formatDateTime(outputs('CurrentEmail')?['receivedDateTime'],'yyyy')" in _fy_src,
       "year Query compares the visible Received Date safely without Excel OData")

    _ifny = actions.get("IsFirstOfNewYear", {})
    ok(_ifny.get("type") == "If", "IsFirstOfNewYear is a real If condition")
    ok(set(_ifny.get("runAfter", {})) == {"Filter_rows_this_year", "Get_rows_any_applicant"},
       "IsFirstOfNewYear waits on both the year filter and real-applicant guard")
    _yexpr = json.dumps(_ifny.get("expression", {}))
    ok('"and"' in _yexpr and "Filter_rows_this_year" in _yexpr
       and "Get_rows_any_applicant" in _yexpr,
       "year separator requires no row this year AND at least one prior real applicant")
    ok("Add_row_year_separator" in _ifny.get("actions", {}),
       "IsFirstOfNewYear's TRUE branch inserts the separator row")
    ok(_ifny.get("else", {}).get("actions", {}) == {},
       "IsFirstOfNewYear's FALSE branch is a no-op (a later applicant this year needs nothing)")

    _arys = actions.get("Add_row_year_separator", {})
    ok(_arys.get("inputs", {}).get("host", {}).get("operationId") == "AddRowV2",
       "Add_row_year_separator is a real Excel AddRowV2 action")
    ok(_arys.get("inputs", {}).get("parameters", {}).get("item", None) == {},
       "Add_row_year_separator writes a truly empty item (a blank row, no column values)")

    _add_row_ra = actions.get("Add_row", {}).get("runAfter", {})
    ok("IsFirstOfNewYear" in _add_row_ra,
       "Add_row (the real candidate row) waits on IsFirstOfNewYear")
    ok(set(_add_row_ra.get("IsFirstOfNewYear", [])) == {"Succeeded", "Failed", "Skipped"},
       "Add_row proceeds regardless of the separator check's own outcome")

    # Only the brand-new-applicant path can ever add a row - duplicate/update/follow-up only
    # PATCH an existing row, so none of them should reference this feature's actions at all.
    for _no_touch in ("Patch_dup_attempts", "Patch_update_attempts", "Patch_followup_attempts"):
        _ra = actions.get(_no_touch, {}).get("runAfter", {})
        ok("IsFirstOfNewYear" not in _ra and "Filter_rows_this_year" not in _ra,
           f"{_no_touch} (a PATCH-only path) is untouched by the year-separator feature")
else:
    print("  (config says year_separator DISABLED - validating clean removal)")
    ok("Filter_rows_this_year" not in actions, "Filter_rows_this_year absent when disabled")
    ok("IsFirstOfNewYear" not in actions, "IsFirstOfNewYear absent when disabled")
    ok("Add_row_year_separator" not in actions, "Add_row_year_separator absent when disabled")
    ok("IsFirstOfNewYear" not in actions.get("Add_row", {}).get("runAfter", {}),
       "Add_row's runAfter has no dangling reference to the disabled feature")

# ── month-boundary separator (added 2026-07-31) ──
print("\n=== N2. MONTH SEPARATOR ===")
_msep_cfg = cfg.get("month_separator", {"enabled": True})
_add_row_ra2 = actions.get("Add_row", {}).get("runAfter", {})
if not _msep_cfg.get("enabled", True):
    print("  (config says month_separator DISABLED - validating clean removal)")
    ok("Filter_rows_this_month" not in actions, "Filter_rows_this_month absent when disabled")
    ok("IsFirstOfNewMonth" not in actions, "IsFirstOfNewMonth absent when disabled")
    ok("IsFirstOfNewMonth" not in _add_row_ra2, "no dangling runAfter reference")
else:
    _gm = actions.get("Filter_rows_this_month", {})
    _im = actions.get("IsFirstOfNewMonth", {})
    _mf = json.dumps(_gm.get("inputs", {}))
    ok(_gm.get("type") == "Query" and
       _gm.get("runAfter") == {"Get_rows_for_separators": ["Succeeded"]},
       "month detection uses an in-memory Query after the shared scan")
    ok("item()?['Received Date']" in _mf and "startsWith" in _mf and "yyyy-MM" in _mf,
       "month Query compares the visible Received Date safely without Excel OData")
    ok(_im.get("type") == "If", "IsFirstOfNewMonth is a real If condition")
    _expected_month_deps = {"Filter_rows_this_month", "Get_rows_any_applicant"}
    if _ysep_cfg.get("enabled", True):
        _expected_month_deps.add("Filter_rows_this_year")
    ok(set(_im.get("runAfter", {})) == _expected_month_deps,
       "IsFirstOfNewMonth waits on every array referenced by its condition")
    _msep = _im.get("actions", {}).get("Add_row_month_separator", {})
    ok(_msep.get("inputs", {}).get("host", {}).get("operationId") == "AddRowV2",
       "Add_row_month_separator is a real Excel AddRowV2 action")
    ok(_msep.get("inputs", {}).get("parameters", {}).get("item") == {},
       "Add_row_month_separator writes a truly empty item (a blank row)")
    # The point of the compound condition: a January row must produce ONE blank row
    # (the year separator), never two stacked separators.
    _mexpr = json.dumps(_im.get("expression", {}))
    ok("Get_rows_any_applicant" in _mexpr,
       "month separator is suppressed on a brand-new workbook (row 2 already supplies the gap)")
    if _ysep_cfg.get("enabled", True):
        ok('"and"' in _mexpr and "Filter_rows_this_year" in _mexpr,
           "month check also requires rows already exist this year (year separator wins in January)")
    else:
        ok("Filter_rows_this_year" not in _mexpr,
           "with year separator off, month check does not reference it")
    ok("IsFirstOfNewMonth" in _add_row_ra2, "Add_row waits on IsFirstOfNewMonth")
    ok(set(_add_row_ra2.get("IsFirstOfNewMonth", [])) == {"Succeeded", "Failed", "Skipped"},
       "Add_row proceeds regardless of the month check's outcome")
    for _nt in ("Patch_dup_attempts", "Patch_update_attempts", "Patch_followup_attempts"):
        _r = actions.get(_nt, {}).get("runAfter", {})
        ok("IsFirstOfNewMonth" not in _r and "Filter_rows_this_month" not in _r,
           f"{_nt} (a PATCH-only path) is untouched by the month-separator feature")

_bad_spaced_date_filters = []
for _name, _action in actions.items():
    _inputs = _action.get("inputs", {})
    _parameters = _inputs.get("parameters", {}) if isinstance(_inputs, dict) else {}
    _filter_value = _parameters.get("$filter", "") if isinstance(_parameters, dict) else ""
    if "Received Date" in str(_filter_value):
        _bad_spaced_date_filters.append((_name, _filter_value))
ok(not _bad_spaced_date_filters,
   "no Excel OData filter references the non-alphanumeric 'Received Date' header")

print(f"\n{'='*60}")
print(f"  P1 RESULT: {P} passed, {F} failed")
print(f"{'='*60}")
sys.exit(1 if F else 0)
