"""P1 end-to-end test — structural validation of the upload ZIP + routing simulation.

Reads the real Power Automate ZIP definition and flow_config.json (no live tenant needed).
Run from any directory:
    python HiringAgent_P1/test_p1.py
"""
import base64, io, json, re, sys, zipfile
from pathlib import Path
from openpyxl import load_workbook
sys.stdout.reconfigure(encoding="utf-8")

# Resolved relative to this file (fixed 2026-08-11). These used to be absolute paths under
# a specific developer's home directory, so the suite only ran on that one machine - on any
# other checkout it died at import before a single assertion. FLOW is always the sibling
# 'flow' directory; P2 is found by walking up until a directory holding HiringAgent_P2
# turns up, which keeps this working whether the tree is nested one level or two (the client
# zip extracts as HiringAgent_P1/HiringAgent_P1/...).
_HERE = Path(__file__).resolve().parent
FLOW = str(_HERE / "flow")

def _find_p2(start: Path) -> str:
    for base in (start, *start.parents):
        cand = base / "HiringAgent_P2"
        if cand.is_dir():
            return str(cand)
    raise SystemExit(
        "Could not locate HiringAgent_P2 near %s - test_p1.py cross-checks P1's column "
        "contract against P2's canonical column lists, so the P2 folder must be present." % start)

P2 = _find_p2(_HERE)

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

# Every user-facing date in the flow is expressed in cfg['timezone']['id'] (2026-08-27).
# Assertions below build the expected expression from that config value rather than
# hardcoding a UTC literal, so switching the zone - or setting it back to 'UTC' - keeps
# the suite meaningful instead of merely red.
TZ_ID = (cfg.get("timezone") or {}).get("id", "UTC")
TZ_OFFSET_H = int((cfg.get("timezone") or {}).get("utc_offset_hours", 0) or 0)


def tz(utc_expr: str) -> str:
    """The build's _local() - mirrored here so drift between the two shows up as a failure."""
    if TZ_OFFSET_H == 0:
        return utc_expr
    # Mirrors build_zip.py::_local EXACTLY. addHours(), not convertFromUtc() - the latter
    # rejects the Outlook connector's timestamp at runtime ("expects its first parameter to
    # be a string that contains a UTC date time"), and normalising it through formatDateTime
    # strips the very Z/offset it needs. Only valid because Arizona never observes DST.
    return "addHours(%s,%d,'yyyy-MM-ddTHH:mm:ss')" % (utc_expr, TZ_OFFSET_H)


RECV = "outputs('CurrentEmail')?['receivedDateTime']"

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
ok(get_unread_params.get("folderPath") == "Inbox",
   "the poll reads the Inbox")
# The INBOX is the queue, not the read flag. Every terminal path moves the message out of
# the Inbox (Archive on success, Junk Email on spam, Archive-unread on hard failure), so
# "still in Inbox" is the accurate definition of "not yet handled". The read flag is not:
# simply OPENING a message in the shared mailbox used to remove it from the queue forever,
# silently, while the flow kept reporting Succeeded. Live case 2026-08-04: a real applicant
# (redacted-applicant@example.edu) - a clean application, PDF attached, zero spam matches - got no reply
# and no row purely because the mailbox was opened to read it.
ok(get_unread_params.get("fetchOnlyUnread") == cfg["trigger"].get("fetch_only_unread", False),
   "the read-flag filter follows trigger.fetch_only_unread")
ok(get_unread_params.get("fetchOnlyUnread") is False,
   "a message that was merely OPENED is still picked up - reading mail cannot strand it")
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

# ── B. COLUMN CONTRACT (table = 33 cols; Add_row writes ONLY the 11 P1-owned ones) ───
# The TABLE has 33 columns (defined by the workbook template header + P2's own schema additions
# since). Add_row only WRITES the 11 P1-owned columns - it does NOT list the other 22 at all (they
# stay blank cells that P2 fills later). This exactly matches the genuine tenant PA exports
# (archive/ne3, ne4) and minimizes the import-time dynamic-schema binding surface that was dropping
# fields in the designer.
# 30 -> 32 on 2026-08-21: 'Education Start Date'/'Education End Date' were added to P2's schema
# 2026-08-18 but this contract and the setup template were never updated to match.
# 32 -> 33 on 2026-09-04: 'Years Exp' added at position #11, directly after
# Country (client request - the header and slot match what the client created BY HAND on the
# live master/Rejected/results sheets, so P2 adopts that column instead of adding a second
# one beside it). P2-owned, so
# Add_row is untouched and the flow package does NOT need rebuilding or re-importing.
print("\n=== B. COLUMN CONTRACT (11-col write, 33-col table) ===")
item   = actions["Add_row"]["inputs"]["parameters"]["item"]
ok(actions["Add_row"].get("inputs", {}).get("retryPolicy", {}).get("type") == "none",
   "Add_row disables automatic replay of the non-idempotent Excel append")
cols   = list(item.keys())
p2cols = ycfg["columns"]
# Reverted to 11 on 2026-09-08. 'Resume URL' (the attachment manifest) and 'Resume Folder
# Path' were briefly added here; two live runs then failed at Add_row with the CV and the
# intake sidecar already written, so the row write was the only step that broke. P2 locates a
# CV without them - by the reference in the filename, and from the intake_*.json sidecar - so
# the columns bought nothing and cost every intake. Add_row stays exactly as the build makes it.
P1_OWNED = ["Application ID", "Received Date", "Last Updated Date", "Full Name", "Email", "Mail Subject",
            "Mail Body", "Status", "Has Resume", "Original Filename", "Application Updates"]
P2_OWNED = ["Category", "Phone", "Location", "Country", "Years Exp",
            "Current Skills", "Education", "Education Start Date", "Education End Date",
            "Looking For Role", "Suggested Role 1", "Suggested Role 2", "Suggested Role 3",
            "Portfolio 1", "Portfolio 2", "Portfolio 3"]
ok(len(cols) == 11, f"Add_row writes exactly the 11 P1-owned columns (got {len(cols)})")
ok(set(cols) == set(P1_OWNED), "Add_row's 11 keys == the P1-owned set")
ok(len(p2cols) == 33, f"P2 config.yaml COLUMNS = 33 (adds Years Exp, got {len(p2cols)})")
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
ok(set(cols).issubset(set(p2cols)), "Every Add_row column exists in the 33-col table schema (no orphan write)")
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


def _alert_body(name):
    """An alert's HTML body, or "" when it is absent or suppressed to a no-op Compose.

    A suppressed alert's inputs is a plain string, so chained .get() on it raised and the
    whole suite died whenever send_admin_failure_alerts was false (found 2026-09-14).
    """
    inputs = actions.get(name, {}).get("inputs")
    params = inputs.get("parameters") if isinstance(inputs, dict) else None
    return str(params.get("emailMessage/Body", "")) if isinstance(params, dict) else ""


def _patch_reaches(source, target):
    """True when `target` runs (transitively) after `source` through runAfter links."""
    seen, frontier = set(), [source]
    while frontier:
        cur = frontier.pop()
        for n, a in actions.items():
            if cur in (a.get("runAfter") or {}) and n not in seen:
                if n == target:
                    return True
                seen.add(n)
                frontier.append(n)
    return False


_suppress_b = not _send_applicant_mail
# An applicant must never be acknowledged for an application that was not recorded. The zip
# base shipped Send_acknowledgment on Add_row [Succeeded, Failed, Skipped], so an Excel write
# failure still mailed the candidate an Application ID with no row behind it - and left
# nothing for P2 to score. Fixed 2026-08-03; pinned in build_zip.py so the base cannot
# reintroduce it.
_ack = actions.get("Send_acknowledgment")
ok(_ack is not None, "Send_acknowledgment action exists")
if _ack is not None:
    ok(list(_ack.get("runAfter", {}).keys()) == ["Add_row"],
       "Send_acknowledgment depends on Add_row only")
    ok(_ack["runAfter"].get("Add_row") == ["Succeeded"],
       "the acknowledgment is sent ONLY when the candidate row was actually written")
    ok("Failed" not in _ack["runAfter"].get("Add_row", []),
       "a failed Excel row write can never still acknowledge the applicant")
    ok("Skipped" not in _ack["runAfter"].get("Add_row", []),
       "a skipped Add_row (upstream save timed out) cannot acknowledge either")

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
# 2026-09-08: the row is now picked by Latest_ref, the deterministic latest of the rows
# Get_rows_ref returned, rather than by first() of an unordered result set. The property
# under test is unchanged - the reference comes from the MATCHED ROW, never from the
# text the sender quoted.
ok("first(body('Latest_ref'))" in actions.get("AppRef_current", {}).get("inputs", ""),
   "AppRef_current is sourced from the matched row (Latest_ref), not the quoted text")
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
ok(p2cols[21] == "Status",           f"Status at position #22 (shifted +2 by Education Start/End Date, +1 more by Years Exp; got '{p2cols[21]}')")
# The two education-date columns sit immediately after 'Education' (added 2026-08-18). Pinned
# by position, not just presence: P2 writes them by name, but the setup template below is
# compared against this exact canonical order, so a silent reorder must fail here too.
ok(p2cols[9]  == "Country",              f"Country at position #10 (got '{p2cols[9]}')")
ok(p2cols[10] == "Years Exp",
   f"Years Exp directly after Country, position #11 (added 2026-09-04; got '{p2cols[10]}')")
ok(p2cols[11] == "Current Skills",       f"Current Skills at position #12 (got '{p2cols[11]}')")
ok(p2cols[12] == "Education",            f"Education at position #13 (got '{p2cols[12]}')")
ok(p2cols[13] == "Education Start Date", f"Education Start Date directly after Education, position #14 (got '{p2cols[13]}')")
ok(p2cols[14] == "Education End Date",   f"Education End Date at position #15 (got '{p2cols[14]}')")
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
   "retired columns absent from the 33-col table too")

# ── C. APPREF — unique per email ─────────────────────────────────────────────
print("\n=== C. APPREF (unique reference per email) ===")
ok(item.get("Application ID") == "@outputs('AppRef')",
   "Add_row Application ID = @outputs('AppRef')")
ar_src = json.dumps(actions["AppRef"]["inputs"])
# DETERMINISTIC since 2026-09-05: the tail is sliced out of the message's own Graph id, not
# minted by guid(). A retry of the same email must produce the SAME reference - with guid()
# it produced a new one, so the resume was saved under one ref and the row keyed to another.
ok("guid(" not in ar_src, "AppRef has no random tail - the same email must mint the same ref")
ok("AppRefSeed" in actions, "AppRefSeed exists (the message id, cleaned and padded)")
seed_src = json.dumps(actions["AppRefSeed"]["inputs"])
ok("CurrentEmail')?['id']" in seed_src,
   "AppRefSeed is built from the message id, the one value that is stable across retries")
for _strip in ("'-',''", "'_',''", "'=',''"):
    ok(_strip in seed_src,
       f"AppRefSeed strips base64url punctuation ({_strip})")
_pad = "0" * (int(cfg["appref"]["hex_length"]) + int(cfg["appref"].get("id_tail_skip", 2)))
ok(f"concat('{_pad}'" in seed_src,
   "AppRefSeed is left-padded, so the slice can never run off the front")
ok("outputs('AppRefSeed')" in ar_src, "AppRef slices that seed")
# The tail sits id_tail_skip characters from the END. Measured across 241 real messages in
# this mailbox: the first, middle and last 4 characters of an Outlook id are IDENTICAL on
# every message (mailbox guid + trailing 'AAA' padding), while the 4 at offset -6 are the
# per-message counter - 185 distinct values, and 0 collisions across every same-minute pair.
_off = int(cfg["appref"]["hex_length"]) + int(cfg["appref"].get("id_tail_skip", 2))
ok(f"sub(length(outputs('AppRefSeed')), {_off})" in ar_src,
   f"AppRef takes its tail {_off} characters from the end, past the low-entropy padding")
ok("receivedDateTime" in ar_src, "AppRef uses received timestamp (not utcNow)")
ok(cfg["appref"]["date_format"] in ar_src, "AppRef date format from config")
ok("HHmm" in ar_src, "AppRef includes time component (HHmm, minute precision)")
ok("HHmmss" not in ar_src, "AppRef time trimmed to minute precision (no seconds)")
ok(f"), {cfg['appref']['hex_length']})" in ar_src,
   "AppRef tail trimmed to the configured length")
ok("toUpper" in ar_src, "AppRef tail is uppercased (consistent round-trip)")
ok(actions["AppRef"]["runAfter"] == {"AppRefSeed": ["Succeeded"]},
   "AppRef runs after the seed it reads")

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
# Get_rows_ref is now a two-branch filter: allow-listed testers look the row up by the quoted
# Application ID, everyone else falls back to matching on sender email. The sender-matching
# branch must still be the identical shared _EMAIL_FILTER expression Get_rows uses - two
# hand-written copies could drift apart - so assert it is embedded verbatim rather than equal.
_ref_filter = actions["Get_rows_ref"]["inputs"]["parameters"].get("$filter", "")
_rows_filter = actions["Get_rows"]["inputs"]["parameters"].get("$filter", "")
ok(_rows_filter.lstrip("@") in _ref_filter,
   "Get_rows_ref still embeds the exact same sender-email filter Get_rows uses (shared constant)")
ok("Application ID eq" in _ref_filter,
   "Get_rows_ref can also look a row up by the quoted Application ID")
# 2026-08-26: the SAME conditional now guards Get_rows, the new-vs-duplicate gate. Keying that
# purely on sender address meant an internal forward was never seen as an existing candidate -
# it minted a fresh row under the FORWARDER's address with the real candidate's resume on it.
# Live case: yashv@driverai.io forwarded Saniya Sekhon's resume quoting APP-20260204-1433-3B16;
# P1 created APP-20260825-1505-F08E under yashv@ and left her real row at Application Updates 0.
_dup_filter = actions["Get_rows"]["inputs"]["parameters"].get("$filter", "")
ok("Application ID eq" in _dup_filter,
   "Get_rows (duplicate gate) can look a row up by the quoted Application ID")
ok("MatchAllowSender" in _dup_filter,
   "by-reference duplicate lookup is restricted to ALLOW-LISTED senders - an external "
   "sender must not be able to overwrite another candidate's row by quoting their ref")
ok("Email eq" in _dup_filter,
   "Get_rows still falls back to the sender-email filter when no ref was quoted")

# The 90-day window must NOT apply to a by-reference match: Saniya's row is dated 2026-02-04,
# so a 2026-08-25 forward quoting her ref is ~200 days old and would fail the window, falling
# through to Add_row - a new row again, just found by a different key.
import json as _json
_dupexpr = _json.dumps(actions["Is_duplicate"]["expression"])
ok('"or"' in _dupexpr,
   "Is_duplicate is an OR: by-reference match OR inside the date window")
ok("MatchAllowSender" in _dupexpr and "QuotedRef" in _dupexpr,
   "Is_duplicate treats an allow-listed quoted reference as decisive, bypassing the window")
# The window is anchored to THIS EMAIL'S OWN received date, not utcNow() (changed 2026-08-28).
# Anchoring to "now" silently disabled duplicate protection for any REPLAY of mail older than
# the window: the gate found the prior row, rejected it as too old, and fell through to Add_row.
# Live case: redacted-applicant2@example.edu had two physical copies of ONE message in the mailbox (identical
# internetMessageId, received 2026-05-08); replayed on 2026-08-28, 112 days later, each copy
# cleared the 90-day test on its own and minted a row (APP-20260508-1642-20A1 and -E766).
# Measured against the email's own date those copies are 0 days apart, so the second is
# correctly a duplicate. Live intake is unaffected (received ~= now).
ok("addDays(%s, -90)" % tz("outputs('CurrentEmail')?['receivedDateTime']") in _dupexpr,
   "the 90-day window is still enforced for ordinary address-matched duplicates, on the same "
   "clock the stored Received Date uses, and anchored to the email's own received date")
ok("addDays(%s, -90)" % tz("utcNow()") not in _dupexpr,
   "the duplicate window is NOT anchored to utcNow() - that disables dedupe on any replay "
   "of mail older than the window (redacted-applicant2 double-row, 2026-08-28)")
ok("body('MatchAllowSender')" in _ref_filter,
   "the Application ID branch is gated on the sender being allow-listed, so an applicant "
   "can never patch another applicant's row by quoting their reference")
ok(_ref_filter.startswith("@if("),
   "Get_rows_ref chooses between the two lookups at runtime, not at build time")

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
ok("stored_name" in cf_src, "Create_file uses the once-generated immutable attachment name")
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
ok("Resume URL" not in pu_item and "Resume URL" not in item,
   "Update path writes no manifest column (the intake sidecar carries the references)")
ok("Application Updates" in pu_item or "ApplicationUpdates" in json.dumps(pu_item),
   "Patch_update_attempts increments Application Updates")
ok("Last Updated Date" in pu_item and "receivedDateTime" in str(pu_item.get("Last Updated Date")),
   "Patch_update_attempts stamps Last Updated Date from the update email time")
# Update save OVERWRITES: files into the ORIGINAL dated folder AND reuses the original stored filename
cu_src = json.dumps(actions.get("Create_update_file", {}))
ok("stored_name" in cu_src, "Update saves under its own immutable filename")
ok("Original Filename" not in cu_src, "Update never reuses the previous CV filename")

# ── G2. NO-REFERENCE RESEND ALSO RE-QUEUES + OVERWRITES ─────────────────────
print("\n=== G2. NO-REFERENCE RESEND (duplicate path) RE-QUEUES + OVERWRITES ===")
ok("Create_dup_update_file" in actions, "Create_dup_update_file exists (resend-without-ref resume save)")
ok("FileRef_from_dup" in actions, "FileRef_from_dup exists (original AppID prefix for resend saves)")
pd      = actions.get("Patch_dup_attempts", {})
pd_item = pd.get("inputs", {}).get("parameters", {}).get("item", {})
pd_item = pd_item if isinstance(pd_item, dict) else {}
ok(pd_item.get("Status") == "New Email Received",
   "Patch_dup_attempts re-queues Status='New Email Received' (a resend without a quoted ref still gets rescored)")
ok("Resume URL" not in pd_item and "Resume URL" not in item,
   "Resend path writes no manifest column (the intake sidecar carries the references)")
ok("Application Updates" in pd_item, "Patch_dup_attempts still increments Application Updates")
ok("Last Updated Date" in pd_item and "receivedDateTime" in str(pd_item.get("Last Updated Date")),
   "Patch_dup_attempts stamps Last Updated Date from the resend email time")
cdu_src = json.dumps(actions.get("Create_dup_update_file", {}))
ok("CurrentEmail" in cdu_src and "receivedDateTime" in cdu_src, "Resend records its own dated folder in the manifest")
ok("stored_name" in cdu_src and "Original Filename" not in cdu_src, "Resend preserves the previous CV")

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
ok(len(sf["bad_senders"])          == 43, f"bad_senders    = 43 (got {len(sf['bad_senders'])})")
# 2026-08-24 audit: 11 outsourcing / bench-sales firm domains added after a full
# read of all 401 Archive messages (incl. extracted resume text) found 21 vendor
# mails filed as applications and 55 auto-replies sent to sales reps.
for _d in ("bacancy.com", "equestsolutions.net", "sigmasolve.com", "logicrays.com",
           "infilon.com", "quebecsol.com", "synapsetechservice.com",
           "itidoltechnologies.com", "atgcorp.com", "pushcam-solution.com",
           "fixelsmedia.com"):
    ok(_d in sf["bad_senders"], f"bad_senders includes vendor domain {_d!r}")
# teknotrain.com was briefly added on the false assumption that anything already
# sitting in Junk was a vendor. redacted-applicant3@example.com is a real applicant (16 yrs
# cloud data architecture) who attached her OWN resume, 'Renu Jaitly.docx', and
# wrote in the first person - the same shape as redacted-vendor@example.com.
# A corporate domain never means agency; only the resume's owner decides that.
ok("teknotrain.com" not in sf["bad_senders"],
   "bad_senders excludes 'teknotrain.com' (real applicant on an employer domain)")
# Regression guard for the 2026-07-31 false-positive fix: the sender gate is an
# unanchored contains() over the raw From header, so a bare role local-part like
# "marketing@" also matches a personal address that merely ends in it
# (redactedmarketing@example.com - a real CMO applicant, junked 3x). Any term
# re-added here must be safe as a substring, not just as a prefix.
ok("marketing@" not in sf["bad_senders"],
   "bad_senders excludes 'marketing@' (substring-matched a real applicant)")
for _t in ("alert", "invoice", "receipt", "payment", "survey"):
    ok(_t not in sf["bad_subjects"],
       f"bad_subjects excludes bare {_t!r} (matches real job titles)")
ok(len(sf["bad_subjects"])         == 27, f"bad_subjects   = 27 (got {len(sf['bad_subjects'])})")

# ── Narrowed 2026-08-24: two more bare tokens qualified, same rule as the 2026-07-31 pass ──
# 'bounce' matched any From containing it - a candidate at a @bounce.com address, or a personal
# address like michael.bounce@gmail.com - and 'digest' is a bare common noun in a subject.
# NOTE the near-miss: 'bounce@' and 'bounces@' look safely qualified but are NOT - they still
# match 'michael.bounce@gmail.com', because the local part ENDS with them. Only the VERP form
# ('bounce-') and the subdomain form ('bounces.') are safe as unanchored substrings. Standard
# RFC bounce senders are already covered by 'mailer-daemon' and 'postmaster' above, and ESP
# bounce traffic by its domain ('amazonses.com', 'sendgrid.net'), so nothing is lost.
for _unsafe in ("bounce", "bounce@", "bounces@"):
    ok(_unsafe not in sf["bad_senders"],
       f"bad_senders excludes {_unsafe!r} (still substring-matches a personal address)")
for _b in ("bounce-", "bounces."):
    ok(_b in sf["bad_senders"], f"bad_senders keeps qualified {_b!r} (VERP / subdomain shapes)")
ok("digest" not in sf["bad_subjects"],
   "bad_subjects excludes bare 'digest' (bare common noun)")
for _d in ("daily digest", "weekly digest", "monthly digest", "your digest"):
    ok(_d in sf["bad_subjects"], f"bad_subjects keeps qualified {_d!r}")
# The qualified forms must still junk what the bare token used to, and must NOT junk a person.
_bs = [s.lower() for s in sf["bad_senders"]]
for _real in ("bounces@sendgrid.net", "bounce-123_HTML@mailer.example.com",
              "no-reply@bounces.example.com", "MAILER-DAEMON@googlemail.com",
              "postmaster@outlook.com"):
    ok(any(b in _real.lower() for b in _bs), f"machine bounce sender still blocked: {_real}")
for _person in ("jane@bounce.com", "michael.bounce@gmail.com", "a.bounces@yahoo.com"):
    ok(not any(b in _person.lower() for b in _bs), f"a real person is NOT blocked: {_person}")
ok(len(sf["spam_phrases"])         == 30, f"spam_phrases   = 30 (got {len(sf['spam_phrases'])})")
ok(len(sf["offensive_phrases"])    == 40, f"offensive      = 40 (got {len(sf['offensive_phrases'])})")
ok(len(sf["malware_phrases"])      == 20, f"malware        = 20 (got {len(sf['malware_phrases'])}) - extensions/macros only, always checked")
ok(len(sf["foreign_scam_phrases"]) == 25, f"foreign_scam   = 25 (got {len(sf['foreign_scam_phrases'])})")
ok(len(sf["link_shortener_phrases"]) == 12,
   f"link_shortener_phrases = 12 (got {len(sf.get('link_shortener_phrases', []))}) - checked separately, bypassed when a real resume is attached")
total = (len(sf["spam_phrases"]) + len(sf["offensive_phrases"]) +
         len(sf["malware_phrases"]) + len(sf["foreign_scam_phrases"]))
ok(total == 115, f"total always-on spam gate phrases = 115 (got {total}) (+ 12 conditional link-shortener phrases)")
ok(not (set(sf["malware_phrases"]) & set(sf["link_shortener_phrases"])),
   "malware_phrases and link_shortener_phrases don't overlap (clean split, nothing lost/duplicated)")

# Every gate term is a BARE SUBSTRING - P1 has no regex and no word boundaries - so a
# term that hides inside an ordinary word junks a real applicant silently: no row, no
# reply, no alert. 'horny' shipped for months and matched 'thorny' (found 2026-09-11).
# These are sentences a real candidate could plausibly write.
_INNOCENT = (
    "We solved a thorny integration problem and shipped it on time.",
    "I led a class of twenty analysts through the migration.",
    "Scaled the platform to assess thousands of documents per hour.",
    "My title was Associate Consultant at a mid-size firm.",
    "Built a document classifier and a shipment tracker in Python.",
    "Experience with Scunthorpe Steel Works as a process engineer.",
)
_ALWAYS_ON = (sf["spam_phrases"] + sf["offensive_phrases"]
              + sf["malware_phrases"] + sf["foreign_scam_phrases"])
for _sentence in _INNOCENT:
    _hits = [t for t in _ALWAYS_ON if t.lower() in _sentence.lower()]
    ok(not _hits,
       f"an ordinary resume sentence is not junked by a substring term "
       f"({_sentence[:44]!r} -> {_hits or 'clean'})")
# ...while the spam those terms exist for is still caught.
for _spam, _why in (("horny singles waiting in your area", "adult spam"),
                    ("you have won the lottery, claim your prize", "scam"),
                    ("open the attached invoice.exe to proceed", "malware")):
    ok(any(t.lower() in _spam.lower() for t in _ALWAYS_ON),
       f"{_why} is still blocked ({_spam[:38]!r})")
ok("apply@driverai.io" in sf["bad_senders"],
   "apply@driverai.io in bad_senders (self-loop protection)")
ok(len(sf["vendor_solicitation_phrases"]) == 12,
   f"vendor_solicitation = 12 (got {len(sf['vendor_solicitation_phrases'])})")
# Regression guard for the 2026-08-24 replay over all 401 archived messages.
# Each term below matched genuine applicants and must never be re-added. The
# 'our ' family is the sharpest trap: contains() is unanchored, and the word
# "your" ends in "our", so "our team" fires on every "your team".
for _t in ("collaborat", "bench", "our team", "agency", "partner with",
           "our company", "we provide", "case study", "ready for relocation",
           "for your review and consideration", "professional development team"):
    ok(_t not in sf["vendor_solicitation_phrases"],
       f"vendor_solicitation excludes {_t!r} (matched real applicants in replay)")
ok(not any(p.startswith("our ") for p in sf["vendor_solicitation_phrases"]),
   "no vendor phrase starts with 'our ' ('your' ends in 'our' - unanchored match)")
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

# ── I2. INTERNAL-TESTER ALLOW-LIST (2026-08-24) ──────────────────────────────
# The allow-list's whole purpose is that an internal tester is never silently junked, so
# the safety-critical assertion is that EVERY ONE of the three spam gates carries the
# not-allow-listed clause. Only the Get_rows_ref lookup branch was covered before; if the
# bypass silently dropped out of one gate on a rebuild, nothing here would have caught it.
print("\n=== I2. INTERNAL-TESTER ALLOW-LIST (spam-gate bypass) ===")
ok("AllowSenders" in actions, "AllowSenders Compose exists")
ok("MatchAllowSender" in actions, "MatchAllowSender Query exists")
ok(actions.get("MatchAllowSender", {}).get("type") == "Query",
   "MatchAllowSender is a Filter-array Query (same pattern as MatchSender)")
ok(actions["MatchAllowSender"]["runAfter"].get("AllowSenders") == ["Succeeded"],
   "MatchAllowSender waits for AllowSenders")
ok("LowerFrom" in json.dumps(actions["MatchAllowSender"]["inputs"]),
   "the allow-list is matched against the lowercased From address, exactly like bad_senders")

_is_subject = actions["IsSystem"]["else"]["actions"]["IsSubject"]
_gate_exprs = {
    "IsSystem (sender screen)": json.dumps(actions["IsSystem"]["expression"]),
    "IsSubject (subject screen)": json.dumps(_is_subject["expression"]),
    "IsSpam (content screen)": is_spam_expr,
}
for _label, _expr in _gate_exprs.items():
    ok("MatchAllowSender" in _expr,
       f"{_label} is bypassed for an allow-listed sender (carries the not-allow-listed clause)")

# An allow-listed sender must not be able to be the trigger mailbox itself - that would
# defeat the sender-based self-loop guard and the flow would answer its own mail forever.
_allow = [str(a).strip().lower() for a in actions["AllowSenders"]["inputs"]]
_trigger_mb = cfg["email"]["trigger_mailbox"].lower()
ok(_trigger_mb not in _allow,
   f"the trigger mailbox ({_trigger_mb}) is NOT allow-listed (self-loop guard intact)")
ok(_trigger_mb in json.dumps(actions["BadSenders"]["inputs"]).lower(),
   "the trigger mailbox is still in bad_senders (the actual self-loop guard)")
ok(all("@" in a for a in _allow),
   f"every allow-list entry is a real address, not a bare token that could match broadly: {_allow}")

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

def _template_sheet_xml(sheet_name: str) -> str:
    """Raw worksheet XML for `sheet_name`, straight out of the .xlsx zip.

    openpyxl normalises a missing cell and an empty cell to the same None, so it cannot see
    the difference between a spacer row that exists and one that only exists in the table's
    declared range. The XML can.
    """
    with zipfile.ZipFile(tpl_path) as _z:
        _wb = _z.read("xl/workbook.xml").decode("utf-8")
        _rels = _z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        _m = re.search(r'<sheet[^>]*name="%s"[^>]*r:id="([^"]+)"' % re.escape(sheet_name), _wb)
        if _m is None:
            raise AssertionError(f"sheet {sheet_name!r} not found in template workbook.xml")
        # Attribute order inside <Relationship .../> is not fixed (this workbook writes
        # Type, Target, Id) - match the whole element by Id, then read Target out of it.
        _rel = next(r for r in re.findall(r"<Relationship\b[^>]*/>", _rels)
                    if 'Id="%s"' % _m.group(1) in r)
        _target = re.search(r'Target="([^"]+)"', _rel).group(1).lstrip("/")
        if not _target.startswith("xl/"):
            _target = "xl/" + _target
        return _z.read(_target).decode("utf-8")


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
    # ...and that the row PHYSICALLY EXISTS, which the check above cannot tell (2026-08-27).
    # openpyxl's ws.cell() returns None for a cell that is not in the sheet at all, so a
    # <row r="2"></row> carrying ZERO <c> elements satisfied "every cell is blank" vacuously.
    # That is exactly what this template drifted into when the schema grew to 32/31 columns:
    # the table ref still claimed A1:AF2, but the sheet held only the header, so the permanent
    # spacer P1's separator logic depends on ("the first-ever applicant uses the header
    # spacer", build_zip.py) was not actually there. Read the sheet XML directly - it is the
    # only view that distinguishes "blank cell" from "no cell".
    _sheet_xml = _template_sheet_xml(sheet_name)
    _row2 = re.search(r'<row r="2"[^>]*>.*?</row>', _sheet_xml, re.S)
    ok(_row2 is not None, f"setup template '{sheet_name}' has a physical row 2 element")
    _n_header = len(re.findall(r'<c r="[A-Z]+1"', _sheet_xml))
    _n_spacer = len(re.findall(r'<c ', _row2.group(0))) if _row2 else 0
    ok(_n_spacer == _n_header,
       f"setup template '{sheet_name}' spacer row has a real cell per column "
       f"({_n_spacer} cells vs {_n_header} header cells) - a cell-less <row/> makes the "
       f"'fully blank' check above pass vacuously")
    ok(ws.max_row == 2,
       f"setup template '{sheet_name}' is exactly header + spacer (max_row={ws.max_row})")

# ── T. LOCAL TIMEZONE (added 2026-08-27) ─────────────────────────────────────
# Outlook hands P1 receivedDateTime in UTC. Every date the flow DERIVES from it is
# user-facing - the Application ID, the two date columns, the Year/Month resume folder,
# the year/month separator boundaries - so all of them must be local wall-clock time.
# Arizona is UTC-7, so mail received before 07:00 UTC belongs to the PREVIOUS local day;
# on the 1st of a month that filed the resume in the wrong month folder entirely
# (a real applicant, APP-20260501-0641-7C95: 2026-05-01 06:41 UTC = 2026-04-30 23:41 local,
# filed under Candidate_Resumes/2026/May instead of .../April).
print("\n=== T. LOCAL TIMEZONE (dates are local wall-clock, not UTC) ===")
_defn_src = json.dumps(defn)

ok(TZ_ID != "", "flow_config.json declares timezone.id")
if TZ_ID != "UTC":
    # The Windows ID matters more than it looks. 'Mountain Standard Time' is DENVER and
    # observes DST (UTC-6 Mar-Nov); under it a real applicant is 2026-05-01 00:41 and stays in
    # May, so it does NOT fix the bug this section exists for. Arizona never shifts.
    ok(TZ_OFFSET_H == -7,
       f"utc_offset_hours is -7 (got {TZ_OFFSET_H}) - a FIXED offset is only exact for a "
       f"zone with no DST, which is why this must stay Arizona and not Denver")
    ok(TZ_ID == "US Mountain Standard Time",
       f"timezone.id is the Arizona (no-DST, UTC-7 year round) Windows ID (got {TZ_ID!r}); "
       f"'Mountain Standard Time' is Denver and would still put a 06:41 UTC May-1 email in May")
    ok(tz(RECV) in _defn_src,
       "the trigger's receivedDateTime is converted to local before any date is derived, "
       "in exactly the shape build_zip.py::_local emits")

# NOTHING may format a raw UTC value into a stored or displayed date.
# Every convertFromUtc() must receive an OFFSET-FREE timestamp. The Outlook connector
# returns receivedDateTime as '2026-05-01T06:41:25+00:00'; convertFromUtc() rejects the
# offset form and FAILS AT RUNTIME while importing perfectly cleanly. That killed two live
# runs on 2026-08-27 - AppRef is unconditional and first in its scope, so its failure took
# the whole branch with it and the alert could only say "Reference: n/a". formatDateTime()
# tolerates the offset, so normalising through it first is what makes the conversion safe.
ok("convertFromUtc(" not in _defn_src,
   "convertFromUtc() appears NOWHERE - it rejects this connector's timestamps at runtime "
   "(proven live 2026-08-27: AppRef failed on '2026-05-01T06:41:25.0000000')")
ok(_defn_src.count(f"addHours(") > 0,
   "the local-time shift is present in the built flow, via addHours()")
ok(f",{TZ_OFFSET_H},'yyyy-MM-ddTHH:mm:ss')" in _defn_src,
   f"the shift uses the configured offset ({TZ_OFFSET_H}h) and an explicit output format")

# formatDateTime(<raw>) is legitimate ONLY as the offset-stripping inner call of
# convertFromUtc(...). Any OTHER occurrence is a value being stored or displayed in UTC.
for _label, _raw in (("utcNow()", "formatDateTime(utcNow()"),
                     ("the raw receivedDateTime", f"formatDateTime({RECV},")):
    _all = _defn_src.count(_raw)
    _wrapped = 0   # addHours wraps the RAW value directly, so any bare format is a real bug
    ok(_all == _wrapped,
       f"no action formats {_label} outside the timezone conversion "
       f"({_all - _wrapped} bare use(s)) - that would store or display a UTC wall-clock")

# The four derived values, each checked against the SAME expression the build emits.
_local_recv = tz(RECV)
_row = actions["Add_row"]["inputs"]["parameters"]["item"]
ok(_row["Received Date"] == "@{formatDateTime(%s,'yyyy-MM-ddTHH:mm:ss')}" % _local_recv,
   "Received Date column is stored in local time")
ok(_row["Last Updated Date"] == "@{formatDateTime(%s,'yyyy-MM-ddTHH:mm:ss')}" % _local_recv,
   "Last Updated Date column is stored in local time")
ok("formatDateTime(%s,'%s')" % (tz("coalesce(%s, utcNow())" % RECV), cfg["appref"]["date_format"])
   in json.dumps(actions["AppRef"]["inputs"]),
   "the Application ID date comes from the local timestamp")

_create = actions["Create_file"]["inputs"]["parameters"]["folderPath"]
ok("receivedDateTime" in _create and "yyyy/MMMM" in _create and "addHours(" in _create,
   "Versioned CV folder derives year/month from the localized source-message date")

# ...and the counterpart that must NOT be converted. A resend/update files into the
# ORIGINAL application's folder, rebuilt from that row's stored Received Date - which this
# same build already wrote in local time. Converting it a second time would shift it another
# 7 hours and send the file to a folder P2 would never look in.
for _name, _src in (("Create_update_file", "Get_rows_ref"),
                    ("Create_dup_update_file", "Filter_submitted")):
    _act = actions.get(_name)
    if _act is None:
        continue
    _fp = _act["inputs"]["parameters"]["folderPath"]
    ok("convertFromUtc" not in _fp,
       f"{_name} folder path does NOT re-convert the row's stored Received Date "
       f"(it is already local - a second conversion would shift it again)")
    ok("receivedDateTime" in _fp and "yyyy/MMMM" in _fp,
       f"{_name} uses the source-message month; the manifest records the exact path")

# Both sides of every timestamp COMPARISON must be on one clock, or the offset silently
# becomes a 7-hour bias.
_lu = json.dumps(actions["Patch_update_attempts"]["inputs"]["parameters"]["item"])
if "Last Updated Date" in _lu:
    ok(_lu.count("addHours(") >= 1 if TZ_OFFSET_H != 0 else True,
       "the never-move-backward Last Updated comparison converts the incoming timestamp "
       "before comparing it against the stored (local) column")

# ── U. ATTACHMENT TRIAGE (resume kept, portfolio/cover letter dropped) ───────
# P1 used to save EVERY .pdf/.docx attachment, and the saved name is
# '<FirstLast>_<AppRef>.<ext>' - the extension is the only per-attachment component. So two
# attachments with DIFFERENT extensions both landed in the folder (live:
# sathishsravanakumar_APP-20260414-0041-9D05.pdf AND .docx), and two with the SAME extension
# resolved to one identical name so the second SILENTLY OVERWROTE the first (live:
# 'Gayuh Nurul Huda - CV - 2026.pdf' vs '... - Portfolio - 2026.pdf'). P2 then concatenated
# the text of every survivor before scoring, so a portfolio polluted the resume's own scores.
print("\n=== U. ATTACHMENT TRIAGE (portfolio/cover-letter exclusion + no silent overwrite) ===")
# rstrip, mirroring build_zip.py. A LEADING space is part of the term: ' cl.' is the
# separator-anchored form that must NOT collapse to the far broader bare 'cl.'. This
# used to .strip() like the build did, so both sides agreed on the wrong answer and the
# suite passed while the package shipped a wider filter than the config asked for.
_EXCL = [str(p).rstrip().lower()
         for p in (cfg.get("resume_attachments") or {}).get("exclude_name_phrases", [])
         if str(p).strip()]
ok(bool(_EXCL), f"flow_config declares resume_attachments.exclude_name_phrases ({len(_EXCL)})")
for _needed in ("portfolio", "cover letter"):
    ok(_needed in _EXCL, f"exclusion list covers {_needed!r}")

# Three actions, because a Query cannot fall back to its own input.
for _a in ("ResumeFilesAll", "ResumeFilesKept", "ResumeFiles"):
    ok(_a in actions and actions[_a].get("type") == "Query", f"{_a} exists and is a Query")

_all_where = actions["ResumeFilesAll"]["inputs"]["where"]
ok(".pdf" in _all_where and ".docx" in _all_where and "contentBytes" in _all_where,
   "ResumeFilesAll is the original unfiltered .pdf/.docx-with-bytes test")
ok("'.doc'" in _all_where,
   "legacy Word (.doc) is accepted - older installs still send it, and P2 reads it "
   "via extraction.py::_word97_text")
ok(_all_where == actions["ResumeFilesEarly"]["inputs"]["where"],
   "the early spam-gate probe and the real save loop use the SAME resume test - if the "
   "probe accepts a file the loop rejects, the mail is treated as an application that "
   "then saves nothing")
_kept_where = actions["ResumeFilesKept"]["inputs"]["where"]
ok(actions["ResumeFilesKept"]["inputs"]["from"] == "@body('ResumeFilesAll')",
   "ResumeFilesKept narrows ResumeFilesAll rather than re-reading the attachments")
# The build compiles each configured phrase into the package verbatim. It used to
# .strip() them, which silently turned the separator-anchored ' cl.' into the far
# broader 'cl.' - the config said one thing and the running flow did another, and
# nothing compared the two. This does.
_zip_excl = [_m for _m in re.findall(r"'([^']*)'", _kept_where) if _m != "name"]
ok(_zip_excl == _EXCL,
   f"every exclusion phrase reaches the package unaltered (config {_EXCL} -> zip {_zip_excl})")
for _p in _EXCL:
    ok("contains(toLower(item()?['name']), '%s')" % _p.replace("'", "''") in _kept_where,
       f"exclusion {_p!r} is applied to the attachment FILENAME, case-insensitively")
ok(_kept_where.startswith("@not(or("),
   "exclusions are OR'd then negated - any one match drops that attachment")

# THE SAFETY PROPERTY. Without this an applicant whose only file is named portfolio.pdf gets
# a 'please attach your resume' reply and no row - the exclusion would have invented a
# no-CV application out of a real one.
_chosen = actions["ResumeFiles"]["inputs"]["from"]
ok("greater(length(body('ResumeFilesKept')), 0)" in _chosen and
   "body('ResumeFilesAll')" in _chosen,
   "ResumeFiles falls back to the UNFILTERED list when the exclusion would leave nothing")
ok("body('ResumeFilesKept')" in _chosen.split("body('ResumeFilesAll')")[0],
   "...and prefers the filtered list whenever it has anything in it")

# Everything downstream must still read body('ResumeFiles'), untouched by this refactor.
# 2026-09-11: each save loop now iterates its OWN Select rather than one shared list, so
# the stored name can carry the tail of the row that branch is writing against instead of
# the tail of the delivering message. What this check protects is unchanged and is asserted
# on the property, not the action name: whatever list a loop iterates must ultimately read
# body('ResumeFiles'), so the portfolio/cover-letter triage still covers every save path.
_BRANCH_ROW_SOURCE = {
    "Save_resumes_to_SharePoint": "outputs('AppRef')",
    "Save_dup_update_resume": "Latest_submitted",
    "Save_update_resume": "Latest_ref",
}
for _loop, _expect_row in _BRANCH_ROW_SOURCE.items():
    if _loop not in actions:
        continue
    _src = str(actions[_loop].get("foreach", ""))
    _sel_name = _src[len("@body('"):-len("')")] if _src.startswith("@body('") else ""
    _sel = actions.get(_sel_name, {})
    ok(_sel.get("inputs", {}).get("from") == "@body('ResumeFiles')",
       f"{_loop} still iterates the triaged list, so the exclusion covers this path too")
    _name_expr = _sel.get("inputs", {}).get("select", {}).get("stored_name", "")
    ok(_expect_row in _name_expr,
       f"{_loop} names the CV after the row it writes to, not the delivering message")

# One resume per person: a resend/update must resolve to the SAME stored name as the
# original application, which is what makes the second CV replace the first instead of
# landing beside it. The dup and update branches therefore must NOT name off AppRef.
for _loop in ("Save_dup_update_resume", "Save_update_resume"):
    if _loop not in actions:
        continue
    _src = str(actions[_loop].get("foreach", ""))
    _sel_name = _src[len("@body('"):-len("')")] if _src.startswith("@body('") else ""
    _expr = actions.get(_sel_name, {}).get("inputs", {}).get("select", {}).get("stored_name", "")
    ok("coalesce(first(body(" in _expr,
       f"{_loop} falls back to AppRef only when its row lookup returned nothing")
ok(actions["ResumeNames"]["inputs"]["from"] == "@body('ResumeFiles')",
   "Original Filename records only the KEPT attachments, so P2's slot count matches the "
   "files that actually exist")

# Collision safety: GUIDs are evaluated once per attachment, before all three upload paths.
_stored = actions["StoredResumeFiles"]["inputs"]["select"]
# 2026-09-10: the version name is no longer guid(). It is
# '<sanitised original>_<AppRef tail>[_<length>].<ext>' - readable, and DETERMINISTIC so a
# re-poll of the same email rewrites the same file instead of minting a second copy, which
# guid() did because it re-evaluates per RUN. Uniqueness is still guaranteed: the ref tail
# separates messages, the length suffix separates same-named attachments inside one message.
_sn = _stored["stored_name"]
ok("outputs('AppRef')" in _sn,
   "the version name carries the application reference, so a later CV cannot overwrite an earlier one")
ok("greater(length(body('ResumeFiles')),1)" in _sn,
   "a second attachment in the same email gets its own discriminator")
ok("item()?['id']" in _sn,
   "the discriminator is the attachment id, so two same-named files of equal size stay distinct")
ok("length(coalesce(item()?['contentBytes']" not in _sn,
   "content length is NOT the discriminator (equal-size documents would collide)")
ok("guid()" not in _sn,
   "the name is deterministic, so re-polling one email cannot create a second copy of its CV")
for _bad in ('"', "*", ":", "<", ">", "?", "/", "|", "#", "%"):
    ok("'%s',''" % _bad in _sn, f"the SharePoint-illegal character {_bad} is stripped from the stored name")
ok("item()?['name']" in _stored["stored_name"], "Each attachment keeps its own extension, including updates")
ok(_stored["name"] == "@item()?['name']", "Original display filename is retained separately")
ok(actions["ResumeManifest"]["inputs"]["select"]["name"] == "@item()?['stored_name']",
   "The manifest and upload refer to the same precomputed filename")
for _n, _loop in (("Create_file", "Save_resumes_to_SharePoint"),
                  ("Create_dup_update_file", "Save_dup_update_resume"),
                  ("Create_update_file", "Save_update_resume")):
    ok(actions[_n]["inputs"]["parameters"]["name"] == "@items('%s')?['stored_name']" % _loop,
       f"{_n} never recomputes a name or overwrites a previous version")
ok(not any("Delete" in a.get("inputs", {}).get("host", {}).get("operationId", "") for a in actions.values() if isinstance(a.get("inputs"), dict)),
   "P1 has no delete action")

# ── V. LOGIC APPS FUNCTION ARITY (added 2026-08-27, after a live outage) ─────
# and() / or() in the Workflow Definition Language are BINARY - exactly two arguments.
# A flat or(a, b, c, ...) is accepted at IMPORT and then fails at RUNTIME, so the flow
# looks healthy in the designer and every run dies inside the offending action. That is
# exactly what happened: a 9-argument or() in ResumeFilesKept's filter took out a live
# run (no row, no resume file, no acknowledgment, "Reference: n/a" in the failure alert
# because AppRef never resolved, and the message quarantined unread to Archive).
# Every multi-term test in this flow is either a Query over the term list
# (MatchAllowSender, MatchSpam) or a right-folded chain of binary or() calls.
print("\n=== V. LOGIC APPS FUNCTION ARITY (and/or must be binary) ===")

def _max_arity(expr: str, fn: str) -> int:
    """Largest top-level argument count of any `fn(...)` call in `expr`."""
    worst = 0
    for m in re.finditer(r"(?<![A-Za-z0-9_])" + fn + r"\(", expr):
        # Quote-aware: a comma inside a '...' literal is DATA, not an argument separator.
        # Without this the check false-alarms on replace(x, ',', '') and on any if() whose
        # branches contain a quoted comma.
        i, depth, args, in_str = m.end(), 1, 1, False
        while i < len(expr) and depth > 0:
            ch = expr[i]
            if ch == "'":
                if in_str and i + 1 < len(expr) and expr[i + 1] == "'":
                    i += 2          # '' is an escaped quote inside a literal
                    continue
                in_str = not in_str
            elif not in_str:
                if ch == "(":
                    depth += 1
                elif ch == ")":
                    depth -= 1
                elif ch == "," and depth == 1:
                    args += 1
            i += 1
        worst = max(worst, args)
    return worst


def _min_arity(expr: str, fn: str) -> int:
    """Smallest top-level argument count of any `fn(...)` call in `expr`."""
    worst = 999
    found_any = False
    for m in re.finditer(r"(?<![A-Za-z0-9_])" + fn + r"\(", expr):
        i, depth, args, in_str = m.end(), 1, 1, False
        found_any = True
        while i < len(expr) and depth > 0:
            ch = expr[i]
            if ch == "'":
                if in_str and i + 1 < len(expr) and expr[i + 1] == "'":
                    i += 2          # '' is an escaped quote inside a literal
                    continue
                in_str = not in_str
            elif not in_str:
                if ch == "(":
                    depth += 1
                elif ch == ")":
                    depth -= 1
                elif ch == "," and depth == 1:
                    args += 1
            i += 1
        worst = min(worst, args)
    return worst if found_any else 2

_defn_all = json.dumps(defn)
for _fn in ("and", "or"):
    _n = _max_arity(_defn_all, _fn)
    ok(_n <= 2,
       f"no {_fn}() in the built flow takes more than 2 arguments (max found: {_n}) - "
       f"a flat multi-arg {_fn}() imports fine and then fails EVERY run at runtime")
    _min_n = _min_arity(_defn_all, _fn)
    ok(_min_n >= 2,
       f"no {_fn}() in the built flow takes fewer than 2 arguments (min found: {_min_n}) - "
       f"a unary {_fn}() fails at runtime in Power Automate")


# And specifically the action that caused the outage.
_kept = actions.get("ResumeFilesKept")
if _kept is not None:
    _w = _kept["inputs"]["where"]
    ok(_max_arity(_w, "or") <= 2,
       "ResumeFilesKept folds its exclusion terms into nested BINARY or() calls")
    ok(_w.count("contains(") == len(_EXCL),
       f"...and still tests every one of the {len(_EXCL)} configured exclusion phrases")


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
ok(len(all_move_v2) == 14,
   f"exactly 14 MoveV2 actions exist (10 outcome + 1 failure + 3 save-failure; got {len(all_move_v2)})")
# 2026-09-08: a resume save that fails must still take the message out of the Inbox, or the
# next poll re-reads it forever. One per save path, all landing unread in the failure folder.
for _sf in ("Move_new_save_failed_unread", "Move_dup_save_failed_unread",
            "Move_update_save_failed_unread"):
    ok(_sf in all_move_v2, f"{_sf} exists so a failed resume save cannot block the queue")
    ok(all_move_v2.get(_sf, {}).get("inputs", {}).get("parameters", {}).get("folderPath")
       == cfg["inbox_tidy"]["failure_destination"],
       f"{_sf} parks the message in the configured failure folder, unread")
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
# Since 2026-09-14 each move also has exactly one failure-only admin alert sibling
# (build_zip.py section 16e): a caught move used to leave the message in the Inbox, to be
# read again on the next poll, with no signal at all.
branch_moves = [name for name in all_moves if name != "Move_to_processed"]
ok(len(branch_moves) == 9, f"9 nested outcome moves require terminal handlers (got {len(branch_moves)})")
for move_name in branch_moves:
    dependents = [
        (name, action) for name, action in actions.items()
        if owner_of.get(name) == owner_of.get(move_name)
        and move_name in action.get("runAfter", {})
    ]
    terminates = [(n, a) for n, a in dependents if a.get("type") == "Terminate"]
    alert_name = "Notify_" + move_name.lower() + "_failed"
    others = [n for n, a in dependents if a.get("type") != "Terminate" and n != alert_name]
    ok(len(terminates) == 1 and not others,
       f"{move_name} has exactly one same-scope Terminate continuation (plus its alert only)")
    ok(any(n == alert_name for n, _ in dependents),
       f"{move_name} has its cleanup-failure alert {alert_name}")
    if terminates:
        ok(set(terminates[0][1].get("runAfter", {}).get(move_name, [])) ==
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
# vendor_solicitation_phrases belongs here: build_zip.py folds it into the single
# SpamPhrases list the flow actually screens on. Omitting it made the simulation screen
# 111 terms against the shipped flow's 123, so a staffing-agency pitch routed to EMAIL1
# here while the real flow junked it.
spam_all     = [s.lower() for s in (sf["spam_phrases"] + sf["offensive_phrases"] +
                                     sf["malware_phrases"] + sf["foreign_scam_phrases"] +
                                     sf.get("vendor_solicitation_phrases", []))]
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
    # Mirrors build_zip.py _RESUME_EXTENSIONS. Legacy .doc joined on 2026-09-11; this model
    # still said .pdf/.docx only, so it kept "proving" .doc was a wrong format.
    pdf_docx   = [a for a in atts if a.endswith((".pdf", ".docx", ".doc"))]
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
     {"from":"redactedmarketing@example.com","subject":"Chief Marketing Officer",
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

    # ── Vendor / staffing-agency solicitation (2026-08-01 live case) ──
    ("vendor pitch — bench resume with a real PDF attached",
     {"from":"sales@somevendor.com","subject":"Excellent developer available",
      "body":"we work as a development partner and have attached their resume for your project",
      "atts":["consultant_cv.pdf"]},
     "SILENT (spam"),
    ("vendor pitch — capability deck",
     {"from":"bd@agency.com","subject":"Partnership",
      "body":"sharing our company profile and corporate deck plus engagement models",
      "atts":["deck.pdf"]},
     "SILENT (spam"),
    ("real applicant saying 'we work as a team' is NOT a vendor pitch",
     {"from":"real@gmail.com","subject":"Application for engineer",
      "body":"in my last role we worked as a team of six. pfa my resume",
      "atts":["cv.pdf"]},
     "EMAIL1"),

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
    ("legacy Word (.doc) is a resume, not a wrong format",
     {"from":"amy@gmail.com","subject":"resume","body":"pfa","atts":["cv.doc"]},
     "EMAIL1"),
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

# ── O3. PATCH-FAILURE ALERTS (added 2026-08-03) ──────────────────────────────
# Notify_failure only watches HasResume. Every candidate-row PATCH is "handled" by whatever
# runs after it (their runAfter lists include Failed so inbox cleanup is unconditional and a
# poisoned message can never block the queue), so the enclosing If reports Succeeded, the
# branch ends on a Terminate whose runStatus is Succeeded, and a failed patch produced a
# GREEN run with no signal anywhere - the applicant's Application Updates counter silently
# never incremented, and on the dup/update paths the row was never re-queued for Phase 2.
# Each alert is a plain SIBLING wired on [Failed, TimedOut]: same nesting depth (the update
# and follow-up patches already sit at the PA ceiling of 8, so a CHILD would fail import),
# nothing depends on it, so it cannot alter any existing path.
print("\n=== O3. PATCH-FAILURE ALERTS (silent-failure guard) ===")
_PATCH_ALERTS_EXPECTED = {
    "Notify_mail_sent_failed":                 "Patch_mail_sent",
    "Notify_dup_attempts_failed":              "Patch_dup_attempts",
    "Notify_dup_attempts_noreply_failed":      "Patch_dup_attempts_noreply",
    "Notify_update_attempts_failed":           "Patch_update_attempts",
    "Notify_update_attempts_noreply_failed":   "Patch_update_attempts_noreply",
    "Notify_followup_attempts_failed":         "Patch_followup_attempts",
    "Notify_followup_attempts_noreply_failed": "Patch_followup_attempts_noreply",
}
_alert_depths = dict(_action_depths(defn.get("actions", {})))
ok(len(_PATCH_ALERTS_EXPECTED) == 7, "all 7 candidate-row PATCH actions are covered by an alert")

for _alert, _patch in _PATCH_ALERTS_EXPECTED.items():
    _a = actions.get(_alert)
    ok(_a is not None, f"{_alert} exists")
    if _a is None:
        continue
    if _send_admin_alerts:
        ok(_a.get("type") == "OpenApiConnection"
           and _a["inputs"]["host"]["operationId"] == "SendEmailV2",
           f"{_alert} is a real admin send")
        ok(_a["inputs"]["parameters"]["emailMessage/To"] == cfg["email"]["admin_email"],
           f"{_alert} targets the configured admin")
        # A send has no mutable-message-ID hazard (that is why MoveV2/MarkAsRead disable
        # retries), and with applicant mail off these alerts are the only live sends - so a
        # transient blip must not lose the only signal.
        ok(_a.get("runtimeConfiguration", {}).get("retryPolicy", {}).get("type") == "exponential",
           f"{_alert} retries on a transient send failure")
    else:
        ok(_a.get("type") == "Compose",
           f"{_alert} is suppressed to a no-op exactly as configured")
    ok(_a.get("runAfter") == {_patch: ["Failed", "TimedOut"]},
       f"{_alert} fires ONLY on {_patch} Failed/TimedOut")
    ok("Succeeded" not in str(_a.get("runAfter")),
       f"{_alert} never fires on a successful patch (no false alarm)")
    ok(_alert_depths.get(_alert) == _alert_depths.get(_patch),
       f"{_alert} is a SIBLING of {_patch} at depth {_alert_depths.get(_patch)} (adds no nesting)")

# Only Terminates may depend on a patch-failure alert, and only unconditionally (all four
# statuses): they WAIT for it so ending the run cannot cancel the send mid-flight (2026-09-14
# dry-run finding - Terminate_update/_noreply and Terminate_noted/_noreply used to race it).
# Anything else depending on an alert would make the alert load-bearing.
_alert_dependents = {n: a for n, a in actions.items()
                     if set(a.get("runAfter") or {}) & set(_PATCH_ALERTS_EXPECTED)}
ok(all(a.get("type") == "Terminate" for a in _alert_dependents.values()),
   f"only Terminates wait for a patch-failure alert (got {sorted(_alert_dependents)})")
for _n, _a in _alert_dependents.items():
    for _dep, _st in _a["runAfter"].items():
        if _dep in _PATCH_ALERTS_EXPECTED:
            ok(sorted(_st) == ["Failed", "Skipped", "Succeeded", "TimedOut"],
               f"{_n} waits for {_dep} unconditionally, so a failed alert cannot block it")
for _patch in _PATCH_ALERTS_EXPECTED.values():
    _alert = "Notify_" + _patch.replace("Patch_", "") + "_failed"
    _terms = [n for n, a in actions.items() if a.get("type") == "Terminate"
              and owner_of.get(n) == owner_of.get(_patch) and _patch_reaches(_patch, n)]
    for _t in _terms:
        ok(_alert in (actions[_t].get("runAfter") or {}),
           f"{_t} (downstream of {_patch}) waits for {_alert}, so the alert is never cancelled")

# ── O3b. INTAKE-WRITE FAILURE ALERTS (added 2026-08-24) ──────────────────────
# O3 covers every PATCH of an EXISTING row. The two writes that CREATE a candidate had no
# watcher at all, and their failure is strictly worse: a failed patch leaves a stale row, a
# failed Add_row leaves NO ROW ANYWHERE. It was silent because Mark_as_read/Move_to_processed
# hang off HasResume rather than Add_row, so cleanup is on a PARALLEL branch and files the
# mail as handled either way, while Send_acknowledgment (correctly gated on Add_row Succeeded,
# see section B) is skipped - no row, no reply, mail archived, run GREEN. The 2026-08-24
# mailbox audit found 11 applicants lost exactly this way, clustered 2-6 June 2026.
print("\n=== O3b. INTAKE-WRITE FAILURE ALERTS (applicant-lost guard) ===")
_WRITE_ALERTS_EXPECTED = {
    "Notify_add_row_failed":     "Add_row",
    "Notify_create_file_failed": "Save_resumes_to_SharePoint",
    # 2026-08-26 pre-import branch audit: Create_file only ever covered the NEW-applicant
    # path. The duplicate/resend and ref-quoted-update paths each have their OWN CreateFile,
    # and both were silent for the same reason the patch failures were - the sibling that
    # follows their Foreach (IsUnderReplyCap_dup / IsUnderReplyCap_update) lists Failed in
    # its runAfter and ABSORBS the failure, so the reply still goes out, the counters still
    # move, Status still resets to 'New Email Received', and P2 then re-scores the STALE file.
    # 2026-09-08 fix: these three used to sit INSIDE their Foreach watching the CreateFile
    # directly, which is exactly why the bug survived - a failure that a later action in the
    # same scope handles is absorbed, so the Foreach still reported Succeeded and the reply
    # branch ran anyway. The alert now watches the LOOP from outside it, which both keeps the
    # signal and lets the reply branch depend on the loop's real status.
    "Notify_create_dup_update_file_failed": "Save_dup_update_resume",
    "Notify_create_update_file_failed":     "Save_update_resume",
}
for _alert, _write in _WRITE_ALERTS_EXPECTED.items():
    _a = actions.get(_alert)
    ok(_a is not None, f"{_alert} exists (guards {_write})")
    if _a is None:
        continue
    if _send_admin_alerts:
        ok(_a.get("type") == "OpenApiConnection"
           and _a["inputs"]["host"]["operationId"] == "SendEmailV2",
           f"{_alert} is a real admin send")
        ok(_a["inputs"]["parameters"]["emailMessage/To"] == cfg["email"]["admin_email"],
           f"{_alert} targets the configured admin")
        ok(_a.get("runtimeConfiguration", {}).get("retryPolicy", {}).get("type") == "exponential",
           f"{_alert} retries on a transient send failure")
        # The body must tell the admin how to get the applicant back, not merely that
        # something broke - the message is still sitting in Archive and is recoverable.
        _body = _a["inputs"]["parameters"]["emailMessage/Body"]
        ok("Inbox" in _body and "unread" in _body,
           f"{_alert} body states the recovery step (move back to Inbox as unread)")
    else:
        ok(_a.get("type") == "Compose",
           f"{_alert} is suppressed to a no-op exactly as configured")
    # The three resume-save alerts watch their LOOP and include Skipped: a loop skipped
    # because the folder step before it timed out used to shelve the applicant with no
    # alert (2026-09-14 dry-run finding). Add_row still fires on Failed/TimedOut only.
    _want = (["Failed", "TimedOut"] if _write == "Add_row"
             else ["Failed", "TimedOut", "Skipped"])
    ok(_a.get("runAfter") == {_write: _want},
       f"{_alert} fires ONLY when {_write} does not complete ({'/'.join(_want)})")
    ok(_alert_depths.get(_alert) == _alert_depths.get(_write),
       f"{_alert} is a SIBLING of {_write} (adds no nesting)")

# The regression this whole section exists to prevent: every action that can lose an
# applicant outright must have SOME watcher on Failed/TimedOut.
_watched = set()
for _n, _a in actions.items():
    if not _n.lower().startswith("notify"):
        continue
    for _dep, _st in (_a.get("runAfter") or {}).items():
        if "Failed" in _st or "TimedOut" in _st:
            _watched.add(_dep)
for _critical in ("Add_row", "Save_resumes_to_SharePoint",
                  "Save_dup_update_resume", "Save_update_resume"):
    ok(_critical in _watched,
       f"{_critical} failure is watched by an alert (a silent failure loses the applicant)")
# The save-failure Terminates deliberately DO depend on their intake-write alert
# (2026-09-10). Both the Move and the Notify hang off the same save loop and run in
# parallel; a MoveV2 finishes well inside the time an Office365 send takes, so ending the
# run on the move alone cancelled the alert mid-send and lost the resume silently. Taking
# the dependency is safe ONLY while it stays unconditional - all four statuses accepted -
# so a failed alert still ends the run green instead of blocking it. Any other dependent,
# or any dependency narrowed to Succeeded, makes the alert load-bearing: that is the
# regression this check exists to catch.
_ALERT_WAITERS = {"Terminate_new_save_failed", "Terminate_dup_save_failed",
                  "Terminate_update_save_failed"}
_ALL_STATUSES = ["Failed", "Skipped", "Succeeded", "TimedOut"]
_write_alert_dependents = [n for n, a in actions.items()
                           if set(a.get("runAfter") or {}) & set(_WRITE_ALERTS_EXPECTED)]
_unexpected = sorted(set(_write_alert_dependents) - _ALERT_WAITERS)
ok(not _unexpected,
   f"only the save-failure terminates depend on an intake-write alert (got {_unexpected})")
for _w in sorted(_write_alert_dependents):
    for _dep, _st in (actions[_w].get("runAfter") or {}).items():
        if _dep in _WRITE_ALERTS_EXPECTED:
            ok(sorted(_st) == _ALL_STATUSES,
               f"{_w} waits for {_dep} unconditionally, so a failed alert cannot block it")

# Which patches are SWALLOWED (a non-alert sibling runs after them on Failed, so the enclosing
# If reports Succeeded and the run goes GREEN - the alert is the only signal) versus which
# PROPAGATE up to HasResume (run goes RED and Notify_failure fires too, making the alert a
# belt-and-braces duplicate). build_zip.py varies each alert's closing paragraph on exactly this
# distinction, so if the wiring ever changes the alert text would start lying. Lock it here.
_EXPECTED_SWALLOWED = {
    "Patch_mail_sent": True,
    "Patch_dup_attempts": True,
    "Patch_dup_attempts_noreply": True,    # Dup_noreply_patch_done absorbs it (2026-08-27)
    "Patch_update_attempts": True,
    "Patch_update_attempts_noreply": True,
    "Patch_followup_attempts": True,
    "Patch_followup_attempts_noreply": True,
}
# EVERY candidate-row patch must be swallowed. Rows and resumes are deleted deliberately
# while the flow polls every 60s (resetting a candidate to re-run them is routine here),
# so Get_rows can read a row that is gone by the time the patch runs. That must never
# turn the run red - the alert is the signal. Guarded as a RULE, not just per-action,
# so a future patch added without a catch sibling fails this suite immediately.
ok(all(_EXPECTED_SWALLOWED.values()),
   "no candidate-row patch propagates - a row deleted mid-run cannot fail the flow")

for _patch, _want_swallowed in _EXPECTED_SWALLOWED.items():
    _followers = [(n, a["runAfter"][_patch]) for n, a in actions.items()
                  if _patch in (a.get("runAfter") or {}) and not n.startswith("Notify_")]
    _got_swallowed = any("Failed" in _st for _n, _st in _followers)
    ok(_got_swallowed == _want_swallowed,
       f"{_patch} failure is "
       f"{'SWALLOWED (alert is the only signal)' if _want_swallowed else 'PROPAGATED (Notify_failure also fires)'}"
       f" - alert wording depends on this; followers={[n for n, _ in _followers]}")

# The wording itself must match the classification.
for _patch, _want_swallowed in _EXPECTED_SWALLOWED.items():
    _alert = "Notify_" + _patch.replace("Patch_", "") + "_failed"
    _body = _alert_body(_alert)
    if not _body:
        continue   # suppressed to Compose when admin alerts are off
    if _want_swallowed:
        ok("only signal" in _body and "propagates" not in _body,
           f"{_alert} tells the admin it is the ONLY signal")
    else:
        ok("propagates" in _body and "two emails" in _body,
           f"{_alert} warns the admin to expect a duplicate Notify_failure")
ok(max(_alert_depths.values()) <= 8,
   f"adding the alerts kept nesting within the PA limit of 8 (got {max(_alert_depths.values())})")

if _send_admin_alerts:
    for _n in ("Notify_failure", "Notify_poll_failure"):
        ok(actions.get(_n, {}).get("runtimeConfiguration", {})
           .get("retryPolicy", {}).get("type") == "exponential",
           f"{_n} retries (changed 2026-08-03; the no-retry rule is for MoveV2, not sends)")

print("\n=== P. YEAR-BOUNDARY SEPARATOR (blank row before the first new-applicant row/year) ===")
_ysep_cfg = cfg.get("year_separator", {"enabled": True})
_msep_cfg = cfg.get("month_separator", {"enabled": True})
_SEP_ON = _ysep_cfg.get("enabled", True) or _msep_cfg.get("enabled", True)

# Both OFF as of 2026-08-27: AddRowV2 can only APPEND, so "a blank before the first row of a
# new month" only holds if emails arrive in date order - which the Inbox does not guarantee
# and a backlog replay actively breaks. The live master ended up with four blanks, none at a
# real month boundary. Grouping needs the sheet REORDERED, so layout is owned by
# HiringAgent_P2/resort_candidate_sheets.py instead. Assert the flow is genuinely clean of
# separator machinery when it is off - a leftover action would put stray blanks back.
if not _SEP_ON:
    for _dead in ("Get_rows_any_applicant", "Get_rows_for_separators",
                  "Filter_rows_this_year", "Filter_rows_this_month",
                  "IsFirstOfNewYear", "IsFirstOfNewMonth",
                  "Add_row_year_separator", "Add_row_month_separator"):
        ok(_dead not in actions,
           f"separators are OFF, so {_dead} is not built (no stray blank rows appended)")
    ok(not any("separator" in str(k).lower() for k in (actions.get("Add_row", {})
                                                       .get("runAfter") or {})),
       "Add_row no longer waits on any separator check")
_any_rows = actions.get("Get_rows_any_applicant", {})
if _SEP_ON:
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
       "formatDateTime(%s,'yyyy')" % tz(RECV) in _fy_src,
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

# ── M3. DUPLICATE CHECK MUST NOT DEPEND ON SHEET ROW ORDER ──────────────────
# The 90-day window used to read last(Filter_submitted), which assumed the sheet is in P1's
# own append order (oldest -> newest). P2's resort_candidate_sheets.py re-sorts the LIVE sheet
# to newest-FIRST, so last() started returning the sender's OLDEST row: a returning applicant
# whose first contact was >90 days ago read as brand new (new row, fresh ack, caps bypassed),
# while all 12 other row expressions read first() and stayed correct. Nothing in this suite
# guarded the invariant, and the two trees have separate suites - so it went unnoticed.
print("\n=== M3. DUPLICATE CHECK IS INDEPENDENT OF ROW ORDER ===")

_dup_act = _find_action(defn, "Is_duplicate") if "_find_action" in dir() else None
if _dup_act is None:
    def _walk_all(o):
        if isinstance(o, dict):
            for k, v in o.items():
                if isinstance(v, dict) and ("expression" in v or "inputs" in v):
                    yield k, v
                yield from _walk_all(v)
        elif isinstance(o, list):
            for i in o:
                yield from _walk_all(i)
    _all_acts = dict(_walk_all(defn))
    _dup_act = _all_acts.get("Is_duplicate")
    _ticks_act = _all_acts.get("Submitted_ticks")

_dup_expr = json.dumps(_dup_act.get("expression", {}))
ok("last(" not in _dup_expr,
   "the duplicate window no longer reads last() (which is the OLDEST row on the live sheet)")
ok("max(" in _dup_expr,
   "the duplicate window takes the MAX received date, so any row order gives the same answer")
ok(_ticks_act is not None and _ticks_act.get("type") == "Select",
   "Submitted_ticks projects each matching row to its tick value")
ok("Submitted_ticks" in _dup_act.get("runAfter", {}),
   "Is_duplicate waits on the projection it reads")
ok("json('[0]')" in _dup_expr or "union(" in _dup_expr,
   "the tick collection is floored so a brand-new applicant never hits max() on an empty array")

# No expression anywhere may pick an END of a row array except the trigger's own single mail.
_ordering_offenders = []
for _n, _a in _all_acts.items():
    _blob = json.dumps(_a)
    for _arr in ("Filter_submitted", "Get_rows_ref", "Get_rows_for_separators"):
        if "last(body('%s')" % _arr in _blob:
            _ordering_offenders.append((_n, _arr))
ok(not _ordering_offenders,
   f"no action reads last() off a candidate-row array (found {_ordering_offenders})")

# Re-implementation of the REAL expression, proving order independence.
def _dup_window_any_order(rows, now=_now, days=dup_days):
    """max(ticks) over the array, floored at 0 - mirrors the shipped expression."""
    ticks = [r for r in rows] or []
    floor = _dt.datetime.min
    return bool(rows) and max(ticks + [floor]) >= now - _dt.timedelta(days=days)

_old_first = _now - _dt.timedelta(days=200)   # first application, long ago
_recent = _now - _dt.timedelta(days=14)       # re-applied two weeks ago
ok(_dup_window_any_order([_old_first, _recent]) is True,
   "append order (oldest first): a recent re-application is still a duplicate")
ok(_dup_window_any_order([_recent, _old_first]) is True,
   "newest-first order (the live sheet): SAME answer - this is the bug that was fixed")
ok(_dup_window_any_order([_old_first]) is False,
   "a sender whose only application is outside the window is not a duplicate")
ok(_dup_window_any_order([]) is False,
   "a brand-new sender with no rows is not a duplicate (and does not error)")

# ── O2. OPERATING POSTURE: LIVE APPLICANT MAIL ──────────────────────────────
# CHANGED 2026-09-05 on user instruction: applicant mail is OFF - silent live intake, i.e.
# process everything, contact nobody. This matches P2's posture (test_mode.suppress_emails
# is also true) so the whole back catalogue can be re-imported and re-scored without any
# candidate hearing from either half of the system. The 2026-08-04 stance was live applicant
# mail for end-to-end testing; restore it by setting email.send_applicant_emails back to
# true and re-running build_zip.py.
#
# NOTE the two systems suppress DIFFERENTLY, on purpose. P1 replaces the six Send_* actions
# with no-op Composes and stamps 'Suppressed (applicant email disabled) <ts>' - a marker that
# cannot be mistaken for a send. P2 keeps its send call and returns early, stamping
# 'TEST-MODE (suppressed) Sent <ts>'; since 2026-09-05 P2 treats its own marker as NOT
# contact, so those rows stay queued and are mailed for real on the first live run.
#
# This section pins whichever posture is committed and asserts the ZIP AGREES WITH IT, so the
# config and the built package can never drift apart unnoticed. Section O above already adapts
# to the flag; what matters here is that flipping it is a deliberate, tested edit rather than
# an accident, in EITHER direction.
_applicant_mail_on = cfg["email"]["send_applicant_emails"] is True
print(f"\n=== O2. OPERATING POSTURE — applicant mail is "
      f"{'ON (real candidates are contacted)' if _applicant_mail_on else 'OFF (silent intake)'} ===")

_posture_src = (Path(FLOW) / "flow_config.json").read_text(encoding="utf-8")
ok(f'"send_applicant_emails": {str(_applicant_mail_on).lower()}' in _posture_src,
   "the posture is committed in flow_config.json, not left to an untracked override")

if _applicant_mail_on:
    # Every applicant-facing send must be a REAL connector action, and the 'Mail Sent' stamps
    # must claim a genuine send - the suppressed build writes a distinct marker instead.
    for _send in ("Send_acknowledgment", "Send_duplicate_notice", "Send_CV_request",
                  "Send_wrong_format", "Send_update_ack", "Send_noted_reply"):
        ok(actions.get(_send, {}).get("type") == "OpenApiConnection",
           f"{_send} is a REAL send (applicant mail is ON)")
    ok("Suppressed (applicant email disabled)" not in blob,
       "no suppression stamp remains in the built ZIP - stamps claim a real 'Sent <timestamp>'")
    ok(blob.count("SharedMailboxSendEmailV2") >= 6,
       f"at least 6 applicant send operations are live "
       f"(got {blob.count('SharedMailboxSendEmailV2')})")
else:
    ok("Suppressed (applicant email disabled)" in blob,
       "suppressed builds stamp a visibly-distinct marker instead of claiming 'Sent'")
    ok(blob.count("SharedMailboxSendEmailV2") == 0,
       "no applicant send operation remains in a suppressed build")

ok(cfg["email"]["send_admin_failure_alerts"] is True,
   "admin failure alerts stay ON regardless of the applicant-mail posture")
ok(cfg["trigger"]["interval_min"] >= 1,
   f"the trigger still polls for real (every {cfg['trigger']['interval_min']} min)")

print("")
print("=== R. 2026-09-08 REVIEW FIXES (regression locks) ===")
# Each block locks in one of the four findings from the package review, so a later edit
# cannot quietly restore the shape that caused the defect.

# R1. Privileged tester matching must be an ADDRESS COMPARISON. contains() matched
# yashv@driverai.io.example.net and any display name carrying that text - and a match
# switches off all three gates AND lets the caller select rows by Application ID.
_allow_where = actions.get("MatchAllowSender", {}).get("inputs", {}).get("where", "")
ok(_allow_where.startswith("@equals("),
   "MatchAllowSender compares the sender address for equality, not by substring")
ok("contains(outputs('LowerFrom'), item())" not in _allow_where,
   "the substring allow-list match is gone")
ok("split(outputs('LowerFrom'),'<')" in _allow_where,
   "the allow-list strips a display-name wrapper before comparing, as the Email column does")

# R2. retryPolicy must appear where Microsoft documents it. Asserting only the builder's own
# placement is what let an unproven location ship on every applicant-facing send.
for _n, _a in actions.items():
    _rc = (_a.get("runtimeConfiguration") or {}).get("retryPolicy")
    if not _rc:
        continue
    ok((_a.get("inputs") or {}).get("retryPolicy") == _rc,
       f"{_n} declares its retryPolicy under inputs, the documented location")
_live_sends = [n for n, a in actions.items()
               if isinstance(a.get("inputs"), dict)
               and a["inputs"].get("host", {}).get("operationId") == "SharedMailboxSendEmailV2"]
for _n in _live_sends:
    ok(actions[_n]["inputs"].get("retryPolicy", {}).get("type") == "none",
       f"{_n} cannot be replayed into a duplicate applicant email")

# R3. No success copy until the resume actually committed.
for _loop, _gated, _move, _term in (
        ("Save_resumes_to_SharePoint", "Add_row",
         "Move_new_save_failed_unread", "Terminate_new_save_failed"),
        ("Save_dup_update_resume", "IsUnderReplyCap_dup",
         "Move_dup_save_failed_unread", "Terminate_dup_save_failed"),
        ("Save_update_resume", "IsUnderReplyCap_update",
         "Move_update_save_failed_unread", "Terminate_update_save_failed")):
    ok(actions.get(_gated, {}).get("runAfter", {}).get(_loop) == ["Succeeded"],
       f"{_gated} runs only when {_loop} genuinely succeeded")
    ok(actions.get(_move, {}).get("runAfter") == {_loop: ["Failed", "TimedOut", "Skipped"]},
       f"{_move} catches a failed or skipped {_loop} so the message leaves the Inbox")
    ok(actions.get(_term, {}).get("inputs", {}).get("runStatus") == "Succeeded",
       f"{_term} ends the run cleanly instead of letting the reply path continue")
ok(actions.get("Send_acknowledgment", {}).get("runAfter", {}).get("Add_row") == ["Succeeded"],
   "the acknowledgment still requires a committed row, which now requires a committed file")

# R3b. The terminate must WAIT for the failure alert, not race it. Move and Notify both
# hang off the save loop in parallel, and a MoveV2 beats an Office365 send comfortably;
# terminating on the move alone cancelled the alert and lost the resume silently.
for _term, _alert in (("Terminate_new_save_failed", "Notify_create_file_failed"),
                      ("Terminate_dup_save_failed", "Notify_create_dup_update_file_failed"),
                      ("Terminate_update_save_failed", "Notify_create_update_file_failed")):
    _after = actions.get(_term, {}).get("runAfter", {})
    ok(_alert in _after,
       f"{_term} waits for {_alert} so the alert cannot be cancelled mid-send")
    ok(sorted(_after.get(_alert, [])) == ["Failed", "Skipped", "Succeeded", "TimedOut"],
       f"{_term} accepts every {_alert} outcome so a failed alert still ends the run green")

# R3b-2. The terminate must NEVER run when the save succeeded. Move is skipped on success,
# so accepting Skipped from the move terminates a successful intake before it can mark
# the email read or move it to Archive.
for _term, _move in (("Terminate_new_save_failed", "Move_new_save_failed_unread"),
                     ("Terminate_dup_save_failed", "Move_dup_save_failed_unread"),
                     ("Terminate_update_save_failed", "Move_update_save_failed_unread")):
    _after = actions.get(_term, {}).get("runAfter", {})
    ok("Skipped" not in _after.get(_move, []),
       f"{_term} must NEVER accept Skipped from {_move} (would kill successful runs)")
    ok(sorted(_after.get(_move, [])) == ["Failed", "Succeeded", "TimedOut"],
       f"{_term} waits for {_move} outcomes on failure only")


# R3c. The alert body rewrite is idempotent: a rebuild over an already-fixed definition
# must not append the recovery copy (or a second </div>) a second time.
for _alert in ("Notify_create_file_failed", "Notify_create_dup_update_file_failed",
               "Notify_create_update_file_failed"):
    _body = _alert_body(_alert)
    if not _send_admin_alerts:
        ok(not _body and actions.get(_alert, {}).get("type") == "Compose",
           f"{_alert} is fully suppressed when admin alerts are off (no stale live copy)")
        continue
    ok(_body.count("The resume was <strong>not saved</strong>") == 1,
       f"{_alert} carries the lost-resume copy exactly once (idempotent rewrite)")
    ok(_body.count("</div>") == 1,
       f"{_alert} body closes exactly one div (no duplicated append)")

# R3d. The alert must not contradict the save gate it sits behind (2026-09-14 review).
# The consequence paragraphs pre-dated require_saved_resume and still claimed Add_row ran
# and the applicant had been told - directly above R3c's "nothing downstream was allowed".
_STALE_CLAIMS = ("Add_row still runs", "already been told", "patched as though",
                 "patched as an update", "claims a resume that does not exist")
for _alert in ("Notify_create_file_failed", "Notify_create_dup_update_file_failed",
               "Notify_create_update_file_failed"):
    _body = _alert_body(_alert)
    if not _body:
        continue   # suppressed to Compose when admin alerts are off
    _stale = [c for c in _STALE_CLAIMS if c in _body]
    ok(not _stale, f"{_alert} makes no pre-gate claim that contradicts the save gate (got {_stale})")
_ms_body = _alert_body("Notify_mail_sent_failed")
if _ms_body:
    ok("Mail Sent stamp was NOT written" in _ms_body and "counter" not in _ms_body,
       "Notify_mail_sent_failed describes the stamp, not a counter Patch_mail_sent never touches")

# R3e. Notify_ignored_mail covers TWO kinds of mail. Has_application_keyword is
# and(no attachments, keyword), so its else branch also receives every email that has
# attachments but no PDF/Word resume - an application HasWrongFormat handles in parallel.
# Calling that "not an application ... sent no reply" was wrong; the wording must follow
# the attachment count, and the reply sentence must follow the applicant-mail setting.
_ig = actions.get("Notify_ignored_mail", {})
_ig_params = _ig.get("inputs", {}).get("parameters", {}) if isinstance(_ig.get("inputs"), dict) else {}
if _ig.get("type") == "OpenApiConnection":
    _ig_subject, _ig_body = str(_ig_params.get("emailMessage/Subject", "")), str(_ig_params.get("emailMessage/Body", ""))
    _att_test = "greater(length(coalesce(outputs('CurrentEmail')?['attachments'], json('[]'))), 0)"
    ok(_att_test in _ig_subject and "Resume not in PDF or Word" in _ig_subject
       and "Non-application mail archived" in _ig_subject,
       "Notify_ignored_mail subject distinguishes a wrong-format resume from non-application mail")
    ok(_att_test in _ig_body and "join(body('AttachNames')" in _ig_body,
       "Notify_ignored_mail body names the unsupported attachments when there are any")
    ok("no Application ID that matches a candidate row" in _ig_body
       and "no application reference" not in _ig_body,
       "Notify_ignored_mail no longer claims 'no reference' (an unmatched ref reaches this branch)")
    ok(("has not been told" in _ig_body) == (not _send_applicant_mail),
       "Notify_ignored_mail states whether a wrong-format sender was told, per applicant-mail config")
    ok(actions.get("HasWrongFormat", {}).get("runAfter", {}).get("Has_application_keyword")
       == ["Succeeded", "Failed", "Skipped"],
       "HasWrongFormat still evaluates after the keyword branch, so the two cases stay parallel")

# R3f. INBOX-CLEANUP AND DUPLICATE-LOOKUP ALERTS (build_zip.py 16e, 2026-09-14).
_CLEANUP_MOVES = ["Move_to_processed"] + [f"Move_to_processed_{s}" for s in
                  ("spam_sender", "spam_subject", "spam_body", "update", "update_noreply",
                   "update_cap", "followup", "followup_noreply", "followup_cap")]
_cleanup_alerts = {"Notify_" + m.lower() + "_failed": m for m in _CLEANUP_MOVES}
_ALL4 = ["Failed", "Skipped", "Succeeded", "TimedOut"]
_review_depths = dict(_action_depths(defn.get("actions", {})))
for _alert, _watched in list(_cleanup_alerts.items()) + [("Notify_get_rows_failed", "Get_rows")]:
    _a = actions.get(_alert)
    ok(_a is not None, f"{_alert} exists")
    if _a is None:
        continue
    if _send_admin_alerts:
        ok(_a.get("type") == "OpenApiConnection"
           and _a["inputs"]["host"]["operationId"] == "SendEmailV2"
           and _a["inputs"]["parameters"]["emailMessage/To"] == cfg["email"]["admin_email"],
           f"{_alert} is a real admin send to the configured admin")
        ok(_a.get("runtimeConfiguration", {}).get("retryPolicy", {}).get("type") == "exponential",
           f"{_alert} retries on a transient send failure")
    else:
        ok(_a.get("type") == "Compose", f"{_alert} is suppressed to a no-op exactly as configured")
    ok(_a.get("runAfter") == {_watched: ["Failed", "TimedOut"]},
       f"{_alert} fires ONLY when {_watched} fails or times out")
    ok(owner_of.get(_alert) == owner_of.get(_watched)
       and _review_depths.get(_alert) == _review_depths.get(_watched),
       f"{_alert} is a same-depth SIBLING of {_watched} (adds no nesting)")
    _waiters = {n: a["runAfter"][_alert] for n, a in actions.items()
                if _alert in (a.get("runAfter") or {})}
    ok(len(_waiters) == 1 and all(sorted(s) == _ALL4 for s in _waiters.values()),
       f"exactly one action waits for {_alert}, unconditionally (got {sorted(_waiters)})")
    if _waiters:
        _w = next(iter(_waiters))
        ok(actions[_w].get("type") in ("Terminate", "Compose"),
           f"{_alert} is caught by a Terminate or catch Compose, so a failed send ends green")
        if actions[_w].get("type") == "Terminate":
            ok(_watched in actions[_w].get("runAfter", {}),
               f"{_w} waits for both {_watched} and its alert, so the alert is never cancelled")
    if _watched.startswith("Move_to_processed_spam"):
        ok("AppRef" not in json.dumps(_a),
           f"{_alert} does not reference AppRef, which the spam gates run before")
ok(actions.get("Processed_cleanup_alert_done", {}).get("runAfter")
   == {"Notify_move_to_processed_failed": ["Succeeded", "Failed", "TimedOut", "Skipped"]},
   "top-level cleanup alert has its catch Compose")
ok(actions.get("Get_rows_alert_done", {}).get("runAfter")
   == {"Notify_get_rows_failed": ["Succeeded", "Failed", "TimedOut", "Skipped"]},
   "duplicate-lookup alert has its catch Compose")
ok(actions.get("Filter_submitted", {}).get("runAfter", {}).get("Get_rows")
   == ["Succeeded", "Failed", "Skipped"],
   "a failed duplicate lookup still does not block intake (the alert only reports it)")
ok(max(_review_depths.values()) <= 8,
   f"cleanup/lookup alerts kept nesting within the PA limit of 8 (got {max(_review_depths.values())})")

# R4. The row that gets patched must be the row the duplicate decision measured.
for _pick in ("Latest_submitted", "Latest_ref"):
    ok(_pick in actions, f"{_pick} exists so the patched row is the latest, not an arbitrary one")
ok("max(union(body('Submitted_ticks')"
   in actions.get("Latest_submitted", {}).get("inputs", {}).get("where", ""),
   "Latest_submitted selects on the same max(ticks) the duplicate decision uses")
ok("max(union(body('Ref_ticks')"
   in actions.get("Latest_ref", {}).get("inputs", {}).get("where", ""),
   "Latest_ref selects the most recent application on the reference path")
ok("first(body('Filter_submitted'))" not in blob,
   "no first() of the unordered duplicate result set survives in the package")
ok("first(body('Get_rows_ref')?['value'])" not in blob,
   "no first() of the unordered reference result set survives in the package")
ok(actions.get("IsKnownSender", {}).get("runAfter", {}).get("Latest_ref")
   == ["Succeeded", "Failed", "Skipped"],
   "a failed reference lookup still reaches the else branch that re-routes to ordinary intake")

print(f"\n{'='*60}")
print(f"  P1 RESULT: {P} passed, {F} failed")
print(f"{'='*60}")
sys.exit(1 if F else 0)
