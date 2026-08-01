"""
zip_audit.py -- Deep audit of the P1 ZIP before uploading to Power Automate.
Checks: ZIP integrity, actions, orphans, waits, outbound-mail controls, config gaps, send actions.
"""
import json, zipfile, re, sys
from pathlib import Path

FLOW = Path("HiringAgent_P1/flow")
ZIP_PATH = FLOW / "DriverAI-Hiring-AutoReply-apply.zip"
CFG_PATH = FLOW / "flow_config.json"

P2 = Path("HiringAgent_App_P2")
sys.path.insert(0, str(P2))

# ── Load files ────────────────────────────────────────────────────────────────
with zipfile.ZipFile(ZIP_PATH) as zf:
    names = zf.namelist()
    zip_def_names = [n for n in names if n.endswith("/definition.json")]
    d = json.loads(zf.read(zip_def_names[0]).decode("utf-8"))
    zip_sizes = {n: zf.getinfo(n).file_size for n in names}

cfg  = json.load(open(CFG_PATH, encoding="utf-8"))
defn = d["properties"]["definition"]
blob = json.dumps(d)

# Walk all actions recursively
actions = {}
def walk(o):
    if isinstance(o, dict):
        for k, v in o.items():
            if k == "actions" and isinstance(v, dict):
                for an, ad in v.items():
                    actions[an] = ad
            walk(v)
    elif isinstance(o, list):
        for x in o:
            walk(x)
walk(defn)

def action_depths(action_map, level=0):
    for action_name, action in action_map.items():
        yield action_name, level
        child_actions = action.get("actions")
        if isinstance(child_actions, dict):
            yield from action_depths(child_actions, level + 1)
        else_actions = action.get("else", {}).get("actions")
        if isinstance(else_actions, dict):
            yield from action_depths(else_actions, level + 1)
        for case in action.get("cases", {}).values():
            case_actions = case.get("actions")
            if isinstance(case_actions, dict):
                yield from action_depths(case_actions, level + 1)
        default_actions = action.get("default", {}).get("actions")
        if isinstance(default_actions, dict):
            yield from action_depths(default_actions, level + 1)

P = F = 0
def ok(c, label):
    global P, F
    sym = "[PASS]" if c else "[FAIL]"
    if c: P += 1
    else: F += 1
    print(f"  {sym} {label}")

# ═══════════════════════════════════════════════════════════════════
print("\n=== 1. ZIP FILE STRUCTURE ===")
for n in sorted(names):
    sz = zip_sizes[n]
    print(f"  {sz:>8}b  {n}")

ok(len(zip_def_names) == 1, f"Exactly one definition.json in ZIP (got {len(zip_def_names)})")
ok(any("manifest.json" in n for n in names), "manifest.json present")
ok(any("apisMap.json" in n for n in names), "apisMap.json present")
ok(any("connectionsMap.json" in n for n in names), "connectionsMap.json present")

# ═══════════════════════════════════════════════════════════════════
print("\n=== 2. PACKAGE IDENTITY ===")
pkg_guid = "36d96cc9-25b1-4e7c-952c-378ba087f142"
ok(d.get("name") == pkg_guid, f"definition.name = package GUID {pkg_guid}")
ok(d.get("id") == f"/providers/Microsoft.Flow/flows/{pkg_guid}", "definition.id matches package GUID")
display = (d.get("properties", {}).get("displayName")
           or defn.get("triggers", {}).get(list(defn.get("triggers", {}).keys() or [""])[0], {}).get("metadata", {}).get("flowDisplayName", ""))
ok(cfg["flow_name"] in blob, f"flow_name '{cfg['flow_name']}' baked into definition")

# ═══════════════════════════════════════════════════════════════════
print("\n=== 3. OUTBOUND MAIL CONTROLS ===")
send_applicant_mail = bool(cfg.get("email", {}).get("send_applicant_emails", True))
send_admin_alerts = bool(cfg.get("email", {}).get("send_admin_failure_alerts", True))
ok(not send_applicant_mail, "Applicant email delivery is disabled (silent live intake)")
ok(send_admin_alerts, "Admin failure alerts remain enabled")
suppressed_stamps = blob.count("Suppressed (applicant email disabled)")
ok(suppressed_stamps >= 4,
   f"Suppression audit stamps present in all Mail Sent patches (got {suppressed_stamps})")
ok(blob.count("formatDateTime") > 0, "Timestamp expressions present")

# ═══════════════════════════════════════════════════════════════════
print("\n=== 4. TRIGGER ===")
triggers = defn.get("triggers", {})
ok(len(triggers) == 1, f"Exactly 1 trigger (got {len(triggers)})")
trg = list(triggers.values())[0]
ok(trg.get("recurrence", {}).get("interval") == cfg["trigger"]["interval_min"],
   f"Poll interval = {cfg['trigger']['interval_min']} min (config-driven)")
ok(cfg["email"]["trigger_mailbox"] in blob, f"Trigger mailbox '{cfg['email']['trigger_mailbox']}' in definition")
ok(trg.get("type") == "Recurrence", "Scheduled Recurrence trigger drains unread Inbox backlog")
ok(trg.get("splitOn") is None, "No new-arrival splitOn watermark remains")
limits = trg.get("runtimeConfiguration", {}).get("concurrency", {})
op_opts = trg.get("operationOptions", "") or ""
ok(limits.get("runs", 0) == 1 or "Single" in op_opts or "1" in json.dumps(limits),
   "Concurrency = 1 (sequential, no double-row writes)")
poll = actions.get("Get_unread_emails", {})
poll_params = poll.get("inputs", {}).get("parameters", {})
ok(poll.get("inputs", {}).get("host", {}).get("operationId") == "GetEmailsV3",
   "Unread poll uses supported Outlook Get emails (V3)")
ok(poll_params.get("mailboxAddress") == cfg["email"]["trigger_mailbox"],
   "Unread poll targets the configured shared mailbox")
ok(poll_params.get("folderPath") == "Inbox" and poll_params.get("fetchOnlyUnread") is True,
   "Unread poll reads only unread Inbox messages")
ok(poll_params.get("includeAttachments") is True,
   "Unread poll includes resume attachment content")
ok(poll_params.get("top") == cfg["trigger"].get("unread_per_run") == 1,
   "Exactly one unread message is processed per run")
ok("triggerOutputs" not in json.dumps(defn) and "triggerBody" not in json.dumps(defn),
   "Processing graph no longer depends on new-arrival trigger payload")

# ═══════════════════════════════════════════════════════════════════
print("\n=== 5. ACTIONS INVENTORY ===")
print(f"  Total actions in definition: {len(actions)}")

type_counts = {}
for a in actions.values():
    t = a.get("type", "?")
    type_counts[t] = type_counts.get(t, 0) + 1
print("  Action types:")
for t, c in sorted(type_counts.items(), key=lambda x: -x[1]):
    print(f"    {c:3}  {t}")

# ═══════════════════════════════════════════════════════════════════
print("\n=== 6. ORPHAN / UNREFERENCED ACTIONS ===")
# Top-level actions in Power Automate: entry points have no runAfter (e.g. CONFIG, BadSenders, etc.), others depend on preceding top-level actions
top_level_acts = defn.get("actions", {})
all_top_runafter = set()
for a in top_level_acts.values():
    all_top_runafter.update(a.get("runAfter", {}).keys())

# Top level orphans = actions defined at top level that have no runAfter AND are not referenced by any top-level action (excluding known initializers)
expected_roots = {"CONFIG", "BadSenders", "BadSubjects", "SpamPhrases", "LinkShortenerPhrases", 
                  "LowerFrom", "LowerSubject", "LowerBody", "AttachNames", "MatchSender"}
top_orphans = {n for n, a in top_level_acts.items() 
               if not a.get("runAfter") and n not in expected_roots and n not in all_top_runafter}

print(f"  Top-level actions: {len(top_level_acts)}")
for n in sorted(top_level_acts):
    print(f"    - {n}")

ok(len(top_orphans) == 0,
   f"No orphaned top-level actions -- {sorted(top_orphans) if top_orphans else 'clean'}")

# ═══════════════════════════════════════════════════════════════════
print("\n=== 7. WAIT / DELAY ACTIONS (should be zero) ===")
waits = [n for n, a in actions.items() if a.get("type") == "Wait"]
ok(len(waits) == 0, f"No Wait/Delay actions (got {len(waits)}) -- delays were removed 2026-07-04")
if waits:
    for n in waits:
        print(f"    WAIT: {n}  {actions[n].get('inputs',{}).get('interval',{})}")

# ═══════════════════════════════════════════════════════════════════
print("\n=== 8. SEND ACTIONS (applicant disabled, admin alert live) ===")
expected_applicant_sends = [
    "Send_acknowledgment", "Send_duplicate_notice", "Send_CV_request",
    "Send_wrong_format", "Send_update_ack", "Send_noted_reply"
]
for s in expected_applicant_sends:
    ok(s in actions, f"Send action present: {s}")
    if s in actions:
        a_type = actions[s].get("type")
        ok(a_type == "Compose", f"  {s} is suppressed Compose (applicant not contacted) -- type={a_type}")
ok(blob.count("SharedMailboxSendEmailV2") == 0,
   "No applicant shared-mailbox send operation remains")
ok("Notify_failure" in actions, "Admin failure action present: Notify_failure")
ok(actions.get("Notify_failure", {}).get("type") == "OpenApiConnection",
   "Notify_failure is a real OpenApiConnection")
ok(actions.get("Notify_poll_failure", {}).get("type") == "OpenApiConnection",
   "Notify_poll_failure is a real admin connector action")
ok(cfg["email"]["admin_email"] in json.dumps(actions.get("Notify_poll_failure", {})),
   "Notify_poll_failure targets the configured admin")
for failure_move in ("Move_failed_to_archive_unread",):
    fm = actions.get(failure_move, {})
    ok(fm.get("inputs", {}).get("host", {}).get("operationId") == "MoveV2",
       f"{failure_move} is a real mailbox move")
    ok(fm.get("inputs", {}).get("parameters", {}).get("folderPath") ==
       cfg["inbox_tidy"].get("failure_destination"),
       f"{failure_move} routes failures to the configured portable folder")
    ok(fm.get("inputs", {}).get("retryPolicy", {}).get("type") == "none",
       f"{failure_move} disables automatic replay of the non-idempotent move")
    ok(fm.get("runAfter") ==
       {"Notify_failure": ["Succeeded", "Failed", "TimedOut"]},
       f"{failure_move} runs only after an actual processing-failure alert attempt")
    ok("Skipped" not in fm.get("runAfter", {}).get("Notify_failure", []),
       f"{failure_move} cannot race the success-path move")

for move_name, move_action in actions.items():
    if (isinstance(move_action.get("inputs"), dict)
            and move_action["inputs"].get("host", {}).get("operationId") == "MoveV2"):
        ok(move_action.get("inputs", {}).get("retryPolicy", {}).get("type") == "none",
           f"{move_name} has retryPolicy=none")

move_v2 = {
    name: action for name, action in actions.items()
    if isinstance(action.get("inputs"), dict)
    and action["inputs"].get("host", {}).get("operationId") == "MoveV2"
}
ok(len(move_v2) == 11, f"Exactly 11 terminal MoveV2 actions exist (got {len(move_v2)})")
for move_name, move_action in move_v2.items():
    params = move_action.get("inputs", {}).get("parameters", {})
    ok(params.get("messageId") == "@outputs('CurrentEmail')?['id']",
       f"{move_name} uses CurrentEmail.id")
    ok(params.get("mailboxAddress") == cfg["email"]["trigger_mailbox"],
       f"{move_name} uses the configured shared mailbox")
    ok(params.get("folderPath") in {"Archive", "Junk Email"},
       f"{move_name} has a portable well-known destination")
    downstream = [
        action for action in actions.values()
        if move_name in action.get("runAfter", {})
        and {"Failed", "TimedOut"}.issubset(
            set(action.get("runAfter", {}).get(move_name, []))
        )
    ]
    ok(len(downstream) == 1,
       f"{move_name} has exactly one explicit failure/timeout continuation")

ok("Move_failed_to_recruiting_review" not in actions,
   "Legacy Recruiting Review failure move is absent")
ok(not any(str(a.get("inputs", {}).get("parameters", {}).get("folderPath", "")).startswith("AAMk")
           for a in move_v2.values()),
   "No tenant-specific Outlook folder ID is baked into MoveV2")

ok(actions.get("Finalize_processed_cleanup", {}).get("runAfter") ==
   {"Move_to_processed": ["Succeeded", "Failed", "TimedOut"]},
   "Normal terminal move is explicitly finalized on success/failure/timeout")
ok(actions.get("Finalize_failed_cleanup", {}).get("runAfter") ==
   {"Move_failed_to_archive_unread": ["Succeeded", "Failed", "TimedOut"]},
   "Failure-route terminal move is explicitly finalized on success/failure/timeout")

depths = list(action_depths(defn["actions"]))
max_depth = max(level for _, level in depths)
ok(max_depth == 8 and all(level <= 8 for _, level in depths),
   f"Maximum action nesting is level {max_depth} (Power Automate limit: 8)")
ok("CONFIG" in defn["actions"] and "CurrentEmail" in defn["actions"],
   "Unread guard does not wrap the legacy processing graph")

# ═══════════════════════════════════════════════════════════════════
print("\n=== 9. SHAREPOINT + EXCEL ADDRESSING ===")
ok(cfg["sharepoint"]["site"] in blob, f"SP site URL baked in: {cfg['sharepoint']['site']}")
ok(cfg["sharepoint"]["resumes_folder"].replace("/Shared Documents", "") in blob,
   f"Resumes folder path present in definition")
ok(cfg["excel"]["file"] in blob, f"Workbook path baked in: {cfg['excel']['file']}")
ok(cfg["excel"]["table"] in blob, f"Table name baked in: {cfg['excel']['table']}")
ok(cfg["excel"]["source"] in blob, "Excel source (site id triple) baked in")
ok(cfg["excel"]["drive"] in blob, "Excel drive id baked in")
# Ensure static literals (no expression tokens) in Excel addressing
for param_key in ["source", "drive", "file", "table"]:
    val = cfg["excel"][param_key]
    ok("@" not in val and "outputs(" not in val,
       f"excel.{param_key} is a static literal (no expression -- prevents DynamicParameterInputInvalid on import)")

# Excel Online only supports alphanumeric column names in OData Filter Query.
# "Received Date" therefore has to be filtered after the row scan.
separator_scan = actions.get("Get_rows_for_separators", {})
separator_inputs = separator_scan.get("inputs", {}) if isinstance(separator_scan.get("inputs"), dict) else {}
separator_params = separator_inputs.get("parameters", {})
separator_paging = separator_scan.get("runtimeConfiguration", {}).get("paginationPolicy", {})
ok(separator_inputs.get("host", {}).get("operationId") == "GetItems" and "$filter" not in separator_params,
   "Separator scan reads Excel rows without an invalid OData filter on 'Received Date'")
ok(separator_params.get("$top") == 5000 and separator_paging.get("minimumItemCount") == 5000,
   "Separator scan is paginated to 5000 rows")
ok(actions.get("Filter_rows_this_year", {}).get("type") == "Query" and
   actions.get("Filter_rows_this_month", {}).get("type") == "Query",
   "Year and month separators filter rows in-flow after the Excel scan")
received_date_odata = []
for action_name, action in actions.items():
    action_inputs = action.get("inputs")
    if not isinstance(action_inputs, dict):
        continue
    filter_query = action_inputs.get("parameters", {}).get("$filter")
    if isinstance(filter_query, str) and "Received Date" in filter_query:
        received_date_odata.append(action_name)
ok(not received_date_odata,
   f"No Excel OData Filter Query references spaced column 'Received Date' (found: {received_date_odata})")

# ═══════════════════════════════════════════════════════════════════
print("\n=== 10. RESUME FOLDER PATH ===")
ok("yyyy" in blob and "MMMM" in blob, "Dated path components yyyy/MMMM baked in")
ok("Wk" not in blob, "No week-level 'Wk' subfolder (removed 2026-07-08)")
ok(cfg["sharepoint"]["dated_resume_subfolders"] is True, "dated_resume_subfolders = true")
ensure_pairs = [
    ("Ensure_resume_folder",        "Save_resumes_to_SharePoint"),
    ("Ensure_update_resume_folder", "Save_update_resume"),
    ("Ensure_dup_resume_folder",    "Save_dup_update_resume"),
]
for ensure, save in ensure_pairs:
    ok(ensure in actions, f"{ensure} exists (auto-creates dated folder before every file save)")
    ok(save in actions,   f"{save} exists")

# ═══════════════════════════════════════════════════════════════════
print("\n=== 11. SPAM GATE COUNTS ===")
sf = cfg["spam_filters"]
ok(len(sf["bad_senders"])           == 31,  f"bad_senders        = 31  (got {len(sf['bad_senders'])})")
ok(len(sf["bad_subjects"])          == 24,  f"bad_subjects       = 24  (got {len(sf['bad_subjects'])})")
ok(len(sf["spam_phrases"])          == 30,  f"spam_phrases       = 30  (got {len(sf['spam_phrases'])})")
ok(len(sf["offensive_phrases"])     == 36,  f"offensive_phrases  = 36  (got {len(sf['offensive_phrases'])})")
ok(len(sf["malware_phrases"])       == 20,  f"malware_phrases    = 20  (got {len(sf['malware_phrases'])})")
ok(len(sf["foreign_scam_phrases"])  == 25,  f"foreign_scam       = 25  (got {len(sf['foreign_scam_phrases'])})")
ok(len(sf["link_shortener_phrases"])== 12,  f"link_shortener     = 12  (got {len(sf.get('link_shortener_phrases',[]))})")
total_always_on = sum(len(sf[k]) for k in ["spam_phrases","offensive_phrases","malware_phrases","foreign_scam_phrases"])
ok(total_always_on == 111, f"Total always-on phrases = 111 (got {total_always_on})")
ok("apply@driverai.io" in sf["bad_senders"], "Self-loop guard (apply@driverai.io) in bad_senders")
ok(len(sf["app_keywords"]) == 21, f"app_keywords = 21 (got {len(sf['app_keywords'])})")

# ═══════════════════════════════════════════════════════════════════
print("\n=== 12. RUNAFTER WIRING (no dangling deps) ===")
dangling = [(n, dep) for n, act in actions.items()
            for dep in act.get("runAfter", {}) if dep not in actions]
ok(len(dangling) == 0,
   f"No dangling runAfter references (got {len(dangling)}) -- all deps resolve to real actions")
if dangling:
    for n, dep in dangling:
        print(f"    DANGLING: {n} -> '{dep}' (not found)")

# ═══════════════════════════════════════════════════════════════════
print("\n=== 13. COLUMN CONTRACT (Add_row writes 11 P1 cols only) ===")
item = actions["Add_row"]["inputs"]["parameters"]["item"]
P1_OWNED = ["Application ID","Received Date","Last Updated Date","Full Name","Email",
            "Mail Subject","Mail Body","Status","Has Resume","Original Filename","Application Updates"]
P2_OWNED = ["Category","Phone","Location","Country","Current Skills","Education",
            "Looking For Role","Suggested Role 1","Suggested Role 2","Suggested Role 3",
            "Portfolio 1","Portfolio 2","Portfolio 3"]
ok(len(item) == 11, f"Add_row writes exactly 11 columns (got {len(item)})")
ok(set(item.keys()) == set(P1_OWNED), "Add_row key set == P1-owned set exactly")
for col in P2_OWNED:
    ok(col not in item, f"P2 column NOT in Add_row: '{col}'")
ok(item.get("Status") == "New Email Received", "Status = 'New Email Received'")
ok(item.get("Has Resume") == "Yes", "Has Resume = 'Yes'")
ok(item.get("Application Updates") == 0, "Application Updates = 0")

# ═══════════════════════════════════════════════════════════════════
print("\n=== 14. WORKBOOK SELF-HEAL REMOVAL (no destructive auto-overwrite) ===")
removed_acts = ["EnsureWb_Probe","WorkbookExists_NoOp","WorkbookTemplate_B64",
                "Create_workbook","Create_workbook_folder",
                "Wait_after_workbook_heal","Wait_after_workbook_folder"]
for r in removed_acts:
    ok(r not in actions, f"Destructive self-heal action absent: {r}")

# ═══════════════════════════════════════════════════════════════════
print("\n=== 15. INBOX TIDY (3 spam->Junk unread, 7 legit->Archive read) ===")
def move_dest(n):
    return actions.get(n, {}).get("inputs", {}).get("parameters", {}).get("folderPath")
all_moves = [n for n in actions if n.startswith("Move_to_processed")]
junk_moves = [n for n in all_moves if move_dest(n) == "Junk Email"]
arch_moves = [n for n in all_moves if move_dest(n) == "Archive"]
mark_reads  = [n for n in actions if n.startswith("Mark_as_read")]
ok(len(all_moves) == 10, f"10 total move actions (3 spam + 7 legit) -- got {len(all_moves)}")
ok(len(junk_moves) == 3, f"3 spam moves -> Junk Email -- got {len(junk_moves)}")
ok(len(arch_moves) == 7, f"7 legit moves -> Archive -- got {len(arch_moves)}")
ok(len(mark_reads)  == 7, f"7 mark-as-read (legit only, spam left unread) -- got {len(mark_reads)}")
for suf in ["spam_sender", "spam_subject", "spam_body"]:
    ok(f"Mark_as_read_{suf}" not in actions, f"spam '{suf}' has NO mark-as-read (left unread in Junk)")

# ═══════════════════════════════════════════════════════════════════
print("\n=== 16. APPREF FORMAT ===")
ar = json.dumps(actions.get("AppRef", {}).get("inputs", {}))
ok("yyyyMMdd" in ar, "AppRef uses yyyyMMdd date format")
ok("HHmm" in ar, "AppRef uses HHmm time (minute precision)")
ok("HHmmss" not in ar, "AppRef does NOT use HHmmss (no seconds - keeps refs short)")
ok("guid()" in ar or "guid(" in ar, "AppRef has guid() random tail")
ok("substring(guid(),0,4)" in ar, "AppRef hex tail = 4 chars")
ok("toUpper" in ar, "AppRef tail is uppercased")
ok(cfg["appref"]["detect_pattern"] in blob, f"detect_pattern '{cfg['appref']['detect_pattern']}' baked in")

# ═══════════════════════════════════════════════════════════════════
print("\n=== 17. CAP SYSTEM (business rules) ===")
rules = cfg["business_rules"]
ok(rules["duplicate_notice_max"] == 5, "duplicate_notice_max = 5")
ok(rules["update_resume_max"]    == 5, "update_resume_max    = 5")
ok(rules["followup_reply_max"]   == 5, "followup_reply_max   = 5")
ok(rules["reply_cap"]            == 3, "reply_cap = 3 (contacts 4-5 silent but still update row)")
for gate in ["IsUnderDupCap","IsUnderUpdateCap","IsUnderFollowupCap"]:
    ok(gate in actions, f"Overall cap gate present: {gate}")
for gate in ["IsUnderReplyCap_dup","IsUnderReplyCap_update","IsUnderReplyCap_followup"]:
    ok(gate in actions, f"Reply-cap gate present:   {gate}")

# ═══════════════════════════════════════════════════════════════════
print("\n=== 18. KNOWN UNUSED / CONFIG-ONLY FIELDS (expected gaps) ===")
# These are intentionally not in the flow (config keeps them for build tooling only)
unused_expected = ["_comment", "_comment_interval", "_comment_loopguard",
                   "_comment_flowname", "_comment_outbound_mail", "_comment_rules",
                   "_comment_dated", "_comment_libid", "_comment_selfheal",
                   "_comment_excel", "_comment_tidy", "_comment_appref"]
print(f"  {len(unused_expected)} config comment/doc fields (all start with '_', ignored by builder):")
for u in unused_expected:
    present = u in cfg
    print(f"    {'[OK doc]' if present else '[MISSING]'} {u}")
# Also: acknowledgment / request_cv email templates in config.yaml are dead (P1 uses build_zip, not config.yaml)
print("  Note: config.yaml 'acknowledgment'/'request_cv' templates are dead (P1 uses build_zip.py hardcoded bodies -- expected)")

# ═══════════════════════════════════════════════════════════════════
print("\n" + "="*60)
print(f"  ZIP AUDIT RESULT: {P} passed, {F} failed")
print("="*60)
if F == 0:
    print("  VERDICT: ZIP is CLEAN -- safe to upload to Power Automate")
else:
    print(f"  VERDICT: {F} issue(s) found -- review FAILs above before uploading")
    sys.exit(1)
