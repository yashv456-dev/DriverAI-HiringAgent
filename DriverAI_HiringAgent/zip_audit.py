"""
zip_audit.py -- Deep audit of the P1 ZIP before uploading to Power Automate.
Checks: ZIP integrity, actions, orphans, waits, test mode, config gaps, send actions.
"""
import json, zipfile, re, sys
from pathlib import Path

FLOW = Path("HiringAgent_P1/flow")
ZIP_PATH = FLOW / "DriverAI-Hiring-AutoReply-apply.zip"
CFG_PATH = FLOW / "flow_config.json"

P2 = Path("HiringAgent_P2")
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
print("\n=== 3. TEST MODE / LIVE MODE CHECK ===")
suppress = bool(cfg.get("test_mode", {}).get("suppress_emails", False))
ok(not suppress, "suppress_emails = false  (LIVE mode, real emails will send)")
test_stamps = blob.count("TEST-MODE (suppressed)")
ok(test_stamps == 0, f"No TEST-MODE stamps in definition (got {test_stamps})")
ok(blob.count("formatDateTime") > 0, "Real formatDateTime timestamp stamps present")

# ═══════════════════════════════════════════════════════════════════
print("\n=== 4. TRIGGER ===")
triggers = defn.get("triggers", {})
ok(len(triggers) == 1, f"Exactly 1 trigger (got {len(triggers)})")
trg = list(triggers.values())[0]
ok(trg.get("recurrence", {}).get("interval") == cfg["trigger"]["interval_min"],
   f"Poll interval = {cfg['trigger']['interval_min']} min (config-driven)")
ok(cfg["email"]["trigger_mailbox"] in blob, f"Trigger mailbox '{cfg['email']['trigger_mailbox']}' in definition")
ok(trg.get("splitOn") is not None, "splitOn = true (per-email fan-out, not batched)")
limits = trg.get("runtimeConfiguration", {}).get("concurrency", {})
op_opts = trg.get("operationOptions", "") or ""
ok(limits.get("runs", 0) == 1 or "Single" in op_opts or "1" in json.dumps(limits),
   "Concurrency = 1 (sequential, no double-row writes)")

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
print("\n=== 8. SEND ACTIONS (all live, no suppressed no-ops) ===")
send_actions = {n: a for n, a in actions.items()
                if a.get("type") == "OpenApiConnection"
                and "Mail" in json.dumps(a.get("inputs", {}))}
ok(len(send_actions) >= 7, f"At least 7 send-mail actions present (got {len(send_actions)})")
expected_sends = [
    "Send_acknowledgment", "Send_duplicate_notice", "Send_CV_request",
    "Send_wrong_format", "Send_update_ack", "Send_noted_reply", "Notify_failure"
]
for s in expected_sends:
    ok(s in actions, f"Send action present: {s}")
    if s in actions:
        a_type = actions[s].get("type")
        ok(a_type == "OpenApiConnection", f"  {s} is real OpenApiConnection (not suppressed Compose) -- type={a_type}")

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
ok(len(sf["bad_senders"])           == 32,  f"bad_senders        = 32  (got {len(sf['bad_senders'])})")
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
                   "_comment_flowname", "_comment_testmode", "_comment_rules",
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
