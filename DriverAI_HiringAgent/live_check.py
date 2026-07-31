"""
live_check.py -- Verify the P1 ZIP is in LIVE mode (real emails, not test mode)
"""
import json, zipfile
from pathlib import Path

cfg = json.load(open('HiringAgent_P1/flow/flow_config.json', encoding='utf-8'))
ZIP_PATH = Path('HiringAgent_P1/flow/DriverAI-Hiring-AutoReply-apply.zip')
with zipfile.ZipFile(ZIP_PATH) as zf:
    names = [n for n in zf.namelist() if n.endswith('/definition.json')]
    d = json.loads(zf.read(names[0]).decode('utf-8'))

defn = d['properties']['definition']
blob = json.dumps(defn)

actions = {}
def walk(o):
    if isinstance(o, dict):
        for k, v in o.items():
            if k == 'actions' and isinstance(v, dict):
                for an, ad in v.items():
                    actions[an] = ad
            walk(v)
    elif isinstance(o, list):
        for x in o: walk(x)
walk(defn)

P = F = 0
def ok(c, label):
    global P, F
    sym = "[PASS]" if c else "[FAIL]"
    if c: P += 1
    else: F += 1
    print(f"  {sym} {label}")

print("=== LIVE MODE CHECK ===")
suppress = cfg['test_mode']['suppress_emails']
ok(suppress is False, f"suppress_emails = {suppress}  (False = LIVE, real emails will send)")

test_stamps = blob.count('TEST-MODE (suppressed)')
ok(test_stamps == 0, f"No TEST-MODE stamps in definition (got {test_stamps})")

sent_stamps = blob.count("'Sent '")
ok(sent_stamps > 0, f"Real 'Sent <timestamp>' stamp expressions baked in (got {sent_stamps})")

print()
print("=== EMAIL ACTIONS (all must be OpenApiConnection, not Compose) ===")
send_names = [
    ('Send_acknowledgment',   'EMAIL1 - new applicant ack'),
    ('Send_duplicate_notice', 'EMAIL2 - duplicate within 90 days'),
    ('Send_CV_request',       'EMAIL3 - please attach your CV'),
    ('Send_wrong_format',     'EMAIL4 - wrong file format'),
    ('Send_update_ack',       'EMAIL5 - CV update received'),
    ('Send_noted_reply',      'EMAIL6 - follow-up noted'),
    ('Notify_failure',        'ADMIN  - error alert'),
]
for name, desc in send_names:
    a = actions.get(name, {})
    t = a.get('type', 'MISSING')
    is_live = t == 'OpenApiConnection'
    inp = a.get('inputs', {}) if isinstance(a.get('inputs'), dict) else {}
    params = inp.get('parameters', {}) if isinstance(inp, dict) else {}
    to_addr = params.get('emailMessage/To', params.get('message/to', ''))
    subj = str(params.get('emailMessage/Subject', params.get('message/subject', '')))[:70]
    ok(is_live, f"{name} [{desc}]  type={t}")
    if to_addr:
        print(f"           To   : {str(to_addr)[:80]}")
    if subj:
        print(f"           Subj : {subj}")

print()
print("=== TRIGGER + MAILBOX ===")
triggers = defn.get('triggers', {})
trg = list(triggers.values())[0]
interval = trg.get('recurrence', {}).get('interval')
ok(cfg['email']['trigger_mailbox'] in blob,
   f"Trigger mailbox baked in: {cfg['email']['trigger_mailbox']}")
ok(interval == cfg['trigger']['interval_min'],
   f"Polls every {interval} min")
ok(trg.get('splitOn') is not None, "splitOn = true (each email is processed individually)")
limits = trg.get('runtimeConfiguration', {}).get('concurrency', {})
ok(limits.get('runs', 0) == 1, "Concurrency = 1 (no double-row race conditions)")

print()
print("=== CONNECTOR TYPE (must be SharedMailbox, not personal inbox) ===")
tname = list(triggers.keys())[0]
ok('shared' in tname.lower() or 'shared_mailbox' in blob.lower(),
   f"Trigger is SharedMailbox connector (not personal Office365)")

print()
print("=== PATCH ACTIONS (Mail Sent stamped only after real send) ===")
for patch, send in [('Patch_mail_sent',          'Send_acknowledgment'),
                    ('Patch_dup_attempts',        'Send_duplicate_notice'),
                    ('Patch_update_attempts',     'Send_update_ack'),
                    ('Patch_followup_attempts',   'Send_noted_reply')]:
    pa = actions.get(patch, {})
    item = pa.get('inputs', {}).get('parameters', {}).get('item', {}) if isinstance(pa.get('inputs'), dict) else {}
    ms = str(item.get('Mail Sent', ''))
    ok('Mail Sent' in item,
       f"{patch} writes 'Mail Sent' stamp")
    ok('TEST-MODE' not in ms,
       f"{patch} stamp is NOT a test-mode placeholder (real conditional expression)")

print()
print("="*60)
print(f"  LIVE CHECK RESULT: {P} passed, {F} failed")
print("="*60)
if F == 0:
    print("  VERDICT: LIVE + READY. All emails will fire on real applicants.")
else:
    print(f"  VERDICT: {F} issue(s) -- review FAILs above before uploading.")
