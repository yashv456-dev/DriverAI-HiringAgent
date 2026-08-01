"""Verify the P1 ZIP is ready for silent live intake.

Applicant-facing email must be disabled, operational processing must remain live,
and the admin failure alert must remain enabled.
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

print("=== SILENT LIVE MODE CHECK ===")
send_applicant = bool(cfg['email'].get('send_applicant_emails', True))
send_admin = bool(cfg['email'].get('send_admin_failure_alerts', True))
ok(send_applicant is False, "Applicant email delivery is disabled")
ok(send_admin is True, "Admin failure alerts are enabled")

suppressed_stamps = blob.count('Suppressed (applicant email disabled)')
ok(suppressed_stamps >= 4,
   f"Suppression audit stamps baked into all Mail Sent patches (got {suppressed_stamps})")

print()
print("=== APPLICANT EMAIL ACTIONS (all must be suppressed Compose) ===")
applicant_send_names = [
    ('Send_acknowledgment',   'EMAIL1 - new applicant ack'),
    ('Send_duplicate_notice', 'EMAIL2 - duplicate within 90 days'),
    ('Send_CV_request',       'EMAIL3 - please attach your CV'),
    ('Send_wrong_format',     'EMAIL4 - wrong file format'),
    ('Send_update_ack',       'EMAIL5 - CV update received'),
    ('Send_noted_reply',      'EMAIL6 - follow-up noted'),
]
for name, desc in applicant_send_names:
    a = actions.get(name, {})
    t = a.get('type', 'MISSING')
    ok(t == 'Compose', f"{name} [{desc}] is suppressed  type={t}")
ok(blob.count('SharedMailboxSendEmailV2') == 0,
   "No applicant shared-mailbox send operation remains")

print()
print("=== ADMIN FAILURE ALERT (must remain live) ===")
notify = actions.get('Notify_failure', {})
ok(notify.get('type') == 'OpenApiConnection', "Notify_failure is a real connector action")
ok(cfg['email']['admin_email'] in json.dumps(notify),
   f"Notify_failure targets {cfg['email']['admin_email']}")
poll_notify = actions.get('Notify_poll_failure', {})
ok(poll_notify.get('type') == 'OpenApiConnection',
   "Notify_poll_failure alerts when the shared Inbox cannot be read")
ok(cfg['email']['admin_email'] in json.dumps(poll_notify),
   f"Notify_poll_failure targets {cfg['email']['admin_email']}")
for failure_move in ('Move_failed_to_archive_unread',):
    fm = actions.get(failure_move, {})
    ok(fm.get('inputs', {}).get('host', {}).get('operationId') == 'MoveV2',
       f"{failure_move} is a real mailbox move")
    ok(fm.get('inputs', {}).get('parameters', {}).get('folderPath') ==
       cfg['inbox_tidy'].get('failure_destination'),
       f"{failure_move} targets the portable failure destination")
    ok(fm.get('inputs', {}).get('retryPolicy', {}).get('type') == 'none',
       f"{failure_move} disables automatic replay")
    ok(fm.get('runAfter') ==
       {'Notify_failure': ['Succeeded', 'Failed', 'TimedOut']},
       f"{failure_move} cannot activate when the failure alert is skipped")

for move_name, move_action in actions.items():
    if (isinstance(move_action.get('inputs'), dict)
            and move_action['inputs'].get('host', {}).get('operationId') == 'MoveV2'):
        ok(move_action.get('inputs', {}).get('retryPolicy', {}).get('type') == 'none',
           f"{move_name} has retryPolicy=none")

move_v2 = {
    name: action for name, action in actions.items()
    if isinstance(action.get('inputs'), dict)
    and action['inputs'].get('host', {}).get('operationId') == 'MoveV2'
}
ok(len(move_v2) == 11, f"Exactly 11 terminal MoveV2 actions exist (got {len(move_v2)})")
for move_name, move_action in move_v2.items():
    params = move_action.get('inputs', {}).get('parameters', {})
    ok(params.get('messageId') == "@outputs('CurrentEmail')?['id']",
       f"{move_name} uses CurrentEmail.id")
    ok(params.get('mailboxAddress') == cfg['email']['trigger_mailbox'],
       f"{move_name} uses the configured shared mailbox")
    ok(params.get('folderPath') in {'Archive', 'Junk Email'},
       f"{move_name} has a portable well-known destination")
    downstream = [
        action for action in actions.values()
        if move_name in action.get('runAfter', {})
        and {'Failed', 'TimedOut'}.issubset(
            set(action.get('runAfter', {}).get(move_name, []))
        )
    ]
    ok(len(downstream) == 1,
       f"{move_name} has exactly one explicit failure/timeout continuation")

ok('Move_failed_to_recruiting_review' not in actions,
   "Legacy Recruiting Review failure move is absent")
ok(not any(str(a.get('inputs', {}).get('parameters', {}).get('folderPath', '')).startswith('AAMk')
           for a in move_v2.values()),
   "No tenant-specific Outlook folder ID is baked into MoveV2")

ok(actions.get('Finalize_processed_cleanup', {}).get('runAfter') ==
   {'Move_to_processed': ['Succeeded', 'Failed', 'TimedOut']},
   "Normal terminal move has handled cleanup outcomes")
ok(actions.get('Finalize_failed_cleanup', {}).get('runAfter') ==
   {'Move_failed_to_archive_unread': ['Succeeded', 'Failed', 'TimedOut']},
   "Failure terminal move has handled cleanup outcomes")

def action_depths(action_map, level=0):
    for action_name, action in action_map.items():
        yield action_name, level
        child_actions = action.get('actions')
        if isinstance(child_actions, dict):
            yield from action_depths(child_actions, level + 1)
        else_actions = action.get('else', {}).get('actions')
        if isinstance(else_actions, dict):
            yield from action_depths(else_actions, level + 1)
        for case in action.get('cases', {}).values():
            case_actions = case.get('actions')
            if isinstance(case_actions, dict):
                yield from action_depths(case_actions, level + 1)
        default_actions = action.get('default', {}).get('actions')
        if isinstance(default_actions, dict):
            yield from action_depths(default_actions, level + 1)

depths = list(action_depths(defn['actions']))
max_depth = max(level for _, level in depths)
ok(max_depth == 8 and all(level <= 8 for _, level in depths),
   f"Maximum action nesting is level {max_depth} (Power Automate limit: 8)")
ok('CONFIG' in defn['actions'] and 'CurrentEmail' in defn['actions'],
   "Unread guard and processing graph are flat top-level siblings")

print()
print("=== TRIGGER + MAILBOX ===")
triggers = defn.get('triggers', {})
trg = list(triggers.values())[0]
interval = trg.get('recurrence', {}).get('interval')
ok(cfg['email']['trigger_mailbox'] in blob,
   f"Trigger mailbox baked in: {cfg['email']['trigger_mailbox']}")
ok(interval == cfg['trigger']['interval_min'],
   f"Polls every {interval} min")
ok(trg.get('type') == 'Recurrence', "Scheduled Recurrence trigger drains existing unread mail")
ok(trg.get('splitOn') is None, "No new-arrival splitOn watermark remains")
limits = trg.get('runtimeConfiguration', {}).get('concurrency', {})
ok(limits.get('runs', 0) == 1, "Concurrency = 1 (no double-row race conditions)")

print()
print("=== UNREAD SHARED-INBOX POLL ===")
poll = actions.get('Get_unread_emails', {})
poll_params = poll.get('inputs', {}).get('parameters', {})
ok(poll.get('inputs', {}).get('host', {}).get('operationId') == 'GetEmailsV3',
   "Get_unread_emails uses Outlook Get emails (V3)")
ok(poll_params.get('mailboxAddress') == cfg['email']['trigger_mailbox'],
   "Poll targets apply@driverai.io")
ok(poll_params.get('folderPath') == 'Inbox' and poll_params.get('fetchOnlyUnread') is True,
   "Poll selects only unread Inbox messages")
ok(poll_params.get('includeAttachments') is True,
   "Poll includes attachment content for resume processing")
ok(poll_params.get('top') == cfg['trigger'].get('unread_per_run') == 1,
   "One unread message is processed per scheduled run")
ok('triggerOutputs' not in blob and 'triggerBody' not in blob,
   "Processing graph reads CurrentEmail instead of a new-arrival trigger payload")

separator_scan = actions.get('Get_rows_for_separators', {})
separator_inputs = separator_scan.get('inputs', {}) if isinstance(separator_scan.get('inputs'), dict) else {}
separator_params = separator_inputs.get('parameters', {})
separator_paging = separator_scan.get('runtimeConfiguration', {}).get('paginationPolicy', {})
ok(separator_inputs.get('host', {}).get('operationId') == 'GetItems' and '$filter' not in separator_params,
   "Separator scan avoids invalid OData filtering on the spaced 'Received Date' column")
ok(separator_params.get('$top') == 5000 and separator_paging.get('minimumItemCount') == 5000,
   "Separator scan has 5000-row pagination")
ok(actions.get('Filter_rows_this_year', {}).get('type') == 'Query' and
   actions.get('Filter_rows_this_month', {}).get('type') == 'Query',
   "Year/month separators use in-flow Query actions")

print()
print("=== PATCH ACTIONS (suppression recorded accurately) ===")
for patch, send in [('Patch_mail_sent',          'Send_acknowledgment'),
                    ('Patch_dup_attempts',        'Send_duplicate_notice'),
                    ('Patch_update_attempts',     'Send_update_ack'),
                    ('Patch_followup_attempts',   'Send_noted_reply')]:
    pa = actions.get(patch, {})
    item = pa.get('inputs', {}).get('parameters', {}).get('item', {}) if isinstance(pa.get('inputs'), dict) else {}
    ms = str(item.get('Mail Sent', ''))
    ok('Mail Sent' in item,
       f"{patch} writes 'Mail Sent' stamp")
    ok('Suppressed (applicant email disabled)' in ms,
       f"{patch} records that applicant email was not sent")

print()
print("="*60)
print(f"  LIVE CHECK RESULT: {P} passed, {F} failed")
print("="*60)
if F == 0:
    print("  VERDICT: SILENT LIVE + READY. Intake runs; applicants are not emailed.")
else:
    print(f"  VERDICT: {F} issue(s) -- review FAILs above before uploading.")
    raise SystemExit(1)
