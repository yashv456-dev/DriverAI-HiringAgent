"""Correctness fixes for the four findings of the 2026-09-08 package review.

Applied last, over the fully assembled flow (after apply_reliability), so the whole
change set is one reviewable file instead of edits scattered through build_zip.py.
Every function is idempotent: re-running a build re-applies the same end state.

  1. exact_allow_list        - allow-listed testers matched by address, not substring
  2. retry_policy_placement  - retryPolicy written where Microsoft documents it
  3. require_saved_resume    - no success copy until the resume actually committed
  4. deterministic_row       - the latest application is the one that gets patched
"""

# ── 1. privileged tester matching ────────────────────────────────────────────
# MatchAllowSender used contains(LowerFrom, item()), so 'yashv@driverai.io.example.net'
# - or any display name carrying that text - matched. A match is not a spam decision: it
# switches OFF all three screening gates including the self-loop guard, AND switches
# Get_rows_ref from sender matching to Application ID matching, which lets the caller
# select any row in the workbook. Compare the extracted address for equality instead.
# LowerFrom is already toLower()'d, so this is the same normalisation _email_norm_expr
# applies to the stored Email cell and the $filter value.
_ALLOW_MATCH = (
    "@equals("
    "toLower(trim(if(contains(outputs('LowerFrom'),'<'),"
    "replace(last(split(outputs('LowerFrom'),'<')),'>',''),"
    "outputs('LowerFrom'))))"
    ", item())"
)


def _walk(action_map):
    """Yield every (name, action, owning dict) in the tree, at any depth."""
    for name, action in list(action_map.items()):
        yield name, action, action_map
        child = action.get("actions")
        if isinstance(child, dict):
            yield from _walk(child)
        for branch in ("else", "default"):
            sub = action.get(branch)
            if isinstance(sub, dict) and isinstance(sub.get("actions"), dict):
                yield from _walk(sub["actions"])
        for case in (action.get("cases") or {}).values():
            if isinstance(case.get("actions"), dict):
                yield from _walk(case["actions"])


def _find(action_map, wanted):
    for name, action, owner in _walk(action_map):
        if name == wanted:
            return action, owner
    return None, None


def _rewrite_expressions(node, old, new):
    """Replace a literal expression fragment everywhere in the tree."""
    if isinstance(node, dict):
        return {k: _rewrite_expressions(v, old, new) for k, v in node.items()}
    if isinstance(node, list):
        return [_rewrite_expressions(v, old, new) for v in node]
    if isinstance(node, str):
        return node.replace(old, new)
    return node


def exact_allow_list(actions):
    match, _ = _find(actions, "MatchAllowSender")
    if match is None:
        raise ValueError("MatchAllowSender missing - allow-list fix cannot be applied")
    match["inputs"]["where"] = _ALLOW_MATCH
    return match


def retry_policy_placement(actions):
    """Mirror every runtimeConfiguration.retryPolicy into inputs.retryPolicy.

    Microsoft's workflow definition language documents retryPolicy under `inputs`
    (learn.microsoft.com/azure/logic-apps/error-exception-handling), and this package
    already uses that location for all 11 MoveV2 actions. The 20 send actions wrote it
    under runtimeConfiguration only, so nothing proved the connector honoured 'none' -
    and if it does not, an ambiguous send timeout is retried by the platform default and
    the applicant receives the same acknowledgment several times. The old key is kept as
    well: whichever one the runtime reads, the intended policy is the one it finds.
    """
    moved = []
    for name, action, _ in _walk(actions):
        policy = (action.get("runtimeConfiguration") or {}).get("retryPolicy")
        if not policy or not isinstance(action.get("inputs"), dict):
            continue
        action["inputs"]["retryPolicy"] = dict(policy)
        moved.append(name)
    return moved


# ── 3. no success copy until the resume actually committed ───────────────────
# Each of the three resume-save Foreach loops carried its OWN failure alert INSIDE the
# loop. In Logic Apps a failure that a later action handles is absorbed, so the Foreach
# still reported Succeeded and the reply branch - which additionally listed Failed and
# Skipped in its runAfter - went ahead: the applicant was told their resume was on file,
# the row was re-queued for Phase 2 against the previous document, and the run went green.
# Lifting the alert OUT of the loop makes the loop's own status truthful, which lets the
# reply branch depend on Succeeded alone. Failure now parks the message unread in Archive,
# the same terminal state every other failure path uses, so the next poll is not blocked
# and the operator can replay it. Depth is unchanged: the alert moves one level SHALLOWER
# and the two cleanup actions are siblings, never children, of the loop.
_SAVE_PATHS = (
    # foreach,                       gated branch,            alert,                                  key
    ("Save_resumes_to_SharePoint", "Add_row", "Notify_create_file_failed", "new"),
    ("Save_dup_update_resume", "IsUnderReplyCap_dup", "Notify_create_dup_update_file_failed", "dup"),
    ("Save_update_resume", "IsUnderReplyCap_update", "Notify_create_update_file_failed", "update"),
)

# Stable cut point for a re-applied build: present in every body this module has
# already rewritten, regardless of which failure folder was configured at the time.
_LOST_MARKER = "<p>The resume was <strong>not saved</strong>"

_LOST_COPY = (
    '<p>The resume was <strong>not saved</strong>, so nothing downstream was allowed to '
    'run: no candidate row was created or re-queued, and <strong>no acknowledgment was '
    'sent</strong>. The message has been moved to %s and left <strong>unread</strong>, '
    'which is this system\'s marker for "failed, needs a replay".</p>'
    '<p><strong>To recover:</strong> find it in %s by the From and Received values above '
    'and move it back to the Inbox, still unread. The next poll re-runs the whole intake '
    'for it, including the applicant reply if applicant mail is on.</p>'
)


def require_saved_resume(actions, failure_folder, mailbox):
    """Gate each reply branch on its resume save, and give failure a clean terminal state."""
    applied = []
    for foreach_name, gated_name, alert_name, key in _SAVE_PATHS:
        loop, owner = _find(actions, foreach_name)
        if loop is None:
            continue
        gated, gated_owner = _find(actions, gated_name)
        if gated is None or gated_owner is not owner:
            raise ValueError(
                "%s and %s are no longer siblings; the success gate would become a "
                "cross-level runAfter, which Power Automate rejects." % (foreach_name, gated_name))

        # (a) lift the alert out of the loop so the loop's status stops lying
        alert = loop.get("actions", {}).pop(alert_name, None)
        if alert is None:
            alert, alert_owner = _find(actions, alert_name)
            if alert is not None and alert_owner is not owner:
                alert_owner.pop(alert_name, None)
        if alert is not None:
            # Skipped too, matching the move in (c). A loop is skipped when the step before it
            # (the Ensure_*_resume_folder create, or the file-list Select) fails or times out;
            # (c) then parks the message unread and terminates green, and without Skipped here
            # that happened with NO alert at all - an applicant silently shelved. On a
            # successful save the loop is Succeeded, so this still never fires then.
            alert["runAfter"] = {foreach_name: ["Failed", "TimedOut", "Skipped"]}
            # With send_admin_failure_alerts=false the alert is already a no-op Compose
            # whose inputs is a plain string, and indexing it crashed the whole build - so
            # that documented switch could not be used at all (found 2026-09-14). The
            # rewiring above still applies; only a real send has a body to rewrite.
            params = (alert.get("inputs") or {}).get("parameters") \
                if isinstance(alert.get("inputs"), dict) else None
            body = params.get("emailMessage/Body") if isinstance(params, dict) else None
            if body is not None:
                # Idempotent, as the module docstring promises. On a pristine body the cut
                # point is the original closing paragraph; on a body this function already
                # rewrote, it is the inserted copy itself. partition() returns the WHOLE
                # string when its marker is absent, so without the second marker a re-run
                # appends _LOST_COPY and another </div> on every pass.
                for _marker in ("<p>The run reports", _LOST_MARKER):
                    if _marker in body:
                        head = body.partition(_marker)[0]
                        break
                else:
                    head = body
                params["emailMessage/Body"] = (
                    head + (_LOST_COPY % (failure_folder, failure_folder)) + "</div>")
            owner.pop(alert_name, None)
            owner[alert_name] = alert

        # (b) the reply/row branch now requires a committed save. Merged, not replaced:
        # Add_row also carries separator dependencies when year/month separators are on,
        # and overwriting its runAfter would silently drop them.
        gated_after = dict(gated.get("runAfter") or {})
        gated_after[foreach_name] = ["Succeeded"]
        gated["runAfter"] = gated_after

        # (c) terminal state for the failure: out of the Inbox, unread, run ends green
        move_name = "Move_%s_save_failed_unread" % key
        term_name = "Terminate_%s_save_failed" % key
        for stale in (move_name, term_name):
            owner.pop(stale, None)
        # Skipped is accepted alongside Failed: the loop is skipped when its own predecessor
        # times out, and a skipped save must not leave the message sitting in the Inbox for
        # the next poll to pick up again. Succeeded is the one status that does NOT route here.
        owner[move_name] = {
            "type": "OpenApiConnection",
            "runAfter": {foreach_name: ["Failed", "TimedOut", "Skipped"]},
            "inputs": {
                # No retry: a replayed MoveV2 races Outlook's mutable message id.
                "retryPolicy": {"type": "none"},
                "parameters": {
                    "messageId": "@outputs('CurrentEmail')?['id']",
                    "mailboxAddress": mailbox,
                    "folderPath": failure_folder,
                },
                "host": {"apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
                         "connectionName": "shared_office365",
                         "operationId": "MoveV2"},
                "authentication": "@parameters('$authentication')",
            },
            "runtimeConfiguration": {"retryPolicy": {"type": "none"}},
        }
        # Terminate waits for the ALERT as well as the move. Both hang off the same
        # loop and therefore run in parallel, and a MoveV2 completes well inside the time
        # an Office365 send takes. Ending the run on the move alone cancels the still-
        # running alert, producing the silent lost resume this whole fix exists to
        # prevent.
        # CRITICAL: move_name must NEVER accept Skipped. On a successful save, move_name
        # is skipped, so Terminate must also be skipped to allow the outer Mark_as_read
        # and Move_to_processed (to Archive) to run.
        term_after = {move_name: ["Succeeded", "Failed", "TimedOut"]}
        if alert is not None:
            term_after[alert_name] = ["Succeeded", "Failed", "TimedOut", "Skipped"]
        owner[term_name] = {
            "type": "Terminate",
            "runAfter": term_after,
            "inputs": {"runStatus": "Succeeded"},
        }
        applied.append(foreach_name)
    return applied


# ── 4. the latest application is the one that gets patched ───────────────────
# Duplicate eligibility was already order-independent: Submitted_ticks projects every
# matching row to a tick value and Is_duplicate compares max(ticks) against the 90-day
# window. The row that then received the patch was first(...) of the same unordered list,
# so a sender holding several applications could have an older one updated. The reference
# path had the same split. Two small Query/Select actions make the selection match the
# decision - latest Received Date wins on both - and every first(...) reference is
# repointed at them. No control action is added, so nesting depth is unchanged.
_TICK = "ticks(coalesce(item()?['Received Date'], '2000-01-01T00:00:00Z'))"


def deterministic_row(actions):
    picked = []

    # (a) duplicate path: pick the latest of the rows Is_duplicate already measured
    ticks, owner = _find(actions, "Submitted_ticks")
    if ticks is not None:
        owner.pop("Latest_submitted", None)
        owner["Latest_submitted"] = {
            "type": "Query",
            "runAfter": {"Submitted_ticks": ["Succeeded"]},
            "inputs": {
                "from": "@body('Filter_submitted')",
                "where": "@equals(%s, max(union(body('Submitted_ticks'), json('[0]'))))" % _TICK,
            },
        }
        dup_if, _ = _find(actions, "Is_duplicate")
        if dup_if is not None:
            # Merged, not replaced, for the same reason require_saved_resume merges:
            # this If can already carry separator dependencies when the year/month
            # separators are on, and overwriting the map would silently drop them. The
            # build still validates in that case; the gate just stops waiting at runtime.
            dup_after = dict(dup_if.get("runAfter") or {})
            dup_after.update({"Filter_submitted": ["Succeeded"],
                              "Submitted_ticks": ["Succeeded"],
                              "Latest_submitted": ["Succeeded"]})
            dup_if["runAfter"] = dup_after
        picked.append("Latest_submitted")

    # (b) reference path: Get_rows_ref is unordered too, so give it the same treatment
    rows_ref, ref_owner = _find(actions, "Get_rows_ref")
    if rows_ref is not None:
        source = "@coalesce(outputs('Get_rows_ref')?['body/value'], json('[]'))"
        for stale in ("Ref_ticks", "Latest_ref"):
            ref_owner.pop(stale, None)
        ref_owner["Ref_ticks"] = {
            "type": "Select",
            "runAfter": {"Get_rows_ref": ["Succeeded"]},
            "inputs": {"from": source, "select": "@" + _TICK},
        }
        ref_owner["Latest_ref"] = {
            "type": "Query",
            "runAfter": {"Ref_ticks": ["Succeeded"]},
            "inputs": {
                "from": source,
                "where": "@equals(%s, max(union(body('Ref_ticks'), json('[0]'))))" % _TICK,
            },
        }
        known, _ = _find(actions, "IsKnownSender")
        if known is not None:
            # Failed/Skipped kept: a failed lookup must still reach the else branch, which
            # is what routes a forged or unmatched reference back into ordinary intake.
            # Merged for the same reason as Is_duplicate above - Latest_ref depends on
            # Get_rows_ref transitively, so keeping any existing predecessor is free.
            known_after = dict(known.get("runAfter") or {})
            known_after["Latest_ref"] = ["Succeeded", "Failed", "Skipped"]
            known["runAfter"] = known_after
        picked.append("Latest_ref")
    return picked


def repoint_row_references(actions):
    """Swap every first(...) row reference onto the deterministic pick.

    Done as a whole-tree string rewrite so no patch item, cap expression, filename or
    reply body can be missed. The two new actions read their source list directly and
    contain no first(...) of their own, so they are not self-rewritten.
    """
    rewritten = _rewrite_expressions(actions,
                                     "first(body('Filter_submitted'))",
                                     "first(body('Latest_submitted'))")
    rewritten = _rewrite_expressions(rewritten,
                                     "first(body('Get_rows_ref')?['value'])",
                                     "first(body('Latest_ref'))")
    return rewritten


def apply_integrity_fixes(definition, failure_folder, mailbox):
    """Entry point. Returns a summary for the build banner."""
    actions = definition["actions"]
    summary = {
        "allow_list": exact_allow_list(actions) is not None,
        "retry_policy": retry_policy_placement(actions),
        "save_gated": require_saved_resume(actions, failure_folder, mailbox),
        "row_pick": deterministic_row(actions),
    }
    definition["actions"] = repoint_row_references(actions)
    return summary
