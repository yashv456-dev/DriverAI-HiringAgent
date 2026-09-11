"""Version every incoming CV; persist its exact attachment references for P2."""


def apply_reliability(definition, config):
    def find_container(tree):
        if 'ResumeFiles' in tree and 'ResumeNames' in tree:
            return tree
        for action in tree.values():
            for child in (action.get('actions'), action.get('else', {}).get('actions')):
                if isinstance(child, dict):
                    found = find_container(child)
                    if found is not None:
                        return found
        return None
    actions = find_container(definition['actions'])
    if actions is None:
        raise ValueError('Resume selection actions missing from P1')
    # Select evaluates each GUID once. Upload retries reuse the selected name; later
    # messages/replays get different names and cannot overwrite an earlier CV.
    # The stored name is READABLE and DETERMINISTIC: '<sanitised original>_<ref tail>.<ext>'.
    #
    # It replaced 'cv_<guid()>' on 2026-09-10. The GUID met this module's versioning goal -
    # a later CV can never overwrite an earlier one - but bought it two real problems:
    #   * guid() is re-evaluated on every RUN, not every message. A re-poll of the SAME
    #     email (the documented outcome when MoveV2 fails and the mail stays in the Inbox)
    #     minted a second random name, so one CV landed twice under two names.
    #   * P2 cannot find it. sharepoint_scoring._resume_name_slots probes two shapes only,
    #     the legacy '<FirstLast>_<FullAppID>' and its own canonical rename, so a GUID name
    #     is invisible to the Excel path and to every maintenance script.
    #
    # The reference tail keeps the versioning guarantee without either problem: AppRef is
    # derived from the Outlook message id, so it is STABLE for a given email (a replay
    # rewrites the same name, which is the idempotent outcome) and DIFFERENT for a genuinely
    # new message (an update or resend cannot overwrite the earlier CV).
    #
    # Sanitising is chained replace() because Power Automate has no regex. Removed: the nine
    # characters SharePoint rejects outright (" * : < > ? / \ |), plus # and % which some
    # tenants still reject, plus the comma, which is the separator P1 itself uses to join
    # 'Original Filename'. Spaces become underscores so the URL carries no %20. The stem is
    # capped so the full path stays inside SharePoint's 400-character limit.
    raw_stem = ("substring(item()?['name'],0,"
                "sub(length(item()?['name']),add(length(last(split(item()?['name'],'.'))),1)))")
    clean = raw_stem
    for bad in ('"', '*', ':', '<', '>', '?', '/', '\\', '|', '#', '%', ',', "'"):
        clean = "replace(%s,'%s','')" % (clean, "''" if bad == "'" else bad)
    clean = "replace(trim(%s),' ','_')" % clean
    # Never emit an empty stem (a name that was nothing but illegal characters).
    clean = "if(equals(length(%s),0),'resume',%s)" % (clean, clean)
    clean = "if(greater(length(%s),60),substring(%s,0,60),%s)" % (clean, clean, clean)
    # ONE RESUME PER PERSON. The tail is the tail of the row this CV belongs to, not of
    # the message that delivered it. On a first application those are the same value. On a
    # resend or an update the duplicate/reference lookup has already resolved the EXISTING
    # row, so the tail is that row's - the new CV therefore writes to the SAME path and
    # replaces the previous copy instead of adding a second file beside it. SharePoint keeps
    # the prior content in the file's own version history, so nothing is actually lost.
    #
    # This is why the expression is built per branch rather than once up front: at the point
    # the old shared Select ran, the flow had not yet looked up which application the CV
    # belonged to and could only name it after the message.
    def tail_of(app_expr):
        return "substring(%s,sub(length(%s),4),4)" % (app_expr, app_expr)
    # The reference tail makes the name unique ACROSS messages. It cannot separate two
    # attachments WITHIN one message that share a filename, which Outlook allows - so a
    # second discriminator is appended, but only when this email actually kept more than one
    # attachment. 274 of 279 live applications carry exactly one and keep the clean
    # '<stem>_<tail>.<ext>' form.
    #
    # The discriminator is the ATTACHMENT ID, not the content length. Length was tried first
    # and is wrong: base64 length is a pure function of byte length, so two different
    # documents of identical size - two one-page PDFs from the same exporter, two scans at a
    # fixed resolution - produce the same name, the second CreateFile silently replaces the
    # first, and the manifest then lists one file for two slots. The attachment id is unique
    # per attachment and stable for a given message, so it keeps the replay-idempotence the
    # ref tail provides. Seeded and stripped exactly like AppRefSeed, so a short or missing
    # id cannot make substring() overrun.
    att_seed = ("concat('000000',replace(replace(replace(coalesce(item()?['id'],''),"
                "'-',''),'_',''),'=',''))")
    multi = ("if(greater(length(body('ResumeFiles')),1),"
             "concat('_',toUpper(substring(%s,sub(length(%s),4),4))),'')" % (att_seed, att_seed))
    def stored_for(app_expr):
        return ("concat(%s,'_',%s,%s,'.',toLower(last(split(item()?['name'],'.'))))"
                % (clean, tail_of(app_expr), multi))

    # Retained for ResumeNames' dependency chain and as the fallback shape.
    actions['StoredResumeFiles'] = {
        'type': 'Select', 'runAfter': {'ResumeFiles': ['Succeeded']},
        'inputs': {'from': "@body('ResumeFiles')", 'select': {
            'name': "@item()?['name']", 'contentBytes': "@item()?['contentBytes']",
            'stored_name': '@' + stored_for("outputs('AppRef')")}}}
    offset = int(config.get('timezone', {}).get('utc_offset_hours', 0))
    stamp = "formatDateTime(addHours(outputs('CurrentEmail')?['receivedDateTime'],%d),'yyyy/MMMM')" % offset
    folder = "concat(outputs('CONFIG')['resumes_folder'],'/',%s)" % stamp
    # The SharePoint CONNECTOR wants the library-qualified path ('/Shared Documents/...')
    # for CreateFile, but for CreateNewFolder (Ensure_*_resume_folder) the action target
    # 'table' is ALREADY the library, so its path must be relative to the library root
    # ('Candidate_Resumes/...'). The manifest is read by P2 through the GRAPH drive API,
    # whose paths are also relative to the drive root.
    library = config['sharepoint']['resumes_folder'].strip('/')
    graph_root = '/' + (library.split('/', 1)[1] if '/' in library else library)
    graph_folder = "concat('%s','/',%s)" % (graph_root, stamp)
    folder_relative = "concat('%s','/',%s)" % (graph_root.lstrip('/'), stamp)
    actions['ResumeManifest'] = {
        'type': 'Select', 'runAfter': {'StoredResumeFiles': ['Succeeded']},
        'inputs': {'from': "@body('StoredResumeFiles')", 'select': {
            'name': "@item()?['stored_name']", 'original_name': "@item()?['name']",
            'folder': '@' + graph_folder}}}
    actions['ResumeNames']['runAfter'] = {'ResumeManifest': ['Succeeded']}
    manifest = "@concat('manifest:',string(body('ResumeManifest')))"

    # Which row each branch is actually writing against. coalesce() keeps a first
    # application working if the lookup returned nothing.
    BRANCH_APP_ID = {
        'Save_resumes_to_SharePoint': ('new', "outputs('AppRef')"),
        'Save_dup_update_resume': (
            'dup', "coalesce(first(body('Latest_submitted'))?['Application ID'],outputs('AppRef'))"),
        'Save_update_resume': (
            'update', "coalesce(first(body('Latest_ref'))?['Application ID'],outputs('AppRef'))"),
    }

    def visit(tree):
        for name, action in list(tree.items()):
            if name in BRANCH_APP_ID:
                key, app_expr = BRANCH_APP_ID[name]
                sel_name, man_name = 'StoredFiles_' + key, 'Manifest_' + key
                # Both Selects sit in the loop's OWN scope, so they can read the row the
                # enclosing branch resolved. runAfter {} is safe: the lookup they reference
                # ran in an ancestor scope, which completed before this scope started.
                tree[sel_name] = {
                    'type': 'Select', 'runAfter': {},
                    'inputs': {'from': "@body('ResumeFiles')", 'select': {
                        'name': "@item()?['name']", 'contentBytes': "@item()?['contentBytes']",
                        'stored_name': '@' + stored_for(app_expr)}}}
                tree[man_name] = {
                    'type': 'Select', 'runAfter': {sel_name: ['Succeeded']},
                    'inputs': {'from': '@body(%r)' % sel_name, 'select': {
                        'name': "@item()?['stored_name']", 'original_name': "@item()?['name']",
                        'folder': '@' + graph_folder}}}
                after = dict(action.get('runAfter') or {})
                after[man_name] = ['Succeeded']
                action['runAfter'] = after
                action['foreach'] = '@body(%r)' % sel_name
                branch_manifest = "@concat('manifest:',string(body(%r)))" % man_name
                for child_name, child in list(action['actions'].items()):
                    if child_name not in ('Create_file', 'Create_update_file', 'Create_dup_update_file'):
                        continue
                    if child.get('inputs', {}).get('host', {}).get('operationId') == 'CreateFile':
                        params = child['inputs']['parameters']
                        params['name'] = "@items('%s')?['stored_name']" % name
                        params['folderPath'] = '@' + folder
                        import copy
                        source = ("first(body('Filter_submitted'))" if name == 'Save_dup_update_resume' else
                                  "first(body('Get_rows_ref')?['value'])" if name == 'Save_update_resume' else None)
                        received = "formatDateTime(addHours(outputs('CurrentEmail')?['receivedDateTime'],%d),'yyyy-MM-ddTHH:mm:ss')" % offset
                        metadata = {
                            'Application ID': '@'+source+"?['Application ID']" if source else "@outputs('AppRef')",
                            'Received Date': '@'+source+"?['Received Date']" if source else '@'+received,
                            'Last Updated Date': '@'+received,
                            'Application Updates': "@add(int(coalesce(%s?['Application Updates'],'0')),1)" % source if source else 0,
                            'Email': "@toLower(trim(if(contains(outputs('CurrentEmail')?['from'],'<'),replace(last(split(outputs('CurrentEmail')?['from'],'<')),'>',''),outputs('CurrentEmail')?['from'])))",
                            'Full Name': '@'+source+"?['Full Name']" if source else '',
                            'Original Filename': "@join(body('ResumeNames'), ', ')",
                            'Resume URL': branch_manifest,
                            'Resume Folder Path': '@'+graph_folder,
                            'Mail Subject': "@coalesce(outputs('CurrentEmail')?['subject'],'')",
                            # bodyPreview, not body: 'body' is the full HTML part, so the cell
                            # fills with '<html><head><meta ...>' markup instead of readable
                            # text, and P2 reads this column as a fallback extraction source.
                            'Mail Body': "@coalesce(outputs('CurrentEmail')?['bodyPreview'],'')",
                            'Status': 'New Email Received', 'Has Resume': 'Yes',
                        }
                        compose_name = 'Capture_'+child_name
                        event_name = 'Preserve_'+child_name
                        action['actions'][compose_name] = {'type':'Compose', 'inputs':metadata,
                                                         'runAfter':{child_name:['Succeeded']}}
                        event = copy.deepcopy(child)
                        event['runAfter'] = {compose_name:['Succeeded']}
                        event['inputs']['parameters']['name'] = "@concat('intake_',items('%s')?['stored_name'],'.json')" % name
                        event['inputs']['parameters']['body'] = "@string(outputs('%s'))" % compose_name
                        action['actions'][event_name] = event
            if name in ('Ensure_resume_folder', 'Ensure_update_resume_folder', 'Ensure_dup_resume_folder'):
                action['inputs']['parameters']['parameters/path'] = '@' + folder_relative
            # Add_row and the Patch_* actions are deliberately NOT touched. Writing
            # 'Resume URL' / 'Resume Folder Path' into the row was tried on 2026-09-08 and
            # broke TWO live runs at Add_row, with the CV and the sidecar already committed,
            # so the row write was the only step that failed. The sidecar carries these
            # references instead. Do not re-add without reproducing that failure first.
            for subtree in (action.get('actions'), action.get('else', {}).get('actions')):
                if isinstance(subtree, dict):
                    visit(subtree)
    visit(actions)
