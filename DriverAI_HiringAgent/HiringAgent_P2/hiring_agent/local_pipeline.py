"""One local owner; remote intake/download and publication surround isolated processing."""
from __future__ import annotations

import contextlib
import hashlib
import json
import multiprocessing
import os
import time
from pathlib import Path

from . import config as cfg
from .sqlite_store import SQLiteCandidateStore, BASE, encoded


@contextlib.contextmanager
def worker_lock(path):
    """OS-owned lock: released even when the worker crashes. Shared by every entry point."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('a+b') as handle:
        handle.seek(0)
        if handle.read(1) == b'':
            handle.write(b'0')
            handle.flush()
        handle.seek(0)
        if os.name == 'nt':
            import msvcrt
            msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            import fcntl
            fcntl.flock(handle, fcntl.LOCK_EX | fcntl.LOCK_NB)
        try:
            yield
        finally:
            handle.seek(0)
            if os.name == 'nt':
                msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(handle, fcntl.LOCK_UN)


class LocalClient:
    """Workbook-shaped adapter over one isolated local application. No Graph methods."""
    local_execution = True
    hostname = 'local SQLite'
    resumes_folder = 'local'
    table = 'HiringAgent_P1_Candidates'

    def __init__(self, values, documents):
        self.candidate_store = SQLiteCandidateStore(':memory:')
        self.candidate_store.add(values)
        self.documents = documents

    def list_rows(self):
        return [{'index': int(r.version), 'values': r.values} for r in self.candidate_store.all_rows()]

    def list_rejected_rows(self):
        return [{'index': int(r.version), 'values': r.values} for r in self.candidate_store.all_rows('rejected')]

    def list_unscored_rows(self, statuses):
        return [r for r in self.list_rows() if r['values'].get('Status') in statuses]

    def download_resume(self, name, subfolder=''):
        from sharepoint_client import SharePointError
        if name not in self.documents:
            if len(self.documents) == 1:
                doc = next(iter(self.documents.values()))
                raw = Path(doc['path']).read_bytes()
                if hashlib.sha256(raw).hexdigest() != doc['sha256']:
                    raise ValueError('Cached attachment hash changed')
                return raw
            raise SharePointError('Not in verified local attachment manifest', status_code=404)
        raw = Path(self.documents[name]['path']).read_bytes()
        if hashlib.sha256(raw).hexdigest() != self.documents[name]['sha256']:
            raise ValueError('Cached attachment hash changed')
        return raw

    def resume_web_url(self, name, subfolder=''):
        return self.documents.get(name, {}).get('url', '')

    def send_mail(self, *args, **kwargs):
        # This adapter must never send mail from a child or mark a message as sent.
        return False


def _child(connection, values, documents, roles):
    try:
        from . import sharepoint_scoring as scoring
        from . import extraction
        client = LocalClient(values, documents)
        # Use the tested scoring/geo logic. Only state and transport adapters differ.
        extraction.STRICT_AI_STAGES = bool(cfg.REQUIRE_AI)
        scoring.get_active_roles = lambda: roles
        result = scoring.score_from_sharepoint(_local_client=client)
        rows = [(sheet, r.values) for sheet in ('main', 'rejected')
                for r in client.candidate_store.all_rows(sheet)]
        connection.send({'summary': result, 'rows': rows})
    except Exception as error:
        connection.send({'error': f'{type(error).__name__}: {error}'})
    finally:
        connection.close()


def isolated_score(values, documents, roles, timeout=None):
    """A killable process bounds OCR plus ALL LLM stages, not just one HTTP request."""
    budget = timeout or float(os.getenv('HIRING_CANDIDATE_TIMEOUT', '900'))
    if budget <= 0:
        raise ValueError('Candidate timeout must be positive')
    ctx = multiprocessing.get_context('spawn')
    reader, writer = ctx.Pipe(duplex=False)
    process = ctx.Process(target=_child, args=(writer, values, documents, roles))
    process.start()
    writer.close()
    try:
        if not reader.poll(budget):
            raise TimeoutError(f'Candidate exceeded {budget:g}s local processing deadline')
        result = reader.recv()
        if 'error' in result:
            raise RuntimeError(result['error'])
        return result
    finally:
        reader.close()
        process.join(timeout=1)
        if process.is_alive():
            process.terminate()
            process.join(timeout=5)
        if process.is_alive():
            process.kill()
            process.join()


def _filename_identity(value, app='', tail=''):
    """Comparable filename/name hint with application tokens and punctuation removed."""
    value = Path(str(value or '')).stem.lower()
    for token in (str(app or '').lower(), str(tail or '').lower()):
        if token:
            value = value.replace(token, '')
    return ''.join(ch for ch in value if ch.isalnum())


def _find_by_app_ref(remote, app, places, hints=None):
    """Locate a CV by the application reference stamped into its filename.

    Every name P1 and P2 save ends '_<AppRef tail>.<ext>', but reconstructing the rest of the
    name needs the candidate's real name - and an unscored row still carries P1's guess from
    the From header. APP-20260713-0025-WPWA asked for 'awaisahmadkallu01_WPWA.pdf' while the
    stored file was 'Awais_Ahmad_WPWA.pdf', so the row could never be scored and never got the
    name that would have found it. The reference does not depend on any of that.

    Returns (folder, name, bytes), or None. A full Application ID match is authoritative.
    A short-tail match is accepted only when an applicant/original-filename hint uniquely
    identifies it. Four-character tails are not unique: APP-20260805-0931-PMCA (Vibha)
    and APP-20260810-1014-PMCA (Shalin) collided live, and the old canonical-shape fallback
    attached Shalin's CV to Vibha's row. Ambiguity now fails closed instead of guessing.
    """
    from sharepoint_client import SharePointError
    tail = str(app or '').rsplit('-', 1)[-1].strip()
    if not tail:
        return None
    for place in places:
        try:
            children = remote.list_folder_children(place)
        except Exception:
            continue
        names = [k.get('name', '') for k in children]
        exact = [n for n in names if app.lower() in n.lower()]
        hits = exact or [n for n in names
                         if Path(n).stem.lower().endswith('_' + tail.lower())]
        if not hits:
            continue
        target = exact[0] if len(exact) == 1 else None
        if target is None:
            # Hints are ordered by authority. Original Filename and sender identity come
            # before a previously-derived Full Name/Resume Link, which might itself be the
            # result of an earlier bad tail match. Use the first hint with one unique hit.
            for value in (hints or []):
                identity = _filename_identity(value, app, tail)
                if not identity:
                    continue
                hinted = []
                for candidate in hits:
                    candidate_id = _filename_identity(candidate, app, tail)
                    if (identity == candidate_id or
                            (len(identity) >= 6 and identity in candidate_id) or
                            (len(candidate_id) >= 6 and candidate_id in identity)):
                        hinted.append(candidate)
                if len(hinted) == 1:
                    target = hinted[0]
                    break
        if target is None:
            cfg.logger.warning(
                '        Attachment: %s has ambiguous short-tail matches in %s: %s',
                app, place, ', '.join(hits))
            continue
        try:
            return place, target, remote.download_file(place, target)
        except SharePointError:
            continue
    return None


def cache_documents(remote, values, cache_dir, *, rejected=False):
    """Resolve every attachment before OCR. Originals are stored by content hash."""
    from . import sharepoint_scoring as scoring
    from sharepoint_client import SharePointError
    app = values['Application ID']
    url = str(values.get('Resume URL', '') or '')
    received = scoring._parse_received(values.get('Received Date'))
    folder = cfg.dated_subpath(received)
    if rejected:
        folder = scoring._rejected_subpath(folder)
    if url.startswith('manifest:'):
        manifest = json.loads(url[len('manifest:'):])
        if not isinstance(manifest, list) or not manifest:
            raise ValueError('Empty or invalid P1 attachment manifest')
        slots = [[entry] for entry in manifest]
    else:
        names = scoring._resume_name_slots(app, values.get('Original Filename', ''),
                                          values.get('Full Name'), values.get('Category'))
        linked = scoring._resume_name_from_url(url)
        if linked:
            if names:
                names[0] = list(dict.fromkeys([linked, *names[0]]))
            else:
                names = [[linked]]
        slots = [[{'name': name, 'folder': f'{remote.resumes_folder}/{folder}', 'legacy': True}
                  for name in slot] for slot in names]
    if not slots:
        raise ValueError(f'{app}: no attachment reference')
    # A rejection moves the CV into the year-level Rejected bucket, but the row itself can
    # stay on the main sheet - a half-applied rejection, or a row restored without its file.
    # The dated folder then holds nothing and the candidate can never be scored again: live
    # 2026-09-08, APP-20260713-0025-WPWA deferred on 'missing attachment slot' while
    # Awais_Ahmad_WPWA.pdf sat in 2026/Rejected. Look in the counterpart bucket before
    # concluding the attachment is gone, in both directions.
    counterpart = (cfg.dated_subpath(received) if rejected
                   else scoring._rejected_subpath(cfg.dated_subpath(received)))
    documents = {}
    cache_dir = Path(cache_dir)
    cache_dir.mkdir(parents=True, exist_ok=True)
    original_names = [part.strip() for part in
                      str(values.get('Original Filename', '') or '').split(',')]
    tail = str(app or '').rsplit('-', 1)[-1].strip()
    email_local = str(values.get('Email', '') or '').split('@', 1)[0]
    for slot_index, slot in enumerate(slots):
        original_hint = (original_names[slot_index]
                         if slot_index < len(original_names) else '')

        def _candidate_rank(entry):
            name = str(entry.get('name', '') or '')
            if str(app).lower() in name.lower():
                return 0
            candidate_id = _filename_identity(name, app, tail)
            original_id = _filename_identity(original_hint, app, tail)
            email_id = _filename_identity(email_local, app, tail)
            if original_id and (original_id == candidate_id or original_id in candidate_id):
                return 1
            if email_id and len(email_id) >= 6 and email_id in candidate_id:
                return 2
            return 3

        slot = sorted(slot, key=_candidate_rank)
        for entry in slot:
            name = entry['name']
            places = [entry['folder'], f'{remote.resumes_folder}/{counterpart}']
            # A manifest written before 2026-09-08 carries the SharePoint CONNECTOR's
            # library-qualified path ('/Shared Documents/Candidate_Resumes/...'). Graph paths
            # are relative to the drive root, which IS that library, so the connector spelling
            # 404s. Retry without the leading library segment rather than deferring the row.
            head = str(entry.get('folder', '') or '').strip('/').split('/', 1)
            if len(head) == 2 and ' ' in head[0]:
                places.insert(1, '/' + head[1])
            if entry.get('legacy'):
                places.append(remote.resumes_folder)
            raw = target_folder = None
            for place in places:
                try:
                    raw = remote.download_file(place, name)
                    target_folder = place
                    break
                except SharePointError as error:
                    if error.status_code != 404:
                        raise
            if raw is None:
                found = _find_by_app_ref(
                    remote, app, places,
                    hints=(original_hint, email_local, values.get('Full Name', ''), name))
                if found:
                    target_folder, name, raw = found
            if raw is None:
                continue
            digest = hashlib.sha256(raw).hexdigest()
            suffix = Path(name).suffix.lower()
            path = cache_dir / (digest + suffix)
            if not path.exists():
                temp = path.with_suffix(path.suffix + '.tmp')
                temp.write_bytes(raw)
                temp.replace(path)
            # A link lookup failure is infrastructure failure, not unreadable content.
            web_url = remote.file_web_url(target_folder, name)
            documents[name] = {'path': str(path), 'sha256': digest, 'url': web_url, 'folder': target_folder}
            break
        else:
            raise FileNotFoundError(f'{app}: missing attachment slot')
    # Explicit filenames eliminate guesses for versioned P1 uploads in the local worker.
    return documents


#: Fields an idle pass tries to recover. These are the ones a resume usually DOES state and
#: the client reads: a blank here is normally a parse miss, not an absent fact. Roles,
#: Category and Status are deliberately excluded - those are decisions, not transcription.
RECOVERABLE_FIELDS = ('Phone', 'Location', 'Country', 'Years Exp',
                      'Education', 'Education Start Date', 'Education End Date')


#: Everything P2 is allowed to write onto P1's intake workbook.
#:
#: The workbook is a STATE fallback, not a data backup. Every field P2 derives lives in the
#: local database and in the two published reports; sync_from_client compares only
#: INPUT_COLUMNS and leaves a matching row untouched, so nothing P2 writes here is ever read
#: back into P2. Two columns are the whole useful contract: workflow Status and the stored
#: Resume Link after P2 renames the file.
MASTER_WRITE_COLUMNS = ('Status', 'Resume Link')


def _master_patch(values):
    """The subset of `values` that may be PATCHed onto a main-sheet row.

    Unlisted columns keep whatever they already hold, so this narrows the write rather
    than blanking anything.
    """
    return {col: values[col] for col in MASTER_WRITE_COLUMNS if col in values}


def _is_gap(value):
    return str(value or '').strip().lower() in ('', 'missing', 'n/a', 'not extracted')


def _recovery_queue(store, limit):
    """Scored rows with a recoverable blank whose CV has not already been re-read.

    Restores the 2026-09-01 rule (see sharepoint_scoring._reevaluate_missing_scored_fields)
    that an idle pass still re-checks scored rows for a missing Education date / Phone /
    Location. The local pipeline lost it: the isolated worker only ever sees the ONE candidate
    it was handed, so nothing swept the catalogue any more.
    """
    from .store import CandidateRow
    tried = {r['app_id']: r['tried_hash'] for r in store.conn.execute('SELECT * FROM recovery')}
    out = []
    for r in store.conn.execute(
            "SELECT * FROM candidates WHERE active=1 AND app_id<>'' ORDER BY id"):
        values = json.loads(r['values_json'])
        if str(values.get('Status', '')).strip() != 'Scored':
            continue
        if not any(_is_gap(values.get(f)) for f in RECOVERABLE_FIELDS):
            continue
        if tried.get(r['app_id']) == (r['source_hash'] or ''):
            continue
        out.append(CandidateRow(app_id=r['app_id'], values=values,
                                version=str(r['id']), sheet=r['sheet']))
        if len(out) >= limit:
            break
    return out


def _file_by_sheet(remote, documents, sheet, received, dry_run):
    """Keep each stored CV in the folder its sheet belongs to.

    A rejection files the CV in the year's Rejected folder, and a candidate who is back on the
    main sheet has it returned to the month folder. The SharePoint pass always did this through
    _move_resumes, but that helper returns early for a local_execution client, the isolated
    worker never touches remote files, and this parent then recorded wherever the file was
    found. Live APP-20260805-0035-MDEA, rejected on 2026-09-15, kept Resume Folder Path
    '/Candidate_Resumes/2026/August' while every other Rejected row read '2026/Rejected'
    (reported 2026-09-17).

    move_resume copies and verifies; the original stays where it was, as every resume has
    since 2026-09-06. Failure is non-fatal and leaves the row pointing at the file it found.
    """
    if dry_run or getattr(remote, 'local_execution', False):
        return documents
    move = getattr(remote, 'move_resume', None)
    if not callable(move):
        return documents
    from . import sharepoint_scoring as scoring
    root = str(getattr(remote, 'resumes_folder', '') or '').rstrip('/')
    dated = f"{root}/{cfg.dated_subpath(scoring._parse_received(received))}"
    rejected = f"{root}/{scoring._rejected_subpath(cfg.dated_subpath(scoring._parse_received(received)))}"
    filed = {}
    for name, doc in documents.items():
        folder = str(doc.get('folder', '') or '')
        in_rejected = folder.strip('/') == rejected.strip('/')
        # Only the two mismatches are corrected; a file P1 stored somewhere else is left alone.
        target = (rejected if sheet == 'rejected' and not in_rejected
                  else dated if sheet != 'rejected' and in_rejected else None)
        if target is None:
            filed[name] = doc
            continue
        try:
            moved = move(name, folder, target)
        except Exception as error:
            cfg.logger.warning('        File move : %s kept in %s (%s)', name, folder, error)
            moved = False
        if moved:
            doc = dict(doc, folder=target, url=remote.file_web_url(target, name))
            cfg.logger.info('        Filed     : %s -> %s', name, target)
        filed[name] = doc
    return filed


def _rename_to_canonical(remote, documents, app_id, full_name, category, dry_run):
    """Give each stored CV the '<First>_<Last>_<AppTail>' name, once the real name is known.

    P1 names the file from the From header, which is only an email local part when the sender
    set no display name - 'awaistahir0001_APP-20260908-1304-4BFA.pdf' against the
    'Yash_Verma_MBMA.pdf' every other row shows. P2 always fixed that after scoring, but
    _rename_scored_resumes returns early for a local_execution client, so the isolated worker
    (which is the only thing that learns the real name) can never do it. It has to run out
    here in the parent, which holds the real Graph client.

    Graph renames in place with a PATCH on the item's name - no copy, no delete, and the file
    keeps its identity and version history. Failure is non-fatal: the score is already good,
    and a row keeping P1's filename is a cosmetic problem, not a lost candidate.
    """
    if dry_run or getattr(remote, 'local_execution', False):
        return documents
    from . import sharepoint_scoring as scoring
    root = str(getattr(remote, 'resumes_folder', '') or '').strip('/')
    renamed = {}
    for name, doc in documents.items():
        target = scoring._canonical_resume_name(full_name, app_id, Path(name).suffix, category)
        if target == name:
            renamed[name] = doc
            continue
        folder = str(doc.get('folder', '') or '').strip('/')
        subfolder = folder[len(root):].strip('/') if root and folder.startswith(root) else folder
        renamed_ok = False
        rename_fn = getattr(remote, 'rename_resume', None)
        if callable(rename_fn):
            try:
                renamed_ok = rename_fn(name, target, subfolder)
            except Exception as error:
                cfg.logger.warning('        Rename error: %s -> %s (%s)', name, target, error)
        if renamed_ok:
            doc = dict(doc, url=remote.file_web_url(doc['folder'], target))
            renamed[target] = doc
            cfg.logger.info('        Renamed   : %s -> %s', name, target)
            continue

        # If rename didn't succeed (e.g. target already exists in SharePoint):
        file_web_url_fn = getattr(remote, 'file_web_url', None)
        if callable(file_web_url_fn):
            try:
                web_url = file_web_url_fn(doc['folder'], target)
                if web_url:
                    doc = dict(doc, url=web_url)
                    renamed[target] = doc
                    cfg.logger.info('        Rename    : %s already exists as %s; using it', name, target)
                    delete_fn = getattr(remote, 'delete_file', None)
                    if callable(delete_fn):
                        try:
                            delete_fn(doc['folder'], name)
                            cfg.logger.info('        Cleaned up intake duplicate: %s', name)
                        except Exception:
                            pass
                    continue
            except Exception:
                pass

        cfg.logger.warning('        Rename    : %s kept its P1 name', name)
        renamed[name] = doc
    return renamed


def run_local_pipeline(remote, *, dry_run=False, process=True, app_ids=None, force=False,
                       scorer=isolated_score):
    from . import sharepoint_scoring as scoring
    from .jd_sources import get_active_roles
    from .publisher import publish_results
    from .sqlite_store import database_path
    path = database_path()
    summary = {'processed': 0, 'rejected': 0, 'errors': 0, 'deferred': 0,
               'location_review': 0, 'client_export': ''}
    # A dry run must not write into the real publish directory. Both share the revision
    # counter, so a preview that reached revision N wrote generation_N there, and the next
    # REAL run at revision N found those files already present, skipped the rebuild and
    # published the preview's stale contents - live 2026-09-08: the client sheet kept
    # 'awaistahir0001_APP-...pdf' after the rename because a preview had built that row first.
    preview_dir = (path.parent / 'P2_Preview') if dry_run else None
    with worker_lock(path.with_suffix('.worker.lock')):
        store = SQLiteCandidateStore(path)
        try:
            if dry_run:
                # No persistent candidate/import/publication writes during previews.
                # Say so plainly: the per-candidate lines below come from the worker, which
                # runs against this throwaway store and reports its writes exactly as a live
                # run would, so without this banner a preview is hard to tell from the real
                # thing in the log.
                cfg.logger.info('  PREVIEW   Dry run: scoring against a throwaway in-memory '
                                'copy. No candidate row, resume file, workbook or email is '
                                'changed; per-candidate "row updated" lines below are '
                                'previews only.')
                memory = SQLiteCandidateStore(':memory:')
                store.conn.backup(memory.conn)
                store.conn.close()
                store = memory
            with store.conn:
                store.conn.execute("UPDATE attempts SET state='interrupted',finished_at=CURRENT_TIMESTAMP WHERE state='running'")
            summary['imported'] = store.sync_from_client(remote)
            # The retired amber doubt fill outlives the code that wrote it, so it has to be
            # cleared off the live sheet once. _ensure_schema does this for the excel
            # backend, but every sqlite entry point short-circuits to this function long
            # before that runs, so the sqlite path has to ask for it here or the colour
            # never goes away. Skipped on a preview: a dry run writes nothing remote.
            if not dry_run:
                scoring._clear_row_highlighting(remote)
            if not process:
                summary['client_export'] = str(publish_results(store, remote, upload=not dry_run,
                                                               output_dir=preview_dir))
                return summary
            roles = get_active_roles()
            # Freeze payload and its hash once; every child sees the same role set.
            role_json = encoded(roles)
            role_hash = hashlib.sha256(role_json.encode()).hexdigest()
            artifact_dir = path.parent / 'processing_artifacts'
            if not dry_run:
                artifact_dir.mkdir(parents=True, exist_ok=True)
                (artifact_dir / f'jd_{role_hash}.json').write_text(role_json, encoding='utf-8')
            queue = [r for r in store.all_rows() if
                     r.values.get('Status') in ('New Email Received', cfg.STATUS_NEEDS_REVIEW)
                     or (r.values.get('Status') == cfg.STATUS_LOCATION_REVIEW
                         and scoring._updated_after_location_request(r.values))]
            if app_ids is not None:
                wanted = set(app_ids)
                queue = [r for sheet in ('main', 'rejected') for r in store.all_rows(sheet)
                         if r.app_id in wanted]
            elif force:
                queue = [r for sheet in ('main', 'rejected') for r in store.all_rows(sheet)]
            conflicts = {r[0] for r in store.conn.execute('SELECT app_id FROM import_conflicts')}
            queue = [r for r in queue if r.app_id not in conflicts]
            # Retry cap. SCORE_RETRY_MAX exists so a row is "never retried forever", but the
            # give-up path that enforces it lives in the excel scorer, which this function
            # short-circuits past - so on sqlite a row that fails deterministically (same
            # resume, same cached JDs, temperature-0 model) was re-attempted on every run
            # for good: APP-20260821-0812-MCQA reached 22 deferred attempts against a cap of
            # 3. Stop re-running a row once its budget is spent and say so, loudly, once per
            # run. An explicit --app-id/force re-run is an operator override and still runs.
            if app_ids is None and not force:
                exhausted = [r for r in queue
                             if store.deferred_attempts(int(r.version)) >= cfg.SCORE_RETRY_MAX]
                if exhausted:
                    cfg.logger.warning(
                        '  STUCK     %d row(s) have spent all %d retries on the current input '
                        'and will NOT be retried until it changes or they are re-run '
                        'explicitly: %s', len(exhausted), cfg.SCORE_RETRY_MAX,
                        ', '.join(r.app_id for r in exhausted))
                    spent = {r.app_id for r in exhausted}
                    queue = [r for r in queue if r.app_id not in spent]
            queue = queue[:cfg.SCORING_BATCH_LIMIT]
            # Idle pass only: with nothing waiting, spend the spare time re-reading the CVs of
            # already-scored rows that still have a recoverable blank. Bounded by
            # HIRING_RECOVERY_LIMIT so an idle poll stays short, and each attempt is recorded
            # against the input hash so an unfillable gap is tried once, not on every poll.
            recovery = []
            if not queue and app_ids is None and not force:
                recovery = [r for r in _recovery_queue(store, int(os.getenv('HIRING_RECOVERY_LIMIT', '5')))
                            if r.app_id not in conflicts]
                queue = recovery
                if recovery:
                    cfg.logger.info('  Idle pass : re-reading %d scored row(s) with missing fields.',
                                    len(recovery))
            for row in queue:
                attempt, version = store.begin_attempt(int(row.version))
                try:
                    values = dict(row.values)
                    documents = cache_documents(remote, values, path.parent / 'resume_cache', rejected=row.sheet == 'rejected')
                    # Manifest is transport metadata; human-facing result URL is first CV.
                    values['_Local Resume Names'] = list(documents)
                    values['Resume URL'] = next(iter(documents.values()))['url']
                    values['Status'] = 'New Email Received'
                    output = scorer(values, documents, roles)
                    outcome = output['summary']
                    if outcome.get('errors') or outcome.get('deferred'):
                        raise RuntimeError('Local processing deferred or failed; prior result retained')
                    rows = output['rows']
                    if len(rows) != 1 or rows[0][1].get('Application ID') != row.app_id:
                        raise ValueError('Worker did not return exactly the requested application')
                    sheet, result_values = rows[0]
                    result_values.pop('_Local Resume Names', None)
                    documents = _rename_to_canonical(
                        remote, documents, row.app_id, result_values.get('Full Name'),
                        result_values.get('Category'), dry_run)
                    documents = _file_by_sheet(remote, documents, sheet,
                                               result_values.get('Received Date'), dry_run)
                    # 'Original Filename' means what the APPLICANT sent. It used to be
                    # overwritten here with the post-rename canonical name, so the column
                    # reported 'Yash_Verma_AF6A.pdf' rather than the
                    # 'Yash_Verma_Senior_UiPath_RPA_Developer_Final_2026.pdf' that actually
                    # arrived - destroying the only record of the sender's own filename and
                    # making the column a duplicate of the file's current name. Resume URL
                    # below already carries where the file now lives, so nothing needs this
                    # cell to track the rename. Only fill it when P1 left it blank.
                    if not str(result_values.get('Original Filename', '') or '').strip():
                        result_values['Original Filename'] = ', '.join(documents)
                    if not str(result_values.get('Last Updated Date', '') or '').strip():
                        result_values['Last Updated Date'] = result_values.get('Received Date') or ''
                    first_document = next(iter(documents.values()))
                    result_values['Resume URL'] = first_document['url']
                    result_values['Resume Folder Path'] = first_document['folder']
                    result_values['Resume Link'] = next(iter(documents.keys()))
                    if result_values.get('Status') == 'New Email Received':
                        raise RuntimeError('Worker did not produce an outcome')
                    # Save traceable input/model context before accepting the result.
                    if not dry_run:
                        (artifact_dir / f'attempt_{attempt}.json').write_text(
                            encoded({'input': row.values, 'documents': documents, 'jd_hash': role_hash,
                                     'model': cfg.OLLAMA_MODEL, 'result': [sheet, result_values]}), encoding='utf-8')
                    state = store.finish_attempt(attempt, version, (sheet, result_values))
                    if state == 'completed':
                        summary['processed'] += 1
                        summary['rejected'] += int(sheet == 'rejected')
                        summary['location_review'] += int(result_values.get('Status') == cfg.STATUS_LOCATION_REVIEW)
                        if not dry_run:
                            try:
                                _sync_row_to_sharepoint(remote, row.app_id, sheet,
                                                        result_values, row.values)
                            except Exception as sync_err:
                                cfg.logger.warning('Could not sync row %s back to SharePoint table: %s', row.app_id, sync_err)
                except Exception as error:
                    store.finish_attempt(attempt, version, detail=f'{type(error).__name__}: {error}')
                    summary['deferred'] += 1
                    cfg.logger.warning('Deferred %s: %s', row.app_id, error)
                if cfg.ROW_DELAY_SECONDS:
                    time.sleep(cfg.ROW_DELAY_SECONDS)
            # Stamp every recovery row with the input hash it was tried at - whether or not the
            # gap filled. Awais Tahir's resume states no total experience and no education
            # dates at all, so three of his four blanks are facts about the document; without
            # this stamp an idle poll would re-OCR him, and every row like him, forever.
            if recovery and not dry_run:
                with store.conn:
                    for row in recovery:
                        store.conn.execute(
                            'INSERT INTO recovery(app_id,tried_hash,tried_at) VALUES(?,'
                            '(SELECT COALESCE(source_hash,\'\') FROM candidates WHERE id=?),'
                            "datetime('now')) ON CONFLICT(app_id) DO UPDATE SET "
                            'tried_hash=excluded.tried_hash,tried_at=excluded.tried_at',
                            (row.app_id, int(row.version)))
                summary['recovered'] = len(recovery)
            summary['client_export'] = str(publish_results(store, remote, upload=not dry_run,
                                                               output_dir=preview_dir))
            return summary
        finally:
            store.conn.close()


def _sync_row_to_sharepoint(remote, app_id, sheet, result_values, source_values=None):
    """Sync the scored result back to the SharePoint intake table if available.

    SharePoint_Master_File.xlsx is P1's intake ledger. P2 keeps its own main/rejected split
    in SQLite and P2-MasterFile.xlsx, so either outcome patches the original P1 row in place
    with MASTER_WRITE_COLUMNS and nothing else. A rejected result changes Status; it never
    copies profile data to, or deletes a row from, P1's workbook.
    """
    if not hasattr(remote, 'list_rows'):
        return
    from .store import ExcelCandidateStore
    sp_store = ExcelCandidateStore(remote)
    if sheet in ('main', 'rejected'):
        sp_store.save_by_id(app_id, _master_patch(result_values))

