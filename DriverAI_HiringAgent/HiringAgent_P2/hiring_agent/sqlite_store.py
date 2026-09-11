"""Durable local state. Excel row offsets never cross this interface."""
from __future__ import annotations

import hashlib
import json
import os
import sqlite3
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
INPUT_COLUMNS = ('Application ID', 'Email', 'Received Date', 'Last Updated Date',
                 'Application Updates', 'Original Filename', 'Mail Subject', 'Mail Body',
                 'Has Resume')


def encoded(value):
    return json.dumps(value, sort_keys=True, ensure_ascii=False, default=str)


def input_key(values):
    data = {k: str(values.get(k, '') or '').strip() for k in INPUT_COLUMNS}
    # P1's immutable attachment manifest is input; P2's presentation URL is not.
    url = str(values.get('Resume URL', '') or '')
    if url.startswith('manifest:'):
        data['manifest'] = url
    return hashlib.sha256(encoded(data).encode()).hexdigest()


def database_path():
    configured = os.getenv('HIRING_SQLITE_PATH')
    path = Path(configured) if configured else BASE / 'candidates.db'
    return path.resolve() if path.is_absolute() else (BASE / path).resolve()


class SQLiteCandidateStore:
    def __init__(self, db_path=None):
        global CandidateRow, RowVanished
        from .store import CandidateRow, RowVanished
        self.db_path = ':memory:' if str(db_path) == ':memory:' else Path(db_path or database_path()).resolve()
        if self.db_path != ':memory:':
            self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(str(self.db_path), timeout=30)
        self.conn.row_factory = sqlite3.Row
        with self.conn:
            self.conn.execute('PRAGMA journal_mode=WAL')
            self.conn.execute('PRAGMA synchronous=FULL')
            self.conn.execute('PRAGMA foreign_keys=ON')
            self.conn.execute('''CREATE TABLE IF NOT EXISTS candidates (
                id INTEGER PRIMARY KEY AUTOINCREMENT, app_id TEXT NOT NULL,
                sheet TEXT NOT NULL DEFAULT 'main', status TEXT NOT NULL,
                values_json TEXT NOT NULL, version INTEGER NOT NULL DEFAULT 1,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
            columns = {r['name'] for r in self.conn.execute('PRAGMA table_info(candidates)')}
            for name, spec in [('active', 'INTEGER NOT NULL DEFAULT 1'),
                               ('source_hash', "TEXT NOT NULL DEFAULT ''")]:
                if name not in columns:
                    self.conn.execute(f'ALTER TABLE candidates ADD COLUMN {name} {spec}')
            self.conn.executescript('''
                CREATE INDEX IF NOT EXISTS idx_candidates_app_id ON candidates(app_id,sheet);
                -- ONE live row per application, so 'which sheet' is a column value and not a
                -- pair of rows that can disagree. A candidate is on CandidateList or Rejected,
                -- never both: the Excel backend rejected by appending to Rejected and then
                -- deleting from main, and when Graph refused the delete (409
                -- InsertDeleteConflict) the candidate stayed on BOTH sheets. Here rejection is
                -- one UPDATE of `sheet`, and this index makes the two-row state unrepresentable
                -- rather than merely unlikely. Keyless rows are exempt - blank spacer rows are
                -- preserved deliberately and are not applications.
                CREATE TABLE IF NOT EXISTS source_events (
                    event_key TEXT PRIMARY KEY, app_id TEXT NOT NULL, payload TEXT NOT NULL,
                    imported_at TEXT DEFAULT CURRENT_TIMESTAMP);
                CREATE TABLE IF NOT EXISTS intake_files (file_key TEXT PRIMARY KEY, payload TEXT NOT NULL);
                CREATE TABLE IF NOT EXISTS candidate_history (
                    id INTEGER PRIMARY KEY, candidate_id INTEGER NOT NULL,
                    version INTEGER NOT NULL, payload TEXT NOT NULL,
                    recorded_at TEXT DEFAULT CURRENT_TIMESTAMP);
                CREATE TABLE IF NOT EXISTS import_conflicts (
                    app_id TEXT PRIMARY KEY, reason TEXT NOT NULL, payload TEXT NOT NULL);
                CREATE TABLE IF NOT EXISTS attempts (
                    id INTEGER PRIMARY KEY, candidate_id INTEGER NOT NULL,
                    input_version INTEGER NOT NULL, state TEXT NOT NULL,
                    detail TEXT NOT NULL DEFAULT '', started_at TEXT DEFAULT CURRENT_TIMESTAMP,
                    finished_at TEXT);
                -- One row per application whose gaps an idle pass has already tried to fill.
                -- Keyed on the input hash so a genuinely NEW resume re-opens the attempt while
                -- an unchanged one is never re-read: without this the same 27 rows whose CVs
                -- simply do not state a total would be re-OCR'd on every idle poll forever.
                CREATE TABLE IF NOT EXISTS recovery (
                    app_id TEXT PRIMARY KEY, tried_hash TEXT NOT NULL DEFAULT '',
                    tried_at TEXT DEFAULT CURRENT_TIMESTAMP);
                CREATE TABLE IF NOT EXISTS publication (
                    id INTEGER PRIMARY KEY CHECK(id=1), revision INTEGER NOT NULL DEFAULT 0,
                    published_revision INTEGER NOT NULL DEFAULT -1, error TEXT NOT NULL DEFAULT '');
                INSERT OR IGNORE INTO publication(id) VALUES(1);
            ''')
            # ONE live row per application, so 'which sheet' is a column value rather than a
            # pair of rows that can disagree. A candidate belongs to CandidateList or Rejected,
            # never both: the Excel backend rejected by appending to Rejected then deleting
            # from main, and when Graph refused the delete (409 InsertDeleteConflict) the
            # candidate stayed on BOTH sheets. Rejection here is one UPDATE of `sheet`, and
            # this index makes the two-row state unrepresentable rather than merely unlikely.
            # Keyless rows are exempt: blank spacer rows are preserved deliberately.
            #
            # Tolerated rather than required, because a database written before this index may
            # already hold duplicates and must still OPEN - refusing to connect would strand
            # the very data that needs repairing. Those databases keep the older defences:
            # _target()'s hinted resolution, and build_master()'s cross-sheet `seen` guard,
            # which fails the publication closed instead of shipping a doubled master.
            try:
                self.conn.execute('CREATE UNIQUE INDEX IF NOT EXISTS idx_candidates_one_live '
                                  "ON candidates(app_id) WHERE active=1 AND app_id<>''")
            except sqlite3.IntegrityError:
                from . import config as cfg
                dupes = [r[0] for r in self.conn.execute(
                    "SELECT app_id FROM candidates WHERE active=1 AND app_id<>'' "
                    'GROUP BY app_id HAVING COUNT(*)>1')]
                cfg.logger.warning('Existing duplicate application rows prevent the one-live-row '
                                   'index; publication fails closed until repaired: %s', dupes)

    def _dirty(self):
        self.conn.execute('UPDATE publication SET revision=revision+1 WHERE id=1')

    def _history(self, row):
        self.conn.execute('INSERT INTO candidate_history(candidate_id,version,payload) VALUES(?,?,?)',
                          (row['id'], row['version'], encoded(dict(row))))

    def _matches(self, app_id, sheet):
        return self.conn.execute('SELECT * FROM candidates WHERE app_id=? AND sheet=? AND active=1 ORDER BY id',
                                 (str(app_id).strip(), sheet)).fetchall()

    def _target(self, app_id, sheet='main', hint=None):
        rows = self._matches(app_id, sheet)
        if hint is not None:
            rows = [r for r in rows if str(r['id']) == str(hint)]
        if len(rows) != 1:
            raise RowVanished(f'{app_id}: expected one {sheet} record, found {len(rows)}; no write performed')
        return rows[0]

    def all_rows(self, sheet='main'):
        return [CandidateRow(app_id=r['app_id'], values=json.loads(r['values_json']),
                             version=str(r['id']), sheet=sheet)
                for r in self.conn.execute('SELECT * FROM candidates WHERE sheet=? AND active=1 ORDER BY id', (sheet,))]

    def get(self, app_id, sheet='main'):
        if not self._matches(app_id, sheet):
            return None
        r = self._target(app_id, sheet)
        return CandidateRow(app_id=r['app_id'], values=json.loads(r['values_json']), version=str(r['id']), sheet=sheet)

    def get_queue(self, statuses, limit=None):
        wanted = {statuses} if isinstance(statuses, str) else set(statuses)
        rows = [r for r in self.all_rows() if r.app_id and r.values.get('Status') in wanted]
        return rows if limit is None else rows[:limit]

    def resolve_position(self, app_id, sheet='main', hint=None):
        return self._target(app_id, sheet, hint)['id']

    def save_by_id(self, app_id, fields, *, current_values=None, sheet='main', hint=None):
        with self.conn:
            r = self._target(app_id, sheet, hint)
            vals = json.loads(r['values_json'])
            # Never merge a caller's stale whole-row snapshot over newer stored values.
            vals.update(fields)
            if str(vals.get('Application ID', app_id)).strip() != app_id:
                raise ValueError('Application identity cannot be changed by a field patch')
            self._history(r)
            self.conn.execute('''UPDATE candidates SET values_json=?,status=?,version=version+1,
                updated_at=CURRENT_TIMESTAMP WHERE id=?''', (encoded(vals), vals.get('Status', ''), r['id']))
            self._dirty()

    def add(self, fields, sheet='main'):
        with self.conn:
            self.conn.execute('INSERT INTO candidates(app_id,sheet,status,values_json) VALUES(?,?,?,?)',
                              (str(fields.get('Application ID', '')).strip(), sheet,
                               str(fields.get('Status', '')), encoded(fields)))
            self._dirty()

    def delete_by_id(self, app_id, sheet='main', hint=None):
        if not self._matches(app_id, sheet):
            return False
        with self.conn:
            r = self._target(app_id, sheet, hint)
            self._history(r)
            self.conn.execute('UPDATE candidates SET active=0,version=version+1 WHERE id=?', (r['id'],))
            self._dirty()
        return True

    def delete_at(self, index, sheet='main'):
        # Local adapter indexes ARE persistent primary keys, never worksheet offsets.
        r = self.conn.execute('SELECT app_id FROM candidates WHERE id=? AND sheet=? AND active=1', (index, sheet)).fetchone()
        if r:
            self.delete_by_id(r['app_id'], sheet, hint=index)

    def move_to_rejected(self, app_id, fields, hint=None):
        if not self._matches(app_id, 'main'):
            existing = self.get(app_id, 'rejected')
            if existing:
                return True
            raise RowVanished(f'{app_id}: no source record to reject')
        with self.conn:
            r = self._target(app_id, 'main', hint)
            if self._matches(app_id, 'rejected'):
                raise RowVanished(f'{app_id}: conflicting rejected record requires reconciliation')
            vals = json.loads(r['values_json'])
            vals.update(fields)
            self._history(r)
            self.conn.execute('''UPDATE candidates SET sheet='rejected',values_json=?,status=?,
                version=version+1 WHERE id=?''', (encoded(vals), vals.get('Status', ''), r['id']))
            self._dirty()
        return True

    def sync_from_client(self, client, force=False):
        """Incremental, non-destructive import. Record every input before reconciliation."""
        # All remote I/O precedes the transaction. Failure cannot partly import a snapshot.
        source = [(sheet, dict(row['values'])) for sheet, reader in
                  [('main', client.list_rows), ('rejected', client.list_rejected_rows)] for row in reader()]
        if hasattr(client, 'list_intake_events'):
            client._intake_event_cache = {r['file_key']: json.loads(r['payload'])
                                         for r in self.conn.execute('SELECT * FROM intake_files')}
            source.extend(('main', dict(row['values'])) for row in client.list_intake_events())
        groups = {}
        for sheet, vals in source:
            app = str(vals.get('Application ID', '')).strip()
            if app:
                groups.setdefault(app, []).append((sheet, vals))
        changed = 0
        with self.conn:
            for file_key, payload in getattr(client, '_intake_event_cache', {}).items():
                self.conn.execute('INSERT OR IGNORE INTO intake_files VALUES(?,?)', (file_key, encoded(payload)))
            for app, incoming in groups.items():
                for sheet, vals in incoming:
                    event = hashlib.sha256(encoded([app, sheet, vals]).encode()).hexdigest()
                    self.conn.execute('INSERT OR IGNORE INTO source_events(event_key,app_id,payload) VALUES(?,?,?)',
                                      (event, app, encoded([sheet, vals])))
                existing = self.conn.execute('SELECT * FROM candidates WHERE app_id=? AND active=1 ORDER BY id', (app,)).fetchall()
                emails = {str(v.get('Email', '')).strip().lower() for _, v in incoming}
                emails |= {str(json.loads(r['values_json']).get('Email', '')).strip().lower() for r in existing}
                emails.discard('')
                def order(pair):
                    _, v = pair
                    return (str(v.get('Last Updated Date') or v.get('Received Date') or ''),
                            str(v.get('Received Date') or ''))
                newest = max(order(pair) for pair in incoming)
                latest = [pair for pair in incoming if order(pair) == newest]
                if len(emails) > 1 or len({input_key(v) for _, v in latest}) > 1:
                    self.conn.execute('INSERT OR REPLACE INTO import_conflicts VALUES(?,?,?)',
                                      (app, 'Conflicting identities or same-time input versions', encoded(incoming)))
                    continue
                sheet, vals = max(latest, key=lambda pair: pair[1].get('Status') != 'New Email Received')
                key = input_key(vals)
                if existing:
                    # Resolve legacy completed/retry copies only when input and identity agree.
                    hashes = {r['source_hash'] or input_key(json.loads(r['values_json'])) for r in existing}
                    if len(existing) > 1 and len(hashes) > 1:
                        self.conn.execute('INSERT OR REPLACE INTO import_conflicts VALUES(?,?,?)',
                                          (app, 'Legacy duplicate inputs need explicit reconciliation', encoded([dict(r) for r in existing])))
                        continue
                    target = max(existing, key=lambda r: (r['status'] != 'New Email Received', r['version']))
                    for other in existing:
                        if other['id'] != target['id']:
                            self._history(other)
                            self.conn.execute('UPDATE candidates SET active=0 WHERE id=?', (other['id'],))
                            self._dirty()
                    old = json.loads(target['values_json'])
                    old_key = target['source_hash'] or input_key(old)
                    if key == old_key:
                        self.conn.execute('UPDATE candidates SET source_hash=? WHERE id=?', (key, target['id']))
                        self.conn.execute('DELETE FROM import_conflicts WHERE app_id=?', (app,))
                        continue
                    if order((sheet, vals)) < order((target['sheet'], old)):
                        continue  # late delivery retained in source_events, never rolled back
                    self._history(target)
                    # New input invalidates the old result. Retain old result in history only.
                    refreshed = dict(vals)
                    refreshed['Status'] = 'New Email Received'
                    refreshed['Retry Count'] = 0
                    self.conn.execute('''UPDATE candidates SET sheet='main',status=?,values_json=?,
                        source_hash=?,version=version+1 WHERE id=?''',
                        (refreshed['Status'], encoded(refreshed), key, target['id']))
                else:
                    self.conn.execute('INSERT INTO candidates(app_id,sheet,status,values_json,source_hash) VALUES(?,?,?,?,?)',
                                      (app, sheet, vals.get('Status', ''), encoded(vals), key))
                self.conn.execute('DELETE FROM import_conflicts WHERE app_id=?', (app,))
                self._dirty()
                changed += 1
        return changed

    def begin_attempt(self, candidate_id):
        with self.conn:
            r = self.conn.execute('SELECT * FROM candidates WHERE id=? AND active=1', (candidate_id,)).fetchone()
            cur = self.conn.execute("INSERT INTO attempts(candidate_id,input_version,state) VALUES(?,?,'running')",
                                    (candidate_id, r['version']))
        return cur.lastrowid, r['version']

    def finish_attempt(self, attempt_id, expected_version, result=None, detail=''):
        with self.conn:
            a = self.conn.execute('SELECT * FROM attempts WHERE id=?', (attempt_id,)).fetchone()
            r = self.conn.execute('SELECT * FROM candidates WHERE id=?', (a['candidate_id'],)).fetchone()
            state = 'deferred'
            if r['version'] != expected_version or not r['active']:
                state = 'superseded'
            elif result:
                self._history(r)
                sheet, vals = result
                self.conn.execute('''UPDATE candidates SET sheet=?,values_json=?,status=?,version=version+1
                    WHERE id=? AND version=?''', (sheet, encoded(vals), vals.get('Status', ''), r['id'], expected_version))
                self._dirty()
                state = 'completed'
            self.conn.execute('UPDATE attempts SET state=?,detail=?,finished_at=CURRENT_TIMESTAMP WHERE id=?',
                              (state, detail, attempt_id))
        return state

    def backup(self, path):
        dest = sqlite3.connect(str(path))
        try:
            self.conn.backup(dest)
        finally:
            dest.close()
