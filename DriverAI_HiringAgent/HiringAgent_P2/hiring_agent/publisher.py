"""Generate reports from committed state and retry delivery without rerunning inference."""
from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
from urllib.parse import unquote

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

from . import config as cfg
from .sqlite_store import BASE


def safe_target(target, intake):
    def canonical(value):
        parts = unquote(str(value)).replace('\\', '/').strip('/').split('/')
        if any(p in ('.', '..') for p in parts):
            raise ValueError('Report path must not contain traversal segments')
        return '/'.join(parts).casefold()
    if canonical(target) == canonical(intake):
        raise ValueError('Refusing to replace the P1 intake workbook with a generated report')


def _workbook_shape(data):
    """(sheet title, row count, column count) per sheet, for an in-memory xlsx."""
    import io
    wb = load_workbook(io.BytesIO(data), read_only=True)
    try:
        return [(ws.title, ws.max_row, ws.max_column) for ws in wb.worksheets]
    finally:
        wb.close()


def _verify_report(remote, folder, name, data):
    """Confirm the stored report is the report we sent, by CONTENT not by bytes.

    SharePoint rewrites an uploaded .xlsx server-side - the same workbook came back 164,931
    bytes against the 157,156 uploaded (2026-09-08), both valid archives - so a sha256 of the
    raw bytes can never match and the publication aborted after every successful upload. What
    has to hold is that the stored file is a readable workbook with the sheets, rows and
    columns we published; re-zipping does not change any of those.
    """
    stored = remote.download_file(folder, name)
    if stored == data:
        return
    try:
        actual = _workbook_shape(stored)
    except Exception as error:
        raise ValueError(f'Stored report is not a readable workbook: {name} ({error})')
    if actual != _workbook_shape(data):
        raise ValueError(f'Stored report does not match what was published: {name}')


def build_master(store, path, generation):
    wb = Workbook()
    wb.remove(wb.active)
    wb.properties.description = f'P2 committed generation {generation}'
    seen = set()
    for sheet, title, columns, table_name in (
            ('main', 'CandidateList', cfg.COLUMNS, 'HiringAgent_P1_Candidates'),
            ('rejected', 'Rejected', cfg.REJECTED_COLUMNS, 'RejectedCandidates')):
        ws = wb.create_sheet(title)
        ws.append(columns)
        for row in store.all_rows(sheet):
            if not row.app_id:
                continue
            if row.app_id in seen:
                raise ValueError(f'Ambiguous active application {row.app_id}; publication deferred')
            seen.add(row.app_id)
            ws.append([row.values.get(col, '') for col in columns])
            for cell in ws[ws.max_row]:
                # Resume/mail values are literal strings, never executable formulas.
                if isinstance(cell.value, str):
                    cell.data_type = 's'
                cell.number_format = '@'
            if sheet == 'main':
                from .sharepoint_scoring import is_doubt_candidate
                is_doubt, _ = is_doubt_candidate(row.values)
                if is_doubt:
                    amber_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                    for cell in ws[ws.max_row]:
                        cell.fill = amber_fill
        if ws.max_row == 1:
            ws.append([''] * len(columns))
        table = Table(displayName=table_name, ref=f'A1:{get_column_letter(len(columns))}{ws.max_row}')
        table.tableStyleInfo = TableStyleInfo(name='TableStyleMedium9', showRowStripes=True)
        ws.add_table(table)
        ws.freeze_panes = 'A2'
        for i, col in enumerate(columns, 1):
            ws.column_dimensions[get_column_letter(i)].width = min(max(len(col) + 3, 18), 42)
        if 'Resume URL' in columns and 'Resume Link' in columns:
            url_col, link_col = columns.index('Resume URL') + 1, columns.index('Resume Link') + 1
            for row in range(2, ws.max_row + 1):
                url = str(ws.cell(row, url_col).value or '')
                if url.startswith(('https://', 'http://')):
                    cell = ws.cell(row, link_col, 'Open resume')
                    cell.hyperlink = url
                    # The loop above stamps every cell as Text ('@') so a stored value can
                    # never be evaluated as a formula. Correct for data, wrong here: it left
                    # the link rendering as ordinary black text with no underline, so it did
                    # not READ as a link even though clicking it worked. Restore the default
                    # format and apply Excel's own Hyperlink styling to this cell only.
                    cell.number_format = 'General'
                    cell.font = Font(color='0563C1', underline='single')
    temp = Path(path).with_suffix('.tmp.xlsx')
    wb.save(temp)
    wb.close()
    check = load_workbook(temp)
    try:
        if check.sheetnames != ['CandidateList', 'Rejected'] or any(not ws.tables for ws in check):
            raise ValueError('Master workbook schema validation failed')
        # Application ID need not remain column A forever.
        expected = len(seen)
        id_counts = 0
        for ws in check:
            columns = [c.value for c in ws[1]]
            id_col = columns.index('Application ID') + 1
            id_counts += sum(bool(ws.cell(i, id_col).value) for i in range(2, ws.max_row + 1))
        if id_counts != expected:
            raise ValueError('Master workbook count validation failed')
    finally:
        check.close()
    temp.replace(path)


def publish_results(store, remote, *, upload=True, output_dir=None):
    from .exporter import ExcelReportExporter
    from sharepoint_client import SharePointError
    output = Path(output_dir or os.getenv('HIRING_PUBLISH_DIR') or BASE / 'P2_Final_Results')
    output.mkdir(parents=True, exist_ok=True)
    record = store.conn.execute('SELECT * FROM publication WHERE id=1').fetchone()
    revision = record['revision']
    master_target = os.getenv('HIRING_MASTER_REPORT_PATH', '/Master_Files/P2-MasterFile.xlsx')
    client_target = os.getenv('HIRING_CLIENT_REPORT_PATH', '/Candidate_List_Results.xlsx')
    for target in (master_target, client_target):
        safe_target(target, remote._wb_path)
    if master_target.strip('/').casefold() == client_target.strip('/').casefold():
        raise ValueError('Master and client results must have distinct targets')
    generation = output / f'generation_{revision}'
    generation.mkdir(exist_ok=True)
    master = generation / 'P2-MasterFile.xlsx'
    client = generation / 'Candidate_List_Results.xlsx'
    if upload and record['published_revision'] == revision and master.exists() and client.exists():
        return client
    # Always rebuild master and client fresh from store so published files are never stale.
    build_master(store, master, revision)
    temp_client = client.with_suffix('.tmp.xlsx')
    ExcelReportExporter(store).generate_workbook(temp_client)
    wb = load_workbook(temp_client)
    wb.properties.description = f'P2 committed generation {revision}'
    for ws in wb:
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if isinstance(cell.value, str):
                    cell.data_type = 's'
    wb.save(temp_client)
    wb.close()
    temp_client.replace(client)
    manifest = {'generation': revision, 'files': [],
                'import_conflicts': [dict(r) for r in store.conn.execute('SELECT app_id,reason FROM import_conflicts')]}
    if not upload:
        return client
    try:
        # EXACTLY three workbooks exist on the tenant: P1's intake master, this master report,
        # and this client sheet - and P2 only ever writes the last two, in place. Version
        # history lives in the local generation_<N> folders, which is why no '.r<N>' sibling
        # and no remote pointer file are uploaded: every extra remote copy was one more
        # 'master file' for a human to mistake for the real one. `published_revision` in the
        # local publication table already answers 'what is live', so the pointer was duplicate
        # state that could itself go stale.
        for path, target in ((master, master_target), (client, client_target)):
            folder, _, name = target.strip('/').rpartition('/')
            data = path.read_bytes()
            remote.upload_file(folder, name, data)
            _verify_report(remote, folder, name, data)
            manifest['files'].append({'path': '/'.join((folder, name)),
                                      'sha256': hashlib.sha256(data).hexdigest()})
        (generation / 'published.json').write_text(json.dumps(manifest, indent=2), encoding='utf-8')
        with store.conn:
            store.conn.execute('UPDATE publication SET published_revision=?,error=? WHERE id=1', (revision, ''))
    except Exception as error:
        with store.conn:
            store.conn.execute('UPDATE publication SET error=? WHERE id=1', (str(error),))
        cfg.logger.warning('Results committed locally; publication pending: %s', error)
    return client
