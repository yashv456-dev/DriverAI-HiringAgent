"""SharePoint / Microsoft Graph client for the hiring scoring worker (APP-ONLY auth).

Authenticates as the APP itself (client-credentials flow) - no user, no browser,
no per-machine login. Drop the .env on any machine and it runs headless. Used by
bot.py's --score-sharepoint mode to:
  - read unscored rows from the candidate Excel table (the work queue),
  - download each resume file,
  - write the scored fields back to the row,
  - (optionally) upload a per-candidate scorecard.

All Graph access goes through the Excel "workbook" and "drive" REST endpoints, so
row updates are surgical (PATCH one row) and never lock or overwrite the whole file
- the Power Automate flow can keep appending rows at the same time.

Requires (in .env):
  TENANT_ID               real tenant GUID (NOT 'common' - client-credentials needs a tenant)
  CLIENT_ID               the app registration's Application (client) ID
  CLIENT_SECRET           a client secret value for that app
  SHAREPOINT_HOSTNAME     e.g. led1234567.sharepoint.com
  SHAREPOINT_SITE_PATH    e.g. /sites/CandidateList_HiringAgent
  SHAREPOINT_TABLE        e.g. HiringAgent_P1_Candidates
  SHAREPOINT_RESUMES_FOLDER  e.g. /Downloaded_Resumes

Optional (in .env):
  SHAREPOINT_WORKBOOK     e.g. /P1P2_SharePoint_Master_Files/Sharepoint_Master_File.xlsx
                          If omitted, defaults to /P1P2_SharePoint_Master_Files/Sharepoint_Master_File.xlsx
                          (single workbook for all years — no <Year> subfolder)
  SENDER_MAILBOX          mailbox to send non-USA decline emails FROM (default: apply@driverai.io)
                          Needs Mail.Send Application permission admin-consented on the app.
"""
import base64
import io
import json
import os
import re

import requests

GRAPH = "https://graph.microsoft.com/v1.0"
_TOKEN_URL = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
_TIMEOUT = 60
_WB_NAME = "Sharepoint_Master_File.xlsx"
_DEFAULT_WORKBOOK_FOLDER = "/Master_Files"


class SharePointError(RuntimeError):
    """Any Graph/config failure - carries a human-readable message for the logs."""

    def __init__(self, message: str, status_code: int | None = None):
        super().__init__(message)
        self.status_code = status_code


def check_graph_reachable() -> tuple[bool, str]:
    """Cheap, short-timeout probe for whether Microsoft Graph/Entra ID is reachable at all.

    Modeled on ollama_health()'s exact shape - never raises, returns (ok, detail). Hits a
    public, credential-free discovery endpoint so it works before a SharePointClient (and
    its .env-sourced tenant/client secret) even exists. Used to fail a --score-sharepoint
    run fast and clearly when there's no internet, instead of letting the raw network
    exception from _get_token() (or a generic SharePointError from the first real Graph
    call) surface a few steps into the run.
    """
    try:
        resp = requests.get(
            "https://login.microsoftonline.com/common/.well-known/openid-configuration",
            timeout=5,
        )
        resp.raise_for_status()
        return True, "Internet/Graph reachable"
    except Exception as e:
        return False, (
            f"No internet - SharePoint mode needs it to read the queue, download resumes, "
            f"and write results back. Nothing to do until connectivity is restored. ({e})"
        )


def _build_candidate_workbook_bytes(columns: list[str], table_name: str,
                                    sheet_name: str = "CandidateList") -> bytes:
    """Build a Graph-safe candidate workbook with one intentionally blank seed row."""
    from openpyxl import Workbook
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    for i, col in enumerate(columns, 1):
        ws.cell(row=1, column=i, value=col)
        ws.cell(row=2, column=i, value="")
    ref = f"A1:{get_column_letter(len(columns))}2"
    tab = Table(displayName=table_name, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9", showFirstColumn=False,
        showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    ws.add_table(tab)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class SharePointClient:
    """Thin, app-only Graph wrapper scoped to ONE SharePoint site + workbook."""

    def __init__(self) -> None:
        self.tenant = (os.getenv("TENANT_ID") or "").strip()
        self.client_id = (os.getenv("CLIENT_ID") or "").strip()
        self.client_secret = (os.getenv("CLIENT_SECRET") or "").strip()
        self.hostname = (os.getenv("SHAREPOINT_HOSTNAME") or "").strip()
        self.site_path = (os.getenv("SHAREPOINT_SITE_PATH") or "").strip()
        self.table = (os.getenv("SHAREPOINT_TABLE") or "").strip()
        self.resumes_folder = (os.getenv("SHAREPOINT_RESUMES_FOLDER") or "").strip().rstrip("/")
        # Mailbox the non-USA decline notice is sent FROM (needs Mail.Send Application consent).
        self.sender_mailbox = (os.getenv("SENDER_MAILBOX") or "apply@driverai.io").strip()

        missing = [k for k, v in {
            "TENANT_ID": self.tenant, "CLIENT_ID": self.client_id,
            "CLIENT_SECRET": self.client_secret, "SHAREPOINT_HOSTNAME": self.hostname,
            "SHAREPOINT_SITE_PATH": self.site_path,
            "SHAREPOINT_TABLE": self.table,
        }.items() if not v]
        if missing:
            raise SharePointError("Missing .env settings: " + ", ".join(missing))
        if self.tenant.lower() == "common":
            raise SharePointError(
                "TENANT_ID must be your real tenant GUID for app-only auth, not 'common'."
            )

        self._token = None
        self._session = requests.Session()
        self._site_id = None

        explicit = (os.getenv("SHAREPOINT_WORKBOOK") or "").strip()
        if explicit:
            self._wb_path = explicit if explicit.startswith("/") else "/" + explicit
        else:
            self._wb_path = f"{_DEFAULT_WORKBOOK_FOLDER}/{_WB_NAME}"
        self._columns_cache = {}
        self._wb_id = None

    # ---- auth --------------------------------------------------------------
    def _get_token(self) -> str:
        """Acquire and cache an app-only access token via client credentials.

        Entra token requests sit outside ``_req``'s Graph retry loop. Retry transient
        network/server failures here as well so a brief connection reset does not abort
        an otherwise healthy scheduled scoring run.
        """
        import time as _time

        if self._token:
            return self._token
        resp = None
        for attempt in range(4):
            try:
                resp = requests.post(
                    _TOKEN_URL.format(tenant=self.tenant),
                    data={
                        "grant_type": "client_credentials",
                        "client_id": self.client_id,
                        "client_secret": self.client_secret,
                        "scope": "https://graph.microsoft.com/.default",
                    },
                    timeout=_TIMEOUT,
                )
            except requests.RequestException as e:
                if attempt >= 3:
                    raise SharePointError(
                        f"Token request could not reach Microsoft after 4 attempts: {e}"
                    ) from e
                wait = 2 * (attempt + 1)
                from hiring_agent.config import logger
                logger.warning(
                    f"   Entra token connection failed - retrying in {wait}s "
                    f"({attempt + 1}/3): {e}"
                )
                _time.sleep(wait)
                continue

            if resp.status_code in (429, 500, 502, 503, 504) and attempt < 3:
                try:
                    retry_after = int(resp.headers.get("Retry-After", "") or 0)
                except ValueError:
                    retry_after = 0
                wait = min(retry_after or 2 * (attempt + 1), 30)
                from hiring_agent.config import logger
                logger.warning(
                    f"   Entra token service returned {resp.status_code} - retrying "
                    f"in {wait}s ({attempt + 1}/3)..."
                )
                _time.sleep(wait)
                continue
            break

        if resp is None:
            raise SharePointError("Token request failed before Microsoft returned a response.")
        if resp.status_code != 200:
            raise SharePointError(
                f"Token request failed ({resp.status_code}). "
                f"Check CLIENT_ID/SECRET/TENANT_ID and that admin consent was granted. "
                f"Detail: {resp.text[:300]}"
            )
        self._token = resp.json()["access_token"]
        self._session.headers.update({"Authorization": f"Bearer {self._token}"})
        return self._token

    def application_roles(self) -> set[str]:
        """Application permissions carried by the current Graph access token.

        This is a read-only diagnostic used by ``bot.py --test-sharepoint`` to
        confirm optional capabilities such as Mail.Send without sending a test
        message to a real mailbox.
        """
        token = self._get_token()
        try:
            payload = token.split(".")[1]
            payload += "=" * (-len(payload) % 4)
            claims = json.loads(base64.urlsafe_b64decode(payload.encode("ascii")))
            return {str(role) for role in claims.get("roles", []) if role}
        except (IndexError, ValueError, TypeError, json.JSONDecodeError) as e:
            raise SharePointError(
                f"Could not inspect Graph application permissions in the access token: {e}"
            ) from e

    def _req(self, method: str, url: str, **kw):
        """Authenticated Graph request with a clear error on failure.

        Throttling/transient errors (429, 503, 504) are retried up to 3 times with
        backoff, honoring Retry-After when Graph sends one — an evening of Excel-API
        throttling used to kill a run mid-pass; now it just slows down.
        """
        import time as _time
        self._get_token()
        r = None
        for attempt in range(4):
            r = self._session.request(method, url, timeout=_TIMEOUT, **kw)
            if r.status_code == 401:  # token expired mid-run - refresh once and retry
                self._token = None
                self._get_token()
                r = self._session.request(method, url, timeout=_TIMEOUT, **kw)
            if r.status_code in (429, 503, 504) and attempt < 3:
                wait = 0
                try:
                    wait = int(r.headers.get("Retry-After", "") or 0)
                except ValueError:
                    pass
                wait = min(wait or 10 * (attempt + 1), 120)
                from hiring_agent.config import logger
                logger.warning(f"   Graph {r.status_code} (throttled/unavailable) — "
                               f"retrying in {wait}s ({attempt + 1}/3)...")
                _time.sleep(wait)
                continue
            break
        if not r.ok:
            raise SharePointError(f"{method} {url.split('/v1.0')[-1]} -> {r.status_code}: {r.text[:300]}",
                                  status_code=r.status_code)
        return r

    # ---- site / workbook plumbing -----------------------------------------
    def site_id(self) -> str:
        """Resolve the SharePoint site's Graph id from hostname + path (cached)."""
        if self._site_id:
            return self._site_id
        url = f"{GRAPH}/sites/{self.hostname}:{self.site_path}"
        self._site_id = self._req("GET", url).json()["id"]
        return self._site_id

    def _wb_base(self) -> str:
        """Base URL for the workbook, resolved and cached by persistent ID to prevent propagation delays."""
        if not self._wb_id:
            url = f"{GRAPH}/sites/{self.site_id()}/drive/root:{self._wb_path}"
            self._wb_id = self._req("GET", url).json()["id"]
        return f"{GRAPH}/sites/{self.site_id()}/drive/items/{self._wb_id}/workbook"

    def table_columns(self) -> list:
        """Column header names in table order (so we build row arrays correctly)."""
        return self._table_columns_of(self.table)

    def workbook_worksheets(self) -> list:
        """Worksheet names in the configured workbook."""
        url = f"{self._wb_base()}/worksheets?$select=name"
        return [s.get("name", "") for s in self._req("GET", url).json().get("value", [])]

    def workbook_tables(self) -> list:
        """Table names in the configured workbook."""
        url = f"{self._wb_base()}/tables?$select=name"
        return [t.get("name", "") for t in self._req("GET", url).json().get("value", [])]

    def workbook_metadata(self) -> dict:
        """Readable workbook/site metadata for diagnostics and the desktop app."""
        return {
            "hostname": self.hostname,
            "site_path": self.site_path,
            "resumes_folder": self.resumes_folder,
            "workbook_path": self._wb_path,
            "workbook_name": self._wb_path.rsplit("/", 1)[-1],
            "table": self.table,
            "worksheets": self.workbook_worksheets(),
            "tables": self.workbook_tables(),
            "columns": self.table_columns(),
        }

    def ensure_columns(self, required: list) -> list:
        """Add any of `required` columns the live table is missing, inserted at its
        canonical position (from config.COLUMNS) relative to whichever neighbors already
        exist - not always tacked onto the far right. A column added to the schema after
        the table was first created (e.g. a brand new field) still lands next to its
        logical neighbors instead of stranded at the end. Existing rows get blank cells.
        Safe for the Power Automate flow (it maps by column name). Returns the list of
        columns that were added."""
        from hiring_agent.config import COLUMNS as _CANON
        existing = list(self.table_columns())
        added = []
        for col in required:
            if col not in existing:
                url = f"{self._wb_base()}/tables/{self.table}/columns/add"
                payload = {"name": col}
                if col in _CANON:
                    # Insertion index = how many canonical-order columns before `col`
                    # are already on the live table (order-independent within this loop -
                    # a column added earlier this pass is already in `existing`).
                    before = _CANON[:_CANON.index(col)]
                    payload["index"] = sum(1 for c in before if c in existing)
                self._req("POST", url, json=payload)
                existing.append(col)
                added.append(col)
        if added:
            self._columns_cache.pop(self.table, None)
            from hiring_agent.config import logger
            logger.info(f"   added missing column(s) to '{self.table}': {', '.join(added)}")
        return added

    def ensure_workbook(self) -> None:
        """Confirm the single candidate workbook exists on SharePoint. NEVER creates or
        overwrites it — mirrors Phase 1's non-destructive design exactly (P1 dropped its own
        workbook auto-create on 2026-07-04 after an unreliable probe kept false-reporting
        'missing' and re-uploaded a blank template over real data, wiping candidate rows).

        Raises SharePointError on ANY failure to confirm the table, 404 or otherwise, so the
        caller always finds out and can alert/abort rather than silently limping forward.
        A genuinely missing workbook is a one-time human setup step
        (P1_Templates/HiringAgent_P1_CandidateList.xlsx), never something either phase
        auto-creates.
        """
        self.table_columns()

    # ---- the work queue ----------------------------------------------------
    def _rows_paged(self, url: str) -> list:
        """GET every row of a table, following @odata.nextLink (Graph pages large tables)."""
        out = []
        while url:
            data = self._req("GET", url).json()
            out.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
        return out

    def list_rows(self) -> list:
        """All table rows as a list of dicts: {'index': int, 'values': {col: val}}.

        index is the row's position in the table (what PATCH itemAt uses). We also
        return the raw values dict keyed by column name for easy filtering."""
        cols = self.table_columns()
        rows = self._rows_paged(f"{self._wb_base()}/tables/{self.table}/rows")
        out = []
        for r in rows:
            vals = r.get("values", [[]])
            cells = vals[0] if vals else []
            mapped = {cols[i]: (cells[i] if i < len(cells) else "") for i in range(len(cols))}
            if not any(str(v or "").strip() for v in mapped.values()):
                continue
            out.append({"index": r.get("index"), "values": mapped})
        return out

    def list_unscored_rows(self, statuses=("New Email Received",)) -> list:
        """Rows still waiting to be scored (Status in the given set — the PA-written
        'new' status plus, for the scoring worker, 'Needs Review' rows so a resume
        that becomes readable after a parser fix is picked up automatically)."""
        if isinstance(statuses, str):
            statuses = (statuses,)
        wanted = {s.strip() for s in statuses}
        return [r for r in self.list_rows()
                if str(r["values"].get("Status", "")).strip() in wanted]

    def _column_formula_at(self, table_name: str, col_name: str, index: int) -> str:
        """The raw formula in one cell (table `col_name`, body row `index`), or '' if none.

        Reading a table row returns computed VALUES, so a =HYPERLINK cell reads back as its
        display text; rewriting that text would flatten the formula. This lets a row rewrite
        put the real formula back so the clickable link survives."""
        import urllib.parse
        colref = f"columns('{urllib.parse.quote(col_name)}')"
        try:
            data = self._req("GET", f"{self._wb_base()}/tables/{table_name}/{colref}"
                                    f"/dataBodyRange?$select=formulas").json()
        except SharePointError:
            return ""
        forms = data.get("formulas", [])
        if 0 <= index < len(forms) and forms[index]:
            return str(forms[index][0] or "")
        return ""

    def _column_letter(self, table_name: str, col_name: str) -> str | None:
        """The worksheet column letter (e.g. 'F') a table column currently occupies, or None
        if it can't be resolved (missing column, throttled, etc.) — shared by every method
        that needs a whole-column worksheet range address for one table column."""
        try:
            addr = self._req("GET", f"{self._wb_base()}/tables/{table_name}/range?$select=address"
                             ).json()["address"]
            m = re.search(r"!([A-Z]+)\d+", addr)
            if not m:
                return None
            n = 0
            for ch in m.group(1):
                n = n * 26 + (ord(ch) - 64)
            cols = self._table_columns_of(table_name)
            if col_name not in cols:
                return None
            n += cols.index(col_name)
            letter = ""
            while n:
                n, r = divmod(n - 1, 26)
                letter = chr(65 + r) + letter
            return letter
        except SharePointError:
            return None

    def rename_column(self, table_name: str, old_name: str, new_name: str) -> None:
        """Rename a table column in place (best-effort) — Graph preserves every row's data,
        just relabels the header. No-op if `old_name` doesn't exist (already renamed, or a
        fresh table created straight under `new_name`)."""
        if old_name not in self._table_columns_of(table_name):
            return
        import urllib.parse
        colref = f"columns('{urllib.parse.quote(old_name)}')"
        try:
            self._req("PATCH", f"{self._wb_base()}/tables/{table_name}/{colref}",
                      json={"name": new_name})
            self._columns_cache.pop(table_name, None)
        except SharePointError:
            pass

    def delete_column(self, table_name: str, col_name: str) -> None:
        """Delete a table column entirely, header and all row data (best-effort). No-op if
        the column doesn't exist (already deleted, or never existed on this table)."""
        if col_name not in self._table_columns_of(table_name):
            return
        import urllib.parse
        colref = f"columns('{urllib.parse.quote(col_name)}')"
        try:
            self._req("DELETE", f"{self._wb_base()}/tables/{table_name}/{colref}")
            self._columns_cache.pop(table_name, None)
        except SharePointError:
            pass

    def clear_column_hyperlink_style(self, table_name: str, col_name: str) -> None:
        """Reset a whole table column's font to plain black / no-underline (best-effort) —
        strips leftover blue+underline 'hyperlink' styling from a column that isn't (or is
        no longer) meant to be a clickable link. No-op on an empty table (dataBodyRange is
        null)."""
        import urllib.parse
        colref = f"columns('{urllib.parse.quote(col_name)}')"
        try:
            self._req("PATCH", f"{self._wb_base()}/tables/{table_name}/{colref}"
                               f"/dataBodyRange/format/font",
                      json={"underline": "None", "color": "#000000"})
        except SharePointError:
            pass

    def set_column_hidden(self, table_name: str, sheet_name: str, col_name: str,
                          hidden: bool) -> None:
        """Show or hide one column of a table (best-effort).

        IMPORTANT: a HIDDEN column breaks the Power Automate Excel connector's
        'List rows present in a table' whenever it uses $filter/$orderby — the connector
        counts visible-vs-actual columns and 400s ('Unable to match columns in the filtered
        view'). Phase 1's duplicate-check (Get_rows_ref, $filter Email eq sender) is exactly
        that call, so the raw 'Resume URL' helper is kept VISIBLE (just narrowed) rather than
        hidden. Left here as a general utility."""
        letter = self._column_letter(table_name, col_name)
        if not letter:
            return
        try:
            self._req("PATCH", f"{self._wb_base()}/worksheets('{sheet_name}')"
                               f"/range(address='{letter}:{letter}')",
                      json={"columnHidden": bool(hidden)})
        except SharePointError:
            pass

    def hide_table_column(self, table_name: str, sheet_name: str, col_name: str) -> None:
        """Back-compat shim — hide one column. Prefer set_column_hidden(...). NOTE: hiding a
        column breaks P1's filtered Get-rows call; see set_column_hidden for why."""
        self.set_column_hidden(table_name, sheet_name, col_name, True)

    def autofit_column(self, table_name: str, sheet_name: str, col_name: str) -> None:
        """Auto-fit one table column's width to its content (best-effort). Used for
        'Resume Link' so the clickable path text isn't clipped by a fixed narrow column —
        adapts on its own as filenames/paths vary in length instead of a hand-tuned width."""
        letter = self._column_letter(table_name, col_name)
        if not letter:
            return
        try:
            self._req("POST", f"{self._wb_base()}/worksheets('{sheet_name}')"
                              f"/range(address='{letter}:{letter}')/format/autofitColumns()")
        except SharePointError:
            pass

    def set_column_width(self, table_name: str, sheet_name: str, col_name: str, width: float) -> None:
        """Set a table column's width explicitly in points."""
        letter = self._column_letter(table_name, col_name)
        if not letter:
            return
        try:
            self._req("PATCH", f"{self._wb_base()}/worksheets('{sheet_name}')"
                               f"/range(address='{letter}:{letter}')", json={"columnWidth": width})
        except SharePointError:
            pass

    def clear_cell_hyperlink_style(self, table_name: str, sheet_name: str, col_name: str,
                                    row_index: int) -> None:
        """Reset one data cell's font to plain black / no-underline (best-effort).

        Fixes a cell that Excel's AutoFormat-As-You-Type auto-linked (blue + underline) when
        someone typed straight into the grid — e.g. a plain 'Email' cell that isn't supposed
        to be a hyperlink at all, unlike other rows written by the API which never trigger
        that autocorrect. `row_index` is the same 0-based table-body index list_rows()/
        list_rejected_rows() return."""
        letter = self._column_letter(table_name, col_name)
        if not letter:
            return
        try:
            addr = self._req("GET", f"{self._wb_base()}/tables/{table_name}"
                                     f"/rows/itemAt(index={row_index})/range?$select=address"
                             ).json()["address"]
            row_num = re.search(r"![A-Z]+(\d+)", addr).group(1)
            self._req("PATCH", f"{self._wb_base()}/worksheets('{sheet_name}')"
                               f"/range(address='{letter}{row_num}')/format/font",
                      json={"underline": "None", "color": "#000000"})
        except SharePointError:
            pass

    def set_column_number_format(self, table_name: str, col_name: str, code: str) -> None:
        """Set one Excel number-format code (e.g. '@' for Text) down a whole table column.

        Used to force the 'Phone' column to Text so Excel stops auto-detecting a bare
        digit-run (a foreign number left un-reformatted) as a NUMBER and rendering it in
        scientific notation / right-aligned. Mirrors set_calculated_column's dataBodyRange
        pattern. No-op on an empty table."""
        import urllib.parse
        colref = f"columns('{urllib.parse.quote(col_name)}')"
        try:
            rng = self._req("GET", f"{self._wb_base()}/tables/{table_name}/{colref}"
                                   f"/dataBodyRange?$select=rowCount").json()
        except SharePointError:
            return
        n = rng.get("rowCount", 0)
        if not n:
            return
        self._req("PATCH", f"{self._wb_base()}/tables/{table_name}/{colref}/dataBodyRange",
                  json={"numberFormat": [[code] for _ in range(n)]})

    def set_calculated_column(self, table_name: str, col_name: str, formula: str) -> None:
        """Set one uniform formula down a whole table column (an Excel 'calculated column').

        Excel auto-fills a formula entered in a table column across every row, so a per-row
        DIFFERENT formula can't survive — but a single formula that uses structured references
        (e.g. =HYPERLINK([@[Resume URL]],[@[Original Filename]])) evaluates per row and is exactly
        what we want. Writes the same formula to the whole data-body range in one PATCH. No-op
        on an empty table (dataBodyRange is null)."""
        import urllib.parse
        colref = f"columns('{urllib.parse.quote(col_name)}')"
        import time
        range_url = (f"{self._wb_base()}/tables/{table_name}/{colref}"
                     "/dataBodyRange")
        # Graph can briefly return the pre-delete rowCount after a table row is moved;
        # the following PATCH then fails because its matrix is one row too tall.  Re-read
        # dimensions before each retry so formula repair remains safe after row shifts.
        for attempt in range(3):
            try:
                rng = self._req("GET", f"{range_url}?$select=rowCount").json()
            except SharePointError:
                return
            n = rng.get("rowCount", 0)
            if not n:
                return
            try:
                self._req("PATCH", range_url,
                          json={"formulas": [[formula] for _ in range(n)]})
                return
            except SharePointError:
                if attempt == 2:
                    raise
                time.sleep(2)

    def _preserve_formulas(self, table_name: str, cols: list, index: int,
                           values_row: list, fields: dict) -> None:
        """Keep a formula cell (currently only 'Resume Link') intact through a full-row
        rewrite: when we're NOT explicitly setting it, swap the flattened display text for
        the cell's actual formula so a =HYPERLINK stays clickable."""
        for fcol in ("Resume Link",):
            if fcol in cols and fcol not in fields:
                f = self._column_formula_at(table_name, fcol, index)
                if f.startswith("="):
                    values_row[cols.index(fcol)] = f

    def update_row(self, index: int, fields: dict, current_values: dict | None = None) -> None:
        """PATCH a single table row in place. `fields` is {column_name: value};
        unspecified columns keep their current value. Pass `current_values` (the row
        dict you already read) to skip an extra Graph round-trip - faster and gentler
        on throttling."""
        cols = self.table_columns()
        if current_values is None:
            current = next((r for r in self.list_rows() if r["index"] == index), None)
            current_values = current["values"] if current else {}
        merged = dict(current_values)
        merged.update(fields)
        values = [[merged.get(c, "") for c in cols]]
        self._preserve_formulas(self.table, cols, index, values[0], fields)
        url = f"{self._wb_base()}/tables/{self.table}/rows/itemAt(index={index})"
        self._req("PATCH", url, json={"values": values})

    # ---- files -------------------------------------------------------------
    def download_file(self, folder: str, name: str) -> bytes:
        """Download a file's bytes from a drive folder (e.g. a saved resume)."""
        path = f"{folder.rstrip('/')}/{name}".lstrip("/")
        url = f"{GRAPH}/sites/{self.site_id()}/drive/root:/{path}:/content"
        return self._req("GET", url).content

    def download_resume(self, name: str, subfolder: str = "") -> bytes:
        """Download a resume from the configured resumes folder.

        `subfolder` (e.g. '2026/June' or '2026/June/Rejected') points at the dated bucket (and,
        for a rejected candidate, its 'Rejected' sub-bucket) the Phase 1 flow / P2 files resumes
        into. If the file isn't there (e.g. saved before dated folders were enabled), fall back
        to the flat resumes folder so older resumes still score.
        """
        subfolder = (subfolder or "").strip("/")
        if subfolder:
            try:
                return self.download_file(f"{self.resumes_folder}/{subfolder}", name)
            except SharePointError:
                pass  # fall back to the flat folder (pre-dated resumes)
        return self.download_file(self.resumes_folder, name)

    def file_web_url(self, folder: str, name: str) -> str:
        """Return the SharePoint web URL that opens a file in the browser (for a clickable link)."""
        path = f"{folder.rstrip('/')}/{name}".lstrip("/")
        url = f"{GRAPH}/sites/{self.site_id()}/drive/root:/{path}:/?$select=webUrl"
        return self._req("GET", url).json().get("webUrl", "")

    def resume_web_url(self, name: str, subfolder: str = "") -> str:
        """Web URL of a saved resume — mirrors download_resume's dated→flat fallback."""
        subfolder = (subfolder or "").strip("/")
        if subfolder:
            try:
                return self.file_web_url(f"{self.resumes_folder}/{subfolder}", name)
            except SharePointError:
                pass  # fall back to the flat folder (pre-dated resumes)
        return self.file_web_url(self.resumes_folder, name)

    def upload_file(self, folder: str, name: str, data: bytes) -> None:
        """Upload (overwrite) a small file's bytes to a drive folder (folders auto-created)."""
        path = f"{folder.rstrip('/')}/{name}".lstrip("/")
        url = f"{GRAPH}/sites/{self.site_id()}/drive/root:/{path}:/content"
        self._req("PUT", url, data=data, headers={"Content-Type": "application/octet-stream"})

    def delete_file(self, folder: str, name: str) -> None:
        """Delete a file from a drive folder (best-effort compensating cleanup)."""
        path = f"{folder.rstrip('/')}/{name}".lstrip("/")
        url = f"{GRAPH}/sites/{self.site_id()}/drive/root:/{path}"
        self._req("DELETE", url)

    def delete_folder(self, path: str) -> None:
        """Delete a drive folder AND everything inside it (Graph deletes recursively).
        Used by the replay-test wipe to remove dated <Year>/ resume trees. Same DELETE
        primitive as delete_file - Graph treats files and folders as driveItems alike."""
        url = f"{GRAPH}/sites/{self.site_id()}/drive/root:/{path.strip('/')}"
        self._req("DELETE", url)

    def list_folder_children(self, path: str) -> list:
        """Names + folder-ness of a drive folder's immediate children (for wipe/verify
        enumeration). Returns [{'name': ..., 'is_folder': bool}]."""
        clean = path.strip("/")
        if clean:
            url = (f"{GRAPH}/sites/{self.site_id()}/drive/root:/{clean}:"
                   f"/children?$select=name,folder&$top=200")
        else:
            url = f"{GRAPH}/sites/{self.site_id()}/drive/root/children?$select=name,folder&$top=200"
        items = []
        while url:
            data = self._req("GET", url).json()
            items.extend({"name": it.get("name", ""), "is_folder": "folder" in it}
                         for it in data.get("value", []))
            url = data.get("@odata.nextLink")
        return items

    def move_resume(self, name: str, from_folder: str, to_folder: str) -> bool:
        """Move one resume file between two fully-resolved folders (e.g. .../2026/July <->
        .../2026/July/Rejected). Download+upload+delete (reuses the existing primitives — Graph
        has no path-based "move" this client already wraps). Returns False (no-op) if the file
        isn't at the source: either it was already moved by an earlier run, or it never existed
        there — either way the caller should treat this as already-correct rather than an error,
        which is what makes every call site idempotent."""
        try:
            data = self.download_file(from_folder, name)
        except SharePointError:
            return False
        self.upload_file(to_folder, name, data)
        try:
            self.delete_file(from_folder, name)
        except SharePointError:
            pass  # copy already succeeded; a leftover source copy is harmless
        return True

    def rename_resume(self, old_name: str, new_name: str, subfolder: str = "") -> bool:
        """Rename a resume file in SharePoint. Returns True if renamed, False if not found/error."""
        subfolder = (subfolder or "").strip("/")
        folder = f"{self.resumes_folder}/{subfolder}" if subfolder else self.resumes_folder
        try:
            # Try to rename in the folder
            path = f"{folder.rstrip('/')}/{old_name}".lstrip("/")
            url = f"{GRAPH}/sites/{self.site_id()}/drive/root:/{path}"
            self._req("PATCH", url, json={"name": new_name})
            return True
        except SharePointError:
            # Fallback to root folder if subfolder attempt failed
            if subfolder:
                try:
                    path = f"{self.resumes_folder.rstrip('/')}/{old_name}".lstrip("/")
                    url = f"{GRAPH}/sites/{self.site_id()}/drive/root:/{path}"
                    self._req("PATCH", url, json={"name": new_name})
                    return True
                except SharePointError:
                    pass
        return False

    # ---- outbound mail (app-only; needs Mail.Send Application permission) ----
    def send_mail(self, to_address: str, subject: str, html_body: str) -> bool:
        """Send an HTML email FROM self.sender_mailbox via Graph (app-only sendMail).

        Returns True on success, False on any failure (never raises) - a failed notice must
        not break a scoring run. Needs Mail.Send (Application) admin-consented on the app; until
        then Graph returns 403 and this logs a warning and returns False.

        This is the ONLY place P2 transmits mail, which makes it the enforcement point for
        config.SUPPRESS_EMAILS (HIRING_SUPPRESS_EMAILS / test_mode.suppress_emails). When
        suppressed it returns True WITHOUT contacting Graph, so the caller proceeds exactly as
        if the send had succeeded - the row still gets its marker and the pass still advances,
        the same way P1's suppressed build keeps every Send_* action's runAfter wiring intact.
        """
        to_address = (to_address or "").strip()
        if not to_address or "@" not in to_address:
            return False
        from hiring_agent.config import SUPPRESS_EMAILS, logger
        if SUPPRESS_EMAILS:
            logger.info(f"       Email     : SUPPRESSED (test mode) - would have emailed "
                        f"{to_address} | {subject}")
            return True
        url = f"{GRAPH}/users/{self.sender_mailbox}/sendMail"
        payload = {
            "message": {
                "subject": subject,
                "body": {"contentType": "HTML", "content": html_body},
                "toRecipients": [{"emailAddress": {"address": to_address}}],
            },
            "saveToSentItems": True,
        }
        try:
            self._req("POST", url, json=payload)
            return True
        except SharePointError as e:
            from hiring_agent.config import logger
            logger.warning(f"       WARNING  : decline email to {to_address} not sent "
                           f"(check Mail.Send app permission): {e}")
            return False

    # ---- rejected sheet ----------------------------------------------------
    def _ensure_rejected_table(self, table_name: str = "RejectedCandidates") -> str:
        """Ensure a 'Rejected' worksheet + table exists in the workbook; return the table name.

        A freshly created Rejected table gets the full REJECTED_COLUMNS schema (every content
        header, plus 'Decline Sent' in place of the main-only 'Mail Sent'), so both sheets carry
        the complete set of headers from the start."""
        url = f"{self._wb_base()}/worksheets"
        sheets = self._req("GET", url).json().get("value", [])
        sheet_names = [s["name"] for s in sheets]
        if "Rejected" not in sheet_names:
            self._req("POST", url, json={"name": "Rejected"})

        tbl_url = f"{self._wb_base()}/worksheets/Rejected/tables"
        try:
            tables = self._req("GET", tbl_url).json().get("value", [])
            if any(t["name"] == table_name for t in tables):
                return table_name
            if tables:
                # A table already exists on the sheet (e.g. an earlier rename failed
                # and it kept its auto-name) — reuse it rather than overlap a new one.
                return tables[0]["name"]
        except SharePointError:
            pass

        from hiring_agent.config import REJECTED_COLUMNS as COLUMNS
        from openpyxl.utils import get_column_letter
        ref = f"A1:{get_column_letter(len(COLUMNS))}1"
        range_url = f"{self._wb_base()}/worksheets/Rejected/range(address='{ref}')"
        self._req("PATCH", range_url, json={"values": [COLUMNS]})
        add_url = f"{self._wb_base()}/worksheets/Rejected/tables/add"
        created = self._req("POST", add_url,
                            json={"address": ref, "hasHeaders": True}).json()
        auto_name = created.get("name") or created.get("id") or ""
        try:
            self._req("PATCH", f"{self._wb_base()}/tables/{auto_name}",
                      json={"name": table_name})
            return table_name
        except SharePointError:
            # Rename failed — the auto-name still works for every rows/add call.
            return auto_name or table_name

    def add_rejected_row(self, fields: dict) -> None:
        """Append a row to the 'Rejected' sheet/table in the same workbook."""
        tbl = self._ensure_rejected_table()
        # Map by the table's OWN columns (not config COLUMNS) — the Rejected table can
        # carry extra columns like 'Decline Sent', and Graph requires the value array
        # length to match the real column count exactly.
        cols = self._table_columns_of(tbl)
        values = [[fields.get(c, "") for c in cols]]
        url = f"{self._wb_base()}/tables/{tbl}/rows/add"
        self._req("POST", url, json={"values": values})

    def ensure_rejected_columns(self, required: list) -> list:
        """Add any missing columns to the Rejected table (if it exists yet), inserted at
        its canonical position (from config.REJECTED_COLUMNS) - same reasoning as
        ensure_columns.

        Never creates the Rejected sheet — a workbook with no rejections doesn't
        need one. Returns the columns that were added."""
        tbl = self._rejected_table_name_if_exists()
        if tbl is None:
            return []
        from hiring_agent.config import REJECTED_COLUMNS as _CANON
        existing = list(self._table_columns_of(tbl))
        added = []
        for col in required:
            if col not in existing:
                payload = {"name": col}
                if col in _CANON:
                    before = _CANON[:_CANON.index(col)]
                    payload["index"] = sum(1 for c in before if c in existing)
                self._req("POST", f"{self._wb_base()}/tables/{tbl}/columns/add",
                          json=payload)
                existing.append(col)
                added.append(col)
        if added:
            from hiring_agent.config import logger
            logger.info(f"   added missing column(s) to Rejected table: {', '.join(added)}")
        return added

    def delete_row(self, index: int) -> None:
        """Delete a row from the main table by index."""
        url = f"{self._wb_base()}/tables/{self.table}/rows/itemAt(index={index})"
        self._req("DELETE", url)

    def add_main_row(self, fields: dict) -> None:
        """Append a row to the MAIN candidate table. Used by --recheck-all to restore a
        candidate that was wrongly moved to Rejected (e.g. geo re-check now says USA)."""
        cols = self.table_columns()
        values = [[fields.get(c, "") for c in cols]]
        url = f"{self._wb_base()}/tables/{self.table}/rows/add"
        self._req("POST", url, json={"values": values})

    # ---- reconcile helpers (used by --recheck-all) -------------------------
    def _table_columns_of(self, table_name: str) -> list:
        """Column header names (in order) for an arbitrary table in this workbook."""
        if table_name in self._columns_cache:
            return self._columns_cache[table_name]
        url = f"{self._wb_base()}/tables/{table_name}/columns?$select=name"
        cols = [c["name"] for c in self._req("GET", url).json().get("value", [])]
        self._columns_cache[table_name] = cols
        return cols

    def _rejected_table_name_if_exists(self, table_name: str = "RejectedCandidates"):
        """Return the Rejected table's name, or None if the sheet/table doesn't exist yet.
        Unlike _ensure_rejected_table this never creates anything (safe for read/dry-run)."""
        try:
            sheets = [s["name"] for s in
                      self._req("GET", f"{self._wb_base()}/worksheets").json().get("value", [])]
            if "Rejected" not in sheets:
                return None
            tables = self._req("GET", f"{self._wb_base()}/worksheets/Rejected/tables"
                               ).json().get("value", [])
            if not tables:
                return None
            for t in tables:
                if t["name"] == table_name:
                    return table_name
            return tables[0]["name"]
        except SharePointError:
            return None

    def list_rejected_rows(self) -> list:
        """All Rejected-sheet rows as {'index', 'values'}; [] if there is no Rejected table."""
        tbl = self._rejected_table_name_if_exists()
        if tbl is None:
            return []
        try:
            cols = self._table_columns_of(tbl)
            rows = self._rows_paged(f"{self._wb_base()}/tables/{tbl}/rows")
        except SharePointError:
            return []
        out = []
        for r in rows:
            vals = r.get("values", [[]])
            cells = vals[0] if vals else []
            mapped = {cols[i]: (cells[i] if i < len(cells) else "") for i in range(len(cols))}
            out.append({"index": r.get("index"), "values": mapped})
        return out

    def update_rejected_row(self, index: int, fields: dict, current_values: dict | None = None) -> None:
        """PATCH a single Rejected-sheet row in place (heal blanks without re-adding)."""
        tbl = self._rejected_table_name_if_exists()
        if tbl is None:
            return
        cols = self._table_columns_of(tbl)
        if current_values is None:
            current = next((r for r in self.list_rejected_rows() if r["index"] == index), None)
            current_values = current["values"] if current else {}
        merged = dict(current_values)
        merged.update(fields)
        values = [[merged.get(c, "") for c in cols]]
        self._preserve_formulas(tbl, cols, index, values[0], fields)
        url = f"{self._wb_base()}/tables/{tbl}/rows/itemAt(index={index})"
        self._req("PATCH", url, json={"values": values})

    def delete_rejected_row(self, index: int) -> None:
        """Delete a row from the Rejected table by index."""
        tbl = self._rejected_table_name_if_exists()
        if tbl is None:
            return
        url = f"{self._wb_base()}/tables/{tbl}/rows/itemAt(index={index})"
        self._req("DELETE", url)
