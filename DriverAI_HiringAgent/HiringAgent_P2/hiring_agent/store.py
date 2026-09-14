"""Row identity for the candidate workbook — the one module allowed to address a row.

WHY THIS EXISTS
---------------
Phase 1 appends to the same Excel table every 60 seconds. Phase 2 used to read a batch of
rows once and then write each result back by POSITION (`itemAt(index=N)`), keeping a
`deleted_idx` list and subtracting an offset to compensate for rows it removed mid-run. That
arithmetic is correct only while EVERY delete path remembers to register itself, and on
2026-08-03 one did not: `_mark_needs_review`'s two delete paths returned None, so every row
after a given-up candidate wrote one index too high — silently pasting each candidate's
results onto the NEXT candidate's row, then failing off the end of the table.

The fix that shipped at the time (`_live_row_index`) verified the arithmetic with a single-row
GET before every write. It worked, but it left positional addressing as the primary mechanism
with verification bolted on top, which means the next forgotten delete path re-opens the same
hole.

This module inverts that. `Application ID` is the primary key, the way Phase 1 has always done
it — every one of P1's seven `PatchItem` actions addresses its row with
`idColumn: "Application ID"`, and both of its reads use a server-side
`$filter` on the same column. P1 has no positional addressing anywhere, and has never suffered
this class of bug. Phase 2 now adopts P1's convention against the same table.

THE CONTRACT
------------
Nothing above this module may hold a row index. Callers pass an Application ID; the store
resolves it to a live position immediately before each write and never caches that position
across an operation. `CandidateRow.version` is deliberately opaque: today it wraps the row
index, and after a migration to a SharePoint List it becomes the ETag, without a single caller
changing.

THE KEY IS NOT UNIQUE
---------------------
Planning this module assumed Application ID was a primary key. It is not, and `test_p2.py`
caught it: re-processing can leave a stale `Needs Review` retry row beside the completed row
for the same candidate, both carrying the same ID. Giving up on the retry row must delete the
retry row, not the Scored one that happens to match first.

So a caller's last-known position is not merely an optimisation — it is the tie-break that
says WHICH matching row was meant. It is still never trusted on its own: the store confirms
the row at that position really carries the key before honouring it, and re-locates by key
when it does not. What was removed is unverified positional arithmetic, not the position.

FAIL CLOSED
-----------
If a row cannot be located by its Application ID, the store raises `RowVanished` rather than
guessing. Callers skip that candidate and leave its Status untouched, so the next run picks it
up. A skipped candidate costs one cycle; a mis-addressed write corrupts two rows.
"""
from __future__ import annotations

from dataclasses import dataclass, replace
from typing import Iterable, Literal, Protocol

from .config import logger

APP_ID_COL = "Application ID"

Sheet = Literal["main", "rejected"]


# Every other module in this package imports sharepoint_client INSIDE a function, so that
# hiring_agent stays importable for local/offline scoring where the Graph client may not be on
# the path at all. A base class cannot be resolved lazily, so this is the one place that has to
# try at import time - and it degrades to RuntimeError when the client is absent, which is
# exactly the offline case that convention protects.
try:                                    # pragma: no cover - depends on deployment shape
    from sharepoint_client import SharePointError as _RowError
except Exception:                       # pragma: no cover
    _RowError = RuntimeError


class RowVanished(_RowError):
    """A row that was read earlier can no longer be found by its Application ID.

    Raised instead of writing to a guessed position. The candidate keeps its current Status
    and is picked up by the next run.

    Deliberately a SharePointError subclass: every caller in sharepoint_scoring.py already
    wraps its writes in `except SharePointError`, so a vanished row degrades exactly like a
    transient Graph failure does today - log it, skip that candidate, carry on with the batch.
    Failing closed needed no new handling anywhere; it inherits the handling that exists.
    """


@dataclass(frozen=True)
class CandidateRow:
    """One candidate row, addressed by key rather than position.

    ``version`` is OPAQUE above this module. Today it carries the row index as a string;
    after a move to a SharePoint List it carries the ETag. Callers hand back what they were
    given and never parse it.
    """

    app_id: str
    values: dict
    version: str
    sheet: Sheet = "main"

    @property
    def is_keyless(self) -> bool:
        """True for a structural row with no Application ID.

        The live sheets carry historical blank rows — the Rejected sheet had drifted to
        roughly 14 orphaned blanks before P2's separator writes were disabled on 2026-09-01.
        They are not candidates and cannot be re-located by key, so they are never returned
        from ``get_queue``/``get`` and are never deleted as "corrupt".
        """
        return not str(self.app_id or "").strip()


class CandidateStore(Protocol):
    """The only interface the pipeline may use to reach candidate rows."""

    def get_queue(self, statuses: tuple, limit: int | None = None) -> list: ...
    def get(self, app_id: str, sheet: Sheet = "main") -> CandidateRow | None: ...
    def all_rows(self, sheet: Sheet = "main") -> list: ...
    def save_by_id(self, app_id: str, fields: dict, *,
                   current_values: dict | None = None, sheet: Sheet = "main") -> None: ...
    def add(self, fields: dict, sheet: Sheet = "main") -> None: ...
    def delete_by_id(self, app_id: str, sheet: Sheet = "main") -> bool: ...
    def move_to_rejected(self, app_id: str, fields: dict) -> bool: ...


def _key_of(values: dict) -> str:
    return str((values or {}).get(APP_ID_COL, "") or "").strip()


class ExcelCandidateStore:
    """`CandidateStore` over today's Excel workbook, via the existing `SharePointClient`.

    Deliberately a thin wrapper: it changes HOW a row is addressed, not what is written. Every
    formula-preservation and row-format side effect stays in `SharePointClient`, unchanged.
    """

    def __init__(self, client):
        self.client = client

    # ---- reads -------------------------------------------------------------

    def _raw_rows(self, sheet: Sheet) -> list:
        if sheet == "rejected":
            return self.client.list_rejected_rows()
        return self.client.list_rows()

    def all_rows(self, sheet: Sheet = "main") -> list:
        """Every row on the sheet, keyless structural rows included."""
        return [CandidateRow(app_id=_key_of(r["values"]), values=r["values"],
                             version=str(r["index"]), sheet=sheet)
                for r in self._raw_rows(sheet)]

    def get(self, app_id: str, sheet: Sheet = "main") -> CandidateRow | None:
        """The row with this Application ID, or None. Never returns a keyless row."""
        want = str(app_id or "").strip()
        if not want:
            return None
        for r in self._raw_rows(sheet):
            if _key_of(r["values"]) == want:
                return CandidateRow(app_id=want, values=r["values"],
                                    version=str(r["index"]), sheet=sheet)
        return None

    def get_queue(self, statuses: tuple, limit: int | None = None) -> list:
        """Main-sheet rows whose Status is in `statuses`, oldest position first.

        Keyless rows are skipped: a blank spacer has no Status and is not a candidate.
        """
        wanted = {str(s).strip() for s in
                  ((statuses,) if isinstance(statuses, str) else statuses)}
        out = []
        for r in self.client.list_rows():
            key = _key_of(r["values"])
            if not key:
                continue
            if str(r["values"].get("Status", "") or "").strip() in wanted:
                out.append(CandidateRow(app_id=key, values=r["values"],
                                        version=str(r["index"]), sheet="main"))
                if limit is not None and len(out) >= limit:
                    break
        return out

    # ---- identity ----------------------------------------------------------

    def _can_verify(self, sheet: Sheet) -> bool:
        """Whether this client can confirm what sits at a position.

        Only the main table has a targeted single-row read. A client that lacks it cannot be
        treated as authoritative about a row being ABSENT — several test stubs implement a
        token `list_rows` that does not model the rows being written at all.
        """
        return sheet == "main" and hasattr(self.client, "row_values_at")

    def _verify_at(self, index: int, app_id: str, sheet: Sheet):
        """Cheap single-row check that `index` still holds `app_id`. None = cannot tell.

        Only the main table has a targeted single-row read. A stripped-down client in a test
        may not implement it at all, in which case the caller falls back to a full scan — the
        same `hasattr` degradation `_live_row_index` performed, preserved deliberately so
        existing fakes keep working.
        """
        if sheet != "main" or not hasattr(self.client, "row_values_at"):
            return None
        try:
            vals = self.client.row_values_at(int(index))
        except Exception:
            return None
        if vals is None:
            return False
        return _key_of(vals) == app_id

    def resolve_position(self, app_id: str, sheet: Sheet = "main",
                         hint: str | int | None = None) -> int:
        """Public form of `_resolve`, for callers that still need a raw position.

        The only caller is `sharepoint_scoring._live_row_index`, which exists so the batch
        loop can log and record a position. New code should use `save_by_id`/`delete_by_id`
        and never see an index at all.
        """
        return self._resolve(app_id, sheet, hint)

    def _resolve(self, app_id: str, sheet: Sheet, hint: str | int | None = None) -> int:
        """The CURRENT index of `app_id`, verified. Raises `RowVanished` if it is gone.

        THE KEY IS NOT UNIQUE. Application IDs are usually one-to-one with rows, but not
        always: re-processing can leave a stale `Needs Review` retry row beside the completed
        row for the same candidate, both carrying the same Application ID. `test_p2.py` pins
        exactly that case — giving up on the retry row at index 9 must delete index 9, not the
        Scored row at index 3 that happens to match the same key and comes first.

        So the hint is not merely an optimisation, it is the tie-breaker:

          * hint verifies clean          -> use it (one cheap GET, the common path)
          * hint is among the matches    -> use it; duplicates exist and the caller meant this row
          * exactly one match            -> use it; the row simply moved
          * no match                     -> RowVanished; never guess
          * several matches, no hint     -> first, with a warning; the caller lost the tie-break

        What is dropped is *unverified positional arithmetic* — an index carried across a batch
        and adjusted by a running offset. A hint is only ever honoured after it has been
        confirmed to hold this candidate.
        """
        key = str(app_id or "").strip()
        if not key:
            raise RowVanished("cannot address a row with no Application ID")

        hinted = None
        if hint is not None:
            try:
                hinted = int(hint)
            except (TypeError, ValueError):
                hinted = None
            if hinted is not None and self._verify_at(hinted, key, sheet) is True:
                return hinted

        # A stripped-down client cannot be scanned. Several test fakes implement only the
        # write half, and `_live_row_index` degraded the same way when `row_values_at` was
        # missing: trust the caller's hint rather than refusing to work. Real runs always
        # carry a full client, so this path is a compatibility shim, not the normal route.
        reader = "list_rejected_rows" if sheet == "rejected" else "list_rows"
        if not hasattr(self.client, reader):
            if hinted is not None:
                return hinted
            raise RowVanished(f"cannot locate {key}: client cannot list the {sheet} sheet")

        matches = [int(r["index"]) for r in self._raw_rows(sheet)
                   if _key_of(r["values"]) == key]
        if not matches:
            # Compatibility shim, MAIN SHEET ONLY. The older stubs answer `list_rows` with a
            # token table that never contained the row being addressed, and the pre-store
            # `_live_row_index` fell back to their arithmetic; keep doing that for them.
            #
            # It must NOT extend to the rejected sheet. `_can_verify` is false there for a
            # different reason - no single-row read endpoint exists for it at all - so
            # phrasing this as `not self._can_verify(sheet)` silently disabled FAIL CLOSED
            # for every rejected-sheet write, letting an unlocatable key land on a stale
            # position. `list_rejected_rows()` IS authoritative about absence.
            if (hinted is not None and sheet == "main"
                    and not hasattr(self.client, "row_values_at")):
                return hinted
            raise RowVanished(f"{key} is no longer on the {sheet} sheet")
        if hinted is not None and hinted in matches:
            return hinted
        if len(matches) > 1:
            logger.warning(
                f"       WARNING  : {key} matches {len(matches)} rows on the {sheet} sheet "
                f"{matches} and no usable position was supplied; addressing the first."
            )
        return matches[0]

    # ---- writes ------------------------------------------------------------

    P1_MASTER_WRITE_COLUMNS = frozenset({"Status", "Resume Link"})

    def _is_p1_intake(self) -> bool:
        return getattr(self.client, "table", None) == "HiringAgent_P1_Candidates"

    def save_by_id(self, app_id: str, fields: dict, *,
                   current_values: dict | None = None, sheet: Sheet = "main",
                   hint: str | int | None = None) -> None:
        """PATCH one row, located by Application ID. Unlisted columns keep their value."""
        if self._is_p1_intake() and sheet == "rejected":
            raise RowVanished("P1's Rejected sheet is read-only to P2")
        index = self._resolve(app_id, sheet, hint)
        if sheet == "rejected":
            self.client.update_rejected_row(index, fields, current_values=current_values)
        else:
            if self._is_p1_intake():
                fields = {k: v for k, v in fields.items()
                          if k in self.P1_MASTER_WRITE_COLUMNS}
            self.client.update_row(index, fields, current_values=current_values)

    def add(self, fields: dict, sheet: Sheet = "main") -> None:
        """Append a row. Appends never race: position is assigned by the server."""
        if sheet == "rejected":
            self.client.add_rejected_row(fields)
        else:
            self.client.add_main_row(fields)

    def delete_by_id(self, app_id: str, sheet: Sheet = "main",
                     hint: str | int | None = None) -> bool:
        """Delete the row with this Application ID. False if it was already gone.

        Excel table rows frequently refuse to delete — Graph answers 409
        `InsertDeleteConflict` — which is why a rejected candidate can leave a ghost row
        behind on the source sheet. That failure is surfaced, never swallowed: this is the
        single place it happens, so a move to a store that CAN delete fixes every caller at
        once.
        """
        try:
            index = self._resolve(app_id, sheet, hint)
        except RowVanished:
            return False
        if sheet == "rejected":
            self.client.delete_rejected_row(index)
        else:
            self.client.delete_row(index)
        return True

    def delete_at(self, index: int, sheet: Sheet = "main") -> None:
        """Positional delete — ONLY valid inside a descending-index batch after all reads.

        The escape hatch for the two end-of-pass cleanup loops that delete merged-away and
        orphaned rows. Deleting from the bottom up can never move a row still queued for
        deletion, so those loops are order-independent and safe as they stand. Do not reach
        for this anywhere else: interleaving a positional delete with per-candidate writes is
        the original bug.
        """
        if sheet == "rejected":
            self.client.delete_rejected_row(int(index))
        else:
            self.client.delete_row(int(index))

    def move_to_rejected(self, app_id: str, fields: dict,
                         hint: str | int | None = None) -> bool:
        """Add to Rejected, then remove from main. Returns False if the delete failed.

        ONE operation on purpose. Today it is still two Graph calls and is NOT atomic: the add
        can succeed and the delete 409, leaving the candidate on both sheets. That is exactly
        the ghost-row symptom, and collapsing it here means the eventual fix lands in one
        place instead of the four call sites that used to open-code the pair.
        """
        if self._is_p1_intake():
            # The rejection belongs to P2's store/master. P1 keeps its intake row and sees
            # only the terminal Status and renamed Resume Link.
            self.save_by_id(app_id, fields, sheet="main", hint=hint)
            return True

        self.client.add_rejected_row(fields)
        try:
            removed = self.delete_by_id(app_id, "main", hint)
        except Exception as e:
            logger.warning(
                f"       WARNING  : {app_id} added to Rejected but could not be removed from "
                f"the main sheet ({e}); it is now on BOTH sheets."
            )
            return False
        if not removed:
            logger.warning(
                f"       WARNING  : {app_id} added to Rejected but its main-sheet row could "
                f"not be located to remove; it may now be on BOTH sheets."
            )
        return removed


from .sqlite_store import SQLiteCandidateStore


_GLOBAL_SQLITE_STORE = None


#: Recognised values for HIRING_STORAGE_BACKEND. Unset still means 'excel': the Azure
#: Functions deploy and most of the suite rely on that default, so it stays. What is NOT
#: tolerated any more is an UNRECOGNISED value. Until 2026-09-09 anything that was not
#: exactly 'sqlite' fell through to ExcelCandidateStore, so a single typo in .env
#: ('sqllite', 'sqlite3', 'Sqlite ') silently pointed every candidate write at P1's live
#: workbook while the operator believed the local database was in use - a misconfiguration
#: that reads as a data-loss incident, not as a config error. It now stops the run.
_VALID_BACKENDS = ("excel", "sqlite")


def storage_backend(backend: str | None = None) -> str:
    """Resolve the storage backend name, refusing anything unrecognised."""
    import os
    chosen = (backend or os.getenv("HIRING_STORAGE_BACKEND") or "excel").lower().strip()
    if chosen not in _VALID_BACKENDS:
        raise ValueError(
            f"HIRING_STORAGE_BACKEND={chosen!r} is not a recognised storage backend "
            f"(expected one of: {', '.join(_VALID_BACKENDS)}). Refusing to fall back to "
            f"'excel', which would send every write straight to P1's live workbook."
        )
    return chosen


def get_store(client=None, backend: str | None = None) -> CandidateStore:
    """Store factory supporting both SQLite and Excel backends."""
    import os
    global _GLOBAL_SQLITE_STORE
    if client is not None and getattr(client, "local_execution", False):
        return client.candidate_store
    # Mock/stub clients in test suites don't have real SharePoint attributes
    if client is not None and not hasattr(client, "site_id"):
        return ExcelCandidateStore(client)
    chosen = storage_backend(backend)
    if chosen == "sqlite":
        if _GLOBAL_SQLITE_STORE is None:
            _GLOBAL_SQLITE_STORE = SQLiteCandidateStore()
        return _GLOBAL_SQLITE_STORE
    return ExcelCandidateStore(client)
