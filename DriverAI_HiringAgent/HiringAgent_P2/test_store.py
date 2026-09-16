"""Tests for the 2026-09-06 additions: row identity, and extraction provenance.

Kept separate from test_p2.py on purpose: that suite's 1393-test count is the regression
baseline for the store migration, and mixing new coverage into it would destroy the
comparison. Run both.

Section A pins the mechanics. Section B reproduces the 2026-08-03 incident directly — the one
where a forgotten delete path shifted every later write by one and pasted each candidate's
results onto the NEXT candidate's row. That is the bug this module exists to make
unrepresentable, so it gets an explicit, named test rather than being implied by the others.
"""
import os
import sys
from pathlib import Path
# Legacy orchestration checks use Excel stubs; local SQLite integration has its own suite.
os.environ['HIRING_STORAGE_BACKEND'] = 'excel'
sys.path.insert(0, str(Path(__file__).resolve().parent))
import test_isolation  # noqa: F401 - must precede every hiring_agent import
from hiring_agent.store import (
    ExcelCandidateStore, CandidateRow, RowVanished, APP_ID_COL,
)

P = F = 0


def ok(c, label):
    global P, F
    if c: P += 1; print(f"  [PASS] {label}")
    else: F += 1; print(f"  [FAIL] {label}")


class FakeClient:
    """An Excel table that behaves like the real one: positional, and it shifts on delete."""

    def __init__(self, main=None, rejected=None):
        self.main = list(main or [])
        self.rejected = list(rejected or [])
        self.writes = []          # (sheet, index, fields)
        self.deletes = []         # (sheet, index)
        self.adds = []            # (sheet, fields)
        self.delete_fails = False  # simulate 409 InsertDeleteConflict

    # -- reads
    def _rows(self, sheet):
        return self.rejected if sheet == "rejected" else self.main

    def list_rows(self):
        return [{"index": i, "values": dict(v)} for i, v in enumerate(self.main)]

    def list_rejected_rows(self):
        return [{"index": i, "values": dict(v)} for i, v in enumerate(self.rejected)]

    def row_values_at(self, index):
        if 0 <= index < len(self.main):
            return dict(self.main[index])
        return None

    # -- writes
    def update_row(self, index, fields, current_values=None):
        self.writes.append(("main", index, dict(fields)))
        self.main[index] = {**self.main[index], **fields}

    def update_rejected_row(self, index, fields, current_values=None):
        self.writes.append(("rejected", index, dict(fields)))
        self.rejected[index] = {**self.rejected[index], **fields}

    def delete_row(self, index):
        if self.delete_fails:
            raise RuntimeError("409 InsertDeleteConflict")
        self.deletes.append(("main", index))
        self.main.pop(index)          # the shift that caused the incident

    def delete_rejected_row(self, index):
        if self.delete_fails:
            raise RuntimeError("409 InsertDeleteConflict")
        self.deletes.append(("rejected", index))
        self.rejected.pop(index)

    def add_main_row(self, fields):
        self.adds.append(("main", dict(fields)))
        self.main.append(dict(fields))

    def add_rejected_row(self, fields):
        self.adds.append(("rejected", dict(fields)))
        self.rejected.append(dict(fields))


class FakeClientNoSingleRead(FakeClient):
    """A stripped-down client with no targeted single-row read.

    Mirrors the fakes already used in test_p2.py, and the `hasattr` degradation the previous
    `_live_row_index` performed. The store must fall back to a full scan, not raise.
    """
    row_values_at = None

    def __getattribute__(self, name):
        if name == "row_values_at":
            raise AttributeError(name)
        return object.__getattribute__(self, name)


def row(app_id, status="New Email Received", **extra):
    return {APP_ID_COL: app_id, "Status": status, **extra}


# ── A. READS ────────────────────────────────────────────────────────────────
print("\n=== A. READS ===")

_c = FakeClient(main=[row("APP-1"), {}, row("APP-2", "Scored"), row("APP-3")])
_s = ExcelCandidateStore(_c)

_q = _s.get_queue(("New Email Received",))
ok([r.app_id for r in _q] == ["APP-1", "APP-3"],
   "get_queue returns only rows whose Status matches")
ok(all(not r.is_keyless for r in _q),
   "get_queue never returns a keyless structural row")
ok(_s.get_queue(("New Email Received",), limit=1)[0].app_id == "APP-1",
   "get_queue honours the batch limit, oldest position first")
ok(_s.get_queue("New Email Received")[0].app_id == "APP-1",
   "get_queue accepts a bare string as well as a tuple")

_all = _s.all_rows()
ok(len(_all) == 4 and sum(1 for r in _all if r.is_keyless) == 1,
   "all_rows preserves the keyless structural row rather than hiding it")

ok(_s.get("APP-2").values["Status"] == "Scored", "get finds a row by Application ID")
ok(_s.get("APP-NOPE") is None, "get returns None for an unknown Application ID")
ok(_s.get("") is None, "get returns None for a blank Application ID")
ok(_s.get_queue(("New Email Received",))[0].version == "0",
   "version carries the position as an opaque string")


# ── B. THE 2026-08-03 INCIDENT ──────────────────────────────────────────────
print("\n=== B. A DELETE MID-BATCH NEVER MIS-ADDRESSES A LATER WRITE ===")

# Read the batch first, exactly as the pipeline does, THEN delete an earlier row. Every
# cached index at or after the deleted position is now stale by one.
_c = FakeClient(main=[row("APP-A"), row("APP-B"), row("APP-C"), row("APP-D")])
_s = ExcelCandidateStore(_c)
_batch = _s.get_queue(("New Email Received",))
ok([r.version for r in _batch] == ["0", "1", "2", "3"], "batch cached positions 0..3")

_s.delete_by_id("APP-B")                      # everything after B shifts down one
ok([v[APP_ID_COL] for v in _c.main] == ["APP-A", "APP-C", "APP-D"],
   "APP-B is gone and the rows below it have shifted")

# APP-D's cached index is 3; it now really lives at 2. A positional write would land on
# nothing (or, with one more row, on the wrong candidate).
_stale = next(r for r in _batch if r.app_id == "APP-D")
_s.save_by_id(_stale.app_id, {"Status": "Scored"}, hint=_stale.version)
ok(_c.main[2][APP_ID_COL] == "APP-D" and _c.main[2]["Status"] == "Scored",
   "a stale cached index is re-resolved and the write lands on the RIGHT candidate")
ok(_c.main[0]["Status"] == "New Email Received" and _c.main[1]["Status"] == "New Email Received",
   "no other candidate's row was touched")

# The same guarantee without any hint at all.
_s.save_by_id("APP-A", {"Status": "Scored"})
ok(_c.main[0]["Status"] == "Scored", "save_by_id works with no positional hint at all")


# ── C. RESOLUTION AND FAILING CLOSED ────────────────────────────────────────
print("\n=== C. RESOLUTION AND FAILING CLOSED ===")

_c = FakeClient(main=[row("APP-1"), row("APP-2")])
_s = ExcelCandidateStore(_c)

try:
    _s.save_by_id("APP-GONE", {"Status": "Scored"})
    _raised = False
except RowVanished:
    _raised = True
ok(_raised, "writing to a vanished Application ID raises RowVanished, never a guessed row")
ok(_c.writes == [], "a vanished row produces no write at all")

try:
    _s.save_by_id("", {"Status": "Scored"})
    _blank = False
except RowVanished:
    _blank = True
ok(_blank, "a blank Application ID can never address a row")

ok(_s.delete_by_id("APP-GONE") is False,
   "delete_by_id reports False for an already-absent row instead of raising")

# A wrong hint must be ignored, not trusted.
_s.save_by_id("APP-1", {"Status": "Scored"}, hint="1")
ok(_c.main[0]["Status"] == "Scored" and _c.main[1]["Status"] == "New Email Received",
   "a hint pointing at the WRONG candidate is discarded and the key wins")

# A client with no single-row read must still resolve, by full scan.
_c2 = FakeClientNoSingleRead(main=[row("APP-1"), row("APP-2")])
_s2 = ExcelCandidateStore(_c2)
_s2.save_by_id("APP-2", {"Status": "Scored"}, hint="0")
ok(_c2.main[1]["Status"] == "Scored",
   "a client without row_values_at degrades to a full scan rather than failing")


# ── C2. THE KEY IS NOT UNIQUE ───────────────────────────────────────────────
print("\n=== C2. DUPLICATE APPLICATION IDs RESOLVE TO THE INTENDED ROW ===")

# Re-processing can leave a stale retry row beside the completed row for the same candidate,
# both carrying the same Application ID. Giving up on the retry row must not delete the
# Scored one. test_p2.py pins this exact shape; it is repeated here because it is the reason
# the hint exists at all.
_dup = [row("APP-DUP", "Scored"), row("OTHER"), row("APP-DUP", "Needs Review - Unreadable Resume")]
_c = FakeClient(main=list(_dup))
_s = ExcelCandidateStore(_c)
_s.delete_by_id("APP-DUP", hint=2)
ok([v[APP_ID_COL] for v in _c.main] == ["APP-DUP", "OTHER"]
   and _c.main[0]["Status"] == "Scored",
   "with duplicate keys the hinted row is deleted, not the first match")

_c = FakeClient(main=list(_dup))
_s = ExcelCandidateStore(_c)
_s.save_by_id("APP-DUP", {"Retry Count": 3}, hint=2)
ok(_c.main[2].get("Retry Count") == 3 and "Retry Count" not in _c.main[0],
   "with duplicate keys the hinted row is patched, not the first match")

# A client that cannot verify positionally must still honour the hint as a tie-break.
_c = FakeClientNoSingleRead(main=list(_dup))
_s = ExcelCandidateStore(_c)
_s.delete_by_id("APP-DUP", hint=2)
ok(len(_c.main) == 2 and _c.main[0]["Status"] == "Scored",
   "the hint breaks the tie even when the row cannot be verified first")

# No hint and an ambiguous key: pick the first, deterministically, and say so.
_c = FakeClient(main=list(_dup))
_s = ExcelCandidateStore(_c)
_s.save_by_id("APP-DUP", {"Retry Count": 1})
ok(_c.main[0].get("Retry Count") == 1,
   "an ambiguous key with no hint resolves to the first match deterministically")

# The hint must still lose to the key when it points somewhere the key is not.
_c = FakeClient(main=list(_dup))
_s = ExcelCandidateStore(_c)
_s.save_by_id("OTHER", {"Status": "Scored"}, hint=0)
ok(_c.main[1]["Status"] == "Scored" and _c.main[0]["Status"] == "Scored"
   and _c.main[1][APP_ID_COL] == "OTHER",
   "a hint that does not hold the key is still discarded in favour of the key")


# Absence is only authoritative from a client that can actually verify a position. The older
# stubs answer list_rows with a token table that never held the row being written, and the
# pre-store _live_row_index trusted their arithmetic. That compatibility is deliberate.
class _EmptyTableNoVerify(FakeClientNoSingleRead):
    def list_rows(self):
        return []

_s = ExcelCandidateStore(_EmptyTableNoVerify())
ok(_s._resolve("APP-X", "main", 2) == 2,
   "an unverifiable client with an empty table falls back to the caller's position")

_verifiable = FakeClient(main=[row("APP-1")])
try:
    ExcelCandidateStore(_verifiable)._resolve("APP-GONE", "main", 2)
    _authoritative = False
except RowVanished:
    _authoritative = True
ok(_authoritative,
   "a client that CAN verify is authoritative: no match really means the row is gone")


# ── D. REJECTION IS ONE OPERATION ───────────────────────────────────────────
print("\n=== D. REJECTION IS ONE OPERATION ===")

_c = FakeClient(main=[row("APP-1"), row("APP-2", "Scored")])
_s = ExcelCandidateStore(_c)
_moved = _s.move_to_rejected("APP-2", {APP_ID_COL: "APP-2", "Status": "Rejected - Non-USA Location"})
ok(_moved is True, "move_to_rejected reports success when both halves land")
ok([v[APP_ID_COL] for v in _c.main] == ["APP-1"], "the candidate left the main sheet")
ok([v[APP_ID_COL] for v in _c.rejected] == ["APP-2"], "the candidate arrived on Rejected")

# The ghost-row symptom: the add succeeds, the delete 409s. It must be reported, not swallowed.
_c = FakeClient(main=[row("APP-9", "Scored")])
_c.delete_fails = True
_s = ExcelCandidateStore(_c)
_moved = _s.move_to_rejected("APP-9", {APP_ID_COL: "APP-9"})
ok(_moved is False,
   "a failed delete after a successful add returns False rather than reporting success")
ok(len(_c.rejected) == 1 and len(_c.main) == 1,
   "the ghost row is visible to the caller — on both sheets, and admitted")


# ── E. THE POSITIONAL ESCAPE HATCH ──────────────────────────────────────────
print("\n=== E. THE POSITIONAL ESCAPE HATCH ===")

# Descending-order batch deletion is order-independent: removing from the bottom up can never
# move a row still queued for deletion. This is the only sanctioned positional use.
_c = FakeClient(main=[row("APP-1"), row("APP-2"), row("APP-3"), row("APP-4")])
_s = ExcelCandidateStore(_c)
for _i in sorted({1, 3}, reverse=True):
    _s.delete_at(_i)
ok([v[APP_ID_COL] for v in _c.main] == ["APP-1", "APP-3"],
   "descending-order positional deletes remove exactly the intended rows")


# ── F. THE REJECTED SHEET ───────────────────────────────────────────────────
print("\n=== F. THE REJECTED SHEET ===")

_c = FakeClient(rejected=[row("APP-1", "Rejected - Non-USA Location"), {}])
_s = ExcelCandidateStore(_c)
ok(_s.get("APP-1", sheet="rejected").values["Status"] == "Rejected - Non-USA Location",
   "get reads the Rejected sheet when asked")
ok(len(_s.all_rows("rejected")) == 2,
   "all_rows keeps the Rejected sheet's orphaned blank rows")
_s.save_by_id("APP-1", {"Decline Sent": "Sent 2026-09-06 10:00"}, sheet="rejected")
ok(_c.rejected[0]["Decline Sent"] == "Sent 2026-09-06 10:00",
   "save_by_id patches a Rejected-sheet row by key")
ok(_c.writes[-1][0] == "rejected", "the Rejected write went to the Rejected sheet")

_s.add({APP_ID_COL: "APP-NEW"}, sheet="main")
ok(_c.adds[-1][0] == "main", "add appends to the sheet it was given")


# -- G. EXTRACTION PROVENANCE ------------------------------------------------
print("\n=== G. A DEGRADED EXTRACTION IS NAMED, NOT SILENT ===")

# Measured 2026-09-06 on the client's model: an ordinary two-page resume takes ~62s against a
# 60s budget, so the timeout fallback is common, not exceptional. It used to be invisible -
# the regex parser's output went out as 'Scored' with nothing recording that the model never
# actually read the resume.
import requests as _rq
from unittest.mock import patch as _patch
import hiring_agent.extraction as _ex

_TXT = "Jane Doe\nAustin, TX | (512) 555-0100 | jane@x.com\nSKILLS\nPython, AWS\n"

with _patch.object(_rq, "post", side_effect=_rq.Timeout("too slow")):
    ok(_ex.extract_with_ollama(_TXT) is None, "a timeout still falls back rather than raising")
ok("timed out" in _ex.EXTRACTION_SOURCE["value"],
   f"a timeout is recorded as a timeout (got {_ex.EXTRACTION_SOURCE['value']!r})")

with _patch.object(_rq, "post", side_effect=_rq.ConnectionError("refused")):
    _ex.extract_with_ollama(_TXT)
ok(_ex.EXTRACTION_SOURCE["value"] == "offline (ConnectionError)",
   "a dead host is distinguished from a timeout")

with _patch.object(_ex, "OLLAMA_ENABLED", False):
    _ex.extract_with_ollama(_TXT)
ok(_ex.EXTRACTION_SOURCE["value"] == "offline (Ollama disabled)",
   "a deliberately disabled model is distinguished from a failure")

# Reset per call, or a healthy candidate inherits the previous one stalling.
with _patch.object(_rq, "post", side_effect=_rq.Timeout("slow")):
    _ex.extract_with_ollama(_TXT)
_stale = _ex.EXTRACTION_SOURCE["value"]
with _patch.object(_ex, "OLLAMA_ENABLED", False):
    _ex.extract_with_ollama(_TXT)
ok(_ex.EXTRACTION_SOURCE["value"] != _stale,
   "the record is reset per call, so a stall never leaks into the next candidate")


# -- H. THE STORE'S "IT DID NOT HAPPEN" ANSWER MUST REACH THE CALLER ----------
print("\n=== H. A BLOCKED DELETE IS NEVER REPORTED AS A COMPLETED REJECTION ===")

# Found in review, 2026-09-06. All three bugs below shipped green: the suites never exercised
# a delete that FAILS, which is the normal case on an Excel table (409 InsertDeleteConflict).

# H1. FAIL CLOSED must hold on the rejected sheet too. `_can_verify` is false there for a
# structural reason - no single-row read endpoint exists for it - so phrasing the stub
# compatibility shim as `not _can_verify(sheet)` silently disabled fail-closed for every
# rejected-sheet write, landing unlocatable keys on a stale position.
_c = FakeClient(rejected=[row("APP-1")])
try:
    ExcelCandidateStore(_c)._resolve("APP-GONE", "rejected", 7)
    _closed = False
except RowVanished:
    _closed = True
ok(_closed, "an unlocatable key on the REJECTED sheet fails closed, it does not use the hint")

_c = FakeClient(main=[row("APP-1")])
ok(ExcelCandidateStore(_c)._resolve("APP-1", "main", 0) == 0,
   "...while a locatable main-sheet key still resolves normally")

# H2. _finish_rejection must not claim success when the main copy could not be removed.
from hiring_agent.sharepoint_scoring import _finish_rejection, _drop_retry_duplicate_if_completed

class _StuckDelete:
    """Add succeeds, delete 409s - the live ghost-row shape."""
    resumes_folder = "/Candidate_Resumes"
    def __init__(self):
        self.rejected_added = []
    def list_rows(self):
        return [{"index": 0, "values": {APP_ID_COL: "APP-STUCK", "Status": "Scored"}}]
    def row_values_at(self, i):
        return {APP_ID_COL: "APP-STUCK"} if i == 0 else None
    def add_rejected_row(self, f):
        self.rejected_added.append(f)
    def delete_row(self, i):
        raise RuntimeError("409 InsertDeleteConflict")

_sd = _StuckDelete()
_res = _finish_rejection(_sd, 0, "APP-STUCK", {APP_ID_COL: "APP-STUCK"},
                         {"Status": "Rejected - Non-USA Location"}, [], "2026/June", False,
                         set(), set(), set(), "non-USA")
ok(_res is False,
   "_finish_rejection returns False when the main row could not be deleted")
ok(len(_sd.rejected_added) == 1,
   "...the candidate is still on the Rejected sheet, so the ghost row is real and reported")

# H3. _drop_retry_duplicate_if_completed must not report a removal that did not occur.
class _AlreadyGone:
    resumes_folder = "/Candidate_Resumes"
    def list_rows(self):
        return [{"index": 0, "values": {APP_ID_COL: "APP-OTHER", "Status": "Scored"}}]
    def list_rejected_rows(self):
        return [{"index": 0, "values": {APP_ID_COL: "APP-DUP", "Status": "Scored"}}]
    def row_values_at(self, i):
        return {APP_ID_COL: "APP-OTHER"} if i == 0 else None
    def delete_row(self, i):
        raise AssertionError("must not be called - the row is not there")

ok(_drop_retry_duplicate_if_completed(_AlreadyGone(), 5, "APP-DUP",
                                      {APP_ID_COL: "APP-DUP"}) is False,
   "a stale retry row that cannot be located reports False, not a phantom deletion")


# -- I. MAINTENANCE SCANS STEP OVER SPACER ROWS -------------------------------
print("\n=== I. A BLANK SPACER IS NOT A CANDIDATE ===")

# Found in review, 2026-09-06. The whole-sheet backfill passes reached client.update_row
# through a `writer` variable, so they escaped the conversion to key-based addressing AND had
# no Application ID filter. list_rejected_rows() does not drop blank rows, so a backfill would
# stamp 'Missing' into all ten columns of each of the ~14 orphaned spacers on Rejected,
# turning cosmetic blanks into fake candidates every later pass then tries to process.
from hiring_agent.sharepoint_scoring import _save_scan_row

class _ScanClient:
    def __init__(self): self.writes = []
    def list_rows(self): return [{"index": 0, "values": {APP_ID_COL: "APP-1", "Phone": ""}}]
    def row_values_at(self, i): return {APP_ID_COL: "APP-1"} if i == 0 else None
    def update_row(self, i, f, current_values=None): self.writes.append(("main", i, f))
    def update_rejected_row(self, i, f, current_values=None): self.writes.append(("rejected", i, f))

_sc = _ScanClient()
ok(_save_scan_row(_sc, {"index": 3, "values": {}}, {"Phone": "Missing"}, sheet="rejected") is False,
   "an all-blank spacer row is skipped, not stamped")
ok(_sc.writes == [], "...and no write is issued for it")

ok(_save_scan_row(_sc, {"index": 9, "values": {APP_ID_COL: "APP-1", "Phone": ""}},
                  {"Phone": "Missing"}, sheet="main") is True,
   "a real candidate row is still patched")
ok(_sc.writes == [("main", 0, {"Phone": "Missing"})],
   "...by Application ID, at its LIVE position 0, not the stale hint 9")

_sc2 = _ScanClient()
ok(_save_scan_row(_sc2, {"index": 0, "values": {"Full Name": "No Id Here"}},
                  {"Phone": "Missing"}, sheet="main") is True,
   "a row with content but no Application ID is still repaired, positionally")


# -- J. AI-ONLY EXTRACTION (require_ai) ---------------------------------------
print("\n=== J. A ROW IS NEVER PUBLISHED FROM ANYTHING BUT A REAL MODEL READ ===")

# Set 2026-09-06. Measured on the client hardware, an ordinary 3.6k-char CV took 62.1s
# against the then-60s budget, so extraction timed out about half the time and SILENTLY
# published regex output as 'Scored'. require_ai removes the silent path entirely.
import hiring_agent.config as _cfg
import hiring_agent.sharepoint_scoring as _ss

ok(_cfg.REQUIRE_AI is True, "require_ai is on")
ok(_cfg.OLLAMA_TIMEOUT >= 180,
   f"the extraction budget is realistic for this hardware (got {_cfg.OLLAMA_TIMEOUT}s)")

# Brain down + require_ai -> the whole run stops. Nothing may be written, because with the
# model unavailable EVERY row in the batch would degrade.
with _patch("sharepoint_client.check_graph_reachable", return_value=(True, "ok")), _patch.object(_ss, "ollama_health", lambda: (False, "Ollama unreachable")):
    _r = _ss.score_from_sharepoint(dry_run=False)
ok("error" in _r and "require_ai" in _r["error"],
   "an unhealthy brain aborts the run with an actionable message")
ok(_r.get("processed") == 0 and _r.get("deferred") == 0,
   "...having scored nothing at all")

# The escape hatch still exists and is honest about what it does.
ok("require_ai=false" in _r["error"],
   "the abort message names the setting that re-enables the offline fallback")

# Deferral must never touch the retry counter: that path ends in
# 'Rejected - Processing Error', and a slow machine must not reject a real applicant.
from pathlib import Path as _Path
_src = _Path("hiring_agent/sharepoint_scoring.py").read_text(encoding="utf-8")
_defer_block = _src[_src.index("if _cfg.REQUIRE_AI:"):]
_defer_block = _defer_block[:_defer_block.index("continue")]
# Strip comments first - the block's own docstring explains WHY it avoids Retry Count, so a
# naive substring check would match the explanation rather than the code.
_defer_code = chr(10).join(l for l in _defer_block.splitlines()
                        if not l.strip().startswith("#"))
ok("Retry Count" not in _defer_code and "_mark_needs_review" not in _defer_code
   and "_give_up" not in _defer_code,
   "the deferral path bumps no retry counter and never routes to give-up/rejection")


# -- K. SQLITE CANDIDATE STORE ------------------------------------------------
print("\n=== K. SQLITE CANDIDATE STORE ===")
from hiring_agent.store import SQLiteCandidateStore

_sql_store = SQLiteCandidateStore(db_path=":memory:")
_sql_store.add({"Application ID": "APP-SQL-1", "Full Name": "Alice", "Status": "New Email Received"})
_sql_store.add({"Application ID": "APP-SQL-2", "Full Name": "Bob", "Status": "Scored"})
_sql_store.add({"Application ID": "", "Full Name": "Spacer", "Status": ""})

# Reads
_q = _sql_store.get_queue(("New Email Received",))
ok(len(_q) == 1 and _q[0].app_id == "APP-SQL-1", "sqlite get_queue filters by status and skips keyless")
ok(_sql_store.get("APP-SQL-1").values.get("Full Name") == "Alice", "sqlite get retrieves by app_id")
ok(len(_sql_store.all_rows("main")) == 3, "sqlite all_rows preserves keyless rows")

# Updates
_sql_store.save_by_id("APP-SQL-1", {"Status": "Scored", "Category": "Engineering"})
ok(_sql_store.get("APP-SQL-1").values.get("Category") == "Engineering", "sqlite save_by_id updates fields")
ok(_sql_store.get("APP-SQL-1").values.get("Status") == "Scored", "sqlite save_by_id updates status")

# Move to rejected (Atomic, zero ghost rows)
_sql_store.move_to_rejected("APP-SQL-1", {"Status": "Rejected - Non-USA Location"})
ok(_sql_store.get("APP-SQL-1", "main") is None, "sqlite move_to_rejected leaves main sheet")
ok(_sql_store.get("APP-SQL-1", "rejected") is not None, "sqlite move_to_rejected arrives on rejected")

# Delete
ok(_sql_store.delete_by_id("APP-SQL-2") is True, "sqlite delete_by_id deletes existing row")
ok(_sql_store.delete_by_id("APP-NONEXISTENT") is False, "sqlite delete_by_id returns False if absent")


print(f"\n{'='*64}")
print(f"  NEW-COVERAGE RESULT: {P} passed, {F} failed")
print(f"{'='*64}")
import sys
sys.exit(1 if F else 0)
