"""Excel date parsing and presentation for generated candidate workbooks."""

import datetime

import pandas as pd
from openpyxl.styles import PatternFill
from openpyxl.worksheet.filters import AutoFilter, FilterColumn

# Every candidate sheet reads header / blank spacer / first real row. The spacer is not
# decoration: P1's intake workbook has shipped with one since it was created, because an
# Excel table whose data body is empty behaves differently from one with a row in it, and
# both phases already skip any row with no Application ID. Painting it light blue makes
# the boundary between the header and the data obvious, and makes it clear the row is
# deliberate rather than someone's stray blank.
SPACER_FILL = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")

# The columns a reader may filter on. Every other header keeps its label but loses its
# dropdown, so the sheet offers the handful of cuts people actually want - by role area,
# by where someone is, by when they applied, and by what they studied - instead of a
# dropdown on all 33. Defined ONCE here and shared by both producers, the master report
# and the client sheet, so the two cannot drift apart the way their row-selection
# filters already have.
FILTERABLE_COLUMNS = ("Category", "Location", "Country", "Received Date",
                      "Education", "Education Start Date", "Education End Date")


def finish_sheet(worksheet, columns, *, table=None, allowed=FILTERABLE_COLUMNS) -> None:
    """Freeze the header row and leave a filter dropdown on `allowed` columns only.

    `table` is the openpyxl Table when the sheet is a real Excel table (the master
    report); omit it for a plain filtered range (the client sheet). Both carry their
    filter state in a different place, which is the whole reason this is one function.

    NO sheet protection. It was applied here briefly and removed on request: these are
    generated files that get overwritten wholesale every run, so protection bought
    nothing an operator wanted and blocked the ordinary housekeeping - deleting rows,
    deleting the file's contents - that people actually do with them.
    """
    ref = table.ref if table is not None else worksheet.dimensions
    hidden = [FilterColumn(colId=index, hiddenButton=True)
              for index, name in enumerate(columns) if name not in allowed]
    if table is not None:
        table.autoFilter = AutoFilter(ref=ref, filterColumn=hidden)
    else:
        worksheet.auto_filter.ref = ref
        worksheet.auto_filter.filterColumn = hidden

    worksheet.freeze_panes = "A2"
    worksheet.protection.sheet = False
    paint_spacer_row(worksheet, len(columns))


def paint_spacer_row(worksheet, width: int) -> bool:
    """Fill row 2 light blue when it is the blank spacer. No-op if it holds data.

    Guarded rather than unconditional so a sheet that was built without a spacer, or one
    whose caller forgot to add it, can never end up with its first candidate painted as
    though it were a separator.
    """
    if worksheet.max_row < 2:
        return False
    if any(str(worksheet.cell(2, col).value or "").strip() for col in range(1, width + 1)):
        return False
    for col in range(1, width + 1):
        worksheet.cell(2, col).fill = SPACER_FILL
    return True


def _parse_received(value) -> datetime.date:
    """Best-effort date from a 'received' value (str/datetime/Excel serial);
    falls back to today.

    Graph returns a raw Excel date serial (days since 1899-12-30, e.g.
    46205.75 = 2026-07-02) when the workbook cell is date-formatted — Excel
    Online coerces date-looking strings on row insert. A naive to_datetime
    reads that number as nanoseconds since 1970 and silently yields
    1970-01-01, which pointed resume downloads at a 1970/January folder
    that never exists, so the row could never score.
    """
    if value not in (None, ""):
        try:
            num = float(value)
        except (TypeError, ValueError):
            num = None
        if num is not None and 20000 <= num <= 80000:   # plausible serial: 1954..2118
            ts = pd.to_datetime(num, unit="D", origin="1899-12-30", errors="coerce")
        else:
            ts = pd.to_datetime(value, errors="coerce")
        if pd.notna(ts):
            return ts.date()
    return datetime.date.today()
