"""Excel date parsing for candidate workbooks (Graph date-serial safe)."""

import datetime

import pandas as pd


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
