"""Parse the WSU vUWS 'Subject Activity Overview' report.

The file extension is .xls but the format is actually SpreadsheetML XML
(same family as the old Overall Summary of Usage report).

This report is run once per reporting window (typically narrowed to a
single teaching week in the Blackboard UI's date picker) and gives, per
student, the total hours spent in the subject during that window. It does
NOT give a daily breakdown and it does NOT give click/interaction counts —
just a single hours figure per student for whatever date range was
selected when the report was generated.

We treat the reported hours as scoped to the window shown in the report's
own 'Date Range' header (not cumulative/lifetime) — running this report
with a one-week window gives that week's hours; running it with a wider
window gives that wider window's total.
"""

import re
import xml.etree.ElementTree as ET
from datetime import date, datetime

import pandas as pd

NS = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
SS = "{urn:schemas-microsoft-com:office:spreadsheet}"

DATE_RANGE_RE = re.compile(
    r"(\d{1,2}/\d{1,2}/\d{2,4})\s*-\s*(\d{1,2}/\d{1,2}/\d{2,4})"
)


def _cells_by_idx(row) -> dict[int, str]:
    """Return {col_index: text_value} for a SpreadsheetML row (0-indexed), honouring ss:Index."""
    out: dict[int, str] = {}
    next_idx = 0
    for c in row.findall("ss:Cell", NS):
        idx_attr = c.get(SS + "Index")
        if idx_attr:
            next_idx = int(idx_attr) - 1
        d = c.find("ss:Data", NS)
        out[next_idx] = (d.text or "") if d is not None else ""
        next_idx += 1
    return out


def _all_rows(path: str):
    tree = ET.parse(path)
    root = tree.getroot()
    ws = root.find("ss:Worksheet", NS)
    table = ws.find("ss:Table", NS)
    return table.findall("ss:Row", NS)


def _parse_short_date(s: str) -> date | None:
    """Parse a 'd/m/yy' or 'd/m/yyyy' style date as used in the report header."""
    s = s.strip()
    for fmt in ("%d/%m/%y", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def parse_date_range(path: str) -> tuple[date | None, date | None]:
    """Find and parse the 'Date Range' header row. Returns (start, end), either may be None."""
    rows = _all_rows(path)
    for row in rows[:15]:
        cells = _cells_by_idx(row)
        joined = " ".join(v for v in cells.values() if v)
        if "Date Range" not in joined:
            continue
        m = DATE_RANGE_RE.search(joined)
        if m:
            return _parse_short_date(m.group(1)), _parse_short_date(m.group(2))
    return None, None


def parse(path: str) -> pd.DataFrame:
    """Return a DataFrame: student_code, hours.

    Locates the 'Student Overview' per-student table by its header row
    (containing both 'Student ID' and a header with 'Hours' in it), then
    reads rows until a blank row / footer is hit.
    """
    rows = _all_rows(path)

    header_idx = None
    col_map: dict[str, int] = {}
    for i, row in enumerate(rows):
        cells = _cells_by_idx(row)
        vals = {k: (v or "").strip() for k, v in cells.items()}
        id_col = next((k for k, v in vals.items() if v == "Student ID"), None)
        hours_col = next((k for k, v in vals.items() if "Hours" in v), None)
        if id_col is not None and hours_col is not None:
            header_idx = i
            col_map = {"id": id_col, "hours": hours_col}
            break

    if header_idx is None:
        raise ValueError(
            "Could not find the per-student 'Student ID' / 'Subject Activity "
            "in Hours' table in this Subject Activity Overview export."
        )

    records = []
    for row in rows[header_idx + 1:]:
        cells = _cells_by_idx(row)
        sid_raw = (cells.get(col_map["id"], "") or "").strip()
        if not sid_raw or not sid_raw.isdigit():
            # Blank row, footer row ('Powered by Blackboard...'), or the
            # table has ended.
            if sid_raw and not sid_raw.isdigit():
                continue
            if not cells:
                continue
            continue
        hours_raw = (cells.get(col_map["hours"], "") or "").strip()
        try:
            hours = float(hours_raw) if hours_raw else 0.0
        except ValueError:
            hours = 0.0
        records.append({"student_code": sid_raw, "hours": hours})

    df = pd.DataFrame(records, columns=["student_code", "hours"])
    if not df.empty:
        df = df.drop_duplicates(subset=["student_code"], keep="first").reset_index(drop=True)
    return df


def parse_with_window(path: str) -> tuple[pd.DataFrame, date | None, date | None]:
    """Convenience wrapper: returns (per-student hours df, window_start, window_end)."""
    df = parse(path)
    start, end = parse_date_range(path)
    return df, start, end
