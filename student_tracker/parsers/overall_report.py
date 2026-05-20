"""Parse the WSU Overall Summary of Usage report.

The file extension is .xls but the format is actually SpreadsheetML XML
(an old MS Office XML format). We parse it as XML.

The report has multiple sections; we extract:
  - Per-student per-date hit counts (the 'Access / Date' section, which has
    one sub-table per calendar month)
  - Per-student per-LMS-area hit counts (the second 'Access / Application'
    sub-table) — used for the 'most active area' context per student.

Supports two date-section layouts:
  1. Native Blackboard: month sub-tables with YYYY-MM header in col 2,
     day numbers in cols 3+, student names in col 2.
  2. Flat ISO (legacy builder): single table with YYYY-MM-DD column
     headers in cols 2+, student names in col 1.
"""

import re
import xml.etree.ElementTree as ET
from datetime import date

import pandas as pd

NS = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
SS = "{urn:schemas-microsoft-com:office:spreadsheet}"

STUDENT_RE = re.compile(r"^(.*?)\s*\((\w+)\)\s*$")
MONTH_RE = re.compile(r"^(\d{4})-(\d{2})$")
ISO_DATE_RE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})$")


def _cells_by_idx(row) -> dict[int, str]:
    """Return {col_index: text_value} for a SpreadsheetML row, honouring ss:Index."""
    out: dict[int, str] = {}
    next_idx = 1
    for c in row.findall("ss:Cell", NS):
        idx_attr = c.get(SS + "Index")
        if idx_attr:
            next_idx = int(idx_attr)
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


def _extract_student_code(cell_value: str) -> str | None:
    if not cell_value:
        return None
    m = STUDENT_RE.match(cell_value.strip())
    return m.group(2).strip() if m else None


def parse_date_section(path: str) -> pd.DataFrame:
    """Return long-format DataFrame: student_code, date, hits.

    Handles two layouts:
      1. Month sub-tables (native Blackboard): YYYY-MM header in col 2,
         day numbers in cols 3+, student names in col 2.
      2. Flat ISO (legacy builder output): YYYY-MM-DD headers in a single
         row, student names in col 1.
    """
    rows = _all_rows(path)
    records: list[dict] = []

    in_date_section = False
    current_year = None
    current_month = None
    current_day_cols: dict[int, int] = {}  # col_idx -> day_number

    # Flat ISO fallback state
    flat_mode = False
    flat_date_cols: dict[int, date] = {}   # col_idx -> date

    for row in rows:
        cells = _cells_by_idx(row)

        # Detect entering the date section
        if cells.get(1, "").strip() == "Access / Date":
            in_date_section = True
            continue
        if not in_date_section:
            continue

        # Detect leaving the date section (next major section)
        col1 = cells.get(1, "").strip()
        if col1.startswith("Access /") and col1 != "Access / Date":
            break

        col2 = cells.get(2, "").strip()

        # ── Try month sub-table header (native format) ────────────
        m = MONTH_RE.match(col2)
        if m:
            flat_mode = False
            current_year = int(m.group(1))
            current_month = int(m.group(2))
            current_day_cols = {}
            for col_idx, val in cells.items():
                if col_idx <= 2:
                    continue
                v = val.strip() if val else ""
                if v.isdigit() and 1 <= int(v) <= 31:
                    current_day_cols[col_idx] = int(v)
            continue

        # ── Try flat ISO header row (legacy builder format) ───────
        # Detected by: col 1 = "Student" and multiple cols have
        # YYYY-MM-DD values.
        if col1 == "Student" and not flat_mode and current_year is None:
            candidate_dates = {}
            for col_idx, val in cells.items():
                if col_idx <= 1:
                    continue
                v = val.strip() if val else ""
                dm = ISO_DATE_RE.match(v)
                if dm:
                    try:
                        candidate_dates[col_idx] = date(
                            int(dm.group(1)), int(dm.group(2)), int(dm.group(3))
                        )
                    except ValueError:
                        pass
            if candidate_dates:
                flat_mode = True
                flat_date_cols = candidate_dates
                continue

        # ── Flat mode: student data rows ──────────────────────────
        if flat_mode:
            student_code = _extract_student_code(col1)
            if not student_code:
                continue
            for col_idx, d in flat_date_cols.items():
                raw = cells.get(col_idx, "")
                if raw is None or raw == "":
                    continue
                try:
                    hits = int(raw)
                except (ValueError, TypeError):
                    continue
                records.append({
                    "student_code": student_code,
                    "date": d,
                    "hits": hits,
                })
            continue

        # ── Month sub-table mode: student data rows ───────────────
        # Skip Total rows and Guest rows
        if col2 in ("Total", "Guest", "") or current_year is None:
            continue

        student_code = _extract_student_code(col2)
        if not student_code:
            continue

        for col_idx, day in current_day_cols.items():
            raw = cells.get(col_idx, "")
            if raw is None or raw == "":
                continue
            try:
                hits = int(raw)
            except (ValueError, TypeError):
                continue
            try:
                d = date(current_year, current_month, day)
            except ValueError:
                continue
            records.append({
                "student_code": student_code,
                "date": d,
                "hits": hits,
            })

    df = pd.DataFrame(records)
    if not df.empty:
        df["date"] = pd.to_datetime(df["date"])
        df = df.groupby(["student_code", "date"], as_index=False)["hits"].sum()
    return df


def parse_area_section(path: str) -> pd.DataFrame:
    """Return per-student per-LMS-area hit counts.

    Long-format: student_code, area, hits. Used for showing each student's
    most-active LMS area as supplementary context.

    Handles student names in either col 1 or col 2.
    """
    rows = _all_rows(path)
    records: list[dict] = []

    in_section = False
    in_per_student_table = False
    area_cols: dict[int, str] = {}
    student_col: int = 2  # default: col 2 (native format)

    for row in rows:
        cells = _cells_by_idx(row)
        col1 = cells.get(1, "").strip()
        col2 = cells.get(2, "").strip()

        # The area section comes first ('Access / Application')
        if col1 == "Access / Application":
            in_section = True
            continue

        # Stop when we hit the next major section
        if in_section and col1.startswith("Access /") and col1 != "Access / Application":
            break

        if not in_section:
            continue

        # Per-student sub-table header detection.
        # Native format: col 2 empty, area names in cols > 2.
        # Legacy builder: col 1 = "Student", area names in cols > 1.
        if not in_per_student_table:
            # Try native format: col 2 empty, many text values in cols > 2
            if not col2 and len(cells) > 5:
                candidate_areas = {
                    idx: val.strip() for idx, val in cells.items()
                    if idx > 2 and val and not val.strip().isdigit()
                }
                if len(candidate_areas) > 5:
                    area_cols = candidate_areas
                    student_col = 2
                    in_per_student_table = True
                    continue

            # Try legacy builder format: col 1 = "Student"
            if col1 == "Student" and len(cells) > 5:
                candidate_areas = {
                    idx: val.strip() for idx, val in cells.items()
                    if idx > 1 and val and not val.strip().isdigit()
                }
                if len(candidate_areas) > 5:
                    area_cols = candidate_areas
                    student_col = 1
                    in_per_student_table = True
                    continue

        if not in_per_student_table or not area_cols:
            continue

        student_val = cells.get(student_col, "").strip()
        student_code = _extract_student_code(student_val)
        if not student_code:
            continue

        for col_idx, area in area_cols.items():
            raw = cells.get(col_idx, "")
            if raw is None or raw == "":
                continue
            try:
                hits = int(raw)
            except (ValueError, TypeError):
                continue
            if hits > 0:
                records.append({
                    "student_code": student_code,
                    "area": area,
                    "hits": hits,
                })

    return pd.DataFrame(records)
