"""Parse the WSU vUWS 'User Activity in Forums' report.

The file extension is .xls but the format is SpreadsheetML XML.

This report is treated as CONTEXT ONLY — it feeds the report's forum
columns but does not drive S1-S8 segmentation (login count and subject
hours do that; see metrics.py / segmentation.py).

Layout, as observed in real exports:
  - 'Access / Forum' section: chart label list, then one or more
    per-student data blocks (each block header row lists a batch of forum
    names in scattered, non-contiguous columns — Blackboard wraps wide
    tables into stacked header/body blocks rather than one wide table).
  - 'Messages / Forum' section: same shape, for message-post counts.
    May be entirely absent from an export if no messages were posted in
    any forum that period.
  - 'Access / Date' section: subject-wide (not per-student) daily hit
    counts for the export's date window — used here only to infer the
    window start/end when the report has no explicit date-range header.

We don't hard-code column positions; we detect header rows in each
section by presence of multiple non-numeric text cells beyond column 1,
and treat every row after a header (until the next header or the section
ends) as a student data row, keyed by the 'Name (ID)' pattern.
"""

import re
import xml.etree.ElementTree as ET
from datetime import date

import pandas as pd

NS = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
SS = "{urn:schemas-microsoft-com:office:spreadsheet}"

STUDENT_RE = re.compile(r"^(.*?)\s*\((\w+)\)\s*$")
ISO_DATE_RE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})$")

SECTION_ACCESS = "Access / Forum"
SECTION_MESSAGES = "Messages / Forum"
SECTION_DATE = "Access / Date"
KNOWN_SECTIONS = (SECTION_ACCESS, SECTION_MESSAGES, SECTION_DATE)


def _cells_by_idx(row) -> dict[int, str]:
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


EXPECTED_SHEET_NAME = "Accesses by Forum"


def _all_rows(path: str):
    tree = ET.parse(path)
    root = tree.getroot()
    ws = root.find("ss:Worksheet", NS)
    table = ws.find("ss:Table", NS)
    return table.findall("ss:Row", NS)


def _sheet_name(path: str) -> str | None:
    tree = ET.parse(path)
    root = tree.getroot()
    ws = root.find("ss:Worksheet", NS)
    return ws.get(SS + "Name") if ws is not None else None


def _extract_student_code(cell_value: str) -> str | None:
    if not cell_value:
        return None
    m = STUDENT_RE.match(cell_value.strip())
    return m.group(2).strip() if m else None


def _is_header_row(cells: dict[int, str]) -> bool:
    """A header row: col 1 empty, several non-numeric text cells beyond col 1."""
    if cells.get(1, "").strip():
        return False
    text_cells = [
        v.strip() for k, v in cells.items()
        if k > 1 and v and not v.strip().lstrip("-").replace(".", "", 1).isdigit()
    ]
    return len(text_cells) >= 2


def _parse_student_table(rows, start_idx, end_idx) -> list[dict]:
    """Parse one or more stacked header/body blocks of a per-student forum table.

    Returns long-format records: student_code, forum, hits.
    """
    records = []
    col_forum: dict[int, str] = {}
    for i in range(start_idx, end_idx):
        cells = _cells_by_idx(rows[i])
        if _is_header_row(cells):
            col_forum = {
                k: v.strip() for k, v in cells.items()
                if k > 1 and v and not v.strip().lstrip("-").replace(".", "", 1).isdigit()
                and v.strip() not in ("Total", "Per cent", "Percent")
            }
            continue
        name_val = cells.get(1, "").strip()
        student_code = _extract_student_code(name_val)
        if not student_code or not col_forum:
            continue
        if name_val in ("Guest", "Total") or "PreviewUser" in name_val:
            continue
        for col_idx, forum in col_forum.items():
            raw = cells.get(col_idx, "")
            if raw is None or raw == "":
                continue
            try:
                hits = int(float(raw))
            except (ValueError, TypeError):
                continue
            if hits:
                records.append({
                    "student_code": student_code,
                    "forum": forum,
                    "hits": hits,
                })
    return records


def _find_marker_rows(rows) -> dict[str, int]:
    """Return {section_marker_text: row_idx} for the top-level chart markers."""
    markers = {}
    for i, row in enumerate(rows):
        cells = _cells_by_idx(row)
        col0 = cells.get(0, "").strip()
        if col0 in KNOWN_SECTIONS and col0 not in markers:
            markers[col0] = i
    return markers


def _find_data_table_clusters(rows, search_start, search_end) -> list[list[int]]:
    """Group header-block row indices into table clusters.

    Blackboard wraps a wide per-student table into several stacked
    header/body blocks (one per batch of forum columns). Consecutive
    header blocks belong to the SAME table when their forum-name sets
    don't overlap (i.e. each new block introduces columns for forums not
    already seen); a header block that repeats an already-seen forum name
    signals the start of a genuinely different table (e.g. Access -> a
    separate Messages table for the same forums).
    """
    header_idxs = []
    for i in range(search_start, search_end):
        cells = _cells_by_idx(rows[i])
        if _is_header_row(cells):
            header_idxs.append(i)

    clusters: list[list[int]] = []
    seen_forums: set[str] = set()
    for idx in header_idxs:
        cells = _cells_by_idx(rows[idx])
        forums_here = {
            v.strip() for k, v in cells.items()
            if k > 1 and v and not v.strip().lstrip("-").replace(".", "", 1).isdigit()
            and v.strip() not in ("Total", "Per cent", "Percent")
        }
        if clusters and not (forums_here & seen_forums):
            clusters[-1].append(idx)
            seen_forums |= forums_here
        else:
            clusters.append([idx])
            seen_forums = set(forums_here)
    return clusters


def parse(path: str) -> dict:
    """Parse the forums report.

    Returns a dict:
      'access':   long-format DataFrame (student_code, forum, hits) — forum views
      'messages': long-format DataFrame (student_code, forum, hits) — forum posts
      'window_start', 'window_end': date | None, inferred from the
                  subject-wide 'Access / Date' section (min/max date seen)

    Data-table identity (Access vs Messages) is assigned by ORDER of
    appearance among detected table clusters, not by proximity to a
    section-title marker — Blackboard lists both charts' forum-name
    labels up front before either table appears, so marker proximity is
    not a reliable way to tell the two tables apart. If only one table
    cluster is found, it's treated as Access (the more commonly populated
    of the two) and 'messages' comes back empty.

    Raises ValueError if the file's internal sheet name doesn't match a
    User Activity in Forums export — most commonly caused by uploading a
    Subject Activity Overview file (sheet name 'Course Activity Overview')
    into the forums slot by mistake.
    """
    sheet_name = _sheet_name(path)
    if sheet_name and sheet_name.strip() != EXPECTED_SHEET_NAME:
        raise ValueError(
            f"this file's internal sheet is named '{sheet_name}', not "
            f"'{EXPECTED_SHEET_NAME}' — it doesn't look like a User "
            f"Activity in Forums export. If this came from the Subject "
            f"Activity Overview report instead, upload it in that slot, "
            f"not this one."
        )

    rows = _all_rows(path)
    markers = _find_marker_rows(rows)

    search_start = markers.get(SECTION_ACCESS, 0) + 1
    search_end = markers.get(SECTION_DATE, len(rows))

    clusters = _find_data_table_clusters(rows, search_start, search_end)

    access_records: list[dict] = []
    messages_records: list[dict] = []
    if len(clusters) >= 1:
        block = clusters[0]
        next_start = clusters[1][0] if len(clusters) > 1 else search_end
        access_records = _parse_student_table(rows, block[0], next_start)
    if len(clusters) >= 2:
        block = clusters[1]
        next_start = clusters[2][0] if len(clusters) > 2 else search_end
        messages_records = _parse_student_table(rows, block[0], next_start)

    window_start = window_end = None
    if SECTION_DATE in markers:
        s = markers[SECTION_DATE] + 1
        dates = []
        for i in range(s, len(rows)):
            cells = _cells_by_idx(rows[i])
            v = (cells.get(1, "") or "").strip()
            m = ISO_DATE_RE.match(v)
            if m:
                dates.append(date(int(m.group(1)), int(m.group(2)), int(m.group(3))))
        if dates:
            window_start, window_end = min(dates), max(dates)

    access_df = pd.DataFrame(access_records, columns=["student_code", "forum", "hits"])
    messages_df = pd.DataFrame(messages_records, columns=["student_code", "forum", "hits"])

    return {
        "access": access_df,
        "messages": messages_df,
        "window_start": window_start,
        "window_end": window_end,
    }


def per_student_totals(parsed: dict) -> pd.DataFrame:
    """Collapse access + messages into one per-student total interaction count.

    Returns: student_code, forum_accesses, forum_messages, forum_interactions
    """
    access = parsed["access"]
    messages = parsed["messages"]

    acc_totals = (
        access.groupby("student_code")["hits"].sum().rename("forum_accesses")
        if not access.empty else pd.Series(name="forum_accesses", dtype=int)
    )
    msg_totals = (
        messages.groupby("student_code")["hits"].sum().rename("forum_messages")
        if not messages.empty else pd.Series(name="forum_messages", dtype=int)
    )

    out = pd.concat([acc_totals, msg_totals], axis=1).fillna(0).astype(int)
    out = out.reset_index().rename(columns={"index": "student_code"})
    if "student_code" not in out.columns and not out.empty:
        out = out.rename(columns={out.columns[0]: "student_code"})
    if out.empty:
        out = pd.DataFrame(columns=["student_code", "forum_accesses", "forum_messages"])
    out["forum_interactions"] = out.get("forum_accesses", 0) + out.get("forum_messages", 0)
    return out
