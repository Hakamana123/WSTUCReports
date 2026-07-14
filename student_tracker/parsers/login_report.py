"""Parse the WSU vUWS Subject Login Report (.xlsx).

Two sections exist (students who logged in this period vs. those who didn't),
each with the same logical columns but at slightly different column offsets.
We locate each section by its header row and stitch them together.

IMPORTANT: DAYS SINCE LAST LOGIN / LAST LOGIN DATE / TOTAL LOGINS are
lifetime figures — they are NOT scoped to whatever date range was picked
in the Blackboard filter when the report was run. Only which SECTION a
student lands in (logged-in-this-period vs not) reflects the chosen
window. This means:
  - S1/S2 checks (which need a true lifetime "ever logged in" / "last
    login before block start") work correctly from ANY single weekly
    pull — no separate lifetime-range pull is needed.
  - A per-week LOGIN COUNT is derived elsewhere (metrics.py) by taking
    the difference between this week's TOTAL LOGINS snapshot and the
    previous week's, across consecutive weekly uploads.
"""

import re
from datetime import date, datetime

import pandas as pd

DATE_WINDOW_RE = re.compile(
    r"between\s+(\d{1,2}/\d{1,2}/\d{4})\s+to\s+(\d{1,2}/\d{1,2}/\d{4})",
    re.IGNORECASE,
)

EXPECTED_COLS = [
    "SURNAME", "FIRST NAME", "STUDENT ID", "EMAIL",
    "DAYS SINCE LAST LOGIN", "LAST LOGIN DATE", "TOTAL LOGINS",
]


def _find_header_rows(raw: pd.DataFrame) -> list[int]:
    """Find rows that contain SURNAME, FIRST NAME, STUDENT ID — section headers."""
    hits = []
    for i in range(len(raw)):
        vals = [str(v) for v in raw.iloc[i].tolist() if pd.notna(v)]
        joined = " ".join(vals).upper()
        if "SURNAME" in joined and "STUDENT ID" in joined and "TOTAL LOGINS" in joined:
            hits.append(i)
    return hits


def _column_map(header_row: pd.Series) -> dict[str, int]:
    """From a header row, return {logical_name: column_index}."""
    mapping = {}
    for col_idx, val in header_row.items():
        if pd.isna(val):
            continue
        v = str(val).strip().upper()
        if v in EXPECTED_COLS:
            mapping[v] = col_idx
    return mapping


def _extract_section(raw: pd.DataFrame, header_idx: int,
                     end_idx: int | None) -> pd.DataFrame:
    cmap = _column_map(raw.iloc[header_idx])
    if not all(c in cmap for c in EXPECTED_COLS):
        return pd.DataFrame(columns=EXPECTED_COLS)
    end = end_idx if end_idx is not None else len(raw)
    body = raw.iloc[header_idx + 1:end]
    out = pd.DataFrame({col: body[cmap[col]] for col in EXPECTED_COLS})
    out = out.dropna(subset=["STUDENT ID"]).reset_index(drop=True)
    return out


def parse(path: str) -> pd.DataFrame:
    """Return one DataFrame combining both sections of the login report.

    Columns: student_code, surname, first_name, email,
             days_since_last_login, last_login_date, total_logins
    """
    raw = pd.read_excel(path, header=None)
    header_rows = _find_header_rows(raw)
    if not header_rows:
        raise ValueError("Could not find any 'SURNAME / STUDENT ID' header rows.")

    sections = []
    for i, hdr in enumerate(header_rows):
        end = header_rows[i + 1] if i + 1 < len(header_rows) else None
        sections.append(_extract_section(raw, hdr, end))
    combined = pd.concat(sections, ignore_index=True)

    out = combined.rename(columns={
        "SURNAME": "surname",
        "FIRST NAME": "first_name",
        "STUDENT ID": "student_code",
        "EMAIL": "email",
        "DAYS SINCE LAST LOGIN": "days_since_last_login",
        "LAST LOGIN DATE": "last_login_date",
        "TOTAL LOGINS": "total_logins",
    })
    out["student_code"] = out["student_code"].astype(str).str.strip()
    # Real date cells arrive from openpyxl as already-parsed datetime objects;
    # only the 'NEVER' sentinel (and occasional blanks) are strings. Replacing
    # those explicitly avoids pandas' noisy mixed-type format-inference
    # warning — and matters more than cosmetics, since this file's dates
    # aren't ambiguous DD/MM strings to begin with, so there's no format to
    # get wrong, just a type to normalise.
    raw_dates = out["last_login_date"]
    non_date_mask = raw_dates.apply(lambda v: not isinstance(v, (pd.Timestamp, datetime, date)))
    cleaned = raw_dates.mask(non_date_mask)
    out["last_login_date"] = pd.to_datetime(cleaned, errors="coerce")
    out["total_logins"] = pd.to_numeric(out["total_logins"], errors="coerce").fillna(0).astype(int)
    out["days_since_last_login"] = pd.to_numeric(out["days_since_last_login"], errors="coerce")
    out = out.drop_duplicates(subset=["student_code"], keep="first").reset_index(drop=True)
    return out


def parse_window(path: str) -> tuple[date | None, date | None]:
    """Find the report's 'between D/M/YYYY to D/M/YYYY' window text.

    Scans the first ~60 rows / 20 cols, matching the pattern used by the
    report's own descriptive text ('...students logged in between X to Y').
    Returns (None, None) if not found.
    """
    raw = pd.read_excel(path, header=None, nrows=60)
    for r in range(min(60, len(raw))):
        for c in range(min(20, raw.shape[1])):
            val = raw.iat[r, c]
            if not isinstance(val, str):
                continue
            m = DATE_WINDOW_RE.search(val)
            if m:
                try:
                    start = datetime.strptime(m.group(1), "%d/%m/%Y").date()
                    end = datetime.strptime(m.group(2), "%d/%m/%Y").date()
                    return start, end
                except ValueError:
                    continue
    return None, None


def parse_with_window(path: str) -> tuple[pd.DataFrame, date | None, date | None]:
    """Convenience wrapper: returns (login df, window_start, window_end)."""
    return parse(path), *parse_window(path)
