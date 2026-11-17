"""Read a class list from whatever spreadsheet holds it.

The class list isn't a vUWS export, so this accepts any of:
  - the university class list (.xls, metadata block then a 'student_code'
    header row, one sheet per class)
  - the Class List Collator output (.xlsx, 'Student ID' / 'First Name' /
    'Last Name' / 'Subject Code' or 'GEDU Subject')
  - any .xlsx / .xls / .csv with a recognisable student-ID column

The header row is found by looking for a student-ID heading in the first
15 rows of each sheet. Sheets named '...unallocated...' are skipped.
"""

from __future__ import annotations

import io
import re

import pandas as pd

ID_HEADERS = {"student_code", "student id", "student_id", "studentid",
              "student number", "student no", "student no.", "id"}
FIRST_HEADERS = ["preferred_name", "preferred name", "first_name", "first name",
                 "given name", "given_name", "firstname"]
LAST_HEADERS = ["last_name", "last name", "surname", "family name", "family_name", "lastname"]
NAME_HEADERS = ["name", "student name", "full name"]
EXTRA = {
    "attend_type": ["attend_type", "attend type", "attendance type", "mode"],
    "course": ["course", "course code", "program", "program code"],
    "email_address": ["email_address", "email", "email address"],
}
SUBJECT_HEADERS = ["display_subject_code", "subject code", "gedu subject", "subject"]
SUBJECT_RE = re.compile(r"([A-Z]{4}\d{4})")
SCAN_ROWS = 15


def _norm(v) -> str:
    return str(v).strip().lower() if pd.notna(v) else ""


def _read_sheets(data: bytes, name: str) -> list[pd.DataFrame]:
    low = name.lower()
    if low.endswith(".csv"):
        return [pd.read_csv(io.BytesIO(data), header=None, dtype=str, encoding="utf-8-sig")]
    engine = "openpyxl" if data[:2] == b"PK" else "xlrd"
    book = pd.read_excel(io.BytesIO(data), sheet_name=None, header=None, dtype=str, engine=engine)
    return [df for sheet, df in book.items() if "unallocated" not in sheet.lower()]


def _first_present(headers: dict[str, int], options: list[str]) -> int | None:
    for o in options:
        if o in headers:
            return headers[o]
    return None


def _clean(s: pd.Series) -> pd.Series:
    s = s.fillna("").astype(str).str.strip()
    return s.where(~s.str.lower().isin(["nan", "none"]), "")


def parse(data: bytes, name: str) -> pd.DataFrame:
    """Return student_code, display_name, subject_code (may be blank) and any
    of attend_type / course / email_address found. One row per student per
    subject."""
    frames = []
    for raw in _read_sheets(data, name):
        header_idx = None
        for i in range(min(SCAN_ROWS, len(raw))):
            if ID_HEADERS & {_norm(v) for v in raw.iloc[i]}:
                header_idx = i
                break
        if header_idx is None:
            continue
        headers = {}
        for c, v in enumerate(raw.iloc[header_idx]):
            h = _norm(v)
            if h and h not in headers:
                headers[h] = c
        body = raw.iloc[header_idx + 1:].reset_index(drop=True)
        id_col = _first_present(headers, sorted(ID_HEADERS, key=lambda h: h == "id"))
        sid = _clean(body[id_col]).str.replace(r"\.0$", "", regex=True)

        first_i = _first_present(headers, FIRST_HEADERS)
        last_i = _first_present(headers, LAST_HEADERS)
        full_i = _first_present(headers, NAME_HEADERS)
        if first_i is not None or last_i is not None:
            first = _clean(body[first_i]) if first_i is not None else ""
            # preferred_name may be blank: fall back to first_name
            if first_i is not None and "first_name" in headers and headers["first_name"] != first_i:
                first = first.where(first != "", _clean(body[headers["first_name"]]))
            last = _clean(body[last_i]) if last_i is not None else ""
            display = (first + " " + last).str.strip() if isinstance(first, pd.Series) else last
        elif full_i is not None:
            display = _clean(body[full_i])
        else:
            display = pd.Series("", index=body.index)

        out = pd.DataFrame({"student_code": sid, "display_name": display})
        subj_i = _first_present(headers, SUBJECT_HEADERS)
        if subj_i is not None:
            out["subject_code"] = _clean(body[subj_i]).str.extract(SUBJECT_RE, expand=False).fillna("")
        else:
            # university class list: subject code sits in the metadata block
            meta = " ".join(_norm(v).upper() for v in raw.iloc[:header_idx].stack().tolist())
            m = SUBJECT_RE.search(meta)
            out["subject_code"] = m.group(1) if m else ""
        for key, opts in EXTRA.items():
            i = _first_present(headers, opts)
            if i is not None:
                out[key] = _clean(body[i])
        frames.append(out[out["student_code"].str.fullmatch(r"\d{6,}")])

    if not frames:
        raise ValueError(
            "couldn't find a student ID column (looked for headings like "
            "'student_code', 'Student ID' or 'Student Number' in the first "
            f"{SCAN_ROWS} rows of each sheet)"
        )
    return pd.concat(frames, ignore_index=True).drop_duplicates(["student_code", "subject_code"])
