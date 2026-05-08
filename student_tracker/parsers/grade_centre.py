"""Parse the WSU vUWS Grade Centre export.

The file extension is .xls but the format is UTF-16 tab-separated text.
Columns:
  Last Name, First Name, Username (= student_code), Last Access, Availability,
  then a variable number of gradebook columns of the form:
      "Item Name [Total Pts: X type] |item_id"

Per the spec, ONLY columns whose Item Name starts with 'Assessment '
(case-insensitive) are treated as summative assessments. Sub-items
('Career Planning', 'Self-Assessment Survey' etc.), formative
'Complete/Incomplete' items and calculated columns ('Weighted Total',
'Total Column', 'Program Code') are excluded.
"""

import re
import pandas as pd

ASSESSMENT_RE = re.compile(r"^assessment\s", re.IGNORECASE)
RESUBMISSION_RE = re.compile(r"resubmission", re.IGNORECASE)
COLUMN_META_RE = re.compile(r"^(.*?)\s*\[Total Pts:\s*([^\]]+)\]\s*\|(\d+)\s*$")
# Values commonly used in non-numeric (Satisfactory-standard) assessments
# that nevertheless indicate a submission was made.
TEXT_SUBMITTED_VALUES = {"satisfactory", "unsatisfactory", "complete", "incomplete"}


def _parse_column_name(col: str) -> dict | None:
    """Extract item name, points/type, and item id from a gradebook column."""
    m = COLUMN_META_RE.match(col)
    if not m:
        return None
    name = m.group(1).strip()
    type_str = m.group(2).strip()
    item_id = m.group(3).strip()
    return {"name": name, "type_str": type_str, "item_id": item_id, "raw": col}


def identify_assessment_columns(columns: list[str]) -> list[dict]:
    """Return the list of summative-assessment column descriptors.

    Excludes 'Resubmission' link columns (treated as repeats of the
    original assessment, not separate items).
    """
    out = []
    for c in columns:
        meta = _parse_column_name(c)
        if not meta:
            continue
        if not ASSESSMENT_RE.match(meta["name"]):
            continue
        if RESUBMISSION_RE.search(meta["name"]):
            continue
        out.append(meta)
    return out


def parse(path: str) -> dict:
    """Parse the Grade Centre file.

    Returns a dict with:
      'students': DataFrame keyed by student_code with last_access (datetime)
                  and one column per summative assessment (raw value as
                  parsed: numeric where possible, else original string,
                  NaN where not submitted).
      'assessments': list of column descriptors {name, type_str, item_id}.
    """
    df = pd.read_csv(path, sep="\t", encoding="utf-16")

    # Drop trailing all-NaN columns (the 'Unnamed: 22' / 'Unnamed: 23' artefacts)
    df = df.dropna(axis=1, how="all")

    if "Username" not in df.columns:
        raise ValueError("Grade centre export missing 'Username' column.")

    df = df.rename(columns={"Username": "student_code"})
    df["student_code"] = df["student_code"].astype(str).str.strip()
    if "Last Access" in df.columns:
        df["last_access"] = pd.to_datetime(df["Last Access"], errors="coerce")
    else:
        df["last_access"] = pd.NaT

    assessments = identify_assessment_columns(df.columns.tolist())

    keep_cols = ["student_code", "last_access"] + [a["raw"] for a in assessments]
    students = df[keep_cols].copy()
    rename_map = {a["raw"]: a["name"] for a in assessments}
    students = students.rename(columns=rename_map)

    # Keep raw values; numeric coercion happens at score-calculation time so
    # that text values like "Satisfactory" are still detected as submissions.
    return {"students": students, "assessments": assessments}


def _is_submitted(val) -> bool:
    """A non-null value, including text like 'Satisfactory', counts as submitted."""
    if pd.isna(val):
        return False
    if isinstance(val, str):
        s = val.strip().lower()
        if not s:
            return False
        # Numeric strings count
        try:
            float(s)
            return True
        except ValueError:
            pass
        return s in TEXT_SUBMITTED_VALUES
    return True


def submission_summary(parsed: dict) -> pd.DataFrame:
    """Per-student: submitted_count, total_count, submission_rate, avg_score (%).

    Submission detection: any non-null value counts, including text values
    like 'Satisfactory' / 'Unsatisfactory' (the student did the work).
    Score calculation: only numeric values are averaged, expressed as a
    percentage of the item's max points. Text-only assessments are
    excluded from the score calculation but counted in submission rate.
    """
    students = parsed["students"]
    assessments = parsed["assessments"]
    if not assessments:
        return pd.DataFrame({
            "student_code": students["student_code"],
            "assessments_submitted": 0,
            "assessments_total": 0,
            "submission_rate": 0.0,
            "avg_score_pct": pd.NA,
        })

    max_pts: dict[str, float] = {}
    for a in assessments:
        m = re.search(r"(\d+(?:\.\d+)?)", a["type_str"])
        if m:
            try:
                max_pts[a["name"]] = float(m.group(1))
            except ValueError:
                pass

    cols = [a["name"] for a in assessments]

    # Submission flags (handles text values like Satisfactory)
    submitted = students[cols].map(_is_submitted)
    submitted_count = submitted.sum(axis=1)
    total_count = len(cols)

    # Score: only numeric values, only items with parseable max_pts
    score_cols = [c for c in cols if c in max_pts and max_pts[c] > 0]
    if score_cols:
        numeric = students[score_cols].apply(pd.to_numeric, errors="coerce")
        score_pct = numeric.copy()
        for c in score_cols:
            score_pct[c] = (numeric[c] / max_pts[c]) * 100.0
        avg_score = score_pct.mean(axis=1, skipna=True)
    else:
        avg_score = pd.Series([pd.NA] * len(students))

    return pd.DataFrame({
        "student_code": students["student_code"].values,
        "assessments_submitted": submitted_count.values,
        "assessments_total": total_count,
        "submission_rate": (submitted_count / total_count).values,
        "avg_score_pct": avg_score.values,
    })
