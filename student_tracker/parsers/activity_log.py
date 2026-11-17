"""Parse the vUWS activity export (one row per student click).

File name pattern: vUWS_activity_<COURSE ID>_<from>_to_<to>.csv
Columns:
  Course ID, Course Name, Student Name, Student ID, Username,
  Event Date (Sydney), Event, Item Title, Content Type,
  Gradebook Category, IP Address

Event values seen so far: COURSE_ACCESS (entered the subject site),
CONTENT_ACCESS (opened an item), ANNOUNCEMENT_READ, ASSESSMENT_SUBMIT,
ASSESSMENT_DRAFT.

The export records WHEN each click happened, never how long the student
stayed. Time is estimated downstream (see activity_metrics).

Cleaning applied here:
  - Staff preview accounts ("<name>_PreviewUser", no Student ID) dropped.
  - Student ID kept as text (no float '.0' artefacts).
  - Username and IP Address dropped (always blank / not needed).
"""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

REQUIRED = [
    "Course ID", "Student Name", "Student ID",
    "Event Date (Sydney)", "Event", "Item Title", "Content Type",
]
DATE_FORMAT = "%d/%m/%Y, %H:%M:%S"
WINDOW_RE = re.compile(r"(\d{4}-\d{2}-\d{2})_to_(\d{4}-\d{2}-\d{2})")
SUBJECT_RE = re.compile(r"^([A-Z]{4}\d{4})")

# Content Type -> plain-English kind shown on the dashboard.
KIND_MAP = {
    "resource/x-bb-document": "Page / document",
    "resource/x-bb-blti-link": "External tool",
    "resource/x-plugin-scormengine": "SCORM module",
    "resource/x-bb-asmt-test-link": "Assessment submission",
    "resource/x-bb-forumlink": "Discussion forum",
}
# Kinds that open outside vUWS: the gap to the next vUWS click is NOT time
# spent on the item, so no dwell time is estimated for them.
EXTERNAL_KINDS = {"External tool", "SCORM module"}


def _kind(row_event: str, content_type) -> str:
    if row_event == "COURSE_ACCESS":
        return "Subject home"
    if row_event == "ANNOUNCEMENT_READ":
        return "Announcement"
    if isinstance(content_type, str):
        return KIND_MAP.get(content_type, content_type.replace("resource/", ""))
    return "Other"


def parse(source, filename: str | None = None) -> dict:
    """Parse one activity export.

    `source` is a path or a file-like object. `filename` (optional) is used
    to read the export window from the name when `source` is a buffer.

    Returns {'events': DataFrame, 'meta': dict}.
      events columns: subject_code, course_id, course_name, student_code,
        student_name, ts, event, item, content_type, kind, gb_category,
        is_external
      meta: subject_code, course_name, filename, window_start, window_end,
        rows_read, dropped_preview, dropped_bad_date, students
    """
    if filename is None and isinstance(source, (str, Path)):
        filename = Path(source).name
    df = pd.read_csv(source, dtype=str, encoding="utf-8-sig")
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(
            "doesn't look like a vUWS activity export - missing column(s): "
            + ", ".join(missing)
        )
    rows_read = len(df)

    sid = df["Student ID"].fillna("").str.strip().str.replace(r"\.0$", "", regex=True)
    name = df["Student Name"].fillna("").str.strip()
    preview = (sid == "") | name.str.endswith("_PreviewUser")
    df = df[~preview].copy()
    sid = sid[~preview]

    ts = pd.to_datetime(df["Event Date (Sydney)"], format=DATE_FORMAT, errors="coerce")
    bad_date = ts.isna()

    course_id = df["Course ID"].fillna("").str.strip()
    subj = course_id.str.extract(SUBJECT_RE, expand=False).fillna(course_id)

    out = pd.DataFrame({
        "subject_code": subj,
        "course_id": course_id,
        "course_name": df.get("Course Name", pd.Series("", index=df.index)).fillna(""),
        "student_code": sid,
        "student_name": df["Student Name"].str.strip(),
        "ts": ts,
        "event": df["Event"].fillna("").str.strip().str.upper(),
        "item": df["Item Title"].fillna("(untitled item)").str.strip(),
        "content_type": df["Content Type"],
        "gb_category": df.get("Gradebook Category", pd.Series(pd.NA, index=df.index)),
    })
    out = out[~bad_date]
    out["kind"] = [_kind(e, c) for e, c in zip(out["event"], out["content_type"])]
    out["is_external"] = out["kind"].isin(EXTERNAL_KINDS)
    out = out.sort_values(["subject_code", "student_code", "ts"]).reset_index(drop=True)

    window_start = window_end = None
    if filename:
        m = WINDOW_RE.search(filename)
        if m:
            window_start = pd.Timestamp(m.group(1)).date()
            window_end = pd.Timestamp(m.group(2)).date()

    codes = out["subject_code"].dropna().unique().tolist()
    meta = {
        "subject_code": codes[0] if len(codes) == 1 else ", ".join(codes),
        "course_name": out["course_name"].iloc[0] if len(out) else "",
        "filename": filename,
        "window_start": window_start,
        "window_end": window_end,
        "rows_read": rows_read,
        "dropped_preview": int(preview.sum()),
        "dropped_bad_date": int(bad_date.sum()),
        "students": int(out["student_code"].nunique()),
    }
    return {"events": out, "meta": meta}
