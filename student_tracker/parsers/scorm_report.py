"""Parse a WSU vUWS SCORM module report (exported as PDF).

OPTIONAL / secondary feature — one file per module (e.g. Module 1.1,
Module 1.2, ...). Does not feed the weekly hours/logins/forums
segmentation engine at all; it's a separate, standalone completion view.

The PDF text is genuinely inconsistent per student, which caps how
reliable this parser can be:
  - Students with zero attempts get NO time/status/score row at all —
    just a bare list of learning-object names.
  - Students who completed the package get a text banner ("This SCORM
    package was completed by...") but the Total Time / Status / Score
    columns are often absent from their row too.
  - Students with an in-progress or failed attempt get the full row
    (Total Time, Status, Scaled Score).
  - Preview/test accounts ("..._PreviewUser") are excluded.

Given that, treat 'total_time_seconds' and 'status' as best-effort —
present when the report shows them, NaN/blank when it doesn't (which is
common for zero-attempt students, so a blank there most often just means
"never opened it", not a parsing failure).
"""

import re
import pandas as pd

USER_RE = re.compile(r"^User:\s*(.+)$")
COMPLETED_RE = re.compile(r"This SCORM package was completed by")
GRADE_RE = re.compile(r"^Grade:\s*(--|[\d.]+)\s*out of\s*([\d.]+)")
DURATION_ROW_RE = re.compile(
    r"^(.*?)\s+"
    r"((?:\d+\s*hours?,?\s*)?(?:\d+\s*minutes?,?\s*)?[\d.]+\s*seconds)\s+"
    r"(complete|incomplete|passed|failed|browsed)\s+"
    r"([\d.]+)%"
)
DURATION_PARTS_RE = re.compile(
    r"(?:(\d+)\s*hours?)?,?\s*(?:(\d+)\s*minutes?)?,?\s*(?:([\d.]+)\s*seconds)?"
)


def _duration_to_seconds(s: str) -> float:
    m = DURATION_PARTS_RE.search(s)
    if not m:
        return 0.0
    hours = float(m.group(1) or 0)
    minutes = float(m.group(2) or 0)
    seconds = float(m.group(3) or 0)
    return hours * 3600 + minutes * 60 + seconds


def _extract_text(path: str) -> str:
    import pdfplumber
    parts = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            parts.append(t)
    return "\n".join(parts)


def _split_student_blocks(text: str) -> list[tuple[str, list[str]]]:
    """Split the report text into (user_line_value, block_lines) per student."""
    lines = text.splitlines()
    blocks: list[tuple[str, list[str]]] = []
    current_user = None
    current_lines: list[str] = []
    for line in lines:
        m = USER_RE.match(line.strip())
        if m:
            if current_user is not None:
                blocks.append((current_user, current_lines))
            current_user = m.group(1).strip()
            current_lines = []
        else:
            current_lines.append(line)
    if current_user is not None:
        blocks.append((current_user, current_lines))
    return blocks


def _parse_name_and_id(user_val: str) -> tuple[str, str, str]:
    """'Surname, First (ID)' or 'Surname, First' -> (surname, first, student_code)."""
    student_code = ""
    m_id = re.search(r"\((\w+)\)\s*$", user_val)
    core = user_val
    if m_id:
        student_code = m_id.group(1)
        core = user_val[:m_id.start()].strip()
    if "," in core:
        surname, first = core.split(",", 1)
        return surname.strip(), first.strip(), student_code
    return core.strip(), "", student_code


def parse(path: str, module_title: str | None = None) -> pd.DataFrame:
    """Return one row per student: surname, first_name, student_code (if
    present in the User line — often absent, since this report identifies
    students by name only), completed, grade, total_time_seconds, status,
    quiz_lines_seen (rough proxy for attempt activity — counts distinct
    quiz-item lines associated with the student, NOT a reliable attempt
    count on its own).
    """
    text = _extract_text(path)
    blocks = _split_student_blocks(text)

    records = []
    for user_val, lines in blocks:
        if "PreviewUser" in user_val:
            continue
        surname, first, student_code = _parse_name_and_id(user_val)
        block_text = "\n".join(lines)

        completed = bool(COMPLETED_RE.search(block_text))

        grade = None
        for line in lines:
            gm = GRADE_RE.match(line.strip())
            if gm:
                grade = None if gm.group(1) == "--" else float(gm.group(1))
                break

        total_time_seconds = None
        status = None
        for line in lines:
            dm = DURATION_ROW_RE.match(line.strip())
            if dm:
                total_time_seconds = _duration_to_seconds(dm.group(2))
                status = dm.group(3)
                break

        quiz_lines_seen = sum(1 for l in lines if l.strip().startswith("Module_Quiz_"))

        records.append({
            "surname": surname,
            "first_name": first,
            "student_code": student_code,
            "module": module_title or "",
            "completed": completed,
            "grade": grade,
            "total_time_seconds": total_time_seconds,
            "status": status,
            "quiz_lines_seen": quiz_lines_seen,
        })

    return pd.DataFrame(records)
