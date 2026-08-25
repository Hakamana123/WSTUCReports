"""
LIVE ROSTER — reads the real file Ashlee's team already sends Grant
directly (e.g. "... Autumn to Spring Re-reg - AUT End progression
Status ... .xlsx", sheet "Query1"), confirmed 2026-08-25 via a recorded
walkthrough with Ashlee to be the actual pre-Grant source: Grant builds
his own "AB4 for SPR" bracket-notation file (history.py) by adding his
judgment columns ON TOP of exactly this file ("he's adding data into his
existing data from previous" - Ashlee). This is earlier and richer than
either format previously supported here (ingest.py's Progression
Outcomes roster, history.py's AB4-for-SPR file) - it carries BOTH the
forward-looking Spring registration data AND explicit per-block/per-prep
pass-fail results for the session just finished, in one file.

Per Josiah (2026-08-25): this REPLACES both older formats going forward.
ingest.py/history.py's loaders are left as-is for anyone still holding an
old-format file, but new work should target this module. Produces BOTH
StudentRecord (ingest.py, for Stage 4/5's bucketing) and SubjectHistory
(history.py, for Stage 2/3's advice) from the SAME parsed file, so one
upload can drive every stage - the point of Josiah's 2026-08-23 ask
("if I upload ashlee's data, I would be able to do a stage 1/2/3/4/5
report").

GRADE CODE MAPPING (confirmed by Josiah 2026-08-25):
    Pass: H, D, C, C+, B, A, P (ordinary grade tiers), E (Exempt -
          requirement satisfied)
    Fail: F (Fail), W (Withdrawn - "for all intents and purposes behave
          like a fail"), FNS (Failed Non-Submit)
Any other code is unrecognized - that student is skipped from whichever
output needs the unresolvable field (see to_student_records/
to_history_records), not guessed.

CUMULATIVE vs THIS-SESSION - important distinction, flagged before this
module was written (2026-08-25 advisor review) and confirmed against the
real file:
  - "Subject 1-8 Status" tracks LIFETIME completion against the
    student's full pattern position (1-8), built from the university's
    grad003 report cross-referenced against Ashlee's program mapping.
    For a continuing student, position 3 there might be something
    completed two semesters ago, NOT "this session's third block" - so
    it is NOT used here for sem1_passed.
  - "Block 1-4 Result"/"Block 1-4 code" are explicitly THIS SESSION's
    four blocks, regardless of which pattern position they happen to be
    for that student - confirmed against real data to be populated the
    same way for both Diploma and Nursing/UPP programs. This is what
    sem1_passed is actually built from.
  - "Prep 1/2 Status" IS safe to use directly for prep_passed, unlike
    Subject N Status - there are only ever two prep subjects total
    (GEDU0016, GEDU0017) for the whole diploma, so "cumulative
    completion" and "has this specific prep subject been passed" are
    the same question for preps, unlike for block positions. Confirmed
    blank for every real Nursing/UPP row (9031/9034 have no prep
    subjects at all, matching history.py's existing Nursing/UPP shape).

sem2_passed is always set to an all-False placeholder of the right
length (2 for diploma, 4 for Nursing/UPP) - this file only ever reports
the session just finished, so "next session's blocks" are always
not-yet-due, exactly like the old bracket-notation file's Sem2 group
before Sem2 runs (see history.py's docstring on why that's "not yet
due", not "failed"). Nothing in history.py's suggested_*/expected_*
functions reads sem2_passed's values, only sem1_passed's - its length is
set purely to satisfy SubjectHistory's documented shape convention.

rereg_principle/template/block_advice/prep_advice are always None here -
this file is upstream of Grant's judgment layer, by design (see module
docstring above). The Stage 2/3 page already has a fallback path for
exactly this ("Suggested Rereg Principle"/"Suggested B1-4/Prep advice"
when Grant's own columns are absent) - built before this module existed,
and it's exactly what this file's output triggers.
"""

from __future__ import annotations
from typing import Optional

import pandas as pd

from .ingest import StudentRecord, UPP_PROGRAM_CODES
from .history import SubjectHistory

PASS_GRADES = {"H", "D", "C", "C+", "B", "A", "P", "E"}
FAIL_GRADES = {"F", "W", "FNS"}

STANDING_MAP = {
    "Good Standing": "GS",
    "At Risk": "AR",
    "Conditional Enrolment": "CE",
    "Exclusion": "EX",
}

REQUIRED_COLUMNS = [
    "student_id", "program", "commencement_period", "progression_outcome",
]

_RENAME_MAP = {
    "STUDENT_ID": "student_id",
    "FIRST_NAME": "first_name",
    "LAST_NAME": "last_name",
    "PREFERRED_NAME": "preferred_name",
    "PROGRAM_CD": "program",
    "COMMENCEMENT_PERIOD": "commencement_period",
    "CAMP_CODE": "campus_code",
    "STUDY_PATH_STATUS": "study_path_status",
    "Prep 1 Status": "prep1_status",
    "Prep 2 Status": "prep2_status",
    "Electives Needed": "electives_needed",
    "Block 1 Result": "block1_result",
    "Block 2 Result": "block2_result",
    "Block 3 Result": "block3_result",
    "Block 4 Result": "block4_result",
    "Prep Registration": "prep_registration",
    "Block 1 Registration": "block1_registration",
    "Block 2 Registration": "block2_registration",
    "Block 3 Registration": "block3_registration",
    "Block 4 Registration": "block4_registration",
    "Progression Outcome": "progression_outcome",
}


def load_live_roster(path, sheet_name=None) -> pd.DataFrame:
    """Load the raw live-roster workbook into a DataFrame with normalized
    column names. Defaults to the first sheet (the real file has exactly
    one, "Query1") - pass sheet_name explicitly for anything else.
    """
    df = pd.read_excel(path, sheet_name=sheet_name if sheet_name is not None else 0)
    original_columns = list(df.columns)
    df = df.rename(columns=_RENAME_MAP)
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(
            "Live roster is missing expected column(s) after "
            f"normalization: {missing}. This file's actual headers are: "
            f"{original_columns}."
        )
    return df


def _is_completed(value) -> bool:
    return isinstance(value, str) and "Completed" in value


def _grade_passed(value) -> Optional[bool]:
    """True/False for a recognized grade code, None for missing or an
    unrecognized code - callers skip rather than guess on None.
    """
    if pd.isna(value):
        return None
    v = str(value).strip()
    if v in PASS_GRADES:
        return True
    if v in FAIL_GRADES:
        return False
    return None


def _student_name(row) -> Optional[str]:
    preferred = row.get("preferred_name")
    first = row.get("first_name")
    last = row.get("last_name")
    given = preferred if not pd.isna(preferred) else first
    parts = [str(p) for p in (given, last) if not pd.isna(p)]
    return " ".join(parts) if parts else None


def to_student_records(df: pd.DataFrame) -> tuple:
    """StudentRecord list for Stage 4/5 (ingest/bucketing/report_builder),
    plus a count of rows skipped for a blank/non-numeric Program.
    Returns (records, skipped_count).
    """
    records = []
    skipped = 0
    for _, row in df.iterrows():
        program_val = row.get("program")
        if pd.isna(program_val):
            skipped += 1
            continue
        try:
            program = int(program_val)
        except (TypeError, ValueError):
            skipped += 1
            continue

        commencement = row.get("commencement_period")
        commencement = None if pd.isna(commencement) else str(commencement)

        standing_desc = row.get("progression_outcome")
        standing_desc = None if pd.isna(standing_desc) else str(standing_desc)
        calculated_standing = STANDING_MAP.get(standing_desc, standing_desc)

        block_registrations = [
            None if pd.isna(row.get(col)) else str(row.get(col))
            for col in [
                "block1_registration", "block2_registration",
                "block3_registration", "block4_registration",
            ]
        ]
        prep_reg = row.get("prep_registration")
        prep_registration = None if pd.isna(prep_reg) else str(prep_reg)

        records.append(
            StudentRecord(
                student_id=str(row.get("student_id")),
                student_name=_student_name(row),
                program=program,
                calculated_standing=calculated_standing,
                standing_desc=standing_desc,
                commencement_period=commencement,
                campus_code=(
                    None if pd.isna(row.get("campus_code")) else str(row.get("campus_code"))
                ),
                study_path_status=(
                    None if pd.isna(row.get("study_path_status")) else str(row.get("study_path_status"))
                ),
                mobile_number=None,   # this file doesn't carry a mobile number column at all
                block_registrations=block_registrations,
                prep_registration=prep_registration,
            )
        )
    return records, skipped


def to_history_records(df: pd.DataFrame) -> tuple:
    """SubjectHistory list for Stage 2/3 (history.py's suggested_*
    functions), plus a count of rows skipped for a blank/non-numeric
    Program or an unresolvable Block 1-4 Result (missing or a grade code
    not in PASS_GRADES/FAIL_GRADES - e.g. a Leave of Absence student with
    no session results at all). Returns (records, skipped_count).

    A student skipped here may still appear in to_student_records's
    output - the two are independent passes, since Stage 4/5 doesn't
    need block-result data at all.
    """
    records = []
    skipped = 0
    for _, row in df.iterrows():
        program_val = row.get("program")
        if pd.isna(program_val):
            skipped += 1
            continue
        try:
            program = int(program_val)
        except (TypeError, ValueError):
            skipped += 1
            continue

        sem1_results = [
            _grade_passed(row.get(col))
            for col in ["block1_result", "block2_result", "block3_result", "block4_result"]
        ]
        if any(r is None for r in sem1_results):
            skipped += 1
            continue
        sem1_passed = sem1_results  # now all bool, no None left

        is_upp = program in UPP_PROGRAM_CODES
        if is_upp:
            prep_passed = []
            sem2_passed = [False, False, False, False]
        else:
            prep_passed = [
                _is_completed(row.get("prep1_status")),
                _is_completed(row.get("prep2_status")),
            ]
            sem2_passed = [False, False]

        electives_val = row.get("electives_needed")
        try:
            electives_outstanding = None if pd.isna(electives_val) else int(electives_val)
        except (TypeError, ValueError):
            electives_outstanding = None   # e.g. "Not Applicable" / "No Elective Required"

        commencement = row.get("commencement_period")
        commencement = None if pd.isna(commencement) else str(commencement)

        records.append(
            SubjectHistory(
                student_id=str(row.get("student_id")),
                student_name=_student_name(row),
                program=program,
                commencement_period=commencement,
                prep_passed=prep_passed,
                sem1_passed=sem1_passed,
                sem2_passed=sem2_passed,
                electives_outstanding=electives_outstanding,
                rereg_principle=None,
                template=None,
                block_advice=[None, None, None, None],
                prep_advice=None,
            )
        )
    return records, skipped
