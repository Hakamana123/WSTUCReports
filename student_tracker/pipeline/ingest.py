"""
Stage 1 of 4 — INGEST
======================
Reads the roster file as it currently comes from Ashlee's team (the
"2026 AUT for 2026 SPR Progression Outcomes" style export) and normalizes
it into a clean internal schema the rest of the pipeline can rely on.

This stage is fully runnable today — nothing here is blocked on Grant.
"""

from __future__ import annotations
from dataclasses import dataclass
from typing import Optional
import pandas as pd

BLOCK_COLS = ["Spring Block 1", "Spring Block 2", "Spring Block 3", "Spring Block 4"]
PREP_COL = "Spring"

# UPP program codes, confirmed 2026-08-06. Diplomas are 7188-7198.
UPP_PROGRAM_CODES = {9031, 9034}

# Canonical (post-rename) columns that must be present - everything
# downstream (bucketing, pattern comparison) depends on these. Student
# name isn't here: to_student_records already treats it as optional
# (falls back to None), which is fine since it's used for display only,
# never classification.
REQUIRED_COLUMNS = [
    "student_id", "program", "calculated_standing", "commencement_period",
    "block_1", "block_2", "block_3", "block_4",
]


@dataclass
class StudentRecord:
    student_id: str
    student_name: Optional[str]
    program: int
    calculated_standing: str          # GS / AR / CE / EX, as supplied
    standing_desc: str
    commencement_period: Optional[str]  # None if missing in source data
    campus_code: Optional[str]
    study_path_status: Optional[str]
    mobile_number: Optional[str]
    block_registrations: list         # length-4 list, each item a subject code or None
    prep_registration: Optional[str]

    @property
    def is_upp(self) -> bool:
        return self.program in UPP_PROGRAM_CODES

    @property
    def has_commencement_period(self) -> bool:
        return self.commencement_period is not None and str(self.commencement_period).strip() != ""

    @property
    def blocks_filled(self) -> int:
        return sum(1 for b in self.block_registrations if b)


def load_roster(path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
    """Load the raw Excel roster into a DataFrame with normalized column names.

    Works against the '2026 AUT for 2026 SPR Progression Outcomes' schema.
    If Ashlee's team changes the export format, this is the one place to
    update column-name mapping.
    """
    df = pd.read_excel(path, sheet_name=sheet_name or 0)

    rename_map = {
        "Student ID (Mock Data)": "student_id",
        "Student ID (Mock)": "student_id",   # tolerate either file's naming
        "Student ID": "student_id",          # real (non-mock) export
        "Student Name (Mock)": "student_name",
        "Student name": "student_name",      # real (non-mock) export - note lowercase "name"
        "Program": "program",
        "Calculated Standing": "calculated_standing",
        "CALC Standing Desc": "standing_desc",
        "COMMENCEMENT_PERIOD": "commencement_period",
        "CAMP_CODE": "campus_code",
        "STUDY_PATH_STATUS": "study_path_status",
        "Mobile Number": "mobile_number",
        "Spring Block 1": "block_1",
        "Spring Block 2": "block_2",
        "Spring Block 3": "block_3",
        "Spring Block 4": "block_4",
        "Spring": "prep",
    }
    original_columns = list(df.columns)
    df = df.rename(columns=rename_map)

    # Not every export will carry every column (e.g. the ALA export also
    # has Academic Period / Part Term / GPA Outcome Message / etc. — those
    # pass through untouched rather than being dropped, in case a later
    # stage wants them). But the columns everything downstream depends on
    # must be present - silently falling back to None (e.g. "Student ID"
    # column missing → every student_id becomes the string "None") is a
    # silent-data-corruption bug, not graceful degradation, so fail loudly
    # here instead.
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(
            "Roster is missing expected column(s) after normalization: "
            f"{missing}. This file's actual headers are: {original_columns}. "
            "If this is a new/different export format, add its real header "
            "name(s) to rename_map in ingest.py."
        )
    return df


def to_student_records(df: pd.DataFrame) -> list[StudentRecord]:
    """Convert a normalized roster DataFrame into StudentRecord objects."""
    records = []
    for _, row in df.iterrows():
        commencement = row.get("commencement_period")
        commencement = None if pd.isna(commencement) else str(commencement)

        block_regs = []
        for col in ["block_1", "block_2", "block_3", "block_4"]:
            val = row.get(col)
            block_regs.append(None if pd.isna(val) else str(val))

        prep_val = row.get("prep")
        prep_val = None if pd.isna(prep_val) else str(prep_val)

        records.append(
            StudentRecord(
                student_id=str(row.get("student_id")),
                student_name=row.get("student_name") if "student_name" in df.columns else None,
                program=int(row.get("program")),
                calculated_standing=str(row.get("calculated_standing")),
                standing_desc=str(row.get("standing_desc")),
                commencement_period=commencement,
                campus_code=row.get("campus_code"),
                study_path_status=(
                    None if pd.isna(row.get("study_path_status")) else str(row.get("study_path_status"))
                ),
                mobile_number=row.get("mobile_number") if "mobile_number" in df.columns else None,
                block_registrations=block_regs,
                prep_registration=prep_val,
            )
        )
    return records
