"""
Re-registration advice — port of Grant's "testing calculator" engine
====================================================================

Grant's calculator (``2025 SUM for 2026 AUT testing calculator``) is a
deterministic lookup engine built in Excel. This module ports its three
reference sheets and reproduces its output:

- ``rereg_data/ref_programs.json``  — from sheet ``REF``: per program, the
  subject code at each pattern position + the offering pattern per planning
  session.
- ``rereg_data/manual_text.json``   — from sheet ``Manual Text``: every
  ``(offering pattern, fail-status bitmask)`` → ``Prep / B1–B4 / Earliest
  Completion`` for diplomas. 4,450 rows, no conflicts.
- ``rereg_data/nursing_text.json``  — sheet ``Nursing Text``: the same for the
  8-subject Nursing pattern.
- ``rereg_data/summer_offering.json`` — sheet ``Summer Offering``: per subject,
  its Summer timetable block + coach comment.

Flow per student:  program + planning session → offering pattern (REF)
                   student's Subject/Prep Status → fail-status bitmask
                   (pattern, bitmask) → advice by position (Manual/Nursing Text)
                   positions → subject codes (REF)

Both of Josiah's 2026-09-02 rules are already baked into Manual/Nursing Text:
positions 1 & 2 always run in Summer, and a clash between position N and N+4
(same timetable slot) is resolved in favour of the cohort subject with the
earlier one pushed to Summer (Earliest Completion moves out).
"""

from __future__ import annotations

import json
import re
from functools import lru_cache
from pathlib import Path

import pandas as pd

_DATA = Path(__file__).with_name("rereg_data")

SHEET_NAME = "Query1"
PREP_SLOTS = ["Prep 1 Status", "Prep 2 Status"]
MODULAR_SLOTS = [f"Subject {i} Status" for i in range(1, 9)]

ADVICE_COLS = [
    "Prep Registration Advice",
    "Block 1 Registration Advice",
    "Block 2 Registration Advice",
    "Block 3 Registration Advice",
    "Block 4 Registration Advice",
]
COMPLETION_COL = "Earliest Completion"
REASON_COL = "Advice Reason"

NURSING_PROGRAMS = {"9031"}
UNSUPPORTED_PROGRAMS = {"9034"}  # Policing — not in Grant's calculator

# planning session -> the REF pattern key to use
PLANNING_SESSIONS = {
    "25 SUM": "25 SUM",
    "26 AUT": "26 AUT",
}
DEFAULT_PLANNING_SESSION = "26 AUT"

NO_REGISTRATION = "You do not need to register in a subject"


# --------------------------------------------------------------------------- #
# Reference data                                                              #
# --------------------------------------------------------------------------- #
@lru_cache(maxsize=1)
def _ref() -> dict:
    return json.loads((_DATA / "ref_programs.json").read_text())


@lru_cache(maxsize=2)
def _lookup(kind: str) -> dict:
    """kind = 'diploma' | 'nursing'. Returns {(pattern, status): row}."""
    fname = "nursing_text.json" if kind == "nursing" else "manual_text.json"
    rows = json.loads((_DATA / fname).read_text())
    return {(r["pattern"], r["status"]): r for r in rows}


@lru_cache(maxsize=1)
def _summer_offering() -> dict:
    rows = json.loads((_DATA / "summer_offering.json").read_text())
    out = {}
    for r in rows:
        if r.get("program") and r.get("position"):
            out[(str(r["program"]), r["position"])] = r
    return out


# --------------------------------------------------------------------------- #
# Status bitmask                                                              #
# --------------------------------------------------------------------------- #
def _is_outstanding(value) -> bool:
    """True when the slot still needs to be passed.

    Passed / currently-registered / blank all count as 0; the bare 1 (str or
    int), or anything that isn't a "… Completed" / "… Currently Registered"
    string, counts as outstanding.
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return False
    s = str(value).strip()
    if s == "":
        return False
    if s.endswith("Completed") or "Currently Registered" in s:
        return False
    return True  # "1", "1.0", etc.


def _elective_count(row) -> int:
    for col in ("Electives Needed", "Elective Completion"):
        if col in row and pd.notna(row[col]):
            try:
                return max(0, min(2, int(float(str(row[col]).strip()))))
            except (TypeError, ValueError):
                return 0
    return 0


def status_mask(row: pd.Series, is_nursing: bool) -> str:
    def bit(slot):
        return "1" if _is_outstanding(row.get(slot)) else "0"

    if is_nursing:
        s = [bit(f"Subject {i} Status") for i in range(1, 9)]
        return f"[{'+'.join(s[:4])}] + [{'+'.join(s[4:])}]"

    p = [bit("Prep 1 Status"), bit("Prep 2 Status")]
    s = [bit(f"Subject {i} Status") for i in range(1, 7)]
    e = _elective_count(row)
    return f"[{p[0]}+{p[1]}] + [{'+'.join(s[:4])}] + [{s[4]}+{s[5]}] + [{e}]"


# --------------------------------------------------------------------------- #
# Advice                                                                      #
# --------------------------------------------------------------------------- #
def pattern_for(program: str, planning_session: str) -> str | None:
    prog = _ref().get(str(program))
    if not prog:
        return None
    key = PLANNING_SESSIONS.get(planning_session, planning_session)
    return prog.get("patterns", {}).get(key)


def _resolve(token: str | None, program: str) -> str:
    """A Manual-Text position token -> a real cell value."""
    if not token:
        return NO_REGISTRATION
    t = token.strip()
    if t.lower() == "elective":
        return "+1 elective"
    m = re.fullmatch(r"Subject (\d)", t)
    if m:
        return _ref().get(str(program), {}).get("subjects", {}).get(m.group(1), t)
    # GEDU0016 / GEDU0017 / "GEDU0016 and GEDU0017" pass through
    return t


def advise_row(row: pd.Series, planning_session: str) -> dict:
    program = str(row["PROGRAM_CD"]).split(".")[0]
    is_nursing = program in NURSING_PROGRAMS

    out = {c: "" for c in ADVICE_COLS}
    out[COMPLETION_COL] = ""
    out[REASON_COL] = ""

    if program in UNSUPPORTED_PROGRAMS:
        out[REASON_COL] = f"Program {program} is not covered by the calculator — refer to coach."
        return out

    pattern = pattern_for(program, planning_session)
    if pattern is None:
        out[REASON_COL] = (
            f"No offering pattern for program {program} / {planning_session} "
            "in the reference data."
        )
        return out

    mask = status_mask(row, is_nursing)
    hit = _lookup("nursing" if is_nursing else "diploma").get((pattern, mask))
    if hit is None:
        out[REASON_COL] = f"No calculator row for pattern {pattern} + status {mask}."
        return out

    tokens = [hit.get("prep"), hit.get("b1"), hit.get("b2"), hit.get("b3"), hit.get("b4")]
    for col, tok in zip(ADVICE_COLS, tokens):
        out[col] = _resolve(tok, program)
    out[COMPLETION_COL] = hit.get("earliest_completion") or ""

    named = [out[c] for c in ADVICE_COLS[1:] if out[c] != NO_REGISTRATION]
    bits = []
    if out[ADVICE_COLS[0]] != NO_REGISTRATION:
        bits.append(f"Prep: {out[ADVICE_COLS[0]]}")
    if named:
        bits.append("Register: " + ", ".join(named))
    if not bits:
        bits.append("Nothing to register this session")
    if hit.get("advice_for_when"):
        bits.append(f"for {hit['advice_for_when']}")
    if out[COMPLETION_COL]:
        bits.append(f"earliest completion {out[COMPLETION_COL]}")
    status = str(row.get("STUDY_PATH_STATUS", "") or "")
    if status and status != "Active Study Path":
        bits.append(f"NOTE: {status}")
    out[REASON_COL] = " | ".join(bits)
    return out


# --------------------------------------------------------------------------- #
# Whole-file entry points                                                     #
# --------------------------------------------------------------------------- #
def load_progression_file(source) -> pd.DataFrame:
    try:
        df = pd.read_excel(source, sheet_name=SHEET_NAME)
    except ValueError as exc:
        raise ValueError(f"Couldn't find a sheet called '{SHEET_NAME}'.") from exc
    need = ["PROGRAM_CD", *PREP_SLOTS, *[f"Subject {i} Status" for i in range(1, 7)]]
    missing = [c for c in need if c not in df.columns]
    if missing:
        raise ValueError("Missing expected columns: " + ", ".join(missing))
    return df


def build_advice(df: pd.DataFrame, planning_session: str = DEFAULT_PLANNING_SESSION) -> pd.DataFrame:
    out = df.copy()
    results = [advise_row(r, planning_session) for _, r in df.iterrows()]
    for col in [*ADVICE_COLS, COMPLETION_COL, REASON_COL]:
        out[col] = [r[col] for r in results]
    return out
