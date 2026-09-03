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

from student_tracker import rereg_sessions as rs

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

    Only ``"… Completed"`` and blank count as done. The bare ``1``, and
    ``"… - Currently Registered"`` (not passed yet -> still needs to be in the
    plan, per Josiah 2026-09-02), both count as outstanding.
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return False
    s = str(value).strip()
    if s == "":
        return False
    if s.endswith("Completed"):
        return False
    return True  # "1", "1.0", "CODE - Currently Registered", "CODE Required"


def _outstanding_strict(value) -> bool:
    """Like ``_is_outstanding`` but a "Currently Registered" subject counts as
    done. Grant's Principle/Template classification was built this way, so the
    classifier uses this; the advice engine uses the lenient version."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return False
    s = str(value).strip()
    return s != "" and not s.endswith("Completed") and "Currently Registered" not in s


def _elective_count(row) -> int:
    """Outstanding elective count, clamped 0..2 (the calculator's mask range).

    Reads ``Electives Needed`` first (the v2 file's column), then
    ``Elective Completion`` (Grant's). Accepts a bare number, or a phrase like
    ``"1 Elective Required"``; ``"No Elective Required"`` / ``"Not Applicable"``
    / blank all mean 0.
    """
    for col in ("Electives Needed", "Elective Completion"):
        if col not in row or pd.isna(row[col]):
            continue
        s = str(row[col]).strip()
        try:
            return max(0, min(2, int(float(s))))
        except (TypeError, ValueError):
            pass
        m = re.match(r"\s*(\d+)", s)
        return max(0, min(2, int(m.group(1)))) if m else 0
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
# Rereg Principle / messaging Template                                        #
# --------------------------------------------------------------------------- #
# From docs/rereg_principles_and_templates.md (reverse-engineered from Grant's
# "2026 AB4 for 2026 SPR" file - the Autumn->Spring round, same as the file
# Josiah runs). Principle -> Template is 1:1.
_PRINCIPLE_TEMPLATE = {
    "On Pattern": "Template 1a",
    "Mostly Progressing": "Template 1c",
    "Stay with Cohort": "Template 1c",
    "Unsatisfactory progress in S1": "Template 1b",
    "Start Again": "Template 1b",
    "3+ Sessions": "Template 2",
    "Overall lack of success": "Template 3",
    "Complete": "Transition",
    "Commencing": "Commencing",
}


def start_semester(commencement) -> str:
    """``"25 - Spring Block 1"`` -> ``"25 SPR"``; anything 2024 or earlier ->
    ``"24"`` (Grant buckets all pre-25 starts together)."""
    m = re.match(r"\s*(\d{2})\s*-\s*(Autumn|Spring|Summer)", str(commencement or ""), re.I)
    if not m:
        return ""
    yy = int(m.group(1))
    if yy <= 24:
        return "24"
    return f"{yy} {m.group(2)[:3].upper()}"


# Autumn->Spring round cohort buckets (docs section "Cohort age"):
_CURRENT_COHORT = {"26 AUT", "25 SPR", "25 SUM"}
_OLD_COHORT = {"25 AUT", "24"}


def sem1_subject_count(row: pd.Series) -> int:
    """How many first-semester subjects the student still owes:
    ``(prep 1 outstanding ? 1 : 0) + (# of slots 1-4 outstanding)``
    (docs A1). Not affected by prep 2, slots 5-6, or electives."""
    p1 = 1 if _outstanding_strict(row.get("Prep 1 Status")) else 0
    s14 = sum(_outstanding_strict(row.get(f"Subject {i} Status")) for i in range(1, 5))
    return p1 + s14


def total_outstanding(row: pd.Series, is_nursing: bool) -> int:
    n = 9 if is_nursing else 7
    subj = sum(_outstanding_strict(row.get(f"Subject {i} Status")) for i in range(1, n))
    if is_nursing:
        return subj
    prep = sum(_outstanding_strict(row.get(f"Prep {i} Status")) for i in (1, 2))
    return prep + subj + _elective_count(row)


def classify(row: pd.Series, is_nursing: bool) -> tuple[str, str]:
    """``(principle, template)`` for the Autumn->Spring round (docs A3).

    Returns ``("", "")`` only when the commencement period is unreadable.
    """
    ss = start_semester(row.get("COMMENCEMENT_PERIOD"))
    if not ss:
        return "", ""

    total = total_outstanding(row, is_nursing)
    if total == 0:
        prin = "Complete"
    elif ss in _CURRENT_COHORT:
        s1 = sem1_subject_count(row)
        if s1 == 0:
            prin = "On Pattern"
        elif s1 <= 2:
            prin = "Stay with Cohort" if is_nursing else "Mostly Progressing"
        else:
            prin = "Start Again" if is_nursing else "Unsatisfactory progress in S1"
    elif ss in _OLD_COHORT:
        prin = "3+ Sessions" if total <= 6 else "Overall lack of success"
    else:  # 26 SPR or later - commenced the session being advised for (or after)
        prin = "Commencing"
    return prin, _PRINCIPLE_TEMPLATE.get(prin, "")


# --------------------------------------------------------------------------- #
# Advice                                                                      #
# --------------------------------------------------------------------------- #
def pattern_for(program: str, planning_session: str) -> tuple[str, str] | tuple[None, str]:
    """``(pattern_string, carried_from)``.

    ``carried_from`` is ``""`` when the session has its own pattern in the REF
    data, or the base session ("26 AUT" / "25 SUM") when the pattern is being
    assumed to carry forward unchanged. ``(None, reason)`` if there's nothing
    to use.
    """
    prog = _ref().get(str(program))
    if not prog:
        return None, f"program {program} not in calculator"
    patterns = prog.get("patterns", {})
    key = PLANNING_SESSIONS.get(planning_session, planning_session)
    if key in patterns:
        return patterns[key], ""
    base = rs.carry_base(planning_session)
    if base and base in patterns:
        return patterns[base], base
    return None, f"no offering pattern for program {program} / {planning_session}"


def _resolve(token: str | None, program: str) -> str:
    """A Manual-Text position token -> a real cell value."""
    if not token:
        return NO_REGISTRATION
    t = token.strip()
    prog = _ref().get(str(program), {})
    if t.lower() == "elective":
        return "+1 elective"
    if t.lower() == "both":  # owes both prep subjects
        p1, p2 = prog.get("prep1"), prog.get("prep2")
        return f"{p1} and {p2}" if p1 and p2 else t
    if t in ("GEDU0016", "Prep 1"):
        return prog.get("prep1", t)
    if t in ("GEDU0017", "Prep 2"):
        return prog.get("prep2", t)
    m = re.fullmatch(r"Subject (\d)", t)
    if m:
        return prog.get("subjects", {}).get(m.group(1), t)
    return t


def advise_row(row: pd.Series, planning_session: str) -> dict:
    """One student -> Grant's calculator advice.

    Keys: the five ``ADVICE_COLS``, ``COMPLETION_COL``, ``REASON_COL``, plus
    ``ok`` (True only when a calculator row was found — callers use this to
    decide whether to fall back to another engine) and ``miss`` (a short reason
    when ``ok`` is False).
    """
    program = str(row["PROGRAM_CD"]).split(".")[0]
    is_nursing = program in NURSING_PROGRAMS

    out = {c: "" for c in ADVICE_COLS}
    out[COMPLETION_COL] = ""
    out[REASON_COL] = ""
    out["ok"] = False
    out["miss"] = ""
    out["total_needed"] = 0
    out["carried_from"] = ""

    if program in UNSUPPORTED_PROGRAMS:
        out["miss"] = f"program {program} not in calculator"
        out[REASON_COL] = f"Program {program} is not covered by the calculator — refer to coach."
        return out

    pattern, carried_from = pattern_for(program, planning_session)
    if pattern is None:
        out["miss"] = carried_from  # holds the reason string in this case
        out[REASON_COL] = f"{carried_from} in the reference data."
        return out
    out["carried_from"] = carried_from

    mask = status_mask(row, is_nursing)
    hit = _lookup("nursing" if is_nursing else "diploma").get((pattern, mask))
    if hit is None:
        out["miss"] = f"no calculator row for {pattern} + {mask}"
        out[REASON_COL] = f"No calculator row for pattern {pattern} + status {mask}."
        return out

    out["ok"] = True
    try:
        out["total_needed"] = int(hit.get("total_needed") or 0)
    except (TypeError, ValueError):
        out["total_needed"] = 0
    tokens = [hit.get("prep"), hit.get("b1"), hit.get("b2"), hit.get("b3"), hit.get("b4")]
    for col, tok in zip(ADVICE_COLS, tokens):
        out[col] = _resolve(tok, program)
    # The calculator's Earliest Completion is anchored to Grant's source file;
    # it is only meaningful for the sessions he actually computed.
    out[COMPLETION_COL] = "" if carried_from else (hit.get("earliest_completion") or "")

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
