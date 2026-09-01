"""
Re-registration advice (v2) - clean rebuild
===========================================

One progression workbook in, the same workbook back out, with the five
registration columns refilled with a recommendation and an ``Advice Reason``
column added.

The whole thing is three steps (Josiah's framing, 2026-09-01):

  1. Read the file.
  2. For each student, list the subjects still to pass (the ``1`` cells,
     turned back into subject codes) plus the outstanding elective count.
  3. Advise which subjects to take next session, from their commencement
     cohort and pattern of study.

Design notes live in docs/rereg_advice_v2_spec.md. No pattern_table.json, no
live-roster merge, no history parsing - everything needed is in the one
workbook, except the Spring offering list (student_tracker/spring_offerings.json,
currently seeded only for Nursing; see that file).
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

SHEET_NAME = "Query1"

PREP_SLOTS = ["Prep 1 Status", "Prep 2 Status"]
MODULAR_SLOTS = [f"Subject {i} Status" for i in range(1, 9)]
ALL_SLOTS = PREP_SLOTS + MODULAR_SLOTS

REG_COLS = [
    "Prep Registration",
    "Block 1 Registration",
    "Block 2 Registration",
    "Block 3 Registration",
    "Block 4 Registration",
]
REASON_COL = "Advice Reason"

BLOCKS_PER_SESSION = 4

# Programs whose "Electives Needed" value is not a real requirement. 9034
# (Policing) is a 2-subject program with no elective structure; the 2 that
# shows up in the file is a data artifact (confirmed by Josiah 2026-09-01).
NO_ELECTIVE_PROGRAMS = {"9034"}

# Commencement period -> pattern slot the student's cohort reaches THIS Spring.
# Only mappings checked against real data in the sample are listed; every other
# commencement string falls through to "backlog mode" (start at slot 1, clear
# earliest-outstanding first) and is flagged for coach review.
COHORT_START_SLOT = {
    "26 - Autumn Block 1": 5,
    "26 - Autumn Block 3": 3,
    "26 - Spring Block 1": 1,
    "26 - Spring": 1,
}

_OFFERINGS_PATH = Path(__file__).with_name("spring_offerings.json")


# --------------------------------------------------------------------------- #
# Loading                                                                     #
# --------------------------------------------------------------------------- #
def load_offerings(path: Path | None = None) -> dict:
    """Load the Spring offering rules. Missing file => everything offered."""
    path = path or _OFFERINGS_PATH
    try:
        raw = json.loads(Path(path).read_text())
    except FileNotFoundError:
        return {"programs": {}, "prep_not_offered": []}
    return {
        "programs": raw.get("programs", {}),
        "prep_not_offered": set(raw.get("prep_not_offered", [])),
    }


def load_progression_file(source) -> pd.DataFrame:
    """Read the progression workbook (sheet 'Query1', header on row 1)."""
    try:
        df = pd.read_excel(source, sheet_name=SHEET_NAME)
    except ValueError as exc:  # sheet not found
        raise ValueError(
            f"Couldn't find a sheet called '{SHEET_NAME}' in this workbook."
        ) from exc

    missing = [c for c in ["COMMENCEMENT_PERIOD", "PROGRAM_CD", *ALL_SLOTS] if c not in df.columns]
    if missing:
        raise ValueError("This file is missing expected columns: " + ", ".join(missing))
    return df


# --------------------------------------------------------------------------- #
# Deriving the per-program slot -> subject-code map from the file itself       #
# --------------------------------------------------------------------------- #
def _code_if_completed(value) -> str | None:
    """'GEDU1001 Completed' -> 'GEDU1001'. Anything else (incl. the bare 1) -> None."""
    if isinstance(value, str) and value.strip().endswith("Completed"):
        return value.split()[0]
    return None


def derive_slot_map(df: pd.DataFrame) -> dict[str, dict[str, str]]:
    """
    {program_cd -> {slot_label -> subject_code}}.

    For each program and slot, the canonical subject is whatever code classmates
    show as '<CODE> Completed' in that slot (the modal value, to shrug off typos).
    """
    slot_map: dict[str, dict[str, str]] = {}
    for program, group in df.groupby("PROGRAM_CD"):
        program = str(program)
        slot_map[program] = {}
        for slot in ALL_SLOTS:
            codes = group[slot].map(_code_if_completed).dropna()
            if not codes.empty:
                slot_map[program][slot] = codes.mode().iat[0]
    return slot_map


# --------------------------------------------------------------------------- #
# Per-student advice                                                          #
# --------------------------------------------------------------------------- #
def _is_outstanding(value) -> bool:
    """A slot is outstanding when it holds the bare 1 (str or int), not a code."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return False
    return str(value).strip() == "1"


def _parse_elective_count(value) -> int:
    try:
        return max(0, int(float(str(value).strip())))
    except (TypeError, ValueError):
        return 0  # "Not Applicable", "No Elective Required", blank


def _slot_number(slot_label: str) -> int:
    """'Subject 3 Status' -> 3."""
    return int(slot_label.split()[1])


@dataclass
class Advice:
    prep: str = ""
    blocks: list[str] = field(default_factory=lambda: ["", "", "", ""])
    reason: str = ""


def advise_student(row: pd.Series, slot_map: dict, offerings: dict) -> Advice:
    program = str(row["PROGRAM_CD"])
    pmap = slot_map.get(program, {})
    prog_offer = offerings.get("programs", {}).get(program, {})
    not_offered_slots = set(prog_offer.get("not_offered_slots", []))
    prep_not_offered = offerings.get("prep_not_offered", set())

    def offered(slot_label: str) -> bool:
        return _slot_number(slot_label) not in not_offered_slots

    # --- program shape -------------------------------------------------------
    has_prep = any(pd.notna(row[s]) for s in PREP_SLOTS)
    modular_used = [s for s in MODULAR_SLOTS if pd.notna(row[s])]

    # --- what's still to pass ----------------------------------------------
    outstanding_prep = [s for s in PREP_SLOTS if _is_outstanding(row[s])]
    outstanding_mod = [s for s in modular_used if _is_outstanding(row[s])]

    notes: list[str] = []

    status = str(row.get("STUDY_PATH_STATUS", "") or "")
    if status and status != "Active Study Path":
        notes.append(f"{status} - confirm the student is returning before acting")
    outcome = str(row.get("Progression Outcome", "") or "")
    if outcome == "Exclusion":
        notes.append("Progression outcome is Exclusion - check eligibility first")

    electives_needed = _parse_elective_count(row.get("Electives Needed"))
    if program in NO_ELECTIVE_PROGRAMS:
        electives_needed = 0

    if not outstanding_prep and not outstanding_mod and electives_needed == 0:
        return Advice(reason="Nothing outstanding - no re-registration needed; confirm completion.")

    # --- cohort clock ------------------------------------------------------
    # Students who started in 2025 or earlier are past their expected finish, so
    # "clear outstanding subjects earliest-first" is exactly right for them and
    # needs no caveat. For 2026 starters we place them at their cohort's slot;
    # an unrecognised 2026 commencement is flagged.
    commencement = str(row["COMMENCEMENT_PERIOD"]).strip()
    year_token = commencement.split(" - ")[0].strip()
    start_slot = COHORT_START_SLOT.get(commencement)
    if start_slot is None:
        start_slot = 1
        if year_token.isdigit() and int(year_token) < 26:
            notes.append("continuing student - plan clears outstanding subjects earliest-first")
        else:
            notes.append(
                f"commencement '{commencement}' not recognised - planned "
                "earliest-outstanding-first, please sanity-check"
            )

    # --- order the outstanding modular subjects: cohort first, then backlog
    cohort_slots = [s for s in outstanding_mod if _slot_number(s) >= start_slot]
    backlog_slots = [s for s in outstanding_mod if _slot_number(s) < start_slot]
    ordered = cohort_slots + backlog_slots

    # --- fill the four block slots ---------------------------------------
    blocks: list[str] = []
    not_offered: list[str] = []   # a constraint - flag it
    rolls_over: list[str] = []    # just more than one session's load - normal
    backlog_placed = 0
    for slot in ordered:
        code = pmap.get(slot, slot)
        if not offered(slot):
            not_offered.append(code)
        elif len(blocks) >= BLOCKS_PER_SESSION:
            rolls_over.append(code)
        else:
            blocks.append(code)
            if slot in backlog_slots:
                backlog_placed += 1

    core_count = len(blocks)

    # --- electives fill whatever block room is left, one per block cell ----
    electives_to_add = max(0, electives_needed - backlog_placed)
    room = BLOCKS_PER_SESSION - len(blocks)
    add_now = min(room, electives_to_add)
    blocks.extend(["+1 elective"] * add_now)
    electives_deferred = electives_needed - add_now

    # --- prep: one prep subject per session, earliest outstanding first ---
    prep_available = [
        pmap.get(s, s) for s in outstanding_prep if pmap.get(s, s) not in prep_not_offered
    ]
    prep_value = prep_available[0] if prep_available else ""
    prep_carried = [c for c in prep_available[1:]] + [
        pmap.get(s, s) for s in outstanding_prep if pmap.get(s, s) in prep_not_offered
    ]

    # --- reason --------------------------------------------------------
    reason_bits: list[str] = []
    core_advised = blocks[:core_count]
    if has_prep and prep_value:
        reason_bits.append(f"Prep: {prep_value}")
    if core_advised:
        reason_bits.append("Subjects: " + ", ".join(core_advised))
    if add_now:
        reason_bits.append(f"+{add_now} elective(s)")
    if electives_deferred:
        if backlog_placed:
            reason_bits.append(
                f"{electives_deferred} elective(s) pushed back - blocks used to re-take "
                + ", ".join(blocks[core_count - backlog_placed:core_count])
            )
        else:
            reason_bits.append(f"{electives_deferred} elective(s) roll to the following session")
    if rolls_over:
        reason_bits.append("Following session: " + ", ".join(rolls_over))
    if not_offered:
        reason_bits.append("Not offered in Spring, carry: " + ", ".join(not_offered))
    if prep_carried:
        reason_bits.append("Prep for a later session: " + ", ".join(prep_carried))
    free = BLOCKS_PER_SESSION - len(blocks)
    if free and not electives_deferred:
        reason_bits.append(f"Light load - {free} block(s) free, coach to place")
    if notes:
        reason_bits.append("NOTE: " + "; ".join(notes))

    padded = (blocks + ["", "", "", ""])[:BLOCKS_PER_SESSION]
    return Advice(prep=prep_value, blocks=padded, reason=" | ".join(reason_bits))


# --------------------------------------------------------------------------- #
# Whole-file entry point                                                      #
# --------------------------------------------------------------------------- #
def build_advice(df: pd.DataFrame, offerings: dict | None = None) -> pd.DataFrame:
    """Return a copy of ``df`` with the registration columns refilled + a reason."""
    offerings = offerings or load_offerings()
    slot_map = derive_slot_map(df)

    out = df.copy()
    for col in REG_COLS:
        if col not in out.columns:
            out[col] = ""
    out[REASON_COL] = ""

    for idx, row in df.iterrows():
        advice = advise_student(row, slot_map, offerings)
        out.at[idx, REG_COLS[0]] = advice.prep
        for block_col, value in zip(REG_COLS[1:], advice.blocks):
            out.at[idx, block_col] = value
        out.at[idx, REASON_COL] = advice.reason

    return out


def to_workbook_bytes(df: pd.DataFrame) -> bytes:
    import io

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=SHEET_NAME, index=False)
    return buffer.getvalue()
