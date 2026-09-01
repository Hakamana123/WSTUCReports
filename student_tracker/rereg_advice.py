"""
Re-registration advice (v2) - clean rebuild
===========================================

One progression workbook in, the same workbook back out, with five
``... Registration Advice`` columns and an ``Advice Reason`` column added
(the file's own ``... Registration`` columns - the student's current
enrolment - are left untouched for side-by-side comparison).

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

# The recommendation is written into its own columns, left of nothing and
# alongside the file's original "... Registration" columns (which hold the
# student's actual current enrolment - kept for comparison, never overwritten).
ADVICE_COLS = [
    "Prep Registration Advice",
    "Block 1 Registration Advice",
    "Block 2 Registration Advice",
    "Block 3 Registration Advice",
    "Block 4 Registration Advice",
]
REASON_COL = "Advice Reason"

BLOCKS_PER_SESSION = 4

# Programs whose "Electives Needed" value is not a real requirement. 9034
# (Policing) is a 2-subject program with no elective structure; the 2 that
# shows up in the file is a data artifact (confirmed by Josiah 2026-09-01).
NO_ELECTIVE_PROGRAMS = {"9034"}

# Progression standing -> registration rules (Josiah, 2026-09-01):
#   Conditional Enrolment: 30cp cap => 3 modular subjects max, and the prep
#     subject (15cp) moves to Summer (prep runs in every Summer block), so the
#     30cp is spent entirely on modular progress.
#   Exclusion: not eligible to re-register - no advice.
#   Good Standing / At Risk / blank (new starters): full load, no restriction.
STANDING_MAX_BLOCKS = {"Conditional Enrolment": 3}
STANDING_PREP_TO_SUMMER = {"Conditional Enrolment"}
STANDING_NO_ADVICE = {"Exclusion"}

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

    outcome = str(row.get("Progression Outcome", "") or "").strip()
    if outcome in STANDING_NO_ADVICE:
        return Advice(reason=f"{outcome} - not eligible to re-register; refer to coach.")

    max_blocks = STANDING_MAX_BLOCKS.get(outcome, BLOCKS_PER_SESSION)
    prep_to_summer = outcome in STANDING_PREP_TO_SUMMER

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
        elif len(blocks) >= max_blocks:
            rolls_over.append(code)
        else:
            blocks.append(code)
            if slot in backlog_slots:
                backlog_placed += 1

    core_count = len(blocks)

    # --- electives fill whatever block room is left, one per block cell ----
    electives_to_add = max(0, electives_needed - backlog_placed)
    room = max_blocks - len(blocks)
    add_now = min(room, electives_to_add)
    blocks.extend(["+1 elective"] * add_now)
    electives_deferred = electives_needed - add_now

    # --- prep: one prep subject per session, earliest outstanding first ---
    prep_available = [
        pmap.get(s, s) for s in outstanding_prep if pmap.get(s, s) not in prep_not_offered
    ]
    prep_summer: list[str] = []
    if prep_to_summer:
        prep_summer = prep_available
        prep_value = ""
        prep_carried = [
            pmap.get(s, s) for s in outstanding_prep if pmap.get(s, s) in prep_not_offered
        ]
    else:
        prep_value = prep_available[0] if prep_available else ""
        prep_carried = list(prep_available[1:]) + [
            pmap.get(s, s) for s in outstanding_prep if pmap.get(s, s) in prep_not_offered
        ]

    # --- reason --------------------------------------------------------
    reason_bits: list[str] = []
    core_advised = blocks[:core_count]
    if outcome in STANDING_MAX_BLOCKS:
        reason_bits.append(f"{outcome}: 30cp cap - {max_blocks} subjects max this session")
    elif outcome == "At Risk":
        reason_bits.append("At Risk (full load allowed - monitor)")
    if has_prep and prep_value:
        reason_bits.append(f"Prep: {prep_value}")
    if prep_summer:
        reason_bits.append("Prep in Summer: " + ", ".join(prep_summer))
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
    free = max_blocks - len(blocks)
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
    """Return a copy of ``df`` with the advice columns + a reason appended.

    The file's original "... Registration" columns (the student's current
    enrolment) are left untouched, so a coach can compare them side by side.
    """
    offerings = offerings or load_offerings()
    slot_map = derive_slot_map(df)

    out = df.copy()
    for col in [*ADVICE_COLS, REASON_COL]:
        out[col] = ""

    for idx, row in df.iterrows():
        advice = advise_student(row, slot_map, offerings)
        out.at[idx, ADVICE_COLS[0]] = advice.prep
        for block_col, value in zip(ADVICE_COLS[1:], advice.blocks):
            out.at[idx, block_col] = value
        out.at[idx, REASON_COL] = advice.reason

    return out


# --------------------------------------------------------------------------- #
# Coach view - a slim, readable second sheet                                   #
# --------------------------------------------------------------------------- #
PASSED_MARK = "✓"       # passed
OUTSTANDING_MARK = "✗"  # still to pass
BAR_FILLED = "█"
BAR_EMPTY = "░"
BAR_WIDTH = 10


def _bar(done: int, total: int) -> str:
    """Text progress bar, e.g. '██████░░░░ 60%'."""
    if total <= 0:
        return ""
    filled = round(done / total * BAR_WIDTH)
    return f"{BAR_FILLED * filled}{BAR_EMPTY * (BAR_WIDTH - filled)} {done / total:.0%}"


COACH_VIEW_SHEET = "Coach View"


def _status(row: pd.Series) -> tuple[str, str, str]:
    """(summary, grid, bar).

    - summary: plain count line, e.g. "Outstanding: 1 prep, 2 core, 2 electives"
    - grid: positional marks, e.g. "Prep ✓✗ | Blk 1-4 ✓✓✓✓ | Blk 5-6 ✗✗ | Elec ✗✗"
    - bar: text progress bar over all required subjects
    ✓ = passed, ✗ = still to pass. Electives always count as outstanding (the
    file only gives a count of what's still needed).
    """
    prep_used = [s for s in PREP_SLOTS if pd.notna(row[s])]
    core_used = [s for s in MODULAR_SLOTS if pd.notna(row[s])]

    prep_out = sum(_is_outstanding(row[s]) for s in prep_used)
    core_out = sum(_is_outstanding(row[s]) for s in core_used)
    elec_out = 0 if str(row["PROGRAM_CD"]) in NO_ELECTIVE_PROGRAMS \
        else _parse_elective_count(row.get("Electives Needed"))

    # summary
    parts = []
    if prep_out:
        parts.append(f"{prep_out} prep")
    if core_out:
        parts.append(f"{core_out} core")
    if elec_out:
        parts.append(f"{elec_out} elective" + ("s" if elec_out > 1 else ""))
    summary = ("Outstanding: " + ", ".join(parts)) if parts else "All passed"

    # grid
    def marks(slots):
        return "".join(OUTSTANDING_MARK if _is_outstanding(row[s]) else PASSED_MARK for s in slots)

    segs = []
    if prep_used:
        segs.append("Prep " + marks(prep_used))
    for i in range(0, len(core_used), 4):
        chunk = core_used[i:i + 4]
        lo, hi = i + 1, i + len(chunk)
        label = f"Blk {lo}-{hi}" if hi > lo else f"Blk {lo}"
        segs.append(f"{label} {marks(chunk)}")
    if elec_out:
        segs.append("Elec " + OUTSTANDING_MARK * elec_out)
    grid = " | ".join(segs)

    # bar: passed / total required (prep + core + electives-needed)
    total = len(prep_used) + len(core_used) + elec_out
    done = (len(prep_used) - prep_out) + (len(core_used) - core_out)
    return summary, grid, _bar(done, total)


COACH_VIEW_COLUMNS = [
    "STUDENT_ID", "FIRST_NAME", "LAST_NAME", "PREFERRED_NAME",
    "INSTITUTION_EMAIL_ADDRESS", "Coach", "PROGRAM_CD", "COMMENCEMENT_PERIOD",
    "Progression Outcome",
]


def build_coach_view(advised: pd.DataFrame, offerings: dict | None = None) -> pd.DataFrame:
    """Slim sheet: identity + success coach + status + the advice columns."""
    base = advised[[c for c in COACH_VIEW_COLUMNS if c in advised.columns]].copy()

    status = [_status(r) for _, r in advised.iterrows()]
    base["Student Status"] = [s for s, _, _ in status]
    base["Progress Bar"] = [b for _, _, b in status]
    base["Progress"] = [g for _, g, _ in status]

    for col in [*ADVICE_COLS, REASON_COL]:
        base[col] = advised[col]
    return base


def to_workbook_bytes(df: pd.DataFrame, coach_view: pd.DataFrame | None = None) -> bytes:
    import io

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        if coach_view is not None:
            coach_view.to_excel(writer, sheet_name=COACH_VIEW_SHEET, index=False)
        df.to_excel(writer, sheet_name=SHEET_NAME, index=False)
    return buffer.getvalue()
