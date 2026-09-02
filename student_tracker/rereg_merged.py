"""
Re-registration advice — Grant's calculator picks the subjects, v2 wraps it
==========================================================================

This is the v2 tool (student_tracker/rereg_advice) with Grant's testing
calculator (student_tracker/rereg_calc) slotted in as the subject picker.
v2 is not replaced — its progression-standing rules, its Coach View, and its
rule-tree all still run; the calculator just supplies the raw Prep / Block
1-4 picks where it has them.

Per student:

  1. Exclusion  ->  no advice (v2 rule), whichever engine.
  2. Ask Grant's calculator for this planning session's Prep / Block 1-4
     picks + Earliest Completion. The calculator already handles the
     per-session offering pattern (e.g. 7197 not running position 4 in
     26 AUT), the timetable clash between position N and N+4 (cohort
     subject kept, the other pushed to Summer), the "positions 1 & 2 run
     every Summer" rule, and elective placement.
  3. Apply v2's standing wrapper to the calculator's ordered picks:
       - Conditional Enrolment: 30cp cap -> keep the first 3 block picks,
         defer the rest; the prep subject (15cp) moves to Summer.
       - At Risk / Good Standing / blank: full load, all 4 picks kept.
  4. If the calculator has no row for this student — a session it doesn't
     cover (Spring, and anything past 26 AUT / 25 SUM), an unported program
     (9034 Policing), or a status combination Grant's own tool returns
     "Not Found" for — fall back to the v2 rule-tree and say so in the
     'Advice Source' column.

Output: the same workbook back with 'Block N Registration Advice',
'Earliest Completion', 'Advice Reason' and 'Advice Source' columns added,
plus the readable 'Coach View' sheet from v2.
"""

from __future__ import annotations

import pandas as pd

from student_tracker import rereg_advice as v2
from student_tracker import rereg_calc as calc
from student_tracker import rereg_sessions as rs

# Same five advice columns as v2, plus two of our own.
ADVICE_COLS = v2.ADVICE_COLS
COMPLETION_COL = "Earliest Completion"
REASON_COL = v2.REASON_COL           # "Advice Reason"
SOURCE_COL = "Advice Source"

# Target sessions the page offers, in cycle order. Grant's calculator has
# offering patterns for 26 AUT + 25 SUM; every other target runs the v2
# rule-tree with the cohort clock advanced to that session.
PLANNING_SESSIONS = rs.NAMED_TARGETS
DEFAULT_PLANNING_SESSION = rs.DEFAULT_TARGET

_SRC_CALC = "Grant calculator"
_SRC_CALC_CE = "Grant calculator + CE 30cp cap"
_SRC_V2 = "v2 rule-tree ({})"
_SRC_EXCLUDED = "standing: Exclusion (no advice)"


# (early position, late position, block index) - the two blocks a diploma
# student can have a position clash in.
_SUMMER_PAIRS = ((1, 5, 0), (2, 6, 1))


def _summer_swap(blocks: list[str], row: pd.Series, program: str) -> tuple[list[str], list[str]]:
    """Positions 1 & 2 always run in Summer; 5 & 6 may not (Josiah 2026-09-02).

    So when a position 5/6 subject is outstanding, do it now in its block
    (1 or 2); if the position 1/2 subject would otherwise sit there, send that
    one to Summer instead. Diploma programs only - nursing has no Summer prep
    cycle, 9034 has no pattern. Anything else bumped out of the block (an
    elective) is picked up by the caller's elective count reconciliation.

    Returns ``(blocks, to_summer)``.
    """
    subj = calc._ref().get(program, {}).get("subjects", {})
    if not subj:
        return blocks, []

    def outstanding(pos: int) -> bool:
        return calc._is_outstanding(row.get(f"Subject {pos} Status"))

    to_summer: list[str] = []
    for early, late, blk in _SUMMER_PAIRS:
        late_code = subj.get(str(late))
        early_code = subj.get(str(early))
        if not late_code or not outstanding(late):
            continue
        if blocks[blk] != late_code:
            blocks[blk] = late_code
        if early_code and outstanding(early) and early_code not in blocks:
            to_summer.append(early_code)
    return blocks, to_summer


def _apply_cap(blocks: list[str], max_blocks: int) -> tuple[list[str], list[str]]:
    """Keep the first ``max_blocks`` non-empty picks **in their block
    positions**; blank the later ones and return them as deferred.

    The calculator positions each pick in the timetable block the subject
    actually runs in (a failed Subject 3 goes in Block 3, not Block 1), so the
    positions must be preserved - only the count is capped.
    """
    kept = list(blocks)
    deferred: list[str] = []
    seen = 0
    for i, b in enumerate(kept):
        if not b:
            continue
        seen += 1
        if seen > max_blocks:
            deferred.append(b)
            kept[i] = ""
    return kept, deferred


def advise_student_merged(row: pd.Series, slot_map: dict, offerings: dict, session: str) -> dict:
    program = str(row["PROGRAM_CD"]).split(".")[0]
    outcome = str(row.get("Progression Outcome", "") or "").strip()

    out = {c: "" for c in ADVICE_COLS}
    out[COMPLETION_COL] = ""
    out[REASON_COL] = ""
    out[SOURCE_COL] = ""

    # 1. Exclusion -> no advice, whichever engine would have run.
    if outcome in v2.STANDING_NO_ADVICE:
        out[REASON_COL] = f"{outcome} - not eligible to re-register; refer to coach."
        out[SOURCE_COL] = _SRC_EXCLUDED
        return out

    # 2. Grant's calculator - only for the sessions it has offering patterns for.
    if rs.uses_calculator(session):
        c = calc.advise_row(row, session)
    else:
        c = {"ok": False, "miss": f"{session} not covered by the calculator"}

    if not c["ok"]:
        # 4. Fall back to the v2 rule-tree (cohort clock -> this target),
        #    tagged with why the calculator didn't run.
        adv = v2.advise_student(row, slot_map, offerings, target=session)
        out[ADVICE_COLS[0]] = adv.prep
        for col, val in zip(ADVICE_COLS[1:], adv.blocks):
            out[col] = val
        out[REASON_COL] = adv.reason
        out[SOURCE_COL] = _SRC_V2.format(c["miss"])
        return out

    # 3. Calculator answered. Keep its block POSITIONS; only cap the count.
    prep_pick = c[ADVICE_COLS[0]]
    prep_pick = "" if prep_pick == calc.NO_REGISTRATION else prep_pick
    positioned = ["" if c[col] == calc.NO_REGISTRATION else c[col] for col in ADVICE_COLS[1:]]

    # Positions 1 & 2 always run in Summer; if 5/6 are outstanding, do them now
    # and push 1/2 to Summer. Skip when the target session *is* Summer.
    to_summer: list[str] = []
    tgt = rs.parse_target(session)
    if not (tgt and tgt[1] == "SUM"):
        positioned, to_summer = _summer_swap(positioned, row, program)

    max_blocks = v2.STANDING_MAX_BLOCKS.get(outcome, v2.BLOCKS_PER_SESSION)
    prep_to_summer = outcome in v2.STANDING_PREP_TO_SUMMER
    capped = outcome in v2.STANDING_MAX_BLOCKS

    kept, deferred = _apply_cap(positioned, max_blocks) if capped else (positioned, [])
    named = [(i + 1, b) for i, b in enumerate(kept) if b]

    # elective bookkeeping: what's needed vs what ended up in a block
    elec_need = calc._elective_count(row)
    elec_now = kept.count("+1 elective")
    subj_deferred = [d for d in deferred if d != "+1 elective"]
    elec_later = max(0, elec_need - elec_now)
    prep_summer = prep_pick if (prep_pick and prep_to_summer) else ""
    prep_now = "" if prep_to_summer else prep_pick

    out[ADVICE_COLS[0]] = prep_now
    for col, val in zip(ADVICE_COLS[1:], kept):
        out[col] = val
    completion = c[COMPLETION_COL]
    out[COMPLETION_COL] = "" if completion in ("", "Not Found") else completion

    nothing = not named and not prep_now and not prep_summer and not to_summer

    bits: list[str] = []
    if capped and not nothing:
        bits.append(f"{outcome}: 30cp cap - {max_blocks} subjects max this session")
    elif outcome == "At Risk" and not nothing:
        bits.append("At Risk (full load allowed - monitor)")
    if prep_now:
        bits.append(f"Prep: {prep_now}")
    if prep_summer:
        bits.append(f"Prep in Summer: {prep_summer}")
    if named:
        bits.append("Register: " + ", ".join(f"Block {n} {b}" for n, b in named))
    if to_summer:
        bits.append("Take in Summer (always offered): " + ", ".join(to_summer))
    if subj_deferred:
        bits.append("Defer to a later session: " + ", ".join(subj_deferred))
    if elec_later:
        bits.append(f"{elec_later} elective(s) to a later session")
    if nothing and c.get("total_needed", 0) > 0:
        bits.append(
            f"Nothing to register in {session} - {c['total_needed']} subject(s) "
            "still to pass are not offered this session"
        )
    elif nothing:
        bits.append("Nothing to register this session")
    if out[COMPLETION_COL]:
        moved = subj_deferred or elec_later or prep_summer or to_summer
        pushed = " (full-load estimate; timing may shift)" if moved else ""
        bits.append(f"Earliest completion {out[COMPLETION_COL]}{pushed}")
    status = str(row.get("STUDY_PATH_STATUS", "") or "")
    if status and status != "Active Study Path":
        bits.append(f"NOTE: {status} - confirm the student is returning before acting")

    carried = c.get("carried_from", "")
    if carried:
        bits.append(
            f"ASSUMED: {carried} offering pattern used for {session} - "
            "sanity-check the subjects run that session; no completion estimate"
        )
    out[REASON_COL] = " | ".join(bits)

    src = _SRC_CALC_CE if capped else _SRC_CALC
    if carried:
        src += f" ({carried} pattern assumed)"
    out[SOURCE_COL] = src
    return out


# --------------------------------------------------------------------------- #
# Whole-file entry points                                                     #
# --------------------------------------------------------------------------- #
load_progression_file = v2.load_progression_file
to_workbook_bytes = v2.to_workbook_bytes
COACH_VIEW_SHEET = v2.COACH_VIEW_SHEET
SHEET_NAME = v2.SHEET_NAME


def build_advice(
    df: pd.DataFrame,
    session: str = DEFAULT_PLANNING_SESSION,
    offerings: dict | None = None,
) -> pd.DataFrame:
    """Copy of ``df`` with the advice + completion + reason + source columns."""
    offerings = offerings or v2.load_offerings()
    slot_map = v2.derive_slot_map(df)

    cols = [*ADVICE_COLS, COMPLETION_COL, REASON_COL, SOURCE_COL]
    out = df.copy()
    for col in cols:
        out[col] = ""

    for idx, row in df.iterrows():
        r = advise_student_merged(row, slot_map, offerings, session)
        for col in cols:
            out.at[idx, col] = r[col]
    return out


def build_coach_view(advised: pd.DataFrame) -> pd.DataFrame:
    """v2's slim readable sheet, plus the Earliest Completion + Advice Source.

    The status grid uses the calculator's outstanding/elective tests (not
    v2's) so the ✓/✗ marks line up with the advice the calculator produced.
    """
    base = advised[[c for c in v2.COACH_VIEW_COLUMNS if c in advised.columns]].copy()

    status = [
        v2._status(r, is_outstanding=calc._is_outstanding, elective_count=calc._elective_count)
        for _, r in advised.iterrows()
    ]
    base["Student Status"] = [s for s, _, _ in status]
    base["Progress Bar"] = [b for _, _, b in status]
    base["Progress"] = [g for _, g, _ in status]

    for col in [*ADVICE_COLS, COMPLETION_COL, SOURCE_COL, REASON_COL]:
        base[col] = advised[col].values
    return base
