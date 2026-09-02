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
_SRC_POSITIONAL = "position-based (fail pattern not in calculator)"
_SRC_V2 = "v2 rule-tree ({})"
_SRC_EXCLUDED = "standing: Exclusion (no advice)"


def _passed_count(row: pd.Series, n_positions: int) -> int:
    return sum(
        str(row.get(f"Subject {p} Status", "")).strip().endswith("Completed")
        for p in range(1, n_positions + 1)
    )


def _positional_fallback(row: pd.Series, program: str, is_nursing: bool) -> tuple[list[str], str]:
    """Advice for a fail pattern Grant's calculator has no row for.

    Each outstanding subject goes in the block it runs in (position P ->
    block ``((P - 1) % 4) + 1``); electives fill the gaps. Passed >= 2 subjects
    total -> keep the cohort (higher) subject when two land in one block;
    passed <= 1 -> restart, keep the lower. Prep 1 before prep 2 (a second
    outstanding prep surfaces in "still to pass").
    """
    prog_ref = calc._ref().get(program, {})
    subj = prog_ref.get("subjects", {})
    n_pos = 8 if is_nursing else 6
    cohort = _passed_count(row, n_pos) >= 2

    out_positions = [
        p for p in range(1, n_pos + 1)
        if calc._is_outstanding(row.get(f"Subject {p} Status"))
    ]
    blocks = ["", "", "", ""]
    for b in range(4):
        contenders = sorted((p for p in out_positions if (p - 1) % 4 == b), reverse=cohort)
        if contenders:
            blocks[b] = subj.get(str(contenders[0]), "")

    elec_need = calc._elective_count(row)
    placed = 0
    for i in range(4):
        if not blocks[i] and placed < elec_need:
            blocks[i] = "+1 elective"
            placed += 1

    prep_pick = ""
    if not is_nursing:
        preps = [
            prog_ref.get(k) for slot, k in (("Prep 1 Status", "prep1"), ("Prep 2 Status", "prep2"))
            if calc._is_outstanding(row.get(slot))
        ]
        if preps:
            prep_pick = preps[0]
    return blocks, prep_pick


CE_CAP_CP = 30  # Conditional Enrolment credit-point cap
_CP_MODULAR = 10
_CP_PREP = 15


def _split_prep(prep_pick: str) -> list[str]:
    """``"GEDU0016 and GEDU0017"`` -> ``["GEDU0016", "GEDU0017"]``."""
    if not prep_pick or prep_pick == calc.NO_REGISTRATION:
        return []
    return [p.strip() for p in prep_pick.replace(" and ", ",").split(",") if p.strip()]


def _ce_fill(
    positioned: list[str], prep_pick: str, elec_need: int
) -> tuple[list[str], list[str], str, str, int]:
    """Conditional Enrolment: fill the 30cp cap in order **modular subjects ->
    electives -> prep** (Josiah 2026-09-02). Modular = elective = 10cp, prep =
    15cp. Modular picks keep their block position; anything over the cap is
    deferred / sent to Summer.

    Returns ``(blocks, modular_deferred, prep_now, prep_summer, electives_now)``.
    """
    cp = 0
    blocks = ["" for _ in positioned]
    modular_deferred: list[str] = []
    for i, b in enumerate(positioned):
        if b and b != "+1 elective":
            if cp + _CP_MODULAR <= CE_CAP_CP:
                blocks[i] = b
                cp += _CP_MODULAR
            else:
                modular_deferred.append(b)

    electives_now = 0
    for i in range(len(blocks)):
        if not blocks[i] and electives_now < elec_need and cp + _CP_MODULAR <= CE_CAP_CP:
            blocks[i] = "+1 elective"
            cp += _CP_MODULAR
            electives_now += 1

    prep_now: list[str] = []
    prep_summer: list[str] = []
    for pc in _split_prep(prep_pick):
        if cp + _CP_PREP <= CE_CAP_CP:
            prep_now.append(pc)
            cp += _CP_PREP
        else:
            prep_summer.append(pc)

    return blocks, modular_deferred, " and ".join(prep_now), " and ".join(prep_summer), electives_now


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

    is_nursing = program in calc.NURSING_PROGRAMS
    in_ref = bool(calc._ref().get(program, {}).get("subjects"))

    # 2. Grant's calculator - only for the sessions it has offering patterns for.
    if rs.uses_calculator(session):
        c = calc.advise_row(row, session)
    else:
        c = {"ok": False, "miss": f"{session} not covered by the calculator"}

    carried = ""
    if c["ok"]:
        # 3a. Calculator answered. Keep its block positions; only cap the count.
        prep_pick = c[ADVICE_COLS[0]]
        prep_pick = "" if prep_pick == calc.NO_REGISTRATION else prep_pick
        positioned = ["" if c[col] == calc.NO_REGISTRATION else c[col] for col in ADVICE_COLS[1:]]
        completion = c[COMPLETION_COL]
        carried = c.get("carried_from", "")
        source_base = _SRC_CALC
    elif rs.uses_calculator(session) and in_ref:
        # 3b. Calculator covers this session but not this fail pattern - place
        #     each outstanding subject in its own block directly.
        positioned, prep_pick = _positional_fallback(row, program, is_nursing)
        completion = ""
        source_base = _SRC_POSITIONAL
    else:
        # 4. Part-way target, program 9034, or program not in the reference
        #    data -> v2 rule-tree.
        adv = v2.advise_student(row, slot_map, offerings, target=session)
        out[ADVICE_COLS[0]] = adv.prep
        for col, val in zip(ADVICE_COLS[1:], adv.blocks):
            out[col] = val
        out[REASON_COL] = adv.reason
        out[SOURCE_COL] = _SRC_V2.format(c["miss"])
        return out

    capped = outcome in v2.STANDING_MAX_BLOCKS  # Conditional Enrolment
    elec_need = calc._elective_count(row)

    if capped:
        kept, deferred, prep_now, prep_summer, elec_now = _ce_fill(positioned, prep_pick, elec_need)
    else:
        kept, deferred = list(positioned), []
        prep_now, prep_summer = prep_pick, ""
        elec_now = kept.count("+1 elective")
    named = [(i + 1, b) for i, b in enumerate(kept) if b]

    out[ADVICE_COLS[0]] = prep_now
    for col, val in zip(ADVICE_COLS[1:], kept):
        out[col] = val
    # a carried-forward pattern's completion estimate is stale (see rereg_calc)
    out[COMPLETION_COL] = "" if completion in ("", "Not Found") or carried else completion

    nothing = not named and not prep_now and not prep_summer

    bits: list[str] = []
    if capped and not nothing:
        bits.append(f"{outcome}: 30cp cap - failed subjects first, then electives, then prep")
    elif outcome == "At Risk" and not nothing:
        bits.append("At Risk (full load allowed - monitor)")
    if prep_now:
        bits.append(f"Prep: {prep_now}")
    if prep_summer:
        bits.append(f"Prep in Summer: {prep_summer}")
    if named:
        bits.append("Register: " + ", ".join(f"Block {n} {b}" for n, b in named))
    if deferred:
        bits.append("Defer to a later session: " + ", ".join(deferred))

    # Everything still outstanding that this session's advice doesn't touch -
    # the calculator only plans four blocks, so a coach needs the rest spelled
    # out (failed subjects that didn't fit, a second prep, unplaced electives).
    prog_ref = calc._ref().get(program, {})
    accounted = {b for _, b in named} | set(deferred) | {prep_now, prep_summer}
    still: list[str] = []
    for pos in range(1, 9):
        code = prog_ref.get("subjects", {}).get(str(pos))
        if code and calc._is_outstanding(row.get(f"Subject {pos} Status")) and code not in accounted:
            still.append(code)
    for pslot, pkey in (("Prep 1 Status", "prep1"), ("Prep 2 Status", "prep2")):
        pcode = prog_ref.get(pkey)
        if (pcode and calc._is_outstanding(row.get(pslot))
                and pcode not in accounted
                and pcode not in str(prep_now) and pcode not in str(prep_summer)
                and pcode not in still):
            still.append(pcode)
    elec_more = max(0, elec_need - elec_now)
    if elec_more:
        still.append(f"+{elec_more} elective")
    if still:
        bits.append("Still to pass (later session): " + ", ".join(still))
    if nothing and c.get("total_needed", 0) > 0:
        bits.append(
            f"Nothing to register in {session} - {c['total_needed']} subject(s) "
            "still to pass are not offered this session"
        )
    elif nothing:
        bits.append("Nothing to register this session")
    if out[COMPLETION_COL]:
        pushed = " (full-load estimate; deferrals push this out)" if (deferred or prep_summer) else ""
        bits.append(f"Earliest completion {out[COMPLETION_COL]}{pushed}")
    status = str(row.get("STUDY_PATH_STATUS", "") or "")
    if status and status != "Active Study Path":
        bits.append(f"NOTE: {status} - confirm the student is returning before acting")

    if carried:
        bits.append(
            f"ASSUMED: {carried} offering pattern used for {session} - "
            "sanity-check the subjects run that session; no completion estimate"
        )
    out[REASON_COL] = " | ".join(bits)

    src = source_base
    if source_base == _SRC_CALC and capped:
        src = _SRC_CALC_CE
    if capped and source_base == _SRC_POSITIONAL:
        src = source_base + " + CE 30cp cap"
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
