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

import json
import re
from pathlib import Path

import pandas as pd

from student_tracker import rereg_advice as v2
from student_tracker import rereg_calc as calc
from student_tracker import rereg_sessions as rs

# Same five advice columns as v2, plus two of our own.
ADVICE_COLS = v2.ADVICE_COLS
COMPLETION_COL = "Earliest Completion"
REASON_COL = v2.REASON_COL           # "Advice Reason"
SOURCE_COL = "Advice Source"
PRINCIPLE_COL = "Rereg Principle"
TEMPLATE_COL = "Messaging Template"
WITHDRAWAL_COL = "Withdrawal Flag"
STUDY_STATUS_COL = "Study Status"
OTHER_ENROL_COL = "Other Enrolment"
# Mid-semester only: what the student is actually registered in for the blocks
# still to come (shown beside that block's advice in the Coach View), and
# whether it matches the advice.
REG_CHECK_COL = "Registration Check"
REGISTERED_COLS = {3: "Block 3 Registered", 4: "Block 4 Registered"}
REG_OK = "OK"

# Deferred / Leave of Absence are the "are you studying with us?" group:
# enrolled on paper, paused in practice, so the subject advice on their row is
# provisional until they confirm they're coming back. The raw STUDY_PATH_STATUS
# column carries this straight from the extract - no relabelling needed.
STUDY_PATH_COL = "STUDY_PATH_STATUS"
ACTIVE_STATUS = "Active Study Path"
_PAUSED_TAB = "Paused - check enrolment"
_COMMENCING_TAB = "Commencing"
# Excluded students get no advice, so they're taken off the Coach View / template
# tabs and listed on a tab of their own (the full sheet still has everyone).
_EXCLUDED_TAB = "Excluded"


def _is_excluded(cv: pd.DataFrame) -> pd.Series:
    """Coach View rows whose Progression Outcome is Exclusion."""
    if STUDY_STATUS_COL not in cv.columns:
        return pd.Series(False, index=cv.index)
    return cv[STUDY_STATUS_COL].astype(str).str.strip().isin(v2.STANDING_NO_ADVICE)

# Progression Outcome is blank for ~a quarter of a mid-semester file. A
# commencing student legitimately has no decision yet; anyone older should have
# one, so they are called out for review rather than defaulted to Good Standing.
NOT_ASSESSED = "Not yet assessed (commencing)"
NOT_ASSESSED_PAUSED = "Not assessed (paused)"
NO_OUTCOME = "No outcome recorded - review"
_PASS_GRADE_LETTERS = set("ABCDHP")  # first letter of a pass grade (A/B/C/C+/D/H/P)
_WITHDRAW_TEXT = "ADVISE WITHDRAWAL"
_CHECK_ENROL_TEXT = "CHECK ENROLMENT"
_AWAITING_TEXT = "AWAITING GRADE"
_SRC_WITHDRAW = "mid-semester withdrawal (commencing student, failed Blocks 1-2)"
_PAUSED_NO_WITHDRAW_NOTE = (
    "No Block 1-2 passes, but paused - no withdrawal advice (not sent to paused students)"
)
_EARLY_FLAG_NOTES = {
    "later": ("CHECK ENROLMENT: not enrolled in Blocks 1-2 but registered later this "
              "session - confirm the enrolment is right (not a withdrawal case)"),
    "nothing": ("CHECK ENROLMENT: no Block 1-2 pass and no later registration this "
                "session - nothing to withdraw from; check they are still studying"),
    _AWAITING_TEXT: ("AWAITING GRADE: no Block 1-2 pass yet but a grade is still to come - "
                     "recheck before any withdrawal advice"),
}


def _block_passed(result) -> bool:
    """True when a ``Block N Result`` cell shows a pass. A block can carry more
    than one comma-separated grade (a student in two subjects that block); the
    block counts as passed if *any* part is a pass grade. Everything else -
    F / FNS / W / E and a blank (no grade recorded) - is 'not passed'."""
    return any(
        part.strip()[:1].upper() in _PASS_GRADE_LETTERS
        for part in str(result or "").split(",")
        if part.strip()
    )


def _block_outcome(row: pd.Series, block: int) -> str:
    """How one already-run block went: ``"pass"``, ``"fail"`` (a grade is
    recorded and none of it is a pass), ``"pending"`` (enrolled, no grade yet)
    or ``"none"`` (not enrolled in that block)."""
    if not _SUBJECT_CODE_RE.search(str(row.get(f"Block {block} code", "") or "")):
        return "none"
    raw = row.get(f"Block {block} Result")
    raw = "" if raw is None or pd.isna(raw) else str(raw)
    if _block_passed(raw):
        return "pass"
    return "fail" if raw.replace(",", "").strip() else "pending"


def _early_block_flag(row: pd.Series, from_block: int) -> str:
    """Mid-session check on a commencing student with no pass so far this
    session. Returns the Withdrawal Flag value, or ``""`` when nothing is wrong.

    Agreed with the coaching team (2026-09-28) - withdrawal only makes sense
    when there is something to withdraw from and something actually failed:

    - any block passed                           -> ``""``
    - a block still waiting on its grade          -> AWAITING GRADE
    - at least one fail, registered in a later block -> ADVISE WITHDRAWAL
      (a fail plus a block they weren't enrolled in counts)
    - otherwise - never enrolled in the early blocks, or nothing left to
      withdraw from                                -> CHECK ENROLMENT
    """
    if from_block < 2:
        return ""
    outcomes = [_block_outcome(row, b) for b in range(1, from_block)]
    if "pass" in outcomes:
        return ""
    if "pending" in outcomes:
        return _AWAITING_TEXT
    later = any(
        _SUBJECT_CODE_RE.search(str(row.get(f"Block {b} code", "") or ""))
        for b in range(from_block, 5)
    )
    if "fail" in outcomes and later:
        return _WITHDRAW_TEXT
    return _CHECK_ENROL_TEXT


def _failed_subjects(row: pd.Series, from_block: int) -> list[str]:
    """``["GEDU1001 (Block 1)", ...]`` - the early-block subjects with a
    non-pass grade, for the withdrawal reason text."""
    out = []
    for b in range(1, from_block):
        codes = _SUBJECT_CODE_RE.findall(str(row.get(f"Block {b} code", "") or ""))
        raw = row.get(f"Block {b} Result")
        raw = "" if raw is None or pd.isna(raw) else str(raw)
        grades = raw.split(",") if raw.strip() else []
        if len(grades) == len(codes):
            out += [f"{c} (Block {b})" for c, g in zip(codes, grades)
                    if _result_mark(g) == RESULT_FAIL]
        elif _result_mark(raw) == RESULT_FAIL:
            out += [f"{c} (Block {b})" for c in codes]
    return out


RESULT_PASS = "✓"
RESULT_FAIL = "✗"
NOT_ENROLLED = "0"


def _result_mark(grade: str) -> str:
    """``✓`` for a pass grade, ``✗`` for any other recorded grade (F / FNS / E /
    W), ``""`` when no grade is recorded yet."""
    grade = grade.strip()
    if not grade:
        return ""
    return RESULT_PASS if _block_passed(grade) else RESULT_FAIL


def _earlier_block_cell(row: pd.Series, block: int) -> str:
    """The greyed advice cell for a block that has already run this session:
    what the student was enrolled in, each subject tagged with its result
    (``GEDU1001 ✓``, ``EDUC1012 ✗``), or ``0`` when they weren't enrolled in
    that block at all - so "didn't take it" and "took it and failed" read
    differently (agreed with the coaching team, 2026-09-28).

    A block can hold more than one subject; its ``Block N Result`` is then a
    comma list in the same order. If the two lists don't line up, the block's
    overall result goes on the end instead of per subject."""
    codes = _SUBJECT_CODE_RE.findall(str(row.get(f"Block {block} code", "") or ""))
    if not codes:
        return _GREY + NOT_ENROLLED
    raw = row.get(f"Block {block} Result")
    raw = "" if raw is None or pd.isna(raw) else str(raw)
    grades = raw.split(",") if raw.strip() else []
    if len(grades) == len(codes):
        parts = [f"{c} {_result_mark(g)}".rstrip() for c, g in zip(codes, grades)]
    else:
        parts = [f"{', '.join(codes)} {_result_mark(raw)}".rstrip()]
    return _GREY + ", ".join(parts)


def _mark_earlier_blocks(out: dict, row: pd.Series, from_block: int) -> None:
    """Mid-semester target: overwrite the advice cells for Blocks
    1..from_block-1 (already started, nothing to register) with what the
    student took there and how it went."""
    for b in range(1, from_block):
        out[ADVICE_COLS[b]] = _earlier_block_cell(row, b)


def _enrolled_earlier_codes(row: pd.Series, from_block: int) -> list[str]:
    """Every subject code the student was enrolled in across Blocks
    1..from_block-1 this session (a block can hold more than one)."""
    return [
        code for b in range(1, from_block)
        for code in _SUBJECT_CODE_RE.findall(str(row.get(f"Block {b} code", "") or ""))
    ]


def _enrolled_earlier_blocks(row: pd.Series, from_block: int) -> dict[str, int]:
    """``{subject_code: block}`` for subjects the student is already enrolled in
    this session, in the blocks before ``from_block`` (the file's ``Block N code``
    columns)."""
    out: dict[str, int] = {}
    for b in range(1, from_block):
        for code in _SUBJECT_CODE_RE.findall(str(row.get(f"Block {b} code", "") or "")):
            out.setdefault(code, b)
    return out


def _commenced_this_session(row: pd.Series, base: str) -> bool:
    """True when the student started their course in the session being advised
    for. The mid-semester withdrawal rule (Grant, 2026-09-09) is for commencing
    students only - a continuing student who fails two blocks in a row is
    advised to re-take, not withdraw."""
    return calc.start_semester(row.get("COMMENCEMENT_PERIOD")) == base

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


# Electives are not offered in Blocks 1 & 2 (Josiah 2026-09-02) - only 3 & 4.
_ELECTIVE_BLOCKS = (2, 3)


def _electives_to_b34(blocks: list[str]) -> list[str]:
    """Move any elective sitting in Block 1/2 to an empty Block 3/4; if there's
    no room it is dropped (and surfaces in "still to pass")."""
    for i in (0, 1):
        if blocks[i] == "+1 elective":
            blocks[i] = ""
            for j in _ELECTIVE_BLOCKS:
                if not blocks[j]:
                    blocks[j] = "+1 elective"
                    break
    return blocks


def _positional_fallback(row: pd.Series, program: str, is_nursing: bool, session: str) -> tuple[list[str], str]:
    """Advice for a fail pattern Grant's calculator has no row for.

    Diploma: each outstanding subject goes in the block it runs in (position P
    -> block ``((P - 1) % 4) + 1``); a block clash is won by the cohort (higher)
    position when the student has passed >= 2 subjects, else by the lower
    (restart). Nursing: the 8 subjects run strictly in sequence, so the first
    four outstanding ones fill Blocks 1-4 in number order. Electives fill empty
    Blocks 3 & 4 only. Prep 1 before prep 2.
    """
    prog_ref = calc._ref().get(program, {})
    subj = calc.subjects_for(program, session)
    n_pos = 8 if is_nursing else 6

    out_positions = [
        p for p in range(1, n_pos + 1)
        if calc._is_outstanding(row.get(f"Subject {p} Status"))
    ]
    blocks = ["", "", "", ""]
    if is_nursing:
        for i, p in enumerate(out_positions[:4]):
            blocks[i] = subj.get(str(p), "")
    else:
        cohort = _passed_count(row, n_pos) >= 2
        for b in range(4):
            contenders = sorted((p for p in out_positions if (p - 1) % 4 == b), reverse=cohort)
            if contenders:
                blocks[b] = subj.get(str(contenders[0]), "")

    elec_need = calc._elective_count(row)
    placed = 0
    for i in _ELECTIVE_BLOCKS:
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


def _preps_in_progress(row: pd.Series, program: str) -> list[str]:
    """The program's prep codes the student is enrolled in right now (a
    ``Prep N Status`` of ``"… Currently Registered"``)."""
    ref = calc._ref().get(program, {})
    return [
        ref[key] for slot, key in (("Prep 1 Status", "prep1"), ("Prep 2 Status", "prep2"))
        if ref.get(key) and "Currently Registered" in str(row.get(slot) or "")
    ]


def _ce_fill(
    positioned: list[str], prep_pick: str, elec_need: int, cp_start: int = 0
) -> tuple[list[str], list[str], str, str, int]:
    """Conditional Enrolment: fill the 30cp cap in order **modular subjects ->
    electives -> prep** (Josiah 2026-09-02). Modular = elective = 10cp, prep =
    15cp. Modular picks keep their block position; anything over the cap is
    deferred / sent to Summer. ``cp_start`` is credit already used this
    session outside the picks (a prep the student is enrolled in now).

    Returns ``(blocks, modular_deferred, prep_now, prep_summer, electives_now)``.
    """
    cp = cp_start
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
    for i in _ELECTIVE_BLOCKS:  # electives only in Blocks 3 & 4
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


def _is_owed(row: pd.Series, program: str, code: str, session: str) -> bool:
    """Is ``code`` a subject / prep this student still needs to pass?"""
    ref = calc._ref().get(program, {})
    for pos, c in calc.subjects_for(program, session).items():
        if c == code:
            return calc._is_outstanding(row.get(f"Subject {pos} Status"))
    if code == ref.get("prep1"):
        return calc._is_outstanding(row.get("Prep 1 Status"))
    if code == ref.get("prep2"):
        return calc._is_outstanding(row.get("Prep 2 Status"))
    return False


_SRC_SUMMER = "Summer"
GREY_MARK = "~"  # advice-cell prefix -> styled grey (for reference, not registered now)
_GREY = GREY_MARK


def strip_grey(df: pd.DataFrame) -> pd.DataFrame:
    """A display copy with the grey marker turned into ``(brackets)``."""
    return df.replace(r"^~(.+)$", r"(\1)", regex=True)


def _grey_cell(value) -> str:
    """Mark one advice cell for-reference-only. Idempotent; leaves a blank and
    the "no registration needed" sentinel alone - neither is a pick to action."""
    text = "" if pd.isna(value) else str(value).strip()
    if not text or text.startswith(GREY_MARK) or text == calc.NO_REGISTRATION:
        return text
    return GREY_MARK + text


def _normalize_summer_offering(summer) -> dict | None:
    """A Summer offering as ``{code: {"campuses": set, "block": str}}``, or None.

    Accepts the campus/block dict from :func:`read_summer_offering_campus`, an
    older ``{code: campuses}`` dict, a plain set of codes (all campuses, no
    block), or ``None`` (no list uploaded). This lets the engine stay campus-
    and block-aware while still accepting a bare code set.
    """
    if summer is None:
        return None
    if isinstance(summer, dict):
        norm = {}
        for code, v in summer.items():
            if isinstance(v, dict):
                norm[code] = {"campuses": set(v.get("campuses") or KNOWN_CAMPUSES),
                              "block": v.get("block", "")}
            else:
                norm[code] = {"campuses": set(v) if v else set(KNOWN_CAMPUSES), "block": ""}
        return norm
    return {c: {"campuses": set(KNOWN_CAMPUSES), "block": ""} for c in summer}


def _summer_advice(
    out: dict, row: pd.Series, program: str, is_nursing: bool, session: str, outcome: str,
    summer_subjects=None,
) -> dict:
    """Summer advice.

    Preps always run in Summer (diplomas only); every other subject runs only if
    it's on the uploaded offering list, **at the student's campus**, and is
    placed in the Summer block (SU1 / SU2) that list gives it - at most one
    subject per Summer block. A subject that's offered but whose Summer block is
    already taken, or that isn't offered at their campus, is greyed for reference
    (take next Autumn). With no list uploaded, only prep + Subjects 1 & 2 are
    assumed to run. Conditional Enrolment keeps the 30cp cap.
    """
    offering = _normalize_summer_offering(summer_subjects)
    assumed = offering is None
    campus = "" if pd.isna(row.get("CAMP_CODE")) else str(row.get("CAMP_CODE")).strip()
    prog_ref = calc._ref().get(program, {})
    subj = calc.subjects_for(program, session)
    capped = outcome in v2.STANDING_MAX_BLOCKS
    cp_cap = CE_CAP_CP if capped else 10 ** 6
    n_pos = 8 if is_nursing else 6

    def offered_block(pos: int) -> tuple[bool, str]:
        """``(offered at this student's campus, Summer block)`` for the subject at
        ``pos``. With no list uploaded, fall back to 'Subjects 1 & 2 run'."""
        code = subj.get(str(pos))
        if not code:
            return (False, "")
        if offering is None:
            return (pos in (1, 2), "")
        entry = offering.get(code)
        if entry and campus in entry["campuses"]:
            return (True, entry["block"])
        return (False, "")

    # preps first (always run in Summer, diplomas only)
    prep_now, cp = [], 0
    for slot, key in (("Prep 1 Status", "prep1"), ("Prep 2 Status", "prep2")):
        code = prog_ref.get(key)
        if not is_nursing and code and calc._is_outstanding(row.get(slot)) and cp + _CP_PREP <= cp_cap:
            prep_now.append(code)
            cp += _CP_PREP

    blocks = ["", "", "", ""]
    grey = [False, False, False, False]
    btag = ["", "", "", ""]
    deferred = [False, False, False, False]  # greyed because offered but Summer block full / cap
    register: list[str] = []
    used_su: set[str] = set()  # a student takes at most one subject per Summer block
    for bi in range(4):
        cands = [
            p for p in (bi + 1, bi + 5)
            if subj.get(str(p)) and calc._is_outstanding(row.get(f"Subject {p} Status"))
        ]
        if not cands:
            continue
        pick = None
        offered_but_stuck = False
        for p in cands:
            ok, blk = offered_block(p)
            if not ok:
                continue
            if (blk and blk in used_su) or cp + _CP_MODULAR > cp_cap:
                offered_but_stuck = True  # it runs at their campus, just can't fit now
                continue
            pick = (subj[str(p)], blk)
            break
        if pick:
            code, blk = pick
            blocks[bi] = code
            btag[bi] = blk
            register.append(code)
            if blk:
                used_su.add(blk)
            cp += _CP_MODULAR
        else:
            blocks[bi] = subj[str(cands[0])]
            grey[bi] = True
            deferred[bi] = offered_but_stuck

    out[ADVICE_COLS[0]] = " and ".join(prep_now)
    for col, val, g, tag in zip(ADVICE_COLS[1:], blocks, grey, btag):
        if not val:
            out[col] = ""
        elif g:
            out[col] = _GREY + val
        else:
            out[col] = f"{val} ({tag})" if tag else val

    scheduled = set(prep_now) | set(register)
    # outstanding items with no block shown at all (a lost 1-vs-5 clash, a
    # second prep) -> next Autumn
    still: list[str] = []
    for pos in range(1, 9):
        code = subj.get(str(pos))
        if code and calc._is_outstanding(row.get(f"Subject {pos} Status")) \
                and code not in scheduled and code not in blocks:
            still.append(code)
    for slot, key in (("Prep 1 Status", "prep1"), ("Prep 2 Status", "prep2")):
        code = prog_ref.get(key)
        if code and calc._is_outstanding(row.get(slot)) and code not in scheduled:
            still.append(code)
    elec_need = calc._elective_count(row)
    if elec_need:
        still.append(f"+{elec_need} elective")
    greyed = [b for b, g in zip(blocks, grey) if g]
    # split the greyed reference subjects: some run in Summer but couldn't fit
    # (their block/cap is full), the rest simply aren't offered at their campus
    deferred_full = [b for b, g, d in zip(blocks, grey, deferred) if g and d]
    not_offered = [b for b, g, d in zip(blocks, grey, deferred) if g and not d]

    # registered subjects with their Summer block, in SU1-then-SU2 order
    reg_by_block = [(btag[i], b) for i, b in enumerate(blocks) if b and not grey[i]]
    reg_display = [f"{code} ({blk})" if blk else code
                   for blk, code in sorted(reg_by_block, key=lambda x: x[0])]

    bits = []
    if capped:
        bits.append(f"{outcome}: 30cp cap")
    if prep_now:
        bits.append("Prep: " + " and ".join(prep_now))
    if reg_display:
        bits.append("Register: " + ", ".join(reg_display))
    if not prep_now and not register:
        bits.append(f"Nothing confirmed for this student runs in {session}")
    if deferred_full:
        bits.append("Runs in Summer but only one subject per block fits - take later: "
                    + ", ".join(deferred_full))
    if not_offered:
        bits.append("Grey = not offered in Summer at their campus - take next Autumn: "
                    + ", ".join(not_offered))
    if assumed and (greyed or register):
        bits.append("ASSUMED: only prep + Subjects 1 & 2 run in Summer - upload the Summer offering list for real advice")
    if still:
        bits.append(f"Also still to pass: " + ", ".join(still))

    # rough completion: everything not registered this Summer (greyed blocks +
    # unblocked items + electives), from the next Autumn at 4 (3 for CE)/session.
    work = len(greyed) + sum(1 for s in still if not s.startswith("+")) + elec_need
    load = 3 if capped else 4
    if work:
        est = rs.advance(session, min(8, -(-work // load)))
        out[COMPLETION_COL] = f"{est} (est.)"
        bits.append(f"Earliest completion ~{est}")
    elif prep_now or register:
        out[COMPLETION_COL] = f"{session} (est.)"

    out[REASON_COL] = " | ".join(bits)
    src = _SRC_SUMMER + (" (assumed offering)" if assumed else " (uploaded offering)")
    out[SOURCE_COL] = src + (" + CE 30cp cap" if capped else "")
    return out


def advise_student_merged(
    row: pd.Series, slot_map: dict, offerings: dict, session: str,
    summer_subjects: set[str] | None = None,
) -> dict:
    program = str(row["PROGRAM_CD"]).split(".")[0]
    outcome = str(row.get("Progression Outcome", "") or "").strip()

    is_nursing = program in calc.NURSING_PROGRAMS
    in_ref = bool(calc._ref().get(program, {}).get("subjects"))
    principle, template = calc.classify(row, is_nursing)
    # statuses in the position order the program runs in for this session
    srow = calc.remap_statuses(row, program, session, slot_map)

    out = {c: "" for c in ADVICE_COLS}
    out[COMPLETION_COL] = ""
    out[REASON_COL] = ""
    out[SOURCE_COL] = ""
    out[PRINCIPLE_COL] = principle
    out[TEMPLATE_COL] = template
    out[WITHDRAWAL_COL] = ""

    # 1. Exclusion -> no advice, whichever engine would have run.
    if outcome in v2.STANDING_NO_ADVICE:
        out[REASON_COL] = f"{outcome} - not eligible to re-register; refer to coach."
        out[SOURCE_COL] = _SRC_EXCLUDED
        return out

    # 1b. Summer target -> preps + whatever the Summer offering list says runs
    #     (or just Subjects 1 & 2 when no list has been uploaded).
    tgt = rs.parse_target(session)
    if tgt and tgt[1] == "SUM":
        return _summer_advice(out, srow, program, is_nursing, session, outcome, summer_subjects)

    # A part-way target ("26 AUT Block 3") uses the whole-session engines - the
    # picks are already locked to the block each subject runs in - and only
    # Blocks <from_block>..4 are actually registered now.
    base = rs.base_session(session)
    from_block = rs.target_block(session)

    # Mid-semester check (Stage 1): a *commencing* student with no pass in the
    # blocks run so far this session (see _early_block_flag). ADVISE WITHDRAWAL
    # -> no subject advice, just flag them for the coach. CHECK ENROLMENT /
    # AWAITING GRADE -> flagged, but the normal advice below still runs. A
    # continuing student who fails two blocks in a row is advised to re-take,
    # not withdraw.
    #
    # UPPAP (program 9034) is excluded: its pattern doesn't start at Block 1, so
    # "hasn't cleared Block 1/2" doesn't mean the same thing - those students are
    # not registered in Blocks 1/2 by design, not because they failed to.
    #
    # A paused (Deferred / Leave of Absence) student is never flagged
    # (coaching team, 2026-09-28): they have no Block 1-2 passes because they
    # stepped away, not because they failed. They fall through to the normal
    # (greyed, provisional) advice with a note for the coach.
    early_flag = (
        _early_block_flag(row, from_block)
        if program not in calc.UNSUPPORTED_PROGRAMS and _commenced_this_session(row, base)
        else ""
    )
    paused_no_withdraw = bool(early_flag) and is_paused(row.get(STUDY_PATH_COL))
    early_note = ""
    if paused_no_withdraw:
        early_flag = ""
    elif early_flag == _WITHDRAW_TEXT:
        _mark_earlier_blocks(out, row, from_block)
        failed = ", ".join(_failed_subjects(row, from_block))
        out[WITHDRAWAL_COL] = _WITHDRAW_TEXT
        out[REASON_COL] = (
            f"** {_WITHDRAW_TEXT} ** - commencing student, no pass in Block(s) "
            f"1-{from_block - 1} this session (failed: {failed}). Drop Block "
            f"{from_block}-4 registrations; restart next semester as a commencing student."
        )
        out[SOURCE_COL] = _SRC_WITHDRAW
        return out
    elif early_flag == _AWAITING_TEXT:
        early_note = _EARLY_FLAG_NOTES[_AWAITING_TEXT]
    elif early_flag == _CHECK_ENROL_TEXT:
        registered_later = any(
            _SUBJECT_CODE_RE.search(str(row.get(f"Block {b} code", "") or ""))
            for b in range(from_block, 5)
        )
        early_note = _EARLY_FLAG_NOTES["later" if registered_later else "nothing"]

    # 2. Grant's calculator - only for the sessions it has offering patterns for.
    if rs.uses_calculator(session):
        c = calc.advise_row(srow, base)
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
        positioned, prep_pick = _positional_fallback(srow, program, is_nursing, session)
        completion = ""
        source_base = _SRC_POSITIONAL
    else:
        # 4. Part-way target, program 9034, or program not in the reference
        #    data -> v2 rule-tree.
        adv = v2.advise_student(row, slot_map, offerings, target=session)
        out[ADVICE_COLS[0]] = adv.prep
        for col, val in zip(ADVICE_COLS[1:], adv.blocks):
            out[col] = val
        _mark_earlier_blocks(out, row, from_block)
        out[REASON_COL] = adv.reason
        if paused_no_withdraw:
            out[REASON_COL] += f" | {_PAUSED_NO_WITHDRAW_NOTE}"
        if early_note:
            out[WITHDRAWAL_COL] = early_flag
            out[REASON_COL] = f"{early_note} | {out[REASON_COL]}"
        out[SOURCE_COL] = _SRC_V2.format(c["miss"])
        return out

    capped = outcome in v2.STANDING_MAX_BLOCKS  # Conditional Enrolment
    elec_need = calc._elective_count(row)

    # Mid-semester, a prep the student is enrolled in right now ("Currently
    # Registered") is part of this session's load: it counts 15cp toward the
    # Conditional Enrolment cap, and isn't advised again or sent to Summer.
    prep_enrolled = _preps_in_progress(row, program) if capped and from_block > 1 else []
    if prep_enrolled:
        prep_pick = " and ".join(p for p in _split_prep(prep_pick) if p not in prep_enrolled)

    # Mid-semester, Blocks 1..from_block-1 have already run: the cap counts what
    # the student was actually enrolled in there (10cp a subject, passed or
    # failed - coaching team, 2026-09-28), not the pattern's picks for those
    # blocks, and applies for the rest of this semester. The early picks are
    # kept out of the fill and put back afterwards, for the "still owes" text.
    early_cp = 0
    early_picks: list[str] = []
    if capped and from_block > 1:
        early_cp = _CP_MODULAR * len(_enrolled_earlier_codes(row, from_block))
        early_picks = [b if b != "+1 elective" else "" for b in positioned[:from_block - 1]]
        positioned = [""] * (from_block - 1) + list(positioned[from_block - 1:])
    session_cp = _CP_PREP * len(prep_enrolled) + early_cp  # already used this session

    if capped:
        kept, deferred, prep_now, prep_summer, elec_now = _ce_fill(
            positioned, prep_pick, elec_need, cp_start=session_cp)
        kept[:len(early_picks)] = early_picks
    else:
        kept, deferred = list(positioned), []
        prep_now, prep_summer = prep_pick, ""
        elec_now = kept.count("+1 elective")

    # electives are not offered in Blocks 1 & 2 (applies to Grant's picks too)
    kept = _electives_to_b34(kept)
    elec_now = kept.count("+1 elective")

    # Part-way target: only Blocks >= from_block are registered now. Force each
    # remaining block to hold its own outstanding subject (the calculator
    # sometimes leaves a backlog subject unplaced), then keep the earlier
    # blocks displayed - the subjects that belong there. For a Conditional
    # Enrolment student the 30cp cap still binds: the in-progress Blocks
    # 1..from_block-1 already eat into it, so only fill a remaining block that
    # _ce_fill left empty while there is room left (an overwrite of an existing
    # pick is credit-neutral). Anything skipped falls through to "still to pass".
    prog_subj = calc.subjects_for(program, session)
    # A subject already being taken in Blocks 1..from_block-1 this session is
    # never advised again in a later block (e.g. a transition session where the
    # same subject runs in two blocks).
    taking = _enrolled_earlier_blocks(row, from_block)
    in_progress: list[tuple[int, str]] = []
    if from_block > 1:
        cp_used = 0
        if capped:
            cp_used = session_cp + _CP_PREP * len(_split_prep(prep_now))
            cp_used += _CP_MODULAR * sum(1 for b in kept[from_block - 1:] if b)
        for bi in range(from_block - 1, 4):
            forced = next(
                (prog_subj[str(pos)] for pos in (bi + 1, bi + 5)
                 if prog_subj.get(str(pos))
                 and prog_subj[str(pos)] not in taking
                 and calc._is_outstanding(srow.get(f"Subject {pos} Status"))),
                None,
            )
            if forced is None:
                continue
            if capped and not kept[bi]:
                if cp_used + _CP_MODULAR > CE_CAP_CP:
                    continue  # cap reached - leave this block empty
                cp_used += _CP_MODULAR
            kept[bi] = forced
        for bi in range(from_block - 1, 4):
            if kept[bi] in taking:
                in_progress.append((taking[kept[bi]], kept[bi]))
                kept[bi] = ""
        elec_now = kept.count("+1 elective")
    named_all = [(i + 1, b) for i, b in enumerate(kept) if b]
    if from_block > 1:
        named = [(n, b) for n, b in named_all if n >= from_block]
        partway_carry = [
            b for n, b in named_all
            if n < from_block and b != "+1 elective" and _is_owed(srow, program, b, session)
        ]
    else:
        named = named_all
        partway_carry = []

    # For a mid-semester target the earlier blocks have already started, so their
    # advice cells show the subject the student actually took in that block and
    # its result (greyed, for reference), not a pattern subject they can no longer
    # register in. What they still owe from those blocks is in the reason text.
    out[ADVICE_COLS[0]] = prep_now
    for col, val in zip(ADVICE_COLS[1:], kept):
        out[col] = val
    _mark_earlier_blocks(out, row, from_block)
    # a carried-forward pattern's completion estimate is stale (see rereg_calc)
    out[COMPLETION_COL] = "" if completion in ("", "Not Found") or carried else completion

    nothing = not named and not prep_now and not prep_summer

    bits: list[str] = []
    if from_block > 1:
        bits.append(f"Advising from Block {from_block} - register Blocks {from_block}-4 only")
    if capped and not nothing:
        bits.append(f"{outcome}: 30cp cap - failed subjects first, then electives, then prep")
    elif outcome == "At Risk" and not nothing:
        bits.append("At Risk (full load allowed - monitor)")
    if prep_enrolled:
        bits.append(f"Prep in progress: {' and '.join(prep_enrolled)} (counted in the 30cp cap)")
    if session_cp:
        parts = ([f"prep {_CP_PREP * len(prep_enrolled)}cp"] if prep_enrolled else []) + (
            [f"Blocks 1-{from_block - 1} {early_cp}cp"] if early_cp else [])
        bits.append(f"Already enrolled this session: {session_cp}cp of 30 ({', '.join(parts)})")
    if prep_now:
        bits.append(f"Prep: {prep_now}")
    if prep_summer:
        bits.append(f"Prep in Summer: {prep_summer}")
    if named:
        bits.append("Register: " + ", ".join(f"Block {n} {b}" for n, b in named))
    if in_progress:
        bits.append(
            "Already enrolled this session (not re-advised): "
            + ", ".join(f"{code} (Block {blk})" for blk, code in sorted(in_progress))
        )
    if partway_carry:
        bits.append(
            f"Blocks 1-{from_block - 1} already in progress; still owes "
            + ", ".join(partway_carry) + " (take next Autumn)"
        )
    if deferred:
        bits.append("Defer to a later session: " + ", ".join(deferred))

    # Everything still outstanding that this session's advice doesn't touch -
    # the calculator only plans four blocks, so a coach needs the rest spelled
    # out (failed subjects that didn't fit, a second prep, unplaced electives).
    prog_ref = calc._ref().get(program, {})
    accounted = ({b for _, b in named} | set(deferred) | set(partway_carry)
                 | {code for _, code in in_progress} | {prep_now, prep_summer}
                 | set(prep_enrolled))
    still: list[str] = []
    for pos in range(1, 9):
        code = prog_subj.get(str(pos))
        if code and calc._is_outstanding(srow.get(f"Subject {pos} Status")) and code not in accounted:
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
    not_offered = c.get("total_needed", 0) - len(in_progress)
    if nothing and session_cp + _CP_MODULAR > CE_CAP_CP:
        bits.append(f"Nothing more to register in {session} - {outcome} 30cp cap reached")
    elif nothing and not_offered > 0:
        bits.append(
            f"Nothing to register in {session} - {not_offered} subject(s) "
            "still to pass are not offered this session"
        )
    elif nothing:
        bits.append("Nothing to register this session")

    # Earliest completion: Grant's own date for 26 AUT; a rough projection
    # (this session + ceil(work left / load) more sessions) for everything else.
    mod_out = sum(calc._is_outstanding(srow.get(f"Subject {i} Status")) for i in range(1, 9))
    prep_out = sum(calc._is_outstanding(row.get(f"Prep {i} Status")) for i in (1, 2))
    sched_mod = sum(1 for _, b in named if b != "+1 elective")
    load = 3 if capped else 4
    mod_left = max(0, (mod_out - sched_mod) + (elec_need - elec_now))
    prep_left = max(0, prep_out - len(_split_prep(prep_now)) - len(prep_enrolled))
    sessions_after = min(8, max(-(-mod_left // load), prep_left))
    estimate = rs.advance(session, sessions_after)

    if out[COMPLETION_COL]:  # Grant's real date (26 AUT)
        pushed = " (full-load estimate; deferrals push this out)" if (deferred or prep_summer) else ""
        bits.append(f"Earliest completion {out[COMPLETION_COL]}{pushed}")
    elif estimate and not nothing:
        out[COMPLETION_COL] = f"{estimate} (est.)"
        bits.append(
            f"Earliest completion ~{estimate}"
            + (f" ({sessions_after} more session(s) after this)" if sessions_after else " (this session)")
        )
    elif estimate and nothing and mod_left == 0 and prep_left == 0:
        out[COMPLETION_COL] = f"{session} (est.)"
    status = str(row.get("STUDY_PATH_STATUS", "") or "")
    if status and status != "Active Study Path":
        bits.append(f"NOTE: {status} - confirm the student is returning before acting")
    if paused_no_withdraw:
        bits.append(_PAUSED_NO_WITHDRAW_NOTE)
    if early_note:
        out[WITHDRAWAL_COL] = early_flag
        bits.insert(0, early_note)

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
def load_progression_file(source) -> pd.DataFrame:
    """v2's loader, with exact duplicate rows dropped.

    A student listed twice under the SAME program with every column identical is
    an extract artefact, not a dual enrolment: advising them twice double-counts
    them and puts the same row on a coach's tab twice. Rows that differ anywhere
    - a genuine second program - are kept, because those are two real study
    paths needing separate advice.

    The number dropped is recorded in ``df.attrs["duplicates_dropped"]``.
    """
    df = v2.load_progression_file(source)
    deduped = df.drop_duplicates().reset_index(drop=True)
    deduped.attrs["duplicates_dropped"] = len(df) - len(deduped)
    return deduped
COACH_VIEW_SHEET = v2.COACH_VIEW_SHEET
SHEET_NAME = v2.SHEET_NAME


def _style_advice_sheet(ws) -> None:
    """Apply the two bits of styling the advice output uses to one worksheet:
    an advice-block cell prefixed with the grey marker becomes grey italic text
    (for-reference, not registered) with the marker stripped, a row flagged
    ``ADVISE WITHDRAWAL`` gets its flag + reason cells in red bold, and a
    ``CHECK ENROLMENT`` flag is amber bold."""
    from openpyxl.styles import Font

    grey_font = Font(color="808080", italic=True)
    red_font = Font(color="C00000", bold=True)
    header = list(next(ws.iter_rows(max_row=1, values_only=True)))
    block_cols = {i for i, c in enumerate(header, 1) if c in ADVICE_COLS}
    flag_col = next((i for i, c in enumerate(header, 1) if c == WITHDRAWAL_COL), None)
    red_cols = {i for i, c in enumerate(header, 1) if c in (WITHDRAWAL_COL, REASON_COL)}
    amber_font = Font(color="C65911", bold=True)
    for rowcells in ws.iter_rows(min_row=2):
        flag = str(rowcells[flag_col - 1].value or "").strip() if flag_col else ""
        for cell in rowcells:
            if cell.column in block_cols and isinstance(cell.value, str) and cell.value.startswith(_GREY):
                cell.value = cell.value[len(_GREY):]
                cell.font = grey_font
            elif flag == _WITHDRAW_TEXT and cell.column in red_cols:
                cell.font = red_font
            elif flag == _CHECK_ENROL_TEXT and cell.column == flag_col:
                cell.font = amber_font


def to_workbook_bytes(df: pd.DataFrame, coach_view: pd.DataFrame | None = None) -> bytes:
    """Same as v2's, plus two bits of styling: an advice-block cell prefixed
    with the grey marker is written as grey italic text (for-reference, not
    registered), and a row flagged ``ADVISE WITHDRAWAL`` gets its flag and
    reason cells in red bold. Excluded students move from the Coach View to an
    ``Excluded`` sheet after it; the full sheet keeps every row."""
    import io

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        sheets = []
        if coach_view is not None:
            excl = _is_excluded(coach_view)
            coach_view[~excl].to_excel(writer, sheet_name=COACH_VIEW_SHEET, index=False)
            sheets.append(COACH_VIEW_SHEET)
            if excl.any():
                coach_view[excl].to_excel(writer, sheet_name=_EXCLUDED_TAB, index=False)
                sheets.append(_EXCLUDED_TAB)
        df.to_excel(writer, sheet_name=SHEET_NAME, index=False)
        sheets.append(SHEET_NAME)

        for name in sheets:
            _style_advice_sheet(writer.sheets[name])
    return buffer.getvalue()


# --------------------------------------------------------------------------- #
# Split the Coach View: one file per success coach, one tab per template       #
# --------------------------------------------------------------------------- #
COACH_COL = "Coach"
_NO_COACH = "(no coach)"
_NO_TEMPLATE = "(no template)"
_BAD_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")
_SUBJECT_CODE_RE = re.compile(r"\b[A-Z]{4}\d{4}\b")
_SUBJECT_NAMES_PATH = Path(__file__).with_name("subject_names.json")


def load_subject_names(path: Path | None = None) -> dict[str, str]:
    """``{subject_code: subject_name}`` from ``subject_names.json`` (``_comment``
    key dropped). Empty dict if the file is missing."""
    try:
        raw = json.loads((path or _SUBJECT_NAMES_PATH).read_text(encoding="utf-8"))
    except FileNotFoundError:
        return {}
    return {k: v for k, v in raw.items() if not k.startswith("_")}


def _name_subject_codes(text: str, names: dict[str, str]) -> str:
    """Rewrite every ``ABCD1234`` token in ``text`` to ``ABCD1234 — Name``.

    Leaves the grey marker prefix, ``+N elective`` placeholders, the ``and`` /
    ``,`` joiners and anything not in ``names`` untouched.
    """
    if not text or not names:
        return text
    return _SUBJECT_CODE_RE.sub(
        lambda m: f"{m.group(0)} — {names[m.group(0)]}" if m.group(0) in names else m.group(0),
        text,
    )


def _add_subject_names(cv: pd.DataFrame, names: dict[str, str] | None = None) -> pd.DataFrame:
    """Copy of ``cv`` with subject names appended to the codes in the five advice
    columns. No-op if the name list is empty."""
    names = load_subject_names() if names is None else names
    cv = cv.copy()
    if not names:
        return cv
    for col in [*ADVICE_COLS, *REGISTERED_COLS.values()]:
        if col in cv.columns:
            cv[col] = cv[col].map(lambda v: _name_subject_codes(v, names) if isinstance(v, str) else v)
    return cv


def _safe_filename(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9_-]+", "_", str(name)).strip("_") or "unnamed"


def _safe_sheet_name(name: str) -> str:
    return _BAD_SHEET_CHARS.sub("_", str(name)).strip()[:31] or "Sheet"


def split_coach_view_by_coach(coach_view: pd.DataFrame) -> dict[str, bytes]:
    """One ``.xlsx`` per success coach (the ``Coach`` column), and inside each
    one worksheet per ``Messaging Template`` the coach actually has students on.

    Coach View columns only, styled like the main workbook. A blank coach lands
    in a ``no_coach.xlsx`` file; a blank template in a ``(no template)`` sheet,
    so no student is silently dropped. The advice columns carry the subject name
    beside each code (from ``subject_names.json``). Returns ``{filename: xlsx_bytes}``.
    """
    import io

    if COACH_COL not in coach_view.columns:
        raise ValueError(f"No '{COACH_COL}' column in the Coach View — nothing to split by.")

    cv = _add_subject_names(coach_view)
    cv["_coach"] = cv[COACH_COL].fillna("").astype(str).str.strip().replace("", _NO_COACH)
    if TEMPLATE_COL in cv.columns:
        cv["_tmpl"] = cv[TEMPLATE_COL].fillna("").astype(str).str.strip().replace("", _NO_TEMPLATE)
    else:
        cv["_tmpl"] = _NO_TEMPLATE

    # Excluded students (no advice) leave every other tab, the Paused one
    # included, for a tab of their own.
    excl = _is_excluded(cv)
    cv.loc[excl, "_tmpl"] = _EXCLUDED_TAB

    # Paused students go on the "are you studying with us?" tab - the
    # check-enrolment worklist - instead of being scattered across the template
    # tabs with only a note at the tail of the reason text to tell them apart.
    #
    # Commencing paused students are a special case: the Commencing tab keeps the
    # whole intake (paused or not), so they STAY there and are ALSO repeated on
    # the Paused tab, which is the coach's complete outreach list. Every other
    # paused student is simply moved onto the Paused tab.
    if STUDY_PATH_COL in cv.columns:
        paused = cv[STUDY_PATH_COL].map(is_paused) & ~excl
        commencing_paused = paused & cv["_tmpl"].eq(_COMMENCING_TAB)
        # a copy of each paused commencing student for the Paused tab
        dup = cv[commencing_paused].copy()
        dup["_tmpl"] = _PAUSED_TAB
        # everyone else paused just moves onto the Paused tab
        cv.loc[paused & ~commencing_paused, "_tmpl"] = _PAUSED_TAB
        cv = pd.concat([cv, dup], ignore_index=True)

    out: dict[str, bytes] = {}
    used_files: dict[str, int] = {}
    for coach, group in cv.groupby("_coach", sort=True):
        fname = _safe_filename(coach)
        if fname in used_files:
            used_files[fname] += 1
            fname = f"{fname}_{used_files[fname]}"
        else:
            used_files[fname] = 1

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            used_sheets: set[str] = set()
            tabs = dict(tuple(group.groupby("_tmpl")))
            # Every coach file gets a Commencing tab, empty (headers only) if the
            # coach has no commencing students, so the files all look the same.
            if TEMPLATE_COL in cv.columns and _COMMENCING_TAB not in tabs:
                tabs[_COMMENCING_TAB] = group.iloc[0:0]
            # Commencing always opens the file, Excluded closes it; the rest
            # follow alphabetically.
            for tmpl in sorted(tabs, key=lambda t: (t != _COMMENCING_TAB, t == _EXCLUDED_TAB, t)):
                sub = tabs[tmpl]
                sheet = _safe_sheet_name(tmpl)
                base, n = sheet, 1
                while sheet.lower() in used_sheets:
                    n += 1
                    sheet = f"{base[:28]}_{n}"
                used_sheets.add(sheet.lower())
                sub.drop(columns=["_coach", "_tmpl"]).to_excel(writer, sheet_name=sheet, index=False)
                _style_advice_sheet(writer.sheets[sheet])
        out[f"{fname}.xlsx"] = buffer.getvalue()
    return out


def split_coach_view_zip_bytes(coach_view: pd.DataFrame) -> bytes:
    """``split_coach_view_by_coach`` packed into a single ``.zip``."""
    import io
    import zipfile

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in sorted(split_coach_view_by_coach(coach_view).items()):
            zf.writestr(name, data)
    return zip_buf.getvalue()


def read_summer_offering(source) -> set[str]:
    """Parse an uploaded Summer offering list -> the set of subject codes that
    run. Accepts .xlsx or .csv; any column whose cells look like subject codes
    (4 letters + 4 digits) is used."""
    import io
    import re as _re

    raw = source if hasattr(source, "read") else source
    try:
        frames = [pd.read_excel(io.BytesIO(raw.getvalue()) if hasattr(raw, "getvalue") else raw,
                                sheet_name=None)]
        cells = pd.concat(frames[0].values(), ignore_index=True).astype(str).values.ravel()
    except Exception:
        cells = pd.read_csv(raw).astype(str).values.ravel()
    return {c.strip().upper() for c in cells if _re.fullmatch(r"[A-Za-z]{4}\d{4}", c.strip())}


# Campus codes the College uses; a Summer offering row lists the campuses a
# subject runs at with these. Anything not one of these isn't read as a campus.
KNOWN_CAMPUSES = {"BK", "CA", "KW", "PC", "LP", "ON", "BL"}


def read_summer_offering_campus(source) -> dict[str, dict]:
    """Parse an uploaded Summer offering list -> ``{code: {"campuses", "block"}}``.

    Same lenient scan as ``read_summer_offering``, but per row: the subject code
    (4 letters + 4 digits) is mapped to whatever campus codes appear in the same
    row, plus the Summer block it runs in if the row names one (``SU1`` / ``SU2``
    / "Summer block 1" / "Summer 2" -> ``"SU1"`` / ``"SU2"``). A subject listed
    with no campus codes runs **everywhere** (all ``KNOWN_CAMPUSES``); with no
    block named, ``block`` is ``""``. So a bare code list still works.
    """
    import io
    import re as _re

    raw = source if hasattr(source, "read") else source
    try:
        sheets = pd.read_excel(io.BytesIO(raw.getvalue()) if hasattr(raw, "getvalue") else raw,
                               sheet_name=None)
        frame = pd.concat(sheets.values(), ignore_index=True)
    except Exception:
        frame = pd.read_csv(raw)

    camps_by: dict[str, set[str]] = {}
    block_by: dict[str, str] = {}
    for _, row in frame.astype(str).iterrows():
        line = " ".join(str(c) for c in row.values)
        # a cell may hold several tokens ("BK CA KW"), so split before matching
        tokens = [t.strip() for cell in row.values for t in _re.split(r"[\s,;/]+", str(cell))]
        codes = [t.upper() for t in tokens if _re.fullmatch(r"[A-Za-z]{4}\d{4}", t)]
        camps = {t.upper() for t in tokens if t.upper() in KNOWN_CAMPUSES}
        bm = _re.search(r"su\s*([12])\b|summer\s*(?:block\s*)?([12])\b", line, _re.I)
        block = f"SU{bm.group(1) or bm.group(2)}" if bm else ""
        for code in codes:
            camps_by.setdefault(code, set()).update(camps)
            if block and not block_by.get(code):
                block_by[code] = block
    return {
        code: {"campuses": (camps or set(KNOWN_CAMPUSES)), "block": block_by.get(code, "")}
        for code, camps in camps_by.items()
    }


# --------------------------------------------------------------------------- #
# Summer early advice - a targeting list, NOT the full Summer engine           #
# --------------------------------------------------------------------------- #
# An "early indicator" (out in SB3, before Summer offerings/results are final):
# which students could use a confirmed Summer subject to either get back on
# pattern (they failed an early subject) or finish sooner (near the end, one
# subject left). Deliberately narrow - not everyone is advised for Summer.
EARLY_GROUP_RESTORE = "Get back on pattern"
EARLY_GROUP_FINISH = "Finish sooner"
# Excluded students still match on subjects but aren't eligible to re-register,
# so they're kept off the contact list and put on a tab of their own.
EARLY_GROUP_EXCLUDED = "Excluded - no advice"
EARLY_GROUPS = (EARLY_GROUP_RESTORE, EARLY_GROUP_FINISH, EARLY_GROUP_EXCLUDED)
EARLY_OUTCOME_COL = "Progression Outcome"
EARLY_GROUP_COL = "Group"
EARLY_CATCHUP_COL = "Summer 1 catch-up"
EARLY_OUTSTANDING_COL = "Outstanding subjects"
_EARLY_MAX_OUTSTANDING = 2  # "off-pattern due to failing one or two subjects"


def summer_early_advice(
    df: pd.DataFrame, offering: dict[str, set[str]], session: str = "26 SUM",
) -> pd.DataFrame:
    """Shortlist of students a confirmed Summer offering could help.

    A candidate has **1-2 outstanding core subjects** (currently-registered
    counts as done, so fresh commencers with everything ahead are excluded) and
    at least one of those is **offered in Summer at their campus** (``offering``
    is ``{code: {campuses}}`` from :func:`read_summer_offering_campus`). Each is
    grouped: *Get back on pattern* when the catch-up subject is an early one
    (position 1-2, they're behind), else *Finish sooner* (a later subject, they
    are near the end). A student whose Progression Outcome is Exclusion goes in
    the *Excluded* group instead - listed for the coach, not advised. Each row
    carries the outcome (blanks labelled as in the Coach View's Study Status).
    Returns one row per candidate; empty frame if none.
    """
    names = load_subject_names()

    def campuses_of(code: str) -> set:
        entry = offering.get(code)
        return entry["campuses"] if isinstance(entry, dict) else (entry or set())

    def block_of(code: str) -> str:
        entry = offering.get(code)
        return entry.get("block", "") if isinstance(entry, dict) else ""

    def label(code: str) -> str:
        base = f"{code} — {names[code]}" if code in names else code
        blk = block_of(code)
        return f"{base} ({blk})" if blk else base

    rows = []
    for _, r in df.iterrows():
        program = str(r["PROGRAM_CD"]).split(".")[0]
        campus = "" if pd.isna(r.get("CAMP_CODE")) else str(r.get("CAMP_CODE")).strip()
        n_pos = 8 if program in calc.NURSING_PROGRAMS else 6
        pattern = calc.subjects_for(program, session)
        outstanding = [
            (pos, pattern.get(str(pos)))
            for pos in range(1, n_pos + 1)
            if pattern.get(str(pos)) and calc._outstanding_strict(r.get(f"Subject {pos} Status"))
        ]
        if not 1 <= len(outstanding) <= _EARLY_MAX_OUTSTANDING:
            continue
        catch = [(pos, code) for pos, code in outstanding
                 if code in offering and campus in campuses_of(code)]
        if not catch:
            continue
        outcome = "" if pd.isna(r.get("Progression Outcome")) else str(r.get("Progression Outcome")).strip()
        if outcome in v2.STANDING_NO_ADVICE:
            group = EARLY_GROUP_EXCLUDED
        else:
            group = EARLY_GROUP_RESTORE if min(p for p, _ in catch) <= 2 else EARLY_GROUP_FINISH
        _, template = calc.classify(r, program in calc.NURSING_PROGRAMS)
        rows.append({
            "STUDENT_ID": r["STUDENT_ID"],
            "FIRST_NAME": r.get("FIRST_NAME"), "LAST_NAME": r.get("LAST_NAME"),
            "PREFERRED_NAME": r.get("PREFERRED_NAME"),
            "INSTITUTION_EMAIL_ADDRESS": r.get("INSTITUTION_EMAIL_ADDRESS"),
            COACH_COL: r.get("Coach"), "PROGRAM_CD": program, "CAMP_CODE": campus,
            "COMMENCEMENT_PERIOD": r.get("COMMENCEMENT_PERIOD"),
            EARLY_OUTCOME_COL: _study_status(outcome, template, r.get(STUDY_PATH_COL)),
            EARLY_GROUP_COL: group,
            EARLY_CATCHUP_COL: ", ".join(label(c) for _, c in catch),
            EARLY_OUTSTANDING_COL: ", ".join(c for _, c in outstanding),
            "# outstanding": len(outstanding),
        })
    cols = ["STUDENT_ID", "FIRST_NAME", "LAST_NAME", "PREFERRED_NAME",
            "INSTITUTION_EMAIL_ADDRESS", COACH_COL, "PROGRAM_CD", "CAMP_CODE",
            "COMMENCEMENT_PERIOD", EARLY_OUTCOME_COL, EARLY_GROUP_COL, EARLY_CATCHUP_COL,
            EARLY_OUTSTANDING_COL, "# outstanding"]
    out = pd.DataFrame(rows, columns=cols)
    if len(out):
        order = {g: i for i, g in enumerate(EARLY_GROUPS)}
        out = (out.assign(_g=out[EARLY_GROUP_COL].map(order))
                  .sort_values(["_g", COACH_COL, "LAST_NAME"])
                  .drop(columns="_g").reset_index(drop=True))
    return out


def summer_early_by_coach_zip(shortlist: pd.DataFrame) -> bytes:
    """The early-advice shortlist as a ``.zip`` of one ``.xlsx`` per coach, a tab
    per group (Excluded on its own tab), so each SSC gets their own candidates."""
    import io
    import zipfile

    cv = shortlist.copy()
    cv["_coach"] = cv[COACH_COL].fillna("").astype(str).str.strip().replace("", _NO_COACH)
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for coach, group in cv.groupby("_coach", sort=True):
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                wrote = False
                for grp_name in EARLY_GROUPS:
                    sub = group[group[EARLY_GROUP_COL] == grp_name].drop(columns=["_coach"])
                    if len(sub):
                        sub.to_excel(writer, sheet_name=_safe_sheet_name(grp_name), index=False)
                        wrote = True
                if not wrote:
                    group.iloc[0:0].drop(columns=["_coach"]).to_excel(
                        writer, sheet_name="No candidates", index=False)
            zf.writestr(f"{_safe_filename(coach)}.xlsx", buffer.getvalue())
    return zip_buf.getvalue()


def build_advice(
    df: pd.DataFrame,
    session: str = DEFAULT_PLANNING_SESSION,
    offerings: dict | None = None,
    summer_subjects: set[str] | None = None,
) -> pd.DataFrame:
    """Copy of ``df`` with the advice + completion + reason + source columns."""
    offerings = offerings or v2.load_offerings()
    slot_map = v2.derive_slot_map(df)

    cols = [*ADVICE_COLS, COMPLETION_COL, PRINCIPLE_COL, TEMPLATE_COL,
            WITHDRAWAL_COL, REASON_COL, SOURCE_COL]
    out = df.copy()
    for col in cols:
        out[col] = ""

    for idx, row in df.iterrows():
        r = advise_student_merged(row, slot_map, offerings, session, summer_subjects)
        for col in cols:
            out.at[idx, col] = r[col]

    # A paused student's picks are provisional: what they WOULD take, not a
    # registration to action, so they are greyed the same way an already-running
    # block subject is. The subjects stay visible - a coach needs them for the
    # "are you coming back?" conversation.
    if STUDY_PATH_COL in out.columns:
        paused = out[STUDY_PATH_COL].map(is_paused)
        for col in ADVICE_COLS:
            out.loc[paused, col] = out.loc[paused, col].map(_grey_cell)

    from_block = rs.target_block(session)
    if from_block > 1:
        for b in range(from_block, 5):
            out[REGISTERED_COLS[b]] = [
                ", ".join(_SUBJECT_CODE_RE.findall(str(v or ""))) if not pd.isna(v) else ""
                for v in out.get(f"Block {b} code", pd.Series([""] * len(out), index=out.index))
            ]
        out[REG_CHECK_COL] = [
            _registration_check(row, session, from_block) for _, row in out.iterrows()
        ]
    return out


def _block_reg_check(advice, registered: list[str], pattern: set[str]) -> str:
    """One block: ``""`` when the registration matches the advice, else what
    the coach needs to do about it."""
    adv = "" if pd.isna(advice) else str(advice).strip()
    if not adv or adv.startswith(_GREY) or adv == calc.NO_REGISTRATION:
        return f"drop {', '.join(registered)}" if registered else ""
    if adv == "+1 elective":
        if not registered:
            return "register an elective"
        if all(c in pattern for c in registered):
            return f"check {', '.join(registered)} (a course subject, elective advised)"
        return ""
    wanted = _SUBJECT_CODE_RE.findall(adv)
    if not registered:
        return f"register {', '.join(wanted)}"
    if set(wanted) & set(registered):
        return ""
    return f"check: registered {', '.join(registered)}, advised {', '.join(wanted)}"


def _registration_check(row: pd.Series, session: str, from_block: int) -> str:
    """Mid-semester: does what the student is registered in for Blocks
    from_block..4 match the advice? ``OK``, or a short to-do per block.

    Any registered subject that isn't one of the program's own subjects or
    preps is taken to be an elective."""
    registered = {
        b: _SUBJECT_CODE_RE.findall(str(row.get(f"Block {b} code", "") or ""))
        for b in range(from_block, 5)
    }
    if str(row.get(SOURCE_COL, "")).startswith(_SRC_EXCLUDED):
        regs = [c for b in registered.values() for c in b]
        return f"Excluded - registered in {', '.join(regs)}" if regs else "Excluded"
    if is_paused(row.get(STUDY_PATH_COL)):
        return "Paused - confirm they're returning first"
    program = str(row["PROGRAM_CD"]).split(".")[0]
    ref = calc._ref().get(program, {})
    pattern = set(calc.subjects_for(program, session).values()) | {ref.get("prep1"), ref.get("prep2")}
    todo = [
        f"B{b}: {msg}" for b in range(from_block, 5)
        if (msg := _block_reg_check(row.get(ADVICE_COLS[b]), registered[b], pattern))
    ]
    return "; ".join(todo) if todo else REG_OK


def _other_enrolments(advised: pd.DataFrame) -> list[str]:
    """Per row: a note naming the student's OTHER enrolment(s), or ``""``.

    After the loader drops exact duplicates, a student with more than one row is
    genuinely in two programs - so a coach can see that someone else is advising
    the same person, and that one of the two may simply never have been
    withdrawn from.
    """
    if "STUDENT_ID" not in advised.columns:
        return [""] * len(advised)

    def col(name):
        return (advised[name] if name in advised.columns
                else pd.Series([""] * len(advised), index=advised.index))

    rows = list(zip(advised["STUDENT_ID"], col("PROGRAM_CD"), col(COACH_COL)))
    by_id: dict = {}
    for i, (sid, _, _) in enumerate(rows):
        by_id.setdefault(sid, []).append(i)

    notes = []
    for i, (sid, _, _) in enumerate(rows):
        others = [j for j in by_id[sid] if j != i]
        if not others:
            notes.append("")
            continue
        bits = []
        for j in others:
            _, prog, coach = rows[j]
            prog = "" if pd.isna(prog) else str(prog).strip()
            coach = "" if pd.isna(coach) else str(coach).strip()
            bits.append(f"program {prog}" + (f" ({coach})" if coach else "") if prog else (coach or "another program"))
        notes.append("Also enrolled: " + "; ".join(bits))
    return notes


def is_paused(status) -> bool:
    """True for a student who is enrolled but not currently studying (Deferred /
    Leave of Absence), read straight off ``STUDY_PATH_STATUS``.

    Anything that isn't the recognised active value counts as paused, so a new
    status in a future extract lands on the check-enrolment tab rather than
    being assumed reachable. A blank (or a file with no such column) is treated
    as active, leaving the output as it was before this column existed.
    """
    text = "" if pd.isna(status) else str(status).strip()
    return bool(text) and text != ACTIVE_STATUS


def _study_status(outcome, template, status=None) -> str:
    """``Progression Outcome`` -> a label with no silent blanks.

    A blank is only a *problem* for an active student who commenced before this
    session - they sat the progression round and should have a decision. A
    commencing student has no decision yet, and a Deferred / Leave-of-Absence
    student was never in the round, so neither is flagged for review.
    """
    text = "" if pd.isna(outcome) else str(outcome).strip()
    if text:
        return text
    if str(template or "").strip() == "Commencing":
        return NOT_ASSESSED
    if is_paused(status):
        return NOT_ASSESSED_PAUSED
    return NO_OUTCOME


def build_coach_view(advised: pd.DataFrame) -> pd.DataFrame:
    """v2's slim readable sheet, plus the Earliest Completion + Advice Source.

    The status grid uses the calculator's outstanding/elective tests (not
    v2's) so the ✓/✗ marks line up with the advice the calculator produced.
    """
    base = advised[[c for c in v2.COACH_VIEW_COLUMNS if c in advised.columns]].copy()

    outcomes = (
        advised["Progression Outcome"] if "Progression Outcome" in advised.columns
        else pd.Series([None] * len(advised), index=advised.index)
    )
    statuses = (
        advised[STUDY_PATH_COL] if STUDY_PATH_COL in advised.columns
        else pd.Series([None] * len(advised), index=advised.index)
    )
    base[STUDY_STATUS_COL] = [
        _study_status(o, t, s)
        for o, t, s in zip(outcomes, advised[TEMPLATE_COL], statuses)
    ]
    base[OTHER_ENROL_COL] = _other_enrolments(advised)

    status = [
        v2._status(r, is_outstanding=calc._is_outstanding, elective_count=calc._elective_count)
        for _, r in advised.iterrows()
    ]
    base["Student Status"] = [s for s, _, _ in status]
    base["Progress Bar"] = [b for _, _, b in status]
    base["Progress"] = [g for _, g, _ in status]

    # Advice Source (which engine ran) is kept in the full sheet for debugging
    # but left off the Coach View - a coach doesn't need it.
    # Mid-semester: each block still to register gets what the student is
    # actually registered in beside the advice, then the Registration Check.
    advice_cols: list[str] = []
    for i, col in enumerate(ADVICE_COLS):
        advice_cols.append(col)
        if REGISTERED_COLS.get(i) in advised.columns:
            advice_cols.append(REGISTERED_COLS[i])
    if REG_CHECK_COL in advised.columns:
        advice_cols.append(REG_CHECK_COL)
    for col in [WITHDRAWAL_COL, TEMPLATE_COL, PRINCIPLE_COL, *advice_cols, COMPLETION_COL, REASON_COL]:
        base[col] = advised[col].values
    return base
