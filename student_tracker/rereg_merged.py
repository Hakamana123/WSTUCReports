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
PRINCIPLE_COL = "Rereg Principle"
TEMPLATE_COL = "Messaging Template"

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


def _positional_fallback(row: pd.Series, program: str, is_nursing: bool) -> tuple[list[str], str]:
    """Advice for a fail pattern Grant's calculator has no row for.

    Diploma: each outstanding subject goes in the block it runs in (position P
    -> block ``((P - 1) % 4) + 1``); a block clash is won by the cohort (higher)
    position when the student has passed >= 2 subjects, else by the lower
    (restart). Nursing: the 8 subjects run strictly in sequence, so the first
    four outstanding ones fill Blocks 1-4 in number order. Electives fill empty
    Blocks 3 & 4 only. Prep 1 before prep 2.
    """
    prog_ref = calc._ref().get(program, {})
    subj = prog_ref.get("subjects", {})
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


def _is_owed(row: pd.Series, program: str, code: str) -> bool:
    """Is ``code`` a subject / prep this student still needs to pass?"""
    ref = calc._ref().get(program, {})
    for pos, c in ref.get("subjects", {}).items():
        if c == code:
            return calc._is_outstanding(row.get(f"Subject {pos} Status"))
    if code == ref.get("prep1"):
        return calc._is_outstanding(row.get("Prep 1 Status"))
    if code == ref.get("prep2"):
        return calc._is_outstanding(row.get("Prep 2 Status"))
    return False


_SRC_SUMMER = "Summer"
_GREY = "‹"  # advice-cell prefix -> styled grey (shown for reference, not registered now)


def _summer_advice(
    out: dict, row: pd.Series, program: str, is_nursing: bool, session: str, outcome: str,
    summer_subjects: set[str] | None = None,
) -> dict:
    """Summer advice.

    Preps and Subjects 1 & 2 are confirmed to run in Summer; anything else is
    unknown until a Summer offering list is uploaded (``summer_subjects``).
    Every outstanding subject is shown in the block it runs in - plain if it's
    being registered this Summer, **greyed** if it's only there for reference
    (not offered / unknown / didn't fit). Conditional Enrolment keeps the 30cp
    cap.
    """
    prog_ref = calc._ref().get(program, {})
    subj = prog_ref.get("subjects", {})
    capped = outcome in v2.STANDING_MAX_BLOCKS
    cp_cap = CE_CAP_CP if capped else 10 ** 6
    assumed = summer_subjects is None
    n_pos = 8 if is_nursing else 6

    def runs_in_summer(pos: int) -> bool:
        code = subj.get(str(pos))
        if summer_subjects is not None:
            return code in summer_subjects
        return pos in (1, 2)

    # preps first (always run in Summer, diplomas only)
    prep_now, cp = [], 0
    for slot, key in (("Prep 1 Status", "prep1"), ("Prep 2 Status", "prep2")):
        code = prog_ref.get(key)
        if not is_nursing and code and calc._is_outstanding(row.get(slot)) and cp + _CP_PREP <= cp_cap:
            prep_now.append(code)
            cp += _CP_PREP

    blocks = ["", "", "", ""]
    grey = [False, False, False, False]
    register: list[str] = []
    for bi in range(4):
        cands = [
            p for p in (bi + 1, bi + 5)
            if subj.get(str(p)) and calc._is_outstanding(row.get(f"Subject {p} Status"))
        ]
        if not cands:
            continue
        offered = [p for p in cands if runs_in_summer(p)]
        if offered and cp + _CP_MODULAR <= cp_cap:
            code = subj[str(offered[0])]
            blocks[bi] = code
            register.append(code)
            cp += _CP_MODULAR
        else:
            blocks[bi] = subj[str(cands[0])]
            grey[bi] = True

    out[ADVICE_COLS[0]] = " and ".join(prep_now)
    for col, val, g in zip(ADVICE_COLS[1:], blocks, grey):
        out[col] = (_GREY + val) if (val and g) else val

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

    bits = []
    if capped:
        bits.append(f"{outcome}: 30cp cap")
    if prep_now:
        bits.append("Prep: " + " and ".join(prep_now))
    if register:
        bits.append("Register: " + ", ".join(register))
    if not prep_now and not register:
        bits.append(f"Nothing confirmed for this student runs in {session}")
    if greyed:
        bits.append("Grey = not offered in Summer (or unknown) - take next Autumn: " + ", ".join(greyed))
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

    out = {c: "" for c in ADVICE_COLS}
    out[COMPLETION_COL] = ""
    out[REASON_COL] = ""
    out[SOURCE_COL] = ""
    out[PRINCIPLE_COL] = principle
    out[TEMPLATE_COL] = template

    # 1. Exclusion -> no advice, whichever engine would have run.
    if outcome in v2.STANDING_NO_ADVICE:
        out[REASON_COL] = f"{outcome} - not eligible to re-register; refer to coach."
        out[SOURCE_COL] = _SRC_EXCLUDED
        return out

    # 1b. Summer target -> preps + whatever the Summer offering list says runs
    #     (or just Subjects 1 & 2 when no list has been uploaded).
    tgt = rs.parse_target(session)
    if tgt and tgt[1] == "SUM":
        return _summer_advice(out, row, program, is_nursing, session, outcome, summer_subjects)

    # A part-way target ("26 AUT Block 3") uses the whole-session engines - the
    # picks are already locked to the block each subject runs in - and only
    # Blocks <from_block>..4 are actually registered now.
    base = rs.base_session(session)
    from_block = rs.target_block(session)

    # 2. Grant's calculator - only for the sessions it has offering patterns for.
    if rs.uses_calculator(session):
        c = calc.advise_row(row, base)
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

    # electives are not offered in Blocks 1 & 2 (applies to Grant's picks too)
    kept = _electives_to_b34(kept)
    elec_now = kept.count("+1 elective")

    # Part-way target: only Blocks >= from_block are registered now. Force each
    # remaining block to hold its own outstanding subject (the calculator
    # sometimes leaves a backlog subject unplaced), then keep the earlier
    # blocks displayed - the subjects that belong there.
    prog_subj = calc._ref().get(program, {}).get("subjects", {})
    if from_block > 1:
        for bi in range(from_block - 1, 4):
            for pos in (bi + 1, bi + 5):
                code = prog_subj.get(str(pos))
                if code and calc._is_outstanding(row.get(f"Subject {pos} Status")):
                    kept[bi] = code
                    break
        elec_now = kept.count("+1 elective")
    named_all = [(i + 1, b) for i, b in enumerate(kept) if b]
    if from_block > 1:
        named = [(n, b) for n, b in named_all if n >= from_block]
        partway_carry = [
            b for n, b in named_all
            if n < from_block and b != "+1 elective" and _is_owed(row, program, b)
        ]
    else:
        named = named_all
        partway_carry = []

    out[ADVICE_COLS[0]] = prep_now
    for i, (col, val) in enumerate(zip(ADVICE_COLS[1:], kept)):
        # part-way: Blocks before from_block are shown greyed (in progress)
        out[col] = (_GREY + val) if (val and from_block > 1 and i + 1 < from_block) else val
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
    if prep_now:
        bits.append(f"Prep: {prep_now}")
    if prep_summer:
        bits.append(f"Prep in Summer: {prep_summer}")
    if named:
        bits.append("Register: " + ", ".join(f"Block {n} {b}" for n, b in named))
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
    accounted = {b for _, b in named} | set(deferred) | set(partway_carry) | {prep_now, prep_summer}
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

    # Earliest completion: Grant's own date for 26 AUT; a rough projection
    # (this session + ceil(work left / load) more sessions) for everything else.
    mod_out = sum(calc._is_outstanding(row.get(f"Subject {i} Status")) for i in range(1, 9))
    prep_out = sum(calc._is_outstanding(row.get(f"Prep {i} Status")) for i in (1, 2))
    sched_mod = sum(1 for _, b in named if b != "+1 elective")
    load = 3 if capped else 4
    mod_left = max(0, (mod_out - sched_mod) + (elec_need - elec_now))
    prep_left = max(0, prep_out - len(_split_prep(prep_now)))
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
COACH_VIEW_SHEET = v2.COACH_VIEW_SHEET
SHEET_NAME = v2.SHEET_NAME


def to_workbook_bytes(df: pd.DataFrame, coach_view: pd.DataFrame | None = None) -> bytes:
    """Same as v2's, but any advice-block cell prefixed with the grey marker
    (``_GREY``) is written as grey italic text - "shown for reference, not
    being registered this session"."""
    import io
    from openpyxl.styles import Font

    grey_font = Font(color="808080", italic=True)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        sheets = []
        if coach_view is not None:
            coach_view.to_excel(writer, sheet_name=COACH_VIEW_SHEET, index=False)
            sheets.append(COACH_VIEW_SHEET)
        df.to_excel(writer, sheet_name=SHEET_NAME, index=False)
        sheets.append(SHEET_NAME)

        for name in sheets:
            ws = writer.sheets[name]
            block_cols = {
                i for i, c in enumerate(next(ws.iter_rows(max_row=1, values_only=True)), 1)
                if c in ADVICE_COLS[1:]
            }
            for rowcells in ws.iter_rows(min_row=2):
                for cell in rowcells:
                    if cell.column in block_cols and isinstance(cell.value, str) and cell.value.startswith(_GREY):
                        cell.value = cell.value[len(_GREY):]
                        cell.font = grey_font
    return buffer.getvalue()


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


def build_advice(
    df: pd.DataFrame,
    session: str = DEFAULT_PLANNING_SESSION,
    offerings: dict | None = None,
    summer_subjects: set[str] | None = None,
) -> pd.DataFrame:
    """Copy of ``df`` with the advice + completion + reason + source columns."""
    offerings = offerings or v2.load_offerings()
    slot_map = v2.derive_slot_map(df)

    cols = [*ADVICE_COLS, COMPLETION_COL, PRINCIPLE_COL, TEMPLATE_COL, REASON_COL, SOURCE_COL]
    out = df.copy()
    for col in cols:
        out[col] = ""

    for idx, row in df.iterrows():
        r = advise_student_merged(row, slot_map, offerings, session, summer_subjects)
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

    # Advice Source (which engine ran) is kept in the full sheet for debugging
    # but left off the Coach View - a coach doesn't need it.
    for col in [TEMPLATE_COL, PRINCIPLE_COL, *ADVICE_COLS, COMPLETION_COL, REASON_COL]:
        base[col] = advised[col].values
    return base
