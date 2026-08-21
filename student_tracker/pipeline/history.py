"""
PAST SUBJECT HISTORY — reads the "AB4 for SPR - Reregistration List" style
file (Grant's actual generated advisory output, discovered 2026-08-21) and
turns its "All Subjects" bracket string into subject-level pass/fail data.

This is a SEPARATE input from Ashlee's roster (ingest.py). The roster is
forward-looking only (what's registered for next semester); this file is
the first source we've seen with actual past pass/fail history. The two
are combined by student_id in report_builder.py, sitting alongside the
existing bucket/on-pattern output rather than replacing it - Josiah
confirmed the real "Rereg Principles" taxonomy this file carries should
sit alongside the bucketing.py taxonomy, not replace it (2026-08-21).

Two "All Subjects" shapes are confirmed, both empirically validated
against real files (v1.1 2026-08-21, v2.1 2026-08-22 - the workbook's
sheet names changed between the two, see load_reregistration_history):

DIPLOMA shape (programs 7188-7198) - 4 bracket groups, sizes (2, 4, 2, 1):
    [P1+P2] + [S1+S2+S3+S4] + [T1+T2] + [E]
  - Group 1 (2 digits): the two general prep subjects, GEDU0016 and
    GEDU0017 - NOT the same as GEDU1001/GEDU1002, which are ordinary
    pattern positions 1-2 for some diploma programs. These two preps
    aren't in pattern_table.json at all.
  - Group 2 (4 digits): Semester 1's 4 core blocks - lines up exactly
    with pattern_table.json positions 1-4.
  - Group 3 (2 digits): Semester 2's remaining 2 core blocks - pattern
    positions 5-6.
  - Group 4 (1 digit): NOT pass/fail. Confirmed empirically (cross-
    tabulated against the real file's B3 Name/B4 Name columns) to be a
    COUNT (0, 1, or 2) of electives still outstanding.

NURSING/UPP shape (program 9031) - 2 bracket groups, sizes (4, 4):
    [S1+S2+S3+S4] + [S5+S6+S7+S8]
  No separate prep group and no elective-count group - confirmed by the
  real v2.1 file's Nursing sheet always having a blank "Prep" column,
  consistent with 9031 having no separate prep subjects at all (its 8
  pattern-table positions are all named subjects, e.g. LANG0051). The two
  groups are pattern positions 1-4 and 5-8, matching 9031's 8-position
  pattern_table.json entry exactly.

Both shapes: each pass/fail digit is 0 = passed, 1 = still needs to
enrol/pass. Any other bracket-group count/size is an unrecognized shape -
parse_all_subjects raises rather than guess.

2026-08-22: the diploma shape's pass/fail decoding (not the elective-count
digit) was independently ground-truth-checked against a different real
file - "2026 AUT Stage 5 Reregistration Advice List v1.1.xlsx", which
decodes the same "All Subjects" string into explicit Subject 1-8
Status/Prep 1-2 Status columns (e.g. "HUMN1070 Completed" vs. a bare 1).
parse_all_subjects's output matched all 387 diploma rows exactly (one row
showed a raw 0 instead of a resolved subject name for program 7191
position 5 - a lookup glitch in that spreadsheet's own formula, not a
decoding disagreement). This is a stronger check than the original v1.1
reverse-engineering, since it's ground truth from a second, independently
generated file rather than inference from B-Name columns. The
elective-count digit's meaning wasn't re-confirmed by this file - its
correlation with "elective" wording in B2-B4 Name was noticeably weaker
there, plausibly because that file has no B1 Name column at all (see
stages.py's Stage 5 note) so elective advice that would show in B1 isn't
visible to check against.

2026-08-22: this file matches AUT STAGE 2 of the Modular Re-registration
Timeline (stages.py) on every checkable point - its filename is exactly
Stage 2's documented example ("2026 AB4 for 2026 SPR"), its population
spans ALL standings (Good Standing/At Risk/Conditional Enrolment/
Exclusion, confirmed by checking Progression Outcome's value counts) not
just Conditional Enrolment, and every row's "Advice for when" reads "26
SPR" - i.e. generated after Autumn's results, advising for the coming
Spring. That's Stage 2's population and timing exactly, not Stage 5's
(which is CE-only, confirmed separately via the different Stage 5 file -
see stages.py). Unlike Stage 5, which our bucketing.py computes from
scratch, we don't know Grant's actual rules for deriving Rereg
Principles/Template from pass/fail history - only the shape of his
finished output. So this module's role for Stage 2 is READING/
DISPLAYING his already-computed recommendation (see
pages/10_Stage2_Recommendations.py), not reproducing his logic the way
bucketing.py does for Stage 5.

2026-08-22, confirmed directly by Josiah (first real classification-rule
evidence, not just structural inference):
  - GEDU0016 and GEDU0017 are NOT interchangeable/concurrent - GEDU0016 is
    the Semester 1 prep, GEDU0017 is the Semester 2 prep. A student can't
    have passed both by the time a Stage 2 (post-AB4, pre-SPR) file is
    generated, since GEDU0017 hasn't run yet. So prep_passed[1] (and,
    generally, any sem2_passed digit) showing as "not yet passed" at this
    checkpoint means NOT YET DUE, not failed - only sem1_passed digits are
    genuine pass/fail signals at this point in the calendar. (This is why
    failed_positions/glossary wording already says "not yet passed"
    rather than "failed" - that hedge turns out to be load-bearing, not
    just cautious phrasing.)
  - Confirmed remediation shape for a student who fails Sem1 Block 3 AND
    4 (passed prep-to-date and Blocks 1-2): the following semester's
    registration is NOT the normal 4-block pattern. Spring Block 1-2 =
    the normal pattern positions 5-6; Spring Block 3-4 = the FAILED
    Semester 1 Block 3/4 subjects moved into the back half of next
    semester as catch-up, alongside GEDU0017 proceeding as normal. This
    is a genuinely different registration shape than what
    pattern_lookup.compare_registration_to_pattern currently checks (it
    only compares Spring Block 1-2 against positions 5-6 for diplomas) -
    a student in exactly this situation would look "off pattern" to our
    Stage 5 bucketing today even though this registration is the CORRECT
    catch-up shape. Not yet acted on in bucketing.py/pattern_lookup.py -
    flagging here as a known gap this scenario exposed.
  - Confirmed Rereg Principles data point: passed prep-to-date, passed
    Sem1 Blocks 1-2, failed Sem1 Blocks 3-4, first semester (no prior
    session history) -> "Unsatisfactory progress in S1". Academic
    standing (GS/AR/CE) was explicitly NOT predictable from this pattern
    alone per Josiah - Rereg Principles classification appears to be
    driven primarily by the pass/fail pattern's shape/timing, not
    directly coupled to Calculated Standing.
  - Confirmed "3+ Sessions" is NOT a pure tenure/duration label ("this
    happens to be their 3rd+ semester") - it requires an actual pattern
    of struggling ACROSS multiple sessions. The identical fail pattern
    above (2 of 4 Sem1 blocks failed), if it belongs to a continuing
    student now on their 3rd+ semester rather than a first-timer, does
    map to "3+ Sessions" instead of "Unsatisfactory progress in S1" -
    but per Josiah's reasoning, that's because being on a 3rd+ attempt
    while still failing blocks is itself evidence of cross-session
    struggle (and correlates with, but isn't guaranteed to be, AR/CE
    standing), not because session-count alone overrides the pattern.
    So the real rule needs BOTH signals together: which session number
    this is for the student, AND whether the current pattern shows
    struggle - neither alone determines the category. We don't yet have
    a clean "which session number is this" signal wired into history.py
    (commencement_period vs. the file's own period is the closest proxy
    but hasn't been validated for this purpose).
  - Confirmed the failure-COUNT boundary between "Mostly Progressing" and
    "Unsatisfactory progress in S1" for a first-semester student: passed
    prep-to-date, passed Sem1 Blocks 1-2, failed only ONE of the
    remaining two Sem1 blocks -> "Mostly Progressing", not
    "Unsatisfactory progress in S1". Josiah's stated reasoning was
    recoverability ("the student can quite easily catch up"), not a
    rigid count rule per se - but for THIS specific block-4 pattern, the
    working threshold is: 1 of 4 Sem1 blocks failed = Mostly Progressing,
    2 of 4 failed = Unsatisfactory progress in S1 (first semester) / 3+
    Sessions (continuing, struggling pattern). Not yet tested: whether
    which specific block failed matters (e.g. an earlier block being a
    prerequisite for a later one), or whether 3-4 failures tips into a
    more severe category like "Overall lack of success".
  - Confirmed "Overall lack of success" for a FIRST-semester student who
    fails everything (GEDU0016 prep AND all 4 Sem1 blocks) - so this
    category is NOT reserved for continuing/multi-session strugglers the
    way "3+ Sessions" is; it can trigger in session 1 given total (or
    near-total) failure. This means severity-within-the-current-session
    and persistence-across-sessions look like two separate axes feeding
    Rereg Principles, not one ladder ordered purely by session count -
    e.g. a mild single-session stumble (Mostly Progressing) and a total
    single-session wipeout (Overall lack of success) are both possible
    in session 1, at opposite ends of severity, with no session-count
    component needed for either. Confirmed action: "ask her to re-enrol
    and try again" - the real-world equivalent of bucketing.py's
    reapply_next_semester bucket (commencing, below Good Standing, zero
    registration), suggesting our own invented taxonomy and Grant's real
    one converge on similar remediation for this case even though the
    category NAMES and trigger conditions differ.
  - PARTIALLY CONFIRMED, needs follow-up: whether the fail-COUNT
    threshold for a category tightens the further along a student is
    (e.g. a session-3+ student with just ONE outstanding subject already
    counting as "Overall lack of success", vs. session 1 needing near-
    total failure for that same label) was floated but not directly
    confirmed or denied by Josiah - his reply ("likely to be CE") only
    confirmed that a session-3+ student still carrying failures tends
    toward Conditional Enrolment standing, not the specific fail-count
    question asked. Treat the "thresholds tighten with session number"
    idea as a live hypothesis, not a confirmed rule, until asked again
    directly with a cleaner yes/no framing.
"""

from __future__ import annotations
from dataclasses import dataclass
from typing import Optional
import re

import pandas as pd

from .pattern_lookup import lookup_pattern

PREP_SUBJECT_CODES = ["GEDU0016", "GEDU0017"]

# Sheet names seen across file versions that hold the FULL population for
# that program group - not the "Bulk"/"Complete"/"Conditional Enrol"/
# "Exclusion"/"No Subject Avail" sheets, which are filtered subsets of
# these (confirmed by row counts summing to match, v2.1 2026-08-22) and
# would double-count students if also read.
_FULL_SHEET_NAMES = ["Diplomas - All", "All Diplomas", "All Nursing", "Nursing - All"]

REQUIRED_COLUMNS = [
    "student_id", "program", "commencement_period", "all_subjects",
]


@dataclass
class SubjectHistory:
    student_id: str
    student_name: Optional[str]
    program: int
    commencement_period: Optional[str]
    prep_passed: list       # matches PREP_SUBJECT_CODES order; [] for the Nursing/UPP shape (no prep group)
    sem1_passed: list       # length 4, pattern positions 1-4 (both shapes)
    sem2_passed: list       # length 2 (diploma, positions 5-6) or 4 (Nursing/UPP, positions 5-8)
    electives_outstanding: Optional[int]   # None for the Nursing/UPP shape (no elective group)
    rereg_principle: Optional[str]
    template: Optional[str]
    block_advice: list      # length 4, B1-B4 Name, next semester's advice
    prep_advice: Optional[str]   # "Prep Name" (diploma) / "Prep" (Nursing, always None) - Stage 3's recommendation

    @property
    def failed_positions(self) -> list:
        """Pattern-table positions not yet passed (1-6 for diploma shape,
        1-8 for Nursing/UPP shape - sem2_passed's length determines the
        upper bound generically). The two prep subjects (diploma shape
        only) aren't pattern-table positions, so a failed prep doesn't
        show up here; check prep_passed separately if needed.
        """
        offset = len(self.sem1_passed)
        failed = [i for i, passed in enumerate(self.sem1_passed, start=1) if not passed]
        failed += [i for i, passed in enumerate(self.sem2_passed, start=offset + 1) if not passed]
        return failed


def failed_subject_codes(history: SubjectHistory, pattern_overrides: Optional[dict] = None) -> str:
    """Resolve a SubjectHistory's failed_positions to subject codes where
    the pattern table covers them, falling back to "Position N" for
    positions it doesn't. "(none)" when nothing is outstanding. Used by
    pages/10_Stage2_Recommendations.py and Stage 3's equivalent page so
    they don't drift into different wording for the same thing.
    """
    if not history.failed_positions:
        return "(none)"
    pattern = lookup_pattern(history.program, history.commencement_period, pattern_overrides)
    parts = [
        pattern.block_sequence.get(pos, f"Position {pos}")
        for pos in history.failed_positions
    ]
    return "; ".join(parts)


def parse_all_subjects(value: str) -> tuple:
    """Parse an "All Subjects" bracket string into (prep, sem1, sem2,
    electives_outstanding). prep/sem1/sem2 are lists of bool (True =
    passed). Recognizes both confirmed shapes (see module docstring):
    4 groups sized (2, 4, 2, 1) for diplomas, or 2 groups sized (4, 4)
    for Nursing/UPP (9031) - prep=[] and electives_outstanding=None for
    the latter, since it has neither. Raises ValueError for any other
    shape rather than guess.
    """
    groups = re.findall(r"\[([^\]]*)\]", str(value))
    try:
        digit_groups = [[int(d) for d in g.split("+")] for g in groups]
    except ValueError as e:
        raise ValueError(f"Non-numeric value in All Subjects string {value!r}") from e
    actual_lengths = [len(g) for g in digit_groups]

    if actual_lengths == [2, 4, 2, 1]:
        prep_digits, sem1_digits, sem2_digits, elective_digits = digit_groups
        return (
            [d == 0 for d in prep_digits],
            [d == 0 for d in sem1_digits],
            [d == 0 for d in sem2_digits],
            elective_digits[0],
        )

    if actual_lengths == [4, 4]:
        sem1_digits, sem2_digits = digit_groups
        return (
            [],
            [d == 0 for d in sem1_digits],
            [d == 0 for d in sem2_digits],
            None,
        )

    raise ValueError(
        f"All Subjects value {value!r} has bracket group sizes "
        f"{actual_lengths}, which doesn't match either confirmed shape - "
        "(2, 4, 2, 1) for diplomas or (4, 4) for Nursing/UPP. See "
        "history.py's module docstring."
    )


def load_reregistration_history(path, sheet_name=None) -> pd.DataFrame:
    """Load the raw "AB4 for SPR" workbook into a single normalized
    DataFrame. If sheet_name is None (the normal case for the real
    multi-sheet workbook), reads every sheet in _FULL_SHEET_NAMES that's
    actually present and concatenates them - covers diplomas + Nursing/UPP
    in one upload without double-counting the "Bulk"/"Complete"/etc.
    breakdown sheets. Pass an explicit sheet_name to read just one sheet
    (e.g. a single-sheet file like the sample data).
    """
    if sheet_name is not None:
        sheet_names = [sheet_name]
    else:
        available = pd.ExcelFile(path).sheet_names
        sheet_names = [s for s in _FULL_SHEET_NAMES if s in available]
        if not sheet_names:
            raise ValueError(
                "Couldn't find any of the expected full-population sheets "
                f"({', '.join(_FULL_SHEET_NAMES)}) in this workbook. This "
                f"file's actual sheets are: {available}."
            )

    rename_map = {
        "STUDENT_ID": "student_id",
        "FIRST_NAME": "first_name",
        "LAST_NAME": "last_name",
        "PREFERRED_NAME": "preferred_name",
        "PROGRAM_CD": "program",
        "COMMENCEMENT_PERIOD": "commencement_period",
        "All Subjects": "all_subjects",
        "Rereg Principles": "rereg_principle",
        "Template": "template",
        "B1 Name": "b1_name",
        "B2 Name": "b2_name",
        "B3 Name": "b3_name",
        "B4 Name": "b4_name",
        "Prep Name": "prep_advice",
        "Prep": "prep_advice",
    }

    frames = []
    for name in sheet_names:
        df = pd.read_excel(path, sheet_name=name)
        original_columns = list(df.columns)
        df = df.rename(columns=rename_map)
        missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
        if missing:
            raise ValueError(
                f"Sheet {name!r} is missing expected column(s) after "
                f"normalization: {missing}. This sheet's actual headers "
                f"are: {original_columns}."
            )
        frames.append(df)
    return pd.concat(frames, ignore_index=True)


def to_history_records(df: pd.DataFrame) -> list:
    """Convert a normalized history DataFrame into SubjectHistory objects.
    Rows whose All Subjects value can't be parsed (an unrecognized shape -
    see module docstring) are skipped rather than raising, since a
    multi-sheet upload shouldn't fail entirely over rows this module
    doesn't yet understand.
    """
    records = []
    for _, row in df.iterrows():
        try:
            prep, sem1, sem2, electives = parse_all_subjects(row.get("all_subjects"))
        except ValueError:
            continue

        commencement = row.get("commencement_period")
        commencement = None if pd.isna(commencement) else str(commencement)

        block_advice = [
            None if pd.isna(row.get(col)) else str(row.get(col))
            for col in ["b1_name", "b2_name", "b3_name", "b4_name"]
        ]

        preferred = row.get("preferred_name")
        first = row.get("first_name")
        last = row.get("last_name")
        given_name = preferred if not pd.isna(preferred) else first
        name_parts = [str(p) for p in (given_name, last) if not pd.isna(p)]
        student_name = " ".join(name_parts) if name_parts else None

        records.append(
            SubjectHistory(
                student_id=str(row.get("student_id")),
                student_name=student_name,
                program=int(row.get("program")),
                commencement_period=commencement,
                prep_passed=prep,
                sem1_passed=sem1,
                sem2_passed=sem2,
                electives_outstanding=electives,
                rereg_principle=(
                    None if pd.isna(row.get("rereg_principle")) else str(row.get("rereg_principle"))
                ),
                template=(
                    None if pd.isna(row.get("template")) else str(row.get("template"))
                ),
                block_advice=block_advice,
                prep_advice=(
                    None if pd.isna(row.get("prep_advice")) else str(row.get("prep_advice"))
                ),
            )
        )
    return records


def history_by_student_id(records: list) -> dict:
    """Index SubjectHistory records by student_id."""
    return {r.student_id: r for r in records}
