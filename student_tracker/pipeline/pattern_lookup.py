"""
Stage 2 of 4 — PATTERN LOOKUP
==============================
STATUS: PARTIALLY RESOLVED. Runs for real on a narrow, confirmed slice;
everything else still falls through to an unresolved placeholder.

Goal: given (program, commencement_period), return the student's Pattern
of Study — the expected subject at each block position.

`pattern_table.json` (in this directory) is extracted from the `REF` sheet
of the 2025 SUM for 2026 AUT testing calculator (Grant's working file).
It contains only program/session/subject codes — no student data.

What's resolved (2026-08-07, confirmed with Josiah):
  - "26 - Autumn" (bare) and "26 - Autumn Block 1" are the same
    commencement — both normalize to the "26 AUT" session key.
  - For that session, positions 1-4 map to Autumn Block 1-4
    respectively (Grant, scoping call: "you'll have subjects 1, 2,
    3, 4 running in blocks 1, 2, 3, 4"). A "-" in the reference pattern
    means the subject doesn't run that session at all (see the 7197
    case in the same call), not "not yet resolved".

What's still open (Grant question #1, unchanged) — these all fall
through to `resolved=False` rather than being guessed:
  - Block 2/3/4 commencers: unconfirmed whether they follow a shifted
    version of the same sequence, or something else entirely.
  - Spring: zero rows exist in the reference table.
  - Program 9034: zero rows exist in the reference table. Also now a
    live discrepancy, not just a gap - the 2026 handbook's 9034 is a
    tiny 20cp/2-subject "Applied Policing" prep program that doesn't
    match the 81-127 students under this code in the roster at all.
    Needs Grant/registrar confirmation of what 9034 actually was for
    those students before touching this at all.
  - "25 SUM" / "25 SB4": the reference table stores these as two
    dual-bracket lists joined with "+" (e.g.
    "[1, 2, -, -, -, -] + [1, -, 3, 4, -, -]") — meaning unconfirmed,
    not modeled.
  - Any commencement code that isn't a plain "YY - Season[ Block N]"
    string (e.g. "SC1, WSTC T1" codes) — these are legacy, out of
    scope (see project memory).

2026-08-21: Program 9031 (nursing/UPP) resolved via the 2026 student
handbook (studenthandbook.westernsydney.edu.au), not Grant. The
handbook shows there's no Prep/Subject split at all for 9031 - it's 8
subjects total, "taught in 4-Week Block sessions on Campus - one
subject per Block session," across 2 semesters (4 blocks each), no
prep subject. Positions 5-8 (Spring Block 1-4) are empirically
validated at 100% against real GS-standing 26-Autumn-Block-1 commencers
(280-283 students each, zero mismatches) - stronger evidence than the
diploma position-5/6 mapping. Positions 1-4 (Autumn) use the same
handbook-listed order but aren't independently checkable (no Autumn
block-level registration data in either roster) - trust is inherited
from the confirmed back half, not separately verified.

7197 (Diploma in Education Studies): the 2026 handbook lists a 6th
required core subject, EDUC1012 (Literacy and Numeracy for Educators),
that isn't in this table at all - position 4 is blank here on the
strength of a direct Grant quote ("the 7197 case," see above), but the
handbook contradicts that framing. Left unresolved intentionally - this
needs to go back to Grant, not be silently overwritten either way.

7188 (Diploma, same 6-subject pattern as 7198) is discontinued for new
2026 intake (404s on the current handbook) but this table's entry is
left as-is: since it was already identical to 7198's, it should still
be the right teach-out pattern for continuing 7188 students, not a gap
- flagging the assumption here rather than treating it as still open.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import Optional
import json
import re
from pathlib import Path

_TABLE_PATH = Path(__file__).parent / "pattern_table.json"
_SEASON_ABBR = {"Autumn": "AUT", "Spring": "SPR", "Summer": "SUM"}
_COMMENCEMENT_RE = re.compile(r"^(\d{2}) - (Autumn|Spring|Summer)(?:\s+Block\s+(\d+))?$")

# Sessions that exist in the reference table but use notation this stage
# doesn't understand yet (dual-bracket "+"-joined patterns for Summer).
_UNMODELED_SESSIONS = {"25 SUM", "25 SB4"}
# Programs whose position ordering isn't confirmed even where a row exists.
# 9031 was here until 2026-08-21 - resolved via the handbook + empirical
# validation, see module docstring. Left as an empty set (rather than
# removed) since the mechanism is still needed for any future case like it.
_UNCONFIRMED_ORDER_PROGRAMS = set()


@dataclass
class PatternOfStudy:
    program: int
    commencement_period: str
    # Position (1-based, as int) -> expected subject code. Empty until
    # resolved — see module docstring for exactly what's covered.
    block_sequence: dict = field(default_factory=dict)
    resolved: bool = False
    unresolved_reason: Optional[str] = None


def _load_table() -> dict:
    with open(_TABLE_PATH) as f:
        return json.load(f)


_PATTERN_TABLE: dict = _load_table()


def _normalize_session(commencement_period: str) -> tuple[Optional[str], Optional[str]]:
    """Return (session_key, unresolved_reason). Exactly one is None.

    Only bare ("26 - Autumn") and explicit Block 1 ("26 - Autumn Block 1")
    commencements are normalized — that's the only case confirmed so far.
    """
    match = _COMMENCEMENT_RE.match(commencement_period.strip())
    if not match:
        return None, (
            f"Commencement code {commencement_period!r} doesn't match the plain "
            "'YY - Season[ Block N]' format (e.g. it may be a legacy SC1/WSTC "
            "T1 code) — not covered by the reference table."
        )

    year, season, block = match.groups()
    if block is not None and block != "1":
        return None, (
            f"Block {block} commencement — unconfirmed whether this follows a "
            "shifted version of the Block 1 pattern or a different one "
            "entirely."
        )

    return f"{year} {_SEASON_ABBR[season]}", None


def lookup_pattern(program: int, commencement_period: Optional[str]) -> PatternOfStudy:
    """Return the Pattern of Study for a given program + commencement period.

    Resolved only for the narrow slice described in the module docstring;
    everything else returns `resolved=False` with a specific reason rather
    than a guess.
    """
    if commencement_period is None:
        return PatternOfStudy(
            program=program,
            commencement_period="",
            resolved=False,
            unresolved_reason="No commencement period on record for this student "
                               "- how these should be handled is still an open "
                               "question.",
        )

    session, reason = _normalize_session(commencement_period)
    if session is None:
        return PatternOfStudy(
            program=program, commencement_period=commencement_period,
            resolved=False, unresolved_reason=reason,
        )

    if session in _UNMODELED_SESSIONS:
        return PatternOfStudy(
            program=program, commencement_period=commencement_period,
            resolved=False,
            unresolved_reason=f"{session} exists in the reference table but uses a "
                               "dual-bracket '+'-joined pattern that isn't understood "
                               "yet.",
        )

    if program in _UNCONFIRMED_ORDER_PROGRAMS:
        return PatternOfStudy(
            program=program, commencement_period=commencement_period,
            resolved=False,
            unresolved_reason=f"Program {program} has an 8-position reference pattern "
                               "but the Prep/Subject ordering isn't confirmed yet.",
        )

    positions = _PATTERN_TABLE.get(session, {}).get(str(program))
    if positions is None:
        return PatternOfStudy(
            program=program, commencement_period=commencement_period,
            resolved=False,
            unresolved_reason=f"No reference table row for program {program} in "
                               f"session {session} yet.",
        )

    return PatternOfStudy(
        program=program,
        commencement_period=commencement_period,
        block_sequence={int(k): v for k, v in positions.items()},
        resolved=True,
    )


# --- Pattern-vs-registration comparison ---------------------------------
# block_registrations index -> pattern position. pattern_lookup only
# confirms positions 1-4 = Autumn Block 1-4 (the semester a student
# *started* in); neither roster has Autumn block-level registration data,
# only Spring Block 1-4 (the *next* semester).
#
# Diplomas: positions 5-6 are the candidate mapping to Spring Block 1-2
# (diploma core patterns only have 6 positions - the remaining 2 Spring
# slots are free-choice electives per the 2026 handbook, not a fixed
# lookup, so nothing to check there). Empirically validated 2026-08-21
# against real GS-standing students: 76% match on position 5 vs Spring
# Block 1 (1076/1422 in the Ashlee sample), 98% on position 6 vs Spring
# Block 2 (1393/1420). Mismatches concentrate in students still
# registered for an earlier-position subject (e.g. repeating GEDU1001) -
# i.e. they look like genuinely off-pattern students, not evidence
# against the mapping. Validated-but-unconfirmed, not Grant-confirmed -
# worth a Grant question, but a much stronger footing than a guess.
#
# 9031 (UPP) has no electives at all - all 8 positions are fixed, so
# positions 7-8 map to Spring Block 3-4 too. Confirmed at 100% (280-283
# students each, zero mismatches) - see pattern_table.json / module
# docstring. Harmless no-op for diplomas: they have no position 7/8, so
# compare_registration_to_pattern's "skip if position not in pattern"
# check just skips these entries for them, same as before this changed.
SPRING_BLOCK_TO_POSITION = {0: 5, 1: 6, 2: 7, 3: 8}
BLOCK_LABELS = ["Spring Block 1", "Spring Block 2", "Spring Block 3", "Spring Block 4"]


@dataclass
class PatternComparison:
    status: str                        # "Y", "N", "Partial", "Unknown"
    advised: str                       # formatted expected Spring Block 1-2 subjects
    reason: Optional[str] = None       # unresolved reason, only set when status=="Unknown"


def compare_registration_to_pattern(
    program: int, commencement_period: Optional[str], block_registrations: list
) -> PatternComparison:
    """Compare a student's actual Spring Block 1-2 registration against
    what the pattern expects there (see SPRING_BLOCK_TO_POSITION above for
    why only blocks 1-2 are checkable). Shared by bucketing.py and
    report_builder.py so the two stay consistent rather than drifting.
    """
    pattern = lookup_pattern(program, commencement_period)
    if not pattern.resolved:
        return PatternComparison(status="Unknown", advised="", reason=pattern.unresolved_reason)

    advised_parts = []
    matches = []
    for block_idx, position in SPRING_BLOCK_TO_POSITION.items():
        expected = pattern.block_sequence.get(position)
        if expected is None:
            continue
        advised_parts.append(f"{BLOCK_LABELS[block_idx]}: {expected}")
        matches.append(block_registrations[block_idx] == expected)

    if matches and all(matches):
        status = "Y"
    elif matches and any(matches):
        status = "Partial"
    else:
        status = "N"

    advised = "; ".join(advised_parts) if advised_parts \
        else "(pattern has no Spring-mapped positions for this program)"
    return PatternComparison(status=status, advised=advised)
