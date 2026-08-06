"""
Stage 2 of 4 — PATTERN LOOKUP
==============================
STATUS: BLOCKED. This stage cannot run for real yet.

Goal: given (program, commencement_period), return the student's Pattern
of Study — the expected subject at each of the 8 block positions, plus
which position corresponds to "this Spring" for a student at this point
in their studies.

What's missing to build this for real:
  - Grant question #1: a maintained reference table of Program x
    Commencement Period -> subject sequence. We have partial diploma/
    subject research (see the earlier scoping artifact) but nothing
    authoritative enough to ship, and no confirmation it stays current
    when WSU changes a sequence.

Until that table exists, this module returns a placeholder pattern so
the rest of the pipeline has something to call — every field is marked
UNRESOLVED rather than guessed, so nothing downstream silently treats a
placeholder as real data.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class PatternOfStudy:
    program: int
    commencement_period: str
    # Position 1..8 -> expected subject code. Empty until the real table
    # exists (Grant question #1).
    block_sequence: dict = field(default_factory=dict)
    resolved: bool = False
    unresolved_reason: Optional[str] = None


# In-memory placeholder for what will eventually be a real reference table
# (flat file, database table, or pulled live from wherever Grant/Ashlee
# maintain it — TBD, also part of question #1).
_PATTERN_TABLE: dict[tuple[int, str], PatternOfStudy] = {}


def lookup_pattern(program: int, commencement_period: Optional[str]) -> PatternOfStudy:
    """Return the Pattern of Study for a given program + commencement period.

    Currently always returns an unresolved placeholder — see module
    docstring. Swap the body of this function out once the reference
    table exists; nothing else in the pipeline should need to change,
    since callers already handle `resolved=False`.
    """
    if commencement_period is None:
        return PatternOfStudy(
            program=program,
            commencement_period="",
            resolved=False,
            unresolved_reason="No commencement period on record for this student "
                               "(Grant question: how should these be handled?).",
        )

    key = (program, commencement_period)
    if key in _PATTERN_TABLE:
        return _PATTERN_TABLE[key]

    return PatternOfStudy(
        program=program,
        commencement_period=commencement_period,
        resolved=False,
        unresolved_reason="Pattern reference table not yet available (Grant question #1).",
    )
