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
  - Program 9034: zero rows exist in the reference table.
  - "25 SUM" / "25 SB4": the reference table stores these as two
    dual-bracket lists joined with "+" (e.g.
    "[1, 2, -, -, -, -] + [1, -, 3, 4, -, -]") — meaning unconfirmed,
    not modeled.
  - Program 9031 (nursing/UPP): the reference pattern has 8 positions
    (vs. 6 for diplomas), but which of Prep 1/2 + Subject 1-6 map to
    which of the 8 positions, in what order, is unconfirmed.
  - Any commencement code that isn't a plain "YY - Season[ Block N]"
    string (e.g. "SC1, WSTC T1" codes) — these are legacy, out of
    scope (see project memory).
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
_UNCONFIRMED_ORDER_PROGRAMS = {9031}


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
            "entirely (Grant question #1)."
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
                               "(Grant question: how should these be handled?).",
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
                               "yet (Grant question #1).",
        )

    if program in _UNCONFIRMED_ORDER_PROGRAMS:
        return PatternOfStudy(
            program=program, commencement_period=commencement_period,
            resolved=False,
            unresolved_reason=f"Program {program} has an 8-position reference pattern "
                               "but the Prep/Subject ordering isn't confirmed "
                               "(Grant question #1).",
        )

    positions = _PATTERN_TABLE.get(session, {}).get(str(program))
    if positions is None:
        return PatternOfStudy(
            program=program, commencement_period=commencement_period,
            resolved=False,
            unresolved_reason=f"No reference table row for program {program} in "
                               f"session {session} (Grant question #1).",
        )

    return PatternOfStudy(
        program=program,
        commencement_period=commencement_period,
        block_sequence={int(k): v for k, v in positions.items()},
        resolved=True,
    )
