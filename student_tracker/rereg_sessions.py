"""
Academic-session arithmetic for the re-registration cohort clock.
================================================================

The teaching year is ``AUT -> SPR -> SUM``. Autumn and Spring each run four
teaching blocks and move an on-pattern student four subjects along their
pattern of study. **Summer is catch-up only** (retakes, Conditional Enrolment
students clearing a backlog) - it does not advance an on-pattern student
(confirmed by Josiah, 2026-09-02), so it contributes zero to the clock.

A student's *cohort position* for a target session is: how many pattern
subjects their cohort has worked through by the time that session starts,
plus one. Position 1 = "start of the pattern"; a position past the last
subject slot means the cohort has finished the core pattern and the student
is in elective / backlog territory.

Prep subjects are handled separately by the advice engine and are not counted
here - the clock is purely subjects-done = blocks-attended (this reproduces
the hand-checked values the v2 tool used: ``26 - Autumn Block 1`` -> 5,
``26 - Autumn Block 3`` -> 3, ``26 - Spring Block 1`` -> 1 for a Spring-26
target).
"""

from __future__ import annotations

import re

SESSION_ORDER = ["AUT", "SPR", "SUM"]        # within one year
ADVANCING = {"AUT", "SPR"}                   # sessions that move the clock; SUM does not
BLOCKS_PER_SESSION = 4

_SESSION_ALIASES = {
    "AUTUMN": "AUT", "AUT": "AUT",
    "SPRING": "SPR", "SPR": "SPR",
    "SUMMER": "SUM", "SUM": "SUM",
}

# The named targets the page offers, in cycle order. 25 SUM / 26 AUT are the
# two Grant's calculator covers; the rest run the v2 rule-tree.
NAMED_TARGETS = ["26 AUT", "26 SPR", "26 SUM", "27 AUT", "27 SPR", "27 SUM"]
DEFAULT_TARGET = "26 SPR"


def parse_period(text) -> tuple[int, str, int] | None:
    """``'26 - Autumn Block 3'`` -> ``(2026, 'AUT', 3)``.

    Block defaults to 1 when the string doesn't name one. Returns ``None`` for
    anything that doesn't start with ``<yy> - <Autumn|Spring|Summer>``.
    """
    if text is None:
        return None
    s = str(text).strip()
    m = re.match(r"\s*(\d{2})\s*-\s*(Autumn|Spring|Summer)", s, re.I)
    if not m:
        return None
    year = 2000 + int(m.group(1))
    sess = _SESSION_ALIASES[m.group(2).upper()]
    bm = re.search(r"Block\s*(\d)", s, re.I)
    block = int(bm.group(1)) if bm else 1
    return year, sess, max(1, min(BLOCKS_PER_SESSION, block))


def parse_target(text) -> tuple[int, str, int] | None:
    """``'27 AUT'`` or ``'27 AUT Block 3'`` -> ``(2027, 'AUT', 1|3)``."""
    if text is None:
        return None
    s = str(text).strip()
    m = re.match(r"(\d{2})\s*[-\s]?\s*(AUT|SPR|SUM|Autumn|Spring|Summer)", s, re.I)
    if not m:
        return None
    year = 2000 + int(m.group(1))
    sess = _SESSION_ALIASES[m.group(2).upper()]
    bm = re.search(r"Block\s*(\d)", s, re.I)
    block = int(bm.group(1)) if bm else 1
    return year, sess, max(1, min(BLOCKS_PER_SESSION, block))


def _slot_index(year: int, sess: str) -> int:
    """Monotonic integer for a (year, session) so they can be ordered."""
    return year * 3 + SESSION_ORDER.index(sess)


def _blocks_completed(commencement: tuple[int, str, int], target: tuple[int, str, int]) -> int:
    """Teaching blocks the cohort works through from ``commencement`` up to the
    point ``target`` begins (Autumn/Spring only)."""
    cy, cs, cb = commencement
    ty, ts, tb = target
    ci, ti = _slot_index(cy, cs), _slot_index(ty, ts)

    if ti < ci:
        return 0  # target is before the student started - treat as pattern start

    blocks = 0
    for idx in range(ci, ti + 1):
        y, s = idx // 3, SESSION_ORDER[idx % 3]
        if s not in ADVANCING:
            continue
        is_start = idx == ci
        is_target = idx == ti
        if is_start and is_target:
            step = tb - cb
        elif is_start:
            step = BLOCKS_PER_SESSION - cb + 1
        elif is_target:
            step = tb - 1
        else:
            step = BLOCKS_PER_SESSION
        blocks += max(0, step)
    return blocks


def cohort_position(
    commencement_text,
    target_text,
    subject_slots: int = 6,
) -> tuple[int, str]:
    """``(position, note)`` - the pattern position the student's cohort reaches
    at the start of the target session/block.

    ``position`` is 1-based and clamped to ``[1, subject_slots + 1]``; a value
    of ``subject_slots + 1`` means "cohort has finished the core pattern".
    ``note`` is a short human sentence, or ``""`` when the placement is clean.
    """
    comm = parse_period(commencement_text)
    tgt = parse_target(target_text)

    if tgt is None:
        return 1, f"target '{target_text}' not understood - placed at pattern start"
    if comm is None:
        return 1, (
            f"commencement '{commencement_text}' not recognised - planned "
            "earliest-outstanding-first, please sanity-check"
        )

    blocks = _blocks_completed(comm, tgt)
    position = blocks + 1
    note = ""
    if position > subject_slots:
        position = subject_slots + 1
        note = "past the core pattern - clearing outstanding subjects earliest-first"
    return position, note


def advance(target_text, steps: int) -> str:
    """The teaching session ``steps`` sessions after ``target_text``.
    ``AUT -> SPR -> AUT -> …``; from Summer the next session is the following
    Autumn. ``advance("26 SPR", 0)`` -> ``"26 SPR"``; ``advance("26 SPR", 1)``
    -> ``"27 AUT"``; ``advance("26 SUM", 1)`` -> ``"27 AUT"``.
    """
    tgt = parse_target(target_text)
    if tgt is None:
        return ""
    y, s, _ = tgt
    for _ in range(max(0, steps)):
        if s == "AUT":
            s = "SPR"
        else:  # SPR or SUM
            y, s = y + 1, "AUT"
    return f"{y % 100:02d} {s}"


def available_blocks(target_text, cap: int = BLOCKS_PER_SESSION) -> int:
    """How many teaching blocks are left in the target session.

    A whole-session target (block 1) leaves 4; a part-way target ("advise as of
    Block 3") leaves ``4 - 3 + 1 = 2``.
    """
    tgt = parse_target(target_text)
    if tgt is None:
        return cap
    _, _, tb = tgt
    return max(1, min(cap, BLOCKS_PER_SESSION - tb + 1))


def uses_calculator(target_text) -> bool:
    """True when Grant's calculator should be tried for this target.

    26 AUT / 25 SUM have exact offering patterns; any other whole-session
    Autumn / Spring / Summer target carries the pattern forward (Autumn &
    Spring from 26 AUT, Summer from 25 SUM) on the assumption that the
    offering is unchanged unless we're told otherwise. The calculator has no
    part-way-through logic, so a ``Block 3`` target still runs the rule-tree.
    """
    tgt = parse_target(target_text)
    if tgt is None:
        return False
    _, s, b = tgt
    return b == 1 and s in {"AUT", "SPR", "SUM"}


def carry_base(target_text) -> str | None:
    """The session whose offering pattern is assumed to still hold for a
    session the calculator doesn't list exactly. ``None`` if not carryable."""
    tgt = parse_target(target_text)
    if tgt is None:
        return None
    _, s, b = tgt
    if b != 1:
        return None
    return {"AUT": "26 AUT", "SPR": "26 AUT", "SUM": "25 SUM"}.get(s)
