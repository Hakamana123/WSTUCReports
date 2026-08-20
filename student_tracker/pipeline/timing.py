"""
TIMING — turns a term start date into "how many of the 4 Spring blocks
should already be registered for by now."

Single source of truth for the day-math, shared by the Streamlit page
(display captions) and bucketing.py (actual classification), so the two
can't drift the way they would if each computed this separately.

The 4-week-block assumption (DAYS_PER_BLOCK) is not Grant-confirmed - see
bucketing.py's docstring for where it's used and what depends on it.
"""

from __future__ import annotations
from datetime import date
from typing import Optional

DAYS_PER_BLOCK = 28  # 4-week blocks, per the course-specific/block subject model


def blocks_due(term_start: Optional[date], today: Optional[date] = None) -> Optional[int]:
    """How many of the 4 Spring blocks should already be registered for,
    given Block 1's start date. None if term_start is unknown (no date
    picked) - callers should treat None as "can't tell, don't change
    behavior", not as 0.

    0   = term hasn't started yet (before Block 1).
    1-4 = that block is currently running (blocks 1..N are due).
    4   = also returned once past the whole 4-block window (16 weeks) -
          callers that need to tell "currently in Block 4" apart from
          "past the window entirely" should compare days_elapsed
          themselves; this function only answers "how many blocks
          should be registered for," and the answer to that is the same
          (all 4) in both cases.
    """
    if term_start is None:
        return None
    today = today or date.today()
    days_elapsed = (today - term_start).days
    if days_elapsed < 0:
        return 0
    return min(days_elapsed // DAYS_PER_BLOCK + 1, 4)
