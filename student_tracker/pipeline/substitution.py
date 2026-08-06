"""
Stage 4 of 4 — SUBSTITUTION (Phase 3: path back to completion)
================================================================
STATUS: BLOCKED. Cannot run for real yet.

Goal: for a student who has failed a subject in their pattern, check
whether that subject is still offered going forward, and if not, find
the nearest available replacement — the single-condition model worked
out in scoping (no independent time-based expiry; the only trigger is
"failed it, and it's not offered going forward").

What's missing to build this for real:
  - Grant question #2: subject-level pass/fail history (which specific
    subject was failed) — not present in either sample file.
  - Grant question #3: a forward-looking feed of what's offered from a
    given point onward — existence unconfirmed.
  - Grant question #4: whether a retired-subject substitution mapping
    already exists anywhere, or needs to be built from scratch.

Left as a real function signature + stub body so the pipeline's shape is
visible, rather than left out entirely.
"""

from __future__ import annotations
from dataclasses import dataclass
from typing import Optional


@dataclass
class SubstitutionResult:
    failed_subject: str
    still_offered: Optional[bool]   # None = unknown, can't check yet
    replacement_subject: Optional[str]
    resolved: bool
    unresolved_reason: Optional[str] = None


def find_substitute(failed_subject: str, from_period: str) -> SubstitutionResult:
    """Find the nearest subject a student can take in place of one they
    failed that's no longer offered.

    Always returns unresolved right now — there's no offering feed to
    query against (Grant question #3) and no confirmed source of failed-
    subject detail to call this with in the first place (Grant question #2).
    """
    return SubstitutionResult(
        failed_subject=failed_subject,
        still_offered=None,
        replacement_subject=None,
        resolved=False,
        unresolved_reason="No forward-looking subject offering feed available "
                           "yet (Grant question #3) — cannot check availability "
                           "or find a replacement.",
    )
