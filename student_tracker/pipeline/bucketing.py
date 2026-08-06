"""
Stage 3 of 4 — BUCKETING
=========================
STATUS: RUNS TODAY, on real data. This is the one stage that doesn't
need anything from Grant to produce a first cut.

It sorts each student into a bucket using only what Ashlee's file
already gives us: Calculated Standing (already computed upstream),
registration completeness for next semester, and a heuristic
commencing/continuing split.

IMPORTANT — read before trusting the output:
Every bucket below is labelled either "grounded" (directly confirmed in
the 4 Aug meeting or follow-up scoping) or "candidate" (my inference
from the data, not yet confirmed by Grant). Candidate buckets are
starting points for the Grant conversation, not a finished rule set —
see bucket.confidence and bucket.grant_question on each result.
"""

from __future__ import annotations
from dataclasses import dataclass
from typing import Optional
from collections import Counter

from .ingest import StudentRecord

AT_RISK_LIKE = {"AR", "CE", "EX"}


@dataclass
class BucketResult:
    student_id: str
    bucket: str
    confidence: str            # "grounded" or "candidate"
    rationale: str
    grant_question: Optional[str] = None   # which open question this rule depends on


def _infer_current_period(records: list[StudentRecord]) -> Optional[str]:
    """Heuristic 'what period is this snapshot taken in' — the most common
    commencement_period value in the file, on the assumption that a large
    share of any snapshot is students who just started this period.

    This is a stand-in for an actual "current period" flag, which doesn't
    exist in the input file we've seen. Needs Grant confirmation of a
    better signal if one exists.
    """
    counts = Counter(r.commencement_period for r in records if r.has_commencement_period)
    if not counts:
        return None
    return counts.most_common(1)[0][0]


def bucket_student(record: StudentRecord, current_period: Optional[str]) -> BucketResult:
    sid = record.student_id

    # --- Exception: no commencement period on record ---
    if not record.has_commencement_period:
        return BucketResult(
            sid, "exception_no_commencement_period", "grounded",
            "No Commencement Period / Study Path Status in source data "
            "(~10% of the ALA snapshot) — cannot pattern-match at all.",
            grant_question="How are these students currently handled?",
        )

    standing = record.calculated_standing
    blocks_filled = record.blocks_filled
    is_commencing = (record.commencement_period == current_period)

    # --- Exclusion but still shows an active registration: known anomaly,
    #     seen directly in the sample data (see scoping notes). Flag rather
    #     than silently drop. ---
    if standing == "EX" and blocks_filled > 0:
        return BucketResult(
            sid, "exception_exclusion_still_registered", "grounded",
            "Standing is Exclusion but student still has a Spring block "
            "registered — matches an anomaly observed directly in the "
            "sample ALA data.",
            grant_question="Is this expected lag, or should registration "
                            "be pulled once Exclusion is set?",
        )

    # --- Good Standing, fully registered: majority-advice, no action ---
    if standing == "GS" and blocks_filled == 4:
        return BucketResult(
            sid, "on_pattern_continuing", "grounded",
            "Good Standing, all 4 Spring blocks registered — matches "
            "'majority advice, no action' from scoping.",
        )

    # --- Good Standing, zero registration: doesn't fit any confirmed rule ---
    if standing == "GS" and blocks_filled == 0:
        return BucketResult(
            sid, "zero_registration_unclear", "candidate",
            "Good Standing but nothing registered for Spring — spans all "
            "standings in the sample data (~14% of the file), cause unclear.",
            grant_question="Is zero registration meaningful, or a timing "
                            "artifact of when the extract is pulled?",
        )

    # --- At risk / conditional / excluded, commencing, nothing registered:
    #     candidate match for the Block 2 checkpoint 'clear-cut' case ---
    if standing in AT_RISK_LIKE and is_commencing and blocks_filled == 0:
        return BucketResult(
            sid, "reapply_next_semester", "candidate",
            "Below Good Standing, commenced this period, no Spring "
            "registration — approximates the 'commencing, failed blocks "
            "1+2' case, but built from the semester aggregate standing, "
            "not actual block-1/block-2 results.",
            grant_question="Does the Block 2 check need block-level pass/"
                            "fail data specifically, and where would that "
                            "come from?",
        )

    # --- At risk / conditional / excluded, continuing: candidate match for
    #     success-coach-outreach case ---
    if standing in AT_RISK_LIKE and not is_commencing:
        return BucketResult(
            sid, "success_coach_outreach", "candidate",
            "Below Good Standing, continuing student (not this period's "
            "intake) — approximates the 'continuing, was fine before' "
            "success-coach case from the meeting.",
            grant_question="Confirm this rule and the full named bucket "
                            "list Grant actually uses.",
        )

    # --- Partial registration (1-3 of 4 blocks): the case flagged as still
    #     needing design — off pattern, remediation path not computable yet ---
    if 0 < blocks_filled < 4:
        return BucketResult(
            sid, "off_pattern_partial", "candidate",
            f"{blocks_filled}/4 Spring blocks registered — partial case "
            "that needs the pattern table + subject-level fail data to "
            "resolve into an actual remediation path.",
            grant_question="Pattern reference table (Q1) + subject-level "
                            "pass/fail history (Q2).",
        )

    # --- Fallback: anything not covered above ---
    return BucketResult(
        sid, "exception_manual_review", "candidate",
        f"Standing={standing}, blocks_filled={blocks_filled}, "
        f"commencing={is_commencing} — doesn't match a confirmed rule.",
        grant_question="Full named bucket list (Q10).",
    )


def bucket_all(records: list[StudentRecord]) -> list[BucketResult]:
    current_period = _infer_current_period(records)
    return [bucket_student(r, current_period) for r in records]
