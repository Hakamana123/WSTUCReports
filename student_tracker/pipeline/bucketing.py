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
Every bucket below carries a confidence tier:
  - "high"   — classification confirmed by Grant, AND the resulting
               action is also confirmed. Safe to act on directly.
  - "medium" — classification confirmed by Grant, but what to actually
               do about it is still open — see grant_question.
  - "low"    — the classification itself is my inference from the
               data, not yet confirmed by Grant at all (formerly
               "candidate"). Starting point for the Grant conversation,
               not a finished rule.
See bucket.confidence and bucket.grant_question on each result.

2026-08-21: added a Conditional Enrolment (CE) credit-load check —
CE students registered over 30cp (confirmed directly by Grant in
discussion) get their own exception bucket instead of falling into
the generic AR/CE/EX "at risk" branches below. It's "medium" rather
than "high" because the cap number is confirmed but the remediation
action isn't. GS and EX needed no new code: GS's full load (55cp) is
already what blocks_filled==4 checks, and EX's "shouldn't be
registered at all" case is already the
exception_exclusion_still_registered bucket above.

2026-08-21: on_pattern_continuing now verifies subject content, not just
block count. blocks_filled==4 alone can't distinguish a student on the
correct pattern from one who filled every block with the wrong subject
(e.g. repeating an earlier-position subject instead of progressing) -
confirmed as a real, non-trivial case (21% of the old bucket in the
Ashlee sample) once subject-level comparison
(pattern_lookup.compare_registration_to_pattern) was available. GS +
blocks_filled==4 now splits three ways: subjects match -> still
on_pattern_continuing/high; subjects confirmed wrong ->
exception_registered_wrong_subjects/medium (classification is certain,
remediation isn't); pattern unresolved so match can't be checked ->
exception_full_registration_unverified/low (this is a *narrower* claim
than "on pattern" - it's "fully registered, verification blocked", not
"probably fine").

2026-08-21: extended the same subject-level check to the partial (0 <
blocks_filled < 4) and remaining AR/CE-commencing-fully-registered cases,
via the shared _pattern_verified_bucket() helper - the previous
count-only versions of both had the identical blind spot:
  - off_pattern_partial used to mean "1-3 blocks filled" full stop. 55%
    of that bucket (Ashlee sample) turned out to be on-pattern once
    checked - just not finished registering yet, not off-pattern at all.
    Now splits into on_pattern_partial_registration/medium (subjects
    match what's registered so far), off_pattern_partial/medium (subjects
    confirmed wrong - same name, bumped confidence since it's now a
    verified claim not a count-based guess), or
    exception_partial_registration_unverified/low (pattern unresolved).
  - exception_manual_review used to catch every AR/CE-under-cap student
    commencing this period with all 4 blocks filled ("doesn't match a
    confirmed rule"). 64% of that bucket turned out to be on-pattern once
    checked. Now the same three-way split as above applies, landing in
    on_pattern_at_risk_monitoring/medium instead of the catch-all.
    exception_manual_review itself is now provably unreachable given only
    GS/AR/CE/EX standings exist in the data - kept as a defensive
    fallback for an unexpected future standing value, not a real bucket
    today.
"""

from __future__ import annotations
from dataclasses import dataclass
from typing import Optional
from collections import Counter

from .ingest import StudentRecord
from .pattern_lookup import compare_registration_to_pattern

AT_RISK_LIKE = {"AR", "CE", "EX"}

# Credit points per subject type (confirmed 2026-08-21): a block/course-specific
# subject is 10cp, a semester-long prep subject is 15cp.
#
# CE_CREDIT_CAP: 30cp, confirmed directly by Grant in discussion (2026-08-21),
# not just inferred. WSU progression policy (policies.westernsydney.edu.au,
# id=27) clause 26 corroborates the mechanism — CE officially works by
# "reducing the amount of credit points" — but clause 29 defers the actual
# number to a separate "Academic Progression webpage," listed "per program
# type" (i.e. may not be one flat cap for every program) — that page hasn't
# been fetched, so whether 30 holds across all programs is still open
# (Grant question #5).
BLOCK_CREDIT_POINTS = 10
PREP_CREDIT_POINTS = 15
CE_CREDIT_CAP = 30


@dataclass
class BucketResult:
    student_id: str
    bucket: str
    confidence: str            # "high", "medium", or "low"
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


def _pattern_verified_bucket(
    sid: str, standing: str, record: StudentRecord, blocks_filled: int, *, partial: bool
) -> BucketResult:
    """Shared subject-level check for both "fully registered" (blocks_filled
    == 4) and "partially registered" (0 < blocks_filled < 4) students.
    blocks_filled alone can't tell a student on the correct pattern apart
    from one who filled the same number of blocks with the wrong subjects
    (e.g. repeating an earlier-position subject instead of progressing),
    or - for the partial case - apart from one who's on pattern but simply
    hasn't finished registering yet. compare_registration_to_pattern
    (pattern_lookup.py) is what actually resolves that.
    """
    comparison = compare_registration_to_pattern(
        record.program, record.commencement_period, record.block_registrations
    )
    reg_desc = f"{blocks_filled}/4 Spring blocks" if partial else "all 4 Spring blocks"

    if comparison.status == "Unknown":
        bucket = "exception_partial_registration_unverified" if partial \
            else "exception_full_registration_unverified"
        return BucketResult(
            sid, bucket, "low",
            f"Standing={standing}, {reg_desc} registered, but the pattern "
            f"can't be resolved to verify subject match: {comparison.reason}",
            grant_question="Same pattern-resolution gaps as Grant question #1.",
        )

    if comparison.status == "Y":
        if partial:
            return BucketResult(
                sid, "on_pattern_partial_registration", "medium",
                f"Standing={standing}, {reg_desc} registered so far, and "
                "what's registered matches the pattern — looks like a "
                "student still in the middle of registering, not one who's "
                "off pattern.",
                grant_question="Confirm no action is needed for students "
                                "who are on-pattern but haven't finished "
                                "Spring registration yet, vs. a reminder to "
                                "complete it.",
            )
        if standing == "GS":
            return BucketResult(
                sid, "on_pattern_continuing", "high",
                "Good Standing, all 4 Spring blocks registered and Spring "
                "Block 1-2 subjects match the pattern — matches 'majority "
                "advice, no action' from scoping.",
            )
        return BucketResult(
            sid, "on_pattern_at_risk_monitoring", "medium",
            f"Standing={standing}, all 4 Spring blocks registered and "
            "subjects match the pattern — but standing is below Good "
            "Standing, so whether correct registration alone is enough to "
            "stand down (vs still needing coach monitoring) isn't confirmed.",
            grant_question="Does an at-risk/conditional student who's "
                            "correctly on-pattern still need outreach, or "
                            "does correct registration mean no action?",
        )

    # comparison.status in ("N", "Partial") - subjects confirmed wrong
    bucket = "off_pattern_partial" if partial else "exception_registered_wrong_subjects"
    return BucketResult(
        sid, bucket, "medium",
        f"Standing={standing}, {reg_desc} registered, but the registered "
        f"subjects don't match the pattern (comparison: {comparison.status}) "
        f"— pattern expects {comparison.advised}.",
        grant_question="What's the fix when a student's registered "
                        "subjects don't match the pattern — swap into the "
                        "correct ones, or something else?",
    )


def bucket_student(record: StudentRecord, current_period: Optional[str]) -> BucketResult:
    sid = record.student_id

    # --- Exception: no commencement period on record ---
    if not record.has_commencement_period:
        return BucketResult(
            sid, "exception_no_commencement_period", "medium",
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
            sid, "exception_exclusion_still_registered", "medium",
            "Standing is Exclusion but student still has a Spring block "
            "registered — matches an anomaly observed directly in the "
            "sample ALA data.",
            grant_question="Is this expected lag, or should registration "
                            "be pulled once Exclusion is set?",
        )

    # --- Conditional Enrolment, registered over the 30cp cap (confirmed by
    #     Grant directly): same "registered before standing change caught
    #     up" shape as the Exclusion case above, but a cap comparison
    #     instead of a binary any-registration check. ---
    if standing == "CE":
        credit_load = BLOCK_CREDIT_POINTS * blocks_filled + (
            PREP_CREDIT_POINTS if record.prep_registration else 0
        )
        if credit_load > CE_CREDIT_CAP:
            return BucketResult(
                sid, "exception_ce_over_credit_cap", "medium",
                f"Conditional Enrolment standing but registered for {credit_load}cp "
                f"({blocks_filled} block subject(s)"
                f"{' + a prep subject' if record.prep_registration else ''}) — over "
                f"the {CE_CREDIT_CAP}cp CE load cap (confirmed by Grant).",
                grant_question="WSU policy id=27 clause 29 says CE caps are set "
                                "'per program type' on the Academic Progression "
                                "webpage — confirm 30cp holds across all programs, "
                                "not just the one(s) discussed. Also confirm the fix "
                                "is withdrawing the excess block(s) (Grant question #5).",
            )

    # --- Good Standing, fully registered: verify it's actually the right
    #     subjects, not just the right count (see _pattern_verified_bucket
    #     docstring for why - the same check now covers the partial and
    #     AR/CE-fully-registered cases below too). ---
    if standing == "GS" and blocks_filled == 4:
        return _pattern_verified_bucket(sid, standing, record, blocks_filled, partial=False)

    # --- Good Standing, zero registration: doesn't fit any confirmed rule ---
    if standing == "GS" and blocks_filled == 0:
        return BucketResult(
            sid, "zero_registration_unclear", "low",
            "Good Standing but nothing registered for Spring — spans all "
            "standings in the sample data (~14% of the file), cause unclear.",
            grant_question="Is zero registration meaningful, or a timing "
                            "artifact of when the extract is pulled?",
        )

    # --- At risk / conditional / excluded, commencing, nothing registered:
    #     candidate match for the Block 2 checkpoint 'clear-cut' case ---
    if standing in AT_RISK_LIKE and is_commencing and blocks_filled == 0:
        return BucketResult(
            sid, "reapply_next_semester", "low",
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
            sid, "success_coach_outreach", "low",
            "Below Good Standing, continuing student (not this period's "
            "intake) — approximates the 'continuing, was fine before' "
            "success-coach case from the meeting.",
            grant_question="Confirm this rule and the full named bucket "
                            "list Grant actually uses.",
        )

    # --- Partial registration (1-3 of 4 blocks): verify subject content
    #     instead of assuming "partial count" means "off pattern" - 55% of
    #     this bucket turned out to be on-pattern-but-not-finished-
    #     registering-yet once checked (2026-08-21), not actually off
    #     pattern at all. ---
    if 0 < blocks_filled < 4:
        return _pattern_verified_bucket(sid, standing, record, blocks_filled, partial=True)

    # --- Remaining case: AR/CE (under cap), commencing, fully registered -
    #     verify subject content instead of falling through to the
    #     catch-all below. This was 64% of the old exception_manual_review
    #     bucket (2026-08-21) once subject-level comparison was available;
    #     everything else that used to land in the catch-all is now
    #     provably covered by an earlier branch given only GS/AR/CE/EX
    #     appear in the standing data. ---
    if blocks_filled == 4:
        return _pattern_verified_bucket(sid, standing, record, blocks_filled, partial=False)

    # --- Fallback: anything not covered above (defensive - e.g. a
    #     standing value other than GS/AR/CE/EX showing up in future data) ---
    return BucketResult(
        sid, "exception_manual_review", "low",
        f"Standing={standing}, blocks_filled={blocks_filled}, "
        f"commencing={is_commencing} — doesn't match a confirmed rule.",
        grant_question="Full named bucket list (Q10).",
    )


def bucket_all(records: list[StudentRecord]) -> list[BucketResult]:
    current_period = _infer_current_period(records)
    return [bucket_student(r, current_period) for r in records]
