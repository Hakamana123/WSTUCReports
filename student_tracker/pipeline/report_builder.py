"""
ADVISORY REPORT — assembles ingest + pattern lookup + bucketing into the
final per-student row: on-pattern status, what the pattern says they
should be registered for, what they're actually registered for, and
(where confirmed) what to advise them.

STATUS: runs today for the ~65% of the roster pattern_lookup can resolve
(diploma programs, "26 - Autumn Block 1" / bare "26 - Autumn" commencers).
The other ~35% gets an explicit "Unknown" on_pattern value carrying the
specific unresolved reason from pattern_lookup, rather than a guess or a
silent blank.

The actual Spring Block 1-2-vs-pattern comparison (position 5-6 mapping,
its empirical validation, and why positions 1-4 and Spring Block 3-4
can't be checked) lives in pattern_lookup.compare_registration_to_pattern
- shared with bucketing.py so the two stay consistent. See that module's
docstring for the full reasoning.

Prep subjects (the "Spring" column, e.g. GEDU0016/GEDU0017) have no
source of truth at all: the pattern table only holds block-subject codes,
no prep entries. Always reported as registered but not evaluated -
separate gap from the position question (Grant question #6).

2026-08-21: optionally merges in past subject history from history.py
(the "AB4 for SPR" file) when a history_by_id lookup is supplied,
matched by student_id. This adds rereg_principle/template/subjects_
failed_to_date columns SITTING ALONGSIDE the existing bucket/on_pattern
output - Josiah confirmed the real "Rereg Principles" taxonomy that file
carries should not replace bucketing.py's categories. All three are None
when no history record matches (no history file uploaded, or this
student isn't in it), leaving every other column unaffected.
"""

from __future__ import annotations
from dataclasses import dataclass
from typing import Optional

from .ingest import StudentRecord
from .pattern_lookup import compare_registration_to_pattern, lookup_pattern, BLOCK_LABELS
from .bucketing import bucket_all, BucketResult
from .history import SubjectHistory


@dataclass
class AdvisoryRow:
    student_id: str
    student_name: Optional[str]
    program: int
    on_pattern: str                    # "Y", "N", "Partial", "Unknown"
    unknown_reason: Optional[str]
    subjects_advised: str
    subjects_registered: str
    bucket: str
    bucket_confidence: str
    advice: str
    rereg_principle: Optional[str] = None
    template: Optional[str] = None
    subjects_failed_to_date: Optional[str] = None


def _format_registered(record: StudentRecord) -> str:
    parts = [
        f"{label}: {subject}"
        for label, subject in zip(BLOCK_LABELS, record.block_registrations)
        if subject
    ]
    if record.prep_registration:
        parts.append(f"Prep (not evaluated): {record.prep_registration}")
    return "; ".join(parts) if parts else "(none registered)"


def _advice_for_bucket(bucket_result: BucketResult) -> str:
    """Only "high" confidence buckets have a confirmed action today - see
    bucketing.py's confidence-tier docstring. Everything else is honestly
    not actionable advice yet, not a bug in this function.
    """
    if bucket_result.confidence == "high":
        return "No action needed."
    if bucket_result.grant_question:
        return f"Flag for advisor review - not yet actionable ({bucket_result.grant_question})"
    return "Flag for advisor review - not yet actionable."


def _format_failed_subjects(history: SubjectHistory, pattern_overrides: Optional[dict]) -> str:
    """Resolve a SubjectHistory's failed pattern positions to subject codes
    where the pattern table covers them, falling back to "Position N" for
    positions it doesn't. Empty string (not "(none)") when nothing is
    outstanding, since this column is blank/None-equivalent for students
    with no history record at all - a fully-passed history should read
    the same as "no gaps", not need a different empty-state word.
    """
    if not history.failed_positions:
        return "(none)"
    pattern = lookup_pattern(history.program, history.commencement_period, pattern_overrides)
    parts = [
        pattern.block_sequence.get(pos, f"Position {pos}")
        for pos in history.failed_positions
    ]
    return "; ".join(parts)


def build_advisory_row(
    record: StudentRecord, bucket_result: BucketResult, pattern_overrides: Optional[dict] = None,
    history_by_id: Optional[dict] = None,
) -> AdvisoryRow:
    comparison = compare_registration_to_pattern(
        record.program, record.commencement_period, record.block_registrations,
        pattern_overrides,
    )
    history = history_by_id.get(record.student_id) if history_by_id else None
    return AdvisoryRow(
        student_id=record.student_id,
        student_name=record.student_name,
        program=record.program,
        on_pattern=comparison.status,
        unknown_reason=comparison.reason,
        subjects_advised=comparison.advised,
        subjects_registered=_format_registered(record),
        bucket=bucket_result.bucket,
        bucket_confidence=bucket_result.confidence,
        advice=_advice_for_bucket(bucket_result),
        rereg_principle=history.rereg_principle if history else None,
        template=history.template if history else None,
        subjects_failed_to_date=(
            _format_failed_subjects(history, pattern_overrides) if history else None
        ),
    )


def build_advisory_report(
    records: list[StudentRecord], blocks_due: Optional[int] = None,
    pattern_overrides: Optional[dict] = None, history_by_id: Optional[dict] = None,
) -> list[AdvisoryRow]:
    """blocks_due (student_tracker.pipeline.timing.blocks_due) and
    pattern_overrides (pattern_lookup.load_pattern_overrides) are both
    passed straight through to bucket_all and build_advisory_row - None
    for either leaves that piece of behavior unchanged from before these
    parameters existed.

    history_by_id (history.history_by_student_id): optional lookup of past
    subject pass/fail history, keyed by student_id. None (no history file
    uploaded) leaves rereg_principle/template/subjects_failed_to_date as
    None on every row - everything else is unaffected either way.
    """
    bucket_results = bucket_all(records, blocks_due, pattern_overrides)
    return [
        build_advisory_row(r, br, pattern_overrides, history_by_id)
        for r, br in zip(records, bucket_results)
    ]
