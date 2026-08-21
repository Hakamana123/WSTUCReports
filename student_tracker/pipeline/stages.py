"""
STAGE REGISTRY — the WSU Modular Re-registration Timeline, encoded.

Source: "Modular Re-registration Timeline.xlsx" (supplied by Josiah,
2026-08-21). Each semester (AUT/SPR/SUM) has 5 stages in that source
sheet - this module is the single place that tracks what each stage
needs and whether we can run it today. Nothing here invents logic for a
blocked/undefined stage; see STATUS below.

STATUS legend:
  - "built"     — real logic exists and runs against real sample data.
  - "blocked"   — the stage needs a source file/data granularity we've
                   never seen a sample of. No guessed parsing here -
                   same discipline as substitution.py's stub.
  - "undefined" — the source Timeline sheet itself never specified this
                   stage (blank row, or literally "etc" / "etc").

Everything the rest of the pipeline (ingest/pattern_lookup/bucketing/
report_builder) already does maps to AUT/SPR Stage 4 (the roster itself
- what ingest.load_roster consumes) and Stage 5 (the advisory report -
what bucket_all/build_advisory_report produce). Stage 5's population per
the Timeline is Conditional Enrolment only; stage5_ce_only() below
*wraps* the general-purpose bucketing/report output to scope it to that
audience, rather than mutating bucketing.py itself - the general version
stays intact for whatever consumes it once Stage 1 (or others) becomes
buildable.
"""

from __future__ import annotations
from dataclasses import dataclass, field


@dataclass
class Stage:
    semester: str          # "AUT", "SPR", "SUM"
    number: int             # 1-5
    name: str
    status: str              # "built", "blocked", "undefined"
    when: str = ""
    who: str = ""
    population: str = ""
    what: str = ""
    data_needs: str = ""
    note: str = ""


STAGES: list[Stage] = [
    Stage(
        semester="AUT", number=1,
        name="Recommendations for students who fail AB1 and AB2",
        status="blocked",
        when="In break week after AB2, before AB3.",
        who="Grant",
        population="Any commencing students who failed to complete AB1 AND "
                    "AB2 (dropped, TFN/SSAF issues, off pattern, etc).",
        what="Advise dropping AB3/AB4 registrations; start over in SB1 as a "
             "commencing student.",
        data_needs="Block-level results for AB1 AND AB2 specifically (not "
                    "aggregate Calculated Standing), future registrations "
                    "for AB3/AB4, AUT prep registration.",
        note="Blocked on the same missing block-level pass/fail data as "
             "bucketing.py's reapply_next_semester grant_question - this "
             "Timeline confirms the need is real, it doesn't resolve it.",
    ),
    Stage(
        semester="AUT", number=2,
        name="Recommendations for Spring Block subjects (Post AB4)",
        status="built",
        when="After AB4 results entered, before AUT results released.",
        who="Grant",
        population="All active students registered in any subject over "
                    "AB1-4 post census. Excludes AUT prep results.",
        what="Results-based recommendation - not detailed further in the "
             "source sheet.",
        data_needs="Results from each Autumn Block + AUT prep registration. "
                    "No future registrations. Needs a results-only file "
                    "(example name: '2026 AB4 for 2026 SPR').",
        note="2026-08-22: a real file matching this exact example filename "
             "was found and confirmed on every checkable point - population "
             "spans all standings (not just CE), and every row's 'Advice "
             "for when' reads '26 SPR' (generated after Autumn's results, "
             "advising for the coming Spring). See "
             "pages/10_Stage2_Recommendations.py and history.py's module "
             "docstring. 'Built' here means the same thing it means for "
             "Stage 4 - we read/display/filter/export Grant's own already-"
             "computed output, not independently reproduce his Rereg "
             "Principles/Template classification rules, which we don't "
             "know. SPR Stage 2 stays blocked - it needs a different real "
             "file ('2025 SB4 for 2025 SU1') we haven't seen a sample of.",
    ),
    Stage(
        semester="AUT", number=3,
        name="Recommendations for Prep subjects (Post AUT Results)",
        status="built",
        when="SB1 Week 0 (break between AUT and SPR).",
        who="Grant",
        population="All active students registered in any subject over "
                    "AB1-4 and Autumn (prep) post census.",
        what="Results-based recommendation - not detailed further in the "
             "source sheet.",
        data_needs="Results from each Autumn Block AND the AUTUMN prep "
                    "subject. No future registrations.",
        note="2026-08-22: confirmed the same AB4-for-SPR file used for "
             "Stage 2 also satisfies this - its 'Prep Name' column follows "
             "the identical finished-recommendation shape as 'B1-B4 Name' "
             "('you don't need to register' vs. a specific GEDU0016/0017 "
             "code), just for the two prep subjects instead of the block "
             "subjects. See pages/11_Stage3_Prep_Recommendations.py - kept "
             "as its own page rather than merged into Stage 2's, per "
             "Josiah's 'split by stage, not by type' preference. Same "
             "'built means displays Grant's output, not reproduces his "
             "rules' caveat as Stage 2. SPR Stage 3 stays blocked - needs "
             "a different file ('2025 SPR for 2025 SUM') not yet seen.",
    ),
    Stage(
        semester="AUT", number=4,
        name="Provision of progression outcomes",
        status="built",
        when="SB1 Week 1.",
        who="UNCLEAR (per source sheet)",
        population="All active students.",
        what="Information on their progression outcome.",
        data_needs="This is the Progression Outcomes roster itself.",
        note="This is exactly what student_tracker.pipeline.ingest."
             "load_roster already consumes - the roster used throughout "
             "this project IS this stage's output.",
    ),
    Stage(
        semester="AUT", number=5,
        name="Applied progression outcomes recommendations",
        status="built",
        when="SB2 Week 1.",
        who="Grant",
        population="All students on Conditional Enrolment.",
        what="Provide manual academic advice on subject registration.",
        data_needs="Results from SB1, future registrations for SB2/SB3/SB4, "
                    "SPR prep registration.",
        note="This is what bucketing.py + report_builder.py already do, "
             "scoped to CE via stage5_ce_only() below rather than mutating "
             "bucketing.py itself - keeps the general-purpose, already-"
             "validated logic intact for other consumers.\n"
             "2026-08-22: confirmed directly against a real file for this "
             "exact stage - '2026 AUT Stage 5 Reregistration Advice List "
             "v1.1.xlsx', whose Stage column literally reads '26 AUT Stage "
             "5' and whose every row is Progression Outcome = 'Conditional "
             "Enrolment', matching population above exactly. It also has no "
             "B1 Name column (only Prep/B2/B3/B4) - consistent with 'SB2 "
             "Week 1' timing, since SB1 is already a locked-in result by "
             "then, not something left to advise on. See history.py's "
             "module docstring for the parser validation this file enabled.",
    ),

    # SPR mirrors AUT exactly, block roles swapped (SB<->AB).
    Stage(
        semester="SPR", number=1,
        name="Recommendations for students who fail SB1 and SB2",
        status="blocked",
        when="In break week after SB2, before SB3.",
        who="Grant",
        population="Any commencing students who failed to complete SB1 AND "
                    "SB2.",
        what="Advise dropping SB3/SB4 registrations; start over in AB1 as a "
             "commencing student.",
        data_needs="Block-level results for SB1 AND SB2, future "
                    "registrations for SB3/SB4, SPR prep registration.",
        note="Same blocker as AUT Stage 1.",
    ),
    Stage(
        semester="SPR", number=2,
        name="Recommendations for Summer subjects (Post SB4)",
        status="blocked",
        when="After SB4 results entered, before SPR results released.",
        who="Grant",
        population="All active students registered in any subject over "
                    "SB1-4 post census. Excludes SPR prep results.",
        what="Results-based recommendation - not detailed further in the "
             "source sheet.",
        data_needs="Results from each Spring Block + SPR prep registration. "
                    "No future registrations. Example file name: '2025 SB4 "
                    "for 2025 SU1'.",
    ),
    Stage(
        semester="SPR", number=3,
        name="Recommendations for Prep subjects (Post SPR Results)",
        status="blocked",
        when="Break week after SPR session, before SU1.",
        who="Grant",
        population="All active students registered in any subject over "
                    "SB1-4 and Spring (prep) post census.",
        what="Results-based recommendation - not detailed further in the "
             "source sheet.",
        data_needs="Results from each Spring Block AND the SPRING prep "
                    "subject. No future registrations. Example file name: "
                    "'2025 SPR for 2025 SUM'.",
    ),
    Stage(
        semester="SPR", number=4,
        name="Provision of progression outcomes",
        status="built",
        when="SUM Week 1.",
        who="UNCLEAR (per source sheet)",
        population="All active students.",
        what="Information on their progression outcome.",
        data_needs="Same shape as AUT Stage 4 - the ALA-style Progression "
                    "Outcomes export.",
    ),
    Stage(
        semester="SPR", number=5,
        name="Applied progression outcomes recommendations",
        status="built",
        when="SU2 Week 2.",
        who="Grant",
        population="All students on Conditional Enrolment.",
        what="Provide manual academic advice on subject registration.",
        data_needs="Results from SB1, future registrations for SB2/SB3/SB4, "
                    "SPR prep registration.",
        note="KNOWN TIMING CONFLICT, flagged directly in the source sheet "
             "and not resolved here: 'Show cause period ends 14th January. "
             "This is obviously after SU1 completion, and after Census for "
             "SUMMER (18th December). Students would already have been "
             "registered in possibly 25 CP.' Preserved verbatim - a real "
             "process problem, not ours to fix.",
    ),

    # SUM only has 3 of 5 rows meaningfully filled in the source sheet.
    Stage(
        semester="SUM", number=1,
        name="SUM Stage 1",
        status="blocked",
        when="After SU1 results released, before SU2 begins.",
        who="Grant",
        population="Students registered in SU1 post census, but NOT "
                    "registered in SU2 or SUM.",
        what="Dropout-detection shape, not fail-both-blocks - genuinely "
             "different logic from AUT/SPR Stage 1.",
        data_needs="Results of SU1, registrations of SU2 and SUM. Example "
                    "file name: '2025 SU1 for 2026 AUT'.",
    ),
    Stage(
        semester="SUM", number=2,
        name="SUM Stage 2",
        status="blocked",
        when="After SU2 results.",
        who="Grant",
        population="Students registered in SU2 and NOT in SUM.",
        what="Same dropout-detection shape as SUM Stage 1.",
        data_needs="Results of SU1 and SU2, registrations of SUM. Example "
                    "file name: '2025 SU2 for 2026 AUT'.",
    ),
    Stage(
        semester="SUM", number=3,
        name="SUM Stage 3",
        status="undefined",
        when="After SUM results, AB1 Week 1.",
        who="Grant",
        note="Not a gap in our reading - the source Timeline sheet "
             "literally reads 'etc' / 'etc' for population and action on "
             "this row.",
    ),
    Stage(
        semester="SUM", number=4,
        name="Provision of progression outcomes",
        status="undefined",
        who="UNCLEAR (per source sheet)",
        note="Blank row in the source sheet.",
    ),
    Stage(
        semester="SUM", number=5,
        name="Applied progression outcomes recommendations",
        status="undefined",
        note="Blank row in the source sheet.",
    ),
]


def stages_by_status(status: str) -> list[Stage]:
    return [s for s in STAGES if s.status == status]


def get_stage(semester: str, number: int) -> Stage:
    for s in STAGES:
        if s.semester == semester and s.number == number:
            return s
    raise KeyError(f"No stage {semester} {number}")


def stage5_ce_only(records, advisory_rows):
    """AUT/SPR Stage 5 ('applied progression outcomes recommendations'):
    scopes the general-purpose advisory report down to Conditional
    Enrolment students only, per the Timeline ('All students on
    Conditional Enrolment... SB2 Week 1'). Wraps rather than mutates
    bucketing.py/report_builder.py - those stay general-purpose for
    other consumers (e.g. a future Stage 1 built once block-level pass/
    fail data exists).

    records: list[StudentRecord] (student_tracker.pipeline.ingest)
    advisory_rows: list[AdvisoryRow] (student_tracker.pipeline.
        report_builder), same order/source as records.
    """
    ce_ids = {r.student_id for r in records if r.calculated_standing == "CE"}
    return [row for row in advisory_rows if row.student_id in ce_ids]
