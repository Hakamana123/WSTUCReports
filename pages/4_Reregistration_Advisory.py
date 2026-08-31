"""
Reregistration Advisory — main hub for the Modular Re-registration Timeline
=============================================================================

One page, one entry point, ONE upload feeding every stage. Restructured
2026-08-25 at Josiah's request ("I want to use the updated (or completed)
Ashlee file as the input file, regardless which stage") - a single real
file from Ashlee's team (student_tracker/pipeline/live_roster.py) now
drives both the Stage 2/3 section and the Stage 5 section, replacing the
two separate uploads (Grant's AB4-for-SPR file, the old Progression
Outcomes roster) that used to be needed. See live_roster.py's module
docstring for the full story of how that file was confirmed (a recorded
walkthrough meeting with Ashlee) and exactly what it contains.

Earlier history: this page used to be a single Stage 5 report; Stage 2/3
was folded in as a second section on 2026-08-22 (previously its own page,
pages/10_Stage2_3_Recommendations.py, now gone); each section had its own
file upload until this 2026-08-25 restructure unified them.

STATUS: skeleton, not a finished tool. See the "Stage coverage" expander
below and student_tracker/pipeline/stages.py for what's built vs blocked
vs undefined, and why.
"""

from __future__ import annotations

import io
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

from student_tracker.pipeline import live_roster
from student_tracker.pipeline.bucketing import bucket_all
from student_tracker.pipeline.report_builder import build_advisory_report
from student_tracker.pipeline.stages import STAGES, stage5_ce_only, get_stage
from student_tracker.pipeline.glossary import glossary_dataframe
from student_tracker.pipeline.timing import blocks_due as compute_blocks_due, DAYS_PER_BLOCK
from student_tracker.pipeline.pattern_lookup import load_pattern_overrides
from student_tracker.pipeline.history import (
    failed_subject_codes_sem1_only, infer_current_commencement_year,
    suggested_rereg_principle, suggested_block_advice, suggested_prep_advice,
)

st.set_page_config(page_title="Reregistration Advisory", layout="wide")


def _style_status(val):
    if val == "built":
        return "background-color: #E7F0EA; color: #2F6B4F"
    if val == "blocked":
        return "background-color: #FAF0DC; color: #A8720F"
    if val == "undefined":
        return "background-color: #F2E8E8; color: #8A4B4B"
    return ""


def _style_confidence(val):
    if val == "high":
        return "background-color: #E7F0EA; color: #2F6B4F"
    if val == "medium":
        return "background-color: #FDEEDC; color: #B5590F"
    if val == "low":
        return "background-color: #FAF0DC; color: #A8720F"
    return ""


SHORT_LABELS = {
    "on_pattern_continuing": "On pattern",
    "off_pattern_partial": "Partial",
    "success_coach_outreach": "Coach outreach",
    "zero_registration_unclear": "Zero reg. (unclear)",
    "zero_registration_too_early": "Zero reg. (too early)",
    "zero_registration_overdue": "Zero reg. (overdue)",
    "exception_manual_review": "Manual review",
    "reapply_next_semester": "Reapply",
    "exception_exclusion_still_registered": "Excl. anomaly",
    "exception_no_commencement_period": "No commencement",
    "exception_ce_over_credit_cap": "CE over cap",
    "exception_registered_wrong_subjects": "Wrong subjects",
    "exception_full_registration_unverified": "Full reg. unverified",
    "on_pattern_partial_registration": "On pattern (partial)",
    "on_pattern_at_risk_monitoring": "On pattern (at-risk)",
    "exception_partial_registration_unverified": "Partial reg. unverified",
}

# Same term calendar as 1_Engagement_Report.py, for consistency across pages.
TERM_OPTIONS = {
    "Autumn 2026": date(2026, 3, 2),
    "Spring 2026": date(2026, 7, 20),
    "Summer 2026": date(2026, 11, 23),
    "Autumn 2027": date(2027, 3, 1),
    "Spring 2027": date(2027, 7, 19),
}


def _render_stage_2_3(records, skipped) -> None:
    stage2 = get_stage("AUT", 2)
    stage3 = get_stage("AUT", 3)
    st.markdown(
        f"**Stage 2 - {stage2.name}**  \n"
        f"When: {stage2.when}  \n"
        f"Population: {stage2.population}  \n\n"
        f"**Stage 3 - {stage3.name}**  \n"
        f"When: {stage3.when}  \n"
        f"Population: {stage3.population}"
    )
    st.caption(
        "Stage 2 is the block-subject recommendation, Stage 3 is the "
        "prep-subject recommendation - both come from the same uploaded "
        "roster above, so this section covers both together. Everything "
        "below is OUR OWN best-effort suggestion (see history.py's "
        "REGISTRATION ADVICE docstring section for the confirmed rules "
        "behind it, and suggested_rereg_principle for the severity label) "
        "- this file is upstream of Grant's own judgment layer, so there "
        "is no 'Grant's own advice' to fall back to or compare against "
        "here anymore, unlike before this page moved to the live roster."
    )
    st.caption(
        "The Suggested B1-4/Prep advice columns deliberately punt to "
        "'speak with your academic learning advisor' for anything more "
        "complex than a single Block 1/2 failure or a single Block 3/4 "
        "failure - a student with 2+ subjects outstanding, or a Block "
        "1/2 failure, gets a generic ALA referral rather than a guessed "
        "schedule, because that's the honest answer for those cases, not "
        "a gap in the logic."
    )

    if records is None:
        st.info("Upload the roster above and click Generate.")
        return

    current_autumn_year = infer_current_commencement_year(records)

    computed_rows = [
        {
            "Student ID": r.student_id,
            "Student Name": r.student_name,
            "Program": r.program,
            "Commencement Period": r.commencement_period,
            "Status": "",
            "Suggested Rereg Principle": suggested_rereg_principle(r, current_autumn_year) or "",
            "Subjects failed in Semester 1": failed_subject_codes_sem1_only(r),
            "Electives outstanding": r.electives_outstanding,
            "Suggested B1 advice": suggested_block_advice(r)[0],
            "Suggested B2 advice": suggested_block_advice(r)[1],
            "Suggested B3 advice": suggested_block_advice(r)[2],
            "Suggested B4 advice": suggested_block_advice(r)[3],
            "Suggested Prep advice": suggested_prep_advice(r) or "",
        }
        for r in records
    ]
    no_advice_rows = [
        {
            "Student ID": s.student_id,
            "Student Name": s.student_name,
            "Program": s.program,
            "Commencement Period": s.commencement_period,
            "Status": s.note,
            "Suggested Rereg Principle": "",
            "Subjects failed in Semester 1": "",
            "Electives outstanding": None,
            "Suggested B1 advice": "",
            "Suggested B2 advice": "",
            "Suggested B3 advice": "",
            "Suggested B4 advice": "",
            "Suggested Prep advice": "",
        }
        for s in (skipped or [])
    ]
    report_df = pd.DataFrame(computed_rows + no_advice_rows)

    total = len(report_df)
    no_advice_count = len(no_advice_rows)
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total students", f"{total:,}")
    m2.metric("Programs represented", report_df["Program"].nunique())
    m3.metric(
        "With a suggested principle",
        f"{int((report_df['Suggested Rereg Principle'] != '').sum()):,}",
    )
    m4.metric("No advice this session", f"{no_advice_count:,}")

    if no_advice_count:
        st.caption(
            f"{no_advice_count:,} student(s) above have no computed advice - "
            "see their Status for why (most commonly: already registered "
            "for a later session and haven't started yet, or on Leave of "
            "Absence/Deferred). They're still included in the table and "
            "export below rather than dropped, since knowing who they are "
            "and why still has value."
        )

    st.info(
        "'Suggested Rereg Principle' is blank when we don't have a "
        "confident basis for this student (Nursing/UPP, a different "
        "commencement cohort, or a genuinely ambiguous fail count) - "
        "it's not a claim that nothing's wrong, just that we can't say."
    )

    st.subheader("Per-student recommendations")
    f1, f2, f3 = st.columns(3)
    program_filter = f1.multiselect(
        "Filter: Program", options=sorted(report_df["Program"].unique()), default=[],
        key="stage23_program_filter",
    )
    principle_filter = f2.multiselect(
        "Filter: Suggested Rereg Principle",
        options=sorted(p for p in report_df["Suggested Rereg Principle"].unique() if p),
        default=[],
        key="stage23_principle_filter",
    )
    status_filter = f3.multiselect(
        "Filter: Status",
        options=sorted(s for s in report_df["Status"].unique() if s),
        default=[],
        key="stage23_status_filter",
    )

    filtered_df = report_df
    if program_filter:
        filtered_df = filtered_df[filtered_df["Program"].isin(program_filter)]
    if principle_filter:
        filtered_df = filtered_df[filtered_df["Suggested Rereg Principle"].isin(principle_filter)]
    if status_filter:
        filtered_df = filtered_df[filtered_df["Status"].isin(status_filter)]

    st.dataframe(filtered_df, use_container_width=True, hide_index=True)
    st.caption(f"Showing {len(filtered_df):,} of {len(report_df):,} students.")

    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        report_df.to_excel(writer, sheet_name="Stage 2-3 Recommendations", index=False)

    st.download_button(
        "Download Stage 2/3 recommendations (Excel)",
        excel_buffer.getvalue(),
        "stage2_3_recommendations.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="stage23_download",
    )


def _render_stage_5(records, pattern_overrides) -> None:
    st.caption(
        "Sorts students into advisory buckets from progression standing "
        "and Spring registration completeness. Medium/low confidence "
        "buckets below are starting points for discussion, not confirmed "
        "rules."
    )

    term_choice = st.selectbox(
        "Which term's Block 1 already started? (used as the Week 1 "
        "reference point - the block you're actually in today is "
        "calculated from this, not assumed to be Block 1)",
        options=list(TERM_OPTIONS.keys()) + ["Custom date"],
        index=None,
        placeholder="Select a term...",
        key="stage5_term_choice",
    )

    term_start = None
    if term_choice == "Custom date":
        term_start = st.date_input("Custom term start date", value=None, key="stage5_custom_date")
    elif term_choice is not None:
        term_start = TERM_OPTIONS[term_choice]

    blocks_due_value = compute_blocks_due(term_start)

    if term_start:
        days_elapsed = (date.today() - term_start).days
        if days_elapsed < 0:
            # %-d is Unix-only (breaks on Windows) - build the date string without it.
            formatted_date = f"{term_start:%A}, {term_start.day} {term_start:%B %Y}"
            st.caption(
                f"Term hasn't started yet — Block 1 begins {formatted_date}. Zero "
                "registration is treated as fully expected until then, not overdue."
            )
        elif days_elapsed >= 4 * DAYS_PER_BLOCK:
            st.caption(
                f"{days_elapsed} days into term — past the assumed 4-block window "
                "(16 weeks). Likely exam period or between semesters. All 4 blocks "
                "are treated as due."
            )
        else:
            current_block = days_elapsed // DAYS_PER_BLOCK + 1
            st.caption(
                f"{days_elapsed} days into term → assuming back-to-back 4-week "
                f"blocks, that's roughly **Spring Block {current_block}** today. "
                f"Blocks 1-{current_block} are now treated as due below - a Good "
                "Standing student with zero registration is flagged as overdue "
                "rather than an unclear timing artifact. (The 4-week-block "
                "assumption itself is still not formally confirmed.)"
            )
    else:
        st.caption(
            "Pick a term (or a custom date) to distinguish overdue zero-"
            "registration from students who simply haven't reached that point "
            "yet - without it, zero registration stays classified as unclear "
            "either way."
        )

    if records is None:
        st.info("Upload the roster above to run the bucketing pass.")
        return

    results = bucket_all(records, blocks_due_value, pattern_overrides)

    result_by_id = {r.student_id: r for r in results}
    rows = [
        {
            "student_id": record.student_id,
            "program": record.program,
            "bucket": result_by_id[record.student_id].bucket,
            "confidence": result_by_id[record.student_id].confidence,
            "rationale": result_by_id[record.student_id].rationale,
            "grant_question": result_by_id[record.student_id].grant_question or "",
        }
        for record in records
    ]
    result_df = pd.DataFrame(rows)

    total_students = len(records)
    high_count = int((result_df["confidence"] == "high").sum())
    medium_count = int((result_df["confidence"] == "medium").sum())
    low_count = int((result_df["confidence"] == "low").sum())
    programs_represented = result_df["program"].nunique()

    # --- Summary metrics ---
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Total students", f"{total_students:,}")
    m2.metric(
        "High confidence",
        f"{high_count:,}",
        f"{high_count / total_students:.0%} of total",
    )
    m3.metric(
        "Medium confidence",
        f"{medium_count:,}",
        f"{medium_count / total_students:.0%} of total",
    )
    m4.metric(
        "Low confidence",
        f"{low_count:,}",
        f"{low_count / total_students:.0%} of total",
    )
    m5.metric("Programs represented", programs_represented)

    # --- Bucket breakdown ---
    st.subheader("Buckets")
    st.caption(
        "High = classification and the resulting action are both confirmed — "
        "safe to act on. Medium = classification confirmed, but what to do about it is "
        "still open — see the open question. Low = the classification itself is an "
        "inference from the data, not yet confirmed at all."
    )

    bucket_summary = (
        result_df.groupby("bucket")
        .agg(
            confidence=("confidence", "first"),
            count=("student_id", "count"),
            rationale=("rationale", "first"),
            grant_question=("grant_question", "first"),
        )
        .reset_index()
        .sort_values("count", ascending=False)
    )
    bucket_summary["% of total"] = (bucket_summary["count"] / total_students * 100).round(1)
    bucket_summary["bucket"] = bucket_summary["bucket"].str.replace("_", " ", regex=False)
    bucket_summary = bucket_summary[
        ["bucket", "confidence", "count", "% of total", "rationale", "grant_question"]
    ]

    st.dataframe(
        bucket_summary.style.map(_style_confidence, subset=["confidence"]),
        use_container_width=True,
        hide_index=True,
        column_config={
            "bucket": st.column_config.TextColumn("Bucket", width="medium"),
            "confidence": st.column_config.TextColumn("Confidence"),
            "count": st.column_config.NumberColumn("Count"),
            "% of total": st.column_config.NumberColumn("% of total", format="%.1f%%"),
            "rationale": st.column_config.TextColumn("Rationale (representative)", width="large"),
            "grant_question": st.column_config.TextColumn("Open question / to be confirmed", width="large"),
        },
    )

    st.bar_chart(bucket_summary.set_index("bucket")["count"])

    # --- Per-program breakdown ---
    st.subheader("By program")
    st.caption("Bucket counts per program code. 9031 / 9034 are UPP; the rest are diplomas.")

    program_pivot = pd.pivot_table(
        result_df,
        index="program",
        columns="bucket",
        values="student_id",
        aggfunc="count",
        fill_value=0,
    )
    program_pivot["Total"] = program_pivot.sum(axis=1)
    program_pivot = program_pivot.sort_values("Total", ascending=False)
    program_pivot = program_pivot.rename(columns=SHORT_LABELS)

    st.dataframe(program_pivot, use_container_width=True)

    # --- Per-student advisory report ---
    st.subheader("Per-student advisory report")
    st.caption(
        "On pattern: Y/Partial/N compares actual Spring Block 1-2 registration "
        "against the pattern's expected subjects (only checkable for diploma "
        "programs commencing '26 - Autumn Block 1' — the position-5/6-to-Spring "
        "mapping is empirically validated, not yet formally confirmed; see "
        "report_builder.py). Unknown = pattern can't be resolved yet for this "
        "student — see Reason. Advice is only populated for high-confidence "
        "buckets; everything else is flagged for advisor review pending confirmation."
    )

    advisory_rows = build_advisory_report(records, blocks_due_value, pattern_overrides)

    stage5_only = st.checkbox(
        "Scope to AUT/SPR Stage 5 (Conditional Enrolment only, per the "
        "Timeline - 'All students on Conditional Enrolment', SB2/SU2 Week "
        "1-2)",
        value=False,
        key="stage5_ce_only_checkbox",
    )
    if stage5_only:
        advisory_rows = stage5_ce_only(records, advisory_rows)

    advisory_df = pd.DataFrame(
        {
            "Student ID": r.student_id,
            "Student Name": r.student_name,
            "Program": r.program,
            "On pattern": r.on_pattern,
            "Reason (if Unknown)": r.unknown_reason or "",
            "Block 1 advised": r.advised_blocks[0],
            "Block 2 advised": r.advised_blocks[1],
            "Block 3 advised": r.advised_blocks[2],
            "Block 4 advised": r.advised_blocks[3],
            "Block 1 registered": r.registered_blocks[0],
            "Block 2 registered": r.registered_blocks[1],
            "Block 3 registered": r.registered_blocks[2],
            "Block 4 registered": r.registered_blocks[3],
            "Prep registered": r.prep_registered or "",
            "Bucket": r.bucket.replace("_", " "),
            "Bucket confidence": r.bucket_confidence,
            "Advice": r.advice,
        }
        for r in advisory_rows
    )

    f1, f2 = st.columns(2)
    pattern_filter = f1.multiselect(
        "Filter: On pattern",
        options=["Y", "Partial", "N", "Unknown"],
        default=[],
        key="stage5_pattern_filter",
    )
    confidence_filter = f2.multiselect(
        "Filter: Bucket confidence",
        options=["high", "medium", "low"],
        default=[],
        key="stage5_confidence_filter",
    )

    filtered_df = advisory_df
    if pattern_filter:
        filtered_df = filtered_df[filtered_df["On pattern"].isin(pattern_filter)]
    if confidence_filter:
        filtered_df = filtered_df[filtered_df["Bucket confidence"].isin(confidence_filter)]

    st.dataframe(filtered_df, use_container_width=True, hide_index=True)
    st.caption(f"Showing {len(filtered_df):,} of {len(advisory_df):,} students.")

    show_glossary = st.checkbox("Show glossary", value=False, key="stage5_show_glossary")
    if show_glossary:
        st.caption(
            "What every column, confidence tier, on-pattern value, and bucket "
            "name in this report actually means. Also included as its own tab "
            "in the downloaded Excel file below."
        )
        st.dataframe(glossary_dataframe(), use_container_width=True, hide_index=True)

    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        advisory_df.to_excel(writer, sheet_name="Advisory Report", index=False)
        glossary_dataframe().to_excel(writer, sheet_name="Glossary", index=False)

    st.download_button(
        "Download advisory report (Excel, with Glossary tab)",
        excel_buffer.getvalue(),
        "reregistration_advisory_report.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="stage5_download",
    )


# ============================= Page body =============================

_TEMPLATE_PATH = Path(__file__).resolve().parent.parent / "templates" / "pattern_of_study_template.xlsx"
if _TEMPLATE_PATH.exists():
    st.sidebar.download_button(
        "Download Pattern of Study template",
        _TEMPLATE_PATH.read_bytes(),
        _TEMPLATE_PATH.name,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Fill this in to propose new/updated pattern-of-study reference "
             "data, then upload it below to try it out.",
    )

st.sidebar.markdown("---")
pattern_upload = st.sidebar.file_uploader(
    "Upload filled Pattern of Study template (optional)",
    type=["xlsx"],
    help="Session-scoped only - this never edits the tool's built-in "
         "reference data, it only affects this run. Uploaded rows "
         "override/extend the built-in pattern table below. Used by "
         "Stage 5's on-pattern check.",
)

pattern_overrides = None
if pattern_upload is not None:
    try:
        pattern_overrides = load_pattern_overrides(io.BytesIO(pattern_upload.getvalue()))
        n_sessions = len(pattern_overrides)
        n_rows = sum(len(programs) for programs in pattern_overrides.values())
        st.sidebar.success(
            f"Loaded {n_rows} program/session row(s) across "
            f"{n_sessions} session(s): {', '.join(sorted(pattern_overrides.keys()))}."
        )
    except ValueError as e:
        st.sidebar.error(f"Couldn't read this template: {e}")

st.title("Reregistration Advisory")
st.caption(
    "One stop for the Modular Re-registration Timeline's advisory reports. "
    "Upload the roster once below - it drives every stage section."
)

with st.expander("Instructions", expanded=False):
    st.markdown(
        "Upload the real roster Ashlee's team sends Grant directly (e.g. "
        "'... Autumn to Spring Re-reg - AUT End progression Status ....xlsx', "
        "sheet 'Query1') - one file drives both the Stage 2/3 and Stage 5 "
        "sections below. See student_tracker/pipeline/live_roster.py for "
        "exactly what columns it needs and how they're used.  \n"
        "The sidebar's **Pattern of Study template** download/upload lets you "
        "propose new or corrected pattern-of-study data (e.g. for a "
        "commencement session or program the built-in table doesn't cover "
        "yet) and try it out immediately in Stage 5's on-pattern check, "
        "without a code change. It's session-scoped - uploading never "
        "edits the tool's actual built-in data, only this run's results."
    )

with st.expander("Stage coverage (Modular Re-registration Timeline)", expanded=False):
    st.caption(
        "From the Modular Re-registration Timeline (supplied 2026-08-21): 5 "
        "stages per semester, only some buildable today. Built = real logic "
        "against real sample data. Blocked = needs a source file/data "
        "granularity we've never seen a sample of - no guessed parsing. "
        "Undefined = the source Timeline sheet itself never filled this row in."
    )
    stage_df = pd.DataFrame(
        {
            "Semester": s.semester,
            "Stage": s.number,
            "Name": s.name,
            "Status": s.status,
            "When": s.when,
            "Population": s.population,
        }
        for s in STAGES
    )

    st.dataframe(
        stage_df.style.map(_style_status, subset=["Status"]),
        use_container_width=True,
        hide_index=True,
    )
    st.caption(
        "AUT Stage 2 & 3 and AUT Stage 5 are the sections below. AUT Stage 4 "
        "(the roster itself) is the input to Stage 5's upload, so it doesn't "
        "get its own section. See student_tracker/pipeline/stages.py for "
        "full per-stage detail, including data needs and notes for each "
        "blocked/undefined row."
    )

st.header("Upload roster")
uploaded = st.file_uploader(
    "Roster (.xlsx) — the live Autumn-to-Spring re-reg file (sheet 'Query1')",
    type=["xlsx"],
    help="The real file Ashlee's team sends Grant directly - carries both "
         "the forward-looking Spring registration AND this session's "
         "per-block/per-prep pass-fail results in one file. See "
         "live_roster.py for the exact expected columns.",
    key="live_upload",
)
run_btn = st.button("Generate", type="primary", disabled=not uploaded, key="live_generate")

if run_btn:
    try:
        df = live_roster.load_live_roster(io.BytesIO(uploaded.getvalue()))
        student_records, student_skipped = live_roster.to_student_records(df)
        history_records, history_skipped = live_roster.to_history_records(df)
        st.session_state["live_student_records"] = student_records
        st.session_state["live_student_skipped"] = student_skipped
        st.session_state["live_history_records"] = history_records
        st.session_state["live_history_skipped"] = history_skipped
    except ValueError as e:
        st.session_state.pop("live_student_records", None)
        st.session_state.pop("live_history_records", None)
        st.error(f"Couldn't read this file: {e}")

_student_records = st.session_state.get("live_student_records")
_history_records = st.session_state.get("live_history_records")
if _student_records is not None and st.session_state.get("live_student_skipped"):
    st.warning(
        f"Skipped {st.session_state['live_student_skipped']:,} row(s) with "
        "a blank/non-numeric Program value."
    )

st.header("Generate advice by stage")

with st.expander("Stage 2 & 3 — Block & Prep recommendations (post-AB4 results)", expanded=False):
    _render_stage_2_3(_history_records, st.session_state.get("live_history_skipped"))

with st.expander("Stage 5 — Applied progression outcomes recommendations (Conditional Enrolment)", expanded=True):
    _render_stage_5(_student_records, pattern_overrides)
