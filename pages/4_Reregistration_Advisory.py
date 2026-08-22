"""
Reregistration Advisory — main hub for the Modular Re-registration Timeline
=============================================================================

One page, one entry point. Each buildable stage gets its own collapsed
section below (with its own file upload — Stage 2/3 and Stage 5 need
genuinely different source files, see stages.py) rather than a separate
page per stage. Restructured 2026-08-22 at Josiah's request:
"Reregistration advice as the main page... generate advice for each
stage... stage 2 3 page should be collapsed basically." Stage 2/3 used
to live at pages/10_Stage2_3_Recommendations.py - that file is gone,
folded into the "Stage 2 & 3" section below.

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

from student_tracker.pipeline.ingest import load_roster, to_student_records
from student_tracker.pipeline.bucketing import bucket_all
from student_tracker.pipeline.report_builder import build_advisory_report
from student_tracker.pipeline.stages import STAGES, stage5_ce_only, get_stage
from student_tracker.pipeline.glossary import glossary_dataframe
from student_tracker.pipeline.timing import blocks_due as compute_blocks_due, DAYS_PER_BLOCK
from student_tracker.pipeline.pattern_lookup import load_pattern_overrides
from student_tracker.pipeline.history import (
    load_reregistration_history, to_history_records, failed_subject_codes_sem1_only,
    infer_current_commencement_year, rereg_principle_mismatch, suggested_rereg_principle,
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


def _render_stage_2_3() -> None:
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
        "Both stages draw on the same file - Stage 2 is the block-subject "
        "recommendation (B1-B4 Name), Stage 3 is the prep-subject "
        "recommendation (Prep Name), so this section covers both together "
        "rather than asking for the file twice. When the file carries "
        "Grant's own Rereg Principles/Template/B1-4 Name/Prep Name "
        "columns, this displays/filters/exports them - it does not "
        "independently derive them the way Stage 5 computes its bucket "
        "from scratch, since we only know the shape of Grant's output, "
        "not his actual classification rules. When those columns are "
        "absent, it falls back to a 'Suggested Rereg Principle' from an "
        "empirically-validated threshold - a best-effort fallback, not a "
        "claim to have reproduced Grant's judgment."
    )

    uploaded = st.file_uploader(
        "Reregistration recommendation list (.xlsx) — 'AB4 for SPR' style "
        "file, or an enhanced roster with the same 'All Subjects' pass/fail "
        "column",
        type=["xlsx"],
        help="Reads every diploma and Nursing/UPP 'full population' sheet "
             "in the workbook automatically (not the Bulk/Complete/"
             "Exclusion breakdown sheets, which are subsets of those). "
             "Rereg Principles/Template/B1-4 Name/Prep Name are all "
             "optional - this still works without them, just with less to "
             "show. This is a different file than Stage 5's Progression "
             "Outcomes roster below - it needs subject-level pass/fail "
             "history, which the roster doesn't carry.",
        key="stage23_upload",
    )

    run_btn = st.button(
        "Generate report", type="primary", disabled=not uploaded, key="stage23_generate",
    )

    if run_btn:
        try:
            df = load_reregistration_history(io.BytesIO(uploaded.getvalue()))
            st.session_state["stage23_records"] = to_history_records(df)
            st.session_state["stage23_skipped"] = len(df) - len(st.session_state["stage23_records"])
        except ValueError as e:
            st.session_state.pop("stage23_records", None)
            st.error(f"Couldn't read this file: {e}")

    if "stage23_records" not in st.session_state:
        st.info("Upload the Stage 2/3 recommendation file (.xlsx) and click Generate.")
        return

    records = st.session_state["stage23_records"]
    skipped = st.session_state["stage23_skipped"]
    if skipped:
        st.warning(
            f"Skipped {skipped} row(s) with an unrecognized subject-history "
            "shape (not diploma or Nursing/UPP)."
        )

    current_autumn_year = infer_current_commencement_year(records)
    has_rereg_data = any(r.rereg_principle is not None for r in records)

    report_df = pd.DataFrame(
        {
            "Student ID": r.student_id,
            "Student Name": r.student_name,
            "Program": r.program,
            "Commencement Period": r.commencement_period,
            "Rereg Principles": r.rereg_principle,
            "Suggested Rereg Principle": suggested_rereg_principle(r, current_autumn_year) or "",
            "Template": r.template,
            "QA check": rereg_principle_mismatch(r, current_autumn_year) or "",
            "Subjects failed in Semester 1": failed_subject_codes_sem1_only(r),
            "Electives outstanding": r.electives_outstanding,
            "B1 advice (Stage 2)": r.block_advice[0],
            "B2 advice (Stage 2)": r.block_advice[1],
            "B3 advice (Stage 2)": r.block_advice[2],
            "B4 advice (Stage 2)": r.block_advice[3],
            "Prep advice (Stage 3)": r.prep_advice or "",
        }
        for r in records
    )

    total = len(report_df)
    flagged_count = int((report_df["QA check"] != "").sum())
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total students", f"{total:,}")
    m2.metric("Programs represented", report_df["Program"].nunique())
    m3.metric("Rereg Principles categories", report_df["Rereg Principles"].nunique())
    m4.metric("QA flags", f"{flagged_count:,}")

    if not has_rereg_data:
        st.info(
            "This file has no Rereg Principles column at all - showing "
            "'Suggested Rereg Principle' instead, computed from the "
            "empirically-validated threshold. Blank means we don't have a "
            "confident basis for this student (Nursing/UPP, a different "
            "commencement cohort, or a genuinely ambiguous fail count) - "
            "it's not a claim that nothing's wrong, just that we can't say."
        )
    elif current_autumn_year is None:
        st.caption(
            "QA check unavailable - couldn't infer which Autumn intake "
            "this file is currently reporting on (no Autumn-commencement "
            "rows found)."
        )
    else:
        st.caption(
            f"QA check compares each Autumn-commencing diploma student's "
            f"Rereg Principles against a pattern empirically confirmed "
            f"against real data (see history.py) - only for "
            f"{current_autumn_year} (this file's inferred current intake) "
            f"and continuing students from {current_autumn_year - 1}, and "
            "only where that pattern was 100% consistent in real data. A "
            "blank QA check means either it matched, or this student's "
            "situation isn't one we have a confident rule for yet - it "
            "does NOT mean confirmed correct. This flags rows worth a "
            "second look, it doesn't override Grant's own classification."
        )

    if has_rereg_data:
        st.subheader("By Rereg Principles")
        principle_summary = (
            report_df.groupby(["Rereg Principles", "Template"])
            .size()
            .reset_index(name="Count")
            .sort_values("Count", ascending=False)
        )
        st.dataframe(principle_summary, use_container_width=True, hide_index=True)
        st.bar_chart(report_df["Rereg Principles"].value_counts())

    st.subheader("Per-student recommendations")
    f1, f2, f3 = st.columns(3)
    program_filter = f1.multiselect(
        "Filter: Program", options=sorted(report_df["Program"].unique()), default=[],
        key="stage23_program_filter",
    )
    principle_filter = f2.multiselect(
        "Filter: Rereg Principles",
        options=sorted(report_df["Rereg Principles"].dropna().unique()),
        default=[],
        key="stage23_principle_filter",
    )
    template_filter = f3.multiselect(
        "Filter: Template",
        options=sorted(report_df["Template"].dropna().unique()),
        default=[],
        key="stage23_template_filter",
    )
    flagged_only = st.checkbox(
        f"Show only QA-flagged rows ({flagged_count:,})", value=False, key="stage23_flagged_only",
    )

    filtered_df = report_df
    if flagged_only:
        filtered_df = filtered_df[filtered_df["QA check"] != ""]
    if program_filter:
        filtered_df = filtered_df[filtered_df["Program"].isin(program_filter)]
    if principle_filter:
        filtered_df = filtered_df[filtered_df["Rereg Principles"].isin(principle_filter)]
    if template_filter:
        filtered_df = filtered_df[filtered_df["Template"].isin(template_filter)]

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


def _render_stage_5(pattern_overrides) -> None:
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

    uploaded = st.file_uploader(
        "Roster (.xlsx) — Progression Outcomes export",
        type=["xlsx"],
        help="The 'Autumn for Spring Progression Outcomes' or ALA-style export "
             "with Calculated Standing and Spring Block 1-4 registration "
             "columns. This is a different file than Stage 2/3's "
             "recommendation list above - it doesn't carry subject-level "
             "pass/fail history, so it can't drive that section.",
        key="stage5_upload",
    )

    if not uploaded:
        st.info("Upload a Progression Outcomes roster (.xlsx) to run the bucketing pass.")
        return

    try:
        df = load_roster(io.BytesIO(uploaded.getvalue()))
        records = to_student_records(df)
    except ValueError as e:
        st.error(f"Couldn't read this roster: {e}")
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
            "Subjects advised": r.subjects_advised,
            "Subjects registered": r.subjects_registered,
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
    "Each stage below has its own upload, since different stages need "
    "different source files - the same roster won't satisfy both sections."
)

with st.expander("Instructions", expanded=False):
    st.markdown(
        "**Stage 2 & 3** need an 'AB4 for SPR'-style recommendation file "
        "(subject-level pass/fail history). **Stage 5** needs a "
        "Progression Outcomes roster (Calculated Standing + Spring Block "
        "registrations). These are different files carrying different "
        "data - uploading one where the other is expected will error out, "
        "that's expected rather than a bug.  \n"
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

st.header("Generate advice by stage")

with st.expander("Stage 2 & 3 — Block & Prep recommendations (post-AB4 results)", expanded=False):
    _render_stage_2_3()

with st.expander("Stage 5 — Applied progression outcomes recommendations (Conditional Enrolment)", expanded=True):
    _render_stage_5(pattern_overrides)
