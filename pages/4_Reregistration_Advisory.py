"""
Reregistration Advisory — pattern-based bucketing
====================================================

Sorts students into advisory buckets based on progression standing and
Spring registration completeness. Replaces the old Login Report page.

STATUS: skeleton, not a finished tool. Bucketing (stage 3 of the pipeline)
runs on real data; pattern lookup and subject substitution (stages 2 and 4)
are stubs blocked on open questions for Grant — see
student_tracker/pipeline/pattern_lookup.py and .../substitution.py. Every
bucket below carries a confidence tier — "high" (classification and action
both confirmed by Grant), "medium" (classification confirmed, action still
open), or "low" (classification itself is an inference, not yet confirmed)
— see the Open question column for medium/low buckets.
"""

from __future__ import annotations

import io
from datetime import date

import pandas as pd
import streamlit as st

from student_tracker.pipeline.ingest import load_roster, to_student_records
from student_tracker.pipeline.bucketing import bucket_all
from student_tracker.pipeline.report_builder import build_advisory_report
from student_tracker.pipeline.stages import STAGES, stage5_ce_only
from student_tracker.pipeline.glossary import glossary_dataframe

st.set_page_config(page_title="Reregistration Advisory", layout="wide")

st.title("Reregistration Advisory")
st.caption(
    "Sorts students into advisory buckets from progression standing and Spring "
    "registration completeness. This is a skeleton, not a finished tool — "
    "medium/low confidence buckets below are starting points for the Grant "
    "conversation, not confirmed rules."
)

SHORT_LABELS = {
    "on_pattern_continuing": "On pattern",
    "off_pattern_partial": "Partial",
    "success_coach_outreach": "Coach outreach",
    "zero_registration_unclear": "Zero reg.",
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
DAYS_PER_BLOCK = 28  # 4-week blocks, per the course-specific/block subject model

with st.expander("Instructions", expanded=False):
    st.markdown(
        "Upload a **Progression Outcomes roster** (.xlsx) — the 'Autumn for "
        "Spring Progression Outcomes' or ALA-style export with Calculated "
        "Standing and Spring Block 1-4 registration columns.  \n"
        "Picking a **term start date** shows which Spring block is likely "
        "running today, assuming back-to-back 4-week blocks — this is "
        "informational only for now (not Grant-confirmed), and isn't yet "
        "wired into the bucketing logic below."
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

    def _style_status(val):
        if val == "built":
            return "background-color: #E7F0EA; color: #2F6B4F"
        if val == "blocked":
            return "background-color: #FAF0DC; color: #A8720F"
        if val == "undefined":
            return "background-color: #F2E8E8; color: #8A4B4B"
        return ""

    st.dataframe(
        stage_df.style.map(_style_status, subset=["Status"]),
        use_container_width=True,
        hide_index=True,
    )
    st.caption(
        "Everything on this page today implements AUT/SPR Stage 4 (the "
        "roster itself) and Stage 5 (the advisory report below) - see "
        "student_tracker/pipeline/stages.py for full per-stage detail, "
        "including data needs and notes for each blocked/undefined row."
    )

term_choice = st.selectbox(
    "Term start date (Spring Block 1, Week 1 reference point)",
    options=list(TERM_OPTIONS.keys()) + ["Custom date"],
    index=None,
    placeholder="Select a term...",
    key="term_choice",
)

term_start = None
if term_choice == "Custom date":
    term_start = st.date_input("Custom term start date", value=None, key="term_custom_date")
elif term_choice is not None:
    term_start = TERM_OPTIONS[term_choice]

if term_start:
    days_elapsed = (date.today() - term_start).days
    if days_elapsed < 0:
        # %-d is Unix-only (breaks on Windows) - build the date string without it.
        formatted_date = f"{term_start:%A}, {term_start.day} {term_start:%B %Y}"
        st.caption(f"Term hasn't started yet — Block 1 begins {formatted_date}.")
    elif days_elapsed >= 4 * DAYS_PER_BLOCK:
        st.caption(
            f"{days_elapsed} days into term — past the assumed 4-block window "
            "(16 weeks). Likely exam period or between semesters."
        )
    else:
        current_block = days_elapsed // DAYS_PER_BLOCK + 1
        st.caption(
            f"{days_elapsed} days into term → assuming back-to-back 4-week "
            f"blocks, that's roughly **Spring Block {current_block}** today "
            "(not Grant-confirmed - informational only)."
        )
else:
    st.caption("Pick a term (or a custom date) to see which Spring block is likely running today.")

uploaded = st.file_uploader(
    "Roster (.xlsx) — Progression Outcomes export",
    type=["xlsx"],
    help="The 'Autumn for Spring Progression Outcomes' or ALA-style export with "
         "Calculated Standing and Spring Block 1-4 registration columns.",
)

if not uploaded:
    st.info("Upload a Progression Outcomes roster (.xlsx) to run the bucketing pass.")
    st.stop()

try:
    df = load_roster(io.BytesIO(uploaded.getvalue()))
except ValueError as e:
    st.error(f"Couldn't read this roster: {e}")
    st.stop()

records = to_student_records(df)
results = bucket_all(records)

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
    "High = classification and the resulting action are both confirmed by Grant — "
    "safe to act on. Medium = classification confirmed, but what to do about it is "
    "still open — see the open question. Low = the classification itself is an "
    "inference from the data, not yet confirmed by Grant at all."
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


def _style_confidence(val):
    if val == "high":
        return "background-color: #E7F0EA; color: #2F6B4F"
    if val == "medium":
        return "background-color: #FDEEDC; color: #B5590F"
    if val == "low":
        return "background-color: #FAF0DC; color: #A8720F"
    return ""


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
        "grant_question": st.column_config.TextColumn("Open question for Grant", width="large"),
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
    "mapping is empirically validated, not Grant-confirmed; see "
    "report_builder.py). Unknown = pattern can't be resolved yet for this "
    "student — see Reason. Advice is only populated for high-confidence "
    "buckets; everything else is flagged for advisor review pending Grant."
)

advisory_rows = build_advisory_report(records)

stage5_only = st.checkbox(
    "Scope to AUT/SPR Stage 5 (Conditional Enrolment only, per the "
    "Timeline - 'All students on Conditional Enrolment', SB2/SU2 Week "
    "1-2)",
    value=False,
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
)
confidence_filter = f2.multiselect(
    "Filter: Bucket confidence",
    options=["high", "medium", "low"],
    default=[],
)

filtered_df = advisory_df
if pattern_filter:
    filtered_df = filtered_df[filtered_df["On pattern"].isin(pattern_filter)]
if confidence_filter:
    filtered_df = filtered_df[filtered_df["Bucket confidence"].isin(confidence_filter)]

st.dataframe(filtered_df, use_container_width=True, hide_index=True)
st.caption(f"Showing {len(filtered_df):,} of {len(advisory_df):,} students.")

with st.expander("Glossary", expanded=False):
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
)
