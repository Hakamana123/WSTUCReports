"""
Reregistration Advisory — pattern-based bucketing
====================================================

Sorts students into advisory buckets based on progression standing and
Spring registration completeness. Replaces the old Login Report page.

STATUS: skeleton, not a finished tool. Bucketing (stage 3 of the pipeline)
runs on real data; pattern lookup and subject substitution (stages 2 and 4)
are stubs blocked on open questions for Grant — see
student_tracker/pipeline/pattern_lookup.py and .../substitution.py. Every
bucket below is tagged "grounded" (confirmed rule) or "candidate"
(inference not yet confirmed by Grant) — see the Open question column for
candidate buckets.
"""

from __future__ import annotations

import io

import pandas as pd
import streamlit as st

from student_tracker.pipeline.ingest import load_roster, to_student_records
from student_tracker.pipeline.bucketing import bucket_all

st.set_page_config(page_title="Reregistration Advisory", layout="wide")

st.title("Reregistration Advisory")
st.caption(
    "Sorts students into advisory buckets from progression standing and Spring "
    "registration completeness. This is a skeleton, not a finished tool — "
    "candidate buckets below are starting points for the Grant conversation, "
    "not confirmed rules."
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
}

uploaded = st.sidebar.file_uploader(
    "Roster (.xlsx) — Progression Outcomes export",
    type=["xlsx"],
    help="The 'Autumn for Spring Progression Outcomes' or ALA-style export with "
         "Calculated Standing and Spring Block 1-4 registration columns.",
)

if not uploaded:
    st.info("Upload a Progression Outcomes roster (.xlsx) to run the bucketing pass.")
    st.stop()

df = load_roster(io.BytesIO(uploaded.getvalue()))
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
grounded_count = int((result_df["confidence"] == "grounded").sum())
candidate_count = int((result_df["confidence"] == "candidate").sum())
programs_represented = result_df["program"].nunique()

# --- Summary metrics ---
m1, m2, m3, m4 = st.columns(4)
m1.metric("Total students", f"{total_students:,}")
m2.metric(
    "Grounded buckets",
    f"{grounded_count:,}",
    f"{grounded_count / total_students:.0%} of total",
)
m3.metric(
    "Candidate buckets",
    f"{candidate_count:,}",
    f"{candidate_count / total_students:.0%} of total",
)
m4.metric("Programs represented", programs_represented)

# --- Bucket breakdown ---
st.subheader("Buckets")
st.caption(
    "Grounded = confirmed rule from the 4 Aug scoping meeting. "
    "Candidate = inferred from the data, not yet confirmed by Grant — see the "
    "open question for each."
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
    if val == "grounded":
        return "background-color: #E7F0EA; color: #2F6B4F"
    if val == "candidate":
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
