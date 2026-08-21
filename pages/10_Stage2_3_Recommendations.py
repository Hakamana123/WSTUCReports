"""
Stage 2/3 Recommendations — AUT Stage 2 ("Recommendations for Spring
Block subjects, post AB4 results") AND Stage 3 ("Recommendations for
Prep subjects, post AUT Results") of the Modular Re-registration Timeline
=====================================================================

Merged into one page (2026-08-22) - both stages need the exact same
uploaded file (the AB4-for-SPR-style recommendation list), so a separate
page per stage just meant uploading it twice for no benefit. See
stages.py for both stages' full entries.

STATUS: when the uploaded file carries Grant's own Rereg Principles/
Template/B1-4 Name/Prep Name columns (his already-computed
recommendation), this page displays/filters/exports them - it does NOT
independently derive them the way the Reregistration Advisory page
(Stage 5) computes its bucket from scratch, since we only know the shape
of Grant's output, not his actual classification rules.

When those columns are absent (e.g. a plain enhanced roster with just
the pass/fail history and no judgment layer applied yet), the page falls
back to a "Suggested Rereg Principle" computed from a threshold that was
100% consistent across thousands of real students (history.py's
EMPIRICAL VALIDATION section) - but ONLY for the cohorts/counts where
that threshold was unambiguous; everywhere else it stays blank rather
than guess. This is a best-effort fallback, not a claim to have
reproduced Grant's judgment.
"""

from __future__ import annotations

import io

import pandas as pd
import streamlit as st

from student_tracker.pipeline.history import (
    load_reregistration_history, to_history_records, failed_subject_codes_sem1_only,
    infer_current_commencement_year, rereg_principle_mismatch, suggested_rereg_principle,
)
from student_tracker.pipeline.stages import get_stage

st.set_page_config(page_title="Stage 2/3 Recommendations", layout="wide")

st.title("Stage 2/3 Recommendations")

_stage2 = get_stage("AUT", 2)
_stage3 = get_stage("AUT", 3)
with st.expander("What are Stage 2 and Stage 3?", expanded=False):
    st.markdown(
        f"**Stage 2 - {_stage2.name}**  \n"
        f"When: {_stage2.when}  \n"
        f"Population: {_stage2.population}  \n\n"
        f"**Stage 3 - {_stage3.name}**  \n"
        f"When: {_stage3.when}  \n"
        f"Population: {_stage3.population}"
    )
    st.caption(
        "Both stages draw on the same file - Stage 2 is the block-subject "
        "recommendation (B1-B4 Name), Stage 3 is the prep-subject "
        "recommendation (Prep Name). This page shows both together."
    )

uploaded = st.file_uploader(
    "Reregistration recommendation list (.xlsx) — 'AB4 for SPR' style file, "
    "or an enhanced roster with the same 'All Subjects' pass/fail column",
    type=["xlsx"],
    help="Reads every diploma and Nursing/UPP 'full population' sheet in "
         "the workbook automatically (not the Bulk/Complete/Exclusion "
         "breakdown sheets, which are subsets of those). Rereg Principles/ "
         "Template/B1-4 Name/Prep Name are all optional - the page still "
         "works without them, just with less to show.",
)

run_btn = st.button("Generate report", type="primary", disabled=not uploaded)

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
    st.stop()

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
        "QA check unavailable - couldn't infer which Autumn intake this "
        "file is currently reporting on (no Autumn-commencement rows "
        "found)."
    )
else:
    st.caption(
        f"QA check compares each Autumn-commencing diploma student's "
        f"Rereg Principles against a pattern empirically confirmed "
        f"against real data (see history.py) - only for {current_autumn_year} "
        "(this file's inferred current intake) and continuing students "
        f"from {current_autumn_year - 1}, and only where that pattern was "
        "100% consistent in real data. A blank QA check means either it "
        "matched, or this student's situation isn't one we have a "
        "confident rule for yet - it does NOT mean confirmed correct. "
        "This flags rows worth a second look, it doesn't override "
        "Grant's own classification."
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
    "Filter: Program", options=sorted(report_df["Program"].unique()), default=[]
)
principle_filter = f2.multiselect(
    "Filter: Rereg Principles",
    options=sorted(report_df["Rereg Principles"].dropna().unique()),
    default=[],
)
template_filter = f3.multiselect(
    "Filter: Template",
    options=sorted(report_df["Template"].dropna().unique()),
    default=[],
)
flagged_only = st.checkbox(f"Show only QA-flagged rows ({flagged_count:,})", value=False)

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
)
