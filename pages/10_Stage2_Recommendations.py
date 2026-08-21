"""
Stage 2 Recommendations — AUT/SPR Stage 2 of the Modular Re-registration
Timeline ("Recommendations for Spring/Summer Block subjects, post AB4/SB4
results")
=====================================================================

STATUS: displays Grant's own already-computed recommendation file - it
does NOT independently derive Rereg Principles/Template the way the
Reregistration Advisory page (Stage 5) computes its bucket from scratch.
We know the shape of Grant's output (confirmed against two real files,
v1.1 and v2.1 - see student_tracker/pipeline/history.py) but not his
actual classification rules, so this page's job is to read/filter/export
his file, not to reproduce it. See stages.py for the full Stage 2 entry
this page implements.

2026-08-22: added a "QA check" column (history.rereg_principle_mismatch)
- flags rows where the assigned Rereg Principle disagrees with a
threshold that was 100% consistent across thousands of real students
(history.py's EMPIRICAL VALIDATION section). This is a second-look aid
for catching data anomalies, NOT a replacement for Grant's judgment - a
blank QA check can mean either "matches" or "no confident rule exists
for this student's cohort", and callers must not conflate the two.
"""

from __future__ import annotations

import io

import pandas as pd
import streamlit as st

from student_tracker.pipeline.history import (
    load_reregistration_history, to_history_records, failed_subject_codes_sem1_only,
    infer_current_commencement_year, rereg_principle_mismatch,
)
from student_tracker.pipeline.stages import get_stage

st.set_page_config(page_title="Stage 2 Recommendations", layout="wide")

st.title("Stage 2 Recommendations")

_stage = get_stage("AUT", 2)
with st.expander("What is Stage 2?", expanded=False):
    st.markdown(
        f"**When:** {_stage.when}  \n"
        f"**Who:** {_stage.who}  \n"
        f"**Population:** {_stage.population}  \n"
        f"**What:** {_stage.what}  \n"
        f"**Data needs:** {_stage.data_needs}"
    )
    st.caption(
        "This page displays/filters/exports the file Grant already produces "
        "for this stage - it doesn't compute Rereg Principles or Template "
        "itself, since we only know the shape of his output, not his "
        "underlying classification rules."
    )

uploaded = st.file_uploader(
    "Reregistration recommendation list (.xlsx) — 'AB4 for SPR' style file",
    type=["xlsx"],
    help="Reads every diploma and Nursing/UPP 'full population' sheet in "
         "the workbook automatically (not the Bulk/Complete/Exclusion "
         "breakdown sheets, which are subsets of those).",
)

run_btn = st.button("Generate report", type="primary", disabled=not uploaded)

if run_btn:
    try:
        df = load_reregistration_history(io.BytesIO(uploaded.getvalue()))
        st.session_state["stage2_records"] = to_history_records(df)
        st.session_state["stage2_skipped"] = len(df) - len(st.session_state["stage2_records"])
    except ValueError as e:
        st.session_state.pop("stage2_records", None)
        st.error(f"Couldn't read this file: {e}")

if "stage2_records" not in st.session_state:
    st.info("Upload the Stage 2 recommendation file (.xlsx) and click Generate.")
    st.stop()

records = st.session_state["stage2_records"]
skipped = st.session_state["stage2_skipped"]
if skipped:
    st.warning(
        f"Skipped {skipped} row(s) with an unrecognized subject-history "
        "shape (not diploma or Nursing/UPP)."
    )

current_autumn_year = infer_current_commencement_year(records)

report_df = pd.DataFrame(
    {
        "Student ID": r.student_id,
        "Student Name": r.student_name,
        "Program": r.program,
        "Commencement Period": r.commencement_period,
        "Rereg Principles": r.rereg_principle,
        "Template": r.template,
        "QA check": rereg_principle_mismatch(r, current_autumn_year) or "",
        "Subjects failed in Semester 1": failed_subject_codes_sem1_only(r),
        "Electives outstanding": r.electives_outstanding,
        "B1 advice": r.block_advice[0],
        "B2 advice": r.block_advice[1],
        "B3 advice": r.block_advice[2],
        "B4 advice": r.block_advice[3],
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

if current_autumn_year is None:
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
        "confident rule for yet (Nursing/UPP, a different commencement "
        "cohort, or a genuinely ambiguous case) - it does NOT mean "
        "confirmed correct. This flags rows worth a second look, it "
        "doesn't override Grant's own classification."
    )

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
    report_df.to_excel(writer, sheet_name="Stage 2 Recommendations", index=False)

st.download_button(
    "Download Stage 2 recommendations (Excel)",
    excel_buffer.getvalue(),
    "stage2_recommendations.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
