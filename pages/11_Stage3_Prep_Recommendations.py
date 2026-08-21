"""
Stage 3 Prep Recommendations — AUT/SPR Stage 3 of the Modular
Re-registration Timeline ("Recommendations for Prep subjects, post
AUT/SPR Results")
=====================================================================

STATUS: same role as pages/10_Stage2_Recommendations.py - displays
Grant's own already-computed prep-subject recommendation ("Prep Name"),
not an independently derived one. Confirmed 2026-08-22 that the same
AB4-for-SPR-style file already used for Stage 2 also carries this - Prep
Name follows the identical "you don't need to register" vs. a specific
subject code pattern as B1-B4 Name, just for the two general prep
subjects (GEDU0016/GEDU0017) instead of the block subjects. Kept as its
own page rather than a tab on the Stage 2 page, per Josiah's "split by
stage, not by type" preference (2026-08-22) - Stage 2 and Stage 3 are
different Timeline rows even though they currently share an input file
shape.

Nursing/UPP (9031) has no prep subjects at all (see history.py's module
docstring), so those rows have no prep advice to show - filtered out
here rather than shown as a confusing blank.
"""

from __future__ import annotations

import io

import pandas as pd
import streamlit as st

from student_tracker.pipeline.history import load_reregistration_history, to_history_records
from student_tracker.pipeline.stages import get_stage

st.set_page_config(page_title="Stage 3 Prep Recommendations", layout="wide")

st.title("Stage 3 Prep Recommendations")

_stage = get_stage("AUT", 3)
with st.expander("What is Stage 3?", expanded=False):
    st.markdown(
        f"**When:** {_stage.when}  \n"
        f"**Who:** {_stage.who}  \n"
        f"**Population:** {_stage.population}  \n"
        f"**What:** {_stage.what}  \n"
        f"**Data needs:** {_stage.data_needs}"
    )
    st.caption(
        "This page displays/filters/exports the prep-subject recommendation "
        "Grant already computes (the 'Prep Name' column) - it doesn't derive "
        "it itself. Uses the same input file as Stage 2 Recommendations."
    )

uploaded = st.file_uploader(
    "Reregistration recommendation list (.xlsx) — 'AB4 for SPR' style file",
    type=["xlsx"],
    help="Same file as the Stage 2 Recommendations page. Diploma programs "
         "only - Nursing/UPP (9031) has no prep subjects to advise on.",
)

run_btn = st.button("Generate report", type="primary", disabled=not uploaded)

if run_btn:
    try:
        df = load_reregistration_history(io.BytesIO(uploaded.getvalue()))
        st.session_state["stage3_records"] = to_history_records(df)
    except ValueError as e:
        st.session_state.pop("stage3_records", None)
        st.error(f"Couldn't read this file: {e}")

if "stage3_records" not in st.session_state:
    st.info("Upload the Stage 2/3 recommendation file (.xlsx) and click Generate.")
    st.stop()

records = [r for r in st.session_state["stage3_records"] if r.prep_advice is not None]
skipped_no_prep = len(st.session_state["stage3_records"]) - len(records)
if skipped_no_prep:
    st.caption(
        f"{skipped_no_prep} student(s) with no prep subject to advise on "
        "(e.g. Nursing/UPP) excluded from this view."
    )

report_df = pd.DataFrame(
    {
        "Student ID": r.student_id,
        "Student Name": r.student_name,
        "Program": r.program,
        "Commencement Period": r.commencement_period,
        "Prep 1 passed (GEDU0016)": r.prep_passed[0] if len(r.prep_passed) > 0 else None,
        "Prep 2 passed (GEDU0017)": r.prep_passed[1] if len(r.prep_passed) > 1 else None,
        "Prep advice": r.prep_advice,
        "Rereg Principles": r.rereg_principle,
        "Template": r.template,
    }
    for r in records
)

total = len(report_df)
needs_prep = int((report_df["Prep advice"] != "You do not need to register in a subject").sum())
m1, m2, m3 = st.columns(3)
m1.metric("Total students", f"{total:,}")
m2.metric("Need prep registration", f"{needs_prep:,}", f"{needs_prep / total:.0%} of total" if total else None)
m3.metric("Programs represented", report_df["Program"].nunique())

st.subheader("Per-student prep recommendations")
f1, f2 = st.columns(2)
program_filter = f1.multiselect(
    "Filter: Program", options=sorted(report_df["Program"].unique()), default=[]
)
advice_filter = f2.multiselect(
    "Filter: Prep advice", options=sorted(report_df["Prep advice"].dropna().unique()), default=[]
)

filtered_df = report_df
if program_filter:
    filtered_df = filtered_df[filtered_df["Program"].isin(program_filter)]
if advice_filter:
    filtered_df = filtered_df[filtered_df["Prep advice"].isin(advice_filter)]

st.dataframe(filtered_df, use_container_width=True, hide_index=True)
st.caption(f"Showing {len(filtered_df):,} of {len(report_df):,} students.")

excel_buffer = io.BytesIO()
with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
    report_df.to_excel(writer, sheet_name="Stage 3 Prep Recommendations", index=False)

st.download_button(
    "Download Stage 3 prep recommendations (Excel)",
    excel_buffer.getvalue(),
    "stage3_prep_recommendations.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
