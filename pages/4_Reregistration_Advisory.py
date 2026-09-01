"""
Reregistration Advisory (v2 - clean rebuild, 2026-09-01)
=======================================================

Three steps:
  1. Upload the progression workbook (sheet 'Query1').
  2. The tool lists each student's outstanding subjects + elective count.
  3. It advises next session's subjects from the student's commencement
     cohort and pattern of study, and hands back the same workbook with five
     '... Registration Advice' columns + an 'Advice Reason' column added
     (the original registration columns are left as-is for comparison).

Logic lives in student_tracker/rereg_advice.py; design in
docs/rereg_advice_v2_spec.md. The previous multi-stage version (and its
student_tracker/pipeline/* modules) was retired here - see git history.
"""

from __future__ import annotations

import io

import pandas as pd
import streamlit as st

from student_tracker import rereg_advice as ra

st.set_page_config(page_title="Reregistration Advisory", layout="wide")

st.title("Reregistration Advisory")
st.caption(
    "Upload the Autumn-to-Spring progression file. The tool recommends each "
    "student's next-session subjects and returns the same workbook with "
    "'... Registration Advice' columns added."
)

with st.expander("How the advice is built", expanded=False):
    st.markdown(
        """
- **Outstanding subjects** = the columns showing `1` (not the ones showing
  `"CODE Completed"`). The subject code for each `1` is recovered from
  classmates in the same program.
- **Prep** — if the program has a prep subject still outstanding, the earliest
  one is advised (one prep per session); any others are noted for later.
- **Blocks 1–4** — the student's cohort subjects for this Spring come first
  (from their commencement period), then any earlier outstanding subjects fill
  the remaining blocks, each displacing one elective to a later session.
- **Electives** — `+1 elective` per leftover block, down to `Electives Needed`.
- **Standing** — `Conditional Enrolment` caps the load at 30cp: 3 subjects max
  and the prep subject moves to Summer. `Exclusion` gets no advice. `At Risk`
  and new starters get a normal full load.
- **Carry** — anything not offered in Spring, or with no room, is named in the
  reason as *carry to a later session*.
- Cohort positions are only confirmed for `26 - Autumn Block 1/3` and
  `26 - Spring Block 1`. Everyone else is planned earliest-outstanding-first
  and flagged in the reason.
- Spring offering rules: `student_tracker/spring_offerings.json` (currently only
  Nursing is seeded — extend it from the handbook).
        """
    )

uploaded = st.file_uploader(
    "Progression workbook (.xlsx, sheet 'Query1')", type=["xlsx"], key="rereg_v2_upload"
)

if uploaded is None:
    st.info("Upload the workbook to generate advice.")
    st.stop()

try:
    df = ra.load_progression_file(io.BytesIO(uploaded.getvalue()))
except ValueError as exc:
    st.error(f"Couldn't read this file: {exc}")
    st.stop()

result = ra.build_advice(df)

# --- summary ---------------------------------------------------------------
total = len(result)
flagged = int(result[ra.REASON_COL].str.contains("NOTE:", regex=False).sum())
nothing = int(result[ra.REASON_COL].str.startswith("Nothing outstanding").sum())
ce = int(result[ra.REASON_COL].str.contains("30cp cap", regex=False).sum())
no_advice = int(result[ra.REASON_COL].str.contains("not eligible", regex=False).sum())

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Students", f"{total:,}")
c2.metric("Nothing outstanding", f"{nothing:,}")
c3.metric("Conditional Enrolment (30cp)", f"{ce:,}")
c4.metric("No advice (excluded)", f"{no_advice:,}")
c5.metric("Flagged for coach review", f"{flagged:,}")

# --- preview -------------------------------------------------------------
preview_cols = [
    "STUDENT_ID", "FIRST_NAME", "LAST_NAME", "PROGRAM_CD", "COMMENCEMENT_PERIOD",
    "Electives Needed", *ra.ADVICE_COLS, ra.REASON_COL,
]
preview_cols = [c for c in preview_cols if c in result.columns]
view = result[preview_cols]

programs = sorted(result["PROGRAM_CD"].dropna().astype(str).unique())
pick = st.multiselect("Filter by program", programs, default=[])
if pick:
    view = view[result["PROGRAM_CD"].astype(str).isin(pick)]
only_flagged = st.checkbox("Only rows flagged for coach review", value=False)
if only_flagged:
    view = view[result.loc[view.index, ra.REASON_COL].str.contains("NOTE:", regex=False)]

st.dataframe(view, use_container_width=True, hide_index=True)
st.caption(f"Showing {len(view):,} of {total:,} students.")

st.download_button(
    "Download workbook with advice",
    ra.to_workbook_bytes(result),
    "reregistration_advice.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)
