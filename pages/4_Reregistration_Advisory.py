"""
Reregistration Advisory
=======================

Grant's testing calculator picks each student's next-session subjects; the v2
rule-tree wraps it with the progression-standing rules and the Coach View.

  1. Upload the progression workbook (sheet 'Query1').
  2. Pick the planning session you're advising for.
  3. For every student the tool asks Grant's calculator for the Prep / Block
     1-4 picks + Earliest Completion, then applies the 30cp Conditional
     Enrolment cap / Exclusion / At Risk rules on top. Where the calculator
     has no row (a session it doesn't cover, program 9034, or a status combo
     Grant's tool also can't resolve) it falls back to the v2 rule-tree and
     says so in 'Advice Source'.
  4. Download the same workbook back with the advice columns + a readable
     'Coach View' sheet.

Engines: student_tracker/rereg_merged.py (wiring), rereg_calc.py (Grant's
calculator port), rereg_advice.py (v2 rule-tree + Coach View).
"""

from __future__ import annotations

import io

import streamlit as st

from student_tracker import rereg_merged as rm

st.set_page_config(page_title="Reregistration Advisory", layout="wide")

st.title("Reregistration Advisory")
st.caption(
    "Upload the progression file and pick the target session. Grant's "
    "calculator drives the subject picks (offering pattern carried forward "
    "for sessions after 26 AUT); standing rules (30cp cap, Exclusion) and the "
    "Coach View are layered on top."
)

with st.expander("How the advice is built", expanded=False):
    st.markdown(
        """
- **Target session** is the session you're advising *for*. The teaching year
  runs `AUT -> SPR -> SUM`; Autumn and Spring each move a student four subjects
  along their pattern, Summer is catch-up only.
- **Subjects** come from Grant's testing calculator: offering pattern +
  fail-status -> Prep / Block 1-4 picks, with timetable clashes, "positions
  1 & 2 run every Summer", and elective placement all handled inside it.
- **Which pattern**: Grant computed `26 AUT` and `25 SUM` exactly. For any
  later Autumn / Spring the `26 AUT` pattern is **assumed to still hold**;
  for any later Summer, `25 SUM` — unless the reference data is updated. Those
  rows are tagged *"… pattern assumed"* and carry no completion estimate.
- **Rule-tree fallback** — used only when the calculator genuinely has no
  answer (a status combo Grant's own tool misses, program 9034, part-way
  targets, or Nursing outside Autumn). It places the student by cohort
  position and advises earliest-outstanding-first.
- **Standing** is applied on top: `Conditional Enrolment` caps the load at
  30cp (3 subjects, prep -> Summer). `Exclusion` gets no advice. `At Risk` /
  blank get the full load.
- **Advice Source** on each row names the engine and, where relevant, which
  pattern was assumed or why the calculator was skipped.
- **Custom target** — tick the box to advise part-way through a session
  (e.g. "as of Autumn Block 3" leaves only 2 blocks to fill; these always
  use the rule-tree).
        """
    )

col_a, col_b = st.columns([2, 3])
with col_a:
    named = st.selectbox(
        "Target session (advising for)",
        rm.PLANNING_SESSIONS,
        index=rm.PLANNING_SESSIONS.index(rm.DEFAULT_PLANNING_SESSION),
    )
    custom = st.checkbox("Custom — part-way through a session", value=False)
    if custom:
        cc1, cc2 = st.columns(2)
        with cc1:
            c_year = st.selectbox("Year", ["2026", "2027", "2028"], index=1)
        with cc2:
            c_sess = st.selectbox("Session", ["AUT", "SPR", "SUM"], index=0)
        c_block = st.slider("Advise from block", 1, 4, 1,
                            help="Block 1 = whole session. Block 3 = only blocks 3–4 left.")
        session = f"{c_year[2:]} {c_sess}" + (f" Block {c_block}" if c_block > 1 else "")
    else:
        session = named
    st.caption(f"Target: **{session}**")
with col_b:
    uploaded = st.file_uploader(
        "Progression workbook (.xlsx, sheet 'Query1')", type=["xlsx"], key="rereg_upload"
    )

if uploaded is None:
    st.info("Upload the workbook to generate advice.")
    st.stop()

try:
    df = rm.load_progression_file(io.BytesIO(uploaded.getvalue()))
except ValueError as exc:
    st.error(f"Couldn't read this file: {exc}")
    st.stop()

result = rm.build_advice(df, session=session)
coach_view = rm.build_coach_view(result)

# --- summary --------------------------------------------------------------
total = len(result)
src = result[rm.SOURCE_COL]
from_calc = int(src.str.startswith("Grant calculator").sum())
assumed = int(src.str.contains("pattern assumed", regex=False).sum())
ce = int(src.str.contains("CE 30cp cap", regex=False).sum())
excluded = int(src.str.contains("Exclusion", regex=False).sum())
fallback = int(src.str.startswith("v2 rule-tree").sum())
flagged = int(result[rm.REASON_COL].str.contains("NOTE:", regex=False).sum())

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Students", f"{total:,}")
c2.metric("From calculator", f"{from_calc:,}", help=f"of which {assumed:,} on an assumed (carried-forward) pattern")
c3.metric("Conditional Enrolment (30cp)", f"{ce:,}")
c4.metric("Rule-tree fallback", f"{fallback:,}")
c5.metric("Flagged for coach review", f"{flagged:,}")
if excluded:
    st.caption(f"{excluded:,} student(s) excluded — no advice.")

# --- preview (the Coach View sheet) --------------------------------------
st.caption(
    "Preview of the **Coach View** sheet. The full workbook (original columns "
    "+ advice appended) is the second sheet in the download."
)
view = coach_view

programs = sorted(coach_view["PROGRAM_CD"].dropna().astype(str).unique())
pick = st.multiselect("Filter by program", programs, default=[])
if pick:
    view = view[view["PROGRAM_CD"].astype(str).isin(pick)]

fcol1, fcol2 = st.columns(2)
with fcol1:
    only_flagged = st.checkbox("Only rows flagged for coach review", value=False)
with fcol2:
    only_fallback = st.checkbox("Only v2-fallback rows", value=False)
if only_flagged:
    view = view[view[rm.REASON_COL].str.contains("NOTE:", regex=False)]
if only_fallback:
    view = view[view[rm.SOURCE_COL].str.startswith("v2 rule-tree")]

st.dataframe(view, use_container_width=True, hide_index=True)
st.caption(
    f"Showing {len(view):,} of {total:,} students.  "
    "Progress key: ✓ = passed, ✗ = still to pass. "
    "Groups: prep · core blocks (in fours) · electives."
)

st.download_button(
    "Download workbook (Coach View + full sheet)",
    rm.to_workbook_bytes(result, coach_view),
    f"reregistration_advice_{session.replace(' ', '_')}.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)
