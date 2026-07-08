"""Streamlit dashboard for student engagement tracking.

This file lives at pages/3_Student_Tracker.py — run the multipage app
from the repo root with `streamlit run app.py`, then pick "Student
Tracker" from the sidebar.

Rebuilt for weekly-snapshot inputs. Each reporting point (a class
session) uploads THREE weekly report bundles, re-uploading ALL prior
weeks' files alongside the new one each time (stateless — nothing
persists between sessions):

  1. Class list (.xls) — once per block, doesn't change week to week
  2. Login Report (.xlsx) — one file per week, narrowed to that week's
     date range in the Blackboard UI. Upload one for the pre-teaching
     baseline (Week 0) plus one per teaching week since.
  3. Subject Activity Overview (.xls) — one file per week, same pattern
  4. User Activity in Forums (.xls) — one file per week, same pattern
  5. Grade Centre (.xls) — optional, periodic (not necessarily weekly)
  6. SCORM module reports (.pdf) — optional, one per module, standalone
     completion view — does not feed the weekly engagement engine

Each weekly file self-identifies its own week (from its own date-range
header / window text), so there's no need to manually tag which upload
is which week.
"""

from datetime import date, timedelta
from pathlib import Path
import tempfile
import yaml

import pandas as pd
import streamlit as st

from student_tracker.parsers import class_list as p_class
from student_tracker.parsers import login_report as p_login
from student_tracker.parsers import subject_activity as p_hours
from student_tracker.parsers import forum_activity as p_forum
from student_tracker.parsers import grade_centre as p_grade
from student_tracker.parsers import scorm_report as p_scorm
from student_tracker import metrics
from student_tracker import segmentation
from student_tracker import report

# ------------------------------------------------------------------
# Setup
# ------------------------------------------------------------------

st.set_page_config(page_title="Student Engagement Tracker", layout="wide")

CONFIG_PATH = Path(__file__).parent.parent / "student_tracker" / "config.yaml"
if CONFIG_PATH.exists():
    CONFIG = yaml.safe_load(CONFIG_PATH.read_text())
else:
    CONFIG = {"fade_threshold": 0.5}


def _save_upload(uploaded_file) -> str:
    suffix = Path(uploaded_file.name).suffix
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.getbuffer())
    tmp.close()
    return tmp.name


# ------------------------------------------------------------------
# Sidebar — inputs
# ------------------------------------------------------------------

st.sidebar.title("Inputs")
st.sidebar.markdown(
    "Re-upload **all** weeks' files each session (stateless — nothing is "
    "saved between runs). E.g. Week 0, then Week 0 + Week 1, then "
    "Week 0 + Week 1 + Week 2, and so on."
)

class_file = st.sidebar.file_uploader("Class list (.xls)", type=["xls"])

st.sidebar.divider()
st.sidebar.subheader("Weekly report bundles")
login_files = st.sidebar.file_uploader(
    "Login Reports (.xlsx) — one per week, narrowed date range",
    type=["xlsx"], accept_multiple_files=True,
)
hours_files = st.sidebar.file_uploader(
    "Subject Activity Overview (.xls) — one per week",
    type=["xls"], accept_multiple_files=True,
)
forum_files = st.sidebar.file_uploader(
    "User Activity in Forums (.xls) — one per week",
    type=["xls"], accept_multiple_files=True,
)

st.sidebar.divider()
st.sidebar.subheader("Optional")
grade_file = st.sidebar.file_uploader("Grade Centre (.xls)", type=["xls"])
scorm_files = st.sidebar.file_uploader(
    "SCORM module reports (.pdf) — one per module",
    type=["pdf"], accept_multiple_files=True,
)

st.sidebar.divider()
block_start = st.sidebar.date_input(
    "Block start date (Week 1, Day 1)",
    value=date.today() - timedelta(days=21),
    help="First day of teaching. The week before this is Week 0 "
         "(pre-teaching baseline — used only as the login-delta anchor).",
)
fade_threshold = st.sidebar.number_input(
    "Fade threshold (S6 cutoff)",
    min_value=0.1, max_value=0.9,
    value=float(CONFIG.get("fade_threshold", 0.5)),
    step=0.05,
    help="A student is 'Fading' (S6) if this week's dominant-signal value "
         "< threshold x last week's value on that same signal.",
)

# ------------------------------------------------------------------
# Main panel
# ------------------------------------------------------------------

st.title("Student Engagement Tracker")

if not all([class_file, login_files, hours_files]):
    st.info(
        "Upload the class list, at least one Login Report, and at least "
        "one Subject Activity Overview to begin. Forums, Grade Centre, "
        "and SCORM reports are optional."
    )
    with st.expander("How segments are defined"):
        st.markdown("""
- **Active (a given week)** = hours in subject > 0 **OR** login count > 0.
- **S1 — Never engaged:** 0 lifetime logins (or missing from every login report).
- **S2 — Pre-block ghost:** last login before block start, none since.
- **S3 — W1 ghost:** active in Week 1, inactive in every later week in the data.
- **S4 — Dropped this week:** active last week, inactive this week.
- **S5 — Returning engager:** inactive last week, active this week.
- **S6 — Fading engager:** active both weeks; this week's value on
  whichever of hours/logins was larger last week has dropped below the
  fade threshold.
- **S7 — True sustainer:** active both weeks, not fading.
- **S8 — Long-tail dropout:** active at some point, inactive in the last
  two weeks present.

Forum activity (accesses + messages) is shown as context — it does not
drive classification, since it's task-dependent and bursty in a way that
would distort the week-over-week comparison.
        """)
    st.stop()

# ------------------------------------------------------------------
# Parse inputs
# ------------------------------------------------------------------

with st.spinner("Parsing files…"):
    try:
        cls_df = p_class.parse(_save_upload(class_file))
        subject_code = p_class.detect_subject_code(cls_df)

        login_snapshots = []
        login_warnings = []
        for f in login_files:
            df, w_start, w_end = p_login.parse_with_window(_save_upload(f))
            if w_start is None:
                login_warnings.append(f.name)
            login_snapshots.append((df, w_start, w_end))

        hours_snapshots = []
        hours_warnings = []
        for f in hours_files:
            df, w_start, w_end = p_hours.parse_with_window(_save_upload(f))
            if w_start is None:
                hours_warnings.append(f.name)
            hours_snapshots.append((df, w_start, w_end))

        forum_snapshots = []
        forum_warnings = []
        for f in (forum_files or []):
            parsed = p_forum.parse(_save_upload(f))
            totals = p_forum.per_student_totals(parsed)
            if parsed["window_start"] is None:
                forum_warnings.append(f.name)
            forum_snapshots.append((totals, parsed["window_start"], parsed["window_end"]))

        if grade_file:
            grade_parsed = p_grade.parse(_save_upload(grade_file))
            grade_summary = p_grade.submission_summary(grade_parsed)
        else:
            grade_summary = None

        scorm_tables = {}
        for f in (scorm_files or []):
            module_title = Path(f.name).stem.replace("_", " ")
            scorm_tables[f.name] = p_scorm.parse(_save_upload(f), module_title=module_title)

    except Exception as e:
        st.error(f"Failed to parse one of the files: {e}")
        st.exception(e)
        st.stop()

for label, warn_list in (
    ("Login Report", login_warnings), ("Subject Activity Overview", hours_warnings),
    ("User Activity in Forums", forum_warnings),
):
    if warn_list:
        st.warning(
            f"Couldn't detect a date window in: {', '.join(warn_list)} "
            f"({label}). These files were skipped for week-tagging — "
            f"check the export still has its date-range header/text intact."
        )

# ------------------------------------------------------------------
# Build weekly tables + classify
# ------------------------------------------------------------------

hours_wide = metrics.stack_weekly(hours_snapshots, "hours", block_start)
cumulative_logins, logins_delta, latest_login_df = metrics.build_login_tables(
    login_snapshots, block_start
)
forum_wide = metrics.stack_weekly(forum_snapshots, "forum_interactions", block_start)

weeks = metrics.weeks_in_data(hours_wide, logins_delta, forum_wide)
teaching_weeks = [w for w in weeks if w >= 1]

if not teaching_weeks:
    st.error(
        "No teaching-week (Week 1+) data detected — check the block start "
        "date matches your files, and that at least one non-baseline "
        "week's files are uploaded."
    )
    st.stop()

summary = metrics.per_student_summary(
    cls_df, hours_wide, logins_delta, forum_wide, latest_login_df, grade_summary
)
summary = metrics.append_weekly_columns(
    summary, hours_wide, logins_delta, forum_wide, weeks, block_start
)
classified = segmentation.classify(
    summary, hours_wide, logins_delta, block_start, weeks, fade_threshold
)

reference_date = date.today()

# ------------------------------------------------------------------
# Display
# ------------------------------------------------------------------

week_list_str = ", ".join(f"W{w}" for w in weeks)
st.subheader(f"{subject_code or 'Subject'} — {len(classified)} enrolled students")
st.caption(
    f"Block start: {block_start} · Weeks in data: {week_list_str} "
    f"(W0 = pre-teaching baseline, used only as the login-delta anchor)"
)

with st.expander("File validation", expanded=False):
    vc1, vc2, vc3, vc4 = st.columns(4)
    vc1.metric("Class list", f"{len(cls_df)} students")
    vc2.metric("Login snapshots", f"{len(login_snapshots)} weeks")
    vc3.metric("Hours snapshots", f"{len(hours_snapshots)} weeks")
    vc4.metric("Forum snapshots", f"{len(forum_snapshots)} weeks")
    st.write(f"**Subject code (from class list):** {subject_code or 'unknown'}")
    st.write(f"**Weeks detected:** {week_list_str}")
    if grade_file:
        st.write(f"**Grade centre:** {grade_parsed['students']['student_code'].nunique()} students")
    if scorm_tables:
        st.write(f"**SCORM modules uploaded:** {len(scorm_tables)}")

# Segment counts
counts = segmentation.segment_counts(classified)
counts_display = counts[counts["count"] > 0]
st.markdown("### Segment distribution")
cols = st.columns(min(len(counts_display), 7) or 1)
for i, (_, row) in enumerate(counts_display.iterrows()):
    with cols[i % len(cols)]:
        st.metric(row["segment"], row["count"])
        st.caption(f"{row['pct']}% of cohort")

# Filters
st.markdown("### Student detail")
fc1, fc2, fc3 = st.columns(3)
seg_filter = fc1.multiselect("Filter by segment", options=segmentation.ALL_SEGMENTS, default=[])
if "attend_type" in classified.columns:
    attend_options = sorted(classified["attend_type"].dropna().unique().tolist())
    attend_filter = fc2.multiselect("Filter by attend type", options=attend_options, default=[])
else:
    attend_filter = []
search = fc3.text_input("Search by name or student code", value="")

filtered = classified.copy()
if seg_filter:
    filtered = filtered[filtered["segment"].isin(seg_filter)]
if attend_filter:
    filtered = filtered[filtered["attend_type"].isin(attend_filter)]
if search.strip():
    s = search.strip().lower()
    mask = (
        filtered["student_code"].astype(str).str.lower().str.contains(s, na=False)
        | filtered["first_name"].astype(str).str.lower().str.contains(s, na=False)
        | filtered["last_name"].astype(str).str.lower().str.contains(s, na=False)
    )
    filtered = filtered[mask]

display_cols = [
    "segment", "student_code", "preferred_name", "last_name", "attend_type",
    "last_active_week", "weeks_since_last_active", "weeks_active",
    "total_logins", "last_login_date", "days_since_last_login",
    "total_hours", "total_period_logins", "total_forum_interactions",
    "submission_rate", "avg_score_pct",
]
display_cols += [c for c in filtered.columns if c.startswith("W") and (
    c.endswith("_hours") or c.endswith("_logins") or c.endswith("_forum")
)]
display_cols = [c for c in display_cols if c in filtered.columns]

st.dataframe(
    filtered[display_cols].sort_values(["segment", "total_hours"], ascending=[True, False]),
    use_container_width=True,
    height=500,
    column_config={
        "submission_rate": st.column_config.ProgressColumn(
            "submission_rate", format="%.0f%%", min_value=0, max_value=1
        ),
    },
)
st.caption(f"Showing {len(filtered)} of {len(classified)} students.")

# SCORM (standalone, optional)
if scorm_tables:
    st.markdown("### SCORM module completion (context only — not part of segmentation)")
    for name, df in scorm_tables.items():
        with st.expander(name, expanded=False):
            if df.empty:
                st.write("No students parsed from this file.")
            else:
                st.dataframe(df, use_container_width=True)
                st.caption(
                    "grade / total_time_seconds / status are best-effort — "
                    "the SCORM export omits these for many students "
                    "depending on completion status; blank usually means "
                    "'no attempt data shown', not a parsing failure."
                )

# Export
st.markdown("### Export")
if st.button("Generate Excel report"):
    with st.spinner("Building workbook…"):
        wb = report.build_workbook(
            classified=classified,
            weeks=weeks,
            block_start_date=block_start,
            subject_code=subject_code,
            reference_date=reference_date,
            fade_threshold=fade_threshold,
        )
        data = report.to_bytes(wb)
    fname = f"engagement_report_{subject_code or 'subject'}_{reference_date.isoformat()}.xlsx"
    st.download_button(
        "Download Excel report", data=data, file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
