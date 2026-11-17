"""Student Engagement Tracker - activity-log dashboard.

Run the multipage app from the repo root with `streamlit run app.py`, then
pick "Student Tracker" from the sidebar.

Built on the vUWS activity export (vUWS_activity_<COURSE>_<from>_to_<to>.csv):
one row per student click, with the item opened and when. A single export
covers the whole block, so there are no weekly re-uploads.

Inputs (nothing is stored between visits):
  1. vUWS activity export(s) (.csv) - required; several subjects at once is fine
  2. Class list(s) - optional; not a vUWS export, so any of the university
     class list (.xls), the Class List Collator output (.xlsx) or a sheet
     with a Student ID column (.xlsx / .csv). Adds 'Never visited'
     students, names, attend type
  3. Grade Centre export(s) (.xls) - optional; adds submissions and marks

Class lists and Grade Centres are matched to a subject by which subject's
students they share the most IDs with, so file names don't matter.

Definitions live in student_tracker/activity_metrics.py and are shown in
the "How it's measured" tab.
"""

from __future__ import annotations

import io
import tempfile
from datetime import date
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

from student_tracker import activity_metrics as am
from student_tracker import activity_report
from student_tracker.parsers import activity_log as p_activity
from student_tracker.parsers import grade_centre as p_grade
from student_tracker.parsers import roster as p_roster

st.set_page_config(page_title="Student Engagement Tracker", layout="wide")

BLUE = "#2a78d6"
MUTED = "#c3c2b7"
SEQ_BLUE = ["#cde2fb", "#9ec5f4", "#6da7ec", "#3987e5", "#256abf", "#184f95", "#0d366b"]
SEGMENT_ICON = {
    "Engaged": "🟢", "Fading": "🟡", "Lapsed": "🟠",
    "Never visited": "🔴", "Not on class list": "⚪",
}
STAGE_ICON = {
    "Submitted": "🟢", "Opened submission point": "🟡",
    "Read About page": "🟠", "Opened other items only": "🟠", "Not opened": "🔴",
}


# ------------------------------------------------------------------
# Parsing (cached on file bytes)
# ------------------------------------------------------------------

def _tmp(data: bytes, suffix: str) -> str:
    f = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    f.write(data)
    f.close()
    return f.name


@st.cache_data(show_spinner=False)
def parse_activity(data: bytes, name: str) -> dict:
    return p_activity.parse(io.BytesIO(data), filename=name)


@st.cache_data(show_spinner=False)
def parse_class_list(data: bytes, name: str) -> pd.DataFrame:
    return p_roster.parse(data, name)


@st.cache_data(show_spinner=False)
def parse_grade_centre(data: bytes, name: str) -> dict:
    return p_grade.parse(_tmp(data, Path(name).suffix))


def _best_subject(codes: set[str], subject_students: dict[str, set[str]]) -> tuple[str | None, int]:
    best, overlap = None, 0
    for subj, studs in subject_students.items():
        n = len(codes & studs)
        if n > overlap:
            best, overlap = subj, n
    return best, overlap


def as_pct(df: pd.DataFrame, cols) -> pd.DataFrame:
    """Proportions -> 0-100 for whole-percent progress columns."""
    df = df.copy()
    for c in cols:
        if c in df.columns:
            df[c] = df[c] * 100
    return df


def fmt_pct(x: float) -> str:
    return f"{x * 100:.0f}%"


def fmt_hours(h: float) -> str:
    if h < 1:
        return f"{h * 60:.0f} min"
    return f"{h:.1f} h"


# ------------------------------------------------------------------
# Sidebar - inputs
# ------------------------------------------------------------------

st.sidebar.title("Inputs")
st.sidebar.caption("Nothing is saved between visits - re-upload each time.")
activity_files = st.sidebar.file_uploader(
    "vUWS activity export(s) (.csv)", type=["csv"], accept_multiple_files=True,
    help="One file per subject. Several subjects at once is fine.",
)
class_files = st.sidebar.file_uploader(
    "Class list(s) - optional", type=["xls", "xlsx", "csv"], accept_multiple_files=True,
    help="The university class list, the Class List Collator output, or any sheet with a "
         "Student ID column. Adds students who never visited, preferred names and attend type.",
)
grade_files = st.sidebar.file_uploader(
    "Grade Centre export(s) (.xls) - optional", type=["xls"], accept_multiple_files=True,
    help="Adds submissions (including Turnitin ones the activity log can't see) and marks.",
)

st.title("Student Engagement Tracker")

if not activity_files:
    st.info(
        "Upload at least one **vUWS activity export** (.csv) to begin. "
        "Class lists and Grade Centre exports are optional extras."
    )
    st.markdown(
        "The activity export is one row per student click - which item they "
        "opened and when. From that the tracker works out visits, bounce "
        "rate, estimated time, which content reaches the class, how far each "
        "student has got with each assessment, and when students study. "
        "See **How it's measured** once a file is loaded."
    )
    st.stop()

# ------------------------------------------------------------------
# Parse
# ------------------------------------------------------------------

events_parts, metas, problems = [], [], []
for f in activity_files:
    try:
        parsed = parse_activity(f.getvalue(), f.name)
    except Exception as e:  # bad file: report and carry on with the rest
        problems.append(f"**{f.name}** - {e}")
        continue
    events_parts.append(parsed["events"])
    metas.append(parsed["meta"])

for msg in problems:
    st.error(msg)
if not events_parts:
    st.stop()

all_events = pd.concat(events_parts, ignore_index=True).drop_duplicates(
    subset=["subject_code", "student_code", "ts", "event", "item"]
)
subjects = sorted(all_events["subject_code"].unique())
subject_students = {
    s: set(all_events.loc[all_events["subject_code"] == s, "student_code"]) for s in subjects
}

rosters: dict[str, pd.DataFrame] = {}
file_notes = []
for f in class_files or []:
    try:
        cls = parse_class_list(f.getvalue(), f.name)
    except Exception as e:
        st.warning(f"Class list **{f.name}** skipped - {e}")
        continue
    # A collated list can hold several subjects: match each subject code it
    # names to an uploaded subject; rows with no / unknown code go to the
    # subject they share the most students with.
    for code, part in cls.groupby("subject_code"):
        codes = set(part["student_code"])
        subj = code if code in subjects else _best_subject(codes, subject_students)[0]
        if subj is None:
            continue
        part = part.drop(columns=["subject_code"])
        rosters[subj] = pd.concat([rosters.get(subj), part]).drop_duplicates("student_code")
        overlap = len(codes & subject_students[subj])
        file_notes.append(
            f"Class list `{f.name}` → **{subj}** ({len(codes)} students, {overlap} in the activity log)"
        )
    if not any(n.startswith(f"Class list `{f.name}`") for n in file_notes):
        st.warning(f"Class list **{f.name}** shares no students with any uploaded subject - skipped.")

grades: dict[str, dict] = {}
for f in grade_files or []:
    try:
        g = parse_grade_centre(f.getvalue(), f.name)
    except Exception as e:
        st.warning(f"Grade Centre **{f.name}** skipped - {e}")
        continue
    codes = set(g["students"]["student_code"])
    subj, overlap = _best_subject(codes, subject_students)
    if subj is None:
        st.warning(f"Grade Centre **{f.name}** shares no students with any uploaded subject - skipped.")
        continue
    grades[subj] = g
    file_notes.append(
        f"Grade Centre `{f.name}` → **{subj}** ({len(g['assessments'])} assessment columns, "
        f"{overlap} students matched)"
    )

# ------------------------------------------------------------------
# Settings
# ------------------------------------------------------------------

window_ends = [m["window_end"] for m in metas if m["window_end"]]
default_as_at = min(max(window_ends), date.today()) if window_ends else all_events["ts"].max().date()

st.sidebar.divider()
st.sidebar.subheader("Settings")
as_at = st.sidebar.date_input(
    "As at", value=default_as_at,
    help="'Days since last visit' and the last-7-days figures count back from this date. "
         "Defaults to the export's end date.",
)
gap_minutes = st.sidebar.slider(
    "New session after a gap of (minutes)", 10, 60, 30, step=5,
    help="Clicks closer together than this count as one visit. 30 is the usual web-analytics setting.",
)
active_days = st.sidebar.number_input("Engaged = visited within (days)", 1, 30, 7)
lapse_days = st.sidebar.number_input(
    "Lapsed = no visit for more than (days)", int(active_days) + 1, 60, max(14, int(active_days) + 1),
)

# ------------------------------------------------------------------
# Subject selection + compute
# ------------------------------------------------------------------


def compute(subject: str) -> dict:
    ev_raw = all_events[all_events["subject_code"] == subject]
    ev = am.sessionise(ev_raw, gap_minutes)
    sessions = am.session_table(ev)
    roster = rosters.get(subject)
    students = am.student_summary(ev, sessions, as_at, roster, int(active_days), int(lapse_days))
    enrolled_codes = students.loc[students["_on_roster"], "student_code"].tolist()
    g = grades.get(subject)
    grade_sub = am.grade_submissions(g) if g else None
    if g:
        summary = p_grade.submission_summary(g)
        students = students.merge(summary, on="student_code", how="left")
    content = am.content_table(ev, len(enrolled_codes))
    funnel, matrix = am.assessment_progress(ev, enrolled_codes, grade_sub)
    return {
        "ev": ev, "sessions": sessions, "students": students,
        "enrolled_codes": enrolled_codes, "content": content,
        "funnel": funnel, "matrix": matrix,
        "weekly": am.weekly_activity(ev, sessions, len(enrolled_codes)),
        "heat": am.hour_heatmap(ev),
        "kpis": am.kpis(students, sessions, ev, int(active_days)),
        "course_name": ev_raw["course_name"].iloc[0] if len(ev_raw) else subject,
        "has_roster": roster is not None,
        "has_grades": g is not None,
    }


if len(subjects) > 1:
    subject = st.selectbox("Subject", subjects)
else:
    subject = subjects[0]

with st.spinner("Crunching clicks…"):
    R = compute(subject)

students = R["students"]
names = dict(zip(students["student_code"], students["display_name"]))
K = R["kpis"]
first_ts, last_ts = R["ev"]["ts"].min(), R["ev"]["ts"].max()

st.subheader(R["course_name"])
st.caption(
    f"As at {as_at:%d %b %Y} · clicks from {first_ts:%d %b} to {last_ts:%d %b %Y} · "
    f"class: {'class list' if R['has_roster'] else 'students in the activity log (no class list uploaded)'}"
    f"{' · Grade Centre loaded' if R['has_grades'] else ''}"
)

with st.expander("Files", expanded=False):
    for m in metas:
        st.markdown(
            f"`{m['filename']}` - **{m['subject_code']}**, {m['rows_read']:,} rows, "
            f"{m['students']} students. Dropped {m['dropped_preview']} staff preview row(s)"
            + (f", {m['dropped_bad_date']} with unreadable dates" if m["dropped_bad_date"] else "")
            + "."
        )
    for n in file_notes:
        st.markdown(n)

if K["not_on_list"]:
    st.warning(
        f"{K['not_on_list']} student(s) have clicks but aren't on the class list - "
        "usually a withdrawal. They're shown as **Not on class list** and left out of the class figures."
    )

tab_names = ["Overview", "Students", "Content", "Assessments", "Timing"]
if len(subjects) > 1:
    tab_names.append("Compare subjects")
tab_names += ["Export", "How it's measured"]
tabs = dict(zip(tab_names, st.tabs(tab_names)))

# ------------------------------------------------------------------
# Overview
# ------------------------------------------------------------------

with tabs["Overview"]:
    n = K["enrolled"]
    c = st.columns(5)
    c[0].metric("Students", n)
    c[1].metric(f"Visited in last {active_days} days",
                f"{K['active_recent']} ({fmt_pct(K['active_recent'] / n)})" if n else 0)
    c[2].metric("Never visited", K["never_visited"] if R["has_roster"] else "-",
                help=None if R["has_roster"] else "Needs a class list - the activity log only contains students who clicked.")
    c[3].metric("Median visits per student", f"{K['median_sessions']:.0f}")
    c[4].metric("Median est. time per student", fmt_hours(K["median_hours"]),
                help="Lower bound - see How it's measured.")
    c = st.columns(5)
    c[0].metric("Bounce rate", fmt_pct(K["bounce_rate"]),
                help="Share of visits where the student entered the subject and opened nothing.")
    c[1].metric("Median visit length", f"{K['median_session_min']:.1f} min",
                help="First click to last click. One-click visits count as 0.")
    c[2].metric("Median content coverage", fmt_pct(K["median_coverage"]),
                help="Share of the subject's items a typical student has opened.")
    c[3].metric("Announcement reach", fmt_pct(K["announcement_reach"]),
                help="Share of the class who opened at least one announcement on vUWS.")
    c[4].metric("Items in subject", len(R["content"]),
                help="Items at least one student opened. Items nobody opened don't appear in the export.")

    st.markdown("#### Where the class is")
    seg_counts = students["segment"].value_counts().reindex(am.SEGMENTS, fill_value=0)
    help_ctx = {"active": active_days, "active_plus": active_days + 1, "lapse": lapse_days}
    shown = [s for s in am.SEGMENTS if seg_counts[s] or s in ("Engaged", "Fading", "Lapsed")
             or (s == "Never visited" and R["has_roster"])]
    cols = st.columns(len(shown))
    for col, s in zip(cols, shown):
        col.metric(f"{SEGMENT_ICON[s]} {s}", int(seg_counts[s]),
                   help=am.SEGMENT_HELP[s].format(**help_ctx))

    w = R["weekly"]
    if not w.empty:
        st.markdown("#### Week by week")
        left, right = st.columns(2)
        order = w["week_label"].tolist()
        x_axis = alt.Axis(labelAngle=0, labelLimit=120, labelOverlap=False)
        active_chart = alt.Chart(w).mark_bar(color=BLUE, cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
            x=alt.X("week_label:N", sort=order, title="Week commencing", axis=x_axis),
            y=alt.Y("pct_active:Q", title="% of class who visited", axis=alt.Axis(format="%"),
                    scale=alt.Scale(domain=[0, 1])),
            tooltip=[alt.Tooltip("week_label:N", title="Week commencing"),
                     alt.Tooltip("active_students:Q", title="Students who visited"),
                     alt.Tooltip("pct_active:Q", title="% of class", format=".0%")],
        ).properties(height=260, title="Students who visited each week")
        left.altair_chart(active_chart, width="stretch")

        sw = R["sessions"].assign(week=R["sessions"]["start"].dt.to_period("W-SUN").dt.start_time)
        sw["Visit type"] = sw["bounce"].map({True: "Bounced (opened nothing)", False: "Opened something"})
        sw = sw.groupby(["week", "Visit type"]).size().rename("visits").reset_index()
        sw = sw.merge(w[["week", "week_label"]], on="week", how="left")
        visits_chart = alt.Chart(sw).mark_bar(stroke="white", strokeWidth=1).encode(
            x=alt.X("week_label:N", sort=order, title="Week commencing", axis=x_axis),
            y=alt.Y("visits:Q", title="Visits", stack="zero"),
            color=alt.Color("Visit type:N", scale=alt.Scale(
                domain=["Opened something", "Bounced (opened nothing)"], range=[BLUE, MUTED]),
                legend=alt.Legend(orient="bottom", title=None)),
            order=alt.Order("Visit type:N", sort="descending"),
            tooltip=[alt.Tooltip("week_label:N", title="Week commencing"), "Visit type:N", "visits:Q"],
        ).properties(height=330, title="Visits each week")  # taller: its legend takes plot space
        right.altair_chart(visits_chart, width="stretch")

# ------------------------------------------------------------------
# Students
# ------------------------------------------------------------------

with tabs["Students"]:
    f1, f2 = st.columns([2, 1])
    seg_filter = f1.multiselect("Segment", am.SEGMENTS, default=[])
    search = f2.text_input("Search name or student ID")
    view = students.copy()
    if seg_filter:
        view = view[view["segment"].isin(seg_filter)]
    if search.strip():
        s = search.strip().lower()
        view = view[view["display_name"].fillna("").str.lower().str.contains(s, regex=False)
                    | view["student_code"].str.contains(s, regex=False)]
    view["segment"] = view["segment"].astype(str).map(lambda x: f"{SEGMENT_ICON.get(x, '')} {x}")

    cols = ["segment", "student_code", "display_name", "attend_type", "last_seen",
            "days_since_last", "sessions", "sessions_last_7d", "sessions_prev_7d",
            "active_days", "est_hours", "avg_session_min", "bounce_rate", "coverage",
            "items_opened", "announcements_read", "submissions_logged",
            "assessments_submitted", "submission_rate", "avg_score_pct"]
    cols = [c for c in cols if c in view.columns]
    st.dataframe(
        as_pct(view[cols], ["bounce_rate", "coverage", "submission_rate"]),
        width="stretch", hide_index=True, height=460,
        column_config={
            "segment": "Segment",
            "student_code": "Student ID",
            "display_name": "Name",
            "attend_type": "Attend type",
            "last_seen": st.column_config.DatetimeColumn("Last visit", format="DD MMM, HH:mm"),
            "days_since_last": st.column_config.NumberColumn("Days since", format="%d"),
            "sessions": st.column_config.NumberColumn("Visits", help="Sessions in the export window"),
            "sessions_last_7d": st.column_config.NumberColumn("Visits last 7d"),
            "sessions_prev_7d": st.column_config.NumberColumn("Visits prev 7d"),
            "active_days": st.column_config.NumberColumn("Days active"),
            "est_hours": st.column_config.NumberColumn("Est. hours", format="%.1f", help="Lower bound"),
            "avg_session_min": st.column_config.NumberColumn("Avg visit (min)", format="%.1f"),
            "bounce_rate": st.column_config.ProgressColumn("Bounce rate", format="%.0f%%", min_value=0, max_value=100),
            "coverage": st.column_config.ProgressColumn("Coverage", format="%.0f%%", min_value=0, max_value=100,
                                                        help="Share of the subject's items opened"),
            "items_opened": st.column_config.NumberColumn("Items opened"),
            "announcements_read": st.column_config.NumberColumn("Announcements read"),
            "submissions_logged": st.column_config.NumberColumn("Submit clicks", help="vUWS assessment submissions in the log (Turnitin not included)"),
            "assessments_submitted": st.column_config.NumberColumn("Submitted (GC)", help="From the Grade Centre"),
            "submission_rate": st.column_config.ProgressColumn("Submission rate (GC)", format="%.0f%%", min_value=0, max_value=100),
            "avg_score_pct": st.column_config.NumberColumn("Avg mark % (GC)", format="%.0f"),
        },
    )
    st.caption(f"Showing {len(view)} of {len(students)} students.")

    st.markdown("#### Student profile")
    options = students["student_code"].tolist()
    pick = st.selectbox(
        "Student", options,
        format_func=lambda c: f"{names.get(c) or '(no name)'} - {c}",
    )
    row = students[students["student_code"] == pick].iloc[0]
    sev = R["ev"][R["ev"]["student_code"] == pick]
    pc = st.columns([1.6, 1, 1, 1, 1])
    pc[0].metric(f"{SEGMENT_ICON[str(row['segment'])]} Segment", str(row["segment"]))
    pc[1].metric("Last visit", f"{row['last_seen']:%d %b}" if pd.notna(row.get("last_seen")) else "never")
    pc[2].metric("Visits", int(row["sessions"]))
    pc[3].metric("Est. time", fmt_hours(float(row["est_hours"])))
    pc[4].metric("Coverage", fmt_pct(float(row["coverage"])))

    if sev.empty:
        st.info("No clicks in the activity export for this student.")
    else:
        daily = am.daily_clicks(sev)
        day_range = pd.DataFrame({"day": pd.date_range(first_ts.normalize(), last_ts.normalize())})
        daily = day_range.merge(daily[["day", "clicks"]], on="day", how="left").fillna({"clicks": 0})
        st.altair_chart(
            alt.Chart(daily).mark_bar(color=BLUE, size=9, cornerRadiusTopLeft=2, cornerRadiusTopRight=2).encode(
                x=alt.X("day:T", title=None, axis=alt.Axis(format="%d %b")),
                y=alt.Y("clicks:Q", title="Clicks"),
                tooltip=[alt.Tooltip("day:T", title="Day", format="%a %d %b"), "clicks:Q"],
            ).properties(height=180, title="Clicks per day"),
            width="stretch",
        )
        p1, p2 = st.columns([3, 2])
        with p1:
            st.markdown("**Visits** (most recent first)")
            ss = R["sessions"][R["sessions"]["student_code"] == pick].sort_values("start", ascending=False)
            st.dataframe(
                ss[["start", "minutes", "clicks", "items"]], width="stretch", hide_index=True, height=300,
                column_config={
                    "start": st.column_config.DatetimeColumn("Started", format="ddd DD MMM, HH:mm"),
                    "minutes": st.column_config.NumberColumn("Minutes", format="%.0f"),
                    "clicks": "Clicks",
                    "items": st.column_config.TextColumn("What they opened", width="large"),
                },
            )
        with p2:
            if not R["matrix"].empty and len(R["matrix"].columns) > 1:
                mrow = R["matrix"][R["matrix"]["student_code"] == pick]
                if not mrow.empty:
                    st.markdown("**Assessments**")
                    for c in [c for c in R["matrix"].columns if c.startswith("A")]:
                        stg = mrow.iloc[0][c]
                        st.markdown(f"{STAGE_ICON.get(stg, '')} Assessment {c[1:]}: {stg}")
            opened = set(sev.loc[sev["event"] == "CONTENT_ACCESS", "item"])
            not_opened = [i for i in R["content"]["item"] if i not in opened] if not R["content"].empty else []
            st.markdown(f"**Not yet opened** ({len(not_opened)} of {len(R['content'])} items)")
            st.dataframe(pd.DataFrame({"Item": not_opened}), width="stretch", hide_index=True, height=200)

# ------------------------------------------------------------------
# Content
# ------------------------------------------------------------------

with tabs["Content"]:
    ct = R["content"]
    if ct.empty:
        st.info("No content opens in this export.")
    else:
        kinds = sorted(ct["kind"].unique())
        kind_filter = st.multiselect("Item type", kinds, default=[])
        cview = ct[ct["kind"].isin(kind_filter)] if kind_filter else ct
        st.dataframe(
            as_pct(cview[["item", "kind", "reach", "students", "session_views", "views_per_student",
                          "median_time_s", "exit_rate", "first_opened", "last_opened"]], ["reach", "exit_rate"]),
            width="stretch", hide_index=True, height=480,
            column_config={
                "item": st.column_config.TextColumn("Item", width="large"),
                "kind": "Type",
                "reach": st.column_config.ProgressColumn("Reach", format="%.0f%%", min_value=0, max_value=100,
                                                         help="Share of the class who opened it at least once"),
                "students": "Students",
                "session_views": st.column_config.NumberColumn("Views", help="Visits in which it was opened (repeat clicks in one visit count once)"),
                "views_per_student": st.column_config.NumberColumn("Views / student", format="%.1f"),
                "median_time_s": st.column_config.NumberColumn("Median time (s)", format="%.0f",
                                                               help="Blank for external tools / SCORM - time spent there isn't recorded"),
                "exit_rate": st.column_config.ProgressColumn("Exit rate", format="%.0f%%", min_value=0, max_value=100,
                                                             help="Share of its views that were the last click of the visit"),
                "first_opened": st.column_config.DatetimeColumn("First opened", format="DD MMM"),
                "last_opened": st.column_config.DatetimeColumn("Last opened", format="DD MMM"),
            },
        )
        st.caption(
            "Items nobody has opened don't appear in the export, so they can't be listed here."
        )
        st.markdown("#### Who hasn't opened an item?")
        item = st.selectbox("Item", ct.sort_values("reach")["item"].tolist())
        missing = am.students_not_opened(R["ev"], item, R["enrolled_codes"])
        if missing:
            st.dataframe(
                pd.DataFrame({"Student ID": missing, "Name": [names.get(m) for m in missing],
                              "Segment": [str(students.set_index("student_code").loc[m, "segment"]) for m in missing]}),
                width="stretch", hide_index=True,
            )
        else:
            st.success("Everyone in the class has opened this.")

# ------------------------------------------------------------------
# Assessments
# ------------------------------------------------------------------

with tabs["Assessments"]:
    fn = R["funnel"]
    if fn.empty:
        st.info("No items with 'Assessment <number>' in their title were found.")
    else:
        stage_cols = {
            "class_size": "1 Class",
            "opened_any": "2 Opened any item",
            "read_about": "3 Read About page",
            "opened_submission_point": "4 Opened submission point",
            "submitted": "5 Submitted",
        }
        long = fn.melt(id_vars=["assessment"], value_vars=list(stage_cols), var_name="stage", value_name="students")
        long["stage"] = long["stage"].map(stage_cols)
        long = long.dropna(subset=["students"])
        long["pct"] = long["students"] / fn["class_size"].max()
        chart = alt.Chart(long).mark_bar(color=BLUE, cornerRadiusTopRight=4, cornerRadiusBottomRight=4).encode(
            y=alt.Y("stage:N", title=None, sort=list(stage_cols.values()),
                    axis=alt.Axis(labelExpr="slice(datum.label, 2)", labelLimit=200)),
            x=alt.X("students:Q", title="Students"),
            tooltip=["assessment:N", alt.Tooltip("stage:N", title="Stage"), "students:Q",
                     alt.Tooltip("pct:Q", title="% of class", format=".0%")],
        ).properties(height=150, width=260)
        st.altair_chart(chart.facet(facet=alt.Facet("assessment:N", title=None), columns=2), width="content")

        st.dataframe(
            fn, width="stretch", hide_index=True,
            column_config={
                "assessment": "Assessment",
                "class_size": "Class",
                "opened_any": "Opened any item",
                "read_about": "Read About page",
                "opened_submission_point": "Opened submission point",
                "submitted": st.column_config.NumberColumn("Submitted", format="%d"),
                "submission_source": "Submissions from",
                "opened_not_submitted": st.column_config.NumberColumn("Opened, not submitted", format="%d"),
                "median_lead_days": st.column_config.NumberColumn(
                    "Median days first-open → submit", format="%.1f",
                    help="From first opening anything for this assessment to the first submit click"),
            },
        )
        if not R["has_grades"]:
            st.caption(
                "Turnitin and other external-tool submissions don't appear as clicks. "
                "Upload the Grade Centre to fill in those assessments."
            )

        st.markdown("#### Student × assessment")
        mx = R["matrix"].copy()
        acols = [c for c in mx.columns if c.startswith("A")]
        a_pick = st.selectbox("Show students who haven't submitted…", ["(everyone)"] + acols,
                              format_func=lambda c: c if c == "(everyone)" else f"Assessment {c[1:]}")
        if a_pick != "(everyone)":
            mx = mx[mx[a_pick] != "Submitted"]
        mx.insert(1, "Name", mx["student_code"].map(names))
        for c in acols:
            mx[c] = mx[c].map(lambda s: f"{STAGE_ICON.get(s, '')} {s}")
        unlogged = set(fn.loc[fn["submission_source"] == am.NOT_LOGGED, "assessment"])

        def a_header(c: str) -> str:
            name = f"Assessment {c[1:]}"
            return f"{name} (submit not logged)" if name in unlogged else name

        st.dataframe(
            mx.rename(columns={"student_code": "Student ID", **{c: a_header(c) for c in acols}}),
            width="stretch", hide_index=True, height=420,
        )

# ------------------------------------------------------------------
# Timing
# ------------------------------------------------------------------

with tabs["Timing"]:
    heat = R["heat"]
    if heat.empty:
        st.info("No clicks to chart.")
    else:
        grid = pd.MultiIndex.from_product([am.WEEKDAYS, range(24)], names=["weekday", "hour"])
        heat = heat.set_index(["weekday", "hour"]).reindex(grid, fill_value=0).reset_index()
        hm = alt.Chart(heat).mark_rect(stroke="white", strokeWidth=2, cornerRadius=2).encode(
            x=alt.X("hour:O", title="Hour of day", axis=alt.Axis(labelAngle=0)),
            y=alt.Y("weekday:N", sort=am.WEEKDAYS, title=None),
            color=alt.Color("clicks:Q", title="Clicks", scale=alt.Scale(range=SEQ_BLUE)),
            tooltip=["weekday:N", alt.Tooltip("hour:O", title="Hour"), "clicks:Q",
                     alt.Tooltip("students:Q", title="Students")],
        ).properties(height=260, title="When students click (Sydney time)")
        st.altair_chart(hm, width="stretch")
        top = heat.sort_values(["students", "clicks"], ascending=False).head(3)
        st.caption("Most students online: " + "; ".join(
            f"{r.weekday} {r.hour:02d}:00-{(r.hour + 1) % 24:02d}:00 ({r.clicks} clicks, {r.students} students)"
            for r in top.itertuples()))

        d = R["ev"].assign(day=R["ev"]["ts"].dt.normalize()).groupby("day")["student_code"].nunique()
        d = d.reindex(pd.date_range(first_ts.normalize(), last_ts.normalize()), fill_value=0)
        d = d.rename_axis("day").rename("students").reset_index()
        st.altair_chart(
            alt.Chart(d).mark_line(color=BLUE, strokeWidth=2, point=alt.OverlayMarkDef(size=30, color=BLUE)).encode(
                x=alt.X("day:T", title=None, axis=alt.Axis(format="%d %b")),
                y=alt.Y("students:Q", title="Students"),
                tooltip=[alt.Tooltip("day:T", title="Day", format="%a %d %b"), "students:Q"],
            ).properties(height=220, title="Students who visited each day"),
            width="stretch",
        )

# ------------------------------------------------------------------
# Compare subjects
# ------------------------------------------------------------------

if "Compare subjects" in tabs:
    with tabs["Compare subjects"]:
        rows = []
        for s in subjects:
            r = R if s == subject else compute(s)
            k = r["kpis"]
            rows.append({
                "Subject": s, "Students": k["enrolled"],
                f"Visited last {active_days}d": k["active_recent"] / k["enrolled"] if k["enrolled"] else 0,
                "Never visited": k["never_visited"] if r["has_roster"] else None,
                "Median visits": k["median_sessions"], "Median est. hours": k["median_hours"],
                "Bounce rate": k["bounce_rate"], "Median visit (min)": k["median_session_min"],
                "Median coverage": k["median_coverage"], "Announcement reach": k["announcement_reach"],
            })
        pct_cfg = lambda label: st.column_config.ProgressColumn(label, format="%.0f%%", min_value=0, max_value=100)
        st.dataframe(
            as_pct(pd.DataFrame(rows), [f"Visited last {active_days}d", "Bounce rate", "Median coverage", "Announcement reach"]),
            width="stretch", hide_index=True,
            column_config={
                f"Visited last {active_days}d": pct_cfg(f"Visited last {active_days}d"),
                "Bounce rate": pct_cfg("Bounce rate"),
                "Median coverage": pct_cfg("Median coverage"),
                "Announcement reach": pct_cfg("Announcement reach"),
                "Median est. hours": st.column_config.NumberColumn(format="%.1f"),
                "Median visit (min)": st.column_config.NumberColumn(format="%.1f"),
            },
        )
        st.caption("Each subject uses its own class list if one was uploaded.")

# ------------------------------------------------------------------
# Export
# ------------------------------------------------------------------

with tabs["Export"]:
    st.markdown(f"Excel workbook for **{subject}**: overview, students, content, assessments, visits and weeks.")
    overview = pd.DataFrame([
        ("Subject", R["course_name"]), ("As at", f"{as_at:%d %b %Y}"),
        ("Clicks from", f"{first_ts:%d %b %Y}"), ("Clicks to", f"{last_ts:%d %b %Y}"),
        ("Class source", "Class list" if R["has_roster"] else "Activity log only"),
        ("Students", K["enrolled"]),
        (f"Visited in last {active_days} days", K["active_recent"]),
        ("Never visited", K["never_visited"] if R["has_roster"] else "n/a (no class list)"),
        ("Not on class list", K["not_on_list"]),
        ("Median visits per student", round(K["median_sessions"], 1)),
        ("Median est. hours per student (lower bound)", round(K["median_hours"], 2)),
        ("Bounce rate", fmt_pct(K["bounce_rate"])),
        ("Median visit length (min)", round(K["median_session_min"], 1)),
        ("Median content coverage", fmt_pct(K["median_coverage"])),
        ("Announcement reach", fmt_pct(K["announcement_reach"])),
        ("Session gap (minutes)", gap_minutes),
    ] + [(f"Segment: {s}", int(v)) for s, v in students["segment"].value_counts().reindex(am.SEGMENTS, fill_value=0).items()],
        columns=["Measure", "Value"])
    out_students = students.drop(columns=["_in_log", "_on_roster"]).copy()
    lead = ["segment", "student_code", "display_name"]
    out_students = out_students[lead + [c for c in out_students.columns if c not in lead]].rename(columns={
        "segment": "Segment", "student_code": "Student ID", "display_name": "Name",
        "attend_type": "Attend type", "course": "Course", "email_address": "Email",
        "first_seen": "First visit", "last_seen": "Last visit", "clicks": "Clicks",
        "active_days": "Days active", "items_opened": "Items opened",
        "announcements_read": "Announcements read", "submissions_logged": "Submit clicks",
        "sessions": "Visits", "est_hours": "Est. hours (lower bound)", "bounce_rate": "Bounce rate",
        "avg_session_min": "Avg visit (min)", "sessions_last_7d": "Visits last 7d",
        "sessions_prev_7d": "Visits prev 7d", "coverage": "Coverage",
        "days_since_last": "Days since last visit", "assessments_submitted": "Submitted (GC)",
        "assessments_total": "Assessments (GC)", "submission_rate": "Submission rate (GC)",
        "avg_score_pct": "Avg mark % (GC)",
    })
    out_matrix = R["matrix"].copy()
    out_matrix.insert(1, "Name", out_matrix["student_code"].map(names))
    out_sessions = R["sessions"].drop(columns=["session_id", "actions"]).copy()
    out_sessions.insert(1, "Name", out_sessions["student_code"].map(names))
    sheets = {
        "Overview": overview,
        "Students": out_students,
        "Content": R["content"],
        "Assessment funnel": R["funnel"],
        "Assessment progress": out_matrix,
        "Visits": out_sessions,
        "Weekly": R["weekly"].drop(columns=["week"]) if not R["weekly"].empty else R["weekly"],
    }
    st.download_button(
        "⬇ Download Excel report",
        data=activity_report.build_workbook(sheets),
        file_name=f"engagement_tracker_{subject}_{as_at.isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption("Contains student names and IDs - handle as you would the class list.")

# ------------------------------------------------------------------
# How it's measured
# ------------------------------------------------------------------

with tabs["How it's measured"]:
    st.markdown(f"""
The activity export records **when** each click happened and what was opened - never how long
anyone stayed. Everything below is built from those timestamps.

| Measure | How it's worked out |
|---|---|
| **Visit** (session) | A run of clicks with no gap longer than **{gap_minutes} minutes**. A longer gap starts a new visit. |
| **Bounce** | A visit where the student entered the subject and opened nothing else. |
| **Visit length / est. time** | First click to last click of each visit, added up. A **lower bound**: time after the last click of a visit isn't recorded, so one-click visits count as 0. |
| **Time on an item** | Gap from opening it to the student's next click in the same visit. Not estimated for **External tool** and **SCORM** items - they open outside vUWS, so the gap is just the launch click. |
| **Reach** | Share of the class who opened the item at least once. Stands in for click-through rate, since the export has no "was shown it" count. |
| **Views** | Visits in which the item was opened. Repeat clicks on the same item within one visit count once. |
| **Exit rate** | Share of an item's views that were the last click of the visit. |
| **Coverage** | Share of the subject's items (those anyone opened) a student has opened. |
| **Announcement reach** | Share of the class who opened at least one announcement on vUWS. Email reads aren't counted. |
| **Segments** | 🟢 Engaged - visited in the last {active_days} days · 🟡 Fading - last visit {active_days + 1}-{lapse_days} days ago · 🟠 Lapsed - no visit for more than {lapse_days} days · 🔴 Never visited - on the class list, no clicks · ⚪ Not on class list - clicks but not enrolled (usually withdrawn). Counted back from the **As at** date. |
| **Assessment stages** | Items are linked to an assessment by "Assessment <number>" in the title. *About page* = "About Assessment N". *Submission point* = the vUWS submission link or the external (Turnitin) link. *Submitted* = Grade Centre entry if uploaded (any mark, or Satisfactory/Complete text), otherwise a submit click in the log. Resubmission links and Q&A forums aren't stages. |
| **Days first-open → submit** | From first opening anything for that assessment to the first submit click. Only where submit clicks are logged. |

**Cleaning:** staff preview accounts (`_PreviewUser`) are dropped. Username and IP address are ignored.
""")
