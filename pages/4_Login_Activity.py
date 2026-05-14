"""
Course at a Glance — Login Activity Dashboard
===============================================

Replicates the Subject Login Report using data from one or more
Overall Usage Reports. Useful when the CAG login report in the LMS
is unavailable.

Upload a class list (source of truth) and one or more Overall Usage
Reports to see login metrics, breakdowns, daily activity, and
multi-report comparisons.
"""

from __future__ import annotations

import io
import tempfile
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

from student_tracker.parsers import class_list as p_class
from student_tracker.parsers import overall_report as p_overall

# ------------------------------------------------------------------
# Config
# ------------------------------------------------------------------

FREQ_BANDS = [
    ("Never accessed", 0, 0),
    ("1 login day", 1, 1),
    ("2–5 login days", 2, 5),
    ("6–10 login days", 6, 10),
    ("11–20 login days", 11, 20),
    ("21+ login days", 21, 9999),
]

DAYS_SINCE_BANDS = [
    ("Active today", 0, 0),
    ("1–3 days ago", 1, 3),
    ("4–7 days ago", 4, 7),
    ("8–14 days ago", 8, 14),
    ("15–21 days ago", 15, 21),
    ("22+ days ago", 22, 9999),
    ("Never accessed", None, None),
]


# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------

def _save_upload(uploaded_file) -> str:
    """Save a Streamlit UploadedFile to a temp file, return its path."""
    suffix = Path(uploaded_file.name).suffix
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.getbuffer())
    tmp.close()
    return tmp.name


def _parse_classlist(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """Parse classlist, auto-detecting .xls (CFB) vs .xlsx (enriched)."""
    suffix = Path(filename).suffix.lower()

    if suffix == ".xlsx" or file_bytes[:2] == b"PK":
        # Enriched .xlsx format
        df = pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl", dtype=str)
        df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]
        col_map = {
            "student_id": "student_code",
            "studentid": "student_code",
            "surname": "last_name",
            "first_name": "first_name",
            "firstname": "first_name",
            "course_code": "course_code",
            "subject_code": "subject_code",
            "class_code": "class_code",
            "email": "email_address",
            "teacher": "teacher",
            "discipline_subject": "discipline_subject",
            "discipline_class": "discipline_class",
            "discipline_teacher": "discipline_teacher",
            "gedu_subject": "subject_code",
        }
        df = df.rename(columns={k: v for k, v in col_map.items() if k in df.columns})
    else:
        # Standard .xls CFB via the existing parser
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xls")
        tmp.write(file_bytes)
        tmp.close()
        df = p_class.parse(tmp.name)

    if "student_code" not in df.columns:
        raise ValueError("Class list must contain a student_code / Student ID column.")

    df["student_code"] = df["student_code"].astype(str).str.strip()
    df = df.dropna(subset=["student_code"])
    df = df[df["student_code"] != ""]
    return df


def _build_summary(
    date_hits: pd.DataFrame,
    classlist: pd.DataFrame,
    report_end_date: date,
) -> pd.DataFrame:
    """
    Build a per-student summary using the classlist as source of truth.

    Students on the classlist but absent from the Overall report appear
    with zero activity. Students in the Overall report but NOT on the
    classlist are excluded.
    """
    cl = classlist.copy()
    cl["student_code"] = cl["student_code"].astype(str).str.strip()
    valid_codes = set(cl["student_code"])

    # Filter hits to classlist only
    if not date_hits.empty:
        real = date_hits[date_hits["student_code"].isin(valid_codes)].copy()
    else:
        real = pd.DataFrame(columns=["student_code", "date", "hits"])

    # Aggregate per student
    if not real.empty:
        agg = (
            real.groupby("student_code")
            .agg(
                total_hits=("hits", "sum"),
                active_days=("date", "nunique"),
                first_active=("date", "min"),
                last_active=("date", "max"),
            )
            .reset_index()
        )
    else:
        agg = pd.DataFrame(columns=[
            "student_code", "total_hits", "active_days",
            "first_active", "last_active",
        ])

    # Start from classlist, left-join activity
    summary = cl.merge(agg, on="student_code", how="left")
    summary["total_hits"] = summary["total_hits"].fillna(0).astype(int)
    summary["active_days"] = summary["active_days"].fillna(0).astype(int)
    summary["logged_in"] = summary["total_hits"] > 0

    # Build display name
    if "last_name" in summary.columns and "first_name" in summary.columns:
        summary["student_name"] = (
            summary["last_name"].fillna("") + ", " + summary["first_name"].fillna("")
        )
    else:
        summary["student_name"] = summary["student_code"]

    # Days since last activity
    def _days_since(row):
        if pd.isna(row.get("last_active")):
            return None
        last = row["last_active"]
        if isinstance(last, pd.Timestamp):
            last = last.date()
        return (report_end_date - last).days

    summary["days_since_last"] = summary.apply(_days_since, axis=1)

    sort_col = "last_name" if "last_name" in summary.columns else "student_name"
    return summary.sort_values(sort_col).reset_index(drop=True)


def _freq_band(active_days: int) -> str:
    for label, lo, hi in FREQ_BANDS:
        if lo <= active_days <= hi:
            return label
    return "21+ login days"


def _days_since_band(days_since) -> str:
    if days_since is None or pd.isna(days_since):
        return "Never accessed"
    days_since = int(days_since)
    for label, lo, hi in DAYS_SINCE_BANDS:
        if lo is None:
            continue
        if lo <= days_since <= hi:
            return label
    return "22+ days ago"


# ------------------------------------------------------------------
# Sidebar
# ------------------------------------------------------------------

st.sidebar.title("Login Activity")

st.sidebar.markdown("**1. Class list** (source of truth)")
class_file = st.sidebar.file_uploader(
    "Class list",
    type=["xls", "xlsx"],
    help="Standard .xls or enriched .xlsx class list.",
    key="la_classlist",
)

st.sidebar.markdown("**2. Overall Usage Report(s)**")
uploaded_files = st.sidebar.file_uploader(
    "Overall Usage Report(s)",
    type=["xls"],
    accept_multiple_files=True,
    help="SpreadsheetML XML format from vUWS. Upload multiple for comparison.",
    key="la_overall",
)

st.sidebar.divider()

report_end_date = st.sidebar.date_input(
    "Report end date",
    value=date.today(),
    help="Used to calculate 'days since last login'. Set to the date the report was pulled.",
    key="la_end_date",
)

# Date range filter
st.sidebar.divider()
st.sidebar.markdown("**Date range filter** (optional)")
use_date_filter = st.sidebar.checkbox("Filter to specific date range", value=False, key="la_date_filter")
if use_date_filter:
    filter_start = st.sidebar.date_input("From", value=date.today() - timedelta(days=7), key="la_from")
    filter_end = st.sidebar.date_input("To", value=date.today(), key="la_to")
else:
    filter_start = None
    filter_end = None


# ------------------------------------------------------------------
# Main
# ------------------------------------------------------------------

st.title("📊 Course at a Glance — Login Activity")

if not class_file or not uploaded_files:
    st.info(
        "Upload a **class list** and one or more **Overall Usage Reports** "
        "(.xls) in the sidebar to get started.\n\n"
        "The class list is used as the source of truth for enrolled students. "
        "Only students on the class list appear in the dashboard — "
        "preview users, staff, and other non-student accounts are "
        "excluded automatically.\n\n"
        "This report replicates the Subject Login Report metrics using "
        "the Overall Summary of Usage data."
    )
    st.stop()


# ── Parse classlist ───────────────────────────────────────────────

with st.spinner("Parsing class list…"):
    try:
        classlist = _parse_classlist(class_file.getvalue(), class_file.name)
        st.sidebar.success(f"Class list: {len(classlist)} students")
    except Exception as e:
        st.error(f"Failed to parse class list: {e}")
        st.stop()

# Detect extra columns for cohort filters
has_course = "course_code" in classlist.columns
teacher_col = (
    "discipline_teacher" if "discipline_teacher" in classlist.columns
    else "teacher" if "teacher" in classlist.columns
    else None
)
has_teacher = teacher_col is not None

if has_course or has_teacher:
    st.sidebar.divider()
    st.sidebar.markdown("**Cohort filters**")

    course_filter = []
    teacher_filter = []

    if has_course:
        courses = sorted(classlist["course_code"].dropna().unique())
        course_filter = st.sidebar.multiselect("Course / Program", courses, key="la_course")

    if has_teacher:
        teachers = sorted(classlist[teacher_col].dropna().unique())
        teacher_filter = st.sidebar.multiselect("Teacher", teachers, key="la_teacher")

    if course_filter:
        classlist = classlist[classlist["course_code"].isin(course_filter)]
    if teacher_filter:
        classlist = classlist[classlist[teacher_col].isin(teacher_filter)]

    if classlist.empty:
        st.warning("No students match the selected filters.")
        st.stop()


# ── Parse Overall reports ─────────────────────────────────────────

reports: list[dict] = []

with st.spinner("Parsing Overall Usage Reports…"):
    for uf in uploaded_files:
        try:
            path = _save_upload(uf)
            date_hits = p_overall.parse_date_section(path)

            # Apply date filter if set
            if use_date_filter and filter_start and filter_end:
                if not date_hits.empty:
                    date_hits = date_hits[
                        (date_hits["date"] >= pd.Timestamp(filter_start)) &
                        (date_hits["date"] <= pd.Timestamp(filter_end))
                    ]

            summary = _build_summary(date_hits, classlist, report_end_date)

            # Count classlist students found in Overall
            hits_ids = set(date_hits["student_code"].unique()) if not date_hits.empty else set()
            class_ids = set(classlist["student_code"])
            matched = len(class_ids & hits_ids)

            reports.append({
                "name": uf.name,
                "date_hits": date_hits,
                "summary": summary,
                "matched": matched,
                "total_in_file": len(hits_ids),
            })
        except Exception as e:
            st.error(f"Failed to parse **{uf.name}**: {e}")

if not reports:
    st.warning("No reports parsed successfully.")
    st.stop()


# ------------------------------------------------------------------
# Single-report rendering
# ------------------------------------------------------------------

def render_single_report(rpt: dict, show_title: bool = True):
    date_hits = rpt["date_hits"]
    summary = rpt["summary"]
    name = rpt["name"]

    # Date range label
    class_ids = set(classlist["student_code"])
    real = date_hits[date_hits["student_code"].isin(class_ids)] if not date_hits.empty else date_hits
    if not real.empty:
        d_min = real["date"].min()
        d_max = real["date"].max()
        if isinstance(d_min, pd.Timestamp):
            d_min = d_min.date()
        if isinstance(d_max, pd.Timestamp):
            d_max = d_max.date()
        date_label = f"{d_min.strftime('%d %b')} – {d_max.strftime('%d %b %Y')}"
    else:
        date_label = "No activity data"

    if show_title:
        subject_code = p_class.detect_subject_code(classlist) if hasattr(p_class, 'detect_subject_code') else None
        title_prefix = f"{subject_code} — " if subject_code else ""
        st.subheader(f"{title_prefix}{name}")
        st.caption(
            f"Data period: {date_label}  ·  "
            f"Enrolled: {len(summary)}  ·  "
            f"Matched in report: {rpt['matched']} of {len(classlist)} classlist students"
        )

    total = len(summary)
    logged_in = int(summary["logged_in"].sum())
    not_logged_in = total - logged_in
    login_rate = (logged_in / total * 100) if total > 0 else 0
    active_only = summary[summary["logged_in"]]
    avg_hits = active_only["total_hits"].mean() if not active_only.empty else 0
    avg_days = active_only["active_days"].mean() if not active_only.empty else 0

    # ── Headline metrics ──────────────────────────────────────────
    cols = st.columns(6)
    cols[0].metric("Enrolled", total)
    cols[1].metric("Logged In", logged_in)
    cols[2].metric("Not Logged In", not_logged_in)
    cols[3].metric("Login Rate", f"{login_rate:.1f}%")
    cols[4].metric("Avg Hits (active)", f"{avg_hits:.1f}")
    cols[5].metric("Avg Active Days", f"{avg_days:.1f}")

    st.divider()

    # ── Tabs ──────────────────────────────────────────────────────
    tab_breakdown, tab_days_since, tab_daily, tab_students, tab_export = st.tabs([
        "Login Breakdown",
        "Days Since Last Login",
        "Daily Activity",
        "Student Detail",
        "Export",
    ])

    # ── Tab 1: Login frequency breakdown ──────────────────────────
    with tab_breakdown:
        st.subheader("Login Frequency Distribution")
        summary["freq_band"] = summary["active_days"].apply(_freq_band)

        band_order = [b[0] for b in FREQ_BANDS]
        band_counts = summary["freq_band"].value_counts().reindex(band_order, fill_value=0)
        band_df = pd.DataFrame({
            "Band": band_counts.index,
            "Count": band_counts.values,
            "Percentage": (band_counts.values / total * 100).round(1),
        })

        col_chart, col_table = st.columns([2, 1])
        with col_chart:
            colours = ["#dc3545", "#fd7e14", "#ffc107", "#20c997", "#0d6efd", "#6610f2"]
            fig = px.bar(
                band_df, x="Band", y="Count", color="Band",
                color_discrete_sequence=colours[:len(band_df)], text="Count",
            )
            fig.update_layout(showlegend=False, xaxis_title="", yaxis_title="Students", height=400)
            fig.update_traces(textposition="outside")
            st.plotly_chart(fig, use_container_width=True)
        with col_table:
            st.dataframe(band_df.set_index("Band"), use_container_width=True, height=300)

        # Hits histogram
        st.subheader("Total Hits Distribution (active students)")
        if not active_only.empty:
            fig_hist = px.histogram(
                active_only, x="total_hits", nbins=30,
                labels={"total_hits": "Total Hits"},
                color_discrete_sequence=["#0d6efd"],
            )
            fig_hist.update_layout(xaxis_title="Total Hits", yaxis_title="Students", height=350)
            st.plotly_chart(fig_hist, use_container_width=True)

    # ── Tab 2: Days since last login ──────────────────────────────
    with tab_days_since:
        st.subheader("Days Since Last Login")
        summary["days_since_band"] = summary["days_since_last"].apply(_days_since_band)

        band_order_ds = [b[0] for b in DAYS_SINCE_BANDS]
        ds_counts = summary["days_since_band"].value_counts().reindex(band_order_ds, fill_value=0)
        ds_df = pd.DataFrame({
            "Band": ds_counts.index,
            "Count": ds_counts.values,
            "Percentage": (ds_counts.values / total * 100).round(1),
        })

        col_c2, col_t2 = st.columns([2, 1])
        with col_c2:
            colours_ds = ["#198754", "#20c997", "#0dcaf0", "#ffc107", "#fd7e14", "#dc3545", "#6c757d"]
            fig_ds = px.bar(
                ds_df, x="Band", y="Count", color="Band",
                color_discrete_sequence=colours_ds[:len(ds_df)], text="Count",
            )
            fig_ds.update_layout(showlegend=False, xaxis_title="", yaxis_title="Students", height=400)
            fig_ds.update_traces(textposition="outside")
            st.plotly_chart(fig_ds, use_container_width=True)
        with col_t2:
            st.dataframe(ds_df.set_index("Band"), use_container_width=True, height=350)

        # At-risk students
        st.subheader("Students — Not Logged In (8+ days or never)")
        at_risk = summary[
            (summary["days_since_last"].isna()) | (summary["days_since_last"] >= 8)
        ].sort_values("days_since_last", ascending=False, na_position="first")

        if not at_risk.empty:
            ar_cols = ["student_name", "student_code", "total_hits", "active_days", "last_active", "days_since_last"]
            ar_cols = [c for c in ar_cols if c in at_risk.columns]
            st.dataframe(
                at_risk[ar_cols].rename(columns={
                    "student_name": "Name", "student_code": "Student ID",
                    "total_hits": "Total Hits", "active_days": "Active Days",
                    "last_active": "Last Active", "days_since_last": "Days Since",
                }),
                use_container_width=True, hide_index=True,
            )
            st.caption(f"{len(at_risk)} students with 8+ days since last activity or never accessed.")
        else:
            st.success("All students have been active within the last 7 days.")

    # ── Tab 3: Daily activity chart ───────────────────────────────
    with tab_daily:
        st.subheader("Daily Activity")
        real_hits = real.copy() if not real.empty else pd.DataFrame()

        if not real_hits.empty:
            daily = (
                real_hits.groupby("date")
                .agg(total_hits=("hits", "sum"), unique_students=("student_code", "nunique"))
                .reset_index().sort_values("date")
            )
            daily["date"] = pd.to_datetime(daily["date"])

            fig_daily = go.Figure()
            fig_daily.add_trace(go.Bar(
                x=daily["date"], y=daily["total_hits"],
                name="Total Hits", marker_color="#0d6efd", opacity=0.7,
            ))
            fig_daily.add_trace(go.Scatter(
                x=daily["date"], y=daily["unique_students"],
                name="Unique Students", yaxis="y2",
                mode="lines+markers", marker_color="#dc3545", line=dict(width=2),
            ))
            fig_daily.update_layout(
                yaxis=dict(title="Total Hits", side="left"),
                yaxis2=dict(title="Unique Students", side="right", overlaying="y"),
                legend=dict(orientation="h", yanchor="bottom", y=1.02),
                height=450, hovermode="x unified",
            )
            st.plotly_chart(fig_daily, use_container_width=True)

            # Day of week
            st.subheader("Activity by Day of Week")
            rh = real_hits.copy()
            rh["date"] = pd.to_datetime(rh["date"])
            rh["dow"] = rh["date"].dt.day_name()
            dow_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
            dow_agg = rh.groupby("dow")["hits"].sum().reindex(dow_order, fill_value=0)
            fig_dow = px.bar(
                x=dow_agg.index, y=dow_agg.values,
                labels={"x": "", "y": "Total Hits"},
                color_discrete_sequence=["#6610f2"],
            )
            fig_dow.update_layout(height=350)
            st.plotly_chart(fig_dow, use_container_width=True)
        else:
            st.info("No daily activity data found for classlist students.")

    # ── Tab 4: Student detail ─────────────────────────────────────
    with tab_students:
        st.subheader("All Students")

        col_f1, col_f2 = st.columns(2)
        with col_f1:
            status_filter = st.selectbox(
                "Status", ["All", "Logged in", "Never accessed"], key=f"sf_{name}",
            )
        with col_f2:
            search = st.text_input("Search by name or ID", key=f"sr_{name}")

        display = summary.copy()
        if status_filter == "Logged in":
            display = display[display["logged_in"]]
        elif status_filter == "Never accessed":
            display = display[~display["logged_in"]]

        if search:
            mask = (
                display["student_name"].str.contains(search, case=False, na=False) |
                display["student_code"].str.contains(search, case=False, na=False)
            )
            display = display[mask]

        show_cols = [
            "student_name", "student_code", "total_hits", "active_days",
            "first_active", "last_active", "days_since_last", "logged_in",
        ]
        # Add classlist metadata if present
        for extra in ["course_code", "discipline_class", "discipline_teacher", "teacher", "attend_type"]:
            if extra in display.columns:
                show_cols.insert(2, extra)
        show_cols = [c for c in show_cols if c in display.columns]

        st.dataframe(
            display[show_cols].rename(columns={
                "student_name": "Name", "student_code": "Student ID",
                "total_hits": "Total Hits", "active_days": "Active Days",
                "first_active": "First Active", "last_active": "Last Active",
                "days_since_last": "Days Since", "logged_in": "Logged In",
                "course_code": "Course", "discipline_class": "Class",
                "discipline_teacher": "Teacher", "attend_type": "Attend Type",
            }),
            use_container_width=True, hide_index=True, height=600,
        )
        st.caption(f"Showing {len(display)} of {total} students.")

    # ── Tab 5: Export ─────────────────────────────────────────────
    with tab_export:
        st.subheader("Export to Excel")

        export = summary.copy()
        export["freq_band"] = export["active_days"].apply(_freq_band)
        export["days_since_band"] = export["days_since_last"].apply(_days_since_band)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            meta = pd.DataFrame({
                "Metric": [
                    "Report", "Date Range", "Enrolled", "Logged In",
                    "Not Logged In", "Login Rate",
                    "Avg Total Hits (active)", "Avg Active Days (active)",
                ],
                "Value": [
                    name, date_label, total, logged_in,
                    not_logged_in, f"{login_rate:.1f}%",
                    f"{avg_hits:.1f}", f"{avg_days:.1f}",
                ],
            })
            meta.to_excel(writer, sheet_name="Summary", index=False)

            export_cols = [
                "student_name", "student_code", "total_hits", "active_days",
                "first_active", "last_active", "days_since_last",
                "logged_in", "freq_band", "days_since_band",
            ]
            for extra in ["course_code", "discipline_class", "discipline_teacher", "teacher"]:
                if extra in export.columns:
                    export_cols.insert(2, extra)
            export_cols = [c for c in export_cols if c in export.columns]
            export[export_cols].to_excel(writer, sheet_name="Per Student", index=False)

            band_df.to_excel(writer, sheet_name="Login Freq Breakdown", index=False)
            ds_df.to_excel(writer, sheet_name="Days Since Breakdown", index=False)

        st.download_button(
            "📥 Download Excel Report",
            data=buffer.getvalue(),
            file_name=f"CAG_Login_{name.replace('.xls', '')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )


# ------------------------------------------------------------------
# Multi-report comparison
# ------------------------------------------------------------------

def render_comparison(reports_list: list[dict]):
    st.subheader("Multi-Report Comparison")

    # Comparison table
    comp_rows = []
    for rpt in reports_list:
        s = rpt["summary"]
        total = len(s)
        logged_in = int(s["logged_in"].sum())
        active = s[s["logged_in"]]
        comp_rows.append({
            "Report": rpt["name"],
            "Enrolled": total,
            "Logged In": logged_in,
            "Not Logged In": total - logged_in,
            "Login Rate": f"{logged_in / total * 100:.1f}%" if total else "—",
            "Avg Hits": f"{active['total_hits'].mean():.1f}" if not active.empty else "—",
            "Avg Active Days": f"{active['active_days'].mean():.1f}" if not active.empty else "—",
        })
    st.dataframe(pd.DataFrame(comp_rows), use_container_width=True, hide_index=True)

    # Login rate chart
    st.subheader("Login Rate Comparison")
    rate_data = []
    for rpt in reports_list:
        s = rpt["summary"]
        total = len(s)
        if total > 0:
            rate_data.append({
                "Report": rpt["name"],
                "Login Rate (%)": s["logged_in"].sum() / total * 100,
            })
    if rate_data:
        rate_df = pd.DataFrame(rate_data)
        fig = px.bar(
            rate_df, x="Report", y="Login Rate (%)",
            text="Login Rate (%)", color_discrete_sequence=["#0d6efd"],
        )
        fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
        fig.update_layout(height=400, yaxis_range=[0, 100])
        st.plotly_chart(fig, use_container_width=True)

    # Stacked frequency bands
    st.subheader("Login Frequency — Stacked")
    band_data = []
    for rpt in reports_list:
        s = rpt["summary"]
        s["freq_band"] = s["active_days"].apply(_freq_band)
        for label, _, _ in FREQ_BANDS:
            band_data.append({
                "Report": rpt["name"],
                "Band": label,
                "Count": int((s["freq_band"] == label).sum()),
            })
    if band_data:
        fig_stack = px.bar(
            pd.DataFrame(band_data), x="Report", y="Count", color="Band",
            barmode="stack",
            color_discrete_sequence=["#dc3545", "#fd7e14", "#ffc107", "#20c997", "#0d6efd", "#6610f2"],
        )
        fig_stack.update_layout(height=450)
        st.plotly_chart(fig_stack, use_container_width=True)

    # Student movement (first vs last report)
    if len(reports_list) >= 2:
        st.subheader("Student Movement (First → Last Report)")

        first_s = reports_list[0]["summary"][["student_code", "student_name", "logged_in", "total_hits"]].rename(
            columns={"logged_in": "logged_in_first", "total_hits": "hits_first"}
        )
        last_s = reports_list[-1]["summary"][["student_code", "logged_in", "total_hits"]].rename(
            columns={"logged_in": "logged_in_last", "total_hits": "hits_last"}
        )
        merged = first_s.merge(last_s, on="student_code", how="outer")

        new_active = merged[
            (merged["logged_in_first"].isna() | ~merged["logged_in_first"]) &
            (merged["logged_in_last"] == True)  # noqa: E712
        ]
        dropped = merged[
            (merged["logged_in_first"] == True) &  # noqa: E712
            (merged["logged_in_last"].isna() | ~merged["logged_in_last"])
        ]

        col_m1, col_m2 = st.columns(2)
        with col_m1:
            st.markdown(f"**Newly active:** {len(new_active)}")
            if not new_active.empty:
                st.dataframe(
                    new_active[["student_name", "student_code", "hits_last"]].rename(
                        columns={"student_name": "Name", "student_code": "ID", "hits_last": "Hits (latest)"}
                    ),
                    hide_index=True, height=300,
                )
        with col_m2:
            st.markdown(f"**Dropped off:** {len(dropped)}")
            if not dropped.empty:
                st.dataframe(
                    dropped[["student_name", "student_code", "hits_first"]].rename(
                        columns={"student_name": "Name", "student_code": "ID", "hits_first": "Hits (first)"}
                    ),
                    hide_index=True, height=300,
                )


# ------------------------------------------------------------------
# Render
# ------------------------------------------------------------------

if len(reports) == 1:
    render_single_report(reports[0])
else:
    view_mode = st.sidebar.radio("View", ["Comparison", "Individual Reports"], key="la_view")
    if view_mode == "Comparison":
        render_comparison(reports)
    else:
        selected = st.sidebar.selectbox(
            "Select report", [r["name"] for r in reports], key="la_select",
        )
        rpt = next(r for r in reports if r["name"] == selected)
        render_single_report(rpt)
