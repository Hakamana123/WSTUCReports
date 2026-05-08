"""Streamlit dashboard for student engagement tracking.

Run with:
    streamlit run app.py

Upload four files for ONE subject for one cumulative reporting point:
  1. Class list (.xls — real CFB format)
  2. Subject Login Report (.xlsx)
  3. Overall Summary of Usage (.xls — actually SpreadsheetML XML)
  4. Full Grade Centre (.xls — actually UTF-16 TSV)

Pick the block start date and the recency half-life, then explore the
dashboard or download the Excel report.
"""

from datetime import date, timedelta
from pathlib import Path
import tempfile
import yaml

import pandas as pd
import streamlit as st

from student_tracker.parsers import class_list as p_class
from student_tracker.parsers import login_report as p_login
from student_tracker.parsers import overall_report as p_overall
from student_tracker.parsers import grade_centre as p_grade
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
    CONFIG = {
        "recency_half_life_days": 7,
        "fade_threshold": 0.5,
    }


def _save_upload(uploaded_file) -> str:
    """Save a Streamlit UploadedFile to a temp file, return its path."""
    suffix = Path(uploaded_file.name).suffix
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.getbuffer())
    tmp.close()
    return tmp.name


# ------------------------------------------------------------------
# Sidebar — inputs
# ------------------------------------------------------------------

st.sidebar.title("Inputs")
st.sidebar.markdown("Upload one bundle per subject per reporting point.")

class_file = st.sidebar.file_uploader("1. Class list (.xls)", type=["xls"])
login_file = st.sidebar.file_uploader("2. Subject Login Report (.xlsx)", type=["xlsx"])
overall_files = st.sidebar.file_uploader(
    "3. Overall Summary of Usage (.xls) — can upload multiple",
    type=["xls"],
    accept_multiple_files=True,
)
grade_file = st.sidebar.file_uploader("4. Full Grade Centre (.xls) — optional", type=["xls"])

st.sidebar.divider()

block_start = st.sidebar.date_input(
    "Block start date",
    value=date.today() - timedelta(days=21),
    help="The first day of the teaching block. Used to classify S2 (pre-block ghost) and S3 (W1 ghost).",
)

half_life = st.sidebar.number_input(
    "Recency half-life (days)",
    min_value=1.0, max_value=60.0,
    value=float(CONFIG.get("recency_half_life_days", 7)),
    step=1.0,
    help="Days at which a hit's weight is halved in the recency-weighted metrics.",
)

fade_threshold = st.sidebar.number_input(
    "Fade threshold (S6 cutoff)",
    min_value=0.1, max_value=0.9,
    value=float(CONFIG.get("fade_threshold", 0.5)),
    step=0.05,
    help="A student is 'Fading' (S6) if this_week_hits < threshold * last_week_hits.",
)

# ------------------------------------------------------------------
# Main panel
# ------------------------------------------------------------------

st.title("Student Engagement Tracker")

if not all([class_file, login_file, overall_files]):
    st.info("Upload the first three files in the sidebar to begin (Grade Centre is optional). You can attach multiple Overall Summary files.")
    with st.expander("How segments are defined"):
        st.markdown("""
- **S1 — Never engaged:** 0 lifetime logins (also catches students missing from the login report).
- **S2 — Pre-block ghost:** Logged in before block start, no login since.
- **S3 — W1 ghost:** Hits in block week 1, none in any later week in the data.
- **S4 — Dropped this week:** Hits last week, none this week.
- **S5 — Returning engager:** No hits last week, hits this week.
- **S6 — Fading engager:** Hits both weeks, this week < fade-threshold of last week.
- **S7 — True sustainer:** Hits both weeks, this week is at or above the fade-threshold ratio.
- **S8 — Long-tail dropout:** Had hits at some point in the data window, but zero in either of the last two weeks. Mainly relevant for long blocks (e.g. 17-week prep subjects); usually empty for 4-week discipline blocks.
        """)
    st.stop()

# Parse files
with st.spinner("Parsing files…"):
    try:
        cls_df_raw = p_class.parse(_save_upload(class_file))
        filter_cfg = CONFIG.get("class_list_filter", {})
        cls_df, filter_stats = p_class.filter_for_real_students(
            cls_df_raw,
            exclude_id_prefixes=filter_cfg.get("exclude_id_prefixes", []),
            exclude_surnames=filter_cfg.get("exclude_surnames", []),
        )
        login_df = p_login.parse(_save_upload(login_file))

        # Combine multiple Overall reports (different reporting windows)
        date_hits_parts = []
        for f in overall_files:
            part = p_overall.parse_date_section(_save_upload(f))
            if not part.empty:
                date_hits_parts.append(part)
        if date_hits_parts:
            date_hits = pd.concat(date_hits_parts, ignore_index=True)
            # Defensive dedup: if windows overlap, sum hits for the same (student, date)
            date_hits = date_hits.groupby(
                ["student_code", "date"], as_index=False
            )["hits"].sum()
        else:
            date_hits = pd.DataFrame(columns=["student_code", "date", "hits"])

        if grade_file:
            grade_parsed = p_grade.parse(_save_upload(grade_file))
            grade_summary = p_grade.submission_summary(grade_parsed)
        else:
            grade_parsed = {"students": pd.DataFrame(columns=["student_code"]), "assessments": []}
            grade_summary = None
    except Exception as e:
        st.error(f"Failed to parse one of the files: {e}")
        st.exception(e)
        st.stop()

# Validation
subject_code = p_class.detect_subject_code(cls_df)
class_ids = set(cls_df["student_code"])
login_ids = set(login_df["student_code"])
hits_ids = set(date_hits["student_code"]) if not date_hits.empty else set()

with st.expander("File validation", expanded=False):
    num_val_cols = 4 if grade_file else 3
    val_cols = st.columns(num_val_cols)
    excluded = filter_stats["by_prefix"] + filter_stats["by_surname"]
    val_cols[0].metric(
        "Class list (after filter)",
        f"{len(class_ids)} students",
        delta=(f"-{excluded} excluded" if excluded > 0 else None),
        delta_color="off",
    )
    val_cols[1].metric("Login report", f"{len(login_ids)} students")
    val_cols[2].metric(f"Overall report ({len(overall_files)} file{'s' if len(overall_files) != 1 else ''})",
              f"{len(hits_ids)} with hits")
    if grade_file:
        grade_ids = set(grade_parsed["students"]["student_code"])
        val_cols[3].metric("Grade centre", f"{len(grade_ids)} students")
    st.write(f"**Subject code (from class list):** {subject_code or 'unknown'}")
    if excluded > 0:
        st.write(
            f"**Class list filter:** started with {filter_stats['started']}, "
            f"removed {filter_stats['by_prefix']} by ID prefix and "
            f"{filter_stats['by_surname']} by surname (configured in `config.yaml`)."
        )
    if not date_hits.empty:
        st.write(f"**Date range across Overall files:** {date_hits['date'].min().date()} to {date_hits['date'].max().date()}")
    in_class_with_hits = len(class_ids & hits_ids)
    st.write(f"Class-list students appearing in Overall reports: **{in_class_with_hits}** of {len(class_ids)}")
    extras = hits_ids - class_ids
    if extras:
        st.write(f"Overall reports contain {len(extras)} IDs not in class list (filtered out — preview, staff, or withdrawn).")

# Build master summary (filtered to class list)
reference_date = date_hits["date"].max().date() if not date_hits.empty else date.today()

summary = metrics.per_student_summary(
    class_list=cls_df,
    date_hits=date_hits,
    login=login_df,
    grade_summary=grade_summary,
    reference_date=reference_date,
    half_life_days=half_life,
)

weeks = metrics.weeks_in_data(date_hits)
weekly_hits = metrics.weekly_hits_table(date_hits)
weekly_active = metrics.weekly_active_days_table(date_hits)
summary = metrics.append_weekly_columns(summary, weekly_hits, weekly_active, weeks, block_start)
summary = metrics.append_recent_week_averages(summary, weekly_hits, weeks, reference_date)

classified = segmentation.classify(
    summary=summary,
    weekly_hits=weekly_hits,
    block_start_date=block_start,
    weeks=weeks,
    fade_threshold=fade_threshold,
)

# ------------------------------------------------------------------
# Display
# ------------------------------------------------------------------

st.subheader(f"{subject_code or 'Subject'} — {len(classified)} enrolled students")
week_labels = [metrics.block_week_label(y, w, block_start) for (y, w) in weeks]
st.caption(f"Reference date: {reference_date} · Block start: {block_start} · Weeks in data: {', '.join(week_labels)}")

# Segment counts
counts = segmentation.segment_counts(classified)
counts_display = counts[counts["count"] > 0]

st.markdown("### Segment distribution")
cols = st.columns(min(len(counts_display), 7))
for i, (_, row) in enumerate(counts_display.iterrows()):
    with cols[i % len(cols)]:
        st.metric(row["segment"], row["count"])
        st.caption(f"{row['pct']}% of cohort")

# Filters
st.markdown("### Student detail")
fc1, fc2, fc3 = st.columns(3)
seg_filter = fc1.multiselect(
    "Filter by segment",
    options=segmentation.ALL_SEGMENTS,
    default=[],
)
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
        filtered["student_code"].str.lower().str.contains(s, na=False)
        | filtered["first_name"].str.lower().str.contains(s, na=False)
        | filtered["last_name"].str.lower().str.contains(s, na=False)
    )
    filtered = filtered[mask]

# Display table
display_cols = [
    "segment", "student_code", "preferred_name", "last_name", "attend_type",
    "last_hit_date", "days_since_last_hit",
    "prior_week_daily_avg", "this_week_daily_avg",
    "total_logins", "last_login_date", "days_since_last_login",
    "total_hits", "weighted_hits", "total_active_days", "weighted_active_days",
    "submission_rate", "avg_score_pct",
]
display_cols += [c for c in filtered.columns if c.startswith("W") and ("_hits" in c or "_active_days" in c)]
display_cols = [c for c in display_cols if c in filtered.columns]

st.dataframe(
    filtered[display_cols].sort_values(["segment", "total_hits"], ascending=[True, False]),
    use_container_width=True,
    height=500,
    column_config={
        "submission_rate": st.column_config.ProgressColumn(
            "submission_rate", format="%.0f%%", min_value=0, max_value=1
        ),
    },
)

st.caption(f"Showing {len(filtered)} of {len(classified)} students.")

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
            half_life_days=half_life,
        )
        data = report.to_bytes(wb)
    fname = f"engagement_report_{subject_code or 'subject'}_{reference_date.isoformat()}.xlsx"
    st.download_button(
        "Download Excel report",
        data=data,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
