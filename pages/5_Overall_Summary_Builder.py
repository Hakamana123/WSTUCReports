"""
Blackboard Overall Summary Builder
====================================

Consolidates many per-student Blackboard activity CSVs into a single
"Overall Summary of User Activity" file in SpreadsheetML 2003 XML
format (.xls), matching Blackboard's native Overall export structure.

Also provides session-based engagement analytics (dwell time, bounce
rate, engagement score).
"""

from __future__ import annotations

import io
import re
from typing import Optional

import pandas as pd
import streamlit as st
import xml.etree.ElementTree as ET


# ---------------------------------------------------------------------------
# Event -> Area ID mapping (Blackboard taxonomy)
# ---------------------------------------------------------------------------

DEFAULT_EVENT_TO_AREA: dict[str, str] = {
    "Announcement read":       "announcements",
    "Document Access":         "content",
    "Scorm Item Access":       "content",
    "LTI Item Access":         "content",
    "Content Folder Access":   "content",
    "Discussion Access":       "discussion_board",
    "Discussion Response":     "discussion_board",
    "Discussion Reply":        "discussion_board",
    "Group Access":            "groups",
    "Journal Access":          "journal",
    "Journal Entry":           "journal",
    "Grade Center":            "instructor_gradebook",
    "My Grades":               "student_gradebook",
    "Subject access":          "",   # course entry -- not an area hit
    "Course access":           "",
}

STANDARD_AREA_IDS: list[str] = [
    "Bb-wiki", "announcements", "bb-achievements", "bb-attendance",
    "bb-collab-ultra", "bb-date-management",
    "bb-dropbox-integration-mashup", "bb-glossary", "bb-grading",
    "bb-item-analysis", "bb-learn-analytics-course-instructor-tool",
    "bb-learn-analytics-course-student-tool", "bb-mashups-flickr",
    "bb-mashups-slideshare", "bb-mashups-youtube", "bb-retention",
    "bb-rubric", "bb-selfpeer", "bb-vtbe-tinymce-matheditor",
    "bb-vtbe-tinymce-spellcheck", "bbcms-portfolio",
    "bbgs-gradejourney-gb_extract", "blogs", "calendar", "chat",
    "content", "control_panel", "course-files-contextmenu",
    "courses", "cplogs", "customization", "discussion_board",
    "external-ultra-pathway", "gradebook", "groups",
    "institution_pages", "instructor_gradebook", "journal",
    "learningstandards", "manage-users", "messages", "modules",
    "organizations", "outcomes-alignments", "periodicwork",
    "pk-bb2lti", "portfolios", "qti", "questionbank", "quota",
    "roles", "safeassign", "send_email", "stdy-studiosity",
    "student_gradebook", "tasks", "turn-turnitin", "turn-turnitin2",
    "turn-turnitin3",
]


# ---------------------------------------------------------------------------
# 1. Loading
# ---------------------------------------------------------------------------

def _read_csv_robust(file) -> pd.DataFrame:
    raw = file.read()
    file.seek(0)
    last_err: Optional[Exception] = None
    for enc in ("utf-8", "utf-8-sig", "utf-16", "latin-1"):
        try:
            text = raw.decode(enc)
        except UnicodeError as e:
            last_err = e
            continue
        first = next((ln for ln in text.splitlines() if ln.strip()), "")
        if "\t" in first:
            sep = "\t"
        elif ";" in first and first.count(";") > first.count(","):
            sep = ";"
        else:
            sep = ","
        try:
            return pd.read_csv(io.StringIO(text), sep=sep)
        except Exception as e:
            last_err = e
            continue
    raise ValueError(
        f"Could not parse {getattr(file, 'name', 'file')}: {last_err}"
    )


def load_activity_files(files) -> pd.DataFrame:
    frames = []
    for f in files:
        df = _read_csv_robust(f)
        df.columns = [str(c).strip() for c in df.columns]
        df["__filename__"] = f.name
        frames.append(df)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True, sort=False)


def load_roster(file) -> pd.DataFrame:
    df = _read_csv_robust(file)
    df.columns = [str(c).strip() for c in df.columns]
    return df


# ---------------------------------------------------------------------------
# 2. Join and clean
# ---------------------------------------------------------------------------

def _parse_datetime(s: pd.Series) -> pd.Series:
    normalized = (
        s.astype(str)
         .str.replace(
             r"\b(am|pm)\b",
             lambda m: m.group(1).upper(),
             regex=True,
             flags=re.IGNORECASE,
         )
    )
    return pd.to_datetime(normalized, dayfirst=True, errors="coerce")


def join_and_clean(
    activity: pd.DataFrame,
    roster: pd.DataFrame,
    roster_filename_col: str,
    roster_name_col: str,
    roster_id_col: str,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    r = roster[[roster_filename_col, roster_name_col, roster_id_col]].copy()
    r.columns = ["__filename__", "name", "student_id"]
    for c in r.columns:
        r[c] = r[c].astype(str).str.strip()

    joined = activity.merge(r, on="__filename__", how="left")

    unmatched = (
        joined[joined["name"].isna()]["__filename__"]
        .drop_duplicates()
        .to_frame()
    )
    joined = joined.dropna(subset=["name"]).copy()

    joined["ts"] = _parse_datetime(joined["Date and Time"])
    joined = joined.dropna(subset=["ts"])
    joined["display"] = joined["name"] + " (" + joined["student_id"] + ")"

    return joined, unmatched


def apply_event_mapping(df: pd.DataFrame, mapping: dict[str, str]) -> pd.DataFrame:
    out = df.copy()
    out["area_id"] = out["Event"].map(mapping).fillna("")
    return out


# ---------------------------------------------------------------------------
# 3. Section builders
# ---------------------------------------------------------------------------

def build_application_aggregate(
    df: pd.DataFrame, area_ids: list[str]
) -> tuple[pd.DataFrame, int]:
    m = df[df["area_id"] != ""]
    counts = m["area_id"].value_counts()
    total = int(counts.sum())
    rows = []
    for a in area_ids:
        hits = int(counts.get(a, 0))
        pct = (hits / total) if total else 0.0
        rows.append({"Area ID": a, "Hits": hits, "Per cent": pct})
    return pd.DataFrame(rows), total


def build_application_crosstab(
    df: pd.DataFrame, students: list[str], area_ids: list[str]
) -> pd.DataFrame:
    m = df[df["area_id"] != ""]
    pivot = (
        m.groupby(["display", "area_id"])
         .size()
         .unstack(fill_value=0)
         .reindex(index=students, columns=area_ids, fill_value=0)
    )
    pivot["Total"] = pivot.sum(axis=1)
    return pivot


def build_date_crosstab(
    df: pd.DataFrame, students: list[str]
) -> tuple[pd.DataFrame, list[pd.Timestamp]]:
    d = df.copy()
    if d.empty:
        return pd.DataFrame(index=students), []
    d["date"] = d["ts"].dt.normalize()
    dmin, dmax = d["date"].min(), d["date"].max()
    all_days = list(pd.date_range(dmin, dmax, freq="D"))
    pivot = (
        d.groupby(["display", "date"])
         .size()
         .unstack(fill_value=0)
         .reindex(index=students, columns=all_days, fill_value=0)
    )
    pivot.columns = [c.date() for c in pivot.columns]
    pivot["Total"] = pivot.sum(axis=1)
    return pivot, all_days


def build_hour_aggregate(df: pd.DataFrame) -> pd.DataFrame:
    hours = (
        df["ts"].dt.hour.value_counts()
                .reindex(range(24), fill_value=0).sort_index()
    )
    total = int(hours.sum())
    return pd.DataFrame({
        "Hour of Day": list(hours.index),
        "Hits": [int(v) for v in hours.values],
        "Per cent": [(v / total) if total else 0.0 for v in hours.values],
    })


def build_dayofweek_aggregate(df: pd.DataFrame) -> pd.DataFrame:
    labels = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]
    dows = (
        df["ts"].dt.dayofweek.value_counts()
                .reindex(range(7), fill_value=0).sort_index()
    )
    total = int(dows.sum())
    return pd.DataFrame({
        "Day of Week": labels,
        "Hits": [int(v) for v in dows.values],
        "Per cent": [(v / total) if total else 0.0 for v in dows.values],
    })


# ---------------------------------------------------------------------------
# 3b. Session analytics
# ---------------------------------------------------------------------------

SESSION_GAP_MINUTES: int = 30
DWELL_CAP_SECONDS: float = 30 * 60
BOUNCE_DWELL_THRESHOLD: float = 10.0


def create_sessions(df: pd.DataFrame) -> pd.DataFrame:
    out = df.sort_values(["display", "ts"]).copy()
    out["_gap"] = out.groupby("display")["ts"].diff()
    out["_new_session"] = (
        out["_gap"].isna()
        | (out["_gap"] > pd.Timedelta(minutes=SESSION_GAP_MINUTES))
    )
    out["session_id"] = out["_new_session"].cumsum()

    out["_next_gap"] = out.groupby("session_id")["_gap"].shift(-1)
    out["dwell_seconds"] = (
        out["_next_gap"]
        .dt.total_seconds()
        .clip(upper=DWELL_CAP_SECONDS)
    )

    sess_stats = out.groupby("session_id").agg(
        n_events=("ts", "size"),
        first_dwell=("dwell_seconds", "first"),
    )
    sess_stats["is_bounce"] = (
        (sess_stats["n_events"] == 1)
        | (
            (sess_stats["n_events"] == 1)
            & (sess_stats["first_dwell"].fillna(0) < BOUNCE_DWELL_THRESHOLD)
        )
    )
    out = out.merge(
        sess_stats[["is_bounce"]], left_on="session_id", right_index=True,
        how="left",
    )

    out.drop(columns=["_gap", "_new_session", "_next_gap"], inplace=True)
    return out


def _normalize_series(s: pd.Series) -> pd.Series:
    mn, mx = s.min(), s.max()
    if mx == mn:
        return pd.Series(0.0, index=s.index)
    return (s - mn) / (mx - mn)


def compute_metrics(df: pd.DataFrame, group_col: str) -> pd.DataFrame:
    g = df[df[group_col].astype(str).str.strip() != ""].copy()
    if g.empty:
        return pd.DataFrame()

    agg = g.groupby(group_col).agg(
        access_count=("ts", "size"),
        unique_students=("display", "nunique"),
        avg_dwell_sec=("dwell_seconds", "mean"),
        total_dwell_sec=("dwell_seconds", "sum"),
    ).reset_index()

    sess_bounce = (
        g.groupby([group_col, "session_id"])["is_bounce"]
         .first()
         .reset_index()
    )
    bounce_rates = (
        sess_bounce.groupby(group_col)["is_bounce"]
                   .mean()
                   .reset_index()
                   .rename(columns={"is_bounce": "bounce_rate"})
    )
    agg = agg.merge(bounce_rates, on=group_col, how="left")
    agg["bounce_rate"] = agg["bounce_rate"].fillna(0.0)

    norm_dwell = _normalize_series(agg["avg_dwell_sec"].fillna(0))
    norm_access = _normalize_series(agg["access_count"])
    norm_retain = 1 - agg["bounce_rate"]

    agg["engagement_score"] = (
        (norm_dwell * 0.5 + norm_access * 0.3 + norm_retain * 0.2) * 100
    ).round(1)

    agg["avg_dwell_sec"] = agg["avg_dwell_sec"].round(1)
    agg["total_dwell_sec"] = agg["total_dwell_sec"].round(1)
    agg["bounce_rate"] = (agg["bounce_rate"] * 100).round(1)

    agg = agg.rename(columns={
        group_col: "Group",
        "access_count": "Access Count",
        "unique_students": "Unique Students",
        "avg_dwell_sec": "Avg Dwell (s)",
        "total_dwell_sec": "Total Dwell (s)",
        "bounce_rate": "Bounce Rate (%)",
        "engagement_score": "Engagement (0-100)",
    })

    return agg.sort_values("Engagement (0-100)", ascending=False).reset_index(drop=True)


def compute_student_metrics(df: pd.DataFrame) -> pd.DataFrame:
    g = df.copy()
    stu = g.groupby("display").agg(
        total_events=("ts", "size"),
        sessions=("session_id", "nunique"),
        avg_dwell_sec=("dwell_seconds", "mean"),
        total_dwell_sec=("dwell_seconds", "sum"),
    ).reset_index()

    sess_bounce = (
        g.groupby(["display", "session_id"])["is_bounce"]
         .first()
         .reset_index()
    )
    br = (
        sess_bounce.groupby("display")["is_bounce"]
                   .mean()
                   .reset_index()
                   .rename(columns={"is_bounce": "bounce_rate"})
    )
    stu = stu.merge(br, on="display", how="left")
    stu["bounce_rate"] = (stu["bounce_rate"].fillna(0) * 100).round(1)
    stu["avg_dwell_sec"] = stu["avg_dwell_sec"].round(1)
    stu["total_dwell_sec"] = stu["total_dwell_sec"].round(1)

    stu = stu.rename(columns={
        "display": "Student",
        "total_events": "Events",
        "sessions": "Sessions",
        "avg_dwell_sec": "Avg Dwell (s)",
        "total_dwell_sec": "Total Dwell (s)",
        "bounce_rate": "Bounce Rate (%)",
    })
    return stu.sort_values("Events", ascending=False).reset_index(drop=True)


# ---------------------------------------------------------------------------
# 4. SpreadsheetML writer
# ---------------------------------------------------------------------------

NS = "urn:schemas-microsoft-com:office:spreadsheet"
SS = f"{{{NS}}}"


def _cell(row, value, cell_type: Optional[str] = None):
    c = ET.SubElement(row, f"{SS}Cell")
    if value is None or (isinstance(value, str) and value == ""):
        return c
    if cell_type is None:
        cell_type = (
            "Number"
            if isinstance(value, (int, float)) and not isinstance(value, bool)
            else "String"
        )
    d = ET.SubElement(c, f"{SS}Data", {f"{SS}Type": cell_type})
    d.text = str(value)
    return c


def _row(table):
    return ET.SubElement(table, f"{SS}Row")


def _blank_row(table):
    ET.SubElement(table, f"{SS}Row")


def write_spreadsheetml(
    app_agg: pd.DataFrame,
    app_total: int,
    app_xtab: pd.DataFrame,
    date_xtab: pd.DataFrame,
    all_days: list[pd.Timestamp],
    hour_agg: pd.DataFrame,
    dow_agg: pd.DataFrame,
    title: str = "Overall Summary of User Activity",
) -> bytes:
    ET.register_namespace("", NS)
    ET.register_namespace("o", "urn:schemas-microsoft-com:office:office")
    ET.register_namespace("x", "urn:schemas-microsoft-com:office:excel")
    ET.register_namespace("ss", NS)
    ET.register_namespace("html", "http://www.w3.org/TR/REC-html40")

    wb = ET.Element(f"{SS}Workbook")
    props = ET.SubElement(
        wb, "{urn:schemas-microsoft-com:office:office}DocumentProperties"
    )
    ET.SubElement(
        props, "{urn:schemas-microsoft-com:office:office}Title"
    ).text = title

    ws = ET.SubElement(
        wb, f"{SS}Worksheet", {f"{SS}Name": "Overall Summary of Usage"}
    )
    table = ET.SubElement(ws, f"{SS}Table")

    _cell(_row(table), title)

    # Section 1a: Access / Application (aggregate)
    _cell(_row(table), "Access / Application")
    r = _row(table)
    _cell(r, "Area ID"); _cell(r, "Hits"); _cell(r, "Per cent")
    for _, row_data in app_agg.iterrows():
        r = _row(table)
        _cell(r, row_data["Area ID"])
        _cell(r, int(row_data["Hits"]))
        _cell(r, float(row_data["Per cent"]))
    r = _row(table)
    _cell(r, "Total"); _cell(r, app_total)
    _cell(r, 1.0 if app_total else 0.0)

    _blank_row(table)

    # Section 1b: Access / Application (per student)
    _cell(_row(table), "Access / Application (per student)")
    r = _row(table)
    _cell(r, "Student")
    for a in app_xtab.columns:
        _cell(r, a)
    for student in app_xtab.index:
        r = _row(table)
        _cell(r, student)
        for v in app_xtab.loc[student].tolist():
            _cell(r, int(v))

    _blank_row(table)

    # Section 2: Access / Date
    _cell(_row(table), "Access / Date")
    r = _row(table)
    _cell(r, "Student")
    for d in all_days:
        _cell(r, d.strftime("%Y-%m-%d"))
    if all_days:
        _cell(r, "Total")
    for student in date_xtab.index:
        r = _row(table)
        _cell(r, student)
        if all_days:
            for d in all_days:
                key = d.date() if hasattr(d, "date") else d
                _cell(r, int(date_xtab.loc[student, key]))
            _cell(r, int(date_xtab.loc[student, "Total"]))

    _blank_row(table)

    # Section 3: Access / Hour of Day
    _cell(_row(table), "Access / Hour of Day")
    r = _row(table)
    _cell(r, "Hour of Day"); _cell(r, "Hits"); _cell(r, "Per cent")
    for _, row_data in hour_agg.iterrows():
        r = _row(table)
        _cell(r, int(row_data["Hour of Day"]))
        _cell(r, int(row_data["Hits"]))
        _cell(r, float(row_data["Per cent"]))

    _blank_row(table)

    # Section 4: Access / Day of Week
    _cell(_row(table), "Access / Day of Week")
    r = _row(table)
    _cell(r, "Day of Week"); _cell(r, "Hits"); _cell(r, "Per cent")
    for _, row_data in dow_agg.iterrows():
        r = _row(table)
        _cell(r, row_data["Day of Week"])
        _cell(r, int(row_data["Hits"]))
        _cell(r, float(row_data["Per cent"]))

    body = ET.tostring(wb, encoding="unicode")
    return (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<?mso-application progid="Excel.Sheet"?>\n'
        + body
    ).encode("utf-8")


# ---------------------------------------------------------------------------
# Roster template helper
# ---------------------------------------------------------------------------

def _roster_template_bytes() -> bytes:
    example = pd.DataFrame([
        {"filename": "student-activity-20260416T122939.csv",
         "name": "Smith, John", "student_id": "22001234"},
        {"filename": "student-activity-20260416T122947.csv",
         "name": "Nguyen, Linh", "student_id": "22005678"},
    ])
    return example.to_csv(index=False).encode("utf-8")


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

st.sidebar.title("Overall Summary Builder")

with st.sidebar:
    st.markdown("**1. Activity files**")
    activity_files = st.file_uploader(
        "Per-student activity CSVs",
        type=["csv", "txt"],
        accept_multiple_files=True,
        help="Columns expected: Event, Item, IP Address, Date and Time",
        key="osb_activity",
    )

    st.markdown("**2. Roster**")
    roster_file = st.file_uploader(
        "Roster CSV (filename / name / student_id)",
        type=["csv", "txt"],
        help="One row per student. The filename column must match the "
             "uploaded activity CSV filenames exactly.",
        key="osb_roster",
    )
    st.download_button(
        "Download roster template",
        data=_roster_template_bytes(),
        file_name="roster_template.csv",
        mime="text/csv",
        key="osb_template",
    )


st.title("📦 Overall Summary Builder")
st.caption(
    "Consolidates per-student Blackboard activity CSVs into an Overall "
    "Summary (.xls / SpreadsheetML 2003 XML) matching the four-section "
    "structure of Blackboard's native Overall export."
)

if not activity_files or not roster_file:
    st.info("Upload per-student activity CSVs and a roster CSV in the sidebar to begin.")
    st.stop()

# ── Roster ────────────────────────────────────────────────────────

try:
    roster = load_roster(roster_file)
except Exception as e:
    st.error(f"Could not read roster: {e}")
    st.stop()

st.subheader("Roster preview")
st.dataframe(roster.head(10), use_container_width=True)
st.caption(f"{len(roster):,} roster rows.")

st.subheader("Roster column mapping")
rcols = list(roster.columns)


def _default_idx(name: str) -> int:
    return rcols.index(name) if name in rcols else 0


c1, c2, c3 = st.columns(3)
with c1:
    rfn = st.selectbox("Filename column", rcols, index=_default_idx("filename"), key="osb_rfn")
with c2:
    rn = st.selectbox("Name column", rcols, index=_default_idx("name"), key="osb_rn")
with c3:
    rid = st.selectbox("Student ID column", rcols, index=_default_idx("student_id"), key="osb_rid")

# ── Activity files ────────────────────────────────────────────────

try:
    activity = load_activity_files(activity_files)
except Exception as e:
    st.error(f"Could not read activity files: {e}")
    st.stop()

required = {"Event", "Item", "Date and Time"}
missing = required - set(activity.columns)
if missing:
    st.error(f"Activity files missing required columns: {sorted(missing)}")
    st.stop()

joined, unmatched = join_and_clean(activity, roster, rfn, rn, rid)

c1, c2, c3, c4 = st.columns(4)
c1.metric("Files uploaded", len(activity_files))
c2.metric("Raw events", f"{len(activity):,}")
c3.metric("Matched events", f"{len(joined):,}")
c4.metric("Unmatched files", len(unmatched))

if len(unmatched):
    with st.expander("Unmatched files — no roster row"):
        st.dataframe(unmatched, use_container_width=True)

if joined.empty:
    st.error(
        "No activity events matched the roster. Check that the roster "
        "filename column contains exact filenames."
    )
    st.stop()

# ── Event → Area ID mapping ──────────────────────────────────────

st.subheader("Event to Area ID mapping")
st.caption(
    "Blackboard's Overall report groups hits by Area ID (internal "
    "building-block name). Edit the right column to change the mapping. "
    "Leave blank to exclude an event from Application hit counts."
)
events_in_data = sorted(joined["Event"].dropna().unique())
default_rows = [
    {"Event": e, "Area ID": DEFAULT_EVENT_TO_AREA.get(e, "")}
    for e in events_in_data
]
edited = st.data_editor(
    pd.DataFrame(default_rows),
    num_rows="fixed",
    use_container_width=True,
    key="osb_event_map",
    disabled=["Event"],
)
mapping = dict(zip(edited["Event"], edited["Area ID"].fillna("")))

mapped = apply_event_mapping(joined, mapping)
c1, c2 = st.columns(2)
c1.metric("Events mapped to an Area ID",
          f"{(mapped['area_id'] != '').sum():,}")
c2.metric("Events excluded (no Area ID)",
          f"{(mapped['area_id'] == '').sum():,}")

# ── Students and area IDs ────────────────────────────────────────

r_slim = roster.rename(columns={rn: "name", rid: "student_id"})
student_display = (
    r_slim["name"].astype(str).str.strip()
    + " (" + r_slim["student_id"].astype(str).str.strip() + ")"
)
students = sorted(student_display.unique())

extra = sorted({
    v for v in mapping.values() if v and v not in STANDARD_AREA_IDS
})
area_ids = STANDARD_AREA_IDS + extra

# ── Build sections ────────────────────────────────────────────────

app_agg, app_total = build_application_aggregate(mapped, area_ids)
app_xtab = build_application_crosstab(mapped, students, area_ids)
date_xtab, all_days = build_date_crosstab(mapped, students)
hour_agg = build_hour_aggregate(mapped)
dow_agg = build_dayofweek_aggregate(mapped)

st.subheader("Section previews")
tabs = st.tabs([
    "Application (aggregate)",
    "Application (per student)",
    "Date (per student)",
    "Hour of Day",
    "Day of Week",
])
with tabs[0]:
    st.caption(f"Total hits: {app_total:,}")
    st.dataframe(
        app_agg[app_agg["Hits"] > 0].reset_index(drop=True),
        use_container_width=True,
    )
    with st.expander("Show all area IDs (including zeros)"):
        st.dataframe(app_agg, use_container_width=True)
with tabs[1]:
    preview = app_xtab.loc[:, (app_xtab != 0).any(axis=0)]
    st.dataframe(preview, use_container_width=True)
    st.caption(
        f"{len(app_xtab):,} students x {len(app_xtab.columns)} area IDs "
        f"(preview hides all-zero columns; export keeps them)."
    )
with tabs[2]:
    if all_days:
        st.caption(
            f"Date range: {all_days[0].date()} to {all_days[-1].date()} "
            f"({len(all_days)} days)"
        )
        st.dataframe(date_xtab, use_container_width=True)
    else:
        st.info("No dated events to show.")
with tabs[3]:
    st.dataframe(hour_agg, use_container_width=True)
with tabs[4]:
    st.dataframe(dow_agg, use_container_width=True)

# ── Analytics ─────────────────────────────────────────────────────

st.subheader("Analytics")
st.caption(
    "Session-based engagement metrics. Sessions are inferred from "
    "timestamp gaps (>30 min = new session). Dwell time is estimated "
    "from inter-event gaps within a session, capped at 30 min. "
    "Bounce = session with only 1 interaction. "
    "Engagement score is normalised 0\u2013100 "
    "(50% avg dwell + 30% access count + 20% retention)."
)

sessioned = create_sessions(mapped)

total_sessions = sessioned["session_id"].nunique()
total_bounced = sessioned.groupby("session_id")["is_bounce"].first().sum()
overall_bounce = (total_bounced / total_sessions * 100) if total_sessions else 0
mc1, mc2, mc3, mc4 = st.columns(4)
mc1.metric("Total sessions", f"{total_sessions:,}")
mc2.metric("Bounced sessions", f"{int(total_bounced):,}")
mc3.metric("Overall bounce rate", f"{overall_bounce:.1f}%")
mc4.metric(
    "Median dwell (s)",
    f"{sessioned['dwell_seconds'].median():.1f}"
    if sessioned["dwell_seconds"].notna().any() else "N/A",
)

# Filters
fc1, fc2 = st.columns(2)
with fc1:
    all_students = sorted(sessioned["display"].unique())
    sel_students = st.multiselect(
        "Filter by student", all_students, default=[],
        key="osb_analytics_student_filter",
    )
with fc2:
    all_items = sorted(sessioned["Item"].dropna().unique())
    sel_items = st.multiselect(
        "Filter by Item", all_items, default=[],
        key="osb_analytics_item_filter",
    )

filtered = sessioned.copy()
if sel_students:
    filtered = filtered[filtered["display"].isin(sel_students)]
if sel_items:
    filtered = filtered[filtered["Item"].isin(sel_items)]

if filtered.empty:
    st.warning("No events match the current filters.")
else:
    analytics_tabs = st.tabs([
        "By Item",
        "By Event Type",
        "By Area ID",
        "Per Student",
    ])

    with analytics_tabs[0]:
        m_item = compute_metrics(filtered, "Item")
        if m_item.empty:
            st.info("No item-level data.")
        else:
            st.dataframe(
                m_item, use_container_width=True, hide_index=True,
            )
            st.caption(f"{len(m_item)} items.")
            ch1, ch2 = st.columns(2)
            with ch1:
                top_n = m_item.head(15)
                st.bar_chart(
                    top_n.set_index("Group")["Access Count"],
                    use_container_width=True,
                )
                st.caption("Access count (top 15 items)")
            with ch2:
                st.bar_chart(
                    top_n.set_index("Group")["Engagement (0-100)"],
                    use_container_width=True,
                )
                st.caption("Engagement score (top 15 items)")

    with analytics_tabs[1]:
        m_event = compute_metrics(filtered, "Event")
        if m_event.empty:
            st.info("No event-type data.")
        else:
            st.dataframe(
                m_event, use_container_width=True, hide_index=True,
            )
            ch1, ch2 = st.columns(2)
            with ch1:
                st.bar_chart(
                    m_event.set_index("Group")["Access Count"],
                    use_container_width=True,
                )
                st.caption("Access count by event type")
            with ch2:
                st.bar_chart(
                    m_event.set_index("Group")["Engagement (0-100)"],
                    use_container_width=True,
                )
                st.caption("Engagement score by event type")

    with analytics_tabs[2]:
        m_area = compute_metrics(filtered, "area_id")
        if m_area.empty:
            st.info(
                "No area-level data. Events without an Area ID mapping "
                "are excluded."
            )
        else:
            st.dataframe(
                m_area, use_container_width=True, hide_index=True,
            )
            ch1, ch2 = st.columns(2)
            with ch1:
                st.bar_chart(
                    m_area.set_index("Group")["Access Count"],
                    use_container_width=True,
                )
                st.caption("Access count by Area ID")
            with ch2:
                st.bar_chart(
                    m_area.set_index("Group")["Bounce Rate (%)"],
                    use_container_width=True,
                )
                st.caption("Bounce rate (%) by Area ID")

    with analytics_tabs[3]:
        m_stu = compute_student_metrics(filtered)
        if m_stu.empty:
            st.info("No student data.")
        else:
            st.dataframe(
                m_stu, use_container_width=True, hide_index=True,
            )
            st.caption(f"{len(m_stu)} students.")

# ── Export ────────────────────────────────────────────────────────

st.subheader("Export")
title = st.text_input(
    "Report title", value="Overall Summary of User Activity",
    key="osb_title",
)
filename = st.text_input("Output filename", value="overall_summary.xls", key="osb_filename")

xml_bytes = write_spreadsheetml(
    app_agg, app_total, app_xtab,
    date_xtab, all_days,
    hour_agg, dow_agg,
    title=title,
)
st.download_button(
    "Download Overall Summary (.xls)",
    data=xml_bytes,
    file_name=filename,
    mime="application/vnd.ms-excel",
    type="primary",
    key="osb_download",
)

# Analytics exports
if not filtered.empty:
    st.caption("Analytics exports")
    ec1, ec2, ec3 = st.columns(3)
    with ec1:
        m_item_exp = compute_metrics(sessioned, "Item")
        if not m_item_exp.empty:
            st.download_button(
                "Metrics by Item (.csv)",
                data=m_item_exp.to_csv(index=False).encode("utf-8"),
                file_name="metrics_by_item.csv",
                mime="text/csv",
                key="osb_exp_item",
            )
    with ec2:
        m_event_exp = compute_metrics(sessioned, "Event")
        if not m_event_exp.empty:
            st.download_button(
                "Metrics by Event (.csv)",
                data=m_event_exp.to_csv(index=False).encode("utf-8"),
                file_name="metrics_by_event.csv",
                mime="text/csv",
                key="osb_exp_event",
            )
    with ec3:
        m_stu_exp = compute_student_metrics(sessioned)
        if not m_stu_exp.empty:
            st.download_button(
                "Metrics by Student (.csv)",
                data=m_stu_exp.to_csv(index=False).encode("utf-8"),
                file_name="metrics_by_student.csv",
                mime="text/csv",
                key="osb_exp_student",
            )
