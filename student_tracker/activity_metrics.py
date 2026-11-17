"""Engagement analytics built from the vUWS activity export.

Online-store analytics, translated to a subject site:

  visit / session  -> a run of clicks with no gap longer than `gap_minutes`
                      (default 30, the usual web-analytics convention)
  bounce           -> a session where the student entered the subject and
                      did nothing else (only COURSE_ACCESS clicks)
  time on site     -> session length, first click to last click. A LOWER
                      BOUND: time after the last click of a session is not
                      recorded anywhere, so one-click sessions count as 0.
  time on page     -> gap from opening an item to the student's next click
                      in the same session. Not estimated for External tool /
                      SCORM items (they open outside vUWS, so the gap is the
                      launch click, not time spent), nor for the last click
                      of a session (no next click to measure against).
  reach            -> share of the class who opened an item at least once
                      (stands in for click-through rate: the export has no
                      'was shown it' count to divide by)
  exit rate        -> share of an item's session-views that ended the session
  recency segment  -> Engaged / Fading / Lapsed by days since last visit,
                      plus Never visited (on the class list, no clicks) and
                      Not on class list (clicks, but not enrolled - usually
                      a withdrawal)

Every function works on ONE subject's events; the page filters first.
"""

from __future__ import annotations

import re
from datetime import date, timedelta

import pandas as pd

SEGMENTS = ["Engaged", "Fading", "Lapsed", "Never visited", "Not on class list"]
SEGMENT_HELP = {
    "Engaged": "Visited in the last {active} days",
    "Fading": "Last visit {active_plus}-{lapse} days ago",
    "Lapsed": "No visit for more than {lapse} days",
    "Never visited": "On the class list, no clicks in the export",
    "Not on class list": "Has clicks but isn't on the class list (withdrawn?)",
}
ASSESS_RE = re.compile(r"assessment\s*(\d+)", re.IGNORECASE)
WEEKDAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]


# ------------------------------------------------------------------
# Sessions
# ------------------------------------------------------------------

def sessionise(events: pd.DataFrame, gap_minutes: int = 30) -> pd.DataFrame:
    """Add session_id, gap_next_s, is_exit and dwell_s to each click."""
    ev = events.sort_values(["subject_code", "student_code", "ts"]).copy()
    keys = [ev["subject_code"], ev["student_code"]]
    gap = ev.groupby(keys)["ts"].diff().dt.total_seconds()
    new_session = gap.isna() | (gap > gap_minutes * 60)
    session_no = new_session.groupby(keys).cumsum().astype(int)
    ev["session_id"] = (
        ev["subject_code"] + "|" + ev["student_code"] + "|" + session_no.astype(str)
    )
    nxt = ev.groupby("session_id")["ts"].shift(-1)
    ev["gap_next_s"] = (nxt - ev["ts"]).dt.total_seconds()
    ev["is_exit"] = nxt.isna()
    ev["dwell_s"] = ev["gap_next_s"].where(~ev["is_external"])
    ev["is_action"] = ev["event"] != "COURSE_ACCESS"
    return ev.reset_index(drop=True)


def session_table(ev: pd.DataFrame) -> pd.DataFrame:
    """One row per session."""
    if ev.empty:
        return pd.DataFrame(columns=[
            "session_id", "student_code", "start", "end", "clicks",
            "actions", "minutes", "bounce", "items",
        ])
    g = ev.groupby("session_id", sort=False)
    s = g.agg(
        student_code=("student_code", "first"),
        start=("ts", "min"),
        end=("ts", "max"),
        clicks=("ts", "size"),
        actions=("is_action", "sum"),
    )
    s["minutes"] = (s["end"] - s["start"]).dt.total_seconds() / 60
    s["bounce"] = s["actions"] == 0
    acts = ev[ev["is_action"]]
    s["items"] = acts.groupby("session_id")["item"].agg(
        lambda x: " → ".join(dict.fromkeys(x))  # unique, in click order
    )
    s["items"] = s["items"].fillna("(entered subject only)")
    return s.reset_index()


# ------------------------------------------------------------------
# Students
# ------------------------------------------------------------------

def _as_at_end(as_at: date) -> pd.Timestamp:
    return pd.Timestamp(as_at) + pd.Timedelta(days=1)


def student_summary(
    ev: pd.DataFrame,
    sessions: pd.DataFrame,
    as_at: date,
    roster: pd.DataFrame | None = None,
    active_days: int = 7,
    lapse_days: int = 14,
) -> pd.DataFrame:
    """One row per student: recency, frequency, time, coverage, segment.

    `roster` (optional) has student_code, display_name and any extra
    class-list columns. Students on it with no clicks become 'Never
    visited'; students with clicks who aren't on it become 'Not on class
    list'.
    """
    end = _as_at_end(as_at)
    cols = ["student_code"]
    if ev.empty:
        base = pd.DataFrame(columns=cols)
    else:
        g = ev.groupby("student_code")
        base = g.agg(
            log_name=("student_name", "first"),
            first_seen=("ts", "min"),
            last_seen=("ts", "max"),
            clicks=("ts", "size"),
        )
        base["active_days"] = g["ts"].agg(lambda t: t.dt.date.nunique())
        content = ev[ev["event"] == "CONTENT_ACCESS"]
        base["items_opened"] = content.groupby("student_code")["item"].nunique()
        base["announcements_read"] = (
            ev[ev["event"] == "ANNOUNCEMENT_READ"].groupby("student_code").size()
        )
        base["submissions_logged"] = (
            ev[ev["event"] == "ASSESSMENT_SUBMIT"].groupby("student_code").size()
        )
        sg = sessions.groupby("student_code")
        base["sessions"] = sg.size()
        base["est_hours"] = sg["minutes"].sum() / 60
        base["bounce_rate"] = sg["bounce"].mean()
        base["avg_session_min"] = sg["minutes"].mean()
        recent = sessions[sessions["start"] >= end - pd.Timedelta(days=7)]
        prior = sessions[
            (sessions["start"] >= end - pd.Timedelta(days=14))
            & (sessions["start"] < end - pd.Timedelta(days=7))
        ]
        base["sessions_last_7d"] = recent.groupby("student_code").size()
        base["sessions_prev_7d"] = prior.groupby("student_code").size()
        n_items = content["item"].nunique()
        base["coverage"] = base["items_opened"] / n_items if n_items else 0.0
        base = base.reset_index()
        fill0 = ["items_opened", "announcements_read", "submissions_logged",
                 "sessions_last_7d", "sessions_prev_7d", "coverage"]
        base[fill0] = base[fill0].fillna(0)

    if roster is not None and not roster.empty:
        base["_in_log"] = True
        r = roster.copy()
        r["_on_roster"] = True
        out = r.merge(base, on="student_code", how="outer")
        out["_in_log"] = out["_in_log"].fillna(False).astype(bool)
        out["_on_roster"] = out["_on_roster"].fillna(False).astype(bool)
        out["display_name"] = out["display_name"].fillna(out.get("log_name"))
    else:
        out = base.copy()
        out["_in_log"] = True
        out["_on_roster"] = True
        out["display_name"] = out.get("log_name")

    for c in ["clicks", "sessions", "active_days", "items_opened",
              "announcements_read", "submissions_logged",
              "sessions_last_7d", "sessions_prev_7d"]:
        if c in out.columns:
            out[c] = out[c].fillna(0).astype(int)
    for c in ["est_hours", "coverage"]:
        if c in out.columns:
            out[c] = out[c].fillna(0.0)

    if "last_seen" in out.columns:
        out["days_since_last"] = (
            (pd.Timestamp(as_at) - out["last_seen"].dt.normalize()).dt.days
        )
    else:
        out["days_since_last"] = pd.NA

    def seg(row):
        if not row["_on_roster"]:
            return "Not on class list"
        if not row["_in_log"]:
            return "Never visited"
        d = row["days_since_last"]
        if d <= active_days:
            return "Engaged"
        if d <= lapse_days:
            return "Fading"
        return "Lapsed"

    out["segment"] = out.apply(seg, axis=1) if len(out) else pd.Series(dtype=str)
    out["segment"] = pd.Categorical(out["segment"], categories=SEGMENTS, ordered=True)
    out = out.drop(columns=[c for c in ["log_name"] if c in out.columns])
    return out.sort_values(["segment", "days_since_last", "display_name"]).reset_index(drop=True)


def daily_clicks(ev: pd.DataFrame) -> pd.DataFrame:
    """Clicks per student per day."""
    d = ev.assign(day=ev["ts"].dt.normalize())
    return d.groupby(["student_code", "day"]).size().rename("clicks").reset_index()


# ------------------------------------------------------------------
# Content
# ------------------------------------------------------------------

def content_table(ev: pd.DataFrame, roster_size: int) -> pd.DataFrame:
    """One row per item students opened (CONTENT_ACCESS only)."""
    c = ev[ev["event"] == "CONTENT_ACCESS"]
    if c.empty:
        return pd.DataFrame()
    g = c.groupby("item")
    t = g.agg(
        kind=("kind", "first"),
        students=("student_code", "nunique"),
        session_views=("session_id", "nunique"),
        clicks=("ts", "size"),
        median_time_s=("dwell_s", "median"),
        first_opened=("ts", "min"),
        last_opened=("ts", "max"),
    )
    exits = c[c["is_exit"]].groupby("item")["session_id"].nunique()
    t["exit_rate"] = (exits / t["session_views"]).fillna(0.0)
    t["reach"] = t["students"] / roster_size if roster_size else 0.0
    t["views_per_student"] = t["session_views"] / t["students"]
    return t.reset_index().sort_values(
        ["students", "session_views"], ascending=False
    ).reset_index(drop=True)


def students_not_opened(ev: pd.DataFrame, item: str, roster_codes: list[str]) -> list[str]:
    opened = set(ev.loc[(ev["event"] == "CONTENT_ACCESS") & (ev["item"] == item), "student_code"])
    return [s for s in roster_codes if s not in opened]


# ------------------------------------------------------------------
# Assessments
# ------------------------------------------------------------------

def assessment_number(title: str) -> int | None:
    m = ASSESS_RE.search(title or "")
    return int(m.group(1)) if m else None


def _role(item: str, kind: str) -> str:
    low = item.lower()
    if "resubmission" in low:
        return "Resubmission"
    if low.startswith("about assessment"):
        return "About page"
    if kind == "Discussion forum":
        return "Q&A forum"
    if kind in ("Assessment submission", "External tool"):
        return "Submission point"
    return "Other"


def assessment_items(ev: pd.DataFrame) -> pd.DataFrame:
    """Each assessment-related item with its assessment number and role."""
    items = ev[ev["event"].isin(["CONTENT_ACCESS", "ASSESSMENT_SUBMIT", "ASSESSMENT_DRAFT"])]
    items = items.groupby("item", as_index=False)["kind"].first()
    items["assessment"] = items["item"].map(assessment_number)
    items = items.dropna(subset=["assessment"])
    items["assessment"] = items["assessment"].astype(int)
    items["role"] = [_role(i, k) for i, k in zip(items["item"], items["kind"])]
    return items.sort_values(["assessment", "role", "item"]).reset_index(drop=True)


NOT_LOGGED = "Not logged (Turnitin / external tool) - add a Grade Centre"
STAGES = ["Not opened", "Opened other items only", "Read About page",
          "Opened submission point", "Submitted"]


def assessment_progress(
    ev: pd.DataFrame,
    roster_codes: list[str],
    grade_submitted: dict[int, set[str]] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Returns (funnel, matrix).

    funnel: one row per assessment with counts at each stage, whether
      submission is logged in the export, and median lead time (days
      from first opening any item for that assessment to submitting).
    matrix: student_code x 'A1', 'A2'... holding the furthest stage
      reached (one of STAGES).

    Submission comes from the Grade Centre when one is given for that
    assessment (any mark or Satisfactory/Complete text counts), else from
    ASSESSMENT_SUBMIT clicks. Turnitin / LTI submissions are not logged as
    clicks, so without a Grade Centre those assessments show no
    submissions.
    """
    items = assessment_items(ev)
    if items.empty:
        return pd.DataFrame(), pd.DataFrame({"student_code": roster_codes})
    role_of = dict(zip(items["item"], items["role"]))
    num_of = dict(zip(items["item"], items["assessment"]))
    a_ev = ev[ev["item"].isin(num_of)].copy()
    a_ev["assessment"] = a_ev["item"].map(num_of)
    a_ev["role"] = a_ev["item"].map(role_of)

    funnel_rows = []
    matrix = pd.DataFrame({"student_code": roster_codes})
    roster = set(roster_codes)
    for n in sorted(items["assessment"].unique()):
        e = a_ev[(a_ev["assessment"] == n) & a_ev["student_code"].isin(roster)]
        opened = e[e["event"] == "CONTENT_ACCESS"]
        any_open = set(opened["student_code"])
        about = set(opened.loc[opened["role"] == "About page", "student_code"])
        point = set(opened.loc[opened["role"] == "Submission point", "student_code"])
        logged_items = items[(items["assessment"] == n) & (items["kind"] == "Assessment submission")]
        submits = e[e["event"] == "ASSESSMENT_SUBMIT"]
        submitted_log = set(submits["student_code"])
        if grade_submitted and n in grade_submitted:
            submitted = grade_submitted[n] & roster
            source = "Grade Centre"
        elif not logged_items.empty:
            submitted = submitted_log
            source = "Activity log"
        else:
            submitted = set()
            source = NOT_LOGGED
        not_logged = source == NOT_LOGGED

        lead = []
        if not submits.empty:
            first_open = opened.groupby("student_code")["ts"].min()
            first_sub = submits.groupby("student_code")["ts"].min()
            both = pd.concat([first_open, first_sub], axis=1, keys=["o", "s"]).dropna()
            lead = ((both["s"] - both["o"]).dt.total_seconds() / 86400).tolist()

        funnel_rows.append({
            "assessment": f"Assessment {n}",
            "class_size": len(roster),
            "opened_any": len(any_open),
            "read_about": len(about),
            "opened_submission_point": len(point),
            "submitted": None if not_logged else len(submitted),
            "submission_source": source,
            "opened_not_submitted": None if not_logged else len(any_open - submitted),
            "median_lead_days": float(pd.Series(lead).median()) if lead else None,
        })

        def stage(code):
            if code in submitted:
                return STAGES[4]
            if code in point:
                return STAGES[3]
            if code in about:
                return STAGES[2]
            if code in any_open:
                return STAGES[1]
            return STAGES[0]

        matrix[f"A{n}"] = matrix["student_code"].map(stage)
    funnel = pd.DataFrame(funnel_rows)
    return funnel, matrix


def grade_submissions(grade_parsed: dict) -> dict[int, set[str]]:
    """Map assessment number -> student codes with a submission in the
    Grade Centre (parsed by parsers.grade_centre)."""
    from student_tracker.parsers.grade_centre import _is_submitted

    students = grade_parsed["students"]
    out: dict[int, set[str]] = {}
    for a in grade_parsed["assessments"]:
        n = assessment_number(a["name"])
        if n is None:
            continue
        mask = students[a["name"]].map(_is_submitted)
        out.setdefault(n, set()).update(students.loc[mask, "student_code"])
    return out


# ------------------------------------------------------------------
# Time patterns
# ------------------------------------------------------------------

def weekly_activity(ev: pd.DataFrame, sessions: pd.DataFrame, roster_size: int) -> pd.DataFrame:
    """Per calendar week (Mon start): active students, sessions, clicks."""
    if ev.empty:
        return pd.DataFrame()
    wk = ev["ts"].dt.to_period("W-SUN").dt.start_time
    w = ev.assign(week=wk).groupby("week").agg(
        active_students=("student_code", "nunique"),
        clicks=("ts", "size"),
    )
    s_wk = sessions["start"].dt.to_period("W-SUN").dt.start_time
    w["sessions"] = sessions.assign(week=s_wk).groupby("week").size()
    w["bounce_rate"] = sessions.assign(week=s_wk).groupby("week")["bounce"].mean()
    w["pct_active"] = w["active_students"] / roster_size if roster_size else 0.0
    w = w.reset_index()
    # Weeks the export only partly covers (first / last) get flagged so a
    # one-day week isn't read as a collapse in activity.
    first_day, last_day = ev["ts"].min().normalize(), ev["ts"].max().normalize()
    w["part_week"] = (w["week"] < first_day) | (w["week"] + pd.Timedelta(days=6) > last_day)
    w["week_label"] = w["week"].dt.strftime("%d %b") + w["part_week"].map(
        {True: " (part)", False: ""}
    )
    return w


def hour_heatmap(ev: pd.DataFrame) -> pd.DataFrame:
    """Clicks and distinct students by weekday x hour."""
    h = ev.assign(weekday=ev["ts"].dt.dayofweek.map(dict(enumerate(WEEKDAYS))),
                  hour=ev["ts"].dt.hour)
    return h.groupby(["weekday", "hour"]).agg(
        clicks=("ts", "size"), students=("student_code", "nunique")
    ).reset_index()


# ------------------------------------------------------------------
# Headline numbers
# ------------------------------------------------------------------

def kpis(students: pd.DataFrame, sessions: pd.DataFrame, ev: pd.DataFrame,
         active_days: int = 7) -> dict:
    enrolled = students[students["_on_roster"]]
    visited = enrolled[enrolled["_in_log"]]
    n = len(enrolled)
    readers = set(ev.loc[ev["event"] == "ANNOUNCEMENT_READ", "student_code"])
    ann = len(readers & set(enrolled["student_code"]))
    return {
        "enrolled": n,
        "active_recent": int((enrolled["segment"] == "Engaged").sum()),
        "never_visited": int((enrolled["segment"] == "Never visited").sum()),
        "not_on_list": int((students["segment"] == "Not on class list").sum()),
        "median_sessions": float(visited["sessions"].median()) if len(visited) else 0.0,
        "median_hours": float(visited["est_hours"].median()) if len(visited) else 0.0,
        "bounce_rate": float(sessions["bounce"].mean()) if len(sessions) else 0.0,
        "median_session_min": float(sessions["minutes"].median()) if len(sessions) else 0.0,
        "announcement_reach": ann / n if n else 0.0,
        "median_coverage": float(visited["coverage"].median()) if len(visited) else 0.0,
    }
