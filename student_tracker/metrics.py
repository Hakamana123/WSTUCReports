"""Per-student metric computations.

Rebuilt for weekly-snapshot inputs (replacing the old daily-hit-count
feed, which depended on a report that's no longer reliably exportable).

Each reporting point now comes from THREE independently-pulled weekly
exports, each narrowed to a single week in the Blackboard UI:
  - Login Report        -> weekly login COUNT, via the delta between this
                            week's and last week's cumulative TOTAL LOGINS
                            snapshot (see parsers/login_report.py for why
                            that field is safe to diff). Also the source
                            of lifetime last_login_date / total_logins for
                            S1/S2, taken from the most recent snapshot.
  - Subject Activity Overview -> hours in subject that week (already
                            window-scoped by the report itself).
  - User Activity in Forums   -> forum accesses + messages that week
                            (context only; does not drive segmentation).

Everything is keyed by BLOCK-RELATIVE WEEK NUMBER (an int): the week
before teaching starts is Week 0 (used only as the login-delta baseline,
per the user's own numbering), the first teaching week is Week 1, etc.
This matches block_week_number()'s existing definition below.

Because inputs are weekly totals rather than daily hit counts, some of
the old daily-precision metrics no longer exist in the same form:
  - total_active_days (daily count)   -> weeks_active (weekly count)
  - last_hit_date / days_since_last_hit (day precision) -> last_active_week
    / weeks_since_last_active (week precision)
  - recency-weighted hits (day half-life) -> dropped. A day-granularity
    half-life doesn't mean much when the underlying data IS the week; see
    README notes in report.py for what replaced it.
"""

from datetime import date
import pandas as pd


def iso_week_key(d: date | pd.Timestamp) -> tuple[int, int]:
    """Return (iso_year, iso_week) for a given date."""
    if isinstance(d, pd.Timestamp):
        d = d.date()
    cal = d.isocalendar()
    return (cal[0], cal[1])


def iso_week_label(year: int, week: int) -> str:
    """Human-readable ISO week label like '2026-W14 (Mon 30 Mar)'."""
    monday = date.fromisocalendar(year, week, 1)
    return f"{year}-W{week:02d} (Mon {monday:%d %b})"


def block_week_number(iso_year: int, iso_week: int, block_start_date: date) -> int:
    """Block-relative week number.

    Block start week = 1; the week immediately before = 0; earlier = -1, -2, …
    """
    target_monday = date.fromisocalendar(iso_year, iso_week, 1)
    bsd_cal = block_start_date.isocalendar()
    block_monday = date.fromisocalendar(bsd_cal[0], bsd_cal[1], 1)
    weeks_diff = (target_monday - block_monday).days // 7
    return weeks_diff + 1


def block_week_label(iso_year: int, iso_week: int, block_start_date: date) -> str:
    """Block-relative week label like 'W1 (Mon 30 Mar)' or 'W0 (Mon 23 Mar)'."""
    bw = block_week_number(iso_year, iso_week, block_start_date)
    monday = date.fromisocalendar(iso_year, iso_week, 1)
    return f"W{bw} (Mon {monday:%d %b})"


def block_week_for_date(d: date, block_start_date: date) -> int:
    """Convenience: block-relative week number directly from a date."""
    yr, wk = iso_week_key(d)
    return block_week_number(yr, wk, block_start_date)

def check_window_span(
    window_start: date | None, window_end: date | None, expected_days: int = 7, tolerance: int = 2
) -> str | None:
    """Return a warning string if a file's detected date window isn't close
    to one week, or None if it looks fine.

    A combined multi-week pull (e.g. someone narrows the Blackboard date
    picker to 'Week 0 to Week 1' instead of exporting each week
    separately) can't be split back into individual weeks after the fact
    — the export already aggregated the data before we ever see it. The
    best we can do is catch it loudly at upload time rather than let it
    silently misattribute everything to whichever week window_start falls
    into.
    """
    if window_start is None or window_end is None:
        return None
    span_days = (window_end - window_start).days + 1
    if abs(span_days - expected_days) > tolerance:
        return (
            f"date range is {span_days} days ({window_start} to {window_end}), "
            f"not ~{expected_days} — this looks like a multi-week export, not "
            f"a single narrowed week. Re-export with the date picker narrowed "
            f"to one week; a combined pull can't be split back into "
            f"individual weeks after the fact."
        )
    return None


# ---------------------------------------------------------------------
# Stacking weekly snapshot uploads into wide (student_code x week) tables
# ---------------------------------------------------------------------

def stack_weekly(
    snapshots: list[tuple[pd.DataFrame, date | None, date | None]],
    value_col: str,
    block_start_date: date,
) -> pd.DataFrame:
    """Combine several (df[student_code, value_col], window_start, window_end)
    uploads — one per week — into one wide table: rows = student_code,
    columns = block-relative week number, values = value_col.

    Snapshots with an undetectable window_start are skipped (with the
    caller expected to have already warned the user in the UI).
    """
    frames = []
    for df, window_start, _window_end in snapshots:
        if window_start is None or df is None or df.empty:
            continue
        bw = block_week_for_date(window_start, block_start_date)
        s = df.set_index("student_code")[value_col]
        s = s.groupby(level=0).sum()  # guard against dup student rows
        frames.append(s.rename(bw))
    if not frames:
        return pd.DataFrame()
    wide = pd.concat(frames, axis=1)
    # If two uploads land on the same block week (re-uploaded / duplicate),
    # keep the last one rather than silently summing.
    wide = wide.T.groupby(level=0).last().T
    wide = wide.fillna(0)
    return wide.sort_index(axis=1)


def build_login_tables(
    login_snapshots: list[tuple[pd.DataFrame, date | None, date | None]],
    block_start_date: date,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """From weekly Login Report uploads, build:

      cumulative_wide : student_code x block_week -> lifetime TOTAL LOGINS
                         at the time of that snapshot
      delta_wide       : student_code x block_week -> LOGIN COUNT during
                         that specific week (this week's cumulative minus
                         the previous available snapshot's cumulative).
                         The earliest snapshot in the batch has no prior
                         snapshot to diff against, so its delta is NaN
                         (excluded from segmentation, not treated as 0).
      latest_login_df  : the single login DataFrame from the
                         highest-block-week snapshot — used for the
                         lifetime last_login_date / total_logins / S1-S2
                         checks, since it's the most current picture.

    Students not present in a given week's file are treated as having no
    new logins that snapshot (their cumulative total simply doesn't
    advance) — forward-filled from their last known cumulative value so
    the delta for surrounding weeks isn't corrupted by a one-week gap.
    """
    frames = []
    latest_bw = None
    latest_df = None
    for df, window_start, _window_end in login_snapshots:
        if window_start is None or df is None or df.empty:
            continue
        bw = block_week_for_date(window_start, block_start_date)
        s = df.set_index("student_code")["total_logins"]
        s = s.groupby(level=0).max()
        frames.append(s.rename(bw))
        if latest_bw is None or bw > latest_bw:
            latest_bw = bw
            latest_df = df

    if not frames:
        empty = pd.DataFrame()
        return empty, empty, pd.DataFrame(columns=[
            "student_code", "surname", "first_name", "email",
            "days_since_last_login", "last_login_date", "total_logins",
        ])

    cumulative_wide = pd.concat(frames, axis=1)
    cumulative_wide = cumulative_wide.T.groupby(level=0).last().T
    cumulative_wide = cumulative_wide.sort_index(axis=1)
    cumulative_wide = cumulative_wide.ffill(axis=1)

    delta_wide = cumulative_wide.diff(axis=1)
    # Clamp any negative deltas (shouldn't happen, but a Blackboard export
    # glitch or a re-run with a shrunk window could produce one) to 0
    # rather than let them silently corrupt the fade/growth comparison.
    delta_wide = delta_wide.clip(lower=0)

    return cumulative_wide, delta_wide, latest_df


# ---------------------------------------------------------------------
# Per-student summary
# ---------------------------------------------------------------------

def weeks_in_data(*wide_tables: pd.DataFrame) -> list[int]:
    """Union of block-week columns present across one or more wide tables."""
    weeks: set[int] = set()
    for t in wide_tables:
        if t is not None and not t.empty:
            weeks |= set(int(c) for c in t.columns)
    return sorted(weeks)


def per_student_summary(
    class_list: pd.DataFrame,
    hours_wide: pd.DataFrame,
    logins_delta_wide: pd.DataFrame,
    forum_wide: pd.DataFrame,
    latest_login_df: pd.DataFrame,
    grade_summary: pd.DataFrame | None,
) -> pd.DataFrame:
    """Build the master per-student table, restricted to enrolled students."""
    desired_cols = [
        "student_code", "first_name", "last_name", "preferred_name",
        "attend_type", "course", "course_type", "email_address",
    ]
    available_cols = [c for c in desired_cols if c in class_list.columns]
    base = class_list[available_cols].copy()
    base["student_code"] = base["student_code"].astype(str).str.strip()

    # --- Lifetime login data (from the most recent weekly snapshot) ---
    login_keep = latest_login_df[[
        "student_code", "last_login_date", "total_logins", "days_since_last_login",
    ]].copy() if not latest_login_df.empty else pd.DataFrame(
        columns=["student_code", "last_login_date", "total_logins", "days_since_last_login"]
    )
    login_keep["in_login_report"] = True
    base = base.merge(login_keep, on="student_code", how="left")
    base["in_login_report"] = base["in_login_report"].fillna(False)
    base["total_logins"] = base["total_logins"].fillna(0).astype(int)

    # --- Weekly totals: hours, logins, forum interactions ---
    def _row_totals(wide: pd.DataFrame, col_name: str) -> pd.Series:
        if wide.empty:
            return pd.Series(0.0, index=base.index)
        totals = wide.sum(axis=1).rename(col_name)
        base_local = base.merge(totals, left_on="student_code", right_index=True, how="left")
        return base_local[col_name].fillna(0)

    base["total_hours"] = _row_totals(hours_wide, "total_hours")
    base["total_hours"] = base["total_hours"].round(2)

    if not logins_delta_wide.empty:
        login_totals = logins_delta_wide.sum(axis=1, skipna=True).rename("total_period_logins")
        base = base.merge(login_totals, left_on="student_code", right_index=True, how="left")
        base["total_period_logins"] = base["total_period_logins"].fillna(0).astype(int)
    else:
        base["total_period_logins"] = 0

    if not forum_wide.empty:
        forum_totals = forum_wide.sum(axis=1).rename("total_forum_interactions")
        base = base.merge(forum_totals, left_on="student_code", right_index=True, how="left")
        base["total_forum_interactions"] = base["total_forum_interactions"].fillna(0).astype(int)
    else:
        base["total_forum_interactions"] = 0

    # --- Weeks active / last active week (week-precision, not day-precision) ---
    weeks = weeks_in_data(hours_wide, logins_delta_wide)
    teaching_weeks = [w for w in weeks if w >= 1]

    def _active_mask_for_week(sid, wk) -> bool:
        h = hours_wide.at[sid, wk] if (not hours_wide.empty and sid in hours_wide.index and wk in hours_wide.columns) else 0
        l = logins_delta_wide.at[sid, wk] if (not logins_delta_wide.empty and sid in logins_delta_wide.index and wk in logins_delta_wide.columns) else 0
        h = 0 if pd.isna(h) else h
        l = 0 if pd.isna(l) else l
        return (h > 0) or (l > 0)

    weeks_active_list = []
    last_active_week_list = []
    for sid in base["student_code"]:
        active_weeks = [w for w in teaching_weeks if _active_mask_for_week(sid, w)]
        weeks_active_list.append(len(active_weeks))
        last_active_week_list.append(max(active_weeks) if active_weeks else pd.NA)
    base["weeks_active"] = weeks_active_list
    base["last_active_week"] = last_active_week_list

    this_week = max(teaching_weeks) if teaching_weeks else None
    if this_week is not None:
        base["weeks_since_last_active"] = base["last_active_week"].apply(
            lambda w: (this_week - w) if pd.notna(w) else pd.NA
        )
    else:
        base["weeks_since_last_active"] = pd.NA

    # --- Grade Centre ---
    if grade_summary is not None and not grade_summary.empty:
        base = base.merge(grade_summary, on="student_code", how="left")
        base["assessments_submitted"] = base["assessments_submitted"].fillna(0).astype(int)
        base["submission_rate"] = base["submission_rate"].fillna(0.0)
    else:
        base["assessments_submitted"] = 0
        base["assessments_total"] = 0
        base["submission_rate"] = 0.0
        base["avg_score_pct"] = pd.NA

    return base


def append_weekly_columns(
    summary: pd.DataFrame,
    hours_wide: pd.DataFrame,
    logins_delta_wide: pd.DataFrame,
    forum_wide: pd.DataFrame,
    weeks: list[int],
    block_start_date: date | None = None,
) -> pd.DataFrame:
    """Append per-week hours / logins / forum-interaction columns."""
    out = summary.copy()
    for wk in weeks:
        label_prefix = f"W{wk}"
        for wide, suffix in (
            (hours_wide, "hours"), (logins_delta_wide, "logins"), (forum_wide, "forum"),
        ):
            label = f"{label_prefix}_{suffix}"
            if not wide.empty and wk in wide.columns:
                col = wide[wk].rename(label)
                out = out.merge(col, left_on="student_code", right_index=True, how="left")
            else:
                out[label] = pd.NA

            if suffix == "logins" and wk == 0:
                # Week 0 is the pre-teaching baseline snapshot — there's no
                # prior week to diff against, so its "login count" is
                # genuinely undefined, not zero. Leave as NaN rather than
                # implying the student logged in 0 times that week.
                continue
            if suffix == "hours":
                out[label] = out[label].fillna(0).round(2)
            else:
                out[label] = out[label].fillna(0).astype("Int64")
    return out
