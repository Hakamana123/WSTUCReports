"""Per-student metric computations.

All metrics derive from:
  - The Overall report's date section (per-student per-date hits)
  - The Login report (last login date, total logins lifetime)
  - The Grade Centre (submission and score)

Weeks are ISO weeks, Monday–Sunday.
Recency weighting uses exponential decay with a configurable half-life.
"""

from datetime import date, timedelta
import math
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


def block_week_colname(iso_year: int, iso_week: int, block_start_date: date,
                       suffix: str) -> str:
    """Block-relative column name like 'W1_hits' or 'W0_active_days'."""
    bw = block_week_number(iso_year, iso_week, block_start_date)
    return f"W{bw}_{suffix}"


def weeks_in_data(date_hits: pd.DataFrame) -> list[tuple[int, int]]:
    """Return sorted list of (iso_year, iso_week) tuples present in the data."""
    if date_hits.empty:
        return []
    iso = date_hits["date"].dt.isocalendar()
    pairs = list(zip(iso["year"].astype(int), iso["week"].astype(int)))
    return sorted(set(pairs))


def weekly_hits_table(date_hits: pd.DataFrame) -> pd.DataFrame:
    """Wide table: rows = student_code, columns = (iso_year, iso_week), values = total hits."""
    if date_hits.empty:
        return pd.DataFrame()
    df = date_hits.copy()
    iso = df["date"].dt.isocalendar()
    df["iso_year"] = iso["year"].astype(int)
    df["iso_week"] = iso["week"].astype(int)
    pivot = df.groupby(["student_code", "iso_year", "iso_week"], as_index=False)["hits"].sum()
    wide = pivot.pivot_table(
        index="student_code",
        columns=["iso_year", "iso_week"],
        values="hits",
        fill_value=0,
    )
    return wide


def weekly_active_days_table(date_hits: pd.DataFrame) -> pd.DataFrame:
    """Wide table: rows = student_code, cols = (iso_year, iso_week), values = active days."""
    if date_hits.empty:
        return pd.DataFrame()
    df = date_hits.copy()
    df = df[df["hits"] > 0]
    iso = df["date"].dt.isocalendar()
    df["iso_year"] = iso["year"].astype(int)
    df["iso_week"] = iso["week"].astype(int)
    counts = df.groupby(["student_code", "iso_year", "iso_week"])["date"].nunique().reset_index()
    counts = counts.rename(columns={"date": "active_days"})
    wide = counts.pivot_table(
        index="student_code",
        columns=["iso_year", "iso_week"],
        values="active_days",
        fill_value=0,
    )
    return wide


def recency_weighted(date_hits: pd.DataFrame, reference_date: date,
                     half_life_days: float = 7.0) -> pd.DataFrame:
    """Per-student recency-weighted hits and active-days.

    Weight for day X = 0.5 ** ((reference_date - X).days / half_life_days).
    Returns DataFrame with columns: student_code, weighted_hits, weighted_active_days.
    """
    if date_hits.empty:
        return pd.DataFrame(columns=["student_code", "weighted_hits", "weighted_active_days"])

    df = date_hits.copy()
    df["days_ago"] = (pd.Timestamp(reference_date) - df["date"]).dt.days
    df["weight"] = 0.5 ** (df["days_ago"] / half_life_days)
    # Clamp negative days_ago (future dates) to weight 1.0
    df.loc[df["days_ago"] < 0, "weight"] = 1.0

    df["weighted_hits"] = df["hits"] * df["weight"]
    df["was_active"] = (df["hits"] > 0).astype(int)
    df["weighted_active_days"] = df["was_active"] * df["weight"]

    out = df.groupby("student_code").agg(
        weighted_hits=("weighted_hits", "sum"),
        weighted_active_days=("weighted_active_days", "sum"),
    ).reset_index()
    return out


def per_student_summary(
    class_list: pd.DataFrame,
    date_hits: pd.DataFrame,
    login: pd.DataFrame,
    grade_summary: pd.DataFrame,
    reference_date: date,
    half_life_days: float = 7.0,
) -> pd.DataFrame:
    """Build the master per-student table, restricted to enrolled students.

    Returns DataFrame keyed by student_code with all per-student metrics
    needed for the dashboard and the segmentation step.
    """
    desired_cols = [
        "student_code", "first_name", "last_name", "preferred_name",
        "attend_type", "course", "course_type", "email_address",
    ]
    available_cols = [c for c in desired_cols if c in class_list.columns]
    base = class_list[available_cols].copy()
    base["student_code"] = base["student_code"].astype(str).str.strip()

    # --- Login data (lifetime) ---
    login_keep = login[[
        "student_code", "last_login_date", "total_logins", "days_since_last_login",
    ]].copy()
    login_keep["in_login_report"] = True
    base = base.merge(login_keep, on="student_code", how="left")
    base["in_login_report"] = base["in_login_report"].fillna(False)
    base["total_logins"] = base["total_logins"].fillna(0).astype(int)

    # --- Total / weekly hits ---
    if not date_hits.empty:
        totals = date_hits.groupby("student_code")["hits"].sum().rename("total_hits")
        base = base.merge(totals, on="student_code", how="left")
        active = date_hits[date_hits["hits"] > 0].groupby("student_code")["date"].nunique().rename("total_active_days")
        base = base.merge(active, on="student_code", how="left")
        # Last hit date — derived from the Overall report. More reliable than
        # the Login Report's last_login_date, which can be stale.
        last_hit = date_hits[date_hits["hits"] > 0].groupby("student_code")["date"].max().rename("last_hit_date")
        base = base.merge(last_hit, on="student_code", how="left")
    else:
        base["total_hits"] = 0
        base["total_active_days"] = 0
        base["last_hit_date"] = pd.NaT
    base["total_hits"] = base["total_hits"].fillna(0).astype(int)
    base["total_active_days"] = base["total_active_days"].fillna(0).astype(int)
    base["last_hit_date"] = pd.to_datetime(base["last_hit_date"])
    # days_since_last_hit relative to the reference date
    base["days_since_last_hit"] = (
        pd.Timestamp(reference_date) - base["last_hit_date"]
    ).dt.days

    # --- Recency-weighted ---
    rw = recency_weighted(date_hits, reference_date, half_life_days)
    base = base.merge(rw, on="student_code", how="left")
    base["weighted_hits"] = base["weighted_hits"].fillna(0.0).round(2)
    base["weighted_active_days"] = base["weighted_active_days"].fillna(0.0).round(2)

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


def append_recent_week_averages(summary: pd.DataFrame,
                                 weekly_hits: pd.DataFrame,
                                 weeks: list[tuple[int, int]],
                                 reference_date: date) -> pd.DataFrame:
    """Add prior_week_daily_avg and this_week_daily_avg to the summary.

    Helps interpret S4 (dropped) and S6 (fading) by showing the baseline
    pace each student was at before the current week.

    'this_week' = most recent ISO week with hits in the data.
    'prior_week' = the ISO week before that.
    Daily average is hits divided by 7 (full ISO week). For the current
    week, if it's incomplete relative to reference_date, we divide by the
    number of days elapsed within that week (1–7) so the average isn't
    artificially low.
    """
    out = summary.copy()
    this_week = weeks[-1] if weeks else None
    prior_week = weeks[-2] if len(weeks) >= 2 else None

    def daily_avg_for(week_key, denom):
        if week_key is None or weekly_hits.empty or week_key not in weekly_hits.columns:
            return pd.Series(0.0, index=out.index)
        col = weekly_hits[week_key]
        merged = out["student_code"].map(col).fillna(0)
        return (merged / denom).round(2)

    # Prior week: always 7 days
    out["prior_week_hits"] = out["student_code"].map(
        weekly_hits[prior_week] if (prior_week and not weekly_hits.empty
                                     and prior_week in weekly_hits.columns)
        else pd.Series(dtype=int)
    ).fillna(0).astype(int) if prior_week else 0
    out["prior_week_daily_avg"] = (out["prior_week_hits"] / 7.0).round(2)

    # Current week: divide by days elapsed within the week
    if this_week is not None:
        this_monday = date.fromisocalendar(this_week[0], this_week[1], 1)
        days_elapsed = max(1, min(7, (reference_date - this_monday).days + 1))
        out["this_week_hits"] = out["student_code"].map(
            weekly_hits[this_week] if (not weekly_hits.empty
                                        and this_week in weekly_hits.columns)
            else pd.Series(dtype=int)
        ).fillna(0).astype(int)
        out["this_week_daily_avg"] = (out["this_week_hits"] / days_elapsed).round(2)
    else:
        out["this_week_hits"] = 0
        out["this_week_daily_avg"] = 0.0

    return out


def append_weekly_columns(summary: pd.DataFrame,
                          weekly_hits: pd.DataFrame,
                          weekly_active: pd.DataFrame,
                          weeks: list[tuple[int, int]],
                          block_start_date: date | None = None) -> pd.DataFrame:
    """Append per-week hits and active-days columns to the per-student summary.

    Column names use block-relative week numbers (W1 = block start week,
    W0 = orientation, W-1 = earlier) when block_start_date is supplied;
    otherwise fall back to ISO week numbering (W{NN}).
    """
    out = summary.copy()
    for (yr, wk) in weeks:
        if block_start_date is not None:
            label = block_week_colname(yr, wk, block_start_date, "hits")
            label2 = block_week_colname(yr, wk, block_start_date, "active_days")
        else:
            label = f"W{wk:02d}_hits"
            label2 = f"W{wk:02d}_active_days"

        if not weekly_hits.empty and (yr, wk) in weekly_hits.columns:
            col = weekly_hits[(yr, wk)].rename(label)
            out = out.merge(col, left_on="student_code", right_index=True, how="left")
        else:
            out[label] = 0
        out[label] = out[label].fillna(0).astype(int)

        if not weekly_active.empty and (yr, wk) in weekly_active.columns:
            col2 = weekly_active[(yr, wk)].rename(label2)
            out = out.merge(col2, left_on="student_code", right_index=True, how="left")
        else:
            out[label2] = 0
        out[label2] = out[label2].fillna(0).astype(int)

    return out
