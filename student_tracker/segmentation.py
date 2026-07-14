"""Classify each student into one of the engagement segments.

Rebuilt for the two-signal (hours + login count) weekly model — Option 1
from the design discussion: hours and login count are each checked
independently for "active this week", and the fade/growth ratio (S6/S7)
is measured on whichever of the two was the larger signal for that
student last week. Forum interaction data does NOT feed classification;
it's supplementary context shown alongside the segment, because forum
activity is task-dependent and bursty (spikes when a discussion board
task is due, zero otherwise) in a way that would corrupt week-over-week
fade comparisons if blended in.

Segment definitions:
  S1 - Never engaged: 0 lifetime logins (also catches students missing
       from the login report entirely)
  S2 - Pre-block ghost: logged in before block start, no login since
  S3 - W1 ghost: active (hours or logins) in block-week-1, inactive in
       every later week present in the data
  S4 - Dropped this week: active last week, inactive this week
  S5 - Returning engager: inactive last week, active this week
  S6 - Fading engager: active both weeks, this week's dominant-signal
       value < fade_threshold * last week's dominant-signal value
  S7 - True sustainer: active both weeks, not fading (includes
       increases — a student more active than last week is a sustainer,
       not a separate 'surging' category, by current definition)
  S8 - Long-tail dropout: had activity (hours or logins) at some point
       across the uploaded weeks, but inactive in both of the last two
       weeks present. Mainly relevant for long blocks; usually empty for
       short discipline blocks.

'Active' in a given week = hours > 0 OR login count > 0.

Edge cases:
  - If fewer than 2 teaching weeks (block week >= 1) are in the uploaded
    data, S4-S8 cannot be evaluated. Such students are placed in
    'Active (single week)' or 'Unclassified — no activity'.
  - Week 0 (pre-teaching baseline) is never itself classified against —
    it exists only as the login-delta baseline.
"""

from datetime import date
import pandas as pd

SEG_S1 = "S1 - Never engaged"
SEG_S2 = "S2 - Pre-block ghost"
SEG_S3 = "S3 - W1 ghost"
SEG_S4 = "S4 - Dropped this week"
SEG_S5 = "S5 - Returning engager"
SEG_S6 = "S6 - Fading engager"
SEG_S7 = "S7 - True sustainer"
SEG_S8 = "S8 - Long-tail dropout"
SEG_SINGLE_WEEK = "Active (single week of data)"
SEG_UNCLASSIFIED = "Unclassified"

ALL_SEGMENTS = [
    SEG_S1, SEG_S2, SEG_S3, SEG_S4, SEG_S5, SEG_S6, SEG_S7, SEG_S8,
    SEG_SINGLE_WEEK, SEG_UNCLASSIFIED,
]


def _safe_get(wide: pd.DataFrame, student_code: str, week: int | None) -> float:
    if week is None or wide.empty:
        return 0.0
    if student_code not in wide.index or week not in wide.columns:
        return 0.0
    val = wide.at[student_code, week]
    if pd.isna(val):
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def _is_active(hours_wide, logins_wide, sid, week) -> bool:
    return (_safe_get(hours_wide, sid, week) > 0) or (_safe_get(logins_wide, sid, week) > 0)


def _dominant_metric_ratio(hours_wide, logins_wide, sid, this_week, last_week) -> float | None:
    """Ratio (this_week / last_week) on whichever metric was larger last week.

    Returns None if the dominant metric's last-week value is 0 (can't form
    a ratio — caller should treat that as 'not fading', since S6/S7 is only
    reached when we already know both weeks are active on the OR signal).
    """
    h_last = _safe_get(hours_wide, sid, last_week)
    l_last = _safe_get(logins_wide, sid, last_week)
    if h_last >= l_last:
        last_val, this_val = h_last, _safe_get(hours_wide, sid, this_week)
    else:
        last_val, this_val = l_last, _safe_get(logins_wide, sid, this_week)
    if last_val <= 0:
        return None
    return this_val / last_val


def classify(
    summary: pd.DataFrame,
    hours_wide: pd.DataFrame,
    logins_wide: pd.DataFrame,
    block_start_date: date,
    weeks: list[int],
    fade_threshold: float = 0.5,
) -> pd.DataFrame:
    """Add a 'segment' column to the per-student summary DataFrame.

    Args:
        summary: per-student summary with student_code, total_logins,
                 last_login_date columns (lifetime, from latest snapshot).
        hours_wide, logins_wide: wide tables (student_code index, block-week
                 columns) from metrics.stack_weekly / build_login_tables.
        block_start_date: first day of the teaching block.
        weeks: sorted list of block-relative week numbers present in data
               (may include 0, the pre-teaching baseline).
        fade_threshold: ratio below which 'this week / last week' on the
                        dominant metric counts as fading. Default 0.5.
    """
    teaching_weeks = [w for w in weeks if w >= 1]
    this_week = teaching_weeks[-1] if teaching_weeks else None
    last_week = teaching_weeks[-2] if len(teaching_weeks) >= 2 else None
    w1_key = 1 if 1 in teaching_weeks else None
    later_than_w1 = [w for w in teaching_weeks if w1_key and w > w1_key]

    segments: list[str] = []
    for _, row in summary.iterrows():
        sc = row["student_code"]
        total_logins = int(row.get("total_logins", 0) or 0)
        last_login = row.get("last_login_date")

        # S1 — never logged in (also covers students missing from the login
        # report entirely).
        if total_logins == 0 or pd.isna(last_login):
            segments.append(SEG_S1)
            continue

        # S2
        last_login_date = (
            last_login.date() if isinstance(last_login, pd.Timestamp) else last_login
        )
        if last_login_date < block_start_date:
            segments.append(SEG_S2)
            continue

        # S3 — only meaningful if there are weeks after W1 in the data
        if w1_key and _is_active(hours_wide, logins_wide, sc, w1_key) and later_than_w1:
            if not any(_is_active(hours_wide, logins_wide, sc, w) for w in later_than_w1):
                segments.append(SEG_S3)
                continue

        # Need at least 2 teaching weeks for S4-S7
        if last_week is None:
            if this_week and _is_active(hours_wide, logins_wide, sc, this_week):
                segments.append(SEG_SINGLE_WEEK)
            else:
                segments.append(SEG_UNCLASSIFIED)
            continue

        active_last = _is_active(hours_wide, logins_wide, sc, last_week)
        active_this = _is_active(hours_wide, logins_wide, sc, this_week)

        # S4
        if active_last and not active_this:
            segments.append(SEG_S4)
            continue

        # S5
        if not active_last and active_this:
            segments.append(SEG_S5)
            continue

        # S6 / S7
        if active_last and active_this:
            ratio = _dominant_metric_ratio(hours_wide, logins_wide, sc, this_week, last_week)
            if ratio is not None and ratio < fade_threshold:
                segments.append(SEG_S6)
            else:
                segments.append(SEG_S7)
            continue

        # S8 — Long-tail dropout: active at some point during TEACHING
        # weeks, inactive in both of the last two weeks. Deliberately
        # excludes Week 0 (pre-teaching baseline) activity — a student
        # who only showed up before teaching started and never during it
        # is not a "dropout from teaching", and total_hours /
        # total_period_logins would wrongly include their W0 numbers.
        if any(_is_active(hours_wide, logins_wide, sc, w) for w in teaching_weeks):
            segments.append(SEG_S8)
            continue

        segments.append(SEG_UNCLASSIFIED)

    out = summary.copy()
    out["segment"] = segments
    return out


def segment_counts(classified: pd.DataFrame) -> pd.DataFrame:
    """Return a counts table of students per segment, in canonical order."""
    counts = classified["segment"].value_counts()
    rows = []
    for seg in ALL_SEGMENTS:
        rows.append({"segment": seg, "count": int(counts.get(seg, 0))})
    out = pd.DataFrame(rows)
    out["pct"] = (out["count"] / out["count"].sum() * 100).round(1)
    return out
