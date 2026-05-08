"""Classify each student into one of the engagement segments.

Segment definitions (as agreed with the user):
  S1 - Never engaged: 0 lifetime logins (also catches students missing
       from the login report entirely)
  S2 - Pre-block ghost: logged in before block start, no login since
  S3 - W1 ghost: had hits in block-week-1, none in any later week in data
  S4 - Dropped this week: hits last week, none this week
  S5 - Returning engager: no hits last week, hits this week
  S6 - Fading engager: hits both weeks, this week < 50% of last week
  S7 - True sustainer: hits both weeks, this week within 50% of last week
       (this includes increases — a student more active than last week is
       a sustainer, not a separate 'surging' category, by current definition)
  S8 - Long-tail dropout: had hits at some point during the data window,
       but zero in either of the last two weeks. Effectively only
       meaningful for long blocks (e.g. 17-week prep subjects); will be
       empty or near-empty for short discipline blocks.

Hits drive the comparison (per user preference); active days are reported
as supplementary context but do not affect classification.

Edge cases:
  - If fewer than 2 weeks of data, S4–S8 cannot be evaluated. Such students
    are placed in 'Active (single week)' or 'Unclassified — no activity'.
  - Students with login activity but zero hits anywhere in the data window
    fall through to 'Unclassified' if they don't match S1/S2/S3.
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


def _safe_get(weekly_hits: pd.DataFrame, student_code: str,
              week_key: tuple[int, int] | None) -> int:
    if week_key is None or weekly_hits.empty:
        return 0
    if student_code not in weekly_hits.index:
        return 0
    if week_key not in weekly_hits.columns:
        return 0
    val = weekly_hits.loc[student_code, week_key]
    try:
        return int(val)
    except (ValueError, TypeError):
        return 0


def classify(
    summary: pd.DataFrame,
    weekly_hits: pd.DataFrame,
    block_start_date: date,
    weeks: list[tuple[int, int]],
    fade_threshold: float = 0.5,
) -> pd.DataFrame:
    """Add a 'segment' column to the per-student summary DataFrame.

    Args:
        summary: per-student summary with student_code, total_logins,
                 last_login_date columns.
        weekly_hits: wide table from metrics.weekly_hits_table.
        block_start_date: first day of the teaching block.
        weeks: sorted list of (iso_year, iso_week) tuples present in data.
        fade_threshold: ratio below which 'this week / last week' counts as
                        fading. Default 0.5 (i.e., >50% drop = S6).

    Returns:
        A copy of summary with an added 'segment' column.
    """
    this_week = weeks[-1] if weeks else None
    last_week = weeks[-2] if len(weeks) >= 2 else None

    bsd_cal = block_start_date.isocalendar()
    block_w1_key = (bsd_cal[0], bsd_cal[1])
    if block_w1_key not in weeks:
        block_w1_key = None

    later_than_w1 = [w for w in weeks if block_w1_key and w > block_w1_key]

    segments: list[str] = []
    for _, row in summary.iterrows():
        sc = row["student_code"]
        total_logins = int(row.get("total_logins", 0) or 0)
        last_login = row.get("last_login_date")

        # S1 — never logged in (also covers students missing from the login
        # report entirely; they get total_logins=0 / last_login=NaN after merge).
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

        h_this = _safe_get(weekly_hits, sc, this_week)
        h_last = _safe_get(weekly_hits, sc, last_week)
        h_w1 = _safe_get(weekly_hits, sc, block_w1_key)

        # S3 — only meaningful if there are weeks after W1 in the data
        if h_w1 > 0 and later_than_w1:
            later_total = sum(_safe_get(weekly_hits, sc, w) for w in later_than_w1)
            if later_total == 0:
                segments.append(SEG_S3)
                continue

        # Need at least 2 weeks for S4–S7
        if last_week is None:
            if h_this > 0:
                segments.append(SEG_SINGLE_WEEK)
            else:
                segments.append(SEG_UNCLASSIFIED)
            continue

        # S4
        if h_last > 0 and h_this == 0:
            segments.append(SEG_S4)
            continue

        # S5
        if h_last == 0 and h_this > 0:
            segments.append(SEG_S5)
            continue

        # S6 / S7
        if h_last > 0 and h_this > 0:
            if h_this < fade_threshold * h_last:
                segments.append(SEG_S6)
            else:
                segments.append(SEG_S7)
            continue

        # S8 — Long-tail dropout: had activity sometime in the data window,
        # but zero in both of the last two weeks. Mainly relevant for long
        # blocks (e.g. 17-week prep subjects).
        total_period_hits = int(row.get("total_hits", 0) or 0)
        if total_period_hits > 0:
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
