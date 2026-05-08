"""Excel export of the per-student tracking report."""

from datetime import date
from io import BytesIO
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

from student_tracker.segmentation import segment_counts, ALL_SEGMENTS

HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="305496")
SEGMENT_FILLS = {
    "S1 - Never engaged":          PatternFill("solid", fgColor="C00000"),
    "S2 - Pre-block ghost":        PatternFill("solid", fgColor="E26B0A"),
    "S3 - W1 ghost":               PatternFill("solid", fgColor="FFC000"),
    "S4 - Dropped this week":      PatternFill("solid", fgColor="FFD966"),
    "S5 - Returning engager":      PatternFill("solid", fgColor="A9D08E"),
    "S6 - Fading engager":         PatternFill("solid", fgColor="F4B084"),
    "S7 - True sustainer":         PatternFill("solid", fgColor="70AD47"),
    "S8 - Long-tail dropout":      PatternFill("solid", fgColor="C65911"),
    "Active (single week of data)": PatternFill("solid", fgColor="BDD7EE"),
    "Unclassified":                PatternFill("solid", fgColor="D9D9D9"),
}


def _write_df(ws, df: pd.DataFrame, start_row: int = 1) -> None:
    """Write a DataFrame to a worksheet starting at start_row, with header styling."""
    # Convert pandas NA / NaT to None for openpyxl compatibility
    df = df.copy()
    df = df.astype(object).where(df.notna(), None)
    rows = list(dataframe_to_rows(df, index=False, header=True))
    for r_idx, row in enumerate(rows, start=start_row):
        for c_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            if r_idx == start_row:
                cell.font = HEADER_FONT
                cell.fill = HEADER_FILL
                cell.alignment = Alignment(horizontal="left", vertical="center")
    # Auto-width columns
    for c_idx in range(1, len(df.columns) + 1):
        col_vals = ["" if v is None else str(v) for v in df.iloc[:, c_idx - 1].tolist()[:200]]
        max_len = max([len(str(df.columns[c_idx - 1]))] + [len(v) for v in col_vals])
        ws.column_dimensions[get_column_letter(c_idx)].width = min(max_len + 2, 40)


def build_workbook(
    classified: pd.DataFrame,
    weeks: list[tuple[int, int]],
    block_start_date: date,
    subject_code: str | None,
    reference_date: date,
    half_life_days: float,
) -> Workbook:
    """Build the Excel workbook for download."""
    wb = Workbook()

    # === Summary sheet ===
    ws = wb.active
    ws.title = "Summary"
    ws["A1"] = "Student Engagement Tracking Report"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f"Subject: {subject_code or 'Unknown'}"
    ws["A3"] = f"Block start: {block_start_date.isoformat()}"
    ws["A4"] = f"Reference date: {reference_date.isoformat()}"
    ws["A5"] = f"Recency half-life (days): {half_life_days}"
    ws["A6"] = f"Weeks in data: {', '.join(f'W{w:02d}' for _, w in weeks)}"
    ws["A7"] = f"Total enrolled students: {len(classified)}"

    counts = segment_counts(classified)
    _write_df(ws, counts, start_row=10)
    # Colour the segment cells in the counts table
    for row_idx, seg_name in enumerate(counts["segment"], start=11):
        fill = SEGMENT_FILLS.get(seg_name)
        if fill:
            ws.cell(row=row_idx, column=1).fill = fill

    # === Per-student sheet ===
    ws2 = wb.create_sheet("Per Student")
    # Order columns sensibly
    front_cols = [
        "segment", "student_code", "preferred_name", "first_name", "last_name",
        "attend_type", "course",
        "last_hit_date", "days_since_last_hit",
        "prior_week_daily_avg", "this_week_daily_avg",
        "total_logins", "last_login_date", "days_since_last_login",
        "total_hits", "total_active_days",
        "weighted_hits", "weighted_active_days",
        "assessments_submitted", "assessments_total", "submission_rate", "avg_score_pct",
        "email_address",
    ]
    week_cols = [c for c in classified.columns if c.startswith("W") and ("_hits" in c or "_active_days" in c)]
    other_cols = [c for c in classified.columns if c not in front_cols and c not in week_cols]
    ordered = [c for c in front_cols if c in classified.columns] + week_cols + other_cols
    out_df = classified[ordered].copy()
    # Format dates as ISO strings for Excel compatibility
    if "last_login_date" in out_df.columns:
        out_df["last_login_date"] = out_df["last_login_date"].dt.strftime("%Y-%m-%d").fillna("")
    if "last_hit_date" in out_df.columns:
        out_df["last_hit_date"] = out_df["last_hit_date"].dt.strftime("%Y-%m-%d").fillna("")
    if "submission_rate" in out_df.columns:
        out_df["submission_rate"] = (out_df["submission_rate"] * 100).round(1)
        out_df = out_df.rename(columns={"submission_rate": "submission_rate_pct"})

    _write_df(ws2, out_df, start_row=1)
    # Colour segment column
    seg_col_idx = ordered.index("segment") + 1
    for r_idx, seg in enumerate(out_df["segment"], start=2):
        fill = SEGMENT_FILLS.get(seg)
        if fill:
            ws2.cell(row=r_idx, column=seg_col_idx).fill = fill

    # === Per-segment sheets ===
    for seg in ALL_SEGMENTS:
        sub = classified[classified["segment"] == seg]
        if sub.empty:
            continue
        sheet_name = seg.split(" - ")[0][:31]  # Excel sheet name limit
        if sheet_name in [s.title for s in wb.worksheets]:
            sheet_name = sheet_name + "_"
        ws_s = wb.create_sheet(sheet_name)
        sub_ordered = sub[[c for c in ordered if c in sub.columns]].copy()
        if "last_login_date" in sub_ordered.columns:
            sub_ordered["last_login_date"] = (
                sub_ordered["last_login_date"].dt.strftime("%Y-%m-%d").fillna("")
            )
        if "last_hit_date" in sub_ordered.columns:
            sub_ordered["last_hit_date"] = (
                sub_ordered["last_hit_date"].dt.strftime("%Y-%m-%d").fillna("")
            )
        if "submission_rate" in sub_ordered.columns:
            sub_ordered["submission_rate"] = (sub_ordered["submission_rate"] * 100).round(1)
            sub_ordered = sub_ordered.rename(columns={"submission_rate": "submission_rate_pct"})
        _write_df(ws_s, sub_ordered, start_row=1)

    return wb


def to_bytes(wb: Workbook) -> bytes:
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
