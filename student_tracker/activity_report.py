"""Excel export for the activity-log Student Engagement Tracker."""

from __future__ import annotations

import io

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="305496")
SEGMENT_FILLS = {
    "Engaged": PatternFill("solid", fgColor="A9D08E"),
    "Fading": PatternFill("solid", fgColor="FFD966"),
    "Lapsed": PatternFill("solid", fgColor="F4B084"),
    "Never visited": PatternFill("solid", fgColor="E06666"),
    "Not on class list": PatternFill("solid", fgColor="D9D9D9"),
}
STAGE_FILLS = {
    "Submitted": PatternFill("solid", fgColor="A9D08E"),
    "Opened submission point": PatternFill("solid", fgColor="FFD966"),
    "Read About page": PatternFill("solid", fgColor="F4B084"),
    "Opened other items only": PatternFill("solid", fgColor="F4B084"),
    "Not opened": PatternFill("solid", fgColor="E06666"),
}


def _style(ws, df: pd.DataFrame) -> None:
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    for i, col in enumerate(df.columns, start=1):
        vals = [str(v) for v in df[col].head(200).tolist()]
        width = min(max([len(str(col))] + [len(v) for v in vals]) + 2, 60)
        ws.column_dimensions[get_column_letter(i)].width = width
    fills = {**SEGMENT_FILLS, **STAGE_FILLS}
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, str) and cell.value in fills:
                cell.fill = fills[cell.value]


def build_workbook(sheets: dict[str, pd.DataFrame]) -> bytes:
    """Write each DataFrame to its own sheet (name <= 31 chars) and style it."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for name, df in sheets.items():
            out = df.copy()
            for c in out.columns:
                if isinstance(out[c].dtype, pd.CategoricalDtype):
                    out[c] = out[c].astype(str)
                elif pd.api.types.is_datetime64_any_dtype(out[c]):
                    out[c] = out[c].dt.tz_localize(None) if out[c].dt.tz else out[c]
            out.to_excel(xw, sheet_name=name[:31], index=False)
            _style(xw.sheets[name[:31]], out)
    return buf.getvalue()
