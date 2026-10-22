import io
import json
import re
import zipfile
from collections import Counter, defaultdict
from copy import copy
from pathlib import Path

import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.hyperlink import Hyperlink

st.set_page_config(page_title="File Splitter", layout="wide")

st.title("File Splitter")
st.caption(
    "Upload a finished Class Report (from the Engagement Report builder) and split it into "
    "one workbook per program — each containing only that program's class sheets, ready to "
    "send straight to the program coordinator."
)

HYPERLINK_SHEET_RE = re.compile(r"^#'?(.+?)'?!")
INDEX_COLUMNS = ["#", "Class", "Teacher", "Enrolled", "Active", "At Risk", "AS1 Sub", "AS2 Sub", "AS3 Sub"]
SUBJECT_INDEX_COLUMNS = ["#", "Subject / Class", "Teacher", "Enrolled", "Active", "At Risk", "AS1 Sub", "AS2 Sub", "AS3 Sub"]
SEG_ORDER = {f"S{i}": i for i in range(1, 8)}
STRIPE_COLOUR = "F7F9FC"  # ALT_ROW in the Engagement Report builder
SUBJECT_ROW_COLOUR = "E5ECF6"
CLASS_ROW_FONT_COLOUR = "808080"
WEEK_NOTE_RE = re.compile(r"(W-?\d+(?:\(partial\))?)\.\s*$")
REF_PROGRAMS_PATH = Path(__file__).resolve().parent.parent / "student_tracker" / "rereg_data" / "ref_programs.json"
FRACTION_RE = re.compile(r"^\s*(\d+)\s*/\s*(\d+)\s*$")


def _short_report_name(base_name):
    """Short prefix for output files, keeping names well under Windows' path limit
    once unzipped: 'GEDU0016_Class_Report_W5_20261020_Partial' -> 'GEDU0016_W5'."""
    m = re.match(r"^([A-Za-z0-9]+)_Class_Report_(W-?\d+)", base_name)
    if m:
        return f"{m.group(1)}_{m.group(2)}"
    return re.sub(r"[^A-Za-z0-9_-]+", "_", base_name)[:20].strip("_") or "Class_Report"


def _resolve_hyperlink_sheet(cell):
    if not cell.hyperlink or not cell.hyperlink.target:
        return None
    m = HYPERLINK_SHEET_RE.match(cell.hyperlink.target)
    return m.group(1) if m else None


def parse_class_index(wb):
    """Read the Class Index sheet and return one entry per class row.

    Returns None if the workbook has no Class Index sheet with a Program column
    (i.e. not a class report produced by the Engagement Report builder).
    """
    if "Class Index" not in wb.sheetnames:
        return None
    ws = wb["Class Index"]

    header_row = None
    for r in range(1, min(15, ws.max_row) + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if vals and vals[0] == "#" and "Program" in vals:
            header_row = r
            break
    if header_row is None:
        return None

    entries = []
    for r in range(header_row + 1, ws.max_row + 1):
        num_val = ws.cell(row=r, column=1).value
        if not isinstance(num_val, (int, float)):
            continue  # group header / blank row, not a class row
        class_cell = ws.cell(row=r, column=2)
        if class_cell.value == "No match":
            continue  # unmatched students, not a real class — never split into a program file
        sheet_name = _resolve_hyperlink_sheet(class_cell) or str(class_cell.value)
        entries.append(
            {
                "sheet": sheet_name,
                "class_label": class_cell.value,
                "program": "" if ws.cell(row=r, column=3).value is None else str(ws.cell(row=r, column=3).value),
                "teacher": ws.cell(row=r, column=4).value,
                "enrolled": ws.cell(row=r, column=5).value,
                "active": ws.cell(row=r, column=6).value,
                "at_risk": ws.cell(row=r, column=7).value,
                "as1_sub": ws.cell(row=r, column=8).value,
                "as2_sub": ws.cell(row=r, column=9).value,
                "as3_sub": ws.cell(row=r, column=10).value,
            }
        )
    return entries


def _write_program_index(wb, program, prog_entries):
    """Build a fresh 'Class Index' sheet listing just this program's classes.

    Built from scratch rather than derived by deleting rows from the original:
    openpyxl's delete_rows does not reliably carry hyperlinks along with the
    rows it shifts, which corrupts a sheet with as many rows as Class Index
    has. A freshly-written sheet has no such risk. It keeps the name
    "Class Index" so each retained class sheet's "Back to Index" link
    (which points at '#'Class Index'!A1') keeps working.
    """
    ws = wb.create_sheet("Class Index", 0)

    ws.cell(row=1, column=1, value=f"Program {program} — Class Index").font = Font(bold=True, size=14)
    ws.cell(
        row=2,
        column=1,
        value=f"{len(prog_entries)} class(es)  •  Click a class name to jump to its sheet",
    ).font = Font(italic=True, color="666666")

    header_row = 4
    header_fill = PatternFill("solid", start_color="1F2937", end_color="1F2937")
    for col, title in enumerate(INDEX_COLUMNS, start=1):
        cell = ws.cell(row=header_row, column=col, value=title)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill

    for i, e in enumerate(sorted(prog_entries, key=lambda e: e["class_label"] or ""), start=1):
        r = header_row + i
        ws.cell(row=r, column=1, value=i)
        class_cell = ws.cell(row=r, column=2, value=e["class_label"])
        if e["sheet"] in wb.sheetnames:
            class_cell.hyperlink = Hyperlink(ref=class_cell.coordinate, location=f"'{e['sheet']}'!A1")
            class_cell.font = Font(color="0563C1", underline="single")
        ws.cell(row=r, column=3, value=e["teacher"])
        ws.cell(row=r, column=4, value=e["enrolled"])
        ws.cell(row=r, column=5, value=e["active"])
        ws.cell(row=r, column=6, value=e["at_risk"])
        ws.cell(row=r, column=7, value=e["as1_sub"])
        ws.cell(row=r, column=8, value=e["as2_sub"])
        ws.cell(row=r, column=9, value=e["as3_sub"])

    ws.column_dimensions["B"].width = 32
    ws.column_dimensions["C"].width = 22
    ws.freeze_panes = "A5"


def _split_class_label(label):
    """'FINC1004_26-SB4_PC_1 — WK-02' -> ('FINC1004_26-SB4_PC_1', 'WK-02')."""
    subject, _, class_group = str(label or "").partition(" — ")
    return subject, class_group


def _fill_colour(cell):
    if cell.fill is None or cell.fill.fill_type is None:
        return None
    rgb = cell.fill.start_color.rgb
    return rgb[-6:].upper() if isinstance(rgb, str) else None


def _sum_counts(values):
    nums = [v for v in values if isinstance(v, (int, float))]
    return sum(nums) if nums else None


def _sum_fractions(values):
    """Sum 'n/d' strings as numerator/denominator: 7/16 + 14/15 -> 21/31."""
    num = den = 0
    found = False
    for v in values:
        m = FRACTION_RE.match(str(v)) if v is not None else None
        if m:
            num += int(m.group(1))
            den += int(m.group(2))
            found = True
    return f"{num}/{den}" if found else None


def _teachers(entries):
    """Distinct teacher names across classes. A class's teacher field can hold
    several comma-separated names, so names are deduplicated individually."""
    seen = []
    for e in entries:
        for t in str(e["teacher"] or "").split(","):
            t = t.strip()
            if t and t not in seen:
                seen.append(t)
    return seen


def _read_student_rows(ws):
    """Data rows (row 6 onward) of a class sheet, as lists of source cells."""
    rows = []
    for r in range(6, ws.max_row + 1):
        cells = [ws.cell(row=r, column=c) for c in range(1, ws.max_column + 1)]
        if any(c.value is not None for c in cells):
            rows.append(cells)
    return rows


def _write_subject_sheet(wb, sheet_name, subject, class_entries):
    """One tab for a subject offering, merging all of its class groups' students.

    Rows and styles come straight from the class sheets the Engagement Report
    built, so columns and formatting match them exactly; only the row order,
    the alternating stripe and the three header lines are new.
    """
    sources = [wb[e["sheet"]] for e in class_entries if e["sheet"] in wb.sheetnames]
    template = sources[0]
    n_cols = template.max_column
    headers = [template.cell(row=5, column=c).value for c in range(1, n_cols + 1)]
    for src in sources[1:]:
        if [src.cell(row=5, column=c).value for c in range(1, src.max_column + 1)] != headers:
            raise ValueError(f"Class sheets for {subject} have different columns; can't merge them.")
    col = {h: i for i, h in enumerate(headers)}
    seg_i, last_i, first_i = col["Segment"], col["Surname"], col["First Name"]

    rows = [row for src in sources for row in _read_student_rows(src)]
    rows.sort(
        key=lambda row: (
            SEG_ORDER.get(row[seg_i].value, 99),
            str(row[last_i].value or "").lower(),
            str(row[first_i].value or "").lower(),
        )
    )

    ws = wb.create_sheet(sheet_name)

    # Header rows 1–5: same merges, heights and styles as the class sheets
    for rng in template.merged_cells.ranges:
        if rng.max_row <= 4:
            ws.merge_cells(rng.coord)
    for r in range(1, 6):
        ws.row_dimensions[r].height = template.row_dimensions[r].height
    for r in range(1, 4):
        ws.cell(row=r, column=1)._style = copy(template.cell(row=r, column=1)._style)
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=5, column=c, value=headers[c - 1])
        cell._style = copy(template.cell(row=5, column=c)._style)

    class_groups = [_split_class_label(e["class_label"])[1] or e["class_label"] for e in class_entries]
    ws.cell(row=1, column=1).value = (
        f"{subject} — {len(rows)} students  •  "
        f"{len(class_entries)} class(es): {', '.join(class_groups)}"
    )

    subtitle = ", ".join(
        f"S{i}: {sum(1 for row in rows if row[seg_i].value == f'S{i}')}" for i in range(1, 8)
    )
    teachers = _teachers(class_entries)
    if teachers:
        subtitle += f"  •  {'Teachers' if len(teachers) > 1 else 'Teacher'}: {', '.join(teachers)}"
    ws.cell(row=2, column=1).value = subtitle

    week_match = WEEK_NOTE_RE.search(str(template.cell(row=3, column=1).value or ""))
    week_note = f" {week_match.group(1)}." if week_match else ""
    back_cell = ws.cell(row=3, column=1)
    back_cell.value = (
        "← Back to Index      All students in this subject offering (all class groups), "
        "sorted by segment then surname. See Disc. Class column for class group."
        f"{week_note}"
    )
    back_cell.hyperlink = Hyperlink(ref=back_cell.coordinate, location="'Subject Index'!A1")

    # Student rows: copy each cell and its style, then redo the stripe for the
    # merged order. Cells with a non-stripe fill (segment colour in column A,
    # AS status colours) keep their own fill.
    stripe = PatternFill("solid", start_color=STRIPE_COLOUR, end_color=STRIPE_COLOUR)
    no_fill = PatternFill(fill_type=None)
    for i, src_cells in enumerate(rows):
        r = 6 + i
        for c, src in enumerate(src_cells, start=1):
            cell = ws.cell(row=r, column=c, value=src.value)
            cell._style = copy(src._style)
            if _fill_colour(src) in (None, STRIPE_COLOUR):
                cell.fill = stripe if i % 2 == 0 else no_fill

    for key, dim in template.column_dimensions.items():
        new_dim = ws.column_dimensions[key]
        new_dim.min, new_dim.max = dim.min, dim.max
        new_dim.width = dim.width
        new_dim.outline_level = dim.outline_level
        new_dim.hidden = dim.hidden
    ws.sheet_properties.outlinePr.summaryRight = template.sheet_properties.outlinePr.summaryRight
    ws.freeze_panes = template.freeze_panes
    return ws


def _write_subject_index(wb, program, subjects, subject_sheets):
    """Build the 'Subject Index' sheet: a bold total row per subject offering
    (linked to its merged tab), with a grey row per class group underneath it
    linking to that class's tab."""
    ws = wb.create_sheet("Subject Index", 0)
    n_classes = sum(len(entries) for _, entries in subjects)

    ws.cell(row=1, column=1, value=f"Program {program} — Subject Index").font = Font(bold=True, size=14)
    ws.cell(
        row=2,
        column=1,
        value=f"{len(subjects)} subject offering(s), {n_classes} class(es)  •  "
              "Click a subject or class to jump to its sheet",
    ).font = Font(italic=True, color="666666")

    header_row = 4
    header_fill = PatternFill("solid", start_color="1F2937", end_color="1F2937")
    for c, title in enumerate(SUBJECT_INDEX_COLUMNS, start=1):
        cell = ws.cell(row=header_row, column=c, value=title)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill

    metric_keys = ["enrolled", "active", "at_risk"]
    as_keys = ["as1_sub", "as2_sub", "as3_sub"]
    subject_fill = PatternFill("solid", start_color=SUBJECT_ROW_COLOUR, end_color=SUBJECT_ROW_COLOUR)
    r = header_row
    for n, (subject, entries) in enumerate(subjects, start=1):
        r += 1
        values = (
            [n, subject, ", ".join(_teachers(entries))]
            + [_sum_counts(e[k] for e in entries) for k in metric_keys]
            + [_sum_fractions(e[k] for e in entries) for k in as_keys]
        )
        for c, v in enumerate(values, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = Font(bold=True)
            cell.fill = subject_fill
        name_cell = ws.cell(row=r, column=2)
        name_cell.hyperlink = Hyperlink(ref=name_cell.coordinate, location=f"'{subject_sheets[subject]}'!A1")
        name_cell.font = Font(bold=True, color="0563C1", underline="single")

        for e in entries:
            r += 1
            class_group = _split_class_label(e["class_label"])[1] or e["class_label"]
            values = (
                [None, f"    {class_group}", e["teacher"]]
                + [e[k] for k in metric_keys]
                + [e[k] for k in as_keys]
            )
            for c, v in enumerate(values, start=1):
                ws.cell(row=r, column=c, value=v).font = Font(color=CLASS_ROW_FONT_COLOUR)
            if e["sheet"] in wb.sheetnames and e["sheet"] not in subject_sheets.values():
                class_cell = ws.cell(row=r, column=2)
                class_cell.hyperlink = Hyperlink(ref=class_cell.coordinate, location=f"'{e['sheet']}'!A1")
                class_cell.font = Font(color=CLASS_ROW_FONT_COLOUR, underline="single")

    ws.column_dimensions["B"].width = 32
    ws.column_dimensions["C"].width = 70
    ws.freeze_panes = "A5"


def _ref_subject_programs():
    """Subject code -> programs whose structure includes it, from the REF sheet
    the Re-registration tool uses (any session's subject list counts)."""
    ref = json.loads(REF_PROGRAMS_PATH.read_text())
    owners = defaultdict(set)
    for program, info in ref.items():
        lists = [info.get("subjects", {})] + [
            change.get("subjects", {}) for change in info.get("from_session", {}).values()
        ]
        for subjects in lists:
            for code in subjects.values():
                owners[code].add(program)
    return owners


def group_entries_by_subject_program(wb, entries):
    """Subject mode's program split: each subject offering goes, with all of its
    class groups, to the program whose REF structure includes its subject code.
    Students from other programs taking it as an elective stay in.

    A subject in several programs' structures (e.g. HUMN1068 in 7188 and 7198)
    goes to whichever of them has the most students in the offering (Course
    column). A subject in none of them goes to the program, among those in the
    Class Index, with the most students in it. Ties go to the lowest code.
    """
    ref_owners = _ref_subject_programs()
    programs = {e["program"] for e in entries if e["program"]}
    by_subject = defaultdict(list)
    for e in entries:
        by_subject[_split_class_label(e["class_label"])[0]].append(e)

    by_program = defaultdict(list)
    for subject, sub_entries in by_subject.items():
        courses = Counter()
        for e in sub_entries:
            if e["sheet"] not in wb.sheetnames:
                continue
            ws = wb[e["sheet"]]
            headers = [c.value for c in ws[5]]
            if "Course" not in headers:
                continue
            ci = headers.index("Course")
            courses.update(
                str(row[ci]) for row in ws.iter_rows(min_row=6, values_only=True)
                if row[ci] is not None
            )

        owners = ref_owners.get(subject.split("_")[0], set())
        candidates = owners or programs
        counts = {p: courses.get(p, 0) for p in candidates}
        if len(owners) == 1:
            program = next(iter(owners))
        elif counts and max(counts.values()) > 0:
            top = max(counts.values())
            program = min(p for p, n in counts.items() if n == top)
        elif owners:
            program = min(owners)
        else:
            # No Course data to go on: fall back to the class rows' own programs
            program = Counter(e["program"] for e in sub_entries).most_common(1)[0][0]
        by_program[program or "(no program)"].extend(sub_entries)
    return by_program


def build_subject_program_workbook(file_bytes, program, prog_entries, all_class_sheets):
    """Same program split as build_program_workbook, but organised by subject
    offering: a 'Subject Index', then for each offering a tab with all of its
    class groups merged, followed by each of its class tabs (kept as built)."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))

    by_subject = defaultdict(list)
    for e in prog_entries:
        if e["sheet"] in wb.sheetnames:
            by_subject[_split_class_label(e["class_label"])[0]].append(e)
    subjects = sorted(
        (
            (subject, sorted(entries, key=lambda e: str(e["class_label"] or "").lower()))
            for subject, entries in by_subject.items()
        ),
        key=lambda item: item[0].lower(),
    )

    # Temporary names while the source sheets still exist (a label without
    # " — " gives a subject name equal to its own class sheet's name)
    temp_sheets = [
        _write_subject_sheet(wb, f"__subject_{i}", subject, entries)
        for i, (subject, entries) in enumerate(subjects)
    ]

    # Keep this offering's class tabs, except a label with no class group,
    # whose single class is the merged tab itself
    class_tabs = {
        subject: [wb[e["sheet"]] for e in entries if _split_class_label(e["class_label"])[1]]
        for subject, entries in subjects
    }
    keep = {ws.title for ws in temp_sheets} | {ws.title for tabs in class_tabs.values() for ws in tabs}
    for name in list(wb.sheetnames):
        if name not in keep:
            del wb[name]

    subject_sheets = {}
    for (subject, _), ws in zip(subjects, temp_sheets):
        name = subject[:31] or "Unknown"
        while name in wb.sheetnames or name == "Subject Index":
            name = (name[:28] + "_2")[:31]
        ws.title = name
        subject_sheets[subject] = name

    # Class tabs' "← Back to Index" pointed at the dropped Class Index
    for ws in (ws for tabs in class_tabs.values() for ws in tabs):
        back_cell = ws.cell(row=3, column=1)
        if back_cell.hyperlink:
            back_cell.hyperlink = Hyperlink(ref=back_cell.coordinate, location="'Subject Index'!A1")

    # Order: each merged subject tab followed by its class tabs
    wb._sheets = [
        ws for (subject, _), subject_ws in zip(subjects, temp_sheets)
        for ws in [subject_ws] + class_tabs[subject]
    ]

    _write_subject_index(wb, program, subjects, subject_sheets)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def build_program_workbook(file_bytes, program, prog_entries, all_class_sheets, group_by="class"):
    """Load a fresh copy of the report and strip it down to just this program.

    Only whole-sheet deletion is used on the original workbook (safe, well
    supported by openpyxl) — kept class sheets are byte-for-byte untouched, so
    all their formatting, conditional formatting and hyperlinks survive. The
    old Class Index, Summary and Assessment Detail sheets mix in every other
    program's data and aren't safely row-filterable (see _write_program_index
    docstring), so they're dropped and replaced with a fresh, program-scoped
    Class Index built from parsed data rather than sheet surgery.

    group_by="subject" builds one tab per subject offering instead (see
    build_subject_program_workbook).
    """
    if group_by == "subject":
        return build_subject_program_workbook(file_bytes, program, prog_entries, all_class_sheets)

    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    keep_sheets = {e["sheet"] for e in prog_entries}

    for name in list(wb.sheetnames):
        if name in ("No match", "Summary", "Assessment Detail", "Class Index"):
            del wb[name]
        elif name in all_class_sheets and name not in keep_sheets:
            del wb[name]

    _write_program_index(wb, program, prog_entries)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


uploaded = st.file_uploader("Class Report (.xlsx)", type=["xlsx"])

if uploaded:
    file_bytes = uploaded.getvalue()
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    entries = parse_class_index(wb)

    if entries is None:
        st.error(
            "Couldn't find a 'Class Index' sheet with a Program column in this file. "
            "Upload a Class Report generated by the Engagement Report builder."
        )
    elif not entries:
        st.warning("Class Index has no class rows to split.")
    else:
        group_by = st.radio(
            "Group sheets by",
            ["class", "subject"],
            format_func=lambda x: {
                "class": "Class group (one tab per class, each class filed under its main program)",
                "subject": "Subject offering (each offering filed whole under the program it belongs "
                           "to in the REF program structures: a tab with all its students, then a tab per class)",
            }[x],
            key="splitter_group_by",
        )

        if group_by == "subject":
            by_program = group_entries_by_subject_program(wb, entries)
        else:
            by_program = defaultdict(list)
            for e in entries:
                by_program[e["program"] or "(no program)"].append(e)
        all_class_sheets = {e["sheet"] for e in entries}

        st.subheader("Programs found")
        summary_rows = [
            {
                "Program": program,
                "Classes": len(prog_entries),
                "Teachers": len({e["teacher"] for e in prog_entries if e["teacher"]}),
            }
            for program, prog_entries in sorted(by_program.items())
        ]
        st.dataframe(summary_rows, use_container_width=True, hide_index=True)

        base_name = re.sub(r"\.xlsx$", "", uploaded.name, flags=re.IGNORECASE)
        short_name = _short_report_name(base_name)
        zip_suffix = "_by_subject" if group_by == "subject" else "_by_program"

        if st.button(f"Build {len(by_program)} program files", type="primary"):
            outputs = {}
            progress = st.progress(0.0)
            for i, (program, prog_entries) in enumerate(sorted(by_program.items())):
                safe_program = re.sub(r"[^A-Za-z0-9_-]+", "_", program)
                buf = build_program_workbook(
                    file_bytes, program, prog_entries, all_class_sheets, group_by=group_by
                )
                outputs[f"{short_name}_{safe_program}.xlsx"] = buf
                progress.progress((i + 1) / len(by_program))
            progress.empty()

            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for filename, buf in outputs.items():
                    zf.writestr(filename, buf.getvalue())
            zip_buf.seek(0)

            st.success(f"Built {len(outputs)} program files.")
            st.download_button(
                "Download all as .zip",
                data=zip_buf,
                file_name=f"{short_name}{zip_suffix}.zip",
                mime="application/zip",
            )

            with st.expander("Download individual files"):
                for filename, buf in outputs.items():
                    st.download_button(
                        filename,
                        data=buf,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_{filename}",
                    )
else:
    st.info("Upload a Class Report to split it by program.")
