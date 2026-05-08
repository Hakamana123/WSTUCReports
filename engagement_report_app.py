"""
WSUTC Student Engagement Report Builder
========================================
Streamlit app for generating weekly engagement reports for GEDU0016 and GEDU0017
(or any single Blackboard subject following the same export format).

Run with:
    streamlit run engagement_report_app.py

Then open the URL it prints (usually http://localhost:8501) in a browser.
"""
import io
import re
import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta

import streamlit as st
import xlrd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ===========================================================================
# CONSTANTS
# ===========================================================================
NS = 'urn:schemas-microsoft-com:office:spreadsheet'
P = f'{{{NS}}}'

# Column index → day-of-month for any month section in the SpreadsheetML usage report.
# Same shape across Feb/Mar/Apr (the source format is consistent).
DAY_TO_COL = {
    1: 5, 2: 7, 3: 8, 4: 9, 5: 10, 6: 11, 7: 13, 8: 15, 9: 17, 10: 19,
    11: 20, 12: 22, 13: 23, 14: 25, 15: 27, 16: 28, 17: 29, 18: 31, 19: 32,
    20: 33, 21: 34, 22: 35, 23: 37, 24: 38, 25: 39, 26: 41, 27: 42, 28: 43,
    29: 45, 30: 46, 31: 47,
}

# Teaching week 1 starts here (Monday). Weeks run Monday–Sunday.
WEEK1_START = date(2026, 3, 2)

# Exclusion rules (apply every run)
EXCLUDE_SURNAMES = {'Curtin', 'Rouillon', 'Turro', 'Tyler', 'Wyborn', 'Wagstaffe', 'Pinkerton'}
USAGE_EXCLUDE_NAMES = {'Guest', 'Total'}

# Style palette
NAVY = '1F2F4E'
ACCENT = '2E5D9F'
LIGHT = 'EEF2F7'
ALT_ROW = 'F7F9FC'
WHITE = 'FFFFFF'
RED = 'C0392B'
ORANGE = 'D68910'
YELLOW = 'F1C40F'
GREEN = '27AE60'
BLUE = '2980B9'
PURPLE = '7D3C98'
SEG_COLOURS = {'S1': RED, 'S2': '8B4513', 'S3': ORANGE, 'S4': YELLOW,
               'S5': BLUE, 'S6': PURPLE, 'S7': GREEN}

# ===========================================================================
# CLASS LIST PARSING
# ===========================================================================
def load_classlist(file_bytes):
    """Auto-detects .xls (CFB) vs .xlsx (enriched) format from file magic bytes."""
    if file_bytes[:2] == b'PK':
        return _load_classlist_xlsx(file_bytes)
    return _load_classlist_xls(file_bytes)


def _load_classlist_xls(file_bytes):
    """Parse the original CFB .xls class list format."""
    book = xlrd.open_workbook(file_contents=file_bytes)
    sheet = book.sheet_by_index(0)

    headers = [str(sheet.cell_value(6, c)).strip().lower() for c in range(sheet.ncols)]
    cmap = {h: i for i, h in enumerate(headers) if h}
    if 'student_code' not in cmap:
        raise ValueError("Class list header row not found at row 7. Check the file format.")

    cm = {
        'sid':    cmap['student_code'],
        'first':  cmap.get('first_name', 1),
        'last':   cmap.get('last_name', 2),
        'course': cmap.get('course', 5),
        'email':  cmap.get('email_address', 6),
    }

    # Detect subject from row 1 (e.g. "GEDU0016_26-AUT_ON_2")
    subject_raw = str(sheet.cell_value(1, 0)).strip()
    subject_match = re.match(r'([A-Z]+\d+)', subject_raw)
    subject_code = subject_match.group(1) if subject_match else 'UNKNOWN'

    students = {}
    for r in range(7, sheet.nrows):
        sid = str(sheet.cell_value(r, cm['sid'])).strip()
        if not sid or sid == 'student_code':
            continue
        if sid.startswith('30') or sid.startswith('96'):
            continue
        last = str(sheet.cell_value(r, cm['last'])).strip()
        if last in EXCLUDE_SURNAMES:
            continue
        course_val = str(sheet.cell_value(r, cm['course'])).strip()
        students[sid] = {
            'sid': sid,
            'first': str(sheet.cell_value(r, cm['first'])).strip(),
            'last': last,
            'course': course_val,
            'course_code': course_val,
            'discipline_subject': '',
            'discipline_class': '',
            'email': str(sheet.cell_value(r, cm['email'])).strip(),
            'discipline_teacher': '',
            'subject_code': '',
        }
    return subject_code, students


def _load_classlist_xlsx(file_bytes):
    """Parse enriched .xlsx class list with Discipline Subject/Class/Teacher columns."""
    wb_cl = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb_cl.active

    # Build header map from row 1
    headers = {}
    for c in range(1, ws.max_column + 1):
        val = ws.cell(1, c).value
        if val:
            headers[str(val).strip()] = c

    required = {'Student ID', 'First Name', 'Last Name'}
    missing = required - set(headers.keys())
    if missing:
        raise ValueError(f"Enriched class list missing columns: {missing}")

    # Detect subject code from GEDU Subject or Subject Code column
    subject_code = 'UNKNOWN'
    subj_col_name = None
    for candidate in ('GEDU Subject', 'Subject Code'):
        if candidate in headers:
            subj_col_name = candidate
            break
    if subj_col_name:
        raw = str(ws.cell(2, headers[subj_col_name]).value or '').strip()
        m = re.match(r'([A-Z]+\d+)', raw)
        if m:
            subject_code = m.group(1)

    def cell_str(row, col_name, default_col=1):
        c = headers.get(col_name, default_col)
        v = ws.cell(row, c).value
        if v is None:
            return ''
        if isinstance(v, (int, float)):
            return str(int(v))
        return str(v).strip()

    students = {}
    for r in range(2, ws.max_row + 1):
        sid_raw = ws.cell(r, headers['Student ID']).value
        if sid_raw is None:
            continue
        sid = str(int(sid_raw)).strip() if isinstance(sid_raw, (int, float)) else str(sid_raw).strip()
        if not sid or not sid.isdigit():
            continue
        if sid.startswith('30') or sid.startswith('96'):
            continue
        last = cell_str(r, 'Last Name')
        if last in EXCLUDE_SURNAMES:
            continue
        course_code = cell_str(r, 'Course Code')
        students[sid] = {
            'sid': sid,
            'first': cell_str(r, 'First Name'),
            'last': last,
            'course': course_code,
            'course_code': course_code,
            'discipline_subject': cell_str(r, 'Discipline Subject'),
            'discipline_class': cell_str(r, 'Discipline Class'),
            'email': cell_str(r, 'Email'),
            'discipline_teacher': cell_str(r, 'Discipline Teacher'),
            'subject_code': cell_str(r, subj_col_name) if subj_col_name else '',
        }
    return subject_code, students

# ===========================================================================
# LOGIN REPORT PARSING
# ===========================================================================
def load_login_report(file_bytes):
    """Auto-detects NLI and LI section header rows."""
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    out = {}
    window_start = window_end = None

    # Try to find the report window from the descriptive rows
    for r in range(1, min(ws.max_row + 1, 60)):
        for c in range(1, min(ws.max_column + 1, 20)):
            v = ws.cell(r, c).value
            if isinstance(v, str) and 'between' in v and 'to' in v:
                m = re.search(r'between\s+(\d{1,2}/\d{1,2}/\d{4})\s+to\s+(\d{1,2}/\d{1,2}/\d{4})', v)
                if m:
                    try:
                        window_start = datetime.strptime(m.group(1), '%d/%m/%Y').date()
                        window_end = datetime.strptime(m.group(2), '%d/%m/%Y').date()
                    except ValueError:
                        pass
                if window_start: break
        if window_start: break

    # Find NLI and LI section header rows by scanning for SURNAME label
    nli_header_row = None
    li_header_row = None
    for r in range(1, ws.max_row + 1):
        v8 = ws.cell(r, 8).value
        v6 = ws.cell(r, 6).value
        if v8 == 'SURNAME' and nli_header_row is None:
            nli_header_row = r
        if v6 == 'SURNAME':
            li_header_row = r

    if nli_header_row is None or li_header_row is None:
        raise ValueError("Could not find login report section headers")

    # NLI: surname=col8, first=col17, sid=col23, email=col28, days=col40, last=col45, total=col49
    for r in range(nli_header_row + 1, li_header_row):
        sid = ws.cell(r, 23).value
        if sid is None: continue
        sid = str(sid).strip()
        if not sid.isdigit(): continue
        days = ws.cell(r, 40).value
        last = ws.cell(r, 45).value
        total = ws.cell(r, 49).value
        never = (str(days).strip().upper() == 'NEVER')
        days_n = None if never else (int(days) if isinstance(days, (int, float)) else None)
        last_d = last if isinstance(last, (datetime, date)) else None
        out[sid] = {
            'days_since': days_n, 'last_login': last_d,
            'total_logins': int(total) if isinstance(total, (int, float)) else 0,
            'never': never, 'in_window': False,
        }

    # LI: surname=col6, first=col16, sid=col22, email=col27, days=col39, last=col44, total=col48
    for r in range(li_header_row + 1, ws.max_row + 1):
        sid = ws.cell(r, 22).value
        if sid is None: continue
        sid = str(sid).strip()
        if not sid.isdigit(): continue
        days = ws.cell(r, 39).value
        last = ws.cell(r, 44).value
        total = ws.cell(r, 48).value
        days_n = int(days) if isinstance(days, (int, float)) else None
        last_d = last if isinstance(last, (datetime, date)) else None
        out[sid] = {
            'days_since': days_n, 'last_login': last_d,
            'total_logins': int(total) if isinstance(total, (int, float)) else 0,
            'never': False, 'in_window': True,
        }
    return out, window_start, window_end

# ===========================================================================
# USAGE FILE PARSING
# ===========================================================================
def get_cells(row):
    cells = row.findall(f'{P}Cell'); result = {}; ci = 0
    for cell in cells:
        idx = cell.get(f'{P}Index')
        if idx: ci = int(idx) - 1
        d = cell.find(f'{P}Data')
        result[ci] = d.text if d is not None else ''
        ci += 1
    return result

def parse_usage_file(file_bytes):
    """Returns {(year, month, day): {sid: hits}} for all populated days in the file."""
    tree = ET.parse(io.BytesIO(file_bytes))
    rows = tree.getroot().findall(f'.//{P}Row')
    data = {}
    current_month = None
    for r in rows:
        c = get_cells(r)
        m = c.get(1)
        if isinstance(m, str) and re.match(r'^\d{4}-\d{2}$', m):
            current_month = m
            continue
        if current_month is None:
            continue
        name = (c.get(1) or '').strip()
        if not name:
            continue
        if (name.startswith('Total') or name.startswith('Overall')
                or name == 'Chart does not appear in Excel'):
            current_month = None
            continue
        if name in USAGE_EXCLUDE_NAMES or 'PreviewUser' in name:
            continue
        if '(' not in name or ')' not in name:
            continue
        sid = name[name.rfind('(')+1:name.rfind(')')].strip()
        if not sid.isdigit():
            continue
        if sid.startswith('30') or sid.startswith('96'):
            continue
        year_str, month_str = current_month.split('-')
        year = int(year_str); month = int(month_str)
        for day, col in DAY_TO_COL.items():
            v = c.get(col)
            if v is None or v == '':
                continue
            try:
                hits = int(float(v))
            except (ValueError, TypeError):
                continue
            key = (year, month, day)
            data.setdefault(key, {})[sid] = hits
    return data

def merge_usage_files(usage_file_list):
    """Merges multiple usage files. Last file wins on overlapping days."""
    merged = {}
    for file_bytes in usage_file_list:
        single = parse_usage_file(file_bytes)
        for date_key, student_hits in single.items():
            merged[date_key] = student_hits  # last write wins
    return merged

# ===========================================================================
# GRADE CENTRE PARSING
# ===========================================================================
def load_grade_centre(file_bytes):
    """Parse a Blackboard Grade Centre export (.xls = UTF-16 TSV).

    Auto-detects primary assessment columns (headers starting with
    'Assessment N' but excluding resubmit columns and sub-component columns
    where a collated total exists).

    Returns
    -------
    gc_data : dict  {sid: {short_label: status_string}}
    gc_labels : list of str  Short labels in column order, e.g. ['AS1', 'AS2', 'AS3']
    """
    import csv as _csv

    # Decode the UTF-16 LE file
    text = file_bytes.decode('utf-16-le')
    # Strip BOM if present
    if text and text[0] == '\ufeff':
        text = text[1:]
    reader = _csv.reader(text.splitlines(), delimiter='\t')
    headers_raw = next(reader)
    rows = list(reader)

    # Clean headers: strip surrounding quotes and whitespace
    headers = [h.strip().strip('"') for h in headers_raw]

    # Find the student-ID column (Username or Student ID)
    sid_col = None
    for i, h in enumerate(headers):
        if h.lower() in ('username', 'student id'):
            sid_col = i
            break
    if sid_col is None:
        raise ValueError("Grade Centre: cannot find Username or Student ID column.")

    # Find availability column
    avail_col = None
    for i, h in enumerate(headers):
        if h.lower() == 'availability':
            avail_col = i
            break

    # ---- Auto-detect assessment columns ----
    # Step 1: find all columns whose short name starts with "Assessment"
    # Separate primary columns from resubmit columns.
    def short_name(h):
        return h.split('[')[0].strip() if '[' in h else h.strip()

    import re as _re

    primary_cols = []
    resubmit_cols = []
    for i, h in enumerate(headers):
        sn = short_name(h).lower()
        if not sn.startswith('assessment'):
            continue
        if 'resubmit' in sn or 'resubmission' in sn:
            resubmit_cols.append((i, short_name(h)))
        else:
            primary_cols.append((i, short_name(h)))

    # Step 2: identify collated/total columns — if a collated total exists for
    # an assessment number, skip the sub-components (e.g. "Assessment 3: Main
    # Project" and "Assessment 3: Peer Review" are dropped when
    # "Assessment 3 COLLATED TOTAL" is present).
    collated_nums = set()
    for _, name in primary_cols:
        if 'collated' in name.lower() or 'total' in name.lower().split('assessment')[-1]:
            m = _re.search(r'Assessment\s+(\d+)', name, _re.IGNORECASE)
            if m:
                collated_nums.add(m.group(1))

    # Step 3: for assessment numbers that have a collated column, keep only the
    # collated column; for others keep the primary column (skip sub-parts that
    # share the same assessment number if there's also a standalone one).
    selected = []
    seen_nums = set()
    # Prioritise collated columns, then standalone
    for i, name in primary_cols:
        m = _re.search(r'Assessment\s+(\d+)', name, _re.IGNORECASE)
        if not m:
            continue
        num = m.group(1)
        is_collated = ('collated' in name.lower() or 'total' in name.lower().split('assessment')[-1].split(':')[0])
        if num in collated_nums:
            if is_collated and num not in seen_nums:
                selected.append((i, name, num))
                seen_nums.add(num)
        else:
            same_num = [n for _, n in primary_cols
                        if _re.search(r'Assessment\s+' + num + r'\b', n, _re.IGNORECASE)]
            if len(same_num) == 1 and num not in seen_nums:
                selected.append((i, name, num))
                seen_nums.add(num)
            elif num not in seen_nums:
                selected.append((i, name, num))
                seen_nums.add(num)

    # Sort by assessment number
    selected.sort(key=lambda x: int(x[2]))

    # Step 4: match resubmit columns to their primary assessment number
    resub_map = {}  # assessment_num -> resubmit_col_index
    for i, name in resubmit_cols:
        m = _re.search(r'Assessment\s+(\d+)', name, _re.IGNORECASE)
        if m:
            num = m.group(1)
            if num in seen_nums:
                resub_map[num] = i

    # Build short labels
    gc_labels = [f'AS{num}' for _, _, num in selected]
    gc_col_indices = [i for i, _, _ in selected]
    gc_nums = [num for _, _, num in selected]

    # ---- Extract per-student data ----
    gc_data = {}
    for row in rows:
        if len(row) <= sid_col:
            continue
        sid = row[sid_col].strip().strip('"')
        if not sid or not sid.isdigit():
            continue
        if sid.startswith('30') or sid.startswith('96'):
            continue
        # Skip unavailable students if we can detect it
        if avail_col is not None and row[avail_col].strip().strip('"').lower() != 'yes':
            continue
        # Skip preview users
        if 'PreviewUser' in (row[0] if row else ''):
            continue

        student_gc = {}
        for label, col_idx, num in zip(gc_labels, gc_col_indices, gc_nums):
            val = row[col_idx].strip().strip('"') if col_idx < len(row) else ''
            # Check resubmit column — if it has a value, use it instead
            resub_idx = resub_map.get(num)
            if resub_idx is not None and resub_idx < len(row):
                resub_val = row[resub_idx].strip().strip('"')
                if resub_val:
                    val = resub_val
            if val == '':
                student_gc[label] = 'No Submission'
            else:
                student_gc[label] = val
        gc_data[sid] = student_gc

    return gc_data, gc_labels


# ===========================================================================
# WEEK DETECTION
# ===========================================================================
def detect_current_week(merged_usage, override_latest=None):
    """Returns (current_week_num, days_in_current_week, latest_date).

    If override_latest is given, the latest date is capped at that value
    (used to align usage data to the login report window).
    """
    if not merged_usage:
        return None, 0, None
    all_dates = [date(y, m, d) for (y, m, d) in merged_usage.keys()]
    if override_latest:
        eligible = [d for d in all_dates if d <= override_latest]
        latest = max(eligible) if eligible else override_latest
    else:
        latest = max(all_dates)
    if latest < WEEK1_START:
        return None, 0, latest
    week_num = ((latest - WEEK1_START).days // 7) + 1
    week_start = WEEK1_START + timedelta(days=(week_num - 1) * 7)
    days_in = (latest - week_start).days + 1
    return week_num, days_in, latest

def bucket_by_week(merged_usage, students, max_week, max_date=None):
    """Returns hits[sid][week_key] for week_key in 'w1'..'w<max_week>'.
    If max_date given, days strictly after it are ignored (used to align with login window).
    """
    hits = {sid: {f'w{i}': 0 for i in range(1, max_week + 1)} for sid in students}
    for (y, m, d), student_hits in merged_usage.items():
        dt = date(y, m, d)
        if dt < WEEK1_START:
            continue
        if max_date and dt > max_date:
            continue
        week_num = ((dt - WEEK1_START).days // 7) + 1
        if week_num > max_week:
            continue
        wkey = f'w{week_num}'
        for sid, h in student_hits.items():
            if sid not in hits:
                continue
            hits[sid][wkey] += h
    return hits

def week_date_range(week_num):
    start = WEEK1_START + timedelta(days=(week_num - 1) * 7)
    end = start + timedelta(days=6)
    return start, end

# ===========================================================================
# SEGMENTATION
# ===========================================================================
def classify(students, login, hits, current_week, prev_days, curr_days):
    """Returns {sid: 'S1'..'S7'}."""
    seg = {}
    week_keys = [f'w{i}' for i in range(1, current_week + 1)]
    curr_key = f'w{current_week}'
    prev_key = f'w{current_week - 1}' if current_week > 1 else None

    # Days-since thresholds calculated from report date
    # Threshold = days from week 1 start to report date
    # For partial week, use the latest known data day
    s2_threshold = (curr_days + (current_week - 1) * 7)  # days back to before Mar 1
    # Wait — we need this from the actual report date. The login window end is the report date.
    # But we don't have it here. We'll just compute it from the latest day.
    # Actually we want: how many days since "before Mar 1"?
    # If latest_data_day is Apr 9 (= 40 days inclusive from Mar 1), then days_since for Feb 28 = 40.
    # General formula: latest = WEEK1_START + (current_week-1)*7 + (curr_days-1)
    # latest_offset_from_mar1 = (current_week-1)*7 + curr_days  (1-indexed inclusive)
    # days_since for Feb 28 = latest_offset_from_mar1 + 0 (Feb 28 is 1 day before Mar 1)
    # So S2_DAYS_THRESHOLD = (current_week - 1) * 7 + curr_days + 0
    # Hmm but we computed days as `report_date - last_login`. Let me verify.
    # If latest = Apr 9 (W6, day 5), and last login = Feb 28, days_since = 40.
    # offset = (6-1)*7 + 5 = 40. Match.
    # So S2 threshold = offset from Mar 1 to latest day = (current_week-1)*7 + curr_days
    s2_threshold = (current_week - 1) * 7 + curr_days
    # S3 range: last login in Mar 1-7 means days_since between (offset-6) and (offset)
    s3_low = s2_threshold - 6
    s3_high = s2_threshold - 0  # actually Mar 1 = offset, Mar 7 = offset - 6
    # Wait — Feb 28 is offset - 0? No.
    # Let latest day be Apr 9 (W6, day 5). offset = 40.
    # Days since: Apr 9 - last_login.
    # Last login Apr 8 → days_since = 1.
    # Last login Mar 1 → days_since = 39.
    # Last login Mar 7 → days_since = 33.
    # Last login Feb 28 → days_since = 40.
    # So S2 threshold ≥ 40 (≥ s2_threshold), S3 range = 33 to 39 inclusive.
    # That means S2 threshold = offset, S3 range = (offset-7, offset-1) inclusive.
    s2_threshold = (current_week - 1) * 7 + curr_days
    s3_low = s2_threshold - 7   # Mar 7 → offset - 6, but we want 33 in example: 40 - 7 = 33 ✓
    s3_high = s2_threshold - 1  # Mar 1 → offset - 1: 40 - 1 = 39 ✓

    for sid in students:
        h = hits[sid]
        l = login.get(sid)
        zero_usage = all(h[k] == 0 for k in week_keys)
        days_since = None; in_window = False; never_flag = False
        if l is not None:
            days_since = l['days_since']; in_window = l['in_window']; never_flag = l['never']

        # S1: never logged in at all (NEVER flag, or no login record) AND zero usage
        if zero_usage and (never_flag or l is None):
            seg[sid] = 'S1'; continue

        # S2: has logged in before Mar 2 but not since AND zero usage
        if zero_usage and days_since is not None and days_since >= s2_threshold:
            seg[sid] = 'S2'; continue

        # Catch-all for any remaining zero-usage students (logged in during teaching but no hits)
        if zero_usage:
            seg[sid] = 'S2'; continue

        # S3: login in W1 range, not in window
        if (days_since is not None and s3_low <= days_since <= s3_high and not in_window):
            seg[sid] = 'S3'; continue

        # S4: previously active, currently zero
        if h[curr_key] == 0:
            if any(h[k] > 0 for k in week_keys[:-1]):
                seg[sid] = 'S4'; continue

        # S5/S6/S7: comparison
        if prev_key:
            prev = h[prev_key]; curr = h[curr_key]
            if prev == 0 and curr > 0:
                seg[sid] = 'S5'; continue
            if prev > 0 and curr > 0:
                prev_avg = prev / prev_days
                curr_avg = curr / curr_days
                seg[sid] = 'S6' if curr_avg < prev_avg * 0.5 else 'S7'
                continue

        # Fallback
        seg[sid] = 'S1'
    return seg, s2_threshold, (s3_low, s3_high)

# ===========================================================================
# WORKBOOK BUILDING
# ===========================================================================
def thin_border():
    side = Side(style='thin', color='D5DBDB')
    return Border(left=side, right=side, top=side, bottom=side)

def write_tab_header(ws, title, subtitle, description, n_cols, seg_code=None):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    c = ws.cell(1, 1, title)
    c.font = Font(name='Arial', size=14, bold=True, color=WHITE)
    c.fill = PatternFill('solid', start_color=NAVY)
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws.row_dimensions[1].height = 26

    seg_colour = SEG_COLOURS.get(seg_code, ACCENT) if seg_code else ACCENT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
    c = ws.cell(2, 1, subtitle or '')
    c.font = Font(name='Arial', size=11, bold=True, color=WHITE)
    c.fill = PatternFill('solid', start_color=seg_colour)
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws.row_dimensions[2].height = 20

    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=n_cols)
    c = ws.cell(3, 1, description)
    c.font = Font(name='Arial', size=10, italic=True, color='2C3E50')
    c.fill = PatternFill('solid', start_color=LIGHT)
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1, wrap_text=True)
    ws.row_dimensions[3].height = 36
    ws.row_dimensions[4].height = 6

def write_col_headers(ws, headers, row=5):
    for i, h in enumerate(headers, 1):
        c = ws.cell(row, i, h)
        c.font = Font(name='Arial', size=10, bold=True, color=WHITE)
        c.fill = PatternFill('solid', start_color=ACCENT)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = thin_border()
    ws.row_dimensions[row].height = 30

def write_data_rows(ws, data_rows, start_row=6):
    for ri, row in enumerate(data_rows):
        excel_row = start_row + ri
        fill = PatternFill('solid', start_color=ALT_ROW) if ri % 2 == 0 else None
        for ci, val in enumerate(row, 1):
            c = ws.cell(excel_row, ci, val)
            c.font = Font(name='Arial', size=10)
            if fill:
                c.fill = fill
            c.alignment = Alignment(horizontal='left' if ci <= 5 else 'center', vertical='center')
            c.border = thin_border()

def autosize(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def fmt_login(l):
    if l is None: return ('NO LOGIN RECORD', '-', 0)
    if l['never']: return ('NEVER', 'NEVER', l['total_logins'] or 0)
    days = l['days_since']; last = l['last_login']
    last_str = last.strftime('%Y-%m-%d') if isinstance(last, (datetime, date)) else (str(last) if last else '-')
    return (days if days is not None else '-', last_str, l['total_logins'] or 0)

def build_workbook(subject_code, students, login, hits, seg, current_week, prev_days, curr_days,
                   is_partial, latest_date, login_window, s2_threshold, s3_range):
    wb = Workbook()
    wb.remove(wb.active)

    enrolled = len(students)
    counts = {f'S{i}': 0 for i in range(1, 8)}
    for s in seg.values():
        counts[s] = counts.get(s, 0) + 1

    week_keys = [f'w{i}' for i in range(1, current_week + 1)]
    week_labels = [f'W{i}' for i in range(1, current_week + 1)]
    curr_key = f'w{current_week}'
    prev_key = f'w{current_week - 1}' if current_week > 1 else None

    curr_start, curr_end = week_date_range(current_week)
    if is_partial:
        curr_label = f'W{current_week} ({curr_start.strftime("%b %#d")}-{latest_date.strftime("%#d")}, {curr_days}d partial)'
    else:
        curr_label = f'W{current_week} ({curr_start.strftime("%b %#d")}-{curr_end.strftime("%#d")}, 7d)'
    if prev_key:
        prev_start, prev_end = week_date_range(current_week - 1)
        prev_label = f'W{current_week - 1} ({prev_start.strftime("%b %#d")}-{prev_end.strftime("%#d")}, 7d)'
    else:
        prev_label = '(none)'

    week_descriptions = ', '.join(
        f'W{i}={week_date_range(i)[0].strftime("%b %#d")}-{week_date_range(i)[1].strftime("%#d")}'
        for i in range(1, current_week + 1)
    )

    # ============ Summary ============
    ws = wb.create_sheet('Summary')
    write_tab_header(ws,
        f'{subject_code} Engagement Report — {curr_label}',
        f'Login window {login_window}  •  Enrolled {enrolled}  •  Latest data: {latest_date.strftime("%b %#d %Y")}',
        ('Partial-week run. S4 inflated, S7 understated by timing artefact.' if is_partial
         else 'Full-week run. Standard segmentation.'),
        6)
    write_col_headers(ws, ['Segment', 'Label', 'Count', '% of Enrolled', 'Status', 'Description'], row=5)

    seg_descriptions = {
        'S1': f'Never engaged: never logged in at all (or no login record) AND zero hits across W1 through W{current_week}.',
        'S2': 'Pre-teaching ghosts: last login was before Mar 2 (i.e. logged in during pre-teaching but never returned) AND zero hits all weeks.',
        'S3': 'W1 early drop-offs: last login fell in W1 and they have not returned in the current login window.',
        'S4': f'Active then absent in W{current_week}: had hits in a previous week but zero in W{current_week} to date. Split into "Just Dropped" (W{current_week-1}>0) and "Long Silent" (W1-W{current_week-2} active, W{current_week-1}+W{current_week} zero).' if current_week >= 3 else f'Active then absent in W{current_week}.',
        'S5': f'Late arrivals + W{current_week-1} returners: zero hits in W{current_week-1} but appearing in W{current_week}.' if current_week > 1 else 'Late arrivals: appearing in W1.',
        'S6': f'Fading engagers: active both weeks but daily-average hit rate fell 50%+ from W{current_week-1} to W{current_week}.' if current_week > 1 else 'Fading engagers (n/a in W1).',
        'S7': f'Sustained participants: active both weeks with daily-average rate held within 50%.' if current_week > 1 else 'Sustained participants (n/a in W1).',
    }

    seg_meta = [
        ('S1', 'Never Engaged', 'Critical'),
        ('S2', 'Pre-Teaching Ghosts', 'Critical'),
        ('S3', 'W1 Early Drop-Offs', 'High Risk'),
        ('S4', f'Active then W{current_week} Absent', 'Watch (partial)' if is_partial else 'Watch'),
        ('S5', 'Late Arrivals + Returners', 'Mixed'),
        ('S6', 'Fading Engagers', 'High Risk'),
        ('S7', 'Sustained Participants', 'Healthy'),
    ]
    rows = [[code, label, counts[code], counts[code] / enrolled, status, seg_descriptions[code]]
            for code, label, status in seg_meta]
    write_data_rows(ws, rows, start_row=6)
    for ri in range(6, 6 + len(rows)):
        ws.cell(ri, 4).number_format = '0.0%'

    total_row = 6 + len(rows)
    ws.cell(total_row, 1, 'TOTAL').font = Font(name='Arial', size=10, bold=True)
    ws.cell(total_row, 3, sum(counts.values())).font = Font(name='Arial', size=10, bold=True)
    ws.cell(total_row, 4, sum(counts.values()) / enrolled).number_format = '0.0%'
    ws.cell(total_row, 4).font = Font(name='Arial', size=10, bold=True)
    for ci in range(1, 7):
        ws.cell(total_row, ci).fill = PatternFill('solid', start_color=LIGHT)
        ws.cell(total_row, ci).border = thin_border()

    # Footer notes
    s3_or_s4_eligible = 0
    for sid in students:
        h = hits[sid]; l = login.get(sid)
        days_since = l['days_since'] if l else None
        in_window = l['in_window'] if l else False
        s3_eligible = (days_since is not None and s3_range[0] <= days_since <= s3_range[1] and not in_window)
        s4_eligible = (h[curr_key] == 0 and any(h[k] > 0 for k in week_keys[:-1]))
        if s3_eligible and s4_eligible:
            s3_or_s4_eligible += 1

    note_row = total_row + 2
    notes = [
        f'Enrolled (after exclusions): {enrolled}',
        f'Login window: {login_window}',
        f'Teaching weeks: {week_descriptions}',
        f'Comparison pair: {prev_label} vs {curr_label}. Daily averages normalised by actual day count.',
        (f'PARTIAL WEEK WARNING: W{current_week} has {curr_days} days of data. S4 inflated by timing. S7 understated. S5 includes both true late arrivals and W{current_week-1} returners. Do NOT use S4 list for outreach until W{current_week} closes.'
         if is_partial else f'Full week W{current_week}: standard run.'),
        f'Days-since thresholds: S2 ≥ {s2_threshold} days (last login on or before Mar 1); S3 in {s3_range[0]}-{s3_range[1]} days (last login Mar 2-8).',
        f'S3 / S4 dual-eligible students: {s3_or_s4_eligible}.',
        f'Students in class list but missing from login report: {sum(1 for s in students if s not in login)}.',
        f'Leaderboard ranking: hits per enrolled student (weighted engagement).',
    ]
    for i, n in enumerate(notes):
        c = ws.cell(note_row + i, 1, n)
        ws.merge_cells(start_row=note_row + i, start_column=1, end_row=note_row + i, end_column=6)
        c.font = Font(name='Arial', size=9, italic=(i == 4), color=(RED if i == 4 else '2C3E50'),
                      bold=(i == 4))
        c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        ws.row_dimensions[note_row + i].height = 32 if i == 4 else 16
    autosize(ws, [10, 28, 10, 14, 18, 70])
    ws.freeze_panes = 'A6'

    def make_seg_tab(code, label):
        sids = sorted([sid for sid, s in seg.items() if s == code],
                      key=lambda x: (students[x]['last'].lower(), students[x]['first'].lower()))
        ws = wb.create_sheet(f'{code} {label}'[:31])
        return ws, sids

    # ============ S1 ============
    ws, sids = make_seg_tab('S1', 'Never Engaged')
    write_tab_header(ws, 'S1 — Never Engaged',
        f'{len(sids)} students with zero hits across all teaching weeks',
        seg_descriptions['S1'], 10, 'S1')
    write_col_headers(ws, ['Surname', 'First Name', 'Student ID', 'Course', 'Disc. Class', 'Disc. Teacher', 'Email',
                           'Total Logins', 'Last Login', 'Action Required'], row=5)
    rows = []
    for sid in sids:
        st = students[sid]; l = login.get(sid)
        days, last_str, total = fmt_login(l)
        rows.append([st['last'], st['first'], sid, st['course'], st.get('discipline_class', ''), st.get('discipline_teacher', ''), st['email'], total, last_str,
                     'Initial outreach — verify enrolment intent'])
    write_data_rows(ws, rows, start_row=6)
    autosize(ws, [22, 18, 12, 10, 10, 18, 38, 12, 14, 38])
    ws.freeze_panes = 'A6'

    # ============ S2 ============
    ws, sids = make_seg_tab('S2', 'Pre-Teaching Ghosts')
    write_tab_header(ws, 'S2 — Pre-Teaching Ghosts',
        f'{len(sids)} students with zero hits all weeks AND last login pre-Mar 2 or NEVER',
        seg_descriptions['S2'], 11, 'S2')
    write_col_headers(ws, ['Surname', 'First Name', 'Student ID', 'Course', 'Disc. Class', 'Disc. Teacher', 'Email',
                           'Total Logins', 'Days Since', 'Last Login', 'Action Required'], row=5)
    rows = []
    for sid in sids:
        st = students[sid]; l = login.get(sid)
        days, last_str, total = fmt_login(l)
        if total == 1:
            action = 'Single login only — likely orientation visit; escalate'
        else:
            action = 'Pre-teaching login then disengaged — urgent outreach'
        rows.append([st['last'], st['first'], sid, st['course'], st.get('discipline_class', ''), st.get('discipline_teacher', ''), st['email'], total, days, last_str, action])
    write_data_rows(ws, rows, start_row=6)
    autosize(ws, [22, 18, 12, 10, 10, 18, 38, 12, 12, 14, 50])
    ws.freeze_panes = 'A6'

    # ============ S3 ============
    ws, sids = make_seg_tab('S3', 'W1 Drop-Offs')
    write_tab_header(ws, 'S3 — W1 Early Drop-Offs',
        f'{len(sids)} students whose last login fell in W1 (Mar 2-8)',
        seg_descriptions['S3'], 12, 'S3')
    write_col_headers(ws, ['Surname', 'First Name', 'Student ID', 'Course', 'Disc. Class', 'Disc. Teacher', 'Email',
                           'Total Logins', 'Days Since', 'Last Login', 'Risk Level', 'Action Required'], row=5)
    rows = []
    for sid in sids:
        st = students[sid]; l = login.get(sid)
        days, last_str, total = fmt_login(l)
        risk = 'High' if isinstance(days, int) and days >= (s3_range[0] + 3) else 'Medium'
        rows.append([st['last'], st['first'], sid, st['course'], st.get('discipline_class', ''), st.get('discipline_teacher', ''), st['email'], total, days, last_str,
                     risk, 'Re-engage; offer support'])
    write_data_rows(ws, rows, start_row=6)
    autosize(ws, [22, 18, 12, 10, 10, 18, 38, 12, 12, 14, 12, 30])
    ws.freeze_panes = 'A6'

    # ============ S4 (split) ============
    ws, sids = make_seg_tab('S4', f'Active W{current_week} Absent')
    just_dropped = []; long_silent = []
    if prev_key:
        for sid in sids:
            h = hits[sid]
            if h[prev_key] > 0 and h[curr_key] == 0:
                just_dropped.append(sid)
            else:
                long_silent.append(sid)
    else:
        long_silent = list(sids)

    write_tab_header(ws, f'S4 — Active Then Absent in W{current_week}',
        f'{len(sids)} total  •  Just Dropped: {len(just_dropped)}  •  Long Silent: {len(long_silent)}',
        seg_descriptions['S4'], current_week + 9, 'S4')
    headers_s4 = ['Surname', 'First Name', 'Student ID', 'Course', 'Disc. Class', 'Disc. Teacher', 'Email'] + week_labels[:-1] + ['Group', 'Priority']
    write_col_headers(ws, headers_s4, row=5)

    def priority_for(h, group):
        if group == 'Just Dropped':
            basis = h[prev_key]
        else:
            basis = 0
            for k in reversed(week_keys[:-2] if prev_key else week_keys[:-1]):
                if h[k] > 0:
                    basis = h[k]; break
        if basis >= 20: return 'High'
        if basis >= 8: return 'Medium'
        return 'Standard'

    def rows_for_group(sid_list, group):
        out = []
        for sid in sid_list:
            st = students[sid]; h = hits[sid]
            row = [st['last'], st['first'], sid, st['course'], st.get('discipline_class', ''), st.get('discipline_teacher', ''), st['email']]
            row += [h[k] for k in week_keys[:-1]]  # all weeks except current
            row += [group, priority_for(h, group)]
            out.append(row)
        pri_order = {'High': 0, 'Medium': 1, 'Standard': 2}
        def sort_key(r):
            h = hits[r[2]]
            if group == 'Just Dropped':
                basis = h[prev_key] if prev_key else 0
            else:
                basis = max((h[k] for k in (week_keys[:-2] if prev_key else week_keys[:-1])), default=0)
            return (pri_order[r[-1]], -basis)
        out.sort(key=sort_key)
        return out

    jd_rows = rows_for_group(just_dropped, 'Just Dropped')
    ls_rows = rows_for_group(long_silent, 'Long Silent')
    n_cols_s4 = len(headers_s4)
    current_row = 6
    if jd_rows:
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=n_cols_s4)
        c = ws.cell(current_row, 1, f'▼ JUST DROPPED — W{current_week-1} activity then silent in W{current_week} ({len(jd_rows)} students). Freshest drop signal but distorted by partial week.' if is_partial else f'▼ JUST DROPPED — W{current_week-1} activity then silent in W{current_week} ({len(jd_rows)} students). Freshest drop signal.')
        c.font = Font(name='Arial', size=10, bold=True, color=WHITE)
        c.fill = PatternFill('solid', start_color=RED)
        c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        ws.row_dimensions[current_row].height = 22
        current_row += 1
        write_data_rows(ws, jd_rows, start_row=current_row)
        current_row += len(jd_rows)
    if ls_rows:
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=n_cols_s4)
        c = ws.cell(current_row, 1, f'▼ LONG SILENT — Active in earlier weeks but zero in W{current_week-1} and W{current_week} ({len(ls_rows)} students). Sustained disengagement; prioritise.')
        c.font = Font(name='Arial', size=10, bold=True, color=WHITE)
        c.fill = PatternFill('solid', start_color=PURPLE)
        c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        ws.row_dimensions[current_row].height = 22
        current_row += 1
        write_data_rows(ws, ls_rows, start_row=current_row)
    autosize(ws, [22, 18, 12, 10, 10, 18, 38] + [8] * (current_week - 1) + [14, 12])
    ws.freeze_panes = 'A6'

    # ============ S5 ============
    ws, sids = make_seg_tab('S5', 'Late Arrivals')
    write_tab_header(ws, f'S5 — Late Arrivals + W{current_week-1} Returners' if prev_key else 'S5 — Late Arrivals',
        f'{len(sids)} students with zero W{current_week-1} hits but appearing in W{current_week}' if prev_key else f'{len(sids)} students',
        seg_descriptions['S5'], 12, 'S5')
    write_col_headers(ws, ['Surname', 'First Name', 'Student ID', 'Course', 'Disc. Class', 'Disc. Teacher', 'Email',
                           'Earlier Total', f'W{current_week} Hits', f'W{current_week} Daily Avg',
                           'Type', 'Watch Status'], row=5)
    rows = []
    for sid in sids:
        st = students[sid]; h = hits[sid]
        prior = sum(h[k] for k in week_keys[:-2]) if current_week >= 2 else 0
        kind = 'True late arrival' if prior == 0 else f'W{current_week-1} absentee returned'
        avg = h[curr_key] / curr_days if curr_days > 0 else 0
        if h[curr_key] >= 15: ws_status = 'Promising'
        elif h[curr_key] >= 5: ws_status = 'Monitor'
        else: ws_status = 'At Risk'
        rows.append([st['last'], st['first'], sid, st['course'], st.get('discipline_class', ''), st.get('discipline_teacher', ''), st['email'], prior, h[curr_key],
                     round(avg, 1), kind, ws_status])
    rows.sort(key=lambda r: -r[8])
    write_data_rows(ws, rows, start_row=6)
    autosize(ws, [22, 18, 12, 10, 10, 18, 38, 12, 10, 12, 22, 14])
    ws.freeze_panes = 'A6'

    # ============ S6 ============
    ws, sids = make_seg_tab('S6', 'Fading Engagers')
    write_tab_header(ws, 'S6 — Fading Engagers',
        f'{len(sids)} students whose daily-average dropped 50%+ from W{current_week-1} to W{current_week}',
        seg_descriptions['S6'], current_week + 9, 'S6')
    headers_s6 = ['Surname', 'First Name', 'Student ID', 'Course', 'Disc. Class', 'Disc. Teacher', 'Email'] + week_labels + [f'W{current_week-1} Daily', f'W{current_week} Daily']
    write_col_headers(ws, headers_s6, row=5)
    rows = []
    for sid in sids:
        st = students[sid]; h = hits[sid]
        row = [st['last'], st['first'], sid, st['course'], st.get('discipline_class', ''), st.get('discipline_teacher', ''), st['email']]
        row += [h[k] for k in week_keys]
        row += [round(h[prev_key] / prev_days, 1), round(h[curr_key] / curr_days, 1)]
        rows.append(row)
    rows.sort(key=lambda r: -(r[-2] - r[-1]))
    write_data_rows(ws, rows, start_row=6)
    autosize(ws, [22, 18, 12, 10, 10, 18, 38] + [8] * current_week + [11, 11])
    ws.freeze_panes = 'A6'

    # ============ S7 ============
    ws, sids = make_seg_tab('S7', 'Sustained Participants')
    write_tab_header(ws, 'S7 — Sustained Participants',
        (f'{len(sids)} students maintaining engagement through W{current_week} (UNDERSTATED — partial week)'
         if is_partial else f'{len(sids)} students maintaining engagement through W{current_week}'),
        seg_descriptions['S7'], current_week + 10, 'S7')
    headers_s7 = ['Surname', 'First Name', 'Student ID', 'Course', 'Disc. Class', 'Disc. Teacher', 'Email'] + week_labels + [f'W{current_week-1} Daily', f'W{current_week} Daily', 'Trend']
    write_col_headers(ws, headers_s7, row=5)
    rows = []
    for sid in sids:
        st = students[sid]; h = hits[sid]
        prev_avg = h[prev_key] / prev_days; curr_avg = h[curr_key] / curr_days
        if curr_avg > prev_avg * 1.1: trend = 'Growing'
        elif curr_avg >= prev_avg * 0.85: trend = 'Stable'
        else: trend = 'Slight Dip'
        row = [st['last'], st['first'], sid, st['course'], st.get('discipline_class', ''), st.get('discipline_teacher', ''), st['email']]
        row += [h[k] for k in week_keys]
        row += [round(prev_avg, 1), round(curr_avg, 1), trend]
        rows.append(row)
    rows.sort(key=lambda r: -r[-2])
    write_data_rows(ws, rows, start_row=6)
    autosize(ws, [22, 18, 12, 10, 10, 18, 38] + [8] * current_week + [11, 11, 12])
    ws.freeze_panes = 'A6'

    # ============ Program Leaderboard ============
    ws = wb.create_sheet('Program Leaderboard')
    write_tab_header(ws, 'Program Leaderboard',
        'Ranked by hits per enrolled (weighted engagement)',
        'Programs ordered by total hits divided by enrolled count. Gold/silver/bronze top 3.',
        6 + current_week)
    headers_pl = ['Rank', 'Course', 'Enrolled', 'Active', 'Active Rate %', 'Hits / Enrolled', 'Total Hits'] + week_labels
    write_col_headers(ws, headers_pl, row=5)
    prog = {}
    for sid, st in students.items():
        course = st['course'] or 'UNKNOWN'; h = hits[sid]
        if course not in prog:
            prog[course] = {'enr': 0, 'act': 0, 'tot': 0, **{k: 0 for k in week_keys}}
        prog[course]['enr'] += 1
        total = sum(h[k] for k in week_keys)
        if total > 0: prog[course]['act'] += 1
        prog[course]['tot'] += total
        for k in week_keys:
            prog[course][k] += h[k]
    plist = sorted(prog.items(), key=lambda kv: -(kv[1]['tot'] / max(kv[1]['enr'], 1)))
    rows = []
    for i, (course, p) in enumerate(plist, 1):
        row = [i, course, p['enr'], p['act'],
               round(100 * p['act'] / max(p['enr'], 1), 1),
               round(p['tot'] / max(p['enr'], 1), 1),
               p['tot']]
        row += [p[k] for k in week_keys]
        rows.append(row)
    write_data_rows(ws, rows, start_row=6)
    for ri in range(6, 6 + len(rows)):
        ws.cell(ri, 6).font = Font(name='Arial', size=10, bold=True)
    medals = ['FFD700', 'C0C0C0', 'CD7F32']
    for i in range(min(3, len(rows))):
        for ci in range(1, len(headers_pl) + 1):
            ws.cell(6 + i, ci).fill = PatternFill('solid', start_color=medals[i])
    autosize(ws, [6, 12, 10, 10, 12, 14, 12] + [10] * current_week)
    ws.freeze_panes = 'A6'

    # ============ Top 20 ============
    ws = wb.create_sheet('Top 20 Individual')
    write_tab_header(ws, 'Top 20 Individual Engagement',
        f'Ranked by total hits across W1-W{current_week}',
        'Top 20 students by total hits with per-week breakdown.',
        8 + current_week)
    headers_t20 = ['Rank', 'Surname', 'First Name', 'Student ID', 'Course', 'Disc. Class', 'Disc. Teacher', 'Total'] + week_labels
    write_col_headers(ws, headers_t20, row=5)
    ranked = sorted(students.keys(), key=lambda s: -sum(hits[s][k] for k in week_keys))[:20]
    rows = []
    for i, sid in enumerate(ranked, 1):
        st = students[sid]; h = hits[sid]
        row = [i, st['last'], st['first'], sid, st['course'], st.get('discipline_class', ''), st.get('discipline_teacher', ''), sum(h[k] for k in week_keys)]
        row += [h[k] for k in week_keys]
        rows.append(row)
    write_data_rows(ws, rows, start_row=6)
    autosize(ws, [6, 22, 18, 12, 10, 10, 18, 10] + [8] * current_week)
    ws.freeze_panes = 'A6'

    return wb, counts


# ===========================================================================
# PROGRAM REPORT WORKBOOK
# ===========================================================================
def build_program_workbook(subject_code, students, login, hits, seg, current_week,
                           prev_days, curr_days, is_partial, latest_date, login_window,
                           gc_data=None, gc_labels=None):
    """Build a separate workbook with one sheet per course_code (program).

    Sheet 1 = cross-tab summary (programs × segments).
    Sheets 2+ = one per program, all students sorted by segment then surname,
    with the full column set.
    """
    wb = Workbook()
    wb.remove(wb.active)

    week_keys = [f'w{i}' for i in range(1, current_week + 1)]
    week_labels = [f'W{i}' for i in range(1, current_week + 1)]
    curr_key = f'w{current_week}'
    prev_key = f'w{current_week - 1}' if current_week > 1 else None
    seg_codes = [f'S{i}' for i in range(1, 8)]

    # Group students by course_code
    programs = {}
    for sid, st in students.items():
        cc = st.get('course_code') or st.get('course') or 'UNKNOWN'
        programs.setdefault(cc, []).append(sid)

    # ============ Summary cross-tab ============
    ws = wb.create_sheet('Summary')
    partial_note = ' (PARTIAL)' if is_partial else ''
    write_tab_header(ws,
        f'{subject_code} Program Report — W{current_week}{partial_note}',
        f'Enrolled {len(students)}  •  {len(programs)} programs  •  Latest data: {latest_date.strftime("%b %#d %Y")}',
        'Segment counts per program (course code). One sheet per program follows.',
        len(seg_codes) + 3)

    headers = ['Program', 'Enrolled'] + seg_codes + ['Active Rate %']
    write_col_headers(ws, headers, row=5)

    prog_order = sorted(programs.keys(), key=lambda cc: -len(programs[cc]))
    rows = []
    for cc in prog_order:
        sids = programs[cc]
        enrolled = len(sids)
        counts_row = [cc, enrolled]
        active = 0
        for sc in seg_codes:
            n = sum(1 for sid in sids if seg.get(sid) == sc)
            counts_row.append(n)
            if sc not in ('S1', 'S2'):
                active += n
        counts_row.append(round(100 * active / max(enrolled, 1), 1))
        rows.append(counts_row)

    # Totals row
    totals = ['TOTAL', len(students)]
    for i, sc in enumerate(seg_codes):
        totals.append(sum(r[i + 2] for r in rows))
    total_active = sum(r[-1] * r[1] / 100 for r in rows)
    totals.append(round(100 * total_active / max(len(students), 1), 1))
    rows.append(totals)

    write_data_rows(ws, rows, start_row=6)

    # Colour segment column headers
    for ci, sc in enumerate(seg_codes, 3):
        fill_colour = SEG_COLOURS.get(sc)
        if fill_colour:
            ws.cell(5, ci).fill = PatternFill('solid', start_color=fill_colour)
            ws.cell(5, ci).font = Font(name='Arial', size=10, bold=True, color=WHITE)

    # Bold totals row
    totals_row = 6 + len(rows) - 1
    for ci in range(1, len(headers) + 1):
        ws.cell(totals_row, ci).font = Font(name='Arial', size=10, bold=True)
        ws.cell(totals_row, ci).fill = PatternFill('solid', start_color=LIGHT)
        ws.cell(totals_row, ci).border = thin_border()

    autosize(ws, [12, 10] + [8] * len(seg_codes) + [14])
    ws.freeze_panes = 'A6'

    # ============ Segment Legend ============
    legend_start = totals_row + 2
    ws.merge_cells(start_row=legend_start, start_column=1,
                   end_row=legend_start, end_column=3)
    c = ws.cell(legend_start, 1, 'Segment Legend')
    c.font = Font(name='Arial', size=11, bold=True, color=WHITE)
    c.fill = PatternFill('solid', start_color=NAVY)
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws.row_dimensions[legend_start].height = 22

    seg_legend = [
        ('S1', 'Never Engaged'),
        ('S2', 'Pre-Teaching Ghosts'),
        ('S3', 'W1 Early Drop-Offs'),
        ('S4', f'Active then W{current_week} Absent'),
        ('S5', 'Late Arrivals + Returners'),
        ('S6', 'Fading Engagers'),
        ('S7', 'Sustained Participants'),
    ]
    for li, (code, label) in enumerate(seg_legend):
        row_i = legend_start + 1 + li
        c_code = ws.cell(row_i, 1, code)
        c_code.font = Font(name='Arial', size=10, bold=True, color=WHITE)
        fill_colour = SEG_COLOURS.get(code, ACCENT)
        c_code.fill = PatternFill('solid', start_color=fill_colour)
        c_code.alignment = Alignment(horizontal='center', vertical='center')
        c_code.border = thin_border()
        ws.merge_cells(start_row=row_i, start_column=2, end_row=row_i, end_column=3)
        c_label = ws.cell(row_i, 2, label)
        c_label.font = Font(name='Arial', size=10)
        c_label.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        c_label.border = thin_border()
        ws.cell(row_i, 3).border = thin_border()

    next_row = legend_start + 1 + len(seg_legend) + 1

    # ============ Assessment Submission Breakdown Tables ============
    gc_active = gc_data is not None and gc_labels is not None and len(gc_labels) > 0
    if gc_active:
        def categorise_grade(val):
            """Categorise a grade value into one of the summary buckets."""
            if val == 'No Submission':
                return 'No Submission'
            v = val.strip().lower()
            if v in ('', 'no submission'):
                return 'No Submission'
            if 'needs grading' in v or 'needs marking' in v:
                return 'Needs Marking'
            if 'unsatisf' in v:
                return 'Unsatisfactory'
            # Anything else (Satisfactory, numeric scores, etc.) = Satisfactory
            return 'Satisfactory'

        status_cols = ['Satisfactory', 'Unsatisfactory', 'Needs Marking', 'No Submission']

        for ai, label in enumerate(gc_labels):
            # Section header
            ws.merge_cells(start_row=next_row, start_column=1,
                           end_row=next_row, end_column=len(status_cols) + 2)
            c = ws.cell(next_row, 1, label.replace('AS', 'Assessment '))
            c.font = Font(name='Arial', size=11, bold=True, color=WHITE)
            c.fill = PatternFill('solid', start_color=NAVY)
            c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
            ws.row_dimensions[next_row].height = 22
            next_row += 1

            # Column headers
            as_headers = ['Program'] + status_cols + ['TOTAL']
            for ci, h in enumerate(as_headers, 1):
                c = ws.cell(next_row, ci, h)
                c.font = Font(name='Arial', size=10, bold=True, color=WHITE)
                c.fill = PatternFill('solid', start_color=ACCENT)
                c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                c.border = thin_border()
            ws.row_dimensions[next_row].height = 30
            next_row += 1

            # Per-program rows
            grand_totals = {s: 0 for s in status_cols}
            grand_total_all = 0
            as_data_rows = []
            for cc in prog_order:
                sids = programs[cc]
                bucket = {s: 0 for s in status_cols}
                for sid in sids:
                    sg = gc_data.get(sid, {})
                    val = sg.get(label, 'No Submission')
                    cat = categorise_grade(val)
                    bucket[cat] += 1
                row_total = sum(bucket.values())
                as_data_rows.append([cc] + [bucket[s] for s in status_cols] + [row_total])
                for s in status_cols:
                    grand_totals[s] += bucket[s]
                grand_total_all += row_total

            # Write data rows with alternating fill
            for ri, row_data in enumerate(as_data_rows):
                fill = PatternFill('solid', start_color=ALT_ROW) if ri % 2 == 0 else None
                for ci, val in enumerate(row_data, 1):
                    c = ws.cell(next_row, ci, val)
                    c.font = Font(name='Arial', size=10)
                    if fill:
                        c.fill = fill
                    c.alignment = Alignment(horizontal='center' if ci > 1 else 'left',
                                            vertical='center')
                    c.border = thin_border()
                next_row += 1

            # TOTAL row
            total_row_data = ['TOTAL'] + [grand_totals[s] for s in status_cols] + [grand_total_all]
            for ci, val in enumerate(total_row_data, 1):
                c = ws.cell(next_row, ci, val)
                c.font = Font(name='Arial', size=10, bold=True)
                c.fill = PatternFill('solid', start_color=LIGHT)
                c.border = thin_border()
                c.alignment = Alignment(horizontal='center' if ci > 1 else 'left',
                                        vertical='center')
            next_row += 2  # blank row before next assessment table

    # ============ Per-program sheets ============
    seg_order = {f'S{i}': i for i in range(1, 8)}

    base_headers = ['Segment', 'Surname', 'First Name', 'Student ID',
                    'Course', 'Disc. Subject', 'Disc. Class', 'Disc. Teacher', 'Email'] + \
                   week_labels + ['Total Hits']
    if prev_key:
        base_headers += [f'W{current_week-1} Daily', f'W{current_week} Daily']
    base_headers += ['Total Logins', 'Last Login', 'Days Since']

    # Append grade centre assessment columns if available
    gc_active = gc_data is not None and gc_labels is not None and len(gc_labels) > 0
    if gc_active:
        base_headers += gc_labels

    for cc in prog_order:
        sids = programs[cc]
        sheet_name = str(cc)[:31] or 'UNKNOWN'
        # Avoid duplicate sheet names
        existing = [s.title for s in wb.worksheets]
        if sheet_name in existing:
            sheet_name = (sheet_name[:28] + '_2')[:31]

        ws_p = wb.create_sheet(sheet_name)

        seg_counts_str = ', '.join(
            f'{sc}: {sum(1 for sid in sids if seg.get(sid) == sc)}'
            for sc in seg_codes
        )
        write_tab_header(ws_p,
            f'Program {cc} — {len(sids)} students',
            seg_counts_str,
            f'All students sorted by segment then surname. W{current_week}{"(partial)" if is_partial else ""}.',
            len(base_headers))
        write_col_headers(ws_p, base_headers, row=5)

        # Sort by segment order, then surname
        sorted_sids = sorted(sids, key=lambda sid: (
            seg_order.get(seg.get(sid, 'S1'), 99),
            students[sid]['last'].lower(),
            students[sid]['first'].lower(),
        ))

        rows_p = []
        for sid in sorted_sids:
            st = students[sid]; h = hits[sid]; l = login.get(sid)
            days_since, last_str, total_logins = fmt_login(l)
            total_hits = sum(h[k] for k in week_keys)
            row = [seg.get(sid, '?'), st['last'], st['first'], sid,
                   st['course'], st.get('discipline_subject', ''),
                   st.get('discipline_class', ''), st.get('discipline_teacher', ''),
                   st['email']]
            row += [h[k] for k in week_keys]
            row += [total_hits]
            if prev_key:
                prev_avg = round(h[prev_key] / prev_days, 1) if prev_days else 0
                curr_avg = round(h[curr_key] / curr_days, 1) if curr_days else 0
                row += [prev_avg, curr_avg]
            row += [total_logins, last_str, days_since]
            if gc_active:
                sg = gc_data.get(sid, {})
                for label in gc_labels:
                    row.append(sg.get(label, 'No Submission'))
            rows_p.append(row)

        write_data_rows(ws_p, rows_p, start_row=6)

        # Colour the segment column
        for ri, sid in enumerate(sorted_sids):
            seg_code = seg.get(sid, '')
            fill_colour = SEG_COLOURS.get(seg_code)
            if fill_colour:
                ws_p.cell(6 + ri, 1).fill = PatternFill('solid', start_color=fill_colour)
                ws_p.cell(6 + ri, 1).font = Font(name='Arial', size=10, bold=True, color=WHITE)

        # Colour "No Submission" cells red in assessment columns
        if gc_active:
            no_sub_fill = PatternFill('solid', start_color='FADBD8')
            no_sub_font = Font(name='Arial', size=10, color=RED, bold=True)
            needs_grading_fill = PatternFill('solid', start_color='FEF9E7')
            gc_start_col = len(base_headers) - len(gc_labels) + 1  # 1-indexed
            for ri in range(len(sorted_sids)):
                for ci_offset in range(len(gc_labels)):
                    cell = ws_p.cell(6 + ri, gc_start_col + ci_offset)
                    if cell.value == 'No Submission':
                        cell.fill = no_sub_fill
                        cell.font = no_sub_font
                    elif isinstance(cell.value, str) and 'needs grading' in cell.value.lower():
                        cell.fill = needs_grading_fill

        widths = [8, 22, 18, 12, 10, 24, 10, 18, 38] + [8] * current_week + [10]
        if prev_key:
            widths += [11, 11]
        widths += [12, 14, 12]
        if gc_active:
            widths += [16] * len(gc_labels)
        autosize(ws_p, widths)
        ws_p.freeze_panes = 'A6'

    return wb

# ===========================================================================
# STREAMLIT UI
# ===========================================================================
st.set_page_config(page_title='WSUTC Engagement Report', layout='wide', page_icon='📊')
st.title('WSUTC Student Engagement Report')
st.caption('Upload Blackboard exports for one subject. Generates a 10-tab Excel workbook with engagement segmentation.')

with st.expander('Instructions', expanded=False):
    st.markdown("""
- Upload the **class list** (.xls or .xlsx), the **login report** (.xlsx), and **all relevant usage report files** (.xls).
- **Enriched class list (.xlsx):** If your class list includes `Course Code`, `Class Code`, and `Teacher` columns, these will appear in all tabs and a **Program Report** download becomes available (one sheet per program with segment tags).
- **Grade Centre (.xls, optional):** If uploaded, assessment submission status columns (AS1, AS2, …) are appended to each per-program sheet. Cells show the grade, "Needs Grading", or "No Submission".
- Usage files can overlap; the most recent data wins where they do.
- The current teaching week is auto-detected from the latest day with data in the usage files.
- If the current week is partial, S4 will be flagged as inflated and S7 as understated.
- Defaults: S2 ≥ pre-Mar 2 days, S3 = W1 login range, S4 split into Just Dropped + Long Silent, exclusions per project spec.
""")

col1, col2 = st.columns(2)
with col1:
    classlist_file = st.file_uploader('Class list (.xls / .xlsx)', type=['xls', 'xlsx'], key='cl')
    login_file = st.file_uploader('Login report (.xlsx)', type=['xlsx'], key='lr')
with col2:
    usage_files = st.file_uploader('Usage report files (.xls) — upload all that apply',
                                    type=['xls'], accept_multiple_files=True, key='uf')
    gc_file = st.file_uploader('Grade Centre (.xls) — optional', type=['xls'], key='gc')

run_btn = st.button('Generate report', type='primary', disabled=not (classlist_file and login_file and usage_files))

if run_btn:
    try:
        with st.spinner('Loading class list...'):
            subject_code, students = load_classlist(classlist_file.getvalue())
        st.success(f'**{subject_code}** • {len(students)} enrolled (after exclusions)')

        with st.spinner('Loading login report...'):
            login, win_start, win_end = load_login_report(login_file.getvalue())
        if win_start and win_end:
            login_window_str = f'{win_start.strftime("%b %#d")} - {win_end.strftime("%b %#d %Y")}'
            st.info(f'Login window detected: **{login_window_str}**')
        else:
            login_window_str = 'unknown (could not parse)'
            st.warning('Could not auto-detect login window from report. Check the file format.')

        with st.spinner(f'Parsing {len(usage_files)} usage file(s)...'):
            merged = merge_usage_files([f.getvalue() for f in usage_files])
        st.write(f'Parsed {len(merged)} unique day-records from usage files.')

        # Cap the latest data day at the login window end if known.
        # Stray usage data past the login window (e.g. midnight extraction bleed) is ignored.
        cap_latest = win_end if win_end else None
        current_week, days_in, latest = detect_current_week(merged, override_latest=cap_latest)
        if current_week is None:
            st.error('No usage data found on or after Mar 2 2026. Cannot determine current week.')
            st.stop()
        is_partial = days_in < 7
        curr_days = days_in
        prev_days = 7
        if current_week == 1:
            prev_days = curr_days  # no comparison possible

        partial_msg = f' (PARTIAL — {curr_days} days)' if is_partial else ' (full)'
        st.success(f'Detected current week: **W{current_week}**{partial_msg}  •  Latest data: **{latest.strftime("%b %#d %Y")}**')
        if current_week == 1:
            st.warning('W1 is the only week with data. S5/S6/S7 will be empty (no comparison week available).')

        with st.spinner('Computing weekly hits...'):
            hits = bucket_by_week(merged, students, current_week, max_date=latest)

        with st.spinner('Classifying students...'):
            seg, s2_thresh, s3_rng = classify(students, login, hits, current_week, prev_days, curr_days)

        # Show segment counts
        counts = {f'S{i}': 0 for i in range(1, 8)}
        for s in seg.values():
            counts[s] = counts.get(s, 0) + 1
        assert sum(counts.values()) == len(students), \
            f'Segment sum {sum(counts.values())} != enrolled {len(students)}'

        st.subheader('Segment counts')
        cols = st.columns(7)
        for i, (code, _) in enumerate([
            ('S1', 'Never'), ('S2', 'Ghosts'), ('S3', 'W1 drop'),
            ('S4', f'W{current_week} absent'), ('S5', 'Late/return'),
            ('S6', 'Fading'), ('S7', 'Sustained')]):
            cols[i].metric(code, counts[code], f'{counts[code]/len(students)*100:.1f}%')

        # Standing checks
        st.subheader('Standing checks')
        missing = sum(1 for s in students if s not in login)
        st.write(f'• Enrolled: **{len(students)}**')
        st.write(f'• Missing from login report: **{missing}**')
        st.write(f'• S2 days-since threshold: **≥ {s2_thresh}**')
        st.write(f'• S3 days-since range: **{s3_rng[0]}-{s3_rng[1]}**')
        st.write(f'• Comparison: W{current_week-1} (7d) vs W{current_week} ({curr_days}d)' if current_week > 1 else '• No prior week to compare')

        with st.spinner('Building workbook...'):
            wb, _ = build_workbook(
                subject_code, students, login, hits, seg,
                current_week, prev_days, curr_days, is_partial, latest,
                login_window_str, s2_thresh, s3_rng,
            )
            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)

        date_str = latest.strftime('%Y%m%d')
        suffix = '_Partial' if is_partial else ''
        filename = f'{subject_code}_Engagement_Report_W{current_week}_{date_str}{suffix}.xlsx'
        st.download_button('⬇ Download workbook', data=buf, file_name=filename,
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                           type='primary')

        # Program report — only available when enriched classlist provides course_code
        has_programs = any(st_data.get('course_code') for st_data in students.values())

        # Parse grade centre if uploaded
        gc_data = None
        gc_labels = None
        if gc_file is not None:
            with st.spinner('Parsing grade centre...'):
                gc_data, gc_labels = load_grade_centre(gc_file.getvalue())
            matched = sum(1 for sid in students if sid in gc_data)
            st.info(f'Grade Centre: detected **{len(gc_labels)}** assessments ({", ".join(gc_labels)}) • matched **{matched}/{len(students)}** students')

        if has_programs:
            with st.spinner('Building program report...'):
                wb_prog = build_program_workbook(
                    subject_code, students, login, hits, seg,
                    current_week, prev_days, curr_days, is_partial, latest,
                    login_window_str,
                    gc_data=gc_data, gc_labels=gc_labels,
                )
                buf_prog = io.BytesIO()
                wb_prog.save(buf_prog)
                buf_prog.seek(0)
            prog_filename = f'{subject_code}_Program_Report_W{current_week}_{date_str}{suffix}.xlsx'
            st.download_button('⬇ Download program report', data=buf_prog, file_name=prog_filename,
                               mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                               key='prog_dl')

    except Exception as e:
        st.error(f'Error: {e}')
        import traceback
        st.code(traceback.format_exc())
