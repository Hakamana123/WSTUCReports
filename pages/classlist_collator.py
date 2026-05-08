import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="APP Class List Collator", layout="wide")

st.title("APP Class List Collator")
st.caption(
    "Upload discipline subject class lists and GEDU prep subject class lists. "
    "Students are cross-referenced by Student ID to tag each GEDU student "
    "with their discipline subject, class, and teacher."
)


def find_header_row(df: pd.DataFrame) -> int | None:
    for i in range(min(10, len(df))):
        row_vals = [str(v).strip().lower() for v in df.iloc[i] if pd.notna(v)]
        if "student_code" in row_vals:
            return i
    return None


def parse_classlist(file) -> pd.DataFrame:
    """Parse a single .xls class list file and return a DataFrame of students."""
    xls = pd.ExcelFile(file, engine="xlrd")
    all_rows = []

    for sheet_name in xls.sheet_names:
        if "unallocated" in sheet_name.lower():
            continue

        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        if len(df) < 7:
            continue

        # Metadata from header block
        subject_code_raw = (
            str(df.iloc[1, 0]).strip() if pd.notna(df.iloc[1, 0]) else ""
        )
        class_session_code = (
            str(df.iloc[2, 0]).strip() if pd.notna(df.iloc[2, 0]) else ""
        )
        class_code = re.sub(r"-P\d+$", "", class_session_code)

        staff_raw = str(df.iloc[4, 1]).strip() if pd.notna(df.iloc[4, 1]) else ""
        teacher_name = re.sub(r"^Staff:\s*", "", staff_raw)

        header_idx = find_header_row(df)
        if header_idx is None:
            continue

        headers = {
            str(df.iloc[header_idx, c]).strip().lower(): c
            for c in range(df.shape[1])
            if pd.notna(df.iloc[header_idx, c])
        }

        for i in range(header_idx + 1, len(df)):
            row = df.iloc[i]
            sid_col = headers.get("student_code")
            if sid_col is None or pd.isna(row.iloc[sid_col]):
                continue

            student_id = row.iloc[sid_col]
            student_id = (
                str(int(student_id))
                if isinstance(student_id, float)
                else str(student_id).strip()
            )

            def get_val(col_name: str) -> str:
                col_idx = headers.get(col_name)
                if col_idx is None:
                    return ""
                val = row.iloc[col_idx]
                if pd.isna(val):
                    return ""
                return str(val).strip()

            course_val = get_val("course")
            try:
                course_val = str(int(float(course_val)))
            except (ValueError, TypeError):
                pass

            all_rows.append(
                {
                    "First Name": get_val("first_name"),
                    "Last Name": get_val("last_name"),
                    "Student ID": student_id,
                    "Course Code": course_val,
                    "Subject Code": subject_code_raw,
                    "Class Code": class_code,
                    "Email": get_val("email_address"),
                    "Teacher": teacher_name,
                }
            )

    result = pd.DataFrame(all_rows)
    # Deduplicate sessions
    if not result.empty:
        result = result.drop_duplicates(subset=["Student ID", "Subject Code", "Class Code"])
    return result


# ── Upload zones ──

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Discipline Subject Class Lists")
    discipline_files = st.file_uploader(
        "Upload discipline .xls files (BEHV, COMM, etc.)",
        type=["xls"],
        accept_multiple_files=True,
        key="discipline",
    )

with col2:
    st.subheader("2. GEDU Prep Subject Class Lists")
    gedu_files = st.file_uploader(
        "Upload GEDU0016 and/or GEDU0017 .xls files",
        type=["xls"],
        accept_multiple_files=True,
        key="gedu",
    )

# ── Processing ──

if discipline_files and gedu_files:
    # Parse discipline files
    disc_frames = []
    disc_summary = []
    for f in discipline_files:
        try:
            result = parse_classlist(f)
            disc_frames.append(result)
            disc_summary.append({"File": f.name, "Students": len(result), "Status": "OK"})
        except Exception as e:
            disc_summary.append({"File": f.name, "Students": 0, "Status": f"Error: {e}"})

    # Parse GEDU files
    gedu_frames = []
    gedu_summary = []
    for f in gedu_files:
        try:
            result = parse_classlist(f)
            gedu_frames.append(result)
            gedu_summary.append({"File": f.name, "Students": len(result), "Status": "OK"})
        except Exception as e:
            gedu_summary.append({"File": f.name, "Students": 0, "Status": f"Error: {e}"})

    # Show file summaries
    st.subheader("Files Processed")
    sum_col1, sum_col2 = st.columns(2)
    with sum_col1:
        st.caption("Discipline files")
        st.dataframe(pd.DataFrame(disc_summary), use_container_width=True, hide_index=True)
    with sum_col2:
        st.caption("GEDU files")
        st.dataframe(pd.DataFrame(gedu_summary), use_container_width=True, hide_index=True)

    if disc_frames and gedu_frames:
        disc_all = pd.concat(disc_frames, ignore_index=True)
        gedu_all = pd.concat(gedu_frames, ignore_index=True)

        # Cross-reference: join GEDU students with their discipline info
        merged = gedu_all.merge(
            disc_all,
            on="Student ID",
            how="left",
            suffixes=("_GEDU", "_Discipline"),
        )

        # Build output
        output = pd.DataFrame(
            {
                "First Name": merged["First Name_GEDU"],
                "Last Name": merged["Last Name_GEDU"],
                "Student ID": merged["Student ID"],
                "Course Code": merged["Course Code_GEDU"],
                "Email": merged["Email_GEDU"].where(
                    merged["Email_GEDU"] != "", merged.get("Email_Discipline", "")
                ),
                "GEDU Subject": merged["Subject Code_GEDU"],
                "GEDU Class": merged["Class Code_GEDU"],
                "GEDU Teacher": merged["Teacher_GEDU"],
                "Discipline Subject": merged["Subject Code_Discipline"].fillna("No match"),
                "Discipline Class": merged["Class Code_Discipline"].fillna(""),
                "Discipline Teacher": merged["Teacher_Discipline"].fillna(""),
            }
        )

        output = output.sort_values(
            ["GEDU Subject", "Discipline Subject", "Last Name", "First Name"]
        ).reset_index(drop=True)

        # Count unmatched
        unmatched = (output["Discipline Subject"] == "No match").sum()

        # ── Display ──
        st.subheader("Collated Report")

        # Filter by GEDU subject
        gedu_subjects = sorted(output["GEDU Subject"].unique())
        selected = st.multiselect(
            "Filter by GEDU subject",
            options=gedu_subjects,
            default=gedu_subjects,
        )
        filtered = output[output["GEDU Subject"].isin(selected)]

        st.caption(
            f"{len(filtered)} students shown"
            + (f" · {unmatched} with no discipline match" if unmatched > 0 else "")
        )
        st.dataframe(filtered, use_container_width=True, hide_index=True)

        # ── Export ──
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            # One sheet per GEDU subject + a combined sheet
            output.to_excel(writer, index=False, sheet_name="All Students")
            for subj in gedu_subjects:
                sheet_name = subj.replace("_", " ")[:31]  # Excel 31-char limit
                subset = output[output["GEDU Subject"] == subj]
                subset.to_excel(writer, index=False, sheet_name=sheet_name)
        buffer.seek(0)

        st.download_button(
            label="Download .xlsx",
            data=buffer,
            file_name="APP_collated_classlist.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.officedocument",
        )

elif discipline_files and not gedu_files:
    st.info("Now upload the GEDU class list files to cross-reference.")
elif gedu_files and not discipline_files:
    st.info("Now upload the discipline subject class list files to cross-reference.")
else:
    st.info("Upload discipline and GEDU class list files to get started.")
