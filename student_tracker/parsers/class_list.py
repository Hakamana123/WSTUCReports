"""Parse the WSU class list .xls file (real CFB format, requires xlrd)."""

import pandas as pd


HEADER_ROW = 6  # 0-indexed; 'student_code' header sits here


def parse(path: str) -> pd.DataFrame:
    """Return a tidy DataFrame keyed by student_code (string).

    Columns: student_code, first_name, last_name, preferred_name,
             attend_type, course, course_type, email_address, major,
             display_subject_code
    """
    raw = pd.read_excel(path, header=None, engine="xlrd")
    headers = raw.iloc[HEADER_ROW].tolist()
    df = raw.iloc[HEADER_ROW + 1:].copy()
    df.columns = headers
    df = df.dropna(subset=["student_code"]).reset_index(drop=True)
    df["student_code"] = df["student_code"].astype(str).str.strip()
    for col in ["first_name", "last_name", "preferred_name",
                "attend_type", "course_type", "major"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    return df


def detect_subject_code(df: pd.DataFrame) -> str | None:
    """Return the display_subject_code if present and consistent."""
    if "display_subject_code" not in df.columns:
        return None
    codes = df["display_subject_code"].dropna().astype(str).str.strip().unique()
    return codes[0] if len(codes) == 1 else None


def filter_for_real_students(df: pd.DataFrame,
                              exclude_id_prefixes: list[str] | None = None,
                              exclude_surnames: list[str] | None = None
                              ) -> tuple[pd.DataFrame, dict]:
    """Drop staff, preview, and explicitly-excluded accounts.

    Returns (filtered_df, stats_dict). stats_dict reports how many rows
    were dropped by each rule, for surfacing in the UI's validation panel.
    """
    out = df.copy()
    stats = {"by_prefix": 0, "by_surname": 0, "kept": 0, "started": len(out)}
    if exclude_id_prefixes:
        prefixes = tuple(str(p) for p in exclude_id_prefixes)
        mask = out["student_code"].str.startswith(prefixes)
        stats["by_prefix"] = int(mask.sum())
        out = out[~mask]
    if exclude_surnames and "last_name" in out.columns:
        mask = out["last_name"].isin(exclude_surnames)
        stats["by_surname"] = int(mask.sum())
        out = out[~mask]
    out = out.reset_index(drop=True)
    stats["kept"] = len(out)
    return out, stats
