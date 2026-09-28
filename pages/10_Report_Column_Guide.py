"""
Report Column Guide
===================

What each column of the Reregistration Advisory report means, for coaches,
team leads and the data team. The page itself is docs/rereg_column_guide.html
(the same content as docs/rereg_report_columns.md), shown here so anyone who
can open the app can read it.
"""
from __future__ import annotations

from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

GUIDE = Path(__file__).resolve().parent.parent / "docs" / "rereg_column_guide.html"

st.set_page_config(page_title="Report Column Guide", layout="wide")

components.html(GUIDE.read_text(encoding="utf-8"), height=1100, scrolling=True)
