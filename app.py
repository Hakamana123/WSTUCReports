import streamlit as st

st.set_page_config(page_title="WSUTC Reports", layout="wide", page_icon="📊")

st.title("WSUTC Reports")
st.markdown("Select a report from the sidebar to get started.")

st.divider()
st.link_button("🧭 Open IA Mapping / Workload Management ↗", "https://wsutcreportmem.streamlit.app/")
