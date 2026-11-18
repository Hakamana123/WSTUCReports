"""
Reregistration Advisory
=======================

Grant's testing calculator picks each student's next-session subjects; the v2
rule-tree wraps it with the progression-standing rules and the Coach View.

  1. Upload the progression workbook (sheet 'Query1').
  2. Pick the planning session you're advising for.
  3. For every student the tool asks Grant's calculator for the Prep / Block
     1-4 picks + Earliest Completion, then applies the 30cp Conditional
     Enrolment cap / Exclusion / At Risk rules on top. Where the calculator
     has no row (a session it doesn't cover, program 9034, or a status combo
     Grant's tool also can't resolve) it falls back to the v2 rule-tree and
     says so in 'Advice Source'.
  4. Download the same workbook back with the advice columns + a readable
     'Coach View' sheet.

Engines: student_tracker/rereg_merged.py (wiring), rereg_calc.py (Grant's
calculator port), rereg_advice.py (v2 rule-tree + Coach View).
"""

from __future__ import annotations

import io

import pandas as pd
import streamlit as st

from student_tracker import rereg_merged as rm

st.set_page_config(page_title="Reregistration Advisory", layout="wide")

# What each report column means (docs/rereg_column_guide.html).
st.sidebar.page_link("pages/10_Report_Column_Guide.py", label="Report column guide")

st.title("Reregistration Advisory")
st.caption(
    "Upload the progression file and pick the target session. Grant's "
    "calculator drives the subject picks (offering pattern carried forward "
    "for sessions after 26 AUT); standing rules (30cp cap, Exclusion) and the "
    "Coach View are layered on top."
)

with st.expander("How the advice is built", expanded=False):
    st.markdown(
        """
- **Target session** is the session you're advising *for*. The teaching year
  runs `AUT -> SPR -> SUM`; Autumn and Spring each move a student four subjects
  along their pattern, Summer is catch-up only.
- **Subjects** come from Grant's testing calculator: offering pattern +
  fail-status -> Prep / Block 1-4 picks, with timetable clashes, "positions
  1 & 2 run every Summer", and elective placement all handled inside it.
- **Which pattern**: Grant computed `26 AUT` and `25 SUM` exactly. For any
  later Autumn / Spring the `26 AUT` pattern is **assumed to still hold**;
  for any later Summer, `25 SUM` — unless the reference data is updated. Those
  rows are tagged *"… pattern assumed"* and carry no completion estimate.
- **Rule-tree fallback** — used only when the calculator genuinely has no
  answer (a status combo Grant's own tool misses, program 9034, part-way
  targets, or Nursing outside Autumn). It places the student by cohort
  position and advises earliest-outstanding-first.
- **Standing** is applied on top: `Conditional Enrolment` caps the load at
  30cp (3 subjects, prep -> Summer). `Exclusion` gets no advice. `At Risk` /
  blank get the full load.
- **Advice Source** on each row names the engine and, where relevant, which
  pattern was assumed or why the calculator was skipped.
- **Custom target** — tick the box to advise part-way through a session
  (e.g. "as of Autumn Block 3" leaves only 2 blocks to fill; these always
  use the rule-tree).
        """
    )

col_a, col_b = st.columns([2, 3])
with col_a:
    named = st.selectbox(
        "Target session (advising for)",
        rm.PLANNING_SESSIONS,
        index=rm.PLANNING_SESSIONS.index(rm.DEFAULT_PLANNING_SESSION),
    )
    part_block = st.checkbox(
        "Mid-semester — advise from Block 3",
        value=False,
        help="Only Blocks 3–4 are registered now; Blocks 1–2 (already done this "
             "session) show greyed with what the student took and how it went: "
             "✓ passed, ✗ failed, 0 not enrolled. A commencing student with no "
             "pass and at least one fail is flagged to withdraw; no enrolment or "
             "nothing to withdraw from -> check enrolment; a grade still to come "
             "-> awaiting grade. Paused students are never flagged.",
    )
    session = f"{named} Block 3" if part_block else named
    st.caption(f"Target: **{session}**")
with col_b:
    uploaded = st.file_uploader(
        "Progression workbook (.xlsx, sheet 'Query1')", type=["xlsx"], key="rereg_upload"
    )

summer_subjects = None
summer_offering = {}
is_summer = "SUM" in session.upper()
summer_round = ""
engine_view = False
_ENGINE_VIEW = "Full re-registration advice (engine view)"
if is_summer:
    # Summer advice goes out in rounds, one per set of coaching templates; the
    # page is organised the same way so each round's download is its own thing.
    pick = st.radio(
        "Summer round",
        [*rm.ROUND_LABELS.values(), _ENGINE_VIEW],
        help="**Round 1** goes out by mid-SB3 on the latest extract: current enrolments count as "
             "passed and a free SB4 slot as used; students are listed if Summer as a whole "
             "helps, SU1 + prep are advised and SU2 picks are shown but not to register yet. "
             "**Round 2** runs on the end-of-SB4 extract: actual results, SU1 reconfirmed and "
             "SU2 advised, core subjects only (prep advice follows Spring results in Round 3), "
             "with what each student has registered beside the advice; plus the 2b list of "
             "students who failed and need Autumn. "
             "**Round 3** runs on the post-Spring-results extract: the prep advice for students "
             "who still owe GEDU0016 / GEDU0017, laid over what they've registered - the prep "
             "takes priority. "
             "**Engine view** is the normal per-student Summer advice workbook.",
    )
    if pick == _ENGINE_VIEW:
        engine_view = True
        stage = st.radio(
            "Results as at", ["mid-SB3 (enrolments assumed to pass)", "after SB4 results (core only)",
                              "after Spring results (preps too)"],
            horizontal=True,
        )
        summer_round = {0: rm.ROUND_1, 1: rm.ROUND_2, 2: rm.ROUND_3}[
            ["mid-SB3", "after SB4", "after Spring"].index(stage.split(" (")[0])]
    else:
        summer_round = next(k for k, v in rm.ROUND_LABELS.items() if v == pick)
    summer_files = st.file_uploader(
        "Summer offering list(s) (.xlsx or .csv) — the SU1 and SU2 subject-list "
        "exports as they come, or the offering template (subject code, SU1/SU2, "
        "campus codes BK CA KW PC LP ON). A subject can run in both blocks.",
        type=["xlsx", "csv"], key="summer_upload", accept_multiple_files=True,
    )
    if summer_files:
        summer_offering = rm.read_summer_offering_campus(summer_files)
        # the full engine now takes the campus/block dict so it can place each
        # subject in its Summer block (SU1/SU2) at the student's campus
        summer_subjects = summer_offering or None
        st.caption(
            f"Summer offering: {len(summer_subjects)} subjects "
            f"({', '.join(sorted(summer_subjects)[:8])}{'…' if len(summer_subjects) > 8 else ''})"
            if summer_subjects else "Couldn't read any subject codes from that file."
        )
    elif not engine_view:
        st.warning("Upload the confirmed Summer offering list to build this round's lists.")
    else:
        st.caption(
            "No Summer offering list — assuming only prep + Subjects 1 & 2 run. "
            "Upload the list once it's known for real advice."
        )

if uploaded is None:
    st.info("Upload the workbook to generate advice.")
    st.stop()

try:
    df = rm.load_progression_file(io.BytesIO(uploaded.getvalue()))
except ValueError as exc:
    st.error(f"Couldn't read this file: {exc}")
    st.stop()

renamed = df.attrs.get("renamed_columns", {})
if renamed:
    st.caption("Renamed columns read as the tool expects them: "
               + ", ".join(f"*{a}* → *{c}*" for a, c in renamed.items()) + ".")
if "Electives Needed" not in df.columns:
    st.warning("This file has no *Electives Needed* (or *Electives Required*) column, so every "
               "student reads as owing **no electives**. Check the extract before sending advice.")

# --- Summer registration check -------------------------------------------------
# What students are actually registered in, where the extract carries it. Needs
# only the Summer columns, so it runs before the check on the Status columns.
if rm.has_summer_registrations(df):
    reg_check = rm.summer_registration_check(df)
    flagged = rm.summer_registration_flagged(reg_check)
    is_ce = flagged[rm.REG_CHECK_COL].str.startswith("Conditional Enrolment")
    st.header("Summer registration check")
    r1, r2, r3, r4 = st.columns(4)
    r1.metric("Registered in Summer", f"{len(reg_check):,}")
    r2.metric(f"Over {rm.SUMMER_NO_APPROVAL_CP}cp", f"{int((reg_check[rm.REG_CP_COL] > rm.SUMMER_NO_APPROVAL_CP).sum()):,}")
    r3.metric(f"Over the {rm.SUMMER_CAP_CP}cp limit",
              f"{int((reg_check[rm.REG_CP_COL] > rm.SUMMER_CAP_CP).sum()):,}")
    r4.metric("Conditional Enrolment, over one subject", f"{int(is_ce.sum()):,}")
    st.caption(
        "From the extract's Summer registration columns: a Summer prep counts 15cp and a "
        f"block subject 10cp. Conditional Enrolment allows {rm.CE_SUMMER_SUBJECTS} Summer subject "
        "(the prep or one other). This shows who is over, not who was approved: compare with "
        "the PC decisions before contacting anyone."
    )
    st.dataframe(flagged, use_container_width=True, hide_index=True)
    st.download_button(
        "Download Summer registration check (.xlsx)", rm.summer_registration_workbook(reg_check),
        "summer_registration_check.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disabled=not len(reg_check),
    )

if not rm.has_progress_statuses(df):
    st.error(
        "This file's Prep and Subject Status columns don't say who has passed: there's no "
        "\"Completed\", \"Currently Registered\" or 1 in them. Advice from it would show every "
        "student owing everything, so it isn't produced. Ask for the extract to be re-run "
        "with the status values."
    )
    st.stop()

# --- Summer rounds -------------------------------------------------------------
# The targeting lists the coaching templates are sent from, not the full
# engine: who a confirmed Summer subject helps finish sooner (and, in Round 2,
# who failed and needs Autumn).
xlsx = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _coach_filter(frame: pd.DataFrame, key: str) -> pd.DataFrame:
    """The frame filtered by the coach multiselect (``key`` keeps widgets apart)."""
    coaches = sorted(frame[rm.COACH_COL].dropna().astype(str).unique())
    picked = st.multiselect("Filter by coach", coaches, default=[], key=key)
    return frame[frame[rm.COACH_COL].astype(str).isin(picked)] if picked else frame


def _flag_captions(frame: pd.DataFrame, as_at: str) -> None:
    n_paused = int(frame[rm.EARLY_PAUSED_COL].astype(bool).sum()) if rm.EARLY_PAUSED_COL in frame else 0
    if n_paused:
        st.caption(f"**{n_paused:,} paused** (Deferred / Leave of Absence) — kept in their group but "
                   f"flagged in *{rm.EARLY_PAUSED_COL}*: confirm they're returning before advising.")
    n_regcheck = int(frame[rm.EARLY_REG_CHECK_COL].astype(bool).sum()) if rm.EARLY_REG_CHECK_COL in frame else 0
    if n_regcheck:
        st.caption(f"**{n_regcheck:,}** have no SB3/SB4 registration in the file — flagged in "
                   f"*{rm.EARLY_REG_CHECK_COL}*. They keep the advice; check they're still "
                   f"studying first{f' (extract of {as_at})' if as_at else ''}.")


if is_summer and not engine_view:
    st.header(rm.ROUND_LABELS[summer_round])
    if not summer_offering:
        st.info("Upload the confirmed Summer offering list above to build this round's lists.")
        st.stop()
    as_at = st.text_input(
        "Extract date (registrations as at)", rm.extract_date(uploaded.name),
        help="Read from the file name when it has one (e.g. 20260923 …). Goes on every row, "
             "so coaches know how current the registrations and results are.",
    ).strip()
    shortlist = rm.summer_early_advice(df, summer_offering, session=session, as_at=as_at,
                                       summer_round=summer_round)
    stem = f"summer_round{summer_round[1]}_{session.replace(' ', '_')}"

if is_summer and not engine_view and summer_round == rm.ROUND_2:
    come_back = rm.summer_come_back(df, summer_offering, session=session, as_at=as_at)
    is_excl = shortlist[rm.EARLY_GROUP_COL] == rm.EARLY_GROUP_EXCLUDED
    n_finish = int((shortlist[rm.EARLY_GROUP_COL] == rm.EARLY_GROUP_FINISH).sum())
    n_earlier = int((shortlist[rm.EARLY_GROUP_COL] == rm.EARLY_GROUP_EARLIER).sum())
    m1, m2, m3, m4 = st.columns(4)
    m1.metric(f"2a — {rm.EARLY_GROUP_FINISH}", f"{n_finish:,}",
              help="Everything they still owe fits into this Summer (a Spring prep they're in "
                   "now counts as passed until Round 3)")
    m2.metric(f"2a — {rm.EARLY_GROUP_EARLIER}", f"{n_earlier:,}",
              help="Summer lightens next year enough that they finish a semester earlier")
    m3.metric(f"2b — {rm.EARLY_GROUP_COME_BACK}", f"{len(come_back):,}",
              help="Failed in SB3/SB4 and Summer doesn't bring their finish forward")
    m4.metric("Excluded", f"{int(is_excl.sum()):,}",
              help="Progression Outcome is Exclusion — not eligible to re-register; own tab, not advised")
    st.caption(
        "**Round 2 — after SB4 results.** Actual results as the file has them; only a status still "
        "*Currently Registered* (the Spring prep) counts as passed. **2a**: students Summer brings "
        f"forward — *{rm.EARLY_SU1_RECONFIRM_COL}* and *{rm.EARLY_CATCHUP_COLS['SU2']}* are the "
        "subjects to be in. Core subjects only: a prep still owed is in "
        f"*{rm.EARLY_PREP_LATER_COL}* and advised in Round 3. **2b**: students who failed a "
        "subject in SB3 or SB4 and whom Summer doesn't help — come back in Autumn; "
        f"*{rm.COME_BACK_WHY_COL}* says whether the Round 1 advice had them finishing with Summer. "
        "Advice is from results alone — what they've registered in Summer never changes it; it's "
        f"shown in the *Registered* columns with an *{rm.ACTION_COL}*."
    )
    if rm.ACTION_COL in shortlist:
        counts = shortlist.loc[~is_excl, rm.ACTION_COL].value_counts()
        n_drop = int((come_back[rm.ACTION_COL] == rm.ACTION_DROP).sum()) if rm.ACTION_COL in come_back else 0
        st.caption(
            "**Against Summer registrations** — 2a: "
            + ", ".join(f"**{int(counts.get(a, 0)):,}** {a}"
                        for a in (rm.ACTION_REGISTER, rm.ACTION_OK, rm.ACTION_CHANGE))
            + f". 2b: **{n_drop:,}** registered in Summer anyway ({rm.ACTION_DROP})."
        )
    else:
        st.caption("This extract has no Summer registration columns, so the *Registered* and "
                   f"*{rm.ACTION_COL}* columns aren't shown.")
    _flag_captions(shortlist, as_at)

    tab_a, tab_b = st.tabs([rm.ROUND_2_TABS["2a"], rm.ROUND_2_TABS["2b"]])
    with tab_a:
        view = shortlist
        groups = st.multiselect("Filter by group", list(rm.EARLY_GROUPS), default=[], key="r2_groups")
        if groups:
            view = view[view[rm.EARLY_GROUP_COL].isin(groups)]
        view = _coach_filter(view, "r2a_coach")
        st.dataframe(view, use_container_width=True, hide_index=True)
        st.caption(f"Showing {len(view):,} of {len(shortlist):,}.")
    with tab_b:
        view_b = _coach_filter(come_back, "r2b_coach") if len(come_back) else come_back
        st.dataframe(view_b, use_container_width=True, hide_index=True)
        st.caption(f"Showing {len(view_b):,} of {len(come_back):,}.")

    st.subheader("Demand by subject and campus")
    st.caption("How many 2a students (not Excluded) each Summer subject would take, by campus.")
    st.dataframe(rm.summer_demand(shortlist), use_container_width=True, hide_index=True)

    tabs = rm.round2_tabs(shortlist, come_back)
    readme = rm.readme_frame(rm.ROUND_2_README)
    d1, d2 = st.columns(2)
    d1.download_button(
        "Download Round 2 workbook (2a, 2b, Excluded, Demand, Read me) (.xlsx)",
        rm.summer_tabs_workbook(tabs, readme, demand_from=shortlist),
        f"{stem}.xlsx", xlsx, type="primary",
    )
    d2.download_button(
        "Download Round 2 split by coach (.zip)", rm.summer_tabs_by_coach_zip(tabs, readme),
        f"{stem}_by_coach.zip", "application/zip",
    )
    st.stop()

if is_summer and not engine_view and summer_round == rm.ROUND_3:
    is_excl = shortlist[rm.EARLY_GROUP_COL] == rm.EARLY_GROUP_EXCLUDED
    live = shortlist[~is_excl]
    needs_pc = rm.summer_pc_approval(shortlist)
    n_both = int(live[rm.EARLY_BOTH_PREPS_COL].astype(bool).sum()) if rm.EARLY_BOTH_PREPS_COL in live else 0
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("3a — prep advised", f"{len(live):,}",
              help="Owe GEDU0016 / GEDU0017 on results and Summer brings their finish forward")
    m2.metric("Need PC approval", f"{len(needs_pc):,}",
              help=f"Prep plus a block subject - over {rm.SUMMER_NO_APPROVAL_CP}cp")
    m3.metric("Both preps owed", f"{n_both:,}", help="One advised; the other is the academic call")
    m4.metric("Excluded", f"{int(is_excl.sum()):,}",
              help="Progression Outcome is Exclusion — not eligible to re-register; own tab, not advised")
    st.caption(
        "**Round 3 — after Spring results.** The prep round: students who still owe a prep on "
        f"results and whom Summer brings forward. *{rm.EARLY_PREP_COL}* is the subject for template "
        "3a; the SU1 / SU2 columns are what fits beside it. **The prep takes priority**: a prep plus "
        f"one block subject is the {rm.SUMMER_CAP_CP}cp limit, so a second registered block subject "
        f"is to drop (see *{rm.ACTION_DETAIL_COL}*); a prep plus a block subject needs the PC's "
        f"approval (*{rm.EARLY_APPROVAL_COL}*). A student owing both preps is advised one - "
        f"*{rm.EARLY_BOTH_PREPS_COL}* flags them for the rule-waiver call. Students Summer helps "
        "with no prep owed were advised in Round 2 and aren't listed again."
    )
    if rm.ACTION_COL in shortlist:
        counts = live[rm.ACTION_COL].value_counts()
        st.caption(
            "**Against Summer registrations** — "
            + ", ".join(f"**{int(counts.get(a, 0)):,}** {a}"
                        for a in (rm.ACTION_REGISTER, rm.ACTION_OK, rm.ACTION_CHANGE))
            + "."
        )
    else:
        st.caption("This extract has no Summer registration columns, so the *Registered* and "
                   f"*{rm.ACTION_COL}* columns aren't shown.")
    _flag_captions(shortlist, as_at)

    view = shortlist
    groups = st.multiselect("Filter by group", list(rm.EARLY_GROUPS), default=[], key="r3_groups")
    if groups:
        view = view[view[rm.EARLY_GROUP_COL].isin(groups)]
    view = _coach_filter(view, "r3_coach") if len(view) else view
    st.dataframe(view, use_container_width=True, hide_index=True)
    st.caption(f"Showing {len(view):,} of {len(shortlist):,}.")

    st.subheader("Demand by subject and campus")
    st.dataframe(rm.summer_demand(shortlist), use_container_width=True, hide_index=True)

    tabs = rm.round3_tabs(shortlist)
    readme = rm.readme_frame(rm.ROUND_3_README)
    d1, d2 = st.columns(2)
    d1.download_button(
        "Download Round 3 workbook (3a, Excluded, Demand, Read me) (.xlsx)",
        rm.summer_tabs_workbook(tabs, readme, demand_from=shortlist),
        f"{stem}.xlsx", xlsx, type="primary",
    )
    d2.download_button(
        "Download Round 3 split by coach (.zip)", rm.summer_tabs_by_coach_zip(tabs, readme),
        f"{stem}_by_coach.zip", "application/zip",
    )
    st.download_button(
        f"Download PC approval files by program ({len(needs_pc):,}) (.zip)",
        rm.summer_pc_by_program_zip(shortlist),
        f"{stem}_pc_approval.zip", "application/zip", disabled=not len(needs_pc),
    )
    st.stop()

if is_summer and not engine_view and summer_round == rm.ROUND_1:
    if not len(shortlist):
        st.warning("No students match — Summer doesn't bring anyone's transition to the "
                   "University forward with this offering.")
        st.stop()

    is_excl = shortlist[rm.EARLY_GROUP_COL] == rm.EARLY_GROUP_EXCLUDED
    n_finish = int((shortlist[rm.EARLY_GROUP_COL] == rm.EARLY_GROUP_FINISH).sum())
    n_earlier = int((shortlist[rm.EARLY_GROUP_COL] == rm.EARLY_GROUP_EARLIER).sum())
    m1, m2, m3 = st.columns(3)
    m1.metric(rm.EARLY_GROUP_FINISH, f"{n_finish:,}", help="Everything they still owe fits into this Summer — they finish their diploma / UPP")
    m2.metric(rm.EARLY_GROUP_EARLIER, f"{n_earlier:,}", help="Summer lightens next year enough that they finish (and move to the University) a semester earlier")
    m3.metric("Excluded", f"{int(is_excl.sum()):,}", help="Progression Outcome is Exclusion — not eligible to re-register; separate tab, not advised")
    st.caption(
        "**Round 1 — broad, early.** Students Summer **materially helps** (judged on SU1 + SU2 "
        "together): they finish in Summer, or a semester earlier. SU1 and prep are advised now; "
        f"SU2 picks are shown in *{rm.EARLY_SU2_LATER_COL}*: students wait for Round 2 to register. "
        f"One subject per block, an elective can take a free block, {rm.SUMMER_CAP_CP}cp limit. "
        f"Over {rm.SUMMER_NO_APPROVAL_CP}cp (a prep plus a subject or elective) needs the PC's "
        f"approval: advise the prep and hold the rest, see *{rm.EARLY_APPROVAL_COL}*. "
        f"Conditional Enrolment gets {rm.CE_SUMMER_SUBJECTS} Summer subject (the prep first), "
        f"with the other choice in *{rm.EARLY_ALT_COL}*. "
        "Subjects they're enrolled in now count as passed, and a free SB4 slot as used. "
        "Messaging stays general (the confirmed offerings); this is the *who to contact* list. "
        "Excluded students are on their own tab for the coach, not contacted about Summer."
    )
    _flag_captions(shortlist, as_at)

    view = shortlist
    groups = st.multiselect("Filter by group", list(rm.EARLY_GROUPS), default=[])
    if groups:
        view = view[view[rm.EARLY_GROUP_COL].isin(groups)]
    view = _coach_filter(view, "r1_coach")
    st.dataframe(view, use_container_width=True, hide_index=True)
    st.caption(f"Showing {len(view):,} of {len(shortlist):,} candidates.")

    demand = rm.summer_demand(shortlist)
    st.subheader("Demand by subject and campus")
    st.caption("How many listed students (not Excluded) each Summer subject would take, by "
               "campus — to check the offering against demand.")
    st.dataframe(demand, use_container_width=True, hide_index=True)

    d1, d2 = st.columns(2)
    d1.download_button(
        "Download shortlist (.xlsx)",
        rm.summer_early_workbook(shortlist, "Round 1 Summer advice"),
        f"{stem}.xlsx", xlsx, type="primary",
    )
    d2.download_button(
        "Download split by coach (.zip)", rm.summer_early_by_coach_zip(shortlist),
        f"{stem}_by_coach.zip", "application/zip",
    )

    # Round 1 extras: PC approval, the 1b list, the prep split
    needs_pc = rm.summer_pc_approval(shortlist)
    st.subheader("PC approval")
    st.caption(
        f"**{len(needs_pc):,}** need PC approval (over {rm.SUMMER_NO_APPROVAL_CP}cp), across "
        f"**{needs_pc['PROGRAM_CD'].nunique():,}** programs. One file per program, each with a "
        f"*{rm.PC_DECISION_COL}* dropdown ({' / '.join(rm.PC_DECISIONS)}). A subject that only "
        f"runs at another campus still needs a rule waiver"
        + (f", see *{rm.EARLY_OTHER_CAMPUS_COL}*." if rm.EARLY_OTHER_CAMPUS_COL in shortlist else ".")
    )
    st.download_button(
        f"Download PC approval files by program ({len(needs_pc):,}) (.zip)",
        rm.summer_pc_by_program_zip(shortlist),
        f"{stem}_pc_approval.zip", "application/zip", disabled=not len(needs_pc),
    )

    one_gedu = rm.summer_one_gedu(df, shortlist, session)
    n_95 = int((one_gedu[rm.EARLY_CP_DONE_COL] == rm.ROUND_1B_CP).sum())
    st.subheader("Round 1b — one GEDU subject left")
    st.caption(
        f"**{n_95:,}** have {rm.ROUND_1B_CP}CP completed and only a prep left; "
        f"**{len(one_gedu) - n_95:,}** have only GEDU1001 or GEDU1002 left. Based on results "
        "alone: anything they're enrolled in now doesn't count as passed, and a student "
        "enrolled in that subject now isn't listed. CP is worked out from a 110CP diploma."
    )
    g1, g2 = st.columns(2)
    g1.download_button(
        f"Download Round 1b list ({len(one_gedu):,}) (.xlsx)",
        rm.summer_early_workbook(one_gedu, "Round 1b"),
        f"{stem}_round_1b.xlsx", xlsx, disabled=not len(one_gedu),
    )
    g2.download_button(
        "Download Round 1b split by coach (.zip)", rm.summer_early_by_coach_zip(one_gedu),
        f"{stem}_round_1b_by_coach.zip", "application/zip", disabled=not len(one_gedu),
    )

    with_prep, no_prep = rm.summer_early_prep_split(shortlist)
    st.caption(
        f"**Split by prep** — **{len(with_prep):,}** need a Summer prep "
        f"(GEDU0016 / GEDU0017; any catch-up subjects they also need are on the same row), "
        f"**{len(no_prep):,}** need catch-up subjects only. Excluded students aren't in either."
    )
    p1, p2 = st.columns(2)
    p1.download_button(
        f"Download prep students only ({len(with_prep):,}) (.xlsx)",
        rm.summer_early_workbook(with_prep, "Round 1 prep"),
        f"{stem}_prep.xlsx", xlsx, disabled=not len(with_prep),
    )
    p2.download_button(
        f"Download no-prep students only ({len(no_prep):,}) (.xlsx)",
        rm.summer_early_workbook(no_prep, "Round 1 no prep"),
        f"{stem}_no_prep.xlsx", xlsx, disabled=not len(no_prep),
    )
    p1.download_button(
        "Download prep students split by coach (.zip)", rm.summer_early_by_coach_zip(with_prep),
        f"{stem}_prep_by_coach.zip", "application/zip", disabled=not len(with_prep),
    )
    p2.download_button(
        "Download no-prep students split by coach (.zip)", rm.summer_early_by_coach_zip(no_prep),
        f"{stem}_no_prep_by_coach.zip", "application/zip", disabled=not len(no_prep),
    )
    st.stop()

dropped = int(df.attrs.get("duplicates_dropped", 0))
result = rm.build_advice(df, session=session, summer_subjects=summer_subjects,
                         summer_round=summer_round)
coach_view = rm.build_coach_view(result)

# --- summary --------------------------------------------------------------
total = len(result)
src = result[rm.SOURCE_COL]
from_calc = int(src.str.startswith("Grant calculator").sum())
assumed = int(src.str.contains("pattern assumed", regex=False).sum())
ce = int(src.str.contains("CE 30cp cap", regex=False).sum())
excluded = int(src.str.contains("Exclusion", regex=False).sum())
fallback = int(src.str.startswith("v2 rule-tree").sum())
flagged = int(result[rm.REASON_COL].str.contains("NOTE:", regex=False).sum())

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Students", f"{total:,}")
c2.metric("From calculator", f"{from_calc:,}", help=f"of which {assumed:,} on an assumed (carried-forward) pattern")
c3.metric("Conditional Enrolment (30cp)", f"{ce:,}")
c4.metric("Rule-tree fallback", f"{fallback:,}")
c5.metric("Flagged for coach review", f"{flagged:,}")
no_benefit = int(src.str.contains("no benefit", regex=False).sum())
if no_benefit:
    st.caption(f"{no_benefit:,} student(s) not advised for Summer — it wouldn't bring their "
               "transition to the University forward (subjects greyed, reason says why).")
if excluded:
    st.caption(f"{excluded:,} student(s) excluded — no advice. They're on an "
               f"*{rm._EXCLUDED_TAB}* tab in the downloads, not the Coach View / template tabs.")
if rm.REG_CHECK_COL in result.columns:
    ok = int((result[rm.REG_CHECK_COL] == rm.REG_OK).sum())
    st.caption(
        f"Registration check (Blocks 3–4): **{ok:,}** already registered as advised; "
        f"**{total - ok:,}** need a look (register, drop, check — or paused / excluded)."
    )
flags = result.loc[result[rm.WITHDRAWAL_COL].ne(""), rm.WITHDRAWAL_COL].value_counts()
if not flags.empty:
    st.caption(
        "Mid-semester flags (commencing students with no pass yet): "
        + ", ".join(f"**{int(flags.get(f, 0)):,}** {f.lower()}"
                    for f in (rm._WITHDRAW_TEXT, rm._CHECK_ENROL_TEXT, rm._AWAITING_TEXT))
        + "."
    )

paused = (
    int(coach_view[rm.STUDY_PATH_COL].map(rm.is_paused).sum())
    if rm.STUDY_PATH_COL in coach_view.columns else 0
)
no_outcome = int((coach_view[rm.STUDY_STATUS_COL] == rm.NO_OUTCOME).sum())
if paused:
    st.caption(
        f"**{paused:,} paused** (Deferred / Leave of Absence) — still get subject "
        "advice (greyed), but confirm they're returning first, and are never "
        "flagged for withdrawal. In the per-coach "
        f"split they're all on the *{rm._PAUSED_TAB}* tab; commencing ones also "
        "stay on the Commencing tab, so they appear in both."
    )
if no_outcome:
    st.warning(
        f"{no_outcome:,} student(s) have **no Progression Outcome** in the file "
        "but commenced before this session — they should have one. Shown as "
        f"*{rm.NO_OUTCOME}* in Study Status; not assumed to be Good Standing."
    )

if dropped:
    st.caption(
        f"{dropped} exact duplicate row(s) dropped on load — same student, same "
        "program, every column identical. Genuine second-program enrolments are kept."
    )
dual = int((coach_view[rm.OTHER_ENROL_COL].astype(str) != "").sum()) if rm.OTHER_ENROL_COL in coach_view.columns else 0
if dual:
    st.caption(
        f"{dual} row(s) belong to students enrolled in a **second program** — see "
        f"the *{rm.OTHER_ENROL_COL}* column. Two coaches may be advising the same "
        "student, and one enrolment may never have been withdrawn."
    )

# --- preview (the Coach View sheet) --------------------------------------
st.caption(
    "Preview of the **Coach View** sheet. The full workbook (original columns "
    "+ advice appended) is the second sheet in the download."
)
view = coach_view

programs = sorted(coach_view["PROGRAM_CD"].dropna().astype(str).unique())
pick = st.multiselect("Filter by program", programs, default=[])
if pick:
    view = view[view["PROGRAM_CD"].astype(str).isin(pick)]

fcol1, fcol2 = st.columns(2)
with fcol1:
    only_flagged = st.checkbox("Only rows flagged for coach review", value=False)
with fcol2:
    only_fallback = st.checkbox("Only v2-fallback rows", value=False)
if only_flagged:
    view = view[view[rm.REASON_COL].str.contains("NOTE:", regex=False)]
if only_fallback:
    view = view[view[rm.SOURCE_COL].str.startswith("v2 rule-tree")]

# the grey marker (~) is only meaningful in the styled .xlsx - show it as (…) here
preview = view.replace(r"^~(.+)$", r"(\1)", regex=True)
st.dataframe(preview, use_container_width=True, hide_index=True)
st.caption(
    f"Showing {len(view):,} of {total:,} students.  "
    "Progress key: ✓ = passed, ◐ = registered now, ○ = still to pass.  "
    "A block subject in (brackets) here is greyed in the download — shown for "
    "reference, not being registered this session. On a mid-semester run the "
    "greyed Block 1–2 cells mark the result: ✓ passed, ✗ failed, 0 not enrolled "
    "(no mark = no grade recorded yet)."
)

st.download_button(
    "Download workbook (Coach View + full sheet)",
    rm.to_workbook_bytes(result, coach_view),
    f"reregistration_advice_{session.replace(' ', '_')}.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

# --- split for distribution ---------------------------------------------------
st.divider()
if rm.COACH_COL not in coach_view.columns:
    st.caption(
        f"No '{rm.COACH_COL}' column in this file — can't split the Coach View by "
        "success coach."
    )
else:
    coaches = coach_view[rm.COACH_COL].fillna("").astype(str).str.strip()
    n_files = coaches.replace("", "(no coach)").nunique()
    n_blank = int((coaches == "").sum())
    st.caption(
        f"**Split for distribution** — one Coach View workbook per success coach "
        f"({n_files} file(s)), with a tab per Messaging Template inside each "
        f"plus a *{rm._PAUSED_TAB}* tab for Deferred / Leave-of-Absence students "
        f"(commencing paused students appear there and on the Commencing tab), "
        f"and an *{rm._EXCLUDED_TAB}* tab for students with no advice. "
        f"Advice columns show the subject name beside each code."
        + (f"  {n_blank} student(s) with no coach go into a `no_coach.xlsx` file." if n_blank else "")
    )
    st.download_button(
        "Download Coach View split by coach (.zip)",
        rm.split_coach_view_zip_bytes(coach_view),
        f"reregistration_advice_{session.replace(' ', '_')}_by_coach.zip",
        "application/zip",
    )

    # Mid-semester only: the same split, but a tab per coach ACTION instead of
    # per template - read off the Registration Check.
    if rm.REG_CHECK_COL in coach_view.columns:
        actions = pd.Series([rm.coach_action(r, session)[0] for _, r in coach_view.iterrows()])
        counts = actions.value_counts()
        st.caption(
            "**Split by action** — the same one-workbook-per-coach split, but with a tab per "
            "thing the coach has to do (from the Registration Check). Each student sits on "
            "one tab — the most urgent action — and *Action detail* lists every to-do: "
            + ", ".join(f"**{int(counts.get(a, 0)):,}** {a}" for a in rm.ACTIONS) + "."
        )
        st.download_button(
            "Download split by coach — by action (.zip)",
            rm.split_coach_view_by_action_zip_bytes(coach_view, session),
            f"reregistration_advice_{session.replace(' ', '_')}_by_action.zip",
            "application/zip",
        )
