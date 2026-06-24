"""
streamlit_app.py

Displays the Mets minor-league stats workbook that GitHub Actions pulls
automatically every morning (see .github/workflows/daily_pull.yml +
pull_and_save.py). Also has a single "Pull now" button for an on-demand
refresh -- one button does the full job (box score, all levels, plus
Statcast for Triple-A/Single-A) with no second step required.

requirements.txt should contain:
    streamlit
    requests
    pandas
    openpyxl
"""

import datetime
import io
import os

import pandas as pd
import streamlit as st

from mets_milb_lib import build_full, write_excel_bytes


def fmt(value, spec, suffix=""):
    """MLB's API sometimes returns numeric stats (avg/obp/slg, etc.) as
    strings like '.267'. This safely coerces to float before formatting,
    instead of letting the f-string format spec choke on a str."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "-"
    try:
        return f"{float(value):{spec}}{suffix}"
    except (TypeError, ValueError):
        return "-"

DATA_DIR = "data"
EXCEL_PATH = os.path.join(DATA_DIR, "mets_milb_latest.xlsx")
TIMESTAMP_PATH = os.path.join(DATA_DIR, "last_updated.txt")

st.set_page_config(page_title="Mets MiLB Stat Pull", layout="wide")
st.title("⚾ Mets Minor League Stat Pull")
st.caption(
    "Every Mets affiliate (Syracuse, Binghamton, Brooklyn, St. Lucie, FCL, DSL) -- "
    "no major leaguers. Box score stats for all levels; Statcast (EV, launch angle, "
    "hard-hit%) for Syracuse and St. Lucie, the only levels with public tracking. "
    "Updates automatically every morning."
)

season = st.number_input("Season", min_value=2015, max_value=2030,
                          value=datetime.date.today().year, step=1)

# --- Show last auto-updated timestamp, if available ---
if os.path.exists(TIMESTAMP_PATH):
    with open(TIMESTAMP_PATH) as f:
        st.info(f"Last automatic update: {f.read().strip()}")
else:
    st.warning("No automated pull has run yet. Use 'Pull now' below, or wait for the next scheduled run.")

# --- Manual on-demand pull: ONE button does everything ---
if st.button("🔄 Pull now (box score + Statcast, all in one step)", type="primary"):
    progress_bar = st.progress(0.0)
    status = st.empty()

    def progress_callback(pct, msg):
        progress_bar.progress(min(max(pct, 0.0), 1.0))
        status.text(msg)

    with st.spinner("Working... this can take a while (pitch + batted-ball data is pulled game by game, for every level)."):
        hitting, pitching, arsenal, teams = build_full(int(season), progress_callback)
        excel_bytes = write_excel_bytes(hitting, pitching, teams, arsenal)

    os.makedirs(DATA_DIR, exist_ok=True)
    with open(EXCEL_PATH, "wb") as f:
        f.write(excel_bytes.getvalue())
    with open(TIMESTAMP_PATH, "w") as f:
        f.write(datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC") + " (manual pull)")

    st.session_state["hitting"] = hitting
    st.session_state["pitching"] = pitching
    st.session_state["arsenal"] = arsenal
    status.text("Done.")
    st.success(f"Pulled {len(hitting)} hitter-rows, {len(pitching)} pitcher-rows, {len(arsenal)} arsenal rows.")
    st.rerun()

# --- Load dataframes (from this session's pull, or from the saved file) ---
hitting = st.session_state.get("hitting")
pitching = st.session_state.get("pitching")
arsenal = st.session_state.get("arsenal")

if (hitting is None or pitching is None) and os.path.exists(EXCEL_PATH):
    try:
        xl = pd.ExcelFile(EXCEL_PATH)
        if "Hitting - All Levels" in xl.sheet_names:
            hitting = pd.read_excel(xl, "Hitting - All Levels")
        if "Pitching - All Levels" in xl.sheet_names:
            pitching = pd.read_excel(xl, "Pitching - All Levels")
        if "Pitch Arsenal" in xl.sheet_names:
            arsenal = pd.read_excel(xl, "Pitch Arsenal")
        st.session_state["hitting"] = hitting
        st.session_state["pitching"] = pitching
        st.session_state["arsenal"] = arsenal
    except Exception as e:
        st.error(f"Couldn't load saved data: {e}")

if os.path.exists(EXCEL_PATH):
    with open(EXCEL_PATH, "rb") as f:
        st.download_button(
            label="📥 Download Excel workbook",
            data=f.read(),
            file_name=f"mets_milb_{season}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

if hitting is None and pitching is None:
    st.info("No data yet -- click 'Pull now' above to generate it.")
    st.stop()

st.divider()
st.header("🔎 Player lookup")

hitter_names = set(hitting["playerName"].dropna().unique()) if hitting is not None and not hitting.empty else set()
pitcher_names = set(pitching["playerName"].dropna().unique()) if pitching is not None and not pitching.empty else set()
all_names = sorted(hitter_names | pitcher_names)

if not all_names:
    st.info("No players found in the data.")
    st.stop()

selected_name = st.selectbox("Search for a player", all_names)

is_hitter = selected_name in hitter_names
is_pitcher = selected_name in pitcher_names

if is_hitter:
    row = hitting[hitting["playerName"] == selected_name].iloc[0]
    st.subheader(f"🏏 {selected_name} — {row.get('team', '')} ({row.get('level', '')})")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Age", row.get("currentAge"))
    col2.metric("Position", row.get("position"))
    col3.metric("Bats", row.get("bats"))
    col4.metric("Height/Weight", f"{row.get('height','')} / {row.get('weight','')}")

    st.markdown("**Stat line**")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("AVG", fmt(row.get("avg"), ".3f", ""))
    c2.metric("OBP", fmt(row.get("obp"), ".3f", ""))
    c3.metric("SLG", fmt(row.get("slg"), ".3f", ""))
    c4.metric("BB%", fmt(row.get("BB_pct"), ".1%", ""))
    c5.metric("K%", fmt(row.get("K_pct"), ".1%", ""))

    if row.get("has_statcast"):
        st.markdown("**Statcast (this level is publicly tracked)**")
        s1, s2, s3, s4 = st.columns(4)
        s1.metric("Avg Exit Velo", fmt(row.get("avg_exit_velocity"), ".1f", " mph"))
        s2.metric("Max Exit Velo", fmt(row.get("max_exit_velocity"), ".1f", " mph"))
        s3.metric("Hard-Hit%", fmt(row.get("hard_hit_pct"), ".1%", ""))
        s4.metric("Barrel% (approx)", fmt(row.get("barrel_pct_approx"), ".1%", ""))
    else:
        st.caption("No Statcast available at this level (only Triple-A and Single-A are publicly tracked).")

    st.markdown("**Scouting grades (20-80 scale)**")
    st.caption(
        "Hit/Power are stat-derived with reasonable confidence. Field uses "
        "box-score fielding only (no OAA-equivalent exists for minors) -- "
        "low-medium confidence. Run uses stolen-base rate only, since no "
        "public sprint speed exists for minors -- low confidence. Arm has "
        "zero public data anywhere and is left blank rather than guessed."
    )
    g1, g2, g3, g4, g5, g6 = st.columns(6)
    g1.metric("Hit", row.get("hit_grade") if pd.notna(row.get("hit_grade")) else "-")
    g2.metric("Power", row.get("power_grade") if pd.notna(row.get("power_grade")) else "-")
    g3.metric("Field", row.get("field_grade") if pd.notna(row.get("field_grade")) else "-")
    g4.metric("Run", row.get("run_grade") if pd.notna(row.get("run_grade")) else "-")
    g5.metric("Arm", "N/A")
    g6.metric("Overall FV", row.get("overall_fv") if pd.notna(row.get("overall_fv")) else "-")

    st.markdown("**Skill vs. Results**")
    skill_z = row.get("skill_z")
    results_z = row.get("results_z")
    underv = row.get("undervalued_score")
    if pd.notna(skill_z) and pd.notna(results_z):
        chart_df = pd.DataFrame({
            "signal": ["Skill (discipline/age/contact quality)", "Results (actual stat line)"],
            "z-score vs. level peers": [skill_z, results_z],
        }).set_index("signal")
        st.bar_chart(chart_df)
        if underv is not None and underv > 0.3:
            st.success(f"Undervalued score: {underv:+.2f} — underlying skill looks better than the stat line shows.")
        elif underv is not None and underv < -0.3:
            st.warning(f"Undervalued score: {underv:+.2f} — stat line currently looks better than the underlying skill.")
        else:
            st.info(f"Undervalued score: {underv:+.2f} — results roughly match the underlying skill.")
    else:
        st.caption("Not enough data at this level to compute a skill-vs-results comparison.")

elif is_pitcher:
    row = pitching[pitching["playerName"] == selected_name].iloc[0]
    st.subheader(f"⚾ {selected_name} — {row.get('team', '')} ({row.get('level', '')})")

    col1, col2, col3 = st.columns(3)
    col1.metric("Age", row.get("currentAge"))
    col2.metric("Throws", row.get("throws"))
    col3.metric("Height/Weight", f"{row.get('height','')} / {row.get('weight','')}")

    st.markdown("**Stat line**")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("ERA", f"{row.get('era', '-')}")
    c2.metric("WHIP", f"{row.get('whip', '-')}")
    c3.metric("K%", fmt(row.get("K_pct"), ".1%", ""))
    c4.metric("K-BB%", fmt(row.get("K_minus_BB_pct"), ".1%", ""))

    if row.get("has_statcast"):
        st.markdown("**Contact allowed (Statcast, this level is publicly tracked)**")
        s1, s2, s3 = st.columns(3)
        s1.metric("Avg Exit Velo Allowed", fmt(row.get("avg_exit_velocity_allowed"), ".1f", " mph"))
        s2.metric("Hard-Hit% Allowed", fmt(row.get("hard_hit_pct_allowed"), ".1%", ""))
        s3.metric("Barrel% Allowed (approx)", fmt(row.get("barrel_pct_approx_allowed"), ".1%", ""))
    else:
        st.caption("No batted-ball Statcast available at this level.")

    st.markdown("**Pitch arsenal**")
    if arsenal is not None and not arsenal.empty and "playerId" in row.index:
        pid = row.get("playerId")
        pitcher_arsenal = arsenal[arsenal["pitcherId"] == pid] if "pitcherId" in arsenal.columns else pd.DataFrame()
        if not pitcher_arsenal.empty:
            display_cols = ["pitch_type", "usage_pct", "avg_velo", "max_velo",
                             "avg_spin_rate", "avg_horiz_break_in", "avg_vert_break_in", "whiff_pct"]
            display_cols = [c for c in display_cols if c in pitcher_arsenal.columns]
            st.dataframe(
                pitcher_arsenal[display_cols].sort_values("usage_pct", ascending=False),
                use_container_width=True, hide_index=True,
            )
            if pitcher_arsenal["avg_spin_rate"].isna().all():
                st.caption("No spin-rate tracking at this level (velocity/movement may still be tracked).")
        else:
            st.caption("No pitch-tracking data available for this pitcher.")
    else:
        st.caption("No pitch-tracking data available.")

    st.markdown("**Scouting grades (20-80 scale)**")
    st.caption(
        f"Control: {row.get('control_grade_confidence', 'n/a')}. "
        f"Stuff: {row.get('stuff_grade_confidence', 'n/a')}."
    )
    g1, g2, g3 = st.columns(3)
    g1.metric("Control", row.get("control_grade") if pd.notna(row.get("control_grade")) else "-")
    g2.metric("Stuff", row.get("stuff_grade") if pd.notna(row.get("stuff_grade")) else "-")
    g3.metric("Overall FV", row.get("pitching_overall_fv") if pd.notna(row.get("pitching_overall_fv")) else "-")

    st.markdown("**Skill vs. Results**")
    skill_z = row.get("skill_z")
    results_z = row.get("results_z")
    underv = row.get("undervalued_score")
    if pd.notna(skill_z) and pd.notna(results_z):
        chart_df = pd.DataFrame({
            "signal": ["Skill (K-BB%/age/contact allowed)", "Results (ERA)"],
            "z-score vs. level peers": [skill_z, results_z],
        }).set_index("signal")
        st.bar_chart(chart_df)
        if underv is not None and underv > 0.3:
            st.success(f"Undervalued score: {underv:+.2f} — underlying skill looks better than the ERA shows.")
        elif underv is not None and underv < -0.3:
            st.warning(f"Undervalued score: {underv:+.2f} — ERA currently looks better than the underlying skill.")
        else:
            st.info(f"Undervalued score: {underv:+.2f} — ERA roughly matches the underlying skill.")
    else:
        st.caption("Not enough data at this level to compute a skill-vs-results comparison.")
