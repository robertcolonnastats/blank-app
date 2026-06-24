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

    with st.spinner("Working... this can take a few minutes (Statcast pulls are per-game)."):
        hitting, pitching, teams = build_full(int(season), progress_callback)
        excel_bytes = write_excel_bytes(hitting, pitching, teams)

    os.makedirs(DATA_DIR, exist_ok=True)
    with open(EXCEL_PATH, "wb") as f:
        f.write(excel_bytes.getvalue())
    with open(TIMESTAMP_PATH, "w") as f:
        f.write(datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC") + " (manual pull)")

    st.session_state["hitting"] = hitting
    st.session_state["pitching"] = pitching
    status.text("Done.")
    st.success(f"Pulled {len(hitting)} hitter-rows and {len(pitching)} pitcher-rows.")
    st.rerun()

# --- Display whatever the latest saved file is (auto or manual) ---
if os.path.exists(EXCEL_PATH):
    with open(EXCEL_PATH, "rb") as f:
        file_bytes = f.read()

    st.download_button(
        label="📥 Download Excel workbook",
        data=file_bytes,
        file_name=f"mets_milb_{season}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    try:
        xl = pd.ExcelFile(io.BytesIO(file_bytes))
        sheet = st.selectbox("Preview a sheet", xl.sheet_names)
        df = pd.read_excel(xl, sheet)
        st.dataframe(df, use_container_width=True)
    except Exception as e:
        st.error(f"Couldn't preview the file: {e}")
else:
    st.info("No data file yet -- click 'Pull now' above to generate one.")
