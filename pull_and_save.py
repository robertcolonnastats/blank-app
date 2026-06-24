"""
pull_and_save.py

Headless (no UI) script that does the full pull -- box score stats for
every Mets affiliate, plus Statcast for Triple-A/Single-A -- and saves the
resulting Excel workbook to data/mets_milb_latest.xlsx, along with a
data/last_updated.txt timestamp.

This is meant to be run on a schedule by GitHub Actions (see
.github/workflows/daily_pull.yml), not by a person directly. The Streamlit
app just reads whatever this script last saved.
"""

import argparse
import datetime
import os

from mets_milb_lib import build_full, write_excel_bytes


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--season", type=int, default=datetime.date.today().year)
    parser.add_argument("--out-dir", type=str, default="data")
    args = parser.parse_args()

    os.makedirs(args.out_dir, exist_ok=True)

    hitting, pitching, teams = build_full(args.season)
    excel_bytes = write_excel_bytes(hitting, pitching, teams)

    out_path = os.path.join(args.out_dir, "mets_milb_latest.xlsx")
    with open(out_path, "wb") as f:
        f.write(excel_bytes.read())

    timestamp_path = os.path.join(args.out_dir, "last_updated.txt")
    with open(timestamp_path, "w") as f:
        f.write(datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC"))

    print(f"Saved {out_path}")
    print(f"  Hitting rows: {len(hitting)}")
    print(f"  Pitching rows: {len(pitching)}")


if __name__ == "__main__":
    main()
