"""
mets_milb_pull.py

Pulls EVERY Mets minor-league affiliate's hitting and pitching season stats
(no major leaguers) from the public MLB Stats API, computes a set of
"undervalued vs. results" metrics, and writes a sortable Excel workbook
with one tab per level plus a combined tab.

Usage:
    pip install requests pandas openpyxl
    python mets_milb_pull.py --season 2026 --out mets_milb_2026.xlsx

Notes:
    - MLB_TEAM_ORG_ID = 121 is the Mets parent org ID. This pulls ONLY
      affiliates (sportId != 1), so major leaguers never enter the dataset.
    - sportId mapping used by MLB Stats API:
        11 = Triple-A
        12 = Double-A
        13 = High-A
        14 = Single-A
        16 = Rookie / Complex (FCL)
        5442 = Dominican Summer League (DSL) -- not always populated/
               not always exposed with full stat depth; script will
               just return an empty/partial frame for it if unavailable.
    - The API does not expose Statcast-level data (exit velo, spin rate,
      etc.) for minors publicly. This pulls the deepest box-score-derived
      stats it does expose, plus derived sabermetric stats computed locally.
"""

import argparse
import time
import requests
import pandas as pd

BASE = "https://statsapi.mlb.com/api/v1"
METS_ORG_ID = 121

SPORT_IDS = {
    11: "Triple-A",
    12: "Double-A",
    13: "High-A",
    14: "Single-A",
    16: "Rookie/Complex (FCL)",
    5442: "Dominican Summer League",
}

SESSION = requests.Session()
SESSION.headers.update({"User-Agent": "personal-research-script/1.0"})


def get_json(url, params=None, retries=3, sleep=0.5):
    for attempt in range(retries):
        try:
            r = SESSION.get(url, params=params, timeout=20)
            r.raise_for_status()
            return r.json()
        except Exception as e:
            if attempt == retries - 1:
                print(f"  [warn] failed {url} params={params}: {e}")
                return {}
            time.sleep(sleep)


def get_affiliate_teams(season):
    """Return list of dicts: {teamId, teamName, sportId, levelName} for all
    Mets affiliates (excluding MLB itself) for the given season."""
    teams = []
    for sport_id, level_name in SPORT_IDS.items():
        data = get_json(f"{BASE}/teams", params={
            "sportId": sport_id,
            "season": season,
        })
        for t in data.get("teams", []):
            # parentOrgId identifies which MLB org a minor league team
            # belongs to.
            if t.get("parentOrgId") == METS_ORG_ID:
                teams.append({
                    "teamId": t["id"],
                    "teamName": t.get("name"),
                    "sportId": sport_id,
                    "levelName": level_name,
                })
    return teams


def get_team_group_stats(team_id, sport_id, season, group):
    """group = 'hitting' or 'pitching'. Returns list of per-player stat
    dicts for that team/season/group."""
    data = get_json(f"{BASE}/stats", params={
        "stats": "season",
        "group": group,
        "season": season,
        "sportId": sport_id,
        "teamId": team_id,
        "limit": 300,
    })
    rows = []
    for block in data.get("stats", []):
        for split in block.get("splits", []):
            player = split.get("player", {})
            stat = split.get("stat", {})
            row = {"playerId": player.get("id"), "playerName": player.get("fullName")}
            row.update(stat)
            rows.append(row)
    return rows


def get_player_bio(player_id):
    data = get_json(f"{BASE}/people/{player_id}")
    people = data.get("people", [])
    if not people:
        return {}
    p = people[0]
    return {
        "birthDate": p.get("birthDate"),
        "currentAge": p.get("currentAge"),
        "position": (p.get("primaryPosition") or {}).get("abbreviation"),
        "bats": (p.get("batSide") or {}).get("code"),
        "throws": (p.get("pitchHand") or {}).get("code"),
        "height": p.get("height"),
        "weight": p.get("weight"),
    }


def safe_div(n, d):
    try:
        n = float(n)
        d = float(d)
        return n / d if d else None
    except (TypeError, ValueError):
        return None


def add_hitting_derived(df):
    if df.empty:
        return df
    for col in ["hits", "homeRuns", "atBats", "strikeOuts", "baseOnBalls",
                "sacFlies", "plateAppearances", "hitByPitch"]:
        if col not in df.columns:
            df[col] = 0
    df["BABIP"] = df.apply(
        lambda r: safe_div(
            (float(r["hits"] or 0) - float(r["homeRuns"] or 0)),
            (float(r["atBats"] or 0) - float(r["strikeOuts"] or 0)
             - float(r["homeRuns"] or 0) + float(r["sacFlies"] or 0)),
        ), axis=1)
    df["K_pct"] = df.apply(
        lambda r: safe_div(r["strikeOuts"], r["plateAppearances"]), axis=1)
    df["BB_pct"] = df.apply(
        lambda r: safe_div(r["baseOnBalls"], r["plateAppearances"]), axis=1)
    if "slg" in df.columns and "avg" in df.columns:
        df["ISO"] = df.apply(
            lambda r: (float(r["slg"]) - float(r["avg"]))
            if r.get("slg") not in (None, "", "-") and r.get("avg") not in (None, "", "-")
            else None, axis=1)
    return df


def add_pitching_derived(df):
    if df.empty:
        return df
    for col in ["strikeOuts", "baseOnBalls", "battersFaced", "hits",
                "homeRuns", "inningsPitched"]:
        if col not in df.columns:
            df[col] = 0
    df["K_pct"] = df.apply(
        lambda r: safe_div(r["strikeOuts"], r["battersFaced"]), axis=1)
    df["BB_pct"] = df.apply(
        lambda r: safe_div(r["baseOnBalls"], r["battersFaced"]), axis=1)
    df["K_minus_BB_pct"] = df.apply(
        lambda r: (r["K_pct"] - r["BB_pct"])
        if r["K_pct"] is not None and r["BB_pct"] is not None else None, axis=1)
    return df


def zscore(series):
    s = pd.to_numeric(series, errors="coerce")
    if s.std(skipna=True) in (0, None) or s.dropna().empty:
        return pd.Series([None] * len(series), index=series.index)
    return (s - s.mean(skipna=True)) / s.std(skipna=True)


def build(season):
    print(f"Fetching Mets affiliate teams for {season}...")
    teams = get_affiliate_teams(season)
    if not teams:
        print("No affiliate teams found -- check season/org id.")
    for t in teams:
        print(f"  {t['levelName']:<28} {t['teamName']} (teamId={t['teamId']}, sportId={t['sportId']})")

    hitting_frames = []
    pitching_frames = []

    for t in teams:
        print(f"Pulling hitting stats: {t['teamName']}...")
        hrows = get_team_group_stats(t["teamId"], t["sportId"], season, "hitting")
        for r in hrows:
            r["level"] = t["levelName"]
            r["team"] = t["teamName"]
        hitting_frames.append(pd.DataFrame(hrows))

        print(f"Pulling pitching stats: {t['teamName']}...")
        prows = get_team_group_stats(t["teamId"], t["sportId"], season, "pitching")
        for r in prows:
            r["level"] = t["levelName"]
            r["team"] = t["teamName"]
        pitching_frames.append(pd.DataFrame(prows))

        time.sleep(0.2)  # be polite to the API

    hitting = pd.concat(hitting_frames, ignore_index=True) if hitting_frames else pd.DataFrame()
    pitching = pd.concat(pitching_frames, ignore_index=True) if pitching_frames else pd.DataFrame()

    # Bio / age enrichment (age-vs-level is one of the strongest public
    # undervaluation proxies available without Statcast).
    print("Enriching with player bio/age data (this takes a bit)...")
    all_ids = set()
    if not hitting.empty:
        all_ids |= set(hitting["playerId"].dropna().unique())
    if not pitching.empty:
        all_ids |= set(pitching["playerId"].dropna().unique())

    bios = {}
    for pid in all_ids:
        bios[pid] = get_player_bio(pid)
        time.sleep(0.05)

    bio_df = pd.DataFrame.from_dict(bios, orient="index")
    bio_df.index.name = "playerId"
    bio_df.reset_index(inplace=True)

    if not hitting.empty:
        hitting = hitting.merge(bio_df, on="playerId", how="left")
        hitting = add_hitting_derived(hitting)
        # age-vs-level z-score within each level
        hitting["age_vs_level_z"] = hitting.groupby("level")["currentAge"].transform(zscore)
        hitting["BABIP_z_in_level"] = hitting.groupby("level")["BABIP"].transform(zscore)
        hitting["K_pct_z_in_level"] = hitting.groupby("level")["K_pct"].transform(zscore)
        hitting["BB_pct_z_in_level"] = hitting.groupby("level")["BB_pct"].transform(zscore)
        # Simple undervaluation score: young for level + good discipline
        # + below-average BABIP (suggests bad luck depressing results)
        hitting["undervalued_score"] = (
            -hitting["age_vs_level_z"].fillna(0)
            + hitting["BB_pct_z_in_level"].fillna(0)
            - hitting["K_pct_z_in_level"].fillna(0)
            - hitting["BABIP_z_in_level"].fillna(0)
        )

    if not pitching.empty:
        pitching = pitching.merge(bio_df, on="playerId", how="left")
        pitching = add_pitching_derived(pitching)
        pitching["age_vs_level_z"] = pitching.groupby("level")["currentAge"].transform(zscore)
        pitching["K_minus_BB_z_in_level"] = pitching.groupby("level")["K_minus_BB_pct"].transform(zscore)
        era_z = pitching.groupby("level")["era"].transform(
            lambda s: zscore(pd.to_numeric(s, errors="coerce")))
        pitching["ERA_z_in_level"] = era_z
        # Undervaluation: young for level + strong K-BB% but ERA not yet
        # reflecting it (positive ERA z = worse than league = "unlucky"
        # if K-BB% is strong)
        pitching["undervalued_score"] = (
            -pitching["age_vs_level_z"].fillna(0)
            + pitching["K_minus_BB_z_in_level"].fillna(0)
            + pitching["ERA_z_in_level"].fillna(0)
        )

    return hitting, pitching, teams


def write_excel(hitting, pitching, teams, out_path):
    level_order = list(dict.fromkeys(t["levelName"] for t in teams))

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        # Combined tabs first
        if not hitting.empty:
            hitting.sort_values("undervalued_score", ascending=False).to_excel(
                writer, sheet_name="Hitting - All Levels", index=False)
        if not pitching.empty:
            pitching.sort_values("undervalued_score", ascending=False).to_excel(
                writer, sheet_name="Pitching - All Levels", index=False)

        # Per-level tabs
        for level in level_order:
            if not hitting.empty:
                sub = hitting[hitting["level"] == level].sort_values(
                    "undervalued_score", ascending=False)
                if not sub.empty:
                    sheet = f"Hit-{level}"[:31]
                    sub.to_excel(writer, sheet_name=sheet, index=False)
            if not pitching.empty:
                sub = pitching[pitching["level"] == level].sort_values(
                    "undervalued_score", ascending=False)
                if not sub.empty:
                    sheet = f"Pitch-{level}"[:31]
                    sub.to_excel(writer, sheet_name=sheet, index=False)

    # Add native Excel tables (gives sort/filter dropdowns on every tab)
    from openpyxl import load_workbook
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.utils import get_column_letter

    wb = load_workbook(out_path)
    for ws in wb.worksheets:
        if ws.max_row < 2 or ws.max_column < 1:
            continue
        last_col = get_column_letter(ws.max_column)
        ref = f"A1:{last_col}{ws.max_row}"
        table_name = "T_" + "".join(ch for ch in ws.title if ch.isalnum())
        tbl = Table(displayName=table_name, ref=ref)
        tbl.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(tbl)
        # Autofit-ish column widths
        for col_cells in ws.columns:
            length = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max(length + 2, 10), 40)
    wb.save(out_path)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--season", type=int, default=2026)
    parser.add_argument("--out", type=str, default="mets_milb_stats.xlsx")
    args = parser.parse_args()

    hitting, pitching, teams = build(args.season)
    write_excel(hitting, pitching, teams, args.out)
    print(f"\nDone. Wrote {args.out}")
    print(f"  Hitting rows: {len(hitting)}")
    print(f"  Pitching rows: {len(pitching)}")


if __name__ == "__main__":
    main()
