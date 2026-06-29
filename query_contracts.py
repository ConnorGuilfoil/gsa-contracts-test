#!/usr/bin/env python3
"""
NFL Contracts Query Tool
Reads NFL_Contracts.xlsx — no API key required.
Usage: python3 query_contracts.py
"""

import sys
import math
import pandas as pd
import numpy as np

EXCEL_PATH = "/Users/connorguilfoil/Desktop/Grady Sports/NFL_Contracts.xlsx"

# ── Stat column definitions ───────────────────────────────────────────────────

PS_DEF = [
    "DEF: Def Interceptions Int", "DEF: Def Interceptions Yds",
    "DEF: Def Interceptions IntTD", "DEF: Def Interceptions Lng",
    "DEF: Def Interceptions PD", "DEF: Fumbles FF", "DEF: Fumbles Fmb",
    "DEF: Fumbles FR", "DEF: Fumbles Yds", "DEF: Fumbles FRTD",
    "DEF: Sk", "DEF: Tackles Comb", "DEF: Tackles Solo", "DEF: Tackles Ast",
    "DEF: Tackles TFL", "DEF: Tackles QBHits", "DEF: Sfty",
]
PS_DEFA = [
    "DEFA: Pass Coverage Int", "DEFA: Pass Coverage Tgt", "DEFA: Pass Coverage Cmp",
    "DEFA: Pass Coverage Cmp%", "DEFA: Pass Coverage Yds", "DEFA: Pass Coverage Yds/Cmp",
    "DEFA: Pass Coverage Yds/Tgt", "DEFA: Pass Coverage TD", "DEFA: Pass Coverage Rat",
    "DEFA: Pass Coverage DADOT", "DEFA: Pass Coverage Air", "DEFA: Pass Coverage YAC",
    "DEFA: Pass Rush Bltz", "DEFA: Pass Rush Hrry", "DEFA: Pass Rush QBKD",
    "DEFA: Pass Rush Bats", "DEFA: Pass Rush Sk", "DEFA: Pass Rush Prss",
    "DEFA: Tackles Comb", "DEFA: Tackles MTkl", "DEFA: Tackles MTkl%",
]
PS_PASS = [
    "PASS: QBrec", "PASS: Cmp", "PASS: Att", "PASS: Cmp%", "PASS: Yds",
    "PASS: TD", "PASS: TD%", "PASS: Int", "PASS: Int%", "PASS: 1D",
    "PASS: Succ%", "PASS: Lng", "PASS: Y/A", "PASS: AY/A", "PASS: Y/C",
    "PASS: Y/G", "PASS: Rate", "PASS: QBR", "PASS: Sk", "PASS: Yds.1",
    "PASS: Sk%", "PASS: NY/A", "PASS: ANY/A", "PASS: 4QC", "PASS: GWD",
]
PS_RUSH = [
    "RUSH: Rushing Att", "RUSH: Rushing Yds", "RUSH: Rushing TD",
    "RUSH: Rushing 1D", "RUSH: Rushing Succ%", "RUSH: Rushing Lng",
    "RUSH: Rushing Y/A", "RUSH: Rushing Y/G", "RUSH: Rushing A/G", "RUSH: Fmb",
]
PS_REC = [
    "REC: Receiving Tgt", "REC: Receiving Rec", "REC: Receiving Yds",
    "REC: Receiving Y/R", "REC: Receiving TD", "REC: Receiving 1D",
    "REC: Receiving Succ%", "REC: Receiving Lng", "REC: Receiving R/G",
    "REC: Receiving Y/G", "REC: Receiving Ctch%", "REC: Receiving Y/Tgt", "REC: Fmb",
]
PS_SCR = [
    "SCR: Touchdowns RshTD", "SCR: Touchdowns RecTD", "SCR: Touchdowns PRTD",
    "SCR: Touchdowns KRTD", "SCR: Touchdowns FRTD", "SCR: Touchdowns IntTD",
    "SCR: Touchdowns OthTD", "SCR: Touchdowns AllTD", "SCR: Touchdowns 2PM",
    "SCR: D2P", "SCR: PAT XPM", "SCR: PAT XPA", "SCR: FG FGM", "SCR: FG FGA",
    "SCR: Sfty", "SCR: Pts", "SCR: Pts/G",
]
PS_ALL_STATS = PS_DEF + PS_DEFA + PS_PASS + PS_RUSH + PS_REC + PS_SCR

# Stats shown on comp cards per position
POSITION_STATS = {
    "QB":  PS_PASS + ["RUSH: Rushing Att", "RUSH: Rushing Yds", "RUSH: Rushing TD", "RUSH: Rushing Y/A"],
    "RB":  PS_RUSH + ["REC: Receiving Tgt", "REC: Receiving Rec", "REC: Receiving Yds", "REC: Receiving TD", "REC: Receiving Y/R", "REC: Receiving Ctch%"],
    "WR":  PS_REC,
    "TE":  PS_REC,
    "FB":  ["RUSH: Rushing Att", "RUSH: Rushing Yds", "RUSH: Rushing TD", "REC: Receiving Rec", "REC: Receiving Yds", "REC: Receiving TD"],
    "ED":  ["DEF: Sk", "DEF: Tackles Comb", "DEF: Tackles TFL", "DEF: Tackles QBHits", "DEF: Def Interceptions PD", "DEFA: Pass Rush Prss", "DEFA: Pass Rush Hrry", "DEFA: Tackles MTkl"],
    "IDL": ["DEF: Sk", "DEF: Tackles Comb", "DEF: Tackles TFL", "DEF: Tackles QBHits", "DEFA: Pass Rush Prss", "DEFA: Pass Rush Hrry", "DEFA: Tackles MTkl"],
    "LB":  ["DEF: Tackles Comb", "DEF: Tackles Solo", "DEF: Sk", "DEF: Tackles TFL", "DEF: Def Interceptions Int", "DEF: Def Interceptions PD", "DEF: Fumbles FF", "DEFA: Tackles MTkl"],
    "CB":  ["DEF: Tackles Comb", "DEF: Def Interceptions Int", "DEF: Def Interceptions PD", "DEF: Tackles TFL", "DEFA: Pass Coverage Tgt", "DEFA: Pass Coverage Cmp%", "DEFA: Pass Coverage Yds", "DEFA: Pass Coverage TD", "DEFA: Pass Coverage Rat"],
    "S":   ["DEF: Tackles Comb", "DEF: Def Interceptions Int", "DEF: Def Interceptions PD", "DEF: Tackles TFL", "DEFA: Pass Coverage Tgt", "DEFA: Pass Coverage Cmp%", "DEFA: Pass Coverage Yds", "DEFA: Pass Coverage TD"],
    "K":   ["SCR: FG FGM", "SCR: FG FGA", "SCR: PAT XPM", "SCR: PAT XPA", "SCR: Pts"],
    "P":   ["SCR: Pts", "SCR: Pts/G"],
    "LT": [], "RT": [], "LG": [], "RG": [], "C": [], "LS": [],
}

# Stats used for similarity scoring per position pair
PAIR_SIM_STATS = {
    ("ED",  "ED"):  ["DEF: Sk", "DEFA: Pass Rush Prss", "DEFA: Pass Rush Hrry", "DEF: Tackles QBHits", "DEF: Tackles TFL", "DEF: Tackles Comb"],
    ("IDL", "IDL"): ["DEF: Sk", "DEFA: Pass Rush Prss", "DEFA: Pass Rush Hrry", "DEF: Tackles QBHits", "DEF: Tackles TFL", "DEF: Tackles Comb"],
    ("LB",  "LB"):  ["DEF: Tackles Comb", "DEF: Sk", "DEF: Tackles TFL", "DEF: Tackles QBHits", "DEF: Def Interceptions Int", "DEF: Def Interceptions PD"],
    ("ED",  "IDL"): ["DEF: Sk", "DEFA: Pass Rush Prss", "DEFA: Pass Rush Hrry", "DEF: Tackles QBHits", "DEF: Tackles TFL", "DEF: Tackles Comb"],
    ("IDL", "ED"):  ["DEF: Sk", "DEFA: Pass Rush Prss", "DEFA: Pass Rush Hrry", "DEF: Tackles QBHits", "DEF: Tackles TFL", "DEF: Tackles Comb"],
    ("ED",  "LB"):  ["DEF: Sk", "DEF: Tackles TFL", "DEF: Tackles QBHits", "DEF: Tackles Comb"],
    ("LB",  "ED"):  ["DEF: Sk", "DEF: Tackles TFL", "DEF: Tackles QBHits", "DEF: Tackles Comb"],
    ("CB",  "CB"):  ["DEF: Def Interceptions Int", "DEF: Def Interceptions PD", "DEF: Tackles Comb", "DEFA: Pass Coverage Tgt"],
    ("S",   "S"):   ["DEF: Def Interceptions Int", "DEF: Def Interceptions PD", "DEF: Tackles Comb", "DEFA: Pass Coverage Tgt"],
    ("CB",  "S"):   ["DEF: Def Interceptions Int", "DEF: Def Interceptions PD", "DEF: Tackles Comb", "DEFA: Pass Coverage Tgt"],
    ("S",   "CB"):  ["DEF: Def Interceptions Int", "DEF: Def Interceptions PD", "DEF: Tackles Comb", "DEFA: Pass Coverage Tgt"],
    ("RB",  "RB"):  ["RUSH: Rushing Yds", "RUSH: Rushing Att", "RUSH: Rushing TD", "REC: Receiving Yds", "REC: Receiving Rec"],
    ("WR",  "WR"):  ["REC: Receiving Yds", "REC: Receiving Rec", "REC: Receiving Tgt", "REC: Receiving TD", "REC: Receiving Ctch%"],
    ("TE",  "TE"):  ["REC: Receiving Yds", "REC: Receiving Rec", "REC: Receiving Tgt", "REC: Receiving TD", "REC: Receiving Ctch%"],
    ("WR",  "TE"):  ["REC: Receiving Yds", "REC: Receiving Rec", "REC: Receiving Tgt", "REC: Receiving TD", "REC: Receiving Ctch%"],
    ("TE",  "WR"):  ["REC: Receiving Yds", "REC: Receiving Rec", "REC: Receiving Tgt", "REC: Receiving TD", "REC: Receiving Ctch%"],
    ("QB",  "QB"):  ["PASS: Yds", "PASS: TD", "PASS: Cmp", "PASS: Cmp%", "PASS: Int"],
}

# Stats used for negotiation adjustment per position
ADJ_STATS = {
    "ED":  ["DEF: Sk", "DEFA: Pass Rush Prss", "DEF: Tackles QBHits", "DEF: Tackles TFL"],
    "IDL": ["DEF: Sk", "DEFA: Pass Rush Prss", "DEF: Tackles QBHits", "DEF: Tackles TFL"],
    "LB":  ["DEF: Tackles Comb", "DEF: Sk", "DEF: Tackles TFL", "DEF: Tackles QBHits"],
    "CB":  ["DEF: Def Interceptions Int", "DEF: Def Interceptions PD", "DEF: Tackles Comb", "DEFA: Pass Coverage Tgt"],
    "S":   ["DEF: Def Interceptions Int", "DEF: Def Interceptions PD", "DEF: Tackles Comb", "DEFA: Pass Coverage Tgt"],
    "RB":  ["RUSH: Rushing Yds", "RUSH: Rushing Att", "RUSH: Rushing TD", "REC: Receiving Yds", "REC: Receiving Rec"],
    "WR":  ["REC: Receiving Yds", "REC: Receiving Rec", "REC: Receiving TD", "REC: Receiving Ctch%"],
    "TE":  ["REC: Receiving Yds", "REC: Receiving Rec", "REC: Receiving TD", "REC: Receiving Ctch%"],
    "QB":  ["PASS: Yds", "PASS: TD", "PASS: Cmp%", "PASS: Int"],
}

# Position proximity multipliers
POS_PROX = {
    ("ED",  "ED"):  1.00, ("IDL", "IDL"): 1.00, ("LB",  "LB"):  1.00,
    ("ED",  "IDL"): 0.85, ("IDL", "ED"):  0.85,
    ("ED",  "LB"):  0.70, ("LB",  "ED"):  0.70,
    ("CB",  "CB"):  1.00, ("S",   "S"):   1.00,
    ("CB",  "S"):   0.85, ("S",   "CB"):  0.85,
    ("RB",  "RB"):  1.00, ("WR",  "WR"):  1.00,
    ("TE",  "TE"):  1.00, ("WR",  "TE"):  0.90, ("TE",  "WR"):  0.90,
    ("QB",  "QB"):  1.00,
}

# Age proximity multipliers (years apart → multiplier)
AGE_MULT = {0: 1.00, 1: 0.95, 2: 0.90, 3: 0.85}

# Recency multipliers (years since signing → multiplier)
RECENCY_MULT = {0: 1.00, 1: 0.95, 2: 0.90, 3: 0.85}

# Positions to expand into if fewer than 5 exact comps found
EXPAND_TO = {"ED": ["IDL"], "IDL": ["ED"], "LB": ["ED"]}

# Positions with no similarity stats (OL, specialists)
NO_SIM_POSITIONS = {"LT", "RT", "LG", "RG", "C", "LS", "OL", "OG", "OT", "K", "P"}

# PFR position codes → normalized position key
POS_MAP = {
    "QB": "QB",
    "RB": "RB", "FB": "RB",
    "WR": "WR",
    "TE": "TE",
    "DE": "ED", "LDE": "ED", "RDE": "ED", "OLB": "ED", "LOLB": "ED", "ROLB": "ED",
    "DT": "IDL", "NT": "IDL", "LDT": "IDL", "RDT": "IDL", "DL": "IDL",
    "LB": "LB", "ILB": "LB", "MLB": "LB", "LLB": "LB", "RLB": "LB", "LILB": "LB", "RILB": "LB",
    "CB": "CB", "RCB": "CB", "LCB": "CB",
    "S": "S", "SS": "S", "FS": "S", "DB": "S",
    "K": "K", "P": "P",
    "LT": "LT", "RT": "RT", "LG": "LG", "RG": "RG", "C": "C", "LS": "LS",
}

RATE_COLS = {
    "PASS: Cmp%", "PASS: TD%", "PASS: Int%", "PASS: Succ%", "PASS: Y/A",
    "PASS: AY/A", "PASS: Y/C", "PASS: Y/G", "PASS: Rate", "PASS: QBR",
    "PASS: Sk%", "PASS: NY/A", "PASS: ANY/A", "PASS: QBrec",
    "RUSH: Rushing Succ%", "RUSH: Rushing Y/A", "RUSH: Rushing Y/G",
    "RUSH: Rushing A/G", "RUSH: Rushing Lng",
    "REC: Receiving Y/R", "REC: Receiving Succ%", "REC: Receiving Lng",
    "REC: Receiving R/G", "REC: Receiving Y/G", "REC: Receiving Ctch%", "REC: Receiving Y/Tgt",
    "DEF: Def Interceptions Lng", "SCR: Pts/G",
    "DEFA: Pass Coverage Cmp%", "DEFA: Pass Coverage Yds/Cmp",
    "DEFA: Pass Coverage Yds/Tgt", "DEFA: Pass Coverage Rat",
    "DEFA: Pass Coverage DADOT", "DEFA: Tackles MTkl%",
}

STAT_LABELS = {
    "DEF: Sk": "Sacks", "DEF: Tackles Comb": "Tackles (Combined)",
    "DEF: Tackles Solo": "Tackles (Solo)", "DEF: Tackles Ast": "Tackles (Assisted)",
    "DEF: Tackles TFL": "TFLs", "DEF: Tackles QBHits": "QB Hits",
    "DEF: Def Interceptions Int": "INTs", "DEF: Def Interceptions Yds": "INT Return Yds",
    "DEF: Def Interceptions IntTD": "INT TDs", "DEF: Def Interceptions Lng": "INT Long",
    "DEF: Def Interceptions PD": "Pass Deflections", "DEF: Fumbles FF": "Forced Fumbles",
    "DEF: Fumbles Fmb": "Fumbles", "DEF: Fumbles FR": "Fumble Recoveries",
    "DEF: Fumbles Yds": "Fumble Return Yds", "DEF: Fumbles FRTD": "Fumble Return TDs",
    "DEF: Sfty": "Safeties",
    "DEFA: Pass Rush Prss": "Pressures", "DEFA: Pass Rush Hrry": "Hurries",
    "DEFA: Pass Rush Sk": "Sacks (Advanced)", "DEFA: Pass Rush Bltz": "Blitzes",
    "DEFA: Pass Rush QBKD": "QB Knockdowns", "DEFA: Pass Rush Bats": "Batted Passes",
    "DEFA: Pass Coverage Tgt": "Coverage Targets", "DEFA: Pass Coverage Cmp": "Coverage Completions",
    "DEFA: Pass Coverage Cmp%": "Coverage Cmp%", "DEFA: Pass Coverage Yds": "Coverage Yards",
    "DEFA: Pass Coverage Yds/Cmp": "Coverage Yds/Cmp", "DEFA: Pass Coverage Yds/Tgt": "Coverage Yds/Tgt",
    "DEFA: Pass Coverage TD": "Coverage TDs Allowed", "DEFA: Pass Coverage Rat": "Coverage QB Rating",
    "DEFA: Pass Coverage DADOT": "Coverage ADOT", "DEFA: Pass Coverage Air": "Coverage Air Yds",
    "DEFA: Pass Coverage YAC": "Coverage YAC", "DEFA: Pass Coverage Int": "Coverage INTs",
    "DEFA: Tackles Comb": "Tackles (Adv Combined)", "DEFA: Tackles MTkl": "Missed Tackles",
    "DEFA: Tackles MTkl%": "Missed Tackle%",
    "PASS: Cmp": "Completions", "PASS: Att": "Attempts", "PASS: Cmp%": "Comp%",
    "PASS: Yds": "Pass Yards", "PASS: TD": "Pass TDs", "PASS: TD%": "TD%",
    "PASS: Int": "INTs Thrown", "PASS: Int%": "INT%", "PASS: 1D": "1st Downs",
    "PASS: Succ%": "Success%", "PASS: Lng": "Long", "PASS: Y/A": "Y/Attempt",
    "PASS: AY/A": "Adj Y/Attempt", "PASS: Y/C": "Y/Completion", "PASS: Y/G": "Pass Y/Game",
    "PASS: Rate": "Passer Rating", "PASS: QBR": "QBR", "PASS: Sk": "Times Sacked",
    "PASS: Yds.1": "Sack Yds Lost", "PASS: Sk%": "Sack%", "PASS: NY/A": "Net Y/Attempt",
    "PASS: ANY/A": "Adj Net Y/Attempt", "PASS: 4QC": "4th Qtr Comebacks",
    "PASS: GWD": "Game-Winning Drives", "PASS: QBrec": "QB Record",
    "RUSH: Rushing Att": "Carries", "RUSH: Rushing Yds": "Rush Yards",
    "RUSH: Rushing TD": "Rush TDs", "RUSH: Rushing 1D": "Rush 1st Downs",
    "RUSH: Rushing Succ%": "Rush Success%", "RUSH: Rushing Lng": "Rush Long",
    "RUSH: Rushing Y/A": "Y/Carry", "RUSH: Rushing Y/G": "Rush Y/Game",
    "RUSH: Rushing A/G": "Carries/Game", "RUSH: Fmb": "Fumbles (Rush)",
    "REC: Receiving Tgt": "Targets", "REC: Receiving Rec": "Receptions",
    "REC: Receiving Yds": "Rec Yards", "REC: Receiving Y/R": "Y/Reception",
    "REC: Receiving TD": "Rec TDs", "REC: Receiving 1D": "Rec 1st Downs",
    "REC: Receiving Succ%": "Rec Success%", "REC: Receiving Lng": "Rec Long",
    "REC: Receiving R/G": "Rec/Game", "REC: Receiving Y/G": "Rec Y/Game",
    "REC: Receiving Ctch%": "Catch%", "REC: Receiving Y/Tgt": "Y/Target",
    "REC: Fmb": "Fumbles (Rec)",
    "SCR: FG FGM": "FG Made", "SCR: FG FGA": "FG Attempted",
    "SCR: PAT XPM": "XP Made", "SCR: PAT XPA": "XP Attempted",
    "SCR: Pts": "Points", "SCR: Pts/G": "Points/Game",
    "SCR: Touchdowns RshTD": "Rush TDs", "SCR: Touchdowns RecTD": "Rec TDs",
    "SCR: Touchdowns AllTD": "Total TDs", "SCR: Touchdowns 2PM": "2-Pt Conversions",
    "SCR: Sfty": "Safeties",
}

WIDE  = "═" * 56
THIN  = "─" * 61
DIV   = "─" * 45
COL_S = 21   # stat label column
COL_C = 17   # comp player column
COL_O = 17   # original player column


# ── Helpers ───────────────────────────────────────────────────────────────────

def fmt(val):
    if pd.isna(val): return ""
    if isinstance(val, float): return f"{val:g}"
    return str(val)


def trunc(s, w):
    s = str(s)
    return s[:w - 1] + "…" if len(s) >= w else s


def stat_label(col):
    return STAT_LABELS.get(col, col)


def norm_pos(pfr_pos):
    return POS_MAP.get(str(pfr_pos).upper().strip(), str(pfr_pos).upper().strip())


def ask(prompt, default=""):
    suffix = f" [{default}]" if default else ""
    try:
        val = input(f"  {prompt}{suffix}: ").strip()
    except (EOFError, KeyboardInterrupt):
        print("\nGoodbye.")
        sys.exit(0)
    return val if val else default


def ask_int(prompt, default):
    while True:
        raw = ask(prompt, str(default))
        try:
            return int(raw)
        except ValueError:
            print("    Please enter a number.")


def ask_float(prompt):
    raw = ask(prompt)
    if not raw:
        return None
    try:
        return float(raw)
    except ValueError:
        print("    Not a number — skipped.")
        return None


# ── Data loading ──────────────────────────────────────────────────────────────

def load_data():
    print("Loading NFL_Contracts.xlsx...", end=" ", flush=True)
    ps = pd.read_excel(EXCEL_PATH, sheet_name="Player Stats")
    cc = pd.read_excel(EXCEL_PATH, sheet_name="Contracts + Context")
    print(f"✓  Player Stats: {len(ps):,} rows  |  Contracts+Context: {len(cc):,} rows")
    return ps, cc


# ── Option 1: Player Stats Lookup ─────────────────────────────────────────────

def option1_player_lookup(ps):
    name = ask("Player name (partial match OK)")
    matches = ps[ps["Player"].str.contains(name, case=False, na=False)]

    if matches.empty:
        print(f"\n  No players found matching '{name}'.\n")
        return

    names_found = matches["Player"].unique()
    if len(names_found) > 1:
        print("\n  Multiple matches:")
        for i, n in enumerate(names_found, 1):
            print(f"    {i}) {n}")
        choice = ask("Enter number to select")
        try:
            chosen = names_found[int(choice) - 1]
        except (ValueError, IndexError):
            print("  Invalid choice.")
            return
        matches = matches[matches["Player"] == chosen]

    player_name = matches["Player"].iloc[0]
    rows = matches.sort_values("Season")

    stat_cols_present = [
        c for c in PS_ALL_STATS
        if c in rows.columns and rows[c].notna().any()
    ]

    print(f"\n  {len(rows)} season(s) on file for {player_name}\n")

    for _, row in rows.iterrows():
        season = int(row["Season"])
        pos    = row.get("Pos", "")
        team   = row.get("Team", "")
        g      = row.get("G")
        gs     = row.get("GS")

        print(DIV)
        print(f"{player_name} — {pos} | {team}")
        print(DIV)
        games_str = ""
        if pd.notna(g):
            games_str = f"Games: {fmt(g)}"
            if pd.notna(gs):
                games_str += f" ({fmt(gs)} started)"
        print(f"{season}   {games_str}")

        for col in stat_cols_present:
            val = row.get(col)
            if pd.notna(val) and val != 0:
                print(f"  {stat_label(col)}: {fmt(val)}")

    summable = [c for c in stat_cols_present if c not in RATE_COLS]
    if len(rows) > 1 and summable:
        yr_min = int(rows["Season"].min())
        yr_max = int(rows["Season"].max())
        print(DIV)
        print(f"{player_name} — Combined ({yr_min}–{yr_max})")
        print(DIV)
        totals = rows[summable].sum(min_count=1)
        for col in summable:
            val = totals.get(col)
            if pd.notna(val) and val != 0:
                print(f"  {stat_label(col)}: {fmt(val)}")

    print(DIV + "\n")


# ── Option 2: Contract Comps by Player ────────────────────────────────────────

def _print_stat_table(label, stat_rows, comp_name, player_name):
    """Print a year/combined stats comparison table."""
    if not stat_rows:
        return
    cn = trunc(comp_name, COL_C)
    pn = trunc(player_name, COL_O)
    print(f"\n{label}")
    print("  " + "Stat".ljust(COL_S) + cn.ljust(COL_C) + pn.ljust(COL_O) + "Diff")
    print("  " + THIN)
    for lbl, cv, ov in stat_rows:
        cs  = fmt(cv)
        os_ = fmt(ov)
        if cv != 0:
            dp = (ov - cv) / abs(cv) * 100
            ds = ("+" if dp >= 0 else "") + f"{dp:.0f}%"
        else:
            ds = "—"
        print("  " + trunc(lbl, COL_S).ljust(COL_S) + cs.ljust(COL_C) + os_.ljust(COL_O) + ds)


def display_comp_card(rank, comp_row, player_name, og_season_rows, display_stats, sim_pct):
    comp_name = comp_row["Player"]
    comp_pos  = comp_row["Position"]
    comp_team = comp_row["Team"]
    yr_signed = comp_row["Year Signed"]
    length    = comp_row["Contract Length (Yrs)"]
    total_val = comp_row["Total Value ($M)"]
    aav       = comp_row["AAV ($M)"]
    gtd       = comp_row["Guaranteed ($M)"]
    age       = comp_row.get("Age at Signing (Rounded)", np.nan)

    print("\n" + WIDE)
    age_str = f" | Age at signing: {int(age)}" if pd.notna(age) else ""
    print(f"#{rank} — {comp_name} | {comp_pos} | {comp_team}{age_str}")

    parts = [f"Signed: {int(yr_signed)}" if pd.notna(yr_signed) else "Signed: ?"]
    if pd.notna(length):    parts.append(f"{fmt(length)} yrs")
    if pd.notna(total_val): parts.append(f"${total_val:.1f}M total")
    if pd.notna(aav):       parts.append(f"${aav:.2f}M AAV")
    if pd.notna(gtd):       parts.append(f"${gtd:.2f}M guaranteed")
    print(" | ".join(parts))
    print(f"Similarity: {sim_pct:.0f}%")
    print(WIDE)

    combined_comp = {}
    combined_og   = {}

    for slot, (og_yr, og_row) in enumerate(og_season_rows, start=1):
        prefix   = f"Y-{slot}: "
        comp_yr  = (int(yr_signed) - slot) if pd.notna(yr_signed) else None
        sec_lbl  = f"{comp_yr} Stats" if comp_yr else f"Y-{slot} Stats"

        stat_rows = []
        for stat in display_stats:
            cv = comp_row.get(f"{prefix}{stat}", np.nan)
            ov = og_row.get(stat, np.nan)
            if not (pd.notna(cv) and cv != 0 and pd.notna(ov) and ov != 0):
                continue
            cv, ov = float(cv), float(ov)
            stat_rows.append((stat_label(stat), cv, ov))
            combined_comp[stat] = combined_comp.get(stat, 0.0) + cv
            combined_og[stat]   = combined_og.get(stat, 0.0)   + ov

        _print_stat_table(sec_lbl, stat_rows, comp_name, player_name)

    if len(og_season_rows) > 1 and combined_comp:
        combined_rows = [
            (stat_label(s), combined_comp[s], combined_og[s])
            for s in display_stats
            if s in combined_comp and combined_comp[s] != 0 and combined_og.get(s, 0) != 0
        ]
        _print_stat_table("Combined", combined_rows, comp_name, player_name)

    print("\n" + WIDE)


def _score_comp_pool(pool, og_pos, og_sums, og_sign_age, og_exp, n_seasons):
    """
    Score a pool of contract comps against the original player.
    Returns pool with _sim, _score columns added, sorted by _sim descending.
    """
    pool = pool.copy()

    # Build per-stat normalization means from same-position players
    all_sim_stats = set()
    for stats in PAIR_SIM_STATS.values():
        all_sim_stats.update(stats)

    same_pos = pool[pool["Position"] == og_pos]
    stat_means = {}
    for stat in all_sim_stats:
        vals = []
        for _, row in same_pos.iterrows():
            v1 = row.get(f"Y-1: {stat}", np.nan)
            v2 = row.get(f"Y-2: {stat}", np.nan) if n_seasons == 2 else np.nan
            total = (float(v1) if pd.notna(v1) else 0.0) + (float(v2) if pd.notna(v2) else 0.0)
            if (pd.notna(v1) or (n_seasons == 2 and pd.notna(v2))) and total != 0:
                vals.append(total)
        stat_means[stat] = float(np.mean(vals)) if vals else 1.0

    sim_scores = []
    pos_proxies = []

    for _, row in pool.iterrows():
        comp_pos   = row["Position"]
        pair_stats = PAIR_SIM_STATS.get((og_pos, comp_pos))
        if pair_stats is None:
            sim_scores.append(np.nan)
            pos_proxies.append(np.nan)
            continue

        og_vec, comp_vec, norm_vec = [], [], []
        for stat in pair_stats:
            og_v = og_sums.get(stat)
            if og_v is None:
                continue
            v1 = row.get(f"Y-1: {stat}", np.nan)
            v2 = row.get(f"Y-2: {stat}", np.nan) if n_seasons == 2 else np.nan
            if not (pd.notna(v1) or (n_seasons == 2 and pd.notna(v2))):
                continue
            comp_v = (float(v1) if pd.notna(v1) else 0.0) + (float(v2) if pd.notna(v2) else 0.0)
            og_vec.append(og_v)
            comp_vec.append(comp_v)
            norm_vec.append(max(stat_means.get(stat, 1.0), 1e-9))

        if len(og_vec) < 2:
            sim_scores.append(np.nan)
            pos_proxies.append(np.nan)
            continue

        a    = np.array(og_vec,   dtype=float)
        b    = np.array(comp_vec, dtype=float)
        n    = np.array(norm_vec, dtype=float)
        dist = float(np.sqrt(np.sum(((a - b) / n) ** 2)))
        sim_scores.append(math.exp(-dist / len(og_vec)) * 100)
        pos_proxies.append(POS_PROX.get((og_pos, comp_pos), 1.0))

    pool["_sim"]      = sim_scores
    pool["_pos_prox"] = pos_proxies
    pool = pool[pool["_sim"].notna()].reset_index(drop=True)

    if pool.empty:
        return pool

    # Recency multiplier
    current_year = pd.Timestamp.now().year
    def recency_mult(yr):
        if pd.isna(yr): return 0.80
        diff = current_year - int(yr)
        return RECENCY_MULT.get(diff, 0.80)

    # Age/experience proximity multiplier
    EXP_MULT = {0: 1.00, 1: 0.90, 2: 0.80}
    if og_sign_age is not None and og_sign_age < 23:
        def age_mult(row):
            comp_exp = row.get("_exp", np.nan)
            if pd.isna(comp_exp) or og_exp is None: return 1.00
            diff = int(round(abs(float(comp_exp) - og_exp)))
            return EXP_MULT.get(diff, 0.70)
        pool["_age_mult"] = pool.apply(age_mult, axis=1)
    else:
        def age_mult_val(comp_age):
            if og_sign_age is None or pd.isna(comp_age): return 1.00
            diff = int(round(abs(float(comp_age) - og_sign_age)))
            return AGE_MULT.get(diff, 0.25)
        age_col = "Age at Signing (Rounded)"
        pool["_age_mult"] = pool[age_col].apply(age_mult_val) if age_col in pool.columns else 1.00

    pool["_recency"] = pool["Year Signed"].apply(recency_mult)
    pool["_score"]   = pool["_sim"] * pool["_recency"] * pool["_age_mult"] * pool["_pos_prox"]
    return pool


def option2_comps(ps, cc):
    # ── Player lookup ──────────────────────────────────────────────────────────
    name = ask("Player name to find comps for (partial match OK)")
    matches = ps[ps["Player"].str.contains(name, case=False, na=False)]

    if matches.empty:
        print(f"\n  '{name}' not found in Player Stats.\n")
        return

    unique_names = matches["Player"].unique()
    if len(unique_names) > 1:
        print("\n  Multiple matches:")
        for i, n in enumerate(unique_names, 1):
            print(f"    {i}) {n}")
        choice = ask("Enter number to select")
        try:
            player_name = unique_names[int(choice) - 1]
        except (ValueError, IndexError):
            print("  Invalid choice.")
            return
    else:
        player_name = unique_names[0]

    ps_player       = ps[ps["Player"] == player_name].sort_values("Season")
    seasons_on_file = sorted(ps_player["Season"].astype(int).tolist())
    print(f"\n  Seasons on file for {player_name}: {seasons_on_file}")

    while True:
        raw = ask("Enter the season year(s) to use, comma-separated (e.g. 2024 or 2024,2025)")
        try:
            chosen_years = sorted(set(int(y.strip()) for y in raw.split(",")))
        except ValueError:
            print("  Please enter years as numbers.")
            continue
        bad = [y for y in chosen_years if y not in seasons_on_file]
        if bad:
            print(f"  Not on file: {bad}. Choose from {seasons_on_file}.")
            continue
        if len(chosen_years) > 2:
            print("  Maximum 2 seasons supported.")
            continue
        break

    # og_season_rows: newest year first → Y-1, Y-2
    og_df = ps_player[ps_player["Season"].astype(int).isin(chosen_years)]
    og_season_rows = [
        (int(r["Season"]), r)
        for _, r in og_df.sort_values("Season", ascending=False).iterrows()
    ]
    og_pos   = norm_pos(og_season_rows[0][1]["Pos"])
    n_seasons = len(chosen_years)

    # ── OL / specialist fallback ───────────────────────────────────────────────
    if og_pos in NO_SIM_POSITIONS:
        print(f"\n  Note: Advanced stats are not available for {og_pos} in this dataset.")
        print("  Showing contract comps by position and age only.\n")
        fallback = cc[cc["Position"] == og_pos].copy()
        fallback = fallback[fallback["Player"] != player_name]
        fallback = (fallback.sort_values("Year Signed", ascending=False)
                            .drop_duplicates(subset=["Player"], keep="first"))
        og_cc = cc[cc["Player"] == player_name]
        if not og_cc.empty and "Age at Signing (Rounded)" in cc.columns:
            ages = og_cc["Age at Signing (Rounded)"].dropna()
            if not ages.empty:
                og_age = float(ages.sort_values().iloc[-1])
                fallback = fallback[(fallback["Age at Signing (Rounded)"] - og_age).abs() <= 3]
        top_fb = fallback.sort_values("Year Signed", ascending=False).head(10)
        if top_fb.empty:
            print("  No comps found.\n")
            return
        for i, (_, r) in enumerate(top_fb.iterrows(), start=1):
            age_s   = r.get("Age at Signing (Rounded)", np.nan)
            age_str = f" | Age at signing: {int(age_s)}" if pd.notna(age_s) else ""
            print(f"#{i} — {r['Player']} | {r['Position']} | {r['Team']}{age_str}")
            parts = [f"Signed: {int(r['Year Signed'])}" if pd.notna(r['Year Signed']) else "Signed: ?"]
            if pd.notna(r['Contract Length (Yrs)']): parts.append(f"{fmt(r['Contract Length (Yrs)'])} yrs")
            if pd.notna(r['Total Value ($M)']):       parts.append(f"${r['Total Value ($M)']:.1f}M total")
            if pd.notna(r['AAV ($M)']):               parts.append(f"${r['AAV ($M)']:.2f}M AAV")
            if pd.notna(r['Guaranteed ($M)']):        parts.append(f"${r['Guaranteed ($M)']:.2f}M guaranteed")
            print("  " + " | ".join(parts))
        print()
        return

    # ── Build OG player's combined stats across chosen seasons ─────────────────
    all_sim_stats = set()
    for stats in PAIR_SIM_STATS.values():
        all_sim_stats.update(stats)

    og_sums = {}
    for _, og_row in og_season_rows:
        for stat in all_sim_stats:
            v = og_row.get(stat, np.nan)
            if pd.notna(v):
                og_sums[stat] = og_sums.get(stat, 0.0) + float(v)

    if not og_sums:
        print(f"\n  No usable stats found for {player_name}.\n")
        return

    # OG signing age from most recent CC contract row
    og_cc_rows  = cc[cc["Player"] == player_name]
    og_sign_age = None
    if not og_cc_rows.empty and "Age at Signing (Rounded)" in cc.columns:
        ages = og_cc_rows["Age at Signing (Rounded)"].dropna()
        if not ages.empty:
            og_sign_age = float(ages.sort_values(ascending=False).iloc[0])

    # OG experience (years from first PS season through most recent chosen season)
    og_exp = None
    og_first_season = int(ps_player["Season"].min()) if not ps_player.empty else None
    if og_first_season is not None:
        og_exp = int(chosen_years[-1]) - og_first_season + 1

    # ── Build base comp pool ───────────────────────────────────────────────────
    base_pool = (cc[cc["Player"] != player_name]
                   .sort_values("Year Signed", ascending=False)
                   .drop_duplicates(subset=["Player"], keep="first")
                   .reset_index(drop=True))

    # Add experience column for young-player mode
    if og_sign_age is not None and og_sign_age < 23:
        ps_first = (ps.groupby("Player")["Season"].min().reset_index()
                      .rename(columns={"Season": "_first_season"}))
        base_pool = base_pool.merge(ps_first, on="Player", how="left")
        base_pool["_exp"] = base_pool["Year Signed"] - base_pool["_first_season"]
        age_col = "Age at Signing (Rounded)"
        if age_col in base_pool.columns:
            base_pool = base_pool[base_pool[age_col].fillna(99) <= (og_sign_age + 6)].reset_index(drop=True)
        if og_exp is not None:
            base_pool = base_pool[(base_pool["_exp"] - og_exp).abs() <= 2].reset_index(drop=True)
    else:
        base_pool["_exp"] = np.nan

    # Filter: minimum AAV $3M
    if "AAV ($M)" in base_pool.columns:
        base_pool = base_pool[base_pool["AAV ($M)"].fillna(0) >= 3.0].reset_index(drop=True)

    # Filter: minimum activity thresholds to remove non-contributors
    if og_pos in {"ED", "IDL"}:
        sk  = base_pool.get("Y-1: DEF: Sk",            pd.Series(0.0, index=base_pool.index)).fillna(0)
        prs = base_pool.get("Y-1: DEFA: Pass Rush Prss", pd.Series(0.0, index=base_pool.index)).fillna(0)
        base_pool = base_pool[(sk + prs) >= 5].reset_index(drop=True)
    elif og_pos == "LB":
        tck = base_pool.get("Y-1: DEF: Tackles Comb", pd.Series(0.0, index=base_pool.index)).fillna(0)
        base_pool = base_pool[tck >= 20].reset_index(drop=True)

    # Filter: must have at least one Y-1 stat
    y1_cols = [c for c in base_pool.columns if c.startswith("Y-1: ")]
    if y1_cols:
        base_pool = base_pool[base_pool[y1_cols].notna().any(axis=1)].reset_index(drop=True)

    # Filter: age within 3 years (23+ players only)
    age_col = "Age at Signing (Rounded)"
    if og_sign_age is not None and og_sign_age >= 23 and age_col in base_pool.columns:
        base_pool = base_pool[(base_pool[age_col] - og_sign_age).abs() <= 3].reset_index(drop=True)

    # ── Score exact position pool first ───────────────────────────────────────
    exact_pool = base_pool[base_pool["Position"] == og_pos].copy()
    scored = _score_comp_pool(exact_pool, og_pos, og_sums, og_sign_age, og_exp, n_seasons) if not exact_pool.empty else pd.DataFrame()

    # Expand to similar positions if fewer than 5 exact comps
    if len(scored) < 5 and og_pos in EXPAND_TO:
        print(f"\n  Fewer than 5 exact position comps — expanding to similar positions.")
        expand_positions = [og_pos] + EXPAND_TO[og_pos]
        expanded_pool = base_pool[base_pool["Position"].isin(expand_positions)].copy()
        scored = _score_comp_pool(expanded_pool, og_pos, og_sums, og_sign_age, og_exp, n_seasons) if not expanded_pool.empty else pd.DataFrame()

    if scored.empty:
        print("\n  Not enough overlapping stats to score any comps.\n")
        return

    # ── Top 10: rank by _score, display ordered by _sim descending ────────────
    top10 = (scored
             .sort_values("_score", ascending=False)
             .head(10)
             .sort_values("_sim", ascending=False)
             .reset_index(drop=True))

    display_stats = [c for c in POSITION_STATS.get(og_pos, []) if c in ps.columns]
    yr_label      = f"{chosen_years[0]}–{chosen_years[-1]}" if len(chosen_years) > 1 else str(chosen_years[0])
    og_age_str    = f" | Age at signing: {int(og_sign_age)}" if og_sign_age is not None else ""

    print(f"\n  Top 10 comps for {player_name} ({og_pos}){og_age_str} | Seasons: {yr_label}\n")

    for rank, (_, comp_row) in enumerate(top10.iterrows(), start=1):
        display_comp_card(
            rank           = rank,
            comp_row       = comp_row,
            player_name    = player_name,
            og_season_rows = og_season_rows,
            display_stats  = display_stats,
            sim_pct        = float(comp_row["_sim"]),
        )

    _neg_summary(player_name, og_pos, og_sign_age, og_season_rows, top10, n_seasons)


def _neg_summary(player_name, og_pos, og_sign_age, og_season_rows, top10, n_seasons):
    adj_stats = ADJ_STATS.get(og_pos)
    if adj_stats is None:
        return

    # OG combined stats across chosen seasons
    og_combined = {}
    for _, og_row in og_season_rows:
        for stat in adj_stats:
            v = og_row.get(stat, np.nan)
            if pd.notna(v) and float(v) != 0:
                og_combined[stat] = og_combined.get(stat, 0.0) + float(v)

    COL_N  = 21
    COL_R  = 10
    COL_A  = 14
    COL_AJ = 10

    rows_out = []  # (name, raw_aav, adj_factor, adj_aav, raw_gtd, adj_gtd, sim)

    for _, comp_row in top10.iterrows():
        raw_aav = comp_row.get("AAV ($M)", np.nan)
        raw_gtd = comp_row.get("Guaranteed ($M)", np.nan)
        sim     = float(comp_row.get("_sim", 0.0))
        if pd.isna(raw_aav) or raw_aav <= 0:
            continue

        # Comp combined stats
        comp_combined = {}
        for stat in adj_stats:
            v1 = comp_row.get(f"Y-1: {stat}", np.nan)
            v2 = comp_row.get(f"Y-2: {stat}", np.nan) if n_seasons == 2 else np.nan
            total = (float(v1) if pd.notna(v1) else 0.0) + (float(v2) if pd.notna(v2) else 0.0)
            if (pd.notna(v1) or (n_seasons == 2 and pd.notna(v2))) and total != 0:
                comp_combined[stat] = total

        # Per-stat % differences where both have real values
        diffs = []
        for stat in adj_stats:
            og_v   = og_combined.get(stat)
            comp_v = comp_combined.get(stat)
            if og_v is None or comp_v is None or comp_v == 0:
                continue
            diffs.append((og_v - comp_v) / comp_v * 100)

        if not diffs:
            adj_factor = 0.0
        else:
            raw_adj = sum(diffs) / len(diffs)
            # Dampen by confidence: fewer matching stats = smaller adjustment
            confidence = len(diffs) / len(adj_stats)
            adj_factor = raw_adj * confidence
            # Cap at ±75%
            adj_factor = max(-75.0, min(75.0, adj_factor))

        adj_aav = raw_aav * (1 + adj_factor / 100)
        adj_gtd = raw_gtd * (1 + adj_factor / 100) if pd.notna(raw_gtd) else np.nan

        rows_out.append((comp_row["Player"], raw_aav, adj_factor, adj_aav, raw_gtd, adj_gtd, sim))

    if not rows_out:
        return

    adj_aavs = [r[3] for r in rows_out]
    sims     = [r[6] for r in rows_out]

    total_sim  = sum(sims)
    target_aav = sum(r[3] * r[6] for r in rows_out) / total_sim if total_sim > 0 else np.nan

    gtd_rows   = [(r[5], r[6]) for r in rows_out if pd.notna(r[5])]
    target_gtd = (sum(g * s for g, s in gtd_rows) / sum(s for _, s in gtd_rows) if gtd_rows else np.nan)

    floor   = float(np.percentile(adj_aavs, 25))
    ceiling = float(np.percentile(adj_aavs, 75))

    age_str = f" | Age: {int(og_sign_age)}" if og_sign_age is not None else ""
    print("\n" + WIDE)
    print(f"NEGOTIATION SUMMARY — {player_name} | {og_pos}{age_str}")
    print(WIDE)

    print(f"\nPerformance-Adjusted Comps:")
    print("  " + "Player".ljust(COL_N) + "Raw AAV".ljust(COL_R) + "Adjustment".ljust(COL_A) + "Adj AAV")
    print("  " + "─" * (COL_N + COL_R + COL_A + COL_AJ))
    for comp_name, raw_aav, adj_factor, adj_aav, _, _, _ in rows_out:
        sign = "+" if adj_factor >= 0 else ""
        print("  " + trunc(comp_name, COL_N).ljust(COL_N)
              + f"${raw_aav:.1f}M".ljust(COL_R)
              + f"{sign}{adj_factor:.0f}%".ljust(COL_A)
              + f"${adj_aav:.1f}M")

    print(f"\nNegotiation Range:")
    print(f"  Floor (25th pct):    ${floor:.1f}M AAV")
    print(f"  Target (weighted):   ${target_aav:.1f}M AAV")
    print(f"  Ceiling (75th pct):  ${ceiling:.1f}M AAV")

    if pd.notna(target_gtd):
        print(f"\nTarget Guaranteed:     ${target_gtd:.1f}M")

    top3 = sorted(rows_out, key=lambda r: r[6], reverse=True)[:3]
    print(f"\nTop 3 Strongest Comps by Similarity:")
    for comp_name, raw_aav, _, _, _, _, sim in top3:
        yr = top10.loc[top10["Player"] == comp_name, "Year Signed"]
        yr_str = str(int(yr.iloc[0])) if not yr.empty and pd.notna(yr.iloc[0]) else "?"
        print(f"  {trunc(comp_name, COL_N).ljust(COL_N)}{sim:.0f}%    ${raw_aav:.1f}M AAV    {yr_str}")

    print(WIDE + "\n")


# ── Option 3: Open Filter Query ───────────────────────────────────────────────

def option3_filter(ps, cc):
    print()
    pool = cc.copy()
    filters = []

    positions = sorted(cc["Position"].dropna().unique())
    print(f"  Positions: {', '.join(positions)}")
    pos_input = ask("Position (or Enter for all)").upper()
    if pos_input and pos_input in positions:
        pool = pool[pool["Position"] == pos_input]
        filters.append(f"Position = {pos_input}")
    elif pos_input:
        print(f"  '{pos_input}' not in data — skipping.")
        pos_input = ""

    yr = ask("Signing year or range e.g. 2023 or 2021-2023 (Enter for all)", "")
    if yr:
        try:
            if "-" in yr:
                lo, hi = yr.split("-")
                pool = pool[pool["Year Signed"].between(int(lo), int(hi))]
                filters.append(f"Year {lo}–{hi}")
            else:
                pool = pool[pool["Year Signed"] == int(yr)]
                filters.append(f"Year = {yr}")
        except ValueError:
            print("  Invalid — skipped.")

    rel_stats = [c for c in POSITION_STATS.get(pos_input, []) if f"Y-1: {c}" in cc.columns] if pos_input else []
    if rel_stats:
        print(f"\n  Stat filters for {pos_input} — Y-1 season (Enter to skip any):")
        for stat in rel_stats:
            col = f"Y-1: {stat}"
            threshold = ask_float(f"  Min {stat_label(stat)}")
            if threshold is not None:
                pool = pool[pool[col].fillna(0) >= threshold]
                filters.append(f"Y-1 {stat_label(stat)} ≥ {threshold}")

    print("\n  Contract filters (Enter to skip):")
    aav_min = ask_float("  Min AAV ($M)")
    if aav_min is not None:
        pool = pool[pool["AAV ($M)"] >= aav_min]
        filters.append(f"AAV ≥ ${aav_min}M")
    aav_max = ask_float("  Max AAV ($M)")
    if aav_max is not None:
        pool = pool[pool["AAV ($M)"] <= aav_max]
        filters.append(f"AAV ≤ ${aav_max}M")
    gtd_min = ask_float("  Min Guaranteed ($M)")
    if gtd_min is not None:
        pool = pool[pool["Guaranteed ($M)"] >= gtd_min]
        filters.append(f"Guaranteed ≥ ${gtd_min}M")

    print("\n  Sort by: 1) AAV  2) Guaranteed  3) Total Value  4) Year Signed")
    sort_map = {"1": "AAV ($M)", "2": "Guaranteed ($M)", "3": "Total Value ($M)", "4": "Year Signed"}
    sort_col = sort_map.get(ask("Choice", "1"), "AAV ($M)")
    pool = pool.sort_values(sort_col, ascending=False)

    top_n = ask_int("How many results?", 20)
    pool = pool.head(top_n)

    if pool.empty:
        print("  (no results)\n")
        return

    title = "Results — " + (", ".join(filters) if filters else "no filters")
    print(f"\n  {title}")
    print(f"  {len(pool)} result(s)\n")

    show_stats = POSITION_STATS.get(pos_input, []) if pos_input else []

    for _, row in pool.iterrows():
        player    = row["Player"]
        pos       = row["Position"]
        team      = row["Team"]
        yr_s      = row["Year Signed"]
        length    = row["Contract Length (Yrs)"]
        total_val = row["Total Value ($M)"]
        aav       = row["AAV ($M)"]
        gtd       = row["Guaranteed ($M)"]

        print(DIV)
        print(f"{player} — {pos} | {team}")
        print(DIV)

        parts = [f"Signed: {int(yr_s)}" if pd.notna(yr_s) else "Signed: ?"]
        if pd.notna(length):    parts.append(f"{fmt(length)} yrs")
        if pd.notna(total_val): parts.append(f"${total_val:.1f}M total")
        if pd.notna(aav):       parts.append(f"${aav:.2f}M AAV")
        if pd.notna(gtd):       parts.append(f"${gtd:.2f}M guaranteed")
        print(" | ".join(parts))

        for prefix, offset_n in [("Y-1: ", 1), ("Y-2: ", 2)]:
            yr_label = f"{int(yr_s) - offset_n} Stats" if pd.notna(yr_s) else f"Y-{offset_n} Stats"
            section_vals = []
            for stat in show_stats:
                col = f"{prefix}{stat}"
                if col not in cc.columns: continue
                val = row.get(col)
                if pd.notna(val) and val != 0:
                    section_vals.append((stat_label(stat), val))
            if section_vals:
                print(f"{yr_label}:")
                for lbl, val in section_vals:
                    print(f"  {lbl}: {fmt(val)}")

    print(DIV + "\n")


# ── Main menu ─────────────────────────────────────────────────────────────────

MENU = """
╔══════════════════════════════════════════╗
║       NFL Contracts Query Tool           ║
╠══════════════════════════════════════════╣
║  1 — Player Stats Lookup                 ║
║  2 — Find Contract Comps by Player       ║
║  3 — Open Filter Query                   ║
║  q — Quit                                ║
╚══════════════════════════════════════════╝"""


def main():
    ps, cc = load_data()
    while True:
        print(MENU)
        try:
            choice = input("  Choice> ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            print("\nGoodbye.")
            break

        if choice in ("q", "quit", "exit"):
            print("Goodbye.")
            break
        elif choice == "1":
            option1_player_lookup(ps)
        elif choice == "2":
            option2_comps(ps, cc)
        elif choice == "3":
            option3_filter(ps, cc)
        else:
            print("  Enter 1, 2, 3, or q.")


if __name__ == "__main__":
    main()
