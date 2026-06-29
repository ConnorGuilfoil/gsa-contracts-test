#!/usr/bin/env python3
"""
GSA Contract Comp Finder — Streamlit App
"""

import math
import streamlit as st
import pandas as pd
import numpy as np

EXCEL_PATH = "/Users/connorguilfoil/Desktop/Grady Sports/NFL_Contracts.xlsx"

# ── Constants (identical to query_contracts.py) ───────────────────────────────

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

POS_PROX = {
    ("ED","ED"):1.00,("IDL","IDL"):1.00,("LB","LB"):1.00,
    ("ED","IDL"):0.85,("IDL","ED"):0.85,
    ("ED","LB"):0.70,("LB","ED"):0.70,
    ("CB","CB"):1.00,("S","S"):1.00,
    ("CB","S"):0.85,("S","CB"):0.85,
    ("RB","RB"):1.00,("WR","WR"):1.00,
    ("TE","TE"):1.00,("WR","TE"):0.90,("TE","WR"):0.90,
    ("QB","QB"):1.00,
}

AGE_MULT     = {0:1.00,1:0.95,2:0.90,3:0.85}
RECENCY_MULT = {0:1.00,1:0.95,2:0.90,3:0.85}
EXPAND_TO    = {"ED":["IDL"],"IDL":["ED"],"LB":["ED"]}
NO_SIM_POS   = {"LT","RT","LG","RG","C","LS","OL","OG","OT","K","P"}

POS_MAP = {
    "QB":"QB","RB":"RB","FB":"RB","WR":"WR","TE":"TE",
    "DE":"ED","LDE":"ED","RDE":"ED","OLB":"ED","LOLB":"ED","ROLB":"ED",
    "DT":"IDL","NT":"IDL","LDT":"IDL","RDT":"IDL","DL":"IDL",
    "LB":"LB","ILB":"LB","MLB":"LB","LLB":"LB","RLB":"LB","LILB":"LB","RILB":"LB",
    "CB":"CB","RCB":"CB","LCB":"CB",
    "S":"S","SS":"S","FS":"S","DB":"S",
    "K":"K","P":"P",
    "LT":"LT","RT":"RT","LG":"LG","RG":"RG","C":"C","LS":"LS",
}

STAT_LABELS = {
    "DEF: Sk":"Sacks","DEF: Tackles Comb":"Tackles (Combined)",
    "DEF: Tackles Solo":"Tackles (Solo)","DEF: Tackles Ast":"Tackles (Assisted)",
    "DEF: Tackles TFL":"TFLs","DEF: Tackles QBHits":"QB Hits",
    "DEF: Def Interceptions Int":"INTs","DEF: Def Interceptions PD":"Pass Deflections",
    "DEF: Fumbles FF":"Forced Fumbles","DEF: Fumbles FR":"Fumble Recoveries",
    "DEF: Sfty":"Safeties",
    "DEFA: Pass Rush Prss":"Pressures","DEFA: Pass Rush Hrry":"Hurries",
    "DEFA: Pass Coverage Tgt":"Coverage Targets","DEFA: Pass Coverage Cmp%":"Coverage Cmp%",
    "DEFA: Pass Coverage Yds":"Coverage Yards","DEFA: Pass Coverage TD":"Coverage TDs Allowed",
    "DEFA: Pass Coverage Rat":"Coverage QB Rating","DEFA: Tackles MTkl":"Missed Tackles",
    "PASS: Cmp":"Completions","PASS: Att":"Attempts","PASS: Cmp%":"Comp%",
    "PASS: Yds":"Pass Yards","PASS: TD":"Pass TDs","PASS: Int":"INTs Thrown",
    "PASS: Y/A":"Y/Attempt","PASS: Rate":"Passer Rating","PASS: QBR":"QBR",
    "RUSH: Rushing Att":"Carries","RUSH: Rushing Yds":"Rush Yards",
    "RUSH: Rushing TD":"Rush TDs","RUSH: Rushing Y/A":"Y/Carry",
    "REC: Receiving Tgt":"Targets","REC: Receiving Rec":"Receptions",
    "REC: Receiving Yds":"Rec Yards","REC: Receiving TD":"Rec TDs",
    "REC: Receiving Ctch%":"Catch%","REC: Receiving Y/R":"Y/Reception",
    "SCR: FG FGM":"FG Made","SCR: FG FGA":"FG Attempted",
    "SCR: PAT XPM":"XP Made","SCR: Pts":"Points",
}

def stat_label(col):
    return STAT_LABELS.get(col, col)

def norm_pos(pfr_pos):
    return POS_MAP.get(str(pfr_pos).upper().strip(), str(pfr_pos).upper().strip())

def fmt(val):
    if pd.isna(val): return ""
    if isinstance(val, float): return f"{val:g}"
    return str(val)


# ── Data loading ──────────────────────────────────────────────────────────────

@st.cache_data
def load_data():
    ps = pd.read_excel(EXCEL_PATH, sheet_name="Player Stats")
    cc = pd.read_excel(EXCEL_PATH, sheet_name="Contracts + Context")
    return ps, cc


# ── Scoring (identical logic to query_contracts.py) ───────────────────────────

def score_pool(pool, og_pos, og_sums, og_sign_age, og_exp, n_seasons):
    pool = pool.copy()
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

    sim_scores, pos_proxies = [], []
    for _, row in pool.iterrows():
        comp_pos   = row["Position"]
        pair_stats = PAIR_SIM_STATS.get((og_pos, comp_pos))
        if pair_stats is None:
            sim_scores.append(np.nan); pos_proxies.append(np.nan); continue

        og_vec, comp_vec, norm_vec = [], [], []
        for stat in pair_stats:
            og_v = og_sums.get(stat)
            if og_v is None: continue
            v1 = row.get(f"Y-1: {stat}", np.nan)
            v2 = row.get(f"Y-2: {stat}", np.nan) if n_seasons == 2 else np.nan
            if not (pd.notna(v1) or (n_seasons == 2 and pd.notna(v2))): continue
            comp_v = (float(v1) if pd.notna(v1) else 0.0) + (float(v2) if pd.notna(v2) else 0.0)
            og_vec.append(og_v); comp_vec.append(comp_v)
            norm_vec.append(max(stat_means.get(stat, 1.0), 1e-9))

        if len(og_vec) < 2:
            sim_scores.append(np.nan); pos_proxies.append(np.nan); continue

        a = np.array(og_vec, dtype=float)
        b = np.array(comp_vec, dtype=float)
        n = np.array(norm_vec, dtype=float)
        dist = float(np.sqrt(np.sum(((a - b) / n) ** 2)))
        sim_scores.append(math.exp(-dist / len(og_vec)) * 100)
        pos_proxies.append(POS_PROX.get((og_pos, comp_pos), 1.0))

    pool["_sim"]      = sim_scores
    pool["_pos_prox"] = pos_proxies
    pool = pool[pool["_sim"].notna()].reset_index(drop=True)
    if pool.empty: return pool

    current_year = pd.Timestamp.now().year
    def rec(yr): 
        if pd.isna(yr): return 0.80
        return RECENCY_MULT.get(current_year - int(yr), 0.80)

    EXP_MULT = {0:1.00,1:0.90,2:0.80}
    if og_sign_age is not None and og_sign_age < 23:
        def am(row):
            ce = row.get("_exp", np.nan)
            if pd.isna(ce) or og_exp is None: return 1.00
            return EXP_MULT.get(int(round(abs(float(ce) - og_exp))), 0.70)
        pool["_age_mult"] = pool.apply(am, axis=1)
    else:
        def amv(ca):
            if og_sign_age is None or pd.isna(ca): return 1.00
            return AGE_MULT.get(int(round(abs(float(ca) - og_sign_age))), 0.25)
        ac = "Age at Signing (Rounded)"
        pool["_age_mult"] = pool[ac].apply(amv) if ac in pool.columns else 1.00

    pool["_recency"] = pool["Year Signed"].apply(rec)
    pool["_score"]   = pool["_sim"] * pool["_recency"] * pool["_age_mult"] * pool["_pos_prox"]
    return pool


def get_comps(ps, cc, player_name, chosen_years):
    ps_player = ps[ps["Player"] == player_name].sort_values("Season")
    og_df     = ps_player[ps_player["Season"].astype(int).isin(chosen_years)]
    og_season_rows = [
        (int(r["Season"]), r)
        for _, r in og_df.sort_values("Season", ascending=False).iterrows()
    ]
    og_pos    = norm_pos(og_season_rows[0][1]["Pos"])
    n_seasons = len(chosen_years)

    all_sim_stats = set()
    for stats in PAIR_SIM_STATS.values():
        all_sim_stats.update(stats)

    og_sums = {}
    for _, og_row in og_season_rows:
        for stat in all_sim_stats:
            v = og_row.get(stat, np.nan)
            if pd.notna(v):
                og_sums[stat] = og_sums.get(stat, 0.0) + float(v)

    og_cc_rows  = cc[cc["Player"] == player_name]
    og_sign_age = None
    if not og_cc_rows.empty and "Age at Signing (Rounded)" in cc.columns:
        ages = og_cc_rows["Age at Signing (Rounded)"].dropna()
        if not ages.empty:
            og_sign_age = float(ages.sort_values(ascending=False).iloc[0])

    og_exp = None
    og_first = int(ps_player["Season"].min()) if not ps_player.empty else None
    if og_first is not None:
        og_exp = int(chosen_years[-1]) - og_first + 1

    base_pool = (cc[cc["Player"] != player_name]
                   .sort_values("Year Signed", ascending=False)
                   .drop_duplicates(subset=["Player"], keep="first")
                   .reset_index(drop=True))

    if og_sign_age is not None and og_sign_age < 23:
        ps_first = (ps.groupby("Player")["Season"].min().reset_index()
                      .rename(columns={"Season": "_first_season"}))
        base_pool = base_pool.merge(ps_first, on="Player", how="left")
        base_pool["_exp"] = base_pool["Year Signed"] - base_pool["_first_season"]
        ac = "Age at Signing (Rounded)"
        if ac in base_pool.columns:
            base_pool = base_pool[base_pool[ac].fillna(99) <= (og_sign_age + 6)].reset_index(drop=True)
        if og_exp is not None:
            base_pool = base_pool[(base_pool["_exp"] - og_exp).abs() <= 2].reset_index(drop=True)
    else:
        base_pool["_exp"] = np.nan

    if "AAV ($M)" in base_pool.columns:
        base_pool = base_pool[base_pool["AAV ($M)"].fillna(0) >= 3.0].reset_index(drop=True)

    if og_pos in {"ED","IDL"}:
        sk  = base_pool.get("Y-1: DEF: Sk",             pd.Series(0.0, index=base_pool.index)).fillna(0)
        prs = base_pool.get("Y-1: DEFA: Pass Rush Prss", pd.Series(0.0, index=base_pool.index)).fillna(0)
        base_pool = base_pool[(sk + prs) >= 5].reset_index(drop=True)
    elif og_pos == "LB":
        tck = base_pool.get("Y-1: DEF: Tackles Comb", pd.Series(0.0, index=base_pool.index)).fillna(0)
        base_pool = base_pool[tck >= 20].reset_index(drop=True)

    y1_cols = [c for c in base_pool.columns if c.startswith("Y-1: ")]
    if y1_cols:
        base_pool = base_pool[base_pool[y1_cols].notna().any(axis=1)].reset_index(drop=True)

    ac = "Age at Signing (Rounded)"
    if og_sign_age is not None and og_sign_age >= 23 and ac in base_pool.columns:
        base_pool = base_pool[(base_pool[ac] - og_sign_age).abs() <= 3].reset_index(drop=True)

    exact_pool = base_pool[base_pool["Position"] == og_pos].copy()
    scored = score_pool(exact_pool, og_pos, og_sums, og_sign_age, og_exp, n_seasons) if not exact_pool.empty else pd.DataFrame()

    expanded = False
    if len(scored) < 5 and og_pos in EXPAND_TO:
        expanded = True
        expand_positions = [og_pos] + EXPAND_TO[og_pos]
        expanded_pool = base_pool[base_pool["Position"].isin(expand_positions)].copy()
        scored = score_pool(expanded_pool, og_pos, og_sums, og_sign_age, og_exp, n_seasons) if not expanded_pool.empty else pd.DataFrame()

    if scored.empty:
        return None, og_pos, og_sign_age, og_season_rows, expanded

    top10 = (scored
             .sort_values("_score", ascending=False)
             .head(10)
             .sort_values("_sim", ascending=False)
             .reset_index(drop=True))

    return top10, og_pos, og_sign_age, og_season_rows, expanded


def calc_neg_summary(og_pos, og_season_rows, top10, n_seasons):
    adj_stats = ADJ_STATS.get(og_pos)
    if adj_stats is None or top10 is None:
        return None

    og_combined = {}
    for _, og_row in og_season_rows:
        for stat in adj_stats:
            v = og_row.get(stat, np.nan)
            if pd.notna(v) and float(v) != 0:
                og_combined[stat] = og_combined.get(stat, 0.0) + float(v)

    rows_out = []
    for _, comp_row in top10.iterrows():
        raw_aav = comp_row.get("AAV ($M)", np.nan)
        raw_gtd = comp_row.get("Guaranteed ($M)", np.nan)
        sim     = float(comp_row.get("_sim", 0.0))
        if pd.isna(raw_aav) or raw_aav <= 0: continue

        comp_combined = {}
        for stat in adj_stats:
            v1 = comp_row.get(f"Y-1: {stat}", np.nan)
            v2 = comp_row.get(f"Y-2: {stat}", np.nan) if n_seasons == 2 else np.nan
            total = (float(v1) if pd.notna(v1) else 0.0) + (float(v2) if pd.notna(v2) else 0.0)
            if (pd.notna(v1) or (n_seasons == 2 and pd.notna(v2))) and total != 0:
                comp_combined[stat] = total

        diffs = []
        for stat in adj_stats:
            og_v   = og_combined.get(stat)
            comp_v = comp_combined.get(stat)
            if og_v is None or comp_v is None or comp_v == 0: continue
            diffs.append((og_v - comp_v) / comp_v * 100)

        if not diffs:
            adj_factor = 0.0
        else:
            confidence = len(diffs) / len(adj_stats)
            adj_factor = max(-75.0, min(75.0, (sum(diffs) / len(diffs)) * confidence))

        adj_aav = raw_aav * (1 + adj_factor / 100)
        adj_gtd = raw_gtd * (1 + adj_factor / 100) if pd.notna(raw_gtd) else np.nan
        rows_out.append((comp_row["Player"], raw_aav, adj_factor, adj_aav, raw_gtd, adj_gtd, sim))

    if not rows_out:
        return None

    adj_aavs   = [r[3] for r in rows_out]
    sims       = [r[6] for r in rows_out]
    total_sim  = sum(sims)
    target_aav = sum(r[3] * r[6] for r in rows_out) / total_sim if total_sim > 0 else np.nan
    gtd_rows   = [(r[5], r[6]) for r in rows_out if pd.notna(r[5])]
    target_gtd = sum(g * s for g, s in gtd_rows) / sum(s for _, s in gtd_rows) if gtd_rows else np.nan
    floor      = float(np.percentile(adj_aavs, 25))
    ceiling    = float(np.percentile(adj_aavs, 75))

    return {
        "rows": rows_out,
        "floor": floor,
        "target_aav": target_aav,
        "ceiling": ceiling,
        "target_gtd": target_gtd,
    }


# ── Streamlit UI ──────────────────────────────────────────────────────────────

st.set_page_config(page_title="GSA Contract Comp Finder", layout="wide")
st.title("GSA Contract Comp Finder")

with st.spinner("Loading data..."):
    ps, cc = load_data()

st.sidebar.header("Search")
player_input = st.sidebar.text_input("Player name", placeholder="e.g. Braden Fiske")

if player_input:
    matches = ps[ps["Player"].str.contains(player_input, case=False, na=False)]
    if matches.empty:
        st.sidebar.error(f"No players found matching '{player_input}'")
    else:
        names_found = sorted(matches["Player"].unique())
        selected_name = st.sidebar.selectbox("Select player", names_found)

        ps_player       = ps[ps["Player"] == selected_name].sort_values("Season")
        seasons_on_file = sorted(ps_player["Season"].astype(int).tolist())

        chosen_years = st.sidebar.multiselect(
            "Season(s) to use (max 2)",
            options=seasons_on_file,
            default=[seasons_on_file[-1]],
        )

        if len(chosen_years) > 2:
            st.sidebar.warning("Maximum 2 seasons — using the 2 most recent.")
            chosen_years = sorted(chosen_years)[-2:]

        run = st.sidebar.button("Find Comps", type="primary")

        if run and chosen_years:
            chosen_years = sorted(chosen_years)
            with st.spinner("Finding comps..."):
                top10, og_pos, og_sign_age, og_season_rows, expanded = get_comps(ps, cc, selected_name, chosen_years)

            if top10 is None:
                st.error("Not enough overlapping stats to score any comps.")
            else:
                if expanded:
                    st.info("Fewer than 5 exact position comps found — expanded to similar positions.")

                yr_label   = f"{chosen_years[0]}–{chosen_years[-1]}" if len(chosen_years) > 1 else str(chosen_years[0])
                age_str    = f" | Age at signing: {int(og_sign_age)}" if og_sign_age is not None else ""
                st.subheader(f"Top 10 comps for {selected_name} ({og_pos}){age_str} | Seasons: {yr_label}")

                display_stats = [c for c in POSITION_STATS.get(og_pos, []) if c in ps.columns]

                for rank, (_, comp_row) in enumerate(top10.iterrows(), start=1):
                    comp_name = comp_row["Player"]
                    comp_pos  = comp_row["Position"]
                    comp_team = comp_row["Team"]
                    yr_signed = comp_row["Year Signed"]
                    length    = comp_row["Contract Length (Yrs)"]
                    total_val = comp_row["Total Value ($M)"]
                    aav       = comp_row["AAV ($M)"]
                    gtd       = comp_row["Guaranteed ($M)"]
                    age       = comp_row.get("Age at Signing (Rounded)", np.nan)
                    sim_pct   = float(comp_row["_sim"])

                    age_s = f" | Age at signing: {int(age)}" if pd.notna(age) else ""
                    parts = []
                    if pd.notna(yr_signed): parts.append(f"Signed: {int(yr_signed)}")
                    if pd.notna(length):    parts.append(f"{fmt(length)} yrs")
                    if pd.notna(total_val): parts.append(f"${total_val:.1f}M total")
                    if pd.notna(aav):       parts.append(f"${aav:.2f}M AAV")
                    if pd.notna(gtd):       parts.append(f"${gtd:.2f}M guaranteed")

                    with st.expander(f"#{rank} — {comp_name} | {comp_pos} | {comp_team}{age_s} | Similarity: {sim_pct:.0f}%"):
                        st.caption(" | ".join(parts))

                        combined_comp = {}
                        combined_og   = {}

                        for slot, (og_yr, og_row) in enumerate(og_season_rows, start=1):
                            prefix  = f"Y-{slot}: "
                            comp_yr = (int(yr_signed) - slot) if pd.notna(yr_signed) else None
                            sec_lbl = f"{comp_yr} Stats" if comp_yr else f"Y-{slot} Stats"

                            table_data = []
                            for stat in display_stats:
                                cv = comp_row.get(f"{prefix}{stat}", np.nan)
                                ov = og_row.get(stat, np.nan)
                                if not (pd.notna(cv) and cv != 0 and pd.notna(ov) and ov != 0):
                                    continue
                                cv, ov = float(cv), float(ov)
                                dp = (ov - cv) / abs(cv) * 100 if cv != 0 else 0
                                ds = ("+" if dp >= 0 else "") + f"{dp:.0f}%"
                                table_data.append({
                                    "Stat": stat_label(stat),
                                    comp_name: fmt(cv),
                                    selected_name: fmt(ov),
                                    "Diff": ds,
                                })
                                combined_comp[stat] = combined_comp.get(stat, 0.0) + cv
                                combined_og[stat]   = combined_og.get(stat, 0.0)   + ov

                            if table_data:
                                st.markdown(f"**{sec_lbl}**")
                                st.dataframe(pd.DataFrame(table_data), hide_index=True, use_container_width=True)

                        if len(og_season_rows) > 1 and combined_comp:
                            combined_rows = []
                            for stat in display_stats:
                                if stat not in combined_comp or combined_comp[stat] == 0 or combined_og.get(stat, 0) == 0:
                                    continue
                                cv, ov = combined_comp[stat], combined_og[stat]
                                dp = (ov - cv) / abs(cv) * 100 if cv != 0 else 0
                                ds = ("+" if dp >= 0 else "") + f"{dp:.0f}%"
                                combined_rows.append({
                                    "Stat": stat_label(stat),
                                    comp_name: fmt(cv),
                                    selected_name: fmt(ov),
                                    "Diff": ds,
                                })
                            if combined_rows:
                                st.markdown("**Combined**")
                                st.dataframe(pd.DataFrame(combined_rows), hide_index=True, use_container_width=True)

                # ── Negotiation Summary ────────────────────────────────────────
                summary = calc_neg_summary(og_pos, og_season_rows, top10, len(chosen_years))
                if summary:
                    st.divider()
                    st.subheader(f"Negotiation Summary — {selected_name} | {og_pos}{age_str}")

                    col1, col2, col3 = st.columns(3)
                    col1.metric("Floor (25th pct)", f"${summary['floor']:.1f}M AAV")
                    col2.metric("Target (weighted)", f"${summary['target_aav']:.1f}M AAV")
                    col3.metric("Ceiling (75th pct)", f"${summary['ceiling']:.1f}M AAV")

                    if pd.notna(summary["target_gtd"]):
                        st.metric("Target Guaranteed", f"${summary['target_gtd']:.1f}M")

                    st.markdown("**Performance-Adjusted Comps**")
                    adj_df = pd.DataFrame([
                        {
                            "Player": r[0],
                            "Raw AAV": f"${r[1]:.1f}M",
                            "Adjustment": f"{'+' if r[2] >= 0 else ''}{r[2]:.0f}%",
                            "Adj AAV": f"${r[3]:.1f}M",
                            "Similarity": f"{r[6]:.0f}%",
                        }
                        for r in summary["rows"]
                    ])
                    st.dataframe(adj_df, hide_index=True, use_container_width=True)

                    st.markdown("**Top 3 Strongest Comps**")
                    top3 = sorted(summary["rows"], key=lambda r: r[6], reverse=True)[:3]
                    for comp_name, raw_aav, _, _, _, _, sim in top3:
                        yr = top10.loc[top10["Player"] == comp_name, "Year Signed"]
                        yr_str = str(int(yr.iloc[0])) if not yr.empty and pd.notna(yr.iloc[0]) else "?"
                        st.write(f"**{comp_name}** — {sim:.0f}% similarity | ${raw_aav:.1f}M AAV | {yr_str}")
