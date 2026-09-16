#!/usr/bin/env python3
"""
04_build_cfi_final.py
Canonical Crop Fragility Index computation used by the final project.

Inputs
------
Master_clean.xlsx, sheet "master"

Outputs
-------
outputs/cfi_results.csv
outputs/reinforcement_capacity.csv
outputs/shock_coverage.csv
outputs/sensitivity_analysis.csv

Method
------
Recent baseline: 2021-2023 mean production by country and crop.
Exposure: leading producer / world production.
Downside volatility: RMS of negative annual production growth; positive years set to zero.
Resilience ratio: sum(max(historical_peak - recent_baseline, 0)) outside the leader
                  divided by leading producer recent output.
Structural shortfall S = max(0, 1 - resilience ratio).

E* = min(Exposure / 0.50, 1)
D* = min(DownsideVolatility / 0.20, 1)
H  = 0.50*E* + 0.50*D*
CFI = 100*S*(0.70 + 0.30*H)
"""
from pathlib import Path
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
MASTER = ROOT / "Master_clean.xlsx"
OUT = ROOT / "outputs"
OUT.mkdir(exist_ok=True)

RECENT = (2021, 2023)
HISTORY = (2000, 2023)
EXPOSURE_CAP = 0.50
DOWNSIDE_CAP = 0.20
ALPHA = 0.70
SHOCKS = [0.20, 0.35, 0.50, 1.00]

def downside_volatility(series):
    g = series.sort_index().pct_change().dropna()
    d = np.minimum(g.to_numpy(), 0.0)
    return float(np.sqrt(np.mean(d*d))) if len(d) else 0.0

def risk_band(x):
    if x >= 50: return "High"
    if x >= 20: return "Elevated"
    if x >= 5: return "Moderate"
    if x > 0: return "Low"
    return "Structurally buffered"

m = pd.read_excel(MASTER, sheet_name="master",
                  usecols=["Year","Country","Crop","Production_tonnes"])
m = m[m.Year.between(*HISTORY)].dropna()

rows, reinforcement, shock_rows = [], [], []
for crop in sorted(m.Crop.unique()):
    g = m[m.Crop == crop]
    recent = (g[g.Year.between(*RECENT)]
              .groupby("Country").Production_tonnes.mean().dropna())
    top = recent.idxmax()
    top_out = float(recent[top])
    exposure = top_out / recent.sum()

    peak = g.groupby("Country").Production_tonnes.max()
    cap = (peak - recent).clip(lower=0).drop(labels=[top], errors="ignore")
    cap = cap[cap > 0].sort_values(ascending=False)
    backfill = float(cap.sum())
    ratio = backfill / top_out
    S = max(0.0, 1.0-ratio)

    hist = (g[g.Country == top].groupby("Year").Production_tonnes.sum())
    dvol = downside_volatility(hist)
    E = min(exposure/EXPOSURE_CAP, 1.0)
    D = min(dvol/DOWNSIDE_CAP, 1.0)
    H = 0.5*E + 0.5*D
    cfi = 100*S*(ALPHA + (1-ALPHA)*H)

    rows.append([crop, top, exposure, dvol, top_out, backfill, ratio, cfi, risk_band(cfi)])
    for country, capacity in cap.items():
        reinforcement.append([crop, top, country, capacity])
    for shock in SHOCKS:
        loss = top_out*shock
        covered = min(loss, backfill)
        shock_rows.append([crop, top, shock, loss, covered, max(0, loss-backfill),
                           "covered" if backfill >= loss else "GAP"])

cols = ["Crop","Top producer","Exposure","Downside volatility","Top output tonnes",
        "Backfill capacity tonnes","Resilience ratio","CFI","Risk band"]
pd.DataFrame(rows, columns=cols).sort_values("CFI", ascending=False).to_csv(
    OUT/"cfi_results.csv", index=False)
pd.DataFrame(reinforcement, columns=["Crop","Top producer","Reinforcement country",
                                    "Available capacity tonnes"]).to_csv(
    OUT/"reinforcement_capacity.csv", index=False)
pd.DataFrame(shock_rows, columns=["Crop","Top producer","Shock","Loss tonnes",
                                 "Covered tonnes","Gap tonnes","Status"]).to_csv(
    OUT/"shock_coverage.csv", index=False)

# Sensitivity grid used to test ranking robustness.
sens = []
for alpha in [0.5,0.6,0.7,0.8]:
    for ew in [0.3,0.4,0.5,0.6,0.7]:
        for ecap in [0.4,0.5,0.6]:
            for dcap in [0.15,0.20,0.25]:
                temp=[]
                for r in rows:
                    crop, top, exposure, dvol, top_out, backfill, ratio, *_ = r
                    S=max(0,1-ratio)
                    E=min(exposure/ecap,1)
                    D=min(dvol/dcap,1)
                    H=ew*E+(1-ew)*D
                    score=100*S*(alpha+(1-alpha)*H)
                    temp.append([crop,score])
                temp=sorted(temp,key=lambda x:x[1],reverse=True)
                for rank,(crop,score) in enumerate(temp,1):
                    sens.append([alpha,ew,ecap,dcap,crop,score,rank])
pd.DataFrame(sens, columns=["alpha","exposure_weight","exposure_cap","downside_cap",
                            "Crop","CFI","Rank"]).to_csv(
    OUT/"sensitivity_analysis.csv", index=False)

print("Wrote final CFI outputs to", OUT)
