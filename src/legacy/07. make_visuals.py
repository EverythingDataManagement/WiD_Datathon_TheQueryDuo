#!/usr/bin/env python3
# =====================================================================
#  make_visuals.py  —  Query Queens, GROW track
#  Generates the core charts for the video/deck from the master
#  (+ the land-use file for the forest scatter). Saves PNGs to ./outputs.
#  Colab:  !pip install pandas numpy matplotlib scipy --quiet
# =====================================================================
import os
import pandas as pd, numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

# ---------- SETTINGS ----------
MASTER_FILE  = "Master_clean.xlsx"     # sheet 'master'  (or your Book*.xlsx)
MASTER_SHEET = "master"
LAND_FILE    = "Inputs_LandUse_E_All_Data (Reviewed).xlsx"   # optional, for the forest scatter
LAND_SHEET   = "Sheet 1 - Inputs_LandUse_E_All_"
OUTDIR       = "outputs"
START, END   = 2000, 2023
os.makedirs(OUTDIR, exist_ok=True)

# brand palette
GREEN="#375623"; MID="#548235"; RED="#C0392B"; AMBER="#E1A100"; GREY="#595959"; LIGHT="#E2EFDA"
plt.rcParams.update({"font.family":"DejaVu Sans","axes.spines.top":False,"axes.spines.right":False,
                     "font.size":11,"figure.dpi":150})

m = pd.read_excel(MASTER_FILE, sheet_name=MASTER_SHEET)
m = m[m["Year"].between(START, END)]
crops = sorted(m["Crop"].unique())

# ------------------------------------------------------------------
# helper: resilience ratio per crop
# ------------------------------------------------------------------
def resilience_table():
    rows=[]
    for cr in crops:
        g=m[m.Crop==cr]; cur=g[g.Year==END].groupby("Country").Production_tonnes.sum()
        top=cur.idxmax(); lost=cur.max(); peak=g.groupby("Country").Production_tonnes.max()
        others=[c for c in cur.index if c!=top]
        bf=(peak.reindex(others).fillna(0)-cur.reindex(others).fillna(0)).clip(lower=0).sum()
        rows.append((cr, top, lost, bf, bf/lost))
    return pd.DataFrame(rows, columns=["Crop","Top","Lost","Backfill","Ratio"])

# ==================================================================
# 1. RESILIENCE BAR  (the headline chart)
# ==================================================================
def viz_resilience():
    d=resilience_table().sort_values("Ratio")
    colors=[RED if r<1 else GREEN for r in d.Ratio]
    fig,ax=plt.subplots(figsize=(9,5.5))
    ax.barh(d.Crop, d.Ratio, color=colors)
    ax.axvline(1.0, color="black", lw=1.5, ls="--")
    ax.text(1.02, -0.6, "coverable  →", fontsize=9, color=GREY)
    ax.text(0.98, -0.6, "←  cannot backfill", ha="right", fontsize=9, color=RED)
    for y,(r,cr) in enumerate(zip(d.Ratio,d.Crop)):
        ax.text(r+0.05, y, f"{r:.2f}", va="center", fontsize=9,
                color=RED if r<1 else GREEN, fontweight="bold")
    ax.set_xlabel("Resilience ratio  (backfill capacity ÷ tonnage lost if #1 producer fails)")
    ax.set_title("Can the world replace a crop's top producer?\n7 of 10 staples: NO — and the crops we eat most are the least replaceable",
                 fontweight="bold", loc="left")
    ax.set_xlim(0, max(4, d.Ratio.max()*1.1))
    fig.tight_layout(); fig.savefig(f"{OUTDIR}/1_resilience_bar.png", bbox_inches="tight"); plt.close()
    print("saved 1_resilience_bar.png")

# ==================================================================
# 2. AREA vs YIELD  (how each leader grew)
# ==================================================================
def viz_area_yield():
    leaders=["Brazil","China, mainland","India","United States of America","Russian Federation",
             "Argentina","Ukraine","Nigeria"]
    rows=[]
    for c in leaders:
        g=m[(m.Country==c)&(m.Crop.isin(["Maize","Rice","Wheat","Soybeans","Potatoes","Sorghum"]))]
        a0=g[g.Year==START].Area_harvested_ha.sum(); a1=g[g.Year==END].Area_harvested_ha.sum()
        p0=g[g.Year==START].Production_tonnes.sum(); p1=g[g.Year==END].Production_tonnes.sum()
        if a0<=0 or p0<=0: continue
        y0,y1=p0/a0,p1/a1; dP=p1-p0
        area=(a1-a0)*y0; yld=(y1-y0)*a0
        tot=abs(area)+abs(yld)
        rows.append((c.replace(", mainland","").replace("United States of America","USA"),
                     area/tot*100, yld/tot*100))
    d=pd.DataFrame(rows,columns=["Country","Area%","Yield%"]).sort_values("Area%")
    fig,ax=plt.subplots(figsize=(9,5.5))
    ax.barh(d.Country, d["Area%"], color=AMBER, label="Area expansion (land)")
    ax.barh(d.Country, d["Yield%"], left=d["Area%"], color=MID, label="Yield improvement")
    ax.set_xlabel("share of production growth 2000–2023 (%)")
    ax.set_title("Two ways to grow more food\nLand-expanders (Brazil, Argentina) vs yield-intensifiers (US, Russia)",
                 fontweight="bold", loc="left")
    ax.legend(loc="lower right", frameon=False); ax.set_xlim(0,100)
    fig.tight_layout(); fig.savefig(f"{OUTDIR}/3_area_vs_yield.png", bbox_inches="tight"); plt.close()
    print("saved 3_area_vs_yield.png")

# ==================================================================
# 3. FRAGILITY INDEX ranking
# ==================================================================
def viz_fragility():
    def norm(x,lo,hi): return np.clip((x-lo)/(hi-lo+1e-9),0,1)
    d=resilience_table()
    exp=[]; haz=[]
    for _,x in d.iterrows():
        g=m[m.Crop==x.Crop]; cur=g[g.Year==END].groupby("Country").Production_tonnes.sum()
        exp.append(cur.max()/cur.sum()*100)
        s=m[(m.Country==x.Top)&(m.Crop==x.Crop)].sort_values("Year").Production_tonnes
        haz.append((s.pct_change()*100).std())
    d["exp"]=exp; d["haz"]=haz
    en=norm(d.exp,d.exp.min(),d.exp.max()); hn=norm(d.haz,d.haz.min(),d.haz.max())
    bn=1-d.Ratio.clip(lower=0,upper=1)
    d["score"]=(100*bn*(0.7+0.3*(0.6*en+0.4*hn))).round(0)
    d=d.sort_values("score")
    def band(s): return RED if s>55 else (AMBER if s>20 else GREEN)
    fig,ax=plt.subplots(figsize=(9,5.5))
    ax.barh(d.Crop, d.score, color=[band(s) for s in d.score])
    for y,s in enumerate(d.score): ax.text(s+1,y,f"{int(s)}",va="center",fontsize=9,fontweight="bold")
    ax.set_xlabel("Staple Fragility Index score (0–100)")
    ax.set_title("Which staples are most at risk?\nHigh: Sugarcane, Maize, Rice, Soybeans  (red = high, amber = moderate, green = low)",
                 fontweight="bold", loc="left")
    ax.set_xlim(0,100)
    fig.tight_layout(); fig.savefig(f"{OUTDIR}/4_fragility_index.png", bbox_inches="tight"); plt.close()
    print("saved 4_fragility_index.png")

# ==================================================================
# 4. CROPLAND vs FOREST scatter (needs the land-use file)
# ==================================================================
def viz_forest():
    if not os.path.exists(LAND_FILE):
        print("skip forest scatter — land-use file not found"); return
    rl=pd.read_excel(LAND_FILE, sheet_name=LAND_SHEET, header=1)
    rl=rl[(rl["Element Code"]==5110)&(rl["Item Code"].isin([6620,6646]))]
    rl.columns=[str(c) for c in rl.columns]; ycols=[c for c in rl.columns if c.startswith("Y") and c[1:].isdigit()]
    long=rl.melt(id_vars=["Area","Item Code"],value_vars=ycols,var_name="Yv",value_name="v")
    long["Year"]=long["Yv"].str[1:].astype(int)
    cov=long.dropna(subset=["v"]).groupby("Year").Area.nunique()
    # latest year with near-full coverage (avoid half-empty projection years like 2025)
    L_END=int(cov[cov>=cov.max()*0.98].index.max())
    def sl(item,yr): return long[(long["Item Code"]==item)&(long.Year==yr)].set_index("Area").v
    allc=pd.DataFrame({"c0":sl(6620,START),"c1":sl(6620,L_END),"f0":sl(6646,START),"f1":sl(6646,L_END)}).dropna()
    allc["dc"]=allc.c1-allc.c0; allc["df"]=allc.f1-allc.f0
    # keep ONLY real countries — use the master's cleaned country list as the whitelist.
    # This avoids trying to blocklist every aggregate spelling (World, South America, LLDCs, ...).
    real_countries=set(m["Country"].unique())
    allc=allc[allc.index.isin(real_countries)]
    from scipy import stats
    r,p=stats.pearsonr(allc.dc,allc.df)
    fig,ax=plt.subplots(figsize=(8.5,6))
    ax.scatter(allc.dc/1e3, allc.df/1e3, s=18, color=GREY, alpha=0.5)
    for c in ["Brazil","Nigeria","Argentina","China, mainland","United States of America","Indonesia"]:
        if c in allc.index:
            ax.scatter(allc.loc[c,"dc"]/1e3, allc.loc[c,"df"]/1e3, s=70, color=RED, zorder=5)
            ax.annotate(c.replace(", mainland","").replace("United States of America","USA"),
                        (allc.loc[c,"dc"]/1e3, allc.loc[c,"df"]/1e3), fontsize=9, fontweight="bold",
                        xytext=(5,5), textcoords="offset points")
    ax.axhline(0,color="black",lw=0.6); ax.axvline(0,color="black",lw=0.6)
    ax.set_xlabel("Cropland change 2000–%d (million ha)"%L_END)
    ax.set_ylabel("Forest change (million ha)")
    ax.set_title(f"Where cropland expanded, forest was lost\nAll countries: r = {r:.2f}  (Brazil is the world's #1 forest loser)",
                 fontweight="bold", loc="left")
    fig.tight_layout(); fig.savefig(f"{OUTDIR}/2_cropland_forest_scatter.png", bbox_inches="tight"); plt.close()
    print("saved 2_cropland_forest_scatter.png")

# ==================================================================
# 5. PER-COUNTRY TREND lines
# ==================================================================
def viz_trend():
    sel=["Argentina","Brazil","China, mainland","India","Nigeria","Pakistan","Russian Federation","Thailand","Ukraine","United States of America"]
    d=m[(m.Crop.isin(["Maize","Rice","Wheat","Soybeans","Potatoes","Sorghum"]))&(m.Country.isin(sel))]
    piv=(d.pivot_table(index="Year",columns="Country",values="Production_tonnes",aggfunc="sum")/1e6)[sel]
    fig,ax=plt.subplots(figsize=(9.5,5.5))
    for c in sel:
        ax.plot(piv.index, piv[c], lw=1.8)
        ax.annotate(c.replace(", mainland","").replace("United States of America","USA"),
                    (piv.index[-1], piv[c].iloc[-1]), fontsize=8, xytext=(4,0), textcoords="offset points", va="center")
    ax.set_xlabel("Year"); ax.set_ylabel("Output of 6 staples (Mt)")
    ax.set_title("How each producer evolved, 2000–2023\nSteady giants (China, India) vs volatile risers (Brazil, Russia, Ukraine, Argentina)",
                 fontweight="bold", loc="left")
    ax.set_xlim(START, END+2)
    fig.tight_layout(); fig.savefig(f"{OUTDIR}/5_country_trends.png", bbox_inches="tight"); plt.close()
    print("saved 5_country_trends.png")

if __name__=="__main__":
    viz_resilience()
    viz_area_yield()
    viz_fragility()
    viz_forest()
    viz_trend()
    print("\nAll charts saved to ./outputs/")
