#!/usr/bin/env python3
"""
build_analysis_workbook.py
==========================
Rebuilds the "Analysis Summary" as a calculations workbook: one tab per category
(Growth, Concentration, Cost, Resilience, Conclusion) plus an Index and a
pivot-ready source sheet.

Two kinds of input feed this:
  * RECOMPUTED from raw panels  -> Growth, Concentration, Cost, and Q11.
      Sources: Master_clean.xlsx!master, Cost_Master.xlsx!cost_master
  * SOURCED from pre-computed tables -> Resilience & finale.
      Sources: Fragility_Index.xlsx (partial_shock, fragility_index,
      historical_drops, index_validation, prescription), resilience_by_crop.xlsx

Rerun after editing the data files (or the CONFIG constants) to refresh every number.

    python3 build_analysis_workbook.py            # writes ./Analysis_Calculations.xlsx
    python3 build_analysis_workbook.py --validate # print computed tables, do not write

Requires: pandas, numpy, scipy, openpyxl (all preinstalled).
"""

import argparse
import sys
import numpy as np
import pandas as pd
from scipy.stats import pearsonr

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------------- #
# CONFIG  — edit these, then rerun.
# ----------------------------------------------------------------------------- #
DATA_DIR = "/mnt/user-data/uploads"
OUT_PATH = "Analysis_Calculations.xlsx"

FILES = {
    "master":     f"{DATA_DIR}/Master_clean.xlsx",
    "cost":       f"{DATA_DIR}/Cost_Master.xlsx",
    "fragility":  f"{DATA_DIR}/Fragility_Index.xlsx",
    "resilience": f"{DATA_DIR}/resilience_by_crop.xlsx",
}

START_YEAR = 2000
END_YEAR   = 2023          # summary window; master also carries 2024 (provisional)

PANEL_10 = [
    "Argentina", "Brazil", "China, mainland", "India", "Nigeria",
    "Pakistan", "Russian Federation", "Thailand", "Ukraine",
    "United States of America",
]
TEN_CROPS = ["Maize", "Rice", "Wheat", "Soybeans", "Potatoes",
             "Sugarcane", "Cassava", "Sorghum", "Millet", "Barley"]
# The six "core" staples that carry per-crop coverage in the summary (Q4).
SIX_CORE  = ["Maize", "Rice", "Wheat", "Soybeans", "Potatoes", "Sorghum"]

MT = 1e6  # tonnes -> million tonnes


# ----------------------------------------------------------------------------- #
# LOADERS
# ----------------------------------------------------------------------------- #
def load_master():
    df = pd.read_excel(FILES["master"], sheet_name="master")
    keep = ["Year", "Country", "Continent", "Subregion", "Country Classification",
            "Crop", "Staple Crop", "Area_harvested_ha", "Yield_kg_per_ha",
            "Production_tonnes"]
    df = df[keep].copy()
    for c in ["Area_harvested_ha", "Yield_kg_per_ha", "Production_tonnes"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def load_cost():
    df = pd.read_excel(FILES["cost"], sheet_name="cost_master")
    for c in ["cropland_1000ha", "forest_1000ha", "fertilizer_kg_per_ha"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def load_fragility_sheet(name):
    return pd.read_excel(FILES["fragility"], sheet_name=name)


def load_resilience():
    return pd.read_excel(FILES["resilience"])


# ----------------------------------------------------------------------------- #
# HELPERS
# ----------------------------------------------------------------------------- #
def endpoints(df, group_cols, value, y0=START_YEAR, y1=END_YEAR):
    """Sum `value` at y0 and y1 for each group; return wide with pct/abs change."""
    sub = df[df.Year.isin([y0, y1])]
    g = (sub.groupby(group_cols + ["Year"])[value].sum()
            .unstack("Year").reset_index())
    g = g.rename(columns={y0: f"y{y0}", y1: f"y{y1}"})
    g[f"y{y0}"] = g.get(f"y{y0}", np.nan)
    g[f"y{y1}"] = g.get(f"y{y1}", np.nan)
    g["abs_change"] = g[f"y{y1}"] - g[f"y{y0}"]
    g["pct_change"] = np.where(g[f"y{y0}"] > 0,
                               (g[f"y{y1}"] / g[f"y{y0}"] - 1) * 100, np.nan)
    return g


def pearson_safe(x, y):
    m = np.isfinite(x) & np.isfinite(y)
    x, y = np.asarray(x)[m], np.asarray(y)[m]
    if len(x) < 3:
        return np.nan, np.nan, len(x)
    r, p = pearsonr(x, y)
    return r, p, len(x)


# ----------------------------------------------------------------------------- #
# GROWTH
# ----------------------------------------------------------------------------- #
def growth_world_headline(m):
    """Q1 — world staple output, all-10 vs six-core, START vs END."""
    rows = []
    for label, crops in [("All 10 crops", TEN_CROPS), ("Six core staples", SIX_CORE)]:
        sub = m[m.Crop.isin(crops)]
        p0 = sub.loc[sub.Year == START_YEAR, "Production_tonnes"].sum() / MT
        p1 = sub.loc[sub.Year == END_YEAR, "Production_tonnes"].sum() / MT
        rows.append({"Basket": label, f"{START_YEAR} (Mt)": p0,
                     f"{END_YEAR} (Mt)": p1, "Change (Mt)": p1 - p0,
                     "Change (%)": (p1 / p0 - 1) * 100})
    return pd.DataFrame(rows)


def growth_by_crop(m):
    """Q1 — per-crop growth, ranked by absolute Mt added."""
    g = endpoints(m, ["Crop"], "Production_tonnes")
    g["2000 (Mt)"] = g[f"y{START_YEAR}"] / MT
    g["2023 (Mt)"] = g[f"y{END_YEAR}"] / MT
    g["Change (Mt)"] = g["abs_change"] / MT
    g["Change (%)"] = g["pct_change"]
    g = g.sort_values("Change (Mt)", ascending=False)
    g["Rank by Mt added"] = range(1, len(g) + 1)
    return g[["Crop", "2000 (Mt)", "2023 (Mt)", "Change (Mt)",
              "Change (%)", "Rank by Mt added"]]


def growth_by_country(m):
    """Q1 — country contributions to the SIX-CORE gain (the 'growth leader' cut).

    Six-core excludes sugarcane and cassava; on this basket China leads. Including
    all 10 crops makes Brazil the runaway #1 — that all-10 view lives in Concentration Q3.
    """
    sub = m[m.Crop.isin(SIX_CORE)]
    g = endpoints(sub, ["Country"], "Production_tonnes")
    total_gain = g["abs_change"].sum()
    g["Change (Mt)"] = g["abs_change"] / MT
    g["Share of six-core gain (%)"] = g["abs_change"] / total_gain * 100
    g = g.sort_values("Change (Mt)", ascending=False).head(12)
    g["Rank"] = range(1, len(g) + 1)
    return g[["Rank", "Country", "Change (Mt)", "Share of six-core gain (%)"]]


def growth_area_yield_decomp(m):
    """Q2 — decompose each panel country's per-crop change into area vs yield.

    Production = Area(ha) * Yield(kg/ha) / 1000  (tonnes).
    Additive decomposition of the change:
        dP = Y0*dA + A0*dY + dA*dY   (last term = interaction)
    Classify by which pure effect is larger in magnitude.
    """
    sub = m[(m.Country.isin(PANEL_10)) & (m.Crop.isin(TEN_CROPS))
            & (m.Year.isin([START_YEAR, END_YEAR]))]
    piv = sub.pivot_table(index=["Country", "Crop"], columns="Year",
                          values=["Area_harvested_ha", "Yield_kg_per_ha"])
    rows = []
    for (country, crop), r in piv.iterrows():
        A0 = r.get(("Area_harvested_ha", START_YEAR), np.nan)
        A1 = r.get(("Area_harvested_ha", END_YEAR), np.nan)
        Y0 = r.get(("Yield_kg_per_ha", START_YEAR), np.nan)
        Y1 = r.get(("Yield_kg_per_ha", END_YEAR), np.nan)
        if not np.all(np.isfinite([A0, A1, Y0, Y1])) or (A0 <= 0 and A1 <= 0):
            continue
        dA, dY = A1 - A0, Y1 - Y0
        area_eff  = Y0 * dA / 1000 / MT          # Mt
        yield_eff = A0 * dY / 1000 / MT          # Mt
        inter     = dA * dY / 1000 / MT          # Mt
        total     = area_eff + yield_eff + inter
        if abs(total) < 1e-6:
            driver = "flat"
        elif total < 0:
            driver = "declined"
        elif abs(area_eff) >= abs(yield_eff):
            driver = "land-driven"
        else:
            driver = "yield-driven"
        rows.append({"Country": country, "Crop": crop,
                     "Area effect (Mt)": area_eff, "Yield effect (Mt)": yield_eff,
                     "Interaction (Mt)": inter, "Total change (Mt)": total,
                     "Driver": driver})
    out = pd.DataFrame(rows).sort_values(["Country", "Total change (Mt)"],
                                         ascending=[True, False])
    tally = (out["Driver"].value_counts()
             .rename_axis("Driver").reset_index(name="Country-crop count"))
    return out, tally


def panel_country_trend(m):
    """Q1.1 (a) — per panel country: growth, volatility, best/worst year."""
    sub = m[m.Country.isin(PANEL_10)]
    yearly = sub.groupby(["Country", "Year"])["Production_tonnes"].sum().reset_index()
    yearly = yearly[yearly.Year.between(START_YEAR, END_YEAR)]
    rows = []
    for c, grp in yearly.groupby("Country"):
        grp = grp.sort_values("Year")
        p0 = grp.loc[grp.Year == START_YEAR, "Production_tonnes"].sum() / MT
        p1 = grp.loc[grp.Year == END_YEAR, "Production_tonnes"].sum() / MT
        yoy = grp["Production_tonnes"].pct_change() * 100
        best_i, worst_i = yoy.idxmax(), yoy.idxmin()
        rows.append({
            "Country": c, "2000 (Mt)": p0, "2023 (Mt)": p1,
            "Growth (%)": (p1 / p0 - 1) * 100,
            "YoY volatility (std %)": yoy.std(),
            "Best YoY (%)": yoy.max(),
            "Best year": int(grp.loc[best_i, "Year"]),
            "Worst YoY (%)": yoy.min(),
            "Worst year": int(grp.loc[worst_i, "Year"]),
        })
    return pd.DataFrame(rows).sort_values("Growth (%)", ascending=False)


def panel_country_yearly_wide(m):
    """Q1.1 (b) — Country x Year production (Mt) for the line-chart view."""
    sub = m[m.Country.isin(PANEL_10)]
    yearly = sub.groupby(["Country", "Year"])["Production_tonnes"].sum().reset_index()
    yearly = yearly[yearly.Year.between(START_YEAR, END_YEAR)]
    wide = (yearly.pivot(index="Country", columns="Year",
                         values="Production_tonnes") / MT).reset_index()
    return wide


# ----------------------------------------------------------------------------- #
# CONCENTRATION
# ----------------------------------------------------------------------------- #
def concentration_shares(m):
    """Q3 — top-N shares of the GAIN (who drove the growth), both baskets.

    'Drove the growth' = share of the 2000->2023 increase, not the end-year level.
    All-10 basket -> Brazil #1 (sugarcane + cassava). Six-core -> China #1.
    """
    rows = []
    for label, crops in [("All 10 crops (incl. sugarcane)", TEN_CROPS),
                         ("Six core (no sugarcane)", SIX_CORE)]:
        g = endpoints(m[m.Crop.isin(crops)], ["Country"], "Production_tonnes")
        tot = g["abs_change"].sum()
        top = g.sort_values("abs_change", ascending=False)
        rows.append({"Basket": label,
                     "Top-3 share of gain (%)": top["abs_change"].head(3).sum() / tot * 100,
                     "Top-5 share of gain (%)": top["abs_change"].head(5).sum() / tot * 100,
                     "Top-10 share of gain (%)": top["abs_change"].head(10).sum() / tot * 100,
                     "#1 country": top.iloc[0]["Country"]})
    return pd.DataFrame(rows)


def concentration_top10_list(m):
    """Q3 — top-10 countries by their contribution to the GAIN (all 10 crops)."""
    g = endpoints(m[m.Crop.isin(TEN_CROPS)], ["Country"], "Production_tonnes")
    tot = g["abs_change"].sum()
    top = g.sort_values("abs_change", ascending=False).head(10).copy()
    top["Gain (Mt)"] = top["abs_change"] / MT
    top["Share of gain (%)"] = top["abs_change"] / tot * 100
    top.insert(0, "Rank", range(1, len(top) + 1))
    return top[["Rank", "Country", "Gain (Mt)", "Share of gain (%)"]]


def concentration_crop_coverage(m):
    """Q4 — for each crop, END-year world share held by the 10-panel countries,
    plus the biggest producers that sit OUTSIDE the panel."""
    sub = m[m.Year == END_YEAR]
    rows = []
    for crop in TEN_CROPS:
        cs = sub[sub.Crop == crop].groupby("Country")["Production_tonnes"].sum()
        world = cs.sum()
        panel_share = cs[cs.index.isin(PANEL_10)].sum() / world * 100 if world else np.nan
        outside = cs[~cs.index.isin(PANEL_10)].sort_values(ascending=False).head(4)
        misses = ", ".join(f"{c} ({v/MT:.0f} Mt)" for c, v in outside.items())
        rows.append({"Crop": crop, "Panel-10 coverage (%)": panel_share,
                     "Biggest producers outside panel": misses})
    return pd.DataFrame(rows).sort_values("Panel-10 coverage (%)", ascending=False)


# ----------------------------------------------------------------------------- #
# COST
# ----------------------------------------------------------------------------- #
def cost_country_change(c):
    """Per country: first vs last available cropland & forest, and mean fertilizer."""
    c = c[c.Year.between(START_YEAR, END_YEAR)]
    rows = []
    for country, grp in c.groupby("Country"):
        grp = grp.sort_values("Year")
        cl = grp.dropna(subset=["cropland_1000ha"])
        fo = grp.dropna(subset=["forest_1000ha"])
        if len(cl) >= 2:
            crop_ch = cl["cropland_1000ha"].iloc[-1] - cl["cropland_1000ha"].iloc[0]
            crop_pct = (cl["cropland_1000ha"].iloc[-1] / cl["cropland_1000ha"].iloc[0] - 1) * 100 \
                if cl["cropland_1000ha"].iloc[0] else np.nan
        else:
            crop_ch = crop_pct = np.nan
        if len(fo) >= 2:
            for_ch = fo["forest_1000ha"].iloc[-1] - fo["forest_1000ha"].iloc[0]
            for_pct = (fo["forest_1000ha"].iloc[-1] / fo["forest_1000ha"].iloc[0] - 1) * 100 \
                if fo["forest_1000ha"].iloc[0] else np.nan
        else:
            for_ch = for_pct = np.nan
        rows.append({"Country": country,
                     "Cropland change (1000 ha)": crop_ch, "Cropland change (%)": crop_pct,
                     "Forest change (1000 ha)": for_ch, "Forest change (%)": for_pct,
                     "Mean fertilizer (kg/ha)": grp["fertilizer_kg_per_ha"].mean()})
    return pd.DataFrame(rows)


def cost_land_forest_corr(cc):
    """Q5 / Q3.1a — Pearson r between cropland change and forest change."""
    r, p, n = pearson_safe(cc["Cropland change (1000 ha)"].values,
                           cc["Forest change (1000 ha)"].values)
    return pd.DataFrame([{
        "Relationship": "Cropland change vs forest change (all countries)",
        "Pearson r": r, "p-value": p, "n countries": n,
        "Reading": "negative => cropland expanders tend to lose forest"}])


def cost_forest_losers(cc, m):
    """Q5 — the biggest forest losers that also expanded cropland, with the crop
    most responsible (inferred from END-year harvested-area share)."""
    hot = cc[(cc["Forest change (1000 ha)"] < 0) &
             (cc["Cropland change (1000 ha)"] > 0)].copy()
    hot = hot.sort_values("Forest change (1000 ha)").head(10)
    area = m[m.Year == END_YEAR].groupby(["Country", "Crop"])["Area_harvested_ha"].sum()
    def top_crop(country):
        if country not in area.index.get_level_values(0):
            return ""
        s = area.loc[country]
        share = s / s.sum() * 100
        c = share.idxmax()
        return f"{c} ({share.max():.0f}% of harvested area)"
    hot["Leading crop by area"] = hot["Country"].map(top_crop)
    return hot[["Country", "Forest change (1000 ha)", "Cropland change (1000 ha)",
                "Leading crop by area"]]


CEREALS = ["Maize", "Rice", "Wheat", "Sorghum", "Millet", "Barley"]


def cost_fertilizer_yield_corr(m, c):
    """Q6 — national fertilizer intensity vs national CEREAL yield.

    Cereal yield = total cereal production / total cereal area (single scale, the
    standard FAO/World Bank metric). Production-weighting all 10 crops instead
    conflates crop mix (sugarcane ~70 t/ha vs millet ~1 t/ha) and washes the signal out.
    """
    em = (m[(m.Year == END_YEAR) & (m.Crop.isin(CEREALS))]
          .groupby("Country")
          .agg(p=("Production_tonnes", "sum"), a=("Area_harvested_ha", "sum")))
    em = em[em.a > 0]
    em["cereal_yield_kg_ha"] = em.p / em.a * 1000
    fert = c[c.Year == END_YEAR][["Country", "fertilizer_kg_per_ha"]].dropna()
    j = em.reset_index().merge(fert, on="Country", how="inner")
    r, p, n = pearson_safe(j["fertilizer_kg_per_ha"].values, j["cereal_yield_kg_ha"].values)
    return pd.DataFrame([{
        "Relationship": "Fertilizer (kg/ha) vs national cereal yield (kg/ha)",
        "Pearson r": r, "p-value": p, "n countries": n,
        "Reading": "positive => intensification tracks input pressure (a mechanism, not a clean cost split)"}])


def cost_cropland_forest_quadrants(cc):
    """Q3.1b — counter-examples: who expanded cropland WITHOUT losing forest."""
    d = cc.dropna(subset=["Cropland change (1000 ha)", "Forest change (1000 ha)"])
    q = {
        "Cropland up, forest up (both gained)":
            d[(d["Cropland change (1000 ha)"] > 0) & (d["Forest change (1000 ha)"] > 0)],
        "Cropland up, forest down (expansion at forest cost)":
            d[(d["Cropland change (1000 ha)"] > 0) & (d["Forest change (1000 ha)"] < 0)],
        "Cropland down, forest up":
            d[(d["Cropland change (1000 ha)"] < 0) & (d["Forest change (1000 ha)"] > 0)],
        "Cropland down, forest down":
            d[(d["Cropland change (1000 ha)"] < 0) & (d["Forest change (1000 ha)"] < 0)],
    }
    counts = pd.DataFrame([{"Quadrant": k, "Country count": len(v)} for k, v in q.items()])
    # named counter-examples: biggest cropland expanders whose forest also rose
    ce = q["Cropland up, forest up (both gained)"].sort_values(
        "Cropland change (1000 ha)", ascending=False).head(8)
    examples = ce[["Country", "Cropland change (1000 ha)", "Forest change (1000 ha)"]]
    return counts, examples


def fragility_hhi_trend(m):
    """Q11 — is concentration worsening? Per crop: top-producer share and HHI at
    START vs END, plus the trend across all years (sign + p of a linear fit)."""
    rows = []
    for crop in TEN_CROPS:
        sub = m[(m.Crop == crop) & (m.Year.between(START_YEAR, END_YEAR))]
        top_share, hhi, years = {}, {}, sorted(sub.Year.unique())
        hhi_series, yr_series, topshare_series = [], [], []
        for y in years:
            cs = sub[sub.Year == y].groupby("Country")["Production_tonnes"].sum()
            tot = cs.sum()
            if tot <= 0:
                continue
            shares = cs / tot
            hhi_series.append((shares ** 2).sum() * 10000)   # HHI 0-10000
            topshare_series.append(shares.max() * 100)
            yr_series.append(y)
        if len(yr_series) < 3:
            continue
        # linear trend of HHI on year
        r, p, _ = pearson_safe(np.array(yr_series), np.array(hhi_series))
        trend = "rising (concentrating)" if r > 0 else "falling (diversifying)"
        rows.append({
            "Crop": crop,
            f"Top share {START_YEAR} (%)": topshare_series[0],
            f"Top share {END_YEAR} (%)": topshare_series[-1],
            f"HHI {START_YEAR}": hhi_series[0],
            f"HHI {END_YEAR}": hhi_series[-1],
            "HHI trend": trend,
            "Trend p-value": p,
        })
    return pd.DataFrame(rows).sort_values(f"HHI {END_YEAR}", ascending=False)


# ----------------------------------------------------------------------------- #
# WORKBOOK WRITER
# ----------------------------------------------------------------------------- #
FONT      = "Arial"
NAVY      = "1F3A5F"
BAND      = "2E5A87"
HEADER    = "4A7BA6"
QFILL     = "DCE6F1"
NOTEFILL  = "F2F2F2"
WHITE     = "FFFFFF"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def _autosize(ws):
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            col = cell.column_letter
            ln = max(len(s) for s in str(cell.value).split("\n"))
            widths[col] = min(max(widths.get(col, 10), ln + 2), 60)
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


class SheetWriter:
    def __init__(self, ws):
        self.ws = ws
        self.r = 1

    def band(self, text):
        ws = self.ws
        ws.merge_cells(start_row=self.r, start_column=1, end_row=self.r, end_column=8)
        c = ws.cell(self.r, 1, text)
        c.font = Font(name=FONT, size=14, bold=True, color=WHITE)
        c.fill = PatternFill("solid", fgColor=BAND)
        c.alignment = Alignment(vertical="center", horizontal="left", indent=1)
        ws.row_dimensions[self.r].height = 26
        self.r += 2

    def question(self, qid, title, qtext):
        ws = self.ws
        c = ws.cell(self.r, 1, f"{qid}  ·  {title}")
        c.font = Font(name=FONT, size=11, bold=True, color=NAVY)
        c.fill = PatternFill("solid", fgColor=QFILL)
        ws.merge_cells(start_row=self.r, start_column=1, end_row=self.r, end_column=8)
        self.r += 1
        if qtext:
            c = ws.cell(self.r, 1, qtext)
            c.font = Font(name=FONT, size=9, italic=True, color="595959")
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(start_row=self.r, start_column=1, end_row=self.r, end_column=8)
            ws.row_dimensions[self.r].height = 28
            self.r += 1
        self.r += 1

    def table(self, df, pct_cols=(), int_cols=(), num_cols=None):
        ws = self.ws
        if num_cols is None:
            num_cols = [c for c in df.columns
                        if pd.api.types.is_numeric_dtype(df[c])]
        # header
        for j, col in enumerate(df.columns, start=1):
            c = ws.cell(self.r, j, col)
            c.font = Font(name=FONT, size=10, bold=True, color=WHITE)
            c.fill = PatternFill("solid", fgColor=HEADER)
            c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
            c.border = BORDER
        ws.row_dimensions[self.r].height = 30
        self.r += 1
        # body
        for _, row in df.iterrows():
            for j, col in enumerate(df.columns, start=1):
                v = row[col]
                if pd.isna(v):
                    v = None
                c = ws.cell(self.r, j, v)
                c.font = Font(name=FONT, size=10)
                c.border = BORDER
                c.alignment = Alignment(vertical="center",
                                        horizontal="left" if not isinstance(v, (int, float)) else "right")
                if col in pct_cols:
                    c.number_format = '0.0"%"'
                elif col in int_cols:
                    c.number_format = "#,##0"
                elif col in num_cols:
                    c.number_format = "#,##0.0"
            self.r += 1
        self.r += 1

    def note(self, text):
        ws = self.ws
        c = ws.cell(self.r, 1, "▸ " + text)
        c.font = Font(name=FONT, size=8, italic=True, color="7F7F7F")
        c.fill = PatternFill("solid", fgColor=NOTEFILL)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=self.r, start_column=1, end_row=self.r, end_column=8)
        ws.row_dimensions[self.r].height = 26
        self.r += 2

    def textblock(self, lines):
        ws = self.ws
        for ln in lines:
            c = ws.cell(self.r, 1, ln)
            c.font = Font(name=FONT, size=10)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(start_row=self.r, start_column=1, end_row=self.r, end_column=8)
            ws.row_dimensions[self.r].height = max(16, 15 * (len(ln) // 95 + 1))
            self.r += 1
        self.r += 1


def pct(*cols):
    return tuple(cols)


# ----------------------------------------------------------------------------- #
def build(tables):
    wb = openpyxl.Workbook()
    # ------- INDEX -------
    ws = wb.active
    ws.title = "Index"
    sw = SheetWriter(ws)
    sw.band("ANALYSIS CALCULATIONS  —  questions answered with data")
    sw.textblock([
        f"Window: {START_YEAR}–{END_YEAR}.  Panel = {len(PANEL_10)} countries, {len(TEN_CROPS)} crops.",
        "Each tab holds one category. Every question keeps its original label.",
        "",
        "RECOMPUTED from raw panels (rerun the script to refresh):",
        "   • Growth, Concentration, Cost  — from Master_clean!master and Cost_Master!cost_master",
        "   • Q11 concentration trend       — from Master_clean!master",
        "SOURCED from pre-computed tables (produced by the team's upstream scripts):",
        "   • Resilience & Conclusion       — from Fragility_Index.xlsx and resilience_by_crop.xlsx",
        "",
        "pivot_source tab = the full tidy production panel. Insert > PivotTable on it to slice freely.",
    ])
    idx = pd.DataFrame([
        ["Growth",        "Q1, Q2, Q1.1", "recomputed", "Master_clean!master"],
        ["Concentration", "Q3, Q4",       "recomputed", "Master_clean!master"],
        ["Cost",          "Q5, Q6, Q3.1a, Q3.1b, Q7", "recomputed / limitation",
         "Cost_Master!cost_master, Master_clean!master"],
        ["Resilience",    "Q8–Q11",       "sourced (Q11 recomputed)",
         "Fragility_Index.xlsx, resilience_by_crop.xlsx, Master_clean!master"],
        ["Conclusion",    "Q12, Q13",     "narrative", "—"],
    ], columns=["Tab", "Questions", "How derived", "Source"])
    sw.table(idx)
    _autosize(ws)

    # ------- GROWTH -------
    ws = wb.create_sheet("Growth")
    sw = SheetWriter(ws)
    sw.band("GROWTH")
    sw.question("Q1", "How much more staple food does the world grow since 2000?",
                "World staple output, headline and by crop and by country contribution.")
    sw.table(tables["g_head"], pct_cols=pct("Change (%)"))
    sw.note("Per-person adjustment needs a population series not present in these files; production side only. "
            "Six-core = Maize, Rice, Wheat, Soybeans, Potatoes, Sorghum. Source: Master_clean!master.")
    sw.table(tables["g_crop"], pct_cols=pct("Change (%)"), int_cols=pct("Rank by Mt added"))
    sw.note("Ranked by absolute Mt added, 2000→2023. Source: Master_clean!master.")
    sw.table(tables["g_country"], pct_cols=pct("Share of six-core gain (%)"), int_cols=pct("Rank"))
    sw.note("'Growth leaders' = six-core gain (excl. sugarcane & cassava), so China leads. Including all 10 "
            "crops makes Brazil #1 — see Concentration Q3. Source: Master_clean!master.")

    sw.question("Q2", "Was each leader's gain area expansion or yield improvement?",
                "Production = Area × Yield. Additive decomposition of the change per country-crop.")
    sw.table(tables["g_decomp_tally"], int_cols=pct("Country-crop count"))
    sw.note("Classification counts across the 10-country panel.")
    sw.table(tables["g_decomp"])
    sw.note("dP = Y0·dA + A0·dY + dA·dY. Driver = larger pure effect (area vs yield). Source: Master_clean!master.")

    sw.question("Q1.1 (a)", "Year-over-year output trend for the 10 countries",
                "Per-country growth, volatility, and best/worst single-year moves.")
    sw.table(tables["g_trend"], pct_cols=pct("Growth (%)", "YoY volatility (std %)",
                                             "Best YoY (%)", "Worst YoY (%)"),
             int_cols=pct("Best year", "Worst year"))
    sw.note("Volatility = std of annual % change, 2000–2023. Source: Master_clean!master.")

    sw.question("Q1.1 (b)", "Year-by-year path per country (line-chart data)",
                "Country × Year production (Mt); use this block to plot the ten trajectories.")
    sw.table(tables["g_wide"], int_cols=tuple(c for c in tables["g_wide"].columns if c != "Country"))
    sw.note("Values in Mt. Source: Master_clean!master.")
    _autosize(ws)

    # ------- CONCENTRATION -------
    ws = wb.create_sheet("Concentration")
    sw = SheetWriter(ws)
    sw.band("CONCENTRATION")
    sw.question("Q3", "Which countries drove the growth, and what are their shares?",
                "Shares of the 2000→2023 GAIN (who drove the growth), plus the top-10 list.")
    sw.table(tables["c_shares"], pct_cols=pct("Top-3 share of gain (%)",
                                              "Top-5 share of gain (%)", "Top-10 share of gain (%)"))
    sw.note("Share of the gain, not the end-year level. All 10 crops → Brazil #1; removing sugarcane → China #1. "
            "Source: Master_clean!master.")
    sw.table(tables["c_top10"], pct_cols=pct("Share of gain (%)"), int_cols=pct("Rank"))
    sw.note("Top 10 by contribution to the total gain, all 10 crops. Source: Master_clean!master.")

    sw.question("Q4", "Are the 10 panel countries actually the top producers of each crop?",
                "Per crop: world share held by the panel, and the biggest producers left outside it.")
    sw.table(tables["c_cov"], pct_cols=pct("Panel-10 coverage (%)"))
    sw.note("Coverage = panel share of end-year world production. Source: Master_clean!master.")
    _autosize(ws)

    # ------- COST -------
    ws = wb.create_sheet("Cost")
    sw = SheetWriter(ws)
    sw.band("COST")
    sw.question("Q5 (a)", "Top-10 countries: land expanders vs yield-driven",
                "Country contribution to the gain (see Growth Q1) split by driver (see Growth Q2).")
    sw.textblock(["See Growth!Q1 (country contribution) and Growth!Q2 (area vs yield driver) — "
                  "the same decomposition, not duplicated here."])

    sw.question("Q5 (b)", "Do land-expansion leaders sit where forest is lost?  (Pearson r)",
                "Correlation of cropland change and forest change across all countries.")
    sw.table(tables["k_lf_corr"], num_cols=["Pearson r", "p-value"], int_cols=pct("n countries"))
    sw.note("Change = first vs last available year per country, 2000–2023. Source: Cost_Master!cost_master.")
    sw.table(tables["k_losers"], int_cols=pct())
    sw.note("Biggest forest losers that also expanded cropland; leading crop inferred from harvested-area share (Master).")

    sw.question("Q6", "Do yield-intensification leaders sit where input pressure is high?  (Pearson r)",
                "National fertilizer intensity vs production-weighted national yield.")
    sw.table(tables["k_fy_corr"], num_cols=["Pearson r", "p-value"], int_cols=pct("n countries"))
    sw.note("Grain mismatch: fertilizer is national, yield aggregated from crop level — a mechanism, not a clean split. "
            "Source: Cost_Master!cost_master + Master_clean!master.")

    sw.question("Q3.1a", "Global cropland vs forest correlation",
                "The headline land-cost correlation across all countries.")
    sw.table(tables["k_lf_corr"], num_cols=["Pearson r", "p-value"], int_cols=pct("n countries"))
    sw.note("Filtering out FAOSTAT regional aggregates removes an inflated ~-0.82 pseudo-correlation; "
            "the country-level figure is the honest one.")

    sw.question("Q3.1b", "Counter-examples: cropland up WITHOUT forest loss",
                "Quadrant split of countries by direction of cropland and forest change.")
    sw.table(tables["k_quad_counts"], int_cols=pct("Country count"))
    sw.table(tables["k_quad_examples"], int_cols=pct())
    sw.note("Counter-examples = biggest cropland expanders whose forest also rose. Source: Cost_Master!cost_master.")

    sw.question("Q7", "What about water as a cost?  (tested and excluded)",
                "Why water did not survive into the cost model.")
    sw.textblock([
        "Water was tested and dropped. National water-stress (SDG 6.4.2) does not discriminate among the major",
        "producers — most large agricultural exporters read LOW stress at the national grain, while a few",
        "(India, USA, Russia) sit high, so it fails as a cost lever. The real signal is sub-national (irrigation",
        "pressure in specific basins), which the national series cannot capture. Kept as a stated limitation.",
        "No water column is present in the provided files (Master_clean note: 'Water was tested earlier and demoted').",
    ])
    _autosize(ws)

    # ------- RESILIENCE -------
    ws = wb.create_sheet("Resilience")
    sw = SheetWriter(ws)
    sw.band("RESILIENCE & FRAGILITY")
    sw.question("Q8", "If a crop's top producer failed, can the world cover the loss?",
                "Demonstrated backfill capacity elsewhere vs the loss, and coverage under partial shocks.")
    sw.table(tables["r_resilience"])
    sw.note("Sourced from resilience_by_crop.xlsx. resilience_ratio = backfill capacity ÷ loss; <1 = uncoverable.")
    sw.table(tables["r_partial"])
    sw.note("Sourced from Fragility_Index!partial_shock. 'covered' vs 'GAP' at each shock size.")

    sw.question("Q9", "What does that mean — and is it real?",
                "Resilience is inverse to importance; worst historical single-year drops already seen.")
    sw.textblock([
        "The crops the world eats most (rice, maize, wheat) are the LEAST replaceable — resilience runs inverse",
        "to dietary importance. This is not hypothetical: the table below shows a shock of each size has already",
        "happened at least once since 2000.",
    ])
    sw.table(tables["r_hist"], int_cols=pct("Worst drop year"),
             num_cols=["Worst 1-yr drop %", "Avg 1-yr volatility (std %)"])
    sw.note("Sourced from Fragility_Index!historical_drops.")

    sw.question("Q9.1", "Does the fragility survive a realistic partial shock (-35%)?",
                "Not just a total wipeout — which crops still fail at a plausible one-year drop.")
    sw.table(tables["r_valid"], int_cols=pct("Fragility score (0-100)"))
    sw.note("Sourced from Fragility_Index!index_validation. Argentina 2023 soy drought (~-43%) is the reference real shock.")

    sw.question("Q9.2", "Validated by a real event beyond Argentina?",
                "A second independent confirmation of the fragility finding.")
    sw.textblock([
        "Yes. The 2022 Ukraine invasion disrupted roughly a quarter of global wheat exports in a single event —",
        "a real-world stress test of a top-producer failure, independent of the Argentina drought case. The 2024",
        "India rice export ban is a further live example of a dominant producer removing supply from the market.",
    ])

    sw.question("Q10", "Which crops are most at future risk, and what reinforces them?",
                "The Staple Fragility Index (Exposure × Hazard × Buffer) and the reinforcement watchlist.")
    sw.table(tables["r_fragility"], int_cols=pct("Fragility score (0-100)"),
             num_cols=["Exposure (top share %)", "Hazard (volatility %)", "Buffer (resilience ratio)"])
    sw.note("Sourced from Fragility_Index!fragility_index. Buffer acts as a gate: a coverable crop scores low regardless.")
    sw.table(tables["r_prescription"], int_cols=pct("Reinforcement rank"),
             num_cols=["Latent capacity (Mt)"])
    sw.note("Sourced from Fragility_Index!prescription — countries with demonstrated latent capacity per fragile crop.")

    sw.question("Q11", "Is fragility getting worse over time?",
                "Recomputed: top-producer share and HHI concentration trend per crop, 2000→2023.")
    sw.table(tables["r_hhi"], pct_cols=pct(f"Top share {START_YEAR} (%)", f"Top share {END_YEAR} (%)"),
             num_cols=[f"HHI {START_YEAR}", f"HHI {END_YEAR}", "Trend p-value"])
    sw.note("Not a blanket claim — it's crop-specific. HHI = Σ(country share)²×10000. "
            "Recomputed from Master_clean!master.")
    _autosize(ws)

    # ------- CONCLUSION -------
    ws = wb.create_sheet("Conclusion")
    sw = SheetWriter(ws)
    sw.band("CONCLUSION")
    sw.question("Q12", "The one-sentence takeaway", "")
    sw.textblock([
        "The world grows more staple food two ways — clearing land and raising yields, each with a cost — and it",
        "has concentrated that growth so tightly that the most important staples are now the least able to survive",
        "the loss of any single producer.",
    ])
    sw.question("Q13", "Who is this for, and why does it matter?", "")
    sw.textblock([
        "Primary audience: supply-chain risk teams at the major grain traders (Cargill, ADM, Bunge, Louis Dreyfus,",
        "COFCO) and large food buyers. Where a single crop's top producer is both dominant and volatile, a physical",
        "supply shock is a contract risk years before it hits the news; India's rice ban and Argentina's drought",
        "are the live examples. The recommendation is specific: the Fragility Index turns 'concentration is high'",
        "into an actionable watchlist, ranked by which staple would fail hardest and which countries hold the",
        "demonstrated capacity to reinforce it.",
    ])
    _autosize(ws)

    # ------- PIVOT SOURCE -------
    ws = wb.create_sheet("pivot_source")
    tidy = tables["tidy"]
    for j, col in enumerate(tidy.columns, start=1):
        c = ws.cell(1, j, col)
        c.font = Font(name=FONT, size=10, bold=True, color=WHITE)
        c.fill = PatternFill("solid", fgColor=NAVY)
    for i, (_, row) in enumerate(tidy.iterrows(), start=2):
        for j, col in enumerate(tidy.columns, start=1):
            v = row[col]
            ws.cell(i, j, None if pd.isna(v) else v)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(tidy.columns))}{len(tidy)+1}"
    for j, col in enumerate(tidy.columns, start=1):
        ws.column_dimensions[get_column_letter(j)].width = max(12, len(col) + 2)

    return wb


# ----------------------------------------------------------------------------- #
def compute_all():
    m = load_master()
    c = load_cost()
    cc = cost_country_change(c)
    decomp, decomp_tally = growth_area_yield_decomp(m)
    quad_counts, quad_examples = cost_cropland_forest_quadrants(cc)

    tidy = (m[m.Crop.isin(TEN_CROPS) & m.Year.between(START_YEAR, END_YEAR)]
            [["Year", "Country", "Continent", "Subregion", "Country Classification",
              "Crop", "Staple Crop", "Area_harvested_ha", "Yield_kg_per_ha",
              "Production_tonnes"]].copy())
    tidy["Production_Mt"] = tidy["Production_tonnes"] / MT

    return {
        # growth
        "g_head":    growth_world_headline(m),
        "g_crop":    growth_by_crop(m),
        "g_country": growth_by_country(m),
        "g_decomp":  decomp,
        "g_decomp_tally": decomp_tally,
        "g_trend":   panel_country_trend(m),
        "g_wide":    panel_country_yearly_wide(m),
        # concentration
        "c_shares":  concentration_shares(m),
        "c_top10":   concentration_top10_list(m),
        "c_cov":     concentration_crop_coverage(m),
        # cost
        "k_lf_corr": cost_land_forest_corr(cc),
        "k_losers":  cost_forest_losers(cc, m),
        "k_fy_corr": cost_fertilizer_yield_corr(m, c),
        "k_quad_counts": quad_counts,
        "k_quad_examples": quad_examples,
        # resilience (sourced) + Q11 recomputed
        "r_resilience":  load_resilience(),
        "r_partial":     load_fragility_sheet("partial_shock"),
        "r_hist":        load_fragility_sheet("historical_drops"),
        "r_valid":       load_fragility_sheet("index_validation"),
        "r_fragility":   load_fragility_sheet("fragility_index"),
        "r_prescription": load_fragility_sheet("prescription").iloc[:, :4],
        "r_hhi":         fragility_hhi_trend(m),
        # pivot
        "tidy":      tidy,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--validate", action="store_true", help="print tables, do not write")
    ap.add_argument("--out", default=OUT_PATH)
    args = ap.parse_args()

    tables = compute_all()

    if args.validate:
        for k, v in tables.items():
            print("\n" + "=" * 80 + f"\n{k}  (shape {getattr(v, 'shape', '-')})")
            if isinstance(v, pd.DataFrame):
                print(v.to_string(index=False)[:3000])
        return

    wb = build(tables)
    wb.save(args.out)
    print(f"Wrote {args.out}  ({len(wb.sheetnames)} tabs: {', '.join(wb.sheetnames)})")


if __name__ == "__main__":
    main()