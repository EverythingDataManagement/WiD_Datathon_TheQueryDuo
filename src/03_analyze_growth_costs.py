#!/usr/bin/env python3
"""
03_analyze_growth_costs.py
Recomputes the final growth-path and environmental/input-cost evidence.

Key outputs:
- country_growth_decomposition.csv
- cropland_forest_correlation.csv
- fertilizer_yield_correlation.csv

Important: correlations are associations, not causal estimates.
Land/fertilizer data are national; crop attribution is inferred from production data.
"""
from pathlib import Path
import numpy as np
import pandas as pd
from scipy.stats import pearsonr

ROOT=Path(__file__).resolve().parents[1]
MASTER=ROOT/"Master_clean.xlsx"
COST=ROOT/"data/processed/Cost_Master.xlsx"
OUT=ROOT/"outputs"; OUT.mkdir(exist_ok=True)

m=pd.read_excel(MASTER,sheet_name="master")
c=pd.read_excel(COST,sheet_name="cost_master")
m=m[m.Year.between(2000,2023)]
c=c[c.Year.between(2000,2023)]

# Exact symmetric two-factor decomposition of aggregate P=A*Y.
rows=[]
for country,g in m.groupby("Country"):
    ep=g[g.Year.isin([2000,2023])]
    vals={}
    for y in [2000,2023]:
        z=ep[ep.Year==y]
        A=z.Area_harvested_ha.sum()
        P=z.Production_tonnes.sum()
        Y=P*1000/A if A else np.nan
        vals[y]=(A,P,Y)
    A0,P0,Y0=vals[2000]; A1,P1,Y1=vals[2023]
    if not np.isfinite([A0,P0,Y0,A1,P1,Y1]).all(): continue
    land=(A1-A0)*(Y0+Y1)/2/1000
    yld=(Y1-Y0)*(A0+A1)/2/1000
    dP=P1-P0
    land_share=land/dP if dP>0 else np.nan
    yield_share=yld/dP if dP>0 else np.nan
    if dP>0:
        if land>0 and yld<=0: path="Land-led"
        elif yld>0 and land<=0: path="Yield-led"
        elif land_share>=0.60: path="Land-led"
        elif yield_share>=0.60: path="Yield-led"
        else: path="Mixed"
    else: path="No positive growth"
    rows.append([country,dP,land,yld,land_share,yield_share,path])
pd.DataFrame(rows,columns=["Country","Production change tonnes","Land contribution tonnes",
                           "Yield contribution tonnes","Land share","Yield share","Growth path"])\
  .sort_values("Production change tonnes",ascending=False).to_csv(
      OUT/"country_growth_decomposition.csv",index=False)

# Country endpoint changes for cropland/forest.
# Restrict to countries retained in the cleaned production master; this removes
# territories/non-sovereign reporting units and reproduces the final 190-country result.
analysis_countries = set(m["Country"].dropna().unique())
cc=[]
for country,g in c[c["Country"].isin(analysis_countries)].groupby("Country"):
    z=g.sort_values("Year")
    if 2000 not in set(z.Year) or 2023 not in set(z.Year): continue
    a=z[z.Year==2000].iloc[0]; b=z[z.Year==2023].iloc[0]
    if pd.notna(a.cropland_1000ha) and pd.notna(b.cropland_1000ha) and \
       pd.notna(a.forest_1000ha) and pd.notna(b.forest_1000ha):
        cc.append([country,b.cropland_1000ha-a.cropland_1000ha,
                   b.forest_1000ha-a.forest_1000ha])
cc=pd.DataFrame(cc,columns=["Country","Cropland change 1000ha","Forest change 1000ha"])
r,p=pearsonr(cc["Cropland change 1000ha"],cc["Forest change 1000ha"])
pd.DataFrame([["Cropland change vs forest change",r,p,len(cc)]],
             columns=["Relationship","Pearson r","p-value","n"]).to_csv(
                 OUT/"cropland_forest_correlation.csv",index=False)

# Fertilizer intensity vs national cereal yield, final-year cross-section.
# This avoids treating repeated country-years as independent observations.
cereals=["Maize","Rice","Wheat","Sorghum","Millet","Barley"]
cm=(m[(m.Year==2023)&m.Crop.isin(cereals)]
    .groupby("Country").agg(
        production=("Production_tonnes","sum"),
        area=("Area_harvested_ha","sum")).reset_index())
cm=cm[cm.area>0]
cm["cereal_yield_kg_ha"]=cm.production*1000/cm.area
fert=c[c.Year==2023][["Country","fertilizer_kg_per_ha"]].dropna()
j=cm.merge(fert,on="Country",how="inner").dropna()
j=j[np.isfinite(j.cereal_yield_kg_ha)&np.isfinite(j.fertilizer_kg_per_ha)]
r2,p2=pearsonr(j.fertilizer_kg_per_ha,j.cereal_yield_kg_ha)
pd.DataFrame([["Fertilizer intensity vs cereal yield (2023)",r2,p2,len(j)]],
             columns=["Relationship","Pearson r","p-value","n countries"]).to_csv(
                 OUT/"fertilizer_yield_correlation.csv",index=False)

# Reproduce the earlier roadmap result for auditability. The old "primary" scope
# was Rice, Sorghum, Soybeans, Wheat and Potatoes. This is NOT the final method;
# it is retained to show why the earlier ~+0.51 differs from the final six-cereal result.
old_scope=["Rice","Sorghum","Soybeans","Wheat","Potatoes"]
om=(m[(m.Year==2023)&m.Crop.isin(old_scope)]
    .groupby("Country").agg(
        production=("Production_tonnes","sum"),
        area=("Area_harvested_ha","sum")).reset_index())
om=om[om.area>0]
om["aggregate_yield_kg_ha"]=om.production*1000/om.area
oldj=om.merge(fert,on="Country",how="inner").dropna()
oldj=oldj[np.isfinite(oldj.aggregate_yield_kg_ha)&np.isfinite(oldj.fertilizer_kg_per_ha)]
ro,po=pearsonr(oldj.fertilizer_kg_per_ha,oldj.aggregate_yield_kg_ha)
pd.DataFrame([
    ["Earlier roadmap primary-scope method", ", ".join(old_scope), len(oldj), ro, po, "Historical / superseded"],
    ["Final six-cereal method", ", ".join(cereals), len(j), r2, p2, "Final / authoritative"],
], columns=["Method","Crop scope","n countries","Pearson r","p-value","Status"]).to_csv(
    OUT/"fertilizer_yield_method_reconciliation.csv",index=False)

print("Wrote growth and cost evidence to",OUT)
