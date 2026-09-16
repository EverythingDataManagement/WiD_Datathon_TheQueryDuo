#!/usr/bin/env python3
# =====================================================================
#  build_master.py  —  Query Queens, GROW track
#  Rebuilds the clean Master from the raw FAOSTAT production file.
#  Run once; get the same result every time. Never hand-edit the output.
#
#  Colab setup (run these two lines in a cell FIRST):
#     !pip install country_converter openpyxl --quiet
#     # then upload the raw file, then paste + run everything below.
#
#  Tabs produced: read_me, master, dropped_areas, cleaning_log,
#                 reference_data, data_quality, dq_scorecard, country_lists
# =====================================================================

import warnings
import pandas as pd
import numpy as np

# country_converter derives continent + UN subregion from country names.
# If the library is missing, USE_CC flips to False and the embedded
# REGION_FALLBACK map (below, built from the same source) is used instead.
USE_CC = True
try:
    import country_converter as coco
except Exception:
    USE_CC = False

# ---------- SETTINGS ----------
RAW_FILE   = "Production_Crops_Livestock_E_All_Data (reviewed).xlsx"
RAW_SHEET  = "Sheet 1 - Production_Crops_Live"
OUT_FILE   = "Master_clean.xlsx"
START_YEAR, END_YEAR = 2000, 2024

CROPMAP = {
    "Maize (corn)": "Maize", "Rice": "Rice", "Wheat": "Wheat",
    "Soya beans": "Soybeans", "Potatoes": "Potatoes", "Sugar cane": "Sugarcane",
    "Cassava, fresh": "Cassava", "Sorghum": "Sorghum", "Millet": "Millet", "Barley": "Barley",
}
CATEGORY = {
    "Maize": ("Cereals", "Coarse grains"), "Sorghum": ("Cereals", "Coarse grains"),
    "Millet": ("Cereals", "Coarse grains"), "Barley": ("Cereals", "Coarse grains"),
    "Rice": ("Cereals", "Rice"), "Wheat": ("Cereals", "Wheat"),
    "Potatoes": ("Roots & Tubers", "Tuber"), "Cassava": ("Roots & Tubers", "Root"),
    "Soybeans": ("Oil Crops", "Oilseed"), "Sugarcane": ("Sugar Crops", "Sugar"),
}
ELEMENTS = {"Area harvested": ("Area_harvested_ha", "ha"),
            "Yield": ("Yield_kg_per_ha", "kg/ha"),
            "Production": ("Production_tonnes", "tonnes")}
# FAOSTAT reviewed-file value flags. Confirm against the file's own notes tab;
# FAOSTAT has revised this scheme across versions.
FLAG_MEANING = {"A": "Official figure", "E": "Estimated value", "I": "Imputed value",
                "M": "Missing value (data cannot exist / not applicable)",
                "P": "Provisional value", "X": "Figure from an international organization",
                "": "No flag recorded"}

PANEL_10 = ["Argentina", "Brazil", "China, mainland", "India", "Nigeria",
            "Pakistan", "Russian Federation", "Thailand", "Ukraine",
            "United States of America"]
# OLD master country set was 196; paste it here if you want the exact diff printed.
# Leave empty to skip the comparison.
OLD_COUNTRIES = []   # e.g. ["Afghanistan", "Albania", ...]

AGGREGATES = [
    "World", "Africa", "Americas", "Asia", "Europe", "Oceania",
    "Eastern Africa", "Middle Africa", "Northern Africa", "Southern Africa", "Western Africa",
    "Northern America", "Central America", "Caribbean", "South America",
    "Latin America and the Caribbean", "Sub-Saharan Africa",
    "Central Asia", "Eastern Asia", "Southern Asia", "South-eastern Asia", "Western Asia",
    "Eastern Europe", "Northern Europe", "Southern Europe", "Western Europe",
    "Australia and New Zealand", "Melanesia", "Micronesia", "Polynesia",
    "European Union (27)", "Least Developed Countries (LDCs)", "Least Developed Countries",
    "Land Locked Developing Countries (LLDCs)", "Land Locked Developing Countries",
    "Low Income Food Deficit Countries (LIFDCs)", "Low Income Food Deficit Countries",
    "Net Food Importing Developing Countries (NFIDCs)", "Net Food Importing Developing Countries",
    "Small Island Developing States (SIDS)", "Small Island Developing States",
    "High-income economies", "Low-income economies", "Lower-middle-income economies",
    "Upper-middle-income economies", "Middle-income economies",
]
DUP_DROP = {
    "China": "kept 'China, mainland' (China aggregate double-counts mainland+Taiwan+HK+Macao)",
    "China, Taiwan Province of": "not a UN reporting member; excluded for consistency",
    "Serbia and Montenegro": "former country; superseded by Serbia + Montenegro",
    "Sudan (former)": "former country; superseded by Sudan + South Sudan",
    "Belgium-Luxembourg": "former unit; superseded by Belgium + Luxembourg",
    "USSR": "former country; superseded by successor states",
    "Czechoslovakia": "former country; superseded by Czechia + Slovakia",
    "Yugoslav SFR": "former country; superseded by successor states",
    "Ethiopia PDR": "former country; superseded by Ethiopia",
}

REGION_FALLBACK = {   # country -> (continent, UN subregion); built from country_converter
    'Afghanistan': ('Asia', 'Southern Asia'),
    'Albania': ('Europe', 'Southern Europe'),
    'Algeria': ('Africa', 'Northern Africa'),
    'American Samoa': ('Oceania', 'Polynesia'),
    'Andorra': ('Europe', 'Southern Europe'),
    'Angola': ('Africa', 'Middle Africa'),
    'Anguilla': ('America', 'Caribbean'),
    'Antigua and Barbuda': ('America', 'Caribbean'),
    'Argentina': ('America', 'South America'),
    'Armenia': ('Asia', 'Western Asia'),
    'Aruba': ('America', 'Caribbean'),
    'Ascension, Saint Helena and Tristan da Cunha': ('Africa', 'Western Africa'),
    'Australia': ('Oceania', 'Australia and New Zealand'),
    'Austria': ('Europe', 'Western Europe'),
    'Azerbaijan': ('Asia', 'Western Asia'),
    'Bahamas': ('America', 'Caribbean'),
    'Bahrain': ('Asia', 'Western Asia'),
    'Bangladesh': ('Asia', 'Southern Asia'),
    'Barbados': ('America', 'Caribbean'),
    'Belarus': ('Europe', 'Eastern Europe'),
    'Belgium': ('Europe', 'Western Europe'),
    'Belize': ('America', 'Central America'),
    'Benin': ('Africa', 'Western Africa'),
    'Bermuda': ('America', 'Northern America'),
    'Bhutan': ('Asia', 'Southern Asia'),
    'Bolivia (Plurinational State of)': ('America', 'South America'),
    'Bonaire, Sint Eustatius and Saba': ('America', 'Caribbean'),
    'Bosnia and Herzegovina': ('Europe', 'Southern Europe'),
    'Botswana': ('Africa', 'Southern Africa'),
    'Brazil': ('America', 'South America'),
    'British Virgin Islands': ('America', 'Caribbean'),
    'Brunei Darussalam': ('Asia', 'South-eastern Asia'),
    'Bulgaria': ('Europe', 'Eastern Europe'),
    'Burkina Faso': ('Africa', 'Western Africa'),
    'Burundi': ('Africa', 'Eastern Africa'),
    'Cabo Verde': ('Africa', 'Western Africa'),
    'Cambodia': ('Asia', 'South-eastern Asia'),
    'Cameroon': ('Africa', 'Middle Africa'),
    'Canada': ('America', 'Northern America'),
    'Cayman Islands': ('America', 'Caribbean'),
    'Central African Republic': ('Africa', 'Middle Africa'),
    'Chad': ('Africa', 'Middle Africa'),
    'Channel Islands': ('Europe', 'Northern Europe'),
    'Chile': ('America', 'South America'),
    'China, Hong Kong SAR': ('Asia', 'Eastern Asia'),
    'China, Macao SAR': ('Asia', 'Eastern Asia'),
    'China, mainland': ('Asia', 'Eastern Asia'),
    'Colombia': ('America', 'South America'),
    'Comoros': ('Africa', 'Eastern Africa'),
    'Congo': ('Africa', 'Middle Africa'),
    'Cook Islands': ('Oceania', 'Polynesia'),
    'Costa Rica': ('America', 'Central America'),
    'Croatia': ('Europe', 'Southern Europe'),
    'Cuba': ('America', 'Caribbean'),
    'Curaçao': ('America', 'Caribbean'),
    'Cyprus': ('Asia', 'Western Asia'),
    'Czechia': ('Europe', 'Eastern Europe'),
    "Côte d'Ivoire": ('Africa', 'Western Africa'),
    "Democratic People's Republic of Korea": ('Asia', 'Eastern Asia'),
    'Democratic Republic of the Congo': ('Africa', 'Middle Africa'),
    'Denmark': ('Europe', 'Northern Europe'),
    'Djibouti': ('Africa', 'Eastern Africa'),
    'Dominica': ('America', 'Caribbean'),
    'Dominican Republic': ('America', 'Caribbean'),
    'Ecuador': ('America', 'South America'),
    'Egypt': ('Africa', 'Northern Africa'),
    'El Salvador': ('America', 'Central America'),
    'Equatorial Guinea': ('Africa', 'Middle Africa'),
    'Eritrea': ('Africa', 'Eastern Africa'),
    'Estonia': ('Europe', 'Northern Europe'),
    'Eswatini': ('Africa', 'Southern Africa'),
    'Ethiopia': ('Africa', 'Eastern Africa'),
    'Falkland Islands (Malvinas)': ('America', 'South America'),
    'Faroe Islands': ('Europe', 'Northern Europe'),
    'Fiji': ('Oceania', 'Melanesia'),
    'Finland': ('Europe', 'Northern Europe'),
    'France': ('Europe', 'Western Europe'),
    'French Guiana': ('America', 'South America'),
    'French Polynesia': ('Oceania', 'Polynesia'),
    'Gabon': ('Africa', 'Middle Africa'),
    'Gambia': ('Africa', 'Western Africa'),
    'Georgia': ('Asia', 'Western Asia'),
    'Germany': ('Europe', 'Western Europe'),
    'Ghana': ('Africa', 'Western Africa'),
    'Gibraltar': ('Europe', 'Southern Europe'),
    'Greece': ('Europe', 'Southern Europe'),
    'Greenland': ('America', 'Northern America'),
    'Grenada': ('America', 'Caribbean'),
    'Guadeloupe': ('America', 'Caribbean'),
    'Guam': ('Oceania', 'Micronesia'),
    'Guatemala': ('America', 'Central America'),
    'Guinea': ('Africa', 'Western Africa'),
    'Guinea-Bissau': ('Africa', 'Western Africa'),
    'Guyana': ('America', 'South America'),
    'Haiti': ('America', 'Caribbean'),
    'Holy See': ('Europe', 'Southern Europe'),
    'Honduras': ('America', 'Central America'),
    'Hungary': ('Europe', 'Eastern Europe'),
    'Iceland': ('Europe', 'Northern Europe'),
    'India': ('Asia', 'Southern Asia'),
    'Indonesia': ('Asia', 'South-eastern Asia'),
    'Iran (Islamic Republic of)': ('Asia', 'Southern Asia'),
    'Iraq': ('Asia', 'Western Asia'),
    'Ireland': ('Europe', 'Northern Europe'),
    'Isle of Man': ('Europe', 'Northern Europe'),
    'Israel': ('Asia', 'Western Asia'),
    'Italy': ('Europe', 'Southern Europe'),
    'Jamaica': ('America', 'Caribbean'),
    'Japan': ('Asia', 'Eastern Asia'),
    'Jordan': ('Asia', 'Western Asia'),
    'Kazakhstan': ('Asia', 'Central Asia'),
    'Kenya': ('Africa', 'Eastern Africa'),
    'Kiribati': ('Oceania', 'Micronesia'),
    'Kuwait': ('Asia', 'Western Asia'),
    'Kyrgyzstan': ('Asia', 'Central Asia'),
    "Lao People's Democratic Republic": ('Asia', 'South-eastern Asia'),
    'Latvia': ('Europe', 'Northern Europe'),
    'Lebanon': ('Asia', 'Western Asia'),
    'Lesotho': ('Africa', 'Southern Africa'),
    'Liberia': ('Africa', 'Western Africa'),
    'Libya': ('Africa', 'Northern Africa'),
    'Liechtenstein': ('Europe', 'Western Europe'),
    'Lithuania': ('Europe', 'Northern Europe'),
    'Luxembourg': ('Europe', 'Western Europe'),
    'Madagascar': ('Africa', 'Eastern Africa'),
    'Malawi': ('Africa', 'Eastern Africa'),
    'Malaysia': ('Asia', 'South-eastern Asia'),
    'Maldives': ('Asia', 'Southern Asia'),
    'Mali': ('Africa', 'Western Africa'),
    'Malta': ('Europe', 'Southern Europe'),
    'Marshall Islands': ('Oceania', 'Micronesia'),
    'Martinique': ('America', 'Caribbean'),
    'Mauritania': ('Africa', 'Western Africa'),
    'Mauritius': ('Africa', 'Eastern Africa'),
    'Mayotte': ('Africa', 'Eastern Africa'),
    'Mexico': ('America', 'Central America'),
    'Micronesia (Federated States of)': ('Oceania', 'Micronesia'),
    'Monaco': ('Europe', 'Western Europe'),
    'Mongolia': ('Asia', 'Eastern Asia'),
    'Montenegro': ('Europe', 'Southern Europe'),
    'Montserrat': ('America', 'Caribbean'),
    'Morocco': ('Africa', 'Northern Africa'),
    'Mozambique': ('Africa', 'Eastern Africa'),
    'Myanmar': ('Asia', 'South-eastern Asia'),
    'Namibia': ('Africa', 'Southern Africa'),
    'Naoero': ('Oceania', 'Micronesia'),
    'Nepal': ('Asia', 'Southern Asia'),
    'Netherlands (Kingdom of the)': ('Europe', 'Western Europe'),
    'New Caledonia': ('Oceania', 'Melanesia'),
    'New Zealand': ('Oceania', 'Australia and New Zealand'),
    'Nicaragua': ('America', 'Central America'),
    'Niger': ('Africa', 'Western Africa'),
    'Nigeria': ('Africa', 'Western Africa'),
    'Niue': ('Oceania', 'Polynesia'),
    'Norfolk Island': ('Oceania', 'Australia and New Zealand'),
    'North Macedonia': ('Europe', 'Southern Europe'),
    'Northern Mariana Islands': ('Oceania', 'Micronesia'),
    'Norway': ('Europe', 'Northern Europe'),
    'Oman': ('Asia', 'Western Asia'),
    'Pakistan': ('Asia', 'Southern Asia'),
    'Palau': ('Oceania', 'Micronesia'),
    'Palestine': ('Asia', 'Western Asia'),
    'Panama': ('America', 'Central America'),
    'Papua New Guinea': ('Oceania', 'Melanesia'),
    'Paraguay': ('America', 'South America'),
    'Peru': ('America', 'South America'),
    'Philippines': ('Asia', 'South-eastern Asia'),
    'Pitcairn': ('Oceania', 'Polynesia'),
    'Poland': ('Europe', 'Eastern Europe'),
    'Portugal': ('Europe', 'Southern Europe'),
    'Puerto Rico': ('America', 'Caribbean'),
    'Qatar': ('Asia', 'Western Asia'),
    'Republic of Korea': ('Asia', 'Eastern Asia'),
    'Republic of Moldova': ('Europe', 'Eastern Europe'),
    'Romania': ('Europe', 'Eastern Europe'),
    'Russian Federation': ('Europe', 'Eastern Europe'),
    'Rwanda': ('Africa', 'Eastern Africa'),
    'Réunion': ('Africa', 'Eastern Africa'),
    'Saint Barthélemy': ('America', 'Caribbean'),
    'Saint Kitts and Nevis': ('America', 'Caribbean'),
    'Saint Lucia': ('America', 'Caribbean'),
    'Saint Martin (French part)': ('America', 'Caribbean'),
    'Saint Pierre and Miquelon': ('America', 'Northern America'),
    'Saint Vincent and the Grenadines': ('America', 'Caribbean'),
    'Samoa': ('Oceania', 'Polynesia'),
    'San Marino': ('Europe', 'Southern Europe'),
    'Sao Tome and Principe': ('Africa', 'Middle Africa'),
    'Saudi Arabia': ('Asia', 'Western Asia'),
    'Senegal': ('Africa', 'Western Africa'),
    'Serbia': ('Europe', 'Southern Europe'),
    'Seychelles': ('Africa', 'Eastern Africa'),
    'Sierra Leone': ('Africa', 'Western Africa'),
    'Singapore': ('Asia', 'South-eastern Asia'),
    'Sint Maarten (Dutch part)': ('America', 'Caribbean'),
    'Slovakia': ('Europe', 'Eastern Europe'),
    'Slovenia': ('Europe', 'Southern Europe'),
    'Solomon Islands': ('Oceania', 'Melanesia'),
    'Somalia': ('Africa', 'Eastern Africa'),
    'South Africa': ('Africa', 'Southern Africa'),
    'South Sudan': ('Africa', 'Eastern Africa'),
    'Spain': ('Europe', 'Southern Europe'),
    'Sri Lanka': ('Asia', 'Southern Asia'),
    'Sudan': ('Africa', 'Northern Africa'),
    'Suriname': ('America', 'South America'),
    'Sweden': ('Europe', 'Northern Europe'),
    'Switzerland': ('Europe', 'Western Europe'),
    'Syrian Arab Republic': ('Asia', 'Western Asia'),
    'Tajikistan': ('Asia', 'Central Asia'),
    'Thailand': ('Asia', 'South-eastern Asia'),
    'Timor-Leste': ('Asia', 'South-eastern Asia'),
    'Togo': ('Africa', 'Western Africa'),
    'Tokelau': ('Oceania', 'Polynesia'),
    'Tonga': ('Oceania', 'Polynesia'),
    'Trinidad and Tobago': ('America', 'Caribbean'),
    'Tunisia': ('Africa', 'Northern Africa'),
    'Turkmenistan': ('Asia', 'Central Asia'),
    'Turks and Caicos Islands': ('America', 'Caribbean'),
    'Tuvalu': ('Oceania', 'Polynesia'),
    'Türkiye': ('Asia', 'Western Asia'),
    'Uganda': ('Africa', 'Eastern Africa'),
    'Ukraine': ('Europe', 'Eastern Europe'),
    'United Arab Emirates': ('Asia', 'Western Asia'),
    'United Kingdom of Great Britain and Northern Ireland': ('Europe', 'Northern Europe'),
    'United Republic of Tanzania': ('Africa', 'Eastern Africa'),
    'United States Virgin Islands': ('America', 'Caribbean'),
    'United States of America': ('America', 'Northern America'),
    'Uruguay': ('America', 'South America'),
    'Uzbekistan': ('Asia', 'Central Asia'),
    'Vanuatu': ('Oceania', 'Melanesia'),
    'Venezuela (Bolivarian Republic of)': ('America', 'South America'),
    'Viet Nam': ('Asia', 'South-eastern Asia'),
    'Wallis and Futuna Islands': ('Oceania', 'Polynesia'),
    'Western Sahara': ('Africa', 'Northern Africa'),
    'Yemen': ('Asia', 'Western Asia'),
    'Zambia': ('Africa', 'Eastern Africa'),
    'Zimbabwe': ('Africa', 'Eastern Africa'),
}

log = []
def rec(step, what, count):
    log.append({"Step": step, "What happened": what, "Count": count})
    print(f"[{step}] {what}: {count}")

# =====================================================================
# 1. LOAD
# =====================================================================
raw = pd.read_excel(RAW_FILE, sheet_name=RAW_SHEET)
rec("1. Load", "Raw rows loaded", len(raw))
rec("1. Load", "Distinct areas in raw file", raw["Area"].nunique())
rec("1. Load", "Distinct items in raw file", raw["Item"].nunique())

# =====================================================================
# 2. KEEP the 10 crops + 3 crop elements
# =====================================================================
d = raw[raw["Item"].isin(CROPMAP) & raw["Element"].isin(ELEMENTS)].copy()
rec("2. Filter crops", "Rows after keeping 10 crops x 3 elements", len(d))

# =====================================================================
# 3. RESHAPE — melt BOTH the value column (YXXXX) and its flag (YXXXXF)
#    so every country-crop-year keeps its real per-year flag.
# =====================================================================
val_cols  = [c for c in d.columns if c.startswith("Y") and c[1:].isdigit()]           # Y2000 ...
# melt values
vlong = d.melt(id_vars=["Area", "Item", "Element"], value_vars=val_cols,
               var_name="Yvar", value_name="value")
vlong["Year"] = vlong["Yvar"].str[1:].astype(int)
vlong["metric"] = vlong["Element"].map(lambda e: ELEMENTS[e][0])
vlong = vlong[vlong["Year"].between(START_YEAR, END_YEAR)]
# melt flags (Y2000F ...) the same way, then join on Area/Item/Element/Year
flag_cols = [c for c in d.columns if c.startswith("Y") and c.endswith("F") and c[1:-1].isdigit()]
flong = d.melt(id_vars=["Area", "Item", "Element"], value_vars=flag_cols,
               var_name="Fvar", value_name="flag")
flong["Year"] = flong["Fvar"].str[1:-1].astype(int)
flong = flong[flong["Year"].between(START_YEAR, END_YEAR)]
flong = flong.drop_duplicates(subset=["Area", "Item", "Element", "Year"])
vlong = vlong.merge(flong[["Area", "Item", "Element", "Year", "flag"]],
                    on=["Area", "Item", "Element", "Year"], how="left")

# pivot the three measures into columns; keep the PRODUCTION flag as the row flag
tidy = vlong.pivot_table(index=["Area", "Item", "Year"], columns="metric",
                         values="value", aggfunc="first").reset_index()
# row-level flag = the flag on the Production value (the headline measure)
prodflag = (vlong[vlong["Element"] == "Production"][["Area", "Item", "Year", "flag"]]
            .drop_duplicates(subset=["Area", "Item", "Year"])
            .rename(columns={"flag": "Flag"}))
tidy = tidy.merge(prodflag, on=["Area", "Item", "Year"], how="left")
rec("3. Reshape", "Tidy rows (country x crop x year)", len(tidy))

# =====================================================================
# 4. DROP aggregates + duplicate/former units (track why)
# =====================================================================
dropped = []
for area in sorted(tidy["Area"].unique()):
    if area in AGGREGATES:
        dropped.append({"Area": area, "Reason": "Regional / economic aggregate (a total, not a country)"})
    elif area in DUP_DROP:
        dropped.append({"Area": area, "Reason": DUP_DROP[area]})
dropped_df = pd.DataFrame(dropped)
before = tidy["Area"].nunique()
tidy = tidy[~tidy["Area"].isin(AGGREGATES) & ~tidy["Area"].isin(DUP_DROP)]
rec("4. Drop aggregates", "Areas removed", before - tidy["Area"].nunique())
rec("4. Drop aggregates", "Countries remaining", tidy["Area"].nunique())

# =====================================================================
# 5. RENAME + DERIVE continent/subregion FROM the country + taxonomy
# =====================================================================
tidy = tidy.rename(columns={"Area": "Country"})
tidy["Crop"] = tidy["Item"].map(CROPMAP)
tidy["Crop Category"]    = tidy["Crop"].map(lambda c: CATEGORY[c][0])
tidy["Crop Subcategory"] = tidy["Crop"].map(lambda c: CATEGORY[c][1])

countries = sorted(tidy["Country"].unique())

VALID_CONTINENTS = {"Africa", "America", "Asia", "Europe", "Oceania", "Antarctica"}
def derive_regions(country_list):
    """country -> (continent, subregion). Uses country_converter when available,
    else the embedded REGION_FALLBACK, so regions populate even offline. Never
    silently blanks: anything unresolved returns (None, None) and is reported."""
    out = {}
    if USE_CC:
        cc = coco.CountryConverter()
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")   # mute "more than one regex match" noise
            cont = cc.convert(country_list, to="continent", not_found=None)
            subr = cc.convert(country_list, to="UNregion",  not_found=None)
        for name, c, u in zip(country_list, cont, subr):
            c = c[0] if isinstance(c, (list, tuple)) and c else c   # ambiguous -> first
            u = u[0] if isinstance(u, (list, tuple)) and u else u
            if c not in VALID_CONTINENTS:      # coco echoes the name when it cannot match
                c, u = REGION_FALLBACK.get(name, (None, None))
            out[name] = (c, u)
    else:
        for name in country_list:
            out[name] = REGION_FALLBACK.get(name, (None, None))
    return out

regions = derive_regions(countries)
tidy["Continent"] = tidy["Country"].map(lambda c: regions[c][0])
tidy["Subregion"] = tidy["Country"].map(lambda c: regions[c][1])
unmatched = sorted({c for c in countries if regions[c][0] is None})
rec("5. Derive region", "Countries with no region (add to REGION_FALLBACK)", len(unmatched))
if unmatched:
    print("   unmatched -> add to REGION_FALLBACK:", unmatched)

# ---- Staple flag + country rankings/classification (enrichment) ----
# All ten selected crops are staples EXCEPT sugarcane, which we treat as an
# industrial / production commodity rather than a dietary staple. Move it into
# STAPLE_CROPS if you want to count it as a staple.
STAPLE_CROPS = {"Maize","Rice","Wheat","Soybeans","Potatoes","Cassava","Sorghum","Millet","Barley"}
tidy["Staple Crop"] = tidy["Crop"].map(lambda c: "Yes" if c in STAPLE_CROPS else "No")

# Country rankings by TOTAL production (tonnes) over the window.
# Ranking 1 = primary-scope crops; Ranking 2 = all ten. 1 = biggest producer.
RANK1_CROPS = ["Rice", "Sorghum", "Soybeans", "Wheat", "Potatoes"]   # primary scope
RANK2_CROPS = list(CROPMAP.values())                                 # all 10
def _country_rank(crop_set):
    tot = tidy[tidy["Crop"].isin(crop_set)].groupby("Country")["Production_tonnes"].sum()
    return tot.rank(ascending=False, method="min")
_r1 = _country_rank(RANK1_CROPS)
_r2 = _country_rank(RANK2_CROPS)
tidy["Country Ranking 1"] = tidy["Country"].map(_r1).astype("Int64")
tidy["Country Ranking 2"] = tidy["Country"].map(_r2).astype("Int64")
def _classify(rk):
    if pd.isna(rk): return "Other"
    rk = int(rk)
    if rk <= 3:  return "Top 3"
    if rk <= 5:  return "Top 5"
    if rk <= 10: return "Top 10"
    return "Other"
# Classification is based on Ranking 2 (all 10 crops). Change to Country Ranking 1 if preferred.
tidy["Country Classification"] = tidy["Country Ranking 2"].map(_classify)

# =====================================================================
# 6. _DQ columns — Zero / Blank / (value) per measure
# =====================================================================
def dq(series):
    out = pd.Series(index=series.index, dtype="object")
    out[series.isna()] = "Blank"
    out[series == 0] = "Zero"
    return out
for clean, _u in ELEMENTS.values():
    tidy[f"{clean}_DQ"] = dq(tidy[clean])

# keep the raw FAOSTAT flag CODE (A/E/I/M/P/X) in Flag; reference_data maps codes -> meaning.
tidy["Flag"] = tidy["Flag"].fillna("").astype(str).str.strip()

# =====================================================================
# 7. FINAL COLUMN ORDER
# =====================================================================
COLS = ["Year", "Country", "Continent", "Subregion",
        "Country Classification", "Country Ranking 1", "Country Ranking 2",
        "Crop", "Item", "Crop Category", "Crop Subcategory", "Staple Crop", "Flag",
        "Area_harvested_ha", "Area_harvested_ha_DQ",
        "Yield_kg_per_ha", "Yield_kg_per_ha_DQ",
        "Production_tonnes", "Production_tonnes_DQ"]
master = tidy[COLS].sort_values(["Country", "Crop", "Year"]).reset_index(drop=True)
rec("6. Master", "FINAL master rows", len(master))
rec("6. Master", "Countries", master["Country"].nunique())
rec("6. Master", "Crops", master["Crop"].nunique())

# ---- country diff vs the old master (answers 'which country is missing?') ----
if OLD_COUNTRIES:
    new = set(master["Country"]); old = set(OLD_COUNTRIES)
    print("\n=== COUNTRY DIFF vs old master ===")
    print("In OLD, not in NEW (dropped):", sorted(old - new))
    print("In NEW, not in OLD (added)  :", sorted(new - old))

# =====================================================================
# 8. DATA QUALITY — per-column metrics
# =====================================================================
N = len(master)
FIELD_META = {
 "Year": ("Key","int","Reporting year","year","2000–2024","Which year the figure is for."),
 "Country": ("Key","text","Producer country","name","195 values","Cleaned to real reporting countries."),
 "Continent": ("Dimension","text","Continent (derived from country)","name","6 values","Derived via country_converter."),
 "Subregion": ("Dimension","text","UN subregion (derived from country)","name","~17 values","Derived via country_converter."),
 "Crop": ("Key","text","Clean crop name","name","10 staples","Canonical analysis key."),
 "Item": ("Provenance","text","Raw FAOSTAT item name","name","10 values","Traceability; join on Crop not Item."),
 "Crop Category": ("Dimension","text","Staple family","name","4 values","Cereals / Roots & Tubers / Oil Crops / Sugar Crops."),
 "Crop Subcategory": ("Dimension","text","Crop sub-family","name","6 values","Coarse grains, Rice, Wheat, Root, Tuber, Oilseed, Sugar."),
 "Country Classification": ("Dimension","text","Top 3 / Top 5 / Top 10 / Other (by total production, all 10 crops)","class","Top 3/5/10/Other","Concentration bucket derived from Country Ranking 2."),
 "Country Ranking 1": ("Dimension","int","Country rank by total production of the primary-scope crops","rank","1..N","Rice, Sorghum, Soybeans, Wheat, Potatoes. 1 = largest."),
 "Country Ranking 2": ("Dimension","int","Country rank by total production of all 10 crops","rank","1..N","1 = largest total producer."),
 "Staple Crop": ("Dimension","text","Whether the crop is a dietary staple","Yes/No","Yes/No","All selected crops except Sugarcane (industrial/production commodity)."),
 "Flag": ("Provenance","text","FAOSTAT value flag (derived from the per-year YxxxxF column)","code","A/E/I/M/P/X","Flag on the Production value for that country-crop-year; see reference_data for meanings."),
 "Area_harvested_ha": ("Measure","float","Harvested area","ha","0 – ~5e7","Input to the area effect."),
 "Area_harvested_ha_DQ": ("QA flag","text","Zero/Blank marker for area","Zero/Blank","Zero/Blank","Marks area = 0, blank, or has a value."),
 "Yield_kg_per_ha": ("Measure","float","Yield","kg/ha","0 – ~1.4e5","Input to the yield effect."),
 "Yield_kg_per_ha_DQ": ("QA flag","text","Zero/Blank marker for yield","Zero/Blank","Zero/Blank","Marks yield = 0, blank, or has a value."),
 "Production_tonnes": ("Measure","float","Production quantity","tonnes","0 – ~8e8","THE headline unit."),
 "Production_tonnes_DQ": ("QA flag","text","Zero/Blank marker for production","Zero/Blank","Zero/Blank","Marks production = 0, blank, or has a value."),
}
rows = []
for col in COLS:
    role, typ, desc, unit, rng, meaning = FIELD_META[col]
    s = master[col]
    blank = int(s.isna().sum())
    zero = int((s == 0).sum()) if pd.api.types.is_numeric_dtype(s) else int(s.astype(str).str.strip().isin(["0","0.0"]).sum())
    rows.append({"Column": col, "Role": role, "Type": typ, "Distinct": int(s.nunique()),
                 "Blank count": blank, "Blank %": round(blank/N*100,2),
                 "Zero count": zero, "Zero %": round(zero/N*100,2),
                 "Completeness %": round(100-blank/N*100,2),
                 "Range / values": rng, "Meaning": meaning, "Unit": unit, "Description": desc})
dq_df = pd.DataFrame(rows)

# =====================================================================
# 9. DQ SCORECARD — score each measure column on DMBOK dimensions,
#    roll up to a dataset DQ score, and explain the zeros/blanks.
# =====================================================================
def score_col(col):
    s = master[col]
    completeness = 100 - s.isna().mean()*100
    # validity: share of non-null values that are >= 0 (no negatives allowed)
    nonnull = s.dropna()
    validity = (nonnull >= 0).mean()*100 if len(nonnull) else 100
    return completeness, validity

# uniqueness + consistency are dataset-level (same for all rows)
dupes = master.duplicated(subset=["Country","Crop","Year"]).sum()
uniqueness = 100 * (1 - dupes/len(master))
# consistency: does Production ~= Area x Yield (yield in kg/ha, area in ha -> kg, /1000 = t)
chk = master.dropna(subset=["Area_harvested_ha","Yield_kg_per_ha","Production_tonnes"])
chk = chk[(chk["Area_harvested_ha"]>0)&(chk["Production_tonnes"]>0)]
implied_t = chk["Area_harvested_ha"]*chk["Yield_kg_per_ha"]/1000
consistency = (np.isclose(implied_t, chk["Production_tonnes"], rtol=0.02)).mean()*100

sc_rows = []
for col in ["Area_harvested_ha","Yield_kg_per_ha","Production_tonnes"]:
    comp, valid = score_col(col)
    sc_rows.append({"Column": col, "Completeness": round(comp,1), "Validity": round(valid,1),
                    "Uniqueness": round(uniqueness,1), "Consistency": round(consistency,1),
                    "Column DQ score": round(np.mean([comp,valid,uniqueness,consistency]),1)})
scorecard = pd.DataFrame(sc_rows)
dataset_dq = round(scorecard["Column DQ score"].mean(),1)

# zero/blank call-outs for the three measures, with interpretation
callouts = []
for col in ["Area_harvested_ha","Yield_kg_per_ha","Production_tonnes"]:
    s = master[col]
    z = int((s==0).sum()); b = int(s.isna().sum())
    callouts.append({"Column": col, "Zero count": z, "Blank count": b,
        "What zeros mean": "Crop genuinely NOT grown in that country-year (a real 0, valid) — NOT missing data.",
        "What blanks mean": "A true reporting gap (the country did not report) — a completeness problem, not a real 0."})
callout_df = pd.DataFrame(callouts)

DMBOK = pd.DataFrame([
    ("Completeness","% of non-blank values.","Per-column above; measures ~98–100%."),
    ("Validity","Values in allowed domain (>= 0).","No negatives; ~100%."),
    ("Uniqueness","No duplicate rows at grain Country x Crop x Year.", f"{dupes} duplicates -> {round(uniqueness,1)}%."),
    ("Consistency","Production ~= Area x Yield (2% tol).", f"{round(consistency,1)}% of rows agree."),
    ("Accuracy","Values reflect the source (FAOSTAT).","No values altered; flags carry source quality (A/E/I)."),
    ("Timeliness","Current enough for the question.","Annual to 2024."),
], columns=["DMBOK dimension","Definition","How this dataset scores"])

# =====================================================================
# 10. REFERENCE DATA + COUNTRY LISTS + READ_ME
# =====================================================================
ref_crop = pd.DataFrame([{"Raw item": k,"Clean crop": v,"Category": CATEGORY[v][0],"Subcategory": CATEGORY[v][1]} for k,v in CROPMAP.items()])
ref_elem = pd.DataFrame([{"Raw element": e,"Clean column": v[0],"Unit": v[1]} for e,v in ELEMENTS.items()])
ref_flag = pd.DataFrame([{"Flag code": k or "(blank)","Meaning": v} for k,v in FLAG_MEANING.items()])
ref_dq   = pd.DataFrame([{"_DQ value":"Zero","Meaning":"measure = 0 (crop not grown that year)"},
                         {"_DQ value":"Blank","Meaning":"measure missing in source"},
                         {"_DQ value":"(empty)","Meaning":"measure has a real non-zero value"}])
panel = pd.DataFrame({"Analysis panel (10)": PANEL_10})
panel["In master?"] = panel["Analysis panel (10)"].isin(master["Country"]).map({True:"yes",False:"MISSING"})

readme = [
 ("Query Queens — master dataset",""),("Women in Data 2026 Datathon, GROW track",""),("",""),
 ("What this file is",""),
 ("One tidy table. Each row = one country, one crop, one year, with area, yield and production.",""),
 ("Rebuilt from the raw FAOSTAT file by build_master.py. DO NOT edit by hand; fix the script and re-run.",""),("",""),
 ("Sheets",""),
 ("master","The dataset. Use this for everything."),
 ("dropped_areas","Every area removed and why."),
 ("cleaning_log","Every step with a count. Read these out in the video."),
 ("reference_data","Lookups: crop names, elements, flag codes, _DQ codes."),
 ("data_quality","Per-column metrics (blank/zero/completeness + meaning)."),
 ("dq_scorecard","DMBOK scores per measure + a dataset DQ score + zero/blank call-outs."),
 ("country_lists","The 10-country analysis panel."),("",""),
 ("Decisions made during cleaning",""),
 ("1. Removed regional/economic aggregates (World, Southern Asia, LDCs...). Totals, not places.",""),
 ("2. Removed duplicate/former units (China aggregate, Serbia and Montenegro, Sudan former...). Kept the canonical country.",""),
 ("3. Kept only the ten agreed crops, matched to FAOSTAT exact item names. Never 'Cereals, primary' (contains wheat/rice/maize).",""),
 ("4. Reshaped the wide file (years across the top) into tidy rows, carrying each year's flag (YxxxxF).",""),
 ("5. Derived Continent + Subregion from the country; derived Flag from the per-year flag column.",""),
 ("6. Added _DQ columns marking each measure as Zero / Blank / has-a-value.",""),("",""),
 ("Known limitations (limitations slide)",""),
 ("A. Production only. Water was tested earlier and demoted; join separately (it ends 2023).",""),
 ("B. Some countries have genuine production gaps in the mid-2000s (Thailand, Bangladesh, Philippines, Peru) - real, not join errors.",""),
 ("C. China, Taiwan Province of excluded (not a UN reporting member).",""),
 ("D. French overseas units kept separate as FAOSTAT reports them.",""),
 ("E. Oil palm (top deforestation crop) is out of scope; crop-level forest cost is inferred, not measured.",""),
]

# =====================================================================
# 11. WRITE + light styling
# =====================================================================
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

def dump(ws, df, start_row, title=None):
    """Write a DataFrame (header + rows) to a worksheet starting at start_row (1-based). Returns next free row."""
    r = start_row
    if title:
        ws.cell(r, 1, title).font = Font(bold=True, color="1F3A22"); r += 1
    for j, col in enumerate(df.columns, 1):
        ws.cell(r, j, str(col))
    r += 1
    for _, row in df.iterrows():
        for j, val in enumerate(row.tolist(), 1):
            # pd.isna catches NaN, pd.NA (from Int64 rank cols) and None alike;
            # the ndim guard keeps it safe if a non-scalar ever slips through.
            is_missing = val is None or (np.ndim(val) == 0 and pd.isna(val))
            ws.cell(r, j, (None if is_missing else val))
        r += 1
    return r + 2   # leave a blank spacer row

wb = Workbook(); wb.remove(wb.active)

# read_me
ws = wb.create_sheet("read_me")
for i, (a, b) in enumerate(readme, 1):
    ws.cell(i, 1, a); ws.cell(i, 2, b)

# simple one-table sheets
for name, df in [("master", master), ("dropped_areas", dropped_df),
                 ("cleaning_log", pd.DataFrame(log)), ("country_lists", panel)]:
    ws = wb.create_sheet(name); dump(ws, df, 1)

# reference_data (stacked lookups)
ws = wb.create_sheet("reference_data"); nr = 1
for t, title in [(ref_crop,"Crop mapping"),(ref_elem,"Element -> column"),(ref_flag,"Flag codes"),(ref_dq,"_DQ codes")]:
    nr = dump(ws, t, nr, title)

# data_quality (per-column table + DMBOK summary)
ws = wb.create_sheet("data_quality")
nr = dump(ws, dq_df, 1, "Per-column data-quality metrics")
dump(ws, DMBOK, nr, "DMBOK dimensions — how we checked each")

# dq_scorecard (scores + dataset score + zero/blank call-outs)
ws = wb.create_sheet("dq_scorecard")
nr = dump(ws, scorecard, 1, "DMBOK score per measure column (0-100)")
nr = dump(ws, pd.DataFrame([{"DATASET DQ SCORE %": dataset_dq}]), nr, "Overall dataset DQ score")
dump(ws, callout_df, nr, "Zero / Blank call-outs — what they point to")

wb.save(OUT_FILE)
wb = load_workbook(OUT_FILE); HDR=PatternFill("solid",fgColor="1F3A22")
for ws in wb.worksheets:
    ws.sheet_view.showGridLines=False
    if ws.title!="read_me":
        for cell in ws[1]:
            if cell.value is not None:
                cell.font=Font(name="Arial",size=10,bold=True,color="FFFFFF"); cell.fill=HDR
                cell.alignment=Alignment(horizontal="left",vertical="center",wrap_text=True)
    for col in ws.columns:
        w=max((len(str(c.value)) for c in col if c.value is not None),default=10)
        ws.column_dimensions[get_column_letter(col[0].column)].width=min(max(w+2,12),55)
wb["read_me"]["A1"].font=Font(name="Arial",size=14,bold=True,color="1F3A22")
wb["read_me"]["A2"].font=Font(name="Arial",size=10,italic=True,color="595959")
wb.save(OUT_FILE)

print(f"\nSaved -> {OUT_FILE}")
print(f"master: {len(master)} rows, {master['Country'].nunique()} countries, {master['Crop'].nunique()} crops")
print(f"DATASET DQ SCORE: {dataset_dq}%")