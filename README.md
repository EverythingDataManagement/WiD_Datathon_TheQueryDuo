# Query Queens — Global Staple Crop Risk Monitor

**Question:** How did the world grow more staple food, what pressures came with that
growth, and how fragile is supply if a leading producer is disrupted?

This project uses FAOSTAT production, land-use and fertilizer data to trace a chain:

**Growth → Concentration → Growth Path → Cost/Pressure → Resilience → Crop Fragility Index**

The final application is a Streamlit risk monitor for 10 staple crops.

## Key findings

- Four countries — Brazil, India, China and the United States — account for roughly
  **64% of positive production growth** across the 10-crop basket from 2000 to 2023.
- Growth followed different paths: Brazil was strongly land-led, while India, China
  and the United States were predominantly yield-led.
- Cropland expansion and forest change are negatively associated across countries
  (**Pearson r ≈ -0.36** in the final country-level analysis).
- Yield/input intensity is positively associated with fertilizer use
  (**Pearson r ≈ +0.58**, n=173). The earlier +0.51 result is retained in the audit trail and reconciled as an earlier five-crop scope, not silently overwritten.
- Under the final recent-baseline resilience model, **6 of 10 crops** have a
  resilience ratio below 1.0.
- The Crop Fragility Index identifies structural vulnerability by combining
  replacement shortfall with exposure and historical downside production risk.

## Repository structure

```text
app.py                         Streamlit application
requirements.txt               Python dependencies
data/
  processed/
    Master_clean.xlsx        Clean production master
    Cost_Master.xlsx           Clean national land/fertilizer master
  raw_sources/
    Inputs_LandUse...xlsx      FAOSTAT Land Use input
    Inputs_Fertilizers...xlsx  FAOSTAT Fertilizer input
src/
  01_build_master.py
  02_build_cost_master.py
  03_analyze_growth_costs.py   Final growth decomposition + correlations
  04_build_cfi_final.py        Canonical final resilience/CFI/sensitivity
  06_build_trade_exposure_optional.py
  08_validate_project.py
  legacy/                      Earlier analytical scripts retained for provenance
outputs/
  Correlation_Proof_Cropland_Forest.xlsx
  Crop_Fragility_Index_Methodology_Summary.xlsx
  fertilizer_yield_method_reconciliation.csv
docs/
  METHODOLOGY.md
  DATA_DICTIONARY.md
  DATA_QUALITY.md
  Data_Quality_Integration_Audit.xlsx
  LIMITATIONS.md
  TRADE_NOTE.md
  PROJECT_AUDIT.md
```

## Reproduce the final analysis

```bash
pip install -r requirements.txt
python src/03_analyze_growth_costs.py
python src/04_build_cfi_final.py
python src/08_validate_project.py
streamlit run app.py
```

`01_build_master.py` requires the original FAOSTAT production export, which is not
included in this package. Re-download it from FAOSTAT and use the filename expected
by the script, or point the script to your local copy.

## Trade

Trade was explored as a supporting importer-exposure illustration. It is **not an input
to the CFI**. The raw Detailed Trade Matrix XLSX used during exploration is not included:
it is above GitHub's normal 100 MB single-file limit and Excel-format exports may be
truncated. See `docs/TRADE_NOTE.md`.

## Interpretation

The CFI is a transparent production-capacity risk model, not a forecast of shortages.
It does not currently model trade availability, inventories, logistics, contracts or
policy restrictions. See `docs/LIMITATIONS.md`.

## Data source

FAOSTAT: Crops and livestock products, Land Use, Fertilizers by Nutrient, and
Detailed Trade Matrix (supporting analysis).
