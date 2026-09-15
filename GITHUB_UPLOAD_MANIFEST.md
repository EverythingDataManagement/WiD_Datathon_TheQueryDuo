# GitHub upload manifest

Upload the contents of this repository folder.

## Required
- `app.py`
- `requirements.txt`
- `README.md`
- `.gitignore`
- `Master_clean.xlsx`
- `data/processed/Cost_Master.xlsx`
- `data/raw_sources/Inputs_LandUse_E_All_Data (Reviewed).xlsx`
- `data/raw_sources/Inputs_FertilizersNutrient_E_All_Data (Reviewed).xlsx`
- `src/01_build_master.py`
- `src/02_build_cost_master.py`
- `src/03_analyze_growth_costs.py`
- `src/04_build_cfi_final.py`
- `src/06_build_trade_exposure_optional.py`
- `src/08_validate_project.py`
- `outputs/*`
- `docs/*` — includes `Data_Quality_Integration_Audit.xlsx` and the final audited roadmap

## Do NOT upload
- `Trade_DetailedTradeMatrix_E_All_Data (updated).xlsx` — ~119 MB, above GitHub's normal single-file limit and may be Excel-row-limit truncated.
- Local virtual environments, caches, `.DS_Store`, secrets.
- Duplicate old app versions.

## Source not currently bundled
The original raw FAOSTAT production export required by `01_build_master.py` was not
among the final mounted project files. The clean production master is included.
For full raw-to-clean reproduction, re-download the FAOSTAT production export and
document its source/version.
