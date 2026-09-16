# Data quality and reproducibility

The production master is rebuilt deterministically rather than hand-edited.
The build script removes FAOSTAT regional/economic aggregates and obsolete duplicate
country units before analysis. The workbook contains cleaning logs, reference data,
data-quality fields and a DQ scorecard.

The project previously found that leaving FAOSTAT aggregate rows in the land-use
analysis could create a misleadingly strong cropland/forest relationship. The final
country-level correlation therefore excludes those aggregates.

Source flags are retained in both the production and cost masters. The cost master
also exposes a tidy `source_detail` sheet and flag summaries so estimated/imputed
figures remain visible.

The project DMBOK-style scorecard is approximately 99.6%; this is a project data-quality
assessment, not an external certification.


## Final reconciliation controls

The final audit explicitly reconciles two correlation versions rather than silently overwriting the earlier result. The roadmap's fertilizer/yield result of about +0.51 is reproducible as r=+0.5034 (n=170, p≈2.63e-12) using the earlier five-crop primary scope. The final canonical analysis uses six cereals and returns r=+0.5822 (n=173, p≈4.40e-17). This is logged as a scope/method change.

The land/forest and fertilizer/yield analyses do not use the same complete-case population. Land/forest requires valid 2000 and 2023 cropland and forest endpoints and yields n=190. Fertilizer/yield requires 2023 fertilizer intensity plus positive production/harvested area in the six-cereal basket and yields n=173. No common n is forced and missing analytical inputs are not imputed.

See `Data_Quality_Integration_Audit.xlsx` for the revision log, sample reconciliation, integration controls, source register, approved claims and residual limitations.
