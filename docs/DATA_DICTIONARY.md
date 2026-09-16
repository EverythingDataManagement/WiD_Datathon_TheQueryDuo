# Data dictionary

## Master_clean.xlsx — `master`

| Field | Meaning |
|---|---|
| Year | Observation year |
| Country | FAOSTAT reporting country after cleaning |
| Continent / Subregion | Geographic classifications |
| Crop | Standardized project crop name |
| Crop Category / Crop Subcategory | Project crop grouping |
| Staple Crop | Project staple flag |
| Flag | FAOSTAT source flag |
| Area_harvested_ha | Harvested area, hectares |
| Area_harvested_ha_DQ | DQ field for harvested area |
| Yield_kg_per_ha | Yield, kg/ha |
| Yield_kg_per_ha_DQ | DQ field for yield |
| Production_tonnes | Production, tonnes |
| Production_tonnes_DQ | DQ field for production |

## Cost_Master.xlsx — `cost_master`

| Field | Meaning |
|---|---|
| Country | Country |
| Year | Year |
| cropland_1000ha | National cropland, thousand hectares |
| cropland_flag | FAOSTAT flag |
| forest_1000ha | National forest land, thousand hectares |
| forest_flag | FAOSTAT flag |
| fertilizer_kg_per_ha | N+P+K nutrient intensity per ha of cropland |
| fertilizer_flag | Combined source flags for nutrient components |

## Derived CFI fields

| Field | Meaning |
|---|---|
| Exposure | Leading producer share of recent world output |
| Downside volatility | RMS of negative annual production changes |
| Backfill capacity | Sum of demonstrated headroom outside leading producer |
| Resilience ratio | Backfill capacity / leading-producer recent output |
| Structural shortfall | max(0, 1 − resilience ratio) |
| CFI | Final Crop Fragility Index, 0–100 |
