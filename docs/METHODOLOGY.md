# Methodology

## Analytical question

How has global staple-crop production grown since 2000, what pressures are associated
with the two main growth paths, and how fragile is supply if a leading producer is disrupted?

The final project covers 10 crops: maize, rice, wheat, soybeans, potatoes, sugarcane,
cassava, sorghum, millet and barley.

## 1. Production spine

FAOSTAT crop production is the spine of the analysis. The clean master is one
Country × Crop × Year record with harvested area, yield and production. Regional
aggregates and obsolete/duplicate country units are removed before analysis.

The presentation window is 2000–2023. The workbook can contain 2024 observations,
but 2023 is used as the final comparison year to avoid mixing provisional/latest-year
coverage into the headline analysis.

## 2. Growth concentration

Country contribution to growth is measured as the change in aggregate production
between 2000 and 2023. The final all-10-crop calculation shows that Brazil, India,
China and the United States account for roughly 64% of positive country-level growth.

"Top by growth" is not the same as "top producer of a crop." Resilience always uses
the true leading producer for each crop.

## 3. Land-led versus yield-led growth

For aggregate country production:

Production = Area × Yield.

The final decomposition uses the symmetric two-factor identity:

Land contribution = (A1 − A0) × (Y0 + Y1)/2

Yield contribution = (Y1 − Y0) × (A0 + A1)/2

after unit conversion. The two contributions sum exactly to the production change.

Presentation classification:
- Land-led: land explains at least 60% of positive growth, or yield contribution is non-positive.
- Yield-led: yield explains at least 60%, or land contribution is non-positive.
- Mixed: each explains 40–60%.

This classification threshold is a communication convention, not a statistically
estimated boundary.

## 4. Cost / pressure layer

### Land path → forest pressure

FAOSTAT Land Use is national, not crop-level. The project compares country-level
change in cropland with country-level change in forest land. Pearson correlation is
used as an association measure. The final presentation result is approximately
r = -0.36 across 190 countries: countries expanding cropland more tended to lose
more forest.

This does not establish causality. National forest totals also net planting against
clearing and do not isolate primary-forest loss. Crop attribution is inferred from
the production master, not directly measured in the land-use source.

### Yield path → fertilizer/input pressure

FAOSTAT Fertilizers by Nutrient provides national N+P+K intensity per hectare of
cropland. The final canonical method compares 2023 fertilizer intensity with national
aggregate cereal yield across six cereals: Maize, Rice, Wheat, Sorghum, Millet and Barley.
It returns r = +0.5822 (≈ +0.58), n=173, p≈4.40e-17.

The earlier roadmap result of about +0.51 was traced and reproduced as r=+0.5034,
n=170, p≈2.63e-12 when the earlier five-crop primary scope (Rice, Sorghum, Soybeans,
Wheat and Potatoes) is used. The move from +0.51 to +0.58 is therefore a documented
scope/method revision, not a source-data correction.

Again, this is association rather than causal attribution. Fertilizer is national
while crop production is crop-level, so the grain mismatch is an explicit limitation.

## 5. Resilience ratio

The final application uses a 2021–2023 average as the recent production baseline.

For each crop:
1. Identify the leading producer using recent average production.
2. For every other producer, calculate demonstrated headroom:
   max(historical peak production − recent baseline production, 0).
3. Sum that headroom.
4. Divide it by the leading producer's recent output.

Resilience ratio = total alternative demonstrated capacity / leading-producer output.

A ratio of 1.0 means modeled alternative capacity equals the full recent output of
the leading producer. A value below 1.0 indicates a structural replacement gap.

This is tonnage-based capacity, not guaranteed exportable or commercially available supply.

## 6. Downside volatility

Only production declines contribute to the hazard measure.

For annual growth g_t:

Downside volatility = sqrt(mean(min(g_t, 0)^2)).

Positive production changes are set to zero. This distinguishes downside production
risk from ordinary two-sided volatility.

## 7. Crop Fragility Index (CFI)

Three concepts are retained:
- Exposure: share of global production supplied by the leading producer.
- Downside risk: historical downside volatility of that producer.
- Buffer/resilience: modeled ability of all other producers to replace it.

Structural shortfall:

S = max(0, 1 − Resilience Ratio)

Fixed normalization ceilings:

E* = min(Exposure / 50%, 1)
D* = min(Downside Volatility / 20%, 1)

Amplifier:

H = 0.50 E* + 0.50 D*

Final score:

CFI = 100 × S × [0.70 + 0.30 H]

The structural replacement gap is therefore the anchor. Exposure and downside risk
can amplify that structural vulnerability by up to 30%.

Working interpretation:
- 50+ High
- 20–49.9 Elevated
- 5–19.9 Moderate
- >0–4.9 Low
- 0 Structurally buffered

CFI = 0 does not mean risk-free. It means the crop is structurally buffered under
this production-capacity model.

## 8. Shock scenarios and sensitivity

The dashboard shock slider is a scenario tool. A 35% shock makes the resilience
ratio intuitive; it is not independent validation because coverage is mechanically
determined by the same replacement-capacity estimate.

Model sensitivity is tested across 180 parameter combinations varying:
- structural weight alpha: 0.5, 0.6, 0.7, 0.8
- exposure weight inside H: 0.3–0.7
- exposure cap: 40%, 50%, 60%
- downside cap: 15%, 20%, 25%

The main high-fragility crops remain near the top across these settings. The largest
modeling choice is the treatment of replacement capacity, not small changes in weights.

## 9. Trade information

Trade is a supporting illustration, not an input to the CFI. The Detailed Trade
Matrix can identify importers that depend heavily on a fragile producer when the
top producer is also an important exporter.

This framing does not work consistently for crops such as wheat, where a leading
producer may consume most output domestically. The final project therefore remains
producer-based. Trade is documented as an extension/supporting illustration rather
than mixed into the core fragility score.

## 10. Scope exclusions

Water stress was tested and excluded from the final cost model because national-level
water stress did not provide a clean mechanism for differentiating the major producers.
A useful water extension would require sub-national/basin-level irrigation pressure.

Future extensions include safe-to-export capacity, inventories, logistics, trade
restrictions, commercial availability and additional crops.
