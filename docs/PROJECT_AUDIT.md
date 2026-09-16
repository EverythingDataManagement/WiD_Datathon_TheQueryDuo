# Final project audit

## What is complete

- Production master and deterministic cleaning pipeline
- 10-crop scope and 2000–2023 analytical window
- Growth and concentration analysis
- Land-versus-yield growth decomposition
- Cropland/forest cost evidence
- Fertilizer/yield input-cost evidence
- Resilience ratio and shock scenarios
- Final Crop Fragility Index using downside volatility and fixed normalization ceilings
- Sensitivity analysis
- Streamlit dashboard
- GitHub-ready documentation and reproducibility structure

## Supporting / optional

- Bilateral trade exposure: useful illustration, not a CFI input
- Water stress: tested and excluded from the final model

## Roadmap corrections required

Several roadmap statements reflect earlier versions and should not be used as final claims:

- Use **~64% / nearly two-thirds** for the top four contributors under the final all-10-crop calculation.
- Use the final country-level cropland/forest result **r ≈ -0.36 across 190 countries**, not older -0.39 / 215-country wording.
- Use **six of ten** staples below a resilience ratio of 1 under the final 2021–2023 baseline, not the older seven-of-ten wording.
- The final CFI uses **downside volatility**, not ordinary standard deviation.
- The final CFI uses **fixed ceilings (50% exposure, 20% downside)**, not sample min-max normalization.
- The final CFI uses **50/50 exposure/downside inside H** and a **70/30 structural/amplifier design**.
- The 35% shock is a **scenario/reality check**, not independent validation of the index.
- Trade is **supporting/optional**. Do not say it was entirely scoped out, but also do not imply it is inside the CFI.
- Do not use stale demo examples that list the top producer among its own backup markets.
- Do not claim the raw trade XLSX is safely reproducible from GitHub; it exceeds normal GitHub file limits and may be truncated.

- **Fertilizer/yield reconciliation resolved:** the older roadmap **+0.51** is reproducible as **r=+0.5034, n=170, p≈2.63e-12** using the earlier five-crop primary scope (Rice, Sorghum, Soybeans, Wheat, Potatoes). The final canonical six-cereal method gives **r=+0.5822, n=173, p≈4.40e-17**. Use **+0.58** for the final presentation and document +0.51 as the superseded scope/method.
- **Sample-size reconciliation:** land/forest n=190 and fertilizer/yield n=173 are different complete-case samples with different eligibility rules; they are not expected to match.
