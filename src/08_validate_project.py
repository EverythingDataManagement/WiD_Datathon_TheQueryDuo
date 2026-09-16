#!/usr/bin/env python3
from pathlib import Path
import sys
ROOT=Path(__file__).resolve().parents[1]
required=[
 ROOT/"app.py",
 ROOT/"Master_clean.xlsx",
 ROOT/"data/processed/Cost_Master.xlsx",
 ROOT/"src/03_analyze_growth_costs.py",
 ROOT/"src/04_build_cfi_final.py",
 ROOT/"docs/METHODOLOGY.md",
]
missing=[str(p.relative_to(ROOT)) for p in required if not p.exists()]
if missing:
    print("Missing required project files:")
    print("\n".join(" - "+x for x in missing))
    sys.exit(1)
print("Core project package is complete.")
