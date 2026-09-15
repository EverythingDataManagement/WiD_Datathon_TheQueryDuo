
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# -----------------------------
# PAGE CONFIG
# -----------------------------
st.set_page_config(
    page_title="Global Staple Crop Risk Monitor",
    page_icon="🌾",
    layout="wide",
)

# Presentation styling
st.markdown("""
<style>
.block-container {padding-top: 1.4rem; padding-bottom: 2rem; max-width: 1500px;}
[data-testid="stMetric"] {
    background: #ffffff;
    border: 1px solid #e5e7eb;
    border-radius: 14px;
    padding: 16px 18px;
    box-shadow: 0 2px 8px rgba(0,0,0,.04);
}
[data-testid="stMetricLabel"] {font-weight: 700;}
.decision-bad {
    background: #fff2f2; border: 1px solid #ffd0d0; border-radius: 14px;
    padding: 17px 21px; margin: 8px 0 16px; font-size: 1.25rem; font-weight: 800;
}
.decision-good {
    background: #f0fff4; border: 1px solid #c8efd2; border-radius: 14px;
    padding: 17px 21px; margin: 8px 0 16px; font-size: 1.25rem; font-weight: 800;
}
.small-note {color: #6b7280; font-size: .9rem;}
</style>
""", unsafe_allow_html=True)

BASE_DIR = Path(__file__).resolve().parent
MASTER_FILE = BASE_DIR / "Master_clean.xlsx"

# Working CFI parameters agreed for this project
EXPOSURE_CAP = 0.50
DOWNSIDE_CAP = 0.20
STRUCTURAL_BASE = 0.70
AMPLIFIER_SHARE = 0.30
RECENT_YEARS = (2021, 2023)
HISTORY_END = 2023

# -----------------------------
# DATA
# -----------------------------
@st.cache_data
def load_master():
    if not MASTER_FILE.exists():
        st.error(
            "Master_clean.xlsx was not found. Put it in the same folder as app.py."
        )
        st.stop()

    df = pd.read_excel(
        MASTER_FILE,
        sheet_name="master",
        usecols=[
            "Year",
            "Country",
            "Crop",
            "Production_tonnes",
        ],
    )
    df = df.dropna(subset=["Year", "Country", "Crop", "Production_tonnes"]).copy()
    df["Year"] = df["Year"].astype(int)
    df = df[df["Year"] <= HISTORY_END]
    return df


def downside_volatility(prod_series: pd.Series) -> float:
    """RMS of negative year-over-year production changes; positive changes = 0."""
    s = prod_series.sort_index()
    g = s.pct_change().dropna()
    if g.empty:
        return 0.0
    downside = np.minimum(g.to_numpy(), 0.0)
    return float(np.sqrt(np.mean(np.square(downside))))


@st.cache_data
def build_crop_metrics(df):
    rows = []
    reinforcement = {}

    for crop in sorted(df["Crop"].unique()):
        c = df[df["Crop"] == crop].copy()

        recent = (
            c[c["Year"].between(*RECENT_YEARS)]
            .groupby("Country")["Production_tonnes"]
            .mean()
            .dropna()
        )

        if recent.empty:
            continue

        top_country = recent.idxmax()
        top_output = float(recent.loc[top_country])
        world_output = float(recent.sum())
        exposure = top_output / world_output if world_output else 0.0

        historical_peak = c.groupby("Country")["Production_tonnes"].max()
        capacity = (historical_peak - recent).clip(lower=0)
        capacity = capacity.drop(labels=[top_country], errors="ignore").sort_values(
            ascending=False
        )
        capacity = capacity[capacity > 0]

        backfill = float(capacity.sum())
        buffer_ratio = backfill / top_output if top_output else 0.0
        shortfall = max(0.0, 1.0 - buffer_ratio)

        top_hist = (
            c[c["Country"] == top_country]
            .groupby("Year")["Production_tonnes"]
            .sum()
            .sort_index()
        )
        dvol = downside_volatility(top_hist)

        e_star = min(exposure / EXPOSURE_CAP, 1.0)
        d_star = min(dvol / DOWNSIDE_CAP, 1.0)
        amplifier = 0.5 * e_star + 0.5 * d_star

        cfi = 100 * shortfall * (
            STRUCTURAL_BASE + AMPLIFIER_SHARE * amplifier
        )

        if cfi >= 50:
            risk = "High fragility"
        elif cfi >= 20:
            risk = "Elevated fragility"
        elif cfi >= 5:
            risk = "Moderate fragility"
        elif cfi > 0:
            risk = "Low fragility"
        else:
            risk = "Structurally buffered"

        rows.append(
            {
                "Crop": crop,
                "Top producer": top_country,
                "Exposure": exposure,
                "Downside volatility": dvol,
                "Top output": top_output,
                "Backfill capacity": backfill,
                "Buffer ratio": buffer_ratio,
                "CFI": cfi,
                "Classification": risk,
            }
        )

        reinforcement[crop] = (
            capacity.rename("Available capacity")
            .reset_index()
            .rename(columns={"Country": "Potential reinforcement market"})
        )

    return pd.DataFrame(rows), reinforcement


master = load_master()
crop_metrics, reinforcement_by_crop = build_crop_metrics(master)

# -----------------------------
# HELPERS
# -----------------------------
def mt(x):
    return x / 1_000_000


def pct(x):
    return f"{x * 100:.1f}%"


DISPLAY_NAMES = {
    "China, mainland": "China",
    "United States of America": "United States",
    "Russian Federation": "Russia",
    "Republic of Korea": "South Korea",
}

def display_country(name):
    return DISPLAY_NAMES.get(name, name)


def risk_icon(label):
    if label.startswith("High"):
        return "🔴"
    if label.startswith("Elevated"):
        return "🟠"
    if label.startswith("Moderate"):
        return "🟡"
    if label.startswith("Low"):
        return "🟢"
    return "🔵"


# -----------------------------
# HEADER
# -----------------------------
st.title("🌾 Global Staple Crop Risk Monitor")
st.caption(
    "A production-based view of exposure, replacement capacity and fragility "
    "for 10 global staple crops."
)

tab1, tab2, tab3 = st.tabs(
    ["Risk Monitor", "Global Fragility", "Methodology"]
)

# -----------------------------
# TAB 1 — RISK MONITOR
# -----------------------------
with tab1:
    left, right = st.columns([1, 2.2], gap="large")

    with left:
        st.subheader("Select a crop")
        crop = st.selectbox(
            "Crop",
            crop_metrics.sort_values("CFI", ascending=False)["Crop"].tolist(),
            label_visibility="collapsed",
        )

        shock_pct = st.slider(
            "Simulate a leading-producer production shock",
            min_value=10,
            max_value=100,
            value=35,
            step=5,
            format="%d%%",
        )

        st.info(
            "The shock slider is a scenario tool. It does not change the CFI; "
            "it shows how the existing replacement-capacity estimate behaves "
            "at different disruption sizes."
        )

    row = crop_metrics[crop_metrics["Crop"] == crop].iloc[0]
    top_raw = row["Top producer"]
    top = display_country(top_raw)
    exposure = row["Exposure"]
    buffer_ratio = row["Buffer ratio"]
    cfi = row["CFI"]
    classification = row["Classification"]
    top_output = row["Top output"]
    backfill = row["Backfill capacity"]

    shock = top_output * shock_pct / 100
    covered = min(shock, backfill)
    gap = max(0.0, shock - backfill)
    coverage_pct = covered / shock if shock else 1.0

    with right:
        st.subheader(
            f"{risk_icon(classification)} {crop.upper()} — {classification.upper()}"
        )

        m1, m2, m3, m4 = st.columns(4)
        m1.metric(
            "EXPOSURE",
            pct(exposure),
            help=f"Share of global {crop.lower()} production produced by {top}.",
        )
        m2.metric(
            "DOWNSIDE RISK",
            pct(row["Downside volatility"]),
            help="Historical downside production volatility of the leading producer.",
        )
        m3.metric(
            "RESILIENCE",
            f"{buffer_ratio:.2f}",
            help="Modeled replacement capacity relative to leading-producer output.",
        )
        m4.metric(
            "CROP FRAGILITY INDEX",
            f"{cfi:.1f} / 100",
            help="Working CFI model. Higher values indicate greater structural fragility.",
        )

        st.markdown(
            f"**Leading producer:** {top}  \n"
            f"**Recent average output (2021–2023):** {mt(top_output):,.1f} Mt"
        )

    st.divider()

    st.subheader(f"What happens under a {shock_pct}% shock to {top}?")
    if gap > 0:
        st.markdown(
            f'<div class="decision-bad">⚠️ NOT FULLY COVERED — {mt(gap):,.1f} Mt supply gap</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            '<div class="decision-good">✓ COVERED BY MODELED CAPACITY</div>',
            unsafe_allow_html=True,
        )

    s1, s2, s3 = st.columns(3)
    s1.metric("Production disrupted", f"{mt(shock):,.1f} Mt")
    s2.metric("Modeled backup available", f"{mt(covered):,.1f} Mt")
    s3.metric(
        "Uncovered gap",
        f"{mt(gap):,.1f} Mt",
        delta="Covered" if gap <= 0 else "Supply gap",
        delta_color="normal" if gap <= 0 else "inverse",
    )

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            y=["Shock coverage"],
            x=[mt(covered)],
            name="Covered",
            orientation="h",
            text=[f"{mt(covered):.1f} Mt covered"],
            textposition="inside",
        )
    )
    fig.add_trace(
        go.Bar(
            y=["Shock coverage"],
            x=[mt(gap)],
            name="Gap",
            orientation="h",
            text=[f"{mt(gap):.1f} Mt gap" if gap else ""],
            textposition="inside",
        )
    )
    fig.update_layout(
        barmode="stack",
        height=220,
        margin=dict(l=10, r=10, t=45, b=20),
        xaxis_title="Million tonnes",
        yaxis_title="",
        legend=dict(orientation="h", y=1.25, x=0),
        title=f"Modeled coverage of a {shock_pct}% production shock",
    )
    st.plotly_chart(fig, use_container_width=True)

    if gap > 0:
        st.warning(
            f"Under this scenario, the modeled backup covers about "
            f"{coverage_pct * 100:.0f}% of the disrupted production, leaving "
            f"an uncovered gap of {mt(gap):,.1f} Mt."
        )
    else:
        st.success(
            "Under this scenario, modeled historical spare capacity is sufficient "
            "to cover the disrupted production."
        )

    st.subheader("Where could reinforcement come from?")
    st.caption(
        "These are production-capacity leads, not guaranteed suppliers. "
        "Trade access, logistics, contracts and policy restrictions require separate analysis."
    )

    reinf = reinforcement_by_crop[crop].head(5).copy()
    reinf["Display market"] = reinf["Potential reinforcement market"].map(display_country)
    reinf["Capacity Mt"] = reinf["Available capacity"].map(mt)
    reinf = reinf.sort_values("Capacity Mt")

    fig2 = go.Figure(
        go.Bar(
            x=reinf["Capacity Mt"],
            y=reinf["Display market"],
            orientation="h",
            text=[f"{x:.1f} Mt" for x in reinf["Capacity Mt"]],
            textposition="outside",
        )
    )
    fig2.update_layout(
        height=310,
        margin=dict(l=10, r=60, t=10, b=35),
        xaxis_title="Modeled available capacity (Mt)",
        yaxis_title="",
        showlegend=False,
    )
    st.plotly_chart(fig2, use_container_width=True)
    st.markdown(
        '<div class="small-note">Production-capacity leads, not guaranteed suppliers. '
        'Trade access, logistics, contracts and policy restrictions require separate analysis.</div>',
        unsafe_allow_html=True,
    )


# -----------------------------
# TAB 2 — GLOBAL FRAGILITY
# -----------------------------
with tab2:
    st.subheader("Fragility across the 10 staples")
    ranked = crop_metrics.sort_values("CFI", ascending=True).copy()

    fig = px.bar(
        ranked,
        x="CFI",
        y="Crop",
        orientation="h",
        text="CFI",
        hover_data={
            "Top producer": True,
            "Exposure": ":.1%",
            "Buffer ratio": ":.2f",
            "Downside volatility": ":.1%",
            "Classification": True,
            "CFI": ":.1f",
        },
        title="Crop Fragility Index",
    )
    fig.update_traces(texttemplate="%{text:.1f}", textposition="outside")
    fig.update_layout(
        height=520,
        xaxis_title="CFI score",
        yaxis_title="",
        margin=dict(l=10, r=40, t=50, b=10),
    )
    st.plotly_chart(fig, use_container_width=True)

    table = crop_metrics.sort_values("CFI", ascending=False).copy()
    table["Exposure (%)"] = table["Exposure"] * 100
    table["Downside volatility (%)"] = table["Downside volatility"] * 100
    table = table[
        [
            "Crop",
            "Top producer",
            "Exposure (%)",
            "Downside volatility (%)",
            "Buffer ratio",
            "CFI",
            "Classification",
        ]
    ]
    st.dataframe(
        table,
        hide_index=True,
        use_container_width=True,
        column_config={
            "Exposure (%)": st.column_config.NumberColumn(format="%.1f"),
            "Downside volatility (%)": st.column_config.NumberColumn(format="%.2f"),
            "Buffer ratio": st.column_config.NumberColumn(format="%.2f"),
            "CFI": st.column_config.NumberColumn(format="%.1f"),
        },
    )

# -----------------------------
# TAB 3 — METHODOLOGY
# -----------------------------
with tab3:
    st.subheader("What the model measures")

    st.markdown(
        """
**Exposure** — How much of global production depends on the leading producer.

**Downside volatility** — How strongly that producer's historical production has tended
to move downward. Positive production changes are set to zero.

**Resilience ratio** — How much modeled historical spare production capacity exists
outside the leading producer relative to the leading producer's output.

**Crop Fragility Index** — Replacement capacity is the structural anchor; exposure
and downside volatility act as amplifiers.
"""
    )

    st.code(
        """S = max(0, 1 - Buffer Ratio)

E* = min(Exposure / 50%, 1)
D* = min(Downside Volatility / 20%, 1)

H = 0.5(E*) + 0.5(D*)

CFI = 100 × S × [0.70 + 0.30(H)]""",
        language="text",
    )

    st.subheader("Key assumptions")
    st.markdown(
        f"""
- Recent production baseline: **{RECENT_YEARS[0]}–{RECENT_YEARS[1]} average**
- Historical capacity window: **2000–{HISTORY_END}**
- Replacement capacity uses each non-leading producer's historical maximum
  minus its recent baseline, floored at zero.
- A resilience ratio of **1.0** means modeled alternative capacity could replace
  100% of the leading producer's recent output.
- A CFI of **0** means **structurally buffered under the model**, not zero real-world risk.
- The shock slider is an intuitive scenario test, **not independent validation** of the CFI.
- The model is production-based. It does **not** yet model trade flows, logistics,
  inventory, policy, or commercial availability.
"""
    )

    st.subheader("Working classification")
    bands = pd.DataFrame(
        {
            "CFI": ["50+", "20–49.9", "5–19.9", ">0–4.9", "0"],
            "Interpretation": [
                "High fragility",
                "Elevated fragility",
                "Moderate fragility",
                "Low fragility",
                "Structurally buffered",
            ],
        }
    )
    st.dataframe(bands, hide_index=True, use_container_width=True)

    st.caption(
        "Source: FAOSTAT production data in Master_clean.xlsx. "
        "Working methodology developed for this analysis."
    )
