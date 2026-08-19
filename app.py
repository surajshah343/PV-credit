"""
Private Credit Quarter-End Valuation App (ASC 820, yield method) — DUMMY DATA.
Users upload the three workbooks; bundled files in /data are used as a fallback if present.
Run:  streamlit run app.py
"""
import datetime as dt
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import valuation as v

st.set_page_config(page_title="PC Fund Valuation — Q-End Model", layout="wide", page_icon="📊")

st.title("Private Credit Fund — Quarter-End Valuation - by Suraj Shah")
st.caption("Yield-method DCF with recovery overlay for watchlist credits. "
           "**IN TESTING MODE- All data are for illustration only.**")

DATA_DIR = Path(__file__).parent / "data"


# ------------------------------------------------------------------- file intake
@st.cache_data(show_spinner=False)
def parse_market(b: bytes) -> dict:
    return v.load_market_data(b)


@st.cache_data(show_spinner=False)
def parse_model(b: bytes) -> dict:
    return v.load_model(b)


def _source(uploaded, fallback_name: str):
    """Uploaded file bytes, else bundled repo file if it exists, else None."""
    if uploaded is not None:
        return uploaded.getvalue(), f"uploaded · {uploaded.name}"
    p = DATA_DIR / fallback_name
    if p.exists():
        return p.read_bytes(), f"bundled · data/{fallback_name}"
    return None, None


with st.sidebar:
    st.header("1 · Upload workbooks")
    up_md = st.file_uploader("Market_Data_Inputs.xlsx", type="xlsx", key="md")
    up_q1 = st.file_uploader("PC_Valuation_Q1_2026.xlsx", type="xlsx", key="q1")
    up_q2 = st.file_uploader("PC_Valuation_Q2_2026.xlsx", type="xlsx", key="q2")

md_bytes, md_src = _source(up_md, "Market_Data_Inputs.xlsx")
q1_bytes, q1_src = _source(up_q1, "PC_Valuation_Q1_2026.xlsx")
q2_bytes, q2_src = _source(up_q2, "PC_Valuation_Q2_2026.xlsx")

if md_bytes is None or (q1_bytes is None and q2_bytes is None):
    st.info("⬅️ **Upload the workbooks in the sidebar to begin.**\n\n"
            "Required: `Market_Data_Inputs.xlsx` plus at least one quarterly model "
            "(`PC_Valuation_Q1_2026.xlsx` / `PC_Valuation_Q2_2026.xlsx`). "
            "Upload both quarterly models to unlock the Q-over-Q bridge.")
    st.stop()

try:
    md_all = parse_market(md_bytes)
    models = {}
    if q1_bytes:
        models["Q1"] = parse_model(q1_bytes)
    if q2_bytes:
        models["Q2"] = parse_model(q2_bytes)
except ValueError as e:
    st.error(f"Couldn't read a workbook: {e}")
    st.stop()
except Exception as e:  # corrupt zip, wrong file type renamed to .xlsx, etc.
    st.error(f"Couldn't open a workbook — is it a valid .xlsx file? Details: {e}")
    st.stop()

with st.sidebar:
    st.caption(f"Market data: {md_src}")
    if "Q1" in models: st.caption(f"Q1 model: {q1_src}")
    if "Q2" in models: st.caption(f"Q2 model: {q2_src}")

    st.header("2 · Valuation controls")
    quarters = list(models.keys())
    quarter = st.radio("Portfolio quarter", quarters, horizontal=True,
                       index=len(quarters) - 1, format_func=lambda q: f"{q} 2026")
    model = models[quarter]
    default_vd = model["valdate"] or v.QUARTER_END[quarter]
    valdate = st.date_input("Valuation date", value=default_vd,
                            min_value=dt.date(2026, 1, 1), max_value=dt.date(2027, 12, 31),
                            help="Drives the quarterly cash-flow projection horizon per position. "
                                 "Default comes from the model's Summary tab.")
    vintage = st.selectbox("Market data vintage", ["Match portfolio quarter", "Q1 2026", "Q2 2026"],
                           help="Which quarter-end spot SOFR, forward curve and spread matrix to pull "
                                "from Market_Data_Inputs.xlsx.")
    vintage_key = quarter if vintage == "Match portfolio quarter" else vintage.split()[0]

    st.divider()
    pik_on = st.toggle("Capitalize PIK", value=True,
                       help="ON: PIK accrues to the balance and repays at maturity (base case). "
                            "OFF: PIK component modeled as cash-pay — isolates the PIK effect on value.")
    incl_illiq = st.toggle("Include illiquidity premium", value=True)
    shock = st.slider("Parallel credit spread shock (bps)", -200, 300, 0, 25,
                      help="Sensitivity: shifts every position's discount spread.")

settings = dict(pik_capitalize=pik_on, spread_shock_bps=float(shock), include_illiquidity=incl_illiq)

md = md_all[vintage_key]
cur = v.value_portfolio(model, md, valdate=valdate, **settings)
prior = None
if quarter == "Q2" and "Q1" in models:
    m1 = models["Q1"]
    prior = v.value_portfolio(m1, md_all["Q1"], valdate=m1["valdate"] or v.QUARTER_END["Q1"], **settings)

# -------------------------------------------------------------------- metrics
c1, c2, c3, c4, c5 = st.columns(5)
tot_fv, tot_funded = cur["Fair Value"].sum(), cur["Funded"].sum()
wavg_yield = (cur["Market Yield"] * cur["Funded"]).sum() / tot_funded
bsl = list(md["bench"].values())[0]
prior_fv = prior["Fair Value"].sum() if prior is not None else None
c1.metric("Total Fair Value ($000)", f"{tot_fv:,.0f}",
          delta=f"{tot_fv - prior_fv:,.0f} vs Q1" if prior_fv else None)
c2.metric("FV % of Funded Par", f"{tot_fv / tot_funded:.2%}")
c3.metric("Wtd Avg Market Yield", f"{wavg_yield:.2%}")
c4.metric("BSL Index (sanity check)", f"{bsl:.2%}", delta=f"+{(wavg_yield - bsl) * 1e4:,.0f} bps premium",
          delta_color="off")
c5.metric("Positions", f"{len(cur)}")

tab_sum, tab_bridge, tab_pos, tab_drill, tab_mkt = st.tabs(
    ["📈 Summary", "🌉 Q-over-Q Bridge", "📋 Positions", "🔍 Position Drill-Down", "🗂 Market Data"])

# -------------------------------------------------------------------- summary
with tab_sum:
    left, right = st.columns(2)
    with left:
        st.subheader("Fair value by seniority")
        by_sen = cur.groupby("Seniority", as_index=False)["Fair Value"].sum()
        st.plotly_chart(go.Figure(go.Pie(labels=by_sen["Seniority"], values=by_sen["Fair Value"], hole=0.5)),
                        width="stretch")
    with right:
        st.subheader("Price (% of par) by position")
        d = cur.sort_values("FV % of Funded")
        fig = go.Figure(go.Bar(x=d["FV % of Funded"] * 100, y=d["Borrower"], orientation="h",
                               marker_color=["#d62728" if w > 0 else "#1f77b4" for w in d["Recovery Weight"]]))
        fig.add_vline(x=100, line_dash="dash", line_color="grey")
        fig.update_layout(xaxis_title="FV % of funded par", height=420, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig, width="stretch")
    st.info("Red bars = watchlist credits valued with a probability-weighted recovery overlay "
            "(see the Recovery tab of the Excel model).")

# --------------------------------------------------------------------- bridge
with tab_bridge:
    if quarter != "Q2":
        st.info("Select **Q2 2026** to see the Q1 → Q2 bridge. "
                "(The Q1 workbook bridges against dummy 12/31/2025 marks — see its Bridge tab.)")
    else:
        br = v.qoq_bridge(prior, cur, model["prior_marks"])
        drivers = ["Repayments", "PIK Capitalized", "New Originations",
                   "Credit-Specific", "Market Re-mark & Accretion"]
        totals = {dr: br[dr].sum() for dr in drivers}
        start, end = br["Prior FV"].sum(), br["Current FV"].sum()
        src_note = ("prior FVs recomputed from the uploaded Q1 model with the same toggles"
                    if prior is not None else
                    "prior FVs taken from the Q2 workbook's Bridge tab (upload the Q1 model for "
                    "toggle-consistent attribution)")
        st.subheader("Q1 2026 → Q2 2026 fair value bridge ($000)")
        fig = go.Figure(go.Waterfall(
            orientation="v",
            measure=["absolute"] + ["relative"] * len(drivers) + ["total"],
            x=["Q1 2026 FV"] + drivers + ["Q2 2026 FV"],
            y=[start] + [totals[dr] for dr in drivers] + [end],
            text=[f"{val:,.0f}" for val in [start] + [totals[dr] for dr in drivers] + [end]],
            textposition="outside",
            connector={"line": {"color": "grey"}},
        ))
        fig.update_layout(height=460, margin=dict(l=10, r=10, t=30, b=10),
                          yaxis_title="Fair value ($000)")
        st.plotly_chart(fig, width="stretch")
        st.caption(f"Credit-specific captures watchlist / rating-migration names; market re-mark & "
                   f"accretion is the residual for performing names — mirroring the Excel Bridge tab. "
                   f"({src_note}.)")
        with st.expander("Position-level bridge detail"):
            st.dataframe(br.style.format({c: "{:,.1f}" for c in br.columns if c != "Borrower"}),
                         width="stretch", hide_index=True)

# ------------------------------------------------------------------ positions
with tab_pos:
    st.subheader(f"Position-level valuation — {quarter} 2026 (valuation date {valdate:%m/%d/%Y})")
    fmt = {"Commitment": "{:,.0f}", "Funded": "{:,.1f}", "DCF FV": "{:,.1f}", "Recovery FV": "{:,.1f}",
           "Fair Value": "{:,.1f}", "FV % of Funded": "{:.2%}", "Market Yield": "{:.2%}",
           "Base": "{:.2%}", "Floor": "{:.2%}", "Recovery Weight": "{:.0%}",
           "Cash Spread (bps)": "{:,.0f}", "PIK (bps)": "{:,.0f}", "Credit Spread (bps)": "{:,.0f}",
           "Illiquidity (bps)": "{:,.0f}", "Credit Adj (bps)": "{:,.0f}"}
    st.dataframe(cur.style.format(fmt), width="stretch", hide_index=True, height=430)
    st.caption("Market yield = quarter-end spot SOFR + credit spread (rating × seniority matrix) "
               "+ illiquidity premium (tranche) + credit-specific adjustment + any user shock.")

# ------------------------------------------------------------------ drilldown
with tab_drill:
    name = st.selectbox("Position", cur["Borrower"].tolist())
    row = cur[cur["Borrower"] == name].iloc[0]
    p = next(x for x in model["positions"] if x.name == name)
    yb = v.market_yield(p, md, float(shock), incl_illiq)
    cfs = v.project_cashflows(p, valdate, md, yb["yield"], pik_on)

    a, b, c = st.columns(3)
    a.metric("Fair Value ($000)", f"{row['Fair Value']:,.1f}")
    b.metric("FV % of Funded", f"{row['FV % of Funded']:.2%}")
    c.metric("Market Yield", f"{row['Market Yield']:.2%}")

    st.markdown("**Discount yield build**")
    yb_df = pd.DataFrame({
        "Component": ["Spot SOFR", "Credit spread", "Illiquidity premium", "Credit adjustment",
                      "User spread shock", "Total market yield"],
        "Value": [f"{yb['base']:.2%}", f"{yb['credit_spread_bps']:,.0f} bps", f"{yb['illiq_bps']:,.0f} bps",
                  f"{yb['credit_adj_bps']:,.0f} bps", f"{yb['shock_bps']:,.0f} bps", f"{yb['yield']:.2%}"]})
    st.dataframe(yb_df, hide_index=True)

    if row["Recovery Weight"] > 0:
        rp = model["recovery"]
        st.warning(f"Watchlist credit — fair value blends {1 - row['Recovery Weight']:.0%} yield-method DCF "
                   f"with {row['Recovery Weight']:.0%} recovery value "
                   f"(EV waterfall: {rp['ebitda']:,.0f} EBITDA × {rp['multiple']:.2f}x, "
                   f"PV'd at {rp['distressed_rate']:.0%} over {rp['resolution_yrs']:.2f} yrs → "
                   f"{v.recovery_value_pct(rp):.1%} of par).")

    st.markdown("**Projected contractual cash flows (quarterly)**")
    fig = go.Figure()
    fig.add_bar(x=cfs["n"], y=cfs["cash_interest"], name="Cash interest")
    fig.add_bar(x=cfs["n"], y=cfs["principal"], name="Principal (incl. capitalized PIK)")
    fig.add_scatter(x=cfs["n"], y=cfs["pv"], name="PV of CF", mode="lines+markers", yaxis="y")
    fig.update_layout(barmode="stack", height=380, xaxis_title="Quarter",
                      yaxis_title="$000", margin=dict(l=10, r=10, t=20, b=10))
    st.plotly_chart(fig, width="stretch")
    with st.expander("Cash flow schedule"):
        st.dataframe(cfs.style.format({c: "{:,.1f}" for c in
                                       ["beg_balance", "pik_accrual", "cash_interest", "principal",
                                        "total_cf", "pv"]} | {"fwd_sofr": "{:.2%}", "base": "{:.2%}",
                                                              "df": "{:.4f}"}),
                     width="stretch", hide_index=True)

# ---------------------------------------------------------------- market data
with tab_mkt:
    l, r = st.columns(2)
    with l:
        st.subheader("3M SOFR forward curves")
        fig = go.Figure()
        for vk in ("Q1", "Q2"):
            fig.add_scatter(x=list(range(1, 21)), y=[f * 100 for f in md_all[vk]["fwd"]],
                            name=f"{vk} 2026 vintage", mode="lines+markers")
        fig.update_layout(height=360, xaxis_title="Forward period (quarters)", yaxis_title="%")
        st.plotly_chart(fig, width="stretch")
    with r:
        st.subheader(f"Credit spread matrix — {vintage_key} 2026 (bps)")
        mat = pd.DataFrame(md["spreads"], index=[1, 2, 3, 4, 5]).T
        mat.columns = [f"Rating {i}" for i in mat.columns]
        st.dataframe(mat, width="stretch")
        st.subheader("Illiquidity premia / benchmarks")
        st.dataframe(pd.Series(md["illiq"], name="bps").to_frame(), width="stretch")
        st.dataframe(pd.Series({k: f"{x:.2%}" for k, x in md["bench"].items()}, name="Yield").to_frame(),
                     width="stretch")

st.divider()
st.caption("Method: ASC 820 Level 3, income approach (discounted cash flow / yield method), with a "
           "market-participant recovery overlay for watchlist credits. Unfunded commitments carried at zero. "
           "Dummy data throughout — not investment advice.")
