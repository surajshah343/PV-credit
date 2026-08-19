"""
Valuation engine for the private credit quarter-end app (DUMMY DATA).

Reads the three linked workbooks:
  data/Market_Data_Inputs.xlsx   - central market assumptions (Q1 + Q2 2026 vintages)
  data/PC_Valuation_Q1_2026.xlsx - Q1 portfolio (positions, recovery params, prior marks)
  data/PC_Valuation_Q2_2026.xlsx - Q2 portfolio

and mirrors the Excel yield-method DCF so the app and the workbooks agree:
  market yield = spot SOFR + credit spread(rating, seniority) + illiquidity premium + credit adj
  quarterly contractual CFs off the SOFR forward curve (with floors); PIK capitalizes; bullet maturity.
"""
from __future__ import annotations
import datetime as dt
import math
from dataclasses import dataclass, replace
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

DATA_DIR = Path(__file__).parent / "data"
QUARTER_END = {"Q1": dt.date(2026, 3, 31), "Q2": dt.date(2026, 6, 30)}
MODEL_FILE = {"Q1": "PC_Valuation_Q1_2026.xlsx", "Q2": "PC_Valuation_Q2_2026.xlsx"}


# ----------------------------------------------------------------- market data
def load_market_data(path: Path | None = None) -> dict:
    """Parse Market_Data_Inputs.xlsx into a dict keyed by vintage ('Q1'/'Q2')."""
    path = path or DATA_DIR / "Market_Data_Inputs.xlsx"
    ws = load_workbook(path, data_only=True)["Market_Data"]
    out = {}
    seniorities = [ws[f"A{32+i}"].value for i in range(4)]
    for j, v in enumerate(("Q1", "Q2")):
        col = "BC"[j]
        spot = ws[f"{col}3"].value
        fwd = [ws[f"{col}{8+i}"].value for i in range(20)]
        base = 30 + j * 8
        matrix = {ws[f"A{base+2+si}"].value: [ws.cell(row=base + 2 + si, column=2 + k).value for k in range(5)]
                  for si in range(4)}
        illiq = {ws[f"A{47+si}"].value: ws[f"B{47+si}"].value for si in range(4)}
        bench = {ws[f"A{55+bi}"].value: ws[f"{col}{55+bi}"].value for bi in range(4)}
        out[v] = {"spot": spot, "fwd": fwd, "spreads": matrix, "illiq": illiq,
                  "bench": bench, "seniorities": seniorities}
    return out


# ------------------------------------------------------------------- positions
@dataclass
class Position:
    name: str
    industry: str
    seniority: str
    rating: int
    origination: dt.date
    maturity: dt.date
    commitment: float
    funded: float
    cash_spread_bps: float
    pik_bps: float
    floor: float
    oid: float
    credit_adj_bps: float
    recovery_weight: float


def _d(v):
    return v.date() if isinstance(v, dt.datetime) else v


def load_positions(quarter: str) -> list[Position]:
    ws = load_workbook(DATA_DIR / MODEL_FILE[quarter], data_only=True)["Positions"]
    rows = []
    r = 5
    while ws[f"B{r}"].value and ws[f"B{r}"].value != "TOTAL / WTD AVG":
        rows.append(Position(
            name=ws[f"B{r}"].value, industry=ws[f"C{r}"].value, seniority=ws[f"D{r}"].value,
            rating=int(ws[f"E{r}"].value), origination=_d(ws[f"F{r}"].value), maturity=_d(ws[f"G{r}"].value),
            commitment=float(ws[f"H{r}"].value), funded=float(ws[f"I{r}"].value),
            cash_spread_bps=float(ws[f"J{r}"].value), pik_bps=float(ws[f"K{r}"].value),
            floor=float(ws[f"L{r}"].value), oid=float(ws[f"M{r}"].value),
            credit_adj_bps=float(ws[f"R{r}"].value), recovery_weight=float(ws[f"U{r}"].value),
        ))
        r += 1
    return rows


def load_recovery_params(quarter: str) -> dict:
    ws = load_workbook(DATA_DIR / MODEL_FILE[quarter], data_only=True)["Recovery"]
    return {"ebitda": ws["B4"].value, "multiple": ws["B5"].value, "revolver": ws["B7"].value,
            "fl_claims": ws["B9"].value, "distressed_rate": ws["B13"].value, "resolution_yrs": ws["B12"].value}


def load_prior_marks(quarter: str) -> pd.DataFrame:
    """Bridge tab: prior-quarter FV / funded / par repaid per borrower."""
    ws = load_workbook(DATA_DIR / MODEL_FILE[quarter], data_only=True)["Bridge"]
    rows, r = [], 4
    while ws[f"A{r}"].value and ws[f"A{r}"].value != "TOTAL":
        rows.append({"name": ws[f"A{r}"].value, "prior_fv": ws[f"B{r}"].value or 0.0,
                     "prior_funded": ws[f"C{r}"].value or 0.0, "par_repaid": ws[f"D{r}"].value or 0.0,
                     "credit_event": ws[f"K{r}"].value == "Y"})
        r += 1
    return pd.DataFrame(rows).set_index("name")


# ---------------------------------------------------------------------- engine
def n_periods(valdate: dt.date, maturity: dt.date) -> int:
    return max(1, math.ceil((maturity - valdate).days / 365 * 4))


def market_yield(p: Position, md: dict, spread_shock_bps: float = 0.0,
                 include_illiquidity: bool = True) -> dict:
    spread = md["spreads"][p.seniority][p.rating - 1]
    illiq = md["illiq"][p.seniority] if include_illiquidity else 0.0
    total = md["spot"] + (spread + illiq + p.credit_adj_bps + spread_shock_bps) / 10000
    return {"base": md["spot"], "credit_spread_bps": spread, "illiq_bps": illiq,
            "credit_adj_bps": p.credit_adj_bps, "shock_bps": spread_shock_bps, "yield": total}


def project_cashflows(p: Position, valdate: dt.date, md: dict, y: float,
                      pik_capitalize: bool = True) -> pd.DataFrame:
    """Quarterly contractual CFs. pik_capitalize=False treats the PIK coupon as cash-pay."""
    N = n_periods(valdate, p.maturity)
    bal = p.funded
    recs = []
    for n in range(1, N + 1):
        fwd = md["fwd"][min(n, 20) - 1]
        base = max(fwd, p.floor)
        cash_rate = base + p.cash_spread_bps / 10000
        pik_rate = p.pik_bps / 10000
        if pik_capitalize:
            pik_acc = bal * pik_rate / 4
            cash_int = bal * cash_rate / 4
        else:
            pik_acc = 0.0
            cash_int = bal * (cash_rate + pik_rate) / 4
        end_bal = bal + pik_acc
        principal = end_bal if n == N else 0.0
        cf = cash_int + principal
        df = 1 / (1 + y / 4) ** n
        recs.append({"n": n, "period_end": valdate + dt.timedelta(days=round(91.25 * n)),
                     "fwd_sofr": fwd, "base": base, "beg_balance": bal, "pik_accrual": pik_acc,
                     "cash_interest": cash_int, "principal": principal, "total_cf": cf,
                     "df": df, "pv": cf * df})
        bal = end_bal
    return pd.DataFrame(recs)


def recovery_value_pct(rp: dict) -> float:
    ev = rp["ebitda"] * rp["multiple"]
    to_fl = max(0.0, ev - rp["revolver"])
    gross = min(1.0, to_fl / rp["fl_claims"])
    return gross / (1 + rp["distressed_rate"]) ** rp["resolution_yrs"]


def value_portfolio(quarter: str, md_all: dict, valdate: dt.date | None = None,
                    vintage: str | None = None, pik_capitalize: bool = True,
                    spread_shock_bps: float = 0.0, include_illiquidity: bool = True) -> pd.DataFrame:
    valdate = valdate or QUARTER_END[quarter]
    md = md_all[vintage or quarter]
    positions = load_positions(quarter)
    rp = load_recovery_params(quarter)
    rec_pct = recovery_value_pct(rp)
    rows = []
    for p in positions:
        yb = market_yield(p, md, spread_shock_bps, include_illiquidity)
        cfs = project_cashflows(p, valdate, md, yb["yield"], pik_capitalize)
        dcf_fv = cfs["pv"].sum()
        rec_fv = rec_pct * p.funded if p.recovery_weight > 0 else 0.0
        final = (1 - p.recovery_weight) * dcf_fv + p.recovery_weight * rec_fv
        rows.append({"Borrower": p.name, "Industry": p.industry, "Seniority": p.seniority,
                     "Rating": p.rating, "Maturity": p.maturity, "Commitment": p.commitment,
                     "Funded": p.funded, "Cash Spread (bps)": p.cash_spread_bps,
                     "PIK (bps)": p.pik_bps, "Floor": p.floor,
                     "Base": yb["base"], "Credit Spread (bps)": yb["credit_spread_bps"],
                     "Illiquidity (bps)": yb["illiq_bps"], "Credit Adj (bps)": yb["credit_adj_bps"],
                     "Market Yield": yb["yield"], "DCF FV": dcf_fv,
                     "Recovery Weight": p.recovery_weight, "Recovery FV": rec_fv,
                     "Fair Value": final, "FV % of Funded": final / p.funded})
    return pd.DataFrame(rows)


def qoq_bridge(prior_df: pd.DataFrame, current_df: pd.DataFrame, current_quarter: str) -> pd.DataFrame:
    """Position-level bridge mirroring the Excel Bridge tab.
    Drivers: repayments, PIK capitalized, new originations, credit-specific, market re-mark & accretion."""
    marks = load_prior_marks(current_quarter)
    prior_fv_engine = prior_df.set_index("Borrower")["Fair Value"] if prior_df is not None else None
    rows = []
    for _, r in current_df.iterrows():
        nm = r["Borrower"]
        m = marks.loc[nm] if nm in marks.index else None
        # prefer the engine's prior FV (consistent with user toggles); fall back to workbook marks
        key = nm.replace(" (NEW)", "")
        if prior_fv_engine is not None and key in prior_fv_engine.index and "(NEW)" not in nm:
            prior_fv = float(prior_fv_engine[key])
        else:
            prior_fv = float(m["prior_fv"]) if m is not None else 0.0
        prior_funded = float(m["prior_funded"]) if m is not None else 0.0
        repaid = float(m["par_repaid"]) if m is not None else 0.0
        credit_event = bool(m["credit_event"]) if m is not None else False
        cur_fv, funded, pik = r["Fair Value"], r["Funded"], r["PIK (bps)"]
        is_new = prior_fv == 0 and prior_funded == 0
        repay_impact = -repaid * (prior_fv / prior_funded) if prior_funded else 0.0
        pik_cap = (funded - (prior_funded or funded)) * (cur_fv / funded) if pik > 0 else 0.0
        new_orig = cur_fv if is_new else 0.0
        resid = cur_fv - prior_fv - repay_impact - pik_cap - new_orig
        rows.append({"Borrower": nm, "Prior FV": prior_fv, "Repayments": repay_impact,
                     "PIK Capitalized": pik_cap, "New Originations": new_orig,
                     "Credit-Specific": resid if credit_event else 0.0,
                     "Market Re-mark & Accretion": 0.0 if credit_event else resid,
                     "Current FV": cur_fv})
    return pd.DataFrame(rows)
