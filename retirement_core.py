"""
retirement_core.py
Core retirement model, scenario generator, charts, and Word report (DOCX).
Designed to be used from a Streamlit UI.
"""
from __future__ import annotations
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from docx import Document
import matplotlib.pyplot as plt
import datetime as _dt
import os
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional, Tuple, Union, IO
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
# ----------------------------
# Core model
# ----------------------------


@dataclass
class Inputs:
    current_age: int = 40
    pension_access_age: int = 57
    state_pension_age: int = 67
    life_expectancy: int = 85
    annual_spending_today: float = 20000.0
    inflation_rate: float = 0.025
    nominal_return: float = 0.05
    include_state_pension: bool = True
    state_pension_annual_today: float = 11502.0
    sizing_method: str = "SWR"  # "SWR" or "PV"
    swr: float = 0.04
    current_isa: float = 0.0
    current_pension: float = 0.0
    # Desired pension balance (today's £) at life_expectancy
    desired_pension_at_end: float = 0.0

    # --- NEW: Scenario triggers ---
    adj_swr: float = 0.035
    adj_inflation: float = 0.005
    adj_return: float = 0.005


@dataclass
class Results:
    real_return: float
    bridge_years: int
    isa_needed_today: float
    pension_needed_today: float
    total_needed_today: float
    pension_needed_at_access: float
    mathematical_min_total_today: float


def _safe_pv_annuity(payment: float, r: float, n: int) -> float:
    """PV of n end-of-period payments at rate r. Handles near-zero r."""
    if n <= 0:
        return 0.0
    if abs(r) < 1e-9:
        return payment * n
    return payment * (1 - (1 + r) ** (-n)) / r


def _real_return(nominal_return: float, inflation_rate: float) -> float:
    return (1 + nominal_return) / (1 + inflation_rate) - 1


def size_capital(inputs: Inputs) -> Tuple[Results, pd.DataFrame, pd.DataFrame]:
    """
    Compute required capital and produce deterministic projections.
    Returns:
        results: summary numbers
        projection: year-by-year cashflow projection
        meta: one-row dataframe with inputs + results
    """
    # Validation
    if inputs.pension_access_age < inputs.current_age:
        raise ValueError("Pension access age must be >= current age")
    if inputs.life_expectancy < inputs.current_age:
        raise ValueError("Life expectancy must be >= current age")
    if inputs.annual_spending_today < 0:
        raise ValueError("Annual spending must be >= 0")
    if inputs.inflation_rate <= -0.99:
        raise ValueError("Inflation rate too low")
    if inputs.nominal_return <= -0.99:
        raise ValueError("Nominal return too low")
    if inputs.sizing_method not in {"PV", "SWR"}:
        raise ValueError("sizing_method must be 'PV' or 'SWR'")
    if inputs.sizing_method == "SWR" and not (0 < inputs.swr < 0.2):
        raise ValueError("SWR must be between 0 and 20%")
    rr = _real_return(inputs.nominal_return, inputs.inflation_rate)
    bridge_years = inputs.pension_access_age - inputs.current_age
    # ISA bridge sizing
    isa_needed_today = _safe_pv_annuity(
        inputs.annual_spending_today, rr, bridge_years) * (1 + rr)
    # Post-access phases
    years_access_to_state = max(
        0, inputs.state_pension_age - inputs.pension_access_age)
    years_state_to_end = max(
        0, inputs.life_expectancy - inputs.state_pension_age + 1)
    pv_pre_state_at_access = _safe_pv_annuity(
        inputs.annual_spending_today, rr, years_access_to_state
    ) * (1 + rr)    
    if inputs.include_state_pension:
        gap_today = max(
            0.0, inputs.annual_spending_today - inputs.state_pension_annual_today
        )
    else:
        gap_today = inputs.annual_spending_today
    pv_gap_at_state_age = _safe_pv_annuity(gap_today, rr, years_state_to_end) * (1 + rr)
    pv_gap_at_access = (
        pv_gap_at_state_age / ((1 + rr) ** years_access_to_state)
        if years_access_to_state > 0
        else pv_gap_at_state_age
    )
    # Desired end-of-life pension balance, discounted to access age using real return
    years_access_to_end = max(
        0, inputs.life_expectancy - inputs.pension_access_age + 1)
    pv_end_at_access = (
        inputs.desired_pension_at_end / ((1 + rr) ** years_access_to_end)
        if years_access_to_end > 0
        else inputs.desired_pension_at_end
    )
    pension_needed_at_access_pv = (
        pv_pre_state_at_access + pv_gap_at_access + pv_end_at_access
    )
    pension_needed_today_pv = (
        pension_needed_at_access_pv / ((1 + rr) ** bridge_years)
        if bridge_years > 0
        else pension_needed_at_access_pv
    )
    mathematical_min_total_today = isa_needed_today + pension_needed_today_pv

    # SWR sizing (add desired end-balance as extra lump sum at access)
    if inputs.sizing_method == "SWR":
        mult = 1.0 / inputs.swr
    
        if years_state_to_end > 0:
            # You reach State Pension age during the model horizon
            if years_access_to_state > 0:
                # Before State Pension: you need full spending; after: only the gap
                temp_shortfall = inputs.annual_spending_today - gap_today  # typically the State Pension amount (capped by spending)
                pv_temp = _safe_pv_annuity(temp_shortfall, rr, years_access_to_state) * (1 + rr)
                pension_needed_at_access_base = (gap_today * mult) + pv_temp
            else:
                # State Pension already active at pension access
                pension_needed_at_access_base = gap_today * mult
        else:
            # You never reach State Pension age (so it must not reduce required pot)
            pension_needed_at_access_base = inputs.annual_spending_today * mult
    
        pension_needed_at_access = pension_needed_at_access_base + pv_end_at_access
        pension_needed_today = (
            pension_needed_at_access / ((1 + rr) ** bridge_years)
            if bridge_years > 0
            else pension_needed_at_access
        )
    else:
        pension_needed_at_access = pension_needed_at_access_pv
        pension_needed_today = pension_needed_today_pv
    total_needed_today = isa_needed_today + pension_needed_today
    results = Results(
        real_return=rr,
        bridge_years=bridge_years,
        isa_needed_today=isa_needed_today,
        pension_needed_today=pension_needed_today,
        total_needed_today=total_needed_today,
        pension_needed_at_access=pension_needed_at_access,
        mathematical_min_total_today=mathematical_min_total_today,
    )
    # Deterministic projection with realistic depletion:
    projection: List[Dict] = []
    bal_isa = isa_needed_today
    bal_pension = pension_needed_today

    has_actuals = inputs.current_isa > 0 or inputs.current_pension > 0
    act_isa = inputs.current_isa
    act_pension = inputs.current_pension

    for t in range(inputs.life_expectancy - inputs.current_age + 1):
        age = inputs.current_age + t
        infl_factor = (1 + inputs.inflation_rate) ** t
        spending_nominal = inputs.annual_spending_today * infl_factor

        state_pension_nominal = 0.0
        if inputs.include_state_pension and age >= inputs.state_pension_age:
            state_pension_nominal = inputs.state_pension_annual_today * infl_factor

        withdrawal_needed = max(0.0, spending_nominal - state_pension_nominal)

        if age < inputs.pension_access_age:
            bal_isa -= spending_nominal
            bal_isa *= (1 + inputs.nominal_return)
            bal_pension *= (1 + inputs.nominal_return)
            source = "ISA"
            portfolio_withdrawal = spending_nominal
            
            if has_actuals:
                act_isa -= spending_nominal
                if act_isa > 0:
                    act_isa *= (1 + inputs.nominal_return)
                act_pension *= (1 + inputs.nominal_return)
        else:
            bal_isa *= (1 + inputs.nominal_return)
            bal_pension -= withdrawal_needed
            bal_pension *= (1 + inputs.nominal_return)
            source = "Pension"
            portfolio_withdrawal = withdrawal_needed

            if has_actuals:
                if act_isa > 0:
                    act_isa *= (1 + inputs.nominal_return)
                act_pension -= withdrawal_needed
                act_pension *= (1 + inputs.nominal_return)

        total_bal = bal_isa + bal_pension
        act_total = (act_isa + act_pension) if has_actuals else None

        projection.append(
            {
                "Age": age,
                "YearIndex": t,
                "Spending": spending_nominal,
                "StatePension": state_pension_nominal,
                "PortfolioWithdrawal": portfolio_withdrawal,
                "Source": source,
                "ISA_Balance": bal_isa,
                "Pension_Balance": bal_pension,
                "Total_Balance": total_bal,
                "Actual_Total_Balance": act_total,
                # --- NEW: Track separate actuals ---
                "Actual_ISA_Balance": act_isa if has_actuals else None,
                "Actual_Pension_Balance": act_pension if has_actuals else None,
            }
        )
    df = pd.DataFrame(projection)
    meta = {**asdict(inputs), **asdict(results)}
    meta_df = pd.DataFrame([meta])
    return results, df, meta_df
# ----------------------------
# Charts
# ----------------------------


def _fmt_k(x, _pos=None):
    return f"£{x/1000:.0f}k"


def plot_single_projection(
    inputs: Inputs, results: Results, df: pd.DataFrame, out_png: str
) -> None:
    """Single-scenario visualisation (PNG)."""
    fig, axes = plt.subplots(2, 1, figsize=(13, 9), sharex=True)
    # Balances
    ax = axes[0]
    ax.plot(df["Age"], df["Total_Balance"], lw=2.6, label="Total (Target)", color="#2E86AB")
    ax.plot(df["Age"], df["ISA_Balance"], lw=1.8, label="ISA (Target)", color="#A23B72")
    ax.plot(df["Age"], df["Pension_Balance"], lw=1.8, label="Pension (Target)", color="#F18F01")
    
    # --- NEW: Plot separate actuals and highlight shortfall ---
    if "Actual_Total_Balance" in df.columns and df["Actual_Total_Balance"].notna().any():
        ax.plot(df["Age"], df["Actual_Total_Balance"], lw=2.6, label="Actual Total", color="black", ls="--")
        ax.plot(df["Age"], df["Actual_ISA_Balance"], lw=1.8, label="Actual ISA", color="#A23B72", ls=":")
        ax.plot(df["Age"], df["Actual_Pension_Balance"], lw=1.8, label="Actual Pension", color="#F18F01", ls=":")
        
        # Highlight ISA Shortfall before pension access
        bridge_mask = df["Age"] <= inputs.pension_access_age
        if (df.loc[bridge_mask, "Actual_ISA_Balance"] < 0).any():
            ax.fill_between(
                df.loc[bridge_mask, "Age"],
                0,
                df.loc[bridge_mask, "Actual_ISA_Balance"],
                where=df.loc[bridge_mask, "Actual_ISA_Balance"] < 0,
                color="red",
                alpha=0.25,
                label="ISA Shortfall (Illiquid)",
                interpolate=True
            )

    ax.axhline(0, color="red", ls="--", lw=1, alpha=0.6)
    ax.axvline(inputs.pension_access_age, color="orange", ls=":", lw=2)
    if inputs.include_state_pension:
        ax.axvline(inputs.state_pension_age, color="green", ls=":", lw=2)
    ax.set_title("Portfolio balances over time",
                 fontsize=13, fontweight="bold")
    ax.set_ylabel("Balance (£)", fontsize=11)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(_fmt_k))
    ax.grid(True, alpha=0.25)
    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1.0), borderaxespad=0.0)
    # Income / spending
    ax = axes[1]
    ax.plot(
        df["Age"],
        df["Spending"],
        lw=2.2,
        color="red",
        ls="--",
        label="Spending",
    )
    ax.fill_between(
        df["Age"],
        0,
        df["StatePension"],
        alpha=0.45,
        color="#06A77D",
        label="State pension",
    )
    ax.fill_between(
        df["Age"],
        df["StatePension"],
        df["StatePension"] + df["PortfolioWithdrawal"],
        alpha=0.45,
        color="#D4B483",
        label="Portfolio withdrawal",
    )
    ax.axvline(
        inputs.pension_access_age, color="orange", ls=":", lw=2, label="Pension access"
    )
    if inputs.include_state_pension:
        ax.axvline(
            inputs.state_pension_age,
            color="green",
            ls=":",
            lw=2,
            label="State pension starts",
        )
    ax.set_title("Annual spending and income sources",
                 fontsize=13, fontweight="bold")
    ax.set_xlabel("Age (years)", fontsize=11)
    ax.set_ylabel("Annual amount (£)", fontsize=11)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(_fmt_k))
    ax.grid(True, alpha=0.25)
    ax.legend(
        loc="upper left", bbox_to_anchor=(1.02, 1.0), borderaxespad=0.0, fontsize=9
    )
    fig.suptitle(
        f"Retirement projection | sizing={inputs.sizing_method}"
        f" | return={inputs.nominal_return:.1%} | inflation={inputs.inflation_rate:.1%}"
        f" | spending=£{inputs.annual_spending_today:,.0f} (today)",
        y=0.98,
        fontsize=12,
    )
    fig.tight_layout(rect=[0, 0, 0.82, 0.96])
    fig.savefig(out_png, dpi=200, bbox_inches="tight")
    plt.close(fig)


def plot_scenario_set(base: Inputs, out_png: str) -> pd.DataFrame:
    """Scenario-comparison visualisation (PNG) + summary dataframe."""
    scenarios: List[Tuple[str, Inputs]] = []
    scenarios.append(("Base (SWR 4%)", base))
    scenarios.append(
        (
            f"Conservative (SWR {base.adj_swr*100:.1f}%)",
            Inputs(**{**asdict(base), "swr": base.adj_swr, "sizing_method": "SWR"}),
        )
    )
    scenarios.append(
        ("PV (mathematical min)", Inputs(**{**asdict(base), "sizing_method": "PV"}))
    )
    scenarios.append(
        (
            f"Inflation adjustment ({base.adj_inflation*100:+.2f}%)",
            Inputs(**{**asdict(base), "inflation_rate": base.inflation_rate + base.adj_inflation}),
        )
    )
    scenarios.append(
        (
            f"Returns adjustment ({base.adj_return*100:+.2f}%)",
            Inputs(**{**asdict(base), "nominal_return": base.nominal_return + base.adj_return}),
        )
    )

    summary_rows: List[Dict] = []
    scenario_results: List[Tuple[str, Inputs, Results, pd.DataFrame]] = []
    for label, inp in scenarios:
        res, df, _meta = size_capital(inp)
        summary_rows.append(
            {
                "Scenario": label,
                "TotalNeededToday": res.total_needed_today,
                "ISANeededToday": res.isa_needed_today,
                "PensionNeededToday": res.pension_needed_today,
            }
        )
        scenario_results.append((label, inp, res, df))
    summary = pd.DataFrame(summary_rows)
    fig, axes = plt.subplots(2, 2, figsize=(16, 11))
    # (1) Total balance over time
    ax = axes[0, 0]
    for label, _inp, _res, df_s in scenario_results:
        ax.plot(df_s["Age"], df_s["Total_Balance"], lw=2, label=label)
    ax.axhline(0, color="red", ls="--", lw=1, alpha=0.6)
    ax.set_title("Total balance over time", fontweight="bold")
    ax.set_xlabel("Age (years)")
    ax.set_ylabel("Balance (£)")
    ax.yaxis.set_major_formatter(plt.FuncFormatter(_fmt_k))
    ax.grid(True, alpha=0.25)
    ax.legend(
        fontsize=8, loc="upper left", bbox_to_anchor=(1.02, 1.0), borderaxespad=0.0
    )
    # (2) ISA vs pension needed today
    ax = axes[0, 1]
    x = np.arange(len(summary))
    w = 0.38
    ax.bar(x - w / 2, summary["ISANeededToday"],
           width=w, label="ISA", color="#A23B72")
    ax.bar(
        x + w / 2,
        summary["PensionNeededToday"],
        width=w,
        label="Pension",
        color="#F18F01",
    )
    ax.set_title("Capital needed today", fontweight="bold")
    ax.set_xticks(x)
    ax.set_xticklabels(summary["Scenario"],
                       rotation=30, ha="right", fontsize=8)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(_fmt_k))
    ax.grid(True, alpha=0.25, axis="y")
    ax.legend(
        loc="upper left", bbox_to_anchor=(1.02, 1.0), borderaxespad=0.0
    )
    # (3) Base scenario income sources
    base_label, base_inp, _base_res, base_df = scenario_results[0]
    ax = axes[1, 0]
    ax.fill_between(
        base_df["Age"],
        0,
        base_df["StatePension"],
        alpha=0.45,
        color="#06A77D",
        label="State pension",
    )
    ax.fill_between(
        base_df["Age"],
        base_df["StatePension"],
        base_df["StatePension"] + base_df["PortfolioWithdrawal"],
        alpha=0.45,
        color="#D4B483",
        label="Portfolio",
    )
    ax.plot(
        base_df["Age"],
        base_df["Spending"],
        color="red",
        ls="--",
        lw=2,
        label="Spending",
    )
    ax.axvline(
        base_inp.pension_access_age,
        color="orange",
        ls=":",
        lw=2,
        label="Pension access",
    )
    if base_inp.include_state_pension:
        ax.axvline(
            base_inp.state_pension_age,
            color="green",
            ls=":",
            lw=2,
            label="State pension",
        )
    ax.set_title("Base scenario: spending and income", fontweight="bold")
    ax.set_xlabel("Age (years)")
    ax.set_ylabel("Annual amount (£)")
    ax.yaxis.set_major_formatter(plt.FuncFormatter(_fmt_k))
    ax.grid(True, alpha=0.25)
    ax.legend(
        fontsize=8, loc="upper left", bbox_to_anchor=(1.02, 1.0), borderaxespad=0.0
    )
    # (4) Summary table
    ax = axes[1, 1]
    ax.axis("off")
    tdf = summary.copy()
    for c in ["TotalNeededToday", "ISANeededToday", "PensionNeededToday"]:
        tdf[c] = tdf[c].map(lambda v: f"£{v:,.0f}")
    tdf.columns = [
        "Scenario",
        "Total\nNeeded Today",
        "ISA\nNeeded Today",
        "Pension\nNeeded Today",
    ]
    tbl = ax.table(
        cellText=tdf.values.tolist(),
        colLabels=list(tdf.columns),
        cellLoc="center",
        loc="center",
    )
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(9)
    tbl.scale(1.0, 1.9)
    tbl.auto_set_column_width(col=list(range(len(tdf.columns))))
    ax.set_title("Capital needed today", fontweight="bold")
    fig.suptitle("Retirement scenario comparison", fontsize=14, fontweight="bold")
    fig.tight_layout(rect=[0, 0, 0.82, 0.96])
    fig.savefig(out_png, dpi=200, bbox_inches="tight")
    plt.close(fig)
    return summary

# ----------------------------
# DOCX report (accepts BytesIO)
# ----------------------------

def create_pension_report(
    out_docx: Union[str, IO[bytes]],
    inputs: Inputs,
    results: Results,
    df: pd.DataFrame,
    chart_png: str,
    scenario_png: Optional[str] = None,
) -> None:
    """
    Generate a professional pension report with embedded charts and commentary.
    out_docx:
      - path string (for scripts), or
      - BytesIO file-like object (for Streamlit download)
    """
    doc = Document()
    # Title
    title = doc.add_heading("Retirement Capital Planning Report", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.runs[0]
    title_run.font.size = Pt(20)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(46, 134, 171)
    # Date
    date_para = doc.add_paragraph(
        f"Report generated: {_dt.datetime.now().strftime('%d %B %Y')}"
    )
    date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    date_para.runs[0].font.size = Pt(11)
    date_para.runs[0].font.italic = True
    doc.add_paragraph()
    # Executive summary
    doc.add_heading("Executive Summary", level=1)
    summary_text = (
        "This report provides a deterministic retirement capital analysis based on your specified assumptions. "
        f"The analysis sizes the capital required in two separate 'pots': an ISA (or accessible account) to bridge from your current age "
        f"({inputs.current_age}) to pension access age ({inputs.pension_access_age}), and a pension pot to fund spending from pension "
        f"access until your planning horizon (age {inputs.life_expectancy})."
    )
    p = doc.add_paragraph(summary_text)
    for run in p.runs:
        run.font.size = Pt(11)
    doc.add_paragraph()
    # Key capital requirements
    doc.add_heading("Key Capital Requirements", level=2)
    table = doc.add_table(rows=4, cols=2)
    table.style = "Light Grid Accent 1"
    cells = [
        ("ISA / Accessible capital needed (today)",
         f"£{results.isa_needed_today:,.0f}"),
        ("Pension pot needed (today)",
         f"£{results.pension_needed_today:,.0f}"),
        ("Total capital needed (today)",
         f"£{results.total_needed_today:,.0f}"),
        (
            "Mathematical minimum (PV method)",
            f"£{results.mathematical_min_total_today:,.0f}",
        ),
    ]
    for i, (label, value) in enumerate(cells):
        row = table.rows[i]
        row.cells[0].text = label
        row.cells[1].text = value
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(10)
    doc.add_paragraph()
    # Capital commentary
    doc.add_heading("Capital Requirement Commentary", level=2)
    gap_isa = (
        results.isa_needed_today - inputs.current_isa
        if inputs.current_isa > 0
        else results.isa_needed_today
    )
    gap_pension = (
        results.pension_needed_today - inputs.current_pension
        if inputs.current_pension > 0
        else results.pension_needed_today
    )
    gap_total = results.total_needed_today - (
        inputs.current_isa + inputs.current_pension
    )
    commentary_paras: List[str] = []
    commentary_paras.append(
        f"**ISA Bridge Requirement:** You need £{results.isa_needed_today:,.0f} in accessible capital (ISA or similar) today to fund "
        f"{results.bridge_years} years of spending from age {inputs.current_age} to {inputs.pension_access_age}, before your pension becomes accessible. "
        f"This represents the present value of £{inputs.annual_spending_today:,.0f} annual spending (in today's terms), growing at {inputs.inflation_rate:.1%} "
        f"inflation and discounted at a real return of {results.real_return:.2%}."
    )
    commentary_paras.append(
        f"**Pension Pot Requirement:** You need £{results.pension_needed_today:,.0f} in your pension today to fund spending from age "
        f"{inputs.pension_access_age} (pension access) until age {inputs.life_expectancy}. This calculation uses the **{inputs.sizing_method}** sizing method."
    )
    if inputs.sizing_method == "SWR":
        commentary_paras.append(
            f"The **Safe Withdrawal Rate (SWR)** method sizes your pension pot based on a {inputs.swr:.1%} withdrawal rate. This is a "
            f"rule-of-thumb approach that aims to provide a buffer against sequence-of-returns risk. At age {inputs.pension_access_age}, you'll need "
            f"£{results.pension_needed_at_access:,.0f} in your pension, which corresponds to £{results.pension_needed_today:,.0f} today "
            f"(discounted over {results.bridge_years} years)."
        )
    else:
        commentary_paras.append(
            f"The **Present Value (PV)** method sizes your pension pot mathematically: it calculates the exact present value of all future "
            f"inflation-adjusted spending needs, reduced by State Pension income (if applicable). This is the theoretical minimum required under "
            f"deterministic assumptions (constant {inputs.nominal_return:.1%} return)."
        )
    # Longevity buffer explanation
    if inputs.desired_pension_at_end > 0:
        approx_years = int(
            inputs.desired_pension_at_end /
            max(inputs.annual_spending_today, 1)
        )
        commentary_paras.append(
            f"**Longevity Buffer:** You have specified a desired pension balance of £{inputs.desired_pension_at_end:,.0f} (in today's terms) to remain "
            f"at age {inputs.life_expectancy}. This buffer provides a safety margin if you live beyond your planning horizon. With your spending assumptions, "
            f"this residual capital could fund approximately {approx_years} additional years at current spending levels (ignoring State Pension). "
            f"If you set this to £0, the model calculates the minimum capital needed to fully utilize your portfolio by age {inputs.life_expectancy}."
        )
    else:
        commentary_paras.append(
            f"**Longevity Buffer:** You have set the desired end-of-life pension balance to £0, meaning the model calculates the **minimum** capital required "
            f"to fully utilize your ISA and pension by age {inputs.life_expectancy}. If you wish to plan for longevity beyond age {inputs.life_expectancy}, "
            f"consider setting a desired pension balance (e.g., £100,000) to create a buffer for additional years."
        )
    if inputs.include_state_pension:
        commentary_paras.append(
            f"**State Pension Impact:** From age {inputs.state_pension_age}, you will receive £{inputs.state_pension_annual_today:,.0f} per year "
            "(in today's terms) from the State Pension. This reduces the withdrawal required from your pension pot, as you only need to draw the 'gap' "
            "between your spending target and State Pension income."
        )
    if inputs.current_isa > 0 or inputs.current_pension > 0:
        commentary_paras.append(
            f"**Your Current Position vs Target:** Based on your stated current holdings (ISA: £{inputs.current_isa:,.0f}, Pension: "
            f"£{inputs.current_pension:,.0f}), your shortfall is: ISA gap £{gap_isa:,.0f}, Pension gap £{gap_pension:,.0f}, "
            f"Total gap £{gap_total:,.0f}."
        )
    for para_text in commentary_paras:
        p = doc.add_paragraph()
        parts = para_text.split("**")
        p.add_run(parts[0])
        for i, segment in enumerate(parts[1:]):
            if i % 2 == 0:
                run = p.add_run(segment)
                run.bold = True
            else:
                run = p.add_run(segment)
        for run in p.runs:
            run.font.size = Pt(11)
    doc.add_page_break()
    # Assumptions
    doc.add_heading("Model Assumptions", level=1)
    doc.add_paragraph(
        "This analysis is based on a **deterministic model** with constant return assumptions. Real-world outcomes will vary due to "
        "market volatility, sequence of returns, longevity uncertainty, and changes in personal circumstances.",
        style="Intense Quote",
    )
    doc.add_heading("Key Assumptions", level=2)
    assumptions_table = doc.add_table(rows=11, cols=2)
    assumptions_table.style = "Light List Accent 1"
    assumptions_data = [
        ("Current age", f"{inputs.current_age} years"),
        ("Pension access age", f"{inputs.pension_access_age} years"),
        ("State Pension age", f"{inputs.state_pension_age} years"),
        (
            "Life expectancy (planning horizon)",
            f"{inputs.life_expectancy} years",
        ),
        ("Annual spending target (today)",
         f"£{inputs.annual_spending_today:,.0f}"),
        ("Inflation rate (constant)", f"{inputs.inflation_rate:.2%} per year"),
        ("Nominal investment return (constant)",
         f"{inputs.nominal_return:.2%} per year"),
        ("Real return (inflation-adjusted)",
         f"{results.real_return:.2%} per year"),
        ("State Pension included", "Yes" if inputs.include_state_pension else "No"),
        (
            "State Pension amount (today)",
            f"£{inputs.state_pension_annual_today:,.0f}"
            if inputs.include_state_pension
            else "N/A",
        ),
        (
            "Desired pension at end of life (today)",
            f"£{inputs.desired_pension_at_end:,.0f}",
        ),
    ]
    for i, (label, value) in enumerate(assumptions_data):
        row = assumptions_table.rows[i]
        row.cells[0].text = label
        row.cells[1].text = value
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(10)
    doc.add_paragraph()
    # Limitations
    doc.add_heading("Important Limitations", level=2)
    limitations = [
        f"**Constant returns:** This model assumes a fixed {inputs.nominal_return:.1%} nominal return every year. Real portfolios experience volatility, "
        "and sequence of returns matters—poor early returns can deplete capital faster.",
        f"**Constant inflation:** Inflation is assumed constant at {inputs.inflation_rate:.1%}. Actual inflation varies year-to-year.",
        "**No flexibility modelled:** The model assumes fixed annual spending (inflation-adjusted). Real retirees can adjust spending "
        "in response to portfolio performance.",
        f"**Longevity risk:** If you live beyond age {inputs.life_expectancy}, you will need additional capital or reduced spending. "
        f"The desired end-of-life pension balance ( £{inputs.desired_pension_at_end:,.0f} ) provides some buffer for this risk.",
        "**Tax not modelled:** ISA withdrawals are tax-free in the UK, but pension withdrawals may be subject to income tax. "
        "This analysis does not model tax explicitly.",
        "**Pension drawdown mechanics:** Post-access, the pension pot is gradually drawn down to fund spending (reduced by State Pension). "
        "The projection shows this depletion over time. Withdrawals are taken before growth each year to create a realistic depletion curve.",
    ]
    for limitation in limitations:
        p = doc.add_paragraph()
        parts = limitation.split("**")
        p.add_run(parts[0])
        for i, segment in enumerate(parts[1:]):
            if i % 2 == 0:
                run = p.add_run(segment)
                run.bold = True
                run.font.size = Pt(11)
            else:
                run = p.add_run(segment)
                run.font.size = Pt(11)
    doc.add_page_break()
    # Charts
    doc.add_heading("Projection Visualisations", level=1)
    doc.add_heading("Single Scenario Projection", level=2)
    doc.add_paragraph(
        "The chart below shows the projected balances and income sources over your retirement horizon under the base assumptions."
    )
    if os.path.exists(chart_png):
        doc.add_picture(chart_png, width=Inches(6.0))
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()
    doc.add_heading("Chart Interpretation", level=3)
    chart_explanation = [
        f"**Top panel (Portfolio Balances):** Shows how your ISA and pension balances evolve over time. "
        f"The ISA is depleted during the bridge period (ages {inputs.current_age} to {inputs.pension_access_age}), after which it remains at zero. "
        f"The pension pot grows until age {inputs.pension_access_age}, then is gradually drawn down to fund your spending gap (after State Pension, if applicable) "
        f"until age {inputs.life_expectancy}. The pension balance at age {inputs.life_expectancy} should be approximately "
        f"£{inputs.desired_pension_at_end:,.0f} (your specified buffer).",
        f"**Bottom panel (Spending and Income):** The red dashed line shows your inflation-adjusted spending target. The filled areas "
        f"show income sources: green represents State Pension (from age {inputs.state_pension_age} onward), and tan represents portfolio withdrawals. "
        "The portfolio withdrawals come entirely from the ISA before pension access, and entirely from the pension thereafter.",
        f"**Vertical lines:** Orange line marks pension access age ({inputs.pension_access_age}), green line marks State Pension start ({inputs.state_pension_age}).",
    ]
    for explanation in chart_explanation:
        p = doc.add_paragraph()
        parts = explanation.split("**")
        p.add_run(parts[0])
        for i, segment in enumerate(parts[1:]):
            if i % 2 == 0:
                run = p.add_run(segment)
                run.bold = True
                run.font.size = Pt(10)
            else:
                run = p.add_run(segment)
                run.font.size = Pt(10)
    doc.add_paragraph()
    doc.add_heading("How the Pension Pot Is Drawn Down", level=3)
    drawdown_text = (
        f"After you reach pension access age ({inputs.pension_access_age}), the model draws from your pension pot each year to "
        f"cover the spending need (net of State Pension). The projection uses a 'withdraw-then-grow' approach: each year, the required withdrawal "
        f"is taken first, then the remaining balance grows at {inputs.nominal_return:.1%}. This creates a realistic depletion curve rather than "
        "an ever-growing pension pot. "
        f"By age {inputs.life_expectancy}, the pension pot should reach approximately £{inputs.desired_pension_at_end:,.0f} "
        "(your specified end-of-life balance). If you set this to £0, the pension is fully consumed by your planning horizon."
    )
    p = doc.add_paragraph(drawdown_text)
    for run in p.runs:
        run.font.size = Pt(11)
    doc.add_page_break()
    # Scenario comparison
    if scenario_png and os.path.exists(scenario_png):
        doc.add_heading("Scenario Comparison", level=2)
        doc.add_paragraph(
            "The chart below compares your base scenario against alternative assumptions (conservative SWR, mathematical minimum PV, "
            "higher inflation, higher returns). This illustrates the range of capital requirements under different conditions."
        )
        doc.add_picture(scenario_png, width=Inches(6.0))
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_page_break()
    # Conclusion
    doc.add_heading("Conclusion and Next Steps", level=1)
    conclusion_text = (
        f"Based on the assumptions provided, you require a total of £{results.total_needed_today:,.0f} in capital today "
        f"(£{results.isa_needed_today:,.0f} in ISA, £{results.pension_needed_today:,.0f} in pension) to fund £{inputs.annual_spending_today:,.0f} "
        f"annual spending (in today's terms) from age {inputs.current_age} to {inputs.life_expectancy}.\n\n"
        "**Recommended next steps:**\n\n"
        "1. **Review assumptions:** Ensure inflation, return, and spending assumptions reflect your expectations and risk tolerance.\n\n"
        "2. **Stress test:** Consider running additional scenarios with lower returns or higher inflation to understand downside risks.\n\n"
        "3. **Tax planning:** Consult a tax adviser to optimize ISA vs pension contributions and withdrawal strategies.\n\n"
        f"4. **Longevity planning:** If you have family history of longevity beyond age {inputs.life_expectancy}, consider extending the planning horizon "
        "or increasing the desired end-of-life pension balance.\n\n"
        "5. **Professional advice:** This is a simplified model. Consider consulting an independent financial adviser (IFA) for comprehensive retirement planning, "
        "including State Pension forecasts, annuity options, and estate planning.\n\n"
        "**Disclaimer:** This report is for illustrative purposes only and does not constitute financial advice. Actual investment returns, inflation, "
        "and personal circumstances will vary. Always seek professional financial advice before making retirement decisions."
    )
    for para_text in conclusion_text.split("\n\n"):
        p = doc.add_paragraph()
        parts = para_text.split("**")
        p.add_run(parts[0])
        for i, segment in enumerate(parts[1:]):
            if i % 2 == 0:
                run = p.add_run(segment)
                run.bold = True
                run.font.size = Pt(11)
            else:
                run = p.add_run(segment)
                run.font.size = Pt(11)
    # Save to filename or BytesIO
    doc.save(out_docx)
