"""
retirement_core.py
Core math, dataclasses, and plotting for the retirement planner.
No Streamlit dependencies here.
"""
import copy
from dataclasses import dataclass
from typing import Dict, List, Tuple
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from docx import Document
from docx.shared import Inches


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
    # --- NEW: Other Income ---
    other_income: float = 0.0
    other_income_years: int = 0


@dataclass
class Results:
    real_return: float
    bridge_years: int
    isa_needed_today: float
    pension_needed_today: float
    total_needed_today: float
    pension_needed_at_access: float
    mathematical_min_total_today: float  # Only used for comparison if method="SWR"


def _safe_pv_annuity(payment: float, r: float, n: int) -> float:
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
        
    # --- NEW: Other Income Offset (ISA Phase) ---
    isa_oi_yrs = min(bridge_years, inputs.other_income_years)
    pv_oi_isa = _safe_pv_annuity(inputs.other_income, inputs.nominal_return, isa_oi_yrs) * (1 + inputs.nominal_return)
    isa_needed_today = max(0.0, isa_needed_today - pv_oi_isa)
    
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

    # --- NEW: Other Income Offset (Pension Phase) ---
    pens_oi_yrs = max(0, inputs.other_income_years - bridge_years)
    if pens_oi_yrs > 0:
        pv_oi_pens_at_access = _safe_pv_annuity(inputs.other_income, inputs.nominal_return, pens_oi_yrs) * (1 + inputs.nominal_return)
        pv_oi_pens_today = pv_oi_pens_at_access / ((1 + inputs.nominal_return) ** bridge_years) if bridge_years > 0 else pv_oi_pens_at_access
        
        pension_needed_at_access = max(0.0, pension_needed_at_access - pv_oi_pens_at_access)
        pension_needed_today = max(0.0, pension_needed_today - pv_oi_pens_today)

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

        # --- NEW: Other Income (Nominal) ---
        other_income_nominal = inputs.other_income if t < inputs.other_income_years else 0.0

        withdrawal_needed = max(0.0, spending_nominal - state_pension_nominal - other_income_nominal)

        if age < inputs.pension_access_age:
            isa_withdrawal = max(0.0, spending_nominal - other_income_nominal)
            bal_isa -= isa_withdrawal
            bal_isa *= (1 + inputs.nominal_return)
            bal_pension *= (1 + inputs.nominal_return)
            source = "ISA"
            portfolio_withdrawal = isa_withdrawal
            
            if has_actuals:
                act_isa -= isa_withdrawal
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
                "OtherIncome": other_income_nominal,
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
    
    # Build meta df
    meta_dict = vars(inputs).copy()
    meta_dict.update(vars(results))
    meta_df = pd.DataFrame([meta_dict])
    return results, df, meta_df


def _fmt_k(val: float, _: int) -> str:
    """Formatter for matplotlib axes (e.g. 500000 -> 500k)"""
    return f"{val/1000:,.0f}k"


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
                label="ISA Shortfall (before access)",
            )
            # Clip actual ISA to 0 in plot so it doesn't drop to minus infinity visually
            df["Actual_ISA_Balance_Plot"] = df["Actual_ISA_Balance"].clip(lower=0)
            ax.plot(df["Age"], df["Actual_ISA_Balance_Plot"], lw=1.8, color="#A23B72", ls=":")

    ax.set_title("Capital progression by phase")
    ax.set_ylabel("Nominal balance (£)")
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
    
    # --- NEW: Other Income shading (Blue) ---
    ax.fill_between(
        df["Age"],
        df["StatePension"],
        df["StatePension"] + df["OtherIncome"],
        alpha=0.3,
        color="#2E86AB", # Blue
        label="Other income",
    )
    ax.fill_between(
        df["Age"],
        df["StatePension"] + df["OtherIncome"],
        df["StatePension"] + df["OtherIncome"] + df["PortfolioWithdrawal"],
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
                 pad=15, fontsize=12, loc="left")
    ax.set_ylabel("Nominal income / spending (£)")
    ax.yaxis.set_major_formatter(plt.FuncFormatter(_fmt_k))
    ax.set_xlabel("Age")
    ax.grid(True, alpha=0.25)
    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1.0), borderaxespad=0.0)

    fig.suptitle(
        f"Retirement Plan (Method: {inputs.sizing_method})"
        f" | Total Target today: £{results.total_needed_today:,.0f}"
        f" | spending=£{inputs.annual_spending_today:,.0f} (today)",
        fontsize=14,
        y=0.98,
    )
    plt.tight_layout(rect=[0, 0, 0.85, 0.95])
    plt.savefig(out_png, dpi=150)
    plt.close()


def plot_scenario_set(base_inputs: Inputs, out_png: str) -> pd.DataFrame:
    """
    Produce a suite of scenarios by tweaking key inputs.
    Returns summary DataFrame of required capital across scenarios.
    Plots a multi-line comparison and saves to out_png.
    """
    scenarios = [
        ("1. Base", copy.deepcopy(base_inputs)),
    ]

    # Scenario 2: Conservative SWR
    if base_inputs.sizing_method == "SWR":
        s2 = copy.deepcopy(base_inputs)
        s2.swr = base_inputs.adj_swr
        scenarios.append((f"2. Conserv SWR ({s2.swr*100:.1f}%)", s2))
    else:
        # PV doesn't use SWR, so just do a lower return scenario instead
        s2 = copy.deepcopy(base_inputs)
        s2.nominal_return -= base_inputs.adj_return
        scenarios.append(("2. Lower return (-0.5%)", s2))

    # Scenario 3: Higher inflation
    s3 = copy.deepcopy(base_inputs)
    s3.inflation_rate += base_inputs.adj_inflation
    scenarios.append(("3. High inflation (+0.5%)", s3))

    # Scenario 4: Lower returns
    s4 = copy.deepcopy(base_inputs)
    s4.nominal_return -= base_inputs.adj_return
    scenarios.append(("4. Low returns (-0.5%)", s4))

    # Scenario 5: Lower returns AND high inflation
    s5 = copy.deepcopy(base_inputs)
    s5.nominal_return -= base_inputs.adj_return
    s5.inflation_rate += base_inputs.adj_inflation
    scenarios.append(("5. Worst case (Low ret + High inf)", s5))

    results_data = []
    fig, ax = plt.subplots(figsize=(10, 6))

    for name, s_inputs in scenarios:
        res, df, _ = size_capital(s_inputs)
        results_data.append(
            {
                "Scenario": name,
                "Method": s_inputs.sizing_method,
                "ISA_Needed": res.isa_needed_today,
                "Pension_Needed": res.pension_needed_today,
                "Total_Needed": res.total_needed_today,
            }
        )
        # Plot target total balance for each scenario
        ax.plot(df["Age"], df["Total_Balance"], lw=2, label=name)

    # Summarise in dataframe
    summary_df = pd.DataFrame(results_data)

    ax.set_title("Total target capital needed over time (Scenarios)")
    ax.set_ylabel("Nominal total balance (£)")
    ax.yaxis.set_major_formatter(plt.FuncFormatter(_fmt_k))
    ax.set_xlabel("Age")
    ax.grid(True, alpha=0.3)
    ax.legend(loc="upper right")

    plt.tight_layout()
    plt.savefig(out_png, dpi=150)
    plt.close()

    return summary_df


def create_pension_report(
    out_docx,
    inputs: Inputs,
    results: Results,
    df: pd.DataFrame,
    chart_png: str,
    scenario_png: str,
):
    """Generate Word Document."""
    doc = Document()
    doc.add_heading("Retirement Plan – Output Report", 0)

    doc.add_heading("Target Capital Required Today", level=1)
    p = doc.add_paragraph()
    p.add_run(f"ISA / accessible needed: £{results.isa_needed_today:,.0f}\n").bold = True
    p.add_run(f"Pension needed: £{results.pension_needed_today:,.0f}\n").bold = True
    p.add_run(f"Total needed today: £{results.total_needed_today:,.0f}\n").bold = True

    doc.add_heading("Assumptions Summary", level=1)
    doc.add_paragraph(
        f"Ages: Current {inputs.current_age} | Access {inputs.pension_access_age} "
        f"| State {inputs.state_pension_age} | Life exp {inputs.life_expectancy}"
    )
    doc.add_paragraph(
        f"Spending: £{inputs.annual_spending_today:,.0f} (today's £) "
        f"| Inflation {inputs.inflation_rate*100:.1f}% "
        f"| Return {inputs.nominal_return*100:.1f}%"
    )
    if inputs.include_state_pension:
        doc.add_paragraph(f"State Pension included: £{inputs.state_pension_annual_today:,.0f} / yr (today's £)")
    if inputs.other_income > 0:
        doc.add_paragraph(f"Other Income: £{inputs.other_income:,.0f} for {inputs.other_income_years} years (nominal)")
    if inputs.desired_pension_at_end > 0:
        doc.add_paragraph(f"Longevity buffer (today's £) at life expectancy: £{inputs.desired_pension_at_end:,.0f}")
        
    doc.add_paragraph(f"Sizing methodology: {inputs.sizing_method}")

    doc.add_heading("Base Scenario Visualisation", level=1)
    doc.add_picture(chart_png, width=Inches(6.0))

    doc.add_heading("Scenario Sensitivities", level=1)
    doc.add_picture(scenario_png, width=Inches(6.0))

    doc.save(out_docx)
