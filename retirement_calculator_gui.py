"""retirement_calculator_gui.py

A retirement capital + cashflow calculator with a Tkinter GUI.

What it does
- Sizes ISA/accessible capital needed for the pre-pension bridge.
- Sizes pension capital needed for post-access spending (optionally reduced by State Pension from a chosen age).
- Supports two sizing methods:
    1) "PV": mathematical present value of inflation-linked spending until life expectancy.
    2) "SWR": rule-of-thumb multiplier at pension access age (4% or 3.5%).
- Produces deterministic year-by-year projections (constant nominal return each year).
- Exports CSV(s) and a PNG visualisation.

Run
    python retirement_calculator_gui.py

Notes
- This is NOT financial advice.
- All figures are in nominal pounds in the year they occur unless stated.
"""

from __future__ import annotations

import datetime as _dt
import math
import os
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional, Tuple

import pandas as pd
import numpy as np

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import matplotlib
matplotlib.use("Agg")  # needed so PNG export works without a display
import matplotlib.pyplot as plt


# ----------------------------
# Core model
# ----------------------------

@dataclass
class Inputs:
    current_age: int = 51
    pension_access_age: int = 57
    state_pension_age: int = 67
    life_expectancy: int = 85

    annual_spending_today: float = 20000.0  # in today's pounds

    inflation_rate: float = 0.015          # 1.5%
    nominal_return: float = 0.05           # 5%

    include_state_pension: bool = True
    state_pension_annual_today: float = 11502.0  # in today's pounds

    sizing_method: str = "SWR"             # "PV" or "SWR"
    swr: float = 0.04                      # used when sizing_method == "SWR"; e.g., 0.04 or 0.035

    # holdings (optional; used for gap display only)
    current_isa: float = 0.0
    current_pension: float = 0.0


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
    """Compute required capital and produce projections.

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
    if inputs.state_pension_age < inputs.current_age:
        # allowed, but odd; we keep it
        pass
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

    # All PV sizing is easiest in real terms (today's pounds), then discounting.
    # Bridge: need full spending (no pension withdrawals allowed)
    isa_needed_today = _safe_pv_annuity(inputs.annual_spending_today, rr, bridge_years)

    # Post-access phases
    years_access_to_state = max(0, inputs.state_pension_age - inputs.pension_access_age)
    years_state_to_end = max(0, inputs.life_expectancy - inputs.state_pension_age + 1)

    # Spending needs at/after access are in *today's pounds* when sizing via real PV.
    # At access age: PV of spending from access->state start
    pv_pre_state_at_access = _safe_pv_annuity(inputs.annual_spending_today, rr, years_access_to_state)

    if inputs.include_state_pension:
        gap_today = max(0.0, inputs.annual_spending_today - inputs.state_pension_annual_today)
    else:
        gap_today = inputs.annual_spending_today

    # At state pension age: PV of (gap) spending until end
    pv_gap_at_state_age = _safe_pv_annuity(gap_today, rr, years_state_to_end)

    # Discount that back to access age
    pv_gap_at_access = pv_gap_at_state_age / ((1 + rr) ** years_access_to_state) if years_access_to_state > 0 else pv_gap_at_state_age

    pension_needed_at_access_pv = pv_pre_state_at_access + pv_gap_at_access
    pension_needed_today_pv = pension_needed_at_access_pv / ((1 + rr) ** bridge_years) if bridge_years > 0 else pension_needed_at_access_pv

    mathematical_min_total_today = isa_needed_today + pension_needed_today_pv

    # SWR sizing (rule-of-thumb): size pension pot at access as gap-appropriate multiplier
    if inputs.sizing_method == "SWR":
        mult = 1.0 / inputs.swr

        # Conservative shortcut: require enough at access to cover the larger of:
        # - full spending (before state pension)
        # - gap spending (after state pension)
        need_pre = inputs.annual_spending_today * mult if years_access_to_state > 0 else 0.0
        need_post = gap_today * mult if years_state_to_end > 0 else 0.0

        pension_needed_at_access = max(need_pre, need_post)
        pension_needed_today = pension_needed_at_access / ((1 + rr) ** bridge_years) if bridge_years > 0 else pension_needed_at_access
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

    # Deterministic projection (uses spending objective; does not implement SWR spending rule)
    projection = []
    bal_isa = isa_needed_today
    bal_pension = pension_needed_today

    for t in range(inputs.life_expectancy - inputs.current_age + 1):
        age = inputs.current_age + t
        infl_factor = (1 + inputs.inflation_rate) ** t

        spending_nominal = inputs.annual_spending_today * infl_factor

        state_pension_nominal = 0.0
        if inputs.include_state_pension and age >= inputs.state_pension_age:
            state_pension_nominal = inputs.state_pension_annual_today * infl_factor

        withdrawal_needed = max(0.0, spending_nominal - state_pension_nominal)

        # grow then withdraw at end of year (simple convention)
        bal_isa *= (1 + inputs.nominal_return)
        bal_pension *= (1 + inputs.nominal_return)

        if age < inputs.pension_access_age:
            # cannot use pension
            bal_isa -= spending_nominal
            source = "ISA"
            portfolio_withdrawal = spending_nominal
        else:
            bal_pension -= withdrawal_needed
            source = "Pension"
            portfolio_withdrawal = withdrawal_needed

        total_bal = bal_isa + bal_pension

        projection.append({
            "Age": age,
            "YearIndex": t,
            "Spending": spending_nominal,
            "StatePension": state_pension_nominal,
            "PortfolioWithdrawal": portfolio_withdrawal,
            "Source": source,
            "ISA_Balance": bal_isa,
            "Pension_Balance": bal_pension,
            "Total_Balance": total_bal,
        })

    df = pd.DataFrame(projection)

    # meta for easy export
    meta = {**asdict(inputs), **asdict(results)}
    meta_df = pd.DataFrame([meta])

    return results, df, meta_df


# ----------------------------
# Charts
# ----------------------------

def _fmt_k(x, _pos=None):
    return f"£{x/1000:.0f}k"


def plot_single_projection(inputs: Inputs, results: Results, df: pd.DataFrame, out_png: str) -> None:
    fig, axes = plt.subplots(2, 1, figsize=(12, 9))

    # Balance
    ax = axes[0]
    ax.plot(df["Age"], df["Total_Balance"], lw=2, label="Total")
    ax.plot(df["Age"], df["ISA_Balance"], lw=1.5, label="ISA")
    ax.plot(df["Age"], df["Pension_Balance"], lw=1.5, label="Pension")
    ax.axhline(0, color="red", ls="--", lw=1, alpha=0.6)
    ax.axvline(inputs.pension_access_age, color="orange", ls=":", lw=2, label="Pension access")
    if inputs.include_state_pension:
        ax.axvline(inputs.state_pension_age, color="green", ls=":", lw=2, label="State pension")
    ax.set_title("Portfolio balances (deterministic)")
    ax.set_xlabel("Age")
    ax.set_ylabel("Balance")
    ax.yaxis.set_major_formatter(plt.FuncFormatter(_fmt_k))
    ax.legend(loc="best")

    # Income / withdrawals
    ax = axes[1]
    ax.plot(df["Age"], df["Spending"], lw=2, color="red", ls="--", label="Spending")
    ax.fill_between(df["Age"], 0, df["StatePension"], alpha=0.5, label="State pension")
    ax.fill_between(df["Age"], df["StatePension"], df["StatePension"] + df["PortfolioWithdrawal"], alpha=0.5, label="Portfolio withdrawal")
    ax.set_title("Annual spending and income sources")
    ax.set_xlabel("Age")
    ax.set_ylabel("Annual amount")
    ax.yaxis.set_major_formatter(plt.FuncFormatter(_fmt_k))
    ax.legend(loc="best")

    fig.suptitle(
        f"Retirement projection | sizing={inputs.sizing_method}"
        f" | return={inputs.nominal_return:.1%} | inflation={inputs.inflation_rate:.1%}"
        f" | spending=£{inputs.annual_spending_today:,.0f} (today)",
        y=0.98,
        fontsize=12,
    )

    fig.tight_layout(rect=[0, 0, 1, 0.96])
    fig.savefig(out_png, dpi=200)
    plt.close(fig)


def plot_scenario_set(base: Inputs, out_png: str) -> pd.DataFrame:
    """Reproduce a scenario-comparison visualisation similar to the earlier script.

    Returns a summary dataframe with capital requirements.
    """

    scenarios: List[Tuple[str, Inputs]] = []

    # Base
    scenarios.append(("Base (SWR 4%)", base))

    # Conservative SWR
    cons = Inputs(**{**asdict(base), "swr": 0.035, "sizing_method": "SWR"})
    scenarios.append(("Conservative (SWR 3.5%)", cons))

    # PV mathematical minimum
    pv = Inputs(**{**asdict(base), "sizing_method": "PV"})
    scenarios.append(("PV (mathematical min)", pv))

    # Higher inflation
    hi = Inputs(**{**asdict(base), "inflation_rate": base.inflation_rate + 0.005})
    scenarios.append((f"Higher inflation ({hi.inflation_rate:.1%})", hi))

    # Lower returns
    lr = Inputs(**{**asdict(base), "nominal_return": max(-0.5, base.nominal_return - 0.01)})
    scenarios.append((f"Lower returns ({lr.nominal_return:.1%})", lr))

    summary_rows = []
    scenario_results = []

    for label, inp in scenarios:
        res, df, meta = size_capital(inp)
        summary_rows.append({
            "Scenario": label,
            "TotalNeededToday": res.total_needed_today,
            "ISANeededToday": res.isa_needed_today,
            "PensionNeededToday": res.pension_needed_today,
        })
        scenario_results.append((label, inp, res, df))

    summary = pd.DataFrame(summary_rows)

    fig, axes = plt.subplots(2, 2, figsize=(15, 11))

    # (1) Total balance over time
    ax = axes[0, 0]
    for label, inp, res, df in scenario_results:
        ax.plot(df["Age"], df["Total_Balance"], lw=2, label=label)
    ax.axhline(0, color="red", ls="--", lw=1, alpha=0.6)
    ax.set_title("Total balance over time")
    ax.set_xlabel("Age")
    ax.set_ylabel("Balance")
    ax.yaxis.set_major_formatter(plt.FuncFormatter(_fmt_k))
    ax.legend(fontsize=8, loc="upper right")

    # (2) Bar chart: ISA vs pension needed today
    ax = axes[0, 1]
    x = np.arange(len(summary))
    w = 0.38
    ax.bar(x - w/2, summary["ISANeededToday"], width=w, label="ISA")
    ax.bar(x + w/2, summary["PensionNeededToday"], width=w, label="Pension")
    ax.set_title("Capital needed today")
    ax.set_xticks(x)
    ax.set_xticklabels(summary["Scenario"], rotation=20, ha="right")
    ax.yaxis.set_major_formatter(plt.FuncFormatter(_fmt_k))
    ax.legend()

    # (3) Income sources for base scenario
    base_label, base_inp, base_res, base_df = scenario_results[0]
    ax = axes[1, 0]
    ax.fill_between(base_df["Age"], 0, base_df["StatePension"], alpha=0.5, label="State pension")
    ax.fill_between(base_df["Age"], base_df["StatePension"], base_df["StatePension"] + base_df["PortfolioWithdrawal"], alpha=0.5, label="Portfolio")
    ax.plot(base_df["Age"], base_df["Spending"], color="red", ls="--", lw=2, label="Spending")
    ax.axvline(base_inp.pension_access_age, color="orange", ls=":", lw=2, label="Pension access")
    if base_inp.include_state_pension:
        ax.axvline(base_inp.state_pension_age, color="green", ls=":", lw=2, label="State pension")
    ax.set_title("Base scenario: spending and income")
    ax.set_xlabel("Age")
    ax.set_ylabel("Annual amount")
    ax.yaxis.set_major_formatter(plt.FuncFormatter(_fmt_k))
    ax.legend(fontsize=8, loc="upper left")

    # (4) Table
    ax = axes[1, 1]
    ax.axis("off")
    table_df = summary.copy()
    table_df["TotalNeededToday"] = table_df["TotalNeededToday"].map(lambda v: f"£{v:,.0f}")
    table_df["ISANeededToday"] = table_df["ISANeededToday"].map(lambda v: f"£{v:,.0f}")
    table_df["PensionNeededToday"] = table_df["PensionNeededToday"].map(lambda v: f"£{v:,.0f}")
    cell_text = table_df.values.tolist()
    col_labels = list(table_df.columns)
    tbl = ax.table(cellText=cell_text, colLabels=col_labels, cellLoc="center", loc="center")
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(8.5)
    tbl.scale(1, 2.0)
    ax.set_title("Capital required (today)")

    fig.suptitle("Retirement scenario set")
    fig.tight_layout(rect=[0, 0, 1, 0.96])
    fig.savefig(out_png, dpi=200)
    plt.close(fig)

    return summary


# ----------------------------
# Export helpers
# ----------------------------

def timestamp_tag() -> str:
    return _dt.datetime.now().strftime("%Y%m%d_%H%M%S")


def export_run(out_dir: str, inputs: Inputs, results: Results, projection: pd.DataFrame, meta: pd.DataFrame,
               export_charts: bool = True) -> Dict[str, str]:
    os.makedirs(out_dir, exist_ok=True)
    tag = timestamp_tag()

    projection_csv = os.path.join(out_dir, f"projection_{tag}.csv")
    meta_csv = os.path.join(out_dir, f"meta_{tag}.csv")
    png = os.path.join(out_dir, f"projection_{tag}.png")

    projection.to_csv(projection_csv, index=False)
    meta.to_csv(meta_csv, index=False)

    if export_charts:
        plot_single_projection(inputs, results, projection, png)

    return {"projection_csv": projection_csv, "meta_csv": meta_csv, "png": png}


def export_scenario_set(out_dir: str, base_inputs: Inputs) -> Dict[str, str]:
    os.makedirs(out_dir, exist_ok=True)
    tag = timestamp_tag()

    png = os.path.join(out_dir, f"scenario_set_{tag}.png")
    summary_csv = os.path.join(out_dir, f"scenario_set_{tag}.csv")

    summary = plot_scenario_set(base_inputs, png)
    summary.to_csv(summary_csv, index=False)

    return {"scenario_png": png, "scenario_csv": summary_csv}


# ----------------------------
# GUI
# ----------------------------

class RetirementGUI(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Retirement Calculator (ISA bridge + pension) — deterministic")
        self.geometry("980x760")

        self.out_dir = tk.StringVar(value=os.path.abspath(os.getcwd()))

        # Variables
        self.vars: Dict[str, tk.Variable] = {
            "current_age": tk.IntVar(value=51),
            "pension_access_age": tk.IntVar(value=57),
            "state_pension_age": tk.IntVar(value=67),
            "life_expectancy": tk.IntVar(value=85),

            "annual_spending_today": tk.DoubleVar(value=20000),
            "inflation_rate": tk.DoubleVar(value=1.5),  # percent
            "nominal_return": tk.DoubleVar(value=5.0),  # percent

            "include_state_pension": tk.BooleanVar(value=True),
            "state_pension_annual_today": tk.DoubleVar(value=11502),

            "sizing_method": tk.StringVar(value="SWR"),
            "swr": tk.DoubleVar(value=4.0),  # percent

            "current_isa": tk.DoubleVar(value=0),
            "current_pension": tk.DoubleVar(value=0),
        }

        self._build()

    def _build(self):
        # Top: output directory
        top = ttk.Frame(self)
        top.pack(fill="x", padx=12, pady=10)

        ttk.Label(top, text="Output folder:").pack(side="left")
        ttk.Entry(top, textvariable=self.out_dir, width=80).pack(side="left", padx=8)
        ttk.Button(top, text="Browse…", command=self._browse_out).pack(side="left")

        # Main: left inputs, right outputs
        main = ttk.Frame(self)
        main.pack(fill="both", expand=True, padx=12, pady=10)

        left = ttk.Frame(main)
        left.pack(side="left", fill="both", expand=True)

        right = ttk.Frame(main)
        right.pack(side="right", fill="both", expand=True, padx=(12, 0))

        # Inputs sections
        self._section_personal(left)
        self._section_financial(left)
        self._section_pension(left)
        self._section_strategy(left)
        self._section_holdings(left)

        # Buttons
        btns = ttk.Frame(left)
        btns.pack(fill="x", pady=10)
        ttk.Button(btns, text="Run single scenario (export CSV + PNG)", command=self._run_single).pack(fill="x", pady=4)
        ttk.Button(btns, text="Run scenario set (export comparison PNG + CSV)", command=self._run_set).pack(fill="x", pady=4)

        # Output box
        ttk.Label(right, text="Results:").pack(anchor="w")
        self.output = tk.Text(right, height=36, wrap="word")
        self.output.pack(fill="both", expand=True)
        self._write("Adjust inputs on the left, then run a scenario.\n\n")

        # Footer notes
        note = ttk.Label(
            right,
            text=(
                "Notes: Deterministic model (constant return). 'SWR' sizing is a rule-of-thumb for robustness, "
                "so deterministic projections often show a surplus."
            ),
            wraplength=420,
        )
        note.pack(anchor="w", pady=(8, 0))

    def _write(self, s: str):
        self.output.insert("end", s)
        self.output.see("end")

    def _browse_out(self):
        d = filedialog.askdirectory(initialdir=self.out_dir.get())
        if d:
            self.out_dir.set(d)

    def _section_personal(self, parent):
        frm = ttk.LabelFrame(parent, text="Ages (years)")
        frm.pack(fill="x", pady=6)

        self._row(frm, "Current age", "current_age", "Your age today")
        self._row(frm, "Pension access age", "pension_access_age", "Earliest age you can draw from pension")
        self._row(frm, "State pension age", "state_pension_age", "Age State Pension starts")
        self._row(frm, "Life expectancy", "life_expectancy", "Plan end age")

    def _section_financial(self, parent):
        frm = ttk.LabelFrame(parent, text="Spending & returns")
        frm.pack(fill="x", pady=6)

        self._row(frm, "Annual spending (today, £)", "annual_spending_today", "Your target spending in today's pounds")
        self._row(frm, "Inflation (%/yr)", "inflation_rate", "Spending increases by this each year")
        self._row(frm, "Return (%/yr)", "nominal_return", "Nominal portfolio return (constant, deterministic)")

    def _section_pension(self, parent):
        frm = ttk.LabelFrame(parent, text="State Pension")
        frm.pack(fill="x", pady=6)

        chk = ttk.Checkbutton(frm, text="Include State Pension", variable=self.vars["include_state_pension"])
        chk.grid(row=0, column=0, columnspan=3, sticky="w", padx=8, pady=(6, 2))

        self._row(frm, "State Pension (today, £/yr)", "state_pension_annual_today", "Amount in today's pounds (inflation-linked here)", start_row=1)

    def _section_strategy(self, parent):
        frm = ttk.LabelFrame(parent, text="Sizing method")
        frm.pack(fill="x", pady=6)

        ttk.Label(frm, text="Method").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        cmb = ttk.Combobox(frm, textvariable=self.vars["sizing_method"], values=["SWR", "PV"], width=10, state="readonly")
        cmb.grid(row=0, column=1, sticky="w", pady=6)
        ttk.Label(frm, text="SWR (% if method=SWR)").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(frm, textvariable=self.vars["swr"], width=12).grid(row=1, column=1, sticky="w", pady=6)
        ttk.Label(frm, text="e.g., 4.0 for 4% or 3.5 for 3.5%.").grid(row=1, column=2, sticky="w", padx=8)

        ttk.Label(frm, text="Descriptions:").grid(row=2, column=0, sticky="nw", padx=8, pady=(6, 8))
        desc = (
            "SWR: sizes the pension pot at access using a withdrawal-rate multiple (rule of thumb).\n"
            "PV: sizes pots using present value of inflation-linked spending to your end age."
        )
        ttk.Label(frm, text=desc, justify="left").grid(row=2, column=1, columnspan=2, sticky="w", pady=(6, 8))

    def _section_holdings(self, parent):
        frm = ttk.LabelFrame(parent, text="Your current holdings (optional, used for gaps)")
        frm.pack(fill="x", pady=6)

        self._row(frm, "Current ISA / accessible (£)", "current_isa", "Liquid funds available before pension access")
        self._row(frm, "Current pension (£)", "current_pension", "Pension value today", start_row=1)

    def _row(self, frm, label, varname, help_text, start_row=None):
        r = start_row if start_row is not None else frm.grid_size()[1]
        ttk.Label(frm, text=label).grid(row=r, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(frm, textvariable=self.vars[varname], width=16).grid(row=r, column=1, sticky="w", pady=6)
        ttk.Label(frm, text=help_text).grid(row=r, column=2, sticky="w", padx=8)

    def _collect_inputs(self) -> Inputs:
        # Convert percent fields
        infl = self.vars["inflation_rate"].get() / 100.0
        ret = self.vars["nominal_return"].get() / 100.0
        swr = self.vars["swr"].get() / 100.0

        return Inputs(
            current_age=int(self.vars["current_age"].get()),
            pension_access_age=int(self.vars["pension_access_age"].get()),
            state_pension_age=int(self.vars["state_pension_age"].get()),
            life_expectancy=int(self.vars["life_expectancy"].get()),

            annual_spending_today=float(self.vars["annual_spending_today"].get()),

            inflation_rate=float(infl),
            nominal_return=float(ret),

            include_state_pension=bool(self.vars["include_state_pension"].get()),
            state_pension_annual_today=float(self.vars["state_pension_annual_today"].get()),

            sizing_method=str(self.vars["sizing_method"].get()),
            swr=float(swr),

            current_isa=float(self.vars["current_isa"].get()),
            current_pension=float(self.vars["current_pension"].get()),
        )

    def _run_single(self):
        self._write("\n--- Running single scenario ---\n")
        try:
            inp = self._collect_inputs()
            res, df, meta = size_capital(inp)

            paths = export_run(self.out_dir.get(), inp, res, df, meta, export_charts=True)

            self._write(f"Sizing method: {inp.sizing_method}\n")
            if inp.sizing_method == "SWR":
                self._write(f"SWR: {inp.swr:.2%}\n")
            self._write(f"Real return: {res.real_return:.2%}\n")
            self._write(f"ISA needed today: £{res.isa_needed_today:,.0f}\n")
            self._write(f"Pension needed today: £{res.pension_needed_today:,.0f}\n")
            self._write(f"Total needed today: £{res.total_needed_today:,.0f}\n")
            self._write(f"Mathematical-min total (PV) today: £{res.mathematical_min_total_today:,.0f}\n")

            if inp.current_isa or inp.current_pension:
                gap_isa = res.isa_needed_today - inp.current_isa
                gap_pen = res.pension_needed_today - inp.current_pension
                self._write(f"\nGap vs your ISA: £{gap_isa:,.0f}\n")
                self._write(f"Gap vs your pension: £{gap_pen:,.0f}\n")
                self._write(f"Gap total: £{(res.total_needed_today - (inp.current_isa+inp.current_pension)):,.0f}\n")

            # sanity
            min_bal = df["Total_Balance"].min()
            self._write(f"\nMinimum total balance during projection: £{min_bal:,.0f}\n")

            self._write("\nExported files:\n")
            for k, v in paths.items():
                self._write(f"  {k}: {v}\n")

        except Exception as e:
            messagebox.showerror("Error", str(e))
            self._write(f"ERROR: {e}\n")

    def _run_set(self):
        self._write("\n--- Running scenario set ---\n")
        try:
            base = self._collect_inputs()
            paths = export_scenario_set(self.out_dir.get(), base)
            self._write("Exported files:\n")
            for k, v in paths.items():
                self._write(f"  {k}: {v}\n")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self._write(f"ERROR: {e}\n")


def main():
    app = RetirementGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
