"""
app.py
Streamlit front-end for the retirement planner.
- Sidebar: all inputs
- Main area: base scenario, scenario set charts + tables
- Single "Save all" button: downloads ZIP with CSVs, PNGs, DOCX.
"""
import io
import tempfile
from pathlib import Path
import streamlit as st
from retirement_core import (
    Inputs,
    size_capital,
    plot_single_projection,
    plot_scenario_set,
    create_pension_report,
)

st.set_page_config(
    page_title="Retirement Planner – ISA Bridge + Pension",
    layout="wide",
)

st.title("Retirement Planner – ISA Bridge + Pension")

# ----------------------------
# Sidebar inputs
# ----------------------------
st.sidebar.header("Assumptions")
current_age = st.sidebar.number_input("Current age", 18, 100, 51)
pension_access_age = st.sidebar.number_input("Pension access age", 18, 100, 57)
state_pension_age = st.sidebar.number_input("State pension age", 18, 100, 67)
life_expectancy = st.sidebar.number_input("Life expectancy", 50, 110, 85)

annual_spending_today = st.sidebar.number_input(
    "Annual spending (today, £)", 0.0, 1_000_000.0, 20_000.0, step=500.0
)

inflation_rate = (
    st.sidebar.number_input("Inflation (%/yr)", 0.0, 20.0, 1.5, step=0.1) / 100.0
)

nominal_return = (
    st.sidebar.number_input("Return (%/yr)", -50.0, 50.0, 5.0, step=0.1) / 100.0
)

include_state_pension = st.sidebar.checkbox("Include State Pension", True)
state_pension_annual_today = st.sidebar.number_input(
    "State Pension (today, £/yr)", 0.0, 50_000.0, 11_502.0, step=100.0
)

sizing_method = st.sidebar.selectbox("Sizing method", ["SWR", "PV"])
swr = st.sidebar.number_input(
    "SWR (% if method=SWR)", 0.1, 20.0, 4.0, step=0.1
) / 100.0

# ----------------------------
# Adjustments Section
# ----------------------------
st.sidebar.header("Adjustments")

# Allowing current holding fields to be left blank
current_isa_val = st.sidebar.number_input("Current ISA / accessible (£)", 0.0, 10_000_000.0, value=None, step=1_000.0)
current_pension_val = st.sidebar.number_input("Current pension (£)", 0.0, 10_000_000.0, value=None, step=1_000.0)

desired_pension_at_end = st.sidebar.number_input(
    "Longevity buffer (£, today's money)", 0.0, 10_000_000.0, 0.0, step=10_000.0
)
st.sidebar.markdown("_Set to £0 for minimum capital, or e.g. £100,000 to protect against living longer._")

st.sidebar.subheader("Scenario Triggers")
adj_swr = st.sidebar.number_input("SWR adjustment (%)", 0.1, 20.0, 3.5, step=0.1) / 100.0
adj_inflation = st.sidebar.number_input("Inflation adjustment (+%)", 0.0, 10.0, 0.5, step=0.1) / 100.0
adj_return = st.sidebar.number_input("Returns adjustment (+%)", -10.0, 10.0, 0.5, step=0.1) / 100.0

inputs = Inputs(
    current_age=current_age,
    pension_access_age=pension_access_age,
    state_pension_age=state_pension_age,
    life_expectancy=life_expectancy,
    annual_spending_today=annual_spending_today,
    inflation_rate=inflation_rate,
    nominal_return=nominal_return,
    include_state_pension=include_state_pension,
    state_pension_annual_today=state_pension_annual_today,
    sizing_method=sizing_method,
    swr=swr,
    current_isa=current_isa_val if current_isa_val is not None else 0.0,
    current_pension=current_pension_val if current_pension_val is not None else 0.0,
    desired_pension_at_end=desired_pension_at_end,
    adj_swr=adj_swr,
    adj_inflation=adj_inflation,
    adj_return=adj_return,
)

# ----------------------------
# Compute base + scenario set
# ----------------------------
results, df, meta = size_capital(inputs)

col1, col2, col3 = st.columns(3)
with col1:
    st.metric("ISA needed today", f"£{results.isa_needed_today:,.0f}")
with col2:
    st.metric("Pension needed today", f"£{results.pension_needed_today:,.0f}")
with col3:
    st.metric("Total needed today", f"£{results.total_needed_today:,.0f}")

final_pension = df.iloc[-1]["Pension_Balance"]
st.caption(
    f"Final pension balance at age {inputs.life_expectancy}: £{final_pension:,.0f}"
)

# Create temporary directory for charts
tmpdir = tempfile.TemporaryDirectory()
tmp_path = Path(tmpdir.name)
single_png = tmp_path / "projection.png"
scenario_png = tmp_path / "scenario_set.png"

plot_single_projection(inputs, results, df, str(single_png))
scenario_summary = plot_scenario_set(inputs, str(scenario_png))

# ----------------------------
# Display charts and tables
# ----------------------------
st.subheader("Base scenario projection")
st.image(str(single_png))

st.subheader("Scenario comparison")
st.image(str(scenario_png))

st.subheader("Projection table (base scenario)")
st.dataframe(df.style.format(precision=2))

st.subheader("Scenario summary table")
st.dataframe(scenario_summary.style.format(precision=2))

@st.dialog("Quick Start Guide")
def show_help_dialog():
    st.markdown('''
    ### Quick Start Guide – Retirement Planner (ISA Bridge + Pension)
    This app helps you size how much capital you need in ISA / accessible money and pension to fund retirement, under simple, deterministic assumptions.
    
    #### 1. Ages
    * **Current age:** Your age today. The projection starts here and runs to life expectancy.
    * **Pension access age:** Earliest age you can draw from your pension (e.g. 57).
    * **State pension age:** Age you start receiving State Pension.
    * **Life expectancy:** Planning horizon (e.g. 85).
    
    #### 2. Spending and returns
    * **Annual spending:** Your desired annual spending in today's money.
    * **Inflation:** How fast your spending rises over time.
    * **Return:** Expected long-run nominal investment return.
    
    #### 3. State Pension
    * **Include State Pension:** If ticked, State Pension reduces how much you must withdraw from your pots.
    
    #### 4. Sizing method: SWR vs PV
    * **Safe Withdrawal Rate (SWR):** Sizes the pot so withdrawals equal your chosen SWR. Provides a probabilistic safety margin.
    * **Present Value (PV):** Calculates the mathematical minimum required assuming constant returns.
    
    #### 5. Adjustments
    * **Current Holdings:** Input your actual ISA and Pension balances. If provided, the main chart plots an "Actual Total" projection line showing when these actual pots might run out.
    * **Longevity Buffer:** Sets a target residual balance at life expectancy to protect against living longer.
    * **Scenario Triggers:** Override the default adjustments for the sensitivity analysis (Conservative SWR, Inflation, Return).
    ''')

# ----------------------------
# Build ZIP for download
# ----------------------------
def build_full_report_zip() -> bytes:
    """Create an in-memory ZIP with CSVs, PNGs and DOCX."""
    # CSV bytes
    proj_csv_bytes = df.to_csv(index=False).encode("utf-8")
    meta_csv_bytes = meta.to_csv(index=False).encode("utf-8")
    scenario_csv_bytes = scenario_summary.to_csv(index=False).encode("utf-8")

    # DOCX into BytesIO
    docx_buffer = io.BytesIO()
    create_pension_report(
        out_docx=docx_buffer,
        inputs=inputs,
        results=results,
        df=df,
        chart_png=str(single_png),
        scenario_png=str(scenario_png),
    )
    docx_bytes = docx_buffer.getvalue()

    # Build ZIP
    zip_buffer = io.BytesIO()
    import zipfile

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("projection.csv", proj_csv_bytes)
        zf.writestr("meta.csv", meta_csv_bytes)
        zf.writestr("scenario_set.csv", scenario_csv_bytes)
        zf.write(single_png, "projection.png")
        zf.write(scenario_png, "scenario_set.png")
        zf.writestr("pension_report.docx", docx_bytes)

    zip_buffer.seek(0)
    return zip_buffer.getvalue()

st.sidebar.markdown("---")
st.sidebar.subheader("Export & Help")

if st.sidebar.button("Prepare full report bundle"):
    zip_bytes = build_full_report_zip()
    st.sidebar.download_button(
        label="Download all (ZIP)",
        data=zip_bytes,
        file_name="pension_report_bundle.zip",
        mime="application/zip",
    )

if st.sidebar.button("Help / Quick Start Guide"):
    show_help_dialog()
