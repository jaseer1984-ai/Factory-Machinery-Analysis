# Enhanced app.py — Paper Cup Factory Dashboard
# Additional Features:
# - Maintenance scheduling integration
# - Seasonality adjustments
# - Risk analysis and scenarios
# - Performance benchmarking
# - Alert system for aging machines

from datetime import datetime, timedelta
import io
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

st.set_page_config(page_title="Paper Cup Factory Dashboard", layout="wide")
st.title("📈 Paper Cup Factory — Advanced Production & Sales Forecast Dashboard")

# ---------------- Sidebar ----------------
with st.sidebar:
    st.header("⚙️ Settings")
    uploaded = st.file_uploader("Upload factory Excel", type=["xlsx"])

    days_per_month = st.number_input("Days per month", 1, 31, value=28)
    hour_scenarios = st.multiselect(
        "Hours per day (select)", options=[12, 16, 20, 24], default=[12, 16, 20, 24]
    )
    machine_life_years = st.number_input("Machine life (years)", 1, 40, value=10)
    forecast_years = st.number_input("Forecast horizon (years)", 1, 30, value=10)

    st.divider()
    st.subheader("📊 Advanced Settings")
    
    # Seasonality factors
    st.markdown("**Seasonality Adjustments**")
    seasonality_enabled = st.checkbox("Enable seasonal demand patterns")
    if seasonality_enabled:
        peak_months = st.multiselect("Peak demand months", 
                                   options=list(range(1, 13)), 
                                   default=[11, 12],
                                   format_func=lambda x: datetime(2023, x, 1).strftime('%B'))
        seasonality_factor = st.slider("Peak season multiplier", 1.0, 2.0, 1.3, 0.1)
    
    # Risk scenarios
    st.markdown("**Risk Analysis**")
    risk_analysis = st.checkbox("Enable risk scenarios")
    if risk_analysis:
        machine_failure_rate = st.slider("Annual machine failure rate (%)", 0.0, 20.0, 5.0, 0.5)
        supply_disruption_risk = st.slider("Supply chain disruption risk (%)", 0.0, 30.0, 10.0, 1.0)

    st.divider()
    st.subheader("💰 Economics")
    unit_price = st.number_input("Unit price (per cup)", min_value=0.0, value=0.0, step=0.001, format="%.4f")
    unit_cost  = st.number_input("Unit cost (per cup)",  min_value=0.0, value=0.0, step=0.001, format="%.4f")
    capex_per_machine = st.number_input("CAPEX per machine", min_value=0.0, value=0.0, step=1000.0, format="%.0f")
    
    # Maintenance costs
    annual_maintenance_per_machine = st.number_input("Annual maintenance per machine", 
                                                   min_value=0.0, value=5000.0, step=500.0)

# ------------- Enhanced Helpers -------------

def find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Return the ACTUAL column name that matches any candidate (case/space tolerant)."""
    norm = {str(c).strip().lower(): c for c in df.columns}
    # exact
    for want in candidates:
        key = str(want).strip().lower()
        if key in norm:
            return norm[key]
    # contains
    for want in candidates:
        key = str(want).strip().lower()
        for k, real in norm.items():
            if key in k:
                return real
    return None

def parse_date_series(s: pd.Series):
    def parse_one(v):
        if pd.isna(v): return pd.NaT
        if isinstance(v, (pd.Timestamp, datetime)): return pd.to_datetime(v)
        v = str(v)
        fmts = ("%Y-%m-%d","%d-%m-%Y","%d/%m/%Y","%m/%d/%Y","%d-%b-%Y","%b-%d-%Y","%Y/%m/%d")
        for fmt in fmts:
            try:
                return datetime.strptime(v, fmt)
            except:
                pass
        try:
            return pd.to_datetime(v, errors="coerce", dayfirst=True)
        except:
            return pd.NaT
    return pd.to_datetime(s.apply(parse_one), errors="coerce")

@st.cache_data(show_spinner=False)
def load_workbook(file_bytes: bytes) -> dict:
    xls = pd.ExcelFile(file_bytes)
    sheets = {s: pd.read_excel(xls, sheet_name=s) for s in xls.sheet_names}
    for s, df in sheets.items():
        df.columns = [str(c).strip() for c in df.columns]
    return sheets

def detect_key_sheets(sheets: dict):
    names = list(sheets.keys())
    cand_m = [s for s in names if any(k in s.lower() for k in
             ["machine","assets","equipment","plant","capacity","list","details"])]
    machines_sheet = cand_m[0] if cand_m else names[0]
    cand_p = [s for s in names if any(k in s.lower() for k in
             ["prod","output","shift","daily","util","run","report","sales"]) and s != machines_sheet]
    production_sheet = cand_p[0] if cand_p else (names[1] if len(names)>1 else names[0])
    return machines_sheet, production_sheet

def extract_core(machines_df: pd.DataFrame, prod_df: pd.DataFrame, life_years: float):
    # Machine identifier
    machine_col = find_col(machines_df, ["machine","machine id","machine_name","name","id"]) or "Machine"
    if machine_col not in machines_df.columns:
        machines_df["Machine"] = [f"M{i+1}" for i in range(len(machines_df))]
        machine_col = "Machine"

    # Capacity (cups/min)
    cap_col = find_col(machines_df, [
        "capacity","rated capacity","rated_capacity","cups/min","cups per min",
        "capacity (cups/min)","capacity_cups_min","capacity_cup_min"
    ])
    if cap_col is None:
        machines_df["Rated_Capacity_cpm"] = 75.0
        cap_col = "Rated_Capacity_cpm"
    else:
        cap_vals = pd.to_numeric(machines_df[cap_col], errors="coerce")
        if cap_vals.isna().all():
            cap_vals = pd.Series([75.0]*len(machines_df))
        machines_df["Rated_Capacity_cpm"] = cap_vals.fillna(cap_vals.median())
        cap_col = "Rated_Capacity_cpm"

    # Utilization
    util_col = find_col(machines_df, ["utilization","util","uptime %","uptime","availability","runtime%"])
    if util_col is None:
        util_p = find_col(prod_df, ["utilization","util","uptime %","uptime","availability","runtime%"])
        if util_p is not None and not prod_df.empty:
            u = pd.to_numeric(prod_df[util_p].astype(str).str.replace("%","",regex=False), errors="coerce")/100.0
            machines_df["Utilization"] = float(u.fillna(u.mean()).clip(0,1).mean())
        else:
            machines_df["Utilization"] = 0.955
    else:
        u = machines_df[util_col]
        if u.dtype == object:
            u = pd.to_numeric(u.astype(str).str.replace("%","",regex=False), errors="coerce")/100.0
        machines_df["Utilization"] = u.fillna(u.mean()).clip(0,1)

    # ----- Robust Age detection -----
    today = pd.Timestamp.today().normalize()

    age_col  = find_col(machines_df, ["age","age (years)","age_years"])
    start_col = find_col(machines_df, [
        "start date","commission date","commission_date","installed on","install date",
        "start_date","commissioning date","commission year","commission_year"
    ])
    year_col = find_col(machines_df, [
        "year","mfg year","manufacture year","manufacturing year","year of make",
        "purchase year","installation year","install year","yom","yop"
    ])

    if age_col and age_col in machines_df.columns:
        machines_df["Age_years"] = pd.to_numeric(machines_df[age_col], errors="coerce")
        machines_df["Start_Date"] = today - pd.to_timedelta((machines_df["Age_years"]*365.25), unit="D")
    elif start_col and start_col in machines_df.columns:
        machines_df["Start_Date"] = parse_date_series(machines_df[start_col])
        age_days = (today - machines_df["Start_Date"]).dt.days
        machines_df["Age_years"] = (age_days/365.25)
    elif year_col and year_col in machines_df.columns:
        yr = pd.to_numeric(machines_df[year_col], errors="coerce")
        machines_df["Start_Date"] = pd.to_datetime(yr.round().astype("Int64").astype(str) + "-01-01", errors="coerce")
        age_days = (today - machines_df["Start_Date"]).dt.days
        machines_df["Age_years"] = (age_days/365.25)
    else:
        machines_df["Start_Date"] = pd.NaT
        machines_df["Age_years"] = np.nan

    machines_df["Age_years"] = machines_df["Age_years"].round(2)
    machines_df["Remaining_Life_years"] = (life_years - machines_df["Age_years"]).round(2)
    machines_df.loc[machines_df["Remaining_Life_years"] < 0, "Remaining_Life_years"] = 0
    machines_df["End_of_Life_Year"] = machines_df["Start_Date"].dt.year + life_years
    machines_df.loc[machines_df["Start_Date"].isna(), "End_of_Life_Year"] = np.nan

    # Add maintenance scheduling
    machines_df["Next_Major_Maintenance"] = machines_df["Start_Date"] + pd.DateOffset(years=2)
    machines_df["Maintenance_Due"] = (machines_df["Next_Major_Maintenance"] < pd.Timestamp.today() + pd.DateOffset(months=6))

    return machine_col, cap_col, machines_df

def scenario_outputs(machines_df: pd.DataFrame, hours: list[int], dpm: int):
    util = float(machines_df["Utilization"].mean())
    base_cpm = machines_df["Rated_Capacity_cpm"].sum()
    rows = []
    for h in hours:
        cups_day = base_cpm * 60 * h * util
        cups_month = cups_day * dpm
        rows.append({"Hours": h, "Daily_Output_cups": cups_day, "Monthly_Output_cups": cups_month})
    return pd.DataFrame(rows, columns=["Hours","Daily_Output_cups","Monthly_Output_cups"])

def enhanced_annual_forecast(machines_df: pd.DataFrame, hours: list[int], dpm: int, years: int, 
                           price=0.0, cost=0.0, maintenance_cost=0.0, 
                           seasonality_enabled=False, peak_months=None, seasonality_factor=1.0,
                           risk_analysis=False, failure_rate=0.0):
    util = float(machines_df["Utilization"].mean())
    base_cpm = machines_df["Rated_Capacity_cpm"].sum()
    current_year = pd.Timestamp.today().year
    ylist = [current_year + i for i in range(years)]
    tables = {}
    
    for h in hours:
        rows = []
        for y in ylist:
            # Base calculation
            cups_day = base_cpm * 60 * h * util
            cups_year = cups_day * dpm * 12
            
            # Apply seasonality if enabled
            if seasonality_enabled and peak_months:
                seasonal_adjustment = 1.0 + (seasonality_factor - 1.0) * (len(peak_months) / 12)
                cups_year *= seasonal_adjustment
            
            # Apply risk factors
            if risk_analysis and failure_rate > 0:
                availability_factor = 1.0 - (failure_rate / 100.0)
                cups_year *= availability_factor
            
            row = {"Year": y, "Hours_per_Day": h, "Output_cups": cups_year}
            
            if price > 0 and cost >= 0:
                row["Revenue"] = cups_year * price
                row["Gross_Margin"] = cups_year * (price - cost)
                if maintenance_cost > 0:
                    total_maintenance = len(machines_df) * maintenance_cost
                    row["Net_Margin"] = row["Gross_Margin"] - total_maintenance
            
            rows.append(row)
        tables[h] = pd.DataFrame(rows, columns=list(rows[0].keys()))
    return tables

def generate_maintenance_schedule(machines_df: pd.DataFrame, years: int):
    """Generate preventive maintenance schedule"""
    current_year = pd.Timestamp.today().year
    schedule = []
    
    for _, machine in machines_df.iterrows():
        machine_id = machine.get("Machine", f"M{machine.name}")
        start_date = machine.get("Start_Date")
        
        if pd.notna(start_date):
            # Schedule major maintenance every 2 years
            for year_offset in range(0, years):
                maintenance_year = current_year + year_offset
                maintenance_date = pd.Timestamp(maintenance_year, 6, 1)  # Mid-year maintenance
                
                schedule.append({
                    "Machine": machine_id,
                    "Year": maintenance_year,
                    "Maintenance_Type": "Major" if year_offset % 2 == 0 else "Minor",
                    "Estimated_Cost": 15000 if year_offset % 2 == 0 else 5000,
                    "Downtime_Days": 5 if year_offset % 2 == 0 else 2
                })
    
    return pd.DataFrame(schedule)

def create_risk_scenario_analysis(base_output: float, scenarios: dict):
    """Create risk scenario analysis"""
    results = []
    for scenario_name, factors in scenarios.items():
        adjusted_output = base_output
        for factor_name, factor_value in factors.items():
            adjusted_output *= factor_value
        
        results.append({
            "Scenario": scenario_name,
            "Output_Multiplier": adjusted_output / base_output,
            "Annual_Output": adjusted_output,
            "Variance_from_Base": ((adjusted_output - base_output) / base_output) * 100
        })
    
    return pd.DataFrame(results)

def capex_schedule(machines_df: pd.DataFrame, years: int, life_years: float, capex_each: float = 0.0):
    current_year = pd.Timestamp.today().year
    ylist = [current_year + i for i in range(years)]
    counts = {y: 0 for y in ylist}
    for _, r in machines_df.iterrows():
        eol = r["End_of_Life_Year"]
        if pd.notna(eol):
            eol = int(eol)
            if eol in counts:
                counts[eol] += 1
    df = pd.DataFrame({"Year": ylist, "Machines_to_Replace": [counts[y] for y in ylist]})
    if capex_each > 0:
        df["CAPEX"] = df["Machines_to_Replace"] * capex_each
    return df

def to_excel_bytes(sheets: dict) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=name[:31] or "Sheet")
    return output.getvalue()

# String format helpers
def fmt_int(x):
    if x is None or (isinstance(x, float) and np.isnan(x)): return ""
    return f"{int(round(x)):,}"

def fmt_money(x, decimals=2):
    if x is None or (isinstance(x, float) and np.isnan(x)): return ""
    return f"{x:,.{decimals}f}"

# ------------- Main Application -------------
if uploaded is None:
    st.info("🔧 Upload your Excel file to begin comprehensive factory analysis")
    st.markdown("""
    ### Expected Excel Structure:
    - **Machines Sheet**: Contains machine details (ID, capacity, age/start date, utilization)
    - **Production Sheet**: Historical production data (optional utilization rates)
    
    ### New Features:
    - 📊 **Seasonality Analysis**: Account for peak/off-peak demand patterns
    - ⚠️ **Risk Scenarios**: Model machine failures and supply disruptions
    - 🔧 **Maintenance Scheduling**: Preventive maintenance planning
    - 📈 **Performance Benchmarking**: Compare against industry standards
    """)
    st.stop()

# Load and process data
wb = load_workbook(uploaded)
machines_sheet, production_sheet = detect_key_sheets(wb)
st.markdown(f"**Detected Sheets:** `{machines_sheet}` (machines) | `{production_sheet}` (production)")

machines_df_raw = wb[machines_sheet].copy()
prod_df_raw = wb[production_sheet].copy()
machine_col, cap_col, machines_df = extract_core(machines_df_raw.copy(), prod_df_raw.copy(), machine_life_years)

# Enhanced KPI Dashboard
st.subheader("📊 Factory Performance Dashboard")
col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    st.metric("Total Machines", f"{len(machines_df):,}")
with col2:
    avg_util = machines_df['Utilization'].mean()
    st.metric("Avg Utilization", f"{avg_util:.2%}", 
              delta=f"{(avg_util - 0.85):.1%}" if avg_util > 0.85 else None)
with col3:
    total_capacity = machines_df['Rated_Capacity_cpm'].sum()
    st.metric("Total Capacity", f"{total_capacity:,.0f} cups/min")
with col4:
    aging_machines = (machines_df["Remaining_Life_years"] < 2).sum()
    st.metric("Aging Machines", f"{aging_machines:,}", 
              delta=f"-{aging_machines}" if aging_machines > 0 else "0", delta_color="inverse")
with col5:
    avg_age = machines_df["Age_years"].dropna().mean()
    st.metric("Avg Age", "N/A" if pd.isna(avg_age) else f"{avg_age:.1f} years")

# Alert system for critical machines
if aging_machines > 0:
    st.warning(f"⚠️ **Alert**: {aging_machines} machines require attention within 2 years!")

# Maintenance due alerts
maintenance_due = machines_df["Maintenance_Due"].sum() if "Maintenance_Due" in machines_df.columns else 0
if maintenance_due > 0:
    st.info(f"🔧 **Maintenance Notice**: {maintenance_due} machines due for major maintenance")

# Ensure hours selection
if not hour_scenarios:
    st.warning("No hours selected; defaulting to 12/16/20/24.")
    hour_scenarios = [12, 16, 20, 24]

# -------- Enhanced Scenario Analysis --------
scen_df = scenario_outputs(machines_df, hour_scenarios, days_per_month)
if "Hours" in scen_df.columns:
    scen_df = scen_df.sort_values("Hours")

st.subheader("🎯 Production Scenarios Analysis")

# Create enhanced visualization with multiple metrics
fig_scenarios = make_subplots(
    rows=1, cols=2,
    subplot_titles=("Daily Output by Hours", "Monthly Output Trend"),
    specs=[[{"secondary_y": False}, {"secondary_y": True}]]
)

fig_scenarios.add_trace(
    go.Bar(x=scen_df["Hours"], y=scen_df["Daily_Output_cups"], 
           name="Daily Output", text=scen_df["Daily_Output_cups"].round().astype(int).map(lambda v: f"{v:,}"),
           textposition="outside"),
    row=1, col=1
)

fig_scenarios.add_trace(
    go.Scatter(x=scen_df["Hours"], y=scen_df["Monthly_Output_cups"], 
               mode="lines+markers", name="Monthly Output",
               text=scen_df["Monthly_Output_cups"].round().astype(int).map(lambda v: f"{v:,}")),
    row=1, col=2
)

fig_scenarios.update_layout(height=400, showlegend=False)
st.plotly_chart(fig_scenarios, use_container_width=True)

# Display formatted table
disp_scen = scen_df.copy()
disp_scen["Daily Output (cups)"] = disp_scen["Daily_Output_cups"].apply(fmt_int)
disp_scen["Monthly Output (cups)"] = disp_scen["Monthly_Output_cups"].apply(fmt_int)
st.table(disp_scen.set_index("Hours")[["Daily Output (cups)", "Monthly Output (cups)"]])

# -------- Enhanced 10-Year Forecast --------
sales_enabled = unit_price > 0 and unit_cost >= 0

# Apply enhanced forecasting parameters
forecast_params = {
    'seasonality_enabled': seasonality_enabled if 'seasonality_enabled' in locals() else False,
    'peak_months': peak_months if 'seasonality_enabled' in locals() and seasonality_enabled else None,
    'seasonality_factor': seasonality_factor if 'seasonality_enabled' in locals() and seasonality_enabled else 1.0,
    'risk_analysis': risk_analysis if 'risk_analysis' in locals() else False,
    'failure_rate': machine_failure_rate if 'risk_analysis' in locals() and risk_analysis else 0.0
}

tables = enhanced_annual_forecast(
    machines_df, hour_scenarios, days_per_month, forecast_years,
    price=unit_price if sales_enabled else 0.0, 
    cost=unit_cost if sales_enabled else 0.0,
    maintenance_cost=annual_maintenance_per_machine,
    **forecast_params
)

st.subheader("📈 Enhanced 10-Year Forecast" + 
             (" with Seasonality & Risk Analysis" if forecast_params['seasonality_enabled'] or forecast_params['risk_analysis'] else ""))

tab_objs = st.tabs([f"{h}h/day" for h in hour_scenarios])
for tab, h in zip(tab_objs, hour_scenarios):
    with tab:
        # Create enhanced forecast visualization
        fig_enhanced = make_subplots(rows=2, cols=1, 
                                   subplot_titles=(f"Production Forecast - {h}h/day", 
                                                 "Financial Projection" if sales_enabled else ""),
                                   vertical_spacing=0.1)
        
        fig_enhanced.add_trace(
            go.Scatter(x=tables[h]["Year"], y=tables[h]["Output_cups"],
                      mode="lines+markers", name="Production Output",
                      line=dict(color="blue", width=3)),
            row=1, col=1
        )
        
        if sales_enabled and "Revenue" in tables[h].columns:
            fig_enhanced.add_trace(
                go.Scatter(x=tables[h]["Year"], y=tables[h]["Revenue"],
                          mode="lines+markers", name="Revenue",
                          line=dict(color="green", width=2)),
                row=2, col=1
            )
            
            if "Net_Margin" in tables[h].columns:
                fig_enhanced.add_trace(
                    go.Scatter(x=tables[h]["Year"], y=tables[h]["Net_Margin"],
                              mode="lines+markers", name="Net Margin",
                              line=dict(color="orange", width=2)),
                    row=2, col=1
                )
        
        fig_enhanced.update_layout(height=600)
        st.plotly_chart(fig_enhanced, use_container_width=True)
        
        # Display formatted table
        disp = tables[h].copy()
        disp["Output (cups)"] = disp["Output_cups"].apply(fmt_int)
        cols = ["Year", "Hours_per_Day", "Output (cups)"]
        
        if "Revenue" in disp.columns:
            disp["Revenue"] = disp["Revenue"].apply(lambda x: fmt_money(x, 0))
            disp["Gross_Margin"] = disp["Gross_Margin"].apply(lambda x: fmt_money(x, 0))
            cols += ["Revenue", "Gross_Margin"]
            
            if "Net_Margin" in disp.columns:
                disp["Net_Margin"] = disp["Net_Margin"].apply(lambda x: fmt_money(x, 0))
                cols += ["Net_Margin"]
        
        st.table(disp[cols].set_index("Year"))

# -------- Risk Scenario Analysis --------
if 'risk_analysis' in locals() and risk_analysis:
    st.subheader("⚠️ Risk Scenario Analysis")
    
    # Define risk scenarios
    risk_scenarios = {
        "Optimistic": {"machine_reliability": 1.05, "demand": 1.1, "supply_chain": 1.0},
        "Base Case": {"machine_reliability": 1.0, "demand": 1.0, "supply_chain": 1.0},
        "Pessimistic": {"machine_reliability": 0.9, "demand": 0.85, "supply_chain": 0.95},
        "Crisis": {"machine_reliability": 0.75, "demand": 0.7, "supply_chain": 0.8}
    }
    
    base_output = scen_df[scen_df["Hours"] == 16]["Monthly_Output_cups"].iloc[0] * 12  # Annual output at 16h
    risk_df = create_risk_scenario_analysis(base_output, risk_scenarios)
    
    col1, col2 = st.columns(2)
    with col1:
        fig_risk = px.bar(risk_df, x="Scenario", y="Variance_from_Base",
                         title="Output Variance by Risk Scenario (%)",
                         color="Variance_from_Base",
                         color_continuous_scale="RdYlGn_r")
        st.plotly_chart(fig_risk, use_container_width=True)
    
    with col2:
        st.markdown("**Risk Scenario Impact**")
        disp_risk = risk_df.copy()
        disp_risk["Annual Output"] = disp_risk["Annual_Output"].apply(fmt_int)
        disp_risk["Variance (%)"] = disp_risk["Variance_from_Base"].apply(lambda x: f"{x:+.1f}%")
        st.table(disp_risk.set_index("Scenario")[["Annual Output", "Variance (%)"]])

# -------- Maintenance Scheduling --------
st.subheader("🔧 Preventive Maintenance Schedule")
maintenance_schedule = generate_maintenance_schedule(machines_df, forecast_years)

if not maintenance_schedule.empty:
    # Group by year for summary
    yearly_maintenance = maintenance_schedule.groupby("Year").agg({
        "Estimated_Cost": "sum",
        "Downtime_Days": "sum",
        "Machine": "count"
    }).reset_index()
    yearly_maintenance.columns = ["Year", "Total_Cost", "Total_Downtime_Days", "Maintenance_Events"]
    
    col1, col2 = st.columns(2)
    with col1:
        fig_maint_cost = px.bar(yearly_maintenance, x="Year", y="Total_Cost",
                               title="Annual Maintenance Costs",
                               text=yearly_maintenance["Total_Cost"].apply(lambda x: f"${x:,.0f}"))
        st.plotly_chart(fig_maint_cost, use_container_width=True)
    
    with col2:
        fig_maint_events = px.line(yearly_maintenance, x="Year", y="Maintenance_Events",
                                  markers=True, title="Maintenance Events per Year")
        st.plotly_chart(fig_maint_events, use_container_width=True)
    
    # Display maintenance schedule table
    disp_maint = yearly_maintenance.copy()
    disp_maint["Total Cost"] = disp_maint["Total_Cost"].apply(lambda x: fmt_money(x, 0))
    disp_maint["Events"] = disp_maint["Maintenance_Events"].apply(fmt_int)
    disp_maint["Downtime (days)"] = disp_maint["Total_Downtime_Days"].apply(fmt_int)
    st.table(disp_maint.set_index("Year")[["Events", "Total Cost", "Downtime (days)"]])

# -------- CAPEX Replacement Plan --------
st.subheader("💰 CAPEX Replacement Schedule & Investment Planning")

capex_df = capex_schedule(machines_df, forecast_years, machine_life_years,
                          capex_each=capex_per_machine if capex_per_machine > 0 else 0.0)

# Enhanced CAPEX visualization with cumulative investment
col1, col2 = st.columns(2)

with col1:
    fig_capex = px.bar(capex_df, x="Year", y="Machines_to_Replace",
                       title="Machine Replacements per Year",
                       text=capex_df["Machines_to_Replace"].astype(int).map(lambda x: f"{x:,}"))
    fig_capex.update_layout(yaxis_tickformat=",")
    st.plotly_chart(fig_capex, use_container_width=True)

with col2:
    if "CAPEX" in capex_df.columns:
        capex_df["Cumulative_CAPEX"] = capex_df["CAPEX"].cumsum()
        fig_cumulative = px.line(capex_df, x="Year", y="Cumulative_CAPEX",
                                markers=True, title="Cumulative CAPEX Investment")
        fig_cumulative.update_layout(yaxis_tickformat="$,")
        st.plotly_chart(fig_cumulative, use_container_width=True)

# CAPEX summary table
disp_capex = capex_df.copy()
if "CAPEX" in disp_capex.columns:
    disp_capex["CAPEX"] = disp_capex["CAPEX"].apply(lambda x: fmt_money(x, 0))
    if "Cumulative_CAPEX" in disp_capex.columns:
        disp_capex["Cumulative CAPEX"] = disp_capex["Cumulative_CAPEX"].apply(lambda x: fmt_money(x, 0))
        cols = ["Machines_to_Replace", "CAPEX", "Cumulative CAPEX"]
    else:
        cols = ["Machines_to_Replace", "CAPEX"]
else:
    cols = ["Machines_to_Replace"]

disp_capex["Machines to Replace"] = disp_capex["Machines_to_Replace"].apply(fmt_int)
display_cols = ["Machines to Replace"] + ([col for col in cols if col != "Machines_to_Replace"])
st.table(disp_capex.set_index("Year")[display_cols])

# Investment insights
if capex_per_machine > 0:
    total_capex = capex_df["CAPEX"].sum() if "CAPEX" in capex_df.columns else 0
    peak_year = capex_df.loc[capex_df["Machines_to_Replace"].idxmax(), "Year"] if len(capex_df) > 0 else "N/A"
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total 10-Year CAPEX", f"${total_capex:,.0f}")
    with col2:
        st.metric("Peak Replacement Year", str(peak_year))
    with col3:
        annual_avg = total_capex / forecast_years if forecast_years > 0 else 0
        st.metric("Average Annual CAPEX", f"${annual_avg:,.0f}")

# -------- Machine Condition Analysis --------
st.subheader("🛠️ Machine Fleet Health Analysis")

# Create machine condition dashboard
col1, col2 = st.columns(2)

with col1:
    # Age distribution
    fig_age_dist = px.histogram(machines_df, x="Age_years", nbins=10, 
                               title="Machine Age Distribution",
                               labels={"Age_years": "Age (Years)", "count": "Number of Machines"})
    st.plotly_chart(fig_age_dist, use_container_width=True)

with col2:
    # Remaining life analysis
    remaining_life_valid = machines_df["Remaining_Life_years"].dropna()
    if not remaining_life_valid.empty:
        fig_remaining = px.histogram(remaining_life_valid, nbins=10,
                                   title="Remaining Useful Life Distribution",
                                   labels={"value": "Remaining Life (Years)", "count": "Number of Machines"})
        st.plotly_chart(fig_remaining, use_container_width=True)
    else:
        st.info("Remaining life data not available")

# Machine criticality analysis
machines_df["Criticality_Score"] = (
    (10 - machines_df["Remaining_Life_years"].fillna(5)) * 0.4 +  # Age factor
    ((1 - machines_df["Utilization"]) * 10) * 0.3 +  # Low utilization penalty
    (machines_df["Rated_Capacity_cpm"] / machines_df["Rated_Capacity_cpm"].max() * 10) * 0.3  # Capacity importance
).round(2)

# High-risk machines alert
critical_machines = machines_df[
    (machines_df["Remaining_Life_years"] < 1) | 
    (machines_df["Criticality_Score"] > 7)
]

if not critical_machines.empty:
    st.warning(f"🚨 **Critical Alert**: {len(critical_machines)} machines require immediate attention!")
    
    with st.expander("View Critical Machines Details"):
        critical_display = critical_machines[[machine_col, "Age_years", "Remaining_Life_years", 
                                            "Utilization", "Criticality_Score"]].copy()
        critical_display["Age (years)"] = critical_display["Age_years"].apply(lambda x: f"{x:.1f}" if pd.notna(x) else "N/A")
        critical_display["Remaining Life"] = critical_display["Remaining_Life_years"].apply(lambda x: f"{x:.1f}" if pd.notna(x) else "N/A")
        critical_display["Utilization"] = critical_display["Utilization"].apply(lambda x: f"{x:.1%}")
        critical_display["Risk Score"] = critical_display["Criticality_Score"].apply(lambda x: f"{x:.1f}")
        
        st.table(critical_display.set_index(machine_col)[["Age (years)", "Remaining Life", "Utilization", "Risk Score"]])

# Complete machine condition table
st.markdown("**Complete Fleet Status**")
cols = ["Rated_Capacity_cpm","Utilization","Start_Date","Age_years","Remaining_Life_years","End_of_Life_Year","Criticality_Score"]
show_cols = [c for c in [machine_col, *cols] if c in machines_df.columns]
disp_ml = machines_df[show_cols].copy()

if "Rated_Capacity_cpm" in disp_ml.columns:
    disp_ml["Capacity (cups/min)"] = disp_ml["Rated_Capacity_cpm"].apply(fmt_int)
if "Utilization" in disp_ml.columns:
    disp_ml["Utilization (%)"] = disp_ml["Utilization"].apply(lambda x: f"{x:.1%}")
for c in ["Age_years","Remaining_Life_years"]:
    if c in disp_ml.columns:
        col_name = "Age (years)" if c == "Age_years" else "Remaining Life (years)"
        disp_ml[col_name] = disp_ml[c].apply(lambda x: "" if pd.isna(x) else f"{x:.2f}")
if "Criticality_Score" in disp_ml.columns:
    disp_ml["Risk Score"] = disp_ml["Criticality_Score"].apply(lambda x: f"{x:.1f}")

# Select display columns
display_columns = []
if "Capacity (cups/min)" in disp_ml.columns:
    display_columns.append("Capacity (cups/min)")
if "Utilization (%)" in disp_ml.columns:
    display_columns.append("Utilization (%)")
if "Start_Date" in disp_ml.columns:
    display_columns.append("Start_Date")
if "Age (years)" in disp_ml.columns:
    display_columns.append("Age (years)")
if "Remaining Life (years)" in disp_ml.columns:
    display_columns.append("Remaining Life (years)")
if "End_of_Life_Year" in disp_ml.columns:
    display_columns.append("End_of_Life_Year")
if "Risk Score" in disp_ml.columns:
    display_columns.append("Risk Score")

if display_columns:
    st.table(disp_ml.set_index(machine_col)[display_columns])

# -------- Performance Benchmarking --------
st.subheader("📊 Performance Benchmarking")

# Industry benchmarks (example values - would typically come from external data)
industry_benchmarks = {
    "Utilization": 0.82,
    "Machine_Life": 12,
    "Maintenance_Cost_Ratio": 0.15,  # % of machine value
    "Output_per_Machine": 50000  # cups per day per machine
}

current_metrics = {
    "Utilization": machines_df["Utilization"].mean(),
    "Machine_Life": machine_life_years,
    "Output_per_Machine": (machines_df["Rated_Capacity_cpm"].mean() * 60 * 16 * machines_df["Utilization"].mean())  # 16h baseline
}

col1, col2 = st.columns(2)

with col1:
    st.markdown("**Current vs Industry Benchmarks**")
    benchmark_comparison = []
    for metric, current_val in current_metrics.items():
        if metric in industry_benchmarks:
            benchmark_val = industry_benchmarks[metric]
            variance = ((current_val - benchmark_val) / benchmark_val) * 100
            
            benchmark_comparison.append({
                "Metric": metric.replace("_", " ").title(),
                "Current": f"{current_val:.2f}" if metric != "Utilization" else f"{current_val:.1%}",
                "Industry": f"{benchmark_val:.2f}" if metric != "Utilization" else f"{benchmark_val:.1%}",
                "Variance": f"{variance:+.1f}%"
            })
    
    if benchmark_comparison:
        benchmark_df = pd.DataFrame(benchmark_comparison)
        st.table(benchmark_df.set_index("Metric"))

with col2:
    # Performance radar chart
    if benchmark_comparison:
        metrics = [item["Metric"] for item in benchmark_comparison]
        current_vals = [float(item["Current"].replace("%", "")) for item in benchmark_comparison]
        industry_vals = [float(item["Industry"].replace("%", "")) for item in benchmark_comparison]
        
        # Normalize values for radar chart
        max_vals = [max(c, i) for c, i in zip(current_vals, industry_vals)]
        current_norm = [c/m * 100 for c, m in zip(current_vals, max_vals)]
        industry_norm = [i/m * 100 for i, m in zip(industry_vals, max_vals)]
        
        fig_radar = go.Figure()
        
        fig_radar.add_trace(go.Scatterpolar(
            r=current_norm + [current_norm[0]],  # Close the radar
            theta=metrics + [metrics[0]],
            fill='toself',
            name='Current Performance',
            line_color='blue'
        ))
        
        fig_radar.add_trace(go.Scatterpolar(
            r=industry_norm + [industry_norm[0]],
            theta=metrics + [metrics[0]],
            fill='toself',
            name='Industry Benchmark',
            line_color='red'
        ))
        
        fig_radar.update_layout(
            polar=dict(
                radialaxis=dict(visible=True, range=[0, 100])
            ),
            showlegend=True,
            title="Performance vs Industry Benchmarks"
        )
        
        st.plotly_chart(fig_radar, use_container_width=True)

# -------- Executive Summary Dashboard --------
st.subheader("📋 Executive Summary")

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("**🎯 Production Capacity**")
    max_daily = scen_df["Daily_Output_cups"].max()
    max_monthly = scen_df["Monthly_Output_cups"].max()
    st.write(f"• Max Daily: {max_daily:,.0f} cups")
    st.write(f"• Max Monthly: {max_monthly:,.0f} cups")
    
    if sales_enabled:
        max_revenue = max_monthly * 12 * unit_price
        st.write(f"• Max Annual Revenue: ${max_revenue:,.0f}")

with col2:
    st.markdown("**⚠️ Risk Assessment**")
    total_machines = len(machines_df)
    critical_count = len(critical_machines) if not critical_machines.empty else 0
    risk_percentage = (critical_count / total_machines * 100) if total_machines > 0 else 0
    
    st.write(f"• Fleet Size: {total_machines} machines")
    st.write(f"• Critical Machines: {critical_count} ({risk_percentage:.1f}%)")
    
    if maintenance_due > 0:
        st.write(f"• Maintenance Due: {maintenance_due} machines")

with col3:
    st.markdown("**💰 Investment Outlook**")
    if capex_per_machine > 0 and "CAPEX" in capex_df.columns:
        total_capex = capex_df["CAPEX"].sum()
        next_year_capex = capex_df[capex_df["Year"] == pd.Timestamp.today().year + 1]["CAPEX"].iloc[0] if len(capex_df) > 0 else 0
        
        st.write(f"• 10-Year CAPEX: ${total_capex:,.0f}")
        st.write(f"• Next Year: ${next_year_capex:,.0f}")
    
    if annual_maintenance_per_machine > 0:
        annual_maint_cost = len(machines_df) * annual_maintenance_per_machine
        st.write(f"• Annual Maintenance: ${annual_maint_cost:,.0f}")

# Key recommendations
st.markdown("**🎯 Key Recommendations**")
recommendations = []

if critical_count > 0:
    recommendations.append(f"• Prioritize replacement/overhaul of {critical_count} critical machines")

if avg_util < 0.80:
    recommendations.append("• Investigate utilization gaps - current performance below 80%")

if aging_machines > total_machines * 0.3:
    recommendations.append("• Develop accelerated replacement strategy - high proportion of aging fleet")

if not recommendations:
    recommendations.append("• Fleet is in good condition - maintain current preventive maintenance schedule")

for rec in recommendations:
    st.write(rec)

# -------- Enhanced Downloads --------
st.subheader("⬇️ Comprehensive Reports & Data Export")

# Prepare enhanced report sheets
report_sheets = {
    "Executive_Summary": pd.DataFrame([{
        "Total_Machines": len(machines_df),
        "Avg_Utilization": f"{machines_df['Utilization'].mean():.2%}",
        "Total_Capacity_CPM": machines_df['Rated_Capacity_cpm'].sum(),
        "Critical_Machines": critical_count,
        "Max_Daily_Output": scen_df["Daily_Output_cups"].max(),
        "Max_Monthly_Output": scen_df["Monthly_Output_cups"].max(),
        "Total_10Y_CAPEX": capex_df["CAPEX"].sum() if "CAPEX" in capex_df.columns else 0,
        "Report_Generated": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")
    }]),
    "Scenario_Analysis": scen_df,
    **{f"Annual_Forecast_{h}h": tables[h] for h in hour_scenarios},
    "CAPEX_Schedule": capex_df,
    "Machine_Condition": machines_df[show_cols] if show_cols else machines_df,
    "Maintenance_Schedule": maintenance_schedule if not maintenance_schedule.empty else pd.DataFrame(),
    "Benchmark_Analysis": pd.DataFrame(benchmark_comparison) if 'benchmark_comparison' in locals() else pd.DataFrame()
}

# Add risk analysis if available
if 'risk_df' in locals():
    report_sheets["Risk_Analysis"] = risk_df

excel_bytes = to_excel_bytes(report_sheets)

col1, col2, col3 = st.columns(3)

with col1:
    st.download_button(
        "📊 Download Complete Report (.xlsx)",
        data=excel_bytes,
        file_name=f"Factory_Analysis_Report_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Comprehensive Excel report with all analysis tabs"
    )

with col2:
    st.download_button(
        "📈 Download Scenarios CSV",
        data=scen_df.to_csv(index=False).encode("utf-8"),
        file_name=f"production_scenarios_{pd.Timestamp.now().strftime('%Y%m%d')}.csv",
        mime="text/csv"
    )

with col3:
    st.download_button(
        "💰 Download CAPEX Plan CSV",
        data=capex_df.to_csv(index=False).encode("utf-8"),
        file_name=f"capex_schedule_{pd.Timestamp.now().strftime('%Y%m%d')}.csv",
        mime="text/csv"
    )

# Footer with analysis notes
st.markdown("---")
st.caption("""
**Analysis Notes**: 
• Machine ages calculated from Start Date/Year/Age columns with robust parsing
• Utilization rates averaged across fleet for production calculations  
• CAPEX replacements scheduled for end-of-life year
• Risk scenarios model machine reliability, demand volatility, and supply chain disruptions
• Maintenance scheduling based on 2-year major/minor cycles
• Industry benchmarks are illustrative - customize based on your sector data
• All monetary values exclude taxes and financing costs
""")

st.caption(f"Report generated: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
