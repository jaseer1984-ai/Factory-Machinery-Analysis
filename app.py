# Paper Cup Factory — Investor Dashboard (10-year horizon, no 20-year forecast)
# Includes:
# - Current Sales Report (LTM, last-month KPIs, ASP trend, top customers)
# - Production scenarios (12/16/20/24h)
# - 10-year forecast (units, optional revenue/GM/Net Margin)
# - Demand vs Capacity coverage (with scrap & CAGR)
# - Project economics (Cash Flows, NPV, Payback)
# - Sensitivity (heatmap + tornado)
# - CAPEX replacement plan (10y life)
# - Maintenance schedule & alerts
# - Fleet health & benchmarking
# - Excel/CSV exports

from datetime import datetime
import io
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

st.set_page_config(page_title="Paper Cup Factory — Investor Dashboard", layout="wide")
st.title("📈 Paper Cup Factory — Investor Dashboard (10-Year View)")

# ---------------- Sidebar ----------------
with st.sidebar:
    st.header("⚙️ Inputs")
    uploaded = st.file_uploader("Upload factory Excel", type=["xlsx"])

    days_per_month = st.number_input("Days per month", 1, 31, value=28)
    hour_scenarios = st.multiselect(
        "Hours per day (select)", options=[12, 16, 20, 24], default=[12, 16, 20, 24]
    )
    machine_life_years = st.number_input("Machine life (years)", 1, 40, value=10)

    # Forecast horizon fixed to 10 years (no 20-year option)
    st.markdown("**Forecast horizon:** 10 years (fixed)")
    forecast_years = 10

    st.divider()
    st.subheader("📊 Advanced Settings")

    # Seasonality factors
    st.markdown("**Seasonality Adjustments**")
    seasonality_enabled = st.checkbox("Enable seasonal demand patterns", value=False)
    peak_months = []
    seasonality_factor = 1.0
    if seasonality_enabled:
        peak_months = st.multiselect(
            "Peak demand months",
            options=list(range(1, 13)),
            default=[11, 12],
            format_func=lambda x: datetime(2024, x, 1).strftime("%B"),
        )
        seasonality_factor = st.slider("Peak season multiplier", 1.0, 2.0, 1.3, 0.1)

    # Risk scenarios
    st.markdown("**Risk Analysis**")
    risk_analysis = st.checkbox("Enable risk scenarios", value=False)
    machine_failure_rate = 0.0
    if risk_analysis:
        machine_failure_rate = st.slider("Annual machine failure rate (%)", 0.0, 20.0, 5.0, 0.5)

    st.divider()
    st.subheader("💰 Economics")
    unit_price = st.number_input("Unit price (per cup)", min_value=0.0, value=0.0, step=0.001, format="%.4f")
    unit_cost  = st.number_input("Unit cost (per cup)",  min_value=0.0, value=0.0, step=0.001, format="%.4f")
    capex_per_machine = st.number_input("CAPEX per machine", min_value=0.0, value=0.0, step=1000.0, format="%.0f")
    annual_maintenance_per_machine = st.number_input(
        "Annual maintenance per machine", min_value=0.0, value=5000.0, step=500.0
    )

    st.divider()
    st.subheader("📈 Demand & Finance")
    demand_base = st.number_input("Base annual demand (cups)", min_value=0.0, value=0.0, step=100000.0, format="%.0f")
    demand_cagr = st.slider("Demand growth (CAGR %)", 0.0, 50.0, 5.0, 0.5)
    reject_rate = st.slider("Scrap/Reject rate %", 0.0, 20.0, 2.0, 0.5)
    discount_rate = st.slider("Discount rate (WACC) %", 0.0, 30.0, 12.0, 0.5)
    finance_hours = st.selectbox("Evaluate finance & coverage at (hours/day)", options=hour_scenarios, index=0)

# ---------------- Helpers ----------------
def find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    norm = {str(c).strip().lower(): c for c in df.columns}
    for want in candidates:
        key = str(want).strip().lower()
        if key in norm: return norm[key]
    for want in candidates:
        key = str(want).strip().lower()
        for k, real in norm.items():
            if key in k: return real
    return None

def parse_date_series(s: pd.Series):
    def parse_one(v):
        if pd.isna(v): return pd.NaT
        if isinstance(v, (pd.Timestamp, datetime)): return pd.to_datetime(v)
        v = str(v)
        fmts = ("%Y-%m-%d","%d-%m-%Y","%d/%m/%Y","%m/%d/%Y","%d-%b-%Y","%b-%d-%Y","%Y/%m/%d")
        for fmt in fmts:
            try: return datetime.strptime(v, fmt)
            except: pass
        return pd.to_datetime(v, errors="coerce", dayfirst=True)
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
             ["sales","prod","output","shift","daily","util","run","report"]) and s != machines_sheet]
    production_sheet = cand_p[0] if cand_p else (names[1] if len(names)>1 else names[0])
    return machines_sheet, production_sheet

def extract_core(machines_df: pd.DataFrame, prod_df: pd.DataFrame, life_years: float):
    machine_col = find_col(machines_df, ["machine","machine id","machine_name","name","id"]) or "Machine"
    if machine_col not in machines_df.columns:
        machines_df["Machine"] = [f"M{i+1}" for i in range(len(machines_df))]
        machine_col = "Machine"

    cap_col = find_col(machines_df, [
        "capacity","rated capacity","rated_capacity","cups/min","cups per min",
        "capacity (cups/min)","capacity_cups_min","capacity_cup_min"
    ])
    if cap_col is None:
        machines_df["Rated_Capacity_cpm"] = 75.0
        cap_col = "Rated_Capacity_cpm"
    else:
        cap_vals = pd.to_numeric(machines_df[cap_col], errors="coerce")
        if cap_vals.isna().all(): cap_vals = pd.Series([75.0]*len(machines_df))
        machines_df["Rated_Capacity_cpm"] = cap_vals.fillna(cap_vals.median())
        cap_col = "Rated_Capacity_cpm"

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

    # Simple maintenance flags
    machines_df["Next_Major_Maintenance"] = machines_df["Start_Date"] + pd.DateOffset(years=2)
    machines_df["Maintenance_Due"] = (
        machines_df["Next_Major_Maintenance"] < pd.Timestamp.today() + pd.DateOffset(months=6)
    )
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

def enhanced_annual_forecast(machines_df, hours, dpm, years,
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
            cups_day = base_cpm * 60 * h * util
            cups_year = cups_day * dpm * 12
            if seasonality_enabled and peak_months:
                cups_year *= (1.0 + (seasonality_factor - 1.0) * (len(peak_months) / 12.0))
            if risk_analysis and failure_rate > 0:
                cups_year *= (1.0 - failure_rate/100.0)
            row = {"Year": y, "Hours_per_Day": h, "Output_cups": cups_year}
            if price > 0 and cost >= 0:
                row["Revenue"] = cups_year * price
                row["Gross_Margin"] = cups_year * (price - cost)
                if maintenance_cost > 0:
                    row["Net_Margin"] = row["Gross_Margin"] - (len(machines_df) * maintenance_cost)
            rows.append(row)
        tables[h] = pd.DataFrame(rows, columns=list(rows[0].keys()))
    return tables

def generate_maintenance_schedule(machines_df: pd.DataFrame, years: int):
    current_year = pd.Timestamp.today().year
    schedule = []
    for idx, r in machines_df.iterrows():
        m_id = r.get("Machine", f"M{idx+1}")
        if pd.isna(r.get("Start_Date")): continue
        for offset in range(years):
            y = current_year + offset
            schedule.append({
                "Machine": m_id,
                "Year": y,
                "Maintenance_Type": "Major" if offset % 2 == 0 else "Minor",
                "Estimated_Cost": 15000 if offset % 2 == 0 else 5000,
                "Downtime_Days": 5 if offset % 2 == 0 else 2
            })
    return pd.DataFrame(schedule)

def create_risk_scenario_analysis(base_output: float, scenarios: dict):
    rows = []
    for name, factors in scenarios.items():
        adj = base_output
        for f in factors.values(): adj *= f
        rows.append({
            "Scenario": name,
            "Output_Multiplier": adj / base_output if base_output else np.nan,
            "Annual_Output": adj,
            "Variance_from_Base": ((adj - base_output) / base_output * 100) if base_output else np.nan
        })
    return pd.DataFrame(rows)

def capex_schedule(machines_df: pd.DataFrame, years: int, life_years: float, capex_each: float = 0.0):
    current_year = pd.Timestamp.today().year
    ylist = [current_year + i for i in range(years)]
    counts = {y: 0 for y in ylist}
    for _, r in machines_df.iterrows():
        eol = r["End_of_Life_Year"]
        if pd.notna(eol):
            eol = int(eol)
            if eol in counts: counts[eol] += 1
    df = pd.DataFrame({"Year": ylist, "Machines_to_Replace": [counts[y] for y in ylist]})
    if capex_each > 0: df["CAPEX"] = df["Machines_to_Replace"] * capex_each
    return df

def to_excel_bytes(sheets: dict) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=name[:31] or "Sheet")
    return output.getvalue()

# display format helpers
def fmt_int(x):
    if x is None or (isinstance(x, float) and np.isnan(x)): return ""
    return f"{int(round(x)):,}"

def fmt_money(x, decimals=2):
    if x is None or (isinstance(x, float) and np.isnan(x)): return ""
    return f"{x:,.{decimals}f}"

# --------- Sales (Current Sales Report) ---------
def extract_sales_kpis(sales_df: pd.DataFrame):
    if sales_df is None or sales_df.empty: return None
    date_col  = find_col(sales_df, ["date","invoice date","sale date","order date","month"])
    qty_col   = find_col(sales_df, ["qty","quantity","units","cups","pcs","pieces","volume"])
    amt_col   = find_col(sales_df, ["amount","revenue","sales","value","total","net","invoice amount"])
    price_col = find_col(sales_df, ["price","rate","unit price","price/unit","avg price"])
    cust_col  = find_col(sales_df, ["customer","client","buyer","account","party"])
    if date_col is None: return None

    df = sales_df.copy()
    df[date_col] = parse_date_series(df[date_col])
    df = df[df[date_col].notna()]
    if qty_col: df["Qty"] = pd.to_numeric(df[qty_col], errors="coerce")
    elif amt_col and price_col:
        df["Qty"] = pd.to_numeric(df[amt_col], errors="coerce") / pd.to_numeric(df[price_col], errors="coerce")
    else: df["Qty"] = np.nan
    if amt_col: df["Amount"] = pd.to_numeric(df[amt_col], errors="coerce")
    elif price_col: df["Amount"] = pd.to_numeric(df[price_col], errors="coerce") * df["Qty"]
    else: df["Amount"] = np.nan

    df["Month"] = df[date_col].dt.to_period("M").dt.to_timestamp()
    monthly = df.groupby("Month", as_index=False).agg(Units=("Qty","sum"), Revenue=("Amount","sum"))
    monthly["ASP"] = np.where(monthly["Units"]>0, monthly["Revenue"]/monthly["Units"], np.nan)

    last_month = monthly["Month"].max() if not monthly.empty else None
    kpis = {}
    if last_month is not None:
        lm = monthly.loc[monthly["Month"]==last_month].iloc[0]
        prev = monthly.loc[monthly["Month"]==(last_month - pd.DateOffset(months=1))]
        kpis["last_month_units"] = float(lm["Units"])
        kpis["last_month_revenue"] = float(lm["Revenue"])
        kpis["last_month_asp"] = float(lm["ASP"]) if not pd.isna(lm["ASP"]) else np.nan
        prev_rev = prev["Revenue"].iloc[0] if not prev.empty else np.nan
        kpis["mom_growth"] = (lm["Revenue"] - prev_rev)/prev_rev if prev_rev not in [0, np.nan] else np.nan
        start_ltm = last_month - pd.DateOffset(months=11)
        ltm = monthly[(monthly["Month"]>=start_ltm) & (monthly["Month"]<=last_month)]
        kpis["ltm_units"] = float(ltm["Units"].sum())
        kpis["ltm_revenue"] = float(ltm["Revenue"].sum())

    top_customers = None
    if cust_col and last_month is not None:
        start = last_month - pd.DateOffset(months=11)
        mask = (df["Month"]>=start) & (df["Month"]<=last_month)
        top = (df.loc[mask].groupby(cust_col, as_index=False)["Amount"]
                 .sum().sort_values("Amount", ascending=False).head(5))
        if not top.empty:
            top["Share_%"] = (top["Amount"]/float(df.loc[mask,"Amount"].sum()))*100.0
            top_customers = top
    return {"monthly": monthly, "kpis": kpis, "top_customers": top_customers}

# ---------------- Main ----------------
if uploaded is None:
    st.info("Upload your Excel to begin. Forecast is fixed at **10 years**.")
    st.stop()

# Load/process
wb = load_workbook(uploaded)
machines_sheet, production_sheet = detect_key_sheets(wb)
st.markdown(f"**Detected Sheets:** `{machines_sheet}` (machines) | `{production_sheet}` (production/sales)")

machines_df_raw = wb[machines_sheet].copy()
prod_df_raw = wb[production_sheet].copy()
machine_col, cap_col, machines_df = extract_core(machines_df_raw.copy(), prod_df_raw.copy(), machine_life_years)

# Fleet metrics
avg_util = float(machines_df['Utilization'].mean())
total_capacity = float(machines_df['Rated_Capacity_cpm'].sum())
avg_age = machines_df["Age_years"].dropna().mean()
aging_machines = int((machines_df["Remaining_Life_years"] < 2).sum())
maintenance_due = int(machines_df["Maintenance_Due"].sum()) if "Maintenance_Due" in machines_df.columns else 0

cap_max = float(machines_df["Rated_Capacity_cpm"].max()) if "Rated_Capacity_cpm" in machines_df.columns else 1.0
cap_max = cap_max if cap_max > 0 else 1.0
machines_df["Criticality_Score"] = (
    (10 - machines_df["Remaining_Life_years"].fillna(5)) * 0.4 +
    ((1 - machines_df["Utilization"].fillna(avg_util)) * 10) * 0.3 +
    (machines_df["Rated_Capacity_cpm"].fillna(0) / cap_max * 10) * 0.3
).round(2)
critical_machines = machines_df[
    (machines_df["Remaining_Life_years"] < 1) |
    (machines_df["Criticality_Score"] > 7)
]

# ---------- KPI ----------
st.subheader("📊 Factory Performance KPIs")
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Total Machines", f"{len(machines_df):,}")
c2.metric("Avg Utilization", f"{avg_util:.2%}")
c3.metric("Total Capacity", f"{total_capacity:,.0f} cups/min")
c4.metric("Aging Machines (<2y life)", f"{aging_machines:,}", delta_color="inverse")
c5.metric("Avg Age", "N/A" if pd.isna(avg_age) else f"{avg_age:.1f} years")
if aging_machines > 0: st.warning(f"⚠️ {aging_machines} machines need attention within 2 years.")
if maintenance_due > 0: st.info(f"🔧 {maintenance_due} machines due for major maintenance soon.")

# ---------- Current Sales Report ----------
st.subheader("🧾 Current Sales Report")
sales_info = extract_sales_kpis(prod_df_raw)
if not sales_info:
    st.info("No parsable sales data found (date/qty/amount).")
else:
    m = sales_info["monthly"]; k = sales_info["kpis"]
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Last Month Units", f"{k.get('last_month_units', 0):,.0f}")
    k2.metric("Last Month Revenue", f"{k.get('last_month_revenue', 0):,.0f}")
    k3.metric("Last Month ASP", f"{k.get('last_month_asp', 0):,.4f}")
    k4.metric("MoM Revenue Growth", f"{k.get('mom_growth', np.nan):+.1%}" if not pd.isna(k.get('mom_growth', np.nan)) else "N/A")
    d1,d2 = st.columns(2)
    d1.metric("LTM Units", f"{k.get('ltm_units', 0):,.0f}")
    d2.metric("LTM Revenue", f"{k.get('ltm_revenue', 0):,.0f}")

    colA, colB = st.columns(2)
    colA.plotly_chart(go.Figure(data=[go.Bar(x=m["Month"], y=m["Revenue"], text=m["Revenue"].apply(lambda x:f"{x:,.0f}"))]).update_layout(title="Monthly Revenue", yaxis_tickformat=","), use_container_width=True)
    colB.plotly_chart(px.line(m, x="Month", y="Units", markers=True, title="Monthly Units (cups)").update_layout(yaxis_tickformat=","), use_container_width=True)
    colC, colD = st.columns(2)
    colC.plotly_chart(px.line(m, x="Month", y="ASP", markers=True, title="Average Selling Price (ASP)"), use_container_width=True)
    tc = sales_info["top_customers"]
    if tc is not None and not tc.empty:
        tc_disp = tc.copy(); tc_disp["Amount"] = tc_disp["Amount"].apply(lambda x:f"{x:,.0f}"); tc_disp["Share_%"] = tc_disp["Share_%"].apply(lambda x:f"{x:.1f}%")
        colD.markdown("**Top Customers (LTM)**"); colD.table(tc_disp.set_index(tc.columns[0]))
    disp = m.copy(); disp["Units"] = disp["Units"].apply(lambda x:f"{x:,.0f}"); disp["Revenue"] = disp["Revenue"].apply(lambda x:f"{x:,.0f}")
    disp["ASP"] = disp["ASP"].apply(lambda x:"" if pd.isna(x) else f"{x:,.4f}")
    st.table(disp.set_index("Month"))

# ---------- Scenarios ----------
if not hour_scenarios:
    st.warning("No hours selected; defaulting to 12/16/20/24.")
    hour_scenarios = [12,16,20,24]
scen_df = scenario_outputs(machines_df, hour_scenarios, days_per_month).sort_values("Hours")
st.subheader("🎯 Production Scenarios (Daily & Monthly)")
fig_scen = make_subplots(rows=1, cols=2, subplot_titles=("Daily Output by Hours", "Monthly Output Trend"))
fig_scen.add_trace(go.Bar(x=scen_df["Hours"], y=scen_df["Daily_Output_cups"],
                          text=scen_df["Daily_Output_cups"].round().astype(int).map(lambda v:f"{v:,}"),
                          textposition="outside"), row=1, col=1)
fig_scen.add_trace(go.Scatter(x=scen_df["Hours"], y=scen_df["Monthly_Output_cups"], mode="lines+markers"), row=1, col=2)
fig_scen.update_layout(height=420, showlegend=False)
st.plotly_chart(fig_scen, use_container_width=True)
disp_scen = scen_df.copy()
disp_scen["Daily Output (cups)"] = disp_scen["Daily_Output_cups"].apply(fmt_int)
disp_scen["Monthly Output (cups)"] = disp_scen["Monthly_Output_cups"].apply(fmt_int)
st.table(disp_scen.set_index("Hours")[["Daily Output (cups)","Monthly Output (cups)"]])

# ---------- 10-year Forecast (locked) ----------
sales_enabled = unit_price > 0 and unit_cost >= 0
forecast_params = dict(
    seasonality_enabled=seasonality_enabled,
    peak_months=peak_months if seasonality_enabled else None,
    seasonality_factor=seasonality_factor if seasonality_enabled else 1.0,
    risk_analysis=risk_analysis,
    failure_rate=machine_failure_rate if risk_analysis else 0.0,
)
tables = enhanced_annual_forecast(
    machines_df, hour_scenarios, days_per_month, forecast_years,
    price=unit_price if sales_enabled else 0.0,
    cost=unit_cost if sales_enabled else 0.0,
    maintenance_cost=annual_maintenance_per_machine,
    **forecast_params
)
st.subheader("📈 10-Year Forecast" + (" (Seasonality/Risk ON)" if (seasonality_enabled or risk_analysis) else ""))
tabs = st.tabs([f"{h}h/day" for h in hour_scenarios])
for tab, h in zip(tabs, hour_scenarios):
    with tab:
        fig_enh = make_subplots(rows=2, cols=1, subplot_titles=(f"Production — {h}h/day", "Financials" if sales_enabled else ""))
        fig_enh.add_trace(go.Scatter(x=tables[h]["Year"], y=tables[h]["Output_cups"], mode="lines+markers", name="Output"), row=1, col=1)
        if sales_enabled and "Revenue" in tables[h].columns:
            fig_enh.add_trace(go.Scatter(x=tables[h]["Year"], y=tables[h]["Revenue"], mode="lines+markers", name="Revenue"), row=2, col=1)
            if "Net_Margin" in tables[h].columns:
                fig_enh.add_trace(go.Scatter(x=tables[h]["Year"], y=tables[h]["Net_Margin"], mode="lines+markers", name="Net Margin"), row=2, col=1)
        fig_enh.update_layout(height=600); st.plotly_chart(fig_enh, use_container_width=True)

        disp = tables[h].copy()
        disp["Output (cups)"] = disp["Output_cups"].apply(fmt_int)
        cols = ["Year","Hours_per_Day","Output (cups)"]
        if "Revenue" in disp.columns:
            disp["Revenue"] = disp["Revenue"].apply(lambda x: fmt_money(x, 0))
            disp["Gross_Margin"] = disp["Gross_Margin"].apply(lambda x: fmt_money(x, 0))
            cols += ["Revenue","Gross_Margin"]
            if "Net_Margin" in disp.columns:
                disp["Net_Margin"] = disp["Net_Margin"].apply(lambda x: fmt_money(x, 0))
                cols += ["Net_Margin"]
        st.table(disp[cols].set_index("Year"))

# ---------- Demand vs Capacity ----------
st.subheader("📦 Demand vs Capacity (Coverage)")
cap_series = tables[finance_hours][["Year","Output_cups"]].copy()
cap_series["Net_Capacity"] = cap_series["Output_cups"] * (1.0 - reject_rate/100.0)
years = cap_series["Year"].tolist()
dem_vals = [demand_base * ((1 + demand_cagr/100.0) ** i) for i in range(len(years))]
demand_df = pd.DataFrame({"Year": years, "Demand": dem_vals})
cov = cap_series.merge(demand_df, on="Year", how="left")
cov["Excess_(+)/Shortfall_(-)"] = cov["Net_Capacity"] - cov["Demand"]
c1, c2 = st.columns(2)
c1.plotly_chart(go.Figure(data=[
    go.Scatter(x=cov["Year"], y=cov["Demand"], mode="lines+markers", name="Demand"),
    go.Scatter(x=cov["Year"], y=cov["Net_Capacity"], mode="lines+markers", name="Net Capacity")
]).update_layout(title=f"Annual Demand vs Net Capacity @ {finance_hours}h/day", yaxis_tickformat=","), use_container_width=True)
disp_cov = cov.copy()
disp_cov["Net Capacity (cups)"] = disp_cov["Net_Capacity"].apply(lambda x:f"{x:,.0f}")
disp_cov["Demand (cups)"] = disp_cov["Demand"].apply(lambda x:f"{x:,.0f}")
disp_cov["Excess(+)/Shortfall(-)"] = disp_cov["Excess_(+)/Shortfall_(-)"].apply(lambda x:f"{x:,.0f}")
c2.table(disp_cov.set_index("Year")[["Demand (cups)","Net Capacity (cups)","Excess(+)/Shortfall(-)"]])

# ---------- Project economics: Cash Flow, NPV, Payback ----------
st.subheader("💵 Project Economics: Cash Flows, NPV & Payback")
if unit_price <= 0:
    st.info("Enter **Unit price** and **Unit cost** in the sidebar to compute revenue, margin, NPV and payback.")
else:
    cf = cov.copy()
    cf["Revenue"] = cf["Net_Capacity"] * unit_price
    cf["Variable_Cost"] = cf["Net_Capacity"] * unit_cost
    cf["Gross_Margin"] = cf["Revenue"] - cf["Variable_Cost"]
    cf["Maintenance"] = len(machines_df) * annual_maintenance_per_machine
    capex_df = capex_schedule(machines_df, forecast_years, machine_life_years,
                              capex_each=capex_per_machine if capex_per_machine>0 else 0.0)
    cap_map = dict(zip(capex_df["Year"], capex_df.get("CAPEX", pd.Series([0]*len(capex_df)))))
    cf["CAPEX"] = cf["Year"].map(cap_map).fillna(0.0)
    cf["Net_Cash_Flow"] = cf["Gross_Margin"] - cf["Maintenance"] - cf["CAPEX"]

    r = discount_rate/100.0
    npv = sum(cf.loc[i, "Net_Cash_Flow"] / ((1+r)**(i+1)) for i in range(len(cf)))
    cf["Cumulative_CF"] = cf["Net_Cash_Flow"].cumsum()
    payback_year = cf.loc[cf["Cumulative_CF"] >= 0, "Year"].min() if (cf["Cumulative_CF"] >= 0).any() else "Not reached"

    k1,k2,k3 = st.columns(3)
    k1.metric("NPV", f"${npv:,.0f}", help=f"Discount rate: {discount_rate:.1f}%")
    k2.metric("Payback Year", str(payback_year))
    k3.metric("Total Net Cash (10y)", f"${cf['Net_Cash_Flow'].sum():,.0f}")

    fig_cf = go.Figure()
    fig_cf.add_trace(go.Bar(x=cf["Year"], y=cf["Net_Cash_Flow"], name="Net Cash Flow",
                            text=cf["Net_Cash_Flow"].apply(lambda x:f"{x:,.0f}")))
    fig_cf.add_trace(go.Scatter(x=cf["Year"], y=cf["Cumulative_CF"], mode="lines+markers", name="Cumulative CF"))
    fig_cf.update_layout(title="Annual & Cumulative Cash Flows", yaxis_tickformat="$,")
    st.plotly_chart(fig_cf, use_container_width=True)

    show = cf[["Year","Revenue","Variable_Cost","Gross_Margin","Maintenance","CAPEX","Net_Cash_Flow","Cumulative_CF"]].copy()
    for c in show.columns:
        if c!="Year": show[c] = show[c].apply(lambda x:f"{x:,.0f}")
    st.table(show.set_index("Year"))

# ---------- Sensitivity ----------
st.subheader("🧪 Sensitivity Analysis")
hours_grid = np.arange(12, 25)
util_grid = np.linspace(0.70, 1.00, 16)
base_cpm = machines_df["Rated_Capacity_cpm"].sum()
Z = [[base_cpm * 60 * h * u * days_per_month * 12 for h in hours_grid] for u in util_grid]
heat = px.imshow(np.array(Z), x=hours_grid, y=(util_grid*100).round(0),
                 labels=dict(x="Hours/Day", y="Utilization %", color="Annual Output (cups)"),
                 title="Annual Output Sensitivity — Hours vs Utilization")
heat.update_layout(coloraxis_colorbar_tickformat=",")
st.plotly_chart(heat, use_container_width=True)

if unit_price > 0:
    def gm(hours, util, price, cost, dpm):
        annual = base_cpm * 60 * hours * util * dpm * 12
        return annual * (price - cost)
    u0 = float(machines_df["Utilization"].mean())
    gm0 = gm(finance_hours, u0, unit_price, unit_cost, days_per_month)
    tests = [
        ("Hours/Day", lambda s: gm(finance_hours*(1+s), u0, unit_price, unit_cost, days_per_month)),
        ("Utilization", lambda s: gm(finance_hours, u0*(1+s), unit_price, unit_cost, days_per_month)),
        ("Unit Price", lambda s: gm(finance_hours, u0, unit_price*(1+s), unit_cost, days_per_month)),
        ("Unit Cost",  lambda s: gm(finance_hours, u0, unit_price, unit_cost*(1+s), days_per_month)),
        ("Days/Month", lambda s: gm(finance_hours, u0, unit_price, unit_cost, days_per_month*(1+s))),
    ]
    rows = []
    for name, f in tests:
        low = f(-0.10) - gm0
        high = f(+0.10) - gm0
        rows.append({"Driver": name, "Low": low, "High": high})
    tor = pd.DataFrame(rows).set_index("Driver").sort_values("High", ascending=True)
    fig_tor = go.Figure()
    fig_tor.add_trace(go.Bar(x=tor["Low"],  y=tor.index, orientation="h", name="-10%"))
    fig_tor.add_trace(go.Bar(x=tor["High"], y=tor.index, orientation="h", name="+10%"))
    fig_tor.update_layout(barmode="relative", title=f"Tornado — GM sensitivity @ {finance_hours}h/day", xaxis_tickformat="$,")
    st.plotly_chart(fig_tor, use_container_width=True)
else:
    st.info("Set price & cost to view GM sensitivity.")

# ---------- CAPEX ----------
st.subheader("💰 CAPEX Replacement Schedule & Investment Planning")
capex_df = capex_schedule(machines_df, forecast_years, machine_life_years,
                          capex_each=capex_per_machine if capex_per_machine>0 else 0.0)
c1,c2 = st.columns(2)
c1.plotly_chart(px.bar(capex_df, x="Year", y="Machines_to_Replace",
                       title="Machine Replacements per Year",
                       text=capex_df["Machines_to_Replace"].astype(int).map(lambda x:f"{x:,}")).update_layout(yaxis_tickformat=","),
                use_container_width=True)
if "CAPEX" in capex_df.columns:
    capex_df["Cumulative_CAPEX"] = capex_df["CAPEX"].cumsum()
    c2.plotly_chart(px.line(capex_df, x="Year", y="Cumulative_CAPEX", markers=True,
                            title="Cumulative CAPEX Investment").update_layout(yaxis_tickformat="$,"),
                    use_container_width=True)
disp_capex = capex_df.copy()
if "CAPEX" in disp_capex.columns:
    disp_capex["CAPEX"] = disp_capex["CAPEX"].apply(lambda x: fmt_money(x, 0))
    if "Cumulative_CAPEX" in disp_capex.columns:
        disp_capex["Cumulative CAPEX"] = disp_capex["Cumulative_CAPEX"].apply(lambda x: fmt_money(x, 0))
st.table(disp_capex.set_index("Year")[([ "Machines_to_Replace"] + [c for c in ["CAPEX","Cumulative CAPEX"] if c in disp_capex.columns])].rename(columns={"Machines_to_Replace":"Machines to Replace"}))

# ---------- Maintenance scheduling ----------
st.subheader("🔧 Preventive Maintenance Schedule")
maintenance_schedule = generate_maintenance_schedule(machines_df, forecast_years)
if not maintenance_schedule.empty:
    yearly = maintenance_schedule.groupby("Year", as_index=False).agg(
        Total_Cost=("Estimated_Cost","sum"),
        Total_Downtime_Days=("Downtime_Days","sum"),
        Maintenance_Events=("Machine","count")
    )
    d1,d2 = st.columns(2)
    d1.plotly_chart(px.bar(yearly, x="Year", y="Total_Cost", title="Annual Maintenance Costs",
                           text=yearly["Total_Cost"].apply(lambda x:f"${x:,.0f}")), use_container_width=True)
    d2.plotly_chart(px.line(yearly, x="Year", y="Maintenance_Events", markers=True,
                            title="Maintenance Events per Year"), use_container_width=True)
    disp_maint = yearly.copy()
    disp_maint["Total Cost"] = disp_maint["Total_Cost"].apply(lambda x: fmt_money(x, 0))
    disp_maint["Events"] = disp_maint["Maintenance_Events"].apply(fmt_int)
    disp_maint["Downtime (days)"] = disp_maint["Total_Downtime_Days"].apply(fmt_int)
    st.table(disp_maint.set_index("Year")[["Events","Total Cost","Downtime (days)"]])

# ---------- Fleet health ----------
st.subheader("🛠️ Machine Fleet Health")
h1,h2 = st.columns(2)
h1.plotly_chart(px.histogram(machines_df, x="Age_years", nbins=10, title="Age Distribution"), use_container_width=True)
if machines_df["Remaining_Life_years"].notna().any():
    h2.plotly_chart(px.histogram(machines_df.dropna(subset=["Remaining_Life_years"]),
                                 x="Remaining_Life_years", nbins=10, title="Remaining Useful Life"), use_container_width=True)
else:
    h2.info("Remaining life data not available")

if not critical_machines.empty:
    st.warning(f"🚨 Critical: {len(critical_machines)} machines need immediate attention!")
    with st.expander("View Critical Machines"):
        cd = critical_machines[[machine_col,"Age_years","Remaining_Life_years","Utilization","Criticality_Score"]].copy()
        cd["Age (years)"] = cd["Age_years"].apply(lambda x: f"{x:.1f}" if pd.notna(x) else "N/A")
        cd["Remaining Life"] = cd["Remaining_Life_years"].apply(lambda x: f"{x:.1f}" if pd.notna(x) else "N/A")
        cd["Utilization"] = cd["Utilization"].apply(lambda x: f"{x:.1%}")
        cd["Risk Score"] = cd["Criticality_Score"].apply(lambda x: f"{x:.1f}")
        st.table(cd.set_index(machine_col)[["Age (years)","Remaining Life","Utilization","Risk Score"]])

st.markdown("**Complete Fleet Status**")
cols = ["Rated_Capacity_cpm","Utilization","Start_Date","Age_years","Remaining_Life_years","End_of_Life_Year","Criticality_Score"]
show_cols = [c for c in [machine_col, *cols] if c in machines_df.columns]
disp_ml = machines_df[show_cols].copy()
if "Rated_Capacity_cpm" in disp_ml.columns: disp_ml["Capacity (cups/min)"] = disp_ml["Rated_Capacity_cpm"].apply(lambda x:f"{int(round(x)):,}")
if "Utilization" in disp_ml.columns: disp_ml["Utilization (%)"] = disp_ml["Utilization"].apply(lambda x:f"{x:.1%}")
for c in ["Age_years","Remaining_Life_years"]:
    if c in disp_ml.columns:
        nm = "Age (years)" if c=="Age_years" else "Remaining Life (years)"
        disp_ml[nm] = disp_ml[c].apply(lambda x:"" if pd.isna(x) else f"{x:.2f}")
if "Criticality_Score" in disp_ml.columns: disp_ml["Risk Score"] = disp_ml["Criticality_Score"].apply(lambda x:f"{x:.1f}")
disp_cols = [x for x in ["Capacity (cups/min)","Utilization (%)","Start_Date","Age (years)","Remaining Life (years)","End_of_Life_Year","Risk Score"] if x in disp_ml.columns]
if disp_cols: st.table(disp_ml.set_index(machine_col)[disp_cols])

# ---------- Benchmarking ----------
st.subheader("📊 Performance Benchmarking")
industry_benchmarks = {"Utilization": 0.82, "Machine_Life": 12, "Output_per_Machine": 50000}
current_metrics = {
    "Utilization": avg_util,
    "Machine_Life": machine_life_years,
    "Output_per_Machine": machines_df["Rated_Capacity_cpm"].mean() * 60 * 16 * avg_util
}
bench_rows=[]
for m, cur in current_metrics.items():
    if m in industry_benchmarks:
        ind=industry_benchmarks[m]; var=((cur-ind)/ind*100) if ind else np.nan
        bench_rows.append({"Metric": m.replace("_"," ").title(),
                           "Current": f"{cur:.2f}" if m!="Utilization" else f"{cur:.1%}",
                           "Industry": f"{ind:.2f}" if m!="Utilization" else f"{ind:.1%}",
                           "Variance": f"{var:+.1f}%"})
if bench_rows:
    b1,b2 = st.columns(2)
    benchmark_df = pd.DataFrame(bench_rows)
    b1.table(benchmark_df.set_index("Metric"))
    metrics = [r["Metric"] for r in bench_rows]
    cur_vals = [float(r["Current"].replace("%","")) for r in bench_rows]
    ind_vals = [float(r["Industry"].replace("%","")) for r in bench_rows]
    max_vals = [max(c,i) if max(c,i)>0 else 1 for c,i in zip(cur_vals, ind_vals)]
    cur_norm = [c/m*100 for c,m in zip(cur_vals, max_vals)]
    ind_norm = [i/m*100 for i,m in zip(ind_vals, max_vals)]
    rad = go.Figure()
    rad.add_trace(go.Scatterpolar(r=cur_norm+[cur_norm[0]], theta=metrics+[metrics[0]], fill="toself", name="Current"))
    rad.add_trace(go.Scatterpolar(r=ind_norm+[ind_norm[0]], theta=metrics+[metrics[0]], fill="toself", name="Industry"))
    rad.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0,100])), showlegend=True, title="Performance vs Industry")
    b2.plotly_chart(rad, use_container_width=True)

# ---------- Executive summary ----------
st.subheader("📋 Executive Summary")
e1,e2,e3 = st.columns(3)
max_daily = float(scen_df["Daily_Output_cups"].max())
max_monthly = float(scen_df["Monthly_Output_cups"].max())
e1.write(f"**🎯 Production Capacity**  \n• Max Daily: {max_daily:,.0f} cups  \n• Max Monthly: {max_monthly:,.0f} cups")
if sales_enabled: e1.write(f"• Max Annual Revenue: ${max_monthly*12*unit_price:,.0f}")
crit_count = len(critical_machines); risk_pct = (crit_count/len(machines_df)*100) if len(machines_df)>0 else 0
e2.write(f"**⚠️ Risk Assessment**  \n• Fleet Size: {len(machines_df)}  \n• Critical Machines: {crit_count} ({risk_pct:.1f}%)")
if maintenance_due>0: e2.write(f"• Maintenance Due Soon: {maintenance_due}")
if capex_per_machine>0 and "CAPEX" in capex_df.columns:
    total_capex = float(capex_df["CAPEX"].sum())
    next_year = pd.Timestamp.today().year + 1
    next_year_capex = float(capex_df.loc[capex_df["Year"]==next_year, "CAPEX"].sum())
    e3.write(f"**💰 Investment Outlook**  \n• 10-Year CAPEX: ${total_capex:,.0f}  \n• Next Year: ${next_year_capex:,.0f}  \n• Annual Maintenance: ${len(machines_df)*annual_maintenance_per_machine:,.0f}")

# Recommendations
recs=[]
if crit_count>0: recs.append(f"• Prioritize overhaul/replacement of {crit_count} critical machines.")
if avg_util<0.80: recs.append("• Improve utilization — currently below 80%.")
if aging_machines > len(machines_df)*0.3: recs.append("• Accelerate replacement plan — high share of aging fleet.")
if not recs: recs.append("• Fleet healthy — maintain preventive maintenance cadence.")
for r in recs: st.write(r)

# ---------- Downloads ----------
st.subheader("⬇️ Exports")
report_sheets = {
    "Executive_Summary": pd.DataFrame([{
        "Total_Machines": len(machines_df),
        "Avg_Utilization": f"{avg_util:.2%}",
        "Total_Capacity_CPM": total_capacity,
        "Critical_Machines": crit_count,
        "Max_Daily_Output": max_daily,
        "Max_Monthly_Output": max_monthly,
        "Total_10Y_CAPEX": capex_df["CAPEX"].sum() if "CAPEX" in capex_df.columns else 0,
        "Report_Generated": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")
    }]),
    "Scenario_Analysis": scen_df,
    **{f"Annual_Forecast_{h}h": tables[h] for h in hour_scenarios},
    "CAPEX_Schedule": capex_df,
    "Machine_Condition": machines_df[[c for c in [machine_col,"Rated_Capacity_cpm","Utilization","Start_Date","Age_years","Remaining_Life_years","End_of_Life_Year","Criticality_Score"] if c in machines_df.columns]],
}
if sales_info:
    report_sheets["Monthly_Sales"] = sales_info["monthly"]
    if sales_info["top_customers"] is not None:
        report_sheets["Top_Customers_LTM"] = sales_info["top_customers"]
excel_bytes = to_excel_bytes(report_sheets)
c1,c2,c3 = st.columns(3)
c1.download_button("📊 Download Report (.xlsx)", data=excel_bytes,
                   file_name=f"Factory_Investor_Report_{pd.Timestamp.now():%Y%m%d}.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
c2.download_button("📈 Download Scenarios CSV",
                   data=scen_df.to_csv(index=False).encode("utf-8"),
                   file_name=f"production_scenarios_{pd.Timestamp.now():%Y%m%d}.csv", mime="text/csv")
c3.download_button("💰 Download CAPEX CSV",
                   data=capex_df.to_csv(index=False).encode("utf-8"),
                   file_name=f"capex_schedule_{pd.Timestamp.now():%Y%m%d}.csv", mime="text/csv")

st.caption("Horizon fixed at 10 years. All amounts use thousand separators; charts keep numeric data.")
