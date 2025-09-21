# app.py
import io
from datetime import datetime
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# ------------------------
# Page / Sidebar controls
# ------------------------
st.set_page_config(page_title="Paper Cup Factory Dashboard", layout="wide")
st.title("📈 Paper Cup Factory — Production & Sales Forecast Dashboard")

with st.sidebar:
    st.header("⚙️ Settings")
    uploaded = st.file_uploader("Upload factory Excel", type=["xlsx"])
    days_per_month = st.number_input("Days per month", 1, 31, value=28)
    hour_scenarios = st.multiselect(
        "Hours per day (select)", options=[12,16,20,24],
        default=[12,16,20,24]
    )
    machine_life_years = st.number_input("Machine life (years)", 1, 40, value=10)
    forecast_years = st.number_input("Forecast horizon (years)", 1, 30, value=10)

    st.divider()
    st.subheader("Optional Economics")
    unit_price = st.number_input("Unit price (per cup)", min_value=0.0, value=0.0, step=0.001, format="%.4f")
    unit_cost  = st.number_input("Unit cost (per cup)",  min_value=0.0, value=0.0, step=0.001, format="%.4f")
    capex_per_machine = st.number_input("CAPEX per machine (currency)", min_value=0.0, value=0.0, step=1000.0)

    st.divider()
    st.caption("Tip: If your workbook has a Machines sheet and a Production sheet, the app will auto-detect them. It looks for columns like capacity (cups/min), utilization, and start/commission date.")

# ---------------
# Helper funcs
# ---------------
def first_present(cols, options):
    for o in options:
        for c in cols:
            if o == str(c).strip().lower():
                return c
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
    out = s.apply(parse_one)
    return pd.to_datetime(out, errors="coerce")

@st.cache_data(show_spinner=False)
def load_workbook(file_bytes: bytes) -> dict:
    xls = pd.ExcelFile(file_bytes)
    sheets = {s: pd.read_excel(xls, sheet_name=s) for s in xls.sheet_names}
    for s, df in sheets.items():
        df.columns = [str(c).strip() for c in df.columns]
    return sheets

def detect_key_sheets(sheets: dict):
    names = list(sheets.keys())
    # machines sheet
    cand_m = [s for s in names if any(k in s.lower() for k in
             ["machine","assets","equipment","plant","capacity","list"])]
    machines_sheet = cand_m[0] if cand_m else names[0]

    # production sheet
    cand_p = [s for s in names if any(k in s.lower() for k in
             ["prod","output","shift","daily","util","run","report"]) and s != machines_sheet]
    production_sheet = cand_p[0] if cand_p else (names[1] if len(names)>1 else names[0])
    return machines_sheet, production_sheet

def extract_core(machines_df: pd.DataFrame, prod_df: pd.DataFrame, life_years: float):
    lm = {c.lower(): c for c in machines_df.columns}
    lp = {c.lower(): c for c in prod_df.columns}

    machine_col = first_present(lm, ["machine","machine id","machine_name","name","id"]) or "Machine"
    if machine_col not in machines_df.columns:
        machines_df["Machine"] = [f"M{i+1}" for i in range(len(machines_df))]
        machine_col = "Machine"

    cap_col = first_present(lm, ["capacity","rated capacity","rated_capacity","cups/min","cups per min","capacity (cups/min)","capacity_cups_min","capacity_cup_min"])
    util_col = first_present(lm, ["utilization","util","uptime %","uptime","availability","runtime%"])
    start_col = first_present(lm, ["start date","commission date","commission_date","installed on","install date","start_date","commissioning date","commissioned on"])

    # Capacity
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
    if util_col is None:
        util_p = first_present(lp, ["utilization","util","uptime %","uptime","availability","runtime%"])
        if util_p is not None and not prod_df.empty:
            u = pd.to_numeric(prod_df[util_p].astype(str).str.replace("%","",regex=False), errors="coerce")/100.0
            util_val = float(u.fillna(u.mean()).clip(0,1).mean())
            machines_df["Utilization"] = util_val
        else:
            machines_df["Utilization"] = 0.955
    else:
        u = machines_df[util_col]
        if u.dtype == object:
            u = pd.to_numeric(u.astype(str).str.replace("%","",regex=False), errors="coerce")/100.0
        machines_df["Utilization"] = u.fillna(u.mean()).clip(0,1)

    # Start Date / Age
    if start_col is None:
        machines_df["Start_Date"] = pd.NaT
    else:
        machines_df["Start_Date"] = parse_date_series(machines_df[start_col])

    today = pd.Timestamp.today().normalize()
    age_days = (today - machines_df["Start_Date"]).dt.days
    machines_df["Age_years"] = (age_days/365.25).round(2)
    machines_df["Remaining_Life_years"] = (life_years - machines_df["Age_years"]).round(2)
    machines_df.loc[machines_df["Remaining_Life_years"] < 0, "Remaining_Life_years"] = 0
    machines_df["End_of_Life_Year"] = machines_df["Start_Date"].dt.year + life_years
    machines_df.loc[machines_df["Start_Date"].isna(), "End_of_Life_Year"] = np.nan

    return machine_col, cap_col, machines_df

def scenario_outputs(machines_df: pd.DataFrame, hours: list[int], dpm: int):
    util = float(machines_df["Utilization"].mean())
    base_cpm = machines_df["Rated_Capacity_cpm"].sum()
    rows = []
    for h in hours:
        cups_day = base_cpm * 60 * h * util
        cups_month = cups_day * dpm
        rows.append({"Hours": h, "Daily_Output_cups": cups_day, "Monthly_Output_cups": cups_month})
    return pd.DataFrame(rows)

def annual_forecast(machines_df: pd.DataFrame, hours: list[int], dpm: int, years: int, price=0.0, cost=0.0):
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
            row = {"Year": y, "Hours_per_Day": h, "Output_cups": cups_year}
            if price > 0 and cost >= 0:
                row["Revenue"] = cups_year * price
                row["Gross_Margin"] = cups_year * (price - cost)
            rows.append(row)
        tables[h] = pd.DataFrame(rows)
    return tables

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

# ---------------------
# Main body
# ---------------------
if uploaded is None:
    st.info("Upload your Excel to start.")
    st.stop()

# Load workbook
wb = load_workbook(uploaded)
machines_sheet, production_sheet = detect_key_sheets(wb)

st.markdown(f"**Detected Machines sheet:** `{machines_sheet}`  |  **Production sheet:** `{production_sheet}`")

machines_df_raw = wb[machines_sheet].copy()
prod_df_raw = wb[production_sheet].copy()

# Core extraction
machine_col, cap_col, machines_df = extract_core(machines_df_raw.copy(), prod_df_raw.copy(), machine_life_years)

# KPI cards
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("Machines", len(machines_df))
with col2:
    st.metric("Avg Utilization", f"{machines_df['Utilization'].mean():.2%}")
with col3:
    st.metric("Total Rated Capacity (cups/min)", f"{machines_df['Rated_Capacity_cpm'].sum():,.0f}")
with col4:
    st.metric("Avg Age (years)", f"{machines_df['Age_years'].mean():.2f}")

# Scenario outputs
scen_df = scenario_outputs(machines_df, hour_scenarios, days_per_month).sort_values("Hours")
st.subheader("🔁 Output by Hours/Day (12h baseline, 28 days/month)")
st.dataframe(scen_df, use_container_width=True)

fig_day = px.bar(scen_df, x="Hours", y="Daily_Output_cups", title="Daily Output by Hours (cups/day)", text=scen_df["Daily_Output_cups"].round().astype(int))
fig_mon = px.bar(scen_df, x="Hours", y="Monthly_Output_cups", title=f"Monthly Output by Hours (cups/month) — {days_per_month} days", text=scen_df["Monthly_Output_cups"].round().astype(int))
st.plotly_chart(fig_day, use_container_width=True)
st.plotly_chart(fig_mon, use_container_width=True)

# 10-year annual forecasts
sales_enabled = unit_price > 0 and unit_cost >= 0
tables = annual_forecast(machines_df, hour_scenarios, days_per_month, forecast_years, price=unit_price if sales_enabled else 0.0, cost=unit_cost if sales_enabled else 0.0)

st.subheader("📅 10-Year Annual Forecast (Units" + (", Revenue & GM" if sales_enabled else "") + ")")
tab_objs = st.tabs([f"{h}h" for h in hour_scenarios])
for tab, h in zip(tab_objs, hour_scenarios):
    with tab:
        st.dataframe(tables[h], use_container_width=True)
        fig_line = px.line(tables[h], x="Year", y="Output_cups", markers=True, title=f"Annual Output — {h}h")
        st.plotly_chart(fig_line, use_container_width=True)

# CAPEX replacement plan
st.subheader("🏭 CAPEX Replacement Schedule (10-year life)")
capex_df = capex_schedule(machines_df, forecast_years, machine_life_years, capex_each=capex_per_machine if capex_per_machine>0 else 0.0)
st.dataframe(capex_df, use_container_width=True)
fig_capex = px.bar(capex_df, x="Year", y="Machines_to_Replace", title="Machines to Replace per Year", text="Machines_to_Replace")
st.plotly_chart(fig_capex, use_container_width=True)

# Machine condition / life
st.subheader("🛠️ Machine Condition & Remaining Life")
cols = ["Rated_Capacity_cpm","Utilization","Start_Date","Age_years","Remaining_Life_years","End_of_Life_Year"]
show_cols = [c for c in [machine_col, *cols] if c in machines_df.columns]
st.dataframe(machines_df[show_cols], use_container_width=True)

fig_hist = px.histogram(machines_df, x="Remaining_Life_years", nbins=10, title="Remaining Useful Life (Years) — Distribution")
st.plotly_chart(fig_hist, use_container_width=True)

fig_cap = px.bar(machines_df.sort_values("Rated_Capacity_cpm", ascending=False), x=machine_col, y="Rated_Capacity_cpm", title="Rated Capacity by Machine (cups/min)")
st.plotly_chart(fig_cap, use_container_width=True)

# ---------------------
# Downloads
# ---------------------
st.subheader("⬇️ Downloads")
report_sheets = {
    "Scenario_Summary": scen_df,
    **{f"Annual_{h}h": tables[h] for h in hour_scenarios},
    "CAPEX_Schedule": capex_df,
    "Machine_Life": machines_df[show_cols]
}
excel_bytes = to_excel_bytes(report_sheets)
st.download_button("Download Excel Report (.xlsx)", data=excel_bytes, file_name="Factory_Forecast_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.download_button("Download Scenario CSV", data=scen_df.to_csv(index=False).encode("utf-8"), file_name="scenario_summary.csv", mime="text/csv")
st.download_button("Download CAPEX CSV", data=capex_df.to_csv(index=False).encode("utf-8"), file_name="capex_schedule.csv", mime="text/csv")

st.caption("Assumptions: output scales linearly with hours; replacement is immediate in EOL year; downstream processes are unconstrained. Adjust price/cost and CAPEX in the sidebar to populate financials.")
