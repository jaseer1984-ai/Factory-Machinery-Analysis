# app.py — Paper Cup Factory Dashboard
# Features:
# - Upload Excel; auto-detect Machines/Production sheets
# - 12/16/20/24h forecast (baseline 12h, 28 days/month by default)
# - 10-year annual forecast (units; optional revenue/GM)
# - 10-year CAPEX replacement plan (10y life, same-year replacement)
# - Robust age detection (Age / Start Date / Year)
# - Comma-formatted tables; numeric data kept for charts & downloads

from datetime import datetime
import io
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Paper Cup Factory Dashboard", layout="wide")
st.title("📈 Paper Cup Factory — Production & Sales Forecast Dashboard")

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
    st.subheader("Optional Economics")
    unit_price = st.number_input("Unit price (per cup)", min_value=0.0, value=0.0, step=0.001, format="%.4f")
    unit_cost  = st.number_input("Unit cost (per cup)",  min_value=0.0, value=0.0, step=0.001, format="%.4f")
    capex_per_machine = st.number_input("CAPEX per machine", min_value=0.0, value=0.0, step=1000.0, format="%.0f")

# ------------- Helpers -------------

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
        tables[h] = pd.DataFrame(rows, columns=list(rows[0].keys()))
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

# string format helpers (for display tables)
def fmt_int(x):
    if x is None or (isinstance(x, float) and np.isnan(x)): return ""
    return f"{int(round(x)):,}"

def fmt_money(x, decimals=2):
    if x is None or (isinstance(x, float) and np.isnan(x)): return ""
    return f"{x:,.{decimals}f}"

# ------------- Main -------------
if uploaded is None:
    st.info("Upload your Excel to start.")
    st.stop()

wb = load_workbook(uploaded)
machines_sheet, production_sheet = detect_key_sheets(wb)
st.markdown(f"**Detected Machines sheet:** `{machines_sheet}`  |  **Production sheet:** `{production_sheet}`")

machines_df_raw = wb[machines_sheet].copy()
prod_df_raw = wb[production_sheet].copy()
machine_col, cap_col, machines_df = extract_core(machines_df_raw.copy(), prod_df_raw.copy(), machine_life_years)

# KPI cards (comma separated)
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("Machines", f"{len(machines_df):,}")
with col2:
    st.metric("Avg Utilization", f"{machines_df['Utilization'].mean():.2%}")
with col3:
    st.metric("Total Rated Capacity (cups/min)", f"{machines_df['Rated_Capacity_cpm'].sum():,.0f}")
with col4:
    avg_age = machines_df["Age_years"].dropna().mean()
    st.metric("Avg Age (years)", "N/A" if pd.isna(avg_age) else f"{avg_age:.2f}")

# Ensure hours selection
if not hour_scenarios:
    st.warning("No hours selected; defaulting to 12/16/20/24.")
    hour_scenarios = [12, 16, 20, 24]

# -------- Scenario outputs (formatted table + charts) --------
scen_df = scenario_outputs(machines_df, hour_scenarios, days_per_month)
if "Hours" in scen_df.columns:
    scen_df = scen_df.sort_values("Hours")

st.subheader("🔁 Output by Hours/Day (12h baseline, 28 days/month)")
disp_scen = scen_df.copy()
disp_scen["Daily Output (cups)"] = disp_scen["Daily_Output_cups"].apply(fmt_int)
disp_scen["Monthly Output (cups)"] = disp_scen["Monthly_Output_cups"].apply(fmt_int)
st.table(disp_scen.set_index("Hours")[["Daily Output (cups)", "Monthly Output (cups)"]])

fig_day = px.bar(
    scen_df, x="Hours", y="Daily_Output_cups",
    title="Daily Output by Hours (cups/day)",
    text=scen_df["Daily_Output_cups"].round().astype(int).map(lambda v: f"{v:,}")
); fig_day.update_layout(yaxis_tickformat=",")
st.plotly_chart(fig_day, use_container_width=True)

fig_mon = px.bar(
    scen_df, x="Hours", y="Monthly_Output_cups",
    title=f"Monthly Output by Hours (cups/month) — {days_per_month} days",
    text=scen_df["Monthly_Output_cups"].round().astype(int).map(lambda v: f"{v:,}")
); fig_mon.update_layout(yaxis_tickformat=",")
st.plotly_chart(fig_mon, use_container_width=True)

# -------- 10-year annual forecasts --------
sales_enabled = unit_price > 0 and unit_cost >= 0
tables = annual_forecast(
    machines_df, hour_scenarios, days_per_month, forecast_years,
    price=unit_price if sales_enabled else 0.0, cost=unit_cost if sales_enabled else 0.0
)

st.subheader("📅 10-Year Annual Forecast" + (" — Units, Revenue & GM" if sales_enabled else " — Units"))
tab_objs = st.tabs([f"{h}h" for h in hour_scenarios])
for tab, h in zip(tab_objs, hour_scenarios):
    with tab:
        disp = tables[h].copy()
        disp["Output (cups)"] = disp["Output_cups"].apply(fmt_int)
        cols = ["Year", "Hours_per_Day", "Output (cups)"]
        if "Revenue" in disp.columns:
            disp["Revenue"] = disp["Revenue"].apply(lambda x: fmt_money(x, 2))
            disp["Gross_Margin"] = disp["Gross_Margin"].apply(lambda x: fmt_money(x, 2))
            cols += ["Revenue", "Gross_Margin"]
        st.table(disp[cols].set_index("Year"))

        fig_line = px.line(tables[h], x="Year", y="Output_cups", markers=True, title=f"Annual Output — {h}h")
        fig_line.update_layout(yaxis_tickformat=",")
        st.plotly_chart(fig_line, use_container_width=True)

# -------- CAPEX replacement plan --------
st.subheader("🏭 CAPEX Replacement Schedule (10-year life)")
capex_df = capex_schedule(machines_df, forecast_years, machine_life_years,
                          capex_each=capex_per_machine if capex_per_machine > 0 else 0.0)
disp_capex = capex_df.copy()
if "CAPEX" in disp_capex.columns:
    disp_capex["CAPEX"] = disp_capex["CAPEX"].apply(lambda x: fmt_money(x, 0))
st.table(disp_capex.set_index("Year"))

fig_capex = px.bar(capex_df, x="Year", y="Machines_to_Replace",
                   title="Machines to Replace per Year",
                   text=capex_df["Machines_to_Replace"].astype(int).map(lambda x: f"{x:,}"))
fig_capex.update_layout(yaxis_tickformat=",")
st.plotly_chart(fig_capex, use_container_width=True)

# -------- Machine condition --------
st.subheader("🛠️ Machine Condition & Remaining Life")
cols = ["Rated_Capacity_cpm","Utilization","Start_Date","Age_years","Remaining_Life_years","End_of_Life_Year"]
machine_col = machine_col if 'machine_col' in locals() else "Machine"
show_cols = [c for c in [machine_col, *cols] if c in machines_df.columns]
disp_ml = machines_df[show_cols].copy()
if "Rated_Capacity_cpm" in disp_ml.columns:
    disp_ml["Rated_Capacity_cpm"] = disp_ml["Rated_Capacity_cpm"].apply(fmt_int)
for c in ["Age_years","Remaining_Life_years"]:
    if c in disp_ml.columns:
        disp_ml[c] = disp_ml[c].apply(lambda x: "" if pd.isna(x) else f"{x:.2f}")
st.table(disp_ml.set_index(show_cols[0]))

fig_hist = px.histogram(machines_df, x="Remaining_Life_years", nbins=10, title="Remaining Useful Life (Years)")
st.plotly_chart(fig_hist, use_container_width=True)

# -------- Downloads --------
st.subheader("⬇️ Downloads")
report_sheets = {
    "Scenario_Summary": scen_df,
    **{f"Annual_{h}h": tables[h] for h in hour_scenarios},
    "CAPEX_Schedule": capex_df,
    "Machine_Life": machines_df[show_cols],
}
excel_bytes = to_excel_bytes(report_sheets)
st.download_button("Download Excel Report (.xlsx)", data=excel_bytes,
                   file_name="Factory_Forecast_Report.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.download_button("Download Scenario CSV", data=scen_df.to_csv(index=False).encode("utf-8"),
                   file_name="scenario_summary.csv", mime="text/csv")
st.download_button("Download CAPEX CSV", data=capex_df.to_csv(index=False).encode("utf-8"),
                   file_name="capex_schedule.csv", mime="text/csv")

st.caption("Notes: robust column matching; ages from Age/Start Date/Year; "
           "immediate replacement in EOL year; numbers show thousands separators.")
