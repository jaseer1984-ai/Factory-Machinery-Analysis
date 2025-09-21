# app.py — Paper Cup Factory Dashboard (robust ages + comma formatting + hours fallback)

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
        "Hours per day (select)", options=[12,16,20,24], default=[12,16,20,24]
    )
    machine_life_years = st.number_input("Machine life (years)", 1, 40, value=10)
    forecast_years = st.number_input("Forecast horizon (years)", 1, 30, value=10)

    st.divider()
    st.subheader("Optional Economics")
    unit_price = st.number_input("Unit price (per cup)", min_value=0.0, value=0.0, step=0.001, format="%.4f")
    unit_cost  = st.number_input("Unit cost (per cup)",  min_value=0.0, value=0.0, step=0.001, format="%.4f")
    capex_per_machine = st.number_input("CAPEX per machine", min_value=0.0, value=0.0, step=1000.0, format="%.0f")

# ------------- Helpers -------------
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
    # Machines
    cand_m = [s for s in names if any(k in s.lower() for k in
             ["machine","assets","equipment","plant","capacity","list","details"])]
    machines_sheet = cand_m[0] if cand_m else names[0]
    # Production
    cand_p = [s for s in names if any(k in s.lower() for k in
             ["prod","output","shift","daily","util","run","report","sales"]) and s != machines_sheet]
    production_sheet = cand_p[0] if cand_p else (names[1] if len(names)>1 else names[0])
    return machines_sheet, production_sheet

def extract_core(machines_df: pd.DataFrame, prod_df: pd.DataFrame, life_years: float):
    lm = {c.lower(): c for c in machines_df.columns}
    lp = {c.lower(): c for c in prod_df.columns}

    machine_col = first_present(lm, ["machine","machine id","machine_name","name","id"]) or "Machine"
    if machine_col not in machines_df.columns:
        machines_df["Machine"] = [f"M{i+1}" for i in range(len(machines_df))]
        machine_col = "Machine"

    # Capacity
    cap_col = first_present(lm, [
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
    util_col = first_present(lm, ["utilization","util","uptime %","uptime","availability","runtime%"])
    if util_col is None:
        util_p = first_present(lp, ["utilization","util","uptime %","uptime","availability","runtime%"])
        if util_p is not None and not prod_df.empty:
            u = pd.to_numeric(prod_df[util_p].astype(str).replace("%","",regex=True), errors="coerce")/100.0
            machines_df["Utilization"] = float(u.fillna(u.mean()).clip(0,1).mean())
        else:
            machines_df["Utilization"] = 0.955
    else:
        u = machines_df[util_col]
        if u.dtype == object:
            u = pd.to_numeric(u.astype(str).str.replace("%","",regex=False), errors="coerce")/100.0
        machines_df["Utilization"] = u.fillna(u.mean()).clip(0,1)

    # ----- Age handling (robust) -----
    # 1) If we have a direct Age column, use it.
    age_col = first_present(lm, ["age","age (years)","age_years"])
    # 2) Otherwise try Start/Commission/Install Date.
    start_col = first_present(lm, [
        "start date","commission date","commission_date","installed on","install date",
        "start_date","commissioning date","commission year","commission_year"
    ])
    # 3) Otherwise try Year columns (MFG year, installation year, etc.)
    year_col = first_present(lm, [
        "year","mfg year","manufacture year","manufacturing year","year of make",
        "purchase year","installation year","install year","yom","yop"
    ])

    today = pd.Timestamp.today().normalize()

    if age_col is not None:
        machines_df["Age_years"] = pd.to_numeric(machines_df[age_col], errors="coerce")
        # fabricate Start_Date from Age if missing
        machines_df["Start_Date"] = today - pd.to_timedelta((machines_df["Age_years"]*365.25), unit="D")
    elif start_col is not None:
        machines_df["Start_Date"] = parse_date_series(machines_df[start_col])
        age_days = (today - machines_df["Start_Date"]).dt.days
        machines_df["Age_years"] = (age_days/365.25)
    elif year_col is not None:
        # build a Start_Date from January 1st of the given year
        yr = pd.to_numeric(machines_df[year_col], errors="coerce").round().astype("Int64")
        machines_df["Start_Date"] = pd.to_datetime(yr.astype("float").astype("Int64").astype(str) + "-01-01", errors="coerce")
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
    # Always return columns to avoid KeyErrors on empty selection
    out = pd.DataFrame(rows, columns=["Hours","Daily_Output_cups","Monthly_Output_cups"])
    return out

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
    hour_scenarios = [12,16,20,24]

# Scenario outputs (with comma formatting in UI)
scen_df = scenario_outputs(machines_df, hour_scenarios, days_per_month)
if "Hours" in scen_df.columns:
    scen_df = scen_df.sort_values("Hours")

st.subheader("🔁 Output by Hours/Day (12h baseline, 28 days/month)")
st.dataframe(
    scen_df,
    use_container_width=True,
    column_config={
        "Hours": st.column_config.NumberColumn("Hours/Day", format="%d"),
        "Daily_Output_cups": st.column_config.NumberColumn("Daily Output (cups)", format="%,.0f"),
        "Monthly_Output_cups": st.column_config.NumberColumn("Monthly Output (cups)", format="%,.0f"),
    },
)

fig_day = px.bar(
    scen_df, x="Hours", y="Daily_Output_cups",
    title="Daily Output by Hours (cups/day)",
    text=scen_df["Daily_Output_cups"].round().astype(int).map(lambda x: f"{x:,}")
)
fig_day.update_layout(yaxis_tickformat=",")
st.plotly_chart(fig_day, use_container_width=True)

fig_mon = px.bar(
    scen_df, x="Hours", y="Monthly_Output_cups",
    title=f"Monthly Output by Hours (cups/month) — {days_per_month} days",
    text=scen_df["Monthly_Output_cups"].round().astype(int).map(lambda x: f"{x:,}")
)
fig_mon.update_layout(yaxis_tickformat=",")
st.plotly_chart(fig_mon, use_container_width=True)

# 10-year annual forecasts (comma formatting)
sales_enabled = unit_price > 0 and unit_cost >= 0
tables = annual_forecast(
    machines_df, hour_scenarios, days_per_month, forecast_years,
    price=unit_price if sales_enabled else 0.0, cost=unit_cost if sales_enabled else 0.0
)

st.subheader("📅 10-Year Annual Forecast" + (" — Units, Revenue & GM" if sales_enabled else " — Units"))
tab_objs = st.tabs([f"{h}h" for h in hour_scenarios])
for tab, h in zip(tab_objs, hour_scenarios):
    with tab:
        cfg = {
            "Year": st.column_config.NumberColumn("Year", format="%d"),
            "Hours_per_Day": st.column_config.NumberColumn("Hours/Day", format="%d"),
            "Output_cups": st.column_config.NumberColumn("Output (cups)", format="%,.0f"),
        }
        if sales_enabled:
            cfg["Revenue"] = st.column_config.NumberColumn("Revenue", format="%,.2f")
            cfg["Gross_Margin"] = st.column_config.NumberColumn("Gross Margin", format="%,.2f")
        st.dataframe(tables[h], use_container_width=True, column_config=cfg)

        fig_line = px.line(tables[h], x="Year", y="Output_cups", markers=True, title=f"Annual Output — {h}h")
        fig_line.update_layout(yaxis_tickformat=",")
        st.plotly_chart(fig_line, use_container_width=True)

# CAPEX replacement plan (comma formatting)
st.subheader("🏭 CAPEX Replacement Schedule (10-year life)")
capex_df = capex_schedule(machines_df, forecast_years, machine_life_years, capex_each=capex_per_machine if capex_per_machine>0 else 0.0)
st.dataframe(
    capex_df,
    use_container_width=True,
    column_config={
        "Year": st.column_config.NumberColumn("Year", format="%d"),
        "Machines_to_Replace": st.column_config.NumberColumn("Machines to Replace", format="%,d"),
        **({"CAPEX": st.column_config.NumberColumn("CAPEX", format="%,.0f")} if "CAPEX" in capex_df.columns else {})
    },
)
fig_capex = px.bar(capex_df, x="Year", y="Machines_to_Replace", title="Machines to Replace per Year",
                   text=capex_df["Machines_to_Replace"].astype(int).map(lambda x: f"{x:,}"))
fig_capex.update_layout(yaxis_tickformat=",")
st.plotly_chart(fig_capex, use_container_width=True)

# Machine condition table (comma formatting)
st.subheader("🛠️ Machine Condition & Remaining Life")
cols = ["Rated_Capacity_cpm","Utilization","Start_Date","Age_years","Remaining_Life_years","End_of_Life_Year"]
show_cols = [c for c in [machine_col, *cols] if c in machines_df.columns]
st.dataframe(
    machines_df[show_cols],
    use_container_width=True,
    column_config={
        machine_col: st.column_config.Column(machine_col),
        "Rated_Capacity_cpm": st.column_config.NumberColumn("Capacity (cups/min)", format="%,.0f"),
        "Utilization": st.column_config.NumberColumn("Utilization", format="%.2f"),
        "Age_years": st.column_config.NumberColumn("Age (years)", format="%.2f"),
        "Remaining_Life_years": st.column_config.NumberColumn("Remaining Life (years)", format="%.2f"),
        "End_of_Life_Year": st.column_config.NumberColumn("EOL Year", format="%d"),
    },
)

# Downloads
st.subheader("⬇️ Downloads")
report_sheets = {
    "Scenario_Summary": scen_df,
    **{f"Annual_{h}h": tables[h] for h in hour_scenarios},
    "CAPEX_Schedule": capex_df,
    "Machine_Life": machines_df[show_cols],
}
excel_bytes = to_excel_bytes(report_sheets)
st.download_button("Download Excel Report (.xlsx)", data=excel_bytes, file_name="Factory_Forecast_Report.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.caption("Notes: (1) If no dates are provided, age is estimated from Age or Year columns if present. "
           "(2) Numbers show thousands separators. (3) If no hours are selected, app defaults to 12/16/20/24.")
