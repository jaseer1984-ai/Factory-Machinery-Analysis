# app.py
# Streamlit Investor Dashboard — Production, Sales, Scenarios, Insights
#
# How to run locally:
#   pip install -r requirements.txt
#   streamlit run app.py
#
# Works with your existing Excel template (v2):
#   C:/Users/User/OneDrive/Documents/Factory_Project.xlsx
# ...or upload a file via the sidebar.

import os
import io
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px

# ------------------ CONFIG ------------------
DEFAULT_PATH = "C:/Users/User/OneDrive/Documents/Factory_Project.xlsx"
SHIFT_HOURS_DEFAULT = 8
SCENARIO_EFF_DEFAULT = {12: 0.80, 16: 0.78, 20: 0.76, 24: 0.75}

st.set_page_config(
    page_title="Factory Investor Dashboard",
    page_icon="📊",
    layout="wide"
)

# ------------------ HELPERS ------------------
def read_workbook(file_or_path) -> dict:
    """Return dict of DataFrames keyed by sheet name."""
    xls = pd.ExcelFile(file_or_path)
    return {name: xls.parse(name) for name in xls.sheet_names}

def num(df, cols):
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df

def theoretical_from_rated(rated_cpm, hours_per_day):
    return rated_cpm * 60 * hours_per_day

def safe_sum(s: pd.Series) -> float:
    return float(pd.to_numeric(s, errors="coerce").fillna(0).sum())

def build_base_model(sheets, shift_hours=SHIFT_HOURS_DEFAULT):
    md = sheets["Machine Details"].copy()
    od = sheets["Operating Data"].copy()
    el = sheets["Efficiency & Losses"].copy()
    fi = sheets["Financial Inputs"].copy()
    sales = sheets.get("Sales Report", pd.DataFrame())

    # numeric coercion
    md = num(md, ["Rated Capacity (cups/min)","Cup Thickness (mm)"])
    od = num(od, ["Current Working Hours per Day","Working Days per Month",
                  "Theoretical Production per Day (cups)","Actual Production per Day (cups)",
                  "Actual Production per Month (cups)","Efficiency (%)"])
    el = num(el, ["Avg Downtime per Day (hrs)","Maintenance Days per Month","Wastage / Rejection Rate (%)"])
    fi = num(fi, ["Raw Material Cost per 1,000 cups (SAR)","Selling Price per 1,000 cups (SAR)",
                  "Labor Cost per Shift (SAR)","Electricity/Utility Cost per Shift (SAR)"])

    df = (md.merge(od, on="Machine Name/ID", how="left")
            .merge(el, on="Machine Name/ID", how="left")
            .merge(fi, on="Machine Name/ID", how="left"))

    # derive missing fields
    if "Theoretical Production per Day (cups)" not in df or df["Theoretical Production per Day (cups)"].isna().any():
        df["Theoretical Production per Day (cups)"] = theoretical_from_rated(
            df["Rated Capacity (cups/min)"], df["Current Working Hours per Day"]
        )
    if "Actual Production per Month (cups)" not in df or df["Actual Production per Month (cups)"].isna().any():
        df["Actual Production per Month (cups)"] = df["Actual Production per Day (cups)"] * df["Working Days per Month"]

    df["Efficiency (Actual)"] = np.where(
        df["Theoretical Production per Day (cups)"] > 0,
        df["Actual Production per Day (cups)"] / df["Theoretical Production per Day (cups)"],
        np.nan
    )

    # margins & cost
    df["Margin per 1k (SAR)"] = df["Selling Price per 1,000 cups (SAR)"] - df["Raw Material Cost per 1,000 cups (SAR)"]
    df["Shifts/day"] = df["Current Working Hours per Day"] / shift_hours
    df["Shift Cost/day (SAR)"] = (df["Labor Cost per Shift (SAR)"] + df["Electricity/Utility Cost per Shift (SAR)"]) * df["Shifts/day"]
    df["Gross Margin/day (SAR)"] = (df["Actual Production per Day (cups)"]/1000.0)*df["Margin per 1k (SAR)"] - df["Shift Cost/day (SAR)"]
    df["Gross Margin/month (SAR)"] = df["Gross Margin/day (SAR)"] * df["Working Days per Month"]

    # sales (optional)
    sales_df = pd.DataFrame()
    if "Sales Report" in sheets:
        sales_df = sheets["Sales Report"].copy()
        first_col = sales_df.columns[0]
        sales_df = sales_df.rename(columns={first_col: "Month"})
        sales_df["Month"] = pd.to_datetime(sales_df["Month"].astype(str) + "-01", errors="coerce")

    return df, sales_df

def project_hours(df, target_hours, eff_assumption, shift_hours=SHIFT_HOURS_DEFAULT):
    rated = df["Rated Capacity (cups/min)"]
    days = df["Working Days per Month"]
    cur_hours = df["Current Working Hours per Day"]
    projected_day = rated * 60 * target_hours * eff_assumption
    projected_month = projected_day * days
    scale = np.where(cur_hours > 0, target_hours / cur_hours, 0.0)
    proj_shift_cost_day = df["Shift Cost/day (SAR)"] * scale
    proj_gm_day = (projected_day/1000.0)*df["Margin per 1k (SAR)"] - proj_shift_cost_day
    proj_gm_month = proj_gm_day * days
    return projected_day, projected_month, proj_gm_month

def portfolio_agg(df, scenarios_eff):
    agg = {
        "Current/day": df["Actual Production per Day (cups)"].sum(),
        "Current/month": df["Actual Production per Month (cups)"].sum(),
        "Current GM/month (SAR)": df["Gross Margin/month (SAR)"].sum()
    }
    for hrs, eff in scenarios_eff.items():
        _, proj_mon, proj_gm_mon = project_hours(df, hrs, eff)
        agg[f"{hrs}h/month"] = proj_mon.sum()
        agg[f"{hrs}h GM/month (SAR)"] = proj_gm_mon.sum()
    return agg

def sales_gap(df, sales_df):
    current_monthly_prod = df["Actual Production per Month (cups)"].sum()
    if sales_df.empty or "Sales Quantity (cups)" not in sales_df.columns:
        return dict(has_sales=False, current_monthly_production=current_monthly_prod,
                    avg_monthly_sales=None, gap_cups=None)
    n_months = sales_df["Month"].nunique()
    total_sales_qty = sales_df["Sales Quantity (cups)"].sum()
    avg_monthly_sales = total_sales_qty / max(1, n_months)
    gap = avg_monthly_sales - current_monthly_prod
    return dict(has_sales=True, current_monthly_production=current_monthly_prod,
                avg_monthly_sales=avg_monthly_sales, gap_cups=gap)

def insights_text(df, agg, gap_info):
    lines = []
    machines = len(df)
    total_cpm = df["Rated Capacity (cups/min)"].sum()
    avg_eff = df["Efficiency (Actual)"].mean(skipna=True)
    cur_month = agg["Current/month"]
    gm_month = agg["Current GM/month (SAR)"]

    lines.append(f"• {machines} machines | Rated capacity: {int(total_cpm):,} cups/min | Avg utilization: {avg_eff:.1%}")
    lines.append(f"• Output: {int(cur_month):,} cups/month | Gross margin: {gm_month:,.0f} SAR/month")
    for label in [12,16,20,24]:
        inc = agg[f"{label}h/month"] - cur_month
        inc_gm = agg[f"{label}h GM/month (SAR)"] - gm_month
        lines.append(f"• {label}h scenario → {int(agg[f'{label}h/month']):,} cups/month "
                     f"(+{int(inc):,}) | GM {agg[f'{label}h GM/month (SAR)']:,.0f} SAR (+{inc_gm:,.0f})")
    if gap_info["has_sales"]:
        gap = gap_info["gap_cups"]
        if gap > 0:
            lines.append(f"• Sales exceed production by {int(gap):,} cups/month — increase hours and/or reduce downtime to close gap.")
        else:
            lines.append(f"• Production exceeds sales by {int(-gap):,} cups/month — focus on demand/pricing before adding hours.")
    else:
        lines.append("• Sales sheet not found — add Sales Report to compare demand vs supply.")
    return "\n".join(lines)

def to_excel_download(df_summary, agg_df, sales_df, insights):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_summary.to_excel(writer, sheet_name="Machine Summary", index=False)
        agg_df.to_excel(writer, sheet_name="Scenarios (Aggregate)", index=False)
        if not sales_df.empty:
            sales_df.to_excel(writer, sheet_name="Sales Report (Raw)", index=False)
        pd.DataFrame({"Insights":[insights]}).to_excel(writer, sheet_name="Insights & Recos", index=False)
    output.seek(0)
    return output

# ------------------ SIDEBAR ------------------
st.sidebar.header("⚙️ Data & Assumptions")
uploaded = st.sidebar.file_uploader("Upload Excel (v2 template)", type=["xlsx"])
use_default = st.sidebar.checkbox("Use default path", value=True, help=DEFAULT_PATH)
excel_path = DEFAULT_PATH if use_default else st.sidebar.text_input("Or paste Excel path", value=DEFAULT_PATH)

st.sidebar.markdown("---")
st.sidebar.subheader("Shift & Scenario Settings")
shift_hours = st.sidebar.number_input("Shift length (hours)", 4, 12, SHIFT_HOURS_DEFAULT, step=1)
eff_12 = st.sidebar.slider("Efficiency @ 12h", 0.60, 0.95, SCENARIO_EFF_DEFAULT[12], 0.01)
eff_16 = st.sidebar.slider("Efficiency @ 16h", 0.60, 0.95, SCENARIO_EFF_DEFAULT[16], 0.01)
eff_20 = st.sidebar.slider("Efficiency @ 20h", 0.60, 0.95, SCENARIO_EFF_DEFAULT[20], 0.01)
eff_24 = st.sidebar.slider("Efficiency @ 24h", 0.60, 0.95, SCENARIO_EFF_DEFAULT[24], 0.01)
scenarios_eff = {12: eff_12, 16: eff_16, 20: eff_20, 24: eff_24}

st.sidebar.markdown("---")
st.sidebar.info("Tip: Upload Sales Report to unlock Sales vs Production comparison.")

# ------------------ LOAD DATA ------------------
if uploaded is not None:
    sheets = read_workbook(uploaded)
else:
    try:
        sheets = read_workbook(excel_path)
    except Exception as e:
        st.error(f"Could not read Excel. Upload a file or check path. Details: {e}")
        st.stop()

# ------------------ BUILD MODEL ------------------
df, sales_df = build_base_model(sheets, shift_hours=shift_hours)
agg = portfolio_agg(df, scenarios_eff)
gap = sales_gap(df, sales_df)

# ------------------ HEADER ------------------
st.title("📊 Factory Investor Dashboard")
st.caption("Production capacity, efficiency, sales comparison, and hour-increase scenarios")

# ------------------ KPI CARDS ------------------
kpi1, kpi2, kpi3, kpi4 = st.columns(4)
kpi1.metric("Machines", f"{len(df):,}")
kpi2.metric("Rated Capacity (cups/min)", f"{int(df['Rated Capacity (cups/min)'].sum()):,}")
kpi3.metric("Avg Utilization", f"{df['Efficiency (Actual)'].mean(skipna=True):.1%}")
kpi4.metric("GM / Month (SAR)", f"{agg['Current GM/month (SAR)']:,.0f}")

kpi5, kpi6, kpi7, kpi8 = st.columns(4)
kpi5.metric("Actual / Day (cups)", f"{int(df['Actual Production per Day (cups)'].sum()):,}")
kpi6.metric("Actual / Month (cups)", f"{int(agg['Current/month']):,}")
kpi7.metric("Idle / Day (cups)", f"{int((df['Theoretical Production per Day (cups)'] - df['Actual Production per Day (cups)']).clip(lower=0).sum()):,}")
kpi8.metric("Shift Hours", f"{shift_hours}h")

st.markdown("---")

# ------------------ CHARTS ------------------
colA, colB = st.columns(2)
with colA:
    fig = px.bar(
        x=["Theoretical/day","Actual/day"],
        y=[df["Theoretical Production per Day (cups)"].sum(),
           df["Actual Production per Day (cups)"].sum()],
        labels={"x":"","y":"Cups per day"},
        title="Capacity vs Actual (Aggregate)"
    )
    st.plotly_chart(fig, use_container_width=True)

with colB:
    fig = px.histogram(
        df, x="Efficiency (Actual)",
        nbins=12, title="Utilization Distribution (Actual/Theoretical)"
    )
    st.plotly_chart(fig, use_container_width=True)

colC, colD = st.columns(2)
with colC:
    scen_labels = ["Current","12h","16h","20h","24h"]
    scen_vals = [agg["Current/month"], agg["12h/month"], agg["16h/month"], agg["20h/month"], agg["24h/month"]]
    fig = px.bar(
        x=scen_labels, y=scen_vals,
        labels={"x":"Scenario","y":"Cups per month"},
        title="Monthly Output by Scenario (Aggregate)"
    )
    st.plotly_chart(fig, use_container_width=True)

with colD:
    if gap["has_sales"]:
        sdf = sales_df.sort_values("Month")
        fig = px.line(
            sdf, x="Month", y="Sales Quantity (cups)",
            title="Sales vs Production (Monthly)"
        )
        fig.add_hline(y=agg["Current/month"], line_dash="dash", annotation_text="Current Production (monthly)")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Upload Sales Report to view Sales vs Production trend.")

st.markdown("---")

# ------------------ INSIGHTS ------------------
st.subheader("🧠 Insights & Recommendations")
insight = insights_text(df, agg, gap)
st.text(insight)

# ------------------ MACHINE TABLE ------------------
st.subheader("📋 Machine Summary (sorted by margin)")
show_cols = [
    "Machine Name/ID","Machine Type","Condition","Cup Thickness (mm)",
    "Rated Capacity (cups/min)","Current Working Hours per Day","Working Days per Month",
    "Theoretical Production per Day (cups)","Actual Production per Day (cups)","Efficiency (Actual)",
    "Actual Production per Month (cups)",
    "Raw Material Cost per 1,000 cups (SAR)","Selling Price per 1,000 cups (SAR)",
    "Margin per 1k (SAR)","Shifts/day","Shift Cost/day (SAR)","Gross Margin/day (SAR)","Gross Margin/month (SAR)",
    "Avg Downtime per Day (hrs)","Wastage / Rejection Rate (%)","Downtime Reasons"
]
table_df = (df[show_cols].copy()
            .sort_values("Gross Margin/month (SAR)", ascending=False))
st.dataframe(table_df, use_container_width=True)

# ------------------ DOWNLOADS ------------------
st.markdown("### ⬇️ Export")
agg_df = pd.DataFrame([
    {"Scenario":"Current","Monthly Output (cups)":agg["Current/month"],"GM/month (SAR)":agg["Current GM/month (SAR)"]},
    {"Scenario":"12h","Monthly Output (cups)":agg["12h/month"],"GM/month (SAR)":agg["12h GM/month (SAR)"]},
    {"Scenario":"16h","Monthly Output (cups)":agg["16h/month"],"GM/month (SAR)":agg["16h GM/month (SAR)"]},
    {"Scenario":"20h","Monthly Output (cups)":agg["20h/month"],"GM/month (SAR)":agg["20h GM/month (SAR)"]},
    {"Scenario":"24h","Monthly Output (cups)":agg["24h/month"],"GM/month (SAR)":agg["24h GM/month (SAR)"]},
])
xls_bytes = to_excel_download(table_df, agg_df, sales_df, insight)
st.download_button(
    label="Download Investor_Project_Report.xlsx",
    data=xls_bytes,
    file_name="Investor_Project_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.caption("Tip: push this repo to GitHub, then deploy on Streamlit Community Cloud.")
