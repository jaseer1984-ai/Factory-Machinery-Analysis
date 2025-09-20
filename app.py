import os
import io
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ------------------ CONFIG ------------------
DEFAULT_PATH = "C:/Users/User/OneDrive/Documents/Factory_Project.xlsx"
SHIFT_HOURS_DEFAULT = 8
SCENARIO_EFF_DEFAULT = {12: 0.80, 16: 0.78, 20: 0.76, 24: 0.75}

st.set_page_config(page_title="🏭 Factory Analytics Hub", page_icon="📊", layout="wide")

# ------------------ STYLE ------------------
def load_custom_css():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');
    .main .block-container { padding-top: 1.25rem; padding-bottom: 2rem; font-family: 'Inter', sans-serif; }
    .main-header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      padding: 1.25rem 1.5rem; border-radius: 14px; color: #fff; margin-bottom: 1rem; }
    .kpi { background:#fff; border-left:4px solid #667eea; border-radius:12px; padding:1rem 1.1rem;
      box-shadow:0 4px 14px rgba(0,0,0,.06); }
    .kpi .v { font-weight:700; font-size:1.6rem; color:#2c3e50; }
    .kpi .l { font-size:.85rem; color:#7f8c8d; text-transform:uppercase; letter-spacing:.4px; }
    .section { background:linear-gradient(90deg, #f8f9fa, #eef2f7); border-left:4px solid #667eea;
      padding:.7rem 1rem; border-radius:10px; margin:1rem 0 .6rem 0; font-weight:600; }
    .panel { background:#fff; border-radius:12px; padding:1rem; box-shadow:0 4px 14px rgba(0,0,0,.06); }
    </style>
    """, unsafe_allow_html=True)

def kpi_card(label, value, suffix=""):
    st.markdown(f"""<div class='kpi'><div class='v'>{value:,.0f}{suffix}</div>
    <div class='l'>{label}</div></div>""", unsafe_allow_html=True)

# ------------------ HELPERS ------------------
def read_workbook(file_or_path) -> dict:
    xls = pd.ExcelFile(file_or_path)
    return {name: xls.parse(name) for name in xls.sheet_names}

def num(df, cols):
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df

def theoretical_from_rated(rated_cpm, hours_per_day):
    return rated_cpm * 60 * hours_per_day

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

# ------------------ BASE MODEL ------------------
def build_base_model(sheets, shift_hours=SHIFT_HOURS_DEFAULT):
    md = sheets["Machine Details"].copy()
    od = sheets["Operating Data"].copy()
    el = sheets["Efficiency & Losses"].copy()
    fi = sheets["Financial Inputs"].copy()
    sales = sheets.get("Sales Report", pd.DataFrame())

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

    df["Margin per 1k (SAR)"] = df["Selling Price per 1,000 cups (SAR)"] - df["Raw Material Cost per 1,000 cups (SAR)"]
    df["Shifts/day"] = df["Current Working Hours per Day"] / shift_hours
    df["Shift Cost/day (SAR)"] = (df["Labor Cost per Shift (SAR)"] + df["Electricity/Utility Cost per Shift (SAR)"]) * df["Shifts/day"]
    df["Gross Margin/day (SAR)"] = (df["Actual Production per Day (cups)"]/1000.0)*df["Margin per 1k (SAR)"] - df["Shift Cost/day (SAR)"]
    df["Gross Margin/month (SAR)"] = df["Gross Margin/day (SAR)"] * df["Working Days per Month"]

    sales_df = pd.DataFrame()
    if "Sales Report" in sheets:
        sales_df = sheets["Sales Report"].copy()
        first_col = sales_df.columns[0]
        sales_df = sales_df.rename(columns={first_col: "Month"})
        sales_df["Month"] = pd.to_datetime(sales_df["Month"].astype(str) + "-01", errors="coerce")

    return df, sales_df

# ------------------ CHARTS (existing) ------------------
def charts_core(df, agg, sales_df, gap):
    charts = {}

    # Capacity vs Actual per machine
    fig1 = go.Figure()
    fig1.add_trace(go.Bar(
        name='Theoretical Capacity',
        x=[f"M{i+1}" for i in range(len(df))],
        y=df["Theoretical Production per Day (cups)"],
        marker_color='rgba(102, 126, 234, 0.6)',
        marker_line=dict(color='rgba(102, 126, 234, 1)', width=2)))
    fig1.add_trace(go.Bar(
        name='Actual Production',
        x=[f"M{i+1}" for i in range(len(df))],
        y=df["Actual Production per Day (cups)"],
        marker_color='rgba(118, 75, 162, 0.8)',
        marker_line=dict(color='rgba(118, 75, 162, 1)', width=2)))
    fig1.update_layout(barmode='group', title="🏭 Production Capacity vs Actual Output",
                       height=460, legend_orientation="h")
    charts['cap_vs_actual'] = fig1

    # Utilization histogram
    fig2 = px.histogram(df, x="Efficiency (Actual)", nbins=15, title="⚡ Utilization Distribution")
    charts['util_hist'] = fig2

    # Scenario aggregates
    fig3 = make_subplots(specs=[[{"secondary_y": True}]])
    scenarios = ["Current", "12h", "16h", "20h", "24h"]
    outputs = [agg["Current/month"], agg["12h/month"], agg["16h/month"], agg["20h/month"], agg["24h/month"]]
    margins = [agg["Current GM/month (SAR)"], agg["12h GM/month (SAR)"], agg["16h GM/month (SAR)"],
               agg["20h GM/month (SAR)"], agg["24h GM/month (SAR)"]]
    fig3.add_trace(go.Bar(x=scenarios, y=outputs, name="Monthly Output"), secondary_y=False)
    fig3.add_trace(go.Scatter(x=scenarios, y=margins, mode='lines+markers', name="Gross Margin"), secondary_y=True)
    fig3.update_layout(title="📈 Scenario Analysis: Output vs Margin", height=460, legend_orientation="h")
    charts['scenario'] = fig3

    # Sales vs production (if sales)
    if gap["has_sales"]:
        srt = sales_df.sort_values("Month")
        fig4 = go.Figure()
        fig4.add_trace(go.Scatter(x=srt["Month"], y=srt["Sales Quantity (cups)"], name="Sales", mode="lines+markers"))
        fig4.add_hline(y=agg["Current/month"], line_dash="dash", annotation_text="Current Production")
        fig4.update_layout(title="📊 Sales vs Production", height=420)
        charts['sales_vs_prod'] = fig4

    return charts

# ------------------ NEW MODULES ------------------
def unit_economics_waterfall(df, price_adj=0.0, paper_adj=0.0, eff_adj=0.0, waste_adj=0.0):
    """
    Compute average per-1k unit economics across machines, allow simple what-ifs.
    price_adj, paper_adj, eff_adj, waste_adj are percentage deltas (e.g., +5 -> +5%).
    """
    # Base per 1k
    price_1k = df["Selling Price per 1,000 cups (SAR)"].mean(skipna=True)
    raw_1k   = df["Raw Material Cost per 1,000 cups (SAR)"].mean(skipna=True)

    # Conversion per 1k: shift-cost/day divided by (actual/day / 1000)
    actual_day_total = df["Actual Production per Day (cups)"].sum()
    shift_cost_day_total = df["Shift Cost/day (SAR)"].sum()
    conv_per_1k = 0.0
    if actual_day_total > 0:
        conv_per_1k = shift_cost_day_total / (actual_day_total / 1000.0)

    waste_rate = df["Wastage / Rejection Rate (%)"].mean(skipna=True) / 100.0

    # Apply what-ifs
    price_1k *= (1 + price_adj/100.0)
    raw_1k   *= (1 + paper_adj/100.0)
    eff_mult = (1 + eff_adj/100.0)
    # efficiency improvement lowers conversion per 1k proportionally (simple model)
    conv_per_1k = conv_per_1k / eff_mult
    waste_rate = max(0.0, waste_rate * (1 + waste_adj/100.0))

    # Wastage cost (assume it mainly hits raw + conversion)
    waste_cost_1k = waste_rate * (raw_1k + conv_per_1k)

    gm_1k = price_1k - raw_1k - conv_per_1k - waste_cost_1k

    steps = [
        ("Price",           price_1k,  "relative"),
        ("Raw Material",   -raw_1k,    "relative"),
        ("Labor + Power",  -conv_per_1k, "relative"),
        ("Wastage Cost",   -waste_cost_1k, "relative"),
        ("Gross Margin",    gm_1k,     "total")
    ]
    fig = go.Figure(go.Waterfall(
        measure=[s[2] for s in steps],
        x=[s[0] for s in steps],
        y=[s[1] for s in steps],
        connector={"line":{"color":"#888"}}
    ))
    fig.update_layout(title="💵 Unit Economics per 1,000 Cups (What-if)", height=420)

    summary = dict(
        price_1k=price_1k, raw_1k=raw_1k, conv_1k=conv_per_1k,
        waste_rate=waste_rate, waste_cost_1k=waste_cost_1k, gm_1k=gm_1k
    )
    return fig, summary

def compute_oee(df):
    # Quality = 1 - wastage
    quality = 1 - (df["Wastage / Rejection Rate (%)"].fillna(0)/100.0)
    # Availability = (hours - downtime)/hours
    avail = (df["Current Working Hours per Day"] - df["Avg Downtime per Day (hrs)"].fillna(0)) / df["Current Working Hours per Day"].replace(0, np.nan)
    # Performance = actual/theoretical
    perf = df["Actual Production per Day (cups)"] / df["Theoretical Production per Day (cups)"].replace(0, np.nan)
    oee = avail * perf * quality
    out = df.copy()
    out["Availability"] = avail
    out["Performance"]  = perf
    out["Quality"]      = quality
    out["OEE"]          = oee
    return out

def breakeven_and_sensitivity(df, fixed_overheads_sar, hours_grid=(np.arange(8,25,1)), eff_grid=(np.arange(0.6,0.96,0.02))):
    # Use portfolio averages for per-1k economics
    price_1k = df["Selling Price per 1,000 cups (SAR)"].mean(skipna=True)
    raw_1k   = df["Raw Material Cost per 1,000 cups (SAR)"].mean(skipna=True)

    actual_day_total = df["Actual Production per Day (cups)"].sum()
    shift_cost_day_total = df["Shift Cost/day (SAR)"].sum()
    conv_per_1k = shift_cost_day_total / (actual_day_total/1000.0) if actual_day_total>0 else 0.0

    waste_rate = df["Wastage / Rejection Rate (%)"].mean(skipna=True)/100.0
    waste_cost_1k = waste_rate * (raw_1k + conv_per_1k)

    var_cost_per_cup = (raw_1k + conv_per_1k + waste_cost_1k) / 1000.0
    price_per_cup = price_1k / 1000.0
    unit_margin = price_per_cup - var_cost_per_cup
    breakeven_monthly_cups = fixed_overheads_sar / max(1e-6, unit_margin)

    # Sensitivity: monthly GM for hours x eff grid
    rated_total_cpm = df["Rated Capacity (cups/min)"].sum()
    days = df["Working Days per Month"].mean(skipna=True)
    Z = np.zeros((len(eff_grid), len(hours_grid)))
    for i, e in enumerate(eff_grid):
        for j, h in enumerate(hours_grid):
            theo_cups_month = rated_total_cpm * 60 * h * days
            actual_cups_month = theo_cups_month * e
            # Scale shift-costs with hours vs current average hours
            cur_hours = df["Current Working Hours per Day"].mean(skipna=True)
            total_shift_cost_month = df["Shift Cost/day (SAR)"].sum() * days * (h / max(1e-6, cur_hours))
            gm_month = actual_cups_month * unit_margin - total_shift_cost_month - fixed_overheads_sar
            Z[i, j] = gm_month
    return breakeven_monthly_cups, hours_grid, eff_grid, Z

def capacity_expansion(df, new_machines:int, capex_per_machine:float, hours:int, eff:float,
                       ramp_months:int=6, added_fixed_opex:float=0.0):
    """
    Simple 12-month expansion model using portfolio averages for new machines.
    Returns DataFrame with month, incr_output, incr_gm, cum_cash, payback_month.
    """
    if new_machines <= 0:
        months = np.arange(1, 13)
        base = pd.DataFrame({"Month": months, "Incremental Output (cups)": 0.0,
                             "Incremental GM (SAR)": 0.0, "Cumulative Cash (SAR)": -capex_per_machine*0})
        return base, None

    avg_cpm = df["Rated Capacity (cups/min)"].mean(skipna=True)
    days = int(df["Working Days per Month"].mean(skipna=True))
    margin_per_1k = df["Margin per 1k (SAR)"].mean(skipna=True)
    cur_hours = df["Current Working Hours per Day"].mean(skipna=True)
    # Avg shift cost/day per machine
    shift_cost_day_pm = df["Shift Cost/day (SAR)"].sum() / max(1, len(df))

    months = np.arange(1, 13)
    outputs = []
    gms = []
    cash = []
    cum = - new_machines * capex_per_machine  # invest upfront
    for m in months:
        # ramp from 40% to 100% over ramp_months
        ramp_factor = min(1.0, 0.4 + 0.6*(m/ramp_months)) if ramp_months>0 else 1.0
        theo_day = avg_cpm * 60 * hours * eff * ramp_factor
        out_month = theo_day * days * new_machines
        # scale shift cost with hours/time vs current
        shift_cost = shift_cost_day_pm * (hours / max(1e-6, cur_hours)) * days * new_machines
        gm = (out_month/1000.0)*margin_per_1k - shift_cost - added_fixed_opex
        cum += gm
        outputs.append(out_month); gms.append(gm); cash.append(cum)

    dfm = pd.DataFrame({
        "Month": months,
        "Incremental Output (cups)": outputs,
        "Incremental GM (SAR)": gms,
        "Cumulative Cash (SAR)": cash
    })
    payback_month = int(dfm.index[dfm["Cumulative Cash (SAR)"]>=0].min()+1) if (dfm["Cumulative Cash (SAR)"]>=0).any() else None
    return dfm, payback_month

# ------------------ MAIN APP ------------------
def main():
    load_custom_css()

    st.markdown("<div class='main-header'><h2>🏭 Factory Analytics Hub</h2><div>Decision-grade insights for investors</div></div>", unsafe_allow_html=True)

    # ==== Sidebar ====
    with st.sidebar:
        st.subheader("📁 Data Source")
        uploaded = st.file_uploader("Upload Excel (template v2)", type=["xlsx"])
        use_default = st.checkbox("Use default path", value=True, help=DEFAULT_PATH)
        excel_path = DEFAULT_PATH if use_default else st.text_input("Path", value=DEFAULT_PATH)

        st.subheader("🔧 Configuration")
        shift_hours = st.number_input("Shift Length (hours)", 4, 12, SHIFT_HOURS_DEFAULT, step=1)
        st.caption("Affects shift-based costs; production comes from actuals or projections.")
        st.markdown("---")
        st.markdown("**Scenario Efficiency Defaults**")
        eff_12 = st.slider("Efficiency @ 12h", 0.60, 0.95, 0.80, 0.01)
        eff_16 = st.slider("Efficiency @ 16h", 0.60, 0.95, 0.78, 0.01)
        eff_20 = st.slider("Efficiency @ 20h", 0.60, 0.95, 0.76, 0.01)
        eff_24 = st.slider("Efficiency @ 24h", 0.60, 0.95, 0.75, 0.01)
        scenarios_eff = {12: eff_12, 16: eff_16, 20: eff_20, 24: eff_24}

        st.markdown("---")
        use_hours_override = st.checkbox("Override operating hours/day for projection", value=False)
        oper_hours = st.slider("Operating Hours per Day (projection)", 8, 24, 12, 1, disabled=not use_hours_override)

    # ==== Load data ====
    try:
        sheets = read_workbook(uploaded if uploaded is not None else excel_path)
        df, sales_df = build_base_model(sheets, shift_hours=shift_hours)
    except Exception as e:
        st.error(f"Could not read Excel. Upload a file or check path.\n\nDetails: {e}")
        st.stop()

    agg = portfolio_agg(df, scenarios_eff)
    gap = sales_gap(df, sales_df)

    # ==== Projection for KPIs ====
    if use_hours_override:
        nearest = min(SCENARIO_EFF_DEFAULT.keys(), key=lambda h: abs(h - oper_hours))
        eff = scenarios_eff.get(nearest, SCENARIO_EFF_DEFAULT[nearest])
        proj_day, proj_month, proj_gm = project_hours(df, oper_hours, eff)
        kpi_day = float(proj_day.sum())
        kpi_month = float(proj_month.sum())
        kpi_idle = float((df["Rated Capacity (cups/min)"]*60*oper_hours - proj_day).clip(lower=0).sum())
        kpi_gm = float(proj_gm.sum())
        hours_label = f"{oper_hours}h (projected)"
    else:
        kpi_day = float(df["Actual Production per Day (cups)"].sum())
        kpi_month = float(agg["Current/month"])
        kpi_idle = float((df["Theoretical Production per Day (cups)"] - df["Actual Production per Day (cups)"]).clip(lower=0).sum())
        kpi_gm = float(agg["Current GM/month (SAR)"])
        hours_label = f"{shift_hours}h (shift length)"

    # ==== Executive KPIs ====
    st.markdown("<div class='section'>📊 Executive Dashboard</div>", unsafe_allow_html=True)
    c1,c2,c3,c4 = st.columns(4)
    with c1: kpi_card("Total Machines", len(df), " units")
    with c2: kpi_card("Rated Capacity", df["Rated Capacity (cups/min)"].sum(), " CPM")
    with c3: kpi_card("Average Utilization", df["Efficiency (Actual)"].mean(skipna=True)*100, "%")
    with c4: kpi_card("Monthly Margin", kpi_gm, " SAR")
    c5,c6,c7,c8 = st.columns(4)
    with c5: kpi_card("Daily Output", kpi_day, " cups")
    with c6: kpi_card("Monthly Output", kpi_month, " cups")
    with c7: kpi_card("Idle Capacity", kpi_idle, " cups/day")
    with c8: kpi_card("Hours Setting", float(oper_hours if use_hours_override else shift_hours), "h")

    # ==== Core charts ====
    st.markdown("<div class='section'>📈 Performance Analytics</div>", unsafe_allow_html=True)
    charts = charts_core(df, agg, sales_df, gap)
    a,b = st.columns(2)
    a.plotly_chart(charts['cap_vs_actual'], use_container_width=True)
    b.plotly_chart(charts['util_hist'], use_container_width=True)
    st.plotly_chart(charts['scenario'], use_container_width=True)
    if gap["has_sales"]:
        st.plotly_chart(charts['sales_vs_prod'], use_container_width=True)

    # ==== Investor Modules ====
    st.markdown("<div class='section'>🧭 Investor Modules</div>", unsafe_allow_html=True)
    tab1, tab2, tab3, tab4 = st.tabs(["Unit Economics", "OEE & Bottlenecks", "Breakeven & Sensitivity", "Capacity Expansion"])

    # ----- Tab 1: Unit Economics -----
    with tab1:
        st.markdown("#### 💵 Unit Economics (per 1,000 cups) — Waterfall & What-if")
        colu = st.columns(4)
        price_adj = colu[0].number_input("Price Δ (%)", -50.0, 50.0, 0.0, 1.0)
        paper_adj = colu[1].number_input("Paper Cost Δ (%)", -50.0, 50.0, 0.0, 1.0)
        eff_adj   = colu[2].number_input("Efficiency Δ (%)", -30.0, 30.0, 0.0, 1.0)
        waste_adj = colu[3].number_input("Wastage Δ (%)", -50.0, 50.0, 0.0, 1.0)
        fig_wf, u = unit_economics_waterfall(df, price_adj, paper_adj, eff_adj, waste_adj)
        st.plotly_chart(fig_wf, use_container_width=True)
        st.caption(f"Price: {u['price_1k']:,.0f} | Raw: {u['raw_1k']:,.0f} | Conv: {u['conv_1k']:,.0f} | Wastage: {u['waste_cost_1k']:,.0f} | **GM/1k: {u['gm_1k']:,.0f} SAR**")

    # ----- Tab 2: OEE & Bottlenecks -----
    with tab2:
        st.markdown("#### 🛠️ OEE (Availability × Performance × Quality)")
        oee_df = compute_oee(df)
        col_o1, col_o2 = st.columns(2)
        col_o1.plotly_chart(px.bar(oee_df.sort_values("OEE", ascending=False).head(15),
                                   x="Machine Name/ID", y="OEE", title="Top Machines by OEE"), use_container_width=True)
        col_o2.plotly_chart(px.bar(oee_df.sort_values("OEE", ascending=True).head(15),
                                   x="Machine Name/ID", y="OEE", title="Lowest Machines by OEE"), use_container_width=True)

        st.markdown("##### Pareto of Downtime Reasons")
        if "Downtime Reasons" in df.columns:
            d = df["Downtime Reasons"].dropna().astype(str).str.split(",")
            reasons = pd.Series([r.strip() for sub in d for r in sub if r.strip()!=""])
            pareto = reasons.value_counts().reset_index()
            pareto.columns = ["Reason","Count"]
            st.plotly_chart(px.bar(pareto.head(12), x="Reason", y="Count", title="Top Downtime Reasons"), use_container_width=True)
        else:
            st.info("No 'Downtime Reasons' column found.")

    # ----- Tab 3: Breakeven & Sensitivity -----
    with tab3:
        st.markdown("#### ⚖️ Breakeven & Profit Sensitivity")
        fixed_ov = st.number_input("Fixed Overheads / month (SAR)", 0.0, 1e9, 0.0, 1000.0)
        be_cups, H, E, Z = breakeven_and_sensitivity(df, fixed_ov)
        st.metric("Breakeven Volume (cups/month)", f"{be_cups:,.0f}")
        heat = px.imshow(Z, x=[int(h) for h in H], y=[f"{e:.0%}" for e in E],
                         color_continuous_scale="RdYlGn", origin="lower",
                         labels=dict(x="Hours / day", y="Efficiency", color="GM / month (SAR)"),
                         title="Monthly Gross Margin Sensitivity (Hours × Efficiency)")
        st.plotly_chart(heat, use_container_width=True)

    # ----- Tab 4: Capacity Expansion & Payback -----
    with tab4:
        st.markdown("#### 🚀 Expansion Model (12-month) & Payback")
        ce1, ce2, ce3 = st.columns(3)
        new_m = ce1.number_input("New Machines", 0, 200, 2, 1)
        capex = ce2.number_input("Capex per Machine (SAR)", 0.0, 1e9, 150000.0, 1000.0)
        ramp  = ce3.number_input("Ramp-up Months (to 100%)", 0, 24, 6, 1)

        ce4, ce5, ce6 = st.columns(3)
        hours_new = ce4.slider("Operating Hours / day (new)", 8, 24, 16, 1)
        eff_new   = ce5.slider("Efficiency (new)", 0.60, 0.95, 0.78, 0.01)
        add_opex  = ce6.number_input("Added Fixed Opex / month (SAR)", 0.0, 1e8, 0.0, 1000.0)

        dfm, payback = capacity_expansion(df, new_m, capex, hours_new, eff_new, ramp, add_opex)
        st.plotly_chart(px.bar(dfm, x="Month", y="Incremental GM (SAR)", title="Incremental Gross Margin (Monthly)"),
                        use_container_width=True)
        st.plotly_chart(px.line(dfm, x="Month", y="Cumulative Cash (SAR)", markers=True,
                                title="Cumulative Cash vs Time (after Capex)"), use_container_width=True)
        if payback is None:
            st.warning("Payback not reached within 12 months.")
        else:
            st.success(f"Estimated **Payback**: Month {payback}")

    # ==== Detailed machine table ====
    st.markdown("<div class='section'>📋 Detailed Machine Analysis</div>", unsafe_allow_html=True)
    show_cols = [
        "Machine Name/ID","Machine Type","Condition","Cup Thickness (mm)",
        "Rated Capacity (cups/min)","Current Working Hours per Day","Working Days per Month",
        "Theoretical Production per Day (cups)","Actual Production per Day (cups)","Efficiency (Actual)",
        "Actual Production per Month (cups)","Raw Material Cost per 1,000 cups (SAR)",
        "Selling Price per 1,000 cups (SAR)","Margin per 1k (SAR)","Shifts/day",
        "Shift Cost/day (SAR)","Gross Margin/day (SAR)","Gross Margin/month (SAR)",
        "Avg Downtime per Day (hrs)","Wastage / Rejection Rate (%)","Downtime Reasons"
    ]
    table_df = df[show_cols].copy().sort_values("Gross Margin/month (SAR)", ascending=False)
    # Ensure numeric for styling
    num_cols = [c for c in table_df.columns if c not in ["Machine Name/ID","Machine Type","Condition","Downtime Reasons"]]
    for c in num_cols: table_df[c] = pd.to_numeric(table_df[c], errors="coerce")
    st.dataframe(
        table_df.style
        .format({
            "Theoretical Production per Day (cups)":"{:,.0f}",
            "Actual Production per Day (cups)":"{:,.0f}",
            "Actual Production per Month (cups)":"{:,.0f}",
            "Raw Material Cost per 1,000 cups (SAR)":"{:,.2f}",
            "Selling Price per 1,000 cups (SAR)":"{:,.2f}",
            "Margin per 1k (SAR)":"{:,.2f}",
            "Shifts/day":"{:,.1f}",
            "Shift Cost/day (SAR)":"{:,.0f}",
            "Gross Margin/day (SAR)":"{:,.0f}",
            "Gross Margin/month (SAR)":"{:,.0f}",
            "Wastage / Rejection Rate (%)":"{:,.1f}%",
            "Efficiency (Actual)":"{:.1%}",
        })
        .background_gradient(subset=["Gross Margin/month (SAR)"], cmap="RdYlGn"),
        use_container_width=True, height=420
    )

if __name__ == "__main__":
    main()
