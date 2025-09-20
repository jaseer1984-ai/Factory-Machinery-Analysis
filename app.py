import os
import io
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.figure_factory as ff

# ------------------ ENHANCED CONFIG ------------------
DEFAULT_PATH = "C:/Users/User/OneDrive/Documents/Factory_Project.xlsx"
SHIFT_HOURS_DEFAULT = 8
SCENARIO_EFF_DEFAULT = {12: 0.80, 16: 0.78, 20: 0.76, 24: 0.75}

# Modern color palette
COLORS = {
    'primary': '#1f77b4',
    'secondary': '#ff7f0e', 
    'success': '#2ca02c',
    'danger': '#d62728',
    'warning': '#ff7f0e',
    'info': '#17a2b8',
    'light': '#f8f9fa',
    'dark': '#343a40',
    'gradient_start': '#667eea',
    'gradient_end': '#764ba2'
}

st.set_page_config(
    page_title="🏭 Factory Analytics Hub",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ------------------ CUSTOM CSS ------------------
def load_custom_css():
    st.markdown("""
    <style>
    /* Import Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    /* Global Styles */
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        font-family: 'Inter', sans-serif;
    }
    
    /* Header Styling */
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 15px;
        margin-bottom: 2rem;
        box-shadow: 0 10px 25px rgba(0,0,0,0.1);
    }
    
    .main-title {
        color: white;
        font-size: 2.5rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
        text-align: center;
    }
    
    .main-subtitle {
        color: rgba(255,255,255,0.9);
        font-size: 1.2rem;
        font-weight: 300;
        text-align: center;
        margin-bottom: 0;
    }
    
    /* KPI Cards */
    .kpi-card {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.08);
        border-left: 4px solid #667eea;
        margin-bottom: 1rem;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    
    .kpi-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 25px rgba(0,0,0,0.15);
    }
    
    .kpi-value {
        font-size: 2.2rem;
        font-weight: 700;
        color: #2c3e50;
        margin-bottom: 0.2rem;
    }
    
    .kpi-label {
        font-size: 0.9rem;
        color: #7f8c8d;
        font-weight: 500;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    .kpi-delta {
        font-size: 0.85rem;
        font-weight: 600;
        margin-top: 0.3rem;
    }
    
    /* Section Headers */
    .section-header {
        background: linear-gradient(90deg, #f8f9fa 0%, #e9ecef 100%);
        padding: 1rem 1.5rem;
        border-radius: 10px;
        border-left: 4px solid #667eea;
        margin: 1.5rem 0 1rem 0;
    }
    
    .section-title {
        font-size: 1.4rem;
        font-weight: 600;
        color: #2c3e50;
        margin: 0;
    }
    
    /* Chart Containers */
    .chart-container {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.08);
        margin-bottom: 1.5rem;
    }
    
    /* Sidebar Styling */
    .sidebar .sidebar-content {
        background: linear-gradient(180deg, #f8f9fa 0%, #e9ecef 100%);
    }
    
    /* Status Badges */
    .status-badge {
        padding: 0.3rem 0.8rem;
        border-radius: 20px;
        font-size: 0.8rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    .status-excellent { background: #d4edda; color: #155724; }
    .status-good { background: #cce5ff; color: #004085; }
    .status-warning { background: #fff3cd; color: #856404; }
    .status-danger { background: #f8d7da; color: #721c24; }
    
    /* Insights Box */
    .insights-container {
        background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
        border-radius: 15px;
        padding: 2rem;
        margin: 2rem 0;
        border: 1px solid #dee2e6;
    }
    
    .insight-item {
        background: white;
        padding: 1rem;
        border-radius: 8px;
        margin-bottom: 0.8rem;
        border-left: 3px solid #667eea;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    
    /* Metrics Grid */
    .metrics-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 1rem;
        margin: 1rem 0;
    }
    
    /* Animation */
    @keyframes fadeInUp {
        from {
            opacity: 0;
            transform: translateY(20px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }
    
    .animate-fade-in {
        animation: fadeInUp 0.6s ease-out;
    }
    
    /* Data Table Styling */
    .dataframe {
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 4px 15px rgba(0,0,0,0.08);
    }
    
    /* Progress Bars */
    .progress-container {
        background: #e9ecef;
        border-radius: 10px;
        height: 8px;
        margin: 0.5rem 0;
    }
    
    .progress-bar {
        height: 100%;
        border-radius: 10px;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        transition: width 0.3s ease;
    }
    </style>
    """, unsafe_allow_html=True)

# ------------------ HELPER FUNCTIONS ------------------
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

def create_modern_kpi_card(title, value, delta=None, delta_color="normal", prefix="", suffix=""):
    """Create a modern KPI card with optional delta"""
    delta_html = ""
    if delta is not None:
        delta_color_map = {
            "normal": "#6c757d",
            "inverse": "#dc3545" if delta >= 0 else "#28a745",
            "positive": "#28a745" if delta >= 0 else "#dc3545"
        }
        color = delta_color_map.get(delta_color, "#6c757d")
        arrow = "↗" if delta >= 0 else "↘"
        delta_html = f'<div class="kpi-delta" style="color: {color};">{arrow} {delta:+.1f}%</div>'
    
    return f"""
    <div class="kpi-card animate-fade-in">
        <div class="kpi-value">{prefix}{value:,.0f}{suffix}</div>
        <div class="kpi-label">{title}</div>
        {delta_html}
    </div>
    """

def create_status_badge(value, thresholds):
    """Create status badge based on value and thresholds"""
    if value >= thresholds.get('excellent', 90):
        return f'<span class="status-badge status-excellent">Excellent</span>'
    elif value >= thresholds.get('good', 70):
        return f'<span class="status-badge status-good">Good</span>'
    elif value >= thresholds.get('warning', 50):
        return f'<span class="status-badge status-warning">Fair</span>'
    else:
        return f'<span class="status-badge status-danger">Poor</span>'

def create_gauge_chart(value, title, max_value=100, color_ranges=None):
    """Create a modern gauge chart"""
    if color_ranges is None:
        color_ranges = [
            {"range": [0, 50], "color": "#ff4757"},
            {"range": [50, 80], "color": "#ffa726"},
            {"range": [80, 100], "color": "#4caf50"}
        ]
    
    fig = go.Figure(go.Indicator(
        mode = "gauge+number+delta",
        value = value,
        domain = {'x': [0, 1], 'y': [0, 1]},
        title = {'text': title, 'font': {'size': 18, 'color': '#2c3e50'}},
        gauge = {
            'axis': {'range': [None, max_value], 'tickwidth': 1, 'tickcolor': "#bdc3c7"},
            'bar': {'color': "#667eea", 'thickness': 0.3},
            'bgcolor': "white",
            'borderwidth': 2,
            'bordercolor': "#ecf0f1",
            'steps': [{'range': [0, max_value * 0.5], 'color': '#f8f9fa'},
                      {'range': [max_value * 0.5, max_value * 0.8], 'color': '#e9ecef'},
                      {'range': [max_value * 0.8, max_value], 'color': '#dee2e6'}],
            'threshold': {
                'line': {'color': "#e74c3c", 'width': 4},
                'thickness': 0.75,
                'value': max_value * 0.9
            }
        }
    ))
    
    fig.update_layout(
        height=300,
        font={'color': "#2c3e50", 'family': "Inter"},
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=20, r=20, t=60, b=20)
    )
    
    return fig

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

def create_enhanced_charts(df, agg, sales_df, gap_info):
    """Create modern, visually appealing charts"""
    charts = {}
    
    # 1. Capacity vs Actual (Modern Bar Chart)
    fig1 = go.Figure()
    
    fig1.add_trace(go.Bar(
        name='Theoretical Capacity',
        x=[f"M{i+1}" for i in range(len(df))],
        y=df["Theoretical Production per Day (cups)"],
        marker_color='rgba(102, 126, 234, 0.6)',
        marker_line=dict(color='rgba(102, 126, 234, 1)', width=2),
        hovertemplate='<b>%{x}</b><br>Theoretical: %{y:,.0f} cups<extra></extra>'
    ))
    
    fig1.add_trace(go.Bar(
        name='Actual Production',
        x=[f"M{i+1}" for i in range(len(df))],
        y=df["Actual Production per Day (cups)"],
        marker_color='rgba(118, 75, 162, 0.8)',
        marker_line=dict(color='rgba(118, 75, 162, 1)', width=2),
        hovertemplate='<b>%{x}</b><br>Actual: %{y:,.0f} cups<extra></extra>'
    ))
    
    fig1.update_layout(
        title={
            'text': '🏭 Production Capacity vs Actual Output',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 20, 'color': '#2c3e50', 'family': 'Inter'}
        },
        xaxis_title="Machines",
        yaxis_title="Cups per Day",
        barmode='group',
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font={'family': 'Inter', 'color': '#2c3e50'},
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ),
        height=500,
        margin=dict(t=80)
    )
    
    fig1.update_xaxes(showgrid=True, gridcolor='rgba(0,0,0,0.1)')
    fig1.update_yaxes(showgrid=True, gridcolor='rgba(0,0,0,0.1)')
    
    charts['capacity_vs_actual'] = fig1
    
    # 2. Enhanced Utilization Distribution
    fig2 = px.histogram(
        df, 
        x="Efficiency (Actual)",
        nbins=15,
        title="⚡ Machine Utilization Distribution",
        labels={'Efficiency (Actual)': 'Utilization Rate', 'count': 'Number of Machines'},
        color_discrete_sequence=['#667eea']
    )
    
    fig2.add_vline(
        x=df["Efficiency (Actual)"].mean(), 
        line_dash="dash", 
        line_color="#e74c3c",
        line_width=3,
        annotation_text=f"Average: {df['Efficiency (Actual)'].mean():.1%}",
        annotation_position="top"
    )
    
    fig2.update_layout(
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font={'family': 'Inter', 'color': '#2c3e50'},
        title_x=0.5,
        height=400
    )
    
    charts['utilization_hist'] = fig2
    
    # 3. Scenario Comparison (Enhanced with Dual Axis)
    fig3 = make_subplots(specs=[[{"secondary_y": True}]])
    
    scenarios = ["Current", "12h", "16h", "20h", "24h"]
    outputs = [agg["Current/month"], agg["12h/month"], agg["16h/month"], agg["20h/month"], agg["24h/month"]]
    margins = [agg["Current GM/month (SAR)"], agg["12h GM/month (SAR)"], agg["16h GM/month (SAR)"], 
               agg["20h GM/month (SAR)"], agg["24h GM/month (SAR)"]]
    
    # Add bars for output
    fig3.add_trace(
        go.Bar(
            x=scenarios,
            y=outputs,
            name="Monthly Output",
            marker_color='rgba(102, 126, 234, 0.7)',
            yaxis='y',
            hovertemplate='<b>%{x}</b><br>Output: %{y:,.0f} cups<extra></extra>'
        ),
        secondary_y=False
    )
    
    # Add line for margins
    fig3.add_trace(
        go.Scatter(
            x=scenarios,
            y=margins,
            mode='lines+markers',
            name="Gross Margin",
            line=dict(color='#e74c3c', width=4),
            marker=dict(size=12, color='#e74c3c', line=dict(width=2, color='white')),
            yaxis='y2',
            hovertemplate='<b>%{x}</b><br>Margin: %{y:,.0f} SAR<extra></extra>'
        ),
        secondary_y=True
    )
    
    fig3.update_layout(
        title={
            'text': '📈 Scenario Analysis: Output vs Profitability',
            'x': 0.5,
            'font': {'size': 20, 'color': '#2c3e50', 'family': 'Inter'}
        },
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font={'family': 'Inter', 'color': '#2c3e50'},
        height=500,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
    )
    
    fig3.update_xaxes(title_text="Operating Scenarios", showgrid=True, gridcolor='rgba(0,0,0,0.1)')
    fig3.update_yaxes(title_text="Monthly Output (Cups)", secondary_y=False, showgrid=True, gridcolor='rgba(0,0,0,0.1)')
    fig3.update_yaxes(title_text="Gross Margin (SAR)", secondary_y=True)
    
    charts['scenario_analysis'] = fig3
    
    # 4. Sales vs Production (if available)
    if gap_info["has_sales"]:
        fig4 = go.Figure()
        
        # Sales line
        fig4.add_trace(go.Scatter(
            x=sales_df.sort_values("Month")["Month"],
            y=sales_df.sort_values("Month")["Sales Quantity (cups)"],
            mode='lines+markers',
            name='Monthly Sales',
            line=dict(color='#28a745', width=3),
            marker=dict(size=8, color='#28a745'),
            fill='tonexty',
            fillcolor='rgba(40, 167, 69, 0.1)',
            hovertemplate='<b>%{x|%b %Y}</b><br>Sales: %{y:,.0f} cups<extra></extra>'
        ))
        
        # Production capacity line
        fig4.add_hline(
            y=agg["Current/month"],
            line_dash="dash",
            line_color="#667eea",
            line_width=3,
            annotation_text=f"Current Production: {agg['Current/month']:,.0f} cups",
            annotation_position="top right"
        )
        
        fig4.update_layout(
            title={
                'text': '📊 Sales vs Production Capacity Trend',
                'x': 0.5,
                'font': {'size': 20, 'color': '#2c3e50', 'family': 'Inter'}
            },
            xaxis_title="Month",
            yaxis_title="Cups",
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font={'family': 'Inter', 'color': '#2c3e50'},
            height=400,
            showlegend=True
        )
        
        charts['sales_vs_production'] = fig4
    
    return charts

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

# ------------------ MAIN APP ------------------
def main():
    # Load custom CSS
    load_custom_css()
    
    # Modern Header
    st.markdown("""
    <div class="main-header">
        <h1 class="main-title">🏭 Factory Analytics Hub</h1>
        <p class="main-subtitle">Advanced Production Intelligence & Investment Analytics</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Enhanced Sidebar
    with st.sidebar:
        st.markdown("""
        <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                    padding: 1.5rem; border-radius: 10px; margin-bottom: 1.5rem;">
            <h3 style="color: white; margin: 0; text-align: center;">⚙️ Control Panel</h3>
        </div>
        """, unsafe_allow_html=True)
        
        st.subheader("📁 Data Source")
        uploaded = st.file_uploader("Upload Factory Data", type=["xlsx"], help="Upload Excel file with factory operations data")
        use_default = st.checkbox("Use Default Path", value=True, help=DEFAULT_PATH)
        if not use_default:
            excel_path = st.text_input("Excel File Path", value=DEFAULT_PATH)
        else:
            excel_path = DEFAULT_PATH
        
        st.markdown("---")
        
        st.subheader("🔧 Configuration")
        shift_hours = st.number_input("Shift Length (hours)", 4, 12, SHIFT_HOURS_DEFAULT, step=1)
        
        st.markdown("**Scenario Efficiency Settings**")
        eff_12 = st.slider("12h Efficiency", 0.60, 0.95, SCENARIO_EFF_DEFAULT[12], 0.01, format="%.0f%%")
        eff_16 = st.slider("16h Efficiency", 0.60, 0.95, SCENARIO_EFF_DEFAULT[16], 0.01, format="%.0f%%")
        eff_20 = st.slider("20h Efficiency", 0.60, 0.95, SCENARIO_EFF_DEFAULT[20], 0.01, format="%.0f%%")
        eff_24 = st.slider("24h Efficiency", 0.60, 0.95, SCENARIO_EFF_DEFAULT[24], 0.01, format="%.0f%%")
        scenarios_eff = {12: eff_12, 16: eff_16, 20: eff_20, 24: eff_24}
        
        st.markdown("---")
        st.info("💡 **Tip**: Upload Sales Report sheet to unlock advanced sales vs production analysis")
    
    # Load and process data
    try:
        if uploaded is not None:
            sheets = read_workbook(uploaded)
        else:
            sheets = read_workbook(excel_path)
        
        df, sales_df = build_base_model(sheets, shift_hours=shift_hours)
        agg = portfolio_agg(df, scenarios_eff)
        gap = sales_gap(df, sales_df)
        
        # Success message
        st.success(f"✅ Successfully loaded data for {len(df)} machines")
        
    except Exception as e:
        st.error(f"❌ Error loading data: {e}")
        st.stop()
    
    # Executive KPI Dashboard
    st.markdown('<div class="section-header"><h2 class="section-title">📊 Executive Dashboard</h2></div>', unsafe_allow_html=True)
    
    # Enhanced KPI Cards
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(
            create_modern_kpi_card(
                "Total Machines", 
                len(df),
                suffix=" units"
            ), 
            unsafe_allow_html=True
        )
    
    with col2:
        rated_capacity = df["Rated Capacity (cups/min)"].sum()
        st.markdown(
            create_modern_kpi_card(
                "Rated Capacity", 
                rated_capacity,
                suffix=" CPM"
            ), 
            unsafe_allow_html=True
        )
    
    with col3:
        avg_utilization = df["Efficiency (Actual)"].mean(skipna=True)
        utilization_status = "excellent" if avg_utilization > 0.9 else "good" if avg_utilization > 0.7 else "warning"
        st.markdown(
            create_modern_kpi_card(
                "Average Utilization", 
                avg_utilization * 100,
                suffix="%"
            ), 
            unsafe_allow_html=True
        )
    
    with col4:
        monthly_margin = agg['Current GM/month (SAR)']
        st.markdown(
            create_modern_kpi_card(
                "Monthly Margin", 
                monthly_margin,
                prefix="",
                suffix=" SAR"
            ), 
            unsafe_allow_html=True
        )
    
    # Second row of KPIs
    col5, col6, col7, col8 = st.columns(4)
    
    with col5:
        daily_output = df["Actual Production per Day (cups)"].sum()
        st.markdown(
            create_modern_kpi_card(
                "Daily Output", 
                daily_output,
                suffix=" cups"
            ), 
            unsafe_allow_html=True
        )
    
    with col6:
        monthly_output = agg['Current/month']
        st.markdown(
            create_modern_kpi_card(
                "Monthly Output", 
                monthly_output,
                suffix=" cups"
            ), 
            unsafe_allow_html=True
        )
    
    with col7:
        idle_capacity = (df["Theoretical Production per Day (cups)"] - df["Actual Production per Day (cups)"]).clip(lower=0).sum()
        st.markdown(
            create_modern_kpi_card(
                "Idle Capacity", 
                idle_capacity,
                suffix=" cups/day"
            ), 
            unsafe_allow_html=True
        )
    
    with col8:
        st.markdown(
            create_modern_kpi_card(
                "Shift Hours", 
                shift_hours,
                suffix="h"
            ), 
            unsafe_allow_html=True
        )
    
    # Advanced Analytics Section
    st.markdown('<div class="section-header"><h2 class="section-title">📈 Performance Analytics</h2></div>', unsafe_allow_html=True)
    
    # Create enhanced charts
    charts = create_enhanced_charts(df, agg, sales_df, gap)
    
    # Display charts in modern containers
    col_left, col_right = st.columns(2)
    
    with col_left:
        with st.container():
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.plotly_chart(charts['capacity_vs_actual'], use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
    
    with col_right:
        with st.container():
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.plotly_chart(charts['utilization_hist'], use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
    
    # Full-width scenario analysis
    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
    st.plotly_chart(charts['scenario_analysis'], use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Sales vs Production (if available)
    if gap["has_sales"]:
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        st.plotly_chart(charts['sales_vs_production'], use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Sales Gap Analysis
        col_gap1, col_gap2, col_gap3 = st.columns(3)
        
        with col_gap1:
            st.markdown(
                create_modern_kpi_card(
                    "Monthly Sales", 
                    gap["avg_monthly_sales"],
                    suffix=" cups"
                ), 
                unsafe_allow_html=True
            )
        
        with col_gap2:
            st.markdown(
                create_modern_kpi_card(
                    "Production Capacity", 
                    gap["current_monthly_production"],
                    suffix=" cups"
                ), 
                unsafe_allow_html=True
            )
        
        with col_gap3:
            gap_value = abs(gap["gap_cups"])
            gap_label = "Overproduction" if gap["gap_cups"] < 0 else "Shortfall"
            st.markdown(
                create_modern_kpi_card(
                    gap_label, 
                    gap_value,
                    suffix=" cups"
                ), 
                unsafe_allow_html=True
            )
        
        # Gap Analysis Alert
        if gap["gap_cups"] < 0:  # Overproduction
            gap_percentage = abs(gap["gap_cups"]) / gap["avg_monthly_sales"] * 100
            st.error(f"⚠️ **CRITICAL OVERPRODUCTION**: Factory is overproducing by {gap_percentage:,.0f}%. Focus on demand generation and market expansion rather than capacity increases.")
        else:  # Shortfall
            st.warning(f"📈 **PRODUCTION SHORTFALL**: Increase capacity by {gap['gap_cups']:,.0f} cups/month to meet demand.")
    
    # Machine Performance Insights
    st.markdown('<div class="section-header"><h2 class="section-title">🏭 Machine Performance Insights</h2></div>', unsafe_allow_html=True)
    
    # Performance Gauges
    col_gauge1, col_gauge2, col_gauge3 = st.columns(3)
    
    with col_gauge1:
        avg_efficiency = df["Efficiency (Actual)"].mean() * 100
        efficiency_gauge = create_gauge_chart(avg_efficiency, "Overall Efficiency", 100)
        st.plotly_chart(efficiency_gauge, use_container_width=True)
    
    with col_gauge2:
        avg_wastage = df["Wastage / Rejection Rate (%)"].mean()
        wastage_gauge = create_gauge_chart(100 - avg_wastage, "Quality Rate", 100)
        st.plotly_chart(wastage_gauge, use_container_width=True)
    
    with col_gauge3:
        avg_uptime = ((df["Current Working Hours per Day"].mean() - df["Avg Downtime per Day (hrs)"].mean()) / 
                     df["Current Working Hours per Day"].mean() * 100)
        uptime_gauge = create_gauge_chart(avg_uptime, "Uptime Rate", 100)
        st.plotly_chart(uptime_gauge, use_container_width=True)
    
    # Top Performers and Issues
    col_perf1, col_perf2 = st.columns(2)
    
    with col_perf1:
        st.markdown("### 🏆 Top Performers")
        top_performers = df.nlargest(5, "Gross Margin/month (SAR)")[["Machine Name/ID", "Gross Margin/month (SAR)", "Efficiency (Actual)"]]
        for _, row in top_performers.iterrows():
            st.markdown(f"""
            <div class="insight-item">
                <strong>Machine {row['Machine Name/ID']}</strong><br>
                💰 {row['Gross Margin/month (SAR)']:,.0f} SAR/month<br>
                ⚡ {row['Efficiency (Actual)']:.1%} efficiency
            </div>
            """, unsafe_allow_html=True)
    
    with col_perf2:
        st.markdown("### ⚠️ Attention Required")
        # Highest wastage
        high_wastage = df.nlargest(3, "Wastage / Rejection Rate (%)")[["Machine Name/ID", "Wastage / Rejection Rate (%)", "Avg Downtime per Day (hrs)"]]
        for _, row in high_wastage.iterrows():
            st.markdown(f"""
            <div class="insight-item">
                <strong>Machine {row['Machine Name/ID']}</strong><br>
                🔥 {row['Wastage / Rejection Rate (%)']:.1f}% wastage<br>
                ⏱️ {row['Avg Downtime per Day (hrs)']:.1f}h downtime
            </div>
            """, unsafe_allow_html=True)
    
    # Scenario Analysis Deep Dive
    st.markdown('<div class="section-header"><h2 class="section-title">🎯 Scenario Planning</h2></div>', unsafe_allow_html=True)
    
    # Scenario comparison table
    scenario_data = []
    current_margin = agg["Current GM/month (SAR)"]
    current_output = agg["Current/month"]
    
    for hours in [8, 12, 16, 20, 24]:
        if hours == 8:
            output = current_output
            margin = current_margin
            output_increase = 0
            margin_increase = 0
        else:
            output = agg[f"{hours}h/month"]
            margin = agg[f"{hours}h GM/month (SAR)"]
            output_increase = ((output - current_output) / current_output) * 100
            margin_increase = ((margin - current_margin) / current_margin) * 100
        
        scenario_data.append({
            "Scenario": f"{hours}h/day",
            "Monthly Output": f"{output:,.0f}",
            "Output Increase": f"{output_increase:+.1f}%",
            "Monthly Margin": f"{margin:,.0f} SAR",
            "Margin Increase": f"{margin_increase:+.1f}%",
            "ROI Score": f"{margin_increase:.0f}"
        })
    
    scenario_df = pd.DataFrame(scenario_data)
    
    # Enhanced scenario table with styling
    st.markdown("### 📊 Scenario Comparison Matrix")
    st.dataframe(
        scenario_df.style.background_gradient(subset=['ROI Score'], cmap='RdYlGn')
                         .format({'Monthly Output': '{:,}', 'Monthly Margin': '{:,}'}),
        use_container_width=True,
        height=220
    )
    
    # Best scenario recommendation
    best_scenario = max(scenarios_eff.keys(), key=lambda x: agg[f"{x}h GM/month (SAR)"])
    best_margin = agg[f"{best_scenario}h GM/month (SAR)"]
    margin_improvement = ((best_margin - current_margin) / current_margin) * 100
    
    st.success(f"""
    ### 🎯 **Optimal Scenario Recommendation**
    **{best_scenario} hours/day operation** delivers maximum profitability:
    - 💰 **Monthly Margin**: {best_margin:,.0f} SAR (+{margin_improvement:.0f}% increase)
    - 📈 **Monthly Output**: {agg[f'{best_scenario}h/month']:,.0f} cups
    - 🎖️ **Efficiency**: {scenarios_eff[best_scenario]:.0%} assumed efficiency
    """)
    
    # Strategic Insights & Recommendations
    st.markdown('<div class="section-header"><h2 class="section-title">🧠 Strategic Intelligence</h2></div>', unsafe_allow_html=True)
    
    # Generate insights
    insights_content = []
    
    # Production insights
    insights_content.append(f"🏭 **Production Status**: {len(df)} machines operating at {avg_utilization:.1%} average utilization")
    insights_content.append(f"📊 **Current Output**: {monthly_output:,.0f} cups/month generating {monthly_margin:,.0f} SAR margin")
    
    # Scenario insights
    for hours in [12, 16, 20, 24]:
        output_inc = ((agg[f"{hours}h/month"] - current_output) / current_output) * 100
        margin_inc = ((agg[f"{hours}h GM/month (SAR)"] - current_margin) / current_margin) * 100
        insights_content.append(f"⚡ **{hours}h Scenario**: +{output_inc:.0f}% output, +{margin_inc:.0f}% margin ({agg[f'{hours}h GM/month (SAR)']:,.0f} SAR)")
    
    # Sales insights
    if gap["has_sales"]:
        if gap["gap_cups"] < 0:
            overproduction = abs(gap["gap_cups"])
            overprod_percentage = (overproduction / gap["avg_monthly_sales"]) * 100
            insights_content.append(f"🚨 **Market Gap**: Overproducing by {overproduction:,.0f} cups ({overprod_percentage:.0f}%). Prioritize demand generation over capacity expansion.")
        else:
            shortfall = gap["gap_cups"]
            insights_content.append(f"📈 **Growth Opportunity**: {shortfall:,.0f} cups demand exceeds capacity. Consider operational hour increases.")
    
    # Bottleneck insights
    worst_wastage = df.loc[df["Wastage / Rejection Rate (%)"].idxmax()]
    worst_downtime = df.loc[df["Avg Downtime per Day (hrs)"].idxmax()]
    insights_content.append(f"⚠️ **Quality Priority**: Machine {worst_wastage['Machine Name/ID']} has {worst_wastage['Wastage / Rejection Rate (%)']:.1f}% wastage")
    insights_content.append(f"🔧 **Reliability Priority**: Machine {worst_downtime['Machine Name/ID']} has {worst_downtime['Avg Downtime per Day (hrs)']:.1f}h daily downtime")
    
    # Display insights in modern format
    st.markdown('<div class="insights-container">', unsafe_allow_html=True)
    for insight in insights_content:
        st.markdown(f'<div class="insight-item">{insight}</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Action Items
    st.markdown("### 🚀 Priority Action Items")
    
    action_items = []
    
    if gap["has_sales"] and gap["gap_cups"] < 0:
        action_items.extend([
            "🎯 **Market Development**: Launch demand generation campaigns to address overproduction",
            "💰 **Pricing Strategy**: Review pricing to stimulate market demand",
            "🌟 **Product Innovation**: Develop premium offerings for higher margins"
        ])
    else:
        action_items.append(f"⚡ **Capacity Expansion**: Implement {best_scenario}h operations for optimal ROI")
    
    action_items.extend([
        f"🔧 **Quality Improvement**: Target Machine {worst_wastage['Machine Name/ID']} for wastage reduction",
        f"⏱️ **Reliability Enhancement**: Address downtime issues on Machine {worst_downtime['Machine Name/ID']}",
        "📊 **Performance Monitoring**: Implement real-time dashboards for continuous improvement"
    ])
    
    for i, action in enumerate(action_items, 1):
        st.markdown(f"**{i}.** {action}")
    
    # Machine Performance Table
    st.markdown('<div class="section-header"><h2 class="section-title">📋 Detailed Machine Analysis</h2></div>', unsafe_allow_html=True)
    
    # Enhanced machine table
    show_cols = [
        "Machine Name/ID", "Machine Type", "Condition", "Cup Thickness (mm)",
        "Rated Capacity (cups/min)", "Current Working Hours per Day", "Working Days per Month",
        "Theoretical Production per Day (cups)", "Actual Production per Day (cups)", "Efficiency (Actual)",
        "Actual Production per Month (cups)", "Raw Material Cost per 1,000 cups (SAR)", 
        "Selling Price per 1,000 cups (SAR)", "Margin per 1k (SAR)", "Shifts/day",
        "Shift Cost/day (SAR)", "Gross Margin/day (SAR)", "Gross Margin/month (SAR)",
        "Avg Downtime per Day (hrs)", "Wastage / Rejection Rate (%)", "Downtime Reasons"
    ]
    
    table_df = df[show_cols].copy().sort_values("Gross Margin/month (SAR)", ascending=False)
    
    # Format the dataframe for better display
    display_df = table_df.copy()
    display_df["Efficiency (Actual)"] = display_df["Efficiency (Actual)"].apply(lambda x: f"{x:.1%}")
    
    st.dataframe(
        display_df.style.format({
            "Theoretical Production per Day (cups)": "{:,.0f}",
            "Actual Production per Day (cups)": "{:,.0f}",
            "Actual Production per Month (cups)": "{:,.0f}",
            "Raw Material Cost per 1,000 cups (SAR)": "{:.2f}",
            "Selling Price per 1,000 cups (SAR)": "{:.2f}",
            "Margin per 1k (SAR)": "{:.2f}",
            "Shift Cost/day (SAR)": "{:,.0f}",
            "Gross Margin/day (SAR)": "{:,.0f}",
            "Gross Margin/month (SAR)": "{:,.0f}",
            "Wastage / Rejection Rate (%)": "{:.1f}%"
        }).background_gradient(subset=["Gross Margin/month (SAR)"], cmap='RdYlGn'),
        use_container_width=True,
        height=400
    )
    
    # Export Section
    st.markdown('<div class="section-header"><h2 class="section-title">📥 Export & Reports</h2></div>', unsafe_allow_html=True)
    
    # Create comprehensive report
    agg_df = pd.DataFrame([
        {"Scenario": "Current", "Monthly Output (cups)": agg["Current/month"], "GM/month (SAR)": agg["Current GM/month (SAR)"]},
        {"Scenario": "12h", "Monthly Output (cups)": agg["12h/month"], "GM/month (SAR)": agg["12h GM/month (SAR)"]},
        {"Scenario": "16h", "Monthly Output (cups)": agg["16h/month"], "GM/month (SAR)": agg["16h GM/month (SAR)"]},
        {"Scenario": "20h", "Monthly Output (cups)": agg["20h/month"], "GM/month (SAR)": agg["20h GM/month (SAR)"]},
        {"Scenario": "24h", "Monthly Output (cups)": agg["24h/month"], "GM/month (SAR)": agg["24h GM/month (SAR)"]},
    ])
    
    # Combined insights text
    insights_text_combined = "\n".join([
        "EXECUTIVE SUMMARY:",
        f"• {len(df)} machines operating at {avg_utilization:.1%} average utilization",
        f"• Monthly output: {monthly_output:,.0f} cups | Gross margin: {monthly_margin:,.0f} SAR",
        "",
        "SCENARIO ANALYSIS:",
        f"• Optimal scenario: {best_scenario}h/day (+{margin_improvement:.0f}% margin improvement)",
        "",
        "KEY INSIGHTS:",
    ] + [f"• {insight.replace('**', '').replace('🏭 ', '').replace('📊 ', '').replace('⚡ ', '').replace('🚨 ', '').replace('📈 ', '').replace('⚠️ ', '').replace('🔧 ', '')}" for insight in insights_content])
    
    col_export1, col_export2 = st.columns(2)
    
    with col_export1:
        # Excel download
        xls_bytes = to_excel_download(table_df, agg_df, sales_df, insights_text_combined)
        st.download_button(
            label="📊 Download Excel Report",
            data=xls_bytes,
            file_name="Factory_Analytics_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    
    with col_export2:
        # Summary text download
        st.download_button(
            label="📄 Download Text Summary",
            data=insights_text_combined,
            file_name="Factory_Insights_Summary.txt",
            mime="text/plain"
        )
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #7f8c8d; font-style: italic; margin-top: 2rem;">
        🏭 Factory Analytics Hub | Built with Streamlit | Deploy on GitHub for instant access
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
