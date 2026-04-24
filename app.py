"""
THOT Retail — Board-Level Business Dashboard
============================================
Upload THOT_Dashboard_Ready.xlsx (or any compatible Excel file) to instantly
visualise Finance, Operations & KPI pillars across channels and months.

Excel Sheet Layout expected
---------------------------
Sheet 1 — Actuals_FY25-26
  Columns: Month, Channel, Revenue, Net_of_Taxes, COGS, GM, GM%, Ads_Spend,
           CM, CM%, Orders, Clients, Units, ASP, AOV, ROAS, CAC, eM%,
           Discount, Taxes

Sheet 2 — Targets_FY26-27
  Columns: Channel, Revenue_Target, Net_Sales_Target, COGS_Target,
           Total_Margin_Target, Orders_Target, Clients_Target,
           Units_Target, Retention_Margin_Target

Sheet 3 — Orders_Detail  (optional)
  Columns: Order_ID, Date, Channel, Status, GMV, NMV, COGS, GM
"""

import io
import warnings
from typing import Optional

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots

warnings.filterwarnings("ignore")

# ---------------------------------------------------------------------------
# Page config – must be first Streamlit call
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="THOT Retail Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# Global styling
# ---------------------------------------------------------------------------
DARK_CSS = """
<style>
  @import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:wght@300;400;500;600&display=swap');

  html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
    background-color: #0d1117;
    color: #e6edf3;
  }
  h1, h2, h3, h4 {
    font-family: 'Syne', sans-serif;
    color: #58a6ff;
  }
  .metric-card {
    background: linear-gradient(135deg, #161b22 0%, #1c2128 100%);
    border: 1px solid #30363d;
    border-radius: 12px;
    padding: 18px 20px;
    text-align: center;
    transition: box-shadow .2s;
  }
  .metric-card:hover { box-shadow: 0 0 14px #58a6ff55; }
  .metric-label {
    font-size: 0.75rem;
    color: #8b949e;
    letter-spacing: .06em;
    text-transform: uppercase;
    margin-bottom: 4px;
  }
  .metric-value {
    font-family: 'Syne', sans-serif;
    font-size: 1.55rem;
    font-weight: 700;
    color: #f0f6fc;
  }
  .metric-delta {
    font-size: 0.78rem;
    margin-top: 4px;
  }
  .delta-pos { color: #3fb950; }
  .delta-neg { color: #f85149; }
  .section-title {
    font-family: 'Syne', sans-serif;
    font-size: 1.1rem;
    font-weight: 700;
    color: #58a6ff;
    margin-top: 28px;
    margin-bottom: 10px;
    border-left: 3px solid #58a6ff;
    padding-left: 10px;
  }
  .stTabs [data-baseweb="tab-list"] {
    gap: 6px;
    background: #0d1117;
    border-bottom: 1px solid #30363d;
  }
  .stTabs [data-baseweb="tab"] {
    background: #161b22;
    border-radius: 8px 8px 0 0;
    padding: 8px 18px;
    color: #8b949e;
  }
  .stTabs [aria-selected="true"] {
    background: #1f6feb !important;
    color: #f0f6fc !important;
  }
  div[data-testid="stSidebar"] {
    background-color: #161b22;
    border-right: 1px solid #30363d;
  }
  .stButton > button {
    background: #1f6feb;
    color: #f0f6fc;
    border: none;
    border-radius: 8px;
    padding: 8px 20px;
    font-family: 'DM Sans', sans-serif;
    font-weight: 500;
  }
  .stButton > button:hover { background: #388bfd; }
  [data-testid="stDataFrame"] { border-radius: 8px; }
</style>
"""
st.markdown(DARK_CSS, unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Constants & colour palette
# ---------------------------------------------------------------------------
CHANNEL_COLORS = {
    "Display": "#58a6ff",
    "Offline": "#3fb950",
    "Online": "#f78166",
    "Franchise": "#d2a8ff",
    "Affiliates": "#ffa657",
    "The Crallery": "#79c0ff",
    "Thot Retail": "#a5d6ff",
    "Others": "#8b949e",
}

PLOTLY_LAYOUT = dict(
    paper_bgcolor="#0d1117",
    plot_bgcolor="#0d1117",
    font=dict(family="DM Sans", color="#e6edf3"),
    xaxis=dict(gridcolor="#21262d", linecolor="#30363d"),
    yaxis=dict(gridcolor="#21262d", linecolor="#30363d"),
    legend=dict(bgcolor="#161b22", bordercolor="#30363d", borderwidth=1),
    margin=dict(l=40, r=20, t=40, b=40),
)

# ---------------------------------------------------------------------------
# Sample-data generation constants (sourced from THOT FY 25-26 actuals)
# ---------------------------------------------------------------------------

# Gross revenue for FY 25-26 (₹127.55 Crores = ₹1,27,55,00,000)
ANNUAL_REVENUE_BASE: int = 127_55_00_000

# FY 26-27 AoP growth target over FY 25-26 (12.1% YoY)
AOV_GROWTH_FACTOR: float = 1.121

# FY 25-26 full-year actuals used as baseline for growth projections
BASE_ANNUAL_ORDERS: int = 2_257
BASE_ANNUAL_CLIENTS: int = 1_895
BASE_ANNUAL_UNITS: int = 10_310

# FY 26-27 AoP growth multipliers vs FY 25-26 baseline
ORDERS_GROWTH_MULTIPLIER: float = 2.38    # +138% orders YoY
CLIENTS_GROWTH_MULTIPLIER: float = 1.57   # +57% clients YoY
UNITS_GROWTH_MULTIPLIER: float = 1.39     # +39% units YoY

# Channel revenue share in FY 25-26 (Display 11.4%, Affiliates 7.6%, Total 100%)
EXTRA_CHANNEL_WEIGHTS: list[float] = [0.076, 1.0]  # Affiliates, Thot Retail

# GM%/CM% columns stored as fractions (0–1) if their max value is below this threshold
PERCENTAGE_FRACTION_THRESHOLD: float = 1.5

MONTH_ORDER = [
    "Apr", "May", "Jun", "Jul", "Aug", "Sep",
    "Oct", "Nov", "Dec", "Jan", "Feb", "Mar",
]

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def fmt_inr(value: float, unit: str = "auto") -> str:
    """Format a number as ₹ Crores or ₹ Lakhs."""
    if pd.isna(value):
        return "—"
    if unit == "auto":
        if abs(value) >= 1e7:
            return f"₹{value/1e7:.2f} Cr"
        if abs(value) >= 1e5:
            return f"₹{value/1e5:.2f} L"
        return f"₹{value:,.0f}"
    if unit == "cr":
        return f"₹{value/1e7:.2f} Cr"
    if unit == "lakh":
        return f"₹{value/1e5:.2f} L"
    return f"₹{value:,.0f}"


def fmt_pct(value: float) -> str:
    if pd.isna(value):
        return "—"
    return f"{value:.1f}%"


def channel_color(ch: str) -> str:
    return CHANNEL_COLORS.get(ch, "#8b949e")


def apply_plotly_theme(fig: go.Figure) -> go.Figure:
    fig.update_layout(**PLOTLY_LAYOUT)
    return fig


def metric_card(label: str, value: str, delta: Optional[str] = None, delta_positive: Optional[bool] = None) -> str:
    delta_html = ""
    if delta:
        cls = "delta-pos" if delta_positive else ("delta-neg" if delta_positive is False else "")
        delta_html = f'<div class="metric-delta {cls}">{delta}</div>'
    return f"""
<div class="metric-card">
  <div class="metric-label">{label}</div>
  <div class="metric-value">{value}</div>
  {delta_html}
</div>
"""


# ---------------------------------------------------------------------------
# Column auto-detection via fuzzy/keyword matching
# ---------------------------------------------------------------------------
COLUMN_ALIASES: dict[str, list[str]] = {
    "Month": ["month", "period", "mon"],
    "Channel": ["channel", "segment", "ch"],
    "Revenue": ["revenue", "gross revenue", "gross_revenue", "gmv", "sales"],
    "Net_of_Taxes": ["net_of_taxes", "net of taxes", "net sales", "netsales", "nmv"],
    "COGS": ["cogs", "cost of goods", "cost_of_goods"],
    "GM": ["gm", "gross margin", "gross_margin"],
    "GM%": ["gm%", "gm_pct", "gm pct", "gross margin %", "gross_margin_pct"],
    "Ads_Spend": ["ads_spend", "ads spend", "ad spend", "marketing spend", "adspend"],
    "CM": ["cm", "contribution margin", "contribution_margin"],
    "CM%": ["cm%", "cm_pct", "cm pct"],
    "Orders": ["orders", "order count", "no. of orders", "num_orders"],
    "Clients": ["clients", "customers", "client count"],
    "Units": ["units", "unit count", "qty", "quantity"],
    "ASP": ["asp", "avg selling price", "average selling price"],
    "AOV": ["aov", "avg order value", "average order value"],
    "ROAS": ["roas", "return on ad spend"],
    "CAC": ["cac", "customer acquisition cost"],
    "eM%": ["em%", "em_pct", "ebitda margin"],
    "Discount": ["discount", "discounts", "discount_commissions"],
    "Taxes": ["taxes", "tax"],
    # Targets sheet
    "Revenue_Target": ["revenue_target", "revenue target"],
    "Net_Sales_Target": ["net_sales_target", "net sales target"],
    "COGS_Target": ["cogs_target", "cogs target"],
    "Total_Margin_Target": ["total_margin_target", "total margin target"],
    "Orders_Target": ["orders_target", "orders target"],
    "Clients_Target": ["clients_target", "clients target"],
    "Units_Target": ["units_target", "units target"],
    "Retention_Margin_Target": ["retention_margin_target", "retention margin target"],
    # Orders detail
    "Order_ID": ["order_id", "order id", "orderid"],
    "Date": ["date", "order_date"],
    "Status": ["status", "order_status"],
    "GMV": ["gmv", "gross merchandise value"],
    "NMV": ["nmv", "net merchandise value"],
}


def _normalise(s: str) -> str:
    return str(s).strip().lower().replace(" ", "_")


def detect_columns(df: pd.DataFrame) -> dict[str, str]:
    """Return {canonical_name: actual_col} for columns found in df."""
    mapping: dict[str, str] = {}
    col_norm = {_normalise(c): c for c in df.columns}
    for canonical, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            key = _normalise(alias)
            if key in col_norm:
                mapping[canonical] = col_norm[key]
                break
        if canonical not in mapping:
            # direct match
            key = _normalise(canonical)
            if key in col_norm:
                mapping[canonical] = col_norm[key]
    return mapping


def remap(df: pd.DataFrame, col_map: dict[str, str]) -> pd.DataFrame:
    """Rename detected columns to canonical names."""
    inv = {v: k for k, v in col_map.items()}
    return df.rename(columns=inv)


# ---------------------------------------------------------------------------
# Sample data generator
# ---------------------------------------------------------------------------

def generate_sample_data() -> dict[str, pd.DataFrame]:
    rng = np.random.default_rng(42)
    months = [f"{m}-25" if m in ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
              else f"{m}-26" for m in MONTH_ORDER]
    channels = ["Display", "Offline", "Online", "Franchise"]
    channel_weights = [0.114, 0.477, 0.292, 0.041]

    rows = []
    for month in months:
        for ch, wt in zip(channels, channel_weights):
            rev_base = ANNUAL_REVENUE_BASE * wt / 12  # annual ÷ 12 months
            rev = rng.normal(rev_base, rev_base * 0.15)
            rev = max(rev, 0)
            discount = rev * rng.uniform(0.04, 0.10)
            taxes = (rev - discount) * 0.18
            net = rev - discount - taxes
            cogs = net * rng.uniform(0.55, 0.70)
            gm = net - cogs
            gm_pct = (gm / net * 100) if net else 0
            ads = rev * (rng.uniform(0.08, 0.15) if ch in ["Display", "Online"] else 0)
            cm = gm - ads
            cm_pct = (cm / net * 100) if net else 0
            orders = max(int(rng.normal(40 * wt, 8)), 1)
            clients = max(int(orders * rng.uniform(0.7, 1.0)), 1)
            units = int(orders * rng.uniform(3, 6))
            asp = rev / units if units else 0
            aov = rev / orders if orders else 0
            roas = (rev / ads) if ads > 0 else np.nan
            cac = (ads / clients) if clients and ads > 0 else np.nan
            em_pct = cm_pct * rng.uniform(0.85, 0.95)
            rows.append({
                "Month": month,
                "Channel": ch,
                "Revenue": round(rev),
                "Net_of_Taxes": round(net),
                "COGS": round(cogs),
                "GM": round(gm),
                "GM%": round(gm_pct, 2),
                "Ads_Spend": round(ads),
                "CM": round(cm),
                "CM%": round(cm_pct, 2),
                "Orders": orders,
                "Clients": clients,
                "Units": units,
                "ASP": round(asp),
                "AOV": round(aov),
                "ROAS": round(roas, 2) if not pd.isna(roas) else np.nan,
                "CAC": round(cac) if not pd.isna(cac) else np.nan,
                "eM%": round(em_pct, 2),
                "Discount": round(discount),
                "Taxes": round(taxes),
            })

    actuals = pd.DataFrame(rows)

    # Targets
    targets_rows = []
    for ch, wt in zip(channels + ["Affiliates", "Thot Retail"], channel_weights + EXTRA_CHANNEL_WEIGHTS):
        scale = AOV_GROWTH_FACTOR
        targets_rows.append({
            "Channel": ch,
            "Revenue_Target": round(ANNUAL_REVENUE_BASE * wt * scale),
            "Net_Sales_Target": round(ANNUAL_REVENUE_BASE * wt * scale * 0.72),
            "COGS_Target": round(ANNUAL_REVENUE_BASE * wt * scale * 0.72 * 0.65),
            "Total_Margin_Target": round(ANNUAL_REVENUE_BASE * wt * scale * 0.72 * 0.35),
            "Orders_Target": max(int(BASE_ANNUAL_ORDERS * wt * ORDERS_GROWTH_MULTIPLIER), 1),
            "Clients_Target": max(int(BASE_ANNUAL_CLIENTS * wt * CLIENTS_GROWTH_MULTIPLIER), 1),
            "Units_Target": max(int(BASE_ANNUAL_UNITS * wt * UNITS_GROWTH_MULTIPLIER), 1),
            "Retention_Margin_Target": round(ANNUAL_REVENUE_BASE * wt * scale * 0.72 * 0.30),
        })
    targets = pd.DataFrame(targets_rows)

    # Orders detail
    statuses = ["Shipped", "WIP", "Cancelled"]
    status_weights = [0.65, 0.25, 0.10]
    order_rows = []
    for i in range(300):
        ch = rng.choice(channels)
        status = rng.choice(statuses, p=status_weights)
        gmv = rng.uniform(15000, 300000)
        nmv = gmv * rng.uniform(0.75, 0.85)
        cogs_o = nmv * rng.uniform(0.55, 0.70)
        gm_o = nmv - cogs_o
        mon = rng.choice(months)
        order_rows.append({
            "Order_ID": f"ORD-{i+1001}",
            "Date": mon,
            "Channel": ch,
            "Status": status,
            "GMV": round(gmv),
            "NMV": round(nmv),
            "COGS": round(cogs_o),
            "GM": round(gm_o),
        })
    orders_detail = pd.DataFrame(order_rows)

    return {
        "Actuals_FY25-26": actuals,
        "Targets_FY26-27": targets,
        "Orders_Detail": orders_detail,
    }


# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def load_excel(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets: dict[str, pd.DataFrame] = {}
    for name in xl.sheet_names:
        try:
            df = xl.parse(name)
            df.columns = [str(c).strip() for c in df.columns]
            df = df.dropna(how="all")
            sheets[name] = df
        except Exception:
            pass
    return sheets


def find_sheet(sheets: dict[str, pd.DataFrame], candidates: list[str]) -> Optional[pd.DataFrame]:
    for name in sheets:
        for c in candidates:
            if c.lower() in name.lower():
                return sheets[name]
    return None


def prepare_actuals(df: pd.DataFrame) -> pd.DataFrame:
    col_map = detect_columns(df)
    df = remap(df, col_map)
    # Ensure numeric cols
    num_cols = ["Revenue", "Net_of_Taxes", "COGS", "GM", "GM%", "Ads_Spend",
                "CM", "CM%", "Orders", "Clients", "Units", "ASP", "AOV",
                "ROAS", "CAC", "eM%", "Discount", "Taxes"]
    for c in num_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    # Normalise GM% — if stored as 0-1 fraction, convert
    if "GM%" in df.columns and df["GM%"].dropna().max() <= PERCENTAGE_FRACTION_THRESHOLD:
        df["GM%"] = df["GM%"] * 100
    if "CM%" in df.columns and df["CM%"].dropna().max() <= PERCENTAGE_FRACTION_THRESHOLD:
        df["CM%"] = df["CM%"] * 100
    return df


def prepare_targets(df: pd.DataFrame) -> pd.DataFrame:
    col_map = detect_columns(df)
    df = remap(df, col_map)
    num_cols = ["Revenue_Target", "Net_Sales_Target", "COGS_Target",
                "Total_Margin_Target", "Orders_Target", "Clients_Target",
                "Units_Target", "Retention_Margin_Target"]
    for c in num_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def month_sort_key(s: pd.Series) -> pd.Series:
    """Return integer sort key for month strings like 'Apr-25'."""
    fy_order = {m: i for i, m in enumerate(MONTH_ORDER)}
    def _key(val: str) -> int:
        parts = str(val).split("-")
        mon = parts[0][:3].capitalize()
        return fy_order.get(mon, 99)
    return s.map(_key)


# ---------------------------------------------------------------------------
# Chart builders
# ---------------------------------------------------------------------------

def chart_donut(df: pd.DataFrame) -> go.Figure:
    ch_rev = df.groupby("Channel")["Revenue"].sum().reset_index()
    colors = [channel_color(c) for c in ch_rev["Channel"]]
    fig = go.Figure(go.Pie(
        labels=ch_rev["Channel"],
        values=ch_rev["Revenue"],
        hole=0.55,
        marker_colors=colors,
        textinfo="label+percent",
        hovertemplate="%{label}: ₹%{value:,.0f}<extra></extra>",
    ))
    fig.update_layout(title="Revenue by Channel", showlegend=False, **PLOTLY_LAYOUT)
    return fig


def chart_monthly_bar(df: pd.DataFrame) -> go.Figure:
    """Stacked bar of Revenue by Month, coloured by Channel."""
    pivot = df.pivot_table(index="Month", columns="Channel", values="Revenue", aggfunc="sum").fillna(0)
    # Sort months
    pivot["_sort"] = month_sort_key(pd.Series(pivot.index))
    pivot = pivot.sort_values("_sort").drop(columns="_sort")

    fig = go.Figure()
    for ch in pivot.columns:
        fig.add_trace(go.Bar(
            name=ch,
            x=pivot.index,
            y=pivot[ch],
            marker_color=channel_color(ch),
            hovertemplate=f"{ch}: ₹%{{y:,.0f}}<extra></extra>",
        ))
    fig.update_layout(
        barmode="stack",
        title="Monthly Revenue by Channel",
        xaxis_title="Month",
        yaxis_title="Revenue (₹)",
        **PLOTLY_LAYOUT,
    )
    return fig


def chart_channel_bar(df: pd.DataFrame, metric: str = "Revenue", title: str = "") -> go.Figure:
    ch = df.groupby("Channel")[metric].sum().reset_index().sort_values(metric, ascending=False)
    colors = [channel_color(c) for c in ch["Channel"]]
    fig = go.Figure(go.Bar(
        x=ch["Channel"],
        y=ch[metric],
        marker_color=colors,
        hovertemplate=f"%{{x}}: ₹%{{y:,.0f}}<extra></extra>",
    ))
    fig.update_layout(title=title or f"{metric} by Channel", **PLOTLY_LAYOUT)
    return fig


def chart_gm_heatmap(df: pd.DataFrame) -> go.Figure:
    pivot = df.pivot_table(index="Month", columns="Channel", values="GM%", aggfunc="mean")
    pivot["_sort"] = month_sort_key(pd.Series(pivot.index))
    pivot = pivot.sort_values("_sort").drop(columns="_sort")

    fig = go.Figure(go.Heatmap(
        z=pivot.values,
        x=pivot.columns.tolist(),
        y=pivot.index.tolist(),
        colorscale="RdYlGn",
        text=[[f"{v:.1f}%" if not pd.isna(v) else "" for v in row] for row in pivot.values],
        texttemplate="%{text}",
        hovertemplate="Month: %{y}<br>Channel: %{x}<br>GM%%: %{z:.1f}<extra></extra>",
        colorbar=dict(title="GM%"),
    ))
    fig.update_layout(title="GM% Heatmap (Month × Channel)", **PLOTLY_LAYOUT)
    return fig


def chart_waterfall(df: pd.DataFrame) -> go.Figure:
    tot = df[["Revenue", "Discount", "Taxes", "COGS", "Ads_Spend", "CM"]].sum()
    labels = ["Revenue", "−Discount", "−Taxes", "−COGS", "−Ads Spend", "CM"]
    values = [
        tot.get("Revenue", 0),
        -tot.get("Discount", 0),
        -tot.get("Taxes", 0),
        -tot.get("COGS", 0),
        -tot.get("Ads_Spend", 0),
        tot.get("CM", 0),
    ]
    measure = ["absolute", "relative", "relative", "relative", "relative", "total"]
    fig = go.Figure(go.Waterfall(
        name="P&L",
        orientation="v",
        measure=measure,
        x=labels,
        y=values,
        connector=dict(line=dict(color="#30363d")),
        increasing=dict(marker_color="#3fb950"),
        decreasing=dict(marker_color="#f85149"),
        totals=dict(marker_color="#58a6ff"),
        hovertemplate="%{x}: ₹%{y:,.0f}<extra></extra>",
    ))
    fig.update_layout(title="P&L Waterfall", **PLOTLY_LAYOUT)
    return fig


def chart_monthly_line(df: pd.DataFrame, metric: str = "GM%", title: str = "") -> go.Figure:
    mon = df.groupby("Month")[metric].mean().reset_index()
    mon["_sort"] = month_sort_key(mon["Month"])
    mon = mon.sort_values("_sort")
    fig = go.Figure(go.Scatter(
        x=mon["Month"], y=mon[metric],
        mode="lines+markers",
        line=dict(color="#58a6ff", width=2),
        marker=dict(size=7),
        hovertemplate=f"%{{x}}: %{{y:.1f}}<extra></extra>",
    ))
    fig.update_layout(title=title or f"Monthly {metric}", **PLOTLY_LAYOUT)
    return fig


def chart_orders_channel(df: pd.DataFrame) -> go.Figure:
    ch = df.groupby("Channel")["Orders"].sum().reset_index().sort_values("Orders", ascending=False)
    colors = [channel_color(c) for c in ch["Channel"]]
    fig = go.Figure(go.Bar(
        x=ch["Channel"], y=ch["Orders"],
        marker_color=colors,
        hovertemplate="%{x}: %{y:,} orders<extra></extra>",
    ))
    fig.update_layout(title="Orders by Channel", **PLOTLY_LAYOUT)
    return fig


def chart_area_trend(df: pd.DataFrame) -> go.Figure:
    mon = df.groupby("Month")["Orders"].sum().reset_index()
    mon["_sort"] = month_sort_key(mon["Month"])
    mon = mon.sort_values("_sort")
    fig = go.Figure(go.Scatter(
        x=mon["Month"], y=mon["Orders"],
        fill="tozeroy",
        line=dict(color="#58a6ff"),
        fillcolor="rgba(88,166,255,0.15)",
        hovertemplate="%{x}: %{y:,} orders<extra></extra>",
    ))
    fig.update_layout(title="Monthly Order Trend", **PLOTLY_LAYOUT)
    return fig


def chart_unit_economics(df: pd.DataFrame) -> go.Figure:
    metrics = [m for m in ["AOV", "ASP", "ROAS", "CAC"] if m in df.columns]
    ch_vals = df.groupby("Channel")[metrics].mean().reset_index()

    fig = make_subplots(rows=2, cols=2, subplot_titles=metrics)
    positions = [(1, 1), (1, 2), (2, 1), (2, 2)]
    for (r, c), m in zip(positions, metrics):
        colors = [channel_color(ch) for ch in ch_vals["Channel"]]
        fig.add_trace(
            go.Bar(x=ch_vals["Channel"], y=ch_vals[m],
                   marker_color=colors, name=m,
                   hovertemplate=f"{m}: %{{y:,.0f}}<extra></extra>"),
            row=r, col=c
        )
    fig.update_layout(title="Unit Economics by Channel", showlegend=False, **PLOTLY_LAYOUT)
    return fig


def chart_bullet(actuals_df: pd.DataFrame, targets_df: pd.DataFrame) -> go.Figure:
    ch_actual = actuals_df.groupby("Channel")["Revenue"].sum().reset_index()
    ch_actual.columns = ["Channel", "Actual"]

    if "Revenue_Target" not in targets_df.columns:
        return go.Figure()

    merged = ch_actual.merge(
        targets_df[["Channel", "Revenue_Target"]].rename(columns={"Revenue_Target": "Target"}),
        on="Channel", how="inner"
    )
    if merged.empty:
        return go.Figure()

    merged["Attainment"] = merged["Actual"] / merged["Target"] * 100

    fig = go.Figure()
    for _, row in merged.iterrows():
        fig.add_trace(go.Bar(
            name=row["Channel"],
            x=[row["Actual"]],
            y=[row["Channel"]],
            orientation="h",
            marker_color=channel_color(row["Channel"]),
            hovertemplate=f"Actual: ₹%{{x:,.0f}}<extra></extra>",
        ))
        fig.add_trace(go.Scatter(
            x=[row["Target"]],
            y=[row["Channel"]],
            mode="markers",
            marker=dict(symbol="line-ns", size=16, color="#f0f6fc",
                        line=dict(width=3, color="#f0f6fc")),
            name=f"{row['Channel']} Target",
            showlegend=False,
            hovertemplate=f"Target: ₹%{{x:,.0f}}<extra></extra>",
        ))
    fig.update_layout(
        barmode="overlay",
        title="Revenue: Actual vs Target by Channel",
        xaxis_title="Revenue (₹)",
        **PLOTLY_LAYOUT,
    )
    return fig


def chart_orders_status(orders_df: pd.DataFrame) -> go.Figure:
    """Stacked bar of orders GMV by status."""
    grp = orders_df.groupby("Status")["GMV"].sum().reset_index()
    status_colors = {"Shipped": "#3fb950", "WIP": "#ffa657", "Cancelled": "#f85149"}
    colors = [status_colors.get(s, "#8b949e") for s in grp["Status"]]
    fig = go.Figure(go.Bar(
        x=grp["Status"], y=grp["GMV"],
        marker_color=colors,
        hovertemplate="%{x}: ₹%{y:,.0f}<extra></extra>",
    ))
    fig.update_layout(title="GMV by Order Status", **PLOTLY_LAYOUT)
    return fig


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------

def render_sidebar(actuals: Optional[pd.DataFrame]) -> tuple[list[str], list[str]]:
    st.sidebar.title("THOT Retail")
    st.sidebar.markdown("---")

    if actuals is None or actuals.empty:
        return [], []

    channels = sorted(actuals["Channel"].dropna().unique().tolist()) if "Channel" in actuals.columns else []
    months = actuals["Month"].dropna().unique().tolist() if "Month" in actuals.columns else []
    # Sort months
    if months:
        months_sorted = pd.Series(months)
        months_sorted = months_sorted.iloc[month_sort_key(months_sorted).argsort()].tolist()
    else:
        months_sorted = []

    st.sidebar.subheader("🔍 Filters")
    selected_channels = st.sidebar.multiselect(
        "Channel", channels, default=channels, key="filter_channels"
    )
    selected_months = st.sidebar.multiselect(
        "Month", months_sorted, default=months_sorted, key="filter_months"
    )

    st.sidebar.markdown("---")
    st.sidebar.markdown(
        "<small style='color:#8b949e'>THOT Retail Board Dashboard<br>FY 25-26 Actuals · FY 26-27 AoP</small>",
        unsafe_allow_html=True,
    )
    return selected_channels, selected_months


# ---------------------------------------------------------------------------
# Tab renderers
# ---------------------------------------------------------------------------

def tab_overview(df: pd.DataFrame) -> None:
    st.markdown('<div class="section-title">Key Performance Indicators</div>', unsafe_allow_html=True)

    total_rev = df["Revenue"].sum() if "Revenue" in df.columns else 0
    total_gm = df["GM"].sum() if "GM" in df.columns else 0
    total_cm = df["CM"].sum() if "CM" in df.columns else 0
    total_ads = df["Ads_Spend"].sum() if "Ads_Spend" in df.columns else 0
    total_orders = int(df["Orders"].sum()) if "Orders" in df.columns else 0
    total_clients = int(df["Clients"].sum()) if "Clients" in df.columns else 0
    avg_aov = df["AOV"].mean() if "AOV" in df.columns else 0
    avg_roas = df["ROAS"].mean() if "ROAS" in df.columns else 0
    gm_pct = (total_gm / total_rev * 100) if total_rev else 0
    cm_pct = (total_cm / total_rev * 100) if total_rev else 0

    cards = [
        ("Revenue", fmt_inr(total_rev)),
        ("Gross Margin", f"{fmt_inr(total_gm)} ({gm_pct:.1f}%)"),
        ("Contribution Margin", f"{fmt_inr(total_cm)} ({cm_pct:.1f}%)"),
        ("Ads Spend", fmt_inr(total_ads)),
        ("Total Orders", f"{total_orders:,}"),
        ("Total Clients", f"{total_clients:,}"),
        ("Avg AOV", fmt_inr(avg_aov)),
        ("Avg ROAS", f"{avg_roas:.1f}×" if not pd.isna(avg_roas) else "—"),
    ]

    cols = st.columns(4)
    for i, (label, value) in enumerate(cards):
        with cols[i % 4]:
            st.markdown(metric_card(label, value), unsafe_allow_html=True)
            st.write("")

    st.markdown('<div class="section-title">Channel & Monthly Breakdown</div>', unsafe_allow_html=True)
    c1, c2 = st.columns([1, 2])
    with c1:
        st.plotly_chart(chart_donut(df), use_container_width=True)
    with c2:
        st.plotly_chart(chart_monthly_bar(df), use_container_width=True)


def tab_revenue_gm(df: pd.DataFrame) -> None:
    st.markdown('<div class="section-title">Revenue & Gross Margin Analysis</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.plotly_chart(chart_channel_bar(df, "Revenue", "Revenue by Channel"), use_container_width=True)
    with c2:
        st.plotly_chart(chart_channel_bar(df, "GM", "Gross Margin by Channel"), use_container_width=True)

    st.markdown('<div class="section-title">GM% Trend & Heatmap</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.plotly_chart(chart_monthly_line(df, "GM%", "Monthly GM% Trend"), use_container_width=True)
    with c2:
        if "GM%" in df.columns:
            st.plotly_chart(chart_gm_heatmap(df), use_container_width=True)

    st.markdown('<div class="section-title">P&L Waterfall</div>', unsafe_allow_html=True)
    req_cols = ["Revenue", "Discount", "Taxes", "COGS", "Ads_Spend", "CM"]
    if all(c in df.columns for c in req_cols):
        st.plotly_chart(chart_waterfall(df), use_container_width=True)
    else:
        st.info("P&L Waterfall requires: " + ", ".join(req_cols))


def tab_orders_ops(df: pd.DataFrame, orders_df: Optional[pd.DataFrame]) -> None:
    st.markdown('<div class="section-title">Orders & Operations</div>', unsafe_allow_html=True)

    if "Orders" in df.columns:
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(chart_orders_channel(df), use_container_width=True)
        with c2:
            st.plotly_chart(chart_area_trend(df), use_container_width=True)

    st.markdown('<div class="section-title">Unit Economics</div>', unsafe_allow_html=True)
    econ_cols = [c for c in ["AOV", "ASP", "ROAS", "CAC"] if c in df.columns]
    if econ_cols:
        st.plotly_chart(chart_unit_economics(df), use_container_width=True)
    else:
        st.info("Unit economics columns (AOV, ASP, ROAS, CAC) not found in data.")

    if orders_df is not None and not orders_df.empty:
        st.markdown('<div class="section-title">Order Status Breakdown</div>', unsafe_allow_html=True)
        col_map = detect_columns(orders_df)
        od = remap(orders_df, col_map)
        if "Status" in od.columns:
            st.plotly_chart(chart_orders_status(od), use_container_width=True)
            status_tbl = od.groupby("Status").agg(
                Orders=("Order_ID", "count") if "Order_ID" in od.columns else ("Status", "count"),
                GMV=("GMV", "sum"),
                NMV=("NMV", "sum"),
                COGS=("COGS", "sum"),
                GM=("GM", "sum"),
            ).reset_index()
            status_tbl["GM%"] = (status_tbl["GM"] / status_tbl["NMV"] * 100).round(1)
            st.dataframe(status_tbl, use_container_width=True)


def tab_actuals_vs_target(df: pd.DataFrame, targets_df: Optional[pd.DataFrame]) -> None:
    st.markdown('<div class="section-title">Actuals vs Target (FY 26-27)</div>', unsafe_allow_html=True)

    if targets_df is None or targets_df.empty:
        st.warning("No target data found. Upload a file with a 'Targets_FY26-27' sheet.")
        return

    fig_bullet = chart_bullet(df, targets_df)
    if fig_bullet.data:
        st.plotly_chart(fig_bullet, use_container_width=True)

    # Attainment table
    ch_actual = df.groupby("Channel").agg(
        Actual_Revenue=("Revenue", "sum"),
        Actual_Orders=("Orders", "sum") if "Orders" in df.columns else ("Revenue", "count"),
        Actual_Clients=("Clients", "sum") if "Clients" in df.columns else ("Revenue", "count"),
    ).reset_index()

    tgt_cols = [c for c in ["Channel", "Revenue_Target", "Orders_Target", "Clients_Target"] if c in targets_df.columns]
    tgt = targets_df[tgt_cols].copy()

    merged = ch_actual.merge(tgt, on="Channel", how="inner")
    if "Revenue_Target" in merged.columns:
        merged["Rev Attainment %"] = (merged["Actual_Revenue"] / merged["Revenue_Target"] * 100).round(1)
    if "Orders_Target" in merged.columns and "Actual_Orders" in merged.columns:
        merged["Orders Attainment %"] = (merged["Actual_Orders"] / merged["Orders_Target"] * 100).round(1)
    if "Clients_Target" in merged.columns and "Actual_Clients" in merged.columns:
        merged["Clients Attainment %"] = (merged["Actual_Clients"] / merged["Clients_Target"] * 100).round(1)

    # Format revenue columns
    for col in ["Actual_Revenue", "Revenue_Target"]:
        if col in merged.columns:
            merged[col] = merged[col].apply(fmt_inr)

    st.markdown('<div class="section-title">Attainment Summary</div>', unsafe_allow_html=True)
    st.dataframe(merged, use_container_width=True)

    # Monthly plan vs actual if monthly plan available
    st.markdown('<div class="section-title">GM% — Actuals vs Baseline Trend</div>', unsafe_allow_html=True)
    if "GM%" in df.columns:
        st.plotly_chart(chart_monthly_line(df, "GM%", "Monthly GM% (Actuals)"), use_container_width=True)


def tab_raw_data(sheets: dict[str, pd.DataFrame]) -> None:
    st.markdown('<div class="section-title">Raw Data Explorer</div>', unsafe_allow_html=True)
    if not sheets:
        st.info("No data loaded yet.")
        return
    selected = st.selectbox("Select Sheet", list(sheets.keys()))
    df = sheets[selected]
    st.write(f"**{len(df):,} rows × {len(df.columns)} columns**")
    st.dataframe(df, use_container_width=True, height=500)

    # Download as CSV
    csv = df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label=f"Download {selected} as CSV",
        data=csv,
        file_name=f"{selected}.csv",
        mime="text/csv",
    )


# ---------------------------------------------------------------------------
# Main app
# ---------------------------------------------------------------------------

def main() -> None:
    st.title("📊 THOT Retail — Board Dashboard")
    st.markdown(
        "<small style='color:#8b949e'>FY 25-26 Actuals · FY 26-27 AoP · Channel-wise · Month-wise</small>",
        unsafe_allow_html=True,
    )

    # ── State initialisation
    if "sheets" not in st.session_state:
        st.session_state.sheets = {}
    if "sample_loaded" not in st.session_state:
        st.session_state.sample_loaded = False

    # ── Data source controls
    col_upload, col_sample = st.columns([3, 1])
    with col_upload:
        uploaded = st.file_uploader(
            "Upload Excel file (THOT_Dashboard_Ready.xlsx or compatible)",
            type=["xlsx", "xls"],
            key="file_uploader",
        )
    with col_sample:
        st.write("")
        st.write("")
        if st.button("🎲 Load Sample Data"):
            st.session_state.sheets = generate_sample_data()
            st.session_state.sample_loaded = True

    if uploaded is not None:
        with st.spinner("Reading Excel..."):
            st.session_state.sheets = load_excel(uploaded.read())
        st.session_state.sample_loaded = False

    sheets = st.session_state.sheets

    if not sheets:
        st.info("👆 Upload your Excel file above, or click **Load Sample Data** to explore with demo data.")
        return

    # ── Find relevant sheets
    actuals_raw = find_sheet(sheets, ["actuals", "fy25", "25-26", "25_26"])
    targets_raw = find_sheet(sheets, ["target", "fy26", "26-27", "26_27"])
    orders_raw = find_sheet(sheets, ["order", "orders_detail"])

    if actuals_raw is None:
        # Fallback: use first sheet
        actuals_raw = next(iter(sheets.values()))

    actuals = prepare_actuals(actuals_raw.copy())
    targets = prepare_targets(targets_raw.copy()) if targets_raw is not None else None

    # ── Sidebar filters
    selected_channels, selected_months = render_sidebar(actuals)

    # ── Apply filters
    df = actuals.copy()
    if selected_channels and "Channel" in df.columns:
        df = df[df["Channel"].isin(selected_channels)]
    if selected_months and "Month" in df.columns:
        df = df[df["Month"].isin(selected_months)]

    if df.empty:
        st.warning("No data matches the current filters.")
        return

    # ── Tabs
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "🏠 Overview",
        "💰 Revenue & GM",
        "📦 Orders & Ops",
        "🎯 Actuals vs Target",
        "📋 Raw Data",
    ])

    with tab1:
        tab_overview(df)
    with tab2:
        tab_revenue_gm(df)
    with tab3:
        tab_orders_ops(df, orders_raw)
    with tab4:
        tab_actuals_vs_target(df, targets)
    with tab5:
        tab_raw_data(sheets)


if __name__ == "__main__":
    main()
