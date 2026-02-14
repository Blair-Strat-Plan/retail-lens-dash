"""
Retail-Lens — Portfolio Intelligence Suite
Modern SaaS-style financial insights dashboard for a tech hardware retailer.

Raw Excel (data.xlsx) required columns:
- Product
- Type
- Units
- Retail Price
- Wholesale Price

All other fields are derived in Python.
"""

from __future__ import annotations

from pathlib import Path
from datetime import datetime
import pandas as pd
import streamlit as st
from streamlit import html as st_html
import plotly.express as px
import plotly.graph_objects as go


# =========================
# App Config
# =========================
st.set_page_config(
    page_title="Retail-Lens — Portfolio Intelligence Suite",
    layout="wide",
)


# =========================
# Theme (Soft Warm Grey — Premium Light Enterprise)
# =========================
THEME = {
    "bg": "#F3F4F6",        # warm light grey
    "panel": "#FFFFFF",     # white panels
    "panel2": "#F8FAFC",    # off-white panel
    "border": "#E5E7EB",    # soft border
    "text": "#111827",      # near-black
    "muted": "#4B5563",     # grey
    "muted2": "#6B7280",
    "accent": "#334155",    # slate
    "accent2": "#64748B",   # slate light
    "pos": "#0F766E",       # muted teal
    "neg": "#B91C1C",       # controlled red
    "warn": "#B45309",      # muted amber
}


# =========================
# Data Source
# =========================
EXCEL_FILE = Path("data.xlsx")


# =========================
# Global Styling (Ribbon default) + Layout spacing
# =========================
st.markdown(
    f"""
    <style>
      /* --- Background / base --- */
      .stApp {{
        background: {THEME["bg"]};
        color: {THEME["text"]};
      }}

      /* Space content down so Streamlit ribbon doesn't overlap our header bar */
      .block-container {{
        padding-top: 4.5rem !important;
        padding-bottom: 2rem !important;
      }}

      /* Sidebar */
      section[data-testid="stSidebar"] {{
        background: {THEME["panel2"]};
        border-right: 1px solid {THEME["border"]};
      }}

      /* Buttons */
      .stButton button {{
        background: {THEME["panel"]} !important;
        color: {THEME["text"]} !important;
        border: 1px solid {THEME["border"]} !important;
        border-radius: 12px !important;
        transition: transform 140ms ease, border-color 140ms ease, box-shadow 140ms ease;
        box-shadow: 0 1px 2px rgba(0,0,0,0.04);
      }}
      .stButton button:hover {{
        transform: translateY(-1px);
        border-color: rgba(51,65,85,0.55) !important;
        box-shadow: 0 8px 18px rgba(0,0,0,0.07);
      }}

      /* Header bar */
      .rl-header {{
        background: {THEME["panel"]};
        border: 1px solid {THEME["border"]};
        border-radius: 16px;
        padding: 12px 14px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 12px;
        box-shadow: 0 8px 22px rgba(0,0,0,0.06);
      }}
      .rl-brand {{
        display: flex;
        align-items: center;
        gap: 10px;
        min-width: 260px;
      }}
      .rl-brand-mark {{
        width: 26px;
        height: 26px;
        flex: 0 0 26px;
      }}
      /* Sora wordmark */
      .rl-wordmark {{
        font-family: "Sora", ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial !important;
        font-weight: 650;
        font-size: 16px;
        letter-spacing: 0.01em;
        color: {THEME["text"]} !important;
        line-height: 1.0;
      }}
      .rl-subtitle {{
        font-size: 12px;
        color: {THEME["muted"]} !important;
        margin-top: 2px;
      }}
      .rl-center-title {{
        font-size: 13px;
        color: {THEME["muted"]} !important;
        letter-spacing: 0.02em;
      }}
      .rl-right-meta {{
        text-align: right;
        color: {THEME["muted2"]} !important;
        font-size: 12px;
        line-height: 1.2;
        min-width: 260px;
      }}

      /* KPI cards */
      .rl-card {{
        background: {THEME["panel"]};
        border: 1px solid {THEME["border"]};
        border-radius: 14px;
        padding: 12px 12px;
        box-shadow: 0 10px 22px rgba(0,0,0,0.06);
        transition: transform 160ms ease, box-shadow 160ms ease, border-color 160ms ease;
        position: relative;
        overflow: hidden;
      }}
      .rl-card::before {{
        content: "";
        position: absolute;
        top: 0;
        left: 0;
        height: 3px;
        width: 100%;
        background: rgba(51,65,85,0.65);
      }}
      .rl-card:hover {{
        transform: translateY(-1px);
        box-shadow: 0 14px 28px rgba(0,0,0,0.08);
        border-color: rgba(51,65,85,0.45);
      }}
      .rl-kpi-label {{
        color: {THEME["muted"]} !important;
        font-size: 11px;
        letter-spacing: 0.08em;
        text-transform: uppercase;
      }}
      .rl-kpi-value {{
        font-size: 22px;
        font-weight: 700;
        margin-top: 6px;
        color: {THEME["text"]} !important;
      }}
      .rl-kpi-sub {{
        color: {THEME["muted2"]} !important;
        font-size: 12px;
        margin-top: 4px;
      }}

      /* Explainer block (Teen explainer: bigger + italics) */
      .rl-guide {{
        background: {THEME["panel"]};
        border: 1px solid {THEME["border"]};
        border-radius: 16px;
        padding: 14px 16px;
        margin-bottom: 12px;
        max-width: 1100px;
        box-shadow: 0 10px 22px rgba(0,0,0,0.06);
      }}
      .rl-guide p {{
        margin: 0 0 10px 0;
        font-size: 19px;
        font-style: italic;
        line-height: 1.55;
        color: {THEME["muted"]} !important;
      }}
      .rl-guide p:last-child {{
        margin-bottom: 0;
      }}

      /* Subsection heading */
      .rl-section-title {{
        font-size: 14px;
        font-weight: 700;
        margin: 8px 0 0 0;
      }}
      .rl-section-sub {{
        font-size: 12px;
        color: {THEME["muted"]} !important;
        margin-top: 2px;
      }}

      /* Page transition (soft fade/slide) */
      .rl-fade {{
        animation: rlFadeIn 220ms ease-out;
      }}
      @keyframes rlFadeIn {{
        from {{ opacity: 0; transform: translateY(4px); }}
        to   {{ opacity: 1; transform: translateY(0px); }}
      }}

      /* DataFrame polish (Balanced SaaS) */
      div[data-testid="stDataFrame"] {{
        border-radius: 16px;
        overflow: hidden;
        border: 1px solid {THEME["border"]};
        box-shadow: 0 10px 22px rgba(0,0,0,0.06);
      }}

      /* Best-effort sticky header + wrap */
      div[data-testid="stDataFrame"] thead th {{
        position: sticky !important;
        top: 0 !important;
        background: {THEME["panel2"]} !important;
        z-index: 3 !important;
        border-bottom: 1px solid {THEME["border"]} !important;
        white-space: normal !important;
        line-height: 1.15 !important;
      }}

      /* Slightly compact rows */
      div[data-testid="stDataFrame"] td,
      div[data-testid="stDataFrame"] th {{
        padding-top: 0.35rem !important;
        padding-bottom: 0.35rem !important;
      }}
      div[data-testid="stDataFrame"] table {{
        font-size: 12px !important;
      }}

      /* Zebra striping (best-effort) */
      div[data-testid="stDataFrame"] tbody tr:nth-child(odd) td {{
        background: rgba(15, 23, 42, 0.02) !important;
      }}
    </style>

    <!-- Load Sora for wordmark -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Sora:wght@500;600;700&display=swap" rel="stylesheet">
    """,
    unsafe_allow_html=True,
)


# =========================
# Cache + Reload
# =========================
@st.cache_data
def load_data(file_path: Path) -> tuple[pd.DataFrame | None, str | None]:
    try:
        df = pd.read_excel(file_path)
        return df, None
    except FileNotFoundError:
        return None, "data.xlsx not found in the project folder."
    except Exception as e:
        return None, f"Error loading Excel: {e}"


def now_stamp() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


if "last_loaded" not in st.session_state:
    st.session_state.last_loaded = None


def trigger_reload():
    st.cache_data.clear()
    st.session_state.last_loaded = now_stamp()
    st.rerun()


# =========================
# Modeling helpers
# =========================
def safe_div(n, d):
    """Elementwise n/d for Series or scalar; returns NaN where denom is 0/NaN."""
    if pd.api.types.is_scalar(d):
        if pd.isna(d) or d == 0:
            return pd.Series(pd.NA, index=n.index, dtype="float64")
        return n / d
    d_clean = d.replace(0, pd.NA)
    return n / d_clean


def pick_tier_count(series: pd.Series, spread_threshold: float) -> int:
    """Return 3 tiers for narrow spread, 5 tiers for wide spread."""
    s = pd.to_numeric(series, errors="coerce").dropna()
    if s.empty:
        return 3
    spread = float(s.max() - s.min())
    return 5 if spread >= spread_threshold else 3


def assign_percentile_tiers(series: pd.Series, tier_count: int) -> pd.Series:
    """Percentile tiers robust to duplicates by ranking before qcut."""
    labels_3 = ["Low", "Medium", "High"]
    labels_5 = ["Very Low", "Low", "Medium", "High", "Very High"]
    labels = labels_5 if tier_count == 5 else labels_3

    s = pd.to_numeric(series, errors="coerce").astype("float64")
    ranks = s.rank(method="average", na_option="keep")

    try:
        cats = pd.qcut(ranks, q=tier_count, labels=labels, duplicates="drop")
        if hasattr(cats, "cat") and len(cats.cat.categories) < tier_count:
            cats = pd.qcut(ranks, q=3, labels=labels_3, duplicates="drop")
        return cats.astype("string")
    except Exception:
        return pd.Series(pd.NA, index=series.index, dtype="string")


def build_financial_model(raw_df: pd.DataFrame) -> pd.DataFrame:
    """
    Builds modelled financial dataset from raw Excel.
    Preserves numeric integrity (formatting is display-only).
    """
    dfm = raw_df.copy()

    for c in ["Units", "Retail Price", "Wholesale Price"]:
        if c in dfm.columns:
            dfm[c] = pd.to_numeric(dfm[c], errors="coerce")

    dfm["Sales"] = dfm["Units"] * dfm["Retail Price"]
    dfm["Profit per unit"] = dfm["Retail Price"] - dfm["Wholesale Price"]
    dfm["Sum Profit"] = dfm["Units"] * dfm["Profit per unit"]

    dfm["Inventory Cost"] = dfm["Units"] * dfm["Wholesale Price"]

    dfm["Gross Margin %"] = safe_div(dfm["Profit per unit"], dfm["Retail Price"]) * 100

    total_sales = dfm["Sales"].sum(min_count=1)
    total_profit = dfm["Sum Profit"].sum(min_count=1)

    dfm["Revenue Share %"] = safe_div(dfm["Sales"], total_sales) * 100
    dfm["Profit Share %"] = safe_div(dfm["Sum Profit"], total_profit) * 100

    dfm["Profit per $ Inventory"] = safe_div(dfm["Sum Profit"], dfm["Inventory Cost"]) * 100

    MARGIN_SPREAD_THRESHOLD_PTS = 15.0
    PRICE_SPREAD_THRESHOLD_DOLLARS = 1000.0

    margin_tier_count = pick_tier_count(dfm["Gross Margin %"], spread_threshold=MARGIN_SPREAD_THRESHOLD_PTS)
    price_tier_count = pick_tier_count(dfm["Retail Price"], spread_threshold=PRICE_SPREAD_THRESHOLD_DOLLARS)

    dfm["Margin Tier"] = assign_percentile_tiers(dfm["Gross Margin %"], margin_tier_count)
    dfm["Price Tier"] = assign_percentile_tiers(dfm["Retail Price"], price_tier_count)

    ordered_cols = [
        "Product",
        "Type",
        "Price Tier",
        "Margin Tier",
        "Units",
        "Retail Price",
        "Wholesale Price",
        "Sales",
        "Revenue Share %",
        "Profit per unit",
        "Gross Margin %",
        "Sum Profit",
        "Profit Share %",
        "Inventory Cost",
        "Profit per $ Inventory",
    ]
    ordered_cols = [c for c in ordered_cols if c in dfm.columns]
    return dfm[ordered_cols]


# =========================
# Insight catalog
# =========================
INSIGHTS = [
    "Portfolio Overview",
    "Revenue & Profit Contribution",
    "Capital Efficiency",
    "Pricing & Margin Discipline",
    "Concentration Risk",
    "Tier Structure",
    "Unit Economics",
    "Volume vs Margin Tradeoff",
    "Category (Type) Performance",
    "Underperformance & Exit Candidates",
]

GUIDE = {
    "Portfolio Overview": [
        "This dashboard separates activity from value. Activity is what we sell. Value is what we keep after costs.",
        "Use this overview to get your bearings on what’s driving the portfolio before drilling into pricing, capital efficiency, and risk.",
    ],
    "Revenue & Profit Contribution": [
        "Revenue tells us where the business spends most of its effort. Profit tells us what actually keeps the business strong.",
        "If revenue share is high but profit share is low, we may be busy without being paid back. If profit share is high but revenue share is low, it’s often a growth opportunity.",
    ],
    "Capital Efficiency": [
        "Inventory is money parked on shelves. The key question is whether that money is working hard enough to justify being tied up.",
        "This view highlights where inventory dollars produce strong returns, and where they drag on the business.",
    ],
    "Pricing & Margin Discipline": [
        "Pricing discipline is one of the fastest ways to improve profitability. If prices and margins don’t move together, something is misaligned.",
        "This view makes it easy to spot margin leaks and identify products where small pricing or cost changes could matter most.",
    ],
    "Concentration Risk": [
        "A business is fragile when too much performance depends on a small number of products.",
        "This view shows how quickly revenue and profit accumulate across the portfolio, and whether we rely too heavily on a few SKUs.",
    ],
    "Tier Structure": [
        "Tiers show the shape of the portfolio: where we compete on price, and where we earn on margin.",
        "This view shows how sales and profit distribute across price and margin tiers, and which combinations dominate performance.",
    ],
    "Unit Economics": [
        "Unit economics answer a simple question: how much value do we create each time we sell one unit?",
        "This view separates high-volume products from high-value products so we can focus effort where it compounds.",
    ],
    "Volume vs Margin Tradeoff": [
        "High volume can look impressive, but if margins are weak, volume can create operational strain without building wealth.",
        "This view reveals where we may be chasing volume at the cost of profitability and where sustainable volume exists.",
    ],
    "Category (Type) Performance": [
        "Categories behave differently. Some drive big-ticket revenue, others drive profitable attachment and repeat purchases.",
        "This view compares categories on revenue contribution, profit contribution, margins, and capital efficiency.",
    ],
    "Underperformance & Exit Candidates": [
        "Pruning underperformers is often the fastest way to improve portfolio quality.",
        "This view surfaces products that contribute little value relative to the attention and capital they consume.",
    ],
}

TABLE_VIEWS = {
    "Portfolio Overview": [
        "Product", "Type", "Sales", "Sum Profit", "Gross Margin %",
        "Revenue Share %", "Profit Share %", "Inventory Cost", "Profit per $ Inventory",
        "Price Tier", "Margin Tier",
    ],
    "Revenue & Profit Contribution": [
        "Product", "Type",
        "Sales", "Revenue Share %",
        "Sum Profit", "Profit Share %",
        "Gross Margin %", "Profit per unit",
        "Price Tier", "Margin Tier",
    ],
    "Capital Efficiency": [
        "Product", "Type",
        "Inventory Cost", "Profit per $ Inventory",
        "Sum Profit", "Sales", "Units", "Gross Margin %",
        "Price Tier", "Margin Tier",
    ],
    "Pricing & Margin Discipline": [
        "Product", "Type",
        "Retail Price", "Wholesale Price",
        "Profit per unit", "Gross Margin %",
        "Units", "Sales", "Sum Profit",
        "Price Tier", "Margin Tier",
    ],
    "Concentration Risk": [
        "Product", "Type",
        "Sales", "Revenue Share %",
        "Sum Profit", "Profit Share %",
        "Gross Margin %", "Units",
    ],
    "Tier Structure": [
        "Product", "Type",
        "Price Tier", "Margin Tier",
        "Sales", "Sum Profit", "Gross Margin %",
        "Revenue Share %", "Profit Share %",
    ],
    "Unit Economics": [
        "Product", "Type",
        "Units", "Profit per unit",
        "Sales", "Sum Profit", "Gross Margin %",
        "Retail Price", "Wholesale Price",
    ],
    "Volume vs Margin Tradeoff": [
        "Product", "Type",
        "Units", "Gross Margin %",
        "Sales", "Sum Profit",
        "Revenue Share %", "Profit Share %",
    ],
    "Category (Type) Performance": [
        "Type",
        "Sales", "Sum Profit",
        "Revenue Share %", "Profit Share %",
        "Gross Margin %", "Profit per $ Inventory",
        "Inventory Cost",
    ],
    "Underperformance & Exit Candidates": [
        "Product", "Type",
        "Sales", "Sum Profit",
        "Gross Margin %", "Profit per $ Inventory",
        "Revenue Share %", "Profit Share %",
        "Inventory Cost", "Units",
    ],
}


# =========================
# Deterministic Product Color Mapping (Unique per dataset)
# =========================
PRODUCT_PALETTE = [
    "#B91C1C", "#0F766E", "#1D4ED8", "#7C3AED", "#B45309",
    "#0E7490", "#BE185D", "#334155", "#166534", "#6D28D9",
    "#9A3412", "#0F172A", "#2563EB", "#059669", "#DC2626",
    "#9333EA", "#D97706", "#0891B2", "#DB2777", "#475569",
    "#16A34A", "#7C2D12", "#0EA5E9", "#A855F7", "#F97316",
]


def product_color_map(products: pd.Series) -> dict[str, str]:
    """
    Assign unique colors to products present in the dataset.
    Deterministic ordering: sorted product names -> stable mapping.
    """
    unique = sorted([p for p in products.dropna().astype(str).unique()])
    return {p: PRODUCT_PALETTE[i % len(PRODUCT_PALETTE)] for i, p in enumerate(unique)}


# =========================
# Quadrants (Auto labels)
# =========================
def compute_quadrants(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()
    x = d["Revenue Share %"]
    y = d["Profit Share %"]
    x_med = x.median()
    y_med = y.median()

    def label(xv, yv):
        if pd.isna(xv) or pd.isna(yv):
            return "Unclassified"
        hi_x = xv >= x_med
        hi_y = yv >= y_med
        if hi_x and hi_y:
            return "Scale & Protect"
        if hi_x and (not hi_y):
            return "Fix Economics"
        if (not hi_x) and hi_y:
            return "Growth Opportunity"
        return "Prune Candidate"

    d["Quadrant"] = [label(xv, yv) for xv, yv in zip(x, y)]
    d["Rev_Median"] = x_med
    d["Prof_Median"] = y_med
    return d


# =========================
# Formatting helpers (display only)
# =========================
def fmt_currency(x) -> str:
    if x is None or pd.isna(x):
        return "—"
    return f"${x:,.2f}"


def fmt_int(x) -> str:
    if x is None or pd.isna(x):
        return "—"
    return f"{int(round(x)):,.0f}"


def fmt_pct_pts(x) -> str:
    if x is None or pd.isna(x):
        return "—"
    return f"{x:,.2f}"


# =========================
# Plot theming (Soft Analytical)
# =========================
def apply_plot_theme(fig):
    fig.update_layout(
        template="plotly_white",
        paper_bgcolor=THEME["bg"],
        plot_bgcolor=THEME["panel"],
        font=dict(color=THEME["text"]),
        title=dict(font=dict(size=14, color=THEME["text"])),
        margin=dict(l=30, r=20, t=55, b=30),
        legend=dict(
            bgcolor="rgba(0,0,0,0)",
            bordercolor="rgba(0,0,0,0)",
            font=dict(color=THEME["muted"]),
        ),
    )
    fig.update_xaxes(
        gridcolor="rgba(17,24,39,0.08)",
        zerolinecolor="rgba(17,24,39,0.15)",
        linecolor="rgba(17,24,39,0.12)",
        tickfont=dict(color=THEME["muted"]),
        title_font=dict(color=THEME["muted"]),
    )
    fig.update_yaxes(
        gridcolor="rgba(17,24,39,0.08)",
        zerolinecolor="rgba(17,24,39,0.15)",
        linecolor="rgba(17,24,39,0.12)",
        tickfont=dict(color=THEME["muted"]),
        title_font=dict(color=THEME["muted"]),
    )
    return fig


# =========================
# UI blocks
# =========================
def lens_svg() -> str:
    return f"""
    <svg class="rl-brand-mark" viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
      <defs>
        <linearGradient id="g" x1="0" y1="0" x2="1" y2="1">
          <stop offset="0" stop-color="{THEME["accent2"]}" stop-opacity="0.85"/>
          <stop offset="1" stop-color="{THEME["accent"]}" stop-opacity="0.95"/>
        </linearGradient>
      </defs>
      <circle cx="28" cy="28" r="18" fill="none" stroke="url(#g)" stroke-width="5"/>
      <circle cx="28" cy="28" r="7" fill="none" stroke="{THEME["accent"]}" stroke-width="3" opacity="0.65"/>
      <path d="M40 40 L56 56" stroke="{THEME["accent"]}" stroke-width="6" stroke-linecap="round" opacity="0.85"/>
    </svg>
    """


def header_bar(view_title: str, last_loaded: str):
    st_html(
        f"""
        <div class="rl-header rl-fade">
          <div class="rl-brand">
            {lens_svg()}
            <div>
              <div class="rl-wordmark">Retail-Lens</div>
              <div class="rl-subtitle">Portfolio Intelligence Suite</div>
            </div>
          </div>
          <div class="rl-center-title">{view_title}</div>
          <div class="rl-right-meta">
            <div>{EXCEL_FILE.name}</div>
            <div style="margin-top:2px;">Last loaded: {last_loaded}</div>
          </div>
        </div>
        """
    )


def guide_block(lines: list[str]):
    st_html(
        f"""
        <div class="rl-guide rl-fade">
          <p>{lines[0]}</p>
          <p>{lines[1] if len(lines) > 1 else ""}</p>
        </div>
        """
    )


def kpi_row(df: pd.DataFrame):
    total_sales = df["Sales"].sum(min_count=1)
    total_profit = df["Sum Profit"].sum(min_count=1)
    avg_margin = df["Gross Margin %"].mean()
    profit_per_inv = safe_div(df["Sum Profit"], df["Inventory Cost"]).mean() * 100
    product_count = df["Product"].nunique() if "Product" in df.columns else len(df)

    cols = st.columns(5)
    kpis = [
        ("Total Sales", fmt_currency(total_sales), "Portfolio revenue"),
        ("Total Profit", fmt_currency(total_profit), "Portfolio gross profit"),
        ("Avg Gross Margin %", fmt_pct_pts(avg_margin), "Percent-points"),
        ("Profit per $ Inventory", fmt_pct_pts(profit_per_inv), "Percent-points"),
        ("Products", fmt_int(product_count), "Distinct products"),
    ]

    for c, (label, value, sub) in zip(cols, kpis):
        with c:
            st_html(
                f"""
                <div class="rl-card rl-fade">
                  <div class="rl-kpi-label">{label}</div>
                  <div class="rl-kpi-value">{value}</div>
                  <div class="rl-kpi-sub">{sub}</div>
                </div>
                """
            )


def view_df_for_insight(df: pd.DataFrame, insight_name: str) -> pd.DataFrame:
    cols = TABLE_VIEWS.get(insight_name, df.columns.tolist())
    cols = [c for c in cols if c in df.columns]
    return df[cols]


def render_table(df: pd.DataFrame, height: int = 520):
    df_view = df.copy()

    text_cols = [c for c in ["Product", "Type", "Price Tier", "Margin Tier"] if c in df_view.columns]
    numeric_cols = [c for c in df_view.columns if c not in text_cols]

    currency_cols = [c for c in ["Retail Price", "Wholesale Price", "Sales", "Profit per unit", "Sum Profit", "Inventory Cost"] if c in df_view.columns]
    int_cols = [c for c in ["Units"] if c in df_view.columns]
    pct_cols = [c for c in ["Gross Margin %", "Revenue Share %", "Profit Share %", "Profit per $ Inventory"] if c in df_view.columns]

    styler = (
        df_view.style
        .set_table_styles([
            {"selector": "th", "props": [
                ("background-color", THEME["panel2"]),
                ("color", THEME["text"]),
                ("border-bottom", f"1px solid {THEME['border']}"),
                ("white-space", "normal"),
                ("line-height", "1.15"),
            ]},
            {"selector": "td", "props": [
                ("border-bottom", "1px solid rgba(17,24,39,0.06)"),
            ]},
            {"selector": "table", "props": [("font-size", "12px")]},
        ])
        .set_properties(subset=text_cols, **{"text-align": "left"})
        .set_properties(subset=numeric_cols, **{"text-align": "right"})
    )

    fmt_map = {}
    for c in currency_cols:
        fmt_map[c] = "${:,.2f}"
    for c in int_cols:
        fmt_map[c] = "{:,.0f}"
    for c in pct_cols:
        fmt_map[c] = "{:,.2f}"

    if fmt_map:
        styler = styler.format(fmt_map, na_rep="—")

    st.dataframe(styler, use_container_width=True, height=height)


def uniform_scatter(fig, size: int = 10):
    fig.update_traces(marker=dict(size=size, opacity=0.9, line=dict(width=0)))
    return fig


# =========================
# Sidebar (Navigation + Reload)
# =========================
with st.sidebar:
    st_html(
        f"""
        <div style="padding: 8px 8px 2px 8px;">
          <div style="font-family:'Sora', ui-sans-serif, system-ui; font-weight:650; font-size:14px;">
            Retail-Lens
          </div>
          <div style="color:{THEME["muted"]}; font-size:12px; margin-top:2px;">
            Portfolio Intelligence Suite
          </div>
        </div>
        """
    )

    st.markdown("---")
    insight = st.radio("Insights", INSIGHTS, index=1)
    st.markdown("---")

    if st.button("Reload data", use_container_width=True):
        trigger_reload()


# =========================
# Load + model data
# =========================
raw_df, err = load_data(EXCEL_FILE)
if err:
    st.error(err)
    st.stop()

raw_df.index = raw_df.index + 1
model_df = build_financial_model(raw_df)

# Product color mapping (unique per dataset, consistent everywhere)
prod_colors = product_color_map(model_df["Product"])


# =========================
# Header bar
# =========================
last_loaded = st.session_state.last_loaded or "Auto (cache) / first run"
header_bar(insight, last_loaded)

# Guide + KPI row
guide_block(GUIDE.get(insight, ["", ""]))
kpi_row(model_df)

st_html("<div class='rl-fade' style='height: 10px;'></div>")


# =========================
# Page render
# =========================
if insight == "Portfolio Overview":
    c1, c2 = st.columns([1.35, 1.0])

    with c1:
        top = model_df.sort_values("Sales", ascending=False).head(12)
        fig = px.bar(
            top,
            x="Product",
            y="Sales",
            color="Product",
            color_discrete_map=prod_colors,
            title="Top Products by Sales",
        )
        fig = apply_plot_theme(fig)
        fig.update_layout(height=480, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        by_type = model_df.groupby("Type", dropna=False)[["Sales", "Sum Profit"]].sum().reset_index()
        fig = go.Figure()
        fig.add_trace(go.Bar(name="Sales", x=by_type["Type"], y=by_type["Sales"]))
        fig.add_trace(go.Bar(name="Profit", x=by_type["Type"], y=by_type["Sum Profit"]))
        fig.update_layout(
            barmode="group",
            title="Sales & Profit by Category (Type)",
            height=480,
            showlegend=True,
        )
        fig = apply_plot_theme(fig)
        st.plotly_chart(fig, use_container_width=True)

    st_html(
        f"""
        <div class="rl-fade" style="margin-top: 10px;">
          <div class="rl-section-title">Portfolio Excerpt</div>
          <div class="rl-section-sub">A compact slice of the most important rows for quick reading.</div>
        </div>
        """
    )

    excerpt = (
        model_df.sort_values("Sales", ascending=False)
        .loc[:, [c for c in TABLE_VIEWS["Portfolio Overview"] if c in model_df.columns]]
        .head(12)
    )
    render_table(excerpt, height=420)

elif insight == "Revenue & Profit Contribution":
    d = compute_quadrants(model_df)

    c1, c2 = st.columns(2)
    with c1:
        fig = px.pie(
            d,
            values="Sales",
            names="Product",
            color="Product",
            color_discrete_map=prod_colors,
            title="Revenue Share (Sales)",
        )
        fig = apply_plot_theme(fig)
        fig.update_layout(height=420, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        fig = px.pie(
            d,
            values="Sum Profit",
            names="Product",
            color="Product",
            color_discrete_map=prod_colors,
            title="Profit Share (Sum Profit)",
        )
        fig = apply_plot_theme(fig)
        fig.update_layout(height=420, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    fig = px.scatter(
        d,
        x="Revenue Share %",
        y="Profit Share %",
        color="Quadrant",
        hover_name="Product",
        hover_data=["Type", "Units", "Gross Margin %", "Profit per $ Inventory"],
        title="Revenue Share vs Profit Share (Auto Quadrants)",
        category_orders={
            "Quadrant": ["Scale & Protect", "Fix Economics", "Growth Opportunity", "Prune Candidate", "Unclassified"]
        },
        color_discrete_map={
            "Scale & Protect": THEME["pos"],
            "Fix Economics": THEME["warn"],
            "Growth Opportunity": THEME["accent"],
            "Prune Candidate": THEME["neg"],
            "Unclassified": THEME["muted2"],
        },
    )
    fig = apply_plot_theme(fig)
    fig.update_layout(height=520, showlegend=True)

    fig = uniform_scatter(fig, size=11)

    x_med = float(d["Rev_Median"].iloc[0]) if len(d) else 0
    y_med = float(d["Prof_Median"].iloc[0]) if len(d) else 0
    fig.add_vline(x=x_med, line_width=1, line_dash="dot", line_color="rgba(75,85,99,0.55)")
    fig.add_hline(y=y_med, line_width=1, line_dash="dot", line_color="rgba(75,85,99,0.55)")

    fig.add_annotation(xref="paper", yref="paper", x=0.02, y=0.98, text="Growth Opportunity", showarrow=False,
                       font=dict(color="rgba(75,85,99,0.75)", size=12))
    fig.add_annotation(xref="paper", yref="paper", x=0.98, y=0.98, text="Scale & Protect", showarrow=False,
                       font=dict(color="rgba(75,85,99,0.75)", size=12), xanchor="right")
    fig.add_annotation(xref="paper", yref="paper", x=0.02, y=0.02, text="Prune Candidate", showarrow=False,
                       font=dict(color="rgba(75,85,99,0.75)", size=12), yanchor="bottom")
    fig.add_annotation(xref="paper", yref="paper", x=0.98, y=0.02, text="Fix Economics", showarrow=False,
                       font=dict(color="rgba(75,85,99,0.75)", size=12), xanchor="right", yanchor="bottom")

    st.plotly_chart(fig, use_container_width=True)

elif insight == "Capital Efficiency":
    c1, c2 = st.columns([1.25, 1.0])
    with c1:
        fig = px.scatter(
            model_df,
            x="Inventory Cost",
            y="Profit per $ Inventory",
            color="Product",
            color_discrete_map=prod_colors,
            hover_name="Product",
            hover_data=["Type", "Sum Profit", "Gross Margin %", "Units"],
            title="Inventory Cost vs Profit per $ Inventory",
        )
        fig = apply_plot_theme(fig)
        fig.update_layout(height=520, showlegend=False)
        fig = uniform_scatter(fig, size=11)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        worst = model_df.sort_values("Profit per $ Inventory", ascending=True).head(12)
        fig = px.bar(
            worst,
            x="Product",
            y="Profit per $ Inventory",
            color="Product",
            color_discrete_map=prod_colors,
            title="Lowest Inventory Efficiency (Bottom Products)",
        )
        fig = apply_plot_theme(fig)
        fig.update_layout(height=520, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        top = model_df.sort_values("Profit per $ Inventory", ascending=False).head(12)
        fig = px.bar(top, x="Product", y="Profit per $ Inventory", color="Product", color_discrete_map=prod_colors,
                     title="Highest Inventory Efficiency")
        fig = apply_plot_theme(fig)
        fig.update_layout(height=380, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    with c4:
        heavy = model_df.sort_values("Inventory Cost", ascending=False).head(12)
        fig = px.bar(heavy, x="Product", y="Inventory Cost", color="Product", color_discrete_map=prod_colors,
                     title="Largest Inventory Tie-Up")
        fig = apply_plot_theme(fig)
        fig.update_layout(height=380, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

elif insight == "Pricing & Margin Discipline":
    c1, c2 = st.columns(2)

    with c1:
        m = model_df.sort_values("Gross Margin %", ascending=False)
        fig = px.bar(
            m,
            x="Product",
            y="Gross Margin %",
            color="Product",
            color_discrete_map=prod_colors,
            title="Gross Margin % by Product (Ranked)",
        )
        fig = apply_plot_theme(fig)
        fig.update_layout(height=520, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        ppu = model_df.sort_values("Profit per unit", ascending=False)
        fig = px.bar(
            ppu,
            x="Product",
            y="Profit per unit",
            color="Product",
            color_discrete_map=prod_colors,
            title="Profit per Unit by Product (Ranked)",
        )
        fig = apply_plot_theme(fig)
        fig.update_layout(height=520, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        rp = model_df.sort_values("Retail Price", ascending=False)
        fig = px.bar(
            rp,
            x="Product",
            y="Retail Price",
            color="Product",
            color_discrete_map=prod_colors,
            title="Retail Price by Product (Ranked)",
        )
        fig = apply_plot_theme(fig)
        fig.update_layout(height=380, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    with c4:
        wp = model_df.sort_values("Wholesale Price", ascending=False)
        fig = px.bar(
            wp,
            x="Product",
            y="Wholesale Price",
            color="Product",
            color_discrete_map=prod_colors,
            title="Wholesale Price by Product (Ranked)",
        )
        fig = apply_plot_theme(fig)
        fig.update_layout(height=380, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

elif insight == "Concentration Risk":
    d = model_df[["Product", "Sales", "Sum Profit"]].copy()
    d = d.sort_values("Sales", ascending=False).reset_index(drop=True)
    d["Revenue Cumulative %"] = safe_div(d["Sales"].cumsum(), d["Sales"].sum(min_count=1)) * 100
    d["Profit Cumulative %"] = safe_div(d["Sum Profit"].cumsum(), d["Sum Profit"].sum(min_count=1)) * 100
    d["Rank"] = d.index + 1

    fig = go.Figure()
    fig.add_trace(go.Bar(x=d["Rank"], y=d["Sales"], name="Sales", marker_color=THEME["accent"]))
    fig.add_trace(go.Scatter(
        x=d["Rank"], y=d["Revenue Cumulative %"], name="Cumulative Revenue %",
        yaxis="y2", mode="lines+markers", line=dict(color=THEME["accent2"])
    ))
    fig.update_layout(
        title="Revenue Concentration (Pareto)",
        height=520,
        yaxis=dict(title="Sales"),
        yaxis2=dict(title="Cumulative %", overlaying="y", side="right", range=[0, 100]),
        xaxis=dict(title="Product rank (by Sales)"),
        showlegend=True,
    )
    fig = apply_plot_theme(fig)
    st.plotly_chart(fig, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        fig = go.Figure()
        fig.add_trace(go.Bar(x=d["Rank"], y=d["Sum Profit"], name="Profit", marker_color=THEME["accent"]))
        fig.add_trace(go.Scatter(
            x=d["Rank"], y=d["Profit Cumulative %"], name="Cumulative Profit %",
            yaxis="y2", mode="lines+markers", line=dict(color=THEME["accent2"])
        ))
        fig.update_layout(
            title="Profit Concentration (Pareto)",
            height=420,
            yaxis=dict(title="Profit"),
            yaxis2=dict(title="Cumulative %", overlaying="y", side="right", range=[0, 100]),
            xaxis=dict(title="Product rank (by Profit)"),
            showlegend=True,
        )
        fig = apply_plot_theme(fig)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        top_n = min(8, len(d))
        top_rev = d.head(top_n)
        fig = px.pie(
            top_rev,
            values="Sales",
            names="Product",
            color="Product",
            color_discrete_map=prod_colors,
            title=f"Top {top_n} Products by Revenue Share",
        )
        fig = apply_plot_theme(fig)
        fig.update_layout(height=420, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

elif insight == "Tier Structure":
    pivot_sales = model_df.pivot_table(
        index="Price Tier", columns="Margin Tier", values="Sales", aggfunc="sum", dropna=False
    ).fillna(0)

    pivot_profit = model_df.pivot_table(
        index="Price Tier", columns="Margin Tier", values="Sum Profit", aggfunc="sum", dropna=False
    ).fillna(0)

    c1, c2 = st.columns(2)
    with c1:
        fig = px.imshow(pivot_sales, title="Sales by Price Tier × Margin Tier", aspect="auto")
        fig = apply_plot_theme(fig)
        fig.update_layout(height=480, coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        fig = px.imshow(pivot_profit, title="Profit by Price Tier × Margin Tier", aspect="auto")
        fig = apply_plot_theme(fig)
        fig.update_layout(height=480, coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        by_price = model_df.groupby("Price Tier", dropna=False)["Sales"].sum().reset_index().sort_values("Sales", ascending=False)
        fig = px.bar(by_price, x="Price Tier", y="Sales", title="Sales by Price Tier", color_discrete_sequence=[THEME["accent"]])
        fig = apply_plot_theme(fig)
        fig.update_layout(height=360, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    with c4:
        by_margin = model_df.groupby("Margin Tier", dropna=False)["Sum Profit"].sum().reset_index().sort_values("Sum Profit", ascending=False)
        fig = px.bar(by_margin, x="Margin Tier", y="Sum Profit", title="Profit by Margin Tier", color_discrete_sequence=[THEME["accent"]])
        fig = apply_plot_theme(fig)
        fig.update_layout(height=360, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

elif insight == "Unit Economics":
    c1, c2 = st.columns([1.25, 1.0])

    with c1:
        fig = px.scatter(
            model_df,
            x="Units",
            y="Profit per unit",
            color="Product",
            color_discrete_map=prod_colors,
            hover_name="Product",
            hover_data=["Type", "Gross Margin %", "Retail Price", "Wholesale Price"],
            title="Units vs Profit per Unit",
        )
        fig = apply_plot_theme(fig)
        fig.update_layout(height=520, showlegend=False)
        fig = uniform_scatter(fig, size=11)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        best = model_df.sort_values("Profit per unit", ascending=False).head(12)
        fig = px.bar(best, x="Product", y="Profit per unit", color="Product", color_discrete_map=prod_colors,
                     title="Highest Profit per Unit")
        fig = apply_plot_theme(fig)
        fig.update_layout(height=520, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

elif insight == "Volume vs Margin Tradeoff":
    c1, c2 = st.columns(2)

    with c1:
        fig = px.scatter(
            model_df,
            x="Units",
            y="Gross Margin %",
            color="Product",
            color_discrete_map=prod_colors,
            hover_name="Product",
            hover_data=["Type", "Revenue Share %", "Profit Share %", "Profit per unit"],
            title="Units vs Gross Margin %",
        )
        fig = apply_plot_theme(fig)
        fig.update_layout(height=480, showlegend=False)
        fig = uniform_scatter(fig, size=11)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        fig = px.scatter(
            model_df,
            x="Sales",
            y="Gross Margin %",
            color="Product",
            color_discrete_map=prod_colors,
            hover_name="Product",
            hover_data=["Type", "Revenue Share %", "Profit Share %", "Profit per unit"],
            title="Sales vs Gross Margin %",
        )
        fig = apply_plot_theme(fig)
        fig.update_layout(height=480, showlegend=False)
        fig = uniform_scatter(fig, size=11)
        st.plotly_chart(fig, use_container_width=True)

elif insight == "Category (Type) Performance":
    by_type = model_df.groupby("Type", dropna=False).agg(
        Sales=("Sales", "sum"),
        Profit=("Sum Profit", "sum"),
        AvgMargin=("Gross Margin %", "mean"),
        Inventory=("Inventory Cost", "sum"),
        ProfitPerInv=("Profit per $ Inventory", "mean"),
        Products=("Product", "nunique"),
    ).reset_index()

    c1, c2 = st.columns([1.2, 1.0])

    with c1:
        fig = go.Figure()
        fig.add_trace(go.Bar(name="Sales", x=by_type["Type"], y=by_type["Sales"], marker_color=THEME["accent"]))
        fig.add_trace(go.Bar(name="Profit", x=by_type["Type"], y=by_type["Profit"], marker_color=THEME["accent2"]))
        fig.update_layout(barmode="group", title="Sales & Profit by Type", height=520, showlegend=True)
        fig = apply_plot_theme(fig)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        fig = px.bar(by_type, x="Type", y="AvgMargin", title="Average Gross Margin % by Type",
                     color_discrete_sequence=[THEME["accent"]])
        fig = apply_plot_theme(fig)
        fig.update_layout(height=520, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        fig = px.bar(by_type, x="Type", y="Inventory", title="Inventory Tie-Up by Type",
                     color_discrete_sequence=[THEME["accent"]])
        fig = apply_plot_theme(fig)
        fig.update_layout(height=380, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    with c4:
        fig = px.bar(by_type, x="Type", y="ProfitPerInv", title="Profit per $ Inventory (Avg) by Type",
                     color_discrete_sequence=[THEME["accent"]])
        fig = apply_plot_theme(fig)
        fig.update_layout(height=380, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

elif insight == "Underperformance & Exit Candidates":
    d = model_df.copy()

    for col in ["Gross Margin %", "Profit Share %", "Profit per $ Inventory"]:
        d[col + " Rank"] = d[col].rank(pct=True, ascending=True, na_option="keep")
    d["Concern Score"] = (1 - d["Gross Margin % Rank"]) + (1 - d["Profit Share % Rank"]) + (1 - d["Profit per $ Inventory Rank"])

    c1, c2 = st.columns([1.25, 1.0])

    with c1:
        fig = px.scatter(
            d,
            x="Revenue Share %",
            y="Profit Share %",
            color="Concern Score",
            hover_name="Product",
            hover_data=["Type", "Gross Margin %", "Profit per $ Inventory", "Inventory Cost"],
            title="Concern Map (Revenue vs Profit Share)",
            color_continuous_scale="RdYlBu_r",
        )
        fig = apply_plot_theme(fig)
        fig.update_layout(height=520, showlegend=False)
        fig = uniform_scatter(fig, size=11)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        worst = d.sort_values("Concern Score", ascending=False).head(12)
        fig = px.bar(
            worst,
            x="Product",
            y="Concern Score",
            color="Product",
            color_discrete_map=prod_colors,
            title="Highest Concern (Top Candidates)",
        )
        fig = apply_plot_theme(fig)
        fig.update_layout(height=520, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

else:
    st.info("Insight view not implemented yet.")


# =========================
# Detail table (category-driven slice)
# =========================
st_html("<div style='height:10px;'></div>")
st_html(
    f"""
    <div class="rl-fade">
      <div class="rl-section-title">Detailed Portfolio View</div>
      <div class="rl-section-sub">Sliced to match the selected insight category.</div>
    </div>
    """
)

table_df = view_df_for_insight(model_df, insight)

row_height = 34
max_height = 520
table_height = min(len(table_df) * row_height + 42, max_height)
render_table(table_df, height=table_height)

st_html(
    f"""
    <div style="text-align:center; color:{THEME["muted2"]}; margin-top:14px; font-size:12px;">
      Retail-Lens — Portfolio Intelligence Suite
    </div>
    """
)
