"""
Pipeline Analytics Dashboard
A professional Streamlit application for offer pipeline analysis.
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import warnings
from datetime import datetime, date
import io

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Bidding & Design Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# THEME / CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

/* Main background — warm off-white cream */
.stApp { background: #f5f7f0; }
.main .block-container { background: #f5f7f0; padding-top: 1.5rem; }

/* Sidebar — deep forest green */
[data-testid="stSidebar"] {
    background: #1b4332;
    border-right: 2px solid #2d6a4f;
}
[data-testid="stSidebar"] * { color: #d8f3dc !important; }
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiSelect label {
    color: #95d5b2 !important;
    font-size: 0.78rem; font-weight: 600;
    letter-spacing: 0.05em; text-transform: uppercase;
}

/* KPI Cards — each with its own bg + contrasting text */
.kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 14px; margin-bottom: 24px; }
.kpi-card {
    border-radius: 12px;
    padding: 18px 20px 16px;
    position: relative;
    overflow: hidden;
    box-shadow: 0 3px 10px rgba(0,0,0,0.10);
    background: var(--bg, #ffffff);
}
.kpi-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 4px;
    background: var(--accent, #2d6a4f);
    border-radius: 12px 12px 0 0;
}
.kpi-label { font-size: 0.72rem; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; color: var(--label, #4a7c59); margin-bottom: 6px; }
.kpi-value { font-size: 1.7rem; font-weight: 700; color: var(--val, #1b4332); line-height: 1.1; font-family: 'DM Mono', monospace; }
.kpi-sub { font-size: 0.75rem; color: var(--sub, #74b49b); margin-top: 4px; }

/* Section headers */
.section-header {
    font-size: 0.7rem; font-weight: 700; letter-spacing: 0.12em;
    text-transform: uppercase; color: #1b4332;
    border-bottom: 2px solid #b7e4c7;
    padding-bottom: 8px; margin: 28px 0 18px;
    background: linear-gradient(90deg, #d8f3dc 0%, transparent 100%);
    padding-left: 10px; border-radius: 4px 0 0 0;
}

/* Selectbox dropdown text — force black for readability */
[data-testid="stSelectbox"] div[data-baseweb="select"] span,
[data-testid="stSelectbox"] div[data-baseweb="select"] div,
[data-baseweb="popover"] li,
[data-baseweb="popover"] ul span,
[role="listbox"] li,
[role="option"] { color: #111111 !important; }
[data-baseweb="popover"] { background: #ffffff !important; }
[role="listbox"] { background: #ffffff !important; }

/* Tabs */
[data-testid="stTabs"] button { font-size: 0.82rem; font-weight: 600; color: #2d6a4f; }
[data-testid="stTabs"] button[aria-selected="true"] { color: #1b4332 !important; border-bottom-color: #2d6a4f !important; }

/* General text — dark charcoal on light bg */
p, div, span, label { color: #2d3a2e; }

/* Hide Streamlit branding */
#MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────
PLOTLY_THEME = dict(
    paper_bgcolor="rgba(255,255,255,0.85)",
    plot_bgcolor="rgba(255,255,255,0.85)",
    font=dict(family="DM Sans", color="#1b5e20", size=12),
    xaxis=dict(gridcolor="#c8e6c9", linecolor="#a5d6a7", tickfont=dict(size=11, color="#2e7d32")),
    yaxis=dict(gridcolor="#c8e6c9", linecolor="#a5d6a7", tickfont=dict(size=11, color="#2e7d32")),
    legend=dict(bgcolor="rgba(255,255,255,0.6)", font=dict(size=11, color="#1b5e20")),
    margin=dict(l=10, r=10, t=40, b=10),
)

PALETTE = ["#4f8ef7", "#3ddbb9", "#f7c054", "#f7545f", "#a78df7", "#54c5f7", "#f77c4f", "#8ef77c"]

TIP_OFERTA_TYPES = ["Bugetara", "Angajanta", "Revizie"]
STATUS_TYPES = ["Design", "KO", "Lost", "Presented", "Qualification", "Signed"]

# ─────────────────────────────────────────────
# DATA LOADING & CLEANING
# ─────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data(file_bytes: bytes) -> pd.DataFrame:
    """Load and clean the Excel file."""
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0)

    # Normalize column names: strip whitespace and normalize spaces
    df.columns = [str(c).strip().replace("\xa0", " ").replace("  ", " ") for c in df.columns]

    # Normalize all string columns: strip trailing/leading spaces
    # This fixes issues like "Stefan Patru " vs "Stefan Patru"
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).str.strip().replace("nan", np.nan)

    # Robust date parser — tries multiple formats
    def parse_date_col(series: pd.Series) -> pd.Series:
        if pd.api.types.is_datetime64_any_dtype(series):
            return series
        parsed = pd.to_datetime(series, errors="coerce", dayfirst=True)
        if parsed.notna().mean() >= 0.8:
            return parsed
        parsed2 = pd.to_datetime(series, errors="coerce", dayfirst=False)
        if parsed2.notna().mean() >= parsed.notna().mean():
            return parsed2
        return parsed

    date_cols = ["Data solicitare oferta", "Data start oferta",
                 "Data transmitere oferta", "Data estimata semnare contract"]
    for col in date_cols:
        if col in df.columns:
            df[col] = parse_date_col(df[col])

    # Numeric columns — NOTE: iKPI/proiect contains unit labels (MWp, MWh, LP, kVA...)
    # so it is kept as string (unit of measure), NOT converted to numeric
    num_cols = ["Revenues [MEuro]", "GM [MEuro]", "GM %", "iKPI [Valoare]",
                "Probabilitate semnare contract [%]"]
    for col in num_cols:
        if col in df.columns:
            if df[col].dtype == object:
                df[col] = df[col].astype(str).str.replace(",", ".").str.strip()
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # iKPI/proiect — keep as unit label string, already stripped above
    # iKPI [Valoare] — numeric value already handled above

    # GM % — if stored as 0-100 range, convert to 0-1
    if "GM %" in df.columns:
        valid = df["GM %"].dropna()
        if len(valid) > 0 and valid.max() > 1:
            df["GM %"] = df["GM %"] / 100

    # Extract Year / Month from Data start oferta (primary date for filtering)
    ref_col = "Data start oferta"
    if ref_col in df.columns:
        df["_year"] = df[ref_col].dt.year
        df["_month"] = df[ref_col].dt.month
    else:
        for fallback in ["Data solicitare oferta", "Data transmitere oferta"]:
            if fallback in df.columns:
                df["_year"] = df[fallback].dt.year
                df["_month"] = df[fallback].dt.month
                st.warning(f"Coloana 'Data start oferta' nu a fost găsită. Se folosește '{fallback}' pentru filtrare.")
                break

    return df


def business_days(start: pd.Timestamp, end: pd.Timestamp) -> int:
    """Calculate number of business days between two dates (excluding weekends)."""
    if pd.isna(start) or pd.isna(end):
        return np.nan
    if end < start:
        return np.nan
    return np.busday_count(start.date(), end.date())


def calc_processing_time(df: pd.DataFrame) -> pd.Series:
    """Vectorised business day calculation for processing time."""
    if "Data start oferta" not in df.columns or "Data transmitere oferta" not in df.columns:
        return pd.Series(np.nan, index=df.index)
    return df.apply(
        lambda r: business_days(r["Data start oferta"], r["Data transmitere oferta"]), axis=1
    )


# ─────────────────────────────────────────────
# BUSINESS RULE ENGINE
# ─────────────────────────────────────────────
def apply_business_rules(df: pd.DataFrame) -> pd.DataFrame:
    """
    Apply the three main business rules before aggregation:
      Rule 1: If Bugetara + Angajanta pair exists for same Client + VAS → keep only Angajanta
      Rule 2: If multiple Revizie → keep only the latest one
      Rule 3: GM % as arithmetic average (applied at aggregation time)
      Rule 4: Exclude rows with missing key fields
    """
    required_cols = ["Revenues [MEuro]", "GM [MEuro]", "VAS"]
    work = df.copy()

    # Rule 4 — drop rows missing key fields
    work = work.dropna(subset=[c for c in required_cols if c in work.columns])

    if "Tip oferta" not in work.columns:
        return work

    # Rule 1 — Bugetara vs Angajanta per Client + VAS
    key_cols = ["CUI client", "VAS"]
    key_cols = [c for c in key_cols if c in work.columns]

    if key_cols and "Tip oferta" in work.columns:
        ang = work[work["Tip oferta"].str.strip() == "Angajanta"][key_cols].drop_duplicates()
        if not ang.empty:
            # Mark Bugetara rows that have a matching Angajanta
            bud_mask = work["Tip oferta"].str.strip() == "Bugetara"
            merged = work[bud_mask].merge(ang, on=key_cols, how="left", indicator=True)
            has_ang = merged["_merge"] == "both"
            bud_to_drop = work[bud_mask].index[has_ang.values]
            work = work.drop(index=bud_to_drop)

    # Rule 2 — Revisions: keep latest only per Client + VAS
    rev_mask = work["Tip oferta"].str.strip() == "Revizie"
    revs = work[rev_mask].copy()
    if not revs.empty and "Data start oferta" in revs.columns and key_cols:
        # Sort descending, keep first (latest) per group
        revs_sorted = revs.sort_values("Data start oferta", ascending=False)
        latest_revs = revs_sorted.drop_duplicates(subset=key_cols, keep="first")
        older_revs_idx = revs.index.difference(latest_revs.index)
        work = work.drop(index=older_revs_idx)

    return work


# ─────────────────────────────────────────────
# FILTER HELPERS
# ─────────────────────────────────────────────
def filter_df(df: pd.DataFrame, year: int, month: int = None) -> pd.DataFrame:
    """Filter dataframe by year and optional month."""
    out = df[df["_year"] == year]
    if month is not None:
        out = out[out["_month"] == month]
    return out


# ─────────────────────────────────────────────
# CHART HELPERS
# ─────────────────────────────────────────────
def styled_bar(fig: go.Figure) -> go.Figure:
    fig.update_layout(**PLOTLY_THEME)
    fig.update_traces(marker_line_width=0)
    return fig


def kpi_card(label: str, value: str, sub: str = "", accent: str = "#2d6a4f",
             bg: str = "#ffffff", val_color: str = "#1b4332",
             label_color: str = "#4a7c59", sub_color: str = "#74b49b") -> str:
    return f"""
    <div class="kpi-card" style="--accent:{accent};--bg:{bg};--val:{val_color};--label:{label_color};--sub:{sub_color};">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-sub">{sub}</div>
    </div>"""


def fmt_num(n, decimals=1, prefix="", suffix=""):
    if pd.isna(n):
        return "—"
    return f"{prefix}{n:,.{decimals}f}{suffix}"


# ─────────────────────────────────────────────
# SECTION 1 — PIPELINE KPIs
# ─────────────────────────────────────────────
def section_pipeline_kpis(df: pd.DataFrame):
    st.markdown('<div class="section-header">01 — Pipeline Key Indicators</div>', unsafe_allow_html=True)

    pt = calc_processing_time(df)
    avg_pt = pt.mean()

    total_oferte = len(df[df["Tip oferta"].str.strip().isin(TIP_OFERTA_TYPES)]) if "Tip oferta" in df.columns else len(df)
    total_rev = df["Revenues [MEuro]"].sum() if "Revenues [MEuro]" in df.columns else 0
    total_gm = df["GM [MEuro]"].sum() if "GM [MEuro]" in df.columns else 0
    avg_gm_pct = df["GM %"].mean() * 100 if "GM %" in df.columns else 0
    total_ikpi_proj = df["iKPI/proiect"].sum() if "iKPI/proiect" in df.columns else 0
    sum_ikpi_val = df["iKPI [Valoare]"].sum() if "iKPI [Valoare]" in df.columns else 0

    cards_html = '<div class="kpi-grid">'
    # Card 1 — deep forest green bg, cream text
    cards_html += kpi_card("Total Oferte", fmt_num(total_oferte, 0), "Bugetara + Angajanta + Revizie",
                           accent="#52b788", bg="#1b4332", val_color="#d8f3dc", label_color="#95d5b2", sub_color="#52b788")
    # Card 2 — warm amber bg, dark brown text
    cards_html += kpi_card("Total Revenue", fmt_num(total_rev, 2, "€", "M"), "MEuro",
                           accent="#e07c24", bg="#fff3e0", val_color="#7b3f00", label_color="#c25e00", sub_color="#e09a50")
    # Card 3 — slate blue bg, white text
    cards_html += kpi_card("Total GM", fmt_num(total_gm, 2, "€", "M"), "MEuro",
                           accent="#4a90d9", bg="#e8f0fe", val_color="#1a3a6b", label_color="#2c5fa8", sub_color="#6aa3e0")
    # Card 4 — soft teal bg, deep teal text
    cards_html += kpi_card("Average GM %", fmt_num(avg_gm_pct, 1, "", "%"), "Arithmetic avg",
                           accent="#2b9d8f", bg="#e0f2f1", val_color="#0d3b36", label_color="#1a7a6e", sub_color="#4dbfb3")
    # Card 5 — sage green bg, dark olive text
    cards_html += kpi_card("Avg Processing Time", fmt_num(avg_pt, 1, "", " days"), "Business days",
                           accent="#6a994e", bg="#f0f4e8", val_color="#2c4a1e", label_color="#4d7a2a", sub_color="#8ab86a")
    cards_html += '</div>'
    st.markdown(cards_html, unsafe_allow_html=True)


# ─────────────────────────────────────────────
# SECTION 2 — OFFER ACTIVITY
# ─────────────────────────────────────────────
def section_offer_activity(df: pd.DataFrame):
    st.markdown('<div class="section-header">02 — Offer Activity</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)

    with col1:
        if "Tip oferta" in df.columns:
            tip_counts = df["Tip oferta"].str.strip().value_counts().reset_index()
            tip_counts.columns = ["Tip oferta", "Count"]
            fig = px.bar(tip_counts, x="Tip oferta", y="Count", color="Tip oferta",
                         title="Offers by Type", color_discrete_sequence=PALETTE,
                         text="Count")
            fig.update_traces(textposition="outside")
            st.plotly_chart(styled_bar(fig), use_container_width=True)

    with col2:
        if "VAS" in df.columns:
            vas_counts = df["VAS"].value_counts().reset_index()
            vas_counts.columns = ["VAS", "Count"]
            top_vas = vas_counts.head(10)
            fig2 = px.pie(top_vas, names="VAS", values="Count",
                          title="Distribution by VAS (Top 10)",
                          color_discrete_sequence=PALETTE, hole=0.4)
            fig2.update_layout(**PLOTLY_THEME)
            st.plotly_chart(fig2, use_container_width=True)


# ─────────────────────────────────────────────
# SECTION 3 — PRODUCT TYPE DISTRIBUTION
# ─────────────────────────────────────────────
def section_product_type(df: pd.DataFrame):
    st.markdown('<div class="section-header">03 — Product Type Distribution</div>', unsafe_allow_html=True)

    if "Tip oferta" not in df.columns or "VAS" not in df.columns:
        st.info("Columns Tip oferta / VAS not found.")
        return

    pivot = (df[df["Tip oferta"].str.strip().isin(["Bugetara", "Angajanta"])]
             .groupby(["VAS", "Tip oferta"])
             .size().reset_index(name="Count"))

    if pivot.empty:
        st.info("No data for this period.")
        return

    # ── Stacked chart: Bugetara vs Angajanta by VAS ──
    fig = px.bar(pivot, x="VAS", y="Count", color="Tip oferta",
                 barmode="stack", title="Bugetara vs Angajanta by VAS",
                 color_discrete_map={"Bugetara": PALETTE[0], "Angajanta": PALETTE[1]},
                 text="Count")
    fig.update_traces(textposition="inside")
    st.plotly_chart(styled_bar(fig), use_container_width=True)

    # ── KPI cards per VAS (same indicators as section 01) ──
    st.markdown('<div class="section-header">03b — Key Indicators per VAS</div>', unsafe_allow_html=True)

    vas_list = sorted(df["VAS"].dropna().unique().tolist())
    if not vas_list:
        return

    # Let user pick which VAS to inspect, default = all
    selected_vas = st.selectbox("Select VAS", ["All"] + vas_list, key="pt_vas_select")
    vas_df = df if selected_vas == "All" else df[df["VAS"] == selected_vas]

    # Build per-VAS summary table
    rows = []
    for vas, grp in df.groupby("VAS"):
        pt = calc_processing_time(grp)
        total_oferte = len(grp[grp["Tip oferta"].str.strip().isin(TIP_OFERTA_TYPES)]) if "Tip oferta" in grp.columns else len(grp)
        # iKPI/proiect is a unit label (MWp, MWh, LP, kVA...) — show most common unit
        ikpi_unit = grp["iKPI/proiect"].mode()[0] if "iKPI/proiect" in grp.columns and not grp["iKPI/proiect"].dropna().empty else "—"
        rows.append({
            "VAS": vas,
            "Total Oferte": total_oferte,
            "Revenue (MEuro)": round(grp["Revenues [MEuro]"].sum(), 2) if "Revenues [MEuro]" in grp.columns else 0,
            "GM (MEuro)": round(grp["GM [MEuro]"].sum(), 2) if "GM [MEuro]" in grp.columns else 0,
            "Avg GM %": round(grp["GM %"].mean() * 100, 1) if "GM %" in grp.columns else 0,
            "iKPI Unit": ikpi_unit,
            "iKPI Value": round(grp["iKPI [Valoare]"].sum(), 2) if "iKPI [Valoare]" in grp.columns else 0,
            "Avg Processing Time (days)": round(pt.mean(), 1) if not pt.isna().all() else 0,
        })
    vas_summary = pd.DataFrame(rows)

    # Display KPI cards for the selected VAS (or global if All)
    if selected_vas != "All":
        row = vas_summary[vas_summary["VAS"] == selected_vas]
        if not row.empty:
            r = row.iloc[0]
            cards_html = '<div class="kpi-grid">'
            cards_html += kpi_card("Total Oferte", fmt_num(r["Total Oferte"], 0), f"VAS: {selected_vas}",
                                   accent="#52b788", bg="#1b4332", val_color="#d8f3dc", label_color="#95d5b2", sub_color="#52b788")
            cards_html += kpi_card("Revenue", fmt_num(r["Revenue (MEuro)"], 2, "€", "M"), "MEuro",
                                   accent="#e07c24", bg="#fff3e0", val_color="#7b3f00", label_color="#c25e00", sub_color="#e09a50")
            cards_html += kpi_card("GM", fmt_num(r["GM (MEuro)"], 2, "€", "M"), "MEuro",
                                   accent="#4a90d9", bg="#e8f0fe", val_color="#1a3a6b", label_color="#2c5fa8", sub_color="#6aa3e0")
            cards_html += kpi_card("Avg GM %", fmt_num(r["Avg GM %"], 1, "", "%"), "Arithmetic avg",
                                   accent="#2b9d8f", bg="#e0f2f1", val_color="#0d3b36", label_color="#1a7a6e", sub_color="#4dbfb3")
            cards_html += kpi_card("iKPI Projects", str(r["iKPI Unit"]), "Unitate iKPI/proiect",
                                   accent="#9b6fa8", bg="#f3e8f8", val_color="#3b1a4a", label_color="#7a3a9a", sub_color="#b890c8")
            cards_html += kpi_card("iKPI Value", fmt_num(r["iKPI Value"], 1), "iKPI [Valoare]",
                                   accent="#c0485a", bg="#fce4ec", val_color="#5c0a1a", label_color="#a0283a", sub_color="#d47080")
            cards_html += kpi_card("Avg Processing Time", fmt_num(r["Avg Processing Time (days)"], 1, "", " days"), "Business days",
                                   accent="#6a994e", bg="#f0f4e8", val_color="#2c4a1e", label_color="#4d7a2a", sub_color="#8ab86a")
            cards_html += '</div>'
            st.markdown(cards_html, unsafe_allow_html=True)
    else:
        # Show summary table for all VAS
        st.dataframe(vas_summary.set_index("VAS"), use_container_width=True)

    # ── Bar charts: Revenue / GM / iKPI per VAS ──
    col1, col2 = st.columns(2)
    with col1:
        fig2 = px.bar(vas_summary, x="VAS", y="Revenue (MEuro)",
                      title="Revenue by VAS (MEuro)", color="VAS",
                      color_discrete_sequence=PALETTE,
                      text=vas_summary["Revenue (MEuro)"].round(2))
        fig2.update_traces(textposition="outside")
        st.plotly_chart(styled_bar(fig2), use_container_width=True)

    with col2:
        fig3 = px.bar(vas_summary, x="VAS", y="GM (MEuro)",
                      title="GM by VAS (MEuro)", color="VAS",
                      color_discrete_sequence=PALETTE,
                      text=vas_summary["GM (MEuro)"].round(2))
        fig3.update_traces(textposition="outside")
        st.plotly_chart(styled_bar(fig3), use_container_width=True)

    col3, col4 = st.columns(2)
    with col3:
        fig4 = px.bar(vas_summary, x="VAS", y="Avg GM %",
                      title="Average GM % by VAS", color="VAS",
                      color_discrete_sequence=PALETTE,
                      text=vas_summary["Avg GM %"].round(1))
        fig4.update_traces(textposition="outside")
        st.plotly_chart(styled_bar(fig4), use_container_width=True)

    with col4:
        fig5 = px.bar(vas_summary, x="VAS", y="Avg Processing Time (days)",
                      title="Avg Processing Time by VAS (days)", color="VAS",
                      color_discrete_sequence=PALETTE,
                      text=vas_summary["Avg Processing Time (days)"].round(1))
        fig5.update_traces(textposition="outside")
        st.plotly_chart(styled_bar(fig5), use_container_width=True)

    col5, col6 = st.columns(2)
    with col5:
        fig6 = px.bar(vas_summary, x="VAS", y="iKPI Value",
                      title="iKPI [Valoare] by VAS", color="VAS",
                      color_discrete_sequence=PALETTE,
                      text=vas_summary["iKPI Value"].round(1))
        fig6.update_traces(textposition="outside")
        st.plotly_chart(styled_bar(fig6), use_container_width=True)

    with col6:
        fig7 = px.bar(vas_summary, x="VAS", y="Total Oferte",
                      title="Total Oferte by VAS", color="VAS",
                      color_discrete_sequence=PALETTE,
                      text=vas_summary["Total Oferte"].astype(int))
        fig7.update_traces(textposition="outside")
        st.plotly_chart(styled_bar(fig7), use_container_width=True)


# ─────────────────────────────────────────────
# SECTION 4 — OFFER STATUS
# ─────────────────────────────────────────────
def section_offer_status(df: pd.DataFrame):
    st.markdown('<div class="section-header">04 — Offer Status</div>', unsafe_allow_html=True)

    if "Status Oferta" not in df.columns:
        st.info("Column 'Status Oferta' not found.")
        return

    status_counts = df["Status Oferta"].str.strip().value_counts().reset_index()
    status_counts.columns = ["Status", "Count"]
    # Only show statuses with count > 0
    status_counts = status_counts[status_counts["Count"] > 0]

    col1, col2 = st.columns(2)
    with col1:
        fig = px.bar(status_counts, x="Status", y="Count", color="Status",
                     title="Offers by Status", color_discrete_sequence=PALETTE,
                     text="Count")
        fig.update_traces(textposition="outside")
        st.plotly_chart(styled_bar(fig), use_container_width=True)

    with col2:
        fig2 = px.pie(status_counts, names="Status", values="Count",
                      title="Status Distribution", color_discrete_sequence=PALETTE, hole=0.35)
        fig2.update_layout(**PLOTLY_THEME)
        st.plotly_chart(fig2, use_container_width=True)


# ─────────────────────────────────────────────
# SECTION 5 — PRODUCT PERFORMANCE
# ─────────────────────────────────────────────
def section_product_performance(df: pd.DataFrame, label: str = ""):
    header = f'<div class="section-header">05 — Product Performance {label}</div>'
    st.markdown(header, unsafe_allow_html=True)

    work = apply_business_rules(df)

    if work.empty:
        st.info("No data after applying business rules.")
        return

    group_col = "VAS" if "VAS" in work.columns else None
    if group_col is None:
        st.info("VAS column not found.")
        return

    agg = work.groupby(group_col).agg(
        Revenue=("Revenues [MEuro]", "sum"),
        GM=("GM [MEuro]", "sum"),
        GM_pct=("GM %", "mean"),
        iKPI=("iKPI [Valoare]", "sum"),
    ).reset_index()
    agg["GM_pct"] = agg["GM_pct"] * 100

    col1, col2 = st.columns(2)
    with col1:
        fig = px.bar(agg, x=group_col, y="Revenue",
                     title="Revenue by VAS (MEuro)", color=group_col,
                     color_discrete_sequence=PALETTE, text=agg["Revenue"].round(2))
        fig.update_traces(textposition="outside")
        st.plotly_chart(styled_bar(fig), use_container_width=True)

    with col2:
        fig2 = px.bar(agg, x=group_col, y="GM",
                      title="GM by VAS (MEuro)", color=group_col,
                      color_discrete_sequence=PALETTE, text=agg["GM"].round(2))
        fig2.update_traces(textposition="outside")
        st.plotly_chart(styled_bar(fig2), use_container_width=True)

    col3, col4 = st.columns(2)
    with col3:
        fig3 = px.bar(agg, x=group_col, y="GM_pct",
                      title="Average GM % by VAS", color=group_col,
                      color_discrete_sequence=PALETTE, text=agg["GM_pct"].round(1))
        fig3.update_traces(textposition="outside")
        st.plotly_chart(styled_bar(fig3), use_container_width=True)

    with col4:
        fig4 = px.bar(agg, x=group_col, y="iKPI",
                      title="iKPI [Valoare] by VAS", color=group_col,
                      color_discrete_sequence=PALETTE, text=agg["iKPI"].round(1))
        fig4.update_traces(textposition="outside")
        st.plotly_chart(styled_bar(fig4), use_container_width=True)


# ─────────────────────────────────────────────
# SECTION 6 — SIGNED CONTRACT PERFORMANCE
# ─────────────────────────────────────────────
def section_signed_contracts(df: pd.DataFrame, label: str = ""):
    header = f'<div class="section-header">06 — Signed Contract Performance {label}</div>'
    st.markdown(header, unsafe_allow_html=True)

    if "Status Oferta" not in df.columns:
        st.info("Status Oferta column not found.")
        return

    signed_df = df[df["Status Oferta"].str.strip() == "Signed"]
    if signed_df.empty:
        st.info("No signed contracts in this period.")
        return

    # Apply business rules to signed subset
    signed_clean = apply_business_rules(signed_df)

    total_signed_all = len(df[df["Status Oferta"].str.strip() == "Signed"])
    total_signed_period = len(signed_clean)
    total_rev = signed_clean["Revenues [MEuro]"].sum() if "Revenues [MEuro]" in signed_clean.columns else 0
    total_gm = signed_clean["GM [MEuro]"].sum() if "GM [MEuro]" in signed_clean.columns else 0
    avg_gm = signed_clean["GM %"].mean() * 100 if "GM %" in signed_clean.columns else 0
    total_ikpi = signed_clean["iKPI [Valoare]"].sum() if "iKPI [Valoare]" in signed_clean.columns else 0
    # iKPI/proiect is a unit label (MWp, MWh, LP...) — show most common
    ikpi_unit = signed_clean["iKPI/proiect"].mode()[0] if "iKPI/proiect" in signed_clean.columns and not signed_clean["iKPI/proiect"].dropna().empty else "—"

    cards_html = '<div class="kpi-grid">'
    cards_html += kpi_card("Total Signed (All)", fmt_num(total_signed_all, 0), "All time in dataset",
                           accent="#52b788", bg="#1b4332", val_color="#d8f3dc", label_color="#95d5b2", sub_color="#52b788")
    cards_html += kpi_card("Signed (Period)", fmt_num(total_signed_period, 0), "After business rules",
                           accent="#2b9d8f", bg="#e0f2f1", val_color="#0d3b36", label_color="#1a7a6e", sub_color="#4dbfb3")
    cards_html += kpi_card("Signed Revenue", fmt_num(total_rev, 2, "€", "M"), "MEuro",
                           accent="#e07c24", bg="#fff3e0", val_color="#7b3f00", label_color="#c25e00", sub_color="#e09a50")
    cards_html += kpi_card("Signed GM", fmt_num(total_gm, 2, "€", "M"), "MEuro",
                           accent="#4a90d9", bg="#e8f0fe", val_color="#1a3a6b", label_color="#2c5fa8", sub_color="#6aa3e0")
    cards_html += kpi_card("Avg GM %", fmt_num(avg_gm, 1, "", "%"), "Arithmetic avg",
                           accent="#c0485a", bg="#fce4ec", val_color="#5c0a1a", label_color="#a0283a", sub_color="#d47080")
    cards_html += kpi_card("iKPI Value", fmt_num(total_ikpi, 2), "iKPI [Valoare]",
                           accent="#6a994e", bg="#f0f4e8", val_color="#2c4a1e", label_color="#4d7a2a", sub_color="#8ab86a")
    cards_html += kpi_card("iKPI Unit", str(ikpi_unit), "Unitate iKPI/proiect",
                           accent="#9b6fa8", bg="#f3e8f8", val_color="#3b1a4a", label_color="#7a3a9a", sub_color="#b890c8")
    cards_html += '</div>'
    st.markdown(cards_html, unsafe_allow_html=True)

    if not signed_clean.empty:
        col1, col2 = st.columns(2)
        with col1:
            if "VAS" in signed_clean.columns:
                vas_rev = signed_clean.groupby("VAS")["Revenues [MEuro]"].sum().reset_index()
                fig = px.bar(vas_rev, x="VAS", y="Revenues [MEuro]",
                             title="Signed Revenue by VAS",
                             color="VAS", color_discrete_sequence=PALETTE,
                             text=vas_rev["Revenues [MEuro]"].round(2))
                fig.update_traces(textposition="outside")
                st.plotly_chart(styled_bar(fig), use_container_width=True)

        with col2:
            if "Tip oferta" in signed_clean.columns:
                tip_signed = signed_clean["Tip oferta"].str.strip().value_counts().reset_index()
                tip_signed.columns = ["Tip oferta", "Count"]
                fig2 = px.pie(tip_signed, names="Tip oferta", values="Count",
                              title="Signed by Product Type",
                              color_discrete_sequence=PALETTE, hole=0.35)
                fig2.update_layout(**PLOTLY_THEME)
                st.plotly_chart(fig2, use_container_width=True)


# ─────────────────────────────────────────────
# SECTION 7 — ENGINEER PERFORMANCE
# ─────────────────────────────────────────────
def section_engineer_performance(df: pd.DataFrame):
    st.markdown('<div class="section-header">07 — Proposal Engineer Performance</div>', unsafe_allow_html=True)

    eng_col = "Inginer Ofertare"
    if eng_col not in df.columns:
        st.info("Column 'Inginer Ofertare' not found.")
        return

    df = df.copy()
    # Ensure engineer names are fully stripped (catches any remaining whitespace)
    df[eng_col] = df[eng_col].astype(str).str.strip()
    df = df[df[eng_col].notna() & (df[eng_col] != "nan") & (df[eng_col] != "")]
    df["_proc_time"] = calc_processing_time(df)

    # Pre-compute signed mask on the period-filtered df
    if "Status Oferta" in df.columns:
        signed_mask = df["Status Oferta"].str.strip().str.lower() == "signed"
    else:
        signed_mask = pd.Series(False, index=df.index)

    eng_stats = []
    for eng, grp in df.groupby(eng_col):
        bugetara  = (grp["Tip oferta"].str.strip() == "Bugetara").sum()  if "Tip oferta" in grp.columns else 0
        angajanta = (grp["Tip oferta"].str.strip() == "Angajanta").sum() if "Tip oferta" in grp.columns else 0
        revizie   = (grp["Tip oferta"].str.strip() == "Revizie").sum()   if "Tip oferta" in grp.columns else 0
        total_rev = grp["Revenues [MEuro]"].sum()  if "Revenues [MEuro]" in grp.columns else 0
        total_gm  = grp["GM [MEuro]"].sum()        if "GM [MEuro]" in grp.columns else 0
        avg_gm    = grp["GM %"].mean() * 100       if "GM %" in grp.columns else 0
        avg_pt    = grp["_proc_time"].mean()
        signed    = int(signed_mask.loc[grp.index].sum())
        ikpi_val  = grp["iKPI [Valoare]"].sum()    if "iKPI [Valoare]" in grp.columns else 0

        eng_stats.append({
            "Engineer":                  eng,
            "Bugetara":                  int(bugetara),
            "Angajanta":                 int(angajanta),
            "Revizie":                   int(revizie),
            "Total Offers":              int(bugetara + angajanta + revizie),
            "Total Revenue (MEuro)":     round(total_rev, 2),
            "Total GM (MEuro)":          round(total_gm, 2),
            "Avg GM %":                  round(avg_gm, 1),
            "iKPI [Valoare]":            round(ikpi_val, 2),
            "Avg Processing Time (days)": round(avg_pt, 1) if pd.notna(avg_pt) else 0,
            "Signed Contracts":          signed,
        })

    eng_df = pd.DataFrame(eng_stats).sort_values("Total Offers", ascending=False)

    # Display table
    st.dataframe(eng_df.set_index("Engineer"), use_container_width=True)

    col1, col2 = st.columns(2)
    with col1:
        fig = px.bar(eng_df, x="Engineer", y="Avg Processing Time (days)",
                     title="Avg Offer Delivery Time (Business Days)",
                     color="Engineer", color_discrete_sequence=PALETTE,
                     text="Avg Processing Time (days)")
        fig.update_traces(textposition="outside")
        st.plotly_chart(styled_bar(fig), use_container_width=True)

    with col2:
        melt_df = eng_df.melt(id_vars="Engineer",
                              value_vars=["Bugetara", "Angajanta", "Revizie"],
                              var_name="Type", value_name="Count")
        fig2 = px.bar(melt_df, x="Engineer", y="Count", color="Type",
                      barmode="stack", title="Offers Submitted per Type per Engineer",
                      color_discrete_sequence=PALETTE)
        st.plotly_chart(styled_bar(fig2), use_container_width=True)

    col3, col4 = st.columns(2)
    with col3:
        fig3 = px.bar(eng_df, x="Engineer", y="Signed Contracts",
                      title="Signed Contracts per Engineer",
                      color="Engineer", color_discrete_sequence=PALETTE,
                      text="Signed Contracts")
        fig3.update_traces(textposition="outside")
        st.plotly_chart(styled_bar(fig3), use_container_width=True)

    with col4:
        fig4 = px.bar(eng_df, x="Engineer", y="Total Revenue (MEuro)",
                      title="Total Revenue per Engineer (MEuro)",
                      color="Engineer", color_discrete_sequence=PALETTE,
                      text=eng_df["Total Revenue (MEuro)"].round(2))
        fig4.update_traces(textposition="outside")
        st.plotly_chart(styled_bar(fig4), use_container_width=True)


# ─────────────────────────────────────────────
# SECTION 8 — YEAR / MONTH COMPARISON
# ─────────────────────────────────────────────
def section_comparison(df_full: pd.DataFrame):
    st.markdown('<div class="section-header">08 — Year / Month Comparison</div>', unsafe_allow_html=True)

    available_years = sorted(df_full["_year"].dropna().unique().astype(int).tolist())
    if len(available_years) < 1:
        st.info("Not enough data for comparison.")
        return

    st.markdown("##### Select periods to compare")
    col_a, col_b = st.columns(2)
    with col_a:
        year_a = st.selectbox("Period A — Year", available_years, key="cmp_year_a")
        month_a = st.selectbox("Period A — Month (optional)", [None] + list(range(1, 13)),
                               format_func=lambda x: "All months" if x is None else datetime(2000, x, 1).strftime("%B"),
                               key="cmp_month_a")
    with col_b:
        year_b = st.selectbox("Period B — Year", available_years, index=min(1, len(available_years)-1), key="cmp_year_b")
        month_b = st.selectbox("Period B — Month (optional)", [None] + list(range(1, 13)),
                               format_func=lambda x: "All months" if x is None else datetime(2000, x, 1).strftime("%B"),
                               key="cmp_month_b")

    df_a = filter_df(df_full, year_a, month_a)
    df_b = filter_df(df_full, year_b, month_b)
    a_label = f"{year_a}" + (f"-{month_a:02d}" if month_a else "")
    b_label = f"{year_b}" + (f"-{month_b:02d}" if month_b else "")

    work_a = apply_business_rules(df_a)
    work_b = apply_business_rules(df_b)

    def summarise(work, label):
        rev = work["Revenues [MEuro]"].sum() if "Revenues [MEuro]" in work.columns else 0
        gm = work["GM [MEuro]"].sum() if "GM [MEuro]" in work.columns else 0
        gm_pct = work["GM %"].mean() * 100 if "GM %" in work.columns else 0
        ikpi = work["iKPI [Valoare]"].sum() if "iKPI [Valoare]" in work.columns else 0
        signed_rev = (work[work["Status Oferta"].str.strip() == "Signed"]["Revenues [MEuro]"].sum()
                      if "Status Oferta" in work.columns and "Revenues [MEuro]" in work.columns else 0)
        return {"Period": label, "Revenue (MEuro)": rev, "GM (MEuro)": gm,
                "GM %": gm_pct, "iKPI": ikpi, "Signed Revenue": signed_rev,
                "Offers": len(work)}

    summ_a = summarise(work_a, a_label)
    summ_b = summarise(work_b, b_label)
    cmp_df = pd.DataFrame([summ_a, summ_b])

    metrics = ["Revenue (MEuro)", "GM (MEuro)", "GM %", "iKPI", "Signed Revenue", "Offers"]
    col1, col2, col3 = st.columns(3)
    cols_cycle = [col1, col2, col3]
    for i, metric in enumerate(metrics):
        with cols_cycle[i % 3]:
            fig = go.Figure(go.Bar(
                x=cmp_df["Period"], y=cmp_df[metric],
                marker_color=[PALETTE[0], PALETTE[2]],
                text=cmp_df[metric].round(2), textposition="outside"
            ))
            fig.update_layout(title=metric, **PLOTLY_THEME, height=300)
            st.plotly_chart(fig, use_container_width=True)

    # VAS-level comparison
    if "VAS" in work_a.columns and "VAS" in work_b.columns:
        st.markdown("##### Revenue by VAS — Comparison")
        vas_a = work_a.groupby("VAS")["Revenues [MEuro]"].sum().rename(a_label)
        vas_b = work_b.groupby("VAS")["Revenues [MEuro]"].sum().rename(b_label)
        vas_cmp = pd.concat([vas_a, vas_b], axis=1).fillna(0).reset_index()
        fig_vas = px.bar(vas_cmp, x="VAS", y=[a_label, b_label],
                         barmode="group", title="Revenue by VAS — Period Comparison",
                         color_discrete_sequence=[PALETTE[0], PALETTE[2]])
        st.plotly_chart(styled_bar(fig_vas), use_container_width=True)


# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
def render_sidebar(df_full: pd.DataFrame):
    st.sidebar.markdown("""
    <div style='padding: 12px 0 20px; border-bottom: 1px solid #a5d6a7; margin-bottom: 16px;'>
        <div style='font-size:1.1rem; font-weight:700; color:#1b5e20; letter-spacing:0.03em;'>Bidding &amp; Design Dashboard</div>
        <div style='font-size:0.72rem; color:#558b2f; margin-top:4px;'>Offer Analytics Platform</div>
    </div>
    """, unsafe_allow_html=True)

    available_years = sorted(df_full["_year"].dropna().unique().astype(int).tolist())
    if not available_years:
        st.sidebar.warning("No year data found.")
        return None, None

    selected_year = st.sidebar.selectbox("📅 Year (required)", available_years, index=len(available_years)-1)

    months_in_year = sorted(df_full[df_full["_year"] == selected_year]["_month"].dropna().unique().astype(int).tolist())
    month_options = [None] + months_in_year
    selected_month = st.sidebar.selectbox(
        "📆 Month (optional)",
        month_options,
        format_func=lambda x: "All Months" if x is None else datetime(2000, x, 1).strftime("%B")
    )

    st.sidebar.markdown("---")
    st.sidebar.markdown(f"""
    <div style='font-size:0.75rem; color:#558b2f;'>
        <b style='color:#2e7d32'>Active filter:</b><br>
        Year: <span style='color:#1b5e20; font-weight:600'>{selected_year}</span><br>
        Month: <span style='color:#1b5e20; font-weight:600'>{'All' if selected_month is None else datetime(2000, selected_month, 1).strftime('%B')}</span>
    </div>
    """, unsafe_allow_html=True)

    return selected_year, selected_month


# ─────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────
def main():
    # ── Header ──────────────────────────────
    st.markdown("""
    <div style='
        background: linear-gradient(90deg, #1b4332 0%, #2d6a4f 60%, #40916c 100%);
        border-radius: 12px;
        padding: 20px 28px 18px;
        margin-bottom: 24px;
        display: flex; align-items: center; gap: 14px;
    '>
        <div>
            <div style='font-size:1.4rem; font-weight:700; color:#d8f3dc; letter-spacing:-0.02em;'>Bidding &amp; Design Dashboard</div>
            <div style='font-size:0.8rem; color:#95d5b2; margin-top:3px;'>Upload your Excel file to generate interactive pipeline analytics</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── File upload ──────────────────────────
    uploaded = st.file_uploader("Upload Excel File", type=["xlsx", "xls"],
                                label_visibility="collapsed")

    if uploaded is None:
        st.markdown("""
        <div style='
            background: #f1f8e9; border: 1px dashed #a5d6a7; border-radius: 10px;
            padding: 40px; text-align: center; color: #558b2f;
        '>
            <div style='font-size: 1rem; font-weight: 600; color: #2e7d32;'>Drop your Excel file above</div>
            <div style='font-size: 0.8rem; margin-top: 8px; color: #558b2f;'>
                Expected columns: Nr. Oferta, Client name, VAS, Revenues [MEuro], GM [MEuro], GM %, Status Oferta, Tip oferta, etc.
            </div>
        </div>
        """, unsafe_allow_html=True)
        return

    with st.spinner("Loading data..."):
        df_full = load_data(uploaded.read())

    if "_year" not in df_full.columns:
        st.error("Nu s-au putut parsa coloanele de dată. Coloane găsite în fișier:")
        st.code(", ".join(df_full.columns.tolist()))
        st.stop()
        return

    # ── Sidebar ──────────────────────────────
    selected_year, selected_month = render_sidebar(df_full)
    if selected_year is None:
        return

    df = filter_df(df_full, selected_year, selected_month)

    if df.empty:
        st.warning("No data found for the selected period.")
        return

    # ── Navigation Tabs ──────────────────────
    tabs = st.tabs([
        "KPIs",
        "Offer Activity",
        "Product Type",
        "Status",
        "Performance",
        "Signed",
        "Engineers",
        "Comparison",
    ])

    with tabs[0]:
        section_pipeline_kpis(df)

    with tabs[1]:
        section_offer_activity(df)

    with tabs[2]:
        section_product_type(df)

    with tabs[3]:
        section_offer_status(df)

    with tabs[4]:
        section_product_performance(df)

    with tabs[5]:
        section_signed_contracts(df)

    with tabs[6]:
        section_engineer_performance(df)

    with tabs[7]:
        section_comparison(df_full)

    # ── Raw Data Expander ─────────────────────
    with st.expander("Explore Raw Data"):
        st.dataframe(df, use_container_width=True)
        csv = df.to_csv(index=False).encode("utf-8")
        st.download_button("Download filtered CSV", csv, "filtered_data.csv", "text/csv")


if __name__ == "__main__":
    main()
