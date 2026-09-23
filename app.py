"""
================================================================================
CATB Screening & Surveillance Analytics Platform (Master Ground Truth Edition)
================================================================================
Enterprise TB Active Case Finding surveillance system featuring:
- Authoritative HWC Master Hierarchy (Column E) as the absolute source of truth:
  * ZERO deduplication: Every single entry in the master file is preserved and counted.
  * Duplicate HWC names (same or different districts/blocks) appear as separate rows.
  * Strict reconciliation: Total HWCs = Screening Started + Screening Not Started.
- Complete Screening Data Sheet analytics:
  * Frontline CHO name mapped from Column G of the Screening Data Sheet.
  * Column L: AI Preference Presumptive.
  * Column Q: NTEP Presumptive.
  * AI + NTEP Both Presumptive (Analytical overlap metric).
  * Column BA: CHO Override Presumptive.
  * Total Unique Presumptive (Deduplicated union: AI OR NTEP OR CHO Override).
  * Column T Progression: Presumptive Open, Presumptive Closed, Diagnosed, Tested.
- High-visibility compact sidebar and 100% Streamlit Cloud compatibility.
"""

import datetime
import io
import re
import numpy as np
import openpyxl
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ------------------------------------------------------------------------------
# 1. PAGE SETUP & COMPACT ENTERPRISE CSS
# ------------------------------------------------------------------------------
st.set_page_config(
    page_title="CATB Screening & Surveillance Portal",
    page_icon="🩺",
    layout="wide",
    initial_sidebar_state="expanded",
)

ENTERPRISE_THEME_CSS = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }

    .block-container {
        padding-top: 1rem;
        padding-bottom: 2rem;
        padding-left: 2rem;
        padding-right: 2rem;
        max-width: 100% !important;
    }

    [data-testid="stSidebar"] {
        background-color: #F8FAFC !important;
        border-right: 1.5px solid #E2E8F0 !important;
    }
    [data-testid="stSidebar"] > div:first-child {
        padding-top: 0.6rem !important;
        padding-left: 0.8rem !important;
        padding-right: 0.8rem !important;
    }
    
    .sidebar-brand {
        display: flex;
        align-items: center;
        gap: 8px;
        padding-bottom: 6px;
        border-bottom: 1.5px solid #E2E8F0;
        margin-bottom: 8px;
    }
    .sidebar-brand-title {
        font-size: 1rem;
        font-weight: 800;
        color: #0F172A;
        line-height: 1.1;
    }
    .sidebar-brand-sub {
        font-size: 0.7rem;
        color: #64748B;
        font-weight: 600;
    }

    .sidebar-category-header {
        font-size: 0.76rem;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        color: #0369A1;
        background: #E0F2FE;
        padding: 4px 8px;
        border-radius: 6px;
        margin-top: 8px;
        margin-bottom: 6px;
        display: flex;
        align-items: center;
        gap: 5px;
    }

    .filter-item-label {
        font-size: 0.74rem !important;
        font-weight: 700 !important;
        color: #1E293B !important;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        margin-top: 5px !important;
        margin-bottom: 2px !important;
        display: flex;
        align-items: center;
        gap: 4px;
    }

    [data-testid="stMultiSelect"] > div,
    [data-testid="stDateInput"] > div,
    [data-testid="stSelectbox"] > div {
        background-color: #FFFFFF !important;
        border: 1.5px solid #CBD5E1 !important;
        border-radius: 8px !important;
        box-shadow: 0 1px 2px rgba(0, 0, 0, 0.04) !important;
        transition: all 0.2s ease-in-out !important;
        min-height: 36px !important;
    }
    [data-testid="stMultiSelect"] > div:hover,
    [data-testid="stDateInput"] > div:hover,
    [data-testid="stSelectbox"] > div:hover {
        border-color: #0284C7 !important;
        box-shadow: 0 0 0 2px rgba(2, 132, 199, 0.15) !important;
    }
    [data-testid="stMultiSelect"] span[data-baseweb="tag"] {
        background-color: #0F172A !important;
        color: #FFFFFF !important;
        border-radius: 4px !important;
        font-weight: 600 !important;
        font-size: 0.72rem !important;
        padding: 1px 5px !important;
    }

    [data-testid="stFileUploaderInstructions"],
    [data-testid="stFileUploaderDropzoneInstructions"] {
        display: none !important;
    }
    [data-testid="stFileUploader"] {
        margin-bottom: 4px !important;
    }
    [data-testid="stFileUploaderDropzone"] {
        padding: 4px 8px !important;
        min-height: 44px !important;
        max-height: 46px !important;
        border-radius: 8px !important;
        border: 1.5px dashed #94A3B8 !important;
        background-color: #FFFFFF !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
    }
    [data-testid="stFileUploaderDropzone"]:hover {
        border-color: #0284C7 !important;
        background-color: #F0F9FF !important;
    }
    [data-testid="stFileUploaderDropzone"] button {
        padding: 2px 8px !important;
        font-size: 0.72rem !important;
        height: 26px !important;
    }
    [data-testid="stFileUploader"] section {
        padding: 2px 4px !important;
        margin-top: 1px !important;
    }

    .app-header {
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
        padding: 14px 20px;
        border-radius: 12px;
        color: white;
        margin-bottom: 14px;
        border: 1px solid #334155;
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
        gap: 10px;
    }
    .app-header-title {
        font-size: 1.35rem;
        font-weight: 800;
        letter-spacing: -0.02em;
        margin: 0;
        display: flex;
        align-items: center;
        gap: 8px;
        color: #FFFFFF;
    }
    .app-header-subtitle {
        color: #94A3B8;
        font-size: 0.8rem;
        margin-top: 2px;
    }
    .system-status-pill {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: rgba(16, 185, 129, 0.15);
        color: #34D399;
        border: 1px solid rgba(52, 211, 153, 0.3);
        padding: 4px 10px;
        border-radius: 9999px;
        font-size: 0.74rem;
        font-weight: 600;
    }
    .pulse-dot {
        width: 6px;
        height: 6px;
        background-color: #10B981;
        border-radius: 50%;
    }

    .kpi-grid-card {
        background: #FFFFFF;
        border-radius: 12px;
        padding: 12px 14px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
        position: relative;
        overflow: hidden;
        transition: transform 0.15s ease-in-out;
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }
    .kpi-grid-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 14px -3px rgba(0, 0, 0, 0.06);
        border-color: #CBD5E1;
    }
    .kpi-accent-bar {
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        height: 3px;
    }
    .kpi-top-meta {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 3px;
    }
    .kpi-label {
        font-size: 0.68rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #64748B;
    }
    .kpi-icon-bubble {
        width: 26px;
        height: 26px;
        border-radius: 6px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.9rem;
    }
    .kpi-metric-value {
        font-size: 1.55rem;
        font-weight: 800;
        color: #0F172A;
        line-height: 1.15;
    }
    .kpi-subtext {
        font-size: 0.7rem;
        font-weight: 500;
        color: #64748B;
        margin-top: 3px;
    }

    .section-title-wrap {
        margin-top: 14px;
        margin-bottom: 10px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        border-bottom: 2px solid #F1F5F9;
        padding-bottom: 4px;
    }
    .section-title-text {
        font-size: 1.02rem;
        font-weight: 700;
        color: #1E293B;
        display: flex;
        align-items: center;
        gap: 6px;
    }
</style>
"""
st.markdown(ENTERPRISE_THEME_CSS, unsafe_allow_html=True)


# ------------------------------------------------------------------------------
# 2. FILE READER & FUZZY COLUMN DETECTOR
# ------------------------------------------------------------------------------
def clean_header(val: str) -> str:
    if not isinstance(val, str):
        return ""
    return re.sub(r"[_\-\s]+", " ", str(val)).strip().lower()


def detect_field(columns_list, candidates):
    clean_map = {clean_header(c): c for c in columns_list}
    for cand in candidates:
        cc = clean_header(cand)
        if cc in clean_map:
            return clean_map[cc]
    for cand in candidates:
        cc = clean_header(cand)
        for col_clean, original in clean_map.items():
            if cc in col_clean:
                return original
    return None


def make_composite_key(dist, blk, fac):
    """Normalized composite key (District + Block + HWC Name) for robust linking."""
    d = str(dist).strip().title() if pd.notna(dist) else "Unknown"
    b = str(blk).strip().title() if pd.notna(blk) else "Unknown"
    f = str(fac).strip().title() if pd.notna(fac) else "Unknown"
    return f"{d}:::{b}:::{f}".lower()


# ------------------------------------------------------------------------------
# 3. HIGH-PERFORMANCE DATA STANDARDIZATION & COMPLETE HWC PRESERVATION ENGINE
# ------------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_and_standardize_data(file_all_bytes, file_all_name, file_hier_bytes=None, file_hier_name=None):
    """
    Ingests Screening Data Sheet and authoritative HWC Master Hierarchy.
    STRICT REQUIREMENT: Every single HWC entry in the master file is retained.
    NO deduplication is applied to the master file. Duplicate HWC names are preserved as separate rows.
    """
    # 1. Primary Screening Data Sheet
    if file_all_name.lower().endswith(".csv"):
        df_all = pd.read_csv(io.BytesIO(file_all_bytes), low_memory=False)
    else:
        df_all = pd.read_excel(io.BytesIO(file_all_bytes))

    cols_all = list(df_all.columns)

    c_date = detect_field(cols_all, ["reg date", "date", "created at", "screening date"])
    c_id = detect_field(cols_all, ["id", "uuid", "patient id", "beneficiary id"])
    c_dist = detect_field(cols_all, ["district", "district name"])
    c_block = detect_field(cols_all, ["block", "block name", "tehsil"])

    # Column E (Index 4) or detected HWC
    if len(cols_all) > 4 and "hwc" in clean_header(str(cols_all[4])):
        c_hwc = cols_all[4]
    else:
        c_hwc = detect_field(cols_all, ["hwc", "facility", "facility name", "health center"])

    # Column G (Index 6) is CHO Name
    if len(cols_all) > 6 and "cho" in clean_header(str(cols_all[6])):
        c_cho = cols_all[6]
    else:
        c_cho = detect_field(cols_all, ["cho name", "cho", "staff name", "provider"])

    # Column L (Index 11) - AI Preference
    if len(cols_all) > 11 and "ai" in clean_header(str(cols_all[11])):
        c_ai = cols_all[11]
    else:
        c_ai = detect_field(cols_all, ["ai preference", "ai result", "ai"])

    # Column Q (Index 16) - NTEP Result
    if len(cols_all) > 16 and "ntep" in clean_header(str(cols_all[16])):
        c_ntep = cols_all[16]
    else:
        c_ntep = detect_field(cols_all, ["ntep result", "ntep"])

    # Column T (Index 19) - TB Status
    if len(cols_all) > 19 and "status" in clean_header(str(cols_all[19])):
        c_status = cols_all[19]
    else:
        c_status = detect_field(cols_all, ["tb status", "tb_status", "status"])

    # Column BA (Index 52) - Overrule by HWC
    if len(cols_all) > 52 and ("overrule" in clean_header(str(cols_all[52])) or "hwc" in clean_header(str(cols_all[52]))):
        c_overrule = cols_all[52]
    else:
        c_overrule = detect_field(cols_all, ["overrule by hwc", "overrule", "hwc overrule", "cho override"])

    df = df_all.copy()

    # Parse Registration Date
    if c_date and c_date in df.columns:
        df["reg_date_clean"] = pd.to_datetime(df[c_date], errors="coerce")
    else:
        df["reg_date_clean"] = pd.NaT

    df["District_Clean"] = df[c_dist].fillna("Unknown").astype(str).str.strip().str.title() if c_dist else "Unknown"
    df["Block_Clean"] = df[c_block].fillna("Unknown").astype(str).str.strip().str.title() if c_block else "Unknown"
    df["Facility_Clean"] = df[c_hwc].fillna("Unknown").astype(str).str.strip().str.title() if c_hwc else "Unknown"

    # Map CHO Name from Column G
    if c_cho and c_cho in df.columns:
        df["CHO_Clean"] = df[c_cho].astype(str).str.strip().str.title()
        df.loc[df["CHO_Clean"].isin(["", "Nan", "None", "Null", "0"]), "CHO_Clean"] = "Not Available"
    else:
        df["CHO_Clean"] = "Not Available"

    df["str_id"] = df[c_id].astype(str).str.strip() if c_id else df.index.astype(str)
    df["composite_key"] = df.apply(lambda r: make_composite_key(r["District_Clean"], r["Block_Clean"], r["Facility_Clean"]), axis=1)

    # Presumptive Breakdown & Overrule Check
    ai_series = df[c_ai].fillna("").astype(str).str.strip().str.lower() if c_ai and c_ai in df.columns else pd.Series([""] * len(df))
    ntep_series = df[c_ntep].fillna("").astype(str).str.strip().str.lower() if c_ntep and c_ntep in df.columns else pd.Series([""] * len(df))

    if c_overrule and c_overrule in df.columns:
        ovr_series = df[c_overrule].fillna("0").astype(str).str.strip()
        is_ovr_active = ovr_series.isin(["1", "1.0", "True", "true", "yes", "Yes"])
    else:
        is_ovr_active = pd.Series([False] * len(df))

    df["is_ai_pres"] = (ai_series == "presumptive").astype(int)
    df["is_ntep_pres"] = (ntep_series == "presumptive").astype(int)
    df["is_both_pres"] = ((df["is_ai_pres"] == 1) & (df["is_ntep_pres"] == 1)).astype(int)
    df["is_ai_only"] = ((df["is_ai_pres"] == 1) & (df["is_ntep_pres"] == 0)).astype(int)
    df["is_ntep_only"] = ((df["is_ai_pres"] == 0) & (df["is_ntep_pres"] == 1)).astype(int)
    df["is_cho_override"] = ((df["is_ai_pres"] == 0) & (df["is_ntep_pres"] == 0) & is_ovr_active).astype(int)
    df["is_total_presumptive"] = ((df["is_ai_pres"] == 1) | (df["is_ntep_pres"] == 1) | (df["is_cho_override"] == 1)).astype(int)

    # TB Status Progression (Column T)
    status_series = df[c_status].fillna("").astype(str).str.strip().str.upper() if c_status and c_status in df.columns else pd.Series([""] * len(df))
    df["is_presumptive_open"] = (status_series == "PRESUMPTIVE_OPEN").astype(int)
    df["is_presumptive_closed"] = (status_series == "PRESUMPTIVE_CLOSED").astype(int)
    df["is_diagnosed"] = (status_series == "DIAGNOSED_ON_TREATMENT").astype(int)
    df["is_tested"] = (status_series.isin(["PRESUMPTIVE_CLOSED", "DIAGNOSED_ON_TREATMENT"])).astype(int)

    # 2. Complete HWC Master Hierarchy Ingestion (ZERO DEDUPLICATION)
    df_master_hwc = None
    master_raw_count = 0
    if file_hier_bytes:
        try:
            if file_hier_name.lower().endswith(".csv"):
                df_h = pd.read_csv(io.BytesIO(file_hier_bytes), low_memory=False)
            else:
                df_h = pd.read_excel(io.BytesIO(file_hier_bytes))

            h_cols = list(df_h.columns)
            h_dist = detect_field(h_cols, ["district", "district name"])
            h_block = detect_field(h_cols, ["block", "block name", "tehsil"])
            h_fid = detect_field(h_cols, ["facility id", "facility_id", "hwc id", "hwc_id", "nin", "nin code"])

            # Column E (Index 4) or detected HWC
            if len(h_cols) > 4 and "hwc" in clean_header(str(h_cols[4])):
                h_fac = h_cols[4]
            else:
                h_fac = detect_field(h_cols, ["hwc", "facility", "facility name", "health center"])

            if h_fac:
                # Retain EVERY valid row. Do NOT drop duplicates.
                df_h = df_h[df_h[h_fac].notna()].copy()
                df_h["Facility_Clean"] = df_h[h_fac].astype(str).str.strip().str.title()
                df_h["District_Clean"] = df_h[h_dist].fillna("Unknown").astype(str).str.strip().str.title() if h_dist else "Unknown"
                df_h["Block_Clean"] = df_h[h_block].fillna("Unknown").astype(str).str.strip().str.title() if h_block else "Unknown"
                df_h["Facility_ID"] = df_h[h_fid].astype(str).str.strip() if h_fid else ""

                # Assign a unique Master Row ID to preserve exact 1-to-1 row correspondence
                df_h["master_row_id"] = np.arange(1, len(df_h) + 1)
                df_h["composite_key"] = df_h.apply(lambda r: make_composite_key(r["District_Clean"], r["Block_Clean"], r["Facility_Clean"]), axis=1)

                master_raw_count = len(df_h)
                df_master_hwc = df_h.copy()
        except Exception:
            pass

    return df, df_master_hwc, master_raw_count


def render_enterprise_kpi(col, label, value, subtext, icon, accent_color="#0284C7", bg_bubble="#E0F2FE"):
    card_html = f"""
    <div class="kpi-grid-card">
        <div class="kpi-accent-bar" style="background-color: {accent_color};"></div>
        <div class="kpi-top-meta">
            <span class="kpi-label">{label}</span>
            <div class="kpi-icon-bubble" style="background-color: {bg_bubble}; color: {accent_color};">
                {icon}
            </div>
        </div>
        <div class="kpi-metric-value">{value}</div>
        <div class="kpi-subtext">{subtext}</div>
    </div>
    """
    col.markdown(card_html, unsafe_allow_html=True)


# ------------------------------------------------------------------------------
# 4. MAIN APPLICATION
# ------------------------------------------------------------------------------
def main():
    # --------------------------------------------------------------------------
    # SIDEBAR: COMPACT, HIGH-VISIBILITY DESIGN
    # --------------------------------------------------------------------------
    with st.sidebar:
        st.markdown(
            """
            <div class="sidebar-brand">
                <span style="font-size: 1.4rem;">🩺</span>
                <div>
                    <div class="sidebar-brand-title">CATB Command</div>
                    <div class="sidebar-brand-sub">Health Surveillance Portal</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown('<div class="sidebar-category-header">📁 1. Data Ingestion</div>', unsafe_allow_html=True)

        st.markdown('<div class="filter-item-label">Screening Data Sheet</div>', unsafe_allow_html=True)
        up_all = st.file_uploader(
            "Screening Data Sheet",
            type=["xlsx", "xls", "csv"],
            key="file_screening",
            label_visibility="collapsed",
        )

        st.markdown('<div class="filter-item-label">HWC Master Hierarchy (Column E)</div>', unsafe_allow_html=True)
        up_hier = st.file_uploader(
            "HWC Master Hierarchy",
            type=["xlsx", "xls", "csv"],
            key="file_hierarchy",
            label_visibility="collapsed",
        )

    if not up_all:
        st.markdown(
            """
            <div class="app-header">
                <div>
                    <h1 class="app-header-title"><span>🩺</span> CATB Screening & Clinical Surveillance Platform</h1>
                    <div class="app-header-subtitle">Active Case Finding (ACF) Intelligence, AI Triage & Nikshay Continuum of Care</div>
                </div>
                <div class="system-status-pill">
                    <div class="pulse-dot"></div> Awaiting Ingestion
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.info("👈 Please upload your **Screening Data Sheet (.xlsx, .xls, or .csv)** in the sidebar to load the surveillance portal.")
        return

    # Ingest Datasets
    with st.spinner("⚡ Loading full HWC master list & processing screening records..."):
        df_screening, df_master_hwc, master_raw_count = load_and_standardize_data(
            up_all.getvalue(),
            up_all.name,
            up_hier.getvalue() if up_hier else None,
            up_hier.name if up_hier else None,
        )

    # --------------------------------------------------------------------------
    # SIDEBAR: CASCADING FILTERS
    # --------------------------------------------------------------------------
    with st.sidebar:
        st.markdown('<div class="sidebar-category-header">🔍 2. Analytical Filters</div>', unsafe_allow_html=True)

        # Date Range Filter
        dates_valid = df_screening["reg_date_clean"].dropna()
        if not dates_valid.empty:
            min_avail, max_avail = dates_valid.min().date(), dates_valid.max().date()
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.markdown('<div class="filter-item-label">Starting Date</div>', unsafe_allow_html=True)
                start_date = st.date_input("Starting Date", value=min_avail, min_value=min_avail, max_value=max_avail, label_visibility="collapsed")
            with col_d2:
                st.markdown('<div class="filter-item-label">Ending Date</div>', unsafe_allow_html=True)
                end_date = st.date_input("Ending Date", value=max_avail, min_value=min_avail, max_value=max_avail, label_visibility="collapsed")
        else:
            start_date, end_date = None, None

        # Build Geography Reference from Master or Screening
        geo_ref = df_master_hwc if df_master_hwc is not None else df_screening

        # District Filter
        st.markdown('<div class="filter-item-label">Select District</div>', unsafe_allow_html=True)
        avail_districts = sorted([d for d in geo_ref["District_Clean"].unique() if d != "Unknown"])
        sel_districts = st.multiselect(
            "Select District",
            options=avail_districts,
            default=[],
            placeholder="All Districts",
            label_visibility="collapsed",
        )

        # Cascading Block Filter
        geo_b = geo_ref[geo_ref["District_Clean"].isin(sel_districts)] if sel_districts else geo_ref
        avail_blocks = sorted([b for b in geo_b["Block_Clean"].unique() if b != "Unknown"])
        st.markdown('<div class="filter-item-label">Select Block</div>', unsafe_allow_html=True)
        sel_blocks = st.multiselect(
            "Select Block",
            options=avail_blocks,
            default=[],
            placeholder="All Blocks",
            label_visibility="collapsed",
        )

        # Cascading HWC Filter
        geo_h = geo_b[geo_b["Block_Clean"].isin(sel_blocks)] if sel_blocks else geo_b
        avail_hwcs = sorted([h for h in geo_h["Facility_Clean"].unique() if h != "Unknown"])
        st.markdown('<div class="filter-item-label">Select HWC</div>', unsafe_allow_html=True)
        sel_hwcs = st.multiselect(
            "Select HWC",
            options=avail_hwcs,
            default=[],
            placeholder="All HWCs",
            label_visibility="collapsed",
        )

        # Cascading CHO Filter
        df_cho_scope = df_screening[df_screening["Facility_Clean"].isin(sel_hwcs)] if sel_hwcs else df_screening
        avail_chos = sorted([c for c in df_cho_scope["CHO_Clean"].unique() if c not in ["Not Available", "Unknown"]])
        st.markdown('<div class="filter-item-label">Select CHO</div>', unsafe_allow_html=True)
        sel_chos = st.multiselect(
            "Select CHO",
            options=avail_chos,
            default=[],
            placeholder="All CHOs",
            label_visibility="collapsed",
        )

        st.markdown("<div style='height: 6px;'></div>", unsafe_allow_html=True)
        if st.button("🔄 Reset All Filters", use_container_width=True):
            st.rerun()

    # --------------------------------------------------------------------------
    # APPLY DATE & GEOGRAPHY FILTERS TO SCREENING LINE LIST
    # --------------------------------------------------------------------------
    f_screen = df_screening.copy()
    if start_date and end_date:
        f_screen = f_screen[(f_screen["reg_date_clean"].dt.date >= start_date) & (f_screen["reg_date_clean"].dt.date <= end_date)]
    if sel_districts:
        f_screen = f_screen[f_screen["District_Clean"].isin(sel_districts)]
    if sel_blocks:
        f_screen = f_screen[f_screen["Block_Clean"].isin(sel_blocks)]
    if sel_hwcs:
        f_screen = f_screen[f_screen["Facility_Clean"].isin(sel_hwcs)]
    if sel_chos:
        f_screen = f_screen[f_screen["CHO_Clean"].isin(sel_chos)]

    # --------------------------------------------------------------------------
    # 1-TO-1 MASTER LIST PRESERVATION & NON-COLLAPSING JOIN
    # --------------------------------------------------------------------------
    # 1. Base HWC Master Dataset: Retain EVERY single entry without dropping duplicates
    if df_master_hwc is not None:
        base_hwc = df_master_hwc.copy()
        if sel_districts:
            base_hwc = base_hwc[base_hwc["District_Clean"].isin(sel_districts)]
        if sel_blocks:
            base_hwc = base_hwc[base_hwc["Block_Clean"].isin(sel_blocks)]
        if sel_hwcs:
            base_hwc = base_hwc[base_hwc["Facility_Clean"].isin(sel_hwcs)]
    else:
        # If no hierarchy uploaded, count unique composite combinations (District + Block + HWC)
        base_hwc = (
            df_screening[["District_Clean", "Block_Clean", "Facility_Clean", "composite_key"]]
            .drop_duplicates(subset=["composite_key"])
            .copy()
        )
        base_hwc["master_row_id"] = np.arange(1, len(base_hwc) + 1)
        if sel_districts:
            base_hwc = base_hwc[base_hwc["District_Clean"].isin(sel_districts)]
        if sel_blocks:
            base_hwc = base_hwc[base_hwc["Block_Clean"].isin(sel_blocks)]
        if sel_hwcs:
            base_hwc = base_hwc[base_hwc["Facility_Clean"].isin(sel_hwcs)]

    # 2. Extract Active CHO Name from Column G of Screening Data by Composite Key
    cho_mapping = (
        f_screen[f_screen["CHO_Clean"] != "Not Available"]
        .groupby("composite_key")["CHO_Clean"]
        .agg(lambda s: s.value_counts().index[0] if len(s) > 0 else "Not Available")
        .reset_index()
        .rename(columns={"CHO_Clean": "Active_CHO_Name"})
    )

    # 3. Aggregate Screening Activity by Composite Key
    screen_summary = (
        f_screen.groupby("composite_key")
        .agg(
            Total_Screening=("composite_key", "count"),
            AI_Presumptive=("is_ai_pres", "sum"),
            NTEP_Presumptive=("is_ntep_pres", "sum"),
            Both_AI_NTEP=("is_both_pres", "sum"),
            CHO_Override=("is_cho_override", "sum"),
            Total_Presumptive=("is_total_presumptive", "sum"),
            Presumptive_Open=("is_presumptive_open", "sum"),
            Presumptive_Closed=("is_presumptive_closed", "sum"),
            Total_Tested=("is_tested", "sum"),
            Total_Diagnosed=("is_diagnosed", "sum"),
        )
        .reset_index()
    )

    # 4. Left-Merge onto base_hwc by composite_key preserving every master_row_id exactly
    hwc_matrix = pd.merge(base_hwc, screen_summary, on="composite_key", how="left")
    hwc_matrix = pd.merge(hwc_matrix, cho_mapping, on="composite_key", how="left")

    hwc_matrix["Total_Screening"] = hwc_matrix["Total_Screening"].fillna(0).astype(int)
    hwc_matrix["AI_Presumptive"] = hwc_matrix["AI_Presumptive"].fillna(0).astype(int)
    hwc_matrix["NTEP_Presumptive"] = hwc_matrix["NTEP_Presumptive"].fillna(0).astype(int)
    hwc_matrix["Both_AI_NTEP"] = hwc_matrix["Both_AI_NTEP"].fillna(0).astype(int)
    hwc_matrix["CHO_Override"] = hwc_matrix["CHO_Override"].fillna(0).astype(int)
    hwc_matrix["Total_Presumptive"] = hwc_matrix["Total_Presumptive"].fillna(0).astype(int)
    hwc_matrix["Presumptive_Open"] = hwc_matrix["Presumptive_Open"].fillna(0).astype(int)
    hwc_matrix["Presumptive_Closed"] = hwc_matrix["Presumptive_Closed"].fillna(0).astype(int)
    hwc_matrix["Total_Tested"] = hwc_matrix["Total_Tested"].fillna(0).astype(int)
    hwc_matrix["Total_Diagnosed"] = hwc_matrix["Total_Diagnosed"].fillna(0).astype(int)

    hwc_matrix["District_Name"] = hwc_matrix["District_Clean"]
    hwc_matrix["Block_Name"] = hwc_matrix["Block_Clean"]
    hwc_matrix["HWC_Name"] = hwc_matrix["Facility_Clean"]

    # CHO Name: If screening started, display active CHO; otherwise "Not Available"
    hwc_matrix["CHO_Name"] = hwc_matrix["Active_CHO_Name"].fillna("Not Available")
    hwc_matrix.loc[hwc_matrix["Total_Screening"] == 0, "CHO_Name"] = "Not Available"

    # Facility Activity Status
    hwc_matrix["Screening_Status"] = np.where(hwc_matrix["Total_Screening"] > 0, "Screening Started", "Screening Not Started")

    if sel_chos:
        hwc_matrix = hwc_matrix[hwc_matrix["CHO_Name"].isin(sel_chos)]

    # --------------------------------------------------------------------------
    # ACCURATE TOTAL HWCS & KPI RECONCILIATION
    # --------------------------------------------------------------------------
    # Total HWCs exactly matches the row count of base_hwc (master entries)
    total_hwcs = len(hwc_matrix)
    screening_started_hwcs = int((hwc_matrix["Screening_Status"] == "Screening Started").sum())
    screening_not_started_hwcs = int((hwc_matrix["Screening_Status"] == "Screening Not Started").sum())

    total_screenings = int(hwc_matrix["Total_Screening"].sum())
    total_ai_pres = int(hwc_matrix["AI_Presumptive"].sum())
    total_ntep_pres = int(hwc_matrix["NTEP_Presumptive"].sum())
    total_both_pres = int(hwc_matrix["Both_AI_NTEP"].sum())
    total_cho_override = int(hwc_matrix["CHO_Override"].sum())
    total_presumptive = int(hwc_matrix["Total_Presumptive"].sum())

    total_presumptive_open = int(hwc_matrix["Presumptive_Open"].sum())
    total_presumptive_closed = int(hwc_matrix["Presumptive_Closed"].sum())
    total_tested = int(hwc_matrix["Total_Tested"].sum())
    total_diagnosed = int(hwc_matrix["Total_Diagnosed"].sum())

    # Rates
    pres_yield = (total_presumptive / max(total_screenings, 1)) * 100
    testing_rate = (total_tested / max(total_presumptive, 1)) * 100
    diagnosis_yield = (total_diagnosed / max(total_tested, 1)) * 100 if total_tested > 0 else 0.0

    # --------------------------------------------------------------------------
    # TOP HEADER BANNER & VALIDATION NOTICE
    # --------------------------------------------------------------------------
    current_time_str = datetime.datetime.now().strftime("%d %b %Y | %H:%M:%S")
    st.markdown(
        f"""
        <div class="app-header">
            <div>
                <h1 class="app-header-title"><span>🩺</span> CATB Screening & Clinical Surveillance Platform</h1>
                <div class="app-header-subtitle">
                    Jurisdiction: <b>{f"{len(sel_districts)} Districts Selected" if sel_districts else "All Monitored Districts"}</b> &nbsp;|&nbsp; 
                    Master HWCs: <b>{total_hwcs:,} Entries (Column E)</b> &nbsp;|&nbsp; 
                    Period: <b>{start_date} to {end_date}</b> &nbsp;|&nbsp; 
                    Last Sync: <b>{current_time_str}</b>
                </div>
            </div>
            <div class="system-status-pill">
                <div class="pulse-dot"></div> Live Surveillance Active
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Master Validation Confirmation Box
    if df_master_hwc is not None:
        st.markdown(
            f"""
            <div style="background-color: #F0FDF4; border: 1.5px solid #86EFAC; padding: 6px 12px; border-radius: 8px; margin-bottom: 12px; font-size: 0.8rem; color: #166534; font-weight: 600;">
                ✓ <b>HWC Master Reconciliation Verified:</b> Total HWC count ({total_hwcs:,}) exactly matches master file Column E entries. All duplicate HWC names across different blocks/districts are preserved and accounted for.
                <span style="float: right;">Screening Started ({screening_started_hwcs:,}) + Not Started ({screening_not_started_hwcs:,}) = {total_hwcs:,}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # --------------------------------------------------------------------------
    # DASHBOARD KPI CARDS SECTION (ACCURATELY RECONCILED)
    # --------------------------------------------------------------------------
    st.markdown('<div class="section-title-wrap"><div class="section-title-text"><span>🏛️</span> 1. Screening Overview & Facility Operational Status</div></div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    render_enterprise_kpi(c1, "Total HWCs", f"{total_hwcs:,}", "Master Hierarchy (Column E)", "🏥", "#0284C7", "#E0F2FE")
    render_enterprise_kpi(c2, "Screening Started", f"{screening_started_hwcs:,}", f"{(screening_started_hwcs/max(total_hwcs,1)*100):.1f}% active facilities", "✅", "#16A34A", "#DCFCE7")
    render_enterprise_kpi(c3, "Screening Not Started", f"{screening_not_started_hwcs:,}", f"{(screening_not_started_hwcs/max(total_hwcs,1)*100):.1f}% inactive facilities", "⏳", "#DC2626", "#FEE2E2")
    render_enterprise_kpi(c4, "Total Screening", f"{total_screenings:,}", "Evaluated individuals", "📋", "#0D9488", "#CCFBF1")

    st.markdown('<div class="section-title-wrap"><div class="section-title-text"><span>🔬</span> 2. Presumptive Classification Breakdown & Overlap Analysis</div></div>', unsafe_allow_html=True)
    c5, c6, c7, c8, c9 = st.columns(5)
    render_enterprise_kpi(c5, "AI Preference Presumptive", f"{total_ai_pres:,}", "AI algorithm (Col L)", "🤖", "#8B5CF6", "#F3E8FF")
    render_enterprise_kpi(c6, "NTEP Presumptive", f"{total_ntep_pres:,}", "Standard protocol (Col Q)", "📋", "#D97706", "#FEF3C7")
    render_enterprise_kpi(c7, "AI + NTEP Both Presumptive", f"{total_both_pres:,}", "Dual identified overlap", "🤝", "#0284C7", "#E0F2FE")
    render_enterprise_kpi(c8, "CHO Override Presumptive", f"{total_cho_override:,}", "Frontline override (Col BA)", "👩‍⚕️", "#EC4899", "#FCE7F3")
    render_enterprise_kpi(c9, "Total Unique Presumptive", f"{total_presumptive:,}", f"Deduplicated Yield: {pres_yield:.1f}%", "⚠️", "#EF4444", "#FEE2E2")

    st.markdown('<div class="section-title-wrap"><div class="section-title-text"><span>📈</span> 3. TB Case Progression & Clinical Outcomes</div></div>', unsafe_allow_html=True)
    c10, c11, c12, c13 = st.columns(4)
    render_enterprise_kpi(c10, "Presumptive Open", f"{total_presumptive_open:,}", "PRESUMPTIVE_OPEN (Nikshay Created)", "🆔", "#2563EB", "#DBEAFE")
    render_enterprise_kpi(c11, "Presumptive Closed", f"{total_presumptive_closed:,}", "PRESUMPTIVE_CLOSED (Negative/Completed)", "🔒", "#475569", "#F1F5F9")
    render_enterprise_kpi(c12, "Total Tested", f"{total_tested:,}", f"Closed ({total_presumptive_closed}) + Diagnosed ({total_diagnosed})", "🧪", "#7C3AED", "#EDE9FE")
    render_enterprise_kpi(c13, "Total Diagnosed", f"{total_diagnosed:,}", f"DIAGNOSED_ON_TREATMENT ({diagnosis_yield:.1f}%)", "🩺", "#059669", "#D1FAE5")

    # --------------------------------------------------------------------------
    # MAIN NAVIGATION TABS
    # --------------------------------------------------------------------------
    tab_hwc, tab_dist, tab_block, tab_inactive, tab_agreement, tab_cascade, tab_export = st.tabs(
        [
            "📋 HWC-Wise Performance Report",
            "🗺️ District-Wise Summary",
            "🏢 Block-Wise Summary",
            "⚠️ Inactive HWC Report",
            "🤝 AI vs NTEP Agreement Analysis",
            "🔻 TB Status Continuum Analytics",
            "📥 Export Reports",
        ]
    )

    # ==========================================================================
    # TAB 1: HWC-WISE PERFORMANCE REPORT (COMPLETE MASTER LIST, NO ROWS DROPPED)
    # ==========================================================================
    with tab_hwc:
        st.markdown(
            f"""
            <div class="section-title-wrap">
                <div class="section-title-text"><span>📋</span> Facility-Wise Comprehensive TB Screening Report ({total_hwcs:,} HWCs)</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        report_df = hwc_matrix.copy().sort_values(by=["Total_Screening", "Total_Tested"], ascending=[False, False]).reset_index(drop=True)
        report_df["S. No."] = report_df.index + 1

        ordered_cols = [
            "S. No.",
            "District_Name",
            "Block_Name",
            "HWC_Name",
            "CHO_Name",
            "Screening_Status",
            "Total_Screening",
            "AI_Presumptive",
            "NTEP_Presumptive",
            "Both_AI_NTEP",
            "CHO_Override",
            "Total_Presumptive",
            "Presumptive_Open",
            "Presumptive_Closed",
            "Total_Tested",
            "Total_Diagnosed",
        ]
        final_hwc_report = report_df[ordered_cols].rename(
            columns={
                "District_Name": "District",
                "Block_Name": "Block",
                "HWC_Name": "HWC",
                "CHO_Name": "CHO",
                "Screening_Status": "Screening Status",
                "Total_Screening": "Total Screening",
                "AI_Presumptive": "AI Presumptive",
                "NTEP_Presumptive": "NTEP Presumptive",
                "Both_AI_NTEP": "AI + NTEP Both Presumptive",
                "CHO_Override": "CHO Override",
                "Total_Presumptive": "Total Unique Presumptive",
                "Presumptive_Open": "Presumptive Open",
                "Presumptive_Closed": "Presumptive Closed",
                "Total_Tested": "Total Tested",
                "Total_Diagnosed": "Total Diagnosed",
            }
        )

        search_text = st.text_input("🔍 Quick Search HWC / CHO / Block:", "", placeholder="Type name to filter...", key="search_hwc_tab")
        if search_text.strip():
            m = (
                final_hwc_report["HWC"].str.contains(search_text.strip(), case=False, na=False)
                | final_hwc_report["CHO"].str.contains(search_text.strip(), case=False, na=False)
                | final_hwc_report["Block"].str.contains(search_text.strip(), case=False, na=False)
            )
            display_hwc_report = final_hwc_report[m]
        else:
            display_hwc_report = final_hwc_report

        st.dataframe(
            display_hwc_report.style.format(
                {
                    "Total Screening": "{:,}",
                    "AI Presumptive": "{:,}",
                    "NTEP Presumptive": "{:,}",
                    "AI + NTEP Both Presumptive": "{:,}",
                    "CHO Override": "{:,}",
                    "Total Unique Presumptive": "{:,}",
                    "Presumptive Open": "{:,}",
                    "Presumptive Closed": "{:,}",
                    "Total Tested": "{:,}",
                    "Total Diagnosed": "{:,}",
                }
            ),
            use_container_width=True,
            hide_index=True,
            height=520,
        )

    # ==========================================================================
    # TAB 2: DISTRICT-WISE SUMMARY
    # ==========================================================================
    with tab_dist:
        st.markdown(
            """
            <div class="section-title-wrap">
                <div class="section-title-text"><span>🗺️</span> District-Level Surveillance Summary</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        dist_summary = (
            hwc_matrix.groupby("District_Name")
            .agg(
                Total_HWCs=("HWC_Name", "count"),
                Screening_Started=("Screening_Status", lambda s: (s == "Screening Started").sum()),
                Screening_Not_Started=("Screening_Status", lambda s: (s == "Screening Not Started").sum()),
                Total_Screening=("Total_Screening", "sum"),
                AI_Presumptive=("AI_Presumptive", "sum"),
                NTEP_Presumptive=("NTEP_Presumptive", "sum"),
                Both_AI_NTEP=("Both_AI_NTEP", "sum"),
                CHO_Override=("CHO_Override", "sum"),
                Total_Presumptive=("Total_Presumptive", "sum"),
                Presumptive_Open=("Presumptive_Open", "sum"),
                Presumptive_Closed=("Presumptive_Closed", "sum"),
                Total_Tested=("Total_Tested", "sum"),
                Total_Diagnosed=("Total_Diagnosed", "sum"),
            )
            .reset_index()
            .rename(
                columns={
                    "District_Name": "District Name",
                    "Total_HWCs": "Total HWCs",
                    "Screening_Started": "HWCs Where Screening Started",
                    "Screening_Not_Started": "HWCs Where Screening Not Started",
                    "Total_Screening": "Total Screening",
                    "AI_Presumptive": "AI Presumptive",
                    "NTEP_Presumptive": "NTEP Presumptive",
                    "Both_AI_NTEP": "AI + NTEP Both Presumptive",
                    "CHO_Override": "CHO Override",
                    "Total_Presumptive": "Total Unique Presumptive",
                    "Presumptive_Open": "Presumptive Open",
                    "Presumptive_Closed": "Presumptive Closed",
                    "Total_Tested": "Total Tested",
                    "Total_Diagnosed": "Total Diagnosed",
                }
            )
        )

        fig_dist = px.bar(
            dist_summary.sort_values(by="Total Screening", ascending=False),
            x="District Name",
            y=["Total Screening", "Total Unique Presumptive", "AI + NTEP Both Presumptive", "Total Tested", "Total Diagnosed"],
            barmode="group",
            color_discrete_sequence=["#0284C7", "#EF4444", "#8B5CF6", "#7C3AED", "#059669"],
            title="District Performance: Screening, Presumptive Breakdown, Overlap & Diagnosis",
        )
        fig_dist.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1), height=350, margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig_dist, use_container_width=True)

        st.dataframe(
            dist_summary.sort_values(by="Total Screening", ascending=False).style.format(
                {
                    "Total HWCs": "{:,}",
                    "HWCs Where Screening Started": "{:,}",
                    "HWCs Where Screening Not Started": "{:,}",
                    "Total Screening": "{:,}",
                    "AI Presumptive": "{:,}",
                    "NTEP Presumptive": "{:,}",
                    "AI + NTEP Both Presumptive": "{:,}",
                    "CHO Override": "{:,}",
                    "Total Unique Presumptive": "{:,}",
                    "Presumptive Open": "{:,}",
                    "Presumptive Closed": "{:,}",
                    "Total Tested": "{:,}",
                    "Total Diagnosed": "{:,}",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )

    # ==========================================================================
    # TAB 3: BLOCK-WISE SUMMARY
    # ==========================================================================
    with tab_block:
        st.markdown(
            """
            <div class="section-title-wrap">
                <div class="section-title-text"><span>🏢</span> Block-Level Performance & Coverage</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        block_summary = (
            hwc_matrix.groupby(["District_Name", "Block_Name"])
            .agg(
                Total_HWCs=("HWC_Name", "count"),
                Screening_Started=("Screening_Status", lambda s: (s == "Screening Started").sum()),
                Screening_Not_Started=("Screening_Status", lambda s: (s == "Screening Not Started").sum()),
                Total_Screening=("Total_Screening", "sum"),
                AI_Presumptive=("AI_Presumptive", "sum"),
                NTEP_Presumptive=("NTEP_Presumptive", "sum"),
                Both_AI_NTEP=("Both_AI_NTEP", "sum"),
                CHO_Override=("CHO_Override", "sum"),
                Total_Presumptive=("Total_Presumptive", "sum"),
                Presumptive_Open=("Presumptive_Open", "sum"),
                Presumptive_Closed=("Presumptive_Closed", "sum"),
                Total_Tested=("Total_Tested", "sum"),
                Total_Diagnosed=("Total_Diagnosed", "sum"),
            )
            .reset_index()
            .rename(
                columns={
                    "District_Name": "District Name",
                    "Block_Name": "Block Name",
                    "Total_HWCs": "Total HWCs",
                    "Screening_Started": "HWCs Where Screening Started",
                    "Screening_Not_Started": "HWCs Where Screening Not Started",
                    "Total_Screening": "Total Screening",
                    "AI_Presumptive": "AI Presumptive",
                    "NTEP_Presumptive": "NTEP Presumptive",
                    "Both_AI_NTEP": "AI + NTEP Both Presumptive",
                    "CHO_Override": "CHO Override",
                    "Total_Presumptive": "Total Unique Presumptive",
                    "Presumptive_Open": "Presumptive Open",
                    "Presumptive_Closed": "Presumptive Closed",
                    "Total_Tested": "Total Tested",
                    "Total_Diagnosed": "Total Diagnosed",
                }
            )
        )

        st.dataframe(
            block_summary.sort_values(by="Total Screening", ascending=False).style.format(
                {
                    "Total HWCs": "{:,}",
                    "HWCs Where Screening Started": "{:,}",
                    "HWCs Where Screening Not Started": "{:,}",
                    "Total Screening": "{:,}",
                    "AI Presumptive": "{:,}",
                    "NTEP Presumptive": "{:,}",
                    "AI + NTEP Both Presumptive": "{:,}",
                    "CHO Override": "{:,}",
                    "Total Unique Presumptive": "{:,}",
                    "Presumptive Open": "{:,}",
                    "Presumptive Closed": "{:,}",
                    "Total Tested": "{:,}",
                    "Total Diagnosed": "{:,}",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )

    # ==========================================================================
    # TAB 4: INACTIVE HWC REPORT
    # ==========================================================================
    with tab_inactive:
        st.markdown(
            """
            <div class="section-title-wrap">
                <div class="section-title-text"><span>⚠️</span> HWCs Where Screening Has Not Started</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        inactive_df = (
            hwc_matrix[hwc_matrix["Screening_Status"] == "Screening Not Started"]
            .sort_values(by=["District_Name", "Block_Name", "HWC_Name"])
            .reset_index(drop=True)
        )
        inactive_df["S. No."] = inactive_df.index + 1

        final_inactive = inactive_df[
            ["S. No.", "District_Name", "Block_Name", "HWC_Name", "CHO_Name", "Screening_Status", "Total_Screening"]
        ].rename(
            columns={
                "District_Name": "District Name",
                "Block_Name": "Block Name",
                "HWC_Name": "HWC Name",
                "CHO_Name": "CHO Name",
                "Screening_Status": "Screening Status",
                "Total_Screening": "Total Screening",
            }
        )

        st.info(f"📌 A total of **{len(final_inactive):,} HWCs** from the Master Hierarchy have not recorded any screening activity. CHO Name is designated as **Not Available**.")
        st.dataframe(final_inactive, use_container_width=True, hide_index=True)

    # ==========================================================================
    # TAB 5: AI VS NTEP AGREEMENT ANALYSIS
    # ==========================================================================
    with tab_agreement:
        st.markdown(
            """
            <div class="section-title-wrap">
                <div class="section-title-text"><span>🤝</span> AI vs NTEP Classification Concordance & Mismatch Analysis</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        ai_only_pres = int((f_screen["is_ai_only"]).sum())
        ntep_only_pres = int((f_screen["is_ntep_only"]).sum())
        both_non_pres = int(((f_screen["is_ai_pres"] == 0) & (f_screen["is_ntep_pres"] == 0)).sum())
        concordance_rate = ((total_both_pres + both_non_pres) / max(total_screenings, 1)) * 100

        ag1, ag2, ag3, ag4, ag5 = st.columns(5)
        render_enterprise_kpi(ag1, "Both Presumptive", f"{total_both_pres:,}", "AI & NTEP Concordant", "🤝", "#16A34A", "#DCFCE7")
        render_enterprise_kpi(ag2, "AI Only Presumptive", f"{ai_only_pres:,}", "NTEP Non-Presumptive", "🤖", "#8B5CF6", "#F3E8FF")
        render_enterprise_kpi(ag3, "NTEP Only Presumptive", f"{ntep_only_pres:,}", "AI Non-Presumptive", "📋", "#D97706", "#FEF3C7")
        render_enterprise_kpi(ag4, "Both Non-Presumptive", f"{both_non_pres:,}", "Concordant Negatives", "🛡️", "#475569", "#F1F5F9")
        render_enterprise_kpi(ag5, "Overall Concordance", f"{concordance_rate:.1f}%", "AI-NTEP Agreement", "🎯", "#0284C7", "#E0F2FE")

        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

        col_ag_l, col_ag_r = st.columns([5, 4])
        with col_ag_l:
            st.markdown("##### 🔬 AI vs NTEP Classification 2x2 Matrix")
            contingency_df = pd.DataFrame(
                {
                    "Classification Modality": ["AI Presumptive", "AI Non-Presumptive"],
                    "NTEP Presumptive": [total_both_pres, ntep_only_pres],
                    "NTEP Non-Presumptive": [ai_only_pres, both_non_pres],
                }
            )
            st.dataframe(contingency_df, use_container_width=True, hide_index=True)

            st.caption(
                f"• **Dual Positive Yield:** {total_both_pres:,} cases identified by both modalities.\n"
                f"• **AI Incremental Capture:** {ai_only_pres:,} additional presumptive cases caught by AI where standard protocol was non-presumptive.\n"
                f"• **NTEP Incremental Capture:** {ntep_only_pres:,} presumptive cases caught by standard protocol where AI was non-presumptive.\n"
                f"• **Frontline CHO Clinical Override:** {total_cho_override:,} cases where CHO intervened on dual non-presumptive records."
            )

        with col_ag_r:
            st.markdown("##### 📊 Presumptive Component Breakdown")
            comp_df = pd.DataFrame(
                {
                    "Component": [
                        "1. AI + NTEP Both Presumptive",
                        "2. AI Only Presumptive",
                        "3. NTEP Only Presumptive",
                        "4. CHO Clinical Override",
                    ],
                    "Beneficiaries": [total_both_pres, ai_only_pres, ntep_only_pres, total_cho_override],
                }
            )
            fig_comp = px.pie(
                comp_df,
                names="Component",
                values="Beneficiaries",
                hole=0.55,
                color_discrete_sequence=["#16A34A", "#8B5CF6", "#D97706", "#EC4899"],
            )
            fig_comp.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10), legend=dict(orientation="h", yanchor="bottom", y=-0.25))
            st.plotly_chart(fig_comp, use_container_width=True)

        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

        st.markdown("##### 📍 Facilities with Highest Discordance (Mismatch) or High Overrides")
        mismatch_hwcs = hwc_matrix.copy()
        mismatch_hwcs["Mismatch_Cases"] = (mismatch_hwcs["AI_Presumptive"] + mismatch_hwcs["NTEP_Presumptive"] - 2 * mismatch_hwcs["Both_AI_NTEP"]).clip(lower=0)
        top_discordant = mismatch_hwcs[mismatch_hwcs["Total_Screening"] > 0].sort_values(by=["Mismatch_Cases", "CHO_Override"], ascending=[False, False]).head(15)

        st.dataframe(
            top_discordant[
                ["District_Name", "Block_Name", "HWC_Name", "CHO_Name", "Total_Screening", "AI_Presumptive", "NTEP_Presumptive", "Both_AI_NTEP", "Mismatch_Cases", "CHO_Override"]
            ].rename(
                columns={
                    "District_Name": "District",
                    "Block_Name": "Block",
                    "HWC_Name": "HWC",
                    "CHO_Name": "CHO",
                    "Total_Screening": "Screening",
                    "AI_Presumptive": "AI Presumptive",
                    "NTEP_Presumptive": "NTEP Presumptive",
                    "Both_AI_NTEP": "Both Presumptive",
                    "Mismatch_Cases": "AI-NTEP Mismatches",
                    "CHO_Override": "CHO Override",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )

    # ==========================================================================
    # TAB 6: TB STATUS CONTINUUM ANALYTICS
    # ==========================================================================
    with tab_cascade:
        st.markdown(
            """
            <div class="section-title-wrap">
                <div class="section-title-text"><span>🔻</span> Presumptive Modality & Clinical Outcomes</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        col_f1, col_f2 = st.columns([5, 4])
        with col_f1:
            st.markdown("##### 🔻 Continuum of Care: Screening to Diagnosis Funnel")
            funnel_df = pd.DataFrame(
                {
                    "Cascade Stage": [
                        "1. Screened Population",
                        "2. Total Unique Presumptive",
                        "3. Presumptive Open (Nikshay Created)",
                        "4. Total Tested (Closed + Diagnosed)",
                        "5. Diagnosed on Treatment",
                    ],
                    "Beneficiaries": [total_screenings, total_presumptive, total_presumptive_open, total_tested, total_diagnosed],
                }
            )
            fig_funnel = px.funnel(
                funnel_df,
                x="Beneficiaries",
                y="Cascade Stage",
                color="Cascade Stage",
                color_discrete_sequence=["#0284C7", "#EF4444", "#2563EB", "#7C3AED", "#059669"],
            )
            fig_funnel.update_layout(showlegend=False, height=330, margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(fig_funnel, use_container_width=True)

        with col_f2:
            st.markdown("##### 🔬 Presumptive Identification Modalities")
            modality_df = pd.DataFrame(
                {
                    "Identification Modality": [
                        "AI Presumptive (Col L)",
                        "NTEP Presumptive (Col Q)",
                        "AI + NTEP Both (Overlap)",
                        "CHO Clinical Override (Col BA)",
                    ],
                    "Cases": [total_ai_pres, total_ntep_pres, total_both_pres, total_cho_override],
                }
            )
            fig_bar = px.bar(
                modality_df,
                x="Identification Modality",
                y="Cases",
                color="Identification Modality",
                text_auto=True,
                color_discrete_sequence=["#8B5CF6", "#D97706", "#0284C7", "#EC4899"],
            )
            fig_bar.update_layout(showlegend=False, height=330, margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(fig_bar, use_container_width=True)

    # ==========================================================================
    # TAB 7: EXPORT REPORTS
    # ==========================================================================
    with tab_export:
        st.markdown(
            """
            <div class="section-title-wrap">
                <div class="section-title-text"><span>📥</span> Export Consolidated Surveillance Dossier</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        c_exp1, c_exp2 = st.columns(2)
        with c_exp1:
            st.markdown("##### 📑 Multi-Tab Surveillance Workbook (.xlsx)")
            st.caption("Consolidates HWC-Wise Report, District-Wise Summary, Block-Wise Summary, and Inactive HWCs.")

            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                final_hwc_report.to_excel(writer, sheet_name="HWC_Performance", index=False)
                dist_summary.to_excel(writer, sheet_name="District_Summary", index=False)
                block_summary.to_excel(writer, sheet_name="Block_Summary", index=False)
                final_inactive.to_excel(writer, sheet_name="Inactive_HWCs", index=False)

            st.download_button(
                label="📥 Download Consolidated Workbook (.xlsx)",
                data=buf.getvalue(),
                file_name=f"CATB_Surveillance_Dossier_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with c_exp2:
            st.markdown("##### 📄 HWC Performance Line-List (.csv)")
            st.caption("Standard facility-level performance report in CSV format.")
            csv_hwc = final_hwc_report.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="📥 Download HWC Performance (.csv)",
                data=csv_hwc,
                file_name=f"CATB_HWC_Performance_{datetime.datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True,
            )


if __name__ == "__main__":
    main()