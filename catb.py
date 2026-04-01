import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO

REPORTLAB_AVAILABLE = True
REPORTLAB_IMPORT_ERROR = None

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.platypus import (
        SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
    )
except Exception as e:
    REPORTLAB_AVAILABLE = False
    REPORTLAB_IMPORT_ERROR = str(e)


st.set_page_config(
    page_title="CATB Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)


st.markdown("""
<style>
:root {
    --primary: #12355b;
    --secondary: #1d6fdc;
    --accent: #4b9cff;
    --soft: #eef4ff;
    --card: rgba(255,255,255,0.98);
    --border: #dfe9fb;
    --text: #16324f;
    --muted: #5e768f;
    --success: #0f8b4c;
    --warning: #d98200;
    --danger: #c92a2a;
}
[data-testid="stAppViewContainer"] {
    background:
        radial-gradient(circle at top right, rgba(75,156,255,0.12), transparent 28%),
        linear-gradient(180deg, #f5f8fd 0%, #edf4ff 45%, #f8fbff 100%);
}
[data-testid="stHeader"] {
    background: rgba(0,0,0,0);
}
.block-container {
    padding-top: 0.8rem;
    padding-bottom: 1.2rem;
    max-width: 98%;
}
.main-title {
    padding: 1.45rem 1.6rem;
    border-radius: 24px;
    background: linear-gradient(135deg, #0d2745 0%, #144f9e 45%, #2e82f0 75%, #67b0ff 100%);
    color: white;
    box-shadow: 0 18px 40px rgba(18,53,91,0.22);
    margin-bottom: 0.8rem;
    border: 1px solid rgba(255,255,255,0.18);
    position: relative;
    overflow: hidden;
}
.main-title::after {
    content: "";
    position: absolute;
    width: 240px;
    height: 240px;
    right: -40px;
    top: -70px;
    background: radial-gradient(circle, rgba(255,255,255,0.18), rgba(255,255,255,0.02));
    border-radius: 50%;
}
.main-title h1 {
    margin: 0;
    padding: 0;
    color: white;
    font-size: 2.25rem;
    letter-spacing: 0.3px;
}
.main-title p {
    margin: 0.45rem 0 0 0;
    color: #edf5ff;
    font-size: 1rem;
    max-width: 900px;
}
.summary-banner {
    background: linear-gradient(90deg, rgba(18,53,91,0.98), rgba(29,111,220,0.94));
    color: white;
    border-radius: 18px;
    padding: 0.9rem 1rem;
    margin-bottom: 1rem;
    box-shadow: 0 12px 28px rgba(18,53,91,0.18);
    border: 1px solid rgba(255,255,255,0.12);
}
.summary-banner span {
    display: inline-block;
    margin-right: 18px;
    margin-bottom: 4px;
    font-size: 0.98rem;
}
div[data-testid="stMetric"] {
    background: linear-gradient(135deg, #ffffff 0%, #f8fbff 100%);
    border: 1px solid #dbe8ff;
    padding: 16px;
    border-radius: 18px;
    box-shadow: 0 10px 24px rgba(20, 50, 90, 0.08);
}
div[data-testid="stMetric"]:hover {
    transform: translateY(-2px);
    transition: 0.25s ease;
    box-shadow: 0 14px 30px rgba(20, 50, 90, 0.12);
}
div[data-testid="stMetricLabel"] {
    color: #4b6480;
    font-weight: 800;
    font-size: 0.93rem;
}
div[data-testid="stMetricValue"] {
    color: #12355b;
}
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #12355b 0%, #194b80 100%);
}
section[data-testid="stSidebar"] * {
    color: white !important;
}
.sidebar-card {
    background: rgba(255,255,255,0.10);
    padding: 12px;
    border-radius: 14px;
    margin-bottom: 12px;
    border: 1px solid rgba(255,255,255,0.14);
}
.chart-card, .table-card, .info-card {
    background: rgba(255,255,255,0.98);
    border-radius: 22px;
    padding: 12px 16px 10px 16px;
    box-shadow: 0 10px 28px rgba(15, 40, 80, 0.07);
    border: 1px solid #e7eefb;
    margin-bottom: 14px;
}
.small-note {
    color: #58708a;
    font-size: 0.95rem;
    background: #ffffffd8;
    padding: 10px 14px;
    border-radius: 14px;
    border: 1px solid #e7eefb;
    box-shadow: 0 8px 20px rgba(15, 40, 80, 0.04);
}
.filter-box {
    background: rgba(255,255,255,0.10);
    padding: 10px;
    border-radius: 10px;
    margin-bottom: 10px;
    border: 1px solid rgba(255,255,255,0.12);
}
button[data-baseweb="tab"] {
    border-radius: 14px;
    padding: 9px 16px;
    background: #eef4ff;
    margin-right: 5px;
    border: 1px solid #d7e6ff;
    font-weight: 600;
}
button[data-baseweb="tab"][aria-selected="true"] {
    background: linear-gradient(135deg, #12355b, #1d6fdc);
    color: white !important;
    box-shadow: 0 10px 18px rgba(18,53,91,0.18);
}
div.stButton > button, div.stDownloadButton > button {
    border-radius: 12px;
    border: none;
    background: linear-gradient(135deg, #12355b, #1d6fdc);
    color: white;
    font-weight: 700;
    padding: 0.65rem 1rem;
    box-shadow: 0 10px 18px rgba(18,53,91,0.16);
}
div.stButton > button:hover, div.stDownloadButton > button:hover {
    background: linear-gradient(135deg, #0f2d4d, #165fc0);
    color: white;
}
.section-title {
    color: #12355b;
    font-weight: 800;
    margin-top: 0.3rem;
    margin-bottom: 0.7rem;
    font-size: 1.1rem;
    padding: 0.75rem 1rem;
    border-radius: 16px;
    background: linear-gradient(90deg, rgba(18,53,91,0.08), rgba(29,111,220,0.12));
    border-left: 6px solid #1d6fdc;
    border-top: 1px solid #d8e7ff;
    border-right: 1px solid #d8e7ff;
    border-bottom: 1px solid #d8e7ff;
}
.subsection-title {
    color: #12355b;
    font-weight: 800;
    margin-bottom: 0.4rem;
    font-size: 1rem;
}
.logic-box {
    background: linear-gradient(135deg, #f8fbff 0%, #edf4ff 100%);
    border: 1px solid #d8e7ff;
    border-radius: 18px;
    padding: 14px 16px;
}
.logic-box h4 {
    margin: 0 0 8px 0;
    color: #12355b;
}
.logic-box p {
    margin: 0.32rem 0;
    color: #4a647d;
}
.metric-ribbon {
    background: linear-gradient(135deg, rgba(255,255,255,0.98), rgba(241,247,255,0.98));
    border: 1px solid #dbe8ff;
    border-radius: 18px;
    padding: 0.85rem 1rem;
    margin-bottom: 1rem;
    box-shadow: 0 10px 24px rgba(20, 50, 90, 0.06);
}
.metric-ribbon p {
    margin: 0;
    color: #47627d;
    font-size: 0.95rem;
}
.warning-box {
    background: #fff8e8;
    color: #7a5a00;
    border: 1px solid #f3dd9a;
    border-radius: 14px;
    padding: 12px 14px;
    margin-bottom: 12px;
}
</style>
""", unsafe_allow_html=True)


# Configuration for column detection
EXPECTED_DATE_COLS = ["Reg Date", "tbdiagnosisdate", "tb_diagnosis_date"]
FACILITY_CANDIDATES = ["HWC", "Facility Name", "Facility", "Subcenter", "HWC ID"]
BLOCK_CANDIDATES = ["Block"]
DISTRICT_CANDIDATES = ["District"]
CHO_CANDIDATES = ["CHO name"]


@st.cache_data
def load_excel(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = xls.sheet_names[0]
    
    # Read raw to find where headers start
    df_raw = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
    
    header_row_idx = 0
    # Scan first 20 rows to find a row containing "Reg Date" or "District"
    for i, row in df_raw.head(20).iterrows():
        row_values = [str(val).strip().lower() for val in row.dropna()]
        if any(keyword in row_values for keyword in ["reg date", "district", "block", "hwc"]):
            header_row_idx = i
            break

    # Re-read with correct header row
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, skiprows=header_row_idx)
    df.columns = [str(c).strip() for c in df.columns]

    # Standardize column types
    for col in EXPECTED_DATE_COLS:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace({
                "nan": np.nan, "None": np.nan, "null": np.nan, "": np.nan
            })

    numeric_candidates = [
        "Age", "ensemblemodelscore", "ensemble_model_score", "height", "weight",
        "Cough duration", "Fever Duration", "Chest pain duration",
        "Hemoptysis duration", "Loss of appetite duration",
        "Night sweat duration", "Shortness of breath duration",
        "Weight loss duration"
    ]
    for col in numeric_candidates:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df, sheet_name


def find_column(df, candidates):
    for col in candidates:
        if col in df.columns:
            return col
    return None


def find_facility_column(df):
    direct_match = find_column(df, FACILITY_CANDIDATES)
    if direct_match and direct_match != "HWC ID":
        return direct_match

    cols = df.columns.tolist()

    if "Block" in cols and "HWC ID" in cols:
        try:
            block_idx = cols.index("Block")
            hwc_idx = cols.index("HWC ID")
            if hwc_idx - block_idx == 2:
                inferred_facility_col = cols[block_idx + 1]
                if inferred_facility_col not in ["Block", "HWC ID"]:
                    return inferred_facility_col
        except Exception:
            pass

    if direct_match:
        return direct_match

    return None


def safe_unique(series):
    return sorted([x for x in series.dropna().astype(str).unique().tolist()])


def format_dates_for_display(df):
    temp = df.copy()
    for col in temp.columns:
        if pd.api.types.is_datetime64_any_dtype(temp[col]):
            temp[col] = temp[col].dt.strftime("%Y-%m-%d")
    return temp


def prepare_export_table(df, max_rows=None):
    if df is None or df.empty:
        return pd.DataFrame()
    out = df.copy()
    for col in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[col]):
            out[col] = out[col].dt.strftime("%Y-%m-%d")
    if max_rows is not None:
        out = out.head(max_rows)
    return out


def apply_filters(df):
    filtered = df.copy()

    district_col = find_column(filtered, DISTRICT_CANDIDATES)
    block_col = find_column(filtered, BLOCK_CANDIDATES)
    facility_col = find_facility_column(filtered)
    cho_col = find_column(filtered, CHO_CANDIDATES)
    date_col = "Reg Date" if "Reg Date" in filtered.columns else None

    st.sidebar.markdown(
        "<div class='sidebar-card'><h3 style='margin-top:0;'>Filters</h3><p style='margin-bottom:0;'>Refine dashboard by geography, facility, CHO, and date.</p></div>",
        unsafe_allow_html=True
    )

    if district_col:
        districts = safe_unique(filtered[district_col])
        selected_districts = st.sidebar.multiselect(
            "District", districts, default=districts,
            help="By default all districts are selected."
        )
        if selected_districts:
            filtered = filtered[filtered[district_col].astype(str).isin(selected_districts)]

    if block_col:
        blocks = safe_unique(filtered[block_col])
        selected_blocks = st.sidebar.multiselect(
            "Block", blocks, default=blocks,
            help="By default all blocks are selected."
        )
        if selected_blocks:
            filtered = filtered[filtered[block_col].astype(str).isin(selected_blocks)]

    if facility_col:
        facilities = safe_unique(filtered[facility_col])
        selected_facilities = st.sidebar.multiselect(
            "Facility / HWC", facilities, default=facilities,
            help="By default all facilities are selected."
        )
        if selected_facilities:
            filtered = filtered[filtered[facility_col].astype(str).isin(selected_facilities)]

    if cho_col:
        chos = safe_unique(filtered[cho_col])
        selected_chos = st.sidebar.multiselect(
            "CHO Name", chos, default=chos,
            help="By default all CHO names are selected."
        )
        if selected_chos:
            filtered = filtered[filtered[cho_col].astype(str).isin(selected_chos)]

    if date_col and filtered[date_col].notna().any():
        min_date = filtered[date_col].min().date()
        max_date = filtered[date_col].max().date()

        st.sidebar.markdown("<div class='filter-box'><b>Date Filter</b></div>", unsafe_allow_html=True)
        d1, d2 = st.sidebar.columns(2)

        with d1:
            start_date = st.date_input(
                "Start Date",
                value=min_date,
                min_value=min_date,
                max_value=max_date
            )

        with d2:
            end_date = st.date_input(
                "End Date",
                value=max_date,
                min_value=min_date,
                max_value=max_date
            )

        if start_date > end_date:
            st.sidebar.error("Start Date cannot be greater than End Date.")
        else:
            filtered = filtered[
                (filtered[date_col].dt.date >= start_date) &
                (filtered[date_col].dt.date <= end_date)
            ]

    return filtered, district_col, block_col, facility_col, cho_col, date_col


def build_grouped_summary(df, group_col, district_col=None, block_col=None, facility_col=None, cho_col=None):
    if not group_col or group_col not in df.columns:
        return pd.DataFrame()

    summary = df.groupby(group_col).size().reset_index(name="Total Records")

    if "Age" in df.columns:
        age_summary = df.groupby(group_col)["Age"].mean().round(1).reset_index(name="Average Age")
        summary = summary.merge(age_summary, on=group_col, how="left")

    if district_col and district_col in df.columns and group_col != district_col:
        district_summary = df.groupby(group_col)[district_col].nunique().reset_index(name="Total Districts")
        summary = summary.merge(district_summary, on=group_col, how="left")

    if block_col and block_col in df.columns and group_col != block_col:
        block_summary = df.groupby(group_col)[block_col].nunique().reset_index(name="Total Blocks")
        summary = summary.merge(block_summary, on=group_col, how="left")

    if facility_col and facility_col in df.columns and group_col != facility_col:
        facility_summary = df.groupby(group_col)[facility_col].nunique().reset_index(name="Total Facilities")
        summary = summary.merge(facility_summary, on=group_col, how="left")

    if cho_col and cho_col in df.columns and group_col != cho_col:
        cho_summary = df.groupby(group_col)[cho_col].nunique().reset_index(name="Total CHOs")
        summary = summary.merge(cho_summary, on=group_col, how="left")

    if "AI preference" in df.columns:
        presumptive = (
            df[df["AI preference"].astype(str).str.lower().eq("presumptive")]
            .groupby(group_col)
            .size()
            .reset_index(name="Presumptive")
        )
        summary = summary.merge(presumptive, on=group_col, how="left")

    if "Sent for testing" in df.columns:
        sent_for_testing = (
            df[df["Sent for testing"].astype(str).str.lower().eq("yes")]
            .groupby(group_col)
            .size()
            .reset_index(name="Sent for Testing")
        )
        summary = summary.merge(sent_for_testing, on=group_col, how="left")

    if "Type" in df.columns:
        symptomatic = (
            df[df["Type"].astype(str).str.lower().str.contains("symptomatic", na=False)]
            .groupby(group_col)
            .size()
            .reset_index(name="Symptomatic")
        )
        summary = summary.merge(symptomatic, on=group_col, how="left")

    summary = summary.fillna(0)

    for col in ["Total Districts", "Total Blocks", "Total Facilities", "Total CHOs", "Presumptive", "Sent for Testing", "Symptomatic"]:
        if col in summary.columns:
            summary[col] = summary[col].astype(int)

    return summary.sort_values("Total Records", ascending=False)


def get_active_entity_counts(df, entity_col, date_col):
    if not entity_col or entity_col not in df.columns or not date_col or date_col not in df.columns:
        return {"7d": 0, "15d": 0, "30d": 0, "latest_date": None}

    temp = df[[entity_col, date_col]].dropna().copy()
    if temp.empty:
        return {"7d": 0, "15d": 0, "30d": 0, "latest_date": None}

    latest_date = temp[date_col].max()
    if pd.isna(latest_date):
        return {"7d": 0, "15d": 0, "30d": 0, "latest_date": None}

    d7 = latest_date - pd.Timedelta(days=7)
    d15 = latest_date - pd.Timedelta(days=15)
    d30 = latest_date - pd.Timedelta(days=30)

    return {
        "7d": temp[temp[date_col] >= d7][entity_col].astype(str).nunique(),
        "15d": temp[temp[date_col] >= d15][entity_col].astype(str).nunique(),
        "30d": temp[temp[date_col] >= d30][entity_col].astype(str).nunique(),
        "latest_date": latest_date
    }


def build_active_user_table(df, entity_col, date_col):
    if not entity_col or entity_col not in df.columns or not date_col or date_col not in df.columns:
        return pd.DataFrame()

    temp = df[[entity_col, date_col]].dropna().copy()
    if temp.empty:
        return pd.DataFrame()

    latest_date = temp[date_col].max()
    d7 = latest_date - pd.Timedelta(days=7)
    d15 = latest_date - pd.Timedelta(days=15)
    d30 = latest_date - pd.Timedelta(days=30)

    summary = (
        temp.groupby(entity_col)[date_col]
        .max()
        .reset_index(name="Last Active Date")
    )

    summary["Days Since Last Activity"] = (latest_date - summary["Last Active Date"]).dt.days
    summary["Weekly Active"] = np.where(summary["Last Active Date"] >= d7, "Yes", "No")
    summary["Biweekly Active"] = np.where(summary["Last Active Date"] >= d15, "Yes", "No")
    summary["Monthly Active"] = np.where(summary["Last Active Date"] >= d30, "Yes", "No")

    summary["Activity Status"] = np.select(
        [
            summary["Last Active Date"] >= d7,
            summary["Last Active Date"] >= d15,
            summary["Last Active Date"] >= d30
        ],
        [
            "Weekly Active",
            "Biweekly Active",
            "Monthly Active"
        ],
        default="Inactive >30 Days"
    )

    summary["Active Score"] = np.select(
        [
            summary["Last Active Date"] >= d7,
            summary["Last Active Date"] >= d15,
            summary["Last Active Date"] >= d30
        ],
        [3, 2, 1],
        default=0
    )

    return summary.sort_values(["Active Score", "Last Active Date"], ascending=[False, False]).reset_index(drop=True)


def build_top_performer_table(df, entity_col, date_col):
    if not entity_col or entity_col not in df.columns:
        return pd.DataFrame()

    base = df[df[entity_col].notna()].copy()
    if base.empty:
        return pd.DataFrame()

    summary = base.groupby(entity_col).size().reset_index(name="Total Records")

    if "AI preference" in base.columns:
        summary = summary.merge(
            base[base["AI preference"].astype(str).str.lower().eq("presumptive")]
            .groupby(entity_col).size().reset_index(name="Presumptive"),
            on=entity_col, how="left"
        )

    if "Sent for testing" in base.columns:
        summary = summary.merge(
            base[base["Sent for testing"].astype(str).str.lower().eq("yes")]
            .groupby(entity_col).size().reset_index(name="Sent for Testing"),
            on=entity_col, how="left"
        )

    if "Type" in base.columns:
        summary = summary.merge(
            base[base["Type"].astype(str).str.lower().str.contains("symptomatic", na=False)]
            .groupby(entity_col).size().reset_index(name="Symptomatic"),
            on=entity_col, how="left"
        )

    if date_col and date_col in base.columns:
        latest_df = base.groupby(entity_col)[date_col].max().reset_index(name="Last Active Date")
        summary = summary.merge(latest_df, on=entity_col, how="left")

        ref_date = base[date_col].max()
        summary["Days Since Last Activity"] = (ref_date - summary["Last Active Date"]).dt.days

    summary = summary.fillna(0)

    for col in ["Presumptive", "Sent for Testing", "Symptomatic", "Days Since Last Activity"]:
        if col in summary.columns:
            summary[col] = pd.to_numeric(summary[col], errors="coerce").fillna(0)

    summary["Performance Score"] = (
        summary["Total Records"] * 1.0 +
        summary.get("Presumptive", 0) * 2.0 +
        summary.get("Sent for Testing", 0) * 2.5 +
        summary.get("Symptomatic", 0) * 1.5
    )

    summary["Performance Rank"] = summary["Performance Score"].rank(method="dense", ascending=False).astype(int)

    ordered_cols = [entity_col, "Performance Rank", "Performance Score", "Total Records"]
    for col in ["Presumptive", "Sent for Testing", "Symptomatic", "Last Active Date", "Days Since Last Activity"]:
        if col in summary.columns:
            ordered_cols.append(col)

    summary = summary[ordered_cols]
    return summary.sort_values(
        ["Performance Score", "Total Records"],
        ascending=[False, False]
    ).reset_index(drop=True)


def active_user_summary_table(df, entity_col, date_col, label="Entity"):
    summary = build_active_user_table(df, entity_col, date_col)
    if summary.empty:
        st.info(f"{label} active summary not available.")
        return
    st.dataframe(format_dates_for_display(summary), use_container_width=True, height=420)


def top_performer_table(df, entity_col, date_col, label="Entity"):
    summary = build_top_performer_table(df, entity_col, date_col)
    if summary.empty:
        st.info(f"{label} top performer summary not available.")
        return
    st.dataframe(format_dates_for_display(summary.head(20)), use_container_width=True, height=420)


def build_kpis(df, district_col, block_col, facility_col, cho_col):
    total_records = len(df)
    total_districts = df[district_col].nunique() if district_col else 0
    total_blocks = df[block_col].nunique() if block_col else 0
    total_facilities = df[facility_col].dropna().astype(str).nunique() if facility_col else 0
    total_chos = df[cho_col].nunique() if cho_col else 0

    presumptive = 0
    if "AI preference" in df.columns:
        presumptive = int(df["AI preference"].astype(str).str.lower().eq("presumptive").sum())

    tested = 0
    if "Sent for testing" in df.columns:
        tested = int(df["Sent for testing"].astype(str).str.lower().eq("yes").sum())

    symptomatic = 0
    if "Type" in df.columns:
        symptomatic = int(df["Type"].astype(str).str.lower().str.contains("symptomatic", na=False).sum())

    avg_age = round(df["Age"].dropna().mean(), 1) if "Age" in df.columns and df["Age"].notna().any() else 0

    return {
        "Total Records": total_records,
        "Total Districts": total_districts,
        "Total Blocks": total_blocks,
        "Total Facilities": total_facilities,
        "Total CHOs": total_chos,
        "Presumptive": presumptive,
        "Sent for Testing": tested,
        "Symptomatic": symptomatic,
        "Average Age": avg_age
    }


def plot_time_trend(df, date_col):
    if not date_col or df[date_col].dropna().empty:
        st.info("Date trend not available.")
        return

    trend = (
        df.dropna(subset=[date_col])
        .groupby(df[date_col].dt.date)
        .size()
        .reset_index(name="Records")
    )
    trend.columns = ["Date", "Records"]

    fig = px.area(
        trend,
        x="Date",
        y="Records",
        title="Registration Trend",
        line_shape="spline"
    )
    fig.update_traces(line_color="#1d6fdc", fillcolor="rgba(29,111,220,0.24)")
    fig.update_layout(template="plotly_white", height=380, margin=dict(l=10, r=10, t=55, b=10))
    st.plotly_chart(fig, use_container_width=True)


def plot_district_summary(df, district_col):
    if not district_col:
        st.info("District summary not available.")
        return

    top_district = (
        df.groupby(district_col)
        .size()
        .reset_index(name="Records")
        .sort_values("Records", ascending=False)
        .head(10)
    )

    fig = px.bar(
        top_district,
        x=district_col,
        y="Records",
        color="Records",
        text_auto=True,
        title="Top Districts by Records",
        color_continuous_scale="Blues"
    )
    fig.update_layout(
        template="plotly_white",
        height=380,
        margin=dict(l=10, r=10, t=55, b=10),
        coloraxis_showscale=False
    )
    st.plotly_chart(fig, use_container_width=True)


def plot_block_summary(df, block_col):
    if not block_col:
        st.info("Block summary not available.")
        return

    top_blocks = (
        df.groupby(block_col)
        .size()
        .reset_index(name="Records")
        .sort_values("Records", ascending=False)
        .head(12)
    )

    fig = px.bar(
        top_blocks,
        x="Records",
        y=block_col,
        orientation="h",
        color="Records",
        text_auto=True,
        title="Top Blocks by Records",
        color_continuous_scale="Tealgrn"
    )
    fig.update_layout(
        template="plotly_white",
        height=420,
        margin=dict(l=10, r=10, t=55, b=10),
        yaxis={"categoryorder": "total ascending"},
        coloraxis_showscale=False
    )
    st.plotly_chart(fig, use_container_width=True)


def plot_facility_summary(df, facility_col):
    if not facility_col:
        st.info("Facility summary not available.")
        return

    top_facility = (
        df.groupby(facility_col)
        .size()
        .reset_index(name="Records")
        .sort_values("Records", ascending=False)
        .head(12)
    )

    fig = px.bar(
        top_facility,
        x="Records",
        y=facility_col,
        orientation="h",
        color="Records",
        text_auto=True,
        title="Top Facilities / HWC by Records",
        color_continuous_scale="Purp"
    )
    fig.update_layout(
        template="plotly_white",
        height=420,
        margin=dict(l=10, r=10, t=55, b=10),
        yaxis={"categoryorder": "total ascending"},
        coloraxis_showscale=False
    )
    st.plotly_chart(fig, use_container_width=True)


def plot_cho_summary(df, cho_col):
    if not cho_col:
        st.info("CHO summary not available.")
        return

    top_cho = (
        df.groupby(cho_col)
        .size()
        .reset_index(name="Records")
        .sort_values("Records", ascending=False)
        .head(12)
    )

    fig = px.bar(
        top_cho,
        x="Records",
        y=cho_col,
        orientation="h",
        color="Records",
        text_auto=True,
        title="Top CHO by Records",
        color_continuous_scale="Sunsetdark"
    )
    fig.update_layout(
        template="plotly_white",
        height=420,
        margin=dict(l=10, r=10, t=55, b=10),
        yaxis={"categoryorder": "total ascending"},
        coloraxis_showscale=False
    )
    st.plotly_chart(fig, use_container_width=True)


def plot_gender(df):
    if "Gender" not in df.columns:
        st.info("Gender distribution not available.")
        return

    gender_df = df.groupby("Gender").size().reset_index(name="Records")
    fig = px.pie(
        gender_df,
        names="Gender",
        values="Records",
        hole=0.5,
        title="Gender Distribution",
        color_discrete_sequence=px.colors.sequential.Blues_r
    )
    fig.update_layout(template="plotly_white", height=380, margin=dict(l=10, r=10, t=55, b=10))
    st.plotly_chart(fig, use_container_width=True)


def plot_ai_preference(df):
    if "AI preference" not in df.columns:
        st.info("AI preference chart not available.")
        return

    pref_df = (
        df.groupby("AI preference")
        .size()
        .reset_index(name="Records")
        .sort_values("Records", ascending=False)
    )

    fig = px.bar(
        pref_df,
        x="AI preference",
        y="Records",
        color="Records",
        text_auto=True,
        title="AI Preference Distribution",
        color_continuous_scale="Viridis"
    )
    fig.update_layout(
        template="plotly_white",
        height=380,
        margin=dict(l=10, r=10, t=55, b=10),
        coloraxis_showscale=False
    )
    st.plotly_chart(fig, use_container_width=True)


def plot_age_distribution(df):
    if "Age" not in df.columns or df["Age"].dropna().empty:
        st.info("Age distribution not available.")
        return

    fig = px.histogram(
        df,
        x="Age",
        nbins=20,
        title="Age Distribution",
        color_discrete_sequence=["#1d6fdc"]
    )
    fig.update_layout(template="plotly_white", height=380, margin=dict(l=10, r=10, t=55, b=10))
    st.plotly_chart(fig, use_container_width=True)


def plot_active_summary_bar(facility_active, cho_active):
    chart_df = pd.DataFrame({
        "Window": ["Weekly", "Biweekly", "Monthly", "Weekly", "Biweekly", "Monthly"],
        "Count": [
            facility_active["7d"], facility_active["15d"], facility_active["30d"],
            cho_active["7d"], cho_active["15d"], cho_active["30d"]
        ],
        "Type": ["Facility", "Facility", "Facility", "CHO", "CHO", "CHO"]
    })

    fig = px.bar(
        chart_df,
        x="Window",
        y="Count",
        color="Type",
        barmode="group",
        text_auto=True,
        title="Active CHO and Facility Status"
    )
    fig.update_layout(template="plotly_white", height=380, margin=dict(l=10, r=10, t=55, b=10))
    st.plotly_chart(fig, use_container_width=True)


def plot_top_performer_chart(df, entity_col, date_col, title):
    perf = build_top_performer_table(df, entity_col, date_col)
    if perf.empty:
        st.info(f"{title} not available.")
        return

    perf = perf.head(10)

    fig = px.bar(
        perf,
        x="Performance Score",
        y=entity_col,
        orientation="h",
        color="Performance Score",
        text_auto=".2f",
        title=title,
        color_continuous_scale="Bluered"
    )
    fig.update_layout(
        template="plotly_white",
        height=430,
        margin=dict(l=10, r=10, t=55, b=10),
        yaxis={"categoryorder": "total ascending"},
        coloraxis_showscale=False
    )
    st.plotly_chart(fig, use_container_width=True)


def blockwise_facility_table(df, block_col, facility_col):
    if not block_col or not facility_col:
        st.info("Blockwise facility summary not available.")
        return

    bf = (
        df.groupby(block_col)[facility_col]
        .nunique()
        .reset_index(name="Total Facilities")
        .sort_values("Total Facilities", ascending=False)
    )
    st.dataframe(bf, use_container_width=True, height=420)


def district_summary_table(df, district_col, block_col, facility_col):
    if not district_col:
        st.info("District summary not available.")
        return

    result = df.groupby(district_col).agg(total_records=(district_col, "count"))

    if block_col:
        result["total_blocks"] = df.groupby(district_col)[block_col].nunique()
    if facility_col:
        result["total_facilities"] = df.groupby(district_col)[facility_col].nunique()
    if "Age" in df.columns:
        result["avg_age"] = df.groupby(district_col)["Age"].mean().round(1)

    result = result.reset_index().sort_values("total_records", ascending=False)
    st.dataframe(result, use_container_width=True, height=420)


def show_data_quality(df):
    missing = df.isnull().sum().reset_index()
    missing.columns = ["Column", "Missing Values"]
    missing["Missing %"] = ((missing["Missing Values"] / len(df)) * 100).round(2)
    missing = missing.sort_values("Missing Values", ascending=False)
    st.dataframe(missing, use_container_width=True, height=500)


def make_pdf_table_from_df(df, title, elements, styles, max_rows=25):
    if not REPORTLAB_AVAILABLE:
        return

    elements.append(Paragraph(title, styles["Heading2"]))
    elements.append(Spacer(1, 0.12 * inch))

    if df is None or df.empty:
        elements.append(Paragraph("No data available.", styles["Normal"]))
        elements.append(Spacer(1, 0.2 * inch))
        return

    pdf_df = prepare_export_table(df, max_rows=max_rows)

    table_data = [list(pdf_df.columns)] + pdf_df.values.tolist()
    col_count = len(pdf_df.columns)
    col_widths = [max(1.0, 10.4 / max(col_count, 1)) * inch for _ in range(col_count)]

    table = Table(table_data, repeatRows=1, colWidths=col_widths)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#12355b")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#d0d7e8")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.HexColor("#eef4ff")]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 7),
        ("TOPPADDING", (0, 0), (-1, 0), 7),
    ]))

    elements.append(table)
    elements.append(Spacer(1, 0.25 * inch))


def to_pdf_download(
    df, sheet_name, district_col=None, block_col=None, facility_col=None, cho_col=None, date_col=None
):
    if not REPORTLAB_AVAILABLE:
        return None

    output = BytesIO()

    doc = SimpleDocTemplate(
        output,
        pagesize=landscape(A4),
        rightMargin=20,
        leftMargin=20,
        topMargin=20,
        bottomMargin=20
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="ReportSubText",
        parent=styles["Normal"],
        fontSize=9,
        textColor=colors.HexColor("#4b6480"),
        leading=12
    ))

    elements = []

    title = Paragraph("CATB Filtered Dashboard Report", styles["Title"])
    elements.append(title)
    elements.append(Spacer(1, 0.12 * inch))

    meta_text = f"Sheet: {sheet_name} | Filtered Rows: {len(df)} | Columns: {df.shape[1]}"
    elements.append(Paragraph(meta_text, styles["ReportSubText"]))
    elements.append(Spacer(1, 0.2 * inch))

    facility_active_df = build_active_user_table(df, facility_col, date_col) if facility_col and date_col else pd.DataFrame()
    cho_active_df = build_active_user_table(df, cho_col, date_col) if cho_col and date_col else pd.DataFrame()
    facility_perf_df = build_top_performer_table(df, facility_col, date_col) if facility_col else pd.DataFrame()
    cho_perf_df = build_top_performer_table(df, cho_col, date_col) if cho_col else pd.DataFrame()

    if facility_active_df.empty and cho_active_df.empty and facility_perf_df.empty and cho_perf_df.empty:
        elements.append(Paragraph("No derived summary data available for export.", styles["Normal"]))
    else:
        make_pdf_table_from_df(facility_active_df.head(25), "Facility Active Status", elements, styles, max_rows=25)
        make_pdf_table_from_df(cho_active_df.head(25), "CHO Active Status", elements, styles, max_rows=25)

        elements.append(PageBreak())

        make_pdf_table_from_df(facility_perf_df.head(20), "Top Performing Facilities", elements, styles, max_rows=20)
        make_pdf_table_from_df(cho_perf_df.head(20), "Top Performing CHOs", elements, styles, max_rows=20)

    elements.append(PageBreak())

    preview_cols = [
        col for col in [
            "Reg Date", "State", "District", "Block", "Facility Name", "HWC", "CHO name",
            "Gender", "Designation", "AI preference", "Ntep result", "Sent for testing"
        ] if col in df.columns
    ]
    if not preview_cols:
        preview_cols = df.columns.tolist()[:12]

    raw_preview = df[preview_cols].head(25).copy() if not df.empty else pd.DataFrame()
    make_pdf_table_from_df(raw_preview, "Filtered Raw Data Preview", elements, styles, max_rows=25)

    doc.build(elements)
    return output.getvalue()


def to_excel_download(df, district_col=None, block_col=None, facility_col=None, cho_col=None, date_col=None):
    output = BytesIO()

    facility_active_df = build_active_user_table(df, facility_col, date_col) if facility_col and date_col else pd.DataFrame()
    cho_active_df = build_active_user_table(df, cho_col, date_col) if cho_col and date_col else pd.DataFrame()
    facility_perf_df = build_top_performer_table(df, facility_col, date_col) if facility_col else pd.DataFrame()
    cho_perf_df = build_top_performer_table(df, cho_col, date_col) if cho_col else pd.DataFrame()

    export_raw_df = prepare_export_table(df)
    export_fac_active = prepare_export_table(facility_active_df)
    export_cho_active = prepare_export_table(cho_active_df)
    export_fac_perf = prepare_export_table(facility_perf_df)
    export_cho_perf = prepare_export_table(cho_perf_df)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_raw_df.to_excel(writer, index=False, sheet_name="Raw Data")

        if not export_fac_active.empty:
            export_fac_active.to_excel(writer, index=False, sheet_name="Facility Active Status")
        else:
            pd.DataFrame({"Message": ["Facility active summary not available."]}).to_excel(
                writer, index=False, sheet_name="Facility Active Status"
            )

        if not export_cho_active.empty:
            export_cho_active.to_excel(writer, index=False, sheet_name="CHO Active Status")
        else:
            pd.DataFrame({"Message": ["CHO active summary not available."]}).to_excel(
                writer, index=False, sheet_name="CHO Active Status"
            )

        if not export_fac_perf.empty:
            export_fac_perf.to_excel(writer, index=False, sheet_name="Top Facility Performer")
        else:
            pd.DataFrame({"Message": ["Top facility performer summary not available."]}).to_excel(
                writer, index=False, sheet_name="Top Facility Performer"
            )

        if not export_cho_perf.empty:
            export_cho_perf.to_excel(writer, index=False, sheet_name="Top CHO Performer")
        else:
            pd.DataFrame({"Message": ["Top CHO performer summary not available."]}).to_excel(
                writer, index=False, sheet_name="Top CHO Performer"
            )

    return output.getvalue()


def main():
    st.markdown("""
    <div class="main-title">
        <h1>CATB - Smart Excel Analysis Dashboard</h1>
        <p>Upload, filter, analyze, visualize, and export district, block, facility, and CHO-level insights in one modern dashboard.</p>
    </div>
    """, unsafe_allow_html=True)

    if not REPORTLAB_AVAILABLE:
        st.markdown(
            "<div class='warning-box'><b>PDF export is currently disabled.</b> The app is running normally, but the <code>reportlab</code> package is not installed in this deployment environment.</div>",
            unsafe_allow_html=True
        )

    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

    if uploaded_file is None:
        st.info("Please upload an Excel file to begin analysis.")
        return

    try:
        df, sheet_name = load_excel(uploaded_file)
    except Exception as e:
        st.error(f"Error loading file: {e}")
        return

    filtered_df, district_col, block_col, facility_col, cho_col, date_col = apply_filters(df)

    st.markdown(
        f"""
        <div class='summary-banner'>
            <span><b>Loaded Sheet:</b> {sheet_name}</span>
            <span><b>Total Rows:</b> {df.shape[0]:,}</span>
            <span><b>Total Columns:</b> {df.shape[1]:,}</span>
            <span><b>Filtered Rows:</b> {filtered_df.shape[0]:,}</span>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        "<p class='small-note'><b>Note:</b> The dashboard keeps all analysis views, while the filtered Excel/PDF report is designed to reduce repetition and focus on raw data, active status, and top performer summaries.</p>",
        unsafe_allow_html=True
    )

    kpis = build_kpis(filtered_df, district_col, block_col, facility_col, cho_col)
    facility_active = get_active_entity_counts(filtered_df, facility_col, date_col)
    cho_active = get_active_entity_counts(filtered_df, cho_col, date_col)

    facility_perf_df = build_top_performer_table(filtered_df, facility_col, date_col)
    cho_perf_df = build_top_performer_table(filtered_df, cho_col, date_col)

    st.markdown("<div class='section-title'>Dashboard Summary</div>", unsafe_allow_html=True)

    row1 = st.columns(5)
    row1[0].metric("Total Records", f"{kpis['Total Records']:,}")
    row1[1].metric("Total Districts", f"{kpis['Total Districts']:,}")
    row1[2].metric("Total Blocks", f"{kpis['Total Blocks']:,}")
    row1[3].metric("Total Facilities", f"{kpis['Total Facilities']:,}")
    row1[4].metric("Total CHOs", f"{kpis['Total CHOs']:,}")

    row2 = st.columns(4)
    row2[0].metric("Presumptive", f"{kpis['Presumptive']:,}")
    row2[1].metric("Sent for Testing", f"{kpis['Sent for Testing']:,}")
    row2[2].metric("Symptomatic", f"{kpis['Symptomatic']:,}")
    row2[3].metric("Average Age", f"{kpis['Average Age']}")

    st.markdown("<div class='section-title'>Active Status Snapshot</div>", unsafe_allow_html=True)

    row3 = st.columns(6)
    row3[0].metric("Weekly Active Facility", facility_active["7d"])
    row3[1].metric("Biweekly Active Facility", facility_active["15d"])
    row3[2].metric("Monthly Active Facility", facility_active["30d"])
    row3[3].metric("Weekly Active CHO", cho_active["7d"])
    row3[4].metric("Biweekly Active CHO", cho_active["15d"])
    row3[5].metric("Monthly Active CHO", cho_active["30d"])

    latest_dates = [x for x in [facility_active.get("latest_date"), cho_active.get("latest_date")] if x is not None]
    if latest_dates:
        last_data_date = max(latest_dates)
        st.markdown(
            f"""
            <div class='metric-ribbon'>
                <p><b>Reference Date:</b> Active status is calculated using the latest available registration date in filtered data: <b>{pd.to_datetime(last_data_date).date()}</b></p>
            </div>
            """,
            unsafe_allow_html=True
        )

    top1, top2 = st.columns(2)
    with top1:
        st.markdown("<div class='info-card'>", unsafe_allow_html=True)
        st.markdown("<div class='subsection-title'>Top CHO Performer</div>", unsafe_allow_html=True)
        if not cho_perf_df.empty and cho_col in cho_perf_df.columns:
            top_cho_name = cho_perf_df.iloc[0][cho_col]
            top_cho_score = round(float(cho_perf_df.iloc[0]["Performance Score"]), 2)
            top_cho_records = int(cho_perf_df.iloc[0]["Total Records"])
            st.metric("Best CHO", f"{top_cho_name}")
            st.caption(f"Performance Score: {top_cho_score} | Records: {top_cho_records}")
        else:
            st.info("CHO performer data not available.")
        st.markdown("</div>", unsafe_allow_html=True)

    with top2:
        st.markdown("<div class='info-card'>", unsafe_allow_html=True)
        st.markdown("<div class='subsection-title'>Top Facility Performer</div>", unsafe_allow_html=True)
        if not facility_perf_df.empty and facility_col in facility_perf_df.columns:
            top_fac_name = facility_perf_df.iloc[0][facility_col]
            top_fac_score = round(float(facility_perf_df.iloc[0]["Performance Score"]), 2)
            top_fac_records = int(facility_perf_df.iloc[0]["Total Records"])
            st.metric("Best Facility", f"{top_fac_name}")
            st.caption(f"Performance Score: {top_fac_score} | Records: {top_fac_records}")
        else:
            st.info("Facility performer data not available.")
        st.markdown("</div>", unsafe_allow_html=True)

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
        ["Overview", "Geography", "Operational View", "Active Users", "Top Performers", "Data Tables"]
    )

    with tab1:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("<div class='chart-card'>", unsafe_allow_html=True)
            plot_time_trend(filtered_df, date_col)
            st.markdown("</div>", unsafe_allow_html=True)
        with c2:
            st.markdown("<div class='chart-card'>", unsafe_allow_html=True)
            plot_gender(filtered_df)
            st.markdown("</div>", unsafe_allow_html=True)

        c3, c4 = st.columns(2)
        with c3:
            st.markdown("<div class='chart-card'>", unsafe_allow_html=True)
            plot_ai_preference(filtered_df)
            st.markdown("</div>", unsafe_allow_html=True)
        with c4:
            st.markdown("<div class='chart-card'>", unsafe_allow_html=True)
            plot_age_distribution(filtered_df)
            st.markdown("</div>", unsafe_allow_html=True)

    with tab2:
        c5, c6 = st.columns(2)
        with c5:
            st.markdown("<div class='chart-card'>", unsafe_allow_html=True)
            plot_district_summary(filtered_df, district_col)
            st.markdown("</div>", unsafe_allow_html=True)
        with c6:
            st.markdown("<div class='chart-card'>", unsafe_allow_html=True)
            plot_block_summary(filtered_df, block_col)
            st.markdown("</div>", unsafe_allow_html=True)

        c7, c8 = st.columns(2)
        with c7:
            st.markdown("<div class='table-card'>", unsafe_allow_html=True)
            st.subheader("District summary")
            district_summary_table(filtered_df, district_col, block_col, facility_col)
            st.markdown("</div>", unsafe_allow_html=True)
        with c8:
            st.markdown("<div class='table-card'>", unsafe_allow_html=True)
            st.subheader("Blockwise facility")
            blockwise_facility_table(filtered_df, block_col, facility_col)
            st.markdown("</div>", unsafe_allow_html=True)

    with tab3:
        c9, c10 = st.columns(2)
        with c9:
            st.markdown("<div class='chart-card'>", unsafe_allow_html=True)
            plot_facility_summary(filtered_df, facility_col)
            st.markdown("</div>", unsafe_allow_html=True)
        with c10:
            st.markdown("<div class='chart-card'>", unsafe_allow_html=True)
            plot_cho_summary(filtered_df, cho_col)
            st.markdown("</div>", unsafe_allow_html=True)

    with tab4:
        c11, c12 = st.columns(2)
        with c11:
            st.markdown("<div class='chart-card'>", unsafe_allow_html=True)
            plot_active_summary_bar(facility_active, cho_active)
            st.markdown("</div>", unsafe_allow_html=True)

        with c12:
            st.markdown("""
            <div class='logic-box'>
                <h4>Standard Active Logic</h4>
                <p><b>Weekly Active:</b> Last activity within 7 days</p>
                <p><b>Biweekly Active:</b> Last activity within 15 days</p>
                <p><b>Monthly Active:</b> Last activity within 30 days</p>
                <p><b>Inactive:</b> No activity in the last 30 days</p>
            </div>
            """, unsafe_allow_html=True)

        c13, c14 = st.columns(2)
        with c13:
            st.markdown("<div class='table-card'>", unsafe_allow_html=True)
            st.subheader("Facility active status")
            active_user_summary_table(filtered_df, facility_col, date_col, label="Facility")
            st.markdown("</div>", unsafe_allow_html=True)

        with c14:
            st.markdown("<div class='table-card'>", unsafe_allow_html=True)
            st.subheader("CHO active status")
            active_user_summary_table(filtered_df, cho_col, date_col, label="CHO")
            st.markdown("</div>", unsafe_allow_html=True)

    with tab5:
        c15, c16 = st.columns(2)
        with c15:
            st.markdown("<div class='chart-card'>", unsafe_allow_html=True)
            plot_top_performer_chart(filtered_df, cho_col, date_col, "Top 10 CHO Performers")
            st.markdown("</div>", unsafe_allow_html=True)

        with c16:
            st.markdown("<div class='chart-card'>", unsafe_allow_html=True)
            plot_top_performer_chart(filtered_df, facility_col, date_col, "Top 10 Facility Performers")
            st.markdown("</div>", unsafe_allow_html=True)

        c17, c18 = st.columns(2)
        with c17:
            st.markdown("<div class='table-card'>", unsafe_allow_html=True)
            st.subheader("Top CHO performers")
            top_performer_table(filtered_df, cho_col, date_col, label="CHO")
            st.markdown("</div>", unsafe_allow_html=True)

        with c18:
            st.markdown("<div class='table-card'>", unsafe_allow_html=True)
            st.subheader("Top Facility performers")
            top_performer_table(filtered_df, facility_col, date_col, label="Facility")
            st.markdown("</div>", unsafe_allow_html=True)

    with tab6:
        c19, c20 = st.columns(2)
        with c19:
            st.markdown("<div class='table-card'>", unsafe_allow_html=True)
            st.subheader("Data quality")
            show_data_quality(filtered_df)
            st.markdown("</div>", unsafe_allow_html=True)

        with c20:
            st.markdown("<div class='table-card'>", unsafe_allow_html=True)
            st.subheader("Filtered data preview")

            default_visible_columns = [
                col for col in [
                    "Reg Date", "State", "District", "Block", "HWC", "CHO name",
                    "Gender", "Designation", "AI preference", "Ntep result",
                    "Sent for testing", "TB status", "Type", "Age",
                    "ensemble_model_score", "height", "weight"
                ] if col in filtered_df.columns
            ]

            visible_columns = st.multiselect(
                "Select visible columns",
                options=filtered_df.columns.tolist(),
                default=default_visible_columns if default_visible_columns else filtered_df.columns.tolist(),
                help="Choose which fields to display in the preview table."
            )

            display_df = filtered_df[visible_columns] if visible_columns else filtered_df.copy()

            st.dataframe(
                format_dates_for_display(display_df),
                use_container_width=True,
                height=500
            )
            st.markdown("</div>", unsafe_allow_html=True)

    excel_data = to_excel_download(
        filtered_df,
        district_col=district_col,
        block_col=block_col,
        facility_col=facility_col,
        cho_col=cho_col,
        date_col=date_col
    )

    pdf_data = to_pdf_download(
        filtered_df,
        sheet_name=sheet_name,
        district_col=district_col,
        block_col=block_col,
        facility_col=facility_col,
        cho_col=cho_col,
        date_col=date_col
    )

    st.markdown("<div class='section-title'>Download Reports</div>", unsafe_allow_html=True)

    d1, d2 = st.columns(2)
    with d1:
        st.download_button(
            label="Download Filtered Excel Report",
            data=excel_data,
            file_name="CATB_Filtered_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with d2:
        if REPORTLAB_AVAILABLE and pdf_data is not None:
            st.download_button(
                label="Download Filtered PDF Report",
                data=pdf_data,
                file_name="CATB_Filtered_Report.pdf",
                mime="application/pdf"
            )
        else:
            st.button("PDF Export Unavailable", disabled=True)

    if not REPORTLAB_AVAILABLE:
        st.caption("To enable PDF export on Streamlit Cloud, add reportlab to your requirements.txt file.")


if __name__ == "__main__":
    main()