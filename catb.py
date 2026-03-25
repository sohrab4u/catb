import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO

st.set_page_config(
    page_title="CATB Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
[data-testid="stAppViewContainer"] {
    background: linear-gradient(180deg, #f4f7fb 0%, #eef4ff 100%);
}
[data-testid="stHeader"] {
    background: rgba(0,0,0,0);
}
.block-container {
    padding-top: 1rem;
    padding-bottom: 1rem;
    max-width: 98%;
}
.main-title {
    padding: 1.2rem 1.4rem;
    border-radius: 18px;
    background: linear-gradient(135deg, #12355b 0%, #1d6fdc 100%);
    color: white;
    box-shadow: 0 10px 30px rgba(18,53,91,0.18);
    margin-bottom: 1rem;
}
.main-title h1 {
    margin: 0;
    padding: 0;
    color: white;
    font-size: 2.1rem;
}
.main-title p {
    margin: 0.35rem 0 0 0;
    color: #e9f2ff;
}
div[data-testid="stMetric"] {
    background: linear-gradient(135deg, #ffffff, #f8fbff);
    border: 1px solid #d7e6ff;
    padding: 14px;
    border-radius: 16px;
    box-shadow: 0 8px 24px rgba(15, 40, 80, 0.07);
}
div[data-testid="stMetricLabel"] {
    color: #46607a;
    font-weight: 600;
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
    background: rgba(255,255,255,0.08);
    padding: 12px;
    border-radius: 12px;
    margin-bottom: 12px;
    border: 1px solid rgba(255,255,255,0.12);
}
.chart-card {
    background: white;
    border-radius: 18px;
    padding: 10px 14px 2px 14px;
    box-shadow: 0 8px 24px rgba(15, 40, 80, 0.06);
    border: 1px solid #ebf1fb;
    margin-bottom: 12px;
}
.table-card {
    background: white;
    border-radius: 18px;
    padding: 12px 14px;
    box-shadow: 0 8px 24px rgba(15, 40, 80, 0.06);
    border: 1px solid #ebf1fb;
    margin-bottom: 12px;
}
.small-note {
    color: #58708a;
    font-size: 0.95rem;
}
.filter-box {
    background: rgba(255,255,255,0.08);
    padding: 10px;
    border-radius: 10px;
    margin-bottom: 10px;
    border: 1px solid rgba(255,255,255,0.10);
}
</style>
""", unsafe_allow_html=True)

EXPECTED_DATE_COLS = ["Reg Date", "tbdiagnosisdate"]
FACILITY_CANDIDATES = ["HWC ID", "Facility", "HWC Name", "Subcenter"]
BLOCK_CANDIDATES = ["Block"]
DISTRICT_CANDIDATES = ["District"]
CHO_CANDIDATES = ["CHO name"]

@st.cache_data
def load_excel(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = xls.sheet_names[0]
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    df.columns = [str(c).strip() for c in df.columns]

    for col in EXPECTED_DATE_COLS:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace({
                "nan": np.nan,
                "None": np.nan,
                "null": np.nan,
                "": np.nan
            })

    numeric_candidates = [
        "Age", "ensemblemodelscore", "height", "weight",
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

def safe_unique(series):
    return sorted([x for x in series.dropna().astype(str).unique().tolist()])

def apply_filters(df):
    filtered = df.copy()

    district_col = find_column(filtered, DISTRICT_CANDIDATES)
    block_col = find_column(filtered, BLOCK_CANDIDATES)
    facility_col = find_column(filtered, FACILITY_CANDIDATES)
    cho_col = find_column(filtered, CHO_CANDIDATES)
    date_col = "Reg Date" if "Reg Date" in filtered.columns else None

    st.sidebar.markdown(
        "<div class='sidebar-card'><h3 style='margin-top:0;'>Filters</h3><p style='margin-bottom:0;'>Refine dashboard by geography, facility, CHO, and date.</p></div>",
        unsafe_allow_html=True
    )

    if district_col:
        districts = safe_unique(filtered[district_col])
        selected_districts = st.sidebar.multiselect(
            "District",
            districts,
            default=districts,
            help="By default all districts are selected."
        )
        if selected_districts:
            filtered = filtered[filtered[district_col].astype(str).isin(selected_districts)]

    if block_col:
        blocks = safe_unique(filtered[block_col])
        selected_blocks = st.sidebar.multiselect(
            "Block",
            blocks,
            default=blocks,
            help="By default all blocks are selected."
        )
        if selected_blocks:
            filtered = filtered[filtered[block_col].astype(str).isin(selected_blocks)]

    if facility_col:
        facilities = safe_unique(filtered[facility_col])
        selected_facilities = st.sidebar.multiselect(
            "Facility / HWC",
            facilities,
            default=facilities,
            help="By default all facilities are selected."
        )
        if selected_facilities:
            filtered = filtered[filtered[facility_col].astype(str).isin(selected_facilities)]

    if cho_col:
        chos = safe_unique(filtered[cho_col])
        selected_chos = st.sidebar.multiselect(
            "CHO Name",
            chos,
            default=chos,
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
                max_value=max_date,
                help="Default is minimum available date."
            )

        with d2:
            end_date = st.date_input(
                "End Date",
                value=max_date,
                min_value=min_date,
                max_value=max_date,
                help="Default is maximum available date."
            )

        if start_date > end_date:
            st.sidebar.error("Start Date cannot be greater than End Date.")
        else:
            filtered = filtered[
                (filtered[date_col].dt.date >= start_date) &
                (filtered[date_col].dt.date <= end_date)
            ]

    return filtered, district_col, block_col, facility_col, cho_col, date_col

def to_excel_download(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Filtered_Data")
    return output.getvalue()

def build_kpis(df, district_col, block_col, facility_col, cho_col):
    total_records = len(df)
    total_districts = df[district_col].nunique() if district_col else 0
    total_blocks = df[block_col].nunique() if block_col else 0
    total_facilities = df[facility_col].nunique() if facility_col else 0
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
    fig.update_traces(line_color="#1d6fdc", fillcolor="rgba(29,111,220,0.22)")
    fig.update_layout(template="plotly_white", height=380, margin=dict(l=10, r=10, t=50, b=10))
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
        margin=dict(l=10, r=10, t=50, b=10),
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
        margin=dict(l=10, r=10, t=50, b=10),
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
        margin=dict(l=10, r=10, t=50, b=10),
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
        margin=dict(l=10, r=10, t=50, b=10),
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
    fig.update_layout(template="plotly_white", height=380, margin=dict(l=10, r=10, t=50, b=10))
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
        margin=dict(l=10, r=10, t=50, b=10),
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
    fig.update_layout(template="plotly_white", height=380, margin=dict(l=10, r=10, t=50, b=10))
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

def main():
    st.markdown("""
    <div class="main-title">
        <h1>CATB - Smart Excel Analysis Dashboard</h1>
        <p>Upload, filter, analyze, visualize, and export district-block-facility level data in one place.</p>
    </div>
    """, unsafe_allow_html=True)

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
        f"<p class='small-note'><b>Loaded Sheet:</b> {sheet_name} &nbsp;&nbsp;|&nbsp;&nbsp; <b>Rows:</b> {df.shape[0]:,} &nbsp;&nbsp;|&nbsp;&nbsp; <b>Columns:</b> {df.shape[1]:,} &nbsp;&nbsp;|&nbsp;&nbsp; <b>Filtered Rows:</b> {filtered_df.shape[0]:,}</p>",
        unsafe_allow_html=True
    )

    kpis = build_kpis(filtered_df, district_col, block_col, facility_col, cho_col)

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

    tab1, tab2, tab3, tab4 = st.tabs(["Overview", "Geography", "Operational View", "Data Tables"])

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
            st.markdown("<div class='table-card'>", unsafe_allow_html=True)
            st.subheader("Data quality")
            show_data_quality(filtered_df)
            st.markdown("</div>", unsafe_allow_html=True)
        with c12:
            st.markdown("<div class='table-card'>", unsafe_allow_html=True)
            st.subheader("Filtered data preview")

            default_visible_columns = [
                col for col in [
                    "Reg Date", "State", "District", "Block", "HWC ID", "CHO name",
                    "Gender", "Designation", "AI preference", "Ntep result",
                    "Sent for testing", "TB status", "Type", "Age",
                    "ensemblemodelscore", "height", "weight"
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
                display_df,
                use_container_width=True,
                height=500
            )
            st.markdown("</div>", unsafe_allow_html=True)

    download_data = to_excel_download(filtered_df)
    st.download_button(
        label="Download Filtered Data",
        data=download_data,
        file_name="CATB_Filtered_Data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    main()