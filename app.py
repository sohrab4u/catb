import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import datetime

# ==========================================
# 1. PAGE SETUP & ADVANCED ENTERPRISE CSS
# ==========================================
st.set_page_config(
    page_title="Enterprise Health Screening MIS Portal",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Premium Global UI/UX Theme and Sidebar Compression Styles
st.markdown("""
    <style>
    /* Base Workspace Background */
    .main { background-color: #F8FAFC; }
    
    /* Top Header Fine-Tuning */
    .header-title { color: #0F172A; font-weight: 800; font-size: 32px !important; margin-top: -40px; margin-bottom: 5px; }
    .header-subtitle { color: #475569; font-size: 15px; margin-bottom: 25px; font-weight: 400; }
    
    /* Structural Typography styling */
    h2, h3 { color: #1E3A8A; font-weight: 700; margin-top: 20px; border-bottom: 2px solid #E2E8F0; padding-bottom: 8px; }
    
    /* High-Visibility Custom Glossy KPI Cards */
    .kpi-container {
        display: flex;
        flex-wrap: wrap;
        gap: 15px;
        margin-bottom: 20px;
    }
    .kpi-card {
        background: #FFFFFF;
        padding: 24px;
        border-radius: 12px;
        border-top: 6px solid #2563EB;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
        flex: 1;
        min-width: 220px;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .kpi-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04);
    }
    .kpi-title { font-size: 12px; font-weight: 700; color: #64748B; text-transform: uppercase; letter-spacing: 0.8px; }
    .kpi-value { font-size: 32px; font-weight: 800; color: #1E3A8A; margin-top: 6px; }
    
    /* High Contrast Professional Compact Sidebar */
    section[data-testid="stSidebar"] {
        background-color: #0F172A !important;
        border-right: 1px solid #1E293B;
    }
    
    /* CRITICAL OVERHAUL: Make file uploader compact and strip spacing */
    section[data-testid="stSidebar"] div[data-testid="stFileUploader"] {
        padding: 0px !important;
        margin-bottom: 10px !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stFileUploader"] section {
        padding: 8px 12px !important;
        background-color: #1E293B !important;
        border: 1px dashed #3B82F6 !important;
        border-radius: 8px !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stFileUploader"] section div {
        display: none !important; /* Hide drag-and-drop instructional texts */
    }
    section[data-testid="stSidebar"] div[data-testid="stFileUploader"] section button {
        width: 100% !important;
        margin: 0 !important;
        padding: 5px !important;
        font-size: 13px !important;
    }
    
    /* Eliminate vertical padding on sidebar content area */
    section[data-testid="stSidebar"] .block-container {
        padding-top: 1rem !important;
        padding-bottom: 1rem !important;
    }
    
    /* Text Color Force Enforcement inside Sidebar */
    section[data-testid="stSidebar"] h1, 
    section[data-testid="stSidebar"] h2, 
    section[data-testid="stSidebar"] h3, 
    section[data-testid="stSidebar"] h4, 
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] .stMarkdown,
    section[data-testid="stSidebar"] div[data-testid="stMarkdownContainer"] p {
        color: #F8FAFC !important;
    }
    
    /* Space-efficient Custom Sidebar Header Styling */
    .sidebar-header-custom {
        font-size: 16px !important;
        font-weight: 700 !important;
        color: #3B82F6 !important;
        margin-top: -10px !important;
        margin-bottom: 12px !important;
        letter-spacing: 0.5px;
        text-transform: uppercase;
        border-bottom: 1px solid #1E293B;
        padding-bottom: 6px;
    }

    /* Style the Sidebar inputs for compactness */
    section[data-testid="stSidebar"] div[data-testid="stWidgetLabel"] {
        margin-bottom: -2px !important;
    }
    
    section[data-testid="stSidebar"] div.stButton > button {
        background-color: #2563EB !important;
        color: #FFFFFF !important;
        border-radius: 8px !important;
        font-weight: 700 !important;
        border: none !important;
        box-shadow: 0 4px 6px -1px rgba(37, 99, 235, 0.3) !important;
        margin-top: 10px !important;
    }
    section[data-testid="stSidebar"] div.stButton > button:hover {
        background-color: #1D4ED8 !important;
    }
    
    /* Clean Table Views */
    .stDataFrame { background-color: #FFFFFF; border-radius: 8px; padding: 5px; }
    </style>
    """, unsafe_allow_html=True)

# Top Title Header Section (Rendered explicitly to avoid empty/blank space at top)
st.markdown('<div class="header-title">📋 CATB Screening Data Summary and Detailed Report (MIS)</div>', unsafe_allow_html=True)

# ==========================================
# 2. DATA PIPELINE & CLEANING VALIDATOR
# ==========================================
@st.cache_data(show_spinner=False)
def process_health_workbook(file_buffer):
    """
    Ingests uploaded screening datasets. Dynamically processes spreadsheet header metrics,
    normalizes column schemas across multiple sheets, and generates validation tracking flags.
    """
    all_sheets_map = {}
    sheet_quality_summaries = {}
    combined_records = []
    
    try:
        xls = pd.ExcelFile(file_buffer, engine='openpyxl')
        
        for sheet in xls.sheet_names:
            df_check = xls.parse(sheet, nrows=5, header=None)
            if df_check.empty:
                continue
            
            # Dynamic Row Skip Calculation to clean out messy/empty header lines
            skip_rows_computed = 0
            for idx, row in df_check.iterrows():
                if row.astype(str).str.contains("Enrollment Data|Screening|Reg Date|District|HWC", case=False).any():
                    skip_rows_computed = idx
                    break
            
            df = xls.parse(sheet, skiprows=skip_rows_computed)
            if df.empty:
                continue
                
            df.columns = [str(c).strip() for c in df.columns]
            
            # Healthcare MIS Schema Normalization Definitions
            attribute_standards = {
                'reg_date': ['Reg Date', 'date', 'screening_date', 'registration_date'],
                'district': ['District', 'dist', 'zila'],
                'block': ['Block', 'block_name', 'tehsil'],
                'facility': ['HWC', 'Facility', 'center', 'health_facility', 'facility_name'],
                'cho_name': ['CHO name', 'cho', 'operator', 'community_health_officer']
            }
            
            for standard_key, alternatives in attribute_standards.items():
                if standard_key not in df.columns:
                    for alt in alternatives:
                        if alt in df.columns:
                            df.rename(columns={alt: standard_key}, inplace=True)
                            break
                    if standard_key not in df.columns:
                        df[standard_key] = "Unknown"

            # Parse and Sanitize Datetime Indices
            df['reg_date'] = pd.to_datetime(df['reg_date'], errors='coerce')
            df['reg_date'] = df['reg_date'].fillna(pd.Timestamp(datetime.date.today()))
            
            # Clean text column capitalization rules
            for text_col in ['district', 'block', 'facility', 'cho_name']:
                df[text_col] = df[text_col].astype(str).str.strip().str.title()
                df[text_col] = df[text_col].replace({'Nan': 'Unknown', 'None': 'Unknown', '': 'Unknown'})
            
            # Quality Audit Checks: Empty spacing logs
            raw_unskipped = xls.parse(sheet, header=None)
            is_blank_series = raw_unskipped.isnull().all(axis=1)
            f_valid = raw_unskipped.first_valid_index()
            l_valid = raw_unskipped.last_valid_index()
            
            blank_rows_count = 0
            inline_anomalies = 0
            if f_valid is not None and l_valid is not None:
                blank_rows_count = is_blank_series.loc[f_valid:l_valid].sum()
                for i in range(f_valid, l_valid):
                    if is_blank_series.iloc[i] and not is_blank_series.iloc[i+1:].all():
                        inline_anomalies += 1

            total_cells = df.size
            missing_cells = df.isnull().sum().sum()
            completeness = ((total_cells - missing_cells) / total_cells) * 100 if total_cells > 0 else 0
            
            sheet_quality_summaries[sheet] = {
                'row_count': len(df),
                'blank_rows': int(blank_rows_count),
                'inline_anomalies': int(inline_anomalies),
                'duplicate_records': int(df.duplicated().sum()),
                'completeness_score': float(completeness),
                'missing_by_column': df.isnull().sum().to_dict()
            }
            
            all_sheets_map[sheet] = df
            combined_records.append(df)
            
        if not combined_records:
            return None, None, None
            
        master_df = pd.concat(combined_records, ignore_index=True)
        return master_df, all_sheets_map, sheet_quality_summaries
        
    except Exception as e:
        st.error(f"Workbook parsing engine failed: {str(e)}")
        return None, None, None

# ==========================================
# 3. INTERACTIVE HELPER UTILITIES
# ==========================================
def calculate_cho_tier(count, median_val):
    if count >= median_val * 1.5: return "Excellent"
    elif count >= median_val: return "Good"
    elif count >= median_val * 0.5: return "Average"
    else: return "Poor"

def build_excel_report(data_bundle):
    out = BytesIO()
    with pd.ExcelWriter(out, engine='xlsxwriter') as wr:
        for tab_name, dataframe in data_bundle.items():
            dataframe.to_excel(wr, sheet_name=tab_name[:31], index=False)
    return out.getvalue()

# ==========================================
# 4. SIDEBAR PANEL INTEGRATION (COMPACT UI)
# ==========================================
st.sidebar.markdown('<div class="sidebar-header-custom">📊 MIS Control Panel</div>', unsafe_allow_html=True)
source_file = st.sidebar.file_uploader("Upload Screening Excel Data File", type=["xlsx", "xls"])

if source_file is not None:
    # Processing state execution decoupled from UI block dialog box alerts
    master_df, sheet_dfs, quality_manifest = process_health_workbook(source_file)
        
    if master_df is not None:
        # Global Sidebar Dropdown Selectors
        st.sidebar.markdown("### 🔍 Filter Scope")
        
        abs_min_date = master_df['reg_date'].min().to_pydatetime()
        abs_max_date = master_df['reg_date'].max().to_pydatetime()
        
        start_select, end_select = st.sidebar.date_input(
            "Reporting Timeline",
            value=(abs_min_date, abs_max_date),
            min_value=abs_min_date,
            max_value=abs_max_date
        )
        
        # Slicing Master Frame
        working_df = master_df[
            (master_df['reg_date'].dt.date >= start_select) & 
            (master_df['reg_date'].dt.date <= end_select)
        ]
        
        # Dropdowns
        dist_opts = sorted(working_df['district'].unique())
        selected_districts = st.sidebar.multiselect("Districts Selection", options=dist_opts, placeholder="All Active")
        if selected_districts:
            working_df = working_df[working_df['district'].isin(selected_districts)]
            
        block_opts = sorted(working_df['block'].unique())
        selected_blocks = st.sidebar.multiselect("Blocks Selection", options=block_opts, placeholder="All Active")
        if selected_blocks:
            working_df = working_df[working_df['block'].isin(selected_blocks)]
            
        fac_opts = sorted(working_df['facility'].unique())
        selected_facilities = st.sidebar.multiselect("Facilities Selection", options=fac_opts, placeholder="All Active")
        if selected_facilities:
            working_df = working_df[working_df['facility'].isin(selected_facilities)]
            
        cho_opts = sorted(working_df['cho_name'].unique())
        selected_chos = st.sidebar.multiselect("CHO Personnel Selection", options=cho_opts, placeholder="All Active")
        if selected_chos:
            working_df = working_df[working_df['cho_name'].isin(selected_chos)]
            
        global_search = st.sidebar.text_input("📝 Custom String Match Box")
        if global_search:
            working_df = working_df[working_df.astype(str).apply(lambda row: row.str.contains(global_search, case=False).any(), axis=1)]

        # Workspace Tab Divisions
        tabs = st.tabs([
            "📊 Executive Summary",
            "🏢 District Breakdown",
            "🧱 Block Performance",
            "👩‍⚕️ CHO Tiers & Workloads",
            "🏥 Facility Contribution",
            "🛠️ Data Audit & Quality",
            "🧠 Operational Insights"
        ])

        # Core Metrics Computation for Visualizations and Cross-Referencing
        total_scr = len(working_df)
        dist_cnt = working_df['district'].nunique()
        blk_cnt = working_df['block'].nunique()
        fac_cnt = working_df['facility'].nunique()
        cho_cnt = working_df['cho_name'].nunique()
        
        avg_per_dist = total_scr / dist_cnt if dist_cnt > 0 else 0
        avg_per_blk = total_scr / blk_cnt if blk_cnt > 0 else 0
        avg_per_cho = total_scr / cho_cnt if cho_cnt > 0 else 0
        
        agg_cells = working_df.size
        agg_nulls = working_df.isnull().sum().sum()
        global_completeness = ((agg_cells - agg_nulls) / agg_cells) * 100 if agg_cells > 0 else 0
        missing_percentage = 100 - global_completeness

        # Aggregate Aggregations Dataframes
        dist_grp = working_df.groupby('district').agg(
            Screenings_Count=('district', 'count'),
            Assigned_CHOs=('cho_name', 'nunique'),
            Assigned_Facilities=('facility', 'nunique')
        ).sort_values(by='Screenings_Count', ascending=False).reset_index()
        dist_grp['Contribution_Share_%'] = (dist_grp['Screenings_Count'] / total_scr * 100) if total_scr > 0 else 0
        dist_grp['Performance_Rank'] = dist_grp['Screenings_Count'].rank(ascending=False, method='min')

        block_grp = working_df.groupby(['block', 'district']).agg(
            Screenings_Count=('block', 'count'),
            Unique_Facilities=('facility', 'nunique')
        ).sort_values(by='Screenings_Count', ascending=False).reset_index()

        cho_raw = working_df.groupby(['cho_name', 'district', 'facility']).size().reset_index(name='Total_Screenings')
        state_median = cho_raw['Total_Screenings'].median() if not cho_raw.empty else 0
        cho_raw['Workforce_Tier'] = cho_raw['Total_Screenings'].apply(lambda count: calculate_cho_tier(count, state_median))
        cho_raw = cho_raw.sort_values(by='Total_Screenings', ascending=False).reset_index(drop=True)

        fac_grp = working_df.groupby(['facility', 'district', 'block']).size().reset_index(name='Total_Screenings').sort_values(by='Total_Screenings', ascending=False).reset_index(drop=True)
        fac_grp['Contribution_Percentage'] = (fac_grp['Total_Screenings'] / total_scr * 100) if total_scr > 0 else 0

        # ==========================================
        # TAB 1: EXECUTIVE SUMMARY
        # ==========================================
        with tabs[0]:
            st.markdown("### 📈 Health Delivery Infrastructure - Core KPIs")
            st.markdown(f"""
            <div class="kpi-container">
                <div class="kpi-card" style="border-top-color: #2563EB;">
                    <div class="kpi-title">Total Screenings Processed</div>
                    <div class="kpi-value">{total_scr:,}</div>
                </div>
                <div class="kpi-card" style="border-top-color: #10B981;">
                    <div class="kpi-title">Active Districts</div>
                    <div class="kpi-value">{dist_cnt}</div>
                </div>
                <div class="kpi-card" style="border-top-color: #F59E0B;">
                    <div class="kpi-title">Total Blocks Covered</div>
                    <div class="kpi-value">{blk_cnt}</div>
                </div>
                <div class="kpi-card" style="border-top-color: #8B5CF6;">
                    <div class="kpi-title">Tracked Facilities</div>
                    <div class="kpi-value">{fac_cnt}</div>
                </div>
                <div class="kpi-card" style="border-top-color: #EC4899;">
                    <div class="kpi-title">Active CHOs</div>
                    <div class="kpi-value">{cho_cnt}</div>
                </div>
            </div>
            <div class="kpi-container">
                <div class="kpi-card" style="border-top-color: #06B6D4;">
                    <div class="kpi-title">Avg Output / District</div>
                    <div class="kpi-value">{avg_per_dist:,.1f}</div>
                </div>
                <div class="kpi-card" style="border-top-color: #3B82F6;">
                    <div class="kpi-title">Avg Output / Block</div>
                    <div class="kpi-value">{avg_per_blk:,.1f}</div>
                </div>
                <div class="kpi-card" style="border-top-color: #6366F1;">
                    <div class="kpi-title">Avg Output / CHO</div>
                    <div class="kpi-value">{avg_per_cho:,.1f}</div>
                </div>
                <div class="kpi-card" style="border-top-color: #15803D;">
                    <div class="kpi-title">Data Completeness Rate</div>
                    <div class="kpi-value">{global_completeness:.2f}%</div>
                </div>
                <div class="kpi-card" style="border-top-color: #B91C1C;">
                    <div class="kpi-title">Data Missingness Index</div>
                    <div class="kpi-value">{missing_percentage:.2f}%</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown("---")
            st.markdown("### 🕒 Daily Volumetric Longitudinal Screening Trend Profile")
            time_series = working_df.groupby(working_df['reg_date'].dt.date).size().reset_index(name='Volume')
            fig_time = px.area(time_series, x='reg_date', y='Volume', title="Screening Volume Curve Timeline Distribution",
                               labels={'reg_date': 'Timeline Window', 'Volume': 'Evaluated Cases'}, color_discrete_sequence=['#1E3A8A'])
            fig_time.update_layout(plot_bgcolor='white', hovermode='x unified')
            st.plotly_chart(fig_time, use_container_width=True)

        # ==========================================
        # TAB 2: DISTRICT BREAKDOWN (FIXED ERROR LAYER)
        # ==========================================
        with tabs[1]:
            st.markdown("### 🏢 Regional District Standings and Metric Contributions")
            col_d1, col_d2 = st.columns([3, 2])
            with col_d1:
                # FIXED: Switched from .style.background_gradient() to a native Streamlit Progress/Bar configuration to drop the Matplotlib runtime environment requirement.
                st.dataframe(
                    dist_grp, 
                    use_container_width=True,
                    column_config={
                        "Screenings_Count": st.column_config.ProgressColumn(
                            "Screenings Count",
                            help="Volume of total screening logs compiled per region",
                            format="%d",
                            min_value=0,
                            max_value=int(dist_grp["Screenings_Count"].max() if not dist_grp.empty else 100)
                        )
                    }
                )
            with col_d2:
                fig_dist = px.bar(dist_grp, x='district', y='Screenings_Count', title="District Output Volumes",
                                  labels={'district': 'District', 'Screenings_Count': 'Screening Counts'},
                                  color='Screenings_Count', color_continuous_scale=px.colors.sequential.Cividis)
                st.plotly_chart(fig_dist, use_container_width=True)

        # ==========================================
        # TAB 3: BLOCK PERFORMANCE
        # ==========================================
        with tabs[2]:
            st.markdown("### 🧱 Block-Level Operations Performance Summary")
            col_b1, col_b2 = st.columns([2, 3])
            with col_b1:
                st.dataframe(block_grp, use_container_width=True)
            with col_b2:
                fig_blk = px.treemap(block_grp.head(30), path=['district', 'block'], values='Screenings_Count',
                                     title="Top 30 Block Production Volumes Hierarchical Treemap Matrix")
                st.plotly_chart(fig_blk, use_container_width=True)

        # ==========================================
        # TAB 4: CHO TIERS & WORKLOADS
        # ==========================================
        with tabs[3]:
            st.markdown("### 👩‍⚕️ Community Health Officer (CHO) Performance Matrix")
            col_ch1, col_ch2 = st.columns(2)
            with col_ch1:
                st.markdown("🏆 **Top 10 High Performing CHOs**")
                st.dataframe(cho_raw.head(10), use_container_width=True)
                st.markdown("📉 **Bottom 10 Lowest Performing CHOs**")
                st.dataframe(cho_raw.tail(10), use_container_width=True)
            with col_ch2:
                fig_cho_pie = px.pie(
                    cho_raw, names='Workforce_Tier', title="Workforce Segmentation Analytics Breakdown",
                    color='Workforce_Tier',
                    color_discrete_map={'Excellent': '#22C55E', 'Good': '#3B82F6', 'Average': '#F59E0B', 'Poor': '#EF4444'},
                    hole=0.4
                )
                st.plotly_chart(fig_cho_pie, use_container_width=True)

        # ==========================================
        # TAB 5: FACILITY CONTRIBUTION
        # ==========================================
        with tabs[4]:
            st.markdown("### 🏥 Health Sub-Center (HWC) Yield Trackers")
            col_f1, col_f2 = st.columns([4, 3])
            with col_f1:
                st.dataframe(fac_grp, use_container_width=True)
            with col_f2:
                fig_fac = px.pie(fac_grp.head(15), values='Total_Screenings', names='facility', title="Yield Allocation of Top 15 Primary Centers", hole=0.4)
                st.plotly_chart(fig_fac, use_container_width=True)

        # ==========================================
        # TAB 6: DATA AUDIT & QUALITY
        # ==========================================
        with tabs[5]:
            st.markdown("### 🛠️ Structural Integrity and Audit Reports by Worksheet")
            for sheet_name, stats in quality_manifest.items():
                with st.expander(f"📋 Sheet Log Audit Diagnostics: {sheet_name}", expanded=True):
                    q_c1, q_c2, q_c3, q_c4 = st.columns(4)
                    q_c1.metric("Rows Scanned", f"{stats['row_count']:,}")
                    q_c2.metric("Isolated Blank Rows", stats['blank_rows'])
                    q_c3.metric("Inline Entry Dropouts", stats['inline_anomalies'])
                    q_c4.metric("Absolute Duplicate Rows", stats['duplicate_records'])
            
            st.markdown("---")
            st.markdown("### 🔍 Missing Fields Grid Heatmap Profile (500 Row Sample Matrix)")
            sample_size = min(len(working_df), 500)
            if sample_size > 0:
                heat_matrix = working_df.isnull().astype(int).sample(sample_size).values
                fig_hm = px.imshow(heat_matrix, aspect='auto', color_continuous_scale=['#E2E8F0', '#EF4444'],
                                   labels=dict(x="Attribute Axis", y="Sample Log Index"))
                fig_hm.update_layout(coloraxis_showscale=False)
                st.plotly_chart(fig_hm, use_container_width=True)

        # ==========================================
        # TAB 7: OPERATIONAL INSIGHTS
        # ==========================================
        with tabs[6]:
            st.markdown("### 🧠 Automated Health Infrastructure Diagnostics Engine")
            mean_dist_val = dist_grp['Screenings_Count'].mean() if not dist_grp.empty else 0
            
            ins_col, rec_col = st.columns(2)
            with ins_col:
                st.markdown("#### 🏢 Identified Delivery Bottlenecks")
                for idx, r in dist_grp.iterrows():
                    if r['Screenings_Count'] < (mean_dist_val * 0.4):
                        st.error(f"⚠️ **Underperforming Region:** District **{r['district']}** is operating below 40% of the state output mean, with a total yield of **{r['Screenings_Count']} cases**.")
                for idx, r in cho_raw.head(5).iterrows():
                    if r['Total_Screenings'] < 5:
                        st.warning(f"📉 **Low Activity Alert:** CHO **{r['cho_name']}** assigned to facility **{r['facility']}** tracking low throughput cases.")
                if not dist_grp.empty:
                    st.success(f"🌟 **State Benchmark:** District **{dist_grp.iloc[0]['district']}** leads metrics state-wide, contributing **{dist_grp.iloc[0]['Contribution_Share_%']:.1f}%** of total screenings.")

            with rec_col:
                st.markdown("#### 🛠️ Direct Strategic Action Items")
                st.info("👉 **Workforce Balancing:** Deploy roaming support health vectors to low performing blocks.")
                st.info("👉 **Validation Enforcement:** Mandate entry constraints inside local apps to scrub out data dropouts.")
                st.info("👉 **Targeted Expansion:** Launch field camps in sub-district blocks underperforming relative to regional medians.")

        # ==========================================
        # 6. CENTRAL EXPORT MANAGEMENT CONTROLS (EXPANDED TABS)
        # ==========================================
        st.sidebar.markdown("---")
        st.sidebar.markdown("### 📥 Reports Generation")
        
        # Comprehensive Data Sheet Compilation Pipeline
        sheet_wise_audit_list = []
        for sh_name, stats in quality_manifest.items():
            sheet_wise_audit_list.append({
                "Sheet Name": sh_name,
                "Total Rows": stats['row_count'],
                "Blank Rows": stats['blank_rows'],
                "Inline Sequential Anomalies": stats['inline_anomalies'],
                "Duplicate Log Rows": stats['duplicate_records'],
                "Completeness Score %": stats['completeness_score']
            })
        df_sheet_audit_report = pd.DataFrame(sheet_wise_audit_list)

        column_wise_nulls = pd.DataFrame(
            list(working_df.isnull().sum().items()),
            columns=['Data Field Attribute', 'Total Null Fields Collected']
        ).sort_values(by='Total Null Fields Collected', ascending=False)

        management_insights_summary = pd.DataFrame([
            {"Metric Indicator Summary": "Global System Completeness Percentage", "Value Metric": f"{global_completeness:.2f}%"},
            {"Metric Indicator Summary": "Data Missingness Index Profile", "Value Metric": f"{missing_percentage:.2f}%"},
            {"Metric Indicator Summary": "Top Performing District System Model", "Value Metric": dist_grp.iloc[0]['district'] if not dist_grp.empty else "N/A"},
            {"Metric Indicator Summary": "Total Districts Below Operational Benchmark", "Value Metric": len(dist_grp[dist_grp['Screenings_Count'] < (mean_dist_val * 0.4)]) if not dist_grp.empty else 0},
            {"Metric Indicator Summary": "Workforce Size in Critical Care Poor Cohort", "Value Metric": len(cho_raw[cho_raw['Workforce_Tier'] == 'Poor']) if not cho_raw.empty else 0}
        ])

        # Expanded Multi-Tab Bundle Blueprint
        bundle = {
            "Executive_KPI_Summary": pd.DataFrame([{
                "Total Screenings": total_scr, "Unique Districts": dist_cnt, "Unique Blocks": blk_cnt, 
                "Active Health Centers": fac_cnt, "Active CHOs": cho_cnt, "Data Completeness %": f"{global_completeness:.2f}%"
            }]),
            "District_Performance_Matrix": dist_grp,
            "Local_Block_Performance_Rank": block_grp,
            "CHO_Workforce_Tiers": cho_raw,
            "Primary_Facility_Outputs": fac_grp,
            "Spreadsheet_Layer_Audits": df_sheet_audit_report,
            "Field_Missingness_Audits": column_wise_nulls,
            "Strategic_System_Insights": management_insights_summary,
            "Master_Cleaned_Records_Log": working_df.astype(str).head(1000) # Capped at 1k rows to ensure export stability
        }
        
        excel_bytes = build_excel_report(bundle)
        st.sidebar.download_button(
            label="📊 Download MIS Bundle",
            data=excel_bytes,
            file_name=f"Health_Screening_MIS_Package_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
else:
    st.info("💡 Ingestion Queue Empty: Please upload your healthcare multi-sheet data file into the source buffer located in the left dark dashboard menu to initialize analytics.")