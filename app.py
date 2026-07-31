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
    page_title="Enterprise Health Screening MIS & Presumptive Analytics Portal",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main { background-color: #F8FAFC; }
    .header-title { color: #0F172A; font-weight: 800; font-size: 30px !important; margin-top: -30px; margin-bottom: 5px; }
    .header-subtitle { color: #475569; font-size: 14px; margin-bottom: 20px; font-weight: 500; }
    h2, h3 { color: #1E3A8A; font-weight: 700; margin-top: 15px; border-bottom: 2px solid #E2E8F0; padding-bottom: 6px; }
    
    /* KPI Styling */
    .kpi-container { display: flex; flex-wrap: wrap; gap: 12px; margin-bottom: 15px; }
    .kpi-card {
        background: #FFFFFF;
        padding: 18px;
        border-radius: 10px;
        border-top: 5px solid #2563EB;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
        flex: 1;
        min-width: 180px;
        transition: transform 0.2s ease;
    }
    .kpi-card:hover { transform: translateY(-3px); }
    .kpi-title { font-size: 11px; font-weight: 700; color: #64748B; text-transform: uppercase; letter-spacing: 0.5px; }
    .kpi-value { font-size: 26px; font-weight: 800; color: #1E3A8A; margin-top: 4px; }
    .kpi-subtext { font-size: 11px; color: #059669; font-weight: 600; margin-top: 2px; }

    /* Compact Sidebar Styling */
    section[data-testid="stSidebar"] { background-color: #0F172A !important; border-right: 1px solid #1E293B; }
    section[data-testid="stSidebar"] div[data-testid="stFileUploader"] { padding: 0px !important; margin-bottom: 10px !important; }
    section[data-testid="stSidebar"] div[data-testid="stFileUploader"] section { padding: 8px 12px !important; background-color: #1E293B !important; border: 1px dashed #3B82F6 !important; border-radius: 8px !important; }
    section[data-testid="stSidebar"] div[data-testid="stFileUploader"] section div { display: none !important; }
    section[data-testid="stSidebar"] div[data-testid="stFileUploader"] section button { width: 100% !important; margin: 0 !important; padding: 4px !important; font-size: 12px !important; }
    section[data-testid="stSidebar"] .block-container { padding-top: 1rem !important; padding-bottom: 1rem !important; }
    section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3, 
    section[data-testid="stSidebar"] h4, section[data-testid="stSidebar"] p, section[data-testid="stSidebar"] label, 
    section[data-testid="stSidebar"] .stMarkdown, section[data-testid="stSidebar"] div[data-testid="stMarkdownContainer"] p { color: #F8FAFC !important; }
    
    .sidebar-header-custom {
        font-size: 15px !important; font-weight: 700 !important; color: #3B82F6 !important;
        margin-top: -10px !important; margin-bottom: 10px !important; text-transform: uppercase;
        border-bottom: 1px solid #1E293B; padding-bottom: 4px;
    }
    section[data-testid="stSidebar"] div.stButton > button {
        background-color: #2563EB !important; color: #FFFFFF !important; border-radius: 6px !important;
        font-weight: 700 !important; border: none !important; margin-top: 8px !important;
    }
    .stDataFrame { background-color: #FFFFFF; border-radius: 8px; padding: 4px; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="header-title">📋 Enterprise Health Screening & Nikshay ID MIS Portal</div>', unsafe_allow_html=True)
st.markdown('<div class="header-subtitle">District & Block-level Presumptive Screening, HWC Overrules, and Nikshay ID Generation Performance Dashboard</div>', unsafe_allow_html=True)

# ==========================================
# 2. DATA PIPELINE & SCHEMATIC CLEANER
# ==========================================
@st.cache_data(show_spinner=False)
def process_health_workbook(file_buffer):
    try:
        xls = pd.ExcelFile(file_buffer, engine='openpyxl')
        combined_records = []
        
        for sheet in xls.sheet_names:
            df = xls.parse(sheet)
            if df.empty:
                continue
            
            df.columns = [str(c).strip() for c in df.columns]
            
            # Healthcare MIS Schema Mapping
            attribute_standards = {
                'reg_date': ['Reg Date', 'date', 'registration_date'],
                'district': ['District', 'dist', 'zila'],
                'block': ['Block', 'block_name'],
                'facility': ['HWC', 'Facility', 'center', 'health_facility'],
                'cho_name': ['CHO name', 'cho', 'operator'],
                'patient_id': ['ID', 'uuid', 'id', 'Screening name'],
                'gender': ['Gender', 'sex'],
                'age': ['Age', 'age'],
                'designation': ['Designation', 'role'],
                'ai_preference': ['AI preference', 'AI_preference', 'ai_result'],
                'ntep_result': ['Ntep result', 'NTEP_result', 'ntep_status'],
                'nikshay_id': ['nikshay_id', 'Nikshay ID', 'X', 'nikshay'],
                'overrule_hwc': ['Overrule by HWC', 'BB', 'overrule_by_hwc']
            }
            
            for standard_key, alternatives in attribute_standards.items():
                if standard_key not in df.columns:
                    for alt in alternatives:
                        if alt in df.columns:
                            df.rename(columns={alt: standard_key}, inplace=True)
                            break
                    if standard_key not in df.columns:
                        df[standard_key] = np.nan

            # Date Normalization
            df['reg_date'] = pd.to_datetime(df['reg_date'], errors='coerce')
            df['reg_date'] = df['reg_date'].fillna(pd.Timestamp(datetime.date.today()))
            
            # String Cleaning
            for text_col in ['district', 'block', 'facility', 'cho_name']:
                df[text_col] = df[text_col].astype(str).str.strip().str.title()
                df[text_col] = df[text_col].replace({'Nan': 'Unknown', 'None': 'Unknown', '': 'Unknown'})
            
            # ACCURATE NIKSHAY ID COUNT CHECK
            # Check if nikshay_id column has an entry (ignoring null, empty string, nan, none, 0)
            nik_str = df['nikshay_id'].astype(str).str.strip().str.lower()
            df['has_nikshay'] = df['nikshay_id'].notnull() & \
                               (nik_str != '') & \
                               (nik_str != 'nan') & \
                               (nik_str != 'none') & \
                               (nik_str != 'null') & \
                               (nik_str != '0') & \
                               (nik_str != '0.0')

            # Standardize Presumptive Flags
            df['ai_presumptive_flag'] = df['ai_preference'].astype(str).str.lower().str.contains('presumptive', na=False) & \
                                        ~df['ai_preference'].astype(str).str.lower().str.contains('non', na=False)
            
            df['ntep_presumptive_flag'] = df['ntep_result'].astype(str).str.lower().str.contains('presumptive', na=False) & \
                                          ~df['ntep_result'].astype(str).str.lower().str.contains('non', na=False)
            
            # Total Presumptive status flag (Identified as presumptive by NTEP or AI)
            df['is_presumptive'] = df['ntep_presumptive_flag'] | df['ai_presumptive_flag']
            
            df['overrule_flag'] = pd.to_numeric(df['overrule_hwc'], errors='coerce').fillna(0).astype(int) == 1
            
            # Calculations for Nikshay Pending (Presumptive cases where Nikshay ID is not yet generated)
            df['nikshay_pending_presumptive'] = df['is_presumptive'] & (~df['has_nikshay'])
            
            combined_records.append(df)
            
        if not combined_records:
            return None
            
        master_df = pd.concat(combined_records, ignore_index=True)
        return master_df
    except Exception as e:
        st.error(f"Data engine error: {str(e)}")
        return None

# Multi-Tab Dynamic District Excel Exporter
def generate_district_excel_bundle(master_df, cho_summary_df, patient_master_df):
    out = BytesIO()
    with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
        # Sheet 1: CHO Operational Summary
        cho_summary_df.to_excel(writer, sheet_name='CHO_Operational_Summary', index=False)
        
        # Sheet 2: Complete Patient-Level Master Data
        patient_master_df.to_excel(writer, sheet_name='Patient_Level_Master_Log', index=False)
        
        # Individual Sheet for Each District
        unique_districts = sorted(master_df['district'].unique())
        for dist in unique_districts:
            clean_dist_name = str(dist)[:30].replace('/', '_').replace('\\', '_')
            dist_subset = patient_master_df[patient_master_df['District'] == dist].copy()
            dist_subset.to_excel(writer, sheet_name=clean_dist_name, index=False)
            
    return out.getvalue()

# ==========================================
# 3. SIDEBAR & FILTERS
# ==========================================
st.sidebar.markdown('<div class="sidebar-header-custom">📊 MIS Control Panel</div>', unsafe_allow_html=True)
source_file = st.sidebar.file_uploader("Upload Screening Excel Data File", type=["xlsx", "xls"])

if source_file is not None:
    master_df = process_health_workbook(source_file)
    
    if master_df is not None:
        st.sidebar.markdown("### 🔍 Filter Controls")
        
        abs_min_date = master_df['reg_date'].min().to_pydatetime()
        abs_max_date = master_df['reg_date'].max().to_pydatetime()
        
        start_select, end_select = st.sidebar.date_input(
            "Reporting Timeline",
            value=(abs_min_date, abs_max_date),
            min_value=abs_min_date,
            max_value=abs_max_date
        )
        
        # Apply Date Filter
        working_df = master_df[
            (master_df['reg_date'].dt.date >= start_select) & 
            (master_df['reg_date'].dt.date <= end_select)
        ].copy()
        
        # Multi-select Filters
        dist_opts = sorted(working_df['district'].unique())
        selected_districts = st.sidebar.multiselect("District Scope", options=dist_opts)
        if selected_districts:
            working_df = working_df[working_df['district'].isin(selected_districts)]
            
        block_opts = sorted(working_df['block'].unique())
        selected_blocks = st.sidebar.multiselect("Block Scope", options=block_opts)
        if selected_blocks:
            working_df = working_df[working_df['block'].isin(selected_blocks)]
            
        global_search = st.sidebar.text_input("📝 Search (ID / CHO / Facility)")
        if global_search:
            working_df = working_df[working_df.astype(str).apply(lambda row: row.str.contains(global_search, case=False).any(), axis=1)]

        # Core Metrics Computation
        total_scr = len(working_df)
        total_pres_cnt = working_df['is_presumptive'].sum()
        total_overrule_cnt = working_df['overrule_flag'].sum()
        nikshay_gen_cnt = working_df['has_nikshay'].sum()  # Total Nikshay IDs present in nikshay_id column
        nikshay_pend_cnt = working_df['nikshay_pending_presumptive'].sum() # Presumptive cases pending Nikshay ID

        # ==========================================
        # CHO OPERATIONAL SUMMARY TABLE
        # ==========================================
        cho_operational_summary = working_df.groupby(['district', 'block', 'cho_name']).agg(
            Total_Count=('district', 'count'),
            Total_Presumptive_Count=('is_presumptive', 'sum'),
            Total_Override_Count=('overrule_flag', 'sum'),
            Nikshay_ID_Generated_Count=('has_nikshay', 'sum'),  # DIRECT COUNT OF NIKSHAY ID COLUMN ENTRY
            Nikshay_ID_Pending_Count=('nikshay_pending_presumptive', 'sum')
        ).reset_index()

        cho_operational_summary.columns = [
            'District Name', 'Block Name', 'CHO Name', 'Total Count', 
            'Total Presumptive Count', 'Total Override Count', 
            'Total Nikshay ID Generated Count', 'Total Nikshay ID Generation Pending Count'
        ]
        cho_operational_summary = cho_operational_summary.sort_values(by='Total Count', ascending=False)

        # ==========================================
        # PATIENT-WISE COMPLETE LINE-LIST REPORT
        # ==========================================
        patient_line_list = working_df.copy()
        patient_line_list['Nikshay_Status'] = np.where(
            patient_line_list['has_nikshay'], 'Generated', 
            np.where(patient_line_list['is_presumptive'], 'Pending', 'Not Applicable')
        )
        
        patient_report_cols = [
            'reg_date', 'district', 'block', 'facility', 'cho_name', 'patient_id', 
            'age', 'gender', 'ai_preference', 'ntep_result', 'overrule_flag', 
            'has_nikshay', 'nikshay_id', 'Nikshay_Status'
        ]
        available_p_cols = [c for c in patient_report_cols if c in patient_line_list.columns]
        
        patient_master_df = patient_line_list[available_p_cols].copy()
        patient_master_df.columns = [
            'Registration Date', 'District', 'Block', 'Facility (HWC)', 'CHO Name', 
            'Patient ID / Name', 'Age', 'Gender', 'AI Preference', 'NTEP Result', 
            'HWC Overruled', 'Has Nikshay ID', 'Nikshay ID', 'Nikshay ID Status'
        ]

        # District Aggregations for District Performance Tab
        dist_grp = working_df.groupby('district').agg(
            Total_Screenings=('district', 'count'),
            Total_Presumptive=('is_presumptive', 'sum'),
            HWC_Overrules=('overrule_flag', 'sum'),
            Nikshay_IDs_Generated=('has_nikshay', 'sum'), # DIRECT NIKSHAY ID COLUMN COUNT
            Nikshay_IDs_Pending=('nikshay_pending_presumptive', 'sum')
        ).reset_index().sort_values(by='Total_Screenings', ascending=False)

        # Multi-District Excel Bundle Preparation
        excel_report_bytes = generate_district_excel_bundle(
            working_df, cho_operational_summary, patient_master_df
        )

        st.sidebar.markdown("---")
        st.sidebar.markdown("### 📥 Analyst Download Center")
        st.sidebar.download_button(
            label="📊 Download Multi-District Report Bundle",
            data=excel_report_bytes,
            file_name=f"Multi_District_Nikshay_Analyst_Report_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        # Dashboard Navigation Tabs
        tabs = st.tabs([
            "📊 Executive KPI Overview",
            "🎯 Presumptive & Nikshay ID Tracker",
            "🏢 District & Block Performance",
            "📁 Multi-District Analyst Reports"
        ])

        # ==========================================
        # TAB 1: EXECUTIVE KPI OVERVIEW
        # ==========================================
        with tabs[0]:
            st.markdown("### 📈 Health Screening & Nikshay Generation Metrics")
            st.markdown(f"""
            <div class="kpi-container">
                <div class="kpi-card" style="border-top-color: #2563EB;">
                    <div class="kpi-title">Total Screenings</div>
                    <div class="kpi-value">{total_scr:,}</div>
                </div>
                <div class="kpi-card" style="border-top-color: #8B5CF6;">
                    <div class="kpi-title">Total Presumptive Cases</div>
                    <div class="kpi-value">{total_pres_cnt:,}</div>
                    <div class="kpi-subtext">{(total_pres_cnt/total_scr*100 if total_scr>0 else 0):.1f}% of total</div>
                </div>
                <div class="kpi-card" style="border-top-color: #EF4444;">
                    <div class="kpi-title">Total Override Count</div>
                    <div class="kpi-value">{total_overrule_cnt:,}</div>
                    <div class="kpi-subtext">{(total_overrule_cnt/total_scr*100 if total_scr>0 else 0):.1f}% overrules</div>
                </div>
                <div class="kpi-card" style="border-top-color: #10B981;">
                    <div class="kpi-title">Nikshay IDs Generated</div>
                    <div class="kpi-value">{nikshay_gen_cnt:,}</div>
                    <div class="kpi-subtext">{(nikshay_gen_cnt/total_scr*100 if total_scr>0 else 0):.1f}% of total screenings</div>
                </div>
                <div class="kpi-card" style="border-top-color: #F59E0B;">
                    <div class="kpi-title">Nikshay ID Pending</div>
                    <div class="kpi-value">{nikshay_pend_cnt:,}</div>
                    <div class="kpi-subtext">{(nikshay_pend_cnt/total_pres_cnt*100 if total_pres_cnt>0 else 0):.1f}% pending rate</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            col_t1, col_t2 = st.columns([3, 2])
            with col_t1:
                st.markdown("#### 🕒 Longitudinal Daily Screening & Nikshay Registration Trend")
                daily_trend = working_df.groupby(working_df['reg_date'].dt.date).agg(
                    Total=('district', 'count'),
                    Nikshay_IDs=('has_nikshay', 'sum')
                ).reset_index()
                fig_trend = px.line(daily_trend, x='reg_date', y=['Total', 'Nikshay_IDs'], 
                                    labels={'value': 'Volume', 'reg_date': 'Date', 'variable': 'Category'},
                                    color_discrete_map={'Total': '#1E3A8A', 'Nikshay_IDs': '#10B981'},
                                    title="Daily Screening vs Nikshay ID Generation")
                fig_trend.update_layout(plot_bgcolor='white', hovermode='x unified')
                st.plotly_chart(fig_trend, use_container_width=True)

            with col_t2:
                st.markdown("#### ⚖️ Presumptive Triage Pipeline")
                pipeline_data = pd.DataFrame({
                    'Stage': ['Total Screened', 'Presumptive Identified', 'Nikshay ID Generated', 'Nikshay ID Pending'],
                    'Count': [total_scr, total_pres_cnt, nikshay_gen_cnt, nikshay_pend_cnt]
                })
                fig_pipe = px.funnel(pipeline_data, x='Count', y='Stage', color_discrete_sequence=['#2563EB'])
                st.plotly_chart(fig_pipe, use_container_width=True)

        # ==========================================
        # TAB 2: PRESUMPTIVE VS NIKSHAY TRACKER
        # ==========================================
        with tabs[1]:
            st.markdown("### 🎯 Presumptive Yield vs. Nikshay ID Conversion Analysis")
            
            c_a1, c_a2 = st.columns(2)
            with c_a1:
                st.markdown("#### 🤖 AI Presumptive vs. Nikshay ID Breakdown")
                ai_cross = pd.crosstab(working_df['ai_preference'], working_df['has_nikshay'], margins=True)
                ai_cross.columns = ['No Nikshay ID', 'Nikshay ID Generated', 'Total']
                st.dataframe(ai_cross, use_container_width=True)
                
            with c_a2:
                st.markdown("#### 🏥 NTEP Result vs. Nikshay ID Breakdown")
                ntep_cross = pd.crosstab(working_df['ntep_result'], working_df['has_nikshay'], margins=True)
                ntep_cross.columns = ['No Nikshay ID', 'Nikshay ID Generated', 'Total']
                st.dataframe(ntep_cross, use_container_width=True)

            st.markdown("---")
            st.markdown("#### 📊 District-wise Presumptive & Nikshay Generation Bar Graph")
            
            fig_bar = px.bar(
                dist_grp, 
                x='district', 
                y=['Total_Presumptive', 'Nikshay_IDs_Generated', 'Nikshay_IDs_Pending'],
                barmode='group',
                title="District Comparison: Total Presumptive vs Nikshay Generated vs Nikshay Pending",
                labels={'district': 'District', 'value': 'Count', 'variable': 'Metric'},
                color_discrete_sequence=['#8B5CF6', '#10B981', '#F59E0B']
            )
            st.plotly_chart(fig_bar, use_container_width=True)

        # ==========================================
        # TAB 3: DISTRICT & BLOCK PERFORMANCE
        # ==========================================
        with tabs[2]:
            st.markdown("### 🏢 Consolidated District Performance Table")
            st.dataframe(dist_grp, use_container_width=True)

            st.markdown("---")
            st.markdown("### 🧱 Block-Level Operations Performance")
            block_summary = working_df.groupby(['block', 'district']).agg(
                Total_Screenings=('block', 'count'),
                Total_Presumptive=('is_presumptive', 'sum'),
                HWC_Overrules=('overrule_flag', 'sum'),
                Nikshay_IDs_Generated=('has_nikshay', 'sum'),
                Nikshay_IDs_Pending=('nikshay_pending_presumptive', 'sum')
            ).reset_index().sort_values(by='Total_Screenings', ascending=False)
            
            st.dataframe(block_summary, use_container_width=True)

        # ==========================================
        # TAB 4: MULTI-DISTRICT ANALYST REPORTS
        # ==========================================
        with tabs[3]:
            st.markdown("### 📁 Comprehensive Multi-District Analyst Reports")
            
            # --- SUB-TAB 1: CHO OPERATIONAL SUMMARY REPORT ---
            st.markdown("#### 👩‍⚕️ 1. CHO Operational Summary Report")
            st.markdown("*Shows CHO wise Screening Count, Presumptive Count, Override Count, Nikshay ID Generated Count, and Nikshay ID Pending Count.*")
            
            st.dataframe(
                cho_operational_summary,
                use_container_width=True,
                column_config={
                    "Total Count": st.column_config.NumberColumn("Total Screenings", format="%d"),
                    "Total Presumptive Count": st.column_config.NumberColumn("Presumptive Count", format="%d"),
                    "Total Override Count": st.column_config.NumberColumn("Override Count", format="%d"),
                    "Total Nikshay ID Generated Count": st.column_config.NumberColumn("Nikshay Generated", format="%d"),
                    "Total Nikshay ID Generation Pending Count": st.column_config.NumberColumn("Nikshay Pending", format="%d")
                }
            )

            st.markdown("---")

            # --- SUB-TAB 2: COMPLETE PATIENT-WISE LINE-LIST REPORT ---
            st.markdown("#### 👤 2. Complete Patient-Wise Master Line-List Report")
            st.markdown("*Detailed line-list of every patient screened with individual AI result, NTEP result, Overrule status, and Nikshay ID details.*")

            selected_report_dist = st.selectbox("Filter Patient Line-List by District", options=["All Districts"] + dist_opts)
            
            if selected_report_dist != "All Districts":
                display_patient_df = patient_master_df[patient_master_df['District'] == selected_report_dist]
            else:
                display_patient_df = patient_master_df

            st.markdown(f"**Showing {len(display_patient_df):,} Patient Records**")
            st.dataframe(display_patient_df, use_container_width=True)

else:
    st.info("💡 Please upload the screening dataset in the left sidebar to initialize the dashboard and run analytics.")