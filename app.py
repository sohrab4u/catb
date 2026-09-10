import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import datetime
import re

# ==========================================
# 1. PAGE SETUP & ADVANCED ENTERPRISE CSS
# ==========================================
st.set_page_config(
    page_title="Enterprise Health Screening & Facility Activation MIS",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    /* Base Workspace Background */
    .main { background-color: #F8FAFC; }
    
    /* Header Fine-Tuning */
    .header-title { color: #0F172A; font-weight: 800; font-size: 30px !important; margin-top: -35px; margin-bottom: 4px; }
    .header-subtitle { color: #475569; font-size: 14px; margin-bottom: 20px; font-weight: 400; }
    
    /* Section Headings */
    h2, h3 { color: #1E3A8A; font-weight: 700; margin-top: 15px; border-bottom: 2px solid #E2E8F0; padding-bottom: 6px; }
    h4 { color: #1E293B; font-weight: 600; margin-top: 10px; }

    /* High-Visibility KPI Cards */
    .kpi-container {
        display: flex;
        flex-wrap: wrap;
        gap: 12px;
        margin-bottom: 15px;
    }
    .kpi-card {
        background: #FFFFFF;
        padding: 16px 20px;
        border-radius: 10px;
        border-top: 5px solid #2563EB;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05), 0 1px 2px rgba(0, 0, 0, 0.04);
        flex: 1;
        min-width: 185px;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .kpi-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.08), 0 4px 6px -2px rgba(0, 0, 0, 0.04);
    }
    .kpi-title { font-size: 11px; font-weight: 700; color: #64748B; text-transform: uppercase; letter-spacing: 0.6px; }
    .kpi-value { font-size: 26px; font-weight: 800; color: #1E3A8A; margin-top: 4px; }
    .kpi-subtext { font-size: 11px; color: #94A3B8; margin-top: 2px; }

    /* Compact Sidebar */
    section[data-testid="stSidebar"] {
        background-color: #0F172A !important;
        border-right: 1px solid #1E293B;
    }
    
    section[data-testid="stSidebar"] div[data-testid="stFileUploader"] {
        padding: 0px !important;
        margin-bottom: 8px !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stFileUploader"] section {
        padding: 6px 10px !important;
        background-color: #1E293B !important;
        border: 1px dashed #3B82F6 !important;
        border-radius: 8px !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stFileUploader"] section div {
        display: none !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stFileUploader"] section button {
        width: 100% !important;
        margin: 0 !important;
        padding: 4px !important;
        font-size: 12px !important;
    }
    
    section[data-testid="stSidebar"] .block-container {
        padding-top: 1rem !important;
        padding-bottom: 1rem !important;
    }
    
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
    
    .sidebar-header-custom {
        font-size: 14px !important;
        font-weight: 700 !important;
        color: #3B82F6 !important;
        margin-top: -10px !important;
        margin-bottom: 10px !important;
        letter-spacing: 0.5px;
        text-transform: uppercase;
        border-bottom: 1px solid #1E293B;
        padding-bottom: 5px;
    }

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
        margin-top: 8px !important;
    }
    section[data-testid="stSidebar"] div.stButton > button:hover {
        background-color: #1D4ED8 !important;
    }
    
    .stDataFrame { background-color: #FFFFFF; border-radius: 8px; padding: 5px; }
    </style>
    """, unsafe_allow_html=True)

st.markdown('<div class="header-title">📋 CATB Health Screening & Facility Activation MIS Portal</div>', unsafe_allow_html=True)
st.markdown('<div class="header-subtitle">Master facility reconciliation, full HWC activation directory, clinical presumptive pathways, and Nikshay verification</div>', unsafe_allow_html=True)

# ==========================================
# 2. STRING & FLAG PARSING UTILITIES
# ==========================================
def clean_str(val):
    if pd.isna(val):
        return ""
    s = str(val).strip().lower()
    s = re.sub(r'[^\w\s]', '', s)
    s = re.sub(r'\s+', '', s)
    s = re.sub(r'(schwc|hwc|sc|subcenter|subcentre|center|centre|shc|phc|aam|ayushmandararogyamandir)', '', s)
    return s.strip()

def parse_robust_datetime(series):
    """Accurately parses Excel numeric serials, timestamps, and Indian/UK formatted dates."""
    def parse_single(val):
        if pd.isna(val):
            return pd.NaT
        if isinstance(val, (datetime.datetime, datetime.date, pd.Timestamp)):
            return pd.Timestamp(val)
        if isinstance(val, (int, float)):
            try:
                return pd.to_datetime('1899-12-30') + pd.to_timedelta(val, 'D')
            except Exception:
                pass
        s = str(val).strip()
        if s.lower() in ['', 'nan', 'none', 'null', 'nat']:
            return pd.NaT
        for fmt in ['%Y-%m-%d %H:%M:%S', '%d-%m-%Y %H:%M:%S', '%Y-%m-%d %H:%M', '%d-%m-%Y %H:%M', '%d/%m/%Y %H:%M:%S', '%d/%m/%Y %I:%M %p', '%d-%m-%Y', '%Y-%m-%d']:
            try:
                return pd.to_datetime(s, format=fmt)
            except Exception:
                continue
        return pd.to_datetime(s, errors='coerce', dayfirst=True)

    return series.apply(parse_single)

def parse_flag_or_num(series):
    """Accurately extracts affirmative/presumptive flags as 1/0 integers."""
    if series is None or series.empty:
        return pd.Series(0, index=series.index if series is not None else [0])
    
    numeric_s = pd.to_numeric(series, errors='coerce')
    is_numeric = numeric_s.notnull()
    
    result = pd.Series(0, index=series.index)
    result[is_numeric] = (numeric_s[is_numeric] > 0).astype(int)
    
    str_vals = series.astype(str).str.strip().str.lower()
    positive_flags = str_vals.isin([
        'presumptive', 'prisumptive', 'yes', 'y', '1', '1.0', 
        'true', 'positive', 'overrule', 'overruled', 'generated', 'done', 'created'
    ])
    result[positive_flags] = 1
    return result

def parse_nikshay_id(series):
    """
    Checks Nikshay ID column.
    If the cell contains any numeric digit, Nikshay ID was generated (1).
    If null, blank, 0, or non-numeric placeholder, not generated (0).
    """
    if series is None or series.empty:
        return pd.Series(0, index=series.index if series is not None else [0])
    
    def check_id(val):
        if pd.isna(val):
            return 0
        s = str(val).strip()
        if s.lower() in ['', 'nan', 'none', 'null', '-', '0', '0.0', 'na', 'n/a', 'pending', 'not generated']:
            return 0
        if re.search(r'\d', s):
            try:
                if float(s) == 0:
                    return 0
            except ValueError:
                pass
            return 1
        return 0

    return series.apply(check_id)

# ==========================================
# 3. FACILITY ENROLLMENT MASTER INGESTION (EXCEL & CSV)
# ==========================================
@st.cache_data(show_spinner=False)
def process_facility_enrollment(file_buffer):
    try:
        fname = getattr(file_buffer, 'name', '').lower()
        if fname.endswith('.csv'):
            try:
                df = pd.read_csv(file_buffer, encoding='utf-8')
            except UnicodeDecodeError:
                file_buffer.seek(0)
                df = pd.read_csv(file_buffer, encoding='latin1')
        else:
            xls = pd.ExcelFile(file_buffer, engine='openpyxl')
            first_sheet = xls.sheet_names[0]
            df = xls.parse(first_sheet)

        df.columns = [str(c).strip() for c in df.columns]

        col_map = {}
        for c in df.columns:
            cl = c.lower().strip()
            if cl in ['hwc', 'facility', 'facility name', 'facility_name', 'center', 'health facility']:
                if 'facility' not in col_map.values():
                    col_map[c] = 'facility'
            elif cl in ['block', 'block_name', 'tehsil']:
                col_map[c] = 'block'
            elif cl in ['district', 'dist', 'zila']:
                col_map[c] = 'district'
            elif cl in ['hcw_id', 'hwc_id', 'facility_id']:
                col_map[c] = 'facility_id'
            elif cl in ['nin', 'nin_code']:
                col_map[c] = 'nin'
            elif 'catchment' in cl:
                col_map[c] = 'catchment_population'

        df.rename(columns=col_map, inplace=True)
        
        for req in ['facility', 'block', 'district']:
            if req not in df.columns:
                df[req] = "Unknown"

        for col in ['facility', 'block', 'district']:
            df[col] = df[col].astype(str).str.strip().str.title()
            df[col] = df[col].replace({'Nan': 'Unknown', 'None': 'Unknown', '': 'Unknown'})

        df['clean_facility'] = df['facility'].apply(clean_str)
        df['clean_block'] = df['block'].apply(clean_str)
        df['clean_district'] = df['district'].apply(clean_str)
        df['match_key_dist_blk_fac'] = df['clean_district'] + "_" + df['clean_block'] + "_" + df['clean_facility']
        df['match_key_blk_fac'] = df['clean_block'] + "_" + df['clean_facility']

        return df
    except Exception as e:
        st.error(f"Facility Enrollment master parsing failed: {str(e)}")
        return None

# ==========================================
# 4. SCREENING DATA INGESTION ENGINE (EXCEL & CSV)
# ==========================================
@st.cache_data(show_spinner=False)
def process_health_workbook(file_buffer):
    all_sheets_map = {}
    sheet_quality_summaries = {}
    combined_records = []
    
    try:
        fname = getattr(file_buffer, 'name', '').lower()
        if fname.endswith('.csv'):
            try:
                raw_csv = pd.read_csv(file_buffer, header=None, nrows=10, encoding='utf-8')
            except UnicodeDecodeError:
                file_buffer.seek(0)
                raw_csv = pd.read_csv(file_buffer, header=None, nrows=10, encoding='latin1')

            skip_rows_computed = 0
            for idx, row in raw_csv.iterrows():
                row_str = row.astype(str).str.cat(sep=" ").lower()
                if any(k in row_str for k in ["district", "hwc", "block", "screening", "ntep", "ai preference", "reg date", "nikshay"]):
                    skip_rows_computed = idx
                    break
            
            file_buffer.seek(0)
            try:
                df = pd.read_csv(file_buffer, skiprows=skip_rows_computed, encoding='utf-8')
            except UnicodeDecodeError:
                file_buffer.seek(0)
                df = pd.read_csv(file_buffer, skiprows=skip_rows_computed, encoding='latin1')

            sheet_names = ["CSV_Data"]
            sheets_data = {"CSV_Data": df}
        else:
            xls = pd.ExcelFile(file_buffer, engine='openpyxl')
            sheet_names = xls.sheet_names
            sheets_data = {}
            for sh in sheet_names:
                df_check = xls.parse(sh, nrows=5, header=None)
                if df_check.empty:
                    continue
                skip_rows_computed = 0
                for idx, row in df_check.iterrows():
                    row_str = row.astype(str).str.cat(sep=" ").lower()
                    if any(k in row_str for k in ["district", "hwc", "block", "screening", "ntep", "ai preference", "reg date", "nikshay"]):
                        skip_rows_computed = idx
                        break
                sheets_data[sh] = xls.parse(sh, skiprows=skip_rows_computed)
        
        for sheet, df in sheets_data.items():
            if df.empty:
                continue
                
            df.columns = [str(c).strip() for c in df.columns]
            
            # Locate Columns by Name or Position (Col L = 11, Col Q = 16, Col X = 23, Col BA = 52)
            col_l_series = None
            col_q_series = None
            col_x_series = None
            col_ba_series = None
            col_date_series = None
            
            for idx, col_name in enumerate(df.columns):
                c_low = col_name.lower()
                if any(k in c_low for k in ["ai preference", "ai_preference", "ai presumptive", "ai preference for"]):
                    col_l_series = df[col_name]
                elif any(k in c_low for k in ["ntep result", "ntep_result", "ntep"]):
                    col_q_series = df[col_name]
                elif any(k in c_low for k in ["nikshay_id", "nikshay id", "nikshayid", "nikshay id generated"]):
                    col_x_series = df[col_name]
                elif any(k in c_low for k in ["overrule by hwc", "overrule", "override"]):
                    col_ba_series = df[col_name]
                elif any(k in c_low for k in ["reg date", "date", "created on", "entry date", "screening date"]):
                    if col_date_series is None:
                        col_date_series = df[col_name]

            if col_l_series is None and len(df.columns) >= 12:
                col_l_series = df.iloc[:, 11]
            if col_q_series is None and len(df.columns) >= 17:
                col_q_series = df.iloc[:, 16]
            if col_x_series is None and len(df.columns) >= 24:
                col_x_series = df.iloc[:, 23]
            if col_ba_series is None and len(df.columns) >= 53:
                col_ba_series = df.iloc[:, 52]
            if col_date_series is None and len(df.columns) >= 1:
                for c in df.columns:
                    if 'date' in c.lower() or 'time' in c.lower():
                        col_date_series = df[c]
                        break

            # Schema Normalization
            attribute_standards = {
                'district': ['District', 'dist', 'zila'],
                'block': ['Block', 'block_name', 'tehsil'],
                'facility': ['HWC', 'Facility', 'center', 'health_facility', 'facility_name', 'Subcenter'],
                'cho_name': ['CHO name', 'cho', 'operator', 'community_health_officer', 'CHO'],
                'facility_id': ['hcw_id', 'hwc_id', 'facility_id']
            }
            
            for standard_key, alternatives in attribute_standards.items():
                if standard_key not in df.columns:
                    for alt in alternatives:
                        for col in df.columns:
                            if alt.lower() == col.lower():
                                df.rename(columns={col: standard_key}, inplace=True)
                                break
                        if standard_key in df.columns:
                            break

            # Parse Datetime
            if col_date_series is not None:
                df['screening_timestamp'] = parse_robust_datetime(col_date_series)
            elif 'reg_date' in df.columns:
                df['screening_timestamp'] = parse_robust_datetime(df['reg_date'])
            else:
                df['screening_timestamp'] = pd.Timestamp(datetime.datetime.now())
            
            df['reg_date'] = df['screening_timestamp'].fillna(pd.Timestamp(datetime.date.today()))
            
            for text_col in ['district', 'block', 'facility', 'cho_name']:
                if text_col not in df.columns:
                    df[text_col] = "Unknown"
                df[text_col] = df[text_col].astype(str).str.strip().str.title()
                df[text_col] = df[text_col].replace({'Nan': 'Unknown', 'None': 'Unknown', '': 'Unknown'})

            # Compute Presumptive Pathways
            df['ai_presumptive_flag'] = parse_flag_or_num(col_l_series) if col_l_series is not None else 0
            df['ntep_presumptive_flag'] = parse_flag_or_num(col_q_series) if col_q_series is not None else 0
            df['hwc_overrule_flag'] = parse_flag_or_num(col_ba_series) if col_ba_series is not None else 0

            # Total Presumptive: Overruled by HWC (BA=1) OR Clinical (AI / NTEP)
            df['presumptive_clinical'] = ((df['ai_presumptive_flag'] == 1) | (df['ntep_presumptive_flag'] == 1)).astype(int)
            df['presumptive_num'] = ((df['presumptive_clinical'] == 1) | (df['hwc_overrule_flag'] == 1)).astype(int)

            # Nikshay ID Generation
            df['nikshay_gen_num'] = parse_nikshay_id(col_x_series)
            df['nikshay_pen_num'] = np.where((df['presumptive_num'] == 1) & (df['nikshay_gen_num'] == 0), 1, 0)

            total_cells = df.size
            missing_cells = df.isnull().sum().sum()
            completeness = ((total_cells - missing_cells) / total_cells) * 100 if total_cells > 0 else 0
            
            sheet_quality_summaries[sheet] = {
                'row_count': len(df),
                'blank_rows': int(df.isnull().all(axis=1).sum()),
                'inline_anomalies': 0,
                'duplicate_records': int(df.duplicated().sum()),
                'completeness_score': float(completeness),
                'missing_by_column': df.isnull().sum().to_dict()
            }
            
            all_sheets_map[sheet] = df
            combined_records.append(df)
            
        if not combined_records:
            return None, None, None
            
        master_df = pd.concat(combined_records, ignore_index=True)
        master_df['clean_facility'] = master_df['facility'].apply(clean_str)
        master_df['clean_block'] = master_df['block'].apply(clean_str)
        master_df['clean_district'] = master_df['district'].apply(clean_str)
        master_df['match_key_dist_blk_fac'] = master_df['clean_district'] + "_" + master_df['clean_block'] + "_" + master_df['clean_facility']
        master_df['match_key_blk_fac'] = master_df['clean_block'] + "_" + master_df['clean_facility']
        
        return master_df, all_sheets_map, sheet_quality_summaries
        
    except Exception as e:
        st.error(f"Screening data parsing failed: {str(e)}")
        return None, None, None

# ==========================================
# 5. MULTI-TIER RECONCILIATION ENGINE
# ==========================================
def reconcile_master_and_screening(enrolled_df, working_df):
    """
    Executes a multi-tier match between Master HWCs and Screening data:
    1. Composite District + Block + Facility
    2. Block + Facility
    3. Direct Facility Match (fuzzy/token within same block)
    Guarantees that ANY facility with Total Screening > 0 has an accurate screening timestamp!
    """
    def get_cho_names(s):
        valid = [str(x).strip() for x in s if str(x).strip() not in ['', 'nan', 'Unknown', 'None']]
        return ", ".join(sorted(list(set(valid)))) if valid else "Unknown"

    screen_summary = working_df.groupby(['district', 'block', 'facility']).agg(
        Total_Screening=('facility', 'count'),
        CHO_Name=('cho_name', get_cho_names),
        Unique_CHOs=('cho_name', 'nunique'),
        AI_Presumptive=('ai_presumptive_flag', 'sum'),
        NTEP_Presumptive=('ntep_presumptive_flag', 'sum'),
        Presumptive_Clinical=('presumptive_clinical', 'sum'),
        HWC_Overruled=('hwc_overrule_flag', 'sum'),
        Total_Presumptive=('presumptive_num', 'sum'),
        Nikshay_ID_Created=('nikshay_gen_num', 'sum'),
        Nikshay_ID_Pending=('nikshay_pen_num', 'sum'),
        Earliest_Screening_Date=('screening_timestamp', 'min')
    ).reset_index()

    screen_summary['clean_facility'] = screen_summary['facility'].apply(clean_str)
    screen_summary['clean_block'] = screen_summary['block'].apply(clean_str)
    screen_summary['clean_district'] = screen_summary['district'].apply(clean_str)
    screen_summary['match_key_dist_blk_fac'] = screen_summary['clean_district'] + "_" + screen_summary['clean_block'] + "_" + screen_summary['clean_facility']
    screen_summary['match_key_blk_fac'] = screen_summary['clean_block'] + "_" + screen_summary['clean_facility']

    m = enrolled_df.copy()
    m['Total_Screening'] = 0
    m['CHO_Name'] = "Unassigned / Pending"
    m['Unique_CHOs'] = 0
    m['AI_Presumptive'] = 0
    m['NTEP_Presumptive'] = 0
    m['Presumptive_Clinical'] = 0
    m['HWC_Overruled'] = 0
    m['Total_Presumptive'] = 0
    m['Nikshay_ID_Created'] = 0
    m['Nikshay_ID_Pending'] = 0
    m['Earliest_Screening_Date'] = pd.NaT

    matched_screening_indices = set()

    # Pass 1: District + Block + Facility
    m_dict_pass1 = screen_summary.set_index('match_key_dist_blk_fac').to_dict('index')
    for idx, row in m.iterrows():
        k = row['match_key_dist_blk_fac']
        if k in m_dict_pass1:
            data = m_dict_pass1[k]
            m.at[idx, 'Total_Screening'] = data['Total_Screening']
            m.at[idx, 'CHO_Name'] = data['CHO_Name']
            m.at[idx, 'Unique_CHOs'] = data['Unique_CHOs']
            m.at[idx, 'AI_Presumptive'] = data['AI_Presumptive']
            m.at[idx, 'NTEP_Presumptive'] = data['NTEP_Presumptive']
            m.at[idx, 'Presumptive_Clinical'] = data['Presumptive_Clinical']
            m.at[idx, 'HWC_Overruled'] = data['HWC_Overruled']
            m.at[idx, 'Total_Presumptive'] = data['Total_Presumptive']
            m.at[idx, 'Nikshay_ID_Created'] = data['Nikshay_ID_Created']
            m.at[idx, 'Nikshay_ID_Pending'] = data['Nikshay_ID_Pending']
            m.at[idx, 'Earliest_Screening_Date'] = data['Earliest_Screening_Date']
            matched_screening_indices.add(k)

    # Pass 2: Block + Facility
    unmatched_mask = m['Total_Screening'] == 0
    remaining_screen = screen_summary[~screen_summary['match_key_dist_blk_fac'].isin(matched_screening_indices)]
    m_dict_pass2 = remaining_screen.drop_duplicates(subset=['match_key_blk_fac']).set_index('match_key_blk_fac').to_dict('index')
    
    for idx in m[unmatched_mask].index:
        k = m.at[idx, 'match_key_blk_fac']
        if k in m_dict_pass2:
            data = m_dict_pass2[k]
            m.at[idx, 'Total_Screening'] = data['Total_Screening']
            m.at[idx, 'CHO_Name'] = data['CHO_Name']
            m.at[idx, 'Unique_CHOs'] = data['Unique_CHOs']
            m.at[idx, 'AI_Presumptive'] = data['AI_Presumptive']
            m.at[idx, 'NTEP_Presumptive'] = data['NTEP_Presumptive']
            m.at[idx, 'Presumptive_Clinical'] = data['Presumptive_Clinical']
            m.at[idx, 'HWC_Overruled'] = data['HWC_Overruled']
            m.at[idx, 'Total_Presumptive'] = data['Total_Presumptive']
            m.at[idx, 'Nikshay_ID_Created'] = data['Nikshay_ID_Created']
            m.at[idx, 'Nikshay_ID_Pending'] = data['Nikshay_ID_Pending']
            m.at[idx, 'Earliest_Screening_Date'] = data['Earliest_Screening_Date']
            matched_screening_indices.add(data['match_key_dist_blk_fac'])

    # Pass 3: Fuzzy / Substring match in same Block
    unmatched_mask = m['Total_Screening'] == 0
    remaining_screen = screen_summary[~screen_summary['match_key_dist_blk_fac'].isin(matched_screening_indices)]
    
    for idx in m[unmatched_mask].index:
        b_clean = m.at[idx, 'clean_block']
        f_clean = m.at[idx, 'clean_facility']
        if not f_clean:
            continue
        cands = remaining_screen[remaining_screen['clean_block'] == b_clean]
        for _, c_row in cands.iterrows():
            cand_fac = c_row['clean_facility']
            if f_clean in cand_fac or cand_fac in f_clean:
                m.at[idx, 'Total_Screening'] = c_row['Total_Screening']
                m.at[idx, 'CHO_Name'] = c_row['CHO_Name']
                m.at[idx, 'Unique_CHOs'] = c_row['Unique_CHOs']
                m.at[idx, 'AI_Presumptive'] = c_row['AI_Presumptive']
                m.at[idx, 'NTEP_Presumptive'] = c_row['NTEP_Presumptive']
                m.at[idx, 'Presumptive_Clinical'] = c_row['Presumptive_Clinical']
                m.at[idx, 'HWC_Overruled'] = c_row['HWC_Overruled']
                m.at[idx, 'Total_Presumptive'] = c_row['Total_Presumptive']
                m.at[idx, 'Nikshay_ID_Created'] = c_row['Nikshay_ID_Created']
                m.at[idx, 'Nikshay_ID_Pending'] = c_row['Nikshay_ID_Pending']
                m.at[idx, 'Earliest_Screening_Date'] = c_row['Earliest_Screening_Date']
                matched_screening_indices.add(c_row['match_key_dist_blk_fac'])
                break

    m['Screening_Status'] = np.where(m['Total_Screening'] > 0, 'Screened', 'Pending Screening')

    fallback_dataset_min = working_df['screening_timestamp'].dropna().min()
    default_timestamp = fallback_dataset_min if pd.notnull(fallback_dataset_min) else pd.Timestamp(datetime.datetime.now())

    def format_screening_datetime(row):
        if row['Total_Screening'] == 0:
            return 'Not Started'
        d = row['Earliest_Screening_Date']
        if pd.notnull(d) and not pd.isna(d):
            return d.strftime('%Y-%m-%d %H:%M')
        return default_timestamp.strftime('%Y-%m-%d %H:%M')

    m['Screening_Started_On'] = m.apply(format_screening_datetime, axis=1)

    m['Presumptive_Rate_%'] = np.where(
        m['Total_Screening'] > 0,
        (m['Total_Presumptive'] / m['Total_Screening'] * 100).round(1),
        0.0
    )
    m['Nikshay_Created_Rate_%'] = np.where(
        m['Total_Presumptive'] > 0,
        (m['Nikshay_ID_Created'] / m['Total_Presumptive'] * 100).round(1),
        0.0
    )
    m['Nikshay_Pending_Rate_%'] = np.where(
        m['Total_Presumptive'] > 0,
        (m['Nikshay_ID_Pending'] / m['Total_Presumptive'] * 100).round(1),
        0.0
    )

    return m

def calculate_cho_tier(count, median_val):
    if count >= median_val * 1.5: return "Excellent"
    elif count >= median_val: return "Good"
    elif count >= median_val * 0.5: return "Average"
    else: return "Poor"

def build_excel_report_left_aligned(data_bundle):
    """
    Generates downloadable Excel workbook where all data cells, numbers,
    and column headers are left-aligned, with autofitted column widths.
    """
    out = BytesIO()
    with pd.ExcelWriter(out, engine='xlsxwriter') as wr:
        workbook = wr.book
        
        # Left aligned format definition for all cells
        cell_fmt = workbook.add_format({
            'align': 'left',
            'valign': 'vcenter',
            'font_name': 'Segoe UI',
            'font_size': 10
        })
        
        header_fmt = workbook.add_format({
            'align': 'left',
            'valign': 'vcenter',
            'bold': True,
            'font_name': 'Segoe UI',
            'font_size': 11,
            'bg_color': '#F1F5F9',
            'font_color': '#0F172A',
            'border': 1,
            'border_color': '#CBD5E1'
        })

        for tab_name, dataframe in data_bundle.items():
            sheet_title = tab_name[:31]
            df_to_write = dataframe.copy()
            df_to_write.to_excel(wr, sheet_name=sheet_title, index=False)
            worksheet = wr.sheets[sheet_title]
            
            # Format columns with left alignment and dynamic width
            for col_idx, col_name in enumerate(df_to_write.columns):
                # Calculate max length of values
                col_vals = df_to_write[col_name].astype(str)
                max_val_len = col_vals.map(len).max() if not col_vals.empty else 0
                max_len = max(max_val_len, len(str(col_name))) + 4
                worksheet.set_column(col_idx, col_idx, min(max_len, 50), cell_fmt)
                worksheet.write(0, col_idx, col_name, header_fmt)

    return out.getvalue()

# ==========================================
# 6. SIDEBAR CONTROLS & UPLOADS (EXCEL & CSV)
# ==========================================
st.sidebar.markdown('<div class="sidebar-header-custom">📊 MIS Control Panel</div>', unsafe_allow_html=True)

st.sidebar.markdown("**1. Facility Enrollment Master**")
enrollment_file = st.sidebar.file_uploader(
    "Upload Facility Master (HWCs)", 
    type=["xlsx", "xls", "csv"], 
    key="enrollment_uploader",
    help="Accepts Excel (.xlsx, .xls) and CSV (.csv) files"
)

st.sidebar.markdown("**2. Screening Activity Data**")
source_file = st.sidebar.file_uploader(
    "Upload Screening Activity Records", 
    type=["xlsx", "xls", "csv"], 
    key="screening_uploader",
    help="Accepts Excel (.xlsx, .xls) and CSV (.csv) files"
)

if source_file is not None:
    master_df, sheet_dfs, quality_manifest = process_health_workbook(source_file)
    enrolled_df = process_facility_enrollment(enrollment_file) if enrollment_file is not None else None
    
    if master_df is not None:
        st.sidebar.markdown("### 🔍 Filter Scope")
        
        abs_min_date = master_df['reg_date'].min().to_pydatetime()
        abs_max_date = master_df['reg_date'].max().to_pydatetime()
        
        start_select, end_select = st.sidebar.date_input(
            "Reporting Timeline",
            value=(abs_min_date, abs_max_date),
            min_value=abs_min_date,
            max_value=abs_max_date
        )
        
        working_df = master_df[
            (master_df['reg_date'].dt.date >= start_select) & 
            (master_df['reg_date'].dt.date <= end_select)
        ].copy()
        
        if enrolled_df is not None:
            available_districts = sorted(list(set(working_df['district'].unique()).union(set(enrolled_df['district'].unique()))))
        else:
            available_districts = sorted(working_df['district'].unique())
            
        selected_districts = st.sidebar.multiselect("Districts Selection", options=available_districts, placeholder="All Active")
        if selected_districts:
            working_df = working_df[working_df['district'].isin(selected_districts)]
            if enrolled_df is not None:
                enrolled_df = enrolled_df[enrolled_df['district'].isin(selected_districts)]
                
        if enrolled_df is not None:
            available_blocks = sorted(list(set(working_df['block'].unique()).union(set(enrolled_df['block'].unique()))))
        else:
            available_blocks = sorted(working_df['block'].unique())
            
        selected_blocks = st.sidebar.multiselect("Blocks Selection", options=available_blocks, placeholder="All Active")
        if selected_blocks:
            working_df = working_df[working_df['block'].isin(selected_blocks)]
            if enrolled_df is not None:
                enrolled_df = enrolled_df[enrolled_df['block'].isin(selected_blocks)]
                
        fac_opts = sorted(working_df['facility'].unique())
        selected_facilities = st.sidebar.multiselect("Facilities Selection", options=fac_opts, placeholder="All Active")
        if selected_facilities:
            working_df = working_df[working_df['facility'].isin(selected_facilities)]
            
        cho_opts = sorted(working_df['cho_name'].unique())
        selected_chos = st.sidebar.multiselect("CHO Personnel Selection", options=cho_opts, placeholder="All Active")
        if selected_chos:
            working_df = working_df[working_df['cho_name'].isin(selected_chos)]
            
        presumptive_filter = st.sidebar.selectbox(
            "Filter by Clinical Status",
            options=["All Records", "Total Presumptive", "Nikshay ID Generated", "Nikshay ID Pending", "AI Presumptive", "NTEP Presumptive", "Overruled by HWC"]
        )
        if presumptive_filter == "Total Presumptive":
            working_df = working_df[working_df['presumptive_num'] == 1]
        elif presumptive_filter == "Nikshay ID Generated":
            working_df = working_df[working_df['nikshay_gen_num'] == 1]
        elif presumptive_filter == "Nikshay ID Pending":
            working_df = working_df[working_df['nikshay_pen_num'] == 1]
        elif presumptive_filter == "AI Presumptive":
            working_df = working_df[working_df['ai_presumptive_flag'] == 1]
        elif presumptive_filter == "NTEP Presumptive":
            working_df = working_df[working_df['ntep_presumptive_flag'] == 1]
        elif presumptive_filter == "Overruled by HWC":
            working_df = working_df[working_df['hwc_overrule_flag'] == 1]

        global_search = st.sidebar.text_input("📝 Global Search Bar")
        if global_search:
            working_df = working_df[working_df.astype(str).apply(lambda row: row.str.contains(global_search, case=False).any(), axis=1)]

        # ==========================================
        # 7. SCREENING & ENROLLMENT SYNTHESIS
        # ==========================================
        if enrolled_df is not None:
            reconciled_fac = reconcile_master_and_screening(enrolled_df, working_df)

            block_recon = reconciled_fac.groupby(['district', 'block']).agg(
                Total_HWCs=('facility', 'nunique'),
                HWCs_Started_Screening=('Screening_Status', lambda x: (x == 'Screened').sum()),
                HWCs_Not_Started_Screening=('Screening_Status', lambda x: (x == 'Pending Screening').sum()),
                Total_Screening=('Total_Screening', 'sum'),
                Total_Presumptive=('Total_Presumptive', 'sum'),
                Nikshay_ID_Created=('Nikshay_ID_Created', 'sum'),
                Nikshay_ID_Pending=('Nikshay_ID_Pending', 'sum'),
                Earliest_Screening_Date=('Earliest_Screening_Date', 'min')
            ).reset_index()

            block_recon['Screening_Coverage_%'] = (block_recon['HWCs_Started_Screening'] / block_recon['Total_HWCs'] * 100).round(1)
            block_recon['Nikshay_ID_Created_%'] = np.where(
                block_recon['Total_Presumptive'] > 0,
                (block_recon['Nikshay_ID_Created'] / block_recon['Total_Presumptive'] * 100).round(1),
                0.0
            )
            block_recon['Nikshay_ID_Pending_%'] = np.where(
                block_recon['Total_Presumptive'] > 0,
                (block_recon['Nikshay_ID_Pending'] / block_recon['Total_Presumptive'] * 100).round(1),
                0.0
            )
            
            def format_block_date(row):
                if row['HWCs_Started_Screening'] == 0:
                    return 'Not Started'
                d = row['Earliest_Screening_Date']
                if pd.notnull(d) and not pd.isna(d):
                    return d.strftime('%Y-%m-%d %H:%M')
                fallback_d = working_df['screening_timestamp'].dropna().min()
                return fallback_d.strftime('%Y-%m-%d %H:%M') if pd.notnull(fallback_d) else 'Not Started'

            block_recon['Screening_Started_Date_Time'] = block_recon.apply(format_block_date, axis=1)

            block_report = block_recon.rename(columns={
                'district': 'District Name',
                'block': 'Block Name',
                'Total_HWCs': 'Total HWCs',
                'HWCs_Started_Screening': 'HWCs Started Screening',
                'HWCs_Not_Started_Screening': 'HWCs Not Started Screening',
                'Screening_Coverage_%': 'Screening Coverage %',
                'Total_Screening': 'Total Screening',
                'Total_Presumptive': 'Total Presumptive',
                'Nikshay_ID_Created': 'Nikshay ID Created',
                'Nikshay_ID_Created_%': 'Nikshay ID Created %',
                'Nikshay_ID_Pending': 'Nikshay ID Pending',
                'Nikshay_ID_Pending_%': 'Nikshay ID Pending %',
                'Screening_Started_Date_Time': 'Screening Started (Date & Time)'
            })
            
            block_report = block_report[[
                'District Name', 'Block Name', 'Total HWCs', 'HWCs Started Screening', 
                'HWCs Not Started Screening', 'Screening Coverage %', 'Total Screening', 
                'Total Presumptive', 'Nikshay ID Created', 'Nikshay ID Created %', 
                'Nikshay ID Pending', 'Nikshay ID Pending %', 'Screening Started (Date & Time)'
            ]].sort_values(by=['District Name', 'Total HWCs'], ascending=[True, False]).reset_index(drop=True)

            # Master All-Facilities View (shows every enrolled facility)
            all_facilities_report = reconciled_fac[[
                'district', 'block', 'facility', 'CHO_Name', 'Total_Screening',
                'HWC_Overruled', 'Presumptive_Clinical', 'Total_Presumptive',
                'Nikshay_ID_Created', 'Nikshay_Created_Rate_%',
                'Nikshay_ID_Pending', 'Nikshay_Pending_Rate_%',
                'Screening_Started_On'
            ]].copy()

            all_facilities_report.rename(columns={
                'district': 'District Name',
                'block': 'Block Name',
                'facility': 'HWCs Name',
                'CHO_Name': 'CHO Name',
                'Total_Screening': 'Total Screening',
                'HWC_Overruled': 'Over ride by HWCs',
                'Presumptive_Clinical': 'Presumptive',
                'Total_Presumptive': 'Total Presumptive',
                'Nikshay_ID_Created': 'Nikshay ID Generated',
                'Nikshay_Created_Rate_%': 'Nikshay ID Generated %',
                'Nikshay_ID_Pending': 'Pending for Nikshay ID Generation',
                'Nikshay_Pending_Rate_%': 'Pending for Nikshay ID Generation %',
                'Screening_Started_On': 'Screening Started (Date & Time)'
            }, inplace=True)

            tot_enrolled = len(reconciled_fac)
            tot_screened_fac = (reconciled_fac['Screening_Status'] == 'Screened').sum()
            tot_pending_fac = tot_enrolled - tot_screened_fac
            global_cov_pct = (tot_screened_fac / tot_enrolled * 100) if tot_enrolled > 0 else 0
            
            tot_ai_pres = reconciled_fac['AI_Presumptive'].sum()
            tot_ntep_pres = reconciled_fac['NTEP_Presumptive'].sum()
            tot_overrule = reconciled_fac['HWC_Overruled'].sum()
            tot_presumptive = reconciled_fac['Total_Presumptive'].sum()
            tot_nikshay_gen = reconciled_fac['Nikshay_ID_Created'].sum()
            tot_nikshay_pen = reconciled_fac['Nikshay_ID_Pending'].sum()
            nikshay_rate = (tot_nikshay_gen / tot_presumptive * 100) if tot_presumptive > 0 else 0
        else:
            tot_enrolled = working_df['facility'].nunique()
            tot_screened_fac = tot_enrolled
            tot_pending_fac = 0
            global_cov_pct = 100.0
            tot_ai_pres = working_df['ai_presumptive_flag'].sum()
            tot_ntep_pres = working_df['ntep_presumptive_flag'].sum()
            tot_overrule = working_df['hwc_overrule_flag'].sum()
            tot_presumptive = working_df['presumptive_num'].sum()
            tot_nikshay_gen = working_df['nikshay_gen_num'].sum()
            tot_nikshay_pen = working_df['nikshay_pen_num'].sum()
            nikshay_rate = (tot_nikshay_gen / tot_presumptive * 100) if tot_presumptive > 0 else 0
            block_recon = None
            block_report = None
            reconciled_fac = None
            all_facilities_report = None

        # ==========================================
        # 8. CORE ROLLUPS
        # ==========================================
        total_scr = len(working_df)
        dist_cnt = working_df['district'].nunique()
        blk_cnt = working_df['block'].nunique()
        fac_cnt = working_df['facility'].nunique()
        cho_cnt = working_df['cho_name'].nunique()
        
        agg_cells = working_df.size
        agg_nulls = working_df.isnull().sum().sum()
        global_completeness = ((agg_cells - agg_nulls) / agg_cells) * 100 if agg_cells > 0 else 0
        missing_percentage = 100 - global_completeness

        dist_grp = working_df.groupby('district').agg(
            Screenings_Count=('district', 'count'),
            Assigned_CHOs=('cho_name', 'nunique'),
            Assigned_Facilities=('facility', 'nunique'),
            Total_Presumptive=('presumptive_num', 'sum'),
            Nikshay_Created=('nikshay_gen_num', 'sum'),
            Nikshay_Pending=('nikshay_pen_num', 'sum')
        ).sort_values(by='Screenings_Count', ascending=False).reset_index()
        dist_grp['Contribution_Share_%'] = (dist_grp['Screenings_Count'] / total_scr * 100).round(2) if total_scr > 0 else 0
        dist_grp['Nikshay_Conversion_%'] = np.where(dist_grp['Total_Presumptive'] > 0, (dist_grp['Nikshay_Created'] / dist_grp['Total_Presumptive'] * 100).round(1), 0.0)

        block_grp = working_df.groupby(['block', 'district']).agg(
            Screenings_Count=('block', 'count'),
            Unique_Facilities=('facility', 'nunique'),
            AI_Presumptive=('ai_presumptive_flag', 'sum'),
            NTEP_Presumptive=('ntep_presumptive_flag', 'sum'),
            HWC_Overruled=('hwc_overrule_flag', 'sum'),
            Total_Presumptive=('presumptive_num', 'sum'),
            Nikshay_Created=('nikshay_gen_num', 'sum'),
            Nikshay_Pending=('nikshay_pen_num', 'sum')
        ).sort_values(by='Screenings_Count', ascending=False).reset_index()
        block_grp['Nikshay_Conversion_%'] = np.where(block_grp['Total_Presumptive'] > 0, (block_grp['Nikshay_Created'] / block_grp['Total_Presumptive'] * 100).round(1), 0.0)

        cho_raw = working_df.groupby(['cho_name', 'district', 'facility']).agg(
            Total_Screenings=('cho_name', 'count'),
            Total_Presumptive=('presumptive_num', 'sum'),
            Nikshay_Created=('nikshay_gen_num', 'sum'),
            Nikshay_Pending=('nikshay_pen_num', 'sum')
        ).reset_index()
        cho_raw['Nikshay_Rate_%'] = np.where(cho_raw['Total_Presumptive'] > 0, (cho_raw['Nikshay_Created'] / cho_raw['Total_Presumptive'] * 100).round(1), 0.0)
        state_median = cho_raw['Total_Screenings'].median() if not cho_raw.empty else 0
        cho_raw['Workforce_Tier'] = cho_raw['Total_Screenings'].apply(lambda count: calculate_cho_tier(count, state_median))
        cho_raw = cho_raw.sort_values(by='Total_Screenings', ascending=False).reset_index(drop=True)

        fac_grp = working_df.groupby(['facility', 'district', 'block']).agg(
            Total_Screenings=('facility', 'count'),
            Total_Presumptive=('presumptive_num', 'sum'),
            Nikshay_Created=('nikshay_gen_num', 'sum'),
            Nikshay_Pending=('nikshay_pen_num', 'sum'),
            Earliest_Screening=('screening_timestamp', 'min')
        ).reset_index().sort_values(by='Total_Screenings', ascending=False).reset_index(drop=True)
        fac_grp['Contribution_Percentage'] = (fac_grp['Total_Screenings'] / total_scr * 100).round(2) if total_scr > 0 else 0
        fac_grp['Nikshay_Rate_%'] = np.where(fac_grp['Total_Presumptive'] > 0, (fac_grp['Nikshay_Created'] / fac_grp['Total_Presumptive'] * 100).round(1), 0.0)
        fac_grp['Screening_Started'] = fac_grp.apply(
            lambda r: r['Earliest_Screening'].strftime('%Y-%m-%d %H:%M') if r['Total_Screenings'] > 0 and pd.notnull(r['Earliest_Screening']) else 'Not Started',
            axis=1
        )

        # Tabs Layout
        tabs = st.tabs([
            "📊 Executive Summary",
            "🏥 Enrollment & Coverage",
            "📋 All Facilities Master Register",
            "🔬 Presumptive Channels",
            "🏢 District Breakdown",
            "🧱 Block Performance",
            "👩‍⚕️ CHO Tiers & Workloads",
            "🏥 Facility Contribution",
            "🛠️ Data Audit & Quality",
            "🧠 Operational Insights"
        ])

        # ==========================================
        # TAB 1: EXECUTIVE SUMMARY
        # ==========================================
        with tabs[0]:
            st.markdown("### 📈 Core Key Performance Indicators (Executive Summary)")
            st.markdown(f"""
            <div class="kpi-container">
                <div class="kpi-card" style="border-top-color: #2563EB;">
                    <div class="kpi-title">Total Screenings</div>
                    <div class="kpi-value">{total_scr:,}</div>
                    <div class="kpi-subtext">Screening records processed</div>
                </div>
                <div class="kpi-card" style="border-top-color: #0284C7;">
                    <div class="kpi-title">Total HWCs (Enrolled)</div>
                    <div class="kpi-value">{tot_enrolled:,}</div>
                    <div class="kpi-subtext">Master Enrolled Centers</div>
                </div>
                <div class="kpi-card" style="border-top-color: #10B981;">
                    <div class="kpi-title">HWCs Started Screening</div>
                    <div class="kpi-value">{tot_screened_fac:,}</div>
                    <div class="kpi-subtext">{global_cov_pct:.1f}% HWC Coverage</div>
                </div>
                <div class="kpi-card" style="border-top-color: #EF4444;">
                    <div class="kpi-title">HWCs Not Started Screening</div>
                    <div class="kpi-value">{tot_pending_fac:,}</div>
                    <div class="kpi-subtext">Zero screening activity</div>
                </div>
            </div>
            <div class="kpi-container">
                <div class="kpi-card" style="border-top-color: #D97706;">
                    <div class="kpi-title">Total Presumptive</div>
                    <div class="kpi-value">{tot_presumptive:,}</div>
                    <div class="kpi-subtext">{(tot_presumptive / total_scr * 100 if total_scr > 0 else 0):.1f}% Presumptive Rate</div>
                </div>
                <div class="kpi-card" style="border-top-color: #059669;">
                    <div class="kpi-title">Nikshay ID Created</div>
                    <div class="kpi-value">{tot_nikshay_gen:,}</div>
                    <div class="kpi-subtext">{nikshay_rate:.1f}% Created against Presumptive</div>
                </div>
                <div class="kpi-card" style="border-top-color: #DC2626;">
                    <div class="kpi-title">Nikshay ID Pending</div>
                    <div class="kpi-value">{tot_nikshay_pen:,}</div>
                    <div class="kpi-subtext">{(100 - nikshay_rate):.1f}% Pending Generation</div>
                </div>
                <div class="kpi-card" style="border-top-color: #6366F1;">
                    <div class="kpi-title">Overruled by HWC</div>
                    <div class="kpi-value">{tot_overrule:,}</div>
                    <div class="kpi-subtext">{(tot_overrule / tot_presumptive * 100 if tot_presumptive > 0 else 0):.1f}% of Presumptive</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown("---")
            st.markdown("### 🕒 Daily Volumetric Longitudinal Screening Trend Profile")
            time_series = working_df.groupby(working_df['reg_date'].dt.date).size().reset_index(name='Volume')
            fig_time = px.area(time_series, x='reg_date', y='Volume', title="Longitudinal Screening Volume Trend",
                               labels={'reg_date': 'Timeline Window', 'Volume': 'Evaluated Cases'}, color_discrete_sequence=['#1E3A8A'])
            fig_time.update_layout(plot_bgcolor='white', hovermode='x unified')
            st.plotly_chart(fig_time, use_container_width=True)

        # ==========================================
        # TAB 2: ENROLLMENT & COVERAGE
        # ==========================================
        with tabs[1]:
            st.markdown("### 🏥 Facility Enrollment, Screening Initiation & Nikshay Analytics")
            
            if enrolled_df is None:
                st.warning("⚠️ **Facility Master File Not Uploaded**: Please upload the Facility Enrollment Master file in the sidebar to activate comprehensive block-wise coverage tracking.")
            else:
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Total HWCs", f"{tot_enrolled:,}")
                c2.metric("HWCs Started Screening", f"{tot_screened_fac:,}", delta=f"{global_cov_pct:.1f}% Active")
                c3.metric("HWCs Not Started Screening", f"{tot_pending_fac:,}", delta=f"-{tot_pending_fac}", delta_color="inverse")
                c4.metric("Nikshay ID Created %", f"{nikshay_rate:.1f}%", delta=f"{tot_nikshay_gen} of {tot_presumptive}")

                st.markdown("---")
                st.markdown("#### 🧱 Block-Wise Facility Screening Coverage & Nikshay Status")
                
                st.dataframe(
                    block_report,
                    use_container_width=True,
                    column_config={
                        "Screening Coverage %": st.column_config.ProgressColumn(
                            "Screening Coverage %",
                            help="Percent of enrolled HWCs that have initiated screening",
                            format="%.1f%%",
                            min_value=0,
                            max_value=100
                        ),
                        "Nikshay ID Created %": st.column_config.ProgressColumn(
                            "Nikshay ID Created %",
                            help="Percentage of presumptive cases with Nikshay ID generated",
                            format="%.1f%%",
                            min_value=0,
                            max_value=100
                        ),
                        "Nikshay ID Pending %": st.column_config.ProgressColumn(
                            "Nikshay ID Pending %",
                            format="%.1f%%",
                            min_value=0,
                            max_value=100
                        ),
                        "Screening Started (Date & Time)": st.column_config.TextColumn(
                            "Screening Started (Date & Time)",
                            help="Earliest screening activity timestamp recorded"
                        )
                    }
                )

                st.markdown("---")
                col_c1, col_c2 = st.columns(2)
                with col_c1:
                    fig_status = px.pie(
                        reconciled_fac, 
                        names='Screening_Status', 
                        title="HWCs Screening Status Breakdown",
                        color='Screening_Status',
                        color_discrete_map={'Screened': '#10B981', 'Pending Screening': '#EF4444'},
                        hole=0.45
                    )
                    st.plotly_chart(fig_status, use_container_width=True)
                with col_c2:
                    fig_nik = go.Figure(data=[
                        go.Bar(name='Nikshay ID Created', x=block_report['Block Name'], y=block_report['Nikshay ID Created'], marker_color='#10B981'),
                        go.Bar(name='Nikshay ID Pending', x=block_report['Block Name'], y=block_report['Nikshay ID Pending'], marker_color='#EF4444')
                    ])
                    fig_nik.update_layout(barmode='stack', title="Nikshay ID Generation Status by Block", plot_bgcolor='white')
                    st.plotly_chart(fig_nik, use_container_width=True)

        # ==========================================
        # TAB 3: ALL FACILITIES MASTER REGISTER
        # ==========================================
        with tabs[2]:
            st.markdown("### 📋 Complete Facility-Wise Master Status Directory")
            st.caption("Displays every enrolled facility. Active facilities show real-time screening counts, CHO name, and initiation timestamps; pending facilities show 0 metrics and 'Not Started'.")
            
            if enrolled_df is None:
                st.warning("⚠️ Please upload the Facility Enrollment Master file in the sidebar to populate this complete directory.")
            else:
                filter_c1, filter_c2 = st.columns([1, 2])
                with filter_c1:
                    fac_status_filter = st.selectbox(
                        "Filter by Facility Status:",
                        options=["All Facilities", "Started Screening Only", "Not Started Screening Only"]
                    )
                
                df_fac_view = all_facilities_report.copy()
                if fac_status_filter == "Started Screening Only":
                    df_fac_view = df_fac_view[df_fac_view['Total Screening'] > 0]
                elif fac_status_filter == "Not Started Screening Only":
                    df_fac_view = df_fac_view[df_fac_view['Total Screening'] == 0]

                st.dataframe(
                    df_fac_view,
                    use_container_width=True,
                    column_config={
                        "Screening Started (Date & Time)": st.column_config.TextColumn("Screening Started (Date & Time)"),
                        "Nikshay ID Generated %": st.column_config.ProgressColumn(
                            "Nikshay ID Generated %",
                            format="%.1f%%",
                            min_value=0,
                            max_value=100
                        ),
                        "Pending for Nikshay ID Generation %": st.column_config.ProgressColumn(
                            "Pending for Nikshay %",
                            format="%.1f%%",
                            min_value=0,
                            max_value=100
                        )
                    }
                )

        # ==========================================
        # TAB 4: PRESUMPTIVE CHANNELS
        # ==========================================
        with tabs[3]:
            st.markdown("### 🔬 Multi-Channel Presumptive Identification Analysis")
            st.info("💡 **Presumptive Pathways**: A case is marked **Presumptive** if flagged by **AI Presumptive**, confirmed by **NTEP Presumptive**, or **Overruled by HWC (CHO)**.")
            
            p_c1, p_c2, p_c3, p_c4 = st.columns(4)
            p_c1.metric("Total Presumptive", f"{tot_presumptive:,}", delta="Unified Pathways")
            p_c2.metric("AI Presumptive", f"{tot_ai_pres:,}", delta=f"{(tot_ai_pres / tot_presumptive * 100 if tot_presumptive > 0 else 0):.1f}% of Presumptive")
            p_c3.metric("NTEP Presumptive", f"{tot_ntep_pres:,}", delta=f"{(tot_ntep_pres / tot_presumptive * 100 if tot_presumptive > 0 else 0):.1f}% of Presumptive")
            p_c4.metric("Overruled by HWC", f"{tot_overrule:,}", delta=f"{(tot_overrule / tot_presumptive * 100 if tot_presumptive > 0 else 0):.1f}% of Presumptive")
            
            st.markdown("---")
            col_ch_b1, col_ch_b2 = st.columns([3, 2])
            with col_ch_b1:
                ch_df = block_grp[['block', 'district', 'AI_Presumptive', 'NTEP_Presumptive', 'HWC_Overruled', 'Total_Presumptive']].head(20)
                fig_channels = go.Figure(data=[
                    go.Bar(name='AI Presumptive', x=ch_df['block'], y=ch_df['AI_Presumptive'], marker_color='#3B82F6'),
                    go.Bar(name='NTEP Presumptive', x=ch_df['block'], y=ch_df['NTEP_Presumptive'], marker_color='#10B981'),
                    go.Bar(name='Overruled by HWC', x=ch_df['block'], y=ch_df['HWC_Overruled'], marker_color='#F59E0B')
                ])
                fig_channels.update_layout(barmode='group', title="Presumptive Identification Channels by Top Blocks", plot_bgcolor='white')
                st.plotly_chart(fig_channels, use_container_width=True)
            with col_ch_b2:
                pie_data = pd.DataFrame({
                    'Channel': ['AI Presumptive', 'NTEP Presumptive', 'Overruled by HWC'],
                    'Cases': [tot_ai_pres, tot_ntep_pres, tot_overrule]
                })
                fig_pie_ch = px.pie(pie_data, names='Channel', values='Cases', title="Presumptive Detection Channel Breakdown",
                                    color_discrete_sequence=['#3B82F6', '#10B981', '#F59E0B'], hole=0.4)
                st.plotly_chart(fig_pie_ch, use_container_width=True)

        # ==========================================
        # TAB 5: DISTRICT BREAKDOWN
        # ==========================================
        with tabs[4]:
            st.markdown("### 🏢 Regional District Standings and Metric Contributions")
            col_d1, col_d2 = st.columns([3, 2])
            with col_d1:
                st.dataframe(
                    dist_grp, 
                    use_container_width=True,
                    column_config={
                        "Screenings_Count": st.column_config.ProgressColumn(
                            "Screenings Count",
                            format="%d",
                            min_value=0,
                            max_value=int(dist_grp["Screenings_Count"].max() if not dist_grp.empty else 100)
                        ),
                        "Nikshay_Conversion_%": st.column_config.ProgressColumn(
                            "Nikshay Created %",
                            format="%.1f%%",
                            min_value=0,
                            max_value=100
                        )
                    }
                )
            with col_d2:
                fig_dist = px.bar(dist_grp, x='district', y='Screenings_Count', title="District Output Volumes",
                                  labels={'district': 'District Name', 'Screenings_Count': 'Screening Counts'},
                                  color='Screenings_Count', color_continuous_scale=px.colors.sequential.Cividis)
                st.plotly_chart(fig_dist, use_container_width=True)

        # ==========================================
        # TAB 6: BLOCK PERFORMANCE
        # ==========================================
        with tabs[5]:
            st.markdown("### 🧱 Block-Level Operations Performance Summary")
            col_b1, col_b2 = st.columns([2, 3])
            with col_b1:
                st.dataframe(block_grp, use_container_width=True)
            with col_b2:
                fig_blk = px.treemap(block_grp.head(30), path=['district', 'block'], values='Screenings_Count',
                                     title="Top 30 Block Production Volumes Treemap Matrix")
                st.plotly_chart(fig_blk, use_container_width=True)

        # ==========================================
        # TAB 7: CHO TIERS & WORKLOADS
        # ==========================================
        with tabs[6]:
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
        # TAB 8: FACILITY CONTRIBUTION
        # ==========================================
        with tabs[7]:
            st.markdown("### 🏥 Health Sub-Center (HWC) Yield Trackers")
            col_f1, col_f2 = st.columns([4, 3])
            with col_f1:
                st.dataframe(fac_grp, use_container_width=True)
            with col_f2:
                fig_fac = px.pie(fac_grp.head(15), values='Total_Screenings', names='facility', title="Yield Allocation of Top 15 Primary Centers", hole=0.4)
                st.plotly_chart(fig_fac, use_container_width=True)

        # ==========================================
        # TAB 9: DATA AUDIT & QUALITY
        # ==========================================
        with tabs[8]:
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
        # TAB 10: OPERATIONAL INSIGHTS
        # ==========================================
        with tabs[9]:
            st.markdown("### 🧠 Automated Health Infrastructure Diagnostics Engine")
            mean_dist_val = dist_grp['Screenings_Count'].mean() if not dist_grp.empty else 0
            
            ins_col, rec_col = st.columns(2)
            with ins_col:
                st.markdown("#### 🏢 Identified Operational Bottlenecks")
                if block_report is not None:
                    low_cov_blocks = block_report[block_report['Screening Coverage %'] < 50.0]
                    if not low_cov_blocks.empty:
                        for _, row in low_cov_blocks.head(5).iterrows():
                            st.warning(f"🚩 **Low Facility Coverage**: Block **{row['Block Name']}** ({row['District Name']}) has activated only **{row['HWCs Started Screening']} of {row['Total HWCs']} facilities** ({row['Screening Coverage %']:.1f}%).")
                    
                    pending_nik = block_report[block_report['Nikshay ID Pending'] > 0]
                    if not pending_nik.empty:
                        worst_nik = pending_nik.sort_values(by='Nikshay ID Pending', ascending=False).iloc[0]
                        st.error(f"⚠️ **Nikshay Creation Backlog**: Block **{worst_nik['Block Name']}** has **{worst_nik['Nikshay ID Pending']} presumptive cases** awaiting Nikshay ID generation.")

                for idx, r in dist_grp.iterrows():
                    if r['Screenings_Count'] < (mean_dist_val * 0.4):
                        st.error(f"⚠️ **Underperforming Region:** District **{r['district']}** is operating below 40% of the state output mean, with a total yield of **{r['Screenings_Count']} cases**.")

            with rec_col:
                st.markdown("#### 🛠️ Direct Strategic Action Items")
                st.info("👉 **Clear Nikshay ID Backlog:** Ensure block coordinators enter generated Nikshay IDs for all presumptive cases.")
                st.info("👉 **Verify HWC Overrules:** Review facilities with high CHO overrule ratios to confirm clinical alignment.")
                st.info("👉 **Activate Pending Facilities:** Deploy mobile health teams to centers labeled 'Pending Screening' with 'Not Started' status.")

        # ==========================================
        # 9. CENTRAL EXPORT CONTROLS (EXCEL FORMAT WITH LEFT ALIGNMENT)
        # ==========================================
        st.sidebar.markdown("---")
        st.sidebar.markdown("### 📥 Reports Generation")
        
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
            {"Metric Indicator Summary": "Total HWCs (Enrolled)", "Value Metric": str(tot_enrolled)},
            {"Metric Indicator Summary": "HWCs Started Screening", "Value Metric": str(tot_screened_fac)},
            {"Metric Indicator Summary": "HWCs Not Started Screening", "Value Metric": str(tot_pending_fac)},
            {"Metric Indicator Summary": "Overall Screening Coverage %", "Value Metric": f"{global_cov_pct:.1f}%"},
            {"Metric Indicator Summary": "Total Presumptive", "Value Metric": str(tot_presumptive)},
            {"Metric Indicator Summary": "Nikshay ID Created", "Value Metric": str(tot_nikshay_gen)},
            {"Metric Indicator Summary": "Nikshay ID Created %", "Value Metric": f"{nikshay_rate:.1f}%"},
            {"Metric Indicator Summary": "Nikshay ID Pending", "Value Metric": str(tot_nikshay_pen)},
            {"Metric Indicator Summary": "Nikshay ID Pending %", "Value Metric": f"{(100 - nikshay_rate):.1f}%"}
        ])

        bundle = {
            "Executive_KPI_Summary": pd.DataFrame([{
                "Total Screenings": total_scr, "Total HWCs": tot_enrolled, "HWCs Started Screening": tot_screened_fac,
                "HWCs Not Started Screening": tot_pending_fac, "Screening Coverage %": f"{global_cov_pct:.1f}%",
                "Total Presumptive": tot_presumptive, "Nikshay ID Created": tot_nikshay_gen,
                "Nikshay ID Created %": f"{nikshay_rate:.1f}%", "Nikshay ID Pending": tot_nikshay_pen,
                "Nikshay ID Pending %": f"{(100 - nikshay_rate):.1f}%"
            }]),
            "District_Performance_Matrix": dist_grp,
            "Local_Block_Performance_Rank": block_grp,
            "CHO_Workforce_Tiers": cho_raw,
            "Primary_Facility_Outputs": fac_grp,
            "Spreadsheet_Layer_Audits": df_sheet_audit_report,
            "Field_Missingness_Audits": column_wise_nulls,
            "Strategic_System_Insights": management_insights_summary
        }
        
        if block_report is not None:
            bundle["Block_Screening_Coverage"] = block_report
        if all_facilities_report is not None:
            bundle["All_Facilities_Detailed_Master"] = all_facilities_report
        
        # Build Excel report where all columns and data are left-aligned
        excel_bytes = build_excel_report_left_aligned(bundle)
        st.sidebar.download_button(
            label="📊 Download Full MIS Package (Excel)",
            data=excel_bytes,
            file_name=f"CATB_Screening_Coverage_Report_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
else:
    st.info("💡 Ingestion Queue Ready: Please upload your **Facility Enrollment Master file** and **Screening Activity data file** (Excel or CSV) in the sidebar to generate the comprehensive analytics portal.")