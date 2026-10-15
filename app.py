"""
================================================================================
CATB Screening & Surveillance Analytics Platform (Senior Analyst Enterprise Edition)
================================================================================
- Complete Master Count: NO duplicate removals. Every hcw_id / NIN is preserved.
- Hierarchical Mapping: ID-first, fallback to District + Block + Normalized HWC.
- Master Hierarchy Audit: Complete benchmark of Started vs Not Started facilities.
- Comprehensive Line-List & KPI Tracking with CHO from Column H (Index 7).
- Dual Export: Multi-Sheet Styled Excel (.xlsx) & Landscape A4 PDF (.pdf).
"""

import datetime
import io
import re
import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# Safe import for ReportLab (prevents ModuleNotFoundError on Cloud)
REPORTLAB_AVAILABLE = False
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.pdfgen import canvas
    from reportlab.platypus import HRFlowable, Paragraph, SimpleDocTemplate, Table, TableStyle
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# ------------------------------------------------------------------------------
# 1. PAGE SETUP & ENHANCED SCROLLABLE ENTERPRISE THEME
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
        padding-left: 1.5rem;
        padding-right: 1.5rem;
        max-width: 100% !important;
    }

    /* Tab Navigation */
    .stTabs [data-baseweb="tab-list"] {
        display: flex !important;
        flex-wrap: nowrap !important;
        overflow-x: auto !important;
        overflow-y: hidden !important;
        white-space: nowrap !important;
        gap: 8px !important;
        border-bottom: 2px solid #E2E8F0 !important;
        padding-bottom: 8px !important;
        padding-top: 4px !important;
        scrollbar-width: thin !important;
        scrollbar-color: #0284C7 #F1F5F9 !important;
    }
    .stTabs [data-baseweb="tab-list"]::-webkit-scrollbar {
        height: 6px !important;
    }
    .stTabs [data-baseweb="tab-list"]::-webkit-scrollbar-thumb {
        background: #94A3B8 !important;
        border-radius: 9999px !important;
    }
    .stTabs [data-baseweb="tab"] {
        flex-shrink: 0 !important;
        white-space: nowrap !important;
        border-radius: 8px !important;
        padding: 8px 18px !important;
        font-weight: 600 !important;
        font-size: 0.85rem !important;
        color: #475569 !important;
        background-color: #F8FAFC !important;
        border: 1px solid #E2E8F0 !important;
        transition: all 0.15s ease-in-out !important;
    }
    .stTabs [data-baseweb="tab"]:hover {
        background-color: #F1F5F9 !important;
        color: #0F172A !important;
        border-color: #CBD5E1 !important;
    }
    .stTabs [aria-selected="true"] {
        background-color: #0284C7 !important;
        color: #FFFFFF !important;
        border-color: #0284C7 !important;
        box-shadow: 0 2px 4px rgba(2, 132, 199, 0.2) !important;
    }
    .stTabs [data-baseweb="tab-highlight"] {
        display: none !important;
    }

    /* Sidebar Styling */
    [data-testid="stSidebar"] {
        background-color: #F8FAFC !important;
        border-right: 1.5px solid #E2E8F0 !important;
    }
    [data-testid="stSidebarContent"] {
        padding-top: 0.6rem !important;
        padding-left: 0.8rem !important;
        padding-right: 0.8rem !important;
        overflow-y: auto !important;
    }
    .sidebar-brand {
        display: flex;
        align-items: center;
        gap: 8px;
        padding-bottom: 8px;
        border-bottom: 1.5px solid #E2E8F0;
        margin-bottom: 10px;
    }
    .sidebar-brand-title {
        font-size: 1.05rem;
        font-weight: 800;
        color: #0F172A;
        line-height: 1.1;
    }
    .sidebar-brand-sub {
        font-size: 0.72rem;
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
        padding: 5px 8px;
        border-radius: 6px;
        margin-top: 10px;
        margin-bottom: 6px;
    }
    .filter-item-label {
        font-size: 0.74rem !important;
        font-weight: 700 !important;
        color: #1E293B !important;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        margin-top: 6px !important;
        margin-bottom: 2px !important;
    }

    /* App Header Banner */
    .app-header {
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
        padding: 16px 22px;
        border-radius: 12px;
        color: white;
        margin-bottom: 16px;
        border: 1px solid #334155;
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
        gap: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .app-header-title {
        font-size: 1.35rem;
        font-weight: 800;
        margin: 0;
        display: flex;
        align-items: center;
        gap: 8px;
        color: #FFFFFF;
    }
    .app-header-subtitle {
        color: #94A3B8;
        font-size: 0.8rem;
        margin-top: 3px;
    }
    .system-status-pill {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: rgba(16, 185, 129, 0.15);
        color: #34D399;
        border: 1px solid rgba(52, 211, 153, 0.3);
        padding: 4px 12px;
        border-radius: 9999px;
        font-size: 0.74rem;
        font-weight: 600;
    }
    .pulse-dot {
        width: 7px;
        height: 7px;
        background-color: #10B981;
        border-radius: 50%;
    }

    /* Metric KPI Cards */
    .kpi-grid-card {
        background: #FFFFFF;
        border-radius: 10px;
        padding: 12px 14px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
        position: relative;
        overflow: hidden;
        transition: transform 0.15s ease-in-out, box-shadow 0.15s ease;
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }
    .kpi-grid-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 14px -3px rgba(0, 0, 0, 0.08);
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
        margin-bottom: 4px;
    }
    .kpi-label {
        font-size: 0.68rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #64748B;
    }
    .kpi-icon-bubble {
        width: 28px;
        height: 28px;
        border-radius: 6px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.95rem;
    }
    .kpi-metric-value {
        font-size: 1.45rem;
        font-weight: 800;
        color: #0F172A;
        line-height: 1.15;
    }
    .kpi-subtext {
        font-size: 0.7rem;
        font-weight: 500;
        color: #64748B;
        margin-top: 4px;
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
        font-size: 1rem;
        font-weight: 700;
        color: #1E293B;
        display: flex;
        align-items: center;
        gap: 6px;
    }
    .export-box {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 14px;
        margin-bottom: 12px;
    }
</style>
"""
st.markdown(ENTERPRISE_THEME_CSS, unsafe_allow_html=True)


# ------------------------------------------------------------------------------
# 2. UTILITY & NORMALIZATION FUNCTIONS
# ------------------------------------------------------------------------------
def clean_header(val: str) -> str:
    if not isinstance(val, str):
        return ""
    return re.sub(r"[_\-\s]+", " ", str(val)).strip().lower()


def clean_name(val: str) -> str:
    """Robust string cleanup for facility and geography matching."""
    if pd.isna(val):
        return "Unknown"
    v = str(val).strip()
    v = re.sub(r"\s+", " ", v)
    return v.title() if v else "Unknown"


def normalize_facility_key(name: str) -> str:
    """
    Intelligently cleans HWC names by stripping prefixes like 'HWC/SUD, ', 'SC-', 'Ayush-', etc.
    Allows exact matching between 'HWC/SUD, PATWAI' and 'Patwai'.
    """
    if pd.isna(name):
        return ""
    s = str(name).lower()
    s = re.sub(r"\b(hwc|sc|phc|chc|sud|ayush|sub center|subcentre|health and wellness center)\b", " ", s)
    s = re.sub(r"[^a-z0-9]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def make_composite_match_key(dist: str, blk: str, fac: str) -> str:
    """Hierarchy-safe composite key: District + Block + Normalized Facility."""
    d = str(dist).strip().lower() if pd.notna(dist) else ""
    b = str(blk).strip().lower() if pd.notna(blk) else ""
    f = normalize_facility_key(fac)
    return f"{d}:::{b}:::{f}"


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


# ------------------------------------------------------------------------------
# 3. PDF GENERATION ENGINE (REPORTLAB WITH FAIL-SAFE FALLBACK)
# ------------------------------------------------------------------------------
if REPORTLAB_AVAILABLE:
    class NumberedCanvas(canvas.Canvas):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self._saved_page_states = []

        def showPage(self):
            self._saved_page_states.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            num_pages = len(self._saved_page_states)
            for state in self._saved_page_states:
                self.__dict__.update(state)
                self.draw_page_decorations(num_pages)
                super().showPage()
            super().save()

        def draw_page_decorations(self, page_count):
            self.saveState()
            self.setFont("Helvetica", 8)
            self.setFillColor(colors.HexColor("#64748B"))
            self.setStrokeColor(colors.HexColor("#CBD5E1"))
            self.setLineWidth(0.5)
            self.line(30, 565, 812, 565)
            self.drawString(30, 570, "National Tuberculosis Elimination Programme (NTEP) | CATB Active Surveillance Dossier")
            self.line(30, 32, 812, 32)
            generated_on = datetime.datetime.now().strftime("%d-%b-%Y %H:%M:%S")
            self.drawString(30, 22, f"Official Surveillance Dossier | Generated on: {generated_on}")
            self.drawRightString(812, 22, f"Page {self._pageNumber} of {page_count}")
            self.restoreState()


def generate_standard_pdf(df_data: pd.DataFrame, title: str, subtitle: str) -> bytes:
    """Generates Landscape A4 PDF using ReportLab or fallback plain-text buffer."""
    if not REPORTLAB_AVAILABLE:
        fallback_str = f"=== {title} ===\n{subtitle}\n\n"
        fallback_str += df_data.to_string(index=False)
        return fallback_str.encode("utf-8")

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(A4),
        leftMargin=30,
        rightMargin=30,
        topMargin=45,
        bottomMargin=45,
    )
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "DocTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=15,
        leading=18,
        textColor=colors.HexColor("#0F172A"),
        spaceAfter=3,
    )
    sub_style = ParagraphStyle(
        "DocSub",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#475569"),
        spaceAfter=12,
    )
    cell_style = ParagraphStyle(
        "TableCell",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7,
        leading=8.5,
        textColor=colors.HexColor("#1E293B"),
        alignment=0,
    )
    cell_num_style = ParagraphStyle(
        "TableNumCell",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7,
        leading=8.5,
        textColor=colors.HexColor("#1E293B"),
        alignment=2,
    )
    header_style = ParagraphStyle(
        "TableHeader",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=7,
        leading=8.5,
        textColor=colors.white,
        alignment=1,
    )

    story = [
        Paragraph(title, title_style),
        Paragraph(subtitle, sub_style),
        HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#0284C7"), spaceAfter=10),
    ]

    columns = list(df_data.columns)
    table_data = [[Paragraph(str(col), header_style) for col in columns]]

    for _, row in df_data.iterrows():
        row_cells = []
        for col in columns:
            val = row[col]
            if pd.isna(val) or val == "":
                val_str = "-"
            elif isinstance(val, (int, np.integer)):
                val_str = f"{val:,}"
            elif isinstance(val, (float, np.floating)):
                val_str = f"{val:,.1f}"
            else:
                val_str = str(val)

            if isinstance(val, (int, float, np.number)) and not str(col).startswith("S."):
                row_cells.append(Paragraph(val_str, cell_num_style))
            else:
                row_cells.append(Paragraph(val_str, cell_style))
        table_data.append(row_cells)

    avail_width = 782.0
    num_cols = len(columns)
    col_widths = [avail_width / num_cols] * num_cols

    if "HWC" in columns and "CHO" in columns:
        col_widths = []
        for col in columns:
            if col in ["S. No."]:
                col_widths.append(26)
            elif col in ["District", "Block"]:
                col_widths.append(48)
            elif col in ["HWC"]:
                col_widths.append(70)
            elif col in ["CHO"]:
                col_widths.append(65)
            elif col in ["Screening Status", "Screening Started Date"]:
                col_widths.append(50)
            else:
                col_widths.append(34)
        total_w = sum(col_widths)
        col_widths = [(w / total_w) * avail_width for w in col_widths]

    report_table = Table(table_data, colWidths=col_widths, repeatRows=1)
    report_table.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0F172A")),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("INNERGRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#E2E8F0")),
            ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#CBD5E1")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#FFFFFF"), colors.HexColor("#F8FAFC")]),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ])
    )
    story.append(report_table)
    doc.build(story, canvasmaker=NumberedCanvas)
    return buf.getvalue()


def generate_styled_excel(sheets_dict: dict) -> bytes:
    """Creates a formatted multi-sheet Excel workbook."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets_dict.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
            ws = writer.sheets[sheet_name[:31]]

            header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
            header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            border_thin = Border(
                left=Side(style="thin", color="E2E8F0"),
                right=Side(style="thin", color="E2E8F0"),
                top=Side(style="thin", color="E2E8F0"),
                bottom=Side(style="thin", color="E2E8F0"),
            )
            data_font = Font(name="Calibri", size=9.5)
            align_center = Alignment(horizontal="center", vertical="center")
            align_right = Alignment(horizontal="right", vertical="center")
            align_left = Alignment(horizontal="left", vertical="center")

            for col_idx in range(1, len(df.columns) + 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = align_center

            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                for cell in row:
                    cell.font = data_font
                    cell.border = border_thin
                    if isinstance(cell.value, (int, float)):
                        cell.alignment = align_right
                        cell.number_format = "#,##0"
                    else:
                        cell.alignment = align_left

            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    val_str = str(cell.value or "")
                    if len(val_str) > max_len:
                        max_len = len(val_str)
                ws.column_dimensions[col_letter].width = max(max_len + 4, 11)
            ws.views.sheetView[0].showGridLines = True
    return output.getvalue()


# ------------------------------------------------------------------------------
# 4. PRIMARY DATA INGESTION ENGINE (SENIOR ANALYST SPEC)
# ------------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_and_standardize_data(file_all_bytes, file_all_name, file_hier_bytes=None, file_hier_name=None):
    # 1. Ingest Screening Data Sheet
    if file_all_name.lower().endswith(".csv"):
        df_all = pd.read_csv(io.BytesIO(file_all_bytes), low_memory=False)
    else:
        df_all = pd.read_excel(io.BytesIO(file_all_bytes))

    cols_all = list(df_all.columns)

    # Column A (Index 0) - Registration Date
    if len(cols_all) > 0 and ("date" in clean_header(str(cols_all[0])) or "reg" in clean_header(str(cols_all[0]))):
        c_date = cols_all[0]
    else:
        c_date = detect_field(cols_all, ["reg date", "registration date", "screening date", "date", "created at"])

    c_id = detect_field(cols_all, ["id", "uuid", "patient id", "beneficiary id"])
    c_dist = detect_field(cols_all, ["district", "district name"])
    c_block = detect_field(cols_all, ["block", "block name", "tehsil"])

    # Column E (Index 4) - HWC / Facility Name
    if len(cols_all) > 4 and "hwc" in clean_header(str(cols_all[4])):
        c_hwc = cols_all[4]
    else:
        c_hwc = detect_field(cols_all, ["hwc", "hwc name", "facility", "facility name", "health center"])

    # Column H (Index 7) - CHO Name
    if len(cols_all) > 7 and any(k in clean_header(str(cols_all[7])) for k in ["cho", "name", "staff", "provider"]):
        c_cho = cols_all[7]
    elif len(cols_all) > 7 and ("unnamed" not in str(cols_all[7]).lower()):
        c_cho = cols_all[7]
    else:
        c_cho = detect_field(cols_all, ["cho name", "cho", "staff name", "provider name", "health worker", "cho_name"])

    # Column L (Index 11) - AI Preference
    if len(cols_all) > 11 and "ai" in clean_header(str(cols_all[11])):
        c_ai = cols_all[11]
    else:
        c_ai = detect_field(cols_all, ["ai preference", "ai result", "ai interpretation", "ai"])

    # Column Q (Index 16) - NTEP Result
    if len(cols_all) > 16 and "ntep" in clean_header(str(cols_all[16])):
        c_ntep = cols_all[16]
    else:
        c_ntep = detect_field(cols_all, ["ntep result", "ntep interpretation", "ntep"])

    # Column T (Index 19) - TB Status
    if len(cols_all) > 19 and "status" in clean_header(str(cols_all[19])):
        c_status = cols_all[19]
    else:
        c_status = detect_field(cols_all, ["tb status", "tb_status", "status", "case status"])

    # Column AC (Index 28) - Type (Symptomatic vs Asymptomatic)
    if len(cols_all) > 28 and "type" in clean_header(str(cols_all[28])):
        c_type = cols_all[28]
    else:
        c_type = detect_field(cols_all, ["type", "patient type", "symptom type", "clinical type"])

    # Column BA (Index 52) - Overrule by HWC
    if len(cols_all) > 52 and ("overrule" in clean_header(str(cols_all[52])) or "override" in clean_header(str(cols_all[52]))):
        c_overrule = cols_all[52]
    else:
        c_overrule = detect_field(cols_all, ["overrule by hwc", "overrule", "hwc overrule", "cho override", "override"])

    # Check for Optional Facility IDs in Screening Sheet
    c_hcw_id = detect_field(cols_all, ["hcw_id", "hcw id", "hwc id", "facility id"])
    c_nin = detect_field(cols_all, ["nin", "nin id", "national id"])

    df = df_all.copy()

    # Parse Dates
    if c_date and c_date in df.columns:
        df["reg_date_clean"] = pd.to_datetime(df[c_date], errors="coerce")
    else:
        df["reg_date_clean"] = pd.NaT

    df["District_Clean"] = df[c_dist].apply(clean_name) if c_dist else "Unknown"
    df["Block_Clean"] = df[c_block].apply(clean_name) if c_block else "Unknown"
    df["Facility_Clean"] = df[c_hwc].apply(clean_name) if c_hwc else "Unknown"
    
    df["norm_hwc"] = df["Facility_Clean"].apply(normalize_facility_key)
    df["composite_key"] = df.apply(lambda r: make_composite_match_key(r["District_Clean"], r["Block_Clean"], r["Facility_Clean"]), axis=1)

    df["hcw_id_clean"] = df[c_hcw_id].astype(str).str.strip() if c_hcw_id else ""
    df["nin_clean"] = df[c_nin].astype(str).str.strip() if c_nin else ""

    if c_cho and c_cho in df.columns:
        df["CHO_Clean"] = df[c_cho].astype(str).str.strip().str.title()
        invalid_tokens = ["", "Nan", "None", "Null", "0", "Na", "N/A", "0.0", "-", "--"]
        df.loc[df["CHO_Clean"].isin(invalid_tokens), "CHO_Clean"] = "Not Available"
    else:
        df["CHO_Clean"] = "Not Available"

    df["str_id"] = df[c_id].astype(str).str.strip() if c_id else df.index.astype(str)

    # Symptomatic vs Asymptomatic
    type_series = df[c_type].fillna("").astype(str).str.strip().str.lower() if c_type and c_type in df.columns else pd.Series([""] * len(df))
    df["is_asymptomatic"] = (type_series.str.contains("asymp") | type_series.str.contains("assymp")).astype(int)
    df["is_symptomatic"] = ((type_series.str.contains("symp")) & (df["is_asymptomatic"] == 0)).astype(int)

    # AI & NTEP Presumptive
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

    # TB Progression Status
    status_series = df[c_status].fillna("").astype(str).str.strip().str.upper() if c_status and c_status in df.columns else pd.Series([""] * len(df))
    df["is_presumptive_open"] = (status_series == "PRESUMPTIVE_OPEN").astype(int)
    df["is_presumptive_closed"] = (status_series == "PRESUMPTIVE_CLOSED").astype(int)
    df["is_diagnosed"] = (status_series == "DIAGNOSED_ON_TREATMENT").astype(int)
    df["is_tested"] = (status_series.isin(["PRESUMPTIVE_CLOSED", "DIAGNOSED_ON_TREATMENT"])).astype(int)

    # 2. Ingest Master Hierarchy (NO DUPLICATE DROPPING - COUNT EVERY ROW)
    df_master_hwc = None
    if file_hier_bytes:
        try:
            if file_hier_name.lower().endswith(".csv"):
                df_h = pd.read_csv(io.BytesIO(file_hier_bytes), low_memory=False)
            else:
                df_h = pd.read_excel(io.BytesIO(file_hier_bytes))

            df_h.columns = [str(c).strip() for c in df_h.columns]
            h_cols = list(df_h.columns)

            h_dist = detect_field(h_cols, ["district", "district name"])
            h_block = detect_field(h_cols, ["block", "block name", "tehsil"])
            h_fac = detect_field(h_cols, ["hwc", "facility", "facility name", "health center"])
            h_id = detect_field(h_cols, ["hcw_id", "hcw id", "hwc id", "facility id"])
            h_nin = detect_field(h_cols, ["nin", "nin id", "national id"])

            if h_fac:
                df_h = df_h[df_h[h_fac].notna()].copy()
                df_h["Facility_Clean"] = df_h[h_fac].apply(clean_name)
                df_h["District_Clean"] = df_h[h_dist].apply(clean_name) if h_dist else "Unknown"
                df_h["Block_Clean"] = df_h[h_block].apply(clean_name) if h_block else "Unknown"
                
                df_h["norm_hwc"] = df_h["Facility_Clean"].apply(normalize_facility_key)
                df_h["composite_key"] = df_h.apply(lambda r: make_composite_match_key(r["District_Clean"], r["Block_Clean"], r["Facility_Clean"]), axis=1)

                df_h["hcw_id_clean"] = df_h[h_id].astype(str).str.strip() if h_id else ""
                df_h["nin_clean"] = df_h[h_nin].astype(str).str.strip() if h_nin else ""

                # Unique Master Identifier per row to protect duplicates with distinct IDs
                df_h["master_uid"] = np.arange(1, len(df_h) + 1)
                df_master_hwc = df_h.copy()  # COMPLETE COUNT PRESERVED
        except Exception:
            pass

    return df, df_master_hwc


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
# 5. MAIN SURVEILLANCE APPLICATION
# ------------------------------------------------------------------------------
def main():
    if not REPORTLAB_AVAILABLE:
        st.warning("⚠️ **ReportLab missing.** Exported PDFs will use clean text mode. Add `reportlab` to `requirements.txt` to enable styled PDFs.")

    with st.sidebar:
        st.markdown(
            """
            <div class="sidebar-brand">
                <span style="font-size: 1.4rem;">🩺</span>
                <div>
                    <div class="sidebar-brand-title">CATB Command</div>
                    <div class="sidebar-brand-sub">Master Facility Surveillance</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown('<div class="sidebar-category-header">📁 1. Data Ingestion</div>', unsafe_allow_html=True)
        st.markdown('<div class="filter-item-label">Screening Data Sheet</div>', unsafe_allow_html=True)
        up_all = st.file_uploader("Screening Data Sheet", type=["xlsx", "xls", "csv"], key="file_screening", label_visibility="collapsed")

        st.markdown('<div class="filter-item-label">HWC Master Hierarchy (Column E)</div>', unsafe_allow_html=True)
        up_hier = st.file_uploader("HWC Master Hierarchy", type=["xlsx", "xls", "csv"], key="file_hierarchy", label_visibility="collapsed")

    if not up_all:
        st.markdown(
            """
            <div class="app-header">
                <div>
                    <h1 class="app-header-title"><span>🩺</span> CATB Screening & Clinical Surveillance Platform</h1>
                    <div class="app-header-subtitle">Active Case Finding (ACF) Intelligence & Master Facility Audit</div>
                </div>
                <div class="system-status-pill">
                    <div class="pulse-dot"></div> Awaiting Screening Data
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.info("👈 कृपया डैशबोर्ड शुरू करने के लिए साइडबार में **Screening Data Sheet** अपलोड करें।")
        return

    with st.spinner("⚡ Auditing complete facility rosters & clinical screening indicators..."):
        df_screening, df_master_hwc = load_and_standardize_data(
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

        dates_valid = df_screening["reg_date_clean"].dropna()
        if not dates_valid.empty:
            min_avail, max_avail = dates_valid.min().date(), dates_valid.max().date()
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.markdown('<div class="filter-item-label">Start Date</div>', unsafe_allow_html=True)
                start_date = st.date_input("Start Date", value=min_avail, min_value=min_avail, max_value=max_avail, label_visibility="collapsed")
            with col_d2:
                st.markdown('<div class="filter-item-label">End Date</div>', unsafe_allow_html=True)
                end_date = st.date_input("End Date", value=max_avail, min_value=min_avail, max_value=max_avail, label_visibility="collapsed")
        else:
            start_date, end_date = None, None

        # Base Reference for Geography Dropdowns (Master List if present, else Screening)
        geo_source = df_master_hwc if df_master_hwc is not None else df_screening

        avail_districts = sorted([d for d in geo_source["District_Clean"].unique() if d != "Unknown"])
        st.markdown('<div class="filter-item-label">Select District</div>', unsafe_allow_html=True)
        sel_districts = st.multiselect("Select District", options=avail_districts, default=[], placeholder="All Districts", label_visibility="collapsed")

        geo_b = geo_source[geo_source["District_Clean"].isin(sel_districts)] if sel_districts else geo_source
        avail_blocks = sorted([b for b in geo_b["Block_Clean"].unique() if b != "Unknown"])
        st.markdown('<div class="filter-item-label">Select Block</div>', unsafe_allow_html=True)
        sel_blocks = st.multiselect("Select Block", options=avail_blocks, default=[], placeholder="All Blocks", label_visibility="collapsed")

        geo_h = geo_b[geo_b["Block_Clean"].isin(sel_blocks)] if sel_blocks else geo_b
        avail_hwcs = sorted([h for h in geo_h["Facility_Clean"].unique() if h != "Unknown"])
        st.markdown('<div class="filter-item-label">Select HWC</div>', unsafe_allow_html=True)
        sel_hwcs = st.multiselect("Select HWC", options=avail_hwcs, default=[], placeholder="All HWCs", label_visibility="collapsed")

        # CHO Filter (Screening Sheet Column H)
        cho_pool = df_screening[df_screening["Facility_Clean"].isin(sel_hwcs)] if sel_hwcs else df_screening
        avail_chos = sorted([c for c in cho_pool["CHO_Clean"].unique() if c not in ["Not Available", "Unknown"]])
        st.markdown('<div class="filter-item-label">Select CHO (Col H)</div>', unsafe_allow_html=True)
        sel_chos = st.multiselect("Select CHO", options=avail_chos, default=[], placeholder="All CHOs", label_visibility="collapsed")

        st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)
        if st.button("🔄 Reset All Filters", use_container_width=True):
            st.rerun()

    # --------------------------------------------------------------------------
    # APPLY FILTERS TO SCREENING DATASET
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
    # FACILITY SCREENING AGGREGATION & ROBUST HIERARCHICAL MATCHING
    # --------------------------------------------------------------------------
    # Aggregate Screening Data at Facility Level
    active_facility_summary = (
        f_screen.groupby("composite_key")
        .agg(
            District=("District_Clean", "first"),
            Block=("Block_Clean", "first"),
            HWC=("Facility_Clean", "first"),
            norm_hwc=("norm_hwc", "first"),
            Total_Screening=("composite_key", "count"),
            Total_Symptomatic=("is_symptomatic", "sum"),
            Total_Asymptomatic=("is_asymptomatic", "sum"),
            AI_Presumptive=("is_ai_pres", "sum"),
            NTEP_Presumptive=("is_ntep_pres", "sum"),
            NTEP_Only_Presumptive=("is_ntep_only", "sum"),
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

    # Active CHO Mapping per composite key
    cho_valid = f_screen[f_screen["CHO_Clean"] != "Not Available"]
    if not cho_valid.empty:
        cho_mapping = (
            cho_valid.groupby("composite_key")["CHO_Clean"]
            .agg(lambda s: s.value_counts().index[0] if len(s.value_counts()) > 0 else "Not Available")
            .reset_index()
            .rename(columns={"CHO_Clean": "Active_CHO_Name"})
        )
    else:
        cho_mapping = pd.DataFrame(columns=["composite_key", "Active_CHO_Name"])

    # Earliest Date Mapping per composite key
    valid_dates_df = f_screen.dropna(subset=["reg_date_clean"])
    if not valid_dates_df.empty:
        earliest_date_mapping = (
            valid_dates_df.groupby("composite_key")["reg_date_clean"]
            .min()
            .reset_index()
            .rename(columns={"reg_date_clean": "Earliest_Date"})
        )
        earliest_date_mapping["Screening_Started_Date_Calc"] = earliest_date_mapping["Earliest_Date"].dt.strftime("%d-%b-%Y")
    else:
        earliest_date_mapping = pd.DataFrame(columns=["composite_key", "Earliest_Date", "Screening_Started_Date_Calc"])

    active_facility_summary = pd.merge(active_facility_summary, cho_mapping, on="composite_key", how="left")
    active_facility_summary = pd.merge(active_facility_summary, earliest_date_mapping[["composite_key", "Screening_Started_Date_Calc"]], on="composite_key", how="left")
    active_facility_summary["Active_CHO_Name"] = active_facility_summary["Active_CHO_Name"].fillna("Not Available")
    active_facility_summary["Screening_Started_Date_Calc"] = active_facility_summary["Screening_Started_Date_Calc"].fillna("Not Available")

    # --------------------------------------------------------------------------
    # COMPLETE MASTER AUDIT INTEGRATION (PRESERVING EVERY MASTER ROW)
    # --------------------------------------------------------------------------
    if df_master_hwc is not None:
        master_filtered = df_master_hwc.copy()
        if sel_districts:
            master_filtered = master_filtered[master_filtered["District_Clean"].isin(sel_districts)]
        if sel_blocks:
            master_filtered = master_filtered[master_filtered["Block_Clean"].isin(sel_blocks)]
        if sel_hwcs:
            master_filtered = master_filtered[master_filtered["Facility_Clean"].isin(sel_hwcs)]

        total_master_count = len(master_filtered)

        # Merge Step 1: Match by Hierarchical Composite Key (District + Block + Norm HWC)
        merged_master = pd.merge(
            master_filtered,
            active_facility_summary,
            on="composite_key",
            how="left",
            suffixes=("", "_scr")
        )

        # Merge Step 2 Fallback: If Block slightly differs, match within same District by norm_hwc
        unmatched_mask = merged_master["Total_Screening"].isna()
        if unmatched_mask.any():
            fallback_map = active_facility_summary.set_index(["District", "norm_hwc"])
            for idx in merged_master[unmatched_mask].index:
                d_val = merged_master.at[idx, "District_Clean"]
                n_val = merged_master.at[idx, "norm_hwc"]
                if (d_val, n_val) in fallback_map.index:
                    matched_row = fallback_map.loc[[(d_val, n_val)]].iloc[0]
                    for col_name in [
                        "Total_Screening", "Total_Symptomatic", "Total_Asymptomatic",
                        "AI_Presumptive", "NTEP_Presumptive", "NTEP_Only_Presumptive",
                        "Both_AI_NTEP", "CHO_Override", "Total_Presumptive",
                        "Presumptive_Open", "Presumptive_Closed", "Total_Tested", "Total_Diagnosed"
                    ]:
                        merged_master.at[idx, col_name] = matched_row[col_name]
                    merged_master.at[idx, "Active_CHO_Name"] = matched_row["Active_CHO_Name"]
                    merged_master.at[idx, "Screening_Started_Date_Calc"] = matched_row["Screening_Started_Date_Calc"]

        # Fill Inactive Values
        num_cols = [
            "Total_Screening", "Total_Symptomatic", "Total_Asymptomatic",
            "AI_Presumptive", "NTEP_Presumptive", "NTEP_Only_Presumptive",
            "Both_AI_NTEP", "CHO_Override", "Total_Presumptive",
            "Presumptive_Open", "Presumptive_Closed", "Total_Tested", "Total_Diagnosed"
        ]
        for c in num_cols:
            merged_master[c] = merged_master[c].fillna(0).astype(int)

        merged_master["CHO"] = merged_master["Active_CHO_Name"].fillna("Not Available")
        merged_master.loc[merged_master["Total_Screening"] == 0, "CHO"] = "Not Available"

        merged_master["Screening Status"] = np.where(merged_master["Total_Screening"] > 0, "Screening Started", "Screening Not Started")
        merged_master["Screening Started Date"] = merged_master["Screening_Started_Date_Calc"].fillna("Not Available")
        merged_master.loc[merged_master["Total_Screening"] == 0, "Screening Started Date"] = "Not Available"

        merged_master["District"] = merged_master["District_Clean"]
        merged_master["Block"] = merged_master["Block_Clean"]
        merged_master["HWC"] = merged_master["Facility_Clean"]

        # Check for any screening facilities not present in Master List and append them
        matched_comp_keys = set(merged_master["composite_key"].unique())
        unmatched_active = active_facility_summary[~active_facility_summary["composite_key"].isin(matched_comp_keys)].copy()
        if not unmatched_active.empty:
            unmatched_active["Screening Status"] = "Screening Started"
            unmatched_active["CHO"] = unmatched_active["Active_CHO_Name"]
            unmatched_active["Screening Started Date"] = unmatched_active["Screening_Started_Date_Calc"]
            final_facility_master = pd.concat([merged_master, unmatched_active], ignore_index=True)
        else:
            final_facility_master = merged_master.copy()

        # Audit Counts
        screening_started_hwcs = int((final_facility_master["Screening Status"] == "Screening Started").sum())
        screening_not_started_hwcs = int((final_facility_master["Screening Status"] == "Screening Not Started").sum())
        total_hwc_units = len(final_facility_master)

        # Inactive Report
        inactive_hwc_report = final_facility_master[final_facility_master["Screening Status"] == "Screening Not Started"][
            ["District", "Block", "HWC", "CHO", "Screening Status", "Total_Screening"]
        ].sort_values(by=["District", "Block", "HWC"]).reset_index(drop=True)
        inactive_hwc_report.rename(columns={"Total_Screening": "Total Screening"}, inplace=True)
        inactive_hwc_report.insert(0, "S. No.", inactive_hwc_report.index + 1)

    else:
        # Fallback if Master sheet is not uploaded
        final_facility_master = active_facility_summary.copy()
        final_facility_master["CHO"] = final_facility_master["Active_CHO_Name"]
        final_facility_master["Screening Status"] = "Screening Started"
        final_facility_master["Screening Started Date"] = final_facility_master["Screening_Started_Date_Calc"]
        total_hwc_units = len(final_facility_master)
        screening_started_hwcs = total_hwc_units
        screening_not_started_hwcs = 0
        inactive_hwc_report = pd.DataFrame(columns=["S. No.", "District", "Block", "HWC", "CHO", "Screening Status", "Total Screening"])

    # --------------------------------------------------------------------------
    # DASHBOARD GLOBAL CLINICAL KPIS (FROM SCREENING DATA)
    # --------------------------------------------------------------------------
    total_screenings = int(f_screen.shape[0])
    total_symptomatic = int(f_screen["is_symptomatic"].sum()) if "is_symptomatic" in f_screen.columns else 0
    total_asymptomatic = int(f_screen["is_asymptomatic"].sum()) if "is_asymptomatic" in f_screen.columns else 0

    total_ai_pres = int(f_screen["is_ai_pres"].sum()) if "is_ai_pres" in f_screen.columns else 0
    total_ntep_pres = int(f_screen["is_ntep_pres"].sum()) if "is_ntep_pres" in f_screen.columns else 0
    total_ntep_only_pres = int(f_screen["is_ntep_only"].sum()) if "is_ntep_only" in f_screen.columns else 0
    total_both_pres = int(f_screen["is_both_pres"].sum()) if "is_both_pres" in f_screen.columns else 0
    total_cho_override = int(f_screen["is_cho_override"].sum()) if "is_cho_override" in f_screen.columns else 0
    total_presumptive = int(f_screen["is_total_presumptive"].sum()) if "is_total_presumptive" in f_screen.columns else 0

    total_presumptive_open = int(f_screen["is_presumptive_open"].sum()) if "is_presumptive_open" in f_screen.columns else 0
    total_presumptive_closed = int(f_screen["is_presumptive_closed"].sum()) if "is_presumptive_closed" in f_screen.columns else 0
    total_tested = int(f_screen["is_tested"].sum()) if "is_tested" in f_screen.columns else 0
    total_diagnosed = int(f_screen["is_diagnosed"].sum()) if "is_diagnosed" in f_screen.columns else 0

    pct_symptomatic = (total_symptomatic / max(total_screenings, 1)) * 100
    pct_asymptomatic = (total_asymptomatic / max(total_screenings, 1)) * 100
    pres_yield = (total_presumptive / max(total_screenings, 1)) * 100
    diagnosis_yield = (total_diagnosed / max(total_tested, 1)) * 100 if total_tested > 0 else 0.0

    # --------------------------------------------------------------------------
    # HEADER BANNER
    # --------------------------------------------------------------------------
    current_time_str = datetime.datetime.now().strftime("%d %b %Y | %H:%M:%S")
    master_info_label = f"Master HWCs (Complete Count): <b>{total_hwc_units:,}</b>" if df_master_hwc is not None else "Master Hierarchy: <i>Not Uploaded</i>"
    st.markdown(
        f"""
        <div class="app-header">
            <div>
                <h1 class="app-header-title"><span>🩺</span> CATB Screening & Clinical Surveillance Platform</h1>
                <div class="app-header-subtitle">
                    Jurisdiction: <b>{f"{len(sel_districts)} Districts Selected" if sel_districts else "All Monitored Districts"}</b> &nbsp;|&nbsp; 
                    {master_info_label} &nbsp;|&nbsp; 
                    Active Facilities: <b>{screening_started_hwcs:,}</b> &nbsp;|&nbsp; 
                    CHO Column: <b>Column H (Verified)</b>
                </div>
            </div>
            <div class="system-status-pill">
                <div class="pulse-dot"></div> Live Surveillance Active
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # --------------------------------------------------------------------------
    # KPI METRIC CARDS
    # --------------------------------------------------------------------------
    st.markdown('<div class="section-title-wrap"><div class="section-title-text"><span>🏛️</span> 1. Screening Overview & Complete Facility Coverage</div></div>', unsafe_allow_html=True)
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    render_enterprise_kpi(c1, "Total HWCs", f"{total_hwc_units:,}", "Complete Count (Col E)", "🏥", "#0284C7", "#E0F2FE")
    render_enterprise_kpi(c2, "Screening Started", f"{screening_started_hwcs:,}", f"{(screening_started_hwcs/max(total_hwc_units,1)*100):.1f}% active facilities", "✅", "#16A34A", "#DCFCE7")
    render_enterprise_kpi(c3, "Not Started", f"{screening_not_started_hwcs:,}", f"{(screening_not_started_hwcs/max(total_hwc_units,1)*100):.1f}% inactive facilities", "⏳", "#DC2626", "#FEE2E2")
    render_enterprise_kpi(c4, "Total Screened", f"{total_screenings:,}", "Evaluated individuals", "📋", "#0D9488", "#CCFBF1")
    render_enterprise_kpi(c5, "Symptomatic", f"{total_symptomatic:,}", f"{pct_symptomatic:.1f}% screened (Col AC)", "🤒", "#D97706", "#FEF3C7")
    render_enterprise_kpi(c6, "Asymptomatic", f"{total_asymptomatic:,}", f"{pct_asymptomatic:.1f}% screened (Col AC)", "🛡️", "#0369A1", "#E0F2FE")

    st.markdown('<div class="section-title-wrap"><div class="section-title-text"><span>🔬</span> 2. Presumptive Classification Breakdown</div></div>', unsafe_allow_html=True)
    c7, c8, c9, c10, c11, c12 = st.columns(6)
    render_enterprise_kpi(c7, "AI Presumptive", f"{total_ai_pres:,}", "AI Algorithm (Col L)", "🤖", "#8B5CF6", "#F3E8FF")
    render_enterprise_kpi(c8, "NTEP Presumptive", f"{total_ntep_pres:,}", "Protocol (Col Q)", "📋", "#D97706", "#FEF3C7")
    render_enterprise_kpi(c9, "NTEP-Only Pres.", f"{total_ntep_only_pres:,}", "NTEP (+), AI (-)", "🔬", "#EA580C", "#FFEDD5")
    render_enterprise_kpi(c10, "Both Presumptive", f"{total_both_pres:,}", "Dual Positive Overlap", "🤝", "#0284C7", "#E0F2FE")
    render_enterprise_kpi(c11, "CHO Override", f"{total_cho_override:,}", "Frontline Override (Col BA)", "👩‍⚕️", "#EC4899", "#FCE7F3")
    render_enterprise_kpi(c12, "Total Unique Pres.", f"{total_presumptive:,}", f"Yield: {pres_yield:.1f}%", "⚠️", "#EF4444", "#FEE2E2")

    st.markdown('<div class="section-title-wrap"><div class="section-title-text"><span>📈</span> 3. Continuum of Care & Clinical Outcomes</div></div>', unsafe_allow_html=True)
    c13, c14, c15, c16 = st.columns(4)
    render_enterprise_kpi(c13, "Presumptive Open", f"{total_presumptive_open:,}", "PRESUMPTIVE_OPEN (Nikshay Created)", "🆔", "#2563EB", "#DBEAFE")
    render_enterprise_kpi(c14, "Presumptive Closed", f"{total_presumptive_closed:,}", "PRESUMPTIVE_CLOSED (Negative/Tested)", "🔒", "#475569", "#F1F5F9")
    render_enterprise_kpi(c15, "Total Tested", f"{total_tested:,}", f"Tested Cases ({total_tested:,})", "🧪", "#7C3AED", "#EDE9FE")
    render_enterprise_kpi(c16, "Total Diagnosed", f"{total_diagnosed:,}", f"Diagnosed on Treatment ({diagnosis_yield:.1f}%)", "🩺", "#059669", "#D1FAE5")

    # --------------------------------------------------------------------------
    # FORMATTED REPORTS PREPARATION
    # --------------------------------------------------------------------------
    # 1. Complete Facility-Wise Report
    report_hwc = final_facility_master.sort_values(by=["Total_Screening", "Total_Tested"], ascending=[False, False]).reset_index(drop=True)
    report_hwc["S. No."] = report_hwc.index + 1

    ordered_hwc_cols = [
        "S. No.", "District", "Block", "HWC", "CHO", "Screening Status",
        "Total_Screening", "Total_Symptomatic", "Total_Asymptomatic",
        "AI_Presumptive", "NTEP_Presumptive", "NTEP_Only_Presumptive", "Both_AI_NTEP",
        "CHO_Override", "Total_Presumptive", "Presumptive_Open", "Presumptive_Closed",
        "Total_Tested", "Total_Diagnosed", "Screening Started Date"
    ]

    # Add optional IDs if available in master list
    for opt_col in ["hcw_id", "NIN"]:
        if opt_col in report_hwc.columns:
            ordered_hwc_cols.insert(4, opt_col)

    final_hwc_report = report_hwc[ordered_hwc_cols].rename(
        columns={
            "Total_Screening": "Total Screening",
            "Total_Symptomatic": "Symptomatic",
            "Total_Asymptomatic": "Asymptomatic",
            "AI_Presumptive": "AI Presumptive",
            "NTEP_Presumptive": "NTEP Presumptive",
            "NTEP_Only_Presumptive": "NTEP-Only Presumptive",
            "Both_AI_NTEP": "AI + NTEP Both",
            "CHO_Override": "CHO Override",
            "Total_Presumptive": "Total Presumptive",
            "Presumptive_Open": "Presumptive Open",
            "Presumptive_Closed": "Presumptive Closed",
            "Total_Tested": "Total Tested",
            "Total_Diagnosed": "Total Diagnosed",
        }
    )

    # 2. District Summary
    dist_summary = (
        final_facility_master.groupby("District")
        .agg(
            Total_HWCs=("HWC", "count"),
            Active_HWCs=("Screening Status", lambda s: (s == "Screening Started").sum()),
            Inactive_HWCs=("Screening Status", lambda s: (s == "Screening Not Started").sum()),
            Total_Screening=("Total_Screening", "sum"),
            Total_Symptomatic=("Total_Symptomatic", "sum"),
            Total_Asymptomatic=("Total_Asymptomatic", "sum"),
            AI_Presumptive=("AI_Presumptive", "sum"),
            NTEP_Presumptive=("NTEP_Presumptive", "sum"),
            NTEP_Only_Presumptive=("NTEP_Only_Presumptive", "sum"),
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
                "District": "District Name",
                "Total_HWCs": "Total HWCs",
                "Active_HWCs": "Active HWCs",
                "Inactive_HWCs": "Inactive HWCs",
                "Total_Screening": "Total Screening",
                "Total_Symptomatic": "Symptomatic",
                "Total_Asymptomatic": "Asymptomatic",
                "AI_Presumptive": "AI Presumptive",
                "NTEP_Presumptive": "NTEP Presumptive",
                "NTEP_Only_Presumptive": "NTEP-Only Presumptive",
                "Both_AI_NTEP": "AI + NTEP Both",
                "CHO_Override": "CHO Override",
                "Total_Presumptive": "Total Presumptive",
                "Presumptive_Open": "Presumptive Open",
                "Presumptive_Closed": "Presumptive Closed",
                "Total_Tested": "Total Tested",
                "Total_Diagnosed": "Total Diagnosed",
            }
        )
    )

    # 3. Block Summary
    block_summary = (
        final_facility_master.groupby(["District", "Block"])
        .agg(
            Total_HWCs=("HWC", "count"),
            Active_HWCs=("Screening Status", lambda s: (s == "Screening Started").sum()),
            Inactive_HWCs=("Screening Status", lambda s: (s == "Screening Not Started").sum()),
            Total_Screening=("Total_Screening", "sum"),
            Total_Symptomatic=("Total_Symptomatic", "sum"),
            Total_Asymptomatic=("Total_Asymptomatic", "sum"),
            AI_Presumptive=("AI_Presumptive", "sum"),
            NTEP_Presumptive=("NTEP_Presumptive", "sum"),
            NTEP_Only_Presumptive=("NTEP_Only_Presumptive", "sum"),
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
                "District": "District Name",
                "Block": "Block Name",
                "Total_HWCs": "Total HWCs",
                "Active_HWCs": "Active HWCs",
                "Inactive_HWCs": "Inactive HWCs",
                "Total_Screening": "Total Screening",
                "Total_Symptomatic": "Symptomatic",
                "Total_Asymptomatic": "Asymptomatic",
                "AI_Presumptive": "AI Presumptive",
                "NTEP_Presumptive": "NTEP Presumptive",
                "NTEP_Only_Presumptive": "NTEP-Only Presumptive",
                "Both_AI_NTEP": "AI + NTEP Both",
                "CHO_Override": "CHO Override",
                "Total_Presumptive": "Total Presumptive",
                "Presumptive_Open": "Presumptive Open",
                "Presumptive_Closed": "Presumptive Closed",
                "Total_Tested": "Total Tested",
                "Total_Diagnosed": "Total Diagnosed",
            }
        )
    )

    # --------------------------------------------------------------------------
    # NAVIGATION TABS
    # --------------------------------------------------------------------------
    tab_hwc, tab_dist, tab_block, tab_inactive, tab_symptom, tab_agreement, tab_cascade, tab_export = st.tabs(
        [
            "📋 HWC-Wise Performance Report",
            "🗺️ District-Wise Summary",
            "🏢 Block-Wise Summary",
            "⚠️ Inactive HWC Report",
            "🤒 Symptomatic vs Asymptomatic",
            "🤝 AI vs NTEP Agreement",
            "🔻 TB Cascade Continuum",
            "📥 Export Reports",
        ]
    )

    # TAB 1: HWC-WISE PERFORMANCE REPORT
    with tab_hwc:
        st.markdown(
            f"""
            <div class="section-title-wrap">
                <div class="section-title-text"><span>📋</span> Facility-Wise Performance & Master Line-List ({len(final_hwc_report):,} Centers)</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        top_c1, top_c2, top_c3 = st.columns([6, 2, 2])
        with top_c1:
            search_text = st.text_input("Quick Filter", "", placeholder="Type HWC / CHO / Block name to search...", label_visibility="collapsed")

        filtered_hwc_report = final_hwc_report
        if search_text.strip():
            m = (
                final_hwc_report["HWC"].str.contains(search_text.strip(), case=False, na=False)
                | final_hwc_report["CHO"].str.contains(search_text.strip(), case=False, na=False)
                | final_hwc_report["Block"].str.contains(search_text.strip(), case=False, na=False)
                | final_hwc_report["District"].str.contains(search_text.strip(), case=False, na=False)
            )
            filtered_hwc_report = final_hwc_report[m]

        with top_c2:
            excel_bytes = generate_styled_excel({"HWC_Performance": filtered_hwc_report})
            st.download_button(
                "📥 Download Excel (.xlsx)",
                data=excel_bytes,
                file_name="HWC_Performance_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with top_c3:
            pdf_bytes = generate_standard_pdf(
                filtered_hwc_report.head(150),
                "HWC-Wise Performance & Screening Report",
                f"Generated on {current_time_str} | Active CHO Source: Column H"
            )
            st.download_button(
                "📄 Download PDF (.pdf)",
                data=pdf_bytes,
                file_name="HWC_Performance_Report.pdf",
                mime="application/pdf" if REPORTLAB_AVAILABLE else "text/plain",
                use_container_width=True,
            )

        st.dataframe(
            filtered_hwc_report.style.format({
                "Total Screening": "{:,}", "Symptomatic": "{:,}", "Asymptomatic": "{:,}",
                "AI Presumptive": "{:,}", "NTEP Presumptive": "{:,}", "NTEP-Only Presumptive": "{:,}",
                "AI + NTEP Both": "{:,}", "CHO Override": "{:,}", "Total Presumptive": "{:,}",
                "Presumptive Open": "{:,}", "Presumptive Closed": "{:,}", "Total Tested": "{:,}",
                "Total Diagnosed": "{:,}"
            }),
            use_container_width=True,
            hide_index=True,
            height=530,
        )

    # TAB 2: DISTRICT-WISE SUMMARY
    with tab_dist:
        st.markdown('<div class="section-title-wrap"><div class="section-title-text"><span>🗺️</span> District-Level Surveillance Summary</div></div>', unsafe_allow_html=True)
        col_db1, col_db2 = st.columns([8, 2])
        with col_db2:
            dist_pdf = generate_standard_pdf(dist_summary, "District-Level Surveillance Summary", f"Jurisdiction Summary | Generated: {current_time_str}")
            st.download_button("📄 Download District PDF", data=dist_pdf, file_name="District_Summary_Report.pdf", mime="application/pdf" if REPORTLAB_AVAILABLE else "text/plain", use_container_width=True)

        fig_dist = px.bar(
            dist_summary.sort_values(by="Total Screening", ascending=False),
            x="District Name",
            y=["Total Screening", "Symptomatic", "Asymptomatic", "Total Presumptive", "Total Tested", "Total Diagnosed"],
            barmode="group",
            color_discrete_sequence=["#0284C7", "#D97706", "#0369A1", "#EF4444", "#7C3AED", "#059669"],
        )
        fig_dist.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1), height=350, margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig_dist, use_container_width=True)
        st.dataframe(dist_summary.style.format({c: "{:,}" for c in dist_summary.columns if c != "District Name"}), use_container_width=True, hide_index=True)

    # TAB 3: BLOCK-WISE SUMMARY
    with tab_block:
        st.markdown('<div class="section-title-wrap"><div class="section-title-text"><span>🏢</span> Block-Level Performance & Coverage</div></div>', unsafe_allow_html=True)
        col_bb1, col_bb2 = st.columns([8, 2])
        with col_bb2:
            block_pdf = generate_standard_pdf(block_summary, "Block-Level Performance Summary", f"Sub-District Coverage | Generated: {current_time_str}")
            st.download_button("📄 Download Block PDF", data=block_pdf, file_name="Block_Summary_Report.pdf", mime="application/pdf" if REPORTLAB_AVAILABLE else "text/plain", use_container_width=True)

        st.dataframe(block_summary.style.format({c: "{:,}" for c in block_summary.columns if c not in ["District Name", "Block Name"]}), use_container_width=True, hide_index=True)

    # TAB 4: INACTIVE HWC REPORT
    with tab_inactive:
        st.markdown('<div class="section-title-wrap"><div class="section-title-text"><span>⚠️</span> HWCs Where Screening Has Not Started (Master Facility Audit)</div></div>', unsafe_allow_html=True)
        
        if df_master_hwc is None:
            st.info("ℹ️ Inactive HWCs की पहचान करने के लिए कृपया साइडबार में **HWC Master Hierarchy (Column E)** फ़ाइल अपलोड करें।")
        else:
            col_ib1, col_ib2 = st.columns([8, 2])
            with col_ib2:
                inact_pdf = generate_standard_pdf(inactive_hwc_report, "Inactive HWCs (Screening Not Started)", f"Zero Screening Audit | Generated: {current_time_str}")
                st.download_button("📄 Download Inactive PDF", data=inact_pdf, file_name="Inactive_HWCs_Report.pdf", mime="application/pdf" if REPORTLAB_AVAILABLE else "text/plain", use_container_width=True)

            st.warning(f"📌 Master Hierarchy के अनुसार कुल **{len(inactive_hwc_report):,} HWCs** में अभी तक कोई स्क्रीनिंग नहीं हुई है। इन केन्द्रों पर CHO Status **Not Available** है।")
            st.dataframe(inactive_hwc_report, use_container_width=True, hide_index=True)

    # TAB 5: SYMPTOMATIC VS ASYMPTOMATIC
    with tab_symptom:
        st.markdown('<div class="section-title-wrap"><div class="section-title-text"><span>🤒</span> Symptomatic vs. Asymptomatic Presentation (Column AC)</div></div>', unsafe_allow_html=True)
        col_s1, col_s2 = st.columns([5, 4])
        with col_s1:
            fig_symp_bar = px.bar(
                dist_summary.sort_values(by="Total Screening", ascending=False),
                x="District Name",
                y=["Symptomatic", "Asymptomatic"],
                barmode="group",
                color_discrete_sequence=["#D97706", "#0369A1"],
            )
            fig_symp_bar.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1), height=340, margin=dict(l=10, r=10, t=30, b=10))
            st.plotly_chart(fig_symp_bar, use_container_width=True)

        with col_s2:
            symp_pie_df = pd.DataFrame({"Presentation Type": ["Symptomatic", "Asymptomatic"], "Count": [total_symptomatic, total_asymptomatic]})
            fig_symp_pie = px.pie(
                symp_pie_df,
                names="Presentation Type",
                values="Count",
                hole=0.55,
                color="Presentation Type",
                color_discrete_map={"Symptomatic": "#D97706", "Asymptomatic": "#0369A1"},
            )
            fig_symp_pie.update_layout(height=340, margin=dict(l=10, r=10, t=10, b=10), legend=dict(orientation="h", yanchor="bottom", y=-0.15))
            st.plotly_chart(fig_symp_pie, use_container_width=True)

    # TAB 6: AI VS NTEP AGREEMENT
    with tab_agreement:
        st.markdown('<div class="section-title-wrap"><div class="section-title-text"><span>🤝</span> AI vs NTEP Classification Concordance Matrix</div></div>', unsafe_allow_html=True)
        col_ag_l, col_ag_r = st.columns([5, 4])
        with col_ag_l:
            ai_only_pres = int(f_screen["is_ai_only"].sum()) if "is_ai_only" in f_screen.columns else 0
            both_non_pres = int(((f_screen["is_ai_pres"] == 0) & (f_screen["is_ntep_pres"] == 0)).sum()) if ("is_ai_pres" in f_screen.columns and "is_ntep_pres" in f_screen.columns) else 0
            contingency_df = pd.DataFrame({
                "Classification Modality": ["AI Presumptive", "AI Non-Presumptive"],
                "NTEP Presumptive": [total_both_pres, total_ntep_only_pres],
                "NTEP Non-Presumptive": [ai_only_pres, both_non_pres],
            })
            st.dataframe(contingency_df, use_container_width=True, hide_index=True)
            st.caption(
                f"• **Dual Positive Yield:** {total_both_pres:,} cases identified by both.\n"
                f"• **AI Incremental Capture:** {ai_only_pres:,} cases caught by AI where standard protocol was non-presumptive.\n"
                f"• **NTEP-Only Incremental Capture:** {total_ntep_only_pres:,} cases identified exclusively by NTEP.\n"
                f"• **CHO Clinical Override:** {total_cho_override:,} frontline field interventions."
            )

        with col_ag_r:
            comp_df = pd.DataFrame({
                "Component": ["AI + NTEP Both", "AI Only", "NTEP Only", "CHO Override"],
                "Cases": [total_both_pres, ai_only_pres, total_ntep_only_pres, total_cho_override],
            })
            fig_comp = px.pie(comp_df, names="Component", values="Cases", hole=0.55, color_discrete_sequence=["#16A34A", "#8B5CF6", "#EA580C", "#EC4899"])
            fig_comp.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10), legend=dict(orientation="h", yanchor="bottom", y=-0.2))
            st.plotly_chart(fig_comp, use_container_width=True)

    # TAB 7: TB STATUS CONTINUUM FUNNEL
    with tab_cascade:
        st.markdown('<div class="section-title-wrap"><div class="section-title-text"><span>🔻</span> Cascade of Care: Screening to Treatment</div></div>', unsafe_allow_html=True)
        col_f1, col_f2 = st.columns([5, 4])
        with col_f1:
            funnel_df = pd.DataFrame({
                "Cascade Stage": [
                    "1. Screened Population",
                    "2. Total Presumptive",
                    "3. Nikshay Created (Open)",
                    "4. Tested (Closed + Diagnosed)",
                    "5. Diagnosed on Treatment",
                ],
                "Beneficiaries": [total_screenings, total_presumptive, total_presumptive_open, total_tested, total_diagnosed],
            })
            fig_funnel = px.funnel(funnel_df, x="Beneficiaries", y="Cascade Stage", color_discrete_sequence=["#0284C7"])
            fig_funnel.update_layout(showlegend=False, height=330, margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(fig_funnel, use_container_width=True)

        with col_f2:
            modality_df = pd.DataFrame({
                "Modality": ["AI Presumptive", "NTEP Presumptive", "NTEP Only", "Dual Overlap", "CHO Override"],
                "Cases": [total_ai_pres, total_ntep_pres, total_ntep_only_pres, total_both_pres, total_cho_override],
            })
            fig_bar = px.bar(modality_df, x="Modality", y="Cases", color="Modality", text_auto=True, color_discrete_sequence=["#8B5CF6", "#D97706", "#EA580C", "#0284C7", "#EC4899"])
            fig_bar.update_layout(showlegend=False, height=330, margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(fig_bar, use_container_width=True)

    # TAB 8: CONSOLIDATED EXPORT DOSSIER
    with tab_export:
        st.markdown('<div class="section-title-wrap"><div class="section-title-text"><span>📥</span> Master Export Dossier (Multi-Tab Excel & Executive PDFs)</div></div>', unsafe_allow_html=True)
        
        c_exp1, c_exp2 = st.columns(2)
        with c_exp1:
            st.markdown('<div class="export-box">', unsafe_allow_html=True)
            st.markdown("##### 📑 Multi-Sheet Master Workbook (.xlsx)")
            st.caption("Consolidates HWC Performance, District Summary, Block Summary, and Inactive Line-List into styled worksheets.")
            
            master_excel = generate_styled_excel({
                "HWC_Performance": final_hwc_report,
                "District_Summary": dist_summary,
                "Block_Summary": block_summary,
                "Inactive_HWCs": inactive_hwc_report,
            })
            st.download_button(
                "📥 Download Consolidated Excel (.xlsx)",
                data=master_excel,
                file_name=f"CATB_Surveillance_Dossier_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            st.markdown('</div>', unsafe_allow_html=True)

        with c_exp2:
            st.markdown('<div class="export-box">', unsafe_allow_html=True)
            st.markdown("##### 📄 Executive District Dossier PDF (.pdf)")
            st.caption("Standard printable Landscape A4 report with running headers, footers, and official pagination.")
            
            master_pdf = generate_standard_pdf(
                dist_summary,
                "CATB Active Surveillance - District Master Dossier",
                f"State / Monitored Area Consolidated Performance Report | Date: {current_time_str}"
            )
            st.download_button(
                "📄 Download District Dossier (.pdf)",
                data=master_pdf,
                file_name=f"CATB_District_Executive_Report_{datetime.datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf" if REPORTLAB_AVAILABLE else "text/plain",
                use_container_width=True,
            )
            st.markdown('</div>', unsafe_allow_html=True)


if __name__ == "__main__":
    main()