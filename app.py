import streamlit as st
import pandas as pd
import numpy as np
import re
import os
from datetime import datetime, timedelta
from io import BytesIO
from docx import Document

# Optimize page viewports across mobile browsers and high-resolution desktop terminals
st.set_page_config(
    page_title="Global Enterprise Multi-Cleaner",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- GLOBAL HORTICULTURAL SYSTEM DATABASE ---
AGRI_MASTER_DB = {
    "soil_profiles": {
        "Mkuranga / Coast Region": "Coastal sand-loam variants, exceptional drainage parameters, low organic baseline, high humidity indexes.",
        "General / Standard Tropical": "Variable clay-loam variations, standard retention capacities."
    },
    "fertilizer_matrix": {
        "Basal Phase": "DAP (Diammonium Phosphate) - Engineered for immediate structural root architecture building.",
        "Vegetative Growth": "UREA / CAN (Calcium Ammonium Nitrate) - High-efficiency nitrogen release systems.",
        "Production Phase": "NPK 15:15:15 / 20:10:10 - Balanced core macronutrient distribution layers.",
        "Quality & Yield Brix Adjustment": "MOP (Muriate of Potash) - Essential for fruit weight optimization, density development, and natural sugar profiling."
    },
    "crop_blueprint": {
        "Tomatoes": {"target_yield_per_acre_tons": 25.0, "spacing": "60cm x 50cm", "density_per_acre": 13300, "base_price_tsh": 1200000},
        "Watermelon": {"target_yield_per_acre_tons": 30.0, "spacing": "150cm x 100cm", "density_per_acre": 2700, "base_price_tsh": 800000},
        "Okra (Bamia)": {"target_yield_per_acre_tons": 8.0, "spacing": "50cm x 30cm", "density_per_acre": 26600, "base_price_tsh": 1500000},
        "Mangoes": {"target_yield_per_acre_tons": 12.0, "spacing": "9m x 9m", "density_per_acre": 50, "base_price_tsh": 2500000},
        "Pineapples": {"target_yield_per_acre_tons": 35.0, "spacing": "90cm x 60cm x 30cm", "density_per_acre": 18000, "base_price_tsh": 900000},
        "Cashew": {"target_yield_per_acre_tons": 1.2, "spacing": "12m x 12m", "density_per_acre": 27, "base_price_tsh": 3000000}
    }
}

# --- ENGINE 1: DATA CLEANER & STRUCTURAL AUDITOR ---
def clean_spreadsheet(uploaded_file, ext):
    """Engine 1: Audits, detects separators automatically, splits columns, and standardizes values."""
    if ext == '.csv':
        # Read the first few lines to sniff out the separator (comma, semicolon, or tab)
        raw_bytes = uploaded_file.read(2048)
        uploaded_file.seek(0) # Reset file pointer
        sample_text = raw_bytes.decode('utf-8', errors='ignore')
        
        sep = ','
        if ';' in sample_text and sample_text.count(';') > sample_text.count(','):
            sep = ';'
        elif '\t' in sample_text:
            sep = '\t'
            
        df = pd.read_csv(uploaded_file, sep=sep)
    else:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
    
    # 1. Enforce database-safe column names (lowercase, underscores, no symbols)
    df.columns = [re.sub(r'[^a-zA-Z0-9_]', '_', str(col).strip().lower()) for col in df.columns]
    # Strip excess underscores from headers caused by trailing formatting spaces
    df.columns = [re.sub(r'_+', '_', col).strip('_') for col in df.columns]
    
    # 2. Complete data normalization cycle row-by-row
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].str.replace(r'\s+', ' ', regex=True)
            df[col] = df[col].replace(['nan', 'NaN', 'None', 'NULL', 'null', ''], np.nan)
        
        # Smart formatting: Clean up names
        if 'name' in col:
            df[col] = df[col].apply(lambda x: str(x).strip().title() if pd.notna(x) else x)
            
        # Smart formatting: Standardize phone entries to international format with '+' prefix
        elif any(keyword in col for keyword in ['phone', 'contact', 'tel', 'mobile', 'num']):
            def _phone_fix(v):
                if pd.isna(v) or str(v).strip().lower() in ['nan', 'none', '-', 'missing', 'invalid', '']: return ""
                s = re.sub(r'[^0-9]', '', str(v)) 
                if s.startswith('0') and len(s) == 10: 
                    return '+255' + s[1:]
                if s.startswith('255') and len(s) == 12:
                    return '+' + s
                if len(s) > 0 and not s.startswith('+'):
                    return '+' + s
                return s
            df[col] = df[col].apply(_phone_fix)
            
        # FINANCIAL MODULE: Extracts numbers and converts to standard accounting format (e.g., 1,200,000.00)
        elif any(keyword in col for keyword in ['sales', 'amount', 'price', 'revenue', 'cost', 'yield', 'finance', 'total']):
            def _currency_formatter(v):
                if pd.isna(v) or str(v).strip() in ['-', '']: return "0.00"
                s = re.sub(r'[^0-9.]', '', str(v).lower())
                if s == '' or s == '.': return "0.00"
                num_val = float(s) if '.' in s else int(s)
                return f"{num_val:,.2f}"
            df[col] = df[col].apply(_currency_formatter)
            
        # FIXED DATE MODULE: Intelligently converts words like 'yesterday' to actual calendar stamps
        elif any(keyword in col for keyword in ['date', 'trans']):
            def _date_fix(v):
                if pd.isna(v) or str(v).strip() == '': return "Invalid Date"
                s = str(v).strip().lower()
                current_time = datetime.now()
                
                if 'yesterday' in s:
                    return (current_time - timedelta(days=1)).strftime('%Y-%m-%d')
                if 'today' in s or 'now' in s:
                    return current_time.strftime('%Y-%m-%d')
                
                parsed_date = pd.to_datetime(v, errors='coerce')
                if pd.notna(parsed_date):
                    return parsed_date.strftime('%Y-%m-%d')
                return "Invalid Date"
                
            df[col] = df[col].apply(_date_fix)

    # 3. Clear identical duplicates safely based on standardized keys
    df.dropna(how='all', inplace=True)
    identity_keys = [c for c in df.columns if 'name' in c or 'phone' in c or 'contact' in c]
    df.drop_duplicates(subset=identity_keys if identity_keys else None, keep='first', inplace=True)
    
    return df

# --- ENGINE 2: DOCUMENT PROCESSING & STYLE ADAPTER ---
def parse_and_reformat_document(uploaded_file, selected_style):
    doc = Document(uploaded_file)
    cleaned_doc = Document()
    cleaned_doc.add_heading(f"REFORMATTED BUSINESS ARTIFACT - STYLE: {selected_style.upper()}", level=1)
    
    for para in doc.paragraphs:
        txt = para.text.strip()
        if not txt: continue
        txt = re.sub(r'\s+', ' ', txt)
        
        if selected_style in ["Business", "Formal"]:
            txt = re.sub(r"\bi'm\b", "I am", txt, flags=re.I)
            txt = re.sub(r"\bcan't\b", "cannot", txt, flags=re.I)
            txt = re.sub(r"\bdon't\b", "do not", txt, flags=re.I)
            txt = re.sub(r"\bhey\b|\bhi\b", "Dear Sir/Madam,", txt, flags=re.I)
        elif selected_style == "Non-Formal":
            txt = re.sub(r"\butilize\b", "use", txt, flags=re.I)
            txt = re.sub(r"\bsubsequent to\b", "after", txt, flags=re.I)
            
        cleaned_doc.add_paragraph(txt)
        
    out = BytesIO()
    cleaned_doc.save(out)
    out.seek(0)
    return out

# --- ENGINE 3: AGRICULTURAL MODELER & CALCULATOR ---
def process_agricultural_matrix(uploaded_file, target_acres, location_profile):
    raw_text = uploaded_file.read().decode("utf-8", errors="ignore")
    
    report = [
        "==========================================================================",
        "          ENTERPRISE HORTICULTURE PLAN & PRECISION MANUAL GENERATOR       ",
        "==========================================================================",
        f"Target Operating Footprint Size: {target_acres} Acre(s)",
        f"Selected Regional Profile: {location_profile}",
        f"Soil Composition Analysis: {AGRI_MASTER_DB['soil_profiles'].get(location_profile, 'Standard Base')}\n",
        "--------------------------------------------------------------------------",
        "1. PRECISION MANAGEMENT MANUAL & NUTRIENT APPLICATION CYCLES",
        "--------------------------------------------------------------------------"
    ]
    for phase, management_plan in AGRI_MASTER_DB['fertilizer_matrix'].items():
        report.append(f"  ⚡ {phase} Matrix -> Use: {management_plan}")
        
    report.append("\n--------------------------------------------------------------------------")
    report.append("2. FINANCIAL FORECAST & PREDICTIVE YIELD MATRIX MODEL")
    report.append("--------------------------------------------------------------------------")
    
    detected_any = False
    for crop, data in AGRI_MASTER_DB['crop_blueprint'].items():
        if crop.lower() in raw_text.lower() or "all" in raw_text.lower() or len(raw_text) < 10:
            detected_any = True
            total_plant_population = int(data['density_per_acre'] * target_acres)
            total_yield_tons = round(data['target_yield_per_acre_tons'] * target_acres, 2)
            projected_gross_revenue = float(total_yield_tons * 1000 * (data['base_price_tsh'] / 1000))
            
            report.append(f"\n● CROP SYSTEM: {crop.upper()}")
            report.append(f"  ▪ Recommended Plant Population Sizing: {total_plant_population:,} plants")
            report.append(f"  ▪ Regional Spacing Configurations: {data['spacing']}")
            report.append(f"  ▪ Expected Operational Harvest Output: {total_yield_tons:,} Tons")
            report.append(f"  ▪ Projected Market Value Index Baseline: TSh {projected_gross_revenue:,.2f}")
            
    if not detected_any:
        report.append("\n*Note: No custom crop targets matched your instructions file. Baseline metrics provided.*")
        
    return "\n".join(report)


# --- INTERACTIVE USER INTERFACE CONSOLE ---
st.title("🧹 Universal Master Data Cleaning Hub")
st.write("Upload any file type below. The unified script automatically smells data delimiters, parses columns, reformats documents, and builds field production blueprints.")

# App Configuration Settings Sidebar
st.sidebar.header("⚙️ System Control Panel")
doc_style = st.sidebar.selectbox("Document Re-Styling Mode", ["Business", "Formal", "Non-Formal"])
agri_scale = st.sidebar.number_input("Target Agricultural Scale (Acres)", min_value=0.5, max_value=500.0, value=1.0, step=0.5)
agri_loc = st.sidebar.selectbox("Target Regional Zone", ["Mkuranga / Coast Region", "General / Standard Tropical"])

# Unified File Upload Interface Element
uploaded_file = st.file_uploader("Upload target ledger, contract, presentation text, or agricultural directive", type=["xlsx", "xls", "csv", "docx", "txt"])

if uploaded_file is not None:
    filename = uploaded_file.name
    _, ext = os.path.splitext(filename.lower())
    st.info(f"📁 System Core verified file extension properties: **{ext.upper()}**")
    
    # ROUTE 1: SPREADSHEETS & DATA LEDGERS
    if ext in ['.xlsx', '.xls', '.csv']:
        try:
            with st.spinner("Executing delimiter auto-sniffing and data cleaning algorithms..."):
                cleaned_df = clean_spreadsheet(uploaded_file, ext)
            st.subheader("👀 Preview Cleaned Grid")
            st.dataframe(cleaned_df.head(50), use_container_width=True)
            
            out_buf = BytesIO()
            if ext == '.csv':
                cleaned_df.to_csv(out_buf, index=False)
                m_type, name_out = "text/csv", "cleaned_master_spreadsheet.csv"
            else:
                cleaned_df.to_excel(out_buf, index=False, engine='openpyxl')
                m_type, name_out = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "cleaned_master_spreadsheet.xlsx"
            out_buf.seek(0)
            
            st.download_button("📥 Download Cleaned Spreadsheet", data=out_buf, file_name=name_out, mime=m_type, use_container_width=True)
        except Exception as e:
            st.error(f"Spreadsheet Clean Sub-system Fault: {str(e)}")
            
    # ROUTE 2: DOCUMENTS, MANUALS & CORPORATE TEXTS
    elif ext == '.docx':
        try:
            with st.spinner("Normalizing text formatting layouts..."):
                doc_stream = parse_and_reformat_document(uploaded_file, doc_style)
            st.subheader("👀 Preview Status")
            st.success(f"Document content parsed successfully. Spacing layouts corrected, and language set to **{doc_style.upper()}** parameters.")
            st.download_button(f"📥 Download Formatted {doc_style} Document", data=doc_stream, file_name="cleaned_master_document.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
        except Exception as e:
            st.error(f"Text Processing Engine Fault: {str(e)}")
            
    # ROUTE 3: AGRICULTURAL MANAGEMENT DIRECTIVES
    elif ext == '.txt':
        try:
            with st.spinner("Processing yield models against regional agronomy charts..."):
                agri_output_report = process_agricultural_matrix(uploaded_file, agri_scale, agri_loc)
            st.subheader("👀 Preview Blueprint")
            st.text_area("Generated Output File Data Display", value=agri_output_report, height=400)
            st.download_button("📥 Download Agri Implementation Plan (.txt)", data=agri_output_report, file_name="agri_precision_production_manual.txt", mime="text/plain", use_container_width=True)
        except Exception as e:
            st.error(f"Agricultural Modeling Engine Fault: {str(e)}")
