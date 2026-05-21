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
    },
    "protection_schedules": {
        "Tomatoes": [
            {"phase": "Nursery / Transplanting", "target": "Damping Off & Early Aphids", "intervention": "Copper Oxychloride + Imidacloprid", "rate": "2g/L + 0.5ml/L", "phi": "N/A"},
            {"phase": "Early Vegetative (Wk 1-3)", "target": "Tuta Absoluta & Leaf Miners", "intervention": "Flubendiamide or Spinosad", "rate": "0.3ml/L", "phi": "3 Days"},
            {"phase": "Flowering to Fruit Set", "target": "Early/Late Blight & Whiteflies", "intervention": "Mancozeb + Acetamiprid", "rate": "2.5g/L + 0.5g/L", "phi": "7 Days"},
            {"phase": "Maturation / Harvest", "target": "Fruit Borers & Powdery Mildew", "intervention": "Indoxacarb + Azoxystrobin", "rate": "0.5ml/L + 1ml/L", "phi": "3 Days"}
        ],
        "Watermelon": [
            {"phase": "Seedling Emergence", "target": "Soil Insects & Damping Off", "intervention": "Metalaxyl drenching", "rate": "2g/L", "phi": "N/A"},
            {"phase": "Vining / Vegetative", "target": "Melon Aphids & Thrips", "intervention": "Thiamethoxam", "rate": "0.4g/L", "phi": "7 Days"},
            {"phase": "Flowering Block", "target": "Downy Mildew (Avoid pollinator disruption)", "intervention": "Propamocarb (Apply late afternoon)", "rate": "1.5ml/L", "phi": "3 Days"},
            {"phase": "Fruit Expansion", "target": "Fruit Flies & Anthracnose", "intervention": "Lambda-Cyhalothrin + Mancozeb", "rate": "0.5ml/L + 2g/L", "phi": "7 Days"}
        ],
        "Okra (Bamia)": [
            {"phase": "Early Establishment", "target": "Flea Beetles & Jassids", "intervention": "Imidacloprid foliar application", "rate": "0.5ml/L", "phi": "7 Days"},
            {"phase": "Vegetative Stretch", "target": "Powdery Mildew & Aphids", "intervention": "Sulfur WG + Acetamiprid", "rate": "3g/L + 0.4g/L", "phi": "3 Days"},
            {"phase": "Flowering Phase", "target": "Bollworms / Pod Borers", "intervention": "Chlorantraniliprole (Coragen)", "rate": "0.4ml/L", "phi": "1 Day"},
            {"phase": "Active Harvest Loop", "target": "Whiteflies & Red Spider Mites", "intervention": "Abamectin (Strict PHI safety sweep)", "rate": "0.5ml/L", "phi": "3 Days"}
        ]
    }
}

# --- ENGINE 1: DATA CLEANER & STRUCTURAL AUDITOR ---
def clean_spreadsheet(uploaded_file, ext):
    if ext == '.csv':
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
    
    if len(df.columns) == 1:
        raw_col = df.columns[0]
        header_line = str(raw_col).replace('"', '').strip()
        new_headers = [re.sub(r'[^a-zA-Z0-9_]', '_', h.strip().lower()) for h in header_line.split(',')]
        
        split_rows = []
        for val in df.iloc[:, 0]:
            row_str = str(val).strip().strip('"')
            row_cells = re.split(r',(?=(?:[^"]*"[^"]*")*[^"]*$)', row_str)
            split_rows.append([c.strip('"').strip() for c in row_cells])
            
        df = pd.DataFrame(split_rows, columns=new_headers[:len(split_rows[0])])
    else:
        df.columns = [re.sub(r'[^a-zA-Z0-9_]', '_', str(col).strip().lower()) for col in df.columns]
    
    df.columns = [re.sub(r'_+', '_', col).strip('_') for col in df.columns]
    
    phone_cols = [c for c in df.columns if any(k in c for k in ['phone', 'contact', 'tel', 'mobile', 'num'])]
    for col in phone_cols:
        def _pre_parse_phone(v):
            if pd.isna(v) or str(v).strip().lower() in ['nan', 'none', '-', 'missing', 'invalid', '']: return ""
            s = re.sub(r'[^0-9]', '', str(v))
            if s.startswith('0') and len(s) == 10: return '+255' + s[1:]
            if s.startswith('255') and len(s) == 12: return '+' + s
            if len(s) > 0 and not s.startswith('+'): return '+' + s
            return s
        df[col] = df[col].apply(_pre_parse_phone)

    for col in df.columns:
        if df[col].dtype == 'object' and col not in phone_cols:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].str.replace(r'\s+', ' ', regex=True)
            df[col] = df[col].replace(['nan', 'NaN', 'None', 'NULL', 'null', ''], np.nan)
        
        if 'name' in col:
            def _clean_name(v):
                if pd.isna(v) or str(v).strip() == '': return ""
                s = str(v).strip().title()
                s = re.sub(r'\b(Dr|Mr|Mrs|Ms|Eng|Prof)\.?\s+', '', s, flags=re.I)
                return s.strip()
            df[col] = df[col].apply(_clean_name)
            
        elif any(keyword in col for keyword in ['sales', 'amount', 'price', 'revenue', 'cost', 'yield', 'finance', 'total']):
            def _currency_formatter(v):
                if pd.isna(v) or str(v).strip() in ['-', '']: return "0.00"
                s = re.sub(r'[^0-9.]', '', str(v).lower())
                if s == '' or s == '.': return "0.00"
                num_val = float(s) if '.' in s else int(s)
                return f"{num_val:,.2f}"
            df[col] = df[col].apply(_currency_formatter)
            
        elif any(keyword in col for keyword in ['date', 'trans']):
            def _date_fix(v):
                if pd.isna(v) or str(v).strip() == '': return "Invalid Date"
                s = str(v).strip().lower()
                current_time = datetime(2026, 5, 21)
                
                if 'yesterday' in s: return (current_time - timedelta(days=1)).strftime('%Y-%m-%d')
                if 'today' in s or 'now' in s: return current_time.strftime('%Y-%m-%d')
                
                if '/' in s:
                    parts = s.split('/')
                    if len(parts) == 3:
                        p1, p2, p3 = int(parts[0]), int(parts[1]), int(parts[2])
                        year = 2000 + p3 if p3 < 100 else p3
                        if p1 > 12 and p2 <= 12: return f"{year}-{p2:02d}-{p1:02d}"
                        elif p1 <= 12 and p2 <= 12: return f"{year}-{p2:02d}-{p1:02d}"
                
                parsed_date = pd.to_datetime(v, errors='coerce')
                if pd.notna(parsed_date): return parsed_date.strftime('%Y-%m-%d')
                return "Invalid Date"
                
            df[col] = df[col].apply(_date_fix)

    df.dropna(how='all', inplace=True)
    
    if phone_cols:
        valid_phone_col = phone_cols[0]
        df = df.loc[~(df[valid_phone_col].duplicated(keep='first') & (df[valid_phone_col] != ""))]
        
    # 🔥 STRATEGIC PURGE LAYER: Filter out completely corrupted dates before visualization
    date_cols = [c for c in df.columns if any(k in c for k in ['date', 'trans'])]
    if date_cols:
        df = df[df[date_cols[0]] != 'Invalid Date']
        
    return df

# --- ENGINE 2: DOCUMENT PROCESSING & STYLE ADAPTER ---
def parse_and_reformat_document(uploaded_file, selected_style):
    doc = Document(uploaded_file)
    cleaned_doc = Document()
    raw_paras = []
    for p in doc.paragraphs:
        t = p.text.strip()
        if t: raw_paras.append(re.sub(r'\s+', ' ', t))
            
    full_text_block = "\n".join(raw_paras)
    is_letter = "to:" in full_text_block.lower() or "dear" in full_text_block.lower()
    
    if is_letter and selected_style in ["Business", "Formal"]:
        cleaned_doc.add_paragraph("[SENDER CONTACT DETAILS]\n[Postal Address Line 1]\nDar es Salaam, Tanzania\n")
        cleaned_doc.add_paragraph(f"Date: May 21, 2026\n")
    else:
        cleaned_doc.add_heading(f"REFORMATTED BUSINESS ARTIFACT - STYLE: {selected_style.upper()}", level=1)
        
    greeting_injected = False
    
    for txt in raw_paras:
        if selected_style in ["Business", "Formal"]:
            if txt.lower().startswith("date:"): continue
            if txt.lower().startswith("to:"):
                txt = re.sub(r"\bto:\s*", "TO:\n", txt, flags=re.I)
                txt = txt.title().replace("Nssf", "NSSF")
                cleaned_doc.add_paragraph(txt)
                continue
                
            contains_greeting = any(k in txt.lower() for k in ["hey there", "hi team", "hey", "hi", "dear sir"])
            if contains_greeting:
                if not greeting_injected:
                    cleaned_doc.add_paragraph("Dear Sir/Madam,")
                    greeting_injected = True
                txt = re.sub(r"\bhey\s+there,?\s*|\bhi\s+team,?\s*|\bhey,?\s*|\bhi,?\s*|\bdear\s+sir/madam,?\s*", "", txt, flags=re.I)
                if not txt.strip(): continue

            txt = re.sub(r"\bi'm\b", "I am", txt, flags=re.I)
            txt = re.sub(r"\bcan't\b", "cannot", txt, flags=re.I)
            txt = re.sub(r"\bdon't\b", "do not", txt, flags=re.I)
            txt = re.sub(r"\basap\b", "as soon as possible", txt, flags=re.I)
            txt = re.sub(r"\bhaven't\b", "have not", txt, flags=re.I)
            txt = re.sub(r"\bask about\b", "inquire regarding", txt, flags=re.I)
            txt = re.sub(r"\bi've\b", "I have", txt, flags=re.I)
            txt = re.sub(r"\bi\b", "I", txt)
            txt = re.sub(r"\b(i\s)", "I ", txt)
            
            sentences = txt.split('.')
            processed_sentences = []
            for s in sentences:
                s_strip = s.strip()
                if len(s_strip) > 0: processed_sentences.append(s_strip[0].upper() + s_strip[1:])
            txt = ". ".join(processed_sentences)
            if len(txt) > 0 and not txt.endswith('.'): txt += '.'
            txt = re.sub(r',+', ',', txt)
            txt = txt.replace("Nssf", "NSSF")
            
        elif selected_style == "Non-Formal":
            txt = re.sub(r"\butilize\b", "use", txt, flags=re.I)
            txt = re.sub(r"\bsubsequent to\b", "after", txt, flags=re.I)
            
        cleaned_doc.add_paragraph(txt)
        
    if is_letter and selected_style in ["Business", "Formal"]:
        cleaned_doc.add_paragraph("\nYours faithfully,\n\n\n_______________________\n[Insert Full Account Name]\nClaimant / Account Holder")
        
    out = BytesIO()
    cleaned_doc.save(out)
    out.seek(0)
    return out

# --- ENGINE 3: AGRICULTURAL MODELER & CALCULATOR ---
def process_agricultural_matrix(uploaded_file, ext, target_acres, location_profile):
    if ext == '.docx':
        doc = Document(uploaded_file)
        raw_text = "\n".join([p.text for p in doc.paragraphs])
    else:
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
        
    detected_any = False
    crop_blocks = []
    
    for crop, data in AGRI_MASTER_DB['crop_blueprint'].items():
        crop_keyword = crop.lower().split(' ')[0]
        if crop_keyword in raw_text.lower() or "all" in raw_text.lower() or len(raw_text.strip()) < 10:
            detected_any = True
            total_plant_population = int(data['density_per_acre'] * target_acres)
            total_yield_tons = round(data['target_yield_per_acre_tons'] * target_acres, 2)
            projected_gross_revenue = float(total_yield_tons * 1000 * (data['base_price_tsh'] / 1000))
            
            cb = [
                f"\n● CROP SYSTEM: {crop.upper()}",
                f"  ▪ Recommended Plant Population Sizing: {total_plant_population:,} plants",
                f"  ▪ Regional Spacing Configurations: {data['spacing']}",
                f"  ▪ Expected Operational Harvest Output: {total_yield_tons:,} Tons",
                f"  ▪ Projected Market Value Index Baseline: TSh {projected_gross_revenue:,.2f}"
            ]
            if crop in AGRI_MASTER_DB["protection_schedules"]:
                cb.append("\n  ⚙️ ENTERPRISE PLANT PROTECTION PROTOCOLS (SCOUTING & INTERVENTION):")
                for schedule in AGRI_MASTER_DB["protection_schedules"][crop]:
                    cb.append(f"    ▪ [{schedule['phase']}] Target: {schedule['target']}")
                    cb.append(f"      Intervention Strategy: {schedule['intervention']} | Field Rate: {schedule['rate']} | PHI Safety Window: {schedule['phi']}")
            crop_blocks.append("\n".join(cb))
            
    report.append("\n--------------------------------------------------------------------------")
    report.append("2. FINANCIAL FORECAST & PREDICTIVE YIELD MATRIX MODEL")
    report.append("--------------------------------------------------------------------------")
    if detected_any: report.append("\n".join(crop_blocks))
    else: report.append("\n*Note: No custom crop targets matched your instructions file. Baseline metrics provided.*")
    return "\n".join(report)

# --- INTERACTIVE USER INTERFACE CONSOLE ---
st.title("🧹 Universal Master Data Cleaning Hub")
st.write("Upload any file type below. The unified script automatically smells data delimiters, parses columns, reformats documents, and builds field production blueprints.")

st.sidebar.header("⚙️ System Control Panel")
doc_style = st.sidebar.selectbox("Document Re-Styling Mode", ["Business", "Formal", "Non-Formal"])
agri_scale = st.sidebar.number_input("Target Agricultural Scale (Acres)", min_value=0.5, max_value=500.0, value=1.0, step=0.5)
agri_loc = st.sidebar.selectbox("Target Regional Zone", ["Mkuranga / Coast Region", "General / Standard Tropical"])

uploaded_file = st.file_uploader("Upload target ledger, contract, presentation text, or agricultural directive", type=["xlsx", "xls", "csv", "docx", "txt"])

if uploaded_file is not None:
    filename = uploaded_file.name
    _, ext = os.path.splitext(filename.lower())
    st.info(f"📁 System Core verified file extension properties: **{ext.upper()}**")
    
    is_agri_doc = False
    if ext == '.docx':
        try:
            check_doc = Document(uploaded_file)
            uploaded_file.seek(0)
            full_text = "\n".join([p.text for p in check_doc.paragraphs]).lower()
            if any(k in full_text for k in ["directive", "okra", "harvest", "field blueprint", "crop"]): is_agri_doc = True
        except: pass

    if ext in ['.xlsx', '.xls', '.csv']:
        try:
            with st.spinner("Executing structural extraction algorithms..."):
                cleaned_df = clean_spreadsheet(uploaded_file, ext)
            st.subheader("👀 Preview Cleaned Grid")
            st.dataframe(cleaned_df, use_container_width=True)
            
            out_buf = BytesIO()
            if ext == '.csv':
                cleaned_df.to_csv(out_buf, index=False)
                m_type, name_out = "text/csv", "cleaned_master_spreadsheet.csv"
            else:
                cleaned_df.to_excel(out_buf, index=False, engine='openpyxl')
                m_type, name_out = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "cleaned_master_spreadsheet.xlsx"
            out_buf.seek(0)
            st.download_button("📥 Download Cleaned Spreadsheet", data=out_buf, file_name=name_out, mime=m_type, use_container_width=True)
            
            # 🔥 SECTION 4: INTEGRATED EXECUTIVE BUSINESS INTELLIGENCE DASHBOARD
            st.markdown("---")
            st.subheader("📊 Executive Data Insights Dashboard")
            
            # Find sales and date column keys dynamically
            sales_cols = [c for c in cleaned_df.columns if any(k in c for k in ['sales', 'amount', 'price', 'revenue', 'cost', 'total'])]
            date_cols = [c for c in cleaned_df.columns if any(k in c for k in ['date', 'trans'])]
            zone_cols = [c for c in cleaned_df.columns if 'zone' in c or 'region' in c]
            
            if sales_cols:
                metric_df = cleaned_df.copy()
                # Cast the sales text string column back to an executable floating number decimal matrix
                metric_df[sales_cols[0]] = metric_df[sales_cols[0]].astype(str).str.replace(',', '').astype(float)
                
                col1, col2 = st.columns(2)
                
                with col1:
                    total_vol = metric_df[sales_cols[0]].sum()
                    st.metric(label="Validated Transaction Volume", value=f"TSh {total_vol:,.2f}")
                    
                with col2:
                    total_records = len(metric_df)
                    st.metric(label="Total Cleaned Safe Records", value=f"{total_records} Active Rows")
                
                # Render interactive visual metrics charts if both dimensions are present
                if date_cols:
                    st.write("📈 **Transaction Volume Vector Over Time**")
                    metric_df[date_cols[0]] = pd.to_datetime(metric_df[date_cols[0]])
                    time_trend = metric_df.groupby(date_cols[0])[sales_cols[0]].sum().reset_index()
                    st.line_chart(data=time_trend, x=date_cols[0], y=sales_cols[0])
                    
                if zone_cols:
                    st.write("🌍 **Regional Revenue Breakdown Matrix**")
                    zone_chart = metric_df.groupby(zone_cols[0])[sales_cols[0]].sum().reset_index()
                    st.bar_chart(data=zone_chart, x=zone_cols[0], y=sales_cols[0])
                    
        except Exception as e: st.error(f"Spreadsheet Clean Sub-system Fault: {str(e)}")
            
    elif ext == '.txt' or is_agri_doc:
        try:
            with st.spinner("Processing yield models with agronomy protection matrices..."):
                agri_output_report = process_agricultural_matrix(uploaded_file, ext, agri_scale, agri_loc)
            st.subheader("👀 Preview Blueprint")
            st.text_area("Generated Output File Data Display", value=agri_output_report, height=500)
            st.download_button("📥 Download Agri Implementation Plan (.txt)", data=agri_output_report, file_name="agri_precision_production_manual.txt", mime="text/plain", use_container_width=True)
        except Exception as e: st.error(f"Agricultural Modeling Engine Fault: {str(e)}")

    elif ext == '.docx':
        try:
            with st.spinner("Normalizing text formatting layouts..."):
                doc_stream = parse_and_reformat_document(uploaded_file, doc_style)
            st.subheader("👀 Preview Status")
            st.success(f"Document content parsed successfully. Spacing layouts corrected, and language set to **{doc_style.upper()}** parameters.")
            st.download_button(f"📥 Download Formatted {doc_style} Document", data=doc_stream, file_name="cleaned_master_document.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
        except Exception as e: st.error(f"Text Processing Engine Fault: {str(e)}")
