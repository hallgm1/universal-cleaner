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
        "Onions (Kitunguu)": {"target_yield_per_acre_tons": 16.0, "spacing": "30cm x 10cm", "density_per_acre": 133000, "base_price_tsh": 1800000},
        "Tomatoes": {"target_yield_per_acre_tons": 25.0, "spacing": "60cm x 50cm", "density_per_acre": 13300, "base_price_tsh": 1200000},
        "Watermelon": {"target_yield_per_acre_tons": 30.0, "spacing": "150cm x 100cm", "density_per_acre": 2700, "base_price_tsh": 800000},
        "Okra (Bamia)": {"target_yield_per_acre_tons": 8.0, "spacing": "50cm x 30cm", "density_per_acre": 26600, "base_price_tsh": 1500000},
        "Mangoes": {"target_yield_per_acre_tons": 12.0, "spacing": "9m x 9m", "density_per_acre": 50, "base_price_tsh": 2500000},
        "Pineapples": {"target_yield_per_acre_tons": 35.0, "spacing": "90cm x 60cm x 30cm", "density_per_acre": 18000, "base_price_tsh": 900000},
        "Cashew": {"target_yield_per_acre_tons": 1.2, "spacing": "12m x 12m", "density_per_acre": 27, "base_price_tsh": 3000000}
    },
    "protection_schedules": {
        "Onions (Kitunguu)": [
            {"phase": "Nursery Phase (Weeks 1-6)", "target": "Damping Off & Nursery Thrips", "intervention": "Metalaxyl + Profenofos drenching loop", "rate": "2g/L + 1ml/L", "phi": "N/A"},
            {"phase": "Transplanting Window (Weeks 6-7)", "target": "Seedling Shock & Root Nematodes", "intervention": "Humic acid root dipping + Bio-nematicides", "rate": "5ml/L", "phi": "N/A"},
            {"phase": "Early Growth & Vining (Weeks 8-12)", "target": "Onion Thrips & Purple Blotch", "intervention": "Lambda-Cyhalothrin + Mancozeb protective spray", "rate": "0.5ml/L + 2.5g/L", "phi": "14 Days"},
            {"phase": "Bulb Expansion Phase (Weeks 13-19)", "target": "Downy Mildew & Storage Rot Risk", "intervention": "Copper Oxychloride + Acetamiprid systematic sweep", "rate": "2g/L + 0.5g/L", "phi": "7 Days"},
            {"phase": "Maturity & Solar Curing (Weeks 20-24+)", "target": "Neck Rot & Post-Harvest Degradation", "intervention": "Stop all irrigation loops completely. Windrow drying field curing.", "rate": "Manual Processing", "phi": "Zero Chemical Pass"}
        ],
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
            
        elif any(keyword in col for keyword in ['sales', 'amount', 'price', 'revenue', 'cost', 'yield', 'finance', 'total', 'salary']):
            def _currency_formatter(v):
                if pd.isna(v) or str(v).strip() in ['-', '']: return "0.00"
                s = re.sub(r'[^0-9.]', '', str(v).lower())
                if s == '' or s == '.': return "0.00"
                try:
                    num_val = float(s)
                    return f"{num_val:,.2f}"
                except ValueError:
                    return "0.00"
            df[col] = df[col].apply(_currency_formatter)
            
        elif any(keyword in col for keyword in ['date', 'trans', 'hire', 'exit']):
            def _date_fix(v):
                if pd.isna(v) or str(v).strip() == '' or str(v).lower() in ['nan', 'none', '-']: return "Invalid Date"
                s = str(v).strip().lower().replace('.', '-')
                current_time = datetime(2026, 5, 22)
                
                if 'yesterday' in s: return (current_time - timedelta(days=1)).strftime('%Y-%m-%d')
                if 'today' in s or 'now' in s: return current_time.strftime('%Y-%m-%d')
                
                separator = '/' if '/' in s else '-'
                if separator in s:
                    parts = s.split(separator)
                    if len(parts) == 3:
                        try:
                            if len(parts[0].strip()) == 4:
                                y = int(parts[0].strip())
                                m = int(parts[1].strip())
                                d = int(parts[2].strip())
                                if m <= 12 and d <= 31: return f"{y}-{m:02d}-{d:02d}"
                                if d <= 12 and m <= 31: return f"{y}-{d:02d}-{m:02d}" # Flipped safety catch
                            
                            p1 = int(parts[0].strip())
                            p2 = int(parts[1].strip())
                            p3 = int(parts[2].strip())
                            y = 2000 + p3 if p3 < 100 else p3
                            
                            if p1 <= 31 and p2 <= 12: return f"{y}-{p2:02d}-{p1:02d}"
                            if p2 <= 31 and p1 <= 12: return f"{y}-{p1:02d}-{p2:02d}"
                        except ValueError:
                            pass
                
                parsed_date = pd.to_datetime(v, errors='coerce')
                if pd.notna(parsed_date): return parsed_date.strftime('%Y-%m-%d')
                return "Invalid Date"
                
            df[col] = df[col].apply(_date_fix)
            
        elif 'zone' in col or 'region' in col:
            def _normalize_region(v):
                if pd.isna(v) or str(v).strip() == '': return "Unknown"
                s = str(v).strip().lower()
                if s in ['dar', 'dsm', 'dar es salaam', 'tanzania']: return "Dar es Salaam"
                if s in ['znz', 'zanzibar']: return "Zanzibar"
                return str(v).strip().title()
            df[col] = df[col].apply(_normalize_region)

    df.dropna(how='all', inplace=True)
    
    if phone_cols:
        valid_phone_col = phone_cols[0]
        df = df.loc[~(df[valid_phone_col].duplicated(keep='first') & (df[valid_phone_col] != ""))]
        
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
        cleaned_doc.add_paragraph(f"Date: May 22, 2026\n")
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
            total_yield_tons = round(data
