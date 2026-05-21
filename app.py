import streamlit as st
import pandas as pd
import re
import os
from io import BytesIO
from docx import Document

# Initialize a responsive UI engine optimized for desktop and mobile screens
st.set_page_config(
    page_title="Universal Master Cleaner",
    page_icon="🧹",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- REGIONAL AGRONOMY DATABASE (Mkuranga District, Coast Region, Tanzania) ---
MKURANGA_AGRI_DB = {
    "soil_profile": "Coastal sandy-loam structures, high drainage, high ambient humidity profiles.",
    "fertilizers_available": [
        "UREA (High Nitrogen - vegetative growth booster)",
        "DAP (Diammonium Phosphate - critical root establishers for basal application)",
        "CAN (Calcium Ammonium Nitrate - soil-neutral nitrogen feed)",
        "NPK 20:10:10 / 15:15:15 (Balanced compound macronutrients)",
        "MOP (Muriate of Potash - critical for fruit weight, sizing, sugar profiling)",
        "Organic Compost / Manure (Vital for moisture retention in sandy coastal matrices)"
    ],
    "varieties": {
        "Vegetables": {
            "Tomatoes": {"varieties": ["Tanya", "Assila F1", "Eden F1"], "spacing": "60cm x 50cm", "plan": "DAP basal setup. Split-apply UREA/CAN at weeks 3 and 6. Balance with NPK at flowering stage."},
            "Okra (Bamia)": {"varieties": ["Pusa Sawani", "Clemson Spineless"], "spacing": "50cm x 30cm", "plan": "Heavy basal organic manure integration. Top-dress with NPK/UREA cycles every 21 days."},
            "Watermelon": {"varieties": ["Sukari F1", "Safari F1"], "spacing": "150cm x 100cm", "plan": "High manure volume + DAP baseline. Shift heavily to MOP and NPK combinations immediately post-fruit-set for high brix/sweetness parameters."}
        },
        "Fruits": {
            "Mangoes": {"varieties": ["Apple Mango", "Kent", "Keitt"], "spacing": "9m x 9m", "plan": "Apply annual compost loops. Top-dress NPK post-harvest pruning sequences and immediately ahead of flower induction."},
            "Pineapples": {"varieties": ["Smooth Cayenne", "Queen"], "spacing": "90cm x 60cm x 30cm", "plan": "Highly responsive to localized high nitrogen feeds. Run micro-dosed UREA/NPK splits frequently across rainy periods."},
            "Cashew": {"varieties": ["Naliendele Improved Clones"], "spacing": "12m x 12m", "plan": "Integrate structured sulfur dusting schedules for powdery mildew vector control. Top-dress NPK during early seasonal rains."}
        }
    }
}

# --- TRANSFORMATION PIPELINE ENGINES ---

def clean_spreadsheet(uploaded_file, ext):
    """Engine 1: Multi-industry global spreadsheet cleaner & structural normalizer."""
    # Read the file safely into memory
    if ext == '.csv':
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
    
    # 1. Universal Text Cleaning: Strip white spaces from columns and text cells safely
    df.columns = [re.sub(r'[^a-zA-Z0-9_]', '_', str(col).strip().lower()) for col in df.columns]
    
    for col in df.columns:
        # If the column contains text/object data, clean it
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str).str.strip()
            # Clean duplicate internal spaces (e.g. "John   Doe" -> "John Doe")
            df[col] = df[col].str.replace(r'\s+', ' ', regex=True)
            # Normalize common empty string placeholders back to true blanks
            df[col] = df[col].replace(['nan', 'NaN', 'None', 'NULL', 'null', ''], None)

    # 2. Universal Data Purge: Drop fully empty rows & remove duplicates
    df.dropna(how='all', inplace=True)
    df.drop_duplicates(inplace=True)
    
    return df

def clean_word_document(uploaded_file, tone_format):
    """Engine 2: Word text parsing engine with automatic tone adjustments."""
    doc = Document(uploaded_file)
    cleaned_doc = Document()
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        text = re.sub(r'\s+', ' ', text)
        if tone_format in ["Business", "Formal"]:
            text = re.sub(r"\bi'm\b", "I am", text, flags=re.I)
            text = re.sub(r"\bcan't\b", "cannot", text, flags=re.I)
            text = re.sub(r"\bdon't\b", "do not", text, flags=re.I)
            text = re.sub(r"\basap\b", "as soon as possible", text, flags=re.I)
            text = re.sub(r"\bhey\b|\bhi\b", "Dear Sir/Madam,", text, flags=re.I)
        elif tone_format == "Non-Formal":
            text = re.sub(r"\butilize\b", "use", text, flags=re.I)
            text = re.sub(r"\bsubsequent to\b", "after", text, flags=re.I)
        cleaned_doc.add_paragraph(text)
    output_stream = BytesIO()
    cleaned_doc.save(output_stream)
    output_stream.seek(0)
    return output_stream

def generate_mkuranga_agri_report(uploaded_file):
    """Engine 3: Evaluates horticultural inputs, matching them to Mkuranga protocols."""
    raw_text = uploaded_file.read().decode("utf-8", errors="ignore")
    report = [
        "==========================================================================",
        "      MKURANGA DISTRICT HORTICULTURE ALIGNED MANAGEMENT BLUEPRINT         ",
        "==========================================================================",
        f"Target Zone Focus: Mkuranga District, Coast Region (Pwani), Tanzania",
        f"Soil Matrix Profile: {MKURANGA_AGRI_DB['soil_profile']}\n",
        "[1. LOCALIZED HIGH-EFFICIENCY AGRI-INPUT AVAILABILITY]"
    ]
    for fert in MKURANGA_AGRI_DB['fertilizers_available']:
        report.append(f"  • {fert}")
    report.append("\n[2. MODERN AGRO-METHODOLOGY TARGET CROPS MATCHED]")
    
    matched = False
    for category, crops in MKURANGA_AGRI_DB['varieties'].items():
        report.append(f"\n--- {category.upper()} MANAGEMENT MAP ---")
        for crop_name, data in crops.items():
            if crop_name.lower() in raw_text.lower() or "all" in raw_text.lower() or len(raw_text) < 15:
                matched = True
                report.append(f"\n● Target crop: {crop_name}")
                report.append(f"  Recommended High-Yield Varietals: {', '.join(data['varieties'])}")
                report.append(f"  Precision Field Spacing Protocols: {data['spacing']}")
                report.append(f"  Modern Nutrient Integration Plan: {data['plan']}")
    if not matched:
        report.append("\n*Note: No specific keywords matched. Listing baseline regional varieties.*")
    return "\n".join(report)

# --- USER INTERFACE APP LAYOUT ---
st.title("🧹 Universal Master Data Cleaning Hub")
st.write("Upload any file. The system automatically routes data to run global industrial cleaning steps, adapt text styles, or output specialized crop plans.")

# Interactive Controls Configuration Sidebar
st.sidebar.header("🎛️ App Controls")
text_tone = st.sidebar.selectbox("Document Formatting Style", ["Business", "Formal", "Non-Formal"])

# Interface UI Element: Upload File Container
uploaded_file = st.file_uploader("Upload file to process (Excel, CSV, Word, or TXT)", type=["csv", "xlsx", "xls", "docx", "txt"])

if uploaded_file is not None:
    filename = uploaded_file.name
    _, file_extension = os.path.splitext(filename.lower())
    st.info(f"📂 Auto-Identified Input Type: **{file_extension.upper()}**")
    
    # Branch 1: Spreadsheet Execution
    if file_extension in ['.csv', '.xlsx', '.xls']:
        try:
            cleaned_df = clean_spreadsheet(uploaded_file, file_extension)
            st.subheader("👀 Preview Cleaned Grid")
            st.dataframe(cleaned_df.head(30), use_container_width=True)
            
            output_buffer = BytesIO()
            if file_extension == '.csv':
                cleaned_df.to_csv(output_buffer, index=False)
                mime_type, out_filename = "text/csv", "cleaned_master_spreadsheet.csv"
            else:
                cleaned_df.to_excel(output_buffer, index=False, engine='openpyxl')
                mime_type, out_filename = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "cleaned_master_spreadsheet.xlsx"
            output_buffer.seek(0)
            
            st.download_button(label="📥 Download Cleaned Spreadsheet", data=output_buffer, file_name=out_filename, mime=mime_type, use_container_width=True)
        except Exception as e:
            st.error(f"Spreadsheet Engine Error: {str(e)}")
            
    # Branch 2: Document Execution
    elif file_extension == '.docx':
        try:
            word_output_stream = clean_word_document(uploaded_file, text_tone)
            st.subheader("👀 Preview Status")
            st.success(f"Word document content streams successfully parsed, optimized, and set to the '{text_tone.upper()}' linguistic layout standard.")
            st.download_button(label=f"📥 Download Re-Formatted {text_tone} Document", data=word_output_stream, file_name="cleaned_master_document.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
        except Exception as e:
            st.error(f"Text Document Engine Error: {str(e)}")
            
    # Branch 3: Horticulture Execution
    elif file_extension == '.txt':
        try:
            agri_report = generate_mkuranga_agri_report(uploaded_file)
            st.subheader("👀 Preview Blueprint")
            st.text_area("Generated Output File Data Display", value=agri_report, height=350)
            st.download_button(label="📥 Download Mkuranga Production Plan (.txt)", data=agri_report, file_name="mkuranga_horticulture_clean_plan.txt", mime="text/plain", use_container_width=True)
        except Exception as e:
            st.error(f"Horticulture Module Error: {str(e)}")
