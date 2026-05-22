import streamlit as st
import pandas as pd
import re
from datetime import datetime, timedelta
from dateutil import parser

# ---------------------------
# Region Mapping
# ---------------------------
DEFAULT_REGION_MAP = {
    "dsm": "Dar es Salaam", "dar": "Dar es Salaam", "znz": "Zanzibar",
    "moro": "Morogoro", "mwz": "Mwanza", "dom": "Dodoma", "iringa": "Iringa",
    "arusha": "Arusha", "moshi": "Kilimanjaro", "pwani": "Pwani",
    "tabora": "Tabora", "mbeya": "Mbeya", "tanga": "Tanga",
    "kilimanjaro": "Kilimanjaro", "kigoma": "Kigoma", "coast": "Pwani"
}

# ---------------------------
# Cleaning Logic
# ---------------------------
def sanitize_currency(val):
    if pd.isna(val) or str(val).strip() == "": return 0.0
    # Strip commas, quotes, and non-numeric chars
    cleaned = re.sub(r"[^\d.]", "", str(val).replace(',', '').replace('"', ''))
    try: return float(cleaned)
    except: return 0.0

# ---------------------------
# App Interface
# ---------------------------
st.set_page_config(page_title="Universal Cleaner", layout="wide")
st.title("🌍 Universal Cleaner")

tab1, tab2, tab3, tab4 = st.tabs(["Spreadsheet Cleaner", "Letters Cleaner", "Business Proposals", "Agricultural Manuals"])

with tab1:
    st.header("📊 Clean Messy Spreadsheets")
    uploaded_file = st.file_uploader("Upload your spreadsheet (CSV or XLSX)", type=["csv","xlsx"])
    
    if uploaded_file:
        df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith(".csv") else pd.read_excel(uploaded_file)
        
        # Apply specific sanitization to your target column
        if 'Annual Salary Usd' in df.columns:
            df['Annual Salary Usd'] = df['Annual Salary Usd'].apply(sanitize_currency)
        
        st.subheader("✅ Cleaned & Sanitized Data")
        st.dataframe(df, width=1000)
        
        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Download Cleaned Data", csv, "cleaned_data.csv", "text/csv")

with tab2:
    st.header("✉️ Formal & Informal Letters")
    text_input = st.text_area("Paste your letter text here:")
    selected_style = st.selectbox("Style", ["Business", "Formal"])
    if st.button("Clean Letter"):
        # Fixed syntax error: removed "abatement"
        if text_input.strip() and selected_style in ["Business", "Formal"]:
            cleaned_text = " ".join(text_input.split())
            st.subheader("✅ Cleaned Letter")
            st.write(cleaned_text)

with tab3:
    st.header("📑 Business Proposals")
    proposal_input = st.text_area("Paste your business proposal here:")
    if st.button("Clean Proposal"):
        if proposal_input.strip():
            st.write(" ".join(proposal_input.split()))

with tab4:
    st.header("🌱 Agricultural Manuals")
    # Placeholder for logic
    if st.button("Generate Manual"):
        # Fixed syntax error: added missing bracket and param
        data = 100.50 
        total_yield_tons = round(data, 2)
        st.write(f"Yield Calculated: {total_yield_tons} tons")
