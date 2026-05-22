import streamlit as st
import pandas as pd
import numpy as np
import re
import os
from datetime import datetime, timedelta
from io import BytesIO

st.set_page_config(page_title="Universal Spreadsheet Cleaner", page_icon="📊", layout="wide")

st.title("📊 UNIVERSAL SPREADSHEET CLEANER")
st.write("Professional-grade data sanitization engine.")

# --- ROBUST FILE HANDLER ---
def process_data(uploaded_file):
    filename = uploaded_file.name
    ext = os.path.splitext(filename.lower())[1]
    raw_bytes = uploaded_file.getvalue()
    
    # Logic to handle Excel vs CSV
    if ext == '.csv':
        df = pd.read_csv(BytesIO(raw_bytes))
    else:
        df = pd.read_excel(BytesIO(raw_bytes), engine='openpyxl')
    
    # Sanitize Columns
    df.columns = [re.sub(r'[^a-zA-Z0-9_]', '_', str(col).strip().lower()) for col in df.columns]
    
    # Force Numeric for Salary (Preventing String-Math crashes)
    salary_cols = [c for c in df.columns if any(k in c for k in ['salary', 'usd', 'amount', 'pay', 'price'])]
    for col in salary_cols:
        df[col] = pd.to_numeric(df[col].replace(r'[\$,]', '', regex=True), errors='coerce').fillna(0.0)
        
    return df, salary_cols

# --- SESSION STATE GUARD ---
if 'data' not in st.session_state:
    st.session_state.data = None

uploaded_file = st.file_uploader("Upload spreadsheet:", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        df, salary_cols = process_data(uploaded_file)
        st.session_state.data = df
        
        st.subheader("✅ Data Sanitized")
        st.dataframe(df.head(10), use_container_width=True)
        
        # Download logic
        out_buf = BytesIO()
        df.to_csv(out_buf, index=False)
        st.download_button("📥 Download Cleaned Data", data=out_buf.getvalue(), file_name="cleaned_data.csv", mime="text/csv")
        
        # Metric Panel
        if salary_cols:
            total = df[salary_cols[0]].sum()
            st.metric("Total Aggregated Financial Value", f"${total:,.2f}")
            
    except Exception as e:
        st.error(f"Engine Error: {e}. Please clear and re-upload.")
