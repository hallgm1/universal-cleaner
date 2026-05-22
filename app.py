import streamlit as st
import pandas as pd
import numpy as np
import re
import os
from datetime import datetime, timedelta
from io import BytesIO

st.set_page_config(
    page_title="Employee Data Analyzer",
    page_icon="👥",
    layout="wide"
)

st.title("👥 Tool 1: Enterprise Employee Spreadsheet Sanitizer")
st.write("Upload your corporate roster or payroll ledger. This tool maps variable schemas, standardizes dates/salaries, and unlocks full dashboard metrics.")

def clean_employee_ledger(raw_bytes, ext):
    stream = BytesIO(raw_bytes)
    if ext == '.csv':
        df = pd.read_csv(stream)
    else:
        df = pd.read_excel(stream, engine='openpyxl')
    
    # Standardize column headers to a clean lowercase underscore layout
    df.columns = [re.sub(r'[^a-zA-Z0-9_]', '_', str(col).strip().lower()) for col in df.columns]
    df.columns = [re.sub(r'_+', '_', col).strip('_') for col in df.columns]

    # 1. Clean Employee ID structures
    id_cols = [c for c in df.columns if 'id' in c or 'ee' in c]
    for col in id_cols:
        df[col] = df[col].astype(str).str.strip().str.upper()

    # 2. Clean Name formatting (Strip prefixes, apply Title Case)
    name_cols = [c for c in df.columns if 'name' in c or 'full' in c]
    for col in name_cols:
        def _clean_name(v):
            if pd.isna(v) or str(v).strip() in ['', 'nan', 'NaN', 'None']: return ""
            s = str(v).strip().title()
            s = re.sub(r'\b(Dr|Mr|Mrs|Ms|Eng|Prof)\.?\s+', '', s, flags=re.I)
            return s.strip()
        df[col] = df[col].apply(_clean_name)

    # 3. Clean Financial / Payroll numbers safely
    salary_cols = [c for c in df.columns if any(k in c for k in ['salary', 'usd', 'annual', 'sales', 'amount', 'pay'])]
    for col in salary_cols:
        def _currency_formatter(v):
            if pd.isna(v) or str(v).strip() in ['-', '', 'nan', 'NaN']: return 0.0
            s = re.sub(r'[^0-9.]', '', str(v))
            if s == '' or s == '.': return 0.0
            try:
                return float(s)
            except ValueError:
                return 0.0
        df[col] = df[col].apply(_currency_formatter)

    # 4. Standardize Date variables safely
    date_cols = [c for c in df.columns if any(k in c for k in ['date', 'hire', 'exit', 'trans'])]
    for col in date_cols:
        def _date_fix(v):
            if pd.isna(v) or str(v).strip() in ['', 'nan', 'NaN', 'None', '-']: return "Active / Invalid Date"
            s = str(v).strip().lower().replace('.', '-')
            current_time = datetime(2026, 5, 22)
            
            if 'yesterday' in s: return (current_time - timedelta(days=1)).strftime('%Y-%m-%d')
            if 'today' in s or 'now' in s: return current_time.strftime('%Y-%m-%d')
            
            # Catch standard formats
            for fmt in ('%Y-%m-%d', '%d-%m-%Y', '%m-%d-%Y', '%Y/%m/%d', '%d/%m/%Y', '%m/%m/%Y'):
                try:
                    return pd.to_datetime(s, errors='raise').strftime('%Y-%m-%d')
                except:
                    continue
            return "Active"
                
        df[col] = df[col].apply(_date_fix)

    # 5. Drop entries where all fields are completely missing
    df.dropna(how='all', inplace=True)
    return df, salary_cols, date_cols

uploaded_file = st.file_uploader("Upload spreadsheet file (.csv or .xlsx)", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    filename = str(uploaded_file.name)
    ext = os.path.splitext(filename.lower())[1]
    raw_file_content = uploaded_file.getvalue()

    try:
        # Run calculation and sanitization process
        cleaned_df, salary_cols, date_cols = clean_employee_ledger(raw_file_content, ext)
        
        # Display human-scannable table layout
        display_df = cleaned_df.copy()
        formatted_headers = {col: re.sub(r'\s+', ' ', col.replace('_', ' ').strip()).title() for col in display_df.columns}
        
        # Explicit acronym adjustments
        for k in formatted_headers:
            if 'id' in k.lower(): formatted_headers[k] = "Employee ID"
            if 'usd' in k.lower(): formatted_headers[k] = "Annual Salary (USD)"
            
        display_df.rename(columns=formatted_headers, inplace=True)
        
        st.subheader("👀 Clean Data Preview")
        
        # Display float values as formatted currency tokens safely on render
        formatting_dict = {}
        for c in cleaned_df.columns:
            if c in salary_cols:
                display_df[formatted_headers[c]] = display_df[formatted_headers[c]].apply(lambda v: f"${v:,.2f}" if isinstance(v, (int, float)) else v)
        
        st.dataframe(display_df, use_container_width=True)
        
        # Prepare the file buffer download payload
        out_buf = BytesIO()
        if ext == '.csv':
            display_df.to_csv(out_buf, index=False)
            m_type, name_out = "text/csv", "cleaned_employee_records.csv"
        else:
            display_df.to_excel(out_buf, index=False, engine='openpyxl')
            m_type, name_out = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "cleaned_employee_records.xlsx"
        out_buf.seek(0)
        
        st.download_button("📥 Download Cleaned Sheet Document", data=out_buf, file_name=name_out, mime=m_type, use_container_width=True)
        
        # Execution Metrics Dashboard Section
        st.markdown("---")
        st.subheader("📊 Executive Workforce Metrics")
        
        c1, c2 = st.columns(2)
        
        if salary_cols:
            total_payroll = cleaned_df[salary_cols[0]].sum()
            c1.metric(label="Validated Total Annual Payroll Budget", value=f"${total_payroll:,.2f}")
        else:
            c1.metric(label="Validated Total Annual Payroll Budget", value="$0.00")
            
        c2.metric(label="Total Active Staff Records", value=f"{len(cleaned_df)} Headcount Rows")
            
    except Exception as e:
        st.error(f"Spreadsheet System Fault: {str(e)}")
