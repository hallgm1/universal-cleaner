import streamlit as st
import pandas as pd
import numpy as np
import re
import os
from datetime import datetime, timedelta
from io import BytesIO

st.set_page_config(
    page_title="Universal Spreadsheet Cleaner",
    page_icon="📊",
    layout="wide"
)

st.title("📊 UNIVERSAL SPREADSHEET CLEANER")
st.write("Upload any CSV, XLS, or XLSX spreadsheet ledger. This system dynamically scales to any column size, auto-detects column data types, strips formatting noise, and standardizes records.")

def clean_universal_spreadsheet(raw_bytes, ext):
    stream = BytesIO(raw_bytes)
    if ext == '.csv':
        df = pd.read_csv(stream)
    else:
        df = pd.read_excel(stream, engine='openpyxl')
    
    # 1. Handle single-column comma-collapsed rows dynamically
    if len(df.columns) == 1:
        raw_col = df.columns[0]
        header_line = str(raw_col).replace('"', '').strip()
        new_headers = [re.sub(r'[^a-zA-Z0-9_]', '_', h.strip().lower()) for h in header_line.split(',')]
        
        split_rows = []
        for val in df.iloc[:, 0]:
            row_str = str(val).strip().strip('"')
            row_cells = re.split(r',(?=(?:[^"]*"[^"]*")*[^"]*$)', row_str)
            split_rows.append([c.strip('"').strip() for c in row_cells])
            
        max_cols = max([len(r) for r in split_rows]) if split_rows else len(new_headers)
        if len(new_headers) < max_cols:
            new_headers += [f"column_{i}" for i in range(len(new_headers), max_cols)]
        df = pd.DataFrame(split_rows, columns=new_headers[:max_cols])
    else:
        df.columns = [re.sub(r'[^a-zA-Z0-9_]', '_', str(col).strip().lower()) for col in df.columns]
    
    df.columns = [re.sub(r'_+', '_', col).strip('_') for col in df.columns]

    # Find specialized column configurations natively
    phone_cols = [c for c in df.columns if any(k in c for k in ['phone', 'contact', 'tel', 'mobile', 'num'])]
    salary_cols = [c for c in df.columns if any(k in c for k in ['salary', 'usd', 'annual', 'sales', 'amount', 'pay', 'price', 'revenue', 'cost', 'total', 'finance'])]
    date_cols = [c for c in df.columns if any(k in c for k in ['date', 'hire', 'exit', 'trans', 'time'])]
    name_cols = [c for c in df.columns if 'name' in c or 'full' in c]
    zone_cols = [c for c in df.columns if any(k in c for k in ['zone', 'region', 'dept', 'unit', 'state', 'city'])]

    # 2. Universal Data Type Sanitizer
    for col in df.columns:
        # Standardize strings, clean whitespaces, and catch missing values
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].str.replace(r'\s+', ' ', regex=True)
        df[col] = df[col].replace(['nan', 'NaN', 'None', 'NULL', 'null', ''], np.nan)
        
        if col in id_cols := [c for c in df.columns if 'id' in c or 'ee' in c]:
            df[col] = df[col].astype(str).str.strip().str.upper()

        elif col in name_cols:
            def _clean_name(v):
                if pd.isna(v) or str(v).strip() in ['', 'nan', 'NaN']: return ""
                s = str(v).strip().title()
                s = re.sub(r'\b(Dr|Mr|Mrs|Ms|Eng|Prof)\.?\s+', '', s, flags=re.I)
                return s.strip()
            df[col] = df[col].apply(_clean_name)
            
        elif col in phone_cols:
            def _pre_parse_phone(v):
                if pd.isna(v) or str(v).strip() in ['', 'nan', 'NaN']: return ""
                s = re.sub(r'[^0-9]', '', str(v))
                if s.startswith('0') and len(s) == 10: return '+255' + s[1:]
                if s.startswith('255') and len(s) == 12: return '+' + s
                if len(s) > 0 and not s.startswith('+'): return '+' + s
                return s
            df[col] = df[col].apply(_pre_parse_phone)
            
        elif col in salary_cols:
            def _currency_formatter(v):
                if pd.isna(v) or str(v).strip() in ['-', '', 'nan', 'NaN']: return 0.0
                s = re.sub(r'[^0-9.]', '', str(v))
                if s == '' or s == '.': return 0.0
                try:
                    return float(s)
                except ValueError:
                    return 0.0
            df[col] = df[col].apply(_currency_formatter)
            
        elif col in date_cols:
            def _date_fix(v):
                if pd.isna(v) or str(v).strip() in ['', 'nan', 'NaN', '-']: return "Active / Invalid Date"
                s = str(v).strip().lower().replace('.', '-')
                current_time = datetime(2026, 5, 22)
                
                if 'yesterday' in s: return (current_time - timedelta(days=1)).strftime('%Y-%m-%d')
                if 'today' in s or 'now' in s: return current_time.strftime('%Y-%m-%d')
                
                try:
                    return pd.to_datetime(s, errors='raise').strftime('%Y-%m-%d')
                except:
                    return "Active"
            df[col] = df[col].apply(_date_fix)
            
        elif col in zone_cols:
            def _normalize_region(v):
                if pd.isna(v) or str(v).strip() in ['', 'nan', 'NaN']: return "Unknown"
                s = str(v).strip().lower()
                if s in ['dar', 'dsm', 'dar es salaam', 'tanzania']: return "Dar es Salaam"
                if s in ['znz', 'zanzibar']: return "Zanzibar"
                return str(v).strip().title()
            df[col] = df[col].apply(_normalize_region)

    df.dropna(how='all', inplace=True)
    return df, salary_cols, date_cols, zone_cols

uploaded_file = st.file_uploader("Upload target sheet document", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    filename = str(uploaded_file.name)
    ext = os.path.splitext(filename.lower())[1]
    raw_file_content = uploaded_file.getvalue()

    try:
        cleaned_df, salary_cols, date_cols, zone_cols = clean_universal_spreadsheet(raw_file_content, ext)
        
        display_df = cleaned_df.copy()
        formatted_headers = {col: re.sub(r'\s+', ' ', col.replace('_', ' ').strip()).title() for col in display_df.columns}
        
        for k in formatted_headers:
            if 'id' in k.lower(): formatted_headers[k] = "Employee ID" if 'ee' in k.lower() else k.upper().replace('_', ' ')
            if 'usd' in k.lower() or 'salary' in k.lower(): formatted_headers[k] = f"{formatted_headers[k]} (USD)"
            
        display_df.rename(columns=formatted_headers, inplace=True)
        
        st.subheader("👀 Cleaned Grid Data View")
        
        render_df = display_df.copy()
        for c in cleaned_df.columns:
            if c in salary_cols:
                render_df[formatted_headers[c]] = pd.to_numeric(cleaned_df[c], errors='coerce').apply(lambda v: f"${v:,.2f}" if pd.notna(v) else "$0.00")
        
        st.dataframe(render_df, use_container_width=True)
        
        out_buf = BytesIO()
        if ext == '.csv':
            display_df.to_csv(out_buf, index=False)
            m_type, name_out = "text/csv", "universal_cleaned_records.csv"
        else:
            display_df.to_excel(out_buf, index=False, engine='openpyxl')
            m_type, name_out = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "universal_cleaned_records.xlsx"
        out_buf.seek(0)
        
        st.download_button("📥 Download Cleaned Spreadsheet File", data=out_buf, file_name=name_out, mime=m_type, use_container_width=True)
        
        st.markdown("---")
        st.subheader("📊 Dynamic Financial KPI Matrix Dashboard")
        
        col1, col2 = st.columns(2)
        
        if salary_cols:
            total_sum = pd.to_numeric(cleaned_df[salary_cols[0]], errors='coerce').sum()
            col1.metric(label="Aggregated Financial Value Summary", value=f"${total_sum:,.2f}")
        else:
            col1.metric(label="Aggregated Financial Value Summary", value="$0.00")
            
        col2.metric(label="Total Cleaned Safe Records", value=f"{len(cleaned_df)} Active Rows")
        
        if salary_cols and date_cols:
            st.write("📈 **Value Aggregation Metric Tracking Over Time**")
            valid_dates_df = cleaned_df[~cleaned_df[date_cols[0]].isin(["Active", "Active / Invalid Date", "Invalid Date"])].copy()
            if not valid_dates_df.empty:
                valid_dates_df[date_cols[0]] = pd.to_datetime(valid_dates_df[date_cols[0]])
                valid_dates_df[salary_cols[0]] = pd.to_numeric(valid_dates_df[salary_cols[0]], errors='coerce')
                time_trend = valid_dates_df.groupby(date_cols[0])[salary_cols[0]].sum().reset_index()
                st.line_chart(data=time_trend, x=date_cols[0], y=salary_cols[0])
            
        if salary_cols and zone_cols:
            st.write("🌍 **Categorical Distribution Breakdown Splits**")
            chart_df = cleaned_df.copy()
            chart_df[salary_cols[0]] = pd.to_numeric(chart_df[salary_cols[0]], errors='coerce')
            zone_chart = chart_df.groupby(zone_cols[0])[salary_cols[0]].sum().reset_index()
            st.bar_chart(data=zone_chart, x=zone_cols[0], y=salary_cols[0])
            
    except Exception as e:
        st.error(f"Universal Spreadsheet Cleaner Engine Fault: {str(e)}")
