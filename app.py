import streamlit as st
import pandas as pd
import numpy as np
import re
import os
from datetime import datetime, timedelta
from io import BytesIO

st.set_page_config(
    page_title="Data Cleaner Pro",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Dedicated Spreadsheet Cleaning Tool")
st.write("Upload any Excel or CSV ledger. This tool automatically handles variable column layouts, sanitizes headers, and builds custom dashboard analytics.")

def clean_spreadsheet(raw_bytes, ext):
    stream = BytesIO(raw_bytes)
    if ext == '.csv':
        df = pd.read_csv(stream)
    else:
        df = pd.read_excel(stream, engine='openpyxl')
    
    # Handle single-string collapsed lines
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
                                if d <= 12 and m <= 31: return f"{y}-{d:02d}-{m:02d}"
                            
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

uploaded_file = st.file_uploader("Upload target sheet ledger", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    filename = str(uploaded_file.name)
    ext = os.path.splitext(filename.lower())[1]
    raw_file_content = uploaded_file.getvalue()

    try:
        cleaned_df = clean_spreadsheet(raw_file_content, ext)
        
        # Display polished presentation headers
        display_df = cleaned_df.copy()
        formatted_headers = {col: re.sub(r'\s+', ' ', col.replace('_', ' ').strip()).title() for col in display_df.columns}
        display_df.rename(columns=formatted_headers, inplace=True)
        
        st.subheader("👀 Cleaned Data Preview")
        st.dataframe(display_df, use_container_width=True)
        
        out_buf = BytesIO()
        if ext == '.csv':
            display_df.to_csv(out_buf, index=False)
            m_type, name_out = "text/csv", "cleaned_data.csv"
        else:
            display_df.to_excel(out_buf, index=False, engine='openpyxl')
            m_type, name_out = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "cleaned_data.xlsx"
        out_buf.seek(0)
        
        st.download_button("📥 Download Cleaned Sheet", data=out_buf, file_name=name_out, mime=m_type, use_container_width=True)
        
        # Dynamic Analysis Engine - Checks dataset columns natively
        sales_cols = [c for c in cleaned_df.columns if any(k in c for k in ['sales', 'amount', 'price', 'revenue', 'cost', 'total', 'salary'])]
        date_cols = [c for c in cleaned_df.columns if any(k in c for k in ['date', 'trans', 'hire'])]
        zone_cols = [c for c in cleaned_df.columns if any(k in c for k in ['zone', 'region', 'dept', 'business_unit', 'unit'])]
        
        st.markdown("---")
        st.subheader("📊 Operational Analytics Metrics")
        
        if sales_cols:
            metric_df = cleaned_df.copy()
            metric_df[sales_cols[0]] = metric_df[sales_cols[0]].astype(str).str.replace(',', '').astype(float)
            
            c1, c2 = st.columns(2)
            c1.metric(label="Sum Value Volume", value=f"{metric_df[sales_cols[0]].sum():,.2f}")
            c2.metric(label="Total Processed Records", value=f"{len(metric_df)} Valid Rows")
            
            if date_cols:
                st.write("📈 **Financial Metrics Over Time**")
                valid_dates = metric_df[metric_df[date_cols[0]] != 'Invalid Date'].copy()
                if not valid_dates.empty:
                    valid_dates[date_cols[0]] = pd.to_datetime(valid_dates[date_cols[0]])
                    trend = valid_dates.groupby(date_cols[0])[sales_cols[0]].sum().reset_index()
                    st.line_chart(data=trend, x=date_cols[0], y=sales_cols[0])
            
            if zone_cols:
                st.write("🌍 **Categorical Distribution Breakdown**")
                zone_chart = metric_df.groupby(zone_cols[0])[sales_cols[0]].sum().reset_index()
                st.bar_chart(data=zone_chart, x=zone_cols[0], y=sales_cols[0])
        else:
            st.info("ℹ️ Structural validation complete. Visual charts are minimized since no currency/numeric columns were found.")
            
    except Exception as e:
        st.error(f"Spreadsheet Engine Fault: {str(e)}")
