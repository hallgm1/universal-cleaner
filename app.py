with tab1:
    st.header("Corporate Spreadsheet Pipeline")
    uploaded = st.file_uploader("Upload Ledger (.csv/.xlsx)", type=["csv", "xlsx"], key="spreadsheet_upload")
    
    if uploaded:
        # Specialized Parser for Mkuranga-Standard Data
        if uploaded.name.endswith('.csv'):
            df = pd.read_csv(uploaded)
        else:
            df = pd.read_excel(uploaded)
            
        # Surgical extraction of financial columns
        for col in df.columns:
            if 'usd' in col.lower() or 'salary' in col.lower():
                # Force cleanup: strip symbols, then force to float
                df[col] = df[col].astype(str).replace(r'[^\d.]', '', regex=True)
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
        
        st.dataframe(df, use_container_width=True)
        
        # Add a download button for the sanitized data
        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button("Download Sanitized Data", csv, "sanitized_data.csv", "text/csv")
