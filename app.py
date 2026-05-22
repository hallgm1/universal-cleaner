with tab1:
    st.header("Corporate Spreadsheet Pipeline")
    uploaded = st.file_uploader("Upload Ledger (.csv/.xlsx)", type=["csv", "xlsx"], key="spreadsheet_upload")
    
    if uploaded:
        # Load data
        if uploaded.name.endswith('.csv'):
            df = pd.read_csv(uploaded)
        else:
            df = pd.read_excel(uploaded)
            
        # Refined Sanitization Engine
        for col in df.columns:
            if 'usd' in col.lower() or 'salary' in col.lower():
                # 1. Convert to string
                # 2. Remove commas
                # 3. Use regex to keep only digits and the first decimal point
                df[col] = df[col].astype(str).str.replace(',', '').str.extract(r'(\d+\.?\d*)')[0]
                # 4. Force to numeric
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
        
        st.success("Data Sanitized Successfully")
        st.dataframe(df, use_container_width=True)
        
        # Export
        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button("Download Sanitized Data", csv, "sanitized_data.csv", "text/csv")
