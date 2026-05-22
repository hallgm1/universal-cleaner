with tab1:
    st.header("Corporate Spreadsheet Pipeline")
    uploaded = st.file_uploader("Upload Ledger (.csv/.xlsx)", type=["csv", "xlsx"], key="spreadsheet_upload")
    
    if uploaded:
        # 1. Load data
        if uploaded.name.endswith('.csv'):
            df = pd.read_csv(uploaded)
        else:
            df = pd.read_excel(uploaded)
            
        # 2. Surgical Sanitization
        for col in df.columns:
            # Targeted cleanup: If it contains 'salary' or 'usd', strip non-numeric
            if any(term in col.lower() for term in ['salary', 'usd', 'amount', 'price']):
                # Regex: Remove everything EXCEPT digits and decimal points
                df[col] = df[col].astype(str).str.replace(r'[^\d.]', '', regex=True)
                # Convert to numeric, turn errors (empty/text) to NaN, then 0.0
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
        
        # 3. Render
        st.success("Data Sanitized & Structured")
        st.dataframe(df, use_container_width=True)
        
        # 4. Export
        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button("Download Sanitized Data", csv, "sanitized_data.csv", "text/csv")
