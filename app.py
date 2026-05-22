with tab1:
    st.header("Corporate Spreadsheet Pipeline")
    uploaded = st.file_uploader("Upload Ledger (.csv/.xlsx)", type=["csv", "xlsx"], key="spreadsheet_upload")
    
    if uploaded:
        # Load the data
        df = pd.read_csv(uploaded) if uploaded.name.endswith('.csv') else pd.read_excel(uploaded)
        
        # Target the specific salary column by partial match
        target_col = [c for c in df.columns if 'salary' in c.lower() or 'usd' in c.lower()]
        
        if target_col:
            col_name = target_col[0]
            # 1. Force the entire column to be string
            # 2. Remove commas, dollar signs, and non-numeric junk characters
            # 3. Handle values with trailing dots or weird endings
            df[col_name] = df[col_name].astype(str).str.replace(r'[^\d.]', '', regex=True)
            
            # 4. Final conversion to numeric
            df[col_name] = pd.to_numeric(df[col_name], errors='coerce')
            
            st.success(f"Sanitized Column: {col_name}")
            st.dataframe(df.head(10), use_container_width=True)
        else:
            st.error("Could not find a column containing 'Salary' or 'USD'. Check your headers.")
