import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime
import re

# --- UI Setup ---
st.set_page_config(page_title="Retail Stock Tracker Tool", layout="centered")
st.title("📦 Retail Stock Tracker Tool")
st.write("Upload your raw inventory data, configure your filters, and download a targeted multi-tab Excel report.")

# 1. File Uploader
uploaded_file = st.file_uploader("Upload your raw data (CSV or Excel)", type=['csv', 'xlsx', 'xls'])

if uploaded_file is not None:
    # Read the uploaded file
    try:
        if uploaded_file.name.lower().endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        st.success("File uploaded successfully!")
    except Exception as e:
        st.error(f"Error reading file: {e}")
        st.stop()

    # 2. Interactive Filters for the User
    st.subheader("Report Settings")
    
    # Get all columns to populate the dropdown
    all_columns = df.columns.tolist()
    default_index = all_columns.index('Brand Code') if 'Brand Code' in all_columns else 0
    
    # Dropdown: What to split by
    split_col = st.selectbox(
        "Which column would you like to split the report tabs by?", 
        all_columns, 
        index=default_index
    )
    
    # Checkbox: Remove generics/blanks
    remove_generics = st.checkbox(
        f"Remove blank cells and 'Generic' values from the '{split_col}' column", 
        value=True
    )

    st.markdown("---")
    st.subheader("Top Contributors Settings")
    st.write("To keep the Excel file manageable, you can limit how many individual tabs are created.")

    # Dropdown: Metric to rank by
    ranking_metric = st.selectbox(
        "Select the metric to rank the top contributors:",
        ["Total_Stock_Value", "Live_SKUs_Count"]
    )

    # Number input: How many tabs to generate
    top_n = st.number_input(
        f"How many top {split_col} tabs do you want to generate?", 
        min_value=1, 
        max_value=200, 
        value=10,
        step=1
    )

    # 3. Process Data Button
    if st.button("Generate Excel Report"):
        with st.spinner("Crunching the numbers and building your Excel file..."):
            
            # --- Data Processing Logic ---
            if remove_generics:
                df = df.dropna(subset=[split_col])
                df = df[~df[split_col].astype(str).str.strip().str.lower().isin(['generic', ''])]
            else:
                df[split_col] = df[split_col].fillna("Unknown_Blank")

            # Select only the requested columns
            cols_to_keep = [
                'SKU', 'Product Subtype', 'Brand Code', 'Psku', 'Gtin', 'Title En', 
                'Offer Price Lcy', 'Lowest Comp Price Tdy', 'Lowest Comp Link', 
                'Unit Cost Lcy', 'Rebate per Unit Ex VAT Lcy', 'Provision Value per Unit Ex VAT', 
                'Ret Live Stock', 'DOH', 'Ret Units Yst', 'Ret Units L7d', 
                'Ret Units L30d', 'Ret Units L60d', 'Ret Units Mtd', 'Age'
            ]
            
            if split_col not in cols_to_keep:
                cols_to_keep.append(split_col)

            missing_cols = [col for col in cols_to_keep if col not in df.columns]
            if missing_cols:
                st.warning(f"Note: The following expected columns are missing from the uploaded file: {', '.join(missing_cols)}")
            
            cols_to_keep = [col for col in cols_to_keep if col in df.columns]
            main_df = df[cols_to_keep].copy()

            # Ensure calculation columns are numeric
            for numeric_col in ['Offer Price Lcy', 'Unit Cost Lcy', 'Ret Live Stock']:
                if numeric_col in main_df.columns:
                    main_df[numeric_col] = pd.to_numeric(main_df[numeric_col], errors='coerce')

            # Calculate Margin and Stock Value
            if 'Offer Price Lcy' in main_df.columns and 'Unit Cost Lcy' in main_df.columns:
                main_df['Margin'] = np.where(
                    (main_df['Unit Cost Lcy'] > 0) & (main_df['Offer Price Lcy'] > 0),
                    ((main_df['Offer Price Lcy'] / main_df['Unit Cost Lcy']) / main_df['Offer Price Lcy']),
                    np.nan
                )
                
            if 'Unit Cost Lcy' in main_df.columns and 'Ret Live Stock' in main_df.columns:
                main_df['Stock Value'] = main_df['Unit Cost Lcy'] * main_df['Ret Live Stock']

            # Create Pivot Table Summary
            if 'Stock Value' in main_df.columns and 'Offer Price Lcy' in main_df.columns:
                pivot_df = main_df.groupby(split_col).agg(
                    Total_Stock_Value=('Stock Value', 'sum'),
                    Live_SKUs_Count=('Offer Price Lcy', 'count') 
                ).reset_index()
            else:
                pivot_df = pd.DataFrame() 

            # --- Extract Top N Contributors ---
            if not pivot_df.empty:
                # Sort the summary dataframe by the chosen metric in descending order
                pivot_df = pivot_df.sort_values(by=ranking_metric, ascending=False)
                # Grab exactly the top N names from the split column
                top_categories = pivot_df[split_col].head(top_n).tolist()
            else:
                # Fallback if metrics are missing
                top_categories = main_df[split_col].unique()[:top_n]

            # --- Excel Generation in Memory ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Write Main Data
                main_df.to_excel(writer, sheet_name='Main Data', index=False)
                
                # Write Pivot Table
                if not pivot_df.empty:
                    summary_sheet_name = f'{split_col[:20]} Summary'
                    pivot_df.to_excel(writer, sheet_name=summary_sheet_name, index=False)
                
                # Write individual sheets ONLY for the top N categories
                for value in top_categories:
                    value_df = main_df[main_df[split_col] == value]
                    
                    # Create safe sheet names for Excel
                    safe_sheet_name = re.sub(r'[\\/*?:\[\]]', '_', str(value))
                    safe_sheet_name = safe_sheet_name[:31] 
                    
                    if not safe_sheet_name.strip():
                        safe_sheet_name = "Data_Sheet"
                        
                    value_df.to_excel(writer, sheet_name=safe_sheet_name, index=False)

            # Prepare for download
            output.seek(0)
            today_date = datetime.now().strftime("%d-%m-%y")
            file_name = f'Retail Stock Tracker - Top {top_n} - {today_date}.xlsx'

            st.success(f"Report generated successfully! Configured for the Top {top_n} tabs.")
            
            # 4. Download Button
            st.download_button(
                label="📥 Download Excel Report",
                data=output,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
