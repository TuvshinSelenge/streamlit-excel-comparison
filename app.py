import streamlit as st
import pandas as pd
import io
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule
from fuzzywuzzy import fuzz, process

# Define your column mappings
column_mappings = {
    'Isin Code': ['ISIN', 'Isin', 'Share ISIN Reference', 'FINANZINSTRUMENT_IDENT', 'Text23'],
    'Provision': ['Comm. Amount', 'Provision', 'Betrag (€)', 'Vergütung', 'EURMonat', 'Client Trailer Fees Amount In Consolidated Currency', 'Amount In Agreement Ccy', 'Bepro', 'Provisionsbetrag in Währung', 'Fee', 'BPROV', 'BpkEUR', 'Kommissionsbetrag', 'Commission Due Payment CCY', 'Fee (Payment Currency)', 'Amount In Partner Currency', 'Betrag (EUR)'],
    'Date': ['Date', 'Datum', 'Booking Date', 'End-Datum', 'Period End Date', 'Holding as of', 'STICHTAG', 'Period', 'Datum_str', 'Positionsdatum', 'Stichtag', 'Retrocession Date']
}

def read_files_from_upload(uploaded_files):
    """Read uploaded files and return a dictionary of DataFrames."""
    files_data = {}
    for uploaded_file in uploaded_files:
        try:
            df = pd.read_excel(uploaded_file, header=None)
            df = date_converter(df)
            df = set_correct_headers(df, column_mappings)
            files_data[uploaded_file.name] = df
            st.write(f"Read {uploaded_file.name} successfully with shape {df.shape}")
        except Exception as e:
            st.error(f"Error reading {uploaded_file.name}: {e}")
    return files_data

def date_converter(df):
    if "Statement Month" in df.columns and "Statement Year" in df.columns:
        df["Date"] = pd.to_datetime(df["Statement Month"].astype(str) + '-' + df["Statement Year"].astype(str), format='%m-%Y')
    return df

def set_correct_headers(df, column_mappings):
    """Set the correct headers by finding the row that contains any of the mapped column names."""
    for i, row in df.iterrows():
        if any(header in row.values for headers in column_mappings.values() for header in headers):
            df.columns = df.iloc[i]
            df = df.drop(i).reset_index(drop=True)
            return df
    return df

def rename_columns(df, column_mappings):
    """Rename columns in the dataframe based on the provided mappings."""
    rename_dict = {}
    for new_name, old_names in column_mappings.items():
        for old_name in old_names:
            if old_name in df.columns:
                rename_dict[old_name] = new_name
    df.rename(columns=rename_dict, inplace=True)
    return df

def filter_valid_rows(df):
    """Filter rows that have both a valid ISIN code and a valid Date."""
    if 'Isin Code' in df.columns and 'Date' in df.columns:
        df = df.dropna(subset=['Isin Code', 'Date'])
        df = df[df['Isin Code'].apply(lambda x: isinstance(x, str) and x.strip() != "")]
    return df

def convert_date_column(df, date_column_name):
    if date_column_name in df.columns:
        df[date_column_name] = pd.to_datetime(df[date_column_name], errors='coerce', format='%d.%m.%Y')
    return df

def compare_data(fundline_data, excel_data, column_mappings):
    """Compare data from matched Fundline and Excel files based on ISIN and Date and create a DataFrame showing differences."""
    comparison_files = []

    fundline_files = list(fundline_data.keys())
    excel_files = list(excel_data.keys())
    matched_files = match_files(fundline_files, excel_files)

    for fundline_file, excel_file in matched_files:
        fundline_df = fundline_data[fundline_file]
        excel_df = excel_data[excel_file]

        # Rename columns
        fundline_df = rename_columns(fundline_df, column_mappings)
        excel_df = rename_columns(excel_df, column_mappings)

        # Convert date columns to datetime64[ns]
        fundline_df = convert_date_column(fundline_df, 'Date')
        excel_df = convert_date_column(excel_df, 'Date')

        fundline_df  = fundline_df.groupby(['Isin Code', 'Date'])['Erwartete Prov. Whg'].sum().reset_index()
        excel_df  = excel_df.groupby(['Isin Code', 'Date'])['Provision'].sum().reset_index()

        # Check if required columns are present
        if 'Isin Code' in fundline_df.columns and 'Date' in fundline_df.columns and 'Isin Code' in excel_df.columns and 'Date' in excel_df.columns:
            # Merge dataframes on ISIN and Date with suffixes to handle common columns
            comparison_df = pd.merge(
                fundline_df, excel_df, 
                left_on=['Isin Code', 'Date'], 
                right_on=['Isin Code', 'Date'], 
                how='inner', 
                suffixes=('_Fundline', '_Excel')
            )

            # Identify the correct column names
            fundline_column = 'Erwartete Prov. Whg_Fundline' if 'Erwartete Prov. Whg_Fundline' in comparison_df.columns else 'Erwartete Prov. Whg'
            excel_column = 'Provision_Excel' if 'Provision_Excel' in comparison_df.columns else 'Provision'

            # Ensure the columns being compared have the same type
            comparison_df[fundline_column] = comparison_df[fundline_column].astype(float)
            comparison_df[excel_column] = comparison_df[excel_column].astype(float)

            # Calculate differences and add a new column for the difference
            comparison_df['Difference'] = comparison_df[excel_column] - comparison_df[fundline_column]
            
            # Aggregate results for Quartal
            fundline_quartal_agg = fundline_df.groupby('Isin Code')['Erwartete Prov. Whg'].sum().reset_index()
            excel_quartal_agg = excel_df.groupby('Isin Code')['Provision'].sum().reset_index()
            quartal_aggregated_df = pd.merge(
                fundline_quartal_agg, excel_quartal_agg, 
                on='Isin Code', 
                how='inner', 
                suffixes=('_Fundline', '_Excel')
            )
            quartal_aggregated_df['Difference'] = quartal_aggregated_df['Provision'] - quartal_aggregated_df['Erwartete Prov. Whg']

            # Save each comparison result to a file in memory
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                quartal_aggregated_df.to_excel(writer, sheet_name='Quartal', index=False)
                comparison_df[['Isin Code', 'Date', fundline_column, excel_column, 'Difference']].to_excel(writer, sheet_name='Einzeln', index=False)
            
            # Apply conditional formatting
            apply_conditional_formatting(output, sheet_name='Quartal', column='D', lower_threshold=-20, upper_threshold=20)
            apply_conditional_formatting(output, sheet_name='Einzeln', column='E', lower_threshold=-20, upper_threshold=20)

            # Save the BytesIO object to a file
            output.seek(0)
            comparison_file_name = f"{os.path.splitext(fundline_file)[0]}_{os.path.splitext(excel_file)[0]}_comparison.xlsx"
            comparison_files.append((comparison_file_name, output))

        else:
            st.write(f"Required columns not found in {fundline_file} or {excel_file}")

    return comparison_files

def match_files(fundline_files, excel_files):
    """Match Fundline and Excel files based on similar names using fuzzy matching."""
    matched_files = []
    for fundline_file in fundline_files:
        fundline_base = os.path.splitext(fundline_file)[0].lower()
        best_match, score = process.extractOne(fundline_base, [os.path.splitext(f)[0].lower() for f in excel_files], scorer=fuzz.partial_ratio)
        if score > 80:  # Adjust this threshold based on your needs
            excel_file = next(f for f in excel_files if os.path.splitext(f)[0].lower() == best_match)
            matched_files.append((fundline_file, excel_file))
    return matched_files

# Streamlit UI
st.title("Excel Comparison Tool")

st.sidebar.header("Upload Files")
fundline_files = st.sidebar.file_uploader("Upload Fundline Files", type=['xlsx'], accept_multiple_files=True)
excel_files = st.sidebar.file_uploader("Upload Excel Files", type=['xlsx'], accept_multiple_files=True)

if st.sidebar.button("Compare Files"):
    if fundline_files and excel_files:
        st.write("Processing files...")
        fundline_data = read_files_from_upload(fundline_files)
        excel_data = read_files_from_upload(excel_files)
        
        comparison_files = compare_data(fundline_data, excel_data, column_mappings)
        
        if comparison_files:
            st.write("Comparison complete. Download the files below:")
            for file_name, file in comparison_files:
                st.download_button(label=f"Download {file_name}", data=file, file_name=file_name)
        else:
            st.write("No discrepancies found.")
    else:
        st.write("Please upload both Fundline and Excel files.")
