import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule
from io import BytesIO
import logging
from fuzzywuzzy import fuzz, process

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s %(message)s')

# Define column mappings
column_mappings = {
    'Isin Code': [
        'ISIN', 'Isin', 'Share ISIN Reference', 'FINANZINSTRUMENT_IDENT', 'Text23'
    ],
    'Provision': [
        'Comm. Amount', 'Provision', 'Betrag (€)', 'Vergütung', 'EURMonat',
        'Client Trailer Fees Amount In Consolidated Currency', 'Amount In Agreement Ccy', 'Bepro',
        'Provisionsbetrag in Währung', 'Fee', 'BPROV', 'BpkEUR', 'Kommissionsbetrag', 'Commission Due Payment CCY', 'Fee (Payment Currency)', 'Amount In Partner Currency',
        'Betrag (EUR)'
    ],
    'Date': [
        'Date', 'Datum', 'Booking Date', 'End-Datum', 'Period End Date', 'Holding as of', 'STICHTAG', 'Period', 'Datum_str', 'Positionsdatum', 'Stichtag', 'Retrocession Date'
    ]
}

# Functions for processing data
def date_converter(df):
    if "Statement Month" in df.columns and "Statement Year" in df.columns:
        df["Date"] = pd.to_datetime(df["Statement Month"].astype(str) + '-' + df["Statement Year"].astype(str), format='%m-%Y')
    return df

def set_correct_headers(df, column_mappings):
    for i, row in df.iterrows():
        if any(header in row.values for headers in column_mappings.values() for header in headers):
            df.columns = df.iloc[i]
            df = df.drop(i).reset_index(drop=True)
            return df
    return df

def rename_columns(df, column_mappings):
    rename_dict = {}
    for new_name, old_names in column_mappings.items():
        for old_name in old_names:
            if old_name in df.columns:
                rename_dict[old_name] = new_name
    df.rename(columns=rename_dict, inplace=True)
    return df

def filter_valid_rows(df):
    if 'Isin Code' in df.columns and 'Date' in df.columns:
        df = df.dropna(subset=['Isin Code', 'Date'])
        df = df[df['Isin Code'].apply(lambda x: isinstance(x, str) and x.strip() != "")]
    return df

def convert_date_column(df, date_column_name):
    if date_column_name in df.columns:
        df[date_column_name] = pd.to_datetime(df[date_column_name], errors='coerce', format='%d.%m.%Y')
    return df

def aggregate_data(data, value_column):
    aggregated_data = {}
    for file_name, df in data.items():
        df = rename_columns(df, column_mappings)
        df = filter_valid_rows(df)
        logging.info(f"\nColumns in {file_name} after renaming and filtering: {df.columns}")
        if 'Isin Code' in df.columns and value_column in df.columns:
            aggregated_df = df.groupby('Isin Code')[value_column].sum().reset_index()
            aggregated_data[file_name] = aggregated_df
            logging.info(f"Aggregated {file_name} successfully with shape {aggregated_df.shape}")
        else:
            logging.info(f"Required columns not found in {file_name}")
    return aggregated_data

def match_files(fundline_files, excel_files):
    matched_files = []
    for fundline_file in fundline_files:
        fundline_base = os.path.splitext(fundline_file)[0].lower()
        best_match, score = process.extractOne(fundline_base, [os.path.splitext(f)[0].lower() for f in excel_files], scorer=fuzz.partial_ratio)
        if score > 80:  # Adjust this threshold based on your needs
            excel_file = next(f for f in excel_files if os.path.splitext(f)[0].lower() == best_match)
            matched_files.append((fundline_file, excel_file))
    logging.info("Matched files: %s", matched_files)
    return matched_files

def apply_conditional_formatting(output, sheet_name='Sheet2', column='A', lower_threshold=None, upper_threshold=None):
    try:
        output.seek(0)
        wb = load_workbook(output)
        ws = wb[sheet_name]

        fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        end_row = ws.max_row
        if end_row < 2:
            return

        cell_range = f'{column}2:{column}{end_row}'

        if lower_threshold is not None:
            lower_rule = CellIsRule(operator='lessThan', formula=[str(lower_threshold)], fill=fill)
            ws.conditional_formatting.add(cell_range, lower_rule)

        if upper_threshold is not None:
            upper_rule = CellIsRule(operator='greaterThan', formula=[str(upper_threshold)], fill=fill)
            ws.conditional_formatting.add(cell_range, upper_rule)

        output.seek(0)
        wb.save(output)
        output.seek(0)

    except Exception as e:
        logging.error(f"Error applying conditional formatting: {e}")
        raise

def compare_data(fundline_data, excel_data):
    comparison_files = []
    fundline_files = fundline_data.keys()
    excel_files = excel_data.keys()
    matched_files = match_files(fundline_files, excel_files)

    for fundline_file, excel_file in matched_files:
        fundline_df = fundline_data[fundline_file]
        excel_df = excel_data[excel_file]

        fundline_df = rename_columns(fundline_df, column_mappings)
        excel_df = rename_columns(excel_df, column_mappings)

        fundline_df = convert_date_column(fundline_df, 'Date')
        excel_df = convert_date_column(excel_df, 'Date')

        logging.info(f"\nColumns in {fundline_file} after renaming: {fundline_df.columns}")
        logging.info(f"Columns in {excel_file} after renaming: {excel_df.columns}")

        if 'Isin Code' in fundline_df.columns and 'Date' in fundline_df.columns:
            fundline_df = fundline_df.groupby(['Isin Code', 'Date'])['Provision'].sum().reset_index()
            logging.info(f"Fundline DataFrame grouped: {fundline_df.head()}")
        else:
            logging.error(f"Missing 'Isin Code' or 'Date' columns in fundline data from {fundline_file}")
            continue

        if 'Isin Code' in excel_df.columns and 'Date' in excel_df.columns:
            excel_df = excel_df.groupby(['Isin Code', 'Date'])['Provision'].sum().reset_index()
            logging.info(f"Excel DataFrame grouped: {excel_df.head()}")
        else:
            logging.error(f"Missing 'Isin Code' or 'Date' columns in excel data from {excel_file}")
            continue

        comparison_df = pd.merge(
            fundline_df, excel_df,
            left_on=['Isin Code', 'Date'],
            right_on=['Isin Code', 'Date'],
            how='inner',
            suffixes=('_Fundline', '_Excel')
        )

        fundline_column = 'Provision_Fundline' if 'Provision_Fundline' in comparison_df.columns else 'Provision'
        excel_column = 'Provision_Excel' if 'Provision_Excel' in comparison_df.columns else 'Provision'

        comparison_df[fundline_column] = comparison_df[fundline_column].astype(float)
        comparison_df[excel_column] = comparison_df[excel_column].astype(float)

        comparison_df['Difference'] = comparison_df[excel_column] - comparison_df[fundline_column]

        fundline_quartal_agg = fundline_df.groupby('Isin Code')['Provision'].sum().reset_index()
        excel_quartal_agg = excel_df.groupby('Isin Code')['Provision'].sum().reset_index()
        quartal_aggregated_df = pd.merge(
            fundline_quartal_agg, excel_quartal_agg,
            on='Isin Code',
            how='inner',
            suffixes=('_Fundline', '_Excel')
        )
        quartal_aggregated_df['Difference'] = quartal_aggregated_df['Provision_Excel'] - quartal_aggregated_df['Provision_Fundline']

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            quartal_aggregated_df.to_excel(writer, sheet_name='Quartal', index=False)
            comparison_df[['Isin Code', 'Date', fundline_column, excel_column, 'Difference']].to_excel(writer, sheet_name='Einzeln', index=False)

        apply_conditional_formatting(output, sheet_name='Quartal', column='F', lower_threshold=-20, upper_threshold=20)
        apply_conditional_formatting(output, sheet_name='Einzeln', column='E', lower_threshold=-20, upper_threshold=20)

        comparison_files.append(output)

    return comparison_files

# Streamlit UI
st.title("Fundline vs Excel Comparison")

uploaded_fundline_files = st.file_uploader("Upload Fundline Files", type=["xlsx"], accept_multiple_files=True)
uploaded_excel_files = st.file_uploader("Upload Excel Files", type=["xlsx"], accept_multiple_files=True)

if uploaded_fundline_files and uploaded_excel_files:
    fundline_data = {file.name: pd.read_excel(file, engine='openpyxl') for file in uploaded_fundline_files}
    excel_data = {file.name: pd.read_excel(file, engine='openpyxl') for file in uploaded_excel_files}

    comparison_files = compare_data(fundline_data, excel_data)

    for i, file in enumerate(comparison_files):
        st.download_button(
            label=f"Download Comparison File {i+1}",
            data=file.getvalue(),
            file_name=f"comparison_{i+1}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
