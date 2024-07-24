import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from fuzzywuzzy import fuzz, process
import io
import os

# Define your column mappings
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

def read_files(files):
    files_data = {}
    for file_name, file_content in files.items():
        try:
            excel_file = pd.ExcelFile(io.BytesIO(file_content))
            all_sheets_df = pd.DataFrame()
            header_set = False
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(io.BytesIO(file_content), sheet_name=sheet_name, header=None)
                df = date_converter(df)
                if not header_set:
                    df = set_correct_headers(df, column_mappings)
                    header_set = True
                else:
                    df.columns = all_sheets_df.columns  # Set same columns as the first sheet
                all_sheets_df = pd.concat([all_sheets_df, df], ignore_index=True)
            files_data[file_name] = all_sheets_df
        except Exception as e:
            st.error(f"Error reading {file_name}: {e}")
    return files_data

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

def convert_date_column(df, date_column_name):
    if date_column_name in df.columns:
        df[date_column_name] = pd.to_datetime(df[date_column_name], errors='coerce', format='%d.%m.%Y')
    return df

def compare_data(fundline_data, excel_data):
    comparison_files = {}

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

        fundline_df = fundline_df.groupby(['Isin Code', 'Date'])['Erwartete Prov. Whg'].sum().reset_index()
        excel_df = excel_df.groupby(['Isin Code', 'Date'])['Provision'].sum().reset_index()

        comparison_df = pd.merge(
            fundline_df, excel_df, 
            left_on=['Isin Code', 'Date'], 
            right_on=['Isin Code', 'Date'], 
            how='inner', 
            suffixes=('_Fundline', '_Excel')
        )

        fundline_column = 'Erwartete Prov. Whg_Fundline' if 'Erwartete Prov. Whg_Fundline' in comparison_df.columns else 'Erwartete Prov. Whg'
        excel_column = 'Provision_Excel' if 'Provision_Excel' in comparison_df.columns else 'Provision'

        comparison_df[fundline_column] = comparison_df[fundline_column].astype(float)
        comparison_df[excel_column] = comparison_df[excel_column].astype(float)

        comparison_df['Difference'] = comparison_df[excel_column] - comparison_df[fundline_column]

        fundline_quartal_agg = fundline_df.groupby('Isin Code')['Erwartete Prov. Whg'].sum().reset_index()
        excel_quartal_agg = excel_df.groupby('Isin Code')['Provision'].sum().reset_index()
        quartal_aggregated_df = pd.merge(
            fundline_quartal_agg, excel_quartal_agg, 
            on='Isin Code', 
            how='inner', 
            suffixes=('_Fundline', '_Excel')
        )
        quartal_aggregated_df['Difference'] = quartal_aggregated_df['Provision'] - quartal_aggregated_df['Erwartete Prov. Whg']

        comparison_file_name = f"{os.path.splitext(fundline_file)[0]}_{os.path.splitext(excel_file)[0]}_comparison.xlsx"
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            quartal_aggregated_df.to_excel(writer, sheet_name='Quartal', index=False)
            comparison_df[['Isin Code', 'Date', fundline_column, excel_column, 'Difference']].to_excel(writer, sheet_name='Einzeln', index=False)
        output.seek(0)
        comparison_files[comparison_file_name] = output.read()

    return comparison_files

def match_files(fundline_files, excel_files):
    matched_files = []
    for fundline_file in fundline_files:
        fundline_base = os.path.splitext(fundline_file)[0].lower()
        best_match, score = process.extractOne(fundline_base, [os.path.splitext(f)[0].lower() for f in excel_files], scorer=fuzz.partial_ratio)
        if score > 80:
            excel_file = next(f for f in excel_files if os.path.splitext(f)[0].lower() == best_match)
            matched_files.append((fundline_file, excel_file))
    return matched_files

# Streamlit UI
st.title("Excel Comparison Tool")

fundline_file = st.file_uploader("Upload Fundline File", type=['xlsx'])
excel_file = st.file_uploader("Upload Excel File", type=['xlsx'])

if st.button('Process Files'):
    if fundline_file and excel_file:
        try:
            fundline_key = f"fundline_excel/{fundline_file.name}"
            excel_key = f"excel_excel/{excel_file.name}"

            # Upload files to S3 directly using the file-like object
            upload_file_to_s3(fundline_file, S3_BUCKET, fundline_key)
            upload_file_to_s3(excel_file, S3_BUCKET, excel_key)

            # Invoke Lambda function
            result = invoke_lambda(fundline_key, excel_key)
            print(f"Lambda result: {result}")

            if 'statusCode' in result and result['statusCode'] == 200:
                st.success('Files processed successfully! Check the output folder in your S3 bucket for the results.')
                
                # List objects in the output bucket
                output_files = list_s3_objects(OUTPUT_BUCKET, "output/")
                st.write(f"Files in output bucket: {output_files}")

                # Download the comparison file
                comparison_key = f"output/{os.path.splitext(fundline_file.name)[0]}_{os.path.splitext(excel_file.name)[0]}_comparison.xlsx"
                comparison_file = download_file_from_s3(OUTPUT_BUCKET, comparison_key)

                if comparison_file:
                    st.download_button(
                        label="Download comparison file",
                        data=comparison_file,
                        file_name=f"{os.path.splitext(fundline_file.name)[0]}_{os.path.splitext(excel_file.name)[0]}_comparison.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error('Failed to download the comparison file from S3.')
            else:
                st.error(f"Error processing files! Lambda returned: {result}")
        except Exception as e:
            st.error(f"An error occurred: {e}")
    else:
        st.error('Please upload both files!')
