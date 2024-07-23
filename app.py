import streamlit as st
import boto3
import json
import os

# Retrieve AWS credentials and S3 bucket name from environment variables
S3_BUCKET = os.getenv('S3_BUCKET')
AWS_REGION = os.getenv('AWS_DEFAULT_REGION')

# Initialize the S3 client
S3_CLIENT = boto3.client(
    's3',
    region_name=AWS_REGION,
    aws_access_key_id=os.getenv('AWS_ACCESS_KEY_ID'),
    aws_secret_access_key=os.getenv('AWS_SECRET_ACCESS_KEY')
)

def upload_file_to_s3(file, bucket, key):
    try:
        # Print debugging information
        print(f"Uploading file to bucket: {bucket}, key: {key}")
        if file is None:
            raise ValueError("The file object is None")
        S3_CLIENT.upload_fileobj(file, bucket, key)
        print("Upload successful")
    except Exception as e:
        print(f"Error uploading file: {e}")
        raise

def invoke_lambda(fundline_key, excel_key):
    try:
        lambda_client = boto3.client('lambda')
        response = lambda_client.invoke(
            FunctionName='bestandsprovision',
            InvocationType='RequestResponse',
            Payload=json.dumps({
                "bucket": S3_BUCKET,
                "fundline_key": fundline_key,
                "excel_key": excel_key
            })
        )
        return json.loads(response['Payload'].read())
    except Exception as e:
        print(f"Error invoking Lambda function: {e}")
        raise

st.title("Excel Comparison Tool")

fundline_file = st.file_uploader("Upload Fundline File", type=['xlsx'])
excel_file = st.file_uploader("Upload Excel File", type=['xlsx'])

if st.button('Process Files'):
    if fundline_file and excel_file:
        try:
            fundline_key = f"fundline_excel/{fundline_file.name}"
            excel_key = f"excel_excel/{excel_file.name}"

            # Ensure files are correctly read as bytes
            fundline_bytes = fundline_file.read()
            excel_bytes = excel_file.read()

            # Check if the files are correctly loaded
            if fundline_bytes is None:
                raise ValueError("Fundline file bytes are None")
            if excel_bytes is None:
                raise ValueError("Excel file bytes are None")

            # Upload files to S3
            upload_file_to_s3(fundline_bytes, S3_BUCKET, fundline_key)
            upload_file_to_s3(excel_bytes, S3_BUCKET, excel_key)

            # Invoke Lambda function
            result = invoke_lambda(fundline_key, excel_key)

            if result['statusCode'] == 200:
                st.success('Files processed successfully! Check the output folder in your S3 bucket for the results.')
            else:
                st.error('Error processing files!')
        except Exception as e:
            st.error(f"An error occurred: {e}")
    else:
        st.error('Please upload both files!')
