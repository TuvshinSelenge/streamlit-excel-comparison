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
    aws_access_key_id=os.getenv('AKIAQEFWAXCCQ3G46WP4'),
    aws_secret_access_key=os.getenv('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')
)

def upload_file_to_s3(file, bucket, key):
    S3_CLIENT.upload_fileobj(file, bucket, key)

def invoke_lambda(fundline_key, excel_key):
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

st.title("Excel Comparison Tool")

fundline_file = st.file_uploader("Upload Fundline File", type=['xlsx'])
excel_file = st.file_uploader("Upload Excel File", type=['xlsx'])

if st.button('Process Files'):
    if fundline_file and excel_file:
        fundline_key = f"fundline_excel/{fundline_file.name}"
        excel_key = f"excel_excel/{excel_file.name}"

        upload_file_to_s3(fundline_file, S3_BUCKET, fundline_key)
        upload_file_to_s3(excel_file, S3_BUCKET, excel_key)

        result = invoke_lambda(fundline_key, excel_key)

        if result['statusCode'] == 200:
            st.success('Files processed successfully! Check the output folder in your S3 bucket for the results.')
        else:
            st.error('Error processing files!')
    else:
        st.error('Please upload both files!')
