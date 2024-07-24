import streamlit as st
import boto3
import json
import os

# Retrieve AWS credentials and S3 bucket name from environment variables
S3_BUCKET = os.getenv('S3_BUCKET')
OUTPUT_BUCKET = "myexcelfund-output"
AWS_REGION = os.getenv('AWS_DEFAULT_REGION')
AWS_ACCESS_KEY_ID = os.getenv('AWS_ACCESS_KEY_ID')
AWS_SECRET_ACCESS_KEY = os.getenv('AWS_SECRET_ACCESS_KEY')

# Debugging: Print environment variable values to ensure they are set
print(f"S3_BUCKET: {S3_BUCKET}")
print(f"OUTPUT_BUCKET: {OUTPUT_BUCKET}")
print(f"AWS_REGION: {AWS_REGION}")
print(f"AWS_ACCESS_KEY_ID: {AWS_ACCESS_KEY_ID}")
print(f"AWS_SECRET_ACCESS_KEY: {AWS_SECRET_ACCESS_KEY}")

# Ensure none of the environment variables are None
if None in (S3_BUCKET, AWS_REGION, AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY):
    st.error("One or more environment variables are not set. Please check the configuration in Streamlit Cloud.")
    st.stop()

# Initialize the S3 client
S3_CLIENT = boto3.client(
    's3',
    region_name=AWS_REGION,
    aws_access_key_id=AWS_ACCESS_KEY_ID,
    aws_secret_access_key=AWS_SECRET_ACCESS_KEY
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
        lambda_client = boto3.client('lambda', region_name=AWS_REGION,
                                     aws_access_key_id=AWS_ACCESS_KEY_ID,
                                     aws_secret_access_key=AWS_SECRET_ACCESS_KEY)
        response = lambda_client.invoke(
            FunctionName='bestandsprovision2',
            InvocationType='RequestResponse',
            Payload=json.dumps({
                "bucket": S3_BUCKET,
                "fundline_key": fundline_key,
                "excel_key": excel_key
            })
        )
        response_payload = response['Payload'].read()
        print(f"Lambda response payload: {response_payload}")
        return json.loads(response_payload)
    except Exception as e:
        print(f"Error invoking Lambda function: {e}")
        raise

def download_file_from_s3(bucket, key):
    try:
        file_obj = S3_CLIENT.get_object(Bucket=bucket, Key=key)
        return file_obj['Body'].read()
    except Exception as e:
        print(f"Error downloading file from S3: {e}")
        return None

def list_s3_objects(bucket, prefix):
    try:
        response = S3_CLIENT.list_objects_v2(Bucket=bucket, Prefix=prefix)
        if 'Contents' in response:
            return [obj['Key'] for obj in response['Contents']]
        else:
            return []
    except Exception as e:
        print(f"Error listing objects in S3: {e}")
        return []

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
