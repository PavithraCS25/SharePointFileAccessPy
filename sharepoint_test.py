# ---
# jupyter:
#   jupytext:
#     text_representation:
#       extension: .py
#       format_name: light
#       format_version: '1.5'
#       jupytext_version: 1.16.1
#   kernelspec:
#     display_name: Python 3 (Local)
#     language: python
#     name: python3
# ---

# !pip install Office365-REST-Python-Client

# Import Libraries
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import pandas as pd
import io
import os
from io import BytesIO

# Sharepoint AD Application Details
site_url =  'https://<your-tenant>.sharepoint.com/sites/<your-site>'
ip_file_relative_path = '/sites/<your-site>/Shared%20Documents/input'
op_folder_relative_path = '/sites/<your-site>/Shared%20Documents/output'
ip_filename = 'sample_data.csv'
op_filename = 'output_'+ip_filename
client_id = "client-id" #replace with the client id generated in appregnew.aspx
client_secret = "client-secret-key" #replace with the client secret generated in appregnew.aspx


# +
#Function get context of the sharepoint using credentials
def get_sharepoint_context():
    credentials = ClientCredential(client_id, client_secret)
    ctx = ClientContext(site_url).with_credentials(credentials)
    return ctx

#Function to read file from sharepoint
def read_file_from_sp(ctx,ip_file_relative_path,ip_filename):
    file_relative_path = ip_file_relative_path+'/'+ip_filename
    response = File.open_binary(ctx, file_relative_path)
    df = pd.read_csv(io.BytesIO(response.content))
    return df

#Function to write file to sharepoint
def write_file_to_sp(ctx,df,op_folder_relative_path,op_filename):
    target_folder = ctx.web.get_folder_by_server_relative_url(op_folder_relative_path)
    # Create a buffer object
    buffer = BytesIO()

    # Write the dataframe to the buffer
    df.to_csv(buffer, index=False)
    buffer.seek(0)
    file_content = buffer.read()
    # Upload the file
    target_file = target_folder.upload_file(op_filename, file_content).execute_query()
    return 'successfully uploaded file in sharepoint output folder.'

#Function to work on the processing logic
def process_df(df):
    df['new'] = 'new column'
    return df


# +
def main():
    ctx = get_sharepoint_context()
    in_df = read_file_from_sp(ctx,ip_file_relative_path,ip_filename)
    updated_df = process_df(in_df)
    print(write_file_to_sp(ctx,updated_df,op_folder_relative_path,op_filename))
    print('process completed')

if __name__ == "__main__":
    main()







