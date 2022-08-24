import pandas as pd
import numpy as np
import streamlit as st
from tabula.io import read_pdf
import os
import glob
import PyPDF2
import re
import fitz
import base64
from io import BytesIO
import timeit

def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter') # pylint: disable=abstract-class-instantiated
    df.to_excel(writer, sheet_name='Sheet1',index=False)
    writer.save()
    processed_data = output.getvalue()
    return processed_data

def get_table_download_link(df):
    """Generates a link allowing the data in a given panda dataframe to be downloaded
    in:  dataframe
    out: href string
    """
    val = to_excel(df)
    b64 = base64.b64encode(val)
    return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="bulk.xlsx">Download Excel file</a>' # decode b'abc' => abc

# Title
st.title('PDF to Excel - Dis-Chem')

# Form
with st.form("Input"):
    # Upload files
    files = st.file_uploader('Upload all orders',accept_multiple_files=True)
    st.write("Est time: ")
    st.write("10 files ~ 10s")
    st.write("50 files ~ 45s")
    st.write("100 files ~ 2m")
    st.write("200 files ~ 4m")
    map_file = st.file_uploader('Retailer Map', type='xlsx')
    notes = st.text_input('Type in your notes here')
    submitted = st.form_submit_button("Submit")

if submitted:

    start = timeit.default_timer()


    # Read data and get address

    # Read data in uploaded files and add filenames
    if files:
        for file in files:
            filenames = [file.name for file in files]
            file.seek(0)
        files_read = [read_pdf(file,pages='all')[0] for file in files]
        for dataframe, filename in zip(files_read,filenames):
            dataframe['filename'] = filename
        
        # Get address
        address_list = []
        for file in files:    
            file.seek(0)
            with fitz.open(stream=file.read(), filetype="pdf") as doc:
                text = ""
                for page in doc:
                    text += page.get_text()
                    first_word = "Address"
                    second_word = "36 SATURN"
                    index1=text.find(first_word)
                    index2=text.find(second_word)
                    addresses = [(text[index1+8:index2-1]) for page in doc]
                    address_list.append(addresses[0])
                    for dataframe, address in zip(files_read,address_list):
                        dataframe['Store Name'] = address
        
        df = pd.concat(files_read)
    
    # Remove blank rows
    df = df.dropna(subset = ["Article No"])
    
    # Tidy data
    df['PO'] = df['filename'].str.replace('.pdf','')
    df['1'] = df['PO'].str.replace('PO - ','')
    # df['1'] = df['PO'].astype(int)
    df['List Cost'] = df['Uom List Cost'].str.split(' ').str[1]
    df['Price'] = df['List Cost'].astype(float)
    df['Notes']  = notes
    df['2'] = ''
    df['3'] = ''
    df["Dischem's Article Code"] = df['Article No'].astype(int)

    # Merge with retailer map

    # Product List
    if map_file:
        df_product_file = pd.read_excel(map_file,"Master Product List")
    df_product_map = df_product_file.astype(str)
    df_product_map["Dischem's Article Code"] = df_product_map["Dischem's Article Code"].astype(int)
    df_merged1 = df.merge(df_product_map,how='left',on="Dischem's Article Code")
    missing_product = df_merged1['SMD Product Code'].isnull()
    df_missing_product = df_merged1[missing_product]
    df_missing_product_list = df_missing_product[["Dischem's Article Code","Description/Vendor Product Code"]]
    df_missing_unique1 = df_missing_product_list.drop_duplicates()
    st.write("The following products are missing the SMD code on the map: ")
    st.table(df_missing_unique1)


    # Store List
    if map_file:
        df_store_file = pd.read_excel(map_file,"Master Store List")
    df_store_map = df_store_file.astype(str)
    df_merged2 = df_merged1.merge(df_store_map,how='left',on="Store Name")
    missing_stores = df_merged2['SMD Store Name'].isnull()
    df_missing_store = df_merged2[missing_stores]
    df_missing_store_list = df_missing_store[["Store Name"]]
    df_missing_unique2 = df_missing_store_list.drop_duplicates()
    st.write("The following stores are missing the SMD code on the map: ")
    st.table(df_missing_unique2)

    df_final = df_merged2[['Notes','Account Number','SMD Store Code','1','2','3','SMD Product Code','SMD Description','Qty','Price']]
    st.dataframe(df_final)


    stop = timeit.default_timer()
    st.write('Time: ', stop - start)  

    # Output to .xlsx
    st.write('Please ensure that no products are missing before downloading!')
    st.markdown(get_table_download_link(df_final), unsafe_allow_html=True)
         
