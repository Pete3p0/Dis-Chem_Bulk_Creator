import json
# from typing import final
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
    st.write("Number of PDF orders uploaded: {}".format(len(files)))
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
    # Read data in uploaded files and add filenames
    if files:
        # st.write("Number of PDF files uploaded: {}".format(len(files)))
        total_num_pages=[]
        address_list = []
        for file in files:
            # st.write("-------------------PDF------------------")
            # st.write("PDF file information: {}".format(file))
            readpdf = PyPDF2.PdfFileReader(file)
            totalpages = readpdf.numPages
            # st.write("Total page number: {}".format(totalpages))
            total_num_pages.append(totalpages)

            ##### Address column population #######
            file.seek(0)
            with fitz.open(stream=file.read(), filetype="pdf") as doc:
                text = ""
                for page in doc:
                    text += page.get_text()
                    first_word = "Address"
                    second_word = "36 SATURN"
                    index1=text.find(first_word)
                    index2=text.find(second_word)
                    addresses = text[index1+8:index2-1]
                    
                    address_list.append(addresses)
                    
                    #for final_dataframe, address in zip(files_read,address_list):
                        #final_dataframe['Store Name'] = address

            for i in range(totalpages):
                filenames = [file.name for file in files]
                file.seek(0)

        # Array holding df's from each pdf
        dfs=[]

        for i in range(len(total_num_pages)):
            # st.write("-------------- DATAFRAME {} -----------------".format(i+1))
            files_read = read_pdf(files[i],pages="all",guess=False,area=[186.6,5.5,367.9,751.0],columns=[257.9,320.1,402.2,447.6,484.5,510.0,572.5,626.3,691.5])
            concat_tables_df = pd.concat(files_read)

            #Add filename column
            concat_tables_df['filename'] = filenames[i]
            #Add address column
            if(total_num_pages[i] < 2):
                concat_tables_df['Store Name'] = address_list[i]
            else:
                concat_tables_df['Store Name'] = address_list[i+1]

            concat_tables_df = concat_tables_df.reset_index(drop=True)
            # st.write(concat_tables_df)

            dfs.append(concat_tables_df)

        final_dataframe = pd.concat(dfs)
        final_dataframe = final_dataframe.reset_index(drop=True)
        # st.write("-------------FINAL DATAFRAME-------------")
        # st.write(final_dataframe)
      
        #df = pd.concat(files_read)
    
    # Remove blank rows
    df = final_dataframe.dropna(subset = ["Article No"])
    
    # Tidy data
    df['PO'] = df['filename'].str.replace('.pdf','')
    df['Order No.'] = df['PO'].str.replace('PO - ','')
    # df['1'] = df['PO'].astype(int)
    # df['List Cost'] = df['Uom List Cost'].str.split(' ').str[1]
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
    df_missing_store_list = df_missing_store[["Store Name","Order No."]]
    df_missing_unique2 = df_missing_store_list.drop_duplicates()
    st.write("The following stores are missing the SMD code on the map: ")
    st.table(df_missing_unique2)

    # df_merged2['Net Value'] = df_merged2['Net Value'].astype(float)
    # st.dataframe(df_merged2)
    
    df_merged2['Total Amount'] = df_merged2['Qty'] * df_merged2['Price']
    st.write('Order value for each store:')
    grouped_df_value = df_merged2.groupby(["Store Name"],as_index=False).agg({"Qty":"sum", "Total Amount":"sum"}).sort_values("Total Amount", ascending=False)
    st.table(grouped_df_value.style.format({'Qty':'{:,.0f}','Total Amount':'R{:,.2f}'}))
    grouped_df_value_less = grouped_df_value[grouped_df_value['Total Amount'] < 1500]
    df_final = df_merged2[['Notes','SMD Store Code','Order No.','2','3','SMD Product Code','SMD Description','Qty','Price']]
    st.write('Final Table')
    st.dataframe(df_final)


    stop = timeit.default_timer()
    st.write('Time: ', stop - start)  

    # Output to .xlsx
    st.write('________________________________________________________________________')
    st.write('Click below to download final table')
    st.markdown(get_table_download_link(df_final), unsafe_allow_html=True)

    # Table less than R1500
    st.write('________________________________________________________________________')
    st.write('Click below to download table with orders less than R1500')
    st.markdown(get_table_download_link(grouped_df_value_less), unsafe_allow_html=True)
         
