import streamlit as st
import pandas as pd
import os
from bs4 import BeautifulSoup
from weasyprint import HTML, CSS
from weasyprint.text.fonts import FontConfiguration
import zipfile
from io import BytesIO

#function to merge the data from the excel file into the html template file
def merge_data(data, template):
    soup = template
    # replace inner text of the id "customer_id" with the value from the Excel file
    columns = ['client_name', 'client_name2', 'client_name3', 'fc_name', 'fc_name2', 'unit', 'registered_address_cis',
        'mailing_address_cis', 'current_address_cis', 'nationality',
        'occupation', 'business_type', 'account_no', 'account_name',
        'ipq_result', 'investment_objective', 'client_classification',
        't_and_c', 'director_authorized']
    for i in range(len(data)):
        print(f"Processing {i} of {len(data)}: {data['id'][i]}")
        for column in columns:
            # print(column)
            try:
                soup.find(id=column).string = data[column][i]
            except:
                print(f"Error with {column}")
                pass

        # Convert modified HTML to PDF and save it for user to download from browser
        print(f"Creating PDF for {data['id'][i]}")
        # print(f'soup content: {soup}')
        # html to pdf using weasyprint with A4 size
        font_config = FontConfiguration()
        styles = CSS(string=
            """
                @font-face {
                    font-family: 'SukhumvitSet';  
                    src: url('font/SukhumvitSet-Light.ttf') format('truetype');
                    font-weight: 300;  /* Adjust for light weight */
                }
                @font-face {
                    font-family: 'SukhumvitSet';  
                    src: url('font/SukhumvitSet-Medium.ttf') format('truetype');
                    font-weight: normal; /* or 400 */
                }
                @font-face {
                    font-family: 'SukhumvitSet';  
                    src: url('font/SukhumvitSet-Bold.ttf') format('truetype');
                    font-weight: bold;  /* or 700 */
                }
                @font-face {
                    font-family: 'SukhumvitSet';  
                    src: url('font/SukhumvitSet-Thin.ttf') format('truetype');
                    font-weight: 100;  /* Adjust for thin weight */
                }
            """, 
            font_config=font_config
        )
        HTML(string=str(soup)).write_pdf(f'{data["id"][i]}.pdf', stylesheets=[styles], font_config=font_config)

#function to download the merged files as zip
def download_file(data):
    memory_zip = BytesIO()
    with zipfile.ZipFile(memory_zip, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for i in range(len(data['id'])): 
            filename = f'{data["id"][i]}.pdf'
            zip_file.write(filename)

    # Provide the Download Button
    memory_zip.seek(0)
    st.download_button(
        label="Download Zip File",
        data=memory_zip,
        file_name="merged_files.zip",
        mime='application/zip'
    )

#function to read the data from the excel file
def read_data(file):
    data = pd.read_excel(file, sheet_name='Sheet1')
    return data

#function to read the html template file
def read_template(html):
    
    soup = BeautifulSoup(html, 'lxml')
    return soup
    # with open(html, 'r') as file:
    #     return BeautifulSoup(file.read(), 'lxml') 

st.title('Mail Merge App')
st.markdown('This app lets you upload an Excel file and an HTML template file, merge the data from the Excel file into the HTML template file, and then download the merged file.')

#upload the excel file
st.subheader('Upload the Excel file')
excel_file = st.file_uploader('Upload the Excel file', type=['xlsx'])
if excel_file is not None:
    data = read_data(excel_file)
    data['client_name2'] = data['client_name']
    data['client_name3'] = data['client_name']
    data['fc_name2'] = data['fc_name']
    st.write(data)

#upload the html template file
st.subheader('Upload the HTML template file')
template_file = st.file_uploader('Upload the HTML template file', type=['html'])
if template_file is not None:
    template = read_template(template_file)
    # st.write(template)

#merge the data from the excel file into the html template file
if st.button('Merge Data'):
    merge_data(data, template)
    download_file(data)
