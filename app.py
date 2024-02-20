import streamlit as st
import pandas as pd
import os
from bs4 import BeautifulSoup
from weasyprint import HTML, CSS
from weasyprint.text.fonts import FontConfiguration
import zipfile
from io import BytesIO
import subprocess
from pikepdf import Pdf, Permissions, Encryption
#function to merge the data from the excel file into the html template file
def merge_data(data, template):
    soup = template
    # replace inner text of the id "customer_id" with the value from the Excel file
    columns = ['client_name', 'client_name2', 'client_name3', 'fc_name', 'fc_name2', 'unit', 'registered_address_cis',
        'mailing_address_cis', 'current_address_cis', 'nationality',
        'occupation', 'business_type', 'account_no', 'account_name',
        'ipq_result', 'investment_objective', 'client_classification',
        't_and_c', 'director_authorized']
    
    font_config = FontConfiguration()
    styles = CSS(string=
        """
            @font-face {
                font-family: 'SukhumvitSet';  
                src: url('https://github.com/bluenex/baansuan_prannok/raw/master/fonts/sukhumvit-set/SukhumvitSet-Light.ttf') format('truetype');
                font-weight: 300;  /* Adjust for light weight */
            }
            @font-face {
                font-family: 'SukhumvitSet';  
                src: url('https://github.com/bluenex/baansuan_prannok/raw/master/fonts/sukhumvit-set/SukhumvitSet-Medium.ttf') format('truetype');
                font-weight: normal; /* or 400 */
            }
            @font-face {
                font-family: 'SukhumvitSet';  
                src: url('https://github.com/bluenex/baansuan_prannok/raw/master/fonts/sukhumvit-set/SukhumvitSet-Bold.ttf') format('truetype');
                font-weight: bold;  /* or 700 */
            }
            @font-face {
                font-family: 'SukhumvitSet';  
                src: url('https://github.com/bluenex/baansuan_prannok/raw/master/fonts/sukhumvit-set/SukhumvitSet-Thin.ttf') format('truetype');
                font-weight: 100;  /* Adjust for thin weight */
            }
        """, 
        font_config=font_config
    )

   
    memory_zip = BytesIO()
    with zipfile.ZipFile(memory_zip, 'w', zipfile.ZIP_DEFLATED) as zip_file:
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
            pdf = HTML(string=str(soup)).write_pdf(
                stylesheets=[styles],
                zoom=1,
                font_config=font_config)
            # encrypt with password 1234 using pikepdf
            with open('temp.pdf', 'wb') as file:
                file.write(pdf)

            pdf = Pdf.open('temp.pdf')
            password = data["pid_last4"][i]
            pdf.save(f'{data["id"][i]}.pdf', encryption=Encryption(user=password, owner=password, R=4))
            zip_file.write(f'{data["id"][i]}.pdf')
            os.remove('temp.pdf')
            os.remove(f'{data["id"][i]}.pdf')

            # encrypt with password 1234 using qpdf
            # with open('temp.pdf', 'wb') as file:
            #     file.write(pdf)
            # subprocess.run(['qpdf', '--encrypt', '', '1234', '256', '--', 'temp.pdf', f'{data["id"][i]}.pdf'])
            # zip_file.write(f'{data["id"][i]}.pdf')
            # os.remove('temp.pdf')
            # filename = f'{data["id"][i]}.pdf'
            # zip_file.writestr(filename, pdf)
        
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
st.markdown("""
    <style>
        @font-face {
            font-family: 'SukhumvitSet';  
            src: url('https://github.com/bluenex/baansuan_prannok/raw/master/fonts/sukhumvit-set/SukhumvitSet-Text.ttf') format('truetype');
            font-weight: normal; 
        }
        html, body, [class*="css"]  {
            font-family: 'SukhumvitSet' !important;
        }
    </style>""", unsafe_allow_html=True)
st.markdown('This app lets you upload an Excel file and an HTML template file, merge the data from the Excel file into the HTML template file, and then download the merged file.')
st.markdown("ภาษาไทย")
#upload the excel file
st.subheader('Upload the Excel file')
excel_file = st.file_uploader('Upload the Excel file', type=['xlsx'])
if excel_file is not None:
    data = read_data(excel_file)
    data['client_name2'] = data['client_name']
    data['client_name3'] = data['client_name']
    data['fc_name2'] = data['fc_name']
    data['pid_last4'] = data['pid'].astype(str).str[-4:]
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