import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

def extract_details_from_pdf(pdffile):
    with pdfplumber.open(pdffile) as pdf:
        extracted_details = []

        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if text:
                text = text.replace(',', '.').lower()

                match_h1 = re.search(r'h\.1 nomor :\s*((?:\d\s*)+)', text)
                specific_text_h1 = match_h1.group(1).replace(" ", "") if match_h1 else "Pattern not found for H.1 NOMOR :"

                match_c2 = re.search(r'c\.2 nama wajib pajak :\s*(.*)', text)
                specific_text_c2 = match_c2.group(1).strip() if match_c2 else "Pattern not found for C.2 Nama Wajib Pajak :"

                match_c3 = re.search(r'c\.1 npwp :\s*(.*)', text)
                specific_text_c3 = match_c3.group(1).strip().replace(" ", "") if match_c3 else "Pattern not found for C.1 NPWP :"

                match_amount1 = re.search(r'(\d{1,3}(?:\.\d{3})*\.\d{1,3})', text)
                specific_amount1 = match_amount1.group(1).replace('.', '.') if match_amount1 else "Amount not found"

                match_doc_ref = re.search(r'b\.7 dokumen referensi : nomor dokumen (\w+)', text)
                specific_doc_ref = match_doc_ref.group(1).strip() if match_doc_ref else "Pattern not found for B.7 Dokumen Referensi : Nomor Dokumen"

                pph = re.search(r'(\d{1,3}(?:\.\d{3})\.\d{1,3}) (\d)(?:\.\d{2})? (\d{1,3}(?:\.\d{3})*\.\d{1,3})', text)

                if pph:
                    match_amount2 = pph
                    specific_amount2 = match_amount2.group(3)
                else:
                    match_amount2 = re.search(r'(\d{1,3}(?:\.\d{3})\.\d{2}) (\d) (\d)(?:\.\d{2})? (\d{1,3}(?:\.\d{3})*\.\d{1,3})', text)
                    if match_amount2:
                        specific_amount2 = match_amount2.group(4)
                    else:
                        specific_amount2 = "Amount not found"

                match_date_1 = re.search(r'c\.3 tanggal : (\d\s\d) dd (\d\s\d) mm (\d\s\d\s\d\s\d)', text)
                if match_date_1:
                    match_date = match_date_1
                    day = match_date.group(1).replace(" ", "")
                    month = match_date.group(2).replace(" ", "")
                    year = match_date.group(3).replace(" ", "")
                    invoice_date = f"{year}-{month}-{day}"
                else:
                    match_date = re.search(r'c\.3 tanggal : (\d\s\d) (\d\s\d) (\d\s\d\s\d\s\d)', text)
                    if match_date:
                        day = match_date.group(1).replace(" ", "")
                        month = match_date.group(2).replace(" ", "")
                        year = match_date.group(3).replace(" ", "")
                        invoice_date = f"{year}-{month}-{day}"
                    else:
                        invoice_date = "Date not found"

                match_date_2 = re.search(r'nama dokumen invoice tanggal (\d\s\d) dd (\d\s\d) mm (\d\s\d\s\d\s\d)', text)
                if match_date_2:
                    match_date_0 = match_date_2
                    day = match_date_0.group(1).replace(" ", "")
                    month = match_date_0.group(2).replace(" ", "")
                    year = match_date_0.group(3).replace(" ", "")
                    invoice_date_2 = f"{year}-{month}-{day}"
                else:
                    match_date_0 = re.search(r'nama dokumen invoice tanggal (\d\s\d) (\d\s\d) (\d\s\d\s\d\s\d)', text)
                    if match_date_0:
                        day = match_date_0.group(1).replace(" ", "")
                        month = match_date_0.group(2).replace(" ", "")
                        year = match_date_0.group(3).replace(" ", "")
                        invoice_date_2 = f"{year}-{month}-{day}"
                    else:
                        invoice_date_2 = "Date not found"

                extracted_details.append({
                    "H.1 NOMOR text": specific_text_h1,
                    "C.2 Nama Wajib Pajak text": specific_text_c2,
                    "C.1 NPWP text": specific_text_c3,
                    "Specific amount1": specific_amount1,
                    "Specific amount2": specific_amount2,
                    "Date": invoice_date,
                    "Tanggal Dokumen": invoice_date_2,
                    "Nomor Dokumen" : specific_doc_ref,
                    "PDF_Name" : pdffile.name
                })
            else:
                extracted_details.append({
                    "H.1 NOMOR text": "Error! Please Check The File",
                    "C.2 Nama Wajib Pajak text": "Error! Please Check The File",
                    "C.1 NPWP text": "Error! Please Check The File",
                    "Specific amount1": "Error! Please Check The File",
                    "Specific amount2": "Error! Please Check The File",
                    "Date": "Error! Please Check The File",
                    "Tanggal Dokumen": "Error! Please Check The File",
                    "Nomor Dokumen": "Error! Please Check The File",
                    "PDF_Name" : pdffile.name
                })

    return extracted_details

# Streamlit app
st.title('OCS - OCR X-Traction')

# File upload
uploaded_files = st.file_uploader("Choose PDF files", type="pdf", accept_multiple_files=True)

if uploaded_files:
    all_extracted_details = []

    for uploaded_file in uploaded_files:
        extracted_details = extract_details_from_pdf(uploaded_file)
        all_extracted_details.extend(extracted_details)

    # Create a DataFrame from the extracted details
    df = pd.DataFrame(all_extracted_details)

    # Adding additional columns
    df['NO'] = df.index + 1
    df['Pasal'] = 23
    df['Pph'] = 24
    df['Alamat'] = "Indonesia"
    df['NTPN'] = 0

    # Renaming columns to match the desired format
    df.rename(columns={
        "C.2 Nama Wajib Pajak text": "Nama_Pemotong",
        "C.1 NPWP text": "NPWP_Pemotong",
        "Specific amount1": "Nilai_Obj_Pemotongan",
        "Specific amount2": "Pph_potput",
        "H.1 NOMOR text": "Nomor_Bukti",
        "Date": "Tanggal",
        "PDF_Name" : "PDF_Name",
        "Tanggal Dokumen": "Tanggal_Dokumen",
        "Nomor Dokumen": "Nomor_Dokumen"
    }, inplace=True)

    # Reordering columns to match the desired format
    df = df[['NO', 'PDF_Name', 'Nama_Pemotong', 'NPWP_Pemotong', 'Pasal', 'Pph', 'Nilai_Obj_Pemotongan', 'Pph_potput', 'Nomor_Bukti', 'Tanggal', 'Tanggal_Dokumen','Nomor_Dokumen', 'Alamat', 'NTPN']]

    # Convert date columns to datetime format
    df['Tanggal'] = pd.to_datetime(df['Tanggal'], format='%Y-%m-%d', errors='coerce').dt.strftime('%Y-%m-%d')
    df['Tanggal_Dokumen'] = pd.to_datetime(df['Tanggal_Dokumen'], format='%Y-%m-%d', errors='coerce').dt.strftime('%Y-%m-%d')
    
    # Convert all string data to uppercase
    df = df.applymap(lambda x: x.upper() if isinstance(x, str) else x)

    def remove_suffix(amount):
        return re.sub(r'\.00$', '', amount)
    
    # Display the DataFrame
    st.write(df)

    # Optionally, you can provide an option to download the DataFrame as a CSV file
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="Download data as CSV",
            data=csv,
            file_name='extracted_details.csv',
            mime='text/csv',
        )

    # Provide an option to download the DataFrame as an Excel file
    with col2:
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            writer.close()
            processed_data = output.getvalue()

        st.download_button(
            label="Download data as Excel",
            data=processed_data,
            file_name='extracted_details.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
