import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO
import re
from azure.core.credentials import AzureKeyCredential
from azure.ai.documentintelligence import DocumentIntelligenceClient
from azure.ai.documentintelligence.models import AnalyzeDocumentRequest
from streamlit_option_menu import option_menu
import time
from azure.core.pipeline.transport import HttpRequest, HttpResponse
from azure.core.exceptions import HttpResponseError



st.set_page_config(initial_sidebar_state="collapsed")

# Top navigation bar
page = option_menu(
    menu_title=None,
    options=["PPH - OLD", "PPH - 2025"],
    icons=["clock-history", "calendar3"],
    orientation="horizontal",
    styles={
        "container": {"padding": "0!important", "background-color": "#293771"},
        "icon": {"color": "white", "font-size": "18px"},
        "nav-link": {
            "font-size": "16px",
            "text-align": "center",
            "margin": "0px",
            "color": "white",
            "border-radius": "0.5rem",
            "padding": "0.4375rem 0.625rem",
        },
        "nav-link-selected": {"background-color": "rgba(255, 255, 255, 0.25)"},
    }
)

if page == "PPH - OLD":
    # st.write("This is the PPH - OLD page. Here you can add the specific job or content related to PPH - OLD.")
    # Add your code for PPH - OLD job here
    def extract_details_from_pdf(pdffile):
        with pdfplumber.open(pdffile) as pdf:
            extracted_details = []

            for i, page in enumerate(pdf.pages):
                text = page.extract_text()
                st.write(text)
                if text:
                    text = text.replace(',', '.').lower()

                    nama_a = re.search(r"a\.(.*) nama :\s*(.*)", text)
                    v_namaa = nama_a.group(2) if nama_a else "Data Not Found"

                    nitku_a = re.search(r"a\.(.*) nitku :\s*(.*)", text)
                    # st.write(nitku_a)
                    v_nitkua = nitku_a.group(2) if nitku_a else "Data Not Found"

                    nitku_a1 = re.findall(r"a\.(.*) nik :\s*(.*)", text)
                    nik_a = re.search(r"a\.(.*) nik :\s*(.*)", text)
                    v_nika = nik_a.group(2)
                    # st.write(v_nika)

                    if v_nika == nitku_a1[0][1]:
                        v_nika = "Data Not Found"
                    else:
                        nik_a = re.search(r"a\.\d+ nik :\s*(\S+)", text)
                        v_nika = nik_a.group(1)

                    # if  nik_a1:
                    #     nik_a = nik_a1
                    #     v_nika = "No Data Found"
                    # else:
                    #     nik_a = re.search(r"a\.(.*) nik :\s*(.*) a\.(.*) nitku :\s*(.*)", text)
                    #     v_nika = nik_a.group(2)

                    match_h1 = re.search(r'h\.1 nomor :\s*((?:\d\s*)+)', text)
                    specific_text_h1 = match_h1.group(1).replace(" ", "") if match_h1 else "Pattern not found for H.1 NOMOR :"

                    match_c2 = re.search(r'c\.2 nama wajib pajak :\s*(.*)', text)
                    specific_text_c2 = match_c2.group(1).strip() if match_c2 else "Pattern not found for C.2 Nama Wajib Pajak :"

                    match_c3 = re.search(r'c\.1 npwp :\s*(.*)', text)
                    specific_text_c3 = match_c3.group(1).strip().replace(" ", "") if match_c3 else "Pattern not found for C.1 NPWP :"

                    match_amount1 = re.search(r'(\d{1,3}(?:\.\d{3})*\.\d{1,3})', text)
                    specific_amount1 = match_amount1.group(1).replace('.', '.') if match_amount1 else "Amount not found"

                    facture_num = re.search(r'nomor faktur pajak : (.*) tanggal', text)
                    facture_num_data = facture_num.group(1).replace('.', '.') if facture_num else "Amount not found"

                    match_doc_ref = re.search(r'b\.7 dokumen referensi : nomor dokumen (.*)', text)
                    specific_doc_ref = match_doc_ref.group(1).strip() if match_doc_ref else "Data not found"

                    pph = re.search(r'(\d{1,2}-\d{4}) (\d{2}-\d{3}-\d{2}) (.*) (.*) (.*)', text)
                    # pph = re.search(r'(\d{1,2}-\d{4}) (\d{2}-\d{3}-\d{2}) (\d{1,3}(?:\.\d{3}){1,2}(?:\.\d{2,3})?) (.*) (\d{1,3}(?:\.\d{3}){1,2}(?:\.\d{2,3})?)', text)
                    # st.write("PPH : ",pph)
                    # st.write("Text : ",text)
                    tarif_tinggi = 0
                    if pph:
                        match_amount2 = pph
                        specific_amount2 = match_amount2.group(5)
                    else:
                        match_amount2 = re.search(r'(\d{1,2}-\d{4}) (\d{2}-\d{3}-\d{2}) (.*) (.*) (.*) (.*)', text)
                        # match_amount2 = re.search(r'(\d{1,2}-\d{4}) (\d{2}-\d{3}-\d{2}) (\d{1,3}(?:\.\d{3}){1,2}(?:\.\d{2,3})?) (\d) (\d{1,2}\.\d{2}) (\d{1,3}(?:\.\d{3}){1,2}(?:\.\d{2,3})?)', text)
                        if match_amount2:
                            specific_amount2 = match_amount2.group(6)
                            tarif_tinggi = match_amount2.group(4)
                        else:
                            specific_amount2 = "Amount not found"
                    # st.write(specific_amount2)
                    # st.write(match_amount2)

                    match_date_1 = re.search(r'c\.\d+ tanggal : (\d\s\d) dd (\d\s\d) mm (\d\s\d\s\d\s\d)', text)
                    if match_date_1:
                        match_date = match_date_1
                        day = match_date.group(1).replace(" ", "")
                        month = match_date.group(2).replace(" ", "")
                        year = match_date.group(3).replace(" ", "")
                        invoice_date = f"{year}-{month}-{day}"
                    else:
                        match_date = re.search(r'c\.\d+ tanggal : (\d\s\d) (\d\s\d) (\d\s\d\s\d\s\d)', text)
                        if match_date:
                            day = match_date.group(1).replace(" ", "")
                            month = match_date.group(2).replace(" ", "")
                            year = match_date.group(3).replace(" ", "")
                            invoice_date = f"{year}-{month}-{day}"
                        else:
                            invoice_date = "Date not found"

                    match_date_2 = re.search(r'nama dokumen (.*) tanggal (\d\s\d) dd (\d\s\d) mm (\d\s\d\s\d\s\d)', text)
                    nama_dokumen = ''
                    if match_date_2:
                        match_date_0 = match_date_2
                        nama_dokumen = match_date_0.group(1).replace(" ", "")
                        day = match_date_0.group(2).replace(" ", "")
                        month = match_date_0.group(3).replace(" ", "")
                        year = match_date_0.group(4).replace(" ", "")
                        invoice_date_2 = f"{year}-{month}-{day}"
                        
                    else:
                        match_date_0 = re.search(r'nama dokumen (.*) tanggal (\d\s\d) (\d\s\d) (\d\s\d\s\d\s\d)', text)
                        if match_date_0:
                            day = match_date_0.group(2).replace(" ", "")
                            month = match_date_0.group(3).replace(" ", "")
                            year = match_date_0.group(4).replace(" ", "")
                            invoice_date_2 = f"{year}-{month}-{day}"
                            nama_dokumen = match_date_0.group(1).replace(" ", "")
                        else:
                            invoice_date_2 = "Date not found"
                    
                    if not nama_dokumen:
                        nama_dokumen = 'Data not found'
                    else: 
                        nama_dokumen = nama_dokumen 

                    if invoice_date_2 == "Date not found":
                        match_date_2 = re.search(r'nomor : tanggal (\d\s\d) dd (\d\s\d) mm (\d\s\d\s\d\s\d)', text)
                        if match_date_2:
                            match_date_0 = match_date_2
                            day = match_date_0.group(1).replace(" ", "")
                            month = match_date_0.group(2).replace(" ", "")
                            year = match_date_0.group(3).replace(" ", "")
                            invoice_date_2 = f"{year}-{month}-{day}"
                        else:
                            match_date_0 = re.search(r'nomor : tanggal (\d\s\d) (\d\s\d) (\d\s\d\s\d\s\d)', text)
                            if match_date_0:
                                day = match_date_0.group(1).replace(" ", "")
                                month = match_date_0.group(2).replace(" ", "")
                                year = match_date_0.group(3).replace(" ", "")
                                invoice_date_2 = f"{year}-{month}-{day}"
                            else:
                                invoice_date_2 = "Date not found1"
                    if invoice_date_2 == "Date not found1":
                        match_date_2 = re.search(re.escape(facture_num_data) + r' tanggal (\d\s\d) dd (\d\s\d) mm (\d\s\d\s\d\s\d)', text)
                        if match_date_2:
                            match_date_0 = match_date_2
                            day = match_date_0.group(1).replace(" ", "")
                            month = match_date_0.group(2).replace(" ", "")
                            year = match_date_0.group(3).replace(" ", "")
                            invoice_date_2 = f"{year}-{month}-{day}"
                        else:
                            match_date_0 = re.search(re.escape(facture_num_data) + r' tanggal (\d\s\d) (\d\s\d) (\d\s\d\s\d\s\d)', text)
                            if match_date_0:
                                day = match_date_0.group(1).replace(" ", "")
                                month = match_date_0.group(2).replace(" ", "")
                                year = match_date_0.group(3).replace(" ", "")
                                invoice_date_2 = f"{year}-{month}-{day}"
                            else:
                                invoice_date_2 = "Date not found2"

                    stat_pembetulan1 = re.search(r'h\.2 (.*) pembetulan (.*) (.*) h\.3 pembatalan h\.5 (.*) pph tidak final', text)
                    status_pembatalan = ''
                    nilai_pembetulan = 0
                    status_pph_tidakfinal = ''
                    status_pembetulan = ''
                    # st.write(i)
                    # st.write(text)
                    if stat_pembetulan1:
                        stat_pembetulan = stat_pembetulan1
                        # st.write(stat_pembetulan)
                        status_pembetulan = stat_pembetulan.group(1).replace(' ','')
                        nilai_pembetulan = stat_pembetulan.group(3).replace(' ','')
                        status_pph_tidakfinal = stat_pembetulan.group(4).replace(' ','')
                        # st.write(status_pembetulan,nilai_pembetulan,status_pph_tidakfinal)
                    else:
                        stat_pembetulan = re.search(r'h\.2 (.*) pembetulan (.*) (.*) h\.3 (.*) pembatalan h\.5 (.*) pph tidak final final', text)
                        status_pembetulan = stat_pembetulan.group(1).replace(' ','')
                        nilai_pembetulan = stat_pembetulan.group(3).replace(' ','')
                        status_pembatalan = stat_pembetulan.group(4).replace(' ','')
                        status_pph_tidakfinal = stat_pembetulan.group(5).replace(' ','')


                    if status_pembetulan == 'x' or status_pembetulan == 'X':
                        status_pembetulan = 'Yes'
                    else: status_pembetulan = 'No'

                    if status_pembatalan == 'x' or status_pembatalan == 'X':
                        status_pembatalan = 'Yes'
                    else: status_pembatalan = 'No'

                    if status_pph_tidakfinal == 'x' or status_pph_tidakfinal == 'X':
                        status_pph_tidakfinal = 'Yes'
                    else: status_pph_tidakfinal = 'No'

                    pph_final1 = re.search(r'h\.4 pph final', text)
                    status_pph = 'No'
                    if pph_final1:
                        pph_final = pph_final1
                        status_pph = pph_final
                    else:
                        pph_final = re.search(r'h\.4 (.*) pph final',text)
                        status_pph = pph_final.group(1).replace(' ','')
                    if status_pph == 'x' or status_pph == 'X':
                        status_pph = 'Yes'
                    else: status_pph = 'No'

                    extracted_details.append({
                        "H.1 NOMOR text": specific_text_h1,
                        "C.2 Nama Wajib Pajak text": specific_text_c2,
                        "C.1 NPWP text": specific_text_c3,
                        "Specific amount1": specific_amount1,
                        "Specific amount2": specific_amount2,
                        "Date": invoice_date,
                        "Tanggal Dokumen": invoice_date_2,
                        "Nomor Dokumen" : specific_doc_ref,
                        "Nama Dokumen" : nama_dokumen,
                        "Nomor Faktur Pajak" : facture_num_data,
                        "Status Pembetulan" : status_pembetulan,
                        "Nilai Pembetulan" : nilai_pembetulan,
                        'Status PPH Final' : status_pph,
                        "Status PPH Tidak Final" : status_pph_tidakfinal,
                        "Status Pembatalan" : status_pembatalan,
                        "Tarif Tanpa NPWP" : tarif_tinggi,
                        "NITKU" : v_nitkua,
                        "A.4 Nama" : v_namaa,
                        "a.2 NIK" : v_nika,
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
                        "Nomor Faktur Pajak" : "Error! Please Check The File",
                        "Nomor Dokumen": "Error! Please Check The File",
                        "Nama Dokumen" : "Error! Please Check The File",
                        "Tarif Tanpa NPWP": "Error! Please Check The File",
                        "Status Pembetulan" : "Error! Please Check The File",
                        "Nilai Pembetulan" : "Error! Please Check The File",
                        "Status PPH Final" : "Error! Please Check The File",
                        "Status PPH Tidak Final" : "Error! Please Check The File",
                        "Status Pembatalan" : "Error! Please Check The File",
                        "NITKU" : "Error! Please Check The File",
                        "A.4 Nama" : "Error! Please Check The File",
                        "a.2 NIK" : "Error! Please Check The File",
                        "PDF_Name" : pdffile.name
                    })

        return extracted_details

    # Streamlit app
    st.title('OCS - PPH OLD')

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
            "Nomor Dokumen": "Nomor_Dokumen",
            "Nama Dokumen": "Nama_Dokumen",
            "Status Pembetulan" : "Status_Pembetulan",
            "Nilai Pembetulan" : "Nilai_Pembetulan",
            'Status PPH Final' : 'Status_PPH_Final',
            "Status PPH Tidak Final" : "Status_PPH_Tidak_Final",
            "Status Pembatalan" : "Status_Pembatalan",
            "Tarif Tanpa NPWP": "Tarif_Tanpa_NPWP",
            "Nomor Faktur Pajak": "Nomor_Faktur_Pajak",
            "NITKU" : "NITKU",
            "A.4 Nama" : "Nama_Dipotong",
            "a.2 NIK" : "NIK_Dipotong"

        }, inplace=True)

        # Reordering columns to match the desired format
        df = df[['NO', 'PDF_Name','NITKU', 'Nama_Dipotong', 'NIK_Dipotong', 'Nama_Pemotong', 'NPWP_Pemotong', 'Pasal', 'Pph', 'Status_Pembetulan', 'Nilai_Pembetulan', 'Status_PPH_Final', 'Status_PPH_Tidak_Final', 'Status_Pembatalan', 'Nilai_Obj_Pemotongan', 'Pph_potput', 'Tarif_Tanpa_NPWP', 'Nomor_Faktur_Pajak', 'Nomor_Bukti', 'Tanggal', 'Tanggal_Dokumen','Nomor_Dokumen','Nama_Dokumen', 'Alamat', 'NTPN']]

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

elif page == "PPH - 2025":
    # st.write("This is the PPH - 2025 page. Here you can add the specific job or content related to PPH - 2025.")
    # Azure credentials
    endpoint = st.secrets["ocrendpoint"]
    key = st.secrets["ocrkey"]
    model_id = st.secrets["ocrmodelid"]

    # Initialize Azure Document Intelligence client
    document_intelligence_client = DocumentIntelligenceClient(
        endpoint=endpoint, credential=AzureKeyCredential(key)
    )

    # File uploader
    uploaded_files = st.file_uploader("Choose PDF file(s)", type=["pdf"], accept_multiple_files=True)

    if uploaded_files:
        all_extracted_fields = []
        total_scan = 0

        for uploaded_file in uploaded_files:
            total_scan +=1
        
            # Toast-style progress updates
            msg = st.toast(f"📄 Scanning {total_scan} of {len(uploaded_files)} files...")
            time.sleep(3)
            msg.toast("🔍 Analyzing document...")
            time.sleep(1)

            max_retries = 5
            retry_delay = 5  # seconds
            result = None

            with st.spinner(f"Analyzing {uploaded_file.name}..."):
                for attempt in range(max_retries):
                    try:
                        file_stream = BytesIO(uploaded_file.getvalue())
                        poller = document_intelligence_client.begin_analyze_document(
                            model_id=model_id,
                            body=file_stream,
                            content_type="application/pdf"
                        )
                        result = poller.result(timeout=30)

                        # Optional: Access raw HTTP response for debugging
                        try:
                            raw_response = poller._polling_method._initial_response
                            status_code = raw_response.http_response.status_code
                            headers = dict(raw_response.http_response.headers)
                            body_preview = raw_response.http_response.text()[:1000]
                            # Uncomment below to debug
                            # st.text(f"HTTP Status Code: {status_code}")
                            # st.text(f"HTTP Headers: {headers}")
                            # st.text(f"Raw Response Body (truncated): {body_preview}")
                        except Exception as debug_error:
                            st.warning(f"Could not retrieve raw HTTP response: {debug_error}")
                        break  # Exit loop if successful

                    except HttpResponseError as e:
                        st.warning(f"Attempt {attempt + 1} failed for {uploaded_file.name}: {e.message}")
                        time.sleep(retry_delay)
            if result:
                extracted_fields = {"Document": uploaded_file.name}
                if result.documents:
                    for document in result.documents:
                        for name, field in document.fields.items():
                            extracted_fields[name] = field.content
                else:
                    st.warning(f"No fields extracted from {uploaded_file.name}. The model may not recognize this layout.")
                all_extracted_fields.append(extracted_fields)
            else:
                st.error(f"Failed to analyze {uploaded_file.name} after {max_retries} attempts.")
            time.sleep(3)  # Respect rate limits

        if all_extracted_fields:
            combined_df = pd.DataFrame(all_extracted_fields).fillna("")

            st.subheader("Combined Extracted Fields")
            st.dataframe(combined_df)

            # CSV download
            st.download_button(
                label="Download as CSV",
                data=combined_df.to_csv(index=False).encode('utf-8'),
                file_name="combined_extracted_fields.csv",
                mime="text/csv"
            )

            # Excel download
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                combined_df.to_excel(writer, index=False, sheet_name='Extracted Data')
            output.seek(0)

            st.download_button(
                label="Download as Excel",
                data=output,
                file_name="combined_extracted_fields.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

