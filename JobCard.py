import streamlit as st
import pandas as pd
import xlsxwriter
from io import BytesIO
import base64
import logging
import os

# Configuração do logger
logging.basicConfig(level=logging.INFO)

def authenticate(username, password):
    """Autentica o usuário utilizando variáveis de ambiente."""
    username = username.lower()  # Convertendo o login para minúsculas
    valid_users = [
        (os.getenv('USERNAME1', '').lower(), os.getenv('PASSWORD1', '')),
        (os.getenv('USERNAME2', '').lower(), os.getenv('PASSWORD2', ''))
    ]
    return (username, password) in valid_users

def process_excel_data(uploaded_file):
    """Processa os dados do arquivo Excel carregado."""
    try:
        df_spool = pd.read_excel(uploaded_file, sheet_name='Spool', header=9).dropna(how='all')
        df_spool = df_spool.iloc[1:]  # Ignorar a primeira linha
        df_spool = df_spool.reset_index(drop=True)  # Resetar índice
        return df_spool
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
        logging.error(f"Erro ao processar o arquivo: {e}")
        return None

def create_formats(workbook):
    """Cria e retorna formatos de célula para o Excel."""
    merge_format = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'})

    header_format = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#D3D3D3'})

    cell_format = workbook.add_format({'border': 1})
    
    return merge_format, header_format, cell_format

def generate_template(jc_number, issue_date, area, spools, sgs_df):
    """Gera o arquivo Excel do Job Card com os dados fornecidos."""
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet()
    
    merge_format, header_format, cell_format = create_formats(workbook)

    worksheet.set_column('A:L', 15)
    worksheet.merge_range('C1:D1', 'PETROBRAS', merge_format)
    worksheet.merge_range('E1:F1', 'FPSO_P-82', merge_format)
    worksheet.merge_range('G1:H1', 'Request For Fabrication', merge_format)
    
    worksheet.insert_image('A1', 'Logo/BR.png', {'x_offset': 10, 'y_offset': 5, 'x_scale': 0.5, 'y_scale': 0.5})
    worksheet.insert_image('I1', 'Logo/Seatrium.png', {'x_offset': 10, 'y_offset': 5, 'x_scale': 0.5, 'y_scale': 0.5})
    
    worksheet.merge_range('A4:H4', f'JC Number : {jc_number}', merge_format)
    worksheet.merge_range('I4:L4', area, merge_format)
    worksheet.merge_range('A5:H5', f'Issue Date : {issue_date}', merge_format)
    
    worksheet.merge_range('A6:L6', 'Special Instruction : Please be informed that Materials for the following. SPOOL PIECE No.[s] are available for Issuance.', merge_format)
    
    headers = ['No.', 'Area / WBS', 'Spool', 'Sheet', 'Size', 'Paint Code', 'REV.', 'Shop ID', 'Weight', 'Base Material', 'Material Status', 'Remarks']
    worksheet.write_row('A7', headers, header_format)
    
    row = 7
    col = 0
    total_weight = 0
    spools_list = [spool.strip() for spool in spools.split('\n') if spool.strip()]
    for idx, spool in enumerate(spools_list):
        sgs_row = sgs_df[sgs_df['PF Code'] == spool.strip()].iloc[0] if not sgs_df[sgs_df['PF Code'] == spool.strip()].empty else {}
        data = [
            idx + 1,
            sgs_row.get('Área', ''),
            spool.strip(),
            '',
            sgs_row.get('Diam. Polegadas', ''),
            '',
            '',
            '',
            sgs_row.get('Peso (Kg)', 0),
            sgs_row.get('Material', ''),
            'Fully Issued',
            ''
        ]
        total_weight += sgs_row.get('Peso (Kg)', 0)
        worksheet.write_row(row, col, data, cell_format)
        row += 1
    
    worksheet.merge_range(f'A{row+1}:B{row+1}', 'Total Weight: (Kg)', merge_format)
    worksheet.write(f'C{row+1}', total_weight, merge_format)
    
    worksheet.merge_range(f'A{row+2}:B{row+2}', 'Prepared by', merge_format)
    worksheet.merge_range(f'C{row+2}:D{row+2}', 'Approved by', merge_format)
    worksheet.merge_range(f'E{row+2}:F{row+2}', 'Received', merge_format)
    
    worksheet.merge_range(f'A{row+3}:B{row+3}', 'Piping Engg.', merge_format)
    worksheet.merge_range(f'C{row+3}:D{row+3}', 'J/C Co-Ordinator', merge_format)
    worksheet.merge_range(f'E{row+3}:F{row+3}', 'Spooling Vendor : EJA', merge_format)
    
    worksheet.merge_range(f'A{row+4}:F{row+4}', 'CC', merge_format)
    
    workbook.close()
    output.seek(0)
    
    return output

def generate_download_link(output, jc_number):
    """Gera link para download do arquivo Excel."""
    val = output.getvalue()
    b64 = base64.b64encode(val).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="JobCard_{jc_number}.xlsx">Download Excel file</a>'
    return href

def login_page():
    st.title('Job Card Generator - Login')
    username = st.text_input('Username')
    password = st.text_input('Password', type='password')
    if st.button('Login'):
        if authenticate(username, password):
            st.session_state.authenticated = True
            st.session_state.step = 2
            st.experimental_rerun()
        else:
            st.error('Invalid username or password')

def upload_page():
    st.title('Job Card Generator')
    st.header("Upload SGS Excel file")
    uploaded_file = st.file_uploader('Upload SGS Excel file', type=['xlsx'])
    if uploaded_file is not None:
        sgs_df = process_excel_data(uploaded_file)
        if sgs_df is not None:
            st.session_state.sgs_df = sgs_df
            st.session_state.uploaded_file = uploaded_file
            st.session_state.step = 3
            st.success("File processed successfully.")
            st.experimental_rerun()  # Forçar a atualização da página para avançar automaticamente

def job_card_info_page():
    sgs_df = st.session_state.sgs_df
    st.title('Job Card Generator')
    st.header("Job Card Information")
    jc_number = st.text_input('JC Number')
    issue_date = st.date_input('Issue Date')
    area = st.text_input('Area')
    spools = st.text_area('Spool\'s (one per line)')
    if st.button(f"Create Job Card ({jc_number})"):
        if not jc_number or not issue_date or not area or not spools:
            st.error('All fields must be filled out.')
        else:
            # Limpeza das linhas em branco nos spools
            spools = '\n'.join([spool.strip() for spool in spools.split('\n') if spool.strip()])

            # Formatação da data para DD/MM/YYYY
            formatted_issue_date = issue_date.strftime('%d/%m/%Y')
            excel_data = generate_template(jc_number, formatted_issue_date, area, spools, sgs_df)

            # Gerar link de download
            download_link = generate_download_link(excel_data, jc_number)
            
            # Exibir link de download
            st.markdown(download_link, unsafe_allow_html=True)

def main():
    if 'step' not in st.session_state:
        st.session_state.step = 1
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False

    if st.session_state.step == 1:
        login_page()
    elif st.session_state.step == 2:
        if st.session_state.authenticated:
            upload_page()
    elif st.session_state.step == 3:
        if st.session_state.authenticated:
            job_card_info_page()

if __name__ == "__main__":
    main()
