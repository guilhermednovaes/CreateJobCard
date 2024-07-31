import streamlit as st
import pandas as pd
import xlsxwriter
from io import BytesIO
import logging
import os

# Configuração do logger
logging.basicConfig(level=logging.INFO)

PASSWORD_FILE = 'password.txt'

def load_credentials():
    credentials = []
    with open(PASSWORD_FILE, 'r') as file:
        lines = file.readlines()
        for i in range(0, len(lines), 2):
            username = lines[i].strip().split('=')[1].strip()
            password = lines[i+1].strip().split('=')[1].strip()
            credentials.append((username, password))
    return credentials

def save_credentials(credentials):
    with open(PASSWORD_FILE, 'w') as file:
        for username, password in credentials:
            file.write(f'USERNAME = {username}\n')
            file.write(f'PASSWORD = {password}\n')

def authenticate(username, password):
    username = username.lower()
    credentials = load_credentials()
    return (username, password) in credentials

def process_excel_data(uploaded_file, sheet_name='Spool', header=9):
    try:
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header).dropna(how='all')
        df = df.iloc[1:]
        df = df.reset_index(drop=True)
        return df
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
        logging.error(f"Erro ao processar o arquivo: {e}")
        return None

def create_formats(workbook):
    merge_format = workbook.add_format({
        'bold': True,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'})

    header_format = workbook.add_format({
        'bold': True,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#D3D3D3'})

    cell_wrap_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True})

    return merge_format, header_format, cell_wrap_format

def apply_print_settings(worksheet, header_row):
    worksheet.fit_to_pages(1, 0)  # Fit to 1 page wide, no limit on height
    worksheet.repeat_rows(header_row - 1)  # Repeat the header row
    worksheet.set_print_scale(100)  # Set print scale to 100%

def generate_spools_template(jc_number, issue_date, area, spools, sgs_df):
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet()

    merge_format, header_format, cell_wrap_format = create_formats(workbook)

    # Definir as larguras das colunas e ativar quebra de texto
    col_widths = {'A': 9.140625, 'B': 11.0, 'C': 35.5703125, 'D': 9.140625, 'E': 13.0, 'F': 13.0, 'G': 13.0, 'H': 13.0, 'I': 11.7109375, 'J': 17.7109375, 'K': 13.86, 'L': 13.140625}
    for col, width in col_widths.items():
        worksheet.set_column(f'{col}:{col}', width, cell_wrap_format)

    # Definir as alturas das linhas antes e depois da tabela
    header_footer_row_heights = {1: 47.25, 2: 47.25, 3: 47.25}
    for row, height in header_footer_row_heights.items():
        worksheet.set_row(row - 1, height)

    worksheet.merge_range('A1:C3', '', merge_format)
    worksheet.merge_range('D1:H1', 'PETROBRAS', merge_format)
    worksheet.merge_range('D2:H2', 'FPSO_P-82', merge_format)
    worksheet.merge_range('D3:H3', 'Request For Fabrication', merge_format)
    worksheet.merge_range('I1:L3', '', merge_format)

    # Inserção das Imagens
    worksheet.insert_image('A1', 'Logo/BR.png', {'x_offset': 80, 'y_offset': 10, 'x_scale': 1, 'y_scale': 1})
    worksheet.insert_image('I1', 'Logo/Seatrium.png', {'x_offset': 80, 'y_offset': 10, 'x_scale': 1, 'y_scale': 1})

    worksheet.merge_range('A4:D4', f'JC Number : {jc_number}', merge_format)
    worksheet.merge_range('G4:L4', area, merge_format)
    worksheet.merge_range('E4:F4', '', merge_format)
    worksheet.merge_range('A5:D5', f'Issue Date : {issue_date}', merge_format)
    worksheet.merge_range('E5:F5', '', merge_format)
    worksheet.merge_range('G5:L5', '', merge_format)

    worksheet.merge_range('A6:L7', 'Special Instruction : Please be informed that Materials for the following. SPOOL PIECE No.[s] are available for Issuance.', merge_format)

    headers = ['No.', 'Area / WBS', 'Spool', 'Sheet', 'Size', 'Paint Code', 'REV.', 'Shop ID', 'Weight', 'Base Material', 'Material Status', 'Remarks']
    worksheet.write_row('A8', headers, header_format)

    row = 8
    col = 0
    total_weight = 0
    spools_list = list(dict.fromkeys([spool.strip() for spool in spools.split('\n') if spool.strip()]))
    for idx, spool in enumerate(spools_list):
        sgs_row = sgs_df[sgs_df['PF Code'] == spool.strip()].iloc[0] if not sgs_df[sgs_df['PF Code'] == spool.strip()].empty else {}
        
        # Garantir que todos os valores são do tipo esperado
        try:
            module = str(sgs_row.get('Módulo', ''))
            size = str(sgs_row.get('Diam. Polegadas', ''))
            paint_code = str(sgs_row.get('Condição Pintura', ''))
            rev = str(sgs_row.get('Rev. Isometrico', ''))
            shop_id = str(sgs_row.get('Dia Inch', ''))
            weight = float(sgs_row.get('Peso (Kg)', 0))
            base_material = str(sgs_row.get('Material', ''))
        except ValueError as e:
            st.error(f"Erro ao converter os valores: {e}")
            logging.error(f"Erro ao converter os valores: {e}")
            continue
        
        data = [
            idx + 1,
            module,
            spool.strip(),
            '',
            size,
            paint_code,
            rev,
            shop_id,
            weight,
            base_material,
            'Fully Issued',
            ''
        ]
        total_weight += weight
        worksheet.write_row(row, col, data, cell_wrap_format)
        row += 1

    # Definir a altura das linhas da tabela e do cabeçalho
    for r in range(8, row):
        worksheet.set_row(r, 30)

    # Linha do total de peso
    worksheet.merge_range(f'A{row+1}:F{row+1}', 'Total Weight: (Kg)', merge_format)
    worksheet.merge_range(f'G{row+1}:L{row+1}', total_weight, merge_format)

    # Linhas de rodapé
    row += 2
    worksheet.merge_range(f'A{row}:B{row}', 'Prepared by', merge_format)
    worksheet.merge_range(f'C{row}:D{row}', 'Approved by', merge_format)
    worksheet.merge_range(f'E{row}:L{row}', 'Received', merge_format)

    row += 1
    worksheet.merge_range(f'A{row}:B{row}', '', merge_format)
    worksheet.merge_range(f'C{row}:D{row}', '', merge_format)
    worksheet.merge_range(f'E{row}:L{row}', '', merge_format)
    
    row += 1
    worksheet.merge_range(f'A{row}:B{row}', 'Piping Engg.', merge_format)
    worksheet.merge_range(f'C{row}:D{row}', 'J/C Co-Ordinator', merge_format)
    worksheet.merge_range(f'E{row}:L{row}', 'Spooling Vendor : EJA', merge_format)
   
    row += 1
    worksheet.merge_range(f'A{row}:E{row}', '', merge_format)
    worksheet.write(f'F{row}', 'CC', merge_format)
    worksheet.merge_range(f'G{row}:L{row}', '', merge_format)

    # Aplicar formatação apenas até a linha do "CC"
    worksheet.set_row(row + 1, None)
    worksheet.set_row(row + 2, None)

    # Configurações de impressão
    apply_print_settings(worksheet, header_row=8)

    workbook.close()
    output.seek(0)

    return output

def generate_material_template(jc_number, issue_date, area, drawing_df, spools):
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet()

    merge_format, header_format, cell_wrap_format = create_formats(workbook)

    # Definir as larguras das colunas específicas
    col_widths = {'A': 35.5703125, 'B': 13.0, 'C': 22.28515625, 'D': 9.140625, 'E': 13.0, 'F': 46.42578125, 'G': 9.140625, 'H': 13.0, 'I': 13.0, 'J': 13.0, 'K': 13.0, 'L': 13.0}
    for col, width in col_widths.items():
        worksheet.set_column(f'{col}:{col}', width, cell_wrap_format)

    # Definir as alturas das linhas antes e depois da tabela
    header_footer_row_heights = {1: 47.25, 2: 47.25, 3: 47.25}
    for row, height in header_footer_row_heights.items():
        worksheet.set_row(row - 1, height)

    worksheet.merge_range('A1:C3', '', merge_format)
    worksheet.merge_range('D1:H1', 'PETROBRAS', merge_format)
    worksheet.merge_range('D2:H2', 'FPSO_P-82', merge_format)
    worksheet.merge_range('D3:H3', 'Material Pick Ticket SpoolWise', merge_format)
    worksheet.merge_range('I1:L3', '', merge_format)

    # Inserção das Imagens
    worksheet.insert_image('A1', 'Logo/BR.png', {'x_offset': 80, 'y_offset': 10, 'x_scale': 1, 'y_scale': 1})
    worksheet.insert_image('I1', 'Logo/Seatrium.png', {'x_offset': 80, 'y_offset': 10, 'x_scale': 1, 'y_scale': 1})

    worksheet.merge_range('A4:D4', f'JC Number : {jc_number}', merge_format)
    worksheet.merge_range('G4:L4', area, merge_format)
    worksheet.merge_range('E4:F4', '', merge_format)
    worksheet.merge_range('A5:D5', f'Issue Date : {issue_date}', merge_format)
    worksheet.merge_range('E5:F5', '', merge_format)
    worksheet.merge_range('G5:L5', '', merge_format)

    worksheet.merge_range('A6:L7', 'Special Instruction : Please be informed that Materials for the following. SPOOL PIECE No.[s] are available for Issuance.', merge_format)

    headers = ['Spool', 'Rev', 'Mat Code 1', 'Mat Code 2', 'Size', 'Description', 'MRIR No', 'Heat No', 'UOM', 'Req Qty', 'Issued Qty', 'Source']
    worksheet.write_row('A8', headers, header_format)

    row = 8
    col = 0
    spools_list = list(dict.fromkeys([spool.strip() for spool in spools.split('\n') if spool.strip()]))
    drawing_filtered_df = drawing_df[drawing_df['SpoolNo'].isin(spools_list)]
    for idx, drawing_row in drawing_filtered_df.iterrows():
        try:
            spool = str(drawing_row.get('SpoolNo', ''))
            rev = str(drawing_row.get('RevNo', ''))
            mat_code_1 = str(drawing_row.get('SapCode', ''))
            size = str(drawing_row.get('Size_Inch', ''))
            description = str(drawing_row.get('Description', ''))
            req_qty = float(drawing_row.get('RequiredQty', 0))
        except ValueError as e:
            st.error(f"Erro ao converter os valores: {e}")
            logging.error(f"Erro ao converter os valores: {e}")
            continue
        
        data = [
            spool,
            rev,
            mat_code_1,
            '',
            size,
            description,
            '',
            '',
            '',
            req_qty,
            '',
            ''
        ]
        worksheet.write_row(row, col, data, cell_wrap_format)
        row += 1

    # Definir a altura das linhas da tabela e do cabeçalho
    for r in range(8, row):
        worksheet.set_row(r, 30)

    # Linhas de rodapé
    row += 2
    worksheet.merge_range(f'A{row}:B{row}', 'Prepared by', merge_format)
    worksheet.merge_range(f'C{row}:D{row}', 'Approved by', merge_format)
    worksheet.merge_range(f'E{row}:L{row}', 'Received', merge_format)

    row += 1
    worksheet.merge_range(f'A{row}:B{row}', '', merge_format)
    worksheet.merge_range(f'C{row}:D{row}', '', merge_format)
    worksheet.merge_range(f'E{row}:L{row}', '', merge_format)
    
    row += 1
    worksheet.merge_range(f'A{row}:B{row}', 'Piping Engg.', merge_format)
    worksheet.merge_range(f'C{row}:D{row}', 'J/C Co-Ordinator', merge_format)
    worksheet.merge_range(f'E{row}:L{row}', 'Spooling Vendor : EJA', merge_format)
   
    row += 1
    worksheet.merge_range(f'A{row}:E{row}', '', merge_format)
    worksheet.write(f'F{row}', 'CC', merge_format)
    worksheet.merge_range(f'G{row}:L{row}', '', merge_format)

    # Aplicar formatação apenas até a linha do "CC"
    worksheet.set_row(row + 1, None)
    worksheet.set_row(row + 2, None)

    # Configurações de impressão
    apply_print_settings(worksheet, header_row=8)

    workbook.close()
    output.seek(0)

    return output

def next_step(step):
    st.session_state.step = step
    st.experimental_set_query_params(step=step)

def main():
    if 'step' not in st.session_state:
        st.session_state.step = 1
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if 'username' not in st.session_state:
        st.session_state.username = ""

    query_params = st.experimental_get_query_params()
    if 'step' in query_params:
        st.session_state.step = int(query_params['step'][0])

    if st.session_state.step == 1:
        login_page()
    elif st.session_state.step == 2:
        if st.session_state.authenticated:
            if st.session_state.password == '123':
                first_access_page()
            else:
                upload_page()
                if st.session_state.get('sgs_df') is not None and st.session_state.get('drawing_df') is not None:
                    st.button('Next', on_click=next_step, args=(3,))
    elif st.session_state.step == 3:
        if st.session_state.authenticated:
            job_card_info_page()
    elif st.session_state.step == 4:
        if st.session_state.authenticated:
            download_page()

def login_page():
    st.title('Job Card Generator - Login')
    username = st.text_input('Username')
    password = st.text_input('Password', type='password')
    if st.button('Login'):
        if authenticate(username, password):
            st.session_state.authenticated = True
            st.session_state.username = username
            st.session_state.password = password
            st.session_state.step = 2
            st.success("Login successful")
            st.experimental_set_query_params(step=2)
        else:
            st.error('Invalid username or password')

def first_access_page():
    st.title('First Access - Change Password')
    st.write("Your current password is '123'. Please change it to a new password.")
    
    current_password = '123'
    new_password = st.text_input('New Password', type='password')
    confirm_password = st.text_input('Confirm New Password', type='password')

    if st.button('Change Password'):
        if new_password != confirm_password:
            st.error('The new passwords do not match.')
        elif new_password == '123':
            st.error('The new password cannot be "123".')
        else:
            # Update the password in the credentials
            credentials = load_credentials()
            new_credentials = [(user, new_password if user == st.session_state.username else passw) for user, passw in credentials]
            save_credentials(new_credentials)
            st.success('Password changed successfully.')
            st.session_state.password = new_password
            next_step(2)

def upload_page():
    st.title('Job Card Generator')
    st.header("Upload SGS Excel file")
    uploaded_file_sgs = st.file_uploader('Upload SGS Excel file', type=['xlsx'])
    uploaded_file_drawing = st.file_uploader('Upload Drawing Part List Excel file', type=['xlsx'])
    if uploaded_file_sgs is not None and uploaded_file_drawing is not None:
        sgs_df = process_excel_data(uploaded_file_sgs)
        drawing_df = process_excel_data(uploaded_file_drawing, sheet_name='Sheet1', header=0)
        if sgs_df is not None and drawing_df is not None:
            st.session_state.sgs_df = sgs_df
            st.session_state.drawing_df = drawing_df
            st.session_state.uploaded_file_sgs = uploaded_file_sgs
            st.session_state.uploaded_file_drawing = uploaded_file_drawing
            st.session_state.step = 3
            st.success("Files processed successfully.")
            st.experimental_set_query_params(step=3)

def job_card_info_page():
    sgs_df = st.session_state.sgs_df
    drawing_df = st.session_state.drawing_df
    st.title('Job Card Generator')
    jc_number = st.text_input('JC Number')
    issue_date = st.date_input('Issue Date')
    area = st.text_input('Area')
    spools = st.text_area('Spool\'s (one per line)', key='spools_input')

    if spools:
        unique_spools = list(dict.fromkeys([spool.strip() for spool in spools.split('\n') if spool.strip()]))
        spool_label = f"Spool's (one per line) ({len(unique_spools)} Spools)"
        st.session_state.spools = '\n'.join(unique_spools)
        st.text_area(spool_label, value=st.session_state.spools, height=100, key="spools_display", disabled=True)

    if st.button(f"Create Job Cards ({jc_number})"):
        if not jc_number or not issue_date or not area or not spools:
            st.error('All fields must be filled out.')
        else:
            formatted_issue_date = issue_date.strftime('%d/%m/%Y')
            spools_excel = generate_spools_template(jc_number, formatted_issue_date, area, st.session_state.spools, sgs_df)
            material_excel = generate_material_template(jc_number, formatted_issue_date, area, drawing_df, st.session_state.spools)
            st.session_state.spools_excel = spools_excel
            st.session_state.material_excel = material_excel
            st.session_state.jc_number = jc_number
            st.success("Job Cards created successfully.")
            st.button('Next', on_click=next_step, args=(4,))

def download_page():
    st.title('Job Card Generator - Download')
    if 'jc_number' not in st.session_state:
        st.error("No job cards generated. Please go back and complete the previous steps.")
        return

    jc_number = st.session_state.jc_number
    st.download_button(
        label="Download Job Card Spools",
        data=st.session_state.spools_excel.getvalue(),
        file_name=f"JobCard_{jc_number}_Spools.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key='download_spools'
    )
    st.download_button(
        label="Download Job Card Material",
        data=st.session_state.material_excel.getvalue(),
        file_name=f"JobCard_{jc_number}_Material.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key='download_material'
    )
    st.button("Back", on_click=next_step, args=(3,))

if __name__ == "__main__":
    main()
