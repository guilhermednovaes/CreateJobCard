import streamlit as st
import pandas as pd
import xlsxwriter
from io import BytesIO
import logging
import os

# Configuração do logger
logging.basicConfig(level=logging.INFO)

def authenticate(username, password):
    username = username.lower()
    valid_users = [
        (os.getenv('USERNAME1', '').lower(), os.getenv('PASSWORD1', '')),
        (os.getenv('USERNAME2', '').lower(), os.getenv('PASSWORD2', '')),
        (os.getenv('USERNAME3', '').lower(), os.getenv('PASSWORD3', ''))
    ]
    return (username, password) in valid_users

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

    cell_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'})

    return merge_format, header_format, cell_format

def generate_spools_template(jc_number, issue_date, area, spools, sgs_df):
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet()

    merge_format, header_format, cell_format = create_formats(workbook)

    # Definir as larguras das colunas
    col_widths = {'A': 9.140625, 'B': 11.0, 'C': 35.5703125, 'D': 9.140625, 'E': 13.0, 'F': 13.0, 'G': 13.0, 'H': 13.0, 'I': 11.7109375, 'J': 17.7109375, 'K': 13.86, 'L': 13.140625}
    for col, width in col_widths.items():
        worksheet.set_column(f'{col}:{col}', width)

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
        worksheet.write_row(row, col, data, cell_format)
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
    
    row += 1
    worksheet.merge_range(f'A{row}:E{row}', '', merge_format)
    worksheet.merge_range(f'F{row}:F{row}', '', merge_format)
    worksheet.merge_range(f'G{row}:L{row}', '', merge_format)

    row += 1
    worksheet.merge_range(f'A{row}:L{row}', '', merge_format)   

    workbook.close()
    output.seek(0)

    return output

def generate_material_template(jc_number, issue_date, area, drawing_df):
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet()

    merge_format, header_format, cell_format = create_formats(workbook)

    # Definir as larguras das colunas
    col_widths = {'A': 25, 'B': 10, 'C': 15, 'D': 15, 'E': 10, 'F': 40, 'G': 15, 'H': 15, 'I': 10, 'J': 10, 'K': 10, 'L': 15}
    for col, width in col_widths.items():
        worksheet.set_column(f'{col}:{col}', width)

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
    for idx, drawing_row in drawing_df.iterrows():
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
        worksheet.write_row(row, col, data, cell_format)
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
    
    row += 1
    worksheet.merge_range(f'A{row}:E{row}', '', merge_format)
    worksheet.merge_range(f'F{row}:F{row}', '', merge_format)
    worksheet.merge_range(f'G{row}:L{row}', '', merge_format)

    row += 1
    worksheet.merge_range(f'A{row}:L{row}', '', merge_format)   

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

    query_params = st.experimental_get_query_params()
    if 'step' in query_params:
        st.session_state.step = int(query_params['step'][0])

    if st.session_state.step == 1:
        login_page()
        if st.session_state.get('authenticated'):
            st.button('Next', on_click=next_step, args=(2,))
    elif st.session_state.step == 2:
        if st.session_state.authenticated:
            upload_page()
            if st.session_state.get('sgs_df') is not None and st.session_state.get('drawing_df') is not None:
                st.button('Next', on_click=next_step, args=(3,))
    elif st.session_state.step == 3:
        if st.session_state.authenticated:
            job_card_info_page()

def login_page():
    st.title('Job Card Generator - Login')
    username = st.text_input('Username')
    password = st.text_input('Password', type='password')
    if st.button('Login'):
        if authenticate(username, password):
            st.session_state.authenticated = True
            st.session_state.step = 2
            st.success("Login successful")
            st.experimental_set_query_params(step=2)
        else:
            st.error('Invalid username or password')

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
            material_excel = generate_material_template(jc_number, formatted_issue_date, area, drawing_df)
            st.success("Job Cards created successfully.")
            st.download_button(
                label="Download Job Card Spools",
                data=spools_excel.getvalue(),
                file_name=f"JobCard_{jc_number}_Spools.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.download_button(
                label="Download Job Card Material",
                data=material_excel.getvalue(),
                file_name=f"JobCard_{jc_number}_Material.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
