import streamlit as st
import pandas as pd
import xlsxwriter
from io import BytesIO
import logging

# Configuração do logger
logging.basicConfig(level=logging.INFO)

# Funções auxiliares
def load_users():
    users = [
        st.secrets["USERNAME1"].lower(),
        st.secrets["USERNAME2"].lower(),
        st.secrets["USERNAME3"].lower(),
    ]
    logging.info(f"Loaded users: {users}")
    return users

def authenticate(username):
    logging.info(f"Authenticating username: {username}")
    return username.lower() in load_users()

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
        'valign': 'vcenter'
    })
    header_format = workbook.add_format({
        'bold': True,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#D3D3D3'
    })
    cell_wrap_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True
    })
    return merge_format, header_format, cell_wrap_format

def apply_print_settings(worksheet, header_row):
    worksheet.fit_to_pages(1, 0)
    worksheet.repeat_rows(header_row - 1)
    worksheet.set_print_scale(100)

def generate_spools_template(jc_number, issue_date, area, spools, sgs_df):
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet()

    merge_format, header_format, cell_wrap_format = create_formats(workbook)

    col_widths = {'A': 9.14, 'B': 11.0, 'C': 35.57, 'D': 9.14, 'E': 13.0, 'F': 13.0, 'G': 13.0, 'H': 13.0, 'I': 11.71, 'J': 17.71, 'K': 13.86, 'L': 13.14}
    for col, width in col_widths.items():
        worksheet.set_column(f'{col}:{col}', width, cell_wrap_format)

    header_footer_row_heights = {1: 47.25, 2: 47.25, 3: 47.25}
    for row, height in header_footer_row_heights.items():
        worksheet.set_row(row - 1, height)

    worksheet.merge_range('A1:C3', '', merge_format)
    worksheet.merge_range('D1:H1', 'PETROBRAS', merge_format)
    worksheet.merge_range('D2:H2', 'FPSO_P-82', merge_format)
    worksheet.merge_range('D3:H3', 'Request For Fabrication', merge_format)
    worksheet.merge_range('I1:L3', '', merge_format)

    worksheet.insert_image('A1', '/mnt/data/image.png', {'x_offset': 80, 'y_offset': 10, 'x_scale': 1, 'y_scale': 1})
    worksheet.insert_image('I1', '/mnt/data/image.png', {'x_offset': 80, 'y_offset': 10, 'x_scale': 1, 'y_scale': 1})

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
        
        try:
            module = str(sgs_row.get('Módulo', ''))
            size = str(sgs_row.get('Diam. Polegadas', ''))
            paint_code = str(sgs_row.get('Condição Pintura', ''))
            rev = str(sgs_row.get('Rev. Isometrico', ''))
            shop_id = str(sgs_row.get('Dia Inch', ''))
            weight = float(sgs_row.get('Peso (Kg)', 0))
            base_material = str(sgs_row.get('Material', ''))
        except ValueError as e:
            st.error(f"Error converting values: {e}")
            logging.error(f"Error converting values: {e}")
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

    for r in range(8, row):
        worksheet.set_row(r, 30)

    worksheet.merge_range(f'A{row+1}:F{row+1}', 'Total Weight: (Kg)', merge_format)
    worksheet.merge_range(f'G{row+1}:L{row+1}', total_weight, merge_format)

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

    worksheet.set_row(row + 1, None)
    worksheet.set_row(row + 2, None)

    apply_print_settings(worksheet, header_row=8)

    workbook.close()
    output.seek(0)

    return output

def generate_material_template(jc_number, issue_date, area, drawing_df, spools):
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet()

    merge_format, header_format, cell_wrap_format = create_formats(workbook)

    col_widths = {'A': 35.57, 'B': 13.0, 'C': 22.28, 'D': 9.14, 'E': 13.0, 'F': 46.42, 'G': 9.14, 'H': 13.0, 'I': 13.0, 'J': 13.0, 'K': 13.0, 'L': 13.0}
    for col, width in col_widths.items():
        worksheet.set_column(f'{col}:{col}', width, cell_wrap_format)

    header_footer_row_heights = {1: 47.25, 2: 47.25, 3: 47.25}
    for row, height in header_footer_row_heights.items():
        worksheet.set_row(row - 1, height)

    worksheet.merge_range('A1:C3', '', merge_format)
    worksheet.merge_range('D1:H1', 'PETROBRAS', merge_format)
    worksheet.merge_range('D2:H2', 'FPSO_P-82', merge_format)
    worksheet.merge_range('D3:H3', 'Material Pick Ticket SpoolWise', merge_format)
    worksheet.merge_range('I1:L3', '', merge_format)

    worksheet.insert_image('A1', '/mnt/data/image.png', {'x_offset': 80, 'y_offset': 10, 'x_scale': 1, 'y_scale': 1})
    worksheet.insert_image('I1', '/mnt/data/image.png', {'x_offset': 80, 'y_offset': 10, 'x_scale': 1, 'y_scale': 1})

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
            st.error(f"Error converting values: {e}")
            logging.error(f"Error converting values: {e}")
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

    for r in range(8, row):
        worksheet.set_row(r, 30)

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

    worksheet.set_row(row + 1, None)
    worksheet.set_row(row + 2, None)

    apply_print_settings(worksheet, header_row=8)

    workbook.close()
    output.seek(0)

    return output

# Página de Login
def login_page():
    st.title('Job Card Generator - Login')
    st.header('Por favor, faça login para continuar')
    username = st.text_input('Username', on_change=login, key='username')
    if st.session_state.get('auth_error'):
        st.error(st.session_state.auth_error)

def login():
    if 'username' in st.session_state and authenticate(st.session_state.username):
        st.session_state.authenticated = True
        st.session_state.step = 2
        st.experimental_set_query_params(step=2)
        st.success("Login bem-sucedido")
        st.session_state.auth_error = None
    else:
        st.session_state.auth_error = 'Usuário inválido'
        st.error('Usuário inválido')

# Página de seleção de base de dados
def select_database_page():
    st.title("Escolha uma Base de Dados")
    st.header("Selecione uma base de dados predefinida ou faça o upload de uma nova base de dados")

    option = st.radio("Selecione uma opção", ["Usar Base de Dados Predefinida", "Fazer Upload de Base de Dados"])

    if option == "Usar Base de Dados Predefinida":
        if "base_loaded" not in st.session_state:
            st.session_state.base_loaded = False

        if not st.session_state.base_loaded:
            if st.button('Usar Base de Dados Predefinida', key='use_preset_db'):
                with st.spinner('Carregando base de dados...'):
                    try:
                        st.session_state.sgs_df = process_excel_data('SGS.xlsx', sheet_name='Spool', header=9)
                        st.session_state.drawing_df = process_excel_data('DrawingPartList.xlsx', sheet_name='Sheet1', header=0)
                        st.session_state.base_loaded = True
                        st.success("Base de dados carregada com sucesso.")
                    except Exception as e:
                        st.error(f"Erro ao carregar a base de dados: {e}")
                        logging.error(f"Erro ao carregar a base de dados: {e}")

        if st.session_state.base_loaded:
            st.success("Base de dados carregada com sucesso.")
            if st.button('Next', key='next1'):
                st.session_state.step = 3
                st.experimental_set_query_params(step=3)

    elif option == "Fazer Upload de Base de Dados":
        uploaded_file_sgs = st.file_uploader('Upload do arquivo Excel SGS', type=['xlsx'], key='uploaded_file_sgs')
        if uploaded_file_sgs is not None:
            with st.spinner('Carregando arquivo...'):
                sgs_df = process_excel_data(uploaded_file_sgs)
                if sgs_df is not None:
                    st.session_state.sgs_df = sgs_df
                    st.success("Arquivo SGS carregado com sucesso.")

        uploaded_file_drawing = st.file_uploader('Upload do arquivo Excel Drawing Part List', type=['xlsx'], key='uploaded_file_drawing')
        if uploaded_file_drawing is not None:
            with st.spinner('Carregando arquivo...'):
                drawing_df = process_excel_data(uploaded_file_drawing, sheet_name='Sheet1', header=0)
                if drawing_df is not None:
                    st.session_state.drawing_df = drawing_df
                    st.success("Arquivo Drawing Part List carregado com sucesso.")

        if 'sgs_df' in st.session_state and 'drawing_df' in st.session_state:
            if st.button('Next', key='next2'):
                st.session_state.step = 3
                st.experimental_set_query_params(step=3)

# Página de informações do cartão de trabalho
def job_card_info_page():
    st.title("Informações do Cartão de Trabalho")
    st.header("Preencha as informações do cartão de trabalho")

    jc_number = st.text_input('Número do JC', value=st.session_state.get('jc_number', ''))
    issue_date = st.date_input('Data de Emissão', value=st.session_state.get('issue_date', pd.to_datetime('today')))
    area = st.text_input('Área', value=st.session_state.get('area', ''))
    spools = st.text_area('Spools (um por linha)', value=st.session_state.get('spools', ''))

    if st.button("Criar Cartões de Trabalho", key='create_job_cards'):
        if not jc_number or not issue_date or not area or not spools:
            st.error('Todos os campos devem ser preenchidos.')
        else:
            with st.spinner('Criando cartões de trabalho...'):
                formatted_issue_date = issue_date.strftime('%Y/%m/%d')
                spools_excel = generate_spools_template(jc_number, formatted_issue_date, area, spools, st.session_state.sgs_df)
                material_excel = generate_material_template(jc_number, formatted_issue_date, area, st.session_state.drawing_df, spools)
                st.session_state.spools_excel = spools_excel
                st.session_state.material_excel = material_excel
                st.session_state.jc_number = jc_number
                st.session_state.issue_date = issue_date
                st.session_state.area = area
                st.session_state.spools = spools
                st.success("Cartões de Trabalho criados com sucesso.")
                if st.button('Next', key='next3'):
                    st.session_state.step = 4
                    st.experimental_set_query_params(step=4)

# Página de download dos cartões de trabalho
def download_page():
    st.title("Download dos Cartões de Trabalho")
    st.header("Baixe os cartões de trabalho gerados")

    if 'jc_number' not in st.session_state:
        st.error("Nenhum cartão de trabalho gerado. Por favor, volte e complete as etapas anteriores.")
        return

    jc_number = st.session_state.jc_number
    st.download_button(
        label="Download dos Cartões de Trabalho - Spools",
        data=st.session_state.spools_excel.getvalue(),
        file_name=f"JobCard_{jc_number}_Spools.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.download_button(
        label="Download dos Cartões de Trabalho - Material",
        data=st.session_state.material_excel.getvalue(),
        file_name=f"JobCard_{jc_number}_Material.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Função principal
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
    elif st.session_state.step == 2:
        select_database_page()
    elif st.session_state.step == 3:
        job_card_info_page()
    elif st.session_state.step == 4:
        download_page()

if __name__ == "__main__":
    main()
