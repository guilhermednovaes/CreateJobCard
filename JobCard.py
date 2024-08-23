import streamlit as st
import pandas as pd
import xlsxwriter
from io import BytesIO
import logging
import time
from datetime import datetime

# Configuração do logger
logging.basicConfig(level=logging.INFO)

# Função para carregar usuários
def load_users():
    users = [
        st.secrets["USERNAME1"].lower(),
        st.secrets["USERNAME2"].lower(),
        st.secrets["USERNAME3"].lower(),
    ]
    logging.info(f"Loaded users: {users}")
    return users

# Função de autenticação sem senha
def authenticate(username):
    users = load_users()
    logging.info(f"Authenticating username: {username}")
    return username.lower() in users

# Função para processar o arquivo Excel
def process_excel_data(uploaded_file, sheet_name='Spool', header=9):
    try:
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header).dropna(how='all')
        df = df.iloc[1:].reset_index(drop=True)
        return df
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
        logging.error(f"Erro ao processar o arquivo: {e}")
        return None

# Função para criar formatos de célula no Excel
def create_formats(workbook):
    formats = {
        'merge': workbook.add_format({
            'bold': True,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        }),
        'header': workbook.add_format({
            'bold': True,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#D3D3D3'
        }),
        'wrap': workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        })
    }
    return formats

# Função para aplicar configurações de impressão
def apply_print_settings(worksheet, header_row):
    worksheet.fit_to_pages(1, 0)
    worksheet.repeat_rows(header_row - 1)
    worksheet.set_print_scale(100)

# Função para filtrar o DrawingPartList
def filter_and_merge_with_tracker(df: pd.DataFrame, df_piping_fitting_master_tracker: pd.DataFrame = None, df_spec_material: pd.DataFrame = None) -> pd.DataFrame:
    df_filtered = df[
        df["Item"].isin(["CAP", "COUPLING", "ELBOW", "FLANGE", "OLET", "PIPE", "REDUCER", "TEE", "UNION"]) 
        & (~df["SpoolNo"].str.contains("ER"))
        & (df["SapCode"] != "-")
        & (~df["SapCode"].str.startswith("C"))
        & (~df["SapCode"].str.startswith("ESP-"))
    ]
    
    return df_filtered

# Classe para geração de Job Cards
class JobCardGenerator:
    def __init__(self, jc_number, issue_date, area, spools, sgs_df, drawing_df):
        self.jc_number = jc_number
        self.issue_date = issue_date
        self.area = area
        self.spools = spools
        self.sgs_df = sgs_df
        self.drawing_df = drawing_df
    
    def generate_spools_template(self):
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output)
        worksheet = workbook.add_worksheet()
        formats = create_formats(workbook)
        
        # Código para gerar template de spools
        self._setup_spools_worksheet(worksheet, formats)
        self._populate_spools_data(worksheet, formats)
        
        apply_print_settings(worksheet, header_row=8)
        workbook.close()
        output.seek(0)
        return output
    
    def _setup_spools_worksheet(self, worksheet, formats):
        col_widths = {'A': 9.14, 'B': 11.0, 'C': 35.57, 'D': 9.14, 'E': 13.0, 'F': 13.0, 'G': 13.0, 'H': 13.0, 'I': 11.71, 'J': 17.71, 'K': 13.86, 'L': 13.14}
        for col, width in col_widths.items():
            worksheet.set_column(f'{col}:{col}', width, formats['wrap'])
        
        header_footer_row_heights = {1: 47.25, 2: 47.25, 3: 47.25}
        for row, height in header_footer_row_heights.items():
            worksheet.set_row(row - 1, height)
        
        worksheet.merge_range('A1:C3', '', formats['merge'])
        worksheet.merge_range('D1:H1', 'PETROBRAS', formats['merge'])
        worksheet.merge_range('D2:H2', 'FPSO_P-82', formats['merge'])
        worksheet.merge_range('D3:H3', 'Request For Fabrication', formats['merge'])
        worksheet.merge_range('I1:L3', '', formats['merge'])
        
        worksheet.insert_image('A1', 'Logo/BR.png', {'x_offset': 80, 'y_offset': 10, 'x_scale': 1, 'y_scale': 1})
        worksheet.insert_image('I1', 'Logo/Seatrium.png', {'x_offset': 80, 'y_offset': 10, 'x_scale': 1, 'y_scale': 1})
        
        worksheet.merge_range('A4:D4', f'JC Number : {self.jc_number}', formats['merge'])
        worksheet.merge_range('G4:L4', self.area, formats['merge'])
        worksheet.merge_range('A5:D5', f'Issue Date : {self.issue_date}', formats['merge'])
        worksheet.merge_range('A6:L7', 'Special Instruction : Please be informed that Materials for the following SPOOL PIECE No.[s] are available for Issuance.', formats['merge'])
        
        headers = ['No.', 'Area / WBS', 'Spool', 'Sheet', 'Size', 'Paint Code', 'REV.', 'Shop ID', 'Weight', 'Base Material', 'Material Status', 'Remarks']
        worksheet.write_row('A8', headers, formats['header'])
    
    def _populate_spools_data(self, worksheet, formats):
        row = 8
        col = 0
        total_weight = 0
        spools_list = list(dict.fromkeys([spool.strip() for spool in self.spools.split('\n') if spool.strip()]))
        for idx, spool in enumerate(spools_list):
            sgs_row = self.sgs_df[self.sgs_df['PF Code'] == spool.strip()].iloc[0] if not self.sgs_df[self.sgs_df['PF Code'] == spool.strip()].empty else {}
            
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
            worksheet.write_row(row, col, data, formats['wrap'])
            row += 1
        
        for r in range(8, row):
            worksheet.set_row(r, 30)
        
        worksheet.merge_range(f'A{row+1}:F{row+1}', 'Total Weight: (Kg)', formats['merge'])
        worksheet.merge_range(f'G{row+1}:L{row+1}', total_weight, formats['merge'])

    def generate_material_template(self):
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output)
        worksheet = workbook.add_worksheet()
        formats = create_formats(workbook)
        
        # Código para gerar template de materiais
        self._setup_material_worksheet(worksheet, formats)
        self._populate_material_data(worksheet, formats)
        
        apply_print_settings(worksheet, header_row=8)
        workbook.close()
        output.seek(0)
        return output
    
    def _setup_material_worksheet(self, worksheet, formats):
        col_widths = {'A': 35.57, 'B': 13.0, 'C': 22.28, 'D': 9.14, 'E': 13.0, 'F': 46.42, 'G': 9.14, 'H': 13.0, 'I': 13.0, 'J': 13.0, 'K': 13.0, 'L': 13.0}
        for col, width in col_widths.items():
            worksheet.set_column(f'{col}:{col}', width, formats['wrap'])
        
        header_footer_row_heights = {1: 47.25, 2: 47.25, 3: 47.25}
        for row, height in header_footer_row_heights.items():
            worksheet.set_row(row - 1, height)
        
        worksheet.merge_range('A1:C3', '', formats['merge'])
        worksheet.merge_range('D1:H1', 'PETROBRAS', formats['merge'])
        worksheet.merge_range('D2:H2', 'FPSO_P-82', formats['merge'])
        worksheet.merge_range('D3:H3', 'Material Pick Ticket SpoolWise', formats['merge'])
        worksheet.merge_range('I1:L3', '', formats['merge'])
        
        worksheet.insert_image('A1', 'Logo/BR.png', {'x_offset': 80, 'y_offset': 10, 'x_scale': 1, 'y_scale': 1})
        worksheet.insert_image('I1', 'Logo/Seatrium.png', {'x_offset': 80, 'y_offset': 10, 'x_scale': 1, 'y_scale': 1})
        
        worksheet.merge_range('A4:D4', f'JC Number : {self.jc_number}', formats['merge'])
        worksheet.merge_range('G4:L4', self.area, formats['merge'])
        worksheet.merge_range('A5:D5', f'Issue Date : {self.issue_date}', formats['merge'])
        worksheet.merge_range('A6:L7', 'Special Instruction : Please be informed that Materials for the following SPOOL PIECE No.[s] are available for Issuance.', formats['merge'])
        
        headers = ['Spool', 'Rev', 'Mat Code 1', 'Mat Code 2', 'Size', 'Description', 'MRIR No', 'Heat No', 'UOM', 'Req Qty', 'Issued Qty', 'Source']
        worksheet.write_row('A8', headers, formats['header'])
    
    def _populate_material_data(self, worksheet, formats):
        row = 8
        col = 0
        spools_list = list(dict.fromkeys([spool.strip() for spool in self.spools.split('\n') if spool.strip()]))
        
        # Aplicando o filtro no drawing_df antes de usá-lo
        drawing_filtered_df = filter_and_merge_with_tracker(self.drawing_df)
        
        drawing_filtered_df = drawing_filtered_df[drawing_filtered_df['SpoolNo'].isin(spools_list)]
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
            worksheet.write_row(row, col, data, formats['wrap'])
            row += 1
        
        for r in range(8, row):
            worksheet.set_row(r, 30)

# Páginas do fluxo da aplicação
def login_page():
    st.title("Login")
    st.subheader("Bem-vindo ao Gerador de Job Cards")
    username = st.text_input('Username', key='username')
    
    if st.button('Login'):
        if authenticate(username):
            st.session_state.authenticated = True
            st.session_state.step = 2
            st.success("Login realizado com sucesso")

def selection_page():
    st.title("Seleção")
    st.subheader("Escolha uma opção para continuar")
    
    st.info("Use a base de dados pré-configurada se quiser começar rapidamente. Caso contrário, faça o upload de seus próprios arquivos.")

    col1, col2 = st.columns(2)
    with col1:
        if st.button('Usar base de dados pré-configurada', key='preset_db'):
            st.session_state.hide_buttons = True
            with st.spinner('Carregando base de dados pré-configurada...'):
                try:
                    st.session_state.sgs_df = process_excel_data('SGS.xlsx', sheet_name='Spool', header=9)
                    st.session_state.drawing_df = process_excel_data('DrawingPartList.xlsx', sheet_name='Sheet1', header=0)
                    st.success("Base de dados pré-configurada carregada com sucesso.")
                    st.session_state.step = 4
                except Exception as e:
                    st.error(f"Erro ao carregar as bases de dados: {e}")
                    st.session_state.hide_buttons = False

    with col2:
        if st.button('Fazer upload de nova base de dados', key='upload_db'):
            st.session_state.hide_buttons = True
            st.session_state.step = 3

def upload_page():
    st.title("Upload de Arquivos")
    st.subheader("Carregue os arquivos de base de dados")

    uploaded_file_sgs = st.file_uploader('Upload SGS Excel file', type=['xlsx'], key='uploaded_file_sgs')
    uploaded_file_drawing = st.file_uploader('Upload Drawing Part List Excel file', type=['xlsx'], key='uploaded_file_drawing')

    if uploaded_file_sgs and uploaded_file_drawing:
        with st.spinner('Processando arquivos...'):
            sgs_df = process_excel_data(uploaded_file_sgs)
            drawing_df = process_excel_data(uploaded_file_drawing, sheet_name='Sheet1', header=0)
            if sgs_df is not None and drawing_df is not None:
                st.session_state.sgs_df = sgs_df
                st.session_state.drawing_df = drawing_df
                st.success("Arquivos carregados e processados com sucesso.")
                st.session_state.step = 4

def job_card_info_page():
    st.title("Informações do Job Card")
    st.subheader("Preencha as informações para criar os Job Cards")
    
    jc_number = st.text_input('JC Number', value=st.session_state.get('jc_number', ''))
    
    # Exibir data no formato brasileiro
    issue_date = st.date_input('Issue Date', value=st.session_state.get('issue_date', pd.to_datetime('today')), format="DD/MM/YYYY")
    
    area = st.text_input('Area', value=st.session_state.get('area', ''))
    spools = st.text_area('Spool\'s (um por linha)', value=st.session_state.get('spools', ''))

    if st.button("Criar Job Cards"):
        if not jc_number or not issue_date or not area or not spools:  # Corrigido: "or" em vez de "ou"
            st.error('Todos os campos devem ser preenchidos.')
        else:
            with st.spinner('Criando Job Cards...'):
                formatted_issue_date = issue_date.strftime('%d/%m/%Y')  # Formatação da data para DD/MM/YYYY
                generator = JobCardGenerator(jc_number, formatted_issue_date, area, spools, st.session_state.sgs_df, st.session_state.drawing_df)
                
                spools_excel = generator.generate_spools_template()
                material_excel = generator.generate_material_template()
                
                st.session_state.spools_excel = spools_excel
                st.session_state.material_excel = material_excel
                st.session_state.jc_number = jc_number
                st.session_state.issue_date = issue_date
                st.session_state.area = area
                st.session_state.spools = spools
                st.success("Job Cards criados com sucesso.")
                st.session_state.step = 5

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Limpar"):
            st.session_state.jc_number = ''
            st.session_state.issue_date = pd.to_datetime('today')
            st.session_state.area = ''
            st.session_state.spools = ''
    with col2:
        if st.session_state.get('spools_excel') and st.session_state.get('material_excel'):
            if st.button('Próximo'):
                st.session_state.step = 5

def download_page():
    st.title("Download")
    st.subheader("Baixe os Job Cards gerados")

    if 'jc_number' not in st.session_state:
        st.error("Nenhum Job Card gerado. Volte e complete as etapas anteriores.")
        return

    jc_number = st.session_state.jc_number
    st.download_button(
        label="Download Job Card Spools",
        data=st.session_state.spools_excel.getvalue(),
        file_name=f"JobCard_{jc_number}_Spools.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.download_button(
        label="Download Job Card Material",
        data=st.session_state.material_excel.getvalue(),
        file_name=f"JobCard_{jc_number}_Material.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    if st.button("Voltar para Edição"):
        st.session_state.step = 4

# Função principal
def main():
    if 'step' not in st.session_state:
        st.session_state.step = 1
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False

    query_params = st.experimental_get_query_params()
    if 'step' in query_params:
        st.session_state.step = int(query_params['step'][0])

    steps = {
        1: login_page,
        2: selection_page,
        3: upload_page,
        4: job_card_info_page,
        5: download_page,
    }
    
    # Lógica para pular automaticamente para a página correta
    current_step = st.session_state.step
    if current_step in steps:
        steps[current_step]()

    st.sidebar.title("Navegação")
    step_names = ["Login", "Seleção", "Upload de Arquivos", "Informações do Job Card", "Download"]
    st.sidebar.markdown("---")
    for i, name in enumerate(step_names, 1):
        if i <= current_step:
            if st.sidebar.button(name, key=f"step_{i}"):
                st.session_state.step = i
                st.experimental_set_query_params(step=i)
    
    progress = st.sidebar.progress(0)
    progress.progress(st.session_state.step / len(steps))

if __name__ == "__main__":
    main()
