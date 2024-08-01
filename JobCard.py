import streamlit as st
import pandas as pd
import xlsxwriter
from io import BytesIO
import os

USERNAME1 = os.getenv("USERNAME1")
USERNAME2 = os.getenv("USERNAME2")
USERNAME3 = os.getenv("USERNAME3")
usernames = [USERNAME1, USERNAME2, USERNAME3]

def authenticate(username):
    return username.lower() in usernames

def process_excel_data(file, sheet_name='Sheet1', header=0):
    try:
        df = pd.read_excel(file, sheet_name=sheet_name, header=header)
        return df.dropna(how='all').iloc[1:].reset_index(drop=True)
    except Exception as e:
        st.error(f"Error processing file: {e}")
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

    worksheet.merge_range('A1:C1', f'JC Number: {jc_number}', merge_format)
    worksheet.merge_range('A2:C2', f'Issue Date: {issue_date}', merge_format)
    worksheet.merge_range('A3:C3', f'Area: {area}', merge_format)

    headers = ['Spool Number', 'Drawing Number', 'Material']
    for col_num, header in enumerate(headers):
        worksheet.write(4, col_num, header, header_format)

    for row_num, spool in enumerate(spools.split('\n'), start=5):
        worksheet.write(row_num, 0, spool, cell_wrap_format)
        drawing_number = sgs_df[sgs_df['SPOOL'] == spool]['DRAWING'].values[0]
        material = sgs_df[sgs_df['SPOOL'] == spool]['MATERIAL'].values[0]
        worksheet.write(row_num, 1, drawing_number, cell_wrap_format)
        worksheet.write(row_num, 2, material, cell_wrap_format)

    apply_print_settings(worksheet, 5)
    workbook.close()
    output.seek(0)
    return output

def generate_material_template(jc_number, issue_date, area, drawing_df, spools):
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet()
    merge_format, header_format, cell_wrap_format = create_formats(workbook)

    worksheet.merge_range('A1:C1', f'JC Number: {jc_number}', merge_format)
    worksheet.merge_range('A2:C2', f'Issue Date: {issue_date}', merge_format)
    worksheet.merge_range('A3:C3', f'Area: {area}', merge_format)

    headers = ['Material', 'Quantity', 'Unit']
    for col_num, header in enumerate(headers):
        worksheet.write(4, col_num, header, header_format)

    for row_num, spool in enumerate(spools.split('\n'), start=5):
        materials = drawing_df[drawing_df['SPOOL'] == spool]
        for _, material in materials.iterrows():
            worksheet.write(row_num, 0, material['MATERIAL'], cell_wrap_format)
            worksheet.write(row_num, 1, material['QUANTITY'], cell_wrap_format)
            worksheet.write(row_num, 2, material['UNIT'], cell_wrap_format)
            row_num += 1

    apply_print_settings(worksheet, 5)
    workbook.close()
    output.seek(0)
    return output

def login_page():
    st.title('Job Card Generator - Login')
    username = st.text_input('Username', on_change=login, key='username_input')
    if st.session_state.get('authenticated'):
        st.success("Login successful")
        st.session_state.step = 2
        st.experimental_set_query_params(step=2)

def login():
    username = st.session_state.get('username_input', '')
    if authenticate(username):
        st.session_state.authenticated = True
        st.session_state.step = 2
        st.experimental_set_query_params(step=2)
    else:
        st.error('Invalid username')

def upload_page():
    st.title('Job Card Generator')
    st.header("Upload SGS Excel file")

    use_sgs_db = st.checkbox("Use Database SGS File", key='use_sgs_db')
    use_drawing_db = st.checkbox("Use Database Drawing Part List File", key='use_drawing_db')

    if use_sgs_db and 'sgs_df' not in st.session_state:
        st.session_state.sgs_df = process_excel_data('SGS.xlsx', sheet_name='Spool', header=9)
        st.success("Using SGS file from database.")
    
    if use_drawing_db and 'drawing_df' not in st.session_state:
        st.session_state.drawing_df = process_excel_data('DrawingPartList.xlsx', sheet_name='Sheet1', header=0)
        st.success("Using Drawing Part List file from database.")
    
    if not use_sgs_db:
        uploaded_file_sgs = st.file_uploader('Upload SGS Excel file', type=['xlsx'], key='uploaded_file_sgs')
        if uploaded_file_sgs is not None:
            sgs_df = process_excel_data(uploaded_file_sgs)
            if sgs_df is not None:
                st.session_state.sgs_df = sgs_df
                st.success("SGS file uploaded successfully.")
    
    if not use_drawing_db:
        uploaded_file_drawing = st.file_uploader('Upload Drawing Part List Excel file', type=['xlsx'], key='uploaded_file_drawing')
        if uploaded_file_drawing is not None:
            drawing_df = process_excel_data(uploaded_file_drawing, sheet_name='Sheet1', header=0)
            if drawing_df is not None:
                st.session_state.drawing_df = drawing_df
                st.success("Drawing Part List file uploaded successfully.")
    
    if st.session_state.get('sgs_df') is not None and st.session_state.get('drawing_df') is not None:
        st.session_state.step = 3
        st.experimental_set_query_params(step=3)
        st.button('Next', on_click=lambda: st.experimental_set_query_params(step=3))

def job_card_info_page():
    st.title('Job Card Generator')
    st.header("Job Card Information")
    
    jc_number = st.text_input('JC Number', value=st.session_state.get('jc_number', ''))
    issue_date = st.date_input('Issue Date', value=st.session_state.get('issue_date', pd.to_datetime('today')))
    area = st.text_input('Area', value=st.session_state.get('area', ''))
    spools = st.text_area('Spool\'s (one per line)', value=st.session_state.get('spools', ''))

    if st.button("Create Job Cards"):
        if not jc_number or not issue_date or not area or not spools:
            st.error('All fields must be filled out.')
        else:
            formatted_issue_date = issue_date.strftime('%Y/%m/%d')
            spools_excel = generate_spools_template(jc_number, formatted_issue_date, area, spools, st.session_state.sgs_df)
            material_excel = generate_material_template(jc_number, formatted_issue_date, area, st.session_state.drawing_df, spools)
            st.session_state.spools_excel = spools_excel
            st.session_state.material_excel = material_excel
            st.session_state.jc_number = jc_number
            st.session_state.issue_date = issue_date
            st.session_state.area = area
            st.session_state.spools = spools
            st.success("Job Cards created successfully.")
            st.session_state.step = 4
            st.experimental_set_query_params(step=4)
    
    if st.button("Clear"):
        st.session_state.jc_number = ''
        st.session_state.issue_date = pd.to_datetime('today')
        st.session_state.area = ''
        st.session_state.spools = ''
        st.experimental_rerun()

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
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.download_button(
        label="Download Job Card Material",
        data=st.session_state.material_excel.getvalue(),
        file_name=f"JobCard_{jc_number}_Material.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    if st.button("Back"):
        st.session_state.step = 3
        st.experimental_set_query_params(step=3)

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
        2: upload_page,
        3: job_card_info_page,
        4: download_page,
    }
    
    st.sidebar.title("Navigation")
    step_names = ["Login", "Upload Files", "Job Card Info", "Download"]
    st.sidebar.markdown("---")
    for i, name in enumerate(step_names, 1):
        if st.sidebar.button(name, key=f"step_{i}", disabled=(i > st.session_state.step)):
            st.session_state.step = i
            st.experimental_set_query_params(step=i)

    progress = st.sidebar.progress(0)
    progress.progress(st.session_state.step / len(steps))

    steps[st.session_state.step]()

if __name__ == "__main__":
    main()
