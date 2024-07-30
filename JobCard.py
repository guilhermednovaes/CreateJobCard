import streamlit as st
import pandas as pd
import xlsxwriter
import base64
from io import BytesIO

# Function to generate the job card Excel file
def generate_template(jc_number, issue_date, area, spools, sgs_df):
    # Create a new Excel file and add a worksheet.
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet()

    # Define the formats
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

    # Set column widths for better appearance
    worksheet.set_column('A:L', 15)

    # Merge cells for the header
    worksheet.merge_range('C1:D1', 'PETROBRAS', merge_format)
    worksheet.merge_range('E1:F1', 'FPSO_P-82', merge_format)
    worksheet.merge_range('G1:H1', 'Request For Fabrication', merge_format)

    # Insert images
    worksheet.insert_image('A1', 'Logo/BR.png', {'x_offset': 10, 'y_offset': 5, 'x_scale': 0.5, 'y_scale': 0.5})
    worksheet.insert_image('I1', 'Logo/Seatrium.png', {'x_offset': 10, 'y_offset': 5, 'x_scale': 0.5, 'y_scale': 0.5})

    # Add main headers
    worksheet.merge_range('A4:H4', f'JC Number : {jc_number}', merge_format)
    worksheet.merge_range('I4:L4', area, merge_format)
    worksheet.merge_range('A5:H5', f'Issue Date :{issue_date}', merge_format)

    # Add special instruction
    worksheet.merge_range('A6:L6', 'Special Instruction : Please be informed that Materials for the following. SPOOL PIECE No.[s] are available for Issuance.', merge_format)

    # Add table headers
    headers = ['No.', 'Area / WBS', 'Spool', 'Sheet', 'Size', 'Paint Code', 'REV.', 'Shop ID', 'Weight', 'Base Material', 'Material Status', 'Remarks']
    worksheet.write_row('A7', headers, header_format)

    # Add data to the worksheet
    row = 7
    col = 0
    total_weight = 0
    for i, spool in enumerate(spools.split(',')):
        sgs_row = sgs_df[sgs_df['PF Code'] == spool].iloc[0] if not sgs_df[sgs_df['PF Code'] == spool].empty else {}
        data = [
            i + 1,
            sgs_row.get('Área', ''),
            spool,
            '',
            sgs_row.get('Diam. Nominal (mm)', ''),
            '',
            sgs_row.get('Rev. Isometrico', ''),
            sgs_row.get('Área', ''),
            sgs_row.get('Peso (Kg)', 0),
            sgs_row.get('Material', ''),
            'Fully Issued',
            ''
        ]
        worksheet.write_row(row, col, data, cell_format)
        total_weight += float(sgs_row.get('Peso (Kg)', 0))
        row += 1

    # Add total weight and footer
    worksheet.merge_range(f'A{row+2}:B{row+2}', 'Total Weight: (Kg)', merge_format)
    worksheet.write(f'C{row+2}', total_weight, merge_format)

    worksheet.merge_range(f'A{row+3}:B{row+3}', 'Prepared by', merge_format)
    worksheet.merge_range(f'C{row+3}:D{row+3}', 'Approved by', merge_format)
    worksheet.merge_range(f'E{row+3}:F{row+3}', 'Received', merge_format)

    worksheet.merge_range(f'A{row+4}:B{row+4}', 'Piping Engg.', merge_format)
    worksheet.merge_range(f'C{row+4}:D{row+4}', 'J/C Co-Ordinator', merge_format)
    worksheet.merge_range(f'E{row+4}:F{row+4}', 'Spooling Vendor : EJA', merge_format)

    worksheet.merge_range(f'A{row+5}:F{row+5}', 'CC', merge_format)

    # Close the workbook
    workbook.close()
    return output.getvalue()

# Function to download the Excel file
def download_excel(data, filename):
    b64 = base64.b64encode(data).decode('utf-8')
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}">Download Excel file</a>'
    return href

# Streamlit app
st.title("Job Card Generator")

jc_number = st.text_input("JC Number")
issue_date = st.date_input("Issue Date")
area = st.text_input("Area")
spools = st.text_area("Spools (comma separated)")

uploaded_file = st.file_uploader("Upload SGS Excel file", type=["xlsx"])
if uploaded_file:
    sgs_df = pd.read_excel(uploaded_file, sheet_name='Spool', skiprows=1)
    sgs_df.columns = sgs_df.iloc[0]
    sgs_df = sgs_df[1:]

    if st.button(f"Create Job Card ({jc_number})"):
        excel_data = generate_template(jc_number, issue_date, area, spools, sgs_df)
        st.markdown(download_excel(excel_data, f"{jc_number}.xlsx"), unsafe_allow_html=True)
