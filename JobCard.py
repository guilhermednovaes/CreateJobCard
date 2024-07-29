import streamlit as st
import pandas as pd
import xlsxwriter
from io import BytesIO

# Define a function to generate the Excel template
def generate_template(jc_number, issue_date, area, spools, sgs_df):
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet()

    # Define the formats
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

    cell_format = workbook.add_format({'border': 1})
    bold_format = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    editable_format = workbook.add_format({'border': 1, 'bg_color': '#FFFFE0'})

    # Set column widths for better appearance
    worksheet.set_column('A:A', 5)
    worksheet.set_column('B:B', 15)
    worksheet.set_column('C:C', 35)
    worksheet.set_column('D:D', 5)
    worksheet.set_column('E:E', 5)
    worksheet.set_column('F:F', 10)
    worksheet.set_column('G:G', 5)
    worksheet.set_column('H:H', 10)
    worksheet.set_column('I:I', 15)
    worksheet.set_column('J:J', 15)
    worksheet.set_column('K:K', 20)
    worksheet.set_column('L:L', 15)

    # Merge cells for the header
    worksheet.merge_range('A1:B5', '', merge_format)  # For Petrobras logo
    worksheet.merge_range('C1:H1', 'PETROBRAS', merge_format)
    worksheet.merge_range('C2:H2', 'FPSO_P-82', merge_format)
    worksheet.merge_range('C3:H3', 'Request For Fabrication', merge_format)
    worksheet.merge_range('I1:L5', '', merge_format)  # For Seatrium logo

    # Add main headers with editable fields
    worksheet.merge_range('A6:H6', f'JC Number : {jc_number}', editable_format)
    worksheet.merge_range('I6:L6', '', editable_format)
    worksheet.merge_range('A7:H7', f'Issue Date : {issue_date}', editable_format)
    worksheet.merge_range('I7:L7', '', editable_format)
    worksheet.merge_range('A8:H8', f'Area : {area}', editable_format)
    worksheet.merge_range('I8:L8', '', editable_format)

    # Add special instruction
    worksheet.merge_range('A9:L9', 'Special Instruction : Please be informed that Materials for the following. SPOOL PIECE No.[s] are available for Issuance.', merge_format)

    # Add table headers
    headers = ['No.', 'Area / WBS', 'Spool', 'Sheet', 'Size', 'Paint Code', 'REV.', 'Shop ID', 'Weight', 'Base Material', 'Material Status', 'Remarks']
    worksheet.write_row('A10', headers, header_format)

    # Process spools
    spools_list = spools.split('\n')
    data = []
    for idx, spool in enumerate(spools_list, start=1):
        if spool in sgs_df['PF Code'].values:
            sgs_row = sgs_df[sgs_df['PF Code'] == spool].iloc[0]
            data.append([
                idx,
                sgs_row.get('Área', ''),
                spool,
                '',  # Sheet
                '',  # Size
                '',  # Paint Code
                '',  # REV.
                '',  # Shop ID
                '',  # Weight
                sgs_row.get('Material', ''),
                '',  # Material Status
                ''   # Remarks
            ])
        else:
            data.append([idx, '', spool, '', '', '', '', '', '', '', '', ''])

    # Add data to the worksheet
    row = 10
    col = 0
    for item in data:
        worksheet.write_row(row, col, item, cell_format)
        row += 1

    # Calculate the total weight (assuming it's provided in the sgs_df)
    total_weight = sum(float(item[8]) for item in data if item[8])

    # Add total weight and footer dynamically after the data
    footer_start = row

    worksheet.merge_range(f'A{footer_start+1}:H{footer_start+1}', 'Total Weight: (Kg)', merge_format)
    worksheet.merge_range(f'I{footer_start+1}:L{footer_start+1}', f'{total_weight:,.3f}', merge_format)

    worksheet.merge_range(f'A{footer_start+2}:B{footer_start+2}', 'Prepared by', bold_format)
    worksheet.merge_range(f'C{footer_start+2}:D{footer_start+2}', 'Approved by', bold_format)
    worksheet.merge_range(f'E{footer_start+2}:L{footer_start+2}', 'Received', bold_format)

    worksheet.merge_range(f'A{footer_start+3}:B{footer_start+3}', 'Piping Engg.', bold_format)
    worksheet.merge_range(f'C{footer_start+3}:D{footer_start+3}', 'J/C Co-Ordinator', bold_format)
    worksheet.merge_range(f'E{footer_start+3}:L{footer_start+3}', 'Spooling Vendor : EJA', bold_format)

    worksheet.merge_range(f'A{footer_start+4}:L{footer_start+4}', 'CC', bold_format)

    workbook.close()

    output.seek(0)
    return output

st.title('Generate Fabrication Request')

# Input fields
jc_number = st.text_input('JC Number:')
issue_date = st.date_input('Issue Date:')
area = st.text_input('Area:')
spools = st.text_area('Spools (one per line):')

# File uploader for SGS Excel
uploaded_file = st.file_uploader("Upload SGS Excel file", type="xlsx")

if uploaded_file:
    # Read the "Spool" sheet starting from row 7, ignoring row 8
    sgs_df = pd.read_excel(uploaded_file, sheet_name='Spool', skiprows=7)
    sgs_df = sgs_df.iloc[1:].reset_index(drop=True)  # Drop row 8 which is the first row of the DataFrame after skiprows

    # Ensure the 'PF Code' column exists in the dataframe
    if 'PF Code' in sgs_df.columns:
        if st.button('Generate Template'):
            output = generate_template(jc_number, issue_date, area, spools, sgs_df)
            st.download_button(label="Download Excel file", data=output, file_name=f'{jc_number}_template.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    else:
        st.error("The uploaded Excel file does not contain a 'PF Code' column in the 'Spool' sheet.")
