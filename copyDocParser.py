from docx import Document
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE
doc = Document("copydocs3.docx")

import xlsxwriter


# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('sheetCopyDoc.xlsx')
worksheet = workbook.add_worksheet()
paragraphs = doc.paragraphs
index = 0
current_section = ""
current_phase = ""
current_title = ""

spreadsheet_rows = []
while index < len(paragraphs):
    row = paragraphs[index]

    current_row = row.text.strip()
    # NEW SECTION 
    if current_row.isupper() and current_row[1] == ".":
        current_section = (row.text)

    elif current_row.isupper():
        if "PHASE" in current_row:
            current_phase = current_row
        else:
            current_title = (row.text)
        pass
    elif current_row == "":
        pass
    else:
        #Normal Content 
        if current_row != "":
            spreadsheet_rows.append([current_section, current_phase, current_title, current_row])
    index+=1

row = 0
col = 0

print("Writing spreadsheet")
# Iterate over the data and write it out row by row.
for section, phase,  title, txt in (spreadsheet_rows):
    if section == "NEW_SECTION":
        print(phase)
        cell_format = workbook.add_format({'bold': True, 'italic': False})
        cell_format.set_align('center')
        cell_format.set_bg_color('#C04ABC')
        worksheet.set_row(row, 18, cell_format)

        worksheet.write(row, phase,"","")
    else:
        cell_format = workbook.add_format({'bold': True, 'italic': False})
        cell_format.set_align('center')
        cell_format.set_align('vcenter')
        cell_format.set_font_color('#C04ABC')
        worksheet.write(row, col, section, cell_format)
        cell_format = workbook.add_format({'bold': True, 'italic': False})
        cell_format.set_align('left')
        
        worksheet.write(row, col+ 1, phase,cell_format)

        # Write Dev Notes 
        cell_format = workbook.add_format({'bold': False, 'italic': False})
        cell_format.set_align('left')
        
        worksheet.write(row, col + 2, title,cell_format)
        cell_format = workbook.add_format()
        cell_format.set_text_wrap()
        worksheet.write(row, col + 3, txt, cell_format)
    row += 1

workbook.close()
