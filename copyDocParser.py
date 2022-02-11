from docx import Document
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE
import xlsxwriter
import os
import json

all_docs = os.listdir('documents/')
for filename in all_docs:

    doc = Document("documents/" + filename) # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook('sheets/' + filename.replace('.docx', '').replace(' ', '') + '_sheet.xlsx')
    worksheet = workbook.add_worksheet()
    paragraphs = doc.paragraphs
    index = 0
    current_section = ""
    prev_section = ""
    current_phase = ""
    current_title = ""

    spreadsheet_rows = []
    print("parsing documents")
    while index < len(paragraphs):
        row = paragraphs[index]
        # NEW SECTION
        current_row = row.text.strip()
#        current_row = current_row.replace('“', '').replace('”', '')

        if len(current_row) > 1 and current_row[1] == ".":
            prev_section = current_section
            current_section = (row.text)
            if (prev_section == ""):
                prev_section = current_section
            if (current_row.count('.') == 1):
                prinent_title = ''
                current_phase = ''
                current_row = ' foo '
        elif (current_row != "") and (current_row.isupper() and current_row[0] != "{" ):
            prev_section = current_section
            if "[PHASE NOTE]" in current_row:
                current_title = (row.text)
            elif "PHASE" in current_row:
                current_phase = current_row
            else:
                current_title = (row.text)
        else:
                        #Normal Content
            if (current_row != "" and current_section != ""):

                spreadsheet_rows.append([current_section, current_phase, current_title, current_row])
                # edge case with empty section start
            elif prev_section != current_section and current_section.count(".") == 1:
                spreadsheet_rows.append([current_section, current_phase, '', current_row])


        index+=1

    row = 0
    col = 0

    print("Writing spreadsheet")
    # Iterate over the data and write it out row by row.
    cell_format = workbook.add_format({'bold': True, 'italic': False})
    worksheet.set_column(1, 4, 25)
    for section, phase,  title, txt in (spreadsheet_rows):
        if section == "NEW_SECTION":
            cell_format.set_align('center')
            cell_format.set_bg_color('#C04ABC')
            worksheet.set_row(row, 25, cell_format)
            if (title.count(".") == 1):
                cell_format.set_bg_color("#f7f797")
            worksheet.write(row, phase,"","")
        else:
            cell_format = workbook.add_format({'bold': True, 'italic': False})
            cell_format.set_align('center')
            cell_format.set_align('vcenter')
            if (section.count(".") == 1):
                cell_format.set_bg_color("#f7f797")
            cell_format.set_color('#eb34d2')
            worksheet.set_row(row, 25)
            worksheet.write(row, col, section, cell_format)
            cell_format.set_align('left')
            cell_format.set_color('#000000')

            worksheet.write(row, col+ 1, phase,cell_format)
            cell_format = workbook.add_format({'bold': False})

            # Write Dev Notes
            cell_format.set_align('left')
            if ("[VO]" in title):
                cell_format.set_bg_color("#6bdee8")

            worksheet.set_row(row, 25, cell_format)
            worksheet.write(row, col + 2, title,cell_format)
            cell_format.set_text_wrap()
            worksheet.write(row, col + 3, txt, cell_format)
#            cell_format = workbook.add_format({'bold': True, 'italic': False})

        worksheet.set_column(0, 3, 25)
        row += 1

    workbook.close()
