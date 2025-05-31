import os
import openpyxl

folder = 'C:\\Invoice_Annexure'
output_file ='C:\\Invoice_Annexure\\Summary.xlsx'


output_wb = openpyxl.Workbook()
output_sheet = output_wb.active
output_sheet.title = 'Summary Data'

cells = ['B5', 'B6', 'B7', 'B8', 'B9', 'C12', 'C13', 'C14', 'C15', 'C16']

for filename in os.listdir(folder):
    if filename.endswith('.xlsx'):
        file = os.path.join(folder, filename)
        workbook = openpyxl.load_workbook(file)
        values= [workbook.active[cell].value for cell in cells]
        output_sheet.append(values)

output_wb.save(output_file)

