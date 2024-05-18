import pandas as pd   #  pip install pandas
import openpyxl
from docxtpl import DocxTemplate
path = 'demo.xlsx'
df = pd.read_excel(path)
wb = openpyxl.load_workbook(filename=path)
sheet = wb['vacation']

doc = DocxTemplate('шаблон.docx')

for num in range(2, len(list(sheet.rows)) + 1):
    name = sheet['B' + str(num)].value
    last_name = sheet['A' + str(num)].value
    position = sheet['D' + str(num)].value
    start_data = sheet['E' + str(num)].value
    end_data = sheet['F' + str(num)].value
    work_schedule = sheet['G' + str(num)].value

    context = {
        'name': name,
        'last_name': last_name,
        'position': position,
        'start_data': start_data,
        'end_data': end_data,
        'work_schedule' : work_schedule
    }

    doc.render(context)
    doc.save(''+str(last_name) + ' заявление на отпуск.docx')

