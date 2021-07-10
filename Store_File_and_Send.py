from datetime import date

import pandas as pd
from openpyxl import load_workbook
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib import colors
counter = 0


def save_as_pdf(user_data: dict, index: str) -> str:
    file_name = index + "_" + user_data['Имя клиента'] + '.pdf'
    document_title = ''
    title = ''
    text_lines = [
        'Name: ' + user_data['Имя клиента'],
        'Phone Number: ' + user_data['Номер телефона'],
        'Issue: ' + user_data['Поломка']
    ]
    pdf = canvas.Canvas(file_name)
    pdf.setTitle(document_title)
    pdfmetrics.registerFont(TTFont('abc', 'VeraBI.ttf'))
    pdf.setFont('abc', 36)
    pdf.drawCentredString(300, 700, title)
    pdf.setFillColorRGB(0, 0, 255)
    pdf.setFont("Courier-Bold", 24)
    pdf.drawCentredString(290, 720, '')
    pdf.line(30, 710, 550, 710)
    text = pdf.beginText(40, 680)
    text.setFont("Courier", 18)
    text.setFillColor(colors.red)
    for line in text_lines:
        text.textLine(line)
    pdf.drawText(text)
    pdf.save()
    return file_name


def assigned_data_to_excel(file_name: str, user_data: dict, repair_number: str, row: int) -> None:
    """
    Задаём значения в строки excel таблицы по номеру последней строки + 1
    """
    wb = load_workbook(file_name)
    sheet = wb.active
    sheet["A" + str(row + 1)] = repair_number
    sheet["B" + str(row + 1)] = user_data['Имя клиента']
    sheet["C" + str(row + 1)] = user_data['Номер телефона']
    sheet["D" + str(row + 1)] = user_data['Поломка']
    wb.save('Customers Data Base.xlsx')


def store_file(user_data) -> str:
    global counter
    counter += 1
    today = date.today()
    wb = load_workbook('Customers Data Base.xlsx')
    sheet = wb.active
    last_row = sheet.max_row
    last_row_value = sheet.cell(column=1, row=last_row).value
    if today.strftime('%d.%m.%y') != last_row_value[:8]:
        counter = 1  # Если дата сменилась, то обнуляем счётчик
    elif int(last_row_value[9:]) > counter:
        counter = int(last_row_value[9:]) + 1  # Если номер ремонта того же дня больше, то увеличиваем счётчик
    index = today.strftime('%d.%m.%y.') + str(counter).zfill(2)  # Создаём номер ремонта типа "dd.mm.yy.client number"
    assigned_data_to_excel(file_name='Customers Data Base.xlsx', user_data=user_data, repair_number=index, row=last_row)
    return str(index) + save_as_pdf(user_data=user_data, index=index)


df = pd.read_excel("Customers Data Base.xlsx")
df_abd = df[df == '08.07.21.01']
print(df_abd)