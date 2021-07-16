from datetime import date
from openpyxl import load_workbook
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib import colors

file_name = 'Excel_And_Pdf/Customers Data Base.xlsx'


def save_as_pdf(user_data: dict, index: str) -> str:
    file_name_pdf = index + "_" + user_data['Имя клиента'] + '.pdf'
    document_title = ''
    title = ''
    text_lines = [
        'Name: ' + user_data['Имя клиента'],
        'Phone Number: ' + user_data['Номер телефона'],
        'Issue: ' + user_data['Поломка']
    ]
    pdf = canvas.Canvas('Excel_And_Pdf/PDF/' + file_name_pdf)
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
    return file_name_pdf


def assignedDataToExcel(file_name_excel: str, user_data, repair_number: str, row: int, done: bool) -> None:
    """
    Задаём в таблицу In_Progress значения по строке + 1 - если done == False
    Задаём в таблицу Done значения из In_Progress - если done == True
    """
    wb = load_workbook(file_name_excel)
    In_Progress_Table = wb['In_Progress']
    Repair_Done = wb['Done']
    if not done:
        In_Progress_Table["A" + str(row + 1)] = repair_number
        In_Progress_Table["B" + str(row + 1)] = user_data['Имя клиента']
        In_Progress_Table["C" + str(row + 1)] = user_data['Номер телефона']
        In_Progress_Table["D" + str(row + 1)] = user_data['Поломка']
    else:
        Repair_Done["A" + str(row + 1)] = repair_number
        Repair_Done["B" + str(row + 1)] = user_data[0].value
        Repair_Done["C" + str(row + 1)] = user_data[1].value
        Repair_Done["D" + str(row + 1)] = user_data[2].value
    wb.save(file_name)


def store_file(user_data) -> str:
    today = date.today()
    wb = load_workbook('Excel_And_Pdf/Customers Data Base.xlsx')
    sheet = wb.active
    last_row = sheet.max_row
    last_row_value = sheet.cell(column=1, row=last_row).value
    if today.strftime('%d.%m.%y') == last_row_value[:8]:
        counter = int(last_row_value[9:]) + 1  # Если дата всё таже, то присваиваем последний номер + 1
    else:
        counter = 1  # Если дата изменилась, то обнуляем счётчик
    index = today.strftime('%d.%m.%y.') + str(counter).zfill(2)  # Создаём номер ремонта типа "dd.mm.yy.client number"
    assignedDataToExcel(file_name, user_data=user_data, repair_number=index, row=last_row, done=False)
    return str(index) + save_as_pdf(user_data=user_data, index=index)

# TODO: Сделать файл в Google Drive, чтобы в него можно было писать и бот подхватывал это.


def save_data_to_another_table(repair_number: str) -> bool:
    wb = load_workbook(file_name)
    In_Progress = wb['In_Progress']
    Done_Repair = wb['Done']
    for row_number in range(1, len(In_Progress['A'])):
        if In_Progress["A"][row_number].value == repair_number:  # Проверяем, есть ли такой ремонт в таблице
            assignedDataToExcel(file_name, user_data=In_Progress[row_number + 1][1:], repair_number=repair_number,
                                row=Done_Repair.max_row, done=True)
            In_Progress.delete_rows(row_number + 1, 1)  # Удаляем строку из In_Progress
            wb.save(file_name)
            break
    return True
