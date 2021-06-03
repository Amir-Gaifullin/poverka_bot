import openpyxl
import pandas
from openpyxl.styles import PatternFill, Font


# создание xlsx файла
def create_excel(number):
    wb = openpyxl.Workbook()
    sheet = wb.active
    stolb_name = ['TypePOV', 'GosNumberPOV', 'NamePOV', 'DesignationSiPOV', 'DeviceMarkPOV', 'type',
           'SerialNumPOV', 'InventarNumPOV', 'CalibrationDatePOV', 'NextcheckDatePOV', 'MarkCipherPOV',
           'DocPOV', 'DeprcatedPOV', 'temperature', 'pressure', 'hymidity', 'other', 'StandartPOV',
           'GpsPOV', 'SiPOV', 'SiPOVetalon', 'SoPOV', 'ReagentNumber', 'ManufactureYear', 'isOwner',
           'gpsTitle', 'lpsTitle', 'npeNumber', 'rank', 'mpTitle', 'regNumber', 'miOwner', 'signPass',
           'signMi', 'metrologist', 'oMethod', 'calibration', 'reasons', 'structure', 'characteristics',
           'additional_info', 'mimetype', 'filename']
    count = -1
    redFill = PatternFill(start_color='F5F5DC',
                          end_color='F5F5DC',
                          fill_type='solid')
    for row in sheet['A1':'AQ1']:
        for cell in row:
            count += 1
            cell.value = stolb_name[count]
            cell.fill = redFill
            cell.font = Font(size='8', name='Arial')
    wb.save(f'poverka{number}.xlsx')


# mas хранит занчение ячеек, например mas = [1, 2, 3 ...], значит в ячейке Ax = 1, Bx = 2
# , где x равен номеру строчки
# date хранит в себе номер строчки, куда нужно добавлять значение
# number хранит в себе номер xlsx файла
def add_line(mas, number):
    wb = openpyxl.load_workbook(filename=f'poverka{number}.xlsx')
    excel_data_df = pandas.read_excel(f'poverka{number}.xlsx')
    number_line = len(excel_data_df.to_dict(orient='record'))
    sheet = wb.active
    count = 0
    for row in sheet[f'A{number_line+1}':f'AQ{number_line+1}']:
        for cell in row:
            cell.value = mas[count]
            cell.font = Font(size='8', name='Arial')
            count += 1
    wb.save(filename=f'poverka{number}.xlsx')
    wb.close()

