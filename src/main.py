# from json import load
import os
from datetime import datetime
from openpyxl import load_workbook, Workbook

CWD = os.getcwd()

SOURCE_DIR = './files/in/'
DESTIN_DIR = './files/out/'

SOURCE_FILE_NAME = 'Карточка счета 51 за 2022 г.xlsx'

DESTIN_FILE_SUFFIX = ' - out'
DESTIN_SHEET_NAME = 'proc'

file = SOURCE_DIR + SOURCE_FILE_NAME
PERIOD, VAT, DATA, four, five, six, AMOUNT = range(7)


def avd(key, val, dict):
    if key in dict:
        dict[key] += val
    else:
        dict.update({key: val})
    return dict

def collect_vocabulary(filename):
    """Загружаем файл.
    """
    wb = load_workbook(filename=filename)
    ws = wb.active
    month = {}
    contract = {}
    ca = {}

    for row in ws.iter_rows(min_row=2, max_row=2, values_only=True):
        current_date = datetime.strptime(row[PERIOD], '%d.%m.%Y').date()
        year, current_month = current_date.strftime('%Y'), current_date.strftime('%m')
        _, current_ca, current_contract, *_ = row[DATA].split('\n')
        current_sum = round(float(row[AMOUNT]), 2)

        avd(
            current_ca,
            avd(
                current_contract,
                avd(
                    current_month, current_sum, month),
                contract),
            ca)
    return ca

print(collect_vocabulary(file))

def dict_to_file(dict, ext):
    """"Запись словаря python в файл."""

    with open('vocab-test.' + ext, 'w', encoding="utf-8") as file:
        file.write('# Beeline xml mapping fields\n')
        file.write('# dict_file settings:\n\n')
        for key, value in dict.items():
            file.write(f"{key} = {value}\n")
            # file.write(f"'{key}': row[mp.{key}],\n")
    return


def dict_intersection(titles, labels):
    """Соответствие названий колонок и номеров столбцов из файла загрузки"""
    mapping = {}
    for title, label in titles.items():
        if labels[label] not in LABELS_NOT_USED:
            mapping.update({title: labels[label]})
    return mapping


# dict_to_file(collect_vocabulary(file_xml, 'xml'), 'py')

def create_form(filename):
    """Creating file with header fields."""

    split = os.path.splitext(filename)
    output_file = split[0] + DESTIN_FILE_SUFFIX + split[1]
    destin_file = DESTIN_DIR + output_file

    form = Workbook()
    sheet = form.active
    sheet.title = DESTIN_SHEET_NAME

    for col in range(1, len(FORM_FIELDS) + 1):
        sheet.cell(row=1, column=col).value = FORM_FIELDS[col-1]

    form.save(destin_file)
    return destin_file
