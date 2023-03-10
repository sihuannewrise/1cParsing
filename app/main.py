import os
import re
import json

from openpyxl import load_workbook, Workbook

CWD = os.getcwd()

vat_search = r'Без налога (НДС)'
ip_search = r'ИП'

SOURCE_DIR = 'app/files/in/'
DESTIN_DIR = 'app/files/'

SOURCE_FILE_NAME = 'knh51-2022.xlsx'
DESTIN_FILE_NAME = 'expenses.xlsx'

file = SOURCE_DIR + SOURCE_FILE_NAME
xlsx_file = DESTIN_DIR + DESTIN_FILE_NAME

PERIOD, VAT, DATA, four, five, six, AMOUNT = range(7)


class Expenses(dict):
    def __missing__(self, key):
        value = self[key] = type(self)()
        return value


def create_dict(filename):
    """Загружаем файл.
    """
    wb = load_workbook(filename=filename)
    ws = wb.active
    expenses = Expenses()

    for row in ws.iter_rows(min_row=2, max_row=2, values_only=True):
        # current_date = datetime.strptime(row[PERIOD], '%d.%m.%Y').date()
        # year = current_date.strftime('%Y')
        # current_month = int(current_date.strftime('%m'))
        _, current_ca, current_contract, *_ = row[DATA].split('\n')
        zero_vat = re.search(vat_search, row[VAT])
        ip_ca = re.search(ip_search, current_ca)
        if zero_vat or ip_ca:
            current_sum = int(row[AMOUNT])
        else:
            current_sum = int(row[AMOUNT]/1.2)

        # с разбивкой по годам
        if current_contract not in expenses[current_ca]:
            expenses[current_ca][current_contract] = current_sum
        else:
            expenses[current_ca][current_contract] += current_sum

        # с разбивкой по месяцам
        # if current_month not in expenses[current_ca][current_contract]:
        #     expenses[current_ca][current_contract][current_month] = round(
        #         current_sum/1.2, 2)
        # else:
        #     expenses[current_ca][current_contract][current_month] += round(
        #         current_sum/1.2, 2)
    return expenses


# print(len(create_dict(file)))


def write_dict_to_json(data):
    with open("app/files/expenses.json", "w", encoding='UTF-8') as outfile:
        json.dump(data, outfile, ensure_ascii=False)
    return None


# write_dict_to_json(create_dict(file))


def write_dict_to_xlsx(dict):
    wb = Workbook()
    ws = wb.active
    start_col = 1
    start_row = 1
    for ca, val in dict.items:
        current_col = start_col
        current_row = start_row
        ws.cell(column=current_col, row=current_row).value = ca
        current_col += 1
        for key, amount in val.items:
            ws.cell(column=current_col, row=current_row).value = key
            ws.cell(column=current_col, row=current_row).value = amount
    wb.save(filename=xlsx_file)


write_dict_to_xlsx(create_dict(file))
