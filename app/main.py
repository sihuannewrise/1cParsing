import os
import json

# from datetime import datetime
from openpyxl import load_workbook

CWD = os.getcwd()

SOURCE_DIR = './files/in/'
DESTIN_DIR = './files/out/'

SOURCE_FILE_NAME = 'knh51-2022.xlsx'

DESTIN_FILE_SUFFIX = ' - out'
DESTIN_SHEET_NAME = 'proc'

file = SOURCE_DIR + SOURCE_FILE_NAME

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

    for row in ws.iter_rows(min_row=2, values_only=True):
        # current_date = datetime.strptime(row[PERIOD], '%d.%m.%Y').date()
        # year = current_date.strftime('%Y')
        # current_month = int(current_date.strftime('%m'))
        _, current_ca, current_contract, *_ = row[DATA].split('\n')
        current_sum = int(row[AMOUNT])

        # по годам
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


dict = create_dict(file)
with open("files/expenses.json", "w", encoding='UTF-8') as outfile:
    json.dump(dict, outfile, ensure_ascii=False)
