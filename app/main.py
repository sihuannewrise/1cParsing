import os
import json
import pandas as pd

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


def write_dict_to_json(data):
    with open("files/expenses.json", "w", encoding='UTF-8') as outfile:
        json.dump(data, outfile, ensure_ascii=False)
    return None


# write_dict_to_json(create_dict(file))


def write_dict_to_xlsx(dict):
    df = pd.DataFrame.from_dict(dict)
    df.to_excel('files/expenses.xlsx')


write_dict_to_xlsx(create_dict(file))
