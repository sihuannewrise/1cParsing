import os
from datetime import datetime
from openpyxl import load_workbook
# from src.core.crud.expense import create_expense

CWD = os.getcwd()

SOURCE_DIR = './files/in/'
DESTIN_DIR = './files/out/'

SOURCE_FILE_NAME = 'knh51-2022.xlsx'

DESTIN_FILE_SUFFIX = ' - out'
DESTIN_SHEET_NAME = 'proc'

file = SOURCE_DIR + SOURCE_FILE_NAME
PERIOD, VAT, DATA, four, five, six, AMOUNT = range(7)

ca = set()
contract = set()


async def get_ca_by_name(name):
    
    return


async def create_new_expense(filename):
    """Загружаем файл.
    """
    wb = load_workbook(filename=filename)
    ws = wb.active

    for row in ws.iter_rows(min_row=2, max_row=3, values_only=True):
        current_date = datetime.strptime(row[PERIOD], '%d.%m.%Y').date()
        # year = current_date.strftime('%Y')
        # current_month = current_date.strftime('%m')
        _, current_ca, current_contract, *_ = row[DATA].split('\n')
        current_sum = round(float(row[AMOUNT]), 2)
        

        expense = {
            'period': current_date,
            'ca_id': current_ca,
            'contract_id': current_contract,
            'amount': current_sum,
        }
        new_expense = await create_expense(expense)
    
    return None
