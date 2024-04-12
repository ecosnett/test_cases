import openpyxl 
import unittest

wb = openpyxl.load_workbook('ImportTemplate.xlsx', data_only=True)

ws = wb.active

sheet_search = input("select sheet: ")
search = int(input("select designation ID: "))

ws = wb[sheet_search]

def get_div(search):
    global count
    count = 0
    for row in ws.iter_rows(min_row=1, min_col=1, max_row=12, max_col=3): 
        count += 1
        for cell in row:
            if cell.value == search:
                            return str(ws["S"+str(count)].value)

            else:
                continue
def get_total():
        return str(ws["S9"].value)

