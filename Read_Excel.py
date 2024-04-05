import openpyxl 
import unittest

wb = openpyxl.load_workbook('C:\\Users\\edward\\OneDrive - Neville Registrars Limited\\Documents\\course work\\ImportTemplate.xlsx', data_only=True)

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
    
print(get_div(search))
    
def get_total():
        return str(ws["S9"].value)
    
print(get_total())

class tests(unittest.TestCase):
    def test_div(self):
        assert get_div(search) == str(ws["S"+str(count)].value)
        assert get_div(" ") == None

    def test_total(self):
        assert get_total() == str(ws["S9"].value)
           
if __name__=='__main__':
    unittest.main()