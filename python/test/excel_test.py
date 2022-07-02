import sys
sys.path.append("..")

from pipe.xlsx import EXCEL

excel = EXCEL("raw_data.xlsx")
print(excel.get_sheets_name)
excel.active_ws_by_name(excel.get_sheets_name[1])
print(excel.elements_slice("A3", "A3"))
