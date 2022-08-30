from openpyxl import load_workbook

open_xls = load_workbook('../../a_excel_automation/sendRequest.xls', data_only=True)

print(open_xls)