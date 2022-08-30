import win32com.client as win32

# xls --> xlsx로 변경
excel = win32.gencache.EnsureDispatch('Excel.Application')
open_xls = excel.Workbooks.Open('sendRequest.xls')

open_xls.SaveAs('sendRequest.xls'+'x', FileFormat= 51)
open_xls.Close()
excel.Application.Quit()