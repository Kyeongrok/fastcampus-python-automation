import openpyxl

# workbook 만들기
wb = openpyxl.Workbook()
new_file = '주간업무계획표.xlsx'

wb.save(new_file)
print('엑셀파일 생성 완료')
