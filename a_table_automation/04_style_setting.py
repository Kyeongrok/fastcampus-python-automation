from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string

path = '주간업무계획표.xlsx'
wb = load_workbook(path)
# 활성화
ws = wb.active

# B~F 열너비 15 설정
ws.column_dimensions['B'].width=15
ws.column_dimensions['C'].width=15
ws.column_dimensions['D'].width=15
ws.column_dimensions['E'].width=15
ws.column_dimensions['F'].width=15

for col in range(1,6):
    ws.column_dimensions[get_column_letter(col)].width = 15

# E 일정 열너비 40 설정
ws.column_dimensions['E'].width=40

# A 열너비 5 설정
ws.column_dimensions['A'].width=5

# 주간업무계획표 폰트 크기 키우기, 굵게
title_f = Font(name='맑은고딕', size=28, bold=True)
# color='ff9999', strikethrough=True, underline='single', italic=True
ws['B5'].font = title_f

# 칼럼명 폰트 굵게
column_f = Font(bold=True)

font_range = ws['B8:F8']

for col in font_range:
    for cell in col:
        cell.font=column_f

# 주간업무계획표, 기간 가운데 정렬
ws['B5'].alignment = Alignment(horizontal='center', vertical='center')
ws['B6'].alignment = Alignment(horizontal='center', vertical='center')

# 칼럼명 가운데 정렬
for col in font_range:
    for cell in col:
        cell.alignment = Alignment(horizontal='center', vertical='center')


#날짜, 요일 가운데 정렬
ws['B9'].alignment = Alignment(horizontal='center', vertical='center')
ws['C9'].alignment = Alignment(horizontal='center', vertical='center')

ws['B14'].alignment = Alignment(horizontal='center', vertical='center')
ws['C14'].alignment = Alignment(horizontal='center', vertical='center')

ws['B19'].alignment = Alignment(horizontal='center', vertical='center')
ws['C19'].alignment = Alignment(horizontal='center', vertical='center')

ws['B24'].alignment = Alignment(horizontal='center', vertical='center')
ws['C24'].alignment = Alignment(horizontal='center', vertical='center')

ws['B29'].alignment = Alignment(horizontal='center', vertical='center')
ws['C29'].alignment = Alignment(horizontal='center', vertical='center')

ws['B34'].alignment = Alignment(horizontal='center', vertical='center')
ws['C34'].alignment = Alignment(horizontal='center', vertical='center')

ws['B39'].alignment = Alignment(horizontal='center', vertical='center')
ws['C39'].alignment = Alignment(horizontal='center', vertical='center')


# # 담당자, 시작일, 컬럼명 색칠
ws['B2'].fill = PatternFill(fgColor='E2EFDA', fill_type='solid')
ws['B3'].fill = PatternFill(fgColor='E2EFDA', fill_type='solid')


# 표 컬럼명 색칠 1번째 방법(주석)
# ws['B8'].fill = PatternFill(fgColor='E2EFDA', fill_type='solid')
# ws['C8'].fill = PatternFill(fgColor='E2EFDA', fill_type='solid')
# ws['D8'].fill = PatternFill(fgColor='E2EFDA', fill_type='solid')
# ws['E8'].fill = PatternFill(fgColor='E2EFDA', fill_type='solid')
# ws['F8'].fill = PatternFill(fgColor='E2EFDA', fill_type='solid')

# 표 컬럼명 색칠 2번째 방법
cell_range = ws['B8:F8']
# print(cell_range)

for col in cell_range:
    for cell in col:
        if cell.row ==8:#8행이면
            cell.fill = PatternFill(fgColor='E2EFDA', fill_type='solid')

# 표 컬럼명 색칠 3번째 방법(주석)
# 지정한 범위 열 반복문
# for col in ws.iter_cols(min_row=8, min_col=2, max_row=8, max_col=6):
#     for cell in col:
#         cell.fill = PatternFill(fgColor='E2EFDA', fill_type='solid')

# 테두리 설정
border_thin = Border(left=Side(style='thin'),
                   right=Side(style='thin'),
                   top=Side(style='thin'),
                   bottom=Side(style='thin'))

# 타이틀 영역
for col in ws.iter_cols(min_row=2, min_col=2, max_row=3, max_col=3):
    for cell in col:
        cell.border = border_thin

# context 영역
for col in ws.iter_cols(min_row=8, min_col=2, max_row=43, max_col=6):
    for cell in col:
        cell.border = border_thin

# 파일 저장하기
wb.save(path)
print('스타일 적용 완료')
