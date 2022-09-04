from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font, Border, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string


# workbook 만들기
def create_excel_file(filename):
    wb = Workbook()

    wb.save(filename)
    print('엑셀파일 생성 완료')


class WeeklyWorkPlan:
    wb = None
    ws = None

    def __init__(self, filename, sheet_no=0):
        self.wb = load_workbook(filename)
        self.ws = self.wb.worksheets[sheet_no]

    def set_title(self, filename):

        # 현재 활성화 되어있는 워크시트 기본값 1번째
        ws = self.ws
        ws.title = '주간업무계획표'
        ws.cell(row=2, column=2, value='담당자')

        # ws.cell(2, 2, value='담당자')
        ws['C2'] = '김경록'
        ws['B3'] = '시작일'
        ws['C3'] = '2022-08-14'

        ws['B5'] = '주간업무계획표'
        ws['B6'] = '(2022-08-14~2022-08-20)'

        # 셀병합
        ws.merge_cells('B5:F5')
        ws.merge_cells('B6:F6')

        self.wb.save(filename)
        print('타이틀 생성 완료')

    def insert_context(self, filename):

        ws = self.ws

        cols_data = ['날짜', '요일', '시간', '일정', '비고']

        for col_idx in range(len(cols_data)):
            # print(col)
            ws.cell(row=8, column=2 + col_idx).value = cols_data[col_idx]

        week_date = ['8/14', '8/15', '8/16', '8/17', '8/18', '8/19', '8/20']

        k = 9
        s = 2
        for week in week_date:
            ws.cell(row=k, column=s).value = week
            k = k + 1

        weekdays = ['Monday', 'Tuesday', 'Wendsday', 'Thursday', 'Friday', 'Saturday', 'Sunday']

        v = 9
        m = 3
        for day in weekdays:
            ws.cell(row=v, column=m).value = day
            v = v + 1

        # 행 중간 삽입
        ws.insert_rows(10, 4)
        ws.insert_rows(15, 4)
        ws.insert_rows(20, 4)
        ws.insert_rows(25, 4)
        ws.insert_rows(30, 4)
        ws.insert_rows(35, 4)
        ws.insert_rows(40, 4)

        # 날짜, 요일, 비고 셀 병합

        ws.merge_cells('B9:B13')  # 날짜
        ws.merge_cells('C9:C13')  # 요일
        ws.merge_cells('F9:F13')  # 비고

        ws.merge_cells('B14:B18')  # 날짜
        ws.merge_cells('C14:C18')  # 요일
        ws.merge_cells('F14:F18')  # 비고

        ws.merge_cells('B19:B23')  # 날짜
        ws.merge_cells('C19:C23')  # 요일
        ws.merge_cells('F19:F23')  # 비고

        ws.merge_cells('B24:B28')  # 날짜
        ws.merge_cells('C24:C28')  # 요일
        ws.merge_cells('F24:F28')  # 비고

        ws.merge_cells('B29:B33')  # 날짜
        ws.merge_cells('C29:C33')  # 요일
        ws.merge_cells('F29:F33')  # 비고

        ws.merge_cells('B34:B38')  # 날짜
        ws.merge_cells('C34:C38')  # 요일
        ws.merge_cells('F34:F38')  # 비고

        ws.merge_cells('B39:B43')  # 날짜
        ws.merge_cells('C39:C43')  # 요일
        ws.merge_cells('F39:F43')  # 비고

        self.wb.save(filename)
        print('context 생성 완료')

    def style_setting(self, filename):

        # 활성화
        ws = self.wb.worksheets[0]

        # B C D E F 열의 너비 15
        for col in range(1, 6):
            ws.column_dimensions[get_column_letter(col)].width = 15

        # E 일정 열너비 40 설정
        ws.column_dimensions['E'].width = 40

        # A 열너비 5 설정
        ws.column_dimensions['A'].width = 5

        # 주간업무계획표 폰트 크기 키우기, 굵게
        title_f = Font(name='맑은고딕', size=28, bold=True)
        # color='ff9999', strikethrough=True, underline='single', italic=True
        ws['B5'].font = title_f

        # 칼럼명 폰트 굵게
        font_range = ws['B8:F8']

        for col in font_range:
            for cell in col:
                cell.font = Font(bold=True)

        # 주간업무계획표, 기간 가운데 정렬
        ws['B5'].alignment = Alignment(horizontal='center', vertical='center')
        ws['B6'].alignment = Alignment(horizontal='center', vertical='center')

        # 칼럼명 가운데 정렬
        for col in font_range:
            for cell in col:
                cell.alignment = Alignment(horizontal='center', vertical='center')

        # 날짜, 요일 가운데 정렬
        for i in range(9, 40, 5):
            ws[f'B{i}'].alignment = Alignment(horizontal='center', vertical='center')
            ws[f'C{i}'].alignment = Alignment(horizontal='center', vertical='center')

        # # 담당자, 시작일, 컬럼명 색칠
        ws['B2'].fill = PatternFill(fgColor='E2EFDA', fill_type='solid')
        ws['B3'].fill = PatternFill(fgColor='E2EFDA', fill_type='solid')

        # 표 컬럼명 색칠 2번째 방법
        cell_range = ws['B8:F8']
        # print(cell_range)

        for col in cell_range:
            for cell in col:
                if cell.row == 8:  # 8행이면
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
        self.wb.save(filename)
        print('스타일 적용 완료')


if __name__ == '__main__':
    filename = '주간업무계획표.xlsx'
    create_excel_file(filename)
    wwp = WeeklyWorkPlan(filename)
    wwp.set_title(filename)
    wwp.insert_context(filename)
    wwp.style_setting(filename)
