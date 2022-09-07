import pandas as pd
import datetime
import os
from openpyxl import workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill #엑셀 스타일

class ClassificationExcel:
    #Constructor
    def __init__(self, order_xlsx_filename, partner_info_xlsx_filename):
        self.order_xlsx_filename = order_xlsx_filename #주문파일(1page)
        self.df2 = partner_info_xlsx_filename #파트너사정보파일
        self.path = './data\\'

    # brands, partners 리스트, dict 만드는 기능
    def make_product_dict(self):

        # sendRequest 불러오기
        df = pd.read_excel(self.order_xlsx_filename, engine='openpyxl')

        # df파일 1번째 행을 칼럼으로 지정
        processed_df = df.rename(columns=df.iloc[1])

        # 0~1행 삭제
        processed_df = processed_df.drop([processed_df.index[0], processed_df.index[1]])
        # print(processed_df)

        # 인덱스 재설정
        processed_df = processed_df.reset_index(drop=True)
        # print(processed_df)

        # 상품명
        # product_name = processed_df['상품명']
        # print(product_name)

        # processed_df '상품명' partners_dict key와 부분일치 문자열 행 데이터 불러오기
        # print(partners[0])

        # 예시 검색어 부분일치 조회
        # baseus = processed_df.loc[df2['상품명'].str.contains('베이스어스')]
        # print(baseus['상품명'])

        self.processed_df = processed_df #가공한 df파일, #find_product에서 활용


        # 브랜드/업체명 파일 불러오기
        partners_name = pd.read_excel(self.df2, engine='openpyxl')

        # '브랜드' 칼럼 결측치있는 행 전체 제거
        partners_name = partners_name.dropna(axis = 'index', how='any', subset=['브랜드'])

        # 브랜드 list 생성
        brands_parsing = partners_name['브랜드'].str.split('/')
        # print(brands_parsing)
        brands = brands_parsing.tolist()

        # 업체명 list 생성
        partners = partners_name['업체명'].tolist()

        # 결측치 확인
        # print(partners_name[partners_name.partners.isnull()])

        # dict 변환
        # partners_dict = dict(zip(brands, partners))
        # for key in partners_dict:
        # print(key, ':', partners_dict[key])# key,value 출력

        self.brands = brands # find_product에서 활용
        self.partners = partners # find_product에서 활용

    # 각 상품명에 brands 요소가 있는지 모두 확인하여 partners 이름으로 각각 excel 저장 기능
    def find_product(self):
        # data폴더 자동 생성 기능 추가하기

        now = datetime.datetime.now()  # 지금시간
        self.nowToday = now.strftime('%Y-%m-%d')  # 일자

        for i, row in self.processed_df.iterrows():#엑셀 주문 건수 갯수만큼 돌리기
           # print(row)#모든 주문건수 출력(세로)
           # 문제: 2가지 브랜드가 들어간 업체명 추출이 안됨, 각 브랜드 주문건수 1건만 들어감
           for j in range(len(self.brands)):#brands의 갯수만큼 돌리기
               if len(self.brands[j]) ==1:#브랜드명 요소가 1개이면
                   # if row['상품명'].startswith(self.brands[j][0]):
                   if self.brands[j][0] in row['상품명']:  # 전체 df 상품명에서 brands 값명이 있으면, 더클래스가 '차바치 더클래스'가 있는 주문을 제외하고
                       first_df = self.processed_df.loc[self.processed_df['상품명'].str.startswith(self.brands[j][0])]
                       print(first_df)#여기까지는 모든 data가 나옴
                       if len(first_df) !=0:  # df가 비어 있지 않으면
                           self.partners[j] == first_df #파트너사 변수와 1:1 맵핑
                           first_df.to_excel(f'{self.path}[웍스컴바인]발주요청서_{self.nowToday}_{self.partners[j]}.xlsx',
                                                index=False)  # partners 이름으로 excel 저장, index없이 저장
               elif len(self.brands[j]) ==2:#브랜드 요소가 2개이상이면
                   # print(self.brands[j][1])  # 리스트의 두번째 요소만 뽑기
                   if self.brands[j][0] in row['상품명'] or self.brands[j][1] in row['상품명']:  # 전체 df 상품명에서 brands 값명이 있으면
                       second_df = self.processed_df.loc[self.processed_df['상품명'].str.startswith(self.brands[j][0])]#0번째 요소 일치한 df 출력
                       third_df = self.processed_df.loc[self.processed_df['상품명'].str.startswith(self.brands[j][1])]#1번째 요소 일치한 df 출력
                       sum_df = pd.concat([second_df, third_df]) #2개 데이터프레임 병합
                       if len(sum_df) != 0:  # df가 비어 있지 않으면
                           self.partners[j] == sum_df  # 파트너사 변수와 1:1 맵핑
                           sum_df.to_excel(f'{self.path}[웍스컴바인]발주요청서_{self.nowToday}_{self.partners[j]}.xlsx',
                                                index=False)  # partners 이름으로 excel 저장, index없이 저장
               elif len(self.brands[j]) ==3:#브랜드 요소가 3개이상이면
               #     print(self.brands[j][2])  # 리스트의 세번째 요소만 뽑기
                   if self.brands[j][0] in row['상품명'] or self.brands[j][1] in row['상품명'] or self.brands[j][2] in row['상품명']:  # 전체 df 상품명에서 brands 값명이 있으면
                       four_df = self.processed_df.loc[self.processed_df['상품명'].str.startswith(self.brands[j][0])]#0번째 요소 일치한 df 출력
                       fifth_df = self.processed_df.loc[self.processed_df['상품명'].str.startswith(self.brands[j][1])]#1번째 요소 일치한 df 출력
                       sixth_df = self.processed_df.loc[self.processed_df['상품명'].str.startswith(self.brands[j][2])]#2번째 요소 일치한 df 출력
                       sum2_df = pd.concat([four_df, fifth_df, sixth_df])#3개 데이터프레임 병합
                       print(sum2_df)
                       if len(sum2_df) != 0:  # df가 비어 있지 않으면
                           self.partners[j] == sum2_df  # 파트너사 변수와 1:1 맵핑
                           sum2_df.to_excel(f'{self.path}[웍스컴바인]발주요청서_{self.nowToday}_{self.partners[j]}.xlsx',
                                                index=False)  # partners 이름으로 excel 저장, index없이 저장

    # 저장한 주문파일 엑셀 폼 설정
    def set_excel_form(self):
        # 폴더 안의 엑셀 파일 하나씩 불러오기

        file_list = os.listdir(self.path)#path폴더에 있는 파일을 리스트로 받기
        print(file_list)

        for file_name_raw in file_list:
            file_name =  f'{self.path}'+file_name_raw
            wb = load_workbook(filename=file_name)#엑셀파일 가져오기
            ws = wb.active #활성화

            # A1내용 삽입
            m_row = ws.max_row - 1  # 주문 건수 세기(칼럼행 제외)
            print(m_row)
            # 1,2행을 새로 삽입한다
            ws.insert_rows(1)
            ws.insert_rows(2)
            ws['A1'].value = f'발송요청내역 [총 {m_row}건, 1페이지] {self.nowToday}' # 전체 행 개수를 세서 num 변수에 삽입
            f = Font(size=11, bold=True)# 1행 폰트 크기 11, bold
            ws['A1'].font = f
            # # A열 설정
            ws.merge_cells('A1:U1') #셀병합
            ws['A1'].alignment = Alignment(horizontal='left') #왼쪽 정렬
            f2 = Font(size=9)

            # 3, 4행 설정
            for row in ws.rows:
                for cell in row:
                    if cell.row != 1 and cell.row != 2:  # 1,2행 제외
                        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True) # 셀의 너비에 맞게 자동 줄바꿈, 모두 가운데 정렬
                        cell.fill = PatternFill(fgColor='ffffcc', fill_type='solid')#노란색 채우기
                        cell.font = Font(size=9)#3행부터 폰트크기 9
                        cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                        if cell.row == 3: #3행이면
                            cell.alignment = Alignment(horizontal='center', vertical='center', shrink_to_fit=True, wrap_text=False)#줄바꿈 제외, 셀 크기에 맞게 글자 축소
                            cell.font = Font(bold=True)  # 볼드체
                            cell.fill = PatternFill(fgColor='c0c0c0', fill_type='solid')#회색 채우기

            # 열 너비 설정
            columns_15 = [1,21]
            columns_35 = [9,10,20]
            columns_10 = [2,3,4,5,6,7,8,11,12,13,14,15,16,17,18,19]
            for col in range(len(columns_15)):
                ws.column_dimensions[get_column_letter(columns_15[col])].width = 17#A, U, 15로 하면 14.38이 됨
            for col in range(len(columns_35)):
                ws.column_dimensions[get_column_letter(columns_35[col])].width = 38
            for col in range(len(columns_10)):
                ws.column_dimensions[get_column_letter(columns_10[col])].width = 12
                #wb.save(f'{self.path}[웍스컴바인]발주요청서_2022-02-24_AUTOCOS.xlsx')#예시
            wb.save(f'{self.path}'+file_name_raw)
            print('엑셀폼 변경 완료')

#인스턴스 생성

CE = ClassificationExcel('주문목록20221112.xlsx', '벤더리스트.xlsx')#sendRequest.xlsx, listOfPartners_name.xlsx->listOfPartners 자동입력
#
# #메소드 호출
CE.make_product_dict()
CE.find_product()
CE.set_excel_form()
