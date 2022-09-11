import pandas as pd
# 1. 브랜드만 있는 리스트
# 2. 업체명만 있는 리스트
# 3. 주문파일 데이터

'''
.iter_rows()로 반복문 돌리면서
    브랜드 개수만큼 돌려서
    브랜드만 있는 리스트에서 검색을 해서
'''

class ClassificationExcel:
    # Constructor
    def __init__(self, order_xlsx_filename, partner_info_xlsx_filename):
        self.path = './data/'

        # 주문목록 불러오기
        df = pd.read_excel(order_xlsx_filename, engine='openpyxl')

        # df파일 1번째 행을 칼럼으로 지정
        df = df.rename(columns=df.iloc[1])

        # 0~1행 삭제
        processed_df = df.drop([df.index[0], df.index[1]])

        # 인덱스 재설정
        self.order_list = processed_df.reset_index(drop=True)  # 가공한 df파일, #find_product에서 활용


        # 파트너목록 불러오기
        df_partners_info = pd.read_excel(partner_info_xlsx_filename, engine='openpyxl')

        # # 브랜드에 값이 비어있다면 제거
        # df_partners_info.dropna(axis='index', how='any', subset=['브랜드'])

        # 브랜드 list 생성
        self.brands = df_partners_info['브랜드'].tolist()

        # 업체명 list 생성
        self.partners = df_partners_info['업체명'].tolist()



    def classify(self):
        print(len(self.brands), self.brands)
        print(len(self.partners), self.partners)

        # for i, row in self.order_list.iterrows():
        #     print(i)
        print(self.order_list)
        print(self.order_list.count())


if __name__ == '__main__':
    ce = ClassificationExcel('주문목록20221112.xlsx', '파트너목록.xlsx')
    ce.classify()