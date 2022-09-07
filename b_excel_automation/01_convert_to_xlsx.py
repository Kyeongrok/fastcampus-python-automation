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
    def __init__(self, order_list_filename, vendor_list_filename):
        # 주문목록, 벤더리스트 불러오기
        df_order_list = pd.read_excel(order_list_filename)
        df_vendor_list = pd.read_excel(vendor_list_filename)


