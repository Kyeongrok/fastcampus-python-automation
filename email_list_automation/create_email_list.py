import os
import pandas as pd
import datetime

class CreateEmailList:

    def __init__(self, filename, title, text, path='data/'):
        self.filename = filename #data_file
        self.title = title#제목
        self.text = text#본문
        self.path = path

        self.set_filenames()

    # filenm_list만들기
    def set_filenames(self):
        # 전체 칼럼 보기 설정
        pd.set_option('display.max_columns', None)

        # 협력사 data 불러오기
        df_partners = pd.read_excel(self.filename, engine='openpyxl')

        # 브랜드 null값 삭제
        df_partners = df_partners.dropna(subset=['브랜드'])

        self.partners = df_partners#make_email_list에서 변수로 활용

        # data폴더 파일 이름 목록 불러오기
        file_list = os.listdir(self.path)#경로

        # 확장자명 제외한 이름 출력
        email_infos = []
        for file in file_list:
            name = file.split('.')[0]
            partner_name = name.split(' ')[-1]
            # partner_name으로 파트너 이메일, 참조 찾기
            found_row = df_partners[df_partners['업체명'].str.contains(partner_name)]
            email1 = str(found_row['이메일1'].values[0])
            email_cc = str(found_row['참조이메일'].values[0])
            email_infos.append({'filename':file,
                                'partner_email':email1,
                                'cc':email_cc, 'title':self.title,
                                'text':self.text})
            if found_row.empty:
                print(f'{partner_name}이 파트너 목록에 없습니다.')

        self.result = email_infos


    def make_email_list(self, filename='email_list2.xlsx'):

        # df 생성(recipient, title, text, attachment)
        email_list = pd.DataFrame(self.result)

        # # 인덱스 컬럼 없이 값만 엑셀 저장
        email_list.to_excel(filename, index=False, header=False)
        print(f'{filename}으로 이메일 전송목록 생성 완료')

now = datetime.datetime.now()  # 지금시간
nowToday = now.strftime('%m/%d')  # 일자

# 인스턴스 생성
ce = CreateEmailList('파트너목록.xlsx',
    '[패스트몰] BENTZ FAST MALL ' + f'{nowToday} 상품발주 확인요청의 件', '''\
    안녕하세요.
    패스트몰 000입니다.
    금일자로 주문건 접수되어 출고요청 전달드립니다.
    확인부탁드리겠습니다.
    항상 많은 도움주셔서 감사드립니다.
    000 드림
    ''', 'data/')
# # .\\listOfPartners.xlsx파일은 자동입력, 이메일 본문과 제목만 외부에서 정보 입력받기
# # 파트너사 정보를 수정할 일이 있으면 listOfPartners.xlsx 파일을 수정하면 됨
#
# # 메소드 호출
ce.make_email_list()
