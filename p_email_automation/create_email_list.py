import os
import pandas as pd
import datetime

class CreateEmailList:

    def __init__(self, df, title, text, path='data/'):
        self.df = df#data_file
        self.title = title#제목
        self.text = text#본문
        self.path = path

    # filenm_list만들기
    def make_filenm_list(self):
        # 전체 칼럼 보기 설정
        pd.set_option('display.max_columns', None)

        # 협력사 data 불러오기
        partners = pd.read_excel(self.df, engine='openpyxl')
        # print(partners)

        # 브랜드 null값 삭제
        partners = partners.dropna(subset=['브랜드'])

        # null값 확인
        # print(partners.isnull().sum())

        self.partners = partners#make_email_list에서 변수로 활용

        # data폴더 파일 이름 목록 불러오기
        file_list = os.listdir(self.path)#경로
        # print(file_list)#attachment에 확장자명까지 기입

        # 확장자명 제외한 이름 출력
        file_name = []
        for file in file_list:
            if file.count(".")==1:
                name= file.split('.')[0]
                file_name.append(name)
            else:
                for i in range(len(file)-1, 0, -1):
                    if file[i]=='.':
                        file_name.append(file[:i])
                        break
        # print(file_name)
        # print(len(file_name))#22개

        self.file_name = file_name#make_email_list에서 변수로 활용
        self.file_list = file_list#make_email_list에서 변수로 활용

    def make_email_list(self):
        # email_list df 만들기
        # 불러온 파일 업체명이 partners의 '브랜드' 업체명과 일치하면
        # 이메일1 -> recipient 변수명에 저장
        recipient= []
        recipient2 =[]
        # 어차피 file_name 갯수만 찾아야하므로 다돌릴 필요가 없음
        for i, row in self.partners[:len(self.file_name)].iterrows():#file_name 갯수만큼 돌리기
            # print(row)#0~5
            # print(row['참조이메일'])
            partners_name = self.file_name[i][24:]#파일명 뒤에 브랜드명만 나오게 하기
            # print(partners_name)#업체명만
            find_brand = self.partners.loc[self.partners['업체명'].str.contains(partners_name)]#filenm과 일치하는 '업체명'칼럼 df 출력, 변수 활용
            find_brand_email = find_brand['이메일1'].values.tolist()#이메일1만 추출
            if len(find_brand_email) == 0:
                find_brand_email.append('없음')
                print(find_brand_email)
                recipient.append(find_brand_email[0])
            else:
                recipient.append(find_brand_email[0])#값만 추출해서 담기
        # print(recipient)
            find_brand_CC_email = find_brand['참조이메일'].tolist()#참조이메일 값만 추출
            print(find_brand_CC_email)
            if len(find_brand_CC_email) == 0:
                find_brand_CC_email.append('없음')
                recipient2.append(find_brand_CC_email[0])
            else:
                recipient2.append(find_brand_CC_email[0])
        print(recipient) #이메일
        print(recipient2) #참조이메일

        # 불러온 파일명은 그대로 첨부파일 명에 기재
        attachment = self.file_list#변수 활용

        # df 생성(recipient, title, text, attachment)
        table_name = {
            'recipient': recipient,#수신자
            'recipient2': recipient2,#참조이메일
            'title': self.title,# title 동일
            'text': self.text,# text 동일
            'attachment': attachment
        }
        email_list = pd.DataFrame(table_name)
        # print(email_list)


        # 인덱스 컬럼 없이 값만 엑셀 저장
        email_list.to_excel('email_list.xlsx', index=False, header=False)
        print('파일 생성 완료')

now = datetime.datetime.now()  # 지금시간
nowToday = now.strftime('%m/%d')  # 일자

# 인스턴스 생성
ce = CreateEmailList('파트너목록.xlsx', '[패스트몰] BENTZ FAST MALL ' + f'{nowToday} 상품발주 확인요청의 件', '''\
    안녕하세요.
    패스트몰 000입니다.
    금일자로 주문건 접수되어 출고요청 전달드립니다.
    확인부탁드리겠습니다.
    항상 많은 도움주셔서 감사드립니다.
    000 드림
    ''', '20221113/')
# # .\\listOfPartners.xlsx파일은 자동입력, 이메일 본문과 제목만 외부에서 정보 입력받기
# # 파트너사 정보를 수정할 일이 있으면 listOfPartners.xlsx 파일을 수정하면 됨
#
# # 메소드 호출
ce.make_filenm_list()
ce.make_email_list()
