import os
import pandas as pd
import datetime


# filenm_list만들기
def make_email_list(data_path, partners_filename, title, target_filename='email_list.xlsx'):
    # 전체 칼럼 보기 설정
    pd.set_option('display.max_columns', None)

    # 협력사 data 불러오기
    df_partners = pd.read_excel(partners_filename, engine='openpyxl')

    # data폴더 파일 이름 목록 불러오기
    file_list = os.listdir(data_path)  # 경로

    # 확장자명 제외한 이름 출력
    email_infos = []
    for file in file_list:
        name = file.split('.')[0]
        partner_name = name.split(' ')[-1]
        # partner_name으로 파트너 이메일, 참조 찾기
        found_row = df_partners[df_partners['업체명'].str.contains(partner_name)]
        email1 = str(found_row['이메일1'].values[0])
        email_cc = str(found_row['참조이메일'].values[0])
        partner_manager_name = str(found_row['컨택담당자'].values[0])
        email_infos.append({'partner_email': email1,
                            'cc': email_cc,
                            'partner_name':partner_manager_name,
                            'title': title,
                            'filename': file})
        if found_row.empty:
            print(f'{partner_name}이 파트너 목록에 없습니다.')

    email_list = pd.DataFrame(email_infos)
    email_list.to_excel(target_filename, index=False, header=False)
    print('파일 생성 완료')


if __name__ == '__main__':

    nowToday = datetime.datetime.now().strftime('%m/%d')  # 일자
    make_email_list('data/',
                    '파트너목록.xlsx',
                    '[패스트몰] BENTZ FAST MALL ' + f'{nowToday} 상품발주 확인요청의 件')
