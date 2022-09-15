import time

from requests import get
from bs4 import BeautifulSoup
from urllib.parse import quote_plus


class TelegramContent:
    def __init__(self, code):
        self.search_key = code

    def get_estate_crawl(self):
        """
        #rletTypeCd: A01=아파트, A02=오피스텔, B01=분양권, 주택=C03, 토지=E03, 원룸=C01, 상가=D02, 사무실=D01, 공장=E02, 재개발=F01, 건물=D03
        # tradeTypeCd (거래종류): all=전체, A1=매매, B1=전세, B2=월세, B3=단기임대
        # hscpTypeCd (매물종류): 아파트=A01, 주상복합=A03, 재건축=A04 (복수 선택 가능)
        # cortarNo(법정동코드): (예: 1168010600 서울시, 강남구, 대치동)
        :param source:
        :return:
        """
        url = f"https://m.land.naver.com/complex/getPriceInfoChangedBySpc?hscpNo={self.search_key}&ptpNo=1&ptpNoForRealPrice=1&tradTpCd=&maxTradYm=99991231&addedItemTotalCnt=0"
        header = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.220 Whale/1.3.51.7 Safari/537.36',
            'Referer': 'https://m.land.naver.com/'
        }
        res = get(url, headers=header).json()
        try:
            trade_year = res['realPriceInfo']['tradeBssYearList']
            trade_list = res['realPriceInfo']['list']
            trade_contents =[]
            for item in trade_year:
                trade_tmp = [item]
                for keyword in trade_list:
                    if item == keyword['tradeBssYear']:
                        if list(keyword.values())[0][0]['tradTpCd'] == "A1":
                            tmp_dict = {'거래종류': '매매', '거래 일자': list(keyword.values())[0][0]['tradYm']+ list(keyword.values())[0][0]['tradeDate'], '가격': list(keyword.values())[0][0]['priceString']}
                            trade_tmp.append(tmp_dict)
                        elif list(keyword.values())[0][0]['tradTpCd'] == "B1":
                            tmp_dict = {'거래종류': '전세', '거래 일자': list(keyword.values())[0][0]['tradYm'] + list(keyword.values())[0][0]['tradeDate'], '가격': list(keyword.values())[0][0]['priceString']}
                            trade_tmp.append(tmp_dict)
                        elif list(keyword.values())[0][0]['tradTpCd'] == "B2":
                            tmp_dict = {'거래종류': '월세',
                                        '거래 일자': list(keyword.values())[0][0]['tradYm'] + list(keyword.values())[0][0][
                                            'tradeDate'], '가격': list(keyword.values())[0][0]['priceString']}
                            trade_tmp.append(tmp_dict)
                        elif list(keyword.values())[0][0]['tradTpCd'] == "B3":
                            tmp_dict = {'거래종류': '단기임대',
                                        '거래 일자': list(keyword.values())[0][0]['tradYm'] + list(keyword.values())[0][0][
                                            'tradeDate'], '가격': list(keyword.values())[0][0]['priceString']}
                            trade_tmp.append(tmp_dict)
                    else:
                        continue
                trade_contents.append(trade_tmp)
            return trade_contents
        except Exception as e:
            got_error = e
            print('실거래가 이뤄지지않았음')
            return e

    def info_to_str(self):
        data = self.get_estate_crawl()
        content = """"""
        try:
            for item in data:
                content += item[0] + "\n"
                for cont in item[1:]:
                    content += str(cont) + "\n"
            return content
        except Exception as e:
            print(e)
            return "실거래가 이뤄지지 않았습니다."


if __name__ == "__main__":
    building_list = {
        '리버나인': 104311,
        '삼성 힐스테이트': 24008
    }
    final_result = """"""
    final_result += '삼성 힐스테이트' + '\n'
    tc = TelegramContent(building_list['삼성 힐스테이트'])
    final_result += tc.info_to_str()
    print(final_result)



