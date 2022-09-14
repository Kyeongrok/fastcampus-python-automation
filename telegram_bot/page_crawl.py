import time

import requests
from bs4 import BeautifulSoup
from urllib.parse import quote_plus


def get_estate_url(area):
    url = f'https://m.land.naver.com/search/result/{quote_plus(area)}'
    header = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.220 Whale/1.3.51.7 Safari/537.36',
        'Referer': 'https://m.land.naver.com/'
    }
    res = requests.get(url, headers=header).text
    bs = BeautifulSoup(res, "html.parser")
    page = bs.select("#complex_list_ul > li > a")
    results = []
    for item in page:
        print(item['href'])
        get_estate_crawl(item['href'])
        time.sleep(3)


def get_estate_crawl(source):
    """
    #rletTypeCd: A01=아파트, A02=오피스텔, B01=분양권, 주택=C03, 토지=E03, 원룸=C01, 상가=D02, 사무실=D01, 공장=E02, 재개발=F01, 건물=D03
    # tradeTypeCd (거래종류): all=전체, A1=매매, B1=전세, B2=월세, B3=단기임대
    # hscpTypeCd (매물종류): 아파트=A01, 주상복합=A03, 재건축=A04 (복수 선택 가능)
    # cortarNo(법정동코드): (예: 1168010600 서울시, 강남구, 대치동)
    :param source:
    :return:
    """
    url = f"https://m.land.naver.com{source}"
    # url = "https://m.land.naver.com/complex/info/453?ptpNo=1"
    header = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.220 Whale/1.3.51.7 Safari/537.36',
        'Referer': 'https://m.land.naver.com/'
    }
    res = requests.get(url, headers=header).text
    bsobj = BeautifulSoup(res, "html.parser")
    recent_price = bsobj.select_one(".complex_price--trade > .data").get_text()
    recent_date = bsobj.select_one(".complex_price--trade > .date").get_text()
    av_price = bsobj.select(".complex_price_wrap > dl")
    av_trade_price = ""
    av_charter_price = ""
    build_name = bsobj.select_one(".detail_complex_title").get_text()
    for item in av_price:
        if item.select_one(".title").get_text() == "매매가":
            av_trade_price = item.select_one(".data").get_text()
        elif item.select_one(".title").get_text() == "전세가":
            av_charter_price = item.select_one(".data").get_text()
    result = {"건물 이름": build_name, "최근 매매가": recent_price, "최근 매매일": recent_date, "평균 매매가": av_trade_price, "평균 전세가": av_charter_price}
    print(result)


def get_finance_crawl(url):
    url = f"https://finance.naver.com{url}"
    res = requests.get(url).text
    bsobj = BeautifulSoup(res, "html.parser")
    div_today = bsobj.find("div", {"class": "today"})
    em = div_today.find("em")

    price = em.find("span", {"class": "blind"}).text
    h_company = bsobj.find("div", {"class": "h_company"})
    name = h_company.a.text
    div_description = h_company.find("div", {"class": "description"})
    code = div_description.span.text

    table_no_info = bsobj.find("table", {"class": "no_info"})
    tds = table_no_info.tr.find_all("td")
    volume = tds[2].find("span", {"class": "blind"}).text

    dic = {"종목명": name, "종목 코드": code, "시가": price, "거래량": volume}
    return dic


def get_finance_code():
    url = 'https://finance.naver.com/'
    res = requests.get(url).text
    bsobj = BeautifulSoup(res, "html.parser")
    table = bsobj.select('#_topItems1 > tr > th > a')
    codes = []
    for item in table:
        codes.append(item["href"])
    return codes


if __name__ == "__main__":
    get_estate_url('반포')



