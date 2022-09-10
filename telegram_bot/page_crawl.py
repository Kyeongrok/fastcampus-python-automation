import requests
from bs4 import BeautifulSoup


def get_crawl(url):
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


def get_code():
    url = 'https://finance.naver.com/'
    res = requests.get(url).text
    bsobj = BeautifulSoup(res, "html.parser")
    table = bsobj.select('#_topItems1 > tr > th > a')
    codes = []
    for item in table:
        codes.append(item["href"])
    return codes



