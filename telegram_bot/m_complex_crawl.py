from requests import get

# 매매
url_template = lambda trade_type:f'https://m.land.naver.com/complex/info/22627?tradTpCd={trade_type}&ptpNo=&bildNo=&articleListYN=Y'
# 전세

url = 'https://new.land.naver.com/api/complexes/overview/104311'


header = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.220 Whale/1.3.51.7 Safari/537.36',
        'Referer': 'https://m.land.naver.com/'
    }

{
    "22627":'잠실엘스'
}
url_real_price_info_tem = lambda complex_cd:f'https://m.land.naver.com/complex/getPriceInfoChangedBySpc?hscpNo={complex_cd}&ptpNo=1&ptpNoForRealPrice=1&tradTpCd=&maxTradYm=99991231&addedItemTotalCnt=0'

url = url_real_price_info_tem('24008')
print(url, get(url, headers=header).json())

