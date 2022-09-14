import requests
import page_crawl

def call():
    codes = page_crawl.get_finance_code()
    result = []
    for code in codes:
        dic = page_crawl.get_finance_crawl(code)
        result.append(dic)
    # print(result)
    to_str = ''.join(str(e) for e in result)

    data = requests.get(f"https://api.telegram.org/<botId>/sendMessage?chat_id=<chat_id>&text={to_str}")
    print(data.json())


def lambda_handler(event, context):
    call()


