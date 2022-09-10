import requests
import page_crawl

codes = page_crawl.get_code()
result = []
for code in codes:
    dic = page_crawl.get_crawl(code)
    result.append(dic)
# print(result)
to_str = ''.join(str(e) for e in result)


def lambda_handler(event, context):
    data = requests.get(f"https://api.telegram.org/<botId>/sendMessage?chat_id=<chat_id>&text={to_str}")
    print(data.json())


lambda_handler()
