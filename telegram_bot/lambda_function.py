import requests
from page_crawl import TelegramContent


def call():
    building_list = {
        '리버나인': 104311,
        '삼성 힐스테이트': 24008
    }
    final_result = """"""
    final_result += '삼성 힐스테이트' + '\n'
    final_result += f'https://m.land.naver.com/complex/info/{building_list["삼성 힐스테이트"]}?ptpNo=1' + '\n'
    tc = TelegramContent(building_list['삼성 힐스테이트'])
    final_result += tc.info_to_str()
    return final_result


def lambda_handler(event, context):
    to_str = call()
    data = requests.get(
        f"https://api.telegram.org/bot5415652060:AAFuzAh7AjTFFtjP6ilk43_2Y_d5YNVVlN8/sendMessage?chat_id=-657109651&text={to_str}")
    print(data.json())