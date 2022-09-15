from page_crawl import TelegramContent


def call():
    tc = TelegramContent('24008')
    print(tc.info_to_str())


# def lambda_handler(event, context):
#     call()


