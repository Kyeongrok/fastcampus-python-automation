from bs4 import BeautifulSoup
import re

with open('resources/email_template_1.html') as f:
    s = f.read()
    s = s.replace('업체명', 'client')
    s = s.replace('담당자명', 'manager')
    print(s)
    # html = BeautifulSoup(s, "html.parser")
    # # origin = html.find('body').get_text()
    # print(html.find('body').s)
    # print(type(html.find('body').get_text()))
    # body = re.search('<body. */body>', origin, re.I | re.S)
    # print(body)

