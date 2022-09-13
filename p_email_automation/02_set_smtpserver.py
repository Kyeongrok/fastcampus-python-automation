import smtplib
from email.message import EmailMessage
from os import getenv


class EmailSender:
    id = None
    password = None
    smtp_server = None
    smtp_server_map = {
        'gmail.com': 'smtp.gmail.com',
        'naver.com': 'smtp.naver.com',
        'outlook.com': 'smtp-mail.outlook.com'
    }

    def __init__(self, id, password):
        self.id = id
        self.password = password
        self.smtp_server = self.smtp_server_map[id.split('@')[1]]

    def send_email(self, to, cc, title, text):
        fr = self.id
        msg = es.make_email(to, cc, title, text)
        print(self.smtp_server)
        with smtplib.SMTP(self.smtp_server, 587) as smtp:
            # smtp.set_debuglevel(True)
            smtp.ehlo()
            smtp.starttls()
            smtp.login(self.id, self.password)
            smtp.sendmail(to_addrs=to, from_addr=fr, msg=msg.as_string())
            smtp.quit()

        print('발송 성공')

    def make_email(self, recipient, recipient2, title, text):

        msg = EmailMessage()

        # 보내는 사람 / 받는 사람 / 참조 이메일 / 제목 입력
        msg['From'] = self.id
        msg['To'] = recipient.split(',')
        if recipient2 is not None:#참조이메일이 비어 있지 않으면 (!=')라고 할 경우 error 발생 AttributeError: 'NoneType' object has no attribute 'split'
           msg['Cc'] = recipient2.split('/')#슬래시 구분해서 input
        msg['Subject'] = title

        # 본문 구성
        msg.set_content(text)

        return msg

if __name__ == '__main__':
    id = 'oceanfog1@gmail.com'
    id = 'oceanfog@naver.com'
    password = getenv('MY_EMAIL_PASSWORD')
    es = EmailSender(id, password)
    es.send_email('oceanfog1@gmail.com', 'oceanfog2@gmail.comg', '테스트', '내용 테스트')
