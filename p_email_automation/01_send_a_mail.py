import smtplib
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

    def send_email(self, msg, fr, to):
        with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
            # smtp.set_debuglevel(True)
            smtp.ehlo()
            smtp.starttls()
            smtp.login(self.id, self.password)
            smtp.sendmail(to_addrs=to, from_addr=fr, msg=msg)
            smtp.quit()

        print("발송 성공")


if __name__ == '__main__':
    id = 'oceanfog1@gmail.com'
    password = getenv('MY_EMAIL_PASSWORD')
    es = EmailSender(id, password)
    es.send_email('hello test emil', id, 'oceanfog1@gmail.com')
