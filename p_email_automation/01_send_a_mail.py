import smtplib
from os import getenv


class EmailSender:
    id = None
    password = None

    def __init__(self, id, password):
        self.id = id
        self.password = password

    def send_email(self, msg, from_addr, to_addr):
        with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
            # smtp.set_debuglevel(True)
            smtp.ehlo()
            smtp.starttls()
            smtp.login(self.id, self.password)
            smtp.sendmail(to_addrs=to_addr, from_addr=from_addr, msg=msg)
            smtp.quit()

        print("발송 성공")


if __name__ == '__main__':
    id = 'oceanfog1@gmail.com'
    password = getenv('MY_EMAIL_PASSWORD')
    es = EmailSender(id, password)
    es.send_email('hello test emil', id, 'oceanfog1@gmail.com')
