# -*- coding: utf-8 -*-
import smtplib
import time
import csv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import config

class Parse(object):

    def __init__(self, name):
        self.file_name = name

    def get_emails(self):
        res = []
        with open(self.file_name) as file_csv:
            read = csv.reader(file_csv)
            for row in read:
                if len(row) > 3 and row[4]  != "":
                    res.append((row[1], row[4].split()[0]))
        return res

class EmailSender():
    def __init__(self, login, password, msg_template):
        self.login = login
        self.password = password
        self.msg_template = msg_template

        self._init_smtp()


    def _init_smtp(self):
        self.smtp = smtplib.SMTP('smtp.gmail.com', 587)
        self.smtp.starttls()
        self.smtp.login(self.login, self.password)

    def parse_template(self, user_name):
        return self.msg_template.format(user_name)

    def send_email(self, email_to, msg):
        message = self.generate_msg({
            "from": self.login,
            "to": email_to,
            "subject": "Предложение о сотрудничестве компания Kvando Technologies",
            "message": msg})
        self.smtp.sendmail(self.login, email_to, message)

    def generate_msg(self, letter):
        msg = MIMEMultipart()
        msg['From'] = f"{letter['from']}\n"
        msg['To'] = f"{letter['to']}\n"
        msg['Subject'] = f"{letter['subject']}\n"
        msg.attach(MIMEText(letter['message'], "plain"))
        return msg.as_string()

    def __del__(self):
        try:
            self.smtp.quit()
        except Exception as e:
            print(e)

if __name__ == '__main__':
   p = Parse("email.csv")
   email = EmailSender(config.SMTP_EMAIL, config.SMTP_PASSWORD, config.TEMPLATE_EMAIL)
    # [("khigor777@yandex.ru", "Игорь")]
   for v in p.get_emails():
       msg = email.parse_template(v[1])
       try:
            email.send_email(v[0], msg)
       except Exception as e:
            raise Exception(e)
       print(v[0], v[1])
       time.sleep(10)

