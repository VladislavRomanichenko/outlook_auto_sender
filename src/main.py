from outlook import Outlook
import pandas as pd
import time
import sys

file_path = r"ПУТЬ ДО ФАЙЛА"

df = pd.read_excel(file_path)

for email_cnt in range(3, 5000):
    email_column = df.iloc[email_cnt, 3]

    if isinstance(email_column, str):
        email_column = email_column.split(",")
        for email in email_column:
            if email.find("@") != -1:
                email = email.strip()
                print(email)

outlook = Outlook()
while True:
    try:
        subject = "Тема сообщения"
        message = "Содержание сообщения"
        recipient = "Получатель сообщения"

        outlook.send_mail(message, subject, recipient)

        print('Pausing for 5 seconds.')
        time.sleep(5)

    except KeyboardInterrupt as e:
        outlook.close()
        sys.exit()