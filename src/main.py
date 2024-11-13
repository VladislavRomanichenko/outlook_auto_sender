from outlook import Outlook
import time
import sys

def main():
    outlook = Outlook()
    file_path = r"C:\Users\vlad\PycharmProjects\pythonProject\src\СПАРК_Выборка_компаний_20241111_1330_1.xlsx"
    subject = "Тема сообщения"
    message = "Содержание сообщения"

    try:
        mail_recipients = outlook.extract_mails(file_path=file_path, mail_cnt=5)

        for recipient in mail_recipients:
            print(f"---- Sending a message to the next mail -> {recipient} ----")
            outlook.send_message(message, subject, recipient)
            print("\n**** Pausing for 1 seconds. ****\n")
            time.sleep(1)

        outlook.close()

    except KeyboardInterrupt as e:
        outlook.close()
        sys.exit()


if __name__ == "__main__":
    main()