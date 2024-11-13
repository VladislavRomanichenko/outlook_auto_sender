import win32com.client
import pandas as pd
import datetime
import logging

class Outlook(object):

    def __init__(self):
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.mapi = self.outlook.GetNamespace("MAPI")

        #Variable for counting sent messages
        self.cnt_messages = 0


    def extract_mails(self, file_path, mail_column_number = 3, mail_cnt = 100):
        df = pd.read_excel(file_path)
        column = mail_column_number
        mails = list()

        for row in range(3, mail_cnt):
            email_column = df.iloc[row, column]
            if isinstance(email_column, str):
                email_column = email_column.split(",")
                for mail in email_column:
                    if mail.find("@") != -1:
                        mail = mail.strip()
                        mails.append(mail)
        return mails


    def total_messages(self):
        return self.cnt_messages


    def send_message(self, message, subject, recipient_mail, sender_mail = "NO SENDER"):
        self.mail = self.outlook.CreateItem(0)

        #Recipient and sender definitions
        if sender_mail != "NO SENDER":
            self.mail.Sender = sender_mail
        self.mail.To = recipient_mail

        #Determining the subject and message
        self.mail.Subject = subject
        self.messageBody = message
        self.mail.Body = self.messageBody

        #Send message
        self.mail.Send()
        self.cnt_messages += 1


    def close(self):
        if self.total_messages() == 0:
            logging.info('\n---- No emails received. ----\n')
        logging.info("\n---- Outlook sender closed ----\n")

