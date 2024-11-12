import datetime
import logging
import win32com.client

class Outlook(object):

    def __init__(self):
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.mapi = self.outlook.GetNamespace("MAPI")

        #Подсчёт отправленных писем
        self.cnt_messages = 0


    def get_total_messages(self):
        return self.cnt_messages


    def send_mail(self, message, subject, recipient, sender = "NO SENDER"):
        self.mail = self.outlook.CreateItem(0)

        #Определения получателя и отправителя
        if sender != "NO SENDER":
            self.mail.Sender = sender
        self.mail.To = recipient

        #Определение темы и содержания сообщения
        self.mail.Subject = subject
        self.messageBody = message
        self.mail.Body = self.messageBody

        #Отправления сообщения
        self.mail.Send()


    def close(self):
        if self.get_total_messages() == 0:
            logging.info('* No emails received.')
        logging.info("\n----Outlook sender closed----\n")

