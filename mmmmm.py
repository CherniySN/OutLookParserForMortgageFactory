import win32com.client
import os

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 6 папка можно изменить на другую
print('Папка: ',inbox)
# получаем список все входящих писем
messages = inbox.Items

for messag in messages:

    try:
        messub = messag.subject
        messendname = messag.SenderName
        messendaddres = messag.Sender.Address
        messbody = messag.body

        print('Тема письма: ',messub)
        print('Имя отправителя: ', messendname)
        print('Адрес отправителя: ', messendaddres)
        print('Содержание письма: ', messbody)

        for att in messag.Attachments:
            print(att)
            att.SaveASFile(os.path.join('C:/Users/pc/Downloads/forABBmailsorted/' + att.FileName))
            print("Mail Successfully Extracted")
    except:
        continue