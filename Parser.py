import win32com.client
import re, os, time, datetime, uuid

NO_REPLY = False  # отвечать автоматом или
IterationDelay = 60  # перерыв в секундах между повторным сканирование почты
#TerBanks = DSConfig.DSConfig.TerBanks


# Определим правило: обрабатывать ли сообщение или пропустить
# В примере простейшее правило, основанное на анализе темы
def myRule(subject):
    if ('SOME SPECIAL TEXT') in subject.upper():
        return 'Catched'
    return None


# Отправка сообщения
def sentReply(to_address, subject, to_name=''):
    # инициализируем объект outlook
    olk = win32com.client.Dispatch("Outlook.Application")
    Msg = olk.CreateItem(0)

    # формируем письма, выставляя адресата, тему и текст
    Msg.To = to_address
    Msg.Subject = "RE: " + subject  # добавляем RE в тему

    Msg.Body = "Здравствуйте, " + str(to_name) + " \n\n" \
                                                 "Ваше письмо получено!. \n" \
                                                 "\n" \
                                                 "С уважением,\n" \
                                                 "Черный С.Н.\n" \
                                                 "(Авто Ответ)"

    # и отправляем
    Msg.Send()


# запускаем бесконечный цикл проверки почты
while True:
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # шестая папка по умолчанию папка входящих писем
    inbox = outlook.GetDefaultFolder(6)

    # получаем список все входящих писем
    messages = inbox.Items
    for message in messages:
        move_it = False
        # тут хитрость, чтобы программа не падала. Почему иногда outlook глючит на некоторых письмах - хз
        try:
            message.To
            message.Subject
            message.Sender.Address
        except:
            continue

        # А здесь используем наше правило
        # усложнять его можно до бесконечности =)
        MR = myRule(message.Subject)
        if MR is None:
            continue

        # здесь можно обработать вложения
        for att in message.Attachments:
            try:
                f_name = att.FileName
            except:
                f_name = None
                continue
            is_lala_file = re.match('(.+)\.lala', att.FileName, re.M | re.I)
            if is_lala_file:
                move_it = True
                # saving every day in new directory
                # if Sunday or Saturday
                if time.strftime('%A') == 'Sunday':
                    t_time = datetime.date.today() + datetime.timedelta(1)
                elif time.strftime('%A') == 'Sunday':
                    t_time = datetime.date.today() + datetime.timedelta(1)
                else:
                    t_time = time
                date_dir = 'C:\\LALA\\' + t_time.strftime('%Y_%m_%d')
                # checking if directry exists
                if not os.path.isdir(date_dir):
                    os.makedirs(date_dir)
                # вложение сохраняем с уникальным имененем
                keyFileName = date_dir + '\\' + str(uuid.uuid4()) + '___' + att.FileName
                att.SaveAsFile(keyFileName)

                # extracting ID
                keyIds.append(id_by_key(keyFileName))

        if move_it:
            # move
            message.Move(move_folder)
            # sending reply
            if NO_REPLY is False:
                sentReply(message.Sender.Address, message.Subject, to_name=message.Sender, TB=TB, key_ids=keyIds)
                print(message.Subject + "reply sent!")
                print(message.Sender)
            # break

    # делаем задержку и снова повторяем проход
    time.sleep(IterationDelay)