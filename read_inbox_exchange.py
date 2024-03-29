"""Скрипт получения почты с Exchange сервера.

При отправке писем на определенный адрес, скрипт получает эти письма, импортирует нужные атрибуты в базу данных

и удаляет письмо с сервера.
"""
from bs4 import BeautifulSoup

from ldap import GetNameFromLdap
import os
import re
from datetime import datetime
import time
import pymssql
import pytz
from exchangelib import DELEGATE, Account, Credentials, Configuration, Message, Mailbox, HTMLBody, FileAttachment, \
    ItemAttachment
import exchangelib.autodiscover
from requests_kerberos import HTTPKerberosAuth
import urllib3
from cfg import SERVEREXCHANGE, SERVERMSSQL, USERMSSQL, PASSWORDMSSQL, \
    DATABASEMSSQL, DIRECTORYATTACHMENTS, SERVERAD, USERAD, PASSWORDAD, CORP

tz = pytz.timezone('Europe/Moscow')


def body_parse(namefolder, item):
    path = f'{DIRECTORYATTACHMENTS}{namefolder}\\inline\\'
    img_lst = []
    for root, dirs, files in os.walk(path):
        for filename in files:
            img_lst.append(filename)

    soup = BeautifulSoup(item, 'html.parser')
    for index, img in enumerate(soup.findAll('img')):

        if img['src'][0:4] == 'http':
            pass
        else:
            # cid = img['src'][4:12] + '.png'

            cid = img_lst[index]
            img['src'] = f'http://10.1.0.31:9800/app_manager/attachments/{namefolder}/inline/{cid}'
    my_html_string = str(soup).replace("'", '')

    return my_html_string


def printer(data):
    log_file = open("HelpDesk_log.txt", "a")
    print(str(datetime.now()) + ' ' + str(data))
    log_file.write(str(datetime.now()) + ' ' + str(data) + '\n')
    log_file.close()
    return printer


def GetNameSpecialist():
    server = SERVERAD
    user = USERAD
    password = PASSWORDAD
    corp = CORP
    result_name = GetNameFromLdap(server, user, password, corp).Cnname()
    mydb = pymssql.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                           database=DATABASEMSSQL)
    mycursor = mydb.cursor()
    mycursor.execute(f"select uid_Division from Staff where employee = '{result_name}'")
    result = mycursor.fetchall()
    return result[0][0]


UID_DIVISION = GetNameSpecialist()


def Data_Division():
    mydb = pymssql.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                           database=DATABASEMSSQL)
    mycursor = mydb.cursor()
    mycursor.execute(f"SELECT * from Division")
    result = mycursor.fetchall()
    return result


# server = SERVEREXCHANGE
# email = EMAILADDRESS
# username = USEREXECHANGE
# password = USEREXECHANGEPASS

# account = connect(server, email, username, password)


class DirCreated(object):
    """Создание директории"""

    def __init__(self, emailDB_id):
        self.emailDB_id = str(emailDB_id)

    """Создание имени директории"""

    def Dir(self):
        dir = os.path.join(DIRECTORYATTACHMENTS, self.emailDB_id)
        if not os.path.exists(dir):
            os.mkdir(dir)
            return dir

    def InsertDirToMssql(self):
        return self.emailDB_id


class GetAttachments(object):
    """Получение вложений"""

    def __init__(self, item, directory):
        self.tab_2 = None
        self.pushButton_attach = None
        self.item = item
        self.directory = directory

    """Имя вложения"""

    # def FileAttachment(self):
    #     Local_path_list = []
    #     for attachment in self.item:
    #         # Проверка вложения на is_inline
    #         if not attachment.is_inline:  # Изменил 19.12.22 добавил - or attach.name == 'image001.png and len(self.item) > 6: '
    #             local_path = os.path.join(self.directory, attachment.name)
    #             Local_path_list.append(local_path)
    #
    #             with open(local_path, 'wb') as f:
    #                 f.write(attachment.content)
    #     f.close()
    #
    #     return Local_path_list

    def FileAttachment(self):
        Local_path_list = []
        for attachment in self.item:
            if isinstance(attachment, FileAttachment):
                local_path = os.path.join(self.directory, attachment.name)
                Local_path_list.append(local_path)
                with open(local_path, 'wb') as f:
                    f.write(attachment.content)

        f.close()

        return Local_path_list

    def AttachmentsName(self):
        attach_name_lst = []
        for attach in self.item:
            # if not attach.is_inline: # Изменил 19.12.22 добавил - or attach.name == 'image001.png'
            name = os.path.join(self.directory, attach.name)
            attach_name_lst.append(name)

        return attach_name_lst


def CheckCreateDir(attachments):
    for i in attachments:
        if not i.is_inline:
            return True
        else:
            return False


class ConnectToExchange(object):
    """docstring"""

    def __init__(self, server, email, username, password, account, uid_Divisioin):
        """Constructor"""

        self.server = server
        self.email = email
        self.username = username
        self.password = password
        self.account = self.connect()
        self.uid_Division = uid_Divisioin

    def auth_model(**kwargs):
        # get kerberos ticket
        return HTTPKerberosAuth()

    urllib3.disable_warnings()

    def connect(self):
        try:
            from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
            # Care! Ignor Exchange self-signed SSL cert
            BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter

            # fill Credential object with empty fields
            creds = Credentials(
                username=self.username,
                password=self.password
            )

            # add kerberos as GSSAPI auth_type
            # exchangelib.transport.AUTH_TYPE_MAP["GSSAPI"] = auth_model

            # Create Config
            config = Configuration(server=SERVEREXCHANGE, credentials=creds)
            # return result
            return Account(primary_smtp_address=email, autodiscover=False, config=config, access_type=DELEGATE)
        except Exception as e:
            printer(e)

    def StatusConnect(self):
        """Подключаемся к Exchange"""
        printer('CONNECT TO EXCHANGE -> ESTABLISHED')

    def SendInfoMessage(self, subject, body):

        try:
            acc = self.connect()

            cursor = conn_database.cursor()
            sql_select_query = cursor.execute(f"select email_Staff from staff where uid_Division = {UID_DIVISION}")
            result = cursor.fetchall()
            recipients = result
            to_recipients = []
            for recipient in recipients:
                to_recipients.append(Mailbox(email_address=recipient[0]))
            # Create message
            m = Message(account=acc,
                        subject=subject,
                        body=body,
                        to_recipients=to_recipients)

            m.send(save_copy=False)
        except Exception as s:
            printer(s)

    def GetItemAccountInbox(self):
        global cursor, conn_database
        # last_date = self.LastDate()
        # date_str = str(datetime.today())
        # date = date_str.replace(':', '')
        # date_str = datetime.today().strftime('%Y-%m-%d %H:%M:%S')
        # date = date_str.replace(':', '')
        for item in self.account.inbox.filter(is_read=False).order_by('-datetime_received'):
            date_str = datetime.today().strftime('%Y-%m-%d-%H:%M:%S')
            date = date_str.replace(':', '')
            try:
                match = re.findall(r'##\w+', item.body)
                print(match)
                ID = int(match[0].replace('##', ''))
                print(ID)
                direct_for_reply = str(ID)
                sender_id = 2
                if ID:  # если найден id письма в subject

                    conn_database = pymssql.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                                                    database=DATABASEMSSQL)
                    cursor = conn_database.cursor()
                    # html_body = body_parse(direct_for_reply, item.body)
                    # sql = f"""UPDATE email SET text_body = '{html_body}' WHERE id = '{ID}' """
                    # result_html = cursor.execute(sql)
                    txt = ''
                    find_indexes = ['[145]', 'From:']
                    for index in find_indexes:
                        ind = item.text_body.find(index)
                        if ind != -1:
                            txt = item.text_body[:ind]
                            break

                    sql_insert_blob_query = """INSERT INTO chat (chat_id, user_name, content, datetime_send,
                                                                                sender_id) VALUES (%s,%s,%s,%s,%s) """
                    insert_blob_tuple = (
                        ID, item.sender.name, txt, item.datetime_sent.astimezone(tz), sender_id)
                    cursor.execute(sql_insert_blob_query, insert_blob_tuple)

                    # sql_query = f"""UPDATE email SET is_chat = True WHERE id = '{ID}'"""
                    # cursor.execute(sql_query)
                    # Инфо письмо
                    subject_update = item.subject
                    body_update = f'Получен ответ по заявке http://10.1.2.129:22345/applications/message/{ID}\nТема: {subject_update}\nОписание: {item.text_body}'
                    # Отправка инфо письма
                    self.SendInfoMessage(subject_update, body_update)
                    printer(f'Ответ по заявке "{ID}" отправлен')
                    item.delete()

                    conn_database.commit()
                    cursor.close()
                    conn_database.close()



            except:

                open_order = False
                close_order = False
                uid_Division = self.uid_Division
                conn_database = pymssql.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                                                database=DATABASEMSSQL)
                cursor = conn_database.cursor()
                printer(f'Получаем автора заявки')
                cursor.execute(f"SELECT id from applications_orderauthor where author = '{item.sender.name}'")
                result_author = cursor.fetchall()
                if result_author:
                    printer(f'Автор заявки уже записан')
                    # Sql query
                    sql_insert_blob_query = """INSERT INTO email (subject, sender_name, sender_email, copy,
                                                                               datetime_send, yes_no_attach,text_body, recipients, uid_Division, open_order, close_order, html_body, author_id) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) """

                    insert_blob_tuple = (
                        item.subject, item.sender.name, item.sender.email_address, item.display_cc,
                        item.datetime_sent.astimezone(tz), item.has_attachments, item.body, item.display_to,
                        uid_Division, open_order, close_order, item.body, int(result_author[0][0]))
                    printer('author already in base')
                else:
                    printer('Автор не найден')
                    author_sql_quary = f"INSERT INTO applications_orderauthor(author, email) VALUES (%s,%s)"
                    author_blob_tuple = (item.sender.name, item.sender.email_address)
                    cursor.execute(author_sql_quary, author_blob_tuple)
                    printer('Новый автор записан в базу')
                    author_id = cursor.lastrowid

                    sql_insert_blob_query = """INSERT INTO email (subject, sender_name, sender_email, copy,
                                                                                                   datetime_send, yes_no_attach,text_body, recipients, uid_Division, open_order, close_order, html_body, author_id) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) """

                    insert_blob_tuple = (
                        item.subject, item.sender.name, item.sender.email_address, item.display_cc,
                        item.datetime_sent.astimezone(tz), item.has_attachments, item.body, item.display_to,
                        uid_Division, open_order, close_order, item.body, author_id)
                    print('author add in base')

                result = cursor.execute(sql_insert_blob_query, insert_blob_tuple)
                emailDB_id = cursor.lastrowid

                printer(f'"{item.subject}" :, Успешно занесено в базу')

                # Инфо письмо
                subject = f'Заявка от {item.sender.name}, {item.subject}'
                new_request = f'Новая заявка "{emailDB_id}" от {item.sender.name}!'.upper()
                body = f'{new_request}\nТема: {item.subject}\nСсылка - http://10.1.2.129:22345/applications/message/{emailDB_id}\nОписание:\n{item.text_body}'

                ############################################ Вложения!!
                # Директория для записи в базу
                direct = str(emailDB_id)
                dir_inline = os.path.join(DIRECTORYATTACHMENTS + direct + '\\inline')
                try:
                    if item.attachments:
                        directory = DirCreated(emailDB_id).Dir()
                        for attach in item.attachments:
                            if not attach.is_inline:
                                local_path = os.path.join(directory, attach.name)
                                with open(local_path, 'wb') as f:
                                    f.write(attach.content)
                                f.close()
                                # Insert Attachments in database
                                dir_to_sql = os.path.join(direct, attach.name)
                                add_Attachments = """INSERT INTO Attachments(link, id_email, name) VALUES (%s, %s, %s)"""
                                data_Attachments = (dir_to_sql, emailDB_id, attach.name)
                                cursor.execute(add_Attachments, data_Attachments)


                            elif attach.is_inline:

                                if not os.path.exists(dir_inline):
                                    os.mkdir(dir_inline)
                                with open(dir_inline + f'\\{attach.name}', 'wb') as w:
                                    w.write(attach.content)
                                w.close()
                        try:
                            html_body = body_parse(direct, item.body)
                            sql = f"""UPDATE email SET text_body = '{html_body}' WHERE id = '{emailDB_id}' """
                            result_html = cursor.execute(sql)
                        except Exception as e:
                            printer(e)
                        printer('Вложение определены и добавлены')
                except Exception as s:
                    printer(s)

                # Отправка инфо письма
                self.SendInfoMessage(subject, body)
                printer('info message send!')
                conn_database.commit()
                item.delete()
                printer('messege delete')
                printer("-----------")
                cursor.close()
                conn_database.close()


while True:
    try:
        account = ConnectToExchange.connect
        for i in Data_Division():
            uid_Division = i[0]
            email = i[2]
            username = i[3]
            password = i[4]
            conn = ConnectToExchange(SERVEREXCHANGE, email, username, password, account,
                                     uid_Division).GetItemAccountInbox()
            time.sleep(30)
    except KeyboardInterrupt as s:
        printer(s)
    except exchangelib.errors.ErrorFolderNotFound as error:
        printer(error)

    except Exception as e:
        printer(e)
