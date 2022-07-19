"""Скрипт получения почты с Exchange сервера.

При отправке писем на определенный адрес, скрипт получает эти письма, импортирует нужные атрибуты в базу данных

и удаляет письмо с сервера.
"""

import os
from datetime import datetime
import time
import pymssql
import pytz
from exchangelib import DELEGATE, Account, Credentials, Configuration
import exchangelib.autodiscover
from requests_kerberos import HTTPKerberosAuth
import urllib3

from cfg import SERVEREXCHANGE, EMAILADDRESS, USEREXECHANGE, USEREXECHANGEPASS, SERVERMSSQL, USERMSSQL, PASSWORDMSSQL, DATABASEMSSQL, DIRECTORYATTACHMENTS

tz = pytz.timezone('Europe/Moscow')


def auth_model(**kwargs):
    # get kerberos ticket
    return HTTPKerberosAuth()


urllib3.disable_warnings()


def connect(server, email, username, password):
    try:
        from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
        # Care! Ignor Exchange self-signed SSL cert
        BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter

        # fill Credential object with empty fields
        creds = Credentials(
            username=username,
            password=password
        )

        # add kerberos as GSSAPI auth_type
        #exchangelib.transport.AUTH_TYPE_MAP["GSSAPI"] = auth_model

        # Create Config
        config = Configuration(server=server, credentials=creds)
        # return result
        return Account(primary_smtp_address=email, autodiscover=False, config=config, access_type=DELEGATE)
    except Exception as e:
        print(e)


server = SERVEREXCHANGE
email = EMAILADDRESS
username = USEREXECHANGE
password = USEREXECHANGEPASS
account = connect(server, email, username, password)


class DirCreated(object):
    """Создание директории"""

    def __init__(self, sender_name, date):
        self.sender_name = sender_name
        self.date = date

    """Создание имени директории"""

    def Dir(self):
        dir = os.path.join(DIRECTORYATTACHMENTS, self.sender_name + self.date)
        if not os.path.exists(dir):
            os.mkdir(dir)
            return dir


class GetAttachments(object):
    """Получение вложений"""

    def __init__(self, item, directory):
        self.tab_2 = None
        self.pushButton_attach = None
        self.item = item
        self.directory = directory

    """Имя вложения"""

    def FileAttachment(self):
        Local_path_list = []
        for attachment in self.item:
            # Проверка вложения на is_inline
            if not attachment.is_inline:
                local_path = os.path.join(self.directory, attachment.name)
                Local_path_list.append(local_path)
                with open(local_path, 'wb') as f:
                    f.write(attachment.content)
        return Local_path_list


def CheckCreateDir(attachments):
    for i in attachments:
        if not i.is_inline:
            return True
        else:
            return False


class ConnectToExchange(object):
    """docstring"""

    def __init__(self, server, email, username, account):
        """Constructor"""

        self.server = server
        self.email = email
        self.username = username
        self.account = account

    def StatusConnect(self):
        """Подключаемся к Exchange"""
        print('CONNECT TO EXCHANGE -> ESTABLISHED')

    def LastDate(self):
        try:
            conn_database = pymssql.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                                            database=DATABASEMSSQL)
            cursor = conn_database.cursor()
            sql_select_query = cursor.execute("select MAX(datetime_send) from email")
            results = cursor.fetchall()
            last_date = results
            return str(last_date[0][0])
        except IndexError:
            print('list index out of range')

    def GetItemAccountInbox(self):
        global cursor, conn_database
        last_date = self.LastDate()
        # date_str = str(datetime.today())
        # date = date_str.replace(':', '')
        date_str = datetime.today().strftime('%Y-%m-%d %H:%M:%S')
        date = date_str.replace(':', '')
        for item in self.account.inbox.filter(is_read=False).order_by('-datetime_received'):
            print(self.StatusConnect())

            try:
                conn_database = pymssql.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                                                database=DATABASEMSSQL)
                cursor = conn_database.cursor()
                # Sql query
                sql_insert_blob_query = """INSERT INTO email (subject, sender_name, sender_email, copy, 
                    datetime_send, yes_no_attach, text_body, recipients) VALUES (%s,%s,%s,%s,%s,%s,%s,%s) """

                # Convert data into tuple format
                insert_blob_tuple = (
                    item.subject, item.sender.name, item.sender.email_address, item.display_cc,
                    item.datetime_sent.astimezone(tz), item.has_attachments, item.body, item.display_to)
                result = cursor.execute(sql_insert_blob_query, insert_blob_tuple)
                emailDB_id = cursor.lastrowid
                print(f'"{item.subject}" :, Успешно занесено в базу')
                # Insert into Attachments DB
                item.is_read = True
                item.save()
                if CheckCreateDir(item.attachments):
                    directory = DirCreated(item.sender.name, date).Dir()

                    attach = GetAttachments(item.attachments, directory).FileAttachment()
                    for i in attach:
                        link = i
                        # Insert Attachments
                        add_Attachments = """INSERT INTO Attachments(link, id_email) VALUES (%s, %s)"""
                        data_Attachments = (link, emailDB_id)
                        cursor.execute(add_Attachments, data_Attachments)
                        print('Вложение определены и добавлены')


                else:
                    print("Вложения нет, директория не создана")

                conn_database.commit()
                item.delete()
                print('messege delete')

            except pymssql.Error as error:
                return error
            except Exception as s:
                pass

            cursor.close()
            conn_database.close()


while True:
    try:
        conn = ConnectToExchange(server, email, username, account).GetItemAccountInbox()
        time.sleep(30)
    except KeyboardInterrupt as s:
        print(s)
    except exchangelib.errors.ErrorFolderNotFound as error:
       print(error)

    except Exception as e:
        print(e)
