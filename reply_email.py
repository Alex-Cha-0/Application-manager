"""Скрипт ответа на email"""

from requests_kerberos import HTTPKerberosAuth
from exchangelib import DELEGATE, Account, Credentials, Configuration, Message, Mailbox

import urllib3
import pytz
import pymssql

from cfg import SERVEREXCHANGE, SERVERMSSQL, USERMSSQL, PASSWORDMSSQL, \
    DATABASEMSSQL, SERVERAD, USERAD, PASSWORDAD, CORP

from PyQt6 import QtCore, QtGui, QtWidgets

from ldap import GetNameFromLdap


class Ui_MainWindow_reply(object):
    def setupUi(self, MainWindow_reply):
        MainWindow_reply.setObjectName("MainWindow_reply")
        MainWindow_reply.resize(1024, 768)
        MainWindow_reply.setStyleSheet("font: 10pt \"Calibri\";")
        self.centralwidget = QtWidgets.QWidget(MainWindow_reply)
        self.centralwidget.setObjectName("centralwidget")
        self.toolButton_send = QtWidgets.QToolButton(self.centralwidget)
        self.toolButton_send.setGeometry(QtCore.QRect(10, 10, 91, 81))
        self.toolButton_send.setStyleSheet("font: 75 12pt \"Calibri\";\n"
                                           "border-radius: 20px;\n"
                                           "background-color: #b0e0e6;\n"
                                           "color:white")
        self.toolButton_send.setObjectName("toolButton_send")
        self.pushButton_send_email = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_send_email.setGeometry(QtCore.QRect(110, 10, 75, 23))
        self.pushButton_send_email.setStyleSheet("font: 75 12pt \"Calibri\";\n"
                                                 "border-radius: 4px;\n"
                                                 "background-color: #b0e0e6;\n"
                                                 "color:white")
        self.pushButton_send_email.setObjectName("pushButton_send_email")
        self.pushButton_copy = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_copy.setGeometry(QtCore.QRect(110, 40, 75, 23))
        self.pushButton_copy.setStyleSheet("font: 75 12pt \"Calibri\";\n"
                                           "border-radius: 4px;\n"
                                           "background-color: #b0e0e6;\n"
                                           "color:white")
        self.pushButton_copy.setObjectName("pushButton_copy")
        self.label_subject = QtWidgets.QLabel(self.centralwidget)
        self.label_subject.setGeometry(QtCore.QRect(110, 70, 47, 13))
        self.label_subject.setObjectName("label_subject")
        self.lineEdit_send_email = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_send_email.setGeometry(QtCore.QRect(190, 10, 821, 20))
        self.lineEdit_send_email.setObjectName("lineEdit_send_email")
        self.lineEdit_copy = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_copy.setGeometry(QtCore.QRect(190, 40, 821, 20))
        self.lineEdit_copy.setObjectName("lineEdit_copy")
        self.lineEdit_subject = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_subject.setGeometry(QtCore.QRect(190, 70, 731, 20))
        self.lineEdit_subject.setObjectName("lineEdit_subject")
        self.textBrowser_reply = QtWidgets.QTextBrowser(self.centralwidget)
        self.textBrowser_reply.setGeometry(QtCore.QRect(10, 330, 1001, 431))
        font = QtGui.QFont()
        font.setFamily("Calibri")
        font.setPointSize(10)
        font.setBold(False)
        font.setItalic(True)
        font.setWeight(3)
        self.textBrowser_reply.setFont(font)
        self.textBrowser_reply.setStyleSheet("font: 25 italic 10pt \"Calibri\";\n"
                                             "background-color: rgb(246, 255, 255);")
        self.textBrowser_reply.setFrameShape(QtWidgets.QFrame.Shape.NoFrame)
        self.textBrowser_reply.setFrameShadow(QtWidgets.QFrame.Shadow.Sunken)
        self.textBrowser_reply.setTextInteractionFlags(QtCore.Qt.TextInteractionFlag.TextBrowserInteraction)
        self.textBrowser_reply.setOpenExternalLinks(True)
        self.textBrowser_reply.setObjectName("textBrowser_reply")
        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(930, 70, 81, 23))
        self.pushButton_3.setStyleSheet("font: 75 12pt \"Calibri\";\n"
                                        "border-radius: 4px;\n"
                                        "background-color: #b0e0e6;\n"
                                        "color:white")
        self.pushButton_3.setObjectName("pushButton_3")
        self.textEdit_perlyemail = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit_perlyemail.setGeometry(QtCore.QRect(10, 140, 1001, 121))
        self.textEdit_perlyemail.setStyleSheet("font: 12pt \"Calibri\";")
        self.textEdit_perlyemail.setFrameShape(QtWidgets.QFrame.Shape.NoFrame)
        self.textEdit_perlyemail.setObjectName("textEdit_perlyemail")
        self.textEdit_from = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit_from.setGeometry(QtCore.QRect(10, 260, 1001, 71))
        self.textEdit_from.setStyleSheet("font: italic 10pt \"Calibri\";")
        self.textEdit_from.setFrameShape(QtWidgets.QFrame.Shape.Panel)
        self.textEdit_from.setFrameShadow(QtWidgets.QFrame.Shadow.Sunken)
        self.textEdit_from.setObjectName("textEdit_from")
        self.label_status = QtWidgets.QLabel(self.centralwidget)
        self.label_status.setGeometry(QtCore.QRect(10, 100, 931, 31))
        self.label_status.setStyleSheet("font: 25 10pt \"Calibri\";")
        self.label_status.setText("")
        self.label_status.setObjectName("label_status")
        self.label_idcell = QtWidgets.QLabel(self.centralwidget)
        self.label_idcell.setGeometry(QtCore.QRect(950, 100, 47, 13))
        self.label_idcell.setStyleSheet("font: 8pt \"MS Shell Dlg 2\";")
        self.label_idcell.setText("")
        self.label_idcell.setObjectName("label_idcell")
        MainWindow_reply.setCentralWidget(self.centralwidget)

        """Кнопка отправить"""
        self.toolButton_send.clicked.connect(self.GetTExtFromWindow)
        self.toolButton_send.clicked.connect(self.ReplyEmail)
        self.toolButton_send.clicked.connect(self.UpdateReplyEmail)
        # self.toolButton_send.clicked.connect(self.UpdateEmailCloseOrder)
        self.label_idcell.setVisible(False)

        self.retranslateUi(MainWindow_reply)
        QtCore.QMetaObject.connectSlotsByName(MainWindow_reply)

    def retranslateUi(self, MainWindow_reply):
        _translate = QtCore.QCoreApplication.translate
        MainWindow_reply.setWindowTitle(_translate("MainWindow_reply", "reply"))
        self.toolButton_send.setText(_translate("MainWindow_reply", "Отправить"))
        self.pushButton_send_email.setText(_translate("MainWindow_reply", "Кому"))
        self.pushButton_copy.setText(_translate("MainWindow_reply", "Копия"))
        self.label_subject.setText(_translate("MainWindow_reply", "Тема :"))
        self.pushButton_3.setText(_translate("MainWindow_reply", "Вложение"))


    def Data_Division(self):
        server = SERVERAD
        user = USERAD
        password = PASSWORDAD
        corp = CORP
        name = GetNameFromLdap(server, user, password, corp).Cnname()

        mydb = pymssql.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                               database=DATABASEMSSQL)
        mycursor = mydb.cursor()
        mycursor.execute(f"SELECT uid,employee,email,user_exchange,password_exchange FROM Staff INNER JOIN Division "
                         f"ON uid_Division = uid WHERE employee = '{name}'")
        result = mycursor.fetchall()
        return result

    def IdCellInAppManger(self):
        id_c = self.label_idcell.text()
        return int(id_c)

    def GetTExtFromWindow(self):
        textBrowser_reply = self.textBrowser_reply.toPlainText()
        textEdit_perlyemail = self.textEdit_perlyemail.toPlainText()
        textFrom = self.textEdit_from.toPlainText()

        body = textEdit_perlyemail + '\n' + '\n' + '-------------' + '\n' + textFrom + '\n' + '\n' + textBrowser_reply
        return body

    def ReplyEmail(self):
        def auth_model(**kwargs):
            # get kerberos ticket
            return HTTPKerberosAuth()

        urllib3.disable_warnings()
        tz = pytz.timezone('Europe/Moscow')

        def connect(server, email, username, password):
            from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
            # Care! Ignor Exchange self-signed SSL cert
            BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter

            # fill Credential object with empty fields
            creds = Credentials(
                username=username,
                password=password
            )

            # add kerberos as GSSAPI auth_type
            # exchangelib.transport.AUTH_TYPE_MAP["GSSAPI"] = auth_model

            # Create Config
            config = Configuration(server=server, credentials=creds)
            # return result
            return Account(primary_smtp_address=email, autodiscover=False, config=config, access_type=DELEGATE)

        server = SERVEREXCHANGE
        # email = EMAILADDRESS
        # username = USEREXECHANGE
        # password = USEREXECHANGEPASS
        email = self.Data_Division()[0][2]
        username = self.Data_Division()[0][3]
        password = self.Data_Division()[0][4]
        account = connect(server, email, username, password)
        recipients = str(self.lineEdit_send_email.text()).replace(';', '').split()
        copy = self.lineEdit_copy.text().replace(';', '').split()
        subject = self.lineEdit_subject.text()
        body = self.GetTExtFromWindow()

        status = self.label_status

        class ConnectToExchange(object):
            """docstring"""
            status.setText('ConnectToExchange')
            status.setStyleSheet('color:green')

            def __init__(self, server, email, username, account):
                """Constructor"""

                self.server = server
                self.email = email
                self.username = username
                self.account = account

            def Send(self):
                try:
                    # m = Message(account=account, subject=subject, body=body,
                    #             to_recipients=[Mailbox(email_address=' '.join(str(e) for e in recipients))],
                    #             cc_recipients=copy)
                    # m.send()
                    to_recipients = []
                    for recipient in recipients:
                        to_recipients.append(Mailbox(email_address=recipient))
                    # Create message
                    m = Message(account=account,
                                subject=subject,
                                body=body,
                                to_recipients=to_recipients,
                                cc_recipients=copy)

                    m.send_and_save()

                    status.setText('Сообщение отправлено!')
                    status.setStyleSheet('color:green')

                except Exception as s:
                    print(s)

        conn = ConnectToExchange(server, email, username, account)
        conn.Send()

    def UpdateReplyEmail(self):
        id = self.IdCellInAppManger()
        body = self.GetTExtFromWindow()
        reply = True
        try:

            conn_database = pymssql.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                                            database=DATABASEMSSQL)
            cursor = conn_database.cursor()
            # Sql query
            sql_insert_blob_query = f"""UPDATE email SET text_body = '{body}', reply_email = '{reply}'  WHERE id = '{id}' """
            # Convert data into tuple format

            result = cursor.execute(sql_insert_blob_query)
            conn_database.commit()
            cursor.close()
            conn_database.close()
        except pymssql.Error as error:
            # self.label_erorr3.setText("Failed inserting BLOB data into MySQL table {}".format(error))
            print(error)
