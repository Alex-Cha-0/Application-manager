"""Библиотека методов"""

import os
import time
import pymssql
from PyQt6 import QtWidgets, QtGui
from PyQt6.QtWidgets import QMainWindow, QTableWidgetItem, QApplication, QDialog, QFileDialog
from PyQt6.QtCore import QSettings, pyqtSignal
from app_manager import Ui_MainWindow
from ldap import GetNameFromLdap
# from requests_kerberos import HTTPKerberosAuth
# from exchangelib import DELEGATE, Account, Credentials, Configuration, FileAttachment, Message, ItemAttachment, Mailbox
# import pytz
# import exchangelib.autodiscover
# import urllib3
import pymssql as mc
from datetime import datetime, timedelta

from reply_email import Ui_MainWindow_reply

from cfg import SERVERAD, USERAD, PASSWORDAD, SERVERMSSQL, USERMSSQL, PASSWORDMSSQL, DATABASEMSSQL, CORP


class System(QMainWindow, Ui_MainWindow):

    def __init__(self):
        super().__init__()
        # Сохранение настроек
        self.settings = QSettings('app_manager', 'Ui_MainWindow')
        try:
            self.resize(self.settings.value('window size'))
            self.move(self.settings.value('window position'))
        except:
            pass
        self.setupUi(self)
        self.show()
        """REFRESH"""
        self.toolButton_refresh.clicked.connect(self.ImportFromDatabase)
        """ACCEPT BUTTON"""
        self.toolButton_accept.clicked.connect(self.AcceptButton)
        """Записать имя специалиста"""
        self.label_specialist.setText(self.GetNameSpecialist())
        """Получение данных из столбцов"""
        # Получение данных из колонки (ID письма)
        self.tableWidget_table.cellClicked.connect(self.CellWasClicked)
        # Выполнение select запроса SQL с ID письма
        self.tableWidget_table.cellClicked.connect(self.cellDoubleClicked)
        # Действия в таблице вложений
        self.tableWidget.cellClicked.connect(self.tableWidget_cellDoubleClicked)
        # Кнопка ответа на письмо
        self.toolButton_reply.clicked.connect(self.toolButton_replyclicked)
        # Кнопка закрытия заявки
        self.toolButton_closeorder.clicked.connect(self.toolButton_closeorderclicked)
        # Кнопка показать все заявки
        self.radioButton_all.clicked.connect(self.ImportFromDatabase)
        # Кнопка показать принятые заявки
        self.radioButton_accepted.clicked.connect(self.SelectFromEmailAcceptedOrder)
        # Кнопка показать закрытые заявки
        self.radioButton_closed.clicked.connect(self.SelectFromEmailClosedOrder)

        self.window = QtWidgets.QMainWindow()
        self.ui = Ui_MainWindow_reply()
        self.ui.setupUi((self.window))
        ##################
    """ФУНКЦИИ СЛОТЫ"""

    def toolButton_closeorderclicked(self):
        self.open_reply_window()
        self.SelectFromEmailForReplyClose()

        # self.UpdateEmailCloseOrder()
        # self.SelectFromEmailAcceptedOrder()

    def toolButton_replyclicked(self):
        self.open_reply_window()
        self.SelectFromEmailForAnswerDialog()

    def tableWidget_cellDoubleClicked(self):
        self.CellWasClickedAttachTable()
        self.StartOpenAttachment()

    def cellDoubleClicked(self):
        self.SecelctFromDataBase()
        self.SelectFromAttachDataToTable()

    def AcceptButton(self):
        self.UpdateEmailOpenOrder()
        self.ImportFromDatabase()

    """-------------------------------------"""

    # Открыть окно ответа на email
    def open_reply_window(self):
        self.window = QtWidgets.QMainWindow()
        self.ui = Ui_MainWindow_reply()
        self.ui.setupUi((self.window))
        self.window.show()

    def closeEvent(self, event):
        self.settings.setValue('window size', self.size())
        self.settings.setValue('window position', self.pos())

    def GetNameSpecialist(self):
        server = SERVERAD
        user = USERAD
        password = PASSWORDAD
        corp = CORP
        result = GetNameFromLdap(server, user, password, corp).Cnname()
        return result

    def GoToPage1(self):
        # Функция - открыть страницу с индексом - 1
        self.tabWidget.setCurrentIndex(1)

    def GoToPage2(self):
        # Функция - открыть страницу с индексом - 2
        self.tabWidget.setCurrentIndex(2)

    def CellWasClicked(self):
        # Функция - получить ID письма
        try:
            current_row = self.tableWidget_table.currentRow()
            # current_column = self.tableWidget_table.currentColumn()
            cell_value = self.tableWidget_table.item(current_row, 0).text()
            return cell_value
        except Exception as s:
            pass

    def CellWasClickedAttachTable(self):
        current_row = self.tableWidget.currentRow()
        current_column = self.tableWidget.currentColumn()
        cell_value = self.tableWidget.item(current_row, 1).text()
        return cell_value

    def StartOpenAttachment(self):
        try:
            cell = self.CellWasClickedAttachTable()
            link = cell
            os.startfile(link)
        except Exception as s:
            print(s)

    # def GetFromCellItem(self):
    #     row = self.tableWidget_table.currentRow()
    #     subject = self.tableWidget_table.item(row, 1).text()
    #     Autor = self.tableWidget_table.item(row, 2).text()
    #     Sender = self.tableWidget_table.item(row, 3).text()
    #     Copy = self.tableWidget_table.item(row, 4).text()
    #     Date = self.tableWidget_table.item(row, 5).text()

    #     self.label_sender.setText(f'{Autor}')
    #     self.label_time_send.setText(f'{Date}')
    #     self.label_subject.setText(subject)

    """--------------------------------"""

    def ImportFromDatabase(self):
        try:
            mydb = mc.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                              database=DATABASEMSSQL)

            mycursor = mydb.cursor()
            mycursor.execute("SELECT * FROM email WHERE open_order is null ORDER BY datetime_send DESC")
            result = mycursor.fetchall()
            self.tableWidget_table.setRowCount(0)
            # Смена имени колонки
            item = self.tableWidget_table.horizontalHeaderItem(0)
            item.setText("ID")
            item = self.tableWidget_table.horizontalHeaderItem(1)
            item.setText("Тема")
            item = self.tableWidget_table.horizontalHeaderItem(2)
            item.setText("Автор")
            item = self.tableWidget_table.horizontalHeaderItem(3)
            item.setText('Отправитель')
            item = self.tableWidget_table.horizontalHeaderItem(4)
            item.setText("Копия")
            item = self.tableWidget_table.horizontalHeaderItem(5)
            item.setText("Дата создания")
            item = self.tableWidget_table.horizontalHeaderItem(6)
            item.setText("Вложение")

            for row_number, row_data in enumerate(result):
                self.tableWidget_table.insertRow(row_number)

                for column_number, data in enumerate(row_data):
                    self.tableWidget_table.setItem(row_number, column_number, QTableWidgetItem(str(data)))
            self.SetBackgroundIDColor()
            self.SetAttachIcon()
            self.SetReplyIcon()
        except mc.Error as e:
            print(e)

    def SecelctFromDataBase(self):
        try:
            mydb = mc.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                              database=DATABASEMSSQL)
            id = self.CellWasClicked()
            if id:
                mycursor = mydb.cursor()
                sql_select_query = mycursor.execute(
                    f"""SELECT text_body, yes_no_attach, copy, sender_name, datetime_send, subject  FROM email WHERE id = {id}""")
                result = mycursor.fetchall()
                self.textBrowser_email_1.setText(f'{result[0][0]}')
                self.label_attachments.setVisible(result[0][1])

                self.label_sender.setText(f'{result[0][3]}')
                self.label_time_send.setText(f'{result[0][4]}')
                self.label_subject.setText(f'{result[0][5]}')

                if result[0][2] is not None:
                    self.label.setVisible(True)
                    self.label_copy.setVisible(True)
                    self.label_copy.setText(result[0][2])
                else:
                    self.label.setVisible(False)
                    self.label_copy.setVisible(False)

                return result[0][0]


        except mc.Error as e:
            pass

    def JoinFromDataBase(self):
        id_cell = self.CellWasClicked()
        if id_cell:
            try:
                mydb = mc.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                                  database=DATABASEMSSQL)

                mycursor = mydb.cursor()
                sql_select_query = mycursor.execute(
                    f"""SELECT id, subject, sender_name FROM email where id = {id_cell}""")
                result = mycursor.fetchall()
                return result[0]
            except pymssql.Error as error:
                # self.label_erorr3.setText("Failed select data from MySQL table {}".format(error))
                pass
        else:
            # self.label_erorr3.setText("Ничего не выбрано")
            pass

    # def InsertBlobInOrderInWork(self):
    #
    #     if not self.CompareSqlTablesEmailAndTblwork():
    #         try:
    #             id, subject, sender = self.JoinFromDataBase()
    #             date = datetime.today().strftime('%Y-%m-%d %H:%M:%S')
    #
    #             specialist = self.GetNameSpecialist()
    #             conn_database = pymssql.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
    #                                             database=DATABASEMSSQL)
    #             cursor = conn_database.cursor()
    #             # Sql query
    #             sql_insert_blob_query = """INSERT INTO work (id_email_in_work, subject, autor, control_period, specialist) VALUES (%s,%s,%s,%s,%s) """
    #             # Convert data into tuple format
    #             insert_blob_tuple = (id, subject, sender, date, specialist)
    #             result = cursor.execute(sql_insert_blob_query, insert_blob_tuple)
    #             self.label_statusneworder.setText(f"'Заявка '{subject}' принята'")
    #             self.label_statusneworder.setStyleSheet('color:green')
    #
    #             conn_database.commit()
    #             cursor.close()
    #             conn_database.close()
    #         except pymssql.Error as error:
    #             # self.label_erorr3.setText("Failed inserting BLOB data into MySQL table {}".format(error))
    #             pass
    #
    #     else:
    #         self.label_statusneworder.setText(f"'Заявка уже в работе'")
    #         self.label_statusneworder.setStyleSheet('color : red')

    # def ImportFromTblWork(self):
    #     try:
    #         mydb = mc.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
    #                           database=DATABASEMSSQL)
    #
    #         mycursor = mydb.cursor()
    #         mycursor.execute("SELECT * FROM work ORDER BY control_period DESC")
    #         result = mycursor.fetchall()
    #
    #         self.tableWidget_order_tblwork.setRowCount(0)
    #
    #         for row_number, row_data in enumerate(result):
    #             self.tableWidget_order_tblwork.insertRow(row_number)
    #
    #             for column_number, data in enumerate(row_data):
    #                 self.tableWidget_order_tblwork.setItem(row_number, column_number, QTableWidgetItem(str(data)))
    #     except mc.Error as e:
    #         # self.label_error.setText("Error ", e)
    #         pass

    # def CompareSqlTablesEmailAndTblwork(self):
    #     try:
    #         id_cell = self.CellWasClicked()
    #         mydb = mc.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
    #                           database=DATABASEMSSQL)
    #
    #         mycursor = mydb.cursor()
    #         sql_select_query = mycursor.execute(f"""SELECT id FROM email WHERE id = ANY (SELECT
    #                                         id_email_in_work FROM work WHERE id_email_in_work = '{id_cell}')""")
    #         result = mycursor.fetchall()
    #         if result:
    #             return True
    #         else:
    #             return False
    #
    #     except pymssql.Error as error:
    #         # print('Такое значение отсутствует', error)
    #         pass
    #     except Exception as s:
    #         # print('Ошибка: ', s)
    #         pass

    def ShowAttach(self):
        id_email = self.CellWasClicked()

        try:
            conn_database = pymssql.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                                            database=DATABASEMSSQL)
            cursor = conn_database.cursor()
            sql_select_query = cursor.execute(f"""SELECT Link FROM Attachments WHERE id_email = {id_email}""")

            result = cursor.fetchall()
            fname = QFileDialog.getOpenFileName(self, "Open file", result[0][0])

            os.startfile(fname[0])

            # if fname[0]:
            #     f = open(fname[0], 'r')
            #
            #     with f:
            #         data = f.read()
        except IndexError:
            print('list index out of range')
        except Exception as s:
            # self.label_error2.setText(f'{s}')
            pass

    def SelectFromAttachDataToTable(self):
        id_email = self.CellWasClicked()
        try:
            if id_email:
                mydb = mc.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                                  database=DATABASEMSSQL)

                mycursor = mydb.cursor()
                mycursor.execute(f"SELECT link FROM Attachments where id_email = {id_email}")
                result = mycursor.fetchall()

                self.tableWidget.setRowCount(0)

                for row_number, row_data in enumerate(result):
                    self.tableWidget.insertRow(row_number)

                    for column_number, data in enumerate(row_data):
                        data_ = '\\'.join(data.split('\\')[-1:])
                        self.tableWidget.setItem(row_number, column_number, QTableWidgetItem(str(data_)))

                        self.tableWidget.setItem(row_number, column_number + 1, QTableWidgetItem(str(data)))

        except mc.Error as e:
            # self.label_error.setText("Error ", e)
            pass

    def SelectFromEmailForAnswerDialog(self):
        id = self.CellWasClicked()
        if id:
            try:
                mydb = mc.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                                  database=DATABASEMSSQL)
                mycursor = mydb.cursor()
                sql_select_query = mycursor.execute(
                    f"""SELECT sender_email, copy, subject, text_body, sender_name, datetime_send, recipients  FROM email WHERE id = {id}""")
                result = mycursor.fetchall()
                subject_email = 'RE: ' + str(result[0][2])

                self.ui.textBrowser_reply.setText(result[0][3])
                self.ui.lineEdit_send_email.setText(result[0][0])
                self.ui.lineEdit_copy.setText(result[0][1])
                self.ui.lineEdit_subject.setText(subject_email)
                self.ui.label_idcell.setText(id)
                self.ui.textEdit_from.setText(
                    f'От кого: {result[0][4]}, {result[0][0]}\nДата: {result[0][5]}\nКому: {result[0][6]}\nТема: {result[0][2]}')


            except mc.Error as error:
                pass
            except Exception as e:
                print(e)
        else:
            pass

    def UpdateEmailOpenOrder(self):
        id = self.CellWasClicked()
        if not self.CheckOrderIsOpen():
            try:


                open_order = True
                specialist = self.GetNameSpecialist()
                conn_database = pymssql.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                                                database=DATABASEMSSQL)
                cursor = conn_database.cursor()
                # Sql query
                sql_select = cursor.execute(f"""SELECT datetime_send FROM email WHERE id = '{id}'""")
                datetime_send = cursor.fetchall()
                """Контрольный срок"""
                ks = datetime_send[0][0] + timedelta(days=10)
                date = ks.strftime('%Y-%m-%d %H:%M:%S')
                """-----------------"""

                sql_insert_blob_query = f"""UPDATE email SET specialist = '{specialist}',control_period = '{date}', open_order = '{open_order}'   WHERE id = '{id}' """
                # Convert data into tuple format

                result = cursor.execute(sql_insert_blob_query)
                self.label_statusneworder.setText(f"'Заявка '{id}' принята в работу'")
                self.label_statusneworder.setStyleSheet('color:green')

                conn_database.commit()
                cursor.close()
                conn_database.close()
            except pymssql.Error as error:
                # self.label_erorr3.setText("Failed inserting BLOB data into MySQL table {}".format(error))
                print(error)

        else:
            self.label_statusneworder.setText(f"'Заявка уже в работе'")
            self.label_statusneworder.setStyleSheet('color : red')

    def CheckOrderIsOpen(self):
        try:
            id_cell = self.CellWasClicked()
            mydb = mc.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                              database=DATABASEMSSQL)

            mycursor = mydb.cursor()
            sql_select_query = mycursor.execute(f"""SELECT open_order, close_order FROM email WHERE id = {id_cell}""")
            result = mycursor.fetchall()
            if result[0][0] or result[0][1]:
                return True
            else:
                return False

        except pymssql.Error as error:
            print(error)
        except Exception as s:
            print(s)

    def SelectFromEmailAcceptedOrder(self):
        try:
            mydb = mc.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                              database=DATABASEMSSQL)
            mycursor = mydb.cursor()
            sql_select_query = mycursor.execute(
                f"""SELECT id, subject, sender_name, specialist, control_period, datetime_send, yes_no_attach FROM email where open_order = 'True' ORDER BY control_period DESC""")
            result = mycursor.fetchall()
            self.tableWidget_table.setRowCount(0)
            # Смена имени колонки
            item = self.tableWidget_table.horizontalHeaderItem(0)
            item.setText("ID")
            item = self.tableWidget_table.horizontalHeaderItem(1)
            item.setText("Тема")
            item = self.tableWidget_table.horizontalHeaderItem(2)
            item.setText("Имя автора")
            item = self.tableWidget_table.horizontalHeaderItem(3)
            item.setText('Специалист')
            item = self.tableWidget_table.horizontalHeaderItem(4)
            item.setText("Контрольный срок")
            item = self.tableWidget_table.horizontalHeaderItem(5)
            item.setText("Дата создания")
            item = self.tableWidget_table.horizontalHeaderItem(6)
            item.setText("Вложение")

            for row_number, row_data in enumerate(result):
                self.tableWidget_table.insertRow(row_number)

                for column_number, data in enumerate(row_data):
                    self.tableWidget_table.setItem(row_number, column_number, QTableWidgetItem(str(data)))
            self.SetBackgroundKSColor()
            self.SetAttachIcon()
            self.SetReplyIcon()
        except mc.Error as e:
            pass
        except Exception as erorr:
            print(erorr)

    def UpdateEmailCloseOrder(self):
        id = self.CellWasClicked()
        if self.CheckOrderIsOpen() == self.CheckOrderIsClose():
            self.label_statusneworder.setText(f"'Сначала нужно принять заявку'")
            self.label_statusneworder.setStyleSheet('color : red')
        else:
            if not self.CheckOrderIsClose():
                try:
                    date = datetime.today().strftime('%Y-%m-%d %H:%M:%S')
                    close_order = True
                    open_order = False
                    # specialist = self.GetNameSpecialist()
                    conn_database = pymssql.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                                                    database=DATABASEMSSQL)
                    cursor = conn_database.cursor()
                    # Sql query
                    sql_insert_blob_query = f"""UPDATE email SET date_complited = '{date}',open_order = '{open_order}', close_order = '{close_order}'   WHERE id = '{id}' """
                    # Convert data into tuple format

                    result = cursor.execute(sql_insert_blob_query)
                    self.label_statusneworder.setText(f"'Заявка {id} закрыта!'")
                    self.label_statusneworder.setStyleSheet('color:green')

                    conn_database.commit()
                    cursor.close()
                    conn_database.close()
                except pymssql.Error as error:
                    # self.label_erorr3.setText("Failed inserting BLOB data into MySQL table {}".format(error))
                    print(error)

            else:
                self.label_statusneworder.setText(f"'Заявка уже в закрыта'")
                self.label_statusneworder.setStyleSheet('color : red')

    def CheckOrderIsClose(self):
        try:
            id_cell = self.CellWasClicked()
            mydb = mc.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                              database=DATABASEMSSQL)

            mycursor = mydb.cursor()
            sql_select_query = mycursor.execute(f"""SELECT open_order, close_order FROM email WHERE id = {id_cell}""")
            result = mycursor.fetchall()
            if result[0][1]:
                return True
            else:
                return False

        except pymssql.Error as error:
            print(error)
        except Exception as s:
            print(s)

    def SelectFromEmailClosedOrder(self):
        try:
            mydb = mc.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                              database=DATABASEMSSQL)
            mycursor = mydb.cursor()
            sql_select_query = mycursor.execute(
                f"""SELECT id, subject, sender_name, specialist, control_period, date_complited, yes_no_attach FROM email where close_order = 'True' ORDER BY control_period DESC""")
            result = mycursor.fetchall()
            self.tableWidget_table.setRowCount(0)
            # Смена имени колонки
            item = self.tableWidget_table.horizontalHeaderItem(3)
            item.setText("Специалист")
            item = self.tableWidget_table.horizontalHeaderItem(4)
            item.setText('Контрольный срок')
            item = self.tableWidget_table.horizontalHeaderItem(2)
            item.setText("Имя автора")
            item = self.tableWidget_table.horizontalHeaderItem(5)
            item.setText("Дата закрытия")


            for row_number, row_data in enumerate(result):
                self.tableWidget_table.insertRow(row_number)

                for column_number, data in enumerate(row_data):
                    self.tableWidget_table.setItem(row_number, column_number, QTableWidgetItem(str(data)))

            self.SetAttachIcon()
            self.SetReplyIcon()
        except mc.Error as e:
            pass
        except Exception as erorr:
            print(erorr)

    def SetBackgroundIDColor(self):
        # columns = self.tableWidget_table.columnCount()
        rows = self.tableWidget_table.rowCount()
        for row in range(rows):
            self.tableWidget_table.item(row, 0).setBackground(QtGui.QColor(176, 224, 230))
        # try:
        #     mydb = mc.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
        #                       database=DATABASEMSSQL)
        #
        #     mycursor = mydb.cursor()
        #
        #     rows = self.tableWidget_table.rowCount()
        #     for row in range(rows):
        #         item_id = self.tableWidget_table.item(row, 0).text()
        #         sql_select_query = mycursor.execute(
        #             f"""SELECT open_order, close_order FROM email WHERE id = {item_id}""")
        #         result = mycursor.fetchall()
        #
        #         if result[0][1] == result[0][0]:
        #             self.tableWidget_table.item(row, 0).setBackground(QtGui.QColor(176, 224, 230))
        # except Exception as erorr:
        #     print(erorr)

    def SetBackgroundKSColor(self):
        try:
            mydb = mc.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                              database=DATABASEMSSQL)

            mycursor = mydb.cursor()

            rows = self.tableWidget_table.rowCount()
            for row in range(rows):
                item_id = self.tableWidget_table.item(row, 0).text()
                sql_select_query = mycursor.execute(
                    f"""SELECT control_period FROM email WHERE id = {item_id}""")
                date_ks = mycursor.fetchall()

                date = datetime.today()

                if date_ks[0][0] >= date:
                    self.tableWidget_table.item(row, 4).setBackground(QtGui.QColor(211, 252, 204))
                else:
                    self.tableWidget_table.item(row, 4).setBackground(QtGui.QColor(252, 204, 208))
        except Exception as erorr:
            print(erorr)

    def SetAttachIcon(self):
        """Иконка вложений в таблице"""
        try:
            mydb = mc.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                              database=DATABASEMSSQL)

            mycursor = mydb.cursor()

            rows = self.tableWidget_table.rowCount()
            for row in range(rows):
                item_id = self.tableWidget_table.item(row, 0).text()
                sql_select_query = mycursor.execute(
                    f"""SELECT yes_no_attach FROM email WHERE id = {item_id}""")
                result = mycursor.fetchall()
                if result[0][0]:
                    icon = QtGui.QIcon()
                    icon.addPixmap(QtGui.QPixmap(":/images/image/folder.png"), QtGui.QIcon.Mode.Normal,
                                   QtGui.QIcon.State.Off)
                    self.tableWidget_table.item(row, 6).setIcon(icon)
        except:
            pass

    def SetReplyIcon(self):
        """Иконка ответа в таблице"""
        try:
            mydb = mc.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                              database=DATABASEMSSQL)

            mycursor = mydb.cursor()

            rows = self.tableWidget_table.rowCount()
            for row in range(rows):
                item_id = self.tableWidget_table.item(row, 0).text()
                sql_select_query = mycursor.execute(
                    f"""SELECT reply_email FROM email WHERE id = {item_id}""")
                result = mycursor.fetchall()
                if result[0][0]:
                    icon = QtGui.QIcon()
                    icon.addPixmap(QtGui.QPixmap(":/images/image/open-email.png"), QtGui.QIcon.Mode.Normal,
                                   QtGui.QIcon.State.Off)
                    self.tableWidget_table.item(row, 0).setIcon(icon)
        except:
            pass

    def SelectFromEmailForReplyClose(self):
        """Открытие ответа при нажатии кнопки закрыть"""
        id = self.CellWasClicked()
        if id:
            try:
                mydb = mc.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                                  database=DATABASEMSSQL)
                mycursor = mydb.cursor()
                sql_select_query = mycursor.execute(
                    f"""SELECT sender_email, copy, subject, text_body, sender_name, datetime_send, recipients  FROM email WHERE id = {id}""")
                result = mycursor.fetchall()
                subject_email = str(result[0][2])

                self.ui.textBrowser_reply.setText(result[0][3])
                self.ui.lineEdit_send_email.setText(result[0][0])
                self.ui.lineEdit_copy.setText(result[0][1])
                self.ui.lineEdit_subject.setText(subject_email)
                self.ui.label_idcell.setText(id)
                self.ui.textEdit_from.setText(
                    f'От кого: {result[0][4]}, {result[0][0]}\nДата: {result[0][5]}\nКому: {result[0][6]}\nТема: {result[0][2]}')

                self.ui.textEdit_perlyemail.setText(f'Ваша заявка - "{subject_email}" закрыта.')
            except mc.Error as error:
                pass
            except Exception as e:
                print(e)
        else:
            pass


