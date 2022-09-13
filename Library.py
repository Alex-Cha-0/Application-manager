"""Библиотека методов"""

import os

import PyQt5
import pymssql
from PyQt6 import QtWidgets, QtGui
from PyQt6.QtGui import QFont, QColor, QBrush
from PyQt6.QtWidgets import QMainWindow, QTableWidgetItem, QFileDialog, QButtonGroup
from PyQt6.QtCore import QSettings
from app_manager import Ui_MainWindow
from ldap import GetNameFromLdap
from requests_kerberos import HTTPKerberosAuth
from exchangelib import DELEGATE, Account, Credentials, Configuration, Message, Mailbox
import pytz

import urllib3
import pymssql as mc
from datetime import datetime, timedelta

from reply_email import Ui_MainWindow_reply

from cfg import SERVERAD, USERAD, PASSWORDAD, SERVERMSSQL, USERMSSQL, PASSWORDMSSQL, DATABASEMSSQL, CORP, \
    SERVEREXCHANGE


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
        self.toolButton_refresh.clicked.connect(self.ImportFromDatabaseAll)
        """ACCEPT BUTTON"""
        self.toolButton_accept.clicked.connect(self.AcceptButton)
        """Записать имя специалиста"""
        self.label_specialist.setText(self.GetNameSpecialist())
        """Получение данных из столбцов"""
        # Получение данных из колонки (ID письма)
        self.tableWidget_table.cellClicked.connect(self.CellWasClicked)
        # Выполнение select запроса SQL с ID письма
        self.tableWidget_table.cellClicked.connect(self.cellDoubleClicked)

        self.tableWidget_table.doubleClicked.connect(self.DoubleClickedAtAllOrder)

        # Действия в таблице вложений
        self.tableWidget.cellClicked.connect(self.tableWidget_cellDoubleClicked)
        # Кнопка ответа на письмо
        self.toolButton_reply.clicked.connect(self.toolButton_replyclicked)
        # Кнопка закрытия заявки
        self.toolButton_closeorder.clicked.connect(self.toolButton_closeorderclicked)
        # self.toolButton_closeorder.clicked.connect(self.ReplyEmail)
        # Кнопка показать все заявки
        self.radioButton_all.clicked.connect(self.ImportFromDatabaseAll)
        # Кнопка показать принятые заявки
        self.radioButton_accepted.clicked.connect(self.SelectFromEmailAcceptedOrder)
        # Кнопка показать закрытые заявки
        self.radioButton_closed.clicked.connect(self.SelectFromEmailClosedOrder)
        # Массив радиокнопок
        self.radio_buttons = [self.radioButton_all, self.radioButton_accepted, self.radioButton_closed]
        self.radio_buttons_functions = [self.ImportFromDatabaseAll, self.SelectFromEmailAcceptedOrder,
                                        self.SelectFromEmailClosedOrder]
        # Группа радиокнопок
        self.radio_group = QButtonGroup()
        self.radio_group.addButton(self.radioButton_all,1)
        self.radio_group.addButton(self.radioButton_accepted,2)
        self.radio_group.addButton(self.radioButton_closed,3)
        self.radio_group.buttonClicked.connect(self.GetIdRadioButtons)

        # Чекбокс
        self.checkBox_allnotclose.clicked.connect(self.ChekboxEvent)

        self.window = QtWidgets.QMainWindow()
        self.ui = Ui_MainWindow_reply()
        self.ui.setupUi(self.window)
        ##################
        # ComboBox специалист
        self.comboBox.activated.connect(self.CurrentTextComboboxSpec)
        self.comboBox.activated.connect(self.ChekBoxEvent)
        self.comboBox.activated.connect(self.CountVrabote)
        # ComboBox показать
        self.comboBox_2.activated.connect(self.CurrentTextComboboxShow)
        # Кол-во колонок в таблице
        self.count_column = self.tableWidget_table.horizontalHeader().count()

    """Сохранение настроек"""

    def closeEvent(self, event):
        self.settings.setValue('window size', self.size())
        self.settings.setValue('window position', self.pos())
        self.saveSetting()
        event.accept()

    def saveSetting(self):
        self.settings.setValue('checkbox', int(self.ChekboxEvent()))
        self.settings.setValue('combobox_spec', self.comboBox.currentIndex())
        self.settings.setValue('combobox_show', self.comboBox_2.currentIndex())
        for b, button in enumerate(self.radio_buttons):
            if button.isChecked():
                self.settings.setValue('radio_button', b)

        for i in range(self.count_column):
            self.settings.setValue(f'column {i}', self.tableWidget_table.columnWidth(i))

    def loadSetting(self):
        self.checkBox_allnotclose.setChecked(self.settings.value('checkbox', 0))
        self.comboBox.setCurrentIndex(self.settings.value('combobox_spec', 0))
        self.comboBox_2.setCurrentIndex(self.settings.value('combobox_show', 0))

        radio_button = self.settings.value('radio_button', 0)
        self.radio_buttons[radio_button].toggled.connect(self.radio_buttons_functions[radio_button])
        self.radio_buttons[radio_button].setChecked(True)

        for i in range(self.count_column):
            self.tableWidget_table.setColumnWidth(i, self.settings.value(f'column {i}', 100))

    ##########################################################
    """ФУНКЦИИ СЛОТЫ"""

    def ChekboxEvent(self):
        return self.checkBox_allnotclose.isChecked()

    def GetIdRadioButtons(self):
        return self.radio_group.checkedId()

    def ChekBoxEvent(self):
        radio_id = self.GetIdRadioButtons()
        if radio_id == 1:
            self.ImportFromDatabaseAll()
        if radio_id == 2:
            self.SelectFromEmailAcceptedOrder()
        if radio_id == 3:
            self.SelectFromEmailClosedOrder()

    def DoubleClickedAtAllOrder(self):
        id = self.CellWasClicked()
        # self.SelectFromEmailAcceptedOrder()
        mydb = mc.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                          database=DATABASEMSSQL)
        mycursor = mydb.cursor()
        mycursor.execute(f'''SELECT open_order, close_order from email where id = {id}''')
        result = mycursor.fetchall()
        open_order = result[0][0]
        close_order = result[0][1]
        if open_order:
            self.SelectFromEmailAcceptedOrder()
            self.radioButton_accepted.click()
        elif close_order:
            self.SelectFromEmailClosedOrder()
            self.radioButton_closed.click()
        for row in range(self.tableWidget_table.rowCount()):
            cell_value = self.tableWidget_table.item(row, 0).text()
            if cell_value == id:
                self.tableWidget_table.selectRow(row)

    def CurrentTextComboboxShow(self):
        current_text = self.comboBox_2.currentText()
        return current_text

    def ComboBoxSpec(self):
        self.comboBox.addItem('Все')

        mydb = mc.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                          database=DATABASEMSSQL)
        mycursor = mydb.cursor()
        mycursor.execute(f'''SELECT employee from Staff where uid_Division = {self.Data_Division()[0][0]}''')
        result = mycursor.fetchall()
        for spec in result:
            self.comboBox.addItem(f'{spec[0]}')

    def CurrentTextComboboxSpec(self):
        current_text = self.comboBox.currentText()
        return current_text

    def CountVrabote(self):
        mydb = mc.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                          database=DATABASEMSSQL)
        mycursor = mydb.cursor()
        spec = self.CurrentTextComboboxSpec()
        if spec == 'Все':
            param = ''
        else:
            param = 'AND specialist =' + "'" + spec + "'"
        mycursor.execute(f'''SELECT COUNT(open_order) FROM email WHERE open_order = 1 {param}''')
        result = mycursor.fetchall()
        self.label_count_v_rabote.setText(str(result[0][0]))

    def Data_Division(self):
        name = self.label_specialist.text()
        mydb = mc.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                          database=DATABASEMSSQL)
        mycursor = mydb.cursor()
        mycursor.execute(f"SELECT uid,employee,email,user_exchange,password_exchange FROM Staff INNER JOIN Division "
                         f"ON uid_Division = uid WHERE employee = '{name}'")
        result = mycursor.fetchall()
        return result

    def toolButton_closeorderclicked(self):
        self.UpdateEmailCloseOrder()
        self.CountVrabote()

    def toolButton_replyclicked(self):
        self.open_reply_window()
        self.SelectFromEmailForAnswerDialog()

    def tableWidget_cellDoubleClicked(self):
        self.CellWasClickedAttachTable()
        self.StartOpenAttachment()

    def cellDoubleClicked(self):
        self.SecelctToTextBrowserEmail()
        self.SelectFromAttachDataToTable()

    def AcceptButton(self):
        self.UpdateEmailOpenOrder()
        self.ImportFromDatabaseAll()
        self.CountVrabote()

    """-------------------------------------"""

    # Открыть окно ответа на email
    def open_reply_window(self):
        self.window = QtWidgets.QMainWindow()
        self.ui = Ui_MainWindow_reply()
        self.ui.setupUi(self.window)
        self.window.show()

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

    """--------------------------------"""

    def ImportFromDatabaseAll(self):
        try:
            show = self.CurrentTextComboboxShow()
            mydb = mc.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                              database=DATABASEMSSQL)

            mycursor = mydb.cursor()

            allnotclose = self.ChekboxEvent()
            param = 'True'
            if allnotclose:
                close_order = 'AND close_order !=' + "'" + param + "'"
            else:
                close_order = ''
            if show == 'All':
                TOP = '*'
            else:
                TOP = 'TOP' + ' ' + show + ' ' + 'id, subject, sender_name, specialist, copy, datetime_send, yes_no_attach'
            mycursor.execute(
                f"SELECT {TOP} FROM email where uid_Division = {self.Data_Division()[0][0]} {close_order} ORDER BY datetime_send DESC ")
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
            item.setText('Назначен специалист')
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
            self.CountVrabote()
            #self.SetAcceptCollor()
        except mc.Error as e:
            print(e)

    def SecelctToTextBrowserEmail(self):
        try:
            mydb = mc.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                              database=DATABASEMSSQL)
            id = self.CellWasClicked()
            if id:
                mycursor = mydb.cursor()
                sql_select_query = mycursor.execute(
                    f"""SELECT text_body, yes_no_attach, copy, sender_name, datetime_send, subject, sender_email  FROM email WHERE id = {id}""")
                result = mycursor.fetchall()
                self.textBrowser_email_1.setText(f'{result[0][0]}')
                self.label_attachments.setVisible(result[0][1])

                self.label_sender.setText(f'{result[0][3]} - {result[0][6]}')
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
                if result[0][2]:
                    subject_email = 'RE: ' + str(result[0][2] + '  ' + f'Id: ##{id}##')
                else:
                    subject_email = 'RE: ' + '  ' + f'Id: ##{id}##'

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
                close_order = False
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

                sql_insert_blob_query = f"""UPDATE email SET specialist = '{specialist}',control_period = '{date}', 
                open_order = '{open_order}', close_order = '{close_order}' WHERE id = '{id}' """
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
            spec = self.CurrentTextComboboxSpec()
            show = self.CurrentTextComboboxShow()
            if show == 'All':
                TOP = ''
            else:
                TOP = 'TOP' + ' ' + show
            if spec == 'Все':
                param = ''
            else:
                param = 'AND specialist =' + "'" + spec + "'"
            sql_select_query = mycursor.execute(
                f"""SELECT {TOP} id, subject, sender_name, specialist, control_period, datetime_send, yes_no_attach FROM 
                email where open_order = 'True' AND uid_Division = {self.Data_Division()[0][0]} {param} ORDER BY control_period DESC""")
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
            # self.radioButton_accepted.setChecked(True)
        except mc.Error as e:
            print(e)
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

                    self.ReplyEmail()

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
            show = self.CurrentTextComboboxShow()
            spec = self.CurrentTextComboboxSpec()
            if show == 'All':
                TOP = ''
            else:
                TOP = 'TOP' + ' ' + show
            if spec == 'Все':
                param = ''
            else:
                param = 'AND specialist =' + "'" + spec + "'"
            sql_select_query = mycursor.execute(
                f"""SELECT {TOP} id, subject, sender_name, specialist, control_period, date_complited, yes_no_attach FROM 
                email where close_order = 'True' AND uid_Division = {self.Data_Division()[0][0]} {param} ORDER BY control_period DESC""")
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
        # rows = self.tableWidget_table.rowCount()
        # for row in range(rows):
        #     self.tableWidget_table.item(row, 0).setBackground(QtGui.QColor(176, 224, 230))
        try:
            mydb = mc.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                              database=DATABASEMSSQL)

            mycursor = mydb.cursor()
            columns = self.tableWidget_table.columnCount()
            rows = self.tableWidget_table.rowCount()
            spec = self.CurrentTextComboboxSpec()
            font = QFont()
            for row in range(rows):
                for column in range(columns):
                    item_id = self.tableWidget_table.item(row, 0).text()
                    sql_select_query = mycursor.execute(
                        f"""SELECT open_order, close_order, specialist FROM email WHERE id = {item_id}""")
                    result = mycursor.fetchall()
                    open_order = result[0][0]
                    close_order = result[0][1]
                    specialist = result[0][2]
                    # if specialist != spec:
                    #     self.tableWidget_table.item(row, column).setBackground(QtGui.QColor(33, 33, 33))
                    if close_order == open_order:
                        font = QFont()
                        font.setBold(True)
                        self.tableWidget_table.item(row, column).setFont(font)
                        self.tableWidget_table.item(row, column).setForeground(QBrush(QColor(0, 88, 150)))
                    if close_order == True and spec == 'Все':
                        font.setStrikeOut(True)
                        self.tableWidget_table.item(row, column).setFont(font)
                    if open_order == True and spec == 'Все':
                        self.tableWidget_table.item(row, column).setBackground(QtGui.QColor(211, 252, 204))

                    if open_order == True and spec == specialist:
                        self.tableWidget_table.item(row, column).setBackground(QtGui.QColor(211, 252, 204))
                        icon = QtGui.QIcon()
                        icon.addPixmap(QtGui.QPixmap(":/images/image/add_icon.png"), QtGui.QIcon.Mode.Normal,
                                       QtGui.QIcon.State.Off)
                        self.tableWidget_table.item(row, 0).setIcon(icon)
                    if close_order == True and spec == specialist:
                        font.setStrikeOut(True)
                        self.tableWidget_table.item(row, column).setFont(font)




        except Exception as erorr:
            print(erorr)

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

    # def SetAcceptCollor(self):
    #     """Подсветка заявок в работе и закрытых во всех заявках"""
    #     try:
    #         mydb = mc.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
    #                           database=DATABASEMSSQL)
    #
    #         mycursor = mydb.cursor()
    #         spec = self.CurrentTextComboboxSpec()
    #         rows = self.tableWidget_table.rowCount()
    #         for row in range(rows):
    #             item_id = self.tableWidget_table.item(row, 0).text()
    #             sql_select_query = mycursor.execute(
    #                 f"SELECT open_order, close_order, specialist FROM email WHERE id = {item_id}")
    #             result = mycursor.fetchall()
    #             open_order = result[0][0]
    #             close_order = result[0][1]
    #             specialist = result[0][2]
    #             if open_order == True and spec == 'Все':
    #                 self.tableWidget_table.item(row, 0).setBackground(QtGui.QColor(211, 252, 204))
    #
    #
    #             if open_order == True and spec == specialist:
    #                 self.tableWidget_table.item(row, 0).setBackground(QtGui.QColor(211, 252, 204))
    #                 icon = QtGui.QIcon()
    #                 icon.addPixmap(QtGui.QPixmap(":/images/image/add_icon.png"), QtGui.QIcon.Mode.Normal,
    #                                QtGui.QIcon.State.Off)
    #                 self.tableWidget_table.item(row, 0).setIcon(icon)
    #             if close_order == True and spec == specialist:
    #                 font = QFont()
    #                 font.setBold(True)
    #                 font.setStrikeOut(True)
    #                 self.tableWidget_table.item(row, 0).setFont(font)
    #                 #self.tableWidget_table.item(row, 0).setBackground(QtGui.QColor(212, 212, 212))
    #                 icon = QtGui.QIcon()
    #                 icon.addPixmap(QtGui.QPixmap(":/images/image/accepted.png"), QtGui.QIcon.Mode.Normal,
    #                                QtGui.QIcon.State.Off)
    #                 self.tableWidget_table.item(row, 0).setIcon(icon)
    #
    #
    #     except Exception as erorr:
    #         print(erorr)

    """Быстрый ответ при закрытии заявки"""

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

        try:
            id_cell = self.CellWasClicked()
            mydb = mc.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                              database=DATABASEMSSQL)

            mycursor = mydb.cursor()
            sql_select_query = mycursor.execute(
                f"""SELECT subject, copy, sender_email, datetime_send, control_period, date_complited, text_body, specialist FROM email WHERE id = {id_cell}""")
            result = mycursor.fetchall()
            if result[0][0]:
                subject = result[0][0] + ' ' + f'Id: ##{id_cell}##'
            else:
                subject = f'Id: ##{id_cell}##'
            if result[0][1]:
                copy = result[0][1].replace(';', '').split()
            else:
                copy = ''
            recipient = result[0][2]
            body = f'Заявка "{id_cell}" закрыта\nЗаявку выполнял - {result[0][7]}\nНазначена - {result[0][3]}\nКонтрольный срок - {result[0][4]}\nЗакрыта - {result[0][5]}\n-------------\n{result[0][6]}'

            sql_insert_blob_query = f"""UPDATE email SET text_body = '{body}' WHERE id = '{id_cell}' """
            res = mycursor.execute(sql_insert_blob_query)
            mydb.commit()
            mycursor.close()
            mydb.close()

        except Exception as s:
            print(s)

        server = SERVEREXCHANGE

        email = self.Data_Division()[0][2]
        username = self.Data_Division()[0][3]
        password = self.Data_Division()[0][4]
        account = connect(server, email, username, password)

        status = self.label_statusneworder

        class ConnectToExchange(object):
            """docstring"""

            def __init__(self, server, email, username, account):
                """Constructor"""

                self.server = server
                self.email = email
                self.username = username
                self.account = account

            def Send(self):

                try:
                    m = Message(account=account, subject=subject, body=body,
                                to_recipients=[Mailbox(email_address=recipient)],
                                cc_recipients=copy)
                    m.send()

                    # status.setText('Сообщение отправлено!')
                    # status.setStyleSheet('color:green')

                except Exception as s:
                    print(s)

        conn = ConnectToExchange(server, email, username, account)
        conn.Send()
