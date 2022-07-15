import sys
# Импортируем наш интерфейс из файла
# from app_manager import Ui_MainWindow  # design
# from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt6.QtWidgets import QApplication, QDialog
from Library import System

app = QApplication(sys.argv)

library = System()
library.ImportFromDatabase()


sys.exit(app.exec())
