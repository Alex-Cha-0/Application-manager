import sys
from PyQt6.QtWidgets import QApplication, QDialog
from Library import System

app = QApplication(sys.argv)

library = System()

library.ImportFromDatabaseAll()
library.ComboBoxSpec()
library.CountVrabote()
library.loadSetting()
sys.exit(app.exec())
