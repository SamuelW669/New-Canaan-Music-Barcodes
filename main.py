import sys
from PySide6.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtCore import QUrl
from SheetsManager import SheetsManager
from GUI import MainWindow
from ButtonHolder import ButtonHolder

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.load_sheet()
    window.show()
    sys.exit(app.exec())
