import sys
from PySide6.QtWidgets import QApplication
from gui import MainWindow

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.load_sheet()
    window.show()
    sys.exit(app.exec())