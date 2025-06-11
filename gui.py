import sys
import os
import webbrowser
import numpy as np
import pandas as pd
from pathlib import Path
from PySide6.QtWidgets import QMainWindow, QWidget, QPushButton, QVBoxLayout, QWidget, QFileDialog, QLabel
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtCore import QUrl
from sheets_manager import SheetsManager
from barcode_functions import main
from googleapiclient.discovery import build
from google.oauth2 import service_account

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS  # Used by PyInstaller when running as exe
    except AttributeError:
        base_path = os.path.abspath(".")  # Normal script run
    return os.path.join(base_path, relative_path)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Google Sheets GUI App")
        self.resize(1200, 800)

        # Spreadsheet info
        self.service_account_file = resource_path("sturdy-willow-428123-s4-427c9bbca449.json")
        self.sheet_id = '1gUvh6oFqm2YfSE7eo7ZhWWkRji607fbPKDjduOIeIyU'  # e.g., 1aBcD... (from the URL)
        self.sheet_url = f"https://docs.google.com/spreadsheets/d/{self.sheet_id}/edit?usp=sharing"
        self.sheet_manager = SheetsManager(self.service_account_file, self.sheet_id) 
        creds = service_account.Credentials.from_service_account_file(
        self.service_account_file,
        scopes=['https://www.googleapis.com/auth/spreadsheets']
        )
        self.service = build('sheets', 'v4', credentials=creds)

        # Setup widgets
        self.instructions = QLabel("NOTE: Stay on Sheet1 (GID = 0) and DO NOT RENAME Sheet1")
        self.instructions.setFixedHeight(50)
        self.instructions.setStyleSheet("font-weight: bold;")
        self.webview = QWebEngineView()
        self.reset_button = QPushButton("Reset Google Sheet")
        self.import_button = QPushButton("Import .csv File")
        self.print_button = QPushButton("Print Barcodes")

        # Layout
        layout = QVBoxLayout()
        layout.addWidget(self.instructions)
        layout.addWidget(self.reset_button)
        layout.addWidget(self.import_button)
        layout.addWidget(self.print_button)
        layout.addWidget(self.webview)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        # Connect
        self.reset_button.clicked.connect(self.reset_sheet)
        self.import_button.clicked.connect(self.import_csv)
        self.print_button.clicked.connect(lambda: self.export_to_pdf(self.service_account_file, self.sheet_id))

    
    def reset_sheet(self):
        # Template content
        template_data = [["Serial Number", "Title"]]

        # Update the sheet
        self.sheet_manager.reset_sheet("Sheet1", template_data)

        # Load the sheet into browser
        self.webview.setUrl(QUrl(self.sheet_url))

    def load_sheet(self):
        self.webview.setUrl(QUrl(self.sheet_url))
    
    def import_csv(self):
        downloads_folder = str(Path.home() / "Downloads")
        fileName, _ = QFileDialog.getOpenFileName(self, "Open CSV File", downloads_folder, "CSV Files (*.csv);;All Files (*)")
        df = pd.read_csv(fileName)
        df = df.replace({np.nan:""})
        rows = [df.columns.tolist()] + df.values.tolist()
        body = {
            'values': rows
        }
        self.service.spreadsheets().values().update(
            spreadsheetId=self.sheet_id,
            range=f"Sheet1!A1",
            valueInputOption="RAW",
            body=body
        ).execute()

    def export_to_pdf(self, service_account_file, file_id):
        creds = service_account.Credentials.from_service_account_file(service_account_file, scopes=["https://www.googleapis.com/auth/drive.readonly"])
        drive_service = build("drive", "v3", credentials=creds)

        # Export the given sheet tab as CSV
        request = drive_service.files().export_media(fileId=file_id, mimeType="text/csv")
        response = request.execute()

        # Save to file
        output_csv_path = "barcodes.csv"
        with open(output_csv_path, "wb") as f:
            f.write(response)
        print(f"Exported sheet to {output_csv_path}")
        current_folder = os.path.dirname(os.path.abspath(__file__))
        pdf_path = current_folder + "/barcodes.pdf"
        main(output_csv_path, pdf_path=pdf_path)
        webbrowser.open_new(r'file://' + pdf_path)