from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

class SheetsManager:
    def __init__(self, creds_path, spreadsheet_id):
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        self.creds = Credentials.from_service_account_file(creds_path, scopes=SCOPES)
        self.service = build('sheets', 'v4', credentials = self.creds)
        self.spreadsheet_id = spreadsheet_id

    def reset_sheet(self, sheet_name, template_data):
        # Clear the sheet
        self.service.spreadsheets().values().clear(
            spreadsheetId=self.spreadsheet_id,
            range=sheet_name
        ).execute()

        # Write new template data
        body = {'values': template_data}
        self.service.spreadsheets().values().update(
            spreadsheetId=self.spreadsheet_id,
            range=f'{sheet_name}!A1',
            valueInputOption='RAW',
            body=body
        ).execute()
