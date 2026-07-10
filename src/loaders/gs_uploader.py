import csv
from pathlib import Path
from typing import Sequence

import gspread
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

class GoogleSheetsUploader:
    def __init__(self, creds_path: Path, spreadsheet_name: str = "Вкусвилл"):
        self.client = gspread.authorize(
            Credentials.from_service_account_file(creds_path, scopes=SCOPES)
        )
        try:
            self.sh = self.client.open(spreadsheet_name)
        except gspread.SpreadsheetNotFound:
            self.sh = self.client.create(spreadsheet_name)

    def upload_csv(self, worksheet_name: str, csv_path: Path):
        ws = self._get_or_create_ws(worksheet_name)
        ws.clear()
        with csv_path.open(newline="") as f:
            reader = csv.reader(f, delimiter=";")
            ws.update("A1", list(reader))
        self.sh.share("", perm_type="anyone", role="reader")  # read-only link

    def _get_or_create_ws(self, name: str) -> gspread.Worksheet:
        if name not in [ws.title for ws in self.sh.worksheets()]:
            self.sh.add_worksheet(title=name, rows=1000, cols=20)
        return self.sh.worksheet(name)