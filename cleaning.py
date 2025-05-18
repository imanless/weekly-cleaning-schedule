import gspread
import random
from datetime import datetime
from google.oauth2.service_account import Credentials
from google.auth.transport.requests import AuthorizedSession

# Setup
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]
creds = Credentials.from_service_account_file("service_account.json", scopes=scope)
client = gspread.authorize(creds)

# Open the Google Spreadsheet
spreadsheet = client.open("cleaning_test")

# Room and task setup
rooms = ["Room 1", "Room 2", "Room 3", "Room 4", "Room 5", "Room 6"]
tasks = [
    "Vaccum",
    "Floor Cleaned",
    "Stove Cleaned",
    "Tables Cleaned (other surfaces)",
    "Ovens",
    "Empty Trash Bins"
]

# Shuffle and assign
random.shuffle(tasks)
assignments = list(zip(rooms, tasks[:6]))

# Create new sheet for the week
# week_name = f"Week of {datetime.today().strftime('%Y-%m-%d')}"
week_name = f"Test {random.randint(1000, 9999)}"
try:
    worksheet = spreadsheet.add_worksheet(title=week_name, rows="10", cols="4")
except gspread.exceptions.APIError:
    worksheet = spreadsheet.worksheet(week_name)
try:
    worksheet = spreadsheet.add_worksheet(title=week_name, rows="10", cols="4")
except gspread.exceptions.APIError:
    worksheet = spreadsheet.worksheet(week_name)

worksheet.clear() 
# Headers
worksheet.update(range_name="A1:C1", values=[["Room", "Assigned Task", "Done"]])
worksheet.update(range_name="A2:B7", values=assignments)


worksheet.batch_clear(["C2:C7"])

# Add checkboxes using Google Sheets API rules
rule = {
    "requests": [
        {
            "setDataValidation": {
                "range": {
                    "sheetId": worksheet._properties["sheetId"],
                    "startRowIndex": 1,  # Row 2
                    "endRowIndex": 7,    # Row 7
                    "startColumnIndex": 2,  # Column C
                    "endColumnIndex": 3,
                },
                "rule": {
                    "condition": {"type": "BOOLEAN"},
                    "showCustomUi": True
                }
            }
        }
    ]
}


authed_session = AuthorizedSession(creds)
authed_session.post(
    f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet.id}:batchUpdate",
    json=rule
)