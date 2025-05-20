import gspread
import os
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from google.oauth2.service_account import Credentials
from google.auth.transport.requests import AuthorizedSession
import random
import requests
import pathlib
import logging

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

def remove_all_protections(sheet, spreadsheet_id, authed_session):
    # Fetch all protected ranges and remove those that belong to this sheet
    sheet_metadata = authed_session.get(
        f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}?fields=sheets(protectedRanges,properties(sheetId,title))"
    ).json()



logging.info("Current datetime: %s", datetime.now())
# Setup
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]
creds = Credentials.from_service_account_file("service_account.json", scopes=scope)
authed_session = AuthorizedSession(creds)
client = gspread.authorize(creds)

logging.info("Opening Google Spreadsheet: cleaning_test")
spreadsheet = client.open("cleaning_test")

# Ensure Meta sheet exists for persistent metadata
meta_sheet = spreadsheet.worksheet("Meta") if "Meta" in [s.title for s in spreadsheet.worksheets()] else spreadsheet.add_worksheet(title="Meta", rows="10", cols="2")

try:
    if not meta_sheet.hidden:
        meta_sheet.hide()
        logging.info("Meta sheet successfully hidden after creation.")
    else:
        logging.info("Meta sheet was already hidden.")
except AttributeError:
    # If hidden attribute doesn't exist, always try to hide and log result
    try:
        meta_sheet.hide()
        logging.info("Meta sheet hidden (attribute not available to check status).")
    except Exception as e:
        logging.warning(f"Could not hide Meta sheet: {e}")


# --- Remove existing "Current Week" and "Previous Week" tabs if they exist (move deletion after renaming) ---
logging.info("Checking and removing 'Current Week' and 'Previous Week' sheets if they exist.")
for title in ["Current Week", "Previous Week"]:
    try:
        temp_sheet = spreadsheet.worksheet(title)
        spreadsheet.del_worksheet(temp_sheet)
    except gspread.exceptions.WorksheetNotFound:
        pass

# Find the latest date from existing "Week of" sheets
latest_date = None
week_sheets = []
for sheet in spreadsheet.worksheets():
    if sheet.title.startswith("Week of "):
        week_sheets.append(sheet)
        try:
            date_part = sheet.title.replace("Week of ", "").split()[0]
            parsed_date = datetime.strptime(date_part, "%Y-%m-%d")
            if latest_date is None or parsed_date > latest_date:
                latest_date = parsed_date
        except ValueError:
            continue
 # Use current week's Monday as the anchor date
#simulated_start = datetime

#simulated_start = datetime.now (ZoneInfo("Europe/Berlin")).replace(hour=0, minute=0, second=0, microsecond=0)

simulated_start = datetime(2025, 6, 24).replace(hour=0, minute=0, second=0, microsecond=0)
logging.info(f"Simulated start date set to: {simulated_start}")
monday_of_this_week = simulated_start - timedelta(days=simulated_start.weekday())
test_date = monday_of_this_week
logging.info(f"Calculated Monday of the week: {test_date}")
week_name = f"Week of {test_date.strftime('%Y-%m-%d')} — Current Week"

# Prevent re-creating current week's sheet if it already exists
logging.info("Checking if current or previous week's sheet already exists.")
existing_titles = [sheet.title for sheet in spreadsheet.worksheets()]
new_week_base = f"Week of {test_date.strftime('%Y-%m-%d')}"
if f"{new_week_base} — Current Week" in existing_titles or f"{new_week_base} — Previous Week" in existing_titles:
    logging.info("Current week's sheet already exists. Skipping creation.")
    exit()



# --- After cleaning up old "Week of" sheets, identify latest and rename to "Previous Week" ---
# latest_sheet = None
# latest_date_candidate = None
# for sheet in spreadsheet.worksheets():
#     if sheet.title.startswith("Week of "):
#         try:
#             date_part = sheet.title.replace("Week of ", "").split()[0]
#             parsed_date = datetime.strptime(date_part, "%Y-%m-%d")
#             if latest_date_candidate is None or parsed_date > latest_date_candidate:
#                 latest_date_candidate = parsed_date
#                 latest_sheet = sheet
#         except Exception:
#             continue
# if latest_sheet:
#     latest_sheet.update_title("Previous Week")

#
# Find the most recent "Week of ..." sheet (before the one being created)
previous_week_sheet = None
if week_sheets:
    previous_week_sheet = sorted(
        week_sheets,
        key=lambda s: datetime.strptime(s.title.replace("Week of ", "").split()[0], "%Y-%m-%d"),
        reverse=True
    )[0]






# Disable protection before modifying sheet titles
logging.info("Disabling protection before making changes.")
requests.post("https://script.google.com/macros/s/AKfycbwgurjkVrpVN3ZFhFMczROPZSTMeX-vt6SLcBxHJAQ1veFQp8GpdSpcPuXgn0pb1bKMHw/exec?action=disable")
import time
time.sleep(0.5)


#red stuff is here nowno 
if previous_week_sheet:
    red_rows = []
    for i in range(3, 9):  # Rows 2 to 7
        status = previous_week_sheet.acell(f"C{i}").value
        if status != "TRUE":
            red_rows.append(i)

    if red_rows:
        red_format = {
            "requests": [
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": previous_week_sheet._properties["sheetId"],
                            "startRowIndex": row - 1,
                            "endRowIndex": row,
                            "startColumnIndex": 0,
                            "endColumnIndex": 3
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {
                                    "red": 1,
                                    "green": 0.8,
                                    "blue": 0.8
                                }
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor)"
                    }
                } for row in red_rows
            ]
        }
        authed_session.post(
            f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet.id}:batchUpdate",
            json=red_format
        )
#the end of red stuff


# Clean up weekly sheets older than 2 weeks
# Use simulated current date based on last generated test week
for sheet in spreadsheet.worksheets():
    if sheet.title.startswith("Week of "):
        try:
            date_part = sheet.title.replace("Week of ", "").split()[0]
            sheet_date = datetime.strptime(date_part, "%Y-%m-%d")
            if (test_date - sheet_date).days > 7:
                remove_all_protections(sheet, spreadsheet.id, authed_session)
                spreadsheet.del_worksheet(sheet)
        except ValueError:
            continue  # skip sheets with invalid date format

# Room and task setup
rooms = ["Room 1", "Room 2", "Room 3", "Room 4", "Room 5", "Room 6"]
tasks = [
    "Vaccum and Floor Cleaned",
    "Stove and Counter Top",
    "Tables Cleaned (other surfaces)",
    "Oven - Microwave",
    "Empty Trash Bins"
]
bathroom_task = "Your Side of the Bathroom"
bathroom_room_sequence = ["Room 1", "Room 4", "Room 2", "Room 5", "Room 3", "Room 6"]

 # Determine bathroom rotation index from meta sheet (persistent)
try:
    bathroom_rotation_index = int(meta_sheet.acell("A1").value)
except:
    bathroom_rotation_index = 0
bathroom_room = bathroom_room_sequence[bathroom_rotation_index % len(bathroom_room_sequence)]
logging.info(f"Current number of 'Week of ...' sheets: {len(week_sheets)}")
logging.info(f"Bathroom rotation index: {bathroom_rotation_index}")
logging.info(f"Bathroom assigned to: {bathroom_room}")

 # Assign bathroom task to the bathroom_room
# Assign other tasks randomly to remaining rooms
other_rooms = [room for room in rooms if room != bathroom_room]

random.shuffle(tasks)
# Create a dictionary of room to task
assignment_dict = dict(zip(other_rooms, tasks))
assignment_dict[bathroom_room] = bathroom_task

# Update bathroom index in meta sheet
meta_sheet.update("A1", [[str((bathroom_rotation_index + 1) % len(bathroom_room_sequence))]])
# Optionally, label column B1 for clarity
meta_sheet.update("B1", [["Bathroom Index"]])

# Ensure assignments are in the order of the original rooms list
assignments = [(room, assignment_dict[room]) for room in rooms]


# Delete existing "Previous Week" sheet if it exists to avoid rename conflict
for sheet in spreadsheet.worksheets():
    if sheet.title.endswith("— Previous Week"):
        spreadsheet.del_worksheet(sheet)
        break

# Rename existing "Current" tab to "Previous"
for sheet in spreadsheet.worksheets():
    if sheet.title.endswith("— Current Week"):
        sheet_date = sheet.title.replace("Week of ", "").replace("— Current Week", "").strip()
        sheet.update_title(f"Week of {sheet_date} — Previous Week")

worksheet = spreadsheet.add_worksheet(title=week_name, rows="10", cols="4")

# Determine start and end of the week (7-day period)
start_of_week = test_date.strftime("%B %d, %Y")
end_of_week = (test_date + timedelta(days=6)).strftime("%B %d, %Y")
date_range_text = f"For the week: {start_of_week} → {end_of_week}"

worksheet.clear()
worksheet.merge_cells("A1:C1")
worksheet.update("A1", [[date_range_text]])

# Bold only "For the week" in A1 using rich text formatting
header_text_format_request = {
    "requests": [
        {
            "updateCells": {
                "range": {
                    "sheetId": worksheet._properties["sheetId"],
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": 1
                },
                "rows": [
                    {
                        "values": [
                            {
                                "userEnteredValue": {
                                    "stringValue": date_range_text
                                },
                                "textFormatRuns": [
                                    {
                                        "startIndex": 0,
                                        "format": {"bold": True}
                                    },
                                    {
                                        "startIndex": len("For the week:"),
                                        "format": {"bold": False}
                                    }
                                ],
                                "userEnteredFormat": {
                                    "horizontalAlignment": "CENTER",
                                    "verticalAlignment": "MIDDLE"
                                }
                            }
                        ]
                    }
                ],
                "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment),textFormatRuns,userEnteredValue"
            }
        }
    ]
}

logging.info("Applying header formatting and styling.")
authed_session.post(
    f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet.id}:batchUpdate",
    json=header_text_format_request
)

# Center and style the header banner in A1 (separate from textFormatRuns)
header_banner_format = {
    "requests": [
        {
            "repeatCell": {
                "range": {
                    "sheetId": worksheet._properties["sheetId"],
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": 3
                },
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                        "backgroundColor": {"red": 0.85, "green": 0.92, "blue": 0.97},
                        "textFormat": {
                            "bold": False,
                            "fontFamily": "Helvetica"
                        }
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
            }
        }
    ]
}

authed_session.post(
    f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet.id}:batchUpdate",
    json=header_banner_format
)
# Set column widths for better readability
column_width_requests = {
    "requests": [
        {"updateDimensionProperties": {
            "range": {
                "sheetId": worksheet._properties["sheetId"],
                "dimension": "COLUMNS",
                "startIndex": 0,
                "endIndex": 1
            },
            "properties": {"pixelSize": 120},
            "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {
                "sheetId": worksheet._properties["sheetId"],
                "dimension": "COLUMNS",
                "startIndex": 1,
                "endIndex": 2
            },
            "properties": {"pixelSize": 300},
            "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {
                "sheetId": worksheet._properties["sheetId"],
                "dimension": "COLUMNS",
                "startIndex": 2,
                "endIndex": 3
            },
            "properties": {"pixelSize": 80},
            "fields": "pixelSize"
        }},
    ]
}

logging.info("Applying column widths.")
authed_session.post(
    f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet.id}:batchUpdate",
    json=column_width_requests
)
# Headers
worksheet.update(range_name="A2:C2", values=[["Room", "Tasks", "Done"]])
worksheet.update(range_name="A3:B9", values=assignments)


worksheet.batch_clear(["C3:C9"])

# Add checkboxes using Google Sheets API rules
rule = {
    "requests": [
        {
            "setDataValidation": {
                "range": {
                    "sheetId": worksheet._properties["sheetId"],
                    "startRowIndex": 2,  # Row 3
                    "endRowIndex": 8,    # Row 8
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


logging.info("Adding checkboxes.")
authed_session.post(
    f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet.id}:batchUpdate",
    json=rule
)

# Conditional formatting rule: if checkbox is TRUE, highlight row green
format_rule = {
    "requests": [
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [
                        {
                            "sheetId": worksheet._properties["sheetId"],
                            "startRowIndex": 2,
                            "endRowIndex": 8,
                            "startColumnIndex": 0,
                            "endColumnIndex": 3
                        }
                    ],
                    "booleanRule": {
                        "condition": {
                            "type": "CUSTOM_FORMULA",
                            "values": [
                                {"userEnteredValue": "=$C3=TRUE"}
                            ]
                        },
                        "format": {
                            "backgroundColor": {
                                "red": 0.8,
                                "green": 1,
                                "blue": 0.8
                            }
                        }
                    }
                },
                "index": 0
            }
        }
    ]
}

logging.info("Adding conditional formatting.")
authed_session.post(
    f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet.id}:batchUpdate",
    json=format_rule
)

# --- Formatting block: header, alternating row colors, borders, center alignment ---
format_requests = {
    "requests": [
        # Format header row (bold white text on blue background)
        {
            "repeatCell": {
                "range": {
                    "sheetId": worksheet._properties["sheetId"],
                    "startRowIndex": 1,
                    "endRowIndex": 2,
                    "startColumnIndex": 0,
                    "endColumnIndex": 3
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": 0.0, "green": 0.48, "blue": 0.75},
                        "horizontalAlignment": "CENTER",
                        "textFormat": {"foregroundColor": {"red": 1, "green": 1, "blue": 1}, "bold": True}
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
            }
        },
        # Alternating row colors for data rows
        {
            "addBanding": {
                "bandedRange": {
                    "range": {
                        "sheetId": worksheet._properties["sheetId"],
                        "startRowIndex": 2,
                        "endRowIndex": 8,
                        "startColumnIndex": 0,
                        "endColumnIndex": 3
                    },
                    "rowProperties": {
                        "headerColor": {"red": 0.9, "green": 0.9, "blue": 0.9},
                        "firstBandColor": {"red": 1, "green": 1, "blue": 1},
                        "secondBandColor": {"red": 0.95, "green": 0.95, "blue": 0.95}
                    }
                }
            }
        },
        # Add borders around all cells
        {
            "updateBorders": {
                "range": {
                    "sheetId": worksheet._properties["sheetId"],
                    "startRowIndex": 1,
                    "endRowIndex": 8,
                    "startColumnIndex": 0,
                    "endColumnIndex": 3
                },
                "top": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
                "bottom": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
                "left": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
                "right": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
                "innerHorizontal": {"style": "SOLID", "width": 1, "color": {"red": 0.8, "green": 0.8, "blue": 0.8}},
                "innerVertical": {"style": "SOLID", "width": 1, "color": {"red": 0.8, "green": 0.8, "blue": 0.8}}
            }
        },
        # Center align all cells in column A (Room) data rows
        {
            "repeatCell": {
                "range": {
                    "sheetId": worksheet._properties["sheetId"],
                    "startRowIndex": 2,
                    "endRowIndex": 8,
                    "startColumnIndex": 0,
                    "endColumnIndex": 1
                },
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE"
                    }
                },
                "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment)"
            }
        },
        # Left align all cells in column B (Tasks) data rows
        {
            "repeatCell": {
                "range": {
                    "sheetId": worksheet._properties["sheetId"],
                    "startRowIndex": 2,
                    "endRowIndex": 8,
                    "startColumnIndex": 1,
                    "endColumnIndex": 2
                },
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": "LEFT",
                        "verticalAlignment": "MIDDLE"
                    }
                },
                "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment)"
            }
        },
        # Center align all cells in column C (Done) data rows
        {
            "repeatCell": {
                "range": {
                    "sheetId": worksheet._properties["sheetId"],
                    "startRowIndex": 2,
                    "endRowIndex": 8,
                    "startColumnIndex": 2,
                    "endColumnIndex": 3
                },
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE"
                    }
                },
                "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment)"
            }
        }
    ]
}

logging.info("Applying final styling and borders.")
authed_session.post(
    f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet.id}:batchUpdate",
    json=format_requests
)

# Reapply protection at the end
logging.info("Reapplying protection after making changes.")
requests.post("https://script.google.com/macros/s/AKfycbwgurjkVrpVN3ZFhFMczROPZSTMeX-vt6SLcBxHJAQ1veFQp8GpdSpcPuXgn0pb1bKMHw/exec?action=enable")
logging.info("✅ Script completed successfully.")