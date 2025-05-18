function onOpen() {
  const sheetName = getOrCreateWeeklySheet();
  assignTasks(sheetName);
}

function getOrCreateWeeklySheet() {
  const ss = SpreadsheetApp.getActive();
  const today = new Date();
  const weekStart = new Date(
    today.getFullYear(),
    today.getMonth(),
    today.getDate() - today.getDay()
  ); // Sunday
  const weekName = `Week of ${weekStart.toISOString().slice(0, 10)}`;

  const existingSheet = ss.getSheets().find((s) => s.getName() === weekName);
  if (existingSheet) return weekName;

  const sheet = ss.insertSheet(weekName);
  sheet.appendRow(["Room", "Assigned Task", "Done"]);
  return weekName;
}

function assignTasks(sheetName) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(sheetName);

  const values = sheet.getRange("A2:A").getValues().flat().filter(String);
  if (values.length >= 6) return;

  const rooms = ["Room 1", "Room 2", "Room 3", "Room 4", "Room 5", "Room 6"];
  const tasks = [
    "Vaccum",
    "Floor Cleaned",
    "Stove Cleaned",
    "Tables Cleaned (other surfaces)",
    "Ovens",
    "Empty Trash Bins",
  ];

  const shuffled = tasks.sort(() => Math.random() - 0.5);
  for (let i = 0; i < rooms.length; i++) {
    sheet.appendRow([rooms[i], shuffled[i], false]);
  }

  const checkboxRange = sheet.getRange("C2:C7");
  checkboxRange.insertCheckboxes();

  const range = sheet.getRange("A2:C7");
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$C2=TRUE")
    .setBackground("#b7fdb7")
    .setRanges([sheet.getRange("A2:C")])
    .build();
  sheet.setConditionalFormatRules([rule]);
}
