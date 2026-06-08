const HEADER_ROW = 1;

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

function doGet(e) {
  const action = e.parameter.action;
  if (action === "getNames")    return getNames();
  if (action === "checkRecord") return checkRecord(parseInt(e.parameter.row));
  return jsonResponse({ error: "Unknown action" });
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  return saveRecord(data.row, data.dob, data.age, data.email, data.mobile, data.gender);
}

function getNames() {
  const sheet   = getSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow <= HEADER_ROW) return jsonResponse({ names: [] });

  // Get all columns A to D from row 2 to last row
  const numRows = lastRow - HEADER_ROW;
  const data    = sheet.getRange(HEADER_ROW + 1, 1, numRows, 4).getValues();

  const names = [];
  data.forEach((row, i) => {
    const lastName   = (row[0] || "").toString().trim();
    const firstName  = (row[1] || "").toString().trim();
    const middleName = (row[2] || "").toString().trim();
    const suffix     = (row[3] || "").toString().trim();

    // Include ANY row that has at least one name field filled
    const anyName = lastName || firstName || middleName;
    if (!anyName) return;

    let fullName = "";
    if (lastName)   fullName += lastName;
    if (firstName)  fullName += (fullName ? ", " : "") + firstName;
    if (middleName) fullName += " " + middleName;
    if (suffix)     fullName += " " + suffix;

    names.push({ name: fullName.trim(), row: HEADER_ROW + 1 + i });
  });

  return jsonResponse({ names: names, total: names.length, lastRow: lastRow });
}

function checkRecord(row) {
  const sheet  = getSheet();
  const values = sheet.getRange(row, 5, 1, 5).getValues()[0];
  const hasRecord = values.some(v => v !== "" && v !== null && v !== undefined);
  return jsonResponse({ hasRecord: hasRecord });
}

function saveRecord(row, dob, age, email, mobile, gender) {
  const sheet = getSheet();
  sheet.getRange(row, 5).setValue(dob);
  sheet.getRange(row, 6).setValue(age);
  sheet.getRange(row, 7).setValue(email);
  sheet.getRange(row, 8).setValue(mobile);
  sheet.getRange(row, 9).setValue(gender);
  return jsonResponse({ success: true });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
