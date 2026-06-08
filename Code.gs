const HEADER_ROW = 1;

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

function doGet(e) {
  const action   = e.parameter.action;
  const callback = e.parameter.callback; // JSONP support

  let result;
  if      (action === "getNames")    result = getNamesData();
  else if (action === "checkRecord") result = checkRecordData(parseInt(e.parameter.row));
  else if (action === "submit")      result = saveRecordData(
    parseInt(e.parameter.row),
    e.parameter.dob, e.parameter.age,
    e.parameter.email, e.parameter.mobile, e.parameter.gender
  );
  else                               result = { error: "Unknown action" };

  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + JSON.stringify(result) + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const result = saveRecordData(data.row, data.dob, data.age, data.email, data.mobile, data.gender);
  const callback = e.parameter && e.parameter.callback;
  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + JSON.stringify(result) + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getNamesData() {
  const sheet   = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) return { names: [], total: 0 };

  const numRows = lastRow - HEADER_ROW;
  const data    = sheet.getRange(HEADER_ROW + 1, 1, numRows, 4).getValues();
  const names   = [];

  data.forEach((row, i) => {
    const lastName   = (row[0] || "").toString().trim();
    const firstName  = (row[1] || "").toString().trim();
    const middleName = (row[2] || "").toString().trim();
    const suffix     = (row[3] || "").toString().trim();

    if (!lastName && !firstName && !middleName) return;

    let fullName = "";
    if (lastName)   fullName += lastName;
    if (firstName)  fullName += (fullName ? ", " : "") + firstName;
    if (middleName) fullName += " " + middleName;
    if (suffix)     fullName += " " + suffix;

    names.push({ name: fullName.trim(), row: HEADER_ROW + 1 + i });
  });

  return { names: names, total: names.length };
}

function checkRecordData(row) {
  const sheet  = getSheet();
  const values = sheet.getRange(row, 5, 1, 5).getValues()[0];
  const hasRecord = values.some(v => v !== "" && v !== null && v !== undefined);
  return { hasRecord: hasRecord };
}

function saveRecordData(row, dob, age, email, mobile, gender) {
  const sheet = getSheet();
  sheet.getRange(row, 5).setValue(dob);
  sheet.getRange(row, 6).setValue(age);
  sheet.getRange(row, 7).setValue(email);
  sheet.getRange(row, 8).setValue(mobile);
  sheet.getRange(row, 9).setValue(gender);
  return { success: true };
}

function testGetNames() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  Logger.log("Sheet: " + sheet.getName());
  Logger.log("Last row: " + sheet.getLastRow());
  const result = getNamesData();
  Logger.log("Names found: " + result.total);
}
