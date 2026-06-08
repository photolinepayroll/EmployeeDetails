// ─────────────────────────────────────────────────────────────
//  SETUP INSTRUCTIONS
//  1. Open your Google Sheet
//  2. Click Extensions → Apps Script
//  3. Paste this entire file, replacing any existing code
//  4. Click Save, then Deploy → New deployment
//     - Type: Web app
//     - Execute as: Me
//     - Who has access: Anyone
//  5. Copy the Web App URL and paste it into index.html
//     where it says: const APPS_SCRIPT_URL = "PASTE_YOUR_URL_HERE";
// ─────────────────────────────────────────────────────────────

const SHEET_NAME = "Sheet1"; // Change if your sheet tab has a different name
const HEADER_ROW = 1;        // Row number of your header

function doGet(e) {
  const action = e.parameter.action;

  if (action === "getNames") {
    return getNames();
  }

  if (action === "checkRecord") {
    const row = parseInt(e.parameter.row);
    return checkRecord(row);
  }

  return jsonResponse({ error: "Unknown action" });
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  return saveRecord(data.row, data.dob, data.age, data.email, data.mobile, data.gender);
}

// Returns all employee names + their row numbers
function getNames() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, 4).getValues();

  const names = [];
  data.forEach((row, i) => {
    const lastName   = (row[0] || "").toString().trim();
    const firstName  = (row[1] || "").toString().trim();
    const middleName = (row[2] || "").toString().trim();
    const suffix     = (row[3] || "").toString().trim();

    if (lastName || firstName) {
      let fullName = lastName + ", " + firstName;
      if (middleName) fullName += " " + middleName;
      if (suffix)     fullName += " " + suffix;
      names.push({ name: fullName.trim(), row: HEADER_ROW + 1 + i });
    }
  });

  return jsonResponse({ names });
}

// Checks if a row already has data in columns E-I
function checkRecord(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const values = sheet.getRange(row, 5, 1, 5).getValues()[0]; // Columns E to I
  const hasRecord = values.some(v => v !== "" && v !== null);
  return jsonResponse({ hasRecord });
}

// Saves DOB, Age, Email, Mobile, Gender into the correct row
function saveRecord(row, dob, age, email, mobile, gender) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  sheet.getRange(row, 5).setValue(dob);    // E - Date of Birth
  sheet.getRange(row, 6).setValue(age);    // F - Age
  sheet.getRange(row, 7).setValue(email);  // G - Personal Email
  sheet.getRange(row, 8).setValue(mobile); // H - Mobile Number
  sheet.getRange(row, 9).setValue(gender); // I - Gender
  return jsonResponse({ success: true });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
