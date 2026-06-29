/**
 * TRYBE-G — Google Apps Script backend for the Profile Sync feature.
 *
 * SET UP
 * 1. Open the Google Sheet you want to use as your member database.
 * 2. Make sure row 1 has these exact headers, in this order:
 *      Timestamp | Name | Email | Phone | Day | Month | Year
 * 3. In the Sheet, go to Extensions -> Apps Script.
 * 4. Delete any starter code and paste in this whole file.
 * 5. Click Deploy -> New deployment.
 *      - Type: Web app
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 6. Copy the Web app URL it gives you (ends in /exec).
 * 7. In index.html, set SHEETS_API_URL to that URL and USE_MOCK_DATA to false.
 *
 * The sheet's columns are intentionally simple so this is easy to read and
 * adapt if your real sheet has different columns or extra fields.
 */

const SHEET_NAME = "Sheet1"; // change if your tab has a different name
const HEADERS = ["Timestamp", "Name", "Email", "Phone", "Day", "Month", "Year"];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet tab "' + SHEET_NAME + '" not found.');
  return sheet;
}

function rowToMember_(row, rowIndex) {
  return {
    rowIndex: rowIndex,           // internal use, not shown to the user
    name: row[1] || "",
    email: row[2] || "",
    phone: row[3] || "",
    day: row[4] || "",
    month: row[5] || "",
    year: row[6] || ""
  };
}

function normalize_(value) {
  return (value || "").toString().trim().toLowerCase();
}

/**
 * GET  ?action=search&query=<email or phone>
 * Returns { found: true, member: {...} } or { found: false }
 */
function doGet(e) {
  const action = e.parameter.action;

  if (action === "search") {
    const query = normalize_(e.parameter.query).replace(/\s/g, "");
    const sheet = getSheet_();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const email = normalize_(row[2]);
      const phone = normalize_(row[3]).replace(/\s/g, "");
      if (email === query || phone === query) {
        return jsonResponse_({ found: true, member: rowToMember_(row, i + 1) });
      }
    }
    return jsonResponse_({ found: false });
  }

  return jsonResponse_({ error: "Unknown action" });
}

/**
 * POST body: { action: "update", member: { rowIndex, name, email, phone, day, month, year } }
 * Updates the matching row in place. Returns { success: true }.
 */
function doPost(e) {
  const body = JSON.parse(e.postData.contents);

  if (body.action === "update") {
    const member = body.member;
    const sheet = getSheet_();

    if (member.rowIndex) {
      // Fast path: we already know which row to update.
      sheet.getRange(member.rowIndex, 2, 1, 6).setValues([[
        member.name, member.email, member.phone, member.day, member.month, member.year
      ]]);
      return jsonResponse_({ success: true });
    }

    // Fallback: find the row by the original email if rowIndex wasn't passed.
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (normalize_(data[i][2]) === normalize_(member.originalEmail || member.email)) {
        sheet.getRange(i + 1, 2, 1, 6).setValues([[
          member.name, member.email, member.phone, member.day, member.month, member.year
        ]]);
        return jsonResponse_({ success: true });
      }
    }
    return jsonResponse_({ success: false, error: "Member not found" });
  }

  return jsonResponse_({ error: "Unknown action" });
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
