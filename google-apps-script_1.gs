/**
 * TRYBE-G — Google Apps Script backend for the Profile Sync feature.
 *
 * SET UP
 * 1. Open the Google Sheet used as the member database.
 * 2. In the Sheet, go to Extensions > Apps Script.
 * 3. Delete any starter code and paste this whole file.
 * 4. Click Deploy > New deployment.
 *      - Type: Web app
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 5. Copy the Web app URL (ends in /exec).
 * 6. In index.html, set SHEETS_API_URL to that URL.
 *
 * Sheet columns (0-indexed):
 *   0  Timestamp
 *   1  Source
 *   2  First Name
 *   3  Last Name
 *   4  Phone Number (WhatsApp enabled)
 *   5  Email Address
 *   6  Gender
 *   7  Age group
 *   8  Upload a recent picture of you
 *   9  Date of Birth (Month only)
 *   10 Date of Birth (Day only)
 *   11 Are you on the Youth WhatsApp group?
 *   12 Wedding Anniversary (Month only)
 *   13 Wedding Anniversary (Day only)
 *   14 Are you married?
 *   15 Employment Status
 */

const SHEET_NAME = "Form Responses 1";

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet tab "' + SHEET_NAME + '" not found.');
  return sheet;
}

function rowToMember_(row, rowIndex) {
  return {
    rowIndex:  rowIndex,
    firstName: (row[2] || "").toString().trim(),
    lastName:  (row[3] || "").toString().trim(),
    phone:     (row[4] || "").toString().trim(),
    email:     (row[5] || "").toString().trim(),
    dobMonth:  (row[9] || "").toString().trim(),
    dobDay:    (row[10] || "").toString().trim()
  };
}

function normalize_(value) {
  return (value || "").toString().trim().toLowerCase();
}

function doGet(e) {
  const action = e.parameter.action;

  if (action === "search") {
    const query = normalize_(e.parameter.query).replace(/\s/g, "");
    const sheet = getSheet_();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const email = normalize_(row[5]);
      const phone = normalize_(row[4]).replace(/\s/g, "");
      if (email === query || phone === query) {
        return jsonResponse_({ found: true, member: rowToMember_(row, i + 1) });
      }
    }
    return jsonResponse_({ found: false });
  }

  return jsonResponse_({ error: "Unknown action" });
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);

  if (body.action === "update") {
    const member = body.member;
    const sheet = getSheet_();

    if (member.rowIndex) {
      sheet.getRange(member.rowIndex, 3, 1, 4).setValues([[
        member.firstName, member.lastName, member.phone, member.email
      ]]);
      sheet.getRange(member.rowIndex, 10, 1, 2).setValues([[
        member.dobMonth, member.dobDay
      ]]);
      return jsonResponse_({ success: true });
    }

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (normalize_(data[i][5]) === normalize_(member.originalEmail || member.email)) {
        sheet.getRange(i + 1, 3, 1, 4).setValues([[
          member.firstName, member.lastName, member.phone, member.email
        ]]);
        sheet.getRange(i + 1, 10, 1, 2).setValues([[
          member.dobMonth, member.dobDay
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
