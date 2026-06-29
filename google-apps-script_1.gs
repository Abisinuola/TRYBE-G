/**
 * TRYBE-G - Google Apps Script backend.
 *
 * Sheet columns (0-indexed) for "Form Responses 1":
 *   0  Timestamp          6  Gender
 *   1  Source             7  Age group
 *   2  First Name         8  Picture upload
 *   3  Last Name          9  Date of Birth (Month)
 *   4  Phone Number       10 Date of Birth (Day)
 *   5  Email Address      11+ other fields
 */

const MEMBER_SHEET   = "Form Responses 1";
const SHIRT_SHEET    = "Shirt Interest";
const FEEDBACK_SHEET = "Meeting Feedback";
const CONTACT_SHEET  = "Contact";
const CONTACT_EMAIL  = "youths@highlandchurch.com.ng";

function getOrCreateSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

function getMemberSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MEMBER_SHEET);
  if (!sheet) throw new Error('Sheet "' + MEMBER_SHEET + '" not found.');
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

function maskEmail_(email) {
  const e = normalize_(email);
  const at = e.indexOf("@");
  if (at < 2) return "***";
  return e.slice(0, 2) + "***" + e.slice(at);
}

function doGet(e) {
  if (e.parameter.action === "search") {
    const query = normalize_(e.parameter.query).replace(/\s/g, "");
    const sheet = getMemberSheet_();
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

  if (e.parameter.action === "fuzzySearch") {
    const firstName = normalize_(e.parameter.firstName || "");
    const lastName  = normalize_(e.parameter.lastName  || "");
    if (!firstName && !lastName) return jsonResponse_({ matches: [] });

    const sheet = getMemberSheet_();
    const data  = sheet.getDataRange().getValues();
    const matches = [];

    for (let i = 1; i < data.length; i++) {
      const row    = data[i];
      const rFirst = normalize_(row[2]);
      const rLast  = normalize_(row[3]);
      const firstMatch = firstName && rFirst === firstName;
      const lastMatch  = lastName  && rLast  === lastName;
      if (firstMatch || lastMatch) {
        matches.push({
          rowIndex:  i + 1,
          firstName: (row[2] || "").toString().trim(),
          lastName:  (row[3] || "").toString().trim(),
          email:     maskEmail_(row[5]),
          phone:     (row[4] || "").toString().replace(/.(?=.{4})/g, "*"),
          dobMonth:  (row[9]  || "").toString().trim(),
          dobDay:    (row[10] || "").toString().trim(),
          _rowIndex: i + 1
        });
      }
    }
    return jsonResponse_({ matches });
  }

  return jsonResponse_({ error: "Unknown action" });
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);

  if (body.action === "update") {
    return handleUpdate_(body.member);
  }

  if (body.action === "register") {
    const sheet = getMemberSheet_();
    sheet.appendRow([
      new Date(),
      "Website",
      body.firstName || "",
      body.lastName  || "",
      body.phone     || "",
      body.email     || "",
      "",
      "",
      "",
      body.dobMonth  || "",
      body.dobDay    || ""
    ]);
    return jsonResponse_({ success: true });
  }

  if (body.action === "shirtInterest") {
    const sheet = getOrCreateSheet_(SHIRT_SHEET, ["Timestamp", "Email"]);
    sheet.appendRow([new Date(), body.email || ""]);
    return jsonResponse_({ success: true });
  }

  if (body.action === "feedback") {
    const sheet = getOrCreateSheet_(FEEDBACK_SHEET, ["Timestamp", "Message"]);
    sheet.appendRow([new Date(), body.message || ""]);
    return jsonResponse_({ success: true });
  }

  if (body.action === "contact") {
    const sheet = getOrCreateSheet_(CONTACT_SHEET, ["Timestamp", "Name", "Email", "Message"]);
    sheet.appendRow([new Date(), body.name || "", body.email || "", body.message || ""]);
    try {
      MailApp.sendEmail({
        to: CONTACT_EMAIL,
        subject: "TRYBE-G Contact: " + (body.name || "Someone"),
        body: "From: " + (body.name || "") + "\nEmail: " + (body.email || "") + "\n\n" + (body.message || "")
      });
    } catch (_) {}
    return jsonResponse_({ success: true });
  }

  return jsonResponse_({ error: "Unknown action" });
}

function handleUpdate_(member) {
  const sheet = getMemberSheet_();

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

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
