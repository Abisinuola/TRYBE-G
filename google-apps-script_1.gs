/**
 * TRYBE-G - Google Apps Script backend.
 *
 * Sheet columns (0-indexed) for "Form Responses 1":
 *   0  Timestamp              8  Age group
 *   1  Source                 9  Picture
 *   2  First Name             10 DOB Month
 *   3  Last Name              11 DOB Day
 *   4  Country Code           12 On WhatsApp group?
 *   5  Phone Number           13 Anniversary Month
 *   6  Email Address          14 Anniversary Day
 *   7  Gender                 15 Married?
 *                             16 Employment Category
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
    rowIndex:        rowIndex,
    firstName:       (row[2]  || "").toString().trim(),
    lastName:        (row[3]  || "").toString().trim(),
    countryCode:     (row[4]  || "234").toString().trim(),
    phone:           (row[5]  || "").toString().trim(),
    email:           (row[6]  || "").toString().trim(),
    gender:          (row[7]  || "").toString().trim(),
    ageGroup:        (row[8]  || "").toString().trim(),
    dobMonth:        (row[10] || "").toString().trim(),
    dobDay:          (row[11] || "").toString().trim(),
    anniversaryMonth:(row[13] || "").toString().trim(),
    anniversaryDay:  (row[14] || "").toString().trim(),
    married:         (row[15] || "").toString().trim(),
    employment:      (row[16] || "").toString().trim()
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
      const email = normalize_(row[6]);
      const phone = normalize_(row[5]).replace(/\s/g, "");
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

    const sheet  = getMemberSheet_();
    const data   = sheet.getDataRange().getValues();
    const matches = [];

    for (let i = 1; i < data.length; i++) {
      const row    = data[i];
      const rFirst = normalize_(row[2]);
      const rLast  = normalize_(row[3]);
      if ((firstName && rFirst === firstName) || (lastName && rLast === lastName)) {
        matches.push({
          rowIndex:  i + 1,
          firstName: (row[2] || "").toString().trim(),
          lastName:  (row[3] || "").toString().trim(),
          email:     maskEmail_(row[6]),
          phone:     (row[5] || "").toString().replace(/.(?=.{4})/g, "*"),
          dobMonth:  (row[10] || "").toString().trim(),
          dobDay:    (row[11] || "").toString().trim(),
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
      body.firstName        || "",
      body.lastName         || "",
      body.countryCode      || "234",
      body.phone            || "",
      body.email            || "",
      body.gender           || "",
      body.ageGroup         || "",
      "",
      body.dobMonth         || "",
      body.dobDay           || "",
      "",
      body.anniversaryMonth || "",
      body.anniversaryDay   || "",
      body.married          || "",
      body.employment       || ""
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

  function writeRow_(rowIndex) {
    sheet.getRange(rowIndex, 3, 1, 5).setValues([[
      member.firstName, member.lastName, member.countryCode, member.phone, member.email
    ]]);
    sheet.getRange(rowIndex, 9, 1, 1).setValues([[member.ageGroup]]);
    sheet.getRange(rowIndex, 11, 1, 2).setValues([[member.dobMonth, member.dobDay]]);
    sheet.getRange(rowIndex, 14, 1, 4).setValues([[
      member.anniversaryMonth, member.anniversaryDay, member.married, member.employment
    ]]);
  }

  if (member.rowIndex) {
    writeRow_(member.rowIndex);
    return jsonResponse_({ success: true });
  }

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (normalize_(data[i][6]) === normalize_(member.originalEmail || member.email)) {
      writeRow_(i + 1);
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
