// Code.gs — Google Apps Script backend for the Reunion Roll Call

var SHEET_NAME = "Responses";
var USERS_NAME = "Users";

var HEADERS = [
  "Submitted at", "Updated at", "Full name", "Nickname", "Gender",
  "Location", "Profession", "Attending", "Photo URL", "Client ID"
];
var USER_HEADERS = ["Username", "Password Hash", "Client ID", "Created At"];

// Salt — change this once before first deploy, then never again.
var PW_SALT = "psc2011reunion";

function setup() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sheet.clearContents();
  sheet.appendRow(HEADERS);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
  sheet.setFrozenRows(1);

  var users = ss.getSheetByName(USERS_NAME) || ss.insertSheet(USERS_NAME);
  users.clearContents();
  users.appendRow(USER_HEADERS);
  users.getRange(1, 1, 1, USER_HEADERS.length).setFontWeight("bold");
  users.setFrozenRows(1);
  Logger.log("Setup complete.");
}

function clean_(v) {
  var s = String(v == null ? "" : v).trim();
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function hashPw_(password) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    password + PW_SALT
  );
  return bytes.map(function(b) {
    return ("0" + (b & 0xFF).toString(16)).slice(-2);
  }).join("");
}

function ensureUsersSheet_(ss) {
  var sheet = ss.getSheetByName(USERS_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(USERS_NAME);
    sheet.appendRow(USER_HEADERS);
    sheet.getRange(1, 1, 1, USER_HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function handleSignup_(d) {
  var username     = String(d.username     || "").trim().toLowerCase();
  var passwordHash = String(d.passwordHash || "");
  var clientId     = clean_(d.clientId     || "");

  if (!username || username.length < 3)
    return respond_({ ok: false, error: "Username must be at least 3 characters." });
  if (!/^[a-z0-9_]+$/.test(username))
    return respond_({ ok: false, error: "Only letters, numbers and underscore allowed." });
  if (!passwordHash)
    return respond_({ ok: false, error: "Password is required." });

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ensureUsersSheet_(ss);
  var rows  = sheet.getDataRange().getValues();

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === username)
      return respond_({ ok: false, error: "Username already taken. Try another." });
  }

  sheet.appendRow([username, passwordHash, clientId, new Date()]);
  return respond_({ ok: true, username: username, clientId: clientId });
}

function handleSignin_(d) {
  var username     = String(d.username     || "").trim().toLowerCase();
  var passwordHash = String(d.passwordHash || "");

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ensureUsersSheet_(ss);
  var rows  = sheet.getDataRange().getValues();

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === username) {
      if (String(rows[i][1]) === passwordHash)
        return respond_({ ok: true, username: username, clientId: String(rows[i][2]) });
      else
        return respond_({ ok: false, error: "Incorrect password." });
    }
  }
  return respond_({ ok: false, error: "Username not found." });
}

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);

    if (d.action === "signup") return handleSignup_(d);
    if (d.action === "signin") return handleSignin_(d);

    // Honeypot
    if (d.website) return respond_({ ok: true });

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) { setup(); sheet = ss.getSheetByName(SHEET_NAME); }

    var clientId    = clean_(d.clientId || "");
    var now         = new Date();
    var values      = sheet.getDataRange().getValues();
    var existingRow = -1;

    if (clientId) {
      for (var i = 1; i < values.length; i++) {
        if (String(values[i][9]) === clientId) { existingRow = i + 1; break; }
      }
    }

    var photoUrl = existingRow > 0 ? String(values[existingRow - 1][8] || "") : "";
    if (d.photo && String(d.photo).indexOf("https://") === 0) photoUrl = d.photo;

    var row = [
      existingRow > 0 ? values[existingRow - 1][0] : now,
      now,
      clean_(d.fullName),
      clean_(d.nickname),
      clean_(d.gender),
      clean_(d.location),
      clean_(d.profession),
      clean_(d.attending),
      photoUrl,
      clientId
    ];

    if (existingRow > 0) {
      sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    return respond_({ ok: true });
  } catch (err) {
    return respond_({ ok: false, error: err.message });
  }
}

function doGet(e) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return respond_({ ok: true, entries: [] });

    var values  = sheet.getDataRange().getValues();
    var entries = [];
    for (var i = 1; i < values.length; i++) {
      var r = values[i];
      if (!r[2]) continue;
      entries.push({
        name:       String(r[2]),
        nickname:   String(r[3] || ""),
        gender:     String(r[4] || ""),
        location:   String(r[5] || ""),
        profession: String(r[6] || ""),
        attending:  String(r[7] || ""),
        photo:      String(r[8] || "")
      });
    }
    return respond_({ ok: true, entries: entries });
  } catch (err) {
    return respond_({ ok: false, error: err.message, entries: [] });
  }
}

function respond_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
