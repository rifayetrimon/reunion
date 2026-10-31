// Code.gs — Google Apps Script backend for the Reunion Roll Call

var SHEET_NAME = "Responses";
var USERS_NAME = "Users";

var HEADERS = [
  "Submitted at", "Updated at", "Full name", "Nickname", "Gender",
  "Location", "Mobile", "Profession", "Attending", "Photo URL", "Client ID"
];
var USER_HEADERS = ["Username", "Password Hash", "Client ID", "Created At"];

// Salt — change this once before first deploy, then never again.
var PW_SALT = "psc2011reunion";

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  // Only initialise if the sheet is truly empty (no header row yet)
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
    Logger.log("Responses sheet created.");
  } else {
    Logger.log("Responses sheet already has data — skipped to protect existing entries.");
  }

  var users = ss.getSheetByName(USERS_NAME) || ss.insertSheet(USERS_NAME);
  if (users.getLastRow() === 0) {
    users.appendRow(USER_HEADERS);
    users.getRange(1, 1, 1, USER_HEADERS.length).setFontWeight("bold");
    users.setFrozenRows(1);
    Logger.log("Users sheet created.");
  } else {
    Logger.log("Users sheet already has data — skipped to protect existing entries.");
  }

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

    // Find clientId column from header
    var header  = values[0] || [];
    var cidCol  = header.indexOf("Client ID");
    if (cidCol < 0) cidCol = 10; // fallback

    if (clientId) {
      for (var i = 1; i < values.length; i++) {
        if (String(values[i][cidCol]) === clientId) { existingRow = i + 1; break; }
      }
    }

    var photoCol = header.indexOf("Photo URL");
    if (photoCol < 0) photoCol = 9; // fallback
    var photoUrl = existingRow > 0 ? String(values[existingRow - 1][photoCol] || "") : "";
    if (d.photo && String(d.photo).indexOf("https://") === 0) photoUrl = d.photo;

    var row = [
      existingRow > 0 ? values[existingRow - 1][0] : now,
      now,
      clean_(d.fullName),
      clean_(d.nickname),
      clean_(d.gender),
      clean_(d.location),
      clean_(d.mobile),
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

    var values   = sheet.getDataRange().getValues();
    var clientId = (e && e.parameter && e.parameter.clientId) ? String(e.parameter.clientId) : "";

    // Detect column layout from the header row
    var header    = values[0] || [];
    var hasMobile = header.indexOf("Mobile") >= 0;
    var cidCol    = hasMobile ? 10 : 9; // Client ID column index

    // Return a single user's own entry for pre-filling the form
    if (clientId) {
      for (var i = 1; i < values.length; i++) {
        var r = values[i];
        if (String(r[cidCol]) === clientId) {
          return respond_({ ok: true, entry: {
            fullName:   String(r[2] || ""),
            nickname:   String(r[3] || ""),
            gender:     String(r[4] || ""),
            location:   String(r[5] || ""),
            mobile:     hasMobile ? String(r[6] || "") : "",
            profession: hasMobile ? String(r[7] || "") : String(r[6] || ""),
            attending:  hasMobile ? String(r[8] || "") : String(r[7] || ""),
            photo:      hasMobile ? String(r[9] || "") : String(r[8] || "")
          }});
        }
      }
      return respond_({ ok: true, entry: null });
    }

    // Detect column layout from the header row
    var header = values[0] || [];
    var hasMobile = header.indexOf("Mobile") >= 0;

    // Return full public list
    var entries = [];
    for (var j = 1; j < values.length; j++) {
      var row = values[j];
      if (!row[2]) continue;
      entries.push({
        name:       String(row[2]),
        nickname:   String(row[3] || ""),
        gender:     String(row[4] || ""),
        location:   String(row[5] || ""),
        mobile:     hasMobile ? String(row[6] || "") : "",
        profession: hasMobile ? String(row[7] || "") : String(row[6] || ""),
        attending:  hasMobile ? String(row[8] || "") : String(row[7] || ""),
        photo:      hasMobile ? String(row[9] || "") : String(row[8] || "")
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
