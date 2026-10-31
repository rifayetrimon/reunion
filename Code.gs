// Code.gs — Google Apps Script backend for the Reunion Roll Call
// Photos are hosted on Cloudinary; this script only stores the returned URL.

var SHEET_NAME = "Responses";

var HEADERS = [
  "Submitted at", "Updated at", "Full name", "Nickname", "Gender",
  "Location", "Profession", "Attending", "Photo URL", "Client ID"
];

// Run once from the Apps Script editor to create the Responses sheet.
function setup() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sheet.clearContents();
  sheet.appendRow(HEADERS);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
  sheet.setFrozenRows(1);
  Logger.log("Setup complete on sheet: " + sheet.getName());
}

// Prepend apostrophe to values that start with a spreadsheet formula character.
function clean_(v) {
  var s = String(v == null ? "" : v).trim();
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

// POST — called by the browser. Body is JSON sent as text/plain (no CORS preflight).
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);

    // Honeypot: bots fill the hidden "website" field; real users leave it empty.
    if (d.website) return respond_({ ok: true });

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) { setup(); sheet = ss.getSheetByName(SHEET_NAME); }

    var clientId = clean_(d.clientId || "");
    var now      = new Date();

    // Find an existing row for this browser so we update instead of duplicate.
    var values      = sheet.getDataRange().getValues();
    var existingRow = -1;
    if (clientId) {
      for (var i = 1; i < values.length; i++) {
        if (String(values[i][9]) === clientId) { existingRow = i + 1; break; }
      }
    }

    // Keep the existing photo URL unless a new one arrives (Cloudinary https://… URL).
    var photoUrl = existingRow > 0 ? String(values[existingRow - 1][8] || "") : "";
    if (d.photo && String(d.photo).indexOf("https://") === 0) {
      photoUrl = d.photo;
    }

    var row = [
      existingRow > 0 ? values[existingRow - 1][0] : now,  // Submitted at (keep original)
      now,                                                   // Updated at
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

// GET — returns the public list (no private data).
function doGet(e) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return respond_({ ok: true, entries: [] });

    var values  = sheet.getDataRange().getValues();
    var entries = [];
    // Columns: [0] Submitted at, [1] Updated at, [2] Full name, [3] Nickname, [4] Gender,
    //          [5] Location, [6] Profession, [7] Attending, [8] Photo URL, [9] Client ID
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
