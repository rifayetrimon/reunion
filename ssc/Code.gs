// Code.gs — Google Apps Script backend for SSC Batch 2016

var SHEET_NAME = "Responses";
var USERS_NAME = "Users";

// Col 0=SubmittedAt 1=UpdatedAt 2=FullName 3=Nickname 4=Gender
// 5=Group 6=BloodGroup 7=Location 8=Mobile 9=Profession
// 10=JobPlace 11=University 12=MaritalStatus 13=PhotoURL 14=ClientID 15=Status
var HEADERS = [
  "Submitted at", "Updated at", "Full name", "Nickname", "Gender",
  "Group", "Blood Group", "Location", "Mobile", "Profession",
  "Job Place", "University", "Marital Status", "Photo URL", "Client ID", "Status"
];
var USER_HEADERS = ["Username", "Password Hash", "Client ID", "Created At"];

var PW_SALT        = "ssc2016reunion";
var ADMIN_USERNAME = "admin";
var ADMIN_PASSWORD = "admin";

var VALID_GROUPS = ["Science", "Commerce", "Humanities", "Arts"];
var VALID_BLOOD  = ["A+", "A-", "B+", "B-", "AB+", "AB-", "O+", "O-"];

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
    Logger.log("Responses sheet created.");
  } else {
    Logger.log("Responses sheet already has data — skipped.");
  }

  var users = ss.getSheetByName(USERS_NAME) || ss.insertSheet(USERS_NAME);
  if (users.getLastRow() === 0) {
    users.appendRow(USER_HEADERS);
    users.getRange(1, 1, 1, USER_HEADERS.length).setFontWeight("bold");
    users.setFrozenRows(1);
    Logger.log("Users sheet created.");
  } else {
    Logger.log("Users sheet already has data — skipped.");
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

// Returns 0-based index of Status column, adding it if missing.
function ensureStatusCol_(sheet, header) {
  var idx = header.indexOf("Status");
  if (idx >= 0) return idx;
  var newIdx = header.length;
  sheet.getRange(1, newIdx + 1).setValue("Status");
  sheet.getRange(1, newIdx + 1).setFontWeight("bold");
  return newIdx;
}

function handleSignup_(d) {
  var username     = String(d.username     || "").trim().toLowerCase();
  var passwordHash = String(d.passwordHash || "");
  var clientId     = clean_(d.clientId     || "");

  if (username === ADMIN_USERNAME)
    return respond_({ ok: false, error: "Username not available." });
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

  // Admin hardcoded check
  if (username === ADMIN_USERNAME && passwordHash === hashPw_(ADMIN_PASSWORD)) {
    return respond_({ ok: true, username: ADMIN_USERNAME, clientId: "admin", isAdmin: true });
  }

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

function handleUpdateStatus_(d) {
  var adminHash = String(d.adminHash || "");
  if (!adminHash || adminHash !== hashPw_(ADMIN_PASSWORD)) {
    return respond_({ ok: false, error: "Unauthorized" });
  }

  var targetClientId = String(d.clientId || "");
  var status         = String(d.status   || "");
  if (["Approved", "Rejected", "Pending"].indexOf(status) < 0) {
    return respond_({ ok: false, error: "Invalid status." });
  }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return respond_({ ok: false, error: "Sheet not found" });

  var values    = sheet.getDataRange().getValues();
  var header    = values[0] || [];
  var cidCol    = header.indexOf("Client ID"); if (cidCol < 0) cidCol = 14;
  var statusCol = ensureStatusCol_(sheet, header);

  for (var i = 1; i < values.length; i++) {
    if (String(values[i][cidCol]) === targetClientId) {
      sheet.getRange(i + 1, statusCol + 1).setValue(status);
      return respond_({ ok: true });
    }
  }
  return respond_({ ok: false, error: "Entry not found" });
}

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);

    if (d.action === "signup")       return handleSignup_(d);
    if (d.action === "signin")       return handleSignin_(d);
    if (d.action === "updateStatus") return handleUpdateStatus_(d);

    // Honeypot
    if (d.website) return respond_({ ok: true });

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) { setup(); sheet = ss.getSheetByName(SHEET_NAME); }

    // Validate required fields
    var group      = clean_(d.group      || "");
    var bloodGroup = clean_(d.bloodGroup || "");
    if (VALID_GROUPS.indexOf(group) < 0)
      return respond_({ ok: false, error: "Please select a valid Group." });
    if (VALID_BLOOD.indexOf(bloodGroup) < 0)
      return respond_({ ok: false, error: "Please select a valid Blood Group." });

    var clientId = clean_(d.clientId || "");
    var now      = new Date();
    var values   = sheet.getDataRange().getValues();

    var header    = values[0] || [];
    var cidCol    = header.indexOf("Client ID"); if (cidCol < 0) cidCol = 14;
    var statusCol = ensureStatusCol_(sheet, header);

    // Re-read if status column was just added
    if (header.indexOf("Status") < 0) {
      values = sheet.getDataRange().getValues();
      header = values[0] || [];
    }

    var existingRow = -1;
    if (clientId) {
      for (var i = 1; i < values.length; i++) {
        if (String(values[i][cidCol]) === clientId) { existingRow = i + 1; break; }
      }
    }

    var photoCol = header.indexOf("Photo URL"); if (photoCol < 0) photoCol = 13;
    var photoUrl = existingRow > 0 ? String(values[existingRow - 1][photoCol] || "") : "";
    if (d.photo && String(d.photo).indexOf("https://") === 0) photoUrl = d.photo;

    // Preserve existing status; new submissions start as Pending
    var currentStatus = "Pending";
    if (existingRow > 0) {
      var prev = values[existingRow - 1][statusCol];
      currentStatus = prev !== undefined ? String(prev || "") : "";
    }

    var row = [
      existingRow > 0 ? values[existingRow - 1][0] : now,
      now,
      clean_(d.fullName),
      clean_(d.nickname),
      clean_(d.gender),
      group,
      bloodGroup,
      clean_(d.location),
      clean_(d.mobile),
      clean_(d.profession),
      clean_(d.jobPlace),
      clean_(d.university),
      clean_(d.maritalStatus),
      photoUrl,
      clientId,
      currentStatus
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

    var values    = sheet.getDataRange().getValues();
    var clientId  = (e && e.parameter && e.parameter.clientId)  ? String(e.parameter.clientId)  : "";
    var adminHash = (e && e.parameter && e.parameter.adminHash) ? String(e.parameter.adminHash) : "";
    var isAdmin   = adminHash && adminHash === hashPw_(ADMIN_PASSWORD);

    var header    = values[0] || [];
    var cidCol    = header.indexOf("Client ID"); if (cidCol < 0) cidCol = 14;
    var statusCol = header.indexOf("Status");

    // Return single user entry for form pre-fill
    if (clientId && !isAdmin) {
      for (var i = 1; i < values.length; i++) {
        var r = values[i];
        if (String(r[cidCol]) === clientId) {
          return respond_({ ok: true, entry: {
            fullName:      String(r[2]  || ""),
            nickname:      String(r[3]  || ""),
            gender:        String(r[4]  || ""),
            group:         String(r[5]  || ""),
            bloodGroup:    String(r[6]  || ""),
            location:      String(r[7]  || ""),
            mobile:        String(r[8]  || ""),
            profession:    String(r[9]  || ""),
            jobPlace:      String(r[10] || ""),
            university:    String(r[11] || ""),
            maritalStatus: String(r[12] || ""),
            photo:         String(r[13] || "")
          }});
        }
      }
      return respond_({ ok: true, entry: null });
    }

    // Return full list (admin or public)
    var entries = [];
    for (var j = 1; j < values.length; j++) {
      var row    = values[j];
      if (!row[2]) continue;

      var status = statusCol >= 0 ? String(row[statusCol] || "") : "";

      if (!isAdmin) {
        // Public: show entries with empty status (legacy/approved) only;
        // hide Pending (not yet reviewed) and Rejected
        if (status === "Pending" || status === "Rejected") continue;
      }

      var entry = {
        name:          String(row[2]),
        nickname:      String(row[3]  || ""),
        gender:        String(row[4]  || ""),
        group:         String(row[5]  || ""),
        bloodGroup:    String(row[6]  || ""),
        location:      String(row[7]  || ""),
        mobile:        String(row[8]  || ""),
        profession:    String(row[9]  || ""),
        jobPlace:      String(row[10] || ""),
        university:    String(row[11] || ""),
        maritalStatus: String(row[12] || ""),
        photo:         String(row[13] || "")
      };

      if (isAdmin) {
        entry.clientId    = String(row[cidCol] || "");
        entry.status      = status;
        entry.submittedAt = row[0] ? String(row[0]) : "";
        entry.updatedAt   = row[1] ? String(row[1]) : "";
      }

      entries.push(entry);
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
