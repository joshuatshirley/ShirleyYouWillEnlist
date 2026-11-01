/**
 * Backend for the Florence EST kiosk's results webhook.
 *
 * Source of truth for the Apps Script deployed at
 * https://script.google.com/ (Extensions > Apps Script, from a Google
 * Sheet). This file is not run from here -- it's kept in the repo so
 * the deployed script has a version-controlled copy to diff against
 * and redeploy from.
 *
 * Setup (one-time):
 *   1. Create a new Google Sheet (e.g. "Florence EST Results").
 *   2. Extensions > Apps Script. Delete the boilerplate, paste this file.
 *   3. Deploy > New deployment > type "Web app".
 *        Execute as: Me
 *        Who has access: Anyone
 *   4. Authorize when prompted, then copy the resulting web app URL
 *      (ends in /exec) into resultsWebhookUrl in index.html's
 *      window.EST_CONFIG block.
 *
 * The kiosk POSTs JSON as Content-Type: text/plain (not
 * application/json) specifically to avoid a CORS preflight -- Apps
 * Script's web app endpoint issues a redirect for POST requests that
 * a preflighted request can choke on. This script parses the body as
 * JSON regardless of the declared content type.
 *
 * Redeploying after an edit: Deploy > Manage deployments > pick the
 * existing deployment > Edit (pencil) > New version > Deploy. This
 * keeps the same /exec URL, so the kiosk doesn't need updating.
 *
 * Note on HEADER_ROW: the full row is written when the sheet is empty,
 * and getOrCreateSheet fills in any header cells missing from the end of
 * an existing sheet's row 1. That's why new columns are only ever
 * appended at the END of HEADER_ROW and appendResultRow -- never
 * inserted, which would shift existing columns.
 *
 * Optional "Roster" tab (each recruiter's tested applicants + phone
 * numbers, grouped by recruiter then last name): add a new sheet tab
 * named "Roster" and paste this into cell A1. It's a live formula, not
 * script output -- no redeploy needed if it ever needs re-adding.
 *
 *   =QUERY(Results!A2:W, "select C, V, W, U, max(A), count(A) group by
 *   C, V, W, U order by C, V, W label C 'Recruiter', V 'Last Name',
 *   W 'First Name', U 'Phone', max(A) 'Last Tested', count(A)
 *   'Times Tested'", 0)
 *
 * Column letters (C=Recruiter, V=Last Name, W=First Name, U=Phone,
 * A=Timestamp) assume HEADER_ROW's current column order below -- if
 * that order changes, update the letters in the formula to match.
 *
 * Notification on each save: MailApp.sendEmail to NOTIFY_EMAIL, as an
 * HTML report (with a plain-text fallback body for clients that don't
 * render HTML) -- a per-section breakdown table, not just the AFQT
 * number, so a section that went poorly is visible without opening
 * the Sheet, and the email itself is presentable enough to forward
 * straight to a recruiter. (This started as a T-Mobile email-to-SMS
 * gateway address -- carriers routinely filter or kill those, and
 * this one stopped delivering, so it now goes to a real inbox
 * instead.) Adding this to an already-deployed script requires a
 * fresh authorization the first time (it now sends email, a new
 * scope) -- expect a consent prompt on the next "New version" deploy.
 * Set NOTIFY_EMAIL to "" to disable.
 */

var SHEET_NAME = "Results";
var NOTIFY_EMAIL = "joshua.t.shirley@army.mil";
var SHEET_URL = "https://docs.google.com/spreadsheets/d/1BftqXMNaw7WurUHmzGtQiUQvFmgnaGe-1Qm_CtW1Tjc/edit";

var SECTION_NAMES = {
  WK: "Word Knowledge",
  AR: "Arithmetic Reasoning",
  PC: "Paragraph Comprehension",
  MK: "Mathematics Knowledge"
};
// Section order + thresholds match the kiosk's own statusBand() in
// index.html -- keep these in sync if that ever changes.
var SECTION_ORDER = ["WK", "AR", "PC", "MK"];

var HEADER_ROW = [
  "Timestamp", "Name", "Recruiter", "AFQT", "VE Score", "AR Score", "MK Score", "Std Sum",
  "WK Correct", "WK Total", "WK Unanswered",
  "AR Correct", "AR Total", "AR Unanswered",
  "PC Correct", "PC Total", "PC Unanswered",
  "MK Correct", "MK Total", "MK Unanswered",
  "Phone",
  "Last Name", "First Name",
  "Test Number",
  "Next Action", "Notes",
  "Test Key"
];

// Practice drills (the take-home practice page) log here, NOT to the
// Results tab -- a drill is not a graded EST, and keeping it off Results
// keeps the Roster QUERY and dashboard limited to real tests. Write-only
// like everything else on this public endpoint: nothing is ever read
// back to the device.
var PRACTICE_SHEET_NAME = "Practice";
var PRACTICE_HEADER_ROW = [
  "Timestamp", "Phone", "Last Name", "First Name", "Recruiter",
  "Section", "Correct", "Total"
];

// Two payload shapes hit this same endpoint: a full result (appendResultRow,
// the default -- old cached kiosk versions never send `type` at all, so the
// default has to stay "append a result row", not "do nothing") and a
// recruiter's next-action note for an already-saved result
// (updateResultNote, data.type === "note"). A third, data.type ===
// "practice", is a take-home practice drill (appendPracticeRow).
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.type === "note") {
      updateResultNote(data);
    } else if (data.type === "practice") {
      appendPracticeRow(data);
    } else {
      appendResultRow(data);
    }
    return jsonResponse({ ok: true });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

// Finds the result row by its original timestamp (effectively unique --
// ISO-8601 with milliseconds) and writes the note into it in place, so a
// note stays attached to its result for the Roster Dashboard and any
// filtering, rather than living in a disconnected list. Throws (caught by
// doPost, reported back to the kiosk as ok:false) if no matching row is
// found, e.g. the result hasn't finished syncing yet.
function updateResultNote(data) {
  var sheet = getOrCreateSheet();
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var tsCol = headers.indexOf("Timestamp");
  var nextActionCol = headers.indexOf("Next Action");
  var notesCol = headers.indexOf("Notes");

  for (var i = 1; i < values.length; i++) {
    if (values[i][tsCol] === data.result_timestamp) {
      var row = i + 1; // 1-based sheet row; values[0] is the header row
      if (nextActionCol !== -1) sheet.getRange(row, nextActionCol + 1).setValue(data.next_action || "");
      if (notesCol !== -1) sheet.getRange(row, notesCol + 1).setValue(data.note_text || "");
      return;
    }
  }
  throw new Error("No result row found for timestamp " + data.result_timestamp);
}

function appendResultRow(data) {
  var sheet = getOrCreateSheet();
  var sections = data.sections || {};

  function field(code, key) {
    var s = sections[code];
    return s && s[key] !== undefined ? s[key] : "";
  }

  // The kiosk sends last_name/first_name separately (so this sheet can sort
  // and group by last name); the "Name" column is still populated here, as
  // a combined "First Last" display string, for continuity with rows saved
  // before that split.
  var displayName = ((data.first_name || "") + " " + (data.last_name || "")).trim();

  sheet.appendRow([
    data.timestamp || new Date().toISOString(),
    displayName,
    data.recruiter || "",
    data.afqt,
    data.ve_score,
    data.ar_score,
    data.mk_score,
    data.std_sum,
    field("WK", "correct"), field("WK", "total"), field("WK", "unanswered"),
    field("AR", "correct"), field("AR", "total"), field("AR", "unanswered"),
    field("PC", "correct"), field("PC", "total"), field("PC", "unanswered"),
    field("MK", "correct"), field("MK", "total"), field("MK", "unanswered"),
    data.phone || "",
    data.last_name || "",
    data.first_name || "",
    data.test_number || 1,
    // Next Action / Notes are filled in later by updateResultNote; left
    // blank here so Test Key lands in its own column after them.
    "", "",
    data.test_key || ""
  ]);

  notifyResult(data);
}

// No per-drill email -- drills are frequent and low-stakes, so they only
// show up in the Friday weekly report (weekly-report.gs).
function appendPracticeRow(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PRACTICE_SHEET_NAME) || ss.insertSheet(PRACTICE_SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(PRACTICE_HEADER_ROW);
  sheet.appendRow([
    data.timestamp || new Date().toISOString(),
    data.phone || "",
    data.last_name || "",
    data.first_name || "",
    data.recruiter || "",
    data.section || "",
    data.correct,
    data.total
  ]);
}

// Best-effort: the row is already saved by the time this runs, so a
// notify failure (e.g. MailApp's daily quota) must not surface as a
// save failure back to the kiosk.
function notifyResult(data) {
  if (!NOTIFY_EMAIL) return;
  try {
    var when = Utilities.formatDate(
      data.timestamp ? new Date(data.timestamp) : new Date(),
      SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),
      "yyyy MMM dd HH:mm"
    ).toUpperCase();
    var lastName = data.last_name || "";
    var firstName = data.first_name || "";
    var recruiter = data.recruiter || "";
    var afqt = data.afqt;
    var phone = formatPhone(data.phone);
    var testNumber = data.test_number || 1;
    var sections = data.sections || {};

    var subject = "EST " + recruiter + " " + lastName + " " + firstName +
      " AFQT " + afqt + " " + phone;

    var plainLines = [when, lastName + ", " + firstName, "AFQT: " + afqt, phone, recruiter];
    if (testNumber > 1) plainLines.push("Test " + testNumber);
    SECTION_ORDER.forEach(function (code) {
      var s = sections[code];
      if (!s) return;
      plainLines.push(SECTION_NAMES[code] + ": " + s.correct + "/" + s.total + " (" + statusBand(s).label + ")");
    });
    plainLines.push(SHEET_URL);
    var body = plainLines.join("\n");

    MailApp.sendEmail(NOTIFY_EMAIL, subject, body, {
      htmlBody: buildHtmlReport(data, when, testNumber),
      name: "Florence EST Kiosk"
    });
  } catch (err) {
    Logger.log("notifyResult failed: " + err);
  }
}

// Same 70%/40% thresholds as the kiosk's own statusBand() in
// index.html, so this email reads the same as what the
// recruiter saw live on the device.
function statusBand(section) {
  var pct = section.total ? (100 * section.correct) / section.total : 0;
  if (pct >= 70) return { color: "#0ca30c", label: "Strong" };
  if (pct >= 40) return { color: "#fab219", label: "Needs review" };
  return { color: "#d03b3b", label: "Keep practicing" };
}

// Mirrors the kiosk's own Results screen (hero AFQT, a color-banded
// progress meter per section, overall %) as closely as HTML email
// permits. Built with nested tables and bgcolor attributes rather than
// CSS gradients/border-radius/flexbox specifically because a .mil
// inbox is very likely Outlook, whose HTML renderer (Word's engine)
// ignores most modern CSS -- the two-cell "table bar" is the standard
// cross-client-safe way to fake a progress bar in email.
function buildHtmlReport(data, when, testNumber) {
  var lastName = esc(data.last_name || "");
  var firstName = esc(data.first_name || "");
  var recruiter = esc(data.recruiter || "");
  var phone = esc(formatPhone(data.phone));
  var sections = data.sections || {};

  var overallCorrect = 0, overallTotal = 0;
  SECTION_ORDER.forEach(function (code) {
    var s = sections[code];
    if (!s) return;
    overallCorrect += s.correct;
    overallTotal += s.total;
  });
  var overallPct = overallTotal ? Math.round((100 * overallCorrect) / overallTotal) : 0;

  var meterRows = SECTION_ORDER.map(function (code) {
    var s = sections[code];
    if (!s) return "";
    var band = statusBand(s);
    var pct = s.total ? Math.round((100 * s.correct) / s.total) : 0;
    var unansweredNote = s.unanswered > 0
      ? "<span style=\"color:#555;\"> &middot; " + s.unanswered + " unanswered</span>"
      : "";

    return "" +
      "<tr><td style=\"padding:0 0 18px;\">" +
      "<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" role=\"presentation\"><tr>" +
      "<td style=\"font-size:15px;color:#1a1a1a;\">" + esc(SECTION_NAMES[code]) + "</td>" +
      "<td align=\"right\" style=\"font-size:14px;color:#555;font-family:Consolas,'SF Mono',monospace;white-space:nowrap;\">" + s.correct + " / " + s.total + "</td>" +
      "</tr></table>" +
      "<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" role=\"presentation\" style=\"margin-top:6px;\"><tr>" +
      "<td width=\"" + pct + "%\" bgcolor=\"" + band.color + "\" style=\"font-size:1px;line-height:10px;\">&nbsp;</td>" +
      "<td width=\"" + (100 - pct) + "%\" bgcolor=\"#e6e6e6\" style=\"font-size:1px;line-height:10px;\">&nbsp;</td>" +
      "</tr></table>" +
      "<p style=\"margin:6px 0 0;font-size:13px;color:" + band.color + ";font-weight:700;\">" + band.label + unansweredNote + "</p>" +
      "</td></tr>";
  }).join("");

  return "" +
    "<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" role=\"presentation\" style=\"max-width:480px;background:#ffffff;font-family:Arial,Helvetica,sans-serif;color:#1a1a1a;\"><tr><td style=\"padding:8px 4px;\">" +

    "<p style=\"margin:0 0 4px;font-size:13px;color:#555;\">" + esc(when) + (testNumber > 1 ? " &middot; Test " + testNumber : "") + "</p>" +
    "<h2 style=\"margin:0 0 4px;font-size:22px;\">" + lastName + ", " + firstName + "</h2>" +
    "<p style=\"margin:0 0 20px;font-size:14px;color:#555;\">" + recruiter + (phone ? " &middot; " + phone : "") + "</p>" +

    "<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" role=\"presentation\" style=\"border-bottom:1px solid #eee;margin-bottom:20px;\"><tr>" +
    "<td align=\"center\" style=\"padding:6px 0 22px;\">" +
    "<div style=\"font-size:44px;font-weight:700;color:#1e4d8c;line-height:1;\">" + data.afqt + "</div>" +
    "<div style=\"font-size:12px;color:#555;text-transform:uppercase;letter-spacing:1px;margin-top:8px;\">Projected AFQT Score</div>" +
    "<div style=\"font-size:11px;color:#888;margin-top:2px;\">Projected on the full-length ASVAB scale</div>" +
    "</td></tr></table>" +

    "<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" role=\"presentation\">" + meterRows + "</table>" +

    "<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" role=\"presentation\" style=\"border-top:1px solid #eee;margin-top:4px;\"><tr>" +
    "<td align=\"center\" style=\"padding:16px 0 4px;\">" +
    "<div style=\"font-size:26px;font-weight:700;color:#1e4d8c;\">" + overallPct + "%</div>" +
    "<div style=\"font-size:11px;color:#555;text-transform:uppercase;letter-spacing:1px;margin-top:2px;\">Overall</div>" +
    "</td></tr></table>" +

    "<p style=\"margin:22px 0 0;font-size:14px;\"><a href=\"" + SHEET_URL + "\" style=\"color:#1e4d8c;\">View full results sheet &#8599;</a></p>" +

    "</td></tr></table>";
}

function esc(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// Formats a 10-digit US number as "(XXX) XXX-XXXX"; a leading country
// code "1" on an 11-digit number is dropped first. Anything that isn't
// a standard 10-digit number (blank, partial, international) is
// returned as typed rather than guessed at.
function formatPhone(phone) {
  if (!phone) return "";
  var digits = String(phone).replace(/\D/g, "");
  if (digits.length === 11 && digits.charAt(0) === "1") digits = digits.slice(1);
  if (digits.length === 10) {
    return "(" + digits.slice(0, 3) + ") " + digits.slice(3, 6) + "-" + digits.slice(6);
  }
  return phone;
}

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADER_ROW);
  } else if (sheet.getLastColumn() < HEADER_ROW.length) {
    // An existing sheet predates the newest columns (see the HEADER_ROW
    // note at the top) -- fill in just the missing header cells.
    var have = sheet.getLastColumn();
    sheet.getRange(1, have + 1, 1, HEADER_ROW.length - have)
      .setValues([HEADER_ROW.slice(have)]);
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
