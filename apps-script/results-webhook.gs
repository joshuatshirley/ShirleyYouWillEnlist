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
 *      Then Files > + > Script, name it "qrcode", and paste
 *      apps-script/qrcode.gs into it (draws the email's QR codes).
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

// Each per-test email also goes to the recruiter the applicant picked on
// the kiosk. Keys must match the kiosk's EST_CONFIG.recruiters names
// exactly ("Lastname RANK"). Leave an address "" until you have it; that
// recruiter's results then go to NOTIFY_EMAIL only. Use official
// (army.mil) addresses: these emails carry applicant names and phones.
// The emails have no link to the Sheet, so a recruiter only sees the
// results sent to them.
var RECRUITER_EMAILS = {
  "Curless SSG": "",
  "Gatling SFC": "",
  "Guiler SFC": "",
  "Recendez SFC": "",
  "Schultz SFC": "",
  "Shirley SFC": "joshua.t.shirley@army.mil"
  // NEW RECRUITER ("Aquaman" template): add a comma after the line above,
  // copy this line, remove the //, and set the same "Lastname RANK" you
  // used in index.html's EST_CONFIG.recruiters, plus their email.
  // , "Aquaman SGT": "first.m.last.mil@army.mil"
};

// Recruiter follow-up codes in each per-test email (followUpCodes):
// one saves the applicant as a contact, the other opens a text to them
// from the recruiter's phone. PRACTICE_SITE_URL is the public GitHub
// Pages address of the kiosk ("?practice" is added to it); leave it ""
// until the site is up. In the template, {recruiter} = "SFC Shirley",
// {score} = projected AFQT, {link} = the practice link, {weakest} = their
// lowest section (Arithmetic Reasoning when it's in a tie).
var PRACTICE_SITE_URL = "";
var APPLICANT_TEXT_TEMPLATE = "Hey this is {recruiter}. You got a {score}. Check this link to practice " +
  "on your own {link} You needed help with {weakest} the most. Focus on that in March2Success.com.";

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
  "Test Key",
  // End-of-test readiness questions (kiosk renderReadiness). Older kiosks
  // and Florence_Practice_Test don't send them, so these stay blank.
  "Ready (1-10)", "Not-10 Reasons", "What Would Make It 10"
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
    data.test_key || "",
    data.ready === undefined ? "" : data.ready,
    sheetText(data.not_ten_reasons),
    sheetText(data.make_it_ten)
  ]);

  notifyResult(data);
}

// Applicant-typed text: a leading =, +, - or @ would make Sheets read
// the cell as a formula (and show #ERROR!), so prefix those with ' to
// store them as plain text.
function sheetText(value) {
  var text = value == null ? "" : String(value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function hasReady(data) {
  return data.ready !== undefined && data.ready !== null && data.ready !== "";
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
// NOTIFY_EMAIL plus the applicant's recruiter, without sending anyone
// two copies.
function notifyRecipients(recruiter) {
  var list = [];
  [NOTIFY_EMAIL, RECRUITER_EMAILS[recruiter] || ""].forEach(function (addr) {
    addr = String(addr).trim();
    if (addr && list.indexOf(addr.toLowerCase()) === -1) list.push(addr.toLowerCase());
  });
  return list.join(",");
}

function notifyResult(data) {
  var to = notifyRecipients(data.recruiter || "");
  if (!to) return;
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
      " Test " + testNumber + " AFQT " + afqt + " " + phone +
      (hasReady(data) ? " Ready " + data.ready + "/10" : "");

    var plainLines = [when, lastName + ", " + firstName, "Test " + testNumber, "AFQT: " + afqt, phone, recruiter];
    if (hasReady(data)) {
      plainLines.push("Ready to join: " + data.ready + "/10");
      if (data.not_ten_reasons) plainLines.push("Why not a 10: " + data.not_ten_reasons);
      if (data.make_it_ten) plainLines.push("What would make it a 10: " + data.make_it_ten);
    }
    SECTION_ORDER.forEach(function (code) {
      var s = sections[code];
      if (!s) return;
      plainLines.push(SECTION_NAMES[code] + ": " + s.correct + "/" + s.total + " (" + statusBand(s).label + ")");
    });
    var codes = followUpCodes(data, testNumber);
    if (codes.text) plainLines.push("", "Text to send: " + codes.text);
    var body = plainLines.join("\n");

    var emailOptions = {
      htmlBody: buildHtmlReport(data, when, testNumber) + codes.html,
      inlineImages: codes.images,
      name: "Florence EST Kiosk"
    };
    if (codes.vcardBlob) {
      emailOptions.attachments = [codes.vcardBlob];
    }

    MailApp.sendEmail(to, subject, body, emailOptions);
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

    "<p style=\"margin:0 0 4px;font-size:13px;color:#555;\">" + esc(when) + " &middot; Test " + esc(testNumber) + "</p>" +
    "<h2 style=\"margin:0 0 4px;font-size:22px;\">" + lastName + ", " + firstName + "</h2>" +
    "<p style=\"margin:0 0 20px;font-size:14px;color:#555;\">" + recruiter + (phone ? " &middot; " + phone : "") + "</p>" +

    "<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" role=\"presentation\" style=\"border-bottom:1px solid #eee;margin-bottom:20px;\"><tr>" +
    "<td align=\"center\" style=\"padding:6px 0 22px;\">" +
    "<div style=\"font-size:44px;font-weight:700;color:#1e4d8c;line-height:1;\">" + data.afqt + "</div>" +
    "<div style=\"font-size:12px;color:#555;text-transform:uppercase;letter-spacing:1px;margin-top:8px;\">Projected AFQT Score</div>" +
    "<div style=\"font-size:11px;color:#888;margin-top:2px;\">Projected on the full-length ASVAB scale</div>" +
    "</td></tr></table>" +

    "<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" role=\"presentation\">" + meterRows + "</table>" +
    readinessHtml(data) +

    "<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" role=\"presentation\" style=\"border-top:1px solid #eee;margin-top:4px;\"><tr>" +
    "<td align=\"center\" style=\"padding:16px 0 4px;\">" +
    "<div style=\"font-size:26px;font-weight:700;color:#1e4d8c;\">" + overallPct + "%</div>" +
    "<div style=\"font-size:11px;color:#555;text-transform:uppercase;letter-spacing:1px;margin-top:2px;\">Overall</div>" +
    "</td></tr></table>" +


    "</td></tr></table>";
}

// The applicant's answers to the end-of-test readiness questions, if the
// kiosk sent them. All of it is applicant-typed, so every value is escaped.
function readinessHtml(data) {
  if (!hasReady(data)) return "";
  var html = "<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" role=\"presentation\" style=\"margin:8px 0 4px;background:#f1e4c7;\"><tr>" +
    "<td style=\"padding:14px 16px;font-size:14px;\">" +
    "<div style=\"font-size:12px;color:#555;text-transform:uppercase;\">Ready to join the Army</div>" +
    "<div style=\"font-size:26px;font-weight:700;margin:2px 0 6px;\">" + esc(data.ready) + " / 10</div>";
  if (data.not_ten_reasons) html += "<p style=\"margin:4px 0;\"><b>Why not a 10:</b> " + esc(data.not_ten_reasons) + "</p>";
  if (data.make_it_ten) html += "<p style=\"margin:4px 0;\"><b>What would make it a 10:</b> " + esc(data.make_it_ten) + "</p>";
  return html + "</td></tr></table>";
}

// ---------- Recruiter follow-up codes (email only) ----------
// "Lastname RANK" -> "RANK Lastname", like the kiosk's
// formatRecruiterDisplay().
function recruiterDisplay(value) {
  var parts = String(value || "").trim().split(/\s+/);
  if (parts.length < 2) return parts.join(" ");
  var rank = parts[parts.length - 1];
  if (rank !== rank.toUpperCase() || !/[A-Z]/.test(rank)) return parts.join(" ");
  return rank + " " + parts.slice(0, -1).join(" ");
}

// 10 digits or "" (named apart from weekly-report.gs's phoneDigits, which
// can share this project).
function tenDigitPhone(phone) {
  var digits = String(phone || "").replace(/\D/g, "");
  if (digits.length === 11 && digits.charAt(0) === "1") digits = digits.slice(1);
  return digits.length === 10 ? digits : "";
}

// Lowest percent-correct section; on a tie, Arithmetic Reasoning if it's
// one of them, else the first in test order. Same pick as the kiosk's
// weakestSectionCode(); keep the two in sync.
function weakestSectionForText(sections) {
  var worst = null, tied = [];
  SECTION_ORDER.forEach(function (code) {
    var s = sections[code];
    if (!s || !s.total) return;
    var pct = s.correct / s.total;
    if (worst === null || pct < worst) { worst = pct; tied = [code]; }
    else if (pct === worst) tied.push(code);
  });
  var code = tied.indexOf("AR") !== -1 ? "AR" : (tied[0] || "AR");
  return SECTION_NAMES[code];
}

function applicantTextMessage(data) {
  var link = PRACTICE_SITE_URL
    ? PRACTICE_SITE_URL.split("?")[0].split("#")[0] + "?practice."
    : "(practice link not set up yet).";
  return APPLICANT_TEXT_TEMPLATE
    .replace("{recruiter}", recruiterDisplay(data.recruiter))
    .replace("{score}", String(data.afqt))
    .replace("{link}", link)
    .replace("{weakest}", weakestSectionForText(data.sections || {}));
}

function vcardValue(value) {
  return String(value || "").replace(/\\/g, "\\\\").replace(/([,;])/g, "\\$1").replace(/\r?\n/g, " ");
}

function applicantVcard(data, testNumber) {
  var first = data.first_name || "", last = data.last_name || "";
  var digits = tenDigitPhone(data.phone);
  var note = "Practice ASVAB Test " + testNumber + ", AFQT " + data.afqt +
    (hasReady(data) ? ", ready " + data.ready + "/10" : "") +
    ", recruiter " + recruiterDisplay(data.recruiter);
  var lines = ["BEGIN:VCARD", "VERSION:3.0",
    "N:" + vcardValue(last) + ";" + vcardValue(first) + ";;;",
    "FN:" + vcardValue((first + " " + last).trim())];
  if (digits) lines.push("TEL;TYPE=CELL:+1" + digits);
  lines.push("NOTE:" + vcardValue(note), "END:VCARD");
  return lines.join("\r\n");
}

// qrcode.gs draws the code as a GIF, all in Apps Script.
function qrImage(text, name) {
  var qr = qrcode(0, "M");
  qrcode.stringToBytes = qrcode.stringToBytesFuncs["UTF-8"];
  qr.addData(text, "Byte");
  qr.make();
  var dataUrl = qr.createDataURL(4, 4);
  return Utilities.newBlob(Utilities.base64Decode(dataUrl.split(",")[1]), "image/gif", name + ".gif");
}

// The codes for this result as inline email images, plus the text
// message on its own for the plain-text body. Empty (and logged) if
// qrcode.gs wasn't added to the project.
function followUpCodes(data, testNumber) {
  var digits = tenDigitPhone(data.phone);
  var text = digits ? applicantTextMessage(data) : "";
  if (typeof qrcode === "undefined") {
    Logger.log("qrcode.gs is missing from this project; emails go out without QR codes.");
    return { html: "", images: {}, text: text, vcardBlob: null };
  }
  var vcard = applicantVcard(data, testNumber);
  var images = { qrSave: qrImage(vcard, "save-applicant") };
  if (text) images.qrText = qrImage("SMSTO:+1" + digits + ":" + text, "text-applicant");

  var vcardBlob = digits
    ? Utilities.newBlob(vcard, "text/vcard", (((data.first_name || "") + "_" + (data.last_name || "Applicant")).trim() || "applicant").replace(/[^a-zA-Z0-9_-]/g, "_") + ".vcf")
    : null;

  var saveHref = digits ? ("tel:+1" + digits) : "";
  var textHref = (digits && text) ? ("sms:+1" + digits + "?body=" + encodeURIComponent(text)) : "";

  function cell(cid, label, sub, href, actionLabel) {
    var imgTag = "<img src=\"cid:" + cid + "\" width=\"180\" height=\"180\" alt=\"" + esc(label) + " QR code\" style=\"display:block;border:0;margin:0 auto;\">";
    var imgHtml = href
      ? "<a href=\"" + esc(href) + "\" style=\"text-decoration:none;display:block;\">" + imgTag + "</a>"
      : imgTag;
    var titleHtml = href
      ? "<a href=\"" + esc(href) + "\" style=\"color:#0b57d0;text-decoration:none;font-weight:700;\">" + esc(label) + "</a>"
      : esc(label);
    var btnHtml = "";
    if (href && actionLabel) {
      btnHtml = "<div style=\"margin-top:8px;\">" +
        "<a href=\"" + esc(href) + "\" style=\"display:inline-block;padding:6px 14px;background:#221f20;color:#ffcc01;text-decoration:none;font-size:13px;font-weight:700;border-radius:4px;\">" +
        esc(actionLabel) + "</a></div>";
    }
    return "<td align=\"center\" valign=\"top\" style=\"padding:8px;width:50%;\">" +
      imgHtml +
      "<div style=\"font-size:14px;font-weight:700;margin-top:6px;\">" + titleHtml + "</div>" +
      "<div style=\"font-size:12px;color:#555;\">" + esc(sub) + "</div>" +
      btnHtml +
      "</td>";
  }

  var html = "<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" role=\"presentation\" style=\"max-width:480px;margin-top:20px;border-top:1px solid #eee;font-family:Arial,Helvetica,sans-serif;color:#1a1a1a;\">" +
    "<tr><td colspan=\"2\" style=\"padding:14px 4px 0;font-size:16px;font-weight:700;\">Recruiter follow-up</td></tr>" +
    "<tr><td colspan=\"2\" style=\"padding:2px 4px 6px;font-size:13px;color:#555;\">Scan with your phone camera, or tap if opening on your phone.</td></tr><tr>" +
    cell("qrSave", "Save applicant", digits ? (formatPhone(digits) + (vcardBlob ? " · vCard attached" : "")) : "No phone number given", saveHref, "Save / Call") +
    (text ? cell("qrText", "Text the applicant", "Opens a text with their score and practice link", textHref, "Send text")
          : "<td style=\"width:50%;\"></td>") +
    "</tr></table>";
  return { html: html, images: images, text: text, vcardBlob: vcardBlob };
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
