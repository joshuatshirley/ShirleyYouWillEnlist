/**
 * Friday-morning weekly EST report, sent by email.
 *
 * Add it as a second file (Files > + > Script, name it "weekly-report")
 * in EITHER Apps Script project, the Sheet-bound results-webhook one or
 * the private roster-dashboard one. Both define the SHEET_NAME, esc()
 * and formatPhone() this file uses. A time-driven trigger runs
 * server-side and only sends an email: it adds no web endpoint, doesn't
 * change either deployment's access level, and nothing about applicants
 * is ever served back to a device.
 *
 * Setup (one-time):
 *   1. Project Settings (gear) > Time zone: (GMT-05:00) Eastern Time.
 *      The trigger hour and every date in the email use this zone.
 *   2. Set REPORT_EMAIL below.
 *   3. Pick sendWeeklyReportNow in the function dropdown > Run. That
 *      triggers the consent prompt (this file sends email, a new scope
 *      for this project) and sends a test copy covering the last 7 days.
 *   4. Pick installWeeklyTrigger > Run. Safe to re-run: it replaces any
 *      earlier copy of the trigger rather than stacking a second one.
 * No redeploy is needed -- triggers don't go through the web app
 * deployment at all.
 *
 * The report covers the 7 days before it runs. Per station: everyone who
 * took a graded test that week, every test they took (date, test number,
 * projected AFQT), their latest projected AFQT, the change since their
 * first test that week, and their weakest section. Then practice drills
 * from the "Practice" tab, if that tab exists yet.
 *
 * People are matched by phone digits, so "(843) 555-0100" and
 * "8435550100" are the same person; a row with no phone falls back to
 * station + name.
 */

// Its own copy of the Sheet ID (not roster-dashboard.gs's SHEET_ID), so
// this file works pasted into either Apps Script project.
var REPORT_SHEET_ID = "1ERX4wPOGAHg5NQ7Sci36z_iWaZtGqWAe-MbkbQbcguo";
var REPORT_EMAIL = "joshua.t.shirley@army.mil";
var REPORT_DAYS = 7;
var PRACTICE_SHEET_NAME = "Practice";

var REPORT_SECTIONS = {
  WK: "Word Knowledge",
  AR: "Arithmetic Reasoning",
  PC: "Paragraph Comprehension",
  MK: "Mathematics Knowledge"
};
var REPORT_SECTION_ORDER = ["WK", "AR", "PC", "MK"];

function installWeeklyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "sendWeeklyReport") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("sendWeeklyReport")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(6)
    .create();
}

function sendWeeklyReportNow() {
  sendWeeklyReport();
}

function sendWeeklyReport() {
  var ss = SpreadsheetApp.openById(REPORT_SHEET_ID);
  var tz = Session.getScriptTimeZone();
  var end = new Date();
  var start = new Date(end.getTime() - REPORT_DAYS * 24 * 60 * 60 * 1000);

  var tests = readRows(ss.getSheetByName(SHEET_NAME)).filter(function (r) {
    var t = parseTimestamp(r["Timestamp"]);
    return t && t >= start && t <= end;
  });
  var practiceSheet = ss.getSheetByName(PRACTICE_SHEET_NAME);
  var drills = readRows(practiceSheet).filter(function (r) {
    var t = parseTimestamp(r["Timestamp"]);
    return t && t >= start && t <= end;
  });

  var people = groupTestsByPerson(tests);
  var range = fmtDate(start, tz) + " - " + fmtDate(end, tz);
  var subject = "Weekly EST Report " + range + ": " +
    tests.length + (tests.length === 1 ? " test, " : " tests, ") +
    people.length + (people.length === 1 ? " applicant" : " applicants");

  MailApp.sendEmail(REPORT_EMAIL, subject, buildPlainReport(people, drills, range, tz), {
    htmlBody: buildHtmlWeeklyReport(people, drills, range, tz, !!practiceSheet),
    name: "Florence EST Kiosk"
  });
}

// Each row as a {header: value} object, so column order in the Sheet
// can change without breaking this file. A missing tab reads as no rows.
function readRows(sheet) {
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  return values.slice(1).map(function (row) {
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  });
}

// The webhook writes ISO-8601 strings, but Sheets may have converted
// some of them to Date cells -- accept either.
function parseTimestamp(value) {
  if (value instanceof Date) return value;
  if (!value) return null;
  var d = new Date(value);
  return isNaN(d.getTime()) ? null : d;
}

// Rows without a Test Key were saved before the 2026-11 renumbering,
// when the original set was Test 1 and the bank sets were 2-5. It's now
// bank sets 1-4 and the original as Test 5. Mirrors currentTestNumber()
// in index.html; keep the two in sync.
function currentTestNumber(row) {
  var n = Number(row["Test Number"]) || 1;
  if (row["Test Key"]) return n;
  return n === 1 ? 5 : n - 1;
}

function phoneDigits(phone) {
  var digits = String(phone || "").replace(/\D/g, "");
  if (digits.length === 11 && digits.charAt(0) === "1") digits = digits.slice(1);
  return digits;
}

function personKey(row) {
  var digits = phoneDigits(row["Phone"]);
  if (digits) return "p:" + digits;
  return ["n", row["Recruiter"], row["Last Name"], row["First Name"]].join("|").toLowerCase();
}

function groupTestsByPerson(rows) {
  var byKey = {};
  rows.forEach(function (row) {
    var key = personKey(row);
    var p = byKey[key];
    if (!p) {
      p = byKey[key] = { tests: [], phone: "" };
    }
    p.tests.push(row);
    if (!p.phone && row["Phone"]) p.phone = row["Phone"];
  });

  return Object.keys(byKey).map(function (key) {
    var p = byKey[key];
    p.tests.sort(function (a, b) { return parseTimestamp(a["Timestamp"]) - parseTimestamp(b["Timestamp"]); });
    var latest = p.tests[p.tests.length - 1];
    return {
      lastName: latest["Last Name"] || "",
      firstName: latest["First Name"] || "",
      recruiter: latest["Recruiter"] || "(no station)",
      phone: formatPhone(p.phone),
      tests: p.tests,
      latestAfqt: Number(latest["AFQT"]),
      change: p.tests.length > 1 ? Number(latest["AFQT"]) - Number(p.tests[0]["AFQT"]) : null,
      weakest: weakestSection(latest),
      nextAction: latest["Next Action"] || "",
      // Readiness answers from their latest test (blank on older rows).
      ready: latest["Ready (1-10)"] === undefined ? "" : latest["Ready (1-10)"],
      notTenReasons: latest["Not-10 Reasons"] || "",
      makeItTen: latest["What Would Make It 10"] || ""
    };
  }).sort(function (a, b) {
    return a.recruiter.localeCompare(b.recruiter) || a.lastName.localeCompare(b.lastName);
  });
}

// Lowest percent-correct section on the given test; ties go to the
// earlier section in test order.
function weakestSection(row) {
  var worst = null;
  REPORT_SECTION_ORDER.forEach(function (code) {
    var total = Number(row[code + " Total"]);
    if (!total) return;
    var pct = Math.round((100 * Number(row[code + " Correct"])) / total);
    if (!worst || pct < worst.pct) worst = { code: code, pct: pct };
  });
  return worst;
}

// Drills grouped by phone (falling back to name), with a per-section
// count and average score so you can see what each person worked on.
function summarizeDrills(drills) {
  var byKey = {};
  drills.forEach(function (row) {
    var key = personKey(row);
    var p = byKey[key] || (byKey[key] = {
      lastName: row["Last Name"] || "", firstName: row["First Name"] || "",
      recruiter: row["Recruiter"] || "", phone: formatPhone(row["Phone"]),
      count: 0, sections: {}
    });
    p.count++;
    var code = row["Section"] || "?";
    var s = p.sections[code] || (p.sections[code] = { count: 0, correct: 0, total: 0 });
    s.count++;
    s.correct += Number(row["Correct"]) || 0;
    s.total += Number(row["Total"]) || 0;
  });
  return Object.keys(byKey).map(function (k) { return byKey[k]; }).sort(function (a, b) {
    return b.count - a.count || a.lastName.localeCompare(b.lastName);
  });
}

function fmtDate(d, tz) {
  return Utilities.formatDate(d, tz, "dd MMM yyyy").toUpperCase();
}

function fmtDateTime(d, tz) {
  return Utilities.formatDate(d, tz, "EEE dd MMM HHmm").toUpperCase();
}

function fmtChange(change) {
  if (change === null || isNaN(change)) return "";
  return (change > 0 ? "+" : "") + change;
}

function displayName(p) {
  return p.lastName + (p.firstName ? ", " + p.firstName : "");
}

function buildPlainReport(people, drills, range, tz) {
  var lines = ["Weekly EST Report", range, ""];
  if (!people.length) lines.push("No graded tests this week.");
  var station = null;
  people.forEach(function (p) {
    if (p.recruiter !== station) {
      station = p.recruiter;
      lines.push("", "== " + station + " ==");
    }
    lines.push(displayName(p) + "  " + p.phone + "  Latest AFQT " + p.latestAfqt +
      (p.change !== null ? " (" + fmtChange(p.change) + ")" : "") +
      (p.weakest ? "  Weakest: " + REPORT_SECTIONS[p.weakest.code] + " " + p.weakest.pct + "%" : "") +
      (p.ready !== "" ? "  Ready " + p.ready + "/10" : ""));
    if (p.notTenReasons) lines.push("    Why not a 10: " + p.notTenReasons);
    if (p.makeItTen) lines.push("    What would make it a 10: " + p.makeItTen);
    p.tests.forEach(function (t) {
      lines.push("    " + fmtDateTime(parseTimestamp(t["Timestamp"]), tz) + "  Test " + currentTestNumber(t) + "  AFQT " + t["AFQT"]);
    });
  });
  var practice = summarizeDrills(drills);
  lines.push("", "== Practice drills ==");
  if (!practice.length) lines.push("None this week.");
  practice.forEach(function (p) {
    lines.push(displayName(p) + "  " + p.phone + "  " + p.count + (p.count === 1 ? " drill" : " drills"));
  });
  lines.push("", reportSheetUrl());
  return lines.join("\n");
}

// Same table-based, Outlook-safe HTML approach as results-webhook.gs's
// per-test email (no flexbox, no CSS gradients).
function buildHtmlWeeklyReport(people, drills, range, tz, hasPracticeTab) {
  var cell = "padding:6px 8px;border-bottom:1px solid #eee;font-size:13px;vertical-align:top;";
  var head = "padding:6px 8px;border-bottom:2px solid #1e4d8c;font-size:12px;color:#555;text-align:left;text-transform:uppercase;";

  var totalTests = people.reduce(function (n, p) { return n + p.tests.length; }, 0);
  var avg = totalTests
    ? Math.round(people.reduce(function (n, p) {
        return n + p.tests.reduce(function (m, t) { return m + (Number(t["AFQT"]) || 0); }, 0);
      }, 0) / totalTests)
    : null;

  var html = "<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" role=\"presentation\" style=\"max-width:720px;font-family:Arial,Helvetica,sans-serif;color:#1a1a1a;\"><tr><td style=\"padding:8px 4px;\">" +
    "<h2 style=\"margin:0 0 4px;font-size:22px;\">Weekly EST Report</h2>" +
    "<p style=\"margin:0 0 16px;font-size:13px;color:#555;\">" + esc(range) + "</p>" +
    "<p style=\"margin:0 0 20px;font-size:15px;\"><b>" + totalTests + "</b> graded tests &middot; <b>" + people.length +
    "</b> applicants" + (avg !== null ? " &middot; average projected AFQT <b>" + avg + "</b>" : "") +
    " &middot; <b>" + drills.length + "</b> practice drills</p>";

  if (!people.length) {
    html += "<p style=\"font-size:14px;\">No graded tests this week.</p>";
  }

  var stations = [];
  people.forEach(function (p) { if (stations.indexOf(p.recruiter) === -1) stations.push(p.recruiter); });
  stations.forEach(function (station) {
    html += "<h3 style=\"margin:20px 0 6px;font-size:17px;color:#1e4d8c;\">" + esc(station) + "</h3>" +
      "<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" role=\"presentation\"><tr>" +
      "<th style=\"" + head + "\">Applicant</th><th style=\"" + head + "\">Tests this week</th>" +
      "<th style=\"" + head + "\">Latest AFQT</th><th style=\"" + head + "\">Weakest section</th>" +
      "<th style=\"" + head + "\">Ready to join</th></tr>";
    people.filter(function (p) { return p.recruiter === station; }).forEach(function (p) {
      var testLines = p.tests.map(function (t) {
        return esc(fmtDateTime(parseTimestamp(t["Timestamp"]), tz)) + " &middot; Test " + currentTestNumber(t) + " &middot; <b>" + esc(t["AFQT"]) + "</b>";
      }).join("<br>");
      var change = p.change !== null
        ? " <span style=\"color:" + (p.change >= 0 ? "#0ca30c" : "#d03b3b") + ";\">(" + esc(fmtChange(p.change)) + ")</span>"
        : "";
      html += "<tr>" +
        "<td style=\"" + cell + "\"><b>" + esc(displayName(p)) + "</b><br><span style=\"color:#555;\">" + esc(p.phone) + "</span>" +
        (p.nextAction ? "<br><span style=\"color:#555;\">Next: " + esc(p.nextAction) + "</span>" : "") + "</td>" +
        "<td style=\"" + cell + "\">" + testLines + "</td>" +
        "<td style=\"" + cell + "font-size:18px;font-weight:700;color:#1e4d8c;\">" + esc(p.latestAfqt) + change + "</td>" +
        "<td style=\"" + cell + "\">" + (p.weakest ? esc(REPORT_SECTIONS[p.weakest.code]) + " (" + p.weakest.pct + "%)" : "&mdash;") + "</td>" +
        "<td style=\"" + cell + "\">" + readinessCell(p) + "</td>" +
        "</tr>";
    });
    html += "</table>";
  });

  html += "<h3 style=\"margin:28px 0 6px;font-size:17px;color:#1e4d8c;\">Practice drills</h3>";
  var practice = summarizeDrills(drills);
  if (!hasPracticeTab) {
    html += "<p style=\"font-size:13px;color:#555;\">No Practice tab yet. It appears once the first practice drill is logged.</p>";
  } else if (!practice.length) {
    html += "<p style=\"font-size:13px;color:#555;\">No practice drills this week.</p>";
  } else {
    html += "<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" role=\"presentation\"><tr>" +
      "<th style=\"" + head + "\">Applicant</th><th style=\"" + head + "\">Station</th>" +
      "<th style=\"" + head + "\">Drills</th><th style=\"" + head + "\">By section (avg)</th></tr>";
    practice.forEach(function (p) {
      var bySection = REPORT_SECTION_ORDER.filter(function (c) { return p.sections[c]; }).map(function (c) {
        var s = p.sections[c];
        return c + " &times;" + s.count + " (" + (s.total ? Math.round((100 * s.correct) / s.total) : 0) + "%)";
      }).join("<br>");
      html += "<tr>" +
        "<td style=\"" + cell + "\"><b>" + esc(displayName(p)) + "</b><br><span style=\"color:#555;\">" + esc(p.phone) + "</span></td>" +
        "<td style=\"" + cell + "\">" + esc(p.recruiter) + "</td>" +
        "<td style=\"" + cell + "font-weight:700;\">" + p.count + "</td>" +
        "<td style=\"" + cell + "\">" + bySection + "</td></tr>";
    });
    html += "</table>";
  }

  html += "<p style=\"margin:24px 0 0;font-size:14px;\"><a href=\"" + reportSheetUrl() + "\" style=\"color:#1e4d8c;\">Open results sheet &#8599;</a></p>" +
    "</td></tr></table>";
  return html;
}

// Latest test's readiness rating plus the applicant's reasons, escaped
// since they typed some of it.
function readinessCell(p) {
  if (p.ready === "") return "&mdash;";
  var html = "<b>" + esc(p.ready) + " / 10</b>";
  if (p.notTenReasons) html += "<br><span style=\"color:#555;\">Why not 10: " + esc(p.notTenReasons) + "</span>";
  if (p.makeItTen) html += "<br><span style=\"color:#555;\">Make it 10: " + esc(p.makeItTen) + "</span>";
  return html;
}

function reportSheetUrl() {
  return "https://docs.google.com/spreadsheets/d/" + REPORT_SHEET_ID + "/edit";
}
