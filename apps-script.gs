// Google Apps Script: collects every group's submission into a Google Sheet.
// Setup: create a Google Sheet > Extensions > Apps Script > paste this file > Deploy > New deployment >
// type "Web app", Execute as "Me", Who has access "Anyone" > copy the web app URL into
// SUBMIT_URL (index.html) and SHEET_URL (scores.html).
var SHEET_NAME = "Results";

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange("A:D").setNumberFormat("@");          // keep everything as plain text
    sh.appendRow(["Group", "Time", "Items (id:qty)", "Received"]);
  }
  return sh;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var d = JSON.parse(e.postData.contents);
    var g = String(d.g || "").slice(0, 60);
    var t = String(Math.floor(Number(d.t) || 0));
    var q = String(d.q || "");
    if (q.length > 2000 || !/^[a-z0-9-]+:\d+(,[a-z0-9-]+:\d+)*$/.test(q)) {
      return ContentService.createTextOutput("rejected");
    }
    sheet_().appendRow([g, t, q, new Date().toISOString()]);
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  var values = sheet_().getDataRange().getValues().slice(1);
  var rows = values.map(function (r) { return { g: String(r[0]), t: Number(r[1]) || 0, q: String(r[2]) }; });
  return ContentService.createTextOutput(JSON.stringify({ rows: rows }))
    .setMimeType(ContentService.MimeType.JSON);
}
