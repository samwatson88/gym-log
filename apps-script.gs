// ASCEND sync script. Paste into Apps Script (replacing the old code), set TOKEN,
// then Deploy > Manage deployments > Edit > New version > Deploy.

const TOKEN = "replace-with-a-long-secret";
const PROGRAMME_SHEET = "Programme";
const PROG_COLS = ["day", "exercise", "sets", "min", "max"];
const LOG_COLS = ["timestamp", "workout_date", "exercise", "set", "weight_kg", "reps", "set_time", "rest_seconds", "session_id", "day_name", "rir", "warmup"];

function doGet(e) {
  if (!e.parameter || e.parameter.token !== TOKEN) return out({ error: "unauthorised" });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sets = readRows(ss.getSheets()[0]).filter(r => r.exercise && r.exercise !== "TEST");
  const program = readRows(programmeSheet(ss)).filter(r => r.day && r.exercise);
  return out({ sets: sets, program: program });
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  if (d.token !== TOKEN) return out({ error: "unauthorised" });
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (d.action === "program") {
    const sh = programmeSheet(ss);
    sh.clearContents();
    sh.getRange(1, 1, 1, PROG_COLS.length).setValues([PROG_COLS]);
    if (d.rows && d.rows.length) {
      sh.getRange(2, 1, d.rows.length, PROG_COLS.length)
        .setValues(d.rows.map(r => [r.day, r.exercise, r.sets, r.min, r.max]));
    }
    return out({ ok: true });
  }

  const sh = ss.getSheets()[0];
  ensureHeader(sh);
  const row = [d.timestamp, d.workoutDate, d.exercise, d.set, d.weight, d.reps, d.setTime, d.restSeconds, d.sessionId || "", d.dayName || "", d.rir ?? "", d.warmup || ""];
  const r = sh.getLastRow() + 1;
  // Store dates and times as plain text so they read back exactly as the app wrote them.
  [1, 2, 7].forEach(c => sh.getRange(r, c).setNumberFormat("@"));
  sh.getRange(r, 1, 1, row.length).setValues([row]);
  return out({ ok: true });
}

function programmeSheet(ss) {
  let sh = ss.getSheetByName(PROGRAMME_SHEET);
  if (!sh) {
    sh = ss.insertSheet(PROGRAMME_SHEET);
    sh.getRange(1, 1, 1, PROG_COLS.length).setValues([PROG_COLS]);
  }
  return sh;
}

function ensureHeader(sh) {
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, LOG_COLS.length).setValues([LOG_COLS]);
    return;
  }
  const lastCol = sh.getLastColumn();
  const head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const missing = LOG_COLS.filter(c => head.indexOf(c) === -1);
  if (missing.length) sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
}

function readRows(sh) {
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  const head = v[0].map(h => String(h).trim());
  return v.slice(1).map(row => {
    const o = {};
    head.forEach((h, i) => { o[h] = norm(row[i], h); });
    return o;
  });
}

function norm(x, header) {
  if (x instanceof Date) {
    return header === "workout_date"
      ? Utilities.formatDate(x, "UTC", "yyyy-MM-dd")
      : Utilities.formatDate(x, "UTC", "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
  }
  return x;
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
