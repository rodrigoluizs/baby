/**
 * Baby betting pool backend — Google Apps Script Web App.
 *
 * Deploy: Extensions > Apps Script on the pool spreadsheet, paste this file,
 * then Deploy > New deployment > Web app (execute as: Me, access: Anyone).
 * The "bets" and "actual" sheets are created automatically on first request.
 *
 * See apps-script/README.md for the full setup guide.
 */

var DUE_DATE = new Date(2026, 8, 7); // 2026-09-07 = week 40+0
var SEX = ["boy", "girl"];
var WEEKDAYS = ["sun", "mon", "tue", "wed", "thu", "fri", "sat"]; // index = Date.getDay()
var DAY_NIGHT = ["day", "night"];
var HAIR = ["bald", "some", "lots"];
var EYES = ["blue", "brown", "green"];
var BETS_HEADERS = ["ts", "name", "sex", "week", "weekday", "dayNight", "weight", "length", "hair", "eyes", "lang"];
var ACTUAL_HEADERS = ["sex", "birth_date", "birth_time", "weight_grams", "length_cm", "hair", "eyes"];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var bet = validateBet_(JSON.parse(e.postData.contents));
    if (!bet) return json_({ ok: false, error: "invalid bet" });
    ensureBetsSheet_().appendRow([
      new Date().toISOString(), bet.name, bet.sex, bet.week, bet.weekday,
      bet.dayNight, bet.weight, bet.length, bet.hair, bet.eyes, bet.lang,
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  try {
    return json_({ ok: true, bets: readBets_(), actual: readActual_() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function validateBet_(data) {
  if (!data || typeof data !== "object") return null;
  var name = String(data.name || "").trim().slice(0, 60);
  var week = Number(data.week);
  var weight = Number(data.weight);
  var length = Number(data.length);
  var valid = name.length > 0 &&
    SEX.indexOf(data.sex) !== -1 &&
    week >= 37 && week <= 42 &&
    WEEKDAYS.indexOf(data.weekday) !== -1 &&
    DAY_NIGHT.indexOf(data.dayNight) !== -1 &&
    weight >= 2000 && weight <= 6000 &&
    length >= 40 && length <= 60 &&
    HAIR.indexOf(data.hair) !== -1 &&
    EYES.indexOf(data.eyes) !== -1;
  if (!valid) return null;
  return {
    name: name,
    sex: data.sex,
    week: Math.round(week),
    weekday: data.weekday,
    dayNight: data.dayNight,
    weight: Math.round(weight),
    length: Math.round(length),
    hair: data.hair,
    eyes: data.eyes,
    lang: String(data.lang || "").slice(0, 5),
  };
}

function readBets_() {
  var rows = ensureBetsSheet_().getDataRange().getValues();
  var bets = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    bets.push({
      ts: String(r[0]), name: String(r[1]), sex: String(r[2]), week: Number(r[3]),
      weekday: String(r[4]), dayNight: String(r[5]), weight: Number(r[6]),
      length: Number(r[7]), hair: String(r[8]), eyes: String(r[9]),
    });
  }
  bets.sort(function (a, b) { return a.ts < b.ts ? -1 : 1; });
  return bets;
}

/**
 * Reads the real birth data and derives week/weekday/dayNight from
 * birth_date + birth_time, so the parents only fill plain facts.
 * Returns null until every column is filled.
 */
function readActual_() {
  var rows = ensureActualSheet_().getDataRange().getValues();
  if (rows.length < 2) return null;
  var r = rows[1];
  var sex = String(r[0] || "").trim();
  var birthDate = parseDate_(r[1]);
  var hours = parseHours_(r[2]);
  var weight = Number(r[3]);
  var length = Number(r[4]);
  var hair = String(r[5] || "").trim();
  var eyes = String(r[6] || "").trim();
  if (SEX.indexOf(sex) === -1 || !birthDate || hours === null ||
      !(weight > 0) || !(length > 0) || HAIR.indexOf(hair) === -1 ||
      EYES.indexOf(eyes) === -1) {
    return null;
  }
  var daysFromDue = Math.round((birthDate.getTime() - DUE_DATE.getTime()) / 86400000);
  var week = Math.floor((280 + daysFromDue) / 7); // completed gestational weeks
  return {
    sex: sex,
    week: Math.max(37, Math.min(42, week)), // bets use 37 = "≤37" and 42 = "42+"
    weekday: WEEKDAYS[birthDate.getDay()],
    dayNight: hours >= 6 && hours < 18 ? "day" : "night",
    weight: Math.round(weight),
    length: Math.round(length),
    hair: hair,
    eyes: eyes,
  };
}

function parseDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  var m = String(v || "").trim().match(/^(\d{1,2})[\/.](\d{1,2})[\/.](\d{4})$/); // dd/mm/yyyy
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  m = String(v || "").trim().match(/^(\d{4})-(\d{2})-(\d{2})$/); // yyyy-mm-dd
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return null;
}

function parseHours_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v.getHours();
  if (typeof v === "number" && v >= 0 && v < 1) return Math.floor(v * 24); // sheet time serial
  var m = String(v || "").trim().match(/^(\d{1,2})[:h](\d{2})?/);
  if (m && Number(m[1]) < 24) return Number(m[1]);
  return null;
}

function ensureBetsSheet_() {
  return ensureSheet_("bets", BETS_HEADERS);
}

function ensureActualSheet_() {
  var sheet = ensureSheet_("actual", ACTUAL_HEADERS);
  if (sheet.getMaxRows() >= 2 && !sheet.getRange(2, 1).getDataValidation()) {
    sheet.getRange(2, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(SEX, true).build());
    sheet.getRange(2, 6).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(HAIR, true).build());
    sheet.getRange(2, 7).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(EYES, true).build());
  }
  return sheet;
}

function ensureSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
