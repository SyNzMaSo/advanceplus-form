// ══════════════════════════════════════════════════════
//  CONFIG
// ══════════════════════════════════════════════════════
const MAIN_SHEET_NAME = 'DATA';
const TYPE_SHEETS     = ['CI1', 'CI2', 'CS1', 'CS2', 'CW', 'CH', 'CR'];
const RESULT_SHEETS   = ['จบงานเรียบร้อยแล้ว', 'ตัดออก'];
const COL_TYPE   = 2;
const COL_PHONE  = 3;
const COL_RESULT = 7;
const COLNUM_TYPE = 3;
const DROPDOWN_COLS = [3, 6, 8];

// ── DEDUP WINDOW (วินาที) ──
// ถ้าข้อมูล name+phone+type เหมือนกันภายใน 15 วินาที = ข้ามไม่บันทึกซ้ำ
const DEDUP_SECONDS = 15;

// ══════════════════════════════════════════════════════
//  doPost — รับข้อมูลจาก HTML form + ป้องกัน duplicate
// ══════════════════════════════════════════════════════
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let data = {};
    if (e.postData && e.postData.contents) data = JSON.parse(e.postData.contents);

    const name     = (data.name     || '').trim();
    const type     = (data.type     || '').trim();
    const phone    = (data.phone    || '').trim();
    const problem  = (data.problem  || '').trim();
    const progress = (data.progress || '').trim();
    const action   = (data.action   || '').trim();
    const result   = (data.result   || '').trim();

    if (!type) return ok({ status:'error', message:'ไม่มีค่า type' });
    if (!name && !phone) return ok({ status:'error', message:'ข้อมูลไม่ครบ' });

    const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
    if (!mainSheet) return ok({ status:'error', message:'ไม่พบชีต DATA' });

    // ── DEDUP CHECK ──
    // ตรวจว่ามีแถวที่ name+phone+type เหมือนกัน และ timestamp ภายใน DEDUP_SECONDS ไหม
    if (isDuplicate(mainSheet, name, phone, type, DEDUP_SECONDS)) {
      Logger.log('⚠️ Duplicate detected: ' + name + ' / ' + phone + ' / ' + type);
      return ok({ status:'duplicate', message:'ข้อมูลนี้เพิ่งถูกบันทึกไปแล้ว' });
    }

    const timestamp = new Date();
    const rowData   = [timestamp, name, type, phone, problem, progress, action, result];

    appendRow(mainSheet, rowData);
    rebuildSheet(ss, type, mainSheet, COL_TYPE);
    if (result && RESULT_SHEETS.indexOf(result) >= 0)
      rebuildSheet(ss, result, mainSheet, COL_RESULT);

    Logger.log('✅ บันทึก: ' + name + ' / ' + type);
    return ok({ status:'ok', type:type });

  } catch(err) {
    Logger.log('❌ doPost: ' + err.message);
    return ok({ status:'error', message:err.message });
  }
}

// ══════════════════════════════════════════════════════
//  isDuplicate — ตรวจแถวล่าสุดว่าซ้ำกันภายใน N วินาทีไหม
// ══════════════════════════════════════════════════════
function isDuplicate(sheet, name, phone, type, seconds) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;

  var now     = new Date();
  var cutoff  = new Date(now.getTime() - seconds * 1000);

  // ตรวจแค่ 20 แถวล่าสุด (เร็วกว่าอ่านทั้งหมด)
  var checkRows = Math.min(20, lastRow - 1);
  var startRow  = lastRow - checkRows + 1;
  var range     = sheet.getRange(startRow, 1, checkRows, 5); // A-E
  var values    = range.getValues();

  for (var i = values.length - 1; i >= 0; i--) {
    var rowTime  = values[i][0]; // A = timestamp
    var rowName  = (values[i][1] || '').toString().trim(); // B = ชื่อ
    var rowType  = (values[i][2] || '').toString().trim(); // C = ประเภท
    var rowPhone = (values[i][3] || '').toString().trim(); // D = เบอร์

    if (!(rowTime instanceof Date)) continue;
    if (rowTime < cutoff) continue; // เก่าเกินไป ไม่นับ

    // ตรวจ name + phone + type ตรงกันหมด
    if (rowName === name && rowPhone === phone && rowType === type) {
      return true; // 🔴 DUPLICATE
    }
  }
  return false;
}

// ══════════════════════════════════════════════════════
//  Triggers
// ══════════════════════════════════════════════════════
function createTriggers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onEditHandler').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('onChangeHandler').forSpreadsheet(ss).onChange().create();
  Logger.log('✅ Triggers created');
}

function onEditHandler(e) {
  try {
    var sheet = e.range.getSheet();
    if (sheet.getName() !== MAIN_SHEET_NAME) return;
    var editedRow = e.range.getRow();
    var editedCol = e.range.getColumn();
    if (editedRow <= 1) return;
    var ss = e.source;
    var mainSheet = sheet;
    var numCols = mainSheet.getLastColumn();
    var rowData = mainSheet.getRange(editedRow, 1, 1, numCols).getValues()[0];
    var currentType   = (rowData[COL_TYPE]   || '').toString().trim();
    var currentResult = (rowData[COL_RESULT] || '').toString().trim();
    if (currentType && TYPE_SHEETS.indexOf(currentType) >= 0)
      rebuildSheet(ss, currentType, mainSheet, COL_TYPE);
    if (editedCol === COLNUM_TYPE) {
      var oldType = (e.oldValue || '').toString().trim();
      if (oldType && oldType !== currentType && TYPE_SHEETS.indexOf(oldType) >= 0)
        rebuildSheet(ss, oldType, mainSheet, COL_TYPE);
    }
    RESULT_SHEETS.forEach(function(r) {
      if (ss.getSheetByName(r)) rebuildSheet(ss, r, mainSheet, COL_RESULT);
    });
  } catch(err) { Logger.log('❌ onEditHandler: ' + err.message); }
}

function onChangeHandler(e) {
  try {
    var changeType = e.changeType;
    if (changeType === 'INSERT_COLUMN' || changeType === 'REMOVE_COLUMN') return;
    var ss = e.source;
    var activeSheet = ss.getActiveSheet();
    if (!activeSheet || activeSheet.getName() !== MAIN_SHEET_NAME) return;
    var mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
    if (!mainSheet) return;
    rebuildAllSheets(ss, mainSheet);
  } catch(err) { Logger.log('❌ onChangeHandler: ' + err.message); }
}

function rebuildAllSheets(ss, mainSheet) {
  TYPE_SHEETS.forEach(function(name) {
    if (ss.getSheetByName(name)) rebuildSheet(ss, name, mainSheet, COL_TYPE);
  });
  RESULT_SHEETS.forEach(function(name) {
    if (ss.getSheetByName(name)) rebuildSheet(ss, name, mainSheet, COL_RESULT);
  });
}

function rebuildSheet(ss, sheetName, mainSheet, filterCol) {
  var targetSheet = ss.getSheetByName(sheetName);
  if (!targetSheet) targetSheet = ss.insertSheet(sheetName);
  var mainLastRow = mainSheet.getLastRow();
  var mainLastCol = mainSheet.getLastColumn();
  if (mainLastCol === 0) return;
  var header = mainSheet.getRange(1, 1, 1, mainLastCol).getValues()[0];
  if (targetSheet.getLastRow() === 0) {
    targetSheet.getRange(1, 1, 1, header.length).setValues([header]);
    targetSheet.getRange(1, 1, 1, header.length)
      .setBackground('#1a5c30').setFontColor('#ffffff').setFontWeight('bold');
  }
  var curLast = targetSheet.getLastRow();
  if (curLast > 1) targetSheet.deleteRows(2, curLast - 1);
  if (mainLastRow <= 1) return;
  var srcRange  = mainSheet.getRange(2, 1, mainLastRow - 1, mainLastCol);
  var allValues = srcRange.getValues();
  var allBgs    = srcRange.getBackgrounds();
  var allFonts  = srcRange.getFontColors();
  var mValues = [], mBgs = [], mFonts = [];
  for (var i = 0; i < allValues.length; i++) {
    var row = allValues[i];
    if (!row[0]) continue;
    if ((row[filterCol] || '').toString().trim() !== sheetName) continue;
    mValues.push(row); mBgs.push(allBgs[i]); mFonts.push(allFonts[i]);
  }
  if (mValues.length === 0) return;
  var destRange = targetSheet.getRange(2, 1, mValues.length, mainLastCol);
  destRange.setValues(mValues);
  destRange.setBackgrounds(mBgs);
  destRange.setFontColors(mFonts);
  DROPDOWN_COLS.forEach(function(col) {
    if (col <= mainLastCol)
      targetSheet.getRange(2, col, mValues.length, 1).setFontColor(null);
  });
  targetSheet.getRange(2, COL_PHONE + 1, mValues.length, 1).setNumberFormat('@STRING@');
}

function appendRow(sheet, rowData) {
  var nextRow = sheet.getLastRow() + 1;
  var numCols = rowData.length;
  var range   = sheet.getRange(nextRow, 1, 1, numCols);
  range.setNumberFormats([rowData.map(function(_, i) {
    return i === COL_PHONE ? '@STRING@' : 'General';
  })]);
  range.setValues([rowData.map(function(v, i) {
    return i === COL_PHONE ? String(v) : v;
  })]);
}

function ensureSubSheetsExist() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!mainSheet) { Logger.log('❌ ไม่พบ DATA'); return; }
  var all = TYPE_SHEETS.concat(RESULT_SHEETS);
  all.forEach(function(name) {
    var fc = TYPE_SHEETS.indexOf(name) >= 0 ? COL_TYPE : COL_RESULT;
    rebuildSheet(ss, name, mainSheet, fc);
  });
  if (mainSheet.getLastRow() > 1)
    mainSheet.getRange('D:D').setNumberFormat('@STRING@');
  Logger.log('✅ Done');
}

function fixExcelDates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!mainSheet) return;
  var lastRow = mainSheet.getLastRow();
  if (lastRow <= 1) return;
  var colA = mainSheet.getRange(2, 1, lastRow - 1, 1);
  var values = colA.getValues();
  var fixed = 0;
  for (var i = 0; i < values.length; i++) {
    var cell = values[i][0];
    if (!(cell instanceof Date)) continue;
    var year = cell.getFullYear();
    if (year >= 1900 && year < 2400) {
      var c = new Date(cell); c.setFullYear(year + 543);
      values[i][0] = c; fixed++;
    }
  }
  if (fixed > 0) {
    colA.setValues(values);
    rebuildAllSheets(ss, mainSheet);
    Logger.log('✅ แก้วันที่ ' + fixed + ' เซลล์');
  }
}

function ok(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}