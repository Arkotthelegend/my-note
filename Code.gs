/**
 * G-12 / G-11 / G-10 paid-user Google Apps Script
 * ------------------------------------------------
 * Paste this WHOLE file into the Apps Script BOUND to TG APP SHEET
 * (Extensions → Apps Script on that spreadsheet). Do not paste the
 * statistics/rank script into this project. Those are two different apps.
 *
 * Deploy as Web App (Execute as: Me, Who has access: Anyone).
 * Then Deploy → Manage deployments → existing Web app → New version.
 *
 * GOOGLE SHEET — Row 1 headers, lowercase:
 *
 *   id
 *   all, mm, en, math, phy, chem, bio, eco
 *   g11_all, g11_mm, ...   g10_all, g10_mm, ...
 *
 * Volunteer unlocks (separate from paid, same date format):
 *   vol_all, vol_mm, vol_en, vol_math, vol_phy, vol_chem, vol_bio, vol_eco
 *   g11_vol_all, g11_vol_mm, ...   g10_vol_all, g10_vol_mm, ...
 *
 * VolunteerStats tab is created when you first save volunteer stats:
 *   id, reports, note, updated
 *
 * After pasting, run setupDailyCleanup() ONCE from the Apps Script editor.
 * Project timezone: File → Project settings → Asia/Yangon
 */

var TZ = 'Asia/Yangon';

var SUBJECT_IDS = ['all', 'mm', 'en', 'math', 'phy', 'chem', 'bio', 'eco'];
var GRADES = [12, 11, 10];

var SUBJECT_LABELS = {
  all: 'All',
  mm: 'Myanmar',
  en: 'English',
  math: 'Mathematics',
  phy: 'Physics',
  chem: 'Chemistry',
  bio: 'Biology',
  eco: 'Economics'
};

var OLD_NAME_TO_ID = {
  all: 'all',
  myanmar: 'mm',
  english: 'en',
  mathematics: 'math',
  physics: 'phy',
  chemistry: 'chem',
  biology: 'bio',
  economics: 'eco'
};

function allColumnNames() {
  var names = [];
  GRADES.forEach(function (grade) {
    var prefix = grade === 12 ? '' : 'g' + grade + '_';
    SUBJECT_IDS.forEach(function (id) {
      names.push(prefix + id);
    });
  });
  return names;
}

function allVolunteerColumnNames() {
  var names = [];
  GRADES.forEach(function (grade) {
    var prefix = grade === 12 ? 'vol_' : 'g' + grade + '_vol_';
    SUBJECT_IDS.forEach(function (id) {
      names.push(prefix + id);
    });
  });
  return names;
}

function columnFor(grade, subjectId) {
  var g = parseInt(grade, 10) || 12;
  var prefix = g === 12 ? '' : 'g' + g + '_';
  return prefix + subjectId;
}

function volunteerColumnFor(grade, subjectId) {
  var g = parseInt(grade, 10) || 12;
  var prefix = g === 12 ? 'vol_' : 'g' + g + '_vol_';
  return prefix + subjectId;
}

function parseColumn(column) {
  var col = String(column || '').trim().toLowerCase();
  var grade = 12;
  var subjectId = col;
  var kind = 'paid';
  var vol = col.match(/^(?:g(10|11)_)?vol_(all|mm|en|math|phy|chem|bio|eco)$/);
  if (vol) {
    kind = 'volunteer';
    if (vol[1]) grade = parseInt(vol[1], 10);
    subjectId = vol[2];
  } else {
    var m = col.match(/^g(10|11)_(.+)$/);
    if (m) {
      grade = parseInt(m[1], 10);
      subjectId = m[2];
    }
  }
  return {
    column: col,
    grade: grade,
    subjectId: subjectId,
    kind: kind,
    label: SUBJECT_LABELS[subjectId] || subjectId
  };
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function todayYmd() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
}

function toYmd(value) {
  if (value === null || value === undefined || value === '') return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TZ, 'yyyy-MM-dd');
  }
  var s = String(value).trim();
  if (!s) return '';
  var m = s.match(/^(\d{4}-\d{2}-\d{2})/);
  if (m) return m[1];
  var d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  return '';
}

function isExpiredYmd(ymd) {
  if (!ymd) return false;
  return ymd < todayYmd();
}

function getUsersSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName('Users') || ss.getSheetByName('Sheet1') || ss.getSheets()[0];
}

function getLogsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Logs') || ss.getSheetByName('Log');
  if (!sh) {
    sh = ss.insertSheet('Logs');
    sh.appendRow(['timestamp', 'id', 'subject', 'months', 'unit', 'expiry', 'grade', 'column', 'role']);
  }
  return sh;
}

function getVolunteerStatsSheet(createIfMissing) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('VolunteerStats');
  if (!sh && createIfMissing) {
    sh = ss.insertSheet('VolunteerStats');
    sh.appendRow(['id', 'reports', 'note', 'updated']);
  }
  return sh;
}

function headerMap(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var key = String(headers[i] || '').trim().toLowerCase();
    if (key) map[key] = i + 1;
  }
  if (!map.id) map.id = 1;
  return { map: map, headers: headers, lastCol: lastCol };
}

function ensureHeaders() {
  var sheet = getUsersSheet();
  var info = headerMap(sheet);
  var needed = ['id'].concat(allColumnNames()).concat(allVolunteerColumnNames());
  var added = [];
  needed.forEach(function (name) {
    if (!info.map[name]) {
      var col = sheet.getLastColumn() + 1;
      if (sheet.getLastColumn() === 0) col = 1;
      sheet.getRange(1, col).setValue(name);
      info.map[name] = col;
      added.push(name);
    }
  });
  return added;
}

function findUserRow(sheet, id) {
  var info = headerMap(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var ids = sheet.getRange(2, info.map.id, lastRow - 1, 1).getValues();
  var want = String(id).trim();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === want) return i + 2;
  }
  return 0;
}

function subjectColumns(info) {
  var known = {};
  allColumnNames().forEach(function (n) { known[n] = true; });
  var cols = [];
  Object.keys(info.map).forEach(function (key) {
    if (key === 'id') return;
    if (known[key] || /^(g1[01]_)?(all|mm|en|math|phy|chem|bio|eco)$/.test(key)) {
      cols.push(key);
    }
  });
  return cols;
}

function volunteerColumns(info) {
  var known = {};
  allVolunteerColumnNames().forEach(function (n) { known[n] = true; });
  var cols = [];
  Object.keys(info.map).forEach(function (key) {
    if (known[key] || /^(g1[01]_)?vol_(all|mm|en|math|phy|chem|bio|eco)$/.test(key)) {
      cols.push(key);
    }
  });
  return cols;
}

function accessColumns(info) {
  return subjectColumns(info).concat(volunteerColumns(info));
}

function rowHasAnySubject(sheet, row, info) {
  var cols = accessColumns(info);
  for (var i = 0; i < cols.length; i++) {
    if (!info.map[cols[i]]) continue;
    var val = sheet.getRange(row, info.map[cols[i]]).getValue();
    if (toYmd(val) || (val !== '' && val !== null)) return true;
  }
  return false;
}

function volunteerAccessKey(volColumn) {
  return String(volColumn || '').replace('vol_', '');
}

function readVolunteerStatsMap() {
  var sh = getVolunteerStatsSheet(false);
  var map = {};
  if (!sh) return map;
  var info = headerMap(sh);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return map;
  var rows = sh.getRange(2, 1, lastRow - 1, info.lastCol).getValues();
  for (var i = 0; i < rows.length; i++) {
    var id = String(rows[i][info.map.id - 1] || '').trim();
    if (!id) continue;
    map[id] = {
      reports: parseInt(rows[i][(info.map.reports || 2) - 1], 10) || 0,
      note: String(rows[i][(info.map.note || 3) - 1] || '')
    };
  }
  return map;
}

function upsertVolunteerStats(id, reports, note) {
  var sh = getVolunteerStatsSheet(true);
  var info = headerMap(sh);
  var want = String(id || '').trim();
  if (!want) return { success: false, message: 'Missing id' };
  var row = findUserRow(sh, want);
  if (!row) {
    row = Math.max(sh.getLastRow() + 1, 2);
    sh.getRange(row, info.map.id).setValue(want);
  }
  if (reports !== undefined && reports !== null && reports !== '') {
    sh.getRange(row, info.map.reports || 2).setValue(parseInt(reports, 10) || 0);
  }
  if (note !== undefined) {
    sh.getRange(row, info.map.note || 3).setValue(String(note));
  }
  sh.getRange(row, info.map.updated || 4).setValue(new Date());
  return { success: true };
}

function bumpVolunteerReport(id) {
  var sh = getVolunteerStatsSheet(true);
  var info = headerMap(sh);
  var want = String(id || '').trim();
  if (!want) return { success: false, message: 'Missing id' };
  var reportsCol = info.map.reports || 2;
  var row = findUserRow(sh, want);
  if (!row) {
    row = Math.max(sh.getLastRow() + 1, 2);
    sh.getRange(row, info.map.id).setValue(want);
    sh.getRange(row, reportsCol).setValue(1);
  } else {
    var current = parseInt(sh.getRange(row, reportsCol).getValue(), 10) || 0;
    sh.getRange(row, reportsCol).setValue(current + 1);
  }
  sh.getRange(row, info.map.updated || 4).setValue(new Date());
  return { success: true };
}

function resolveSubject(item) {
  if (item.column) return parseColumn(item.column);
  var raw = String(item.subject || '').trim();
  var grade = parseInt(item.grade, 10) || 12;
  var id = OLD_NAME_TO_ID[raw.toLowerCase()] || String(raw).toLowerCase();
  var col = columnFor(grade, id);
  return parseColumn(col);
}

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || '';
    if (action === 'getUsers') return jsonOut(buildAppUsers());
    if (action === 'cleanupExpired') return jsonOut(cleanupExpired());
    return jsonOut(buildAdminData());
  } catch (err) {
    return jsonOut({ success: false, message: String(err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var data = {};
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    }
    var action = data.action || '';
    if (action === 'saveMulti') return jsonOut(saveMulti(data));
    if (action === 'delete') return jsonOut(deleteSubject(data));
    if (action === 'cleanupExpired') return jsonOut(cleanupExpired());
    if (action === 'saveVolunteerStats') return jsonOut(upsertVolunteerStats(data.id, data.reports, data.note));
    if (action === 'bumpVolunteerReport') return jsonOut(bumpVolunteerReport(data.id));
    return jsonOut({ success: false, message: 'Unknown action' });
  } catch (err) {
    return jsonOut({ success: false, message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function buildAppUsers() {
  var sheet = getUsersSheet();
  var info = headerMap(sheet);
  var lastRow = sheet.getLastRow();
  var result = {};
  var volStats = readVolunteerStatsMap();
  if (lastRow < 2) return result;

  var range = sheet.getRange(2, 1, lastRow - 1, info.lastCol).getValues();
  var paidCols = subjectColumns(info);
  var volCols = volunteerColumns(info);

  for (var r = 0; r < range.length; r++) {
    var id = String(range[r][info.map.id - 1] || '').trim();
    if (!id) continue;
    var entry = {};
    paidCols.forEach(function (key) {
      if (!info.map[key]) return;
      var ymd = toYmd(range[r][info.map[key] - 1]);
      if (ymd && !isExpiredYmd(ymd)) entry[key] = ymd;
    });
    var vol = {};
    volCols.forEach(function (key) {
      if (!info.map[key]) return;
      var ymd = toYmd(range[r][info.map[key] - 1]);
      if (ymd && !isExpiredYmd(ymd)) vol[volunteerAccessKey(key)] = ymd;
    });
    var reports = volStats[id] ? parseInt(volStats[id].reports, 10) || 0 : 0;
    if (Object.keys(vol).length) {
      entry.vol = vol;
      entry.isVolunteer = true;
      entry.volReports = reports;
    } else if (reports > 0) {
      entry.isVolunteer = true;
      entry.volReports = reports;
    }
    if (Object.keys(entry).length) result[id] = entry;
  }
  return result;
}

function buildAdminData() {
  var sheet = getUsersSheet();
  var info = headerMap(sheet);
  var lastRow = sheet.getLastRow();
  var active = [];
  var cols = accessColumns(info);

  if (lastRow >= 2) {
    var range = sheet.getRange(2, 1, lastRow - 1, info.lastCol).getValues();
    for (var r = 0; r < range.length; r++) {
      var id = String(range[r][info.map.id - 1] || '').trim();
      if (!id) continue;
      cols.forEach(function (key) {
        if (!info.map[key]) return;
        var ymd = toYmd(range[r][info.map[key] - 1]);
        if (!ymd) return;
        var meta = parseColumn(key);
        active.push({
          id: id,
          subject: meta.label,
          grade: meta.grade,
          column: meta.column,
          expiry: ymd,
          kind: meta.kind,
          volunteer: meta.kind === 'volunteer'
        });
      });
    }
  }

  var logs = [];
  var logSheet = getLogsSheet();
  var logLast = logSheet.getLastRow();
  if (logLast >= 2) {
    var logLastCol = Math.max(logSheet.getLastColumn(), 1);
    var logHeaders = logSheet.getRange(1, 1, 1, logLastCol).getValues()[0].map(function (h) {
      return String(h || '').trim().toLowerCase();
    });
    var rows = logSheet.getRange(2, 1, logLast - 1, logLastCol).getValues();
    function logCol(name, fallbackIndex) {
      var i = logHeaders.indexOf(name);
      return i >= 0 ? i : fallbackIndex;
    }
    var iTs = logCol('timestamp', 0);
    var iId = logCol('id', 1);
    var iSub = logCol('subject', 2);
    var iMonths = logCol('months', 3);
    var iUnit = logCol('unit', 4);
    var iExp = logCol('expiry', 5);
    var iGrade = logCol('grade', 6);
    var iCol = logCol('column', 7);
    var iRole = logCol('role', 8);
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      if (!row[iId] && !row[iTs]) continue;
      logs.push({
        timestamp: row[iTs] instanceof Date ? row[iTs].toISOString() : row[iTs],
        id: row[iId],
        subject: row[iSub],
        months: row[iMonths],
        unit: row[iUnit] || 'months',
        expiry: row[iExp] instanceof Date ? toYmd(row[iExp]) : row[iExp],
        grade: row[iGrade] || 12,
        column: row[iCol] || '',
        role: row[iRole] || ''
      });
    }
  }

  return { active: active, logs: logs, volunteerStats: readVolunteerStatsMap() };
}

function saveMulti(data) {
  ensureHeaders();
  var sheet = getUsersSheet();
  var info = headerMap(sheet);
  var id = String(data.id || '').trim();
  if (!id) return { success: false, message: 'Missing id' };

  var row = findUserRow(sheet, id);
  if (!row) {
    row = sheet.getLastRow() + 1;
    if (row < 2) row = 2;
    sheet.getRange(row, info.map.id).setValue(id);
  }

  var items = data.subjects || [];
  var logSheet = getLogsSheet();
  var grantRole = data.role === 'volunteer' ? 'volunteer' : 'paid';
  items.forEach(function (item) {
    var itemRole = item.role === 'volunteer' ? 'volunteer' : grantRole;
    var meta = resolveSubject(item);
    if (itemRole === 'volunteer' && meta.kind !== 'volunteer') {
      meta = parseColumn(volunteerColumnFor(meta.grade, meta.subjectId));
    }
    var col = info.map[meta.column];
    if (!col) return;
    var expiry = toYmd(item.expiry) || item.expiry;
    sheet.getRange(row, col).setValue(expiry);
    logSheet.appendRow([
      new Date(),
      id,
      meta.label,
      item.duration || item.months || '',
      item.unit || 'months',
      expiry,
      meta.grade,
      meta.column,
      itemRole
    ]);
  });

  return { success: true };
}

function deleteSubject(data) {
  var sheet = getUsersSheet();
  var info = headerMap(sheet);
  var id = String(data.id || '').trim();
  var row = findUserRow(sheet, id);
  if (!row) return { success: false, message: 'User not found' };

  var meta = resolveSubject(data);
  var col = info.map[meta.column];
  if (col) sheet.getRange(row, col).setValue('');

  info = headerMap(sheet);
  if (!rowHasAnySubject(sheet, row, info)) {
    sheet.deleteRow(row);
  }
  return { success: true };
}

function cleanupExpired() {
  ensureHeaders();
  var sheet = getUsersSheet();
  var info = headerMap(sheet);
  var lastRow = sheet.getLastRow();
  var cleared = 0;
  var deletedRows = 0;
  if (lastRow < 2) return { success: true, cleared: 0, deletedRows: 0 };

  var cols = accessColumns(info);
  var width = info.lastCol;
  var height = lastRow - 1;
  var values = sheet.getRange(2, 1, height, width).getValues();
  var toDelete = [];

  for (var r = 0; r < values.length; r++) {
    var stillHas = false;
    cols.forEach(function (key) {
      var idx = info.map[key] - 1;
      var ymd = toYmd(values[r][idx]);
      if (!ymd) return;
      if (isExpiredYmd(ymd)) {
        values[r][idx] = '';
        cleared++;
      } else {
        stillHas = true;
      }
    });
    if (!stillHas) toDelete.push(r + 2);
  }

  sheet.getRange(2, 1, height, width).setValues(values);

  for (var i = toDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(toDelete[i]);
    deletedRows++;
  }

  return { success: true, cleared: cleared, deletedRows: deletedRows };
}

function setupDailyCleanup() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (t) {
    if (t.getHandlerFunction() === 'cleanupExpired') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('cleanupExpired').timeBased().everyDays(1).atHour(1).create();
}
