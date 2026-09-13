/**
 * DE TEAM — ระบบเช็คชื่อพนักงานและคำนวณยอดค่าจ้าง
 * Repository.gs: การเชื่อมต่อ Google Sheets, LockService, batchUpdate และ Data Mapping
 */

function spreadsheet_() {
  var id = getConfigProperty_('SPREADSHEET_ID');
  if (id && typeof id === 'string' && /^[\w-]+$/.test(id)) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (e) {}
  }
  if (typeof SpreadsheetApp !== 'undefined' && SpreadsheetApp.getActiveSpreadsheet) {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  }
  requireValue_(id && typeof id === 'string' && /^[\w-]+$/.test(id), 'CONFIG_REQUIRED');
  return SpreadsheetApp.openById(id);
}

function bangkokToday_() {
  return Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
}

function bangkokNowIso_() {
  return new Date().toISOString();
}

function locked_(operation) {
  var lock = LockService.getScriptLock();
  requireValue_(lock.tryLock(3000), 'BUSY');
  try {
    return operation();
  } finally {
    lock.releaseLock();
  }
}

function table_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  requireValue_(sheet !== null, 'SCHEMA_REQUIRED');
  var data = sheet.getDataRange().getValues();
  requireValue_(data.length && data[0].length === headers.length && headers.every(function (h, i) { return data[0][i] === h; }), 'SCHEMA_MISMATCH');

  var rows = [];
  data.slice(1).forEach(function (values, index) {
    if (values.every(function (v) { return v === ''; })) return;
    var obj = { rowNumber: index + 2 };
    headers.forEach(function (h, i) {
      obj[h] = values[i];
    });
    rows.push(obj);
  });
  return { sheet: sheet, rows: rows };
}

function unique_(rows, field, value) {
  var found = rows.filter(function (r) { return r[field] === value; });
  requireValue_(found.length <= 1, 'DUPLICATE_KEY');
  return found[0] || null;
}

function suffix_(month) {
  return monthKey_(month).replace('-', '_');
}

function fingerprint_(data) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(data), Utilities.Charset.UTF_8);
  return bytes.map(function (b) {
    return ('0' + ((b + 256) % 256).toString(16)).slice(-2);
  }).join('');
}

function cells_(values) {
  return values.map(function (v) {
    if (typeof v === 'number') {
      requireValue_(Number.isFinite(v), 'INVALID_NUMBER');
      return { userEnteredValue: { numberValue: v } };
    }
    // Explicit stringValue ป้องกัน Formula Injection เช่น '=IMPORTXML(...)'
    return { userEnteredValue: { stringValue: String(v === undefined || v === null ? '' : v) } };
  });
}

function append_(sheet, values) {
  return {
    appendCells: {
      sheetId: sheet.getSheetId(),
      rows: [{ values: cells_(values) }],
      fields: 'userEnteredValue'
    }
  };
}

function updateRow_(sheet, rowIndex, values) {
  return {
    updateCells: {
      start: {
        sheetId: sheet.getSheetId(),
        rowIndex: rowIndex,
        columnIndex: 0
      },
      rows: [{ values: cells_(values) }],
      fields: 'userEnteredValue'
    }
  };
}
