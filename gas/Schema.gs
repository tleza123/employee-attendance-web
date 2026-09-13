/**
 * DE TEAM — ระบบเช็คชื่อพนักงานและคำนวณยอดค่าจ้าง
 * Schema.gs: การเริ่มต้นระบบ, ตรวจสอบและสร้างโครงสร้าง Google Sheets แบบ Idempotent
 */

function ensureSheetWithHeaders_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    // Format header row to make it clear
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    return sheet;
  }
  // If already exists, check headers
  var data = sheet.getDataRange().getValues();
  if (data.length === 0 || data[0].length === 0 || (data.length === 1 && data[0].every(function(c) { return c === ''; }))) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return sheet;
  }
  requireValue_(
    data[0].length === headers.length && headers.every(function (h, i) { return data[0][i] === h; }),
    'SCHEMA_MISMATCH: ' + name
  );
  return sheet;
}

function initMonthTables_(ss, month) {
  monthKey_(month);
  var tag = suffix_(month);

  ensureSheetWithHeaders_(ss, 'Attendance_' + tag, ATTENDANCE_HEADERS_);
  ensureSheetWithHeaders_(ss, 'MonthlyExtras_' + tag, MONTHLY_EXTRA_HEADERS_);
  ensureSheetWithHeaders_(ss, 'ExtraReviews_' + tag, EXTRA_REVIEW_HEADERS_);
  ensureSheetWithHeaders_(ss, 'Snapshots_' + tag, SNAPSHOT_HEADERS_);
  ensureSheetWithHeaders_(ss, 'Audit_' + tag, AUDIT_HEADERS_);
  ensureSheetWithHeaders_(ss, 'Requests_' + tag, REQUEST_HEADERS_);

  // ตรวจสอบและเพิ่มแถวใน MonthState ถ้ายังไม่มี
  var monthStateTable = table_(ss, 'MonthState', MONTH_HEADERS_);
  var stateRow = unique_(monthStateTable.rows, 'monthKey', month);
  if (!stateRow) {
    monthStateTable.sheet.appendRow([month, 'OPEN', 1, 0, '', '']);
  }
}

function setupSystem_(optionalSpreadsheetId) {
  var ss = optionalSpreadsheetId ? SpreadsheetApp.openById(optionalSpreadsheetId) : spreadsheet_();
  var now = bangkokNowIso_();
  var today = bangkokToday_();
  var currentMonth = today.slice(0, 7);

  // 1. Settings
  var settingsSheet = ensureSheetWithHeaders_(ss, 'Settings', SETTINGS_HEADERS_);
  var settingsTable = table_(ss, 'Settings', SETTINGS_HEADERS_);
  var defaults = [
    { key: 'schemaVersion', value: SCHEMA_VERSION_ },
    { key: 'shopName', value: DEFAULT_SHOP_NAME_ },
    { key: 'workweekJson', value: DEFAULT_WORKWEEK_JSON_ },
    { key: 'calendarEffectiveFrom', value: today }
  ];

  defaults.forEach(function (def) {
    var existing = unique_(settingsTable.rows, 'key', def.key);
    if (!existing) {
      settingsSheet.appendRow([def.key, def.value, 1, now]);
    }
  });

  // 2. Master Tables
  ensureSheetWithHeaders_(ss, 'Employees', EMPLOYEE_HEADERS_);
  ensureSheetWithHeaders_(ss, 'RateHistory', RATE_HEADERS_);
  ensureSheetWithHeaders_(ss, 'ExtraTemplates', EXTRA_TEMPLATE_HEADERS_);
  ensureSheetWithHeaders_(ss, 'Photos', PHOTO_HEADERS_);
  ensureSheetWithHeaders_(ss, 'WorkCalendar', CALENDAR_HEADERS_);
  ensureSheetWithHeaders_(ss, 'MonthState', MONTH_HEADERS_);

  // 3. Current Month Tables
  initMonthTables_(ss, currentMonth);

  // Remove default "Sheet1" or "แผ่นงาน1" if present and other sheets exist
  var defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('แผ่นงาน1');
  if (defaultSheet && ss.getSheets().length > 1) {
    try {
      ss.deleteSheet(defaultSheet);
    } catch (e) {
      // Ignore if cannot delete
    }
  }

  return { ok: true, message: 'SETUP_COMPLETED', currentMonth: currentMonth };
}
