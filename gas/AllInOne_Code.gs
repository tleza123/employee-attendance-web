/**
 * DE TEAM — ระบบเช็คชื่อพนักงานและคำนวณยอดค่าจ้าง (Single-File Complete System)
 * รวมโค้ดระบบหลังบ้าน + หน้าเว็บ HTML/CSS/JS + สคริปต์สร้างฐานข้อมูล ไว้ในไฟล์ Code.gs ไฟล์เดียว!
 * 
 * Spreadsheet ID: 1D9gYkF6uw9UdDDd4mrTBd3W7JU_kRtIawlqwLfce8bM
 */


// ==========================================
// Module: Config.gs
// ==========================================
/**
 * DE TEAM — ระบบเช็คชื่อพนักงานและคำนวณยอดค่าจ้าง
 * Config.gs: ค่าคงที่, โครงสร้าง Headers และการเข้าถึง Script Properties
 */

var SETTINGS_HEADERS_ = ['key', 'value', 'revision', 'updatedAt'];
var EMPLOYEE_HEADERS_ = ['employeeId', 'name', 'nickname', 'position', 'startDate', 'endDate', 'notes', 'photoVersion', 'revision', 'updatedAt'];
var RATE_HEADERS_ = ['rateId', 'employeeId', 'effectiveFrom', 'dailySatang', 'revision', 'updatedAt'];
var EXTRA_TEMPLATE_HEADERS_ = ['templateId', 'employeeId', 'label', 'amountSatang', 'effectiveFromMonth', 'effectiveToMonth', 'revision', 'updatedAt'];
var PHOTO_HEADERS_ = ['employeeId', 'jpegBase64', 'width', 'height', 'revision', 'updatedAt'];
var CALENDAR_HEADERS_ = ['dateKey', 'kind', 'note', 'revision', 'updatedAt'];

var ATTENDANCE_HEADERS_ = ['key', 'dateKey', 'employeeId', 'status', 'revision', 'updatedAt', 'updatedBy', 'requestId'];
var MONTHLY_EXTRA_HEADERS_ = ['extraId', 'employeeId', 'label', 'amountSatang', 'sourceTemplateId', 'sourceTemplateVersion', 'revision', 'updatedAt'];
var EXTRA_REVIEW_HEADERS_ = ['employeeId', 'extrasRevision', 'confirmedAt', 'confirmedBy'];
var MONTH_HEADERS_ = ['monthKey', 'state', 'revision', 'snapshotVersion', 'closedAt', 'closedBy'];
var SNAPSHOT_HEADERS_ = ['snapshotVersion', 'employeeId', 'jsonPartIndex', 'jsonPartCount', 'snapshotJsonPart'];
var AUDIT_HEADERS_ = ['auditId', 'entityKey', 'beforeJson', 'afterJson', 'updatedAt', 'updatedBy', 'requestId'];
var REQUEST_HEADERS_ = ['requestId', 'fingerprint', 'responseJson', 'updatedAt'];

var DEFAULT_WORKWEEK_JSON_ = '[1,2,3,4,5,6]'; // จันทร์-เสาร์ (Sunday=0)
var DEFAULT_SHOP_NAME_ = 'DE TEAM';
var SCHEMA_VERSION_ = '1';
var MAX_SNAPSHOT_PART_LEN_ = 30000;
var MAX_PHOTO_BASE64_LEN_ = 32768;
var MAX_PHOTO_BYTES_ = 24576; // 24 KiB

var DEFAULT_SPREADSHEET_ID_ = '1D9gYkF6uw9UdDDd4mrTBd3W7JU_kRtIawlqwLfce8bM';
var DEFAULT_OWNER_EMAIL_ = 'tlextle23@gmail.com';

function getConfigProperty_(key) {
  var props = PropertiesService.getScriptProperties();
  var val = props ? props.getProperty(key) : null;
  if (!val) {
    if (key === 'SPREADSHEET_ID') return DEFAULT_SPREADSHEET_ID_;
    if (key === 'OWNER_EMAIL') return DEFAULT_OWNER_EMAIL_;
    if (key === 'ADMIN_KEY') return 'DE_TEAM_SECURE_ADMIN_KEY';
  }
  return val;
}

function setConfigProperty_(key, value) {
  var props = PropertiesService.getScriptProperties();
  if (props) {
    props.setProperty(key, value);
  }
}


// ==========================================
// Module: Repository.gs
// ==========================================
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
      var val = values[i];
      if (val instanceof Date && !isNaN(val.getTime())) {
        if (/Month/.test(h) || h === 'month' || h === 'monthKey') {
          val = Utilities.formatDate(val, 'Asia/Bangkok', 'yyyy-MM');
        } else if (/Date/.test(h) || h === 'date' || h === 'dateKey' || h === 'effectiveFrom') {
          val = Utilities.formatDate(val, 'Asia/Bangkok', 'yyyy-MM-dd');
        }
      }
      obj[h] = val;
    });
    rows.push(obj);
  });
  return { sheet: sheet, rows: rows };
}

function unique_(rows, field, value) {
  var target = value;
  if (target instanceof Date && !isNaN(target.getTime())) {
    target = Utilities.formatDate(target, 'Asia/Bangkok', 'yyyy-MM');
  }
  var found = rows.filter(function (r) {
    var itemVal = r[field];
    if (itemVal instanceof Date && !isNaN(itemVal.getTime())) {
      itemVal = Utilities.formatDate(itemVal, 'Asia/Bangkok', 'yyyy-MM');
    }
    return itemVal === target;
  });
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


// ==========================================
// Module: Auth.gs
// ==========================================
/**
 * DE TEAM — ระบบเช็คชื่อพนักงานและคำนวณยอดค่าจ้าง
 * Auth.gs: การตรวจสอบสิทธิ์การใช้งานสำหรับผู้ดูแลระบบคนเดียว (Owner Only)
 */

function owner_() {
  var configured = getConfigProperty_('OWNER_EMAIL');
  var activeUser = Session.getActiveUser();
  var email = activeUser ? activeUser.getEmail() : '';

  requireValue_(configured && typeof configured === 'string' && configured.trim().length > 0, 'CONFIG_REQUIRED');
  requireValue_(email && typeof email === 'string' && email.trim().length > 0, 'AUTH_REQUIRED');
  requireValue_(email.trim().toLowerCase() === configured.trim().toLowerCase(), 'AUTH_REQUIRED');

  return email.trim().toLowerCase();
}


// ==========================================
// Module: Schema.gs
// ==========================================
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


// ==========================================
// Module: Payroll.gs
// ==========================================
/**
 * DE TEAM — ระบบเช็คชื่อพนักงานและคำนวณยอดค่าจ้าง
 * Payroll.gs: Pure Calculation Engine คำนวณค่าจ้าง, ปัดเศษ, วันทำงาน, ตรวจสอบความถูกต้องของข้อมูล
 */

function requireValue_(condition, code) {
  if (!condition) throw new Error(code || 'INVALID_INPUT');
}

function dateKey_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    value = Utilities.formatDate(value, 'Asia/Bangkok', 'yyyy-MM-dd');
  } else if (value && typeof value !== 'string') {
    value = String(value);
  }
  if (typeof value === 'string') {
    value = value.trim();
  }
  requireValue_(typeof value === 'string' && /^20\d{2}-\d{2}-\d{2}$/.test(value), 'INVALID_DATE');
  var parsed = new Date(value + 'T00:00:00Z');
  requireValue_(Number.isFinite(parsed.getTime()) && parsed.toISOString().slice(0, 10) === value, 'INVALID_DATE');
  return value;
}

function monthKey_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    value = Utilities.formatDate(value, 'Asia/Bangkok', 'yyyy-MM');
  } else if (value && typeof value !== 'string') {
    value = String(value);
  }
  if (typeof value === 'string') {
    value = value.trim();
  }
  requireValue_(typeof value === 'string' && /^20\d{2}-(0[1-9]|1[0-2])$/.test(value), 'INVALID_MONTH');
  return value;
}

function moneySatang_(text) {
  requireValue_(typeof text === 'string' && /^(0|[1-9]\d{0,6})(\.\d{1,2})?$/.test(text), 'INVALID_MONEY');
  var parts = text.split('.');
  var amount = Number(parts[0]) * 100 + Number(((parts[1] || '') + '00').slice(0, 2));
  requireValue_(Number.isSafeInteger(amount) && amount <= 100000000, 'INVALID_MONEY');
  return amount;
}

function validateSatang_(amount) {
  requireValue_(Number.isSafeInteger(amount) && amount >= 0 && amount <= 100000000, 'INVALID_MONEY');
  return amount;
}

function monthDates_(month) {
  monthKey_(month);
  var cursor = new Date(month + '-01T00:00:00Z');
  var result = [];
  while (cursor.toISOString().slice(0, 7) === month) {
    result.push(cursor.toISOString().slice(0, 10));
    cursor.setUTCDate(cursor.getUTCDate() + 1);
  }
  return result;
}

function employedOn_(employee, date) {
  return date >= employee.startDate && (!employee.endDate || date <= employee.endDate);
}

function validateEmployeeDates_(employee) {
  requireValue_(employee && typeof employee.employeeId === 'string' && employee.employeeId.length > 0, 'INVALID_EMPLOYEE');
  dateKey_(employee.startDate);
  if (employee.endDate) {
    requireValue_(dateKey_(employee.endDate) >= employee.startDate, 'INVALID_EMPLOYMENT_DATES');
  }
}

function isWorkday_(date, weekdays, overrides) {
  dateKey_(date);
  requireValue_(
    Array.isArray(weekdays) &&
    weekdays.every(function (d) { return Number.isInteger(d) && d >= 0 && d <= 6; }) &&
    new Set(weekdays).size === weekdays.length,
    'INVALID_CALENDAR'
  );
  var kind = overrides && overrides[date];
  if (kind !== undefined) {
    requireValue_(kind === 'WORKDAY' || kind === 'HOLIDAY', 'INVALID_CALENDAR');
    return kind === 'WORKDAY';
  }
  return weekdays.indexOf(new Date(date + 'T00:00:00Z').getUTCDay()) >= 0;
}

function ratesFor_(rates, employeeId) {
  var seen = new Set();
  return rates.filter(function (r) { return r.employeeId === employeeId; }).map(function (r) {
    dateKey_(r.effectiveFrom);
    validateSatang_(r.dailySatang);
    requireValue_(!seen.has(r.effectiveFrom), 'DUPLICATE_RATE');
    seen.add(r.effectiveFrom);
    return r;
  }).sort(function (a, b) { return a.effectiveFrom.localeCompare(b.effectiveFrom); });
}

function rateOn_(sortedRates, date) {
  var selected = null;
  sortedRates.forEach(function (r) {
    if (r.effectiveFrom <= date) selected = r;
  });
  requireValue_(selected !== null, 'MISSING_RATE');
  return selected.dailySatang;
}

function dailyPay_(status, rate) {
  validateSatang_(rate);
  requireValue_(['FULL', 'HALF', 'ABSENT'].indexOf(status) >= 0, 'INVALID_STATUS');
  // กติกาอ้างอิง: ปัดครึ่งสตางค์ขึ้นต่อวัน (เช่น 500.01 บาท ครึ่งวัน = 250.01 บาท)
  return status === 'FULL' ? rate : status === 'HALF' ? Math.floor((rate + 1) / 2) : 0;
}

/**
 * คำนวณสรุปยอดรายเดือนของพนักงาน 1 คน
 * input: month, today, employee, rates[], attendance[], extras[], weekdays[], calendar{}, systemStartDate
 */
function calculateEmployeeMonth_(input) {
  monthKey_(input.month);
  dateKey_(input.today);
  dateKey_(input.systemStartDate);
  validateEmployeeDates_(input.employee);
  var employee = input.employee;
  var rates = ratesFor_(input.rates, employee.employeeId);
  var attendance = new Map();

  input.attendance.filter(function (r) {
    return r.employeeId === employee.employeeId && r.dateKey.slice(0, 7) === input.month;
  }).forEach(function (r) {
    dateKey_(r.dateKey);
    requireValue_(['FULL', 'HALF', 'ABSENT', 'UNMARKED'].indexOf(r.status) >= 0, 'INVALID_STATUS');
    requireValue_(!attendance.has(r.dateKey), 'DUPLICATE_ATTENDANCE');
    requireValue_(employedOn_(employee, r.dateKey), 'OUTSIDE_EMPLOYMENT');
    requireValue_(r.dateKey >= input.systemStartDate, 'BEFORE_SYSTEM_START');
    requireValue_(r.status === 'UNMARKED' || r.dateKey <= input.today, 'FUTURE_ATTENDANCE');
    requireValue_(r.status === 'UNMARKED' || isWorkday_(r.dateKey, input.weekdays, input.calendar), 'ATTENDANCE_ON_HOLIDAY');
    attendance.set(r.dateKey, r);
  });

  var result = {
    employeeId: employee.employeeId,
    month: input.month,
    full: 0,
    half: 0,
    absent: 0,
    pending: 0,
    workedDays: 0,
    paidDayUnits: 0,
    baseSatang: 0,
    extraSatang: 0,
    totalSatang: 0,
    days: [],
    extras: []
  };

  monthDates_(input.month).forEach(function (date) {
    if (!employedOn_(employee, date) || date < input.systemStartDate || date > input.today) return;
    var workday = isWorkday_(date, input.weekdays, input.calendar);
    var row = attendance.get(date);
    var status = row ? row.status : 'UNMARKED';

    if (!workday) {
      result.days.push({ dateKey: date, status: 'HOLIDAY', amountSatang: null });
      return;
    }

    if (status === 'UNMARKED') {
      result.pending += 1;
      result.days.push({ dateKey: date, status: 'UNMARKED', amountSatang: null });
      return;
    }

    var rate = rateOn_(rates, date);
    var amount = dailyPay_(status, rate);

    if (status === 'FULL') result.full += 1;
    if (status === 'HALF') result.half += 1;
    if (status === 'ABSENT') result.absent += 1;

    result.baseSatang += amount;
    result.days.push({ dateKey: date, status: status, dailySatang: rate, amountSatang: amount });
  });

  var extraIds = new Set();
  input.extras.filter(function (x) {
    return x.employeeId === employee.employeeId && x.monthKey === input.month;
  }).forEach(function (x) {
    requireValue_(typeof x.extraId === 'string' && x.extraId.length > 0 && !extraIds.has(x.extraId), 'DUPLICATE_EXTRA');
    requireValue_(typeof x.label === 'string' && x.label.trim().length > 0 && x.label.length <= 80, 'INVALID_EXTRA');
    extraIds.add(x.extraId);
    result.extraSatang += validateSatang_(x.amountSatang);
    result.extras.push({ extraId: x.extraId, label: x.label, amountSatang: x.amountSatang });
  });

  result.workedDays = result.full + result.half;
  result.paidDayUnits = result.full + result.half / 2;
  result.totalSatang = result.baseSatang + result.extraSatang;
  requireValue_(Number.isSafeInteger(result.totalSatang), 'MONEY_OVERFLOW');

  return result;
}


// ==========================================
// Module: AttendanceService.gs
// ==========================================
/**
 * DE TEAM — ระบบเช็คชื่อพนักงานและคำนวณยอดค่าจ้าง
 * AttendanceService.gs: บริการเช็คชื่อพนักงาน (getDay, saveAttendance)
 */

function monthState_(ss, month) {
  var stateTable = table_(ss, 'MonthState', MONTH_HEADERS_);
  var row = unique_(stateTable.rows, 'monthKey', month);
  if (!row) {
    // ถ้ายังไม่มีสถานะเดือน ให้สร้างเป็น OPEN
    initMonthTables_(ss, month);
    stateTable = table_(ss, 'MonthState', MONTH_HEADERS_);
    row = unique_(stateTable.rows, 'monthKey', month);
  }
  requireValue_(row && ['OPEN', 'CLOSED'].indexOf(row.state) >= 0, 'MONTH_NOT_INITIALIZED');
  return row;
}

function normalizeAttendance_(payload) {
  requireValue_(payload && typeof payload === 'object' && !Array.isArray(payload), 'INVALID_INPUT');
  var fields = ['dateKey', 'employeeId', 'status', 'expectedRevision', 'requestId'];
  requireValue_(
    Object.keys(payload).length === fields.length &&
    Object.keys(payload).every(function (k) { return fields.indexOf(k) >= 0; }),
    'INVALID_INPUT'
  );
  dateKey_(payload.dateKey);
  requireValue_(typeof payload.employeeId === 'string' && /^[a-zA-Z0-9-]{1,64}$/.test(payload.employeeId), 'INVALID_EMPLOYEE');
  requireValue_(['FULL', 'HALF', 'ABSENT', 'UNMARKED'].indexOf(payload.status) >= 0, 'INVALID_STATUS');
  requireValue_(Number.isSafeInteger(payload.expectedRevision) && payload.expectedRevision >= 0, 'INVALID_REVISION');
  requireValue_(
    typeof payload.requestId === 'string' &&
    /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(payload.requestId),
    'INVALID_REQUEST_ID'
  );
  return {
    dateKey: payload.dateKey,
    employeeId: payload.employeeId,
    status: payload.status,
    expectedRevision: payload.expectedRevision,
    requestId: payload.requestId
  };
}

function publicAttendanceRow_(row) {
  var value = {};
  ATTENDANCE_HEADERS_.forEach(function (h) { value[h] = row[h]; });
  return value;
}

function getDay_(ss, date) {
  dateKey_(date);
  var month = date.slice(0, 7);
  var monthState = monthState_(ss, month);
  var calendar = calendar_(ss);

  // ตรวจสอบและสร้างชีตประจำเดือนหากยังไม่มี
  var tag = suffix_(month);
  var attSheetName = 'Attendance_' + tag;
  if (!ss.getSheetByName(attSheetName)) {
    initMonthTables_(ss, month);
  }

  var dayRows = table_(ss, attSheetName, ATTENDANCE_HEADERS_).rows.filter(function (r) {
    return r.dateKey === date;
  });

  var seen = new Set();
  dayRows.forEach(function (r) {
    requireValue_(r.key === date + '|' + r.employeeId && !seen.has(r.key), 'DUPLICATE_KEY');
    requireValue_(
      ['FULL', 'HALF', 'ABSENT', 'UNMARKED'].indexOf(r.status) >= 0 &&
      Number.isSafeInteger(r.revision) && r.revision > 0,
      'INVALID_DATA'
    );
    seen.add(r.key);
  });

  var people = table_(ss, 'Employees', EMPLOYEE_HEADERS_).rows;
  var ids = new Set();
  people.forEach(function (e) {
    validateEmployeeDates_(e);
    requireValue_(!ids.has(e.employeeId), 'DUPLICATE_KEY');
    ids.add(e.employeeId);
  });
  requireValue_(dayRows.every(function (r) { return ids.has(r.employeeId); }), 'ORPHAN_ATTENDANCE');

  return {
    dateKey: date,
    serverToday: bangkokToday_(),
    isClosed: monthState.state === 'CLOSED',
    isWorkday: date >= calendar.startDate && isWorkday_(date, calendar.weekdays, calendar.overrides),
    employees: people.filter(function (e) { return employedOn_(e, date); }).map(function (e) {
      return {
        employeeId: e.employeeId,
        name: e.name,
        nickname: e.nickname,
        position: e.position,
        photoVersion: e.photoVersion
      };
    }),
    attendance: dayRows.map(publicAttendanceRow_)
  };
}

function getDay(date) {
  owner_();
  dateKey_(date);
  return locked_(function () {
    var ss = spreadsheet_();
    return getDay_(ss, date);
  });
}

function saveAttendance(payload) {
  var actor = owner_();
  var data = normalizeAttendance_(payload);

  return locked_(function () {
    var ss = spreadsheet_();
    var month = data.dateKey.slice(0, 7);
    var tag = suffix_(month);

    if (!ss.getSheetByName('Attendance_' + tag)) {
      initMonthTables_(ss, month);
    }

    var fingerprint = fingerprint_(data);
    var requests = table_(ss, 'Requests_' + tag, REQUEST_HEADERS_);
    var receipt = unique_(requests.rows, 'requestId', data.requestId);
    if (receipt) {
      requireValue_(receipt.fingerprint === fingerprint, 'REQUEST_ID_REUSED');
      return JSON.parse(receipt.responseJson);
    }

    requireValue_(monthState_(ss, month).state === 'OPEN', 'MONTH_CLOSED');
    requireValue_(data.dateKey <= bangkokToday_(), 'FUTURE_ATTENDANCE');

    var employee = unique_(table_(ss, 'Employees', EMPLOYEE_HEADERS_).rows, 'employeeId', data.employeeId);
    requireValue_(employee, 'EMPLOYEE_NOT_FOUND');
    validateEmployeeDates_(employee);
    requireValue_(employedOn_(employee, data.dateKey), 'OUTSIDE_EMPLOYMENT');

    var calendar = calendar_(ss);
    requireValue_(data.dateKey >= calendar.startDate, 'BEFORE_SYSTEM_START');
    requireValue_(data.status === 'UNMARKED' || isWorkday_(data.dateKey, calendar.weekdays, calendar.overrides), 'ATTENDANCE_ON_HOLIDAY');

    var attendance = table_(ss, 'Attendance_' + tag, ATTENDANCE_HEADERS_);
    var key = data.dateKey + '|' + data.employeeId;
    var before = unique_(attendance.rows, 'key', key);
    requireValue_(!before || (before.dateKey === data.dateKey && before.employeeId === data.employeeId &&
      Number.isSafeInteger(before.revision) && before.revision > 0), 'INVALID_DATA');

    var revision = before ? before.revision : 0;
    requireValue_(revision === data.expectedRevision, 'CONFLICT');

    var timestamp = bangkokNowIso_();
    var after = {
      key: key,
      dateKey: data.dateKey,
      employeeId: data.employeeId,
      status: data.status,
      revision: revision + 1,
      updatedAt: timestamp,
      updatedBy: actor,
      requestId: data.requestId
    };

    var response = { ok: true, attendance: after };
    var values = ATTENDANCE_HEADERS_.map(function (h) { return after[h]; });
    var change = before
      ? updateRow_(attendance.sheet, before.rowNumber - 1, values)
      : append_(attendance.sheet, values);

    var audit = table_(ss, 'Audit_' + tag, AUDIT_HEADERS_);
    Sheets.Spreadsheets.batchUpdate({
      requests: [
        change,
        append_(audit.sheet, [Utilities.getUuid(), key, JSON.stringify(before ? publicAttendanceRow_(before) : null), JSON.stringify(after), timestamp, actor, data.requestId]),
        append_(requests.sheet, [data.requestId, fingerprint, JSON.stringify(response), timestamp])
      ]
    }, ss.getId());

    return response;
  });
}


// ==========================================
// Module: EmployeeService.gs
// ==========================================
/**
 * DE TEAM — ระบบเช็คชื่อพนักงานและคำนวณยอดค่าจ้าง
 * EmployeeService.gs: บริการจัดการข้อมูลพนักงาน, ประวัติอัตราค่าแรง และการสิ้นสุดการทำงาน
 */

function normalizeEmployeePayload_(payload) {
  requireValue_(payload && typeof payload === 'object', 'INVALID_INPUT');
  requireValue_(typeof payload.requestId === 'string' && /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(payload.requestId), 'INVALID_REQUEST_ID');

  var name = typeof payload.name === 'string' ? payload.name.trim() : '';
  requireValue_(name.length >= 1 && name.length <= 100, 'INVALID_NAME');

  var nickname = typeof payload.nickname === 'string' ? payload.nickname.trim() : '';
  requireValue_(nickname.length <= 40, 'INVALID_NICKNAME');

  var position = typeof payload.position === 'string' ? payload.position.trim() : '';
  requireValue_(position.length >= 1 && position.length <= 80, 'INVALID_POSITION');

  dateKey_(payload.startDate);
  if (payload.endDate) {
    dateKey_(payload.endDate);
    requireValue_(payload.endDate >= payload.startDate, 'INVALID_EMPLOYMENT_DATES');
  }

  var notes = typeof payload.notes === 'string' ? payload.notes.trim() : '';
  requireValue_(notes.length <= 500, 'INVALID_NOTES');

  var dailySatang;
  if (typeof payload.dailySatang === 'number') {
    dailySatang = validateSatang_(payload.dailySatang);
  } else if (typeof payload.dailyWage === 'string') {
    dailySatang = moneySatang_(payload.dailyWage);
  } else {
    throw new Error('INVALID_MONEY');
  }

  var rateEffectiveFrom = payload.rateEffectiveFrom ? dateKey_(payload.rateEffectiveFrom) : payload.startDate;
  requireValue_(rateEffectiveFrom >= payload.startDate, 'INVALID_RATE_DATE');
  if (payload.endDate) {
    requireValue_(rateEffectiveFrom <= payload.endDate, 'INVALID_RATE_DATE');
  }

  return {
    employeeId: payload.employeeId || ('e-' + Utilities.getUuid()),
    name: name,
    nickname: nickname,
    position: position,
    startDate: payload.startDate,
    endDate: payload.endDate || '',
    notes: notes,
    dailySatang: dailySatang,
    rateEffectiveFrom: rateEffectiveFrom,
    extraTemplates: Array.isArray(payload.extraTemplates) ? payload.extraTemplates : [],
    expectedRevision: Number.isSafeInteger(payload.expectedRevision) ? payload.expectedRevision : 0,
    requestId: payload.requestId
  };
}

function saveEmployee(payload) {
  var actor = owner_();
  var data = normalizeEmployeePayload_(payload);

  return locked_(function () {
    var ss = spreadsheet_();
    var today = bangkokToday_();
    var currentMonth = today.slice(0, 7);
    var currentTag = suffix_(currentMonth);

    var fingerprint = fingerprint_(payload);
    var requests = table_(ss, 'Requests_' + currentTag, REQUEST_HEADERS_);
    var receipt = unique_(requests.rows, 'requestId', data.requestId);
    if (receipt) {
      requireValue_(receipt.fingerprint === fingerprint, 'REQUEST_ID_REUSED');
      return JSON.parse(receipt.responseJson);
    }

    var empTable = table_(ss, 'Employees', EMPLOYEE_HEADERS_);
    var existing = unique_(empTable.rows, 'employeeId', data.employeeId);
    var revision = existing ? existing.revision : 0;
    requireValue_(revision === data.expectedRevision, 'CONFLICT');

    // ถ้าตั้ง endDate ให้ตรวจว่ามี attendance หลังจากนั้นหรือไม่
    if (data.endDate) {
      checkNoAttendanceAfterEnd_(ss, data.employeeId, data.endDate);
    }

    var timestamp = bangkokNowIso_();
    var afterEmployee = {
      employeeId: data.employeeId,
      name: data.name,
      nickname: data.nickname,
      position: data.position,
      startDate: data.startDate,
      endDate: data.endDate,
      notes: data.notes,
      photoVersion: existing ? (existing.photoVersion || 0) : 0,
      revision: revision + 1,
      updatedAt: timestamp
    };

    var batchRequests = [];
    var empValues = EMPLOYEE_HEADERS_.map(function (h) { return afterEmployee[h]; });
    var empChange = existing
      ? updateRow_(empTable.sheet, existing.rowNumber - 1, empValues)
      : append_(empTable.sheet, empValues);
    batchRequests.push(empChange);

    // จัดการ RateHistory
    var rateTable = table_(ss, 'RateHistory', RATE_HEADERS_);
    var empRates = rateTable.rows.filter(function (r) { return r.employeeId === data.employeeId; });
    var sameDateRate = unique_(empRates, 'effectiveFrom', data.rateEffectiveFrom);

    if (sameDateRate) {
      // มีอัตราในวันที่เดียวกันแล้ว ตรวจสอบว่าจำนวนเงินตรงกันหรือไม่
      if (sameDateRate.dailySatang !== data.dailySatang) {
        // อัปเดตอัตราเดิมในวันเดียวกัน
        var updatedRate = {
          rateId: sameDateRate.rateId,
          employeeId: data.employeeId,
          effectiveFrom: data.rateEffectiveFrom,
          dailySatang: data.dailySatang,
          revision: sameDateRate.revision + 1,
          updatedAt: timestamp
        };
        batchRequests.push(updateRow_(rateTable.sheet, sameDateRate.rowNumber - 1, RATE_HEADERS_.map(function (h) { return updatedRate[h]; })));
      }
    } else {
      // เพิ่มแถวใหม่ใน RateHistory
      var newRate = {
        rateId: 'r-' + Utilities.getUuid(),
        employeeId: data.employeeId,
        effectiveFrom: data.rateEffectiveFrom,
        dailySatang: data.dailySatang,
        revision: 1,
        updatedAt: timestamp
      };
      batchRequests.push(append_(rateTable.sheet, RATE_HEADERS_.map(function (h) { return newRate[h]; })));
    }

    // จัดการ ExtraTemplates
    if (data.extraTemplates && data.extraTemplates.length > 0) {
      var templateTable = table_(ss, 'ExtraTemplates', EXTRA_TEMPLATE_HEADERS_);
      data.extraTemplates.forEach(function (tpl) {
        requireValue_(typeof tpl.label === 'string' && tpl.label.trim().length > 0, 'INVALID_EXTRA');
        var amountSatang = typeof tpl.amountSatang === 'number' ? validateSatang_(tpl.amountSatang) : moneySatang_(tpl.amount);
        var tplId = tpl.templateId || ('xt-' + Utilities.getUuid());
        var existingTpl = unique_(templateTable.rows, 'templateId', tplId);
        var fromMonth = tpl.effectiveFromMonth ? monthKey_(tpl.effectiveFromMonth) : currentMonth;
        var toMonth = tpl.effectiveToMonth ? monthKey_(tpl.effectiveToMonth) : '';

        var tplRecord = {
          templateId: tplId,
          employeeId: data.employeeId,
          label: tpl.label.trim(),
          amountSatang: amountSatang,
          effectiveFromMonth: fromMonth,
          effectiveToMonth: toMonth,
          revision: existingTpl ? existingTpl.revision + 1 : 1,
          updatedAt: timestamp
        };

        var tplValues = EXTRA_TEMPLATE_HEADERS_.map(function (h) { return tplRecord[h]; });
        if (existingTpl) {
          batchRequests.push(updateRow_(templateTable.sheet, existingTpl.rowNumber - 1, tplValues));
        } else {
          batchRequests.push(append_(templateTable.sheet, tplValues));
        }
      });
    }

    var audit = table_(ss, 'Audit_' + currentTag, AUDIT_HEADERS_);
    var response = { ok: true, employee: afterEmployee };

    batchRequests.push(append_(audit.sheet, [
      Utilities.getUuid(),
      'EMPLOYEE|' + data.employeeId,
      JSON.stringify(existing || null),
      JSON.stringify(afterEmployee),
      timestamp,
      actor,
      data.requestId
    ]));
    batchRequests.push(append_(requests.sheet, [
      data.requestId,
      fingerprint,
      JSON.stringify(response),
      timestamp
    ]));

    Sheets.Spreadsheets.batchUpdate({ requests: batchRequests }, ss.getId());

    return response;
  });
}

function checkNoAttendanceAfterEnd_(ss, employeeId, endDate) {
  var endMonth = endDate.slice(0, 7);
  var allSheets = ss.getSheets();
  allSheets.forEach(function (sheet) {
    var name = sheet.getName();
    if (name.indexOf('Attendance_') === 0) {
      var monthPart = name.slice('Attendance_'.length).replace('_', '-');
      if (monthPart >= endMonth) {
        var data = sheet.getDataRange().getValues();
        if (data.length > 1) {
          data.slice(1).forEach(function (row) {
            var rowEmpId = row[2]; // employeeId
            var rowDate = row[1]; // dateKey
            var rowStatus = row[3]; // status
            if (rowEmpId === employeeId && rowDate > endDate && rowStatus && rowStatus !== 'UNMARKED') {
              throw new Error('ATTENDANCE_AFTER_END_DATE: ' + rowDate + ' (' + rowStatus + ')');
            }
          });
        }
      }
    }
  });
}

function setEmploymentEnd(payload) {
  var actor = owner_();
  requireValue_(payload && typeof payload === 'object', 'INVALID_INPUT');
  requireValue_(typeof payload.employeeId === 'string' && payload.employeeId.length > 0, 'INVALID_EMPLOYEE');
  dateKey_(payload.endDate);
  requireValue_(typeof payload.requestId === 'string' && /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(payload.requestId), 'INVALID_REQUEST_ID');
  requireValue_(Number.isSafeInteger(payload.expectedRevision), 'INVALID_REVISION');

  return locked_(function () {
    var ss = spreadsheet_();
    var currentMonth = bangkokToday_().slice(0, 7);
    var currentTag = suffix_(currentMonth);

    var fingerprint = fingerprint_(payload);
    var requests = table_(ss, 'Requests_' + currentTag, REQUEST_HEADERS_);
    var receipt = unique_(requests.rows, 'requestId', payload.requestId);
    if (receipt) {
      requireValue_(receipt.fingerprint === fingerprint, 'REQUEST_ID_REUSED');
      return JSON.parse(receipt.responseJson);
    }

    var empTable = table_(ss, 'Employees', EMPLOYEE_HEADERS_);
    var employee = unique_(empTable.rows, 'employeeId', payload.employeeId);
    requireValue_(employee, 'EMPLOYEE_NOT_FOUND');
    requireValue_(employee.revision === payload.expectedRevision, 'CONFLICT');
    requireValue_(payload.endDate >= employee.startDate, 'INVALID_EMPLOYMENT_DATES');

    // ตรวจสอบความขัดแย้งของ attendance หลัง endDate
    checkNoAttendanceAfterEnd_(ss, payload.employeeId, payload.endDate);

    var timestamp = bangkokNowIso_();
    var afterEmployee = {
      employeeId: employee.employeeId,
      name: employee.name,
      nickname: employee.nickname,
      position: employee.position,
      startDate: employee.startDate,
      endDate: payload.endDate,
      notes: employee.notes,
      photoVersion: employee.photoVersion,
      revision: employee.revision + 1,
      updatedAt: timestamp
    };

    var change = updateRow_(empTable.sheet, employee.rowNumber - 1, EMPLOYEE_HEADERS_.map(function (h) { return afterEmployee[h]; }));
    var audit = table_(ss, 'Audit_' + currentTag, AUDIT_HEADERS_);
    var response = { ok: true, employee: afterEmployee };

    Sheets.Spreadsheets.batchUpdate({
      requests: [
        change,
        append_(audit.sheet, [Utilities.getUuid(), 'EMPLOYEE|' + payload.employeeId, JSON.stringify(employee), JSON.stringify(afterEmployee), timestamp, actor, payload.requestId]),
        append_(requests.sheet, [payload.requestId, fingerprint, JSON.stringify(response), timestamp])
      ]
    }, ss.getId());

    return response;
  });
}

function addRate(payload) {
  var actor = owner_();
  requireValue_(payload && typeof payload === 'object', 'INVALID_INPUT');
  requireValue_(typeof payload.employeeId === 'string' && payload.employeeId.length > 0, 'INVALID_EMPLOYEE');
  dateKey_(payload.effectiveFrom);
  var dailySatang = typeof payload.dailySatang === 'number' ? validateSatang_(payload.dailySatang) : moneySatang_(payload.dailyWage);
  requireValue_(typeof payload.requestId === 'string' && /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(payload.requestId), 'INVALID_REQUEST_ID');

  return locked_(function () {
    var ss = spreadsheet_();
    var effMonth = payload.effectiveFrom.slice(0, 7);
    var effTag = suffix_(effMonth);

    // ตรวจสอบว่าเดือนของ effectiveFrom ปิดแล้วหรือไม่
    var mState = monthState_(ss, effMonth);
    requireValue_(mState.state === 'OPEN', 'MONTH_CLOSED');

    var fingerprint = fingerprint_(payload);
    var requests = table_(ss, 'Requests_' + effTag, REQUEST_HEADERS_);
    var receipt = unique_(requests.rows, 'requestId', payload.requestId);
    if (receipt) {
      requireValue_(receipt.fingerprint === fingerprint, 'REQUEST_ID_REUSED');
      return JSON.parse(receipt.responseJson);
    }

    var empTable = table_(ss, 'Employees', EMPLOYEE_HEADERS_);
    var employee = unique_(empTable.rows, 'employeeId', payload.employeeId);
    requireValue_(employee, 'EMPLOYEE_NOT_FOUND');
    requireValue_(payload.effectiveFrom >= employee.startDate, 'INVALID_RATE_DATE');
    if (employee.endDate) {
      requireValue_(payload.effectiveFrom <= employee.endDate, 'INVALID_RATE_DATE');
    }

    var rateTable = table_(ss, 'RateHistory', RATE_HEADERS_);
    var empRates = rateTable.rows.filter(function (r) { return r.employeeId === payload.employeeId; });
    var duplicateRate = unique_(empRates, 'effectiveFrom', payload.effectiveFrom);
    requireValue_(!duplicateRate, 'DUPLICATE_RATE');

    var timestamp = bangkokNowIso_();
    var newRate = {
      rateId: 'r-' + Utilities.getUuid(),
      employeeId: payload.employeeId,
      effectiveFrom: payload.effectiveFrom,
      dailySatang: dailySatang,
      revision: 1,
      updatedAt: timestamp
    };

    var change = append_(rateTable.sheet, RATE_HEADERS_.map(function (h) { return newRate[h]; }));
    var audit = table_(ss, 'Audit_' + effTag, AUDIT_HEADERS_);
    var response = { ok: true, rate: newRate };

    Sheets.Spreadsheets.batchUpdate({
      requests: [
        change,
        append_(audit.sheet, [Utilities.getUuid(), 'RATE|' + newRate.rateId, JSON.stringify(null), JSON.stringify(newRate), timestamp, actor, payload.requestId]),
        append_(requests.sheet, [payload.requestId, fingerprint, JSON.stringify(response), timestamp])
      ]
    }, ss.getId());

    return response;
  });
}


// ==========================================
// Module: ExtrasService.gs
// ==========================================
/**
 * DE TEAM — ระบบเช็คชื่อพนักงานและคำนวณยอดค่าจ้าง
 * ExtrasService.gs: บริการเงินพิเศษ (Templates ประจำ, เงินพิเศษรายเดือน, การตรวจสอบและยืนยัน)
 */

function saveExtraTemplate(payload) {
  var actor = owner_();
  requireValue_(payload && typeof payload === 'object', 'INVALID_INPUT');
  requireValue_(typeof payload.employeeId === 'string' && payload.employeeId.length > 0, 'INVALID_EMPLOYEE');
  requireValue_(typeof payload.label === 'string' && payload.label.trim().length > 0 && payload.label.length <= 80, 'INVALID_EXTRA');
  var amountSatang = typeof payload.amountSatang === 'number' ? validateSatang_(payload.amountSatang) : moneySatang_(payload.amount);
  requireValue_(typeof payload.requestId === 'string' && /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(payload.requestId), 'INVALID_REQUEST_ID');

  var fromMonth = payload.effectiveFromMonth ? monthKey_(payload.effectiveFromMonth) : bangkokToday_().slice(0, 7);
  var toMonth = payload.effectiveToMonth ? monthKey_(payload.effectiveToMonth) : '';

  return locked_(function () {
    var ss = spreadsheet_();
    var currentMonth = bangkokToday_().slice(0, 7);
    var currentTag = suffix_(currentMonth);

    var fingerprint = fingerprint_(payload);
    var requests = table_(ss, 'Requests_' + currentTag, REQUEST_HEADERS_);
    var receipt = unique_(requests.rows, 'requestId', payload.requestId);
    if (receipt) {
      requireValue_(receipt.fingerprint === fingerprint, 'REQUEST_ID_REUSED');
      return JSON.parse(receipt.responseJson);
    }

    var tplTable = table_(ss, 'ExtraTemplates', EXTRA_TEMPLATE_HEADERS_);
    var tplId = payload.templateId || ('xt-' + Utilities.getUuid());
    var existing = unique_(tplTable.rows, 'templateId', tplId);
    var revision = existing ? existing.revision : 0;
    requireValue_(revision === (payload.expectedRevision || 0), 'CONFLICT');

    var timestamp = bangkokNowIso_();
    var after = {
      templateId: tplId,
      employeeId: payload.employeeId,
      label: payload.label.trim(),
      amountSatang: amountSatang,
      effectiveFromMonth: fromMonth,
      effectiveToMonth: toMonth,
      revision: revision + 1,
      updatedAt: timestamp
    };

    var values = EXTRA_TEMPLATE_HEADERS_.map(function (h) { return after[h]; });
    var change = existing
      ? updateRow_(tplTable.sheet, existing.rowNumber - 1, values)
      : append_(tplTable.sheet, values);

    var audit = table_(ss, 'Audit_' + currentTag, AUDIT_HEADERS_);
    var response = { ok: true, template: after };

    Sheets.Spreadsheets.batchUpdate({
      requests: [
        change,
        append_(audit.sheet, [Utilities.getUuid(), 'EXTRA_TEMPLATE|' + tplId, JSON.stringify(existing || null), JSON.stringify(after), timestamp, actor, payload.requestId]),
        append_(requests.sheet, [payload.requestId, fingerprint, JSON.stringify(response), timestamp])
      ]
    }, ss.getId());

    return response;
  });
}

function prepareMonthExtras(payload) {
  var actor = owner_();
  requireValue_(payload && typeof payload === 'object', 'INVALID_INPUT');
  monthKey_(payload.monthKey);
  requireValue_(typeof payload.requestId === 'string' && /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(payload.requestId), 'INVALID_REQUEST_ID');

  return locked_(function () {
    var ss = spreadsheet_();
    var month = payload.monthKey;
    var tag = suffix_(month);

    if (!ss.getSheetByName('MonthlyExtras_' + tag)) {
      initMonthTables_(ss, month);
    }

    var fingerprint = fingerprint_(payload);
    var requests = table_(ss, 'Requests_' + tag, REQUEST_HEADERS_);
    var receipt = unique_(requests.rows, 'requestId', payload.requestId);
    if (receipt) {
      requireValue_(receipt.fingerprint === fingerprint, 'REQUEST_ID_REUSED');
      return JSON.parse(receipt.responseJson);
    }

    requireValue_(monthState_(ss, month).state === 'OPEN', 'MONTH_CLOSED');

    var tplTable = table_(ss, 'ExtraTemplates', EXTRA_TEMPLATE_HEADERS_);
    var monthlyTable = table_(ss, 'MonthlyExtras_' + tag, MONTHLY_EXTRA_HEADERS_);

    // หา templates ที่มีผลในเดือนนี้
    var applicableTemplates = tplTable.rows.filter(function (tpl) {
      if (tpl.effectiveFromMonth && tpl.effectiveFromMonth > month) return false;
      if (tpl.effectiveToMonth && tpl.effectiveToMonth < month) return false;
      return true;
    });

    var existingSourceMap = new Set();
    monthlyTable.rows.forEach(function (row) {
      if (row.sourceTemplateId) {
        existingSourceMap.add(row.employeeId + '|' + row.sourceTemplateId);
      }
    });

    var timestamp = bangkokNowIso_();
    var batchRequests = [];
    var createdCount = 0;

    applicableTemplates.forEach(function (tpl) {
      var key = tpl.employeeId + '|' + tpl.templateId;
      if (!existingSourceMap.has(key)) {
        var extraRecord = {
          extraId: 'mx-' + Utilities.getUuid(),
          employeeId: tpl.employeeId,
          label: tpl.label,
          amountSatang: tpl.amountSatang,
          sourceTemplateId: tpl.templateId,
          sourceTemplateVersion: tpl.revision,
          revision: 1,
          updatedAt: timestamp
        };
        var values = MONTHLY_EXTRA_HEADERS_.map(function (h) { return extraRecord[h]; });
        batchRequests.push(append_(monthlyTable.sheet, values));
        createdCount++;
      }
    });

    var audit = table_(ss, 'Audit_' + tag, AUDIT_HEADERS_);
    var response = { ok: true, createdCount: createdCount };

    if (batchRequests.length > 0) {
      batchRequests.push(append_(audit.sheet, [
        Utilities.getUuid(),
        'MONTH_EXTRAS|' + month,
        JSON.stringify(null),
        JSON.stringify({ createdCount: createdCount }),
        timestamp,
        actor,
        payload.requestId
      ]));
    }
    batchRequests.push(append_(requests.sheet, [
      payload.requestId,
      fingerprint,
      JSON.stringify(response),
      timestamp
    ]));

    Sheets.Spreadsheets.batchUpdate({ requests: batchRequests }, ss.getId());

    return response;
  });
}

function saveMonthExtra(payload) {
  var actor = owner_();
  requireValue_(payload && typeof payload === 'object', 'INVALID_INPUT');
  monthKey_(payload.monthKey);
  requireValue_(typeof payload.employeeId === 'string' && payload.employeeId.length > 0, 'INVALID_EMPLOYEE');
  requireValue_(typeof payload.requestId === 'string' && /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(payload.requestId), 'INVALID_REQUEST_ID');

  var isDelete = Boolean(payload.isDelete);
  var label = '';
  var amountSatang = 0;

  if (!isDelete) {
    requireValue_(typeof payload.label === 'string' && payload.label.trim().length > 0 && payload.label.length <= 80, 'INVALID_EXTRA');
    label = payload.label.trim();
    amountSatang = typeof payload.amountSatang === 'number' ? validateSatang_(payload.amountSatang) : moneySatang_(payload.amount);
  }

  return locked_(function () {
    var ss = spreadsheet_();
    var month = payload.monthKey;
    var tag = suffix_(month);

    if (!ss.getSheetByName('MonthlyExtras_' + tag)) {
      initMonthTables_(ss, month);
    }

    var fingerprint = fingerprint_(payload);
    var requests = table_(ss, 'Requests_' + tag, REQUEST_HEADERS_);
    var receipt = unique_(requests.rows, 'requestId', payload.requestId);
    if (receipt) {
      requireValue_(receipt.fingerprint === fingerprint, 'REQUEST_ID_REUSED');
      return JSON.parse(receipt.responseJson);
    }

    requireValue_(monthState_(ss, month).state === 'OPEN', 'MONTH_CLOSED');

    var monthlyTable = table_(ss, 'MonthlyExtras_' + tag, MONTHLY_EXTRA_HEADERS_);
    var extraId = payload.extraId || ('mx-' + Utilities.getUuid());
    var existing = unique_(monthlyTable.rows, 'extraId', extraId);
    var revision = existing ? existing.revision : 0;
    requireValue_(revision === (payload.expectedRevision || 0), 'CONFLICT');

    var timestamp = bangkokNowIso_();
    var batchRequests = [];
    var after = null;

    if (isDelete) {
      requireValue_(existing, 'EXTRA_NOT_FOUND');
      // เพื่อคง revision และ audit ในชีต ให้ตั้ง amountSatang = 0 หรือเคลียร์แถวอย่างปลอดภัย
      // หรือหากใช้การลบ ให้เก็บ audit
      after = {
        extraId: extraId,
        employeeId: payload.employeeId,
        label: existing.label,
        amountSatang: 0,
        sourceTemplateId: existing.sourceTemplateId || '',
        sourceTemplateVersion: existing.sourceTemplateVersion || 0,
        revision: revision + 1,
        updatedAt: timestamp
      };
      batchRequests.push(updateRow_(monthlyTable.sheet, existing.rowNumber - 1, MONTHLY_EXTRA_HEADERS_.map(function (h) { return after[h]; })));
    } else {
      after = {
        extraId: extraId,
        employeeId: payload.employeeId,
        label: label,
        amountSatang: amountSatang,
        sourceTemplateId: existing ? existing.sourceTemplateId : '',
        sourceTemplateVersion: existing ? existing.sourceTemplateVersion : 0,
        revision: revision + 1,
        updatedAt: timestamp
      };
      var change = existing
        ? updateRow_(monthlyTable.sheet, existing.rowNumber - 1, MONTHLY_EXTRA_HEADERS_.map(function (h) { return after[h]; }))
        : append_(monthlyTable.sheet, MONTHLY_EXTRA_HEADERS_.map(function (h) { return after[h]; }));
      batchRequests.push(change);
    }

    // เมื่อแก้เงินพิเศษของพนักงานคนนี้ ให้ล้าง review confirmation เดิมเพื่อให้ต้องตรวจใหม่ก่อนปิดเดือน
    var reviewTable = table_(ss, 'ExtraReviews_' + tag, EXTRA_REVIEW_HEADERS_);
    var existingReview = unique_(reviewTable.rows, 'employeeId', payload.employeeId);
    if (existingReview) {
      batchRequests.push(updateRow_(reviewTable.sheet, existingReview.rowNumber - 1, [payload.employeeId, 0, '', '']));
    }

    var audit = table_(ss, 'Audit_' + tag, AUDIT_HEADERS_);
    var response = { ok: true, extra: after };

    batchRequests.push(append_(audit.sheet, [
      Utilities.getUuid(),
      'MONTHLY_EXTRA|' + extraId,
      JSON.stringify(existing || null),
      JSON.stringify(after),
      timestamp,
      actor,
      payload.requestId
    ]));
    batchRequests.push(append_(requests.sheet, [
      payload.requestId,
      fingerprint,
      JSON.stringify(response),
      timestamp
    ]));

    Sheets.Spreadsheets.batchUpdate({ requests: batchRequests }, ss.getId());

    return response;
  });
}

function confirmMonthExtras(payload) {
  var actor = owner_();
  requireValue_(payload && typeof payload === 'object', 'INVALID_INPUT');
  monthKey_(payload.monthKey);
  requireValue_(typeof payload.employeeId === 'string' && payload.employeeId.length > 0, 'INVALID_EMPLOYEE');
  requireValue_(typeof payload.requestId === 'string' && /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(payload.requestId), 'INVALID_REQUEST_ID');

  return locked_(function () {
    var ss = spreadsheet_();
    var month = payload.monthKey;
    var tag = suffix_(month);

    if (!ss.getSheetByName('ExtraReviews_' + tag)) {
      initMonthTables_(ss, month);
    }

    var fingerprint = fingerprint_(payload);
    var requests = table_(ss, 'Requests_' + tag, REQUEST_HEADERS_);
    var receipt = unique_(requests.rows, 'requestId', payload.requestId);
    if (receipt) {
      requireValue_(receipt.fingerprint === fingerprint, 'REQUEST_ID_REUSED');
      return JSON.parse(receipt.responseJson);
    }

    requireValue_(monthState_(ss, month).state === 'OPEN', 'MONTH_CLOSED');

    var reviewTable = table_(ss, 'ExtraReviews_' + tag, EXTRA_REVIEW_HEADERS_);
    var existing = unique_(reviewTable.rows, 'employeeId', payload.employeeId);

    var timestamp = bangkokNowIso_();
    var after = [payload.employeeId, 1, timestamp, actor];
    var change = existing
      ? updateRow_(reviewTable.sheet, existing.rowNumber - 1, after)
      : append_(reviewTable.sheet, after);

    var audit = table_(ss, 'Audit_' + tag, AUDIT_HEADERS_);
    var response = { ok: true, confirmedAt: timestamp, confirmedBy: actor };

    Sheets.Spreadsheets.batchUpdate({
      requests: [
        change,
        append_(audit.sheet, [Utilities.getUuid(), 'EXTRA_REVIEW|' + payload.employeeId, JSON.stringify(existing || null), JSON.stringify({ confirmedAt: timestamp }), timestamp, actor, payload.requestId]),
        append_(requests.sheet, [payload.requestId, fingerprint, JSON.stringify(response), timestamp])
      ]
    }, ss.getId());

    return response;
  });
}


// ==========================================
// Module: CalendarService.gs
// ==========================================
/**
 * DE TEAM — ระบบเช็คชื่อพนักงานและคำนวณยอดค่าจ้าง
 * CalendarService.gs: การอ่านและบันทึกตารางวันทำงานและวันหยุด
 */

function calendar_(ss) {
  var settings = table_(ss, 'Settings', SETTINGS_HEADERS_).rows;
  var week = unique_(settings, 'key', 'workweekJson');
  var start = unique_(settings, 'key', 'calendarEffectiveFrom');
  requireValue_(week && start, 'CALENDAR_REQUIRED');
  dateKey_(start.value);

  var overrides = Object.create(null);
  var calendarRows = table_(ss, 'WorkCalendar', CALENDAR_HEADERS_).rows;
  calendarRows.forEach(function (r) {
    dateKey_(r.dateKey);
    requireValue_(!Object.prototype.hasOwnProperty.call(overrides, r.dateKey), 'DUPLICATE_KEY');
    requireValue_(['WORKDAY', 'HOLIDAY'].indexOf(r.kind) >= 0, 'INVALID_CALENDAR');
    overrides[r.dateKey] = r.kind;
  });

  return {
    weekdays: JSON.parse(week.value),
    overrides: overrides,
    startDate: start.value
  };
}

function saveCalendar(payload) {
  var actor = owner_();
  requireValue_(payload && typeof payload === 'object', 'INVALID_INPUT');
  requireValue_(typeof payload.requestId === 'string' && /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(payload.requestId), 'INVALID_REQUEST_ID');

  return locked_(function () {
    var ss = spreadsheet_();
    var today = bangkokToday_();

    if (payload.action === 'SET_OVERRIDE' || payload.dateKey) {
      dateKey_(payload.dateKey);
      requireValue_(['WORKDAY', 'HOLIDAY'].indexOf(payload.kind) >= 0, 'INVALID_CALENDAR');
      var month = payload.dateKey.slice(0, 7);
      var tag = suffix_(month);

      // ตรวจสอบ receipt เดิม
      var fingerprint = fingerprint_(payload);
      var requests = table_(ss, 'Requests_' + tag, REQUEST_HEADERS_);
      var receipt = unique_(requests.rows, 'requestId', payload.requestId);
      if (receipt) {
        requireValue_(receipt.fingerprint === fingerprint, 'REQUEST_ID_REUSED');
        return JSON.parse(receipt.responseJson);
      }

      // เดือนต้องยังไม่ปิด
      requireValue_(monthState_(ss, month).state === 'OPEN', 'MONTH_CLOSED');

      var calTable = table_(ss, 'WorkCalendar', CALENDAR_HEADERS_);
      var existing = unique_(calTable.rows, 'dateKey', payload.dateKey);
      var revision = existing ? existing.revision : 0;
      requireValue_(revision === payload.expectedRevision, 'CONFLICT');

      var now = bangkokNowIso_();
      var after = {
        dateKey: payload.dateKey,
        kind: payload.kind,
        note: payload.note || '',
        revision: revision + 1,
        updatedAt: now
      };

      var values = CALENDAR_HEADERS_.map(function (h) { return after[h]; });
      var change = existing
        ? updateRow_(calTable.sheet, existing.rowNumber - 1, values)
        : append_(calTable.sheet, values);

      var audit = table_(ss, 'Audit_' + tag, AUDIT_HEADERS_);
      var response = { ok: true, calendar: after };

      Sheets.Spreadsheets.batchUpdate({
        requests: [
          change,
          append_(audit.sheet, [Utilities.getUuid(), 'CALENDAR|' + payload.dateKey, JSON.stringify(existing || null), JSON.stringify(after), now, actor, payload.requestId]),
          append_(requests.sheet, [payload.requestId, fingerprint, JSON.stringify(response), now])
        ]
      }, ss.getId());

      return response;
    }

    if (payload.action === 'SET_WORKWEEK') {
      requireValue_(Array.isArray(payload.weekdays) && payload.weekdays.every(function (d) { return Number.isInteger(d) && d >= 0 && d <= 6; }), 'INVALID_CALENDAR');
      var curMonth = today.slice(0, 7);
      var curTag = suffix_(curMonth);

      var fp = fingerprint_(payload);
      var reqs = table_(ss, 'Requests_' + curTag, REQUEST_HEADERS_);
      var rc = unique_(reqs.rows, 'requestId', payload.requestId);
      if (rc) {
        requireValue_(rc.fingerprint === fp, 'REQUEST_ID_REUSED');
        return JSON.parse(rc.responseJson);
      }

      var setTable = table_(ss, 'Settings', SETTINGS_HEADERS_);
      var weekRow = unique_(setTable.rows, 'key', 'workweekJson');
      requireValue_(weekRow, 'SETTINGS_NOT_FOUND');
      requireValue_(weekRow.revision === payload.expectedRevision, 'CONFLICT');

      var time = bangkokNowIso_();
      var weekVal = JSON.stringify(payload.weekdays);
      var afterWeek = {
        key: 'workweekJson',
        value: weekVal,
        revision: weekRow.revision + 1,
        updatedAt: time
      };

      var changeWeek = updateRow_(setTable.sheet, weekRow.rowNumber - 1, SETTINGS_HEADERS_.map(function (h) { return afterWeek[h]; }));
      var auditCur = table_(ss, 'Audit_' + curTag, AUDIT_HEADERS_);
      var res = { ok: true, workweek: payload.weekdays, revision: afterWeek.revision };

      Sheets.Spreadsheets.batchUpdate({
        requests: [
          changeWeek,
          append_(auditCur.sheet, [Utilities.getUuid(), 'SETTINGS|workweekJson', JSON.stringify(weekRow), JSON.stringify(afterWeek), time, actor, payload.requestId]),
          append_(reqs.sheet, [payload.requestId, fp, JSON.stringify(res), time])
        ]
      }, ss.getId());

      return res;
    }

    throw new Error('INVALID_INPUT');
  });
}


// ==========================================
// Module: PhotoService.gs
// ==========================================
/**
 * DE TEAM — ระบบเช็คชื่อพนักงานและคำนวณยอดค่าจ้าง
 * PhotoService.gs: บริการรูปภาพย่อพนักงาน (เก็บบน Google Sheets เท่านั้น ไม่ใช้ Drive)
 */

function validateJpegData_(base64Text) {
  requireValue_(typeof base64Text === 'string', 'INVALID_PHOTO');
  // ตัด data URL prefix ออกหากมีส่งมา
  var cleanBase64 = base64Text.indexOf('base64,') >= 0 ? base64Text.split('base64,')[1] : base64Text;
  cleanBase64 = cleanBase64.replace(/\s/g, '');

  requireValue_(cleanBase64.length > 0 && cleanBase64.length <= MAX_PHOTO_BASE64_LEN_, 'PHOTO_TOO_LARGE');

  var bytes = Utilities.base64Decode(cleanBase64);
  requireValue_(bytes.length > 4 && bytes.length <= MAX_PHOTO_BYTES_, 'PHOTO_TOO_LARGE');

  // ตรวจสอบ JPEG Magic Bytes: 0xFF, 0xD8 ที่ต้นไฟล์ และ 0xFF, 0xD9 ที่ท้ายไฟล์
  var b0 = (bytes[0] + 256) % 256;
  var b1 = (bytes[1] + 256) % 256;
  var bEnd1 = (bytes[bytes.length - 2] + 256) % 256;
  var bEnd2 = (bytes[bytes.length - 1] + 256) % 256;

  requireValue_(b0 === 0xFF && b1 === 0xD8, 'NOT_JPEG');
  requireValue_(bEnd1 === 0xFF && bEnd2 === 0xD9, 'NOT_JPEG');

  return cleanBase64;
}

function getPhotos(payload) {
  owner_();
  requireValue_(payload && typeof payload === 'object', 'INVALID_INPUT');
  requireValue_(Array.isArray(payload.employeeIds), 'INVALID_INPUT');
  requireValue_(payload.employeeIds.length <= 8, 'TOO_MANY_PHOTOS');

  var knownVersions = payload.knownVersions || {};

  return locked_(function () {
    var ss = spreadsheet_();
    var photoTable = table_(ss, 'Photos', PHOTO_HEADERS_);
    var result = {};

    payload.employeeIds.forEach(function (empId) {
      var photoRow = unique_(photoTable.rows, 'employeeId', empId);
      if (photoRow) {
        var currentVersion = photoRow.revision || 1;
        if (knownVersions[empId] !== currentVersion) {
          result[empId] = {
            employeeId: empId,
            jpegBase64: photoRow.jpegBase64,
            width: photoRow.width,
            height: photoRow.height,
            version: currentVersion
          };
        }
      }
    });

    return { ok: true, photos: result };
  });
}

function savePhoto(payload) {
  var actor = owner_();
  requireValue_(payload && typeof payload === 'object', 'INVALID_INPUT');
  requireValue_(typeof payload.employeeId === 'string' && payload.employeeId.length > 0, 'INVALID_EMPLOYEE');
  requireValue_(typeof payload.requestId === 'string' && /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(payload.requestId), 'INVALID_REQUEST_ID');

  var width = Number(payload.width) || 0;
  var height = Number(payload.height) || 0;
  requireValue_(width > 0 && width <= 192 && height > 0 && height <= 192, 'INVALID_DIMENSIONS');

  var cleanBase64 = validateJpegData_(payload.jpegBase64);

  return locked_(function () {
    var ss = spreadsheet_();
    var currentMonth = bangkokToday_().slice(0, 7);
    var currentTag = suffix_(currentMonth);

    var fingerprint = fingerprint_(payload);
    var requests = table_(ss, 'Requests_' + currentTag, REQUEST_HEADERS_);
    var receipt = unique_(requests.rows, 'requestId', payload.requestId);
    if (receipt) {
      requireValue_(receipt.fingerprint === fingerprint, 'REQUEST_ID_REUSED');
      return JSON.parse(receipt.responseJson);
    }

    var empTable = table_(ss, 'Employees', EMPLOYEE_HEADERS_);
    var employee = unique_(empTable.rows, 'employeeId', payload.employeeId);
    requireValue_(employee, 'EMPLOYEE_NOT_FOUND');

    var photoTable = table_(ss, 'Photos', PHOTO_HEADERS_);
    var existingPhoto = unique_(photoTable.rows, 'employeeId', payload.employeeId);
    var photoRevision = existingPhoto ? existingPhoto.revision : 0;
    requireValue_(photoRevision === (payload.expectedRevision || 0), 'CONFLICT');

    var timestamp = bangkokNowIso_();
    var newPhotoRevision = photoRevision + 1;

    var afterPhoto = {
      employeeId: payload.employeeId,
      jpegBase64: cleanBase64,
      width: width,
      height: height,
      revision: newPhotoRevision,
      updatedAt: timestamp
    };

    var batchRequests = [];
    var photoValues = PHOTO_HEADERS_.map(function (h) { return afterPhoto[h]; });
    var photoChange = existingPhoto
      ? updateRow_(photoTable.sheet, existingPhoto.rowNumber - 1, photoValues)
      : append_(photoTable.sheet, photoValues);
    batchRequests.push(photoChange);

    // อัปเดต photoVersion ใน Employees
    var updatedEmp = {
      employeeId: employee.employeeId,
      name: employee.name,
      nickname: employee.nickname,
      position: employee.position,
      startDate: employee.startDate,
      endDate: employee.endDate,
      notes: employee.notes,
      photoVersion: newPhotoRevision,
      revision: employee.revision + 1,
      updatedAt: timestamp
    };
    batchRequests.push(updateRow_(empTable.sheet, employee.rowNumber - 1, EMPLOYEE_HEADERS_.map(function (h) { return updatedEmp[h]; })));

    var audit = table_(ss, 'Audit_' + currentTag, AUDIT_HEADERS_);
    var response = { ok: true, employeeId: payload.employeeId, photoVersion: newPhotoRevision };

    batchRequests.push(append_(audit.sheet, [
      Utilities.getUuid(),
      'PHOTO|' + payload.employeeId,
      JSON.stringify(existingPhoto ? { width: existingPhoto.width, height: existingPhoto.height, revision: existingPhoto.revision } : null),
      JSON.stringify({ width: width, height: height, revision: newPhotoRevision }),
      timestamp,
      actor,
      payload.requestId
    ]));
    batchRequests.push(append_(requests.sheet, [
      payload.requestId,
      fingerprint,
      JSON.stringify(response),
      timestamp
    ]));

    Sheets.Spreadsheets.batchUpdate({ requests: batchRequests }, ss.getId());

    return response;
  });
}


// ==========================================
// Module: ReportService.gs
// ==========================================
/**
 * DE TEAM — ระบบเช็คชื่อพนักงานและคำนวณยอดค่าจ้าง
 * ReportService.gs: บริการรายงานสรุปค่าจ้างรายเดือน และรายละเอียดรายบุคคล
 */

function readMonthSnapshots_(ss, month, targetVersion) {
  var tag = suffix_(month);
  var snapSheet = ss.getSheetByName('Snapshots_' + tag);
  if (!snapSheet) return null;

  var snapTable = table_(ss, 'Snapshots_' + tag, SNAPSHOT_HEADERS_);
  var rows = snapTable.rows.filter(function (r) {
    return Number(r.snapshotVersion) === Number(targetVersion);
  });

  if (rows.length === 0) return null;

  var byEmployee = {};
  rows.forEach(function (r) {
    if (!byEmployee[r.employeeId]) {
      byEmployee[r.employeeId] = [];
    }
    byEmployee[r.employeeId].push(r);
  });

  var result = [];
  Object.keys(byEmployee).forEach(function (empId) {
    var parts = byEmployee[empId].sort(function (a, b) {
      return Number(a.jsonPartIndex) - Number(b.jsonPartIndex);
    });
    var fullJson = parts.map(function (p) { return p.snapshotJsonPart; }).join('');
    try {
      result.push(JSON.parse(fullJson));
    } catch (e) {
      throw new Error('CORRUPT_SNAPSHOT: ' + empId);
    }
  });

  return result;
}

function getMonthReport(monthKey) {
  owner_();
  monthKey_(monthKey);

  return locked_(function () {
    var ss = spreadsheet_();
    var mState = monthState_(ss, monthKey);
    var isClosed = mState.state === 'CLOSED';

    if (isClosed && mState.snapshotVersion > 0) {
      // เดือนปิดแล้ว: อ่านข้อมูลจาก Snapshots
      var snapshots = readMonthSnapshots_(ss, monthKey, mState.snapshotVersion);
      requireValue_(snapshots !== null, 'SNAPSHOT_NOT_FOUND');

      var totalSatang = 0;
      var baseSatang = 0;
      var extraSatang = 0;
      var employeeSummaries = snapshots.map(function (s) {
        totalSatang += s.totalSatang || 0;
        baseSatang += s.baseSatang || 0;
        extraSatang += s.extraSatang || 0;
        return {
          employeeId: s.employeeId,
          name: s.name,
          nickname: s.nickname,
          position: s.position,
          photoVersion: s.photoVersion,
          full: s.full,
          half: s.half,
          absent: s.absent,
          pending: s.pending,
          workedDays: s.workedDays,
          paidDayUnits: s.paidDayUnits,
          baseSatang: s.baseSatang,
          extraSatang: s.extraSatang,
          totalSatang: s.totalSatang
        };
      });

      return {
        monthKey: monthKey,
        isClosed: true,
        snapshotVersion: mState.snapshotVersion,
        closedAt: mState.closedAt,
        closedBy: mState.closedBy,
        serverToday: bangkokToday_(),
        totalSatang: totalSatang,
        baseSatang: baseSatang,
        extraSatang: extraSatang,
        pendingCount: 0,
        employees: employeeSummaries
      };
    }

    // เดือนยังเปิดอยู่: คำนวณแบบ Real-time
    var today = bangkokToday_();
    var cal = calendar_(ss);
    var tag = suffix_(monthKey);

    if (!ss.getSheetByName('Attendance_' + tag)) {
      initMonthTables_(ss, monthKey);
    }

    var empTable = table_(ss, 'Employees', EMPLOYEE_HEADERS_);
    var rateTable = table_(ss, 'RateHistory', RATE_HEADERS_);
    var attTable = table_(ss, 'Attendance_' + tag, ATTENDANCE_HEADERS_);
    var extraTable = table_(ss, 'MonthlyExtras_' + tag, MONTHLY_EXTRA_HEADERS_);

    var monthStart = monthKey + '-01';
    var monthDatesList = monthDates_(monthKey);
    var monthEnd = monthDatesList[monthDatesList.length - 1];

    // กรองพนักงานที่อยู่ในช่วงการจ้างงานของเดือนนี้
    var eligibleEmployees = empTable.rows.filter(function (e) {
      validateEmployeeDates_(e);
      if (e.startDate > monthEnd) return false;
      if (e.endDate && e.endDate < monthStart) return false;
      return true;
    });

    var totalSatangSum = 0;
    var baseSatangSum = 0;
    var extraSatangSum = 0;
    var pendingCountSum = 0;

    var employeeSummaries = eligibleEmployees.map(function (emp) {
      var calc = calculateEmployeeMonth_({
        month: monthKey,
        today: today,
        employee: emp,
        rates: rateTable.rows,
        attendance: attTable.rows,
        extras: extraTable.rows,
        weekdays: cal.weekdays,
        calendar: cal.overrides,
        systemStartDate: cal.startDate
      });

      totalSatangSum += calc.totalSatang;
      baseSatangSum += calc.baseSatang;
      extraSatangSum += calc.extraSatang;
      pendingCountSum += calc.pending;

      return {
        employeeId: emp.employeeId,
        name: emp.name,
        nickname: emp.nickname,
        position: emp.position,
        photoVersion: emp.photoVersion,
        full: calc.full,
        half: calc.half,
        absent: calc.absent,
        pending: calc.pending,
        workedDays: calc.workedDays,
        paidDayUnits: calc.paidDayUnits,
        baseSatang: calc.baseSatang,
        extraSatang: calc.extraSatang,
        totalSatang: calc.totalSatang
      };
    });

    return {
      monthKey: monthKey,
      isClosed: false,
      snapshotVersion: 0,
      closedAt: '',
      closedBy: '',
      serverToday: today,
      totalSatang: totalSatangSum,
      baseSatang: baseSatangSum,
      extraSatang: extraSatangSum,
      pendingCount: pendingCountSum,
      employees: employeeSummaries
    };
  });
}

function getEmployeeMonth(employeeId, monthKey) {
  owner_();
  requireValue_(typeof employeeId === 'string' && employeeId.length > 0, 'INVALID_EMPLOYEE');
  monthKey_(monthKey);

  return locked_(function () {
    var ss = spreadsheet_();
    var mState = monthState_(ss, monthKey);
    var isClosed = mState.state === 'CLOSED';

    if (isClosed && mState.snapshotVersion > 0) {
      var snapshots = readMonthSnapshots_(ss, monthKey, mState.snapshotVersion);
      requireValue_(snapshots !== null, 'SNAPSHOT_NOT_FOUND');
      var empSnap = snapshots.find(function (s) { return s.employeeId === employeeId; });
      requireValue_(empSnap, 'EMPLOYEE_NOT_IN_SNAPSHOT');
      return { ok: true, isClosed: true, detail: empSnap };
    }

    var today = bangkokToday_();
    var cal = calendar_(ss);
    var tag = suffix_(monthKey);

    if (!ss.getSheetByName('Attendance_' + tag)) {
      initMonthTables_(ss, monthKey);
    }

    var empTable = table_(ss, 'Employees', EMPLOYEE_HEADERS_);
    var emp = unique_(empTable.rows, 'employeeId', employeeId);
    requireValue_(emp, 'EMPLOYEE_NOT_FOUND');

    var rateTable = table_(ss, 'RateHistory', RATE_HEADERS_);
    var attTable = table_(ss, 'Attendance_' + tag, ATTENDANCE_HEADERS_);
    var extraTable = table_(ss, 'MonthlyExtras_' + tag, MONTHLY_EXTRA_HEADERS_);
    var reviewTable = table_(ss, 'ExtraReviews_' + tag, EXTRA_REVIEW_HEADERS_);

    var calc = calculateEmployeeMonth_({
      month: monthKey,
      today: today,
      employee: emp,
      rates: rateTable.rows,
      attendance: attTable.rows,
      extras: extraTable.rows,
      weekdays: cal.weekdays,
      calendar: cal.overrides,
      systemStartDate: cal.startDate
    });

    var review = unique_(reviewTable.rows, 'employeeId', employeeId);
    var isExtrasConfirmed = Boolean(review && review.confirmedAt);

    var rateHistory = ratesFor_(rateTable.rows, employeeId);

    var detail = {
      employeeId: emp.employeeId,
      name: emp.name,
      nickname: emp.nickname,
      position: emp.position,
      startDate: emp.startDate,
      endDate: emp.endDate,
      photoVersion: emp.photoVersion,
      month: monthKey,
      full: calc.full,
      half: calc.half,
      absent: calc.absent,
      pending: calc.pending,
      workedDays: calc.workedDays,
      paidDayUnits: calc.paidDayUnits,
      baseSatang: calc.baseSatang,
      extraSatang: calc.extraSatang,
      totalSatang: calc.totalSatang,
      days: calc.days,
      extras: calc.extras,
      rates: rateHistory,
      extrasConfirmed: isExtrasConfirmed,
      confirmedAt: review ? review.confirmedAt : ''
    };

    return { ok: true, isClosed: false, detail: detail };
  });
}


// ==========================================
// Module: MonthService.gs
// ==========================================
/**
 * DE TEAM — ระบบเช็คชื่อพนักงานและคำนวณยอดค่าจ้าง
 * MonthService.gs: บริการปิดเดือน (Snapshots) และเปิดเดือนเพื่อแก้ไข
 */

function closeMonth(payload) {
  var actor = owner_();
  requireValue_(payload && typeof payload === 'object', 'INVALID_INPUT');
  monthKey_(payload.monthKey);
  requireValue_(typeof payload.requestId === 'string' && /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(payload.requestId), 'INVALID_REQUEST_ID');

  return locked_(function () {
    var ss = spreadsheet_();
    var month = payload.monthKey;
    var tag = suffix_(month);

    var fingerprint = fingerprint_(payload);
    var requests = table_(ss, 'Requests_' + tag, REQUEST_HEADERS_);
    var receipt = unique_(requests.rows, 'requestId', payload.requestId);
    if (receipt) {
      requireValue_(receipt.fingerprint === fingerprint, 'REQUEST_ID_REUSED');
      return JSON.parse(receipt.responseJson);
    }

    var mStateTable = table_(ss, 'MonthState', MONTH_HEADERS_);
    var mState = unique_(mStateTable.rows, 'monthKey', month);
    requireValue_(mState, 'MONTH_NOT_FOUND');
    requireValue_(mState.state === 'OPEN', 'MONTH_ALREADY_CLOSED');
    requireValue_(mState.revision === payload.expectedRevision, 'CONFLICT');

    // 1. เดือนต้องสิ้นสุดแล้วตามเวลา Bangkok
    var dates = monthDates_(month);
    var lastDate = dates[dates.length - 1];
    var today = bangkokToday_();
    requireValue_(today > lastDate, 'MONTH_NOT_ENDED');

    // 2. ดึงข้อมูลพนักงาน ปฏิทิน อัตรา การเช็คชื่อ และเงินพิเศษ
    var cal = calendar_(ss);
    var empTable = table_(ss, 'Employees', EMPLOYEE_HEADERS_);
    var rateTable = table_(ss, 'RateHistory', RATE_HEADERS_);
    var attTable = table_(ss, 'Attendance_' + tag, ATTENDANCE_HEADERS_);
    var extraTable = table_(ss, 'MonthlyExtras_' + tag, MONTHLY_EXTRA_HEADERS_);
    var reviewTable = table_(ss, 'ExtraReviews_' + tag, EXTRA_REVIEW_HEADERS_);

    var monthStart = month + '-01';
    var eligibleEmployees = empTable.rows.filter(function (e) {
      validateEmployeeDates_(e);
      if (e.startDate > lastDate) return false;
      if (e.endDate && e.endDate < monthStart) return false;
      return true;
    });

    // 3. ตรวจสอบว่าเช็คชื่อครบทุกคน ไม่มี pending workdays
    // 4. ตรวจสอบว่าเงินพิเศษได้รับการยืนยันครบทุกคน
    var timestamp = bangkokNowIso_();
    var newSnapshotVersion = (mState.snapshotVersion || 0) + 1;
    var snapshotRows = [];
    var totalMonthSatang = 0;

    eligibleEmployees.forEach(function (emp) {
      var calc = calculateEmployeeMonth_({
        month: month,
        today: today,
        employee: emp,
        rates: rateTable.rows,
        attendance: attTable.rows,
        extras: extraTable.rows,
        weekdays: cal.weekdays,
        calendar: cal.overrides,
        systemStartDate: cal.startDate
      });

      requireValue_(calc.pending === 0, 'UNMARKED_WORKDAYS_REMAIN: ' + emp.name);

      var review = unique_(reviewTable.rows, 'employeeId', emp.employeeId);
      requireValue_(review && review.confirmedAt, 'EXTRAS_NOT_REVIEWED: ' + emp.name);

      totalMonthSatang += calc.totalSatang;

      var empSnapshot = {
        snapshotVersion: newSnapshotVersion,
        employeeId: emp.employeeId,
        name: emp.name,
        nickname: emp.nickname,
        position: emp.position,
        photoVersion: emp.photoVersion,
        startDate: emp.startDate,
        endDate: emp.endDate,
        month: month,
        full: calc.full,
        half: calc.half,
        absent: calc.absent,
        pending: calc.pending,
        workedDays: calc.workedDays,
        paidDayUnits: calc.paidDayUnits,
        baseSatang: calc.baseSatang,
        extraSatang: calc.extraSatang,
        totalSatang: calc.totalSatang,
        days: calc.days,
        extras: calc.extras,
        rates: ratesFor_(rateTable.rows, emp.employeeId),
        closedAt: timestamp,
        closedBy: actor
      };

      var jsonStr = JSON.stringify(empSnapshot);
      // แบ่ง chunk ไม่เกิน MAX_SNAPSHOT_PART_LEN_ (30,000 ตัวอักษร)
      var partCount = Math.ceil(jsonStr.length / MAX_SNAPSHOT_PART_LEN_);
      for (var p = 0; p < partCount; p++) {
        var partText = jsonStr.slice(p * MAX_SNAPSHOT_PART_LEN_, (p + 1) * MAX_SNAPSHOT_PART_LEN_);
        snapshotRows.push([newSnapshotVersion, emp.employeeId, p + 1, partCount, partText]);
      }
    });

    // 5. บันทึก Snapshots และอัปเดต MonthState เป็น CLOSED ใน batch เดียวกัน
    var snapTable = table_(ss, 'Snapshots_' + tag, SNAPSHOT_HEADERS_);
    var batchRequests = [];

    snapshotRows.forEach(function (rowValues) {
      batchRequests.push(append_(snapTable.sheet, rowValues));
    });

    var afterMonthState = {
      monthKey: month,
      state: 'CLOSED',
      revision: mState.revision + 1,
      snapshotVersion: newSnapshotVersion,
      closedAt: timestamp,
      closedBy: actor
    };
    var mStateValues = MONTH_HEADERS_.map(function (h) { return afterMonthState[h]; });
    batchRequests.push(updateRow_(mStateTable.sheet, mState.rowNumber - 1, mStateValues));

    var audit = table_(ss, 'Audit_' + tag, AUDIT_HEADERS_);
    var response = {
      ok: true,
      state: 'CLOSED',
      snapshotVersion: newSnapshotVersion,
      closedAt: timestamp,
      closedBy: actor,
      totalSatang: totalMonthSatang
    };

    batchRequests.push(append_(audit.sheet, [
      Utilities.getUuid(),
      'MONTH_CLOSE|' + month,
      JSON.stringify(mState),
      JSON.stringify(afterMonthState),
      timestamp,
      actor,
      payload.requestId
    ]));
    batchRequests.push(append_(requests.sheet, [
      payload.requestId,
      fingerprint,
      JSON.stringify(response),
      timestamp
    ]));

    Sheets.Spreadsheets.batchUpdate({ requests: batchRequests }, ss.getId());

    return response;
  });
}

function reopenMonth(payload) {
  var actor = owner_();
  requireValue_(payload && typeof payload === 'object', 'INVALID_INPUT');
  monthKey_(payload.monthKey);
  requireValue_(typeof payload.reason === 'string' && payload.reason.trim().length > 0 && payload.reason.length <= 200, 'INVALID_REASON');
  requireValue_(typeof payload.requestId === 'string' && /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(payload.requestId), 'INVALID_REQUEST_ID');

  return locked_(function () {
    var ss = spreadsheet_();
    var month = payload.monthKey;
    var tag = suffix_(month);

    var fingerprint = fingerprint_(payload);
    var requests = table_(ss, 'Requests_' + tag, REQUEST_HEADERS_);
    var receipt = unique_(requests.rows, 'requestId', payload.requestId);
    if (receipt) {
      requireValue_(receipt.fingerprint === fingerprint, 'REQUEST_ID_REUSED');
      return JSON.parse(receipt.responseJson);
    }

    var mStateTable = table_(ss, 'MonthState', MONTH_HEADERS_);
    var mState = unique_(mStateTable.rows, 'monthKey', month);
    requireValue_(mState, 'MONTH_NOT_FOUND');
    requireValue_(mState.state === 'CLOSED', 'MONTH_ALREADY_OPEN');
    requireValue_(mState.revision === payload.expectedRevision, 'CONFLICT');

    var timestamp = bangkokNowIso_();
    var afterMonthState = {
      monthKey: month,
      state: 'OPEN',
      revision: mState.revision + 1,
      snapshotVersion: mState.snapshotVersion, // คง snapshotVersion เดิมไว้
      closedAt: '',
      closedBy: ''
    };

    var mStateValues = MONTH_HEADERS_.map(function (h) { return afterMonthState[h]; });
    var change = updateRow_(mStateTable.sheet, mState.rowNumber - 1, mStateValues);

    var audit = table_(ss, 'Audit_' + tag, AUDIT_HEADERS_);
    var response = { ok: true, state: 'OPEN', revision: afterMonthState.revision };

    Sheets.Spreadsheets.batchUpdate({
      requests: [
        change,
        append_(audit.sheet, [
          Utilities.getUuid(),
          'MONTH_REOPEN|' + month,
          JSON.stringify(mState),
          JSON.stringify({ state: 'OPEN', reason: payload.reason.trim() }),
          timestamp,
          actor,
          payload.requestId
        ]),
        append_(requests.sheet, [
          payload.requestId,
          fingerprint,
          JSON.stringify(response),
          timestamp
        ])
      ]
    }, ss.getId());

    return response;
  });
}


// ==========================================
// Module: Code.gs
// ==========================================
/**
 * DE TEAM — ระบบเช็คชื่อพนักงานและคำนวณยอดค่าจ้าง
 * Code.gs: จุดเริ่มต้น Web App (doGet) และ Public Endpoints Gateway
 */

function doGet(e) {
  var output;
  try {
    var template = HtmlService.createTemplateFromFile('Index');
    output = template.evaluate();
  } catch (err) {
    output = HtmlService.createHtmlOutput(getAppHtml_());
  }
  return output
    .setTitle('DE TEAM — เช็คชื่อและสรุปค่าจ้าง')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getBootstrap() {
  owner_();
  return locked_(function () {
    var ss = spreadsheet_();
    var today = bangkokToday_();
    var curMonth = today.slice(0, 7);

    // ตรวจสอบและสร้างชีตประจำเดือนปัจจุบันถ้ายังไม่มี
    if (!ss.getSheetByName('Attendance_' + suffix_(curMonth))) {
      initMonthTables_(ss, curMonth);
    }

    var settings = table_(ss, 'Settings', SETTINGS_HEADERS_).rows;
    var shopNameRow = unique_(settings, 'key', 'shopName');
    var shopName = shopNameRow ? shopNameRow.value : DEFAULT_SHOP_NAME_;

    var cal = calendar_(ss);
    var empTable = table_(ss, 'Employees', EMPLOYEE_HEADERS_);
    var employees = empTable.rows.map(function (e) {
      validateEmployeeDates_(e);
      return {
        employeeId: e.employeeId,
        name: e.name,
        nickname: e.nickname,
        position: e.position,
        startDate: e.startDate,
        endDate: e.endDate,
        notes: e.notes,
        photoVersion: e.photoVersion,
        revision: e.revision
      };
    });

    var mState = monthState_(ss, curMonth);
    var dayData = getDay_(ss, today);

    return {
      serverToday: today,
      monthKey: curMonth,
      shopName: shopName,
      isClosed: mState.state === 'CLOSED',
      monthState: {
        monthKey: curMonth,
        state: mState.state,
        revision: mState.revision,
        snapshotVersion: mState.snapshotVersion
      },
      calendar: {
        weekdays: cal.weekdays,
        startDate: cal.startDate
      },
      employees: employees,
      dayAttendance: dayData
    };
  });
}

function setupSystem() {
  owner_();
  return locked_(function () {
    return setupSystem_();
  });
}


// ==========================================
// Module: SetupBackend.gs
// ==========================================
/**
 * DE TEAM — สคริปต์ติดตั้งระบบฐานข้อมูล Google Sheets อัตโนมัติ (One-Click Setup)
 * รันฟังก์ชัน setupAttendanceBackend() ครั้งเดียวเพื่อสร้างตารางทั้งหมด 13 ตารางในชีตนี้
 * 
 * ชีตเป้าหมาย: https://docs.google.com/spreadsheets/d/1D9gYkF6uw9UdDDd4mrTBd3W7JU_kRtIawlqwLfce8bM/edit
 * Spreadsheet ID: 1D9gYkF6uw9UdDDd4mrTBd3W7JU_kRtIawlqwLfce8bM
 */

function setupAttendanceBackend() {
  var SPREADSHEET_ID = '1D9gYkF6uw9UdDDd4mrTBd3W7JU_kRtIawlqwLfce8bM';
  var ss = null;
  
  if (typeof SpreadsheetApp !== 'undefined') {
    try {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    } catch (e) {}
    if (!ss) {
      ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    }
  }

  if (!ss) {
    throw new Error('ไม่สามารถเปิด Google Spreadsheet ได้ กรุณาตรวจสอบสิทธิ์การเข้าถึง');
  }

  // 1. ตั้งค่า Script Properties อัตโนมัติ
  var userEmail = 'tlextle23@gmail.com';
  if (typeof Session !== 'undefined' && Session.getActiveUser) {
    var activeEmail = Session.getActiveUser().getEmail();
    if (activeEmail && activeEmail.trim().length > 0) {
      userEmail = activeEmail.trim();
    }
  }

  if (typeof PropertiesService !== 'undefined' && PropertiesService.getScriptProperties) {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('SPREADSHEET_ID', SPREADSHEET_ID);
    props.setProperty('OWNER_EMAIL', userEmail);
    props.setProperty('ADMIN_KEY', 'DE_TEAM_SECURE_ADMIN_KEY');
  }

  var nowIso = new Date().toISOString();
  var today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  var currentMonth = today.slice(0, 7);
  var monthTag = currentMonth.replace('-', '_');

  // โครงสร้างตารางและหัวคอลัมน์ทั้งหมดตาม BLUEPRINT.md
  var tables = [
    {
      name: 'Settings',
      headers: ['key', 'value', 'revision', 'updatedAt'],
      initialRows: [
        ['schemaVersion', '1', 1, nowIso],
        ['shopName', 'DE TEAM', 1, nowIso],
        ['workweekJson', '[1,2,3,4,5,6]', 1, nowIso],
        ['calendarEffectiveFrom', today, 1, nowIso]
      ]
    },
    {
      name: 'Employees',
      headers: ['employeeId', 'name', 'nickname', 'position', 'startDate', 'endDate', 'notes', 'photoVersion', 'revision', 'updatedAt'],
      initialRows: [
        ['e1', 'สมชาย ใจดี', 'ชาย', 'ช่างติดตั้ง', '2026-08-01', '', '', 0, 1, nowIso],
        ['e2', 'มาลี แสงทอง', 'ลี', 'ผู้ช่วยช่าง', '2026-08-01', '', '', 0, 1, nowIso],
        ['e3', 'วิชัย มีสุข', 'ชัย', 'คนขับรถ', '2026-08-01', '', '', 0, 1, nowIso]
      ]
    },
    {
      name: 'RateHistory',
      headers: ['rateId', 'employeeId', 'effectiveFrom', 'dailySatang', 'revision', 'updatedAt'],
      initialRows: [
        ['r1', 'e1', '2026-08-01', 50000, 1, nowIso],
        ['r2', 'e2', '2026-08-01', 45000, 1, nowIso],
        ['r3', 'e3', '2026-08-01', 45000, 1, nowIso]
      ]
    },
    {
      name: 'ExtraTemplates',
      headers: ['templateId', 'employeeId', 'label', 'amountSatang', 'effectiveFromMonth', 'effectiveToMonth', 'revision', 'updatedAt'],
      initialRows: [
        ['xt1', 'e1', 'ค่าเดินทาง', 150000, '2026-08', '', 1, nowIso],
        ['xt2', 'e1', 'ค่าอาหาร', 100000, '2026-08', '', 1, nowIso],
        ['xt3', 'e2', 'ค่าเบี้ยเลี้ยง', 100000, '2026-08', '', 1, nowIso]
      ]
    },
    {
      name: 'Photos',
      headers: ['employeeId', 'jpegBase64', 'width', 'height', 'revision', 'updatedAt'],
      initialRows: []
    },
    {
      name: 'WorkCalendar',
      headers: ['dateKey', 'kind', 'note', 'revision', 'updatedAt'],
      initialRows: []
    },
    {
      name: 'MonthState',
      headers: ['monthKey', 'state', 'revision', 'snapshotVersion', 'closedAt', 'closedBy'],
      initialRows: [
        [currentMonth, 'OPEN', 1, 0, '', '']
      ]
    },
    {
      name: 'Attendance_' + monthTag,
      headers: ['key', 'dateKey', 'employeeId', 'status', 'revision', 'updatedAt', 'updatedBy', 'requestId'],
      initialRows: []
    },
    {
      name: 'MonthlyExtras_' + monthTag,
      headers: ['extraId', 'employeeId', 'label', 'amountSatang', 'sourceTemplateId', 'sourceTemplateVersion', 'revision', 'updatedAt'],
      initialRows: []
    },
    {
      name: 'ExtraReviews_' + monthTag,
      headers: ['employeeId', 'extrasRevision', 'confirmedAt', 'confirmedBy'],
      initialRows: []
    },
    {
      name: 'Snapshots_' + monthTag,
      headers: ['snapshotVersion', 'employeeId', 'jsonPartIndex', 'jsonPartCount', 'snapshotJsonPart'],
      initialRows: []
    },
    {
      name: 'Audit_' + monthTag,
      headers: ['auditId', 'entityKey', 'beforeJson', 'afterJson', 'updatedAt', 'updatedBy', 'requestId'],
      initialRows: [
        ['audit_init', 'System', '', JSON.stringify({ action: 'SETUP_SYSTEM', target: SPREADSHEET_ID }), nowIso, userEmail, 'req_init']
      ]
    },
    {
      name: 'Requests_' + monthTag,
      headers: ['requestId', 'fingerprint', 'responseJson', 'updatedAt'],
      initialRows: []
    }
  ];

  // วนลูปสร้างชีต ตรวจสอบ และใส่หัวตาราง
  tables.forEach(function(tbl) {
    var sheet = ss.getSheetByName(tbl.name);
    if (!sheet) {
      sheet = ss.insertSheet(tbl.name);
    }
    
    // ถ้าชีตยังว่างเปล่า ให้ใส่หัวคอลัมน์และข้อมูลเริ่มต้น
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, tbl.headers.length).setValues([tbl.headers]);
      
      // จัดรูปแบบหัวตาราง: พื้นหลังสีกรมท่า ตัวอักษรสีขาว หนา ตรึงแถวบนสุด
      var headerRange = sheet.getRange(1, 1, 1, tbl.headers.length);
      headerRange.setBackground('#003566');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setFontWeight('bold');
      sheet.setFrozenRows(1);
      
      // ใส่ข้อมูลเริ่มต้น (ถ้ามี)
      if (tbl.initialRows && tbl.initialRows.length > 0) {
        sheet.getRange(2, 1, tbl.initialRows.length, tbl.headers.length).setValues(tbl.initialRows);
      }
    }
  });

  // ลบ Sheet1 หรือ แผ่นงาน1 เริ่มต้นทิ้ง หากมีตารางอื่นแล้ว
  var defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('แผ่นงาน1');
  if (defaultSheet && ss.getSheets().length > 1) {
    try {
      ss.deleteSheet(defaultSheet);
    } catch (e) {}
  }

  Logger.log('==========================================');
  Logger.log('✅ ติดตั้งฐานข้อมูลสำเร็จเรียบร้อยครบ 13 ตาราง!');
  Logger.log('📌 Spreadsheet ID: ' + SPREADSHEET_ID);
  Logger.log('👤 Owner Email: ' + userEmail);
  Logger.log('==========================================');

  return {
    ok: true,
    spreadsheetId: SPREADSHEET_ID,
    ownerEmail: userEmail,
    tableCount: tables.length,
    currentMonth: currentMonth
  };
}


// ==========================================
// Embedded Web Application HTML/CSS/JS (Self-Contained)
// ==========================================
function getAppHtml_() {
  return "<!doctype html>\n<html lang=\"th\">\n<head>\n  <meta charset=\"utf-8\">\n  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1, viewport-fit=cover\">\n  <meta name=\"description\" content=\"ระบบเช็คชื่อพนักงานและคำนวณสรุปยอดค่าจ้าง DE TEAM\">\n  <title>DE TEAM — เช็คชื่อและสรุปค่าจ้าง</title>\n  <!--\n  DE TEAM — Fonts.html\n  การประกาศฟอนต์ Sarabun (แบบมีหัว) 400 และ 700 พร้อม fallback Tahoma, sans-serif\n  หมายเหตุ: ตามข้อกำหนดใน BLUEPRINT.md หากยังไม่มีไฟล์ WOFF2 ที่ได้รับอนุญาต\n  ระบบจะใช้ fallback เป็น Tahoma, sans-serif เพื่อคงตัวหนังสือไทยแบบมีหัวอย่างซื่อสัตย์\n  และไม่ใส่ base64 ปลอม หรือเรียก CDN ภายนอกทุกครั้งที่เปิดหน้าเว็บ\n-->\n<style>\n  /* เมื่อมีไฟล์ฟอนต์ Sarabun WOFF2 ที่ได้รับอนุญาต สามารถนำ base64 มาใส่ใน url(data:font/woff2;base64,...) */\n  @font-face {\n    font-family: 'Sarabun';\n    font-style: normal;\n    font-weight: 400;\n    font-display: swap;\n    src: local('Sarabun Regular'), local('Sarabun'), local('Tahoma');\n  }\n\n  @font-face {\n    font-family: 'Sarabun';\n    font-style: normal;\n    font-weight: 700;\n    font-display: swap;\n    src: local('Sarabun Bold'), local('Sarabun-Bold'), local('Tahoma Bold'), local('Tahoma');\n  }\n\n  /* กำหนด fallback stack ให้ตัวอักษรภาษาไทยมีหัวเสมอ */\n  :root {\n    --team-font-family: 'Sarabun', Tahoma, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;\n  }\n</style>\n\n  <style>\n  /* ==========================================================================\n     DE TEAM — Design System & CSS Tokens\n     ออกแบบเฉพาะสำหรับผู้ดูแล Gen X: ตัวอักษรใหญ่ จุดกดกว้าง คอนทราสต์ชัด\n     ========================================================================== */\n\n  :root {\n    --team-primary: #003566;\n    --team-primary-dark: #001D3D;\n    --team-accent: #FFC300;\n    --team-bg: #F3F6FA;\n    --team-surface: #FFFFFF;\n    --team-text: #172B43;\n    --team-muted: #526277;\n    --team-border: #D7E0EA;\n\n    /* Status Tokens */\n    --team-full-text: #166534;\n    --team-full-bg: #F0FDF4;\n    --team-half-text: #854D0E;\n    --team-half-bg: #FFFBEB;\n    --team-absent-text: #991B1B;\n    --team-absent-bg: #FEF2F2;\n\n    /* Spacing & Sizes */\n    --team-copy: 1.2rem;       /* ~18px บนมือถือ */\n    --team-radius-card: 1.067rem;\n    --team-radius-btn: 0.8rem;\n    --team-radius-input: 0.65rem;\n    --team-touch-min: 3.467rem; /* ~52px สำหรับปุ่มหลักและปุ่มสถานะ */\n    --team-touch-sec: 3rem;     /* ~45px สำหรับจุดกดรอง */\n\n    color-scheme: light;\n  }\n\n  /* Root Font Size Scaling (REM-first) */\n  html {\n    font-size: 15px;\n    scrollbar-gutter: stable;\n    background-color: var(--team-bg);\n    color: var(--team-text);\n    -webkit-text-size-adjust: 100%;\n    margin: 0;\n    padding: 0;\n  }\n\n  @media (min-width: 640px) { html { font-size: 15.5px; } }\n  @media (min-width: 1024px) { html { font-size: 16px; } }\n  @media (min-width: 1440px) { html { font-size: 16.5px; } }\n  @media (min-width: 1920px) { html { font-size: 17.5px; } }\n\n  *, *::before, *::after {\n    box-sizing: border-box;\n  }\n\n  body {\n    margin: 0;\n    padding: 0;\n    min-height: 100dvh;\n    font-family: var(--team-font-family, Sarabun, Tahoma, sans-serif);\n    font-size: var(--team-copy);\n    line-height: 1.65;\n    letter-spacing: normal;\n    background-color: var(--team-bg);\n    color: var(--team-text);\n    display: flex;\n    flex-direction: column;\n  }\n\n  /* Accessibility & Focus States */\n  button:focus-visible,\n  input:focus-visible,\n  select:focus-visible,\n  textarea:focus-visible,\n  a:focus-visible {\n    outline: 3px solid var(--team-primary);\n    outline-offset: 2px;\n  }\n\n  /* Container Structure */\n  #team-app {\n    width: 100%;\n    max-width: 64rem;\n    margin-inline: auto;\n    display: flex;\n    flex-direction: column;\n    min-height: 100dvh;\n    background-color: var(--team-bg);\n  }\n\n  /* App Header */\n  .team-header {\n    background-color: var(--team-primary);\n    color: #FFFFFF;\n    padding: 1.15rem 1.2rem;\n    display: flex;\n    align-items: center;\n    justify-content: space-between;\n    gap: 1rem;\n    position: sticky;\n    top: 0;\n    z-index: 20;\n    border-bottom: 2px solid var(--team-accent);\n  }\n\n  .team-brand {\n    font-size: 1.25rem;\n    font-weight: 700;\n    letter-spacing: 0.04em;\n    display: flex;\n    align-items: center;\n    gap: 0.4rem;\n  }\n\n  .team-brand span {\n    color: var(--team-accent);\n  }\n\n  .team-header-subtitle {\n    font-size: 1.067rem;\n    color: #E2E8F0;\n  }\n\n  /* Main Content Area */\n  .team-main {\n    flex: 1;\n    padding: 1.2rem 1.067rem calc(6.5rem + env(safe-area-inset-bottom, 0px));\n  }\n\n  @media (min-width: 768px) {\n    .team-main {\n      padding: 1.5rem 1.5rem 2.5rem;\n    }\n  }\n\n  /* Typography */\n  h1, h2, h3, h4 {\n    margin: 0;\n    color: var(--team-primary-dark);\n    line-height: 1.3;\n    font-weight: 700;\n  }\n\n  h2 {\n    font-size: 1.67rem;\n    margin-bottom: 1.2rem;\n  }\n\n  h3 {\n    font-size: 1.33rem;\n    margin-bottom: 0.6rem;\n  }\n\n  p {\n    margin: 0.25rem 0;\n  }\n\n  .team-secondary {\n    color: var(--team-muted);\n    font-size: 1.067rem;\n  }\n\n  /* Buttons */\n  button {\n    font: inherit;\n    letter-spacing: normal;\n    min-height: var(--team-touch-min);\n    border: 1px solid var(--team-border);\n    border-radius: var(--team-radius-btn);\n    padding: 0.6rem 1rem;\n    background-color: var(--team-surface);\n    color: var(--team-primary);\n    cursor: pointer;\n    touch-action: manipulation;\n    user-select: none;\n    -webkit-tap-highlight-color: transparent;\n    display: inline-flex;\n    align-items: center;\n    justify-content: center;\n    gap: 0.5rem;\n    font-weight: 700;\n  }\n\n  button:hover:not(:disabled) {\n    border-color: var(--team-primary);\n  }\n\n  button:disabled {\n    opacity: 0.5;\n    cursor: not-allowed;\n  }\n\n  .team-btn-primary {\n    background-color: var(--team-primary);\n    color: #FFFFFF;\n    border-color: var(--team-primary);\n  }\n\n  .team-btn-primary:hover:not(:disabled) {\n    background-color: var(--team-primary-dark);\n    border-color: var(--team-primary-dark);\n  }\n\n  .team-btn-secondary {\n    background-color: #FFFFFF;\n    color: var(--team-text);\n    border-color: var(--team-border);\n  }\n\n  .team-btn-accent {\n    background-color: var(--team-accent);\n    color: var(--team-primary-dark);\n    border-color: var(--team-accent);\n  }\n\n  .team-btn-danger {\n    background-color: #FEF2F2;\n    color: var(--team-absent-text);\n    border-color: #FCA5A5;\n  }\n\n  .team-btn-wide {\n    width: 100%;\n  }\n\n  .team-btn-text {\n    border-color: transparent;\n    background: transparent;\n    min-height: var(--team-touch-sec);\n    padding: 0.3rem 0.5rem;\n    font-size: 1.067rem;\n    text-decoration: underline;\n    text-underline-offset: 0.25rem;\n    font-weight: 400;\n  }\n\n  /* Forms & Inputs */\n  label {\n    display: block;\n    margin-bottom: 1.15rem;\n    font-weight: 700;\n    color: var(--team-text);\n  }\n\n  input, select, textarea {\n    display: block;\n    min-height: var(--team-touch-min);\n    margin-top: 0.4rem;\n    padding: 0.6rem 0.8rem;\n    border: 1px solid #A7B7CA;\n    border-radius: var(--team-radius-input);\n    background-color: #FFFFFF;\n    color: var(--team-text);\n    width: 100%;\n    min-width: 0;\n    max-width: 100%;\n    font: inherit;\n  }\n\n  @media (pointer: coarse), (max-width: 640px) {\n    input, select, textarea {\n      font-size: max(var(--team-copy), 16px) !important;\n    }\n  }\n\n  textarea {\n    min-height: 6.5rem;\n    resize: vertical;\n  }\n\n  /* Cards & Layout Items */\n  .team-card {\n    background-color: var(--team-surface);\n    border: 1px solid var(--team-border);\n    border-radius: var(--team-radius-card);\n    padding: 1.15rem;\n    margin-bottom: 1rem;\n  }\n\n  .team-grid-2 {\n    display: grid;\n    grid-template-columns: 1fr;\n    gap: 1rem;\n  }\n\n  @media (min-width: 768px) {\n    .team-grid-2 {\n      grid-template-columns: 1fr 1fr;\n    }\n  }\n\n  /* Toolbar */\n  .team-toolbar {\n    display: flex;\n    align-items: center;\n    justify-content: space-between;\n    flex-wrap: wrap;\n    gap: 0.6rem;\n    margin-bottom: 1rem;\n  }\n\n  /* Person Attendance Card */\n  .team-person-card {\n    background-color: #FFFFFF;\n    border: 1px solid var(--team-border);\n    border-radius: var(--team-radius-card);\n    padding: 1.15rem;\n    margin-bottom: 1rem;\n  }\n\n  .team-person-head {\n    display: flex;\n    align-items: center;\n    gap: 0.9rem;\n    margin-bottom: 0.9rem;\n  }\n\n  .team-avatar {\n    flex: 0 0 3.8rem;\n    width: 3.8rem;\n    height: 4.4rem;\n    border-radius: 0.75rem;\n    background-color: #EDF2F7;\n    color: var(--team-muted);\n    display: flex;\n    align-items: center;\n    justify-content: center;\n    font-weight: 700;\n    font-size: 1.5rem;\n    overflow: hidden;\n    border: 1px solid var(--team-border);\n  }\n\n  .team-avatar img {\n    max-width: 100%;\n    width: auto;\n    height: auto;\n    max-height: 100%;\n    object-fit: contain;\n  }\n\n  .team-person-info {\n    min-width: 0;\n    flex: 1;\n  }\n\n  .team-person-name {\n    font-size: 1.33rem;\n    font-weight: 700;\n    color: var(--team-primary-dark);\n  }\n\n  /* Status Buttons Grid */\n  .team-statuses {\n    display: grid;\n    grid-template-columns: repeat(3, minmax(0, 1fr));\n    gap: 0.533rem;\n    margin-top: 0.5rem;\n  }\n\n  .team-status-btn {\n    min-height: var(--team-touch-min);\n    padding: 0.55rem 0.2rem;\n    font-size: var(--team-copy);\n    border: 1.5px solid var(--team-border);\n    background-color: #FFFFFF;\n    color: var(--team-primary);\n    border-radius: var(--team-radius-btn);\n  }\n\n  /* Status Pressed State */\n  .team-status-btn[aria-pressed=\"true\"] {\n    font-weight: 700;\n    border-width: 2.5px;\n  }\n\n  .team-status-btn[data-status=\"FULL\"][aria-pressed=\"true\"] {\n    background-color: var(--team-full-bg);\n    color: var(--team-full-text);\n    border-color: var(--team-full-text);\n  }\n\n  .team-status-btn[data-status=\"HALF\"][aria-pressed=\"true\"] {\n    background-color: var(--team-half-bg);\n    color: var(--team-half-text);\n    border-color: var(--team-half-text);\n  }\n\n  .team-status-btn[data-status=\"ABSENT\"][aria-pressed=\"true\"] {\n    background-color: var(--team-absent-bg);\n    color: var(--team-absent-text);\n    border-color: var(--team-absent-text);\n  }\n\n  /* Saved Status & Actions */\n  .team-saved-row {\n    display: flex;\n    align-items: center;\n    justify-content: space-between;\n    flex-wrap: wrap;\n    gap: 0.4rem;\n    margin-top: 0.75rem;\n    min-height: 2.8rem;\n    font-size: 1.067rem;\n    color: var(--team-muted);\n  }\n\n  .team-saved-badge {\n    display: inline-flex;\n    align-items: center;\n    gap: 0.35rem;\n    font-weight: 700;\n  }\n\n  .team-saved-badge.status-FULL { color: var(--team-full-text); }\n  .team-saved-badge.status-HALF { color: var(--team-half-text); }\n  .team-saved-badge.status-ABSENT { color: var(--team-absent-text); }\n  .team-saved-badge.status-SAVING { color: var(--team-primary); font-style: italic; }\n\n  /* Summary Big Numbers */\n  .team-summary-box {\n    background-color: var(--team-primary);\n    color: #FFFFFF;\n    border-radius: var(--team-radius-card);\n    padding: 1.3rem;\n    margin-bottom: 1.2rem;\n  }\n\n  .team-money-big {\n    font-size: 2.2rem;\n    line-height: 1.25;\n    font-weight: 700;\n    margin: 0.3rem 0;\n    font-variant-numeric: tabular-nums;\n    color: #FFFFFF;\n  }\n\n  .team-summary-box .team-secondary {\n    color: #E2E8F0;\n  }\n\n  /* Report Row Buttons */\n  .team-report-row {\n    display: block;\n    text-align: left;\n    width: 100%;\n    padding: 1.15rem;\n    margin-bottom: 1rem;\n    border: 1px solid var(--team-border);\n    background-color: #FFFFFF;\n    border-radius: var(--team-radius-card);\n    cursor: pointer;\n  }\n\n  .team-report-row:hover {\n    border-color: var(--team-primary);\n  }\n\n  .team-report-header {\n    display: flex;\n    justify-content: space-between;\n    align-items: center;\n    gap: 0.8rem;\n  }\n\n  .team-row-money {\n    font-size: 1.55rem;\n    font-weight: 700;\n    color: var(--team-primary-dark);\n    margin: 0.3rem 0;\n    font-variant-numeric: tabular-nums;\n  }\n\n  .team-counts-row {\n    display: flex;\n    flex-wrap: wrap;\n    gap: 0.3rem 1.2rem;\n    color: var(--team-muted);\n    font-size: 1.067rem;\n  }\n\n  /* Detail KPIs */\n  .team-kpis {\n    display: grid;\n    grid-template-columns: 1fr 1fr;\n    gap: 0.8rem;\n    margin-bottom: 1.2rem;\n  }\n\n  .team-kpi-card {\n    padding: 0.85rem;\n    background-color: #FFFFFF;\n    border: 1px solid var(--team-border);\n    border-radius: var(--team-radius-btn);\n  }\n\n  .team-kpi-card strong {\n    display: block;\n    font-size: 1.5rem;\n    color: var(--team-primary-dark);\n    font-variant-numeric: tabular-nums;\n  }\n\n  /* Line Breakdown */\n  .team-line-item {\n    display: flex;\n    justify-content: space-between;\n    align-items: center;\n    gap: 0.8rem;\n    padding: 0.7rem 0;\n    border-bottom: 1px solid #E1E7EE;\n  }\n\n  .team-line-item span:first-child {\n    min-width: 0;\n    overflow-wrap: anywhere;\n  }\n\n  .team-line-item span:last-child {\n    flex-shrink: 0;\n    font-variant-numeric: tabular-nums;\n  }\n\n  /* Notices & Alerts */\n  .team-notice-box {\n    background-color: #FFF9E6;\n    border: 1px solid #DAC576;\n    border-radius: var(--team-radius-btn);\n    padding: 0.9rem;\n    color: #63480C;\n    margin-bottom: 1rem;\n    font-weight: 700;\n  }\n\n  .team-notice-danger {\n    background-color: #FEF2F2;\n    border: 1px solid #FCA5A5;\n    color: #991B1B;\n  }\n\n  /* Bottom Navigation Bar */\n  .team-navigation {\n    display: grid;\n    grid-template-columns: repeat(3, minmax(0, 1fr));\n    position: fixed;\n    bottom: 0;\n    left: 0;\n    right: 0;\n    margin-inline: auto;\n    max-width: 64rem;\n    background-color: #FFFFFF;\n    border-top: 1px solid var(--team-border);\n    padding: 0.5rem 0.6rem calc(0.5rem + env(safe-area-inset-bottom, 0px));\n    gap: 0.5rem;\n    z-index: 30;\n  }\n\n  .team-nav-btn {\n    display: flex;\n    flex-direction: column;\n    align-items: center;\n    justify-content: center;\n    gap: 0.2rem;\n    border-color: transparent;\n    font-size: 1.067rem;\n    padding: 0.5rem 0.2rem;\n    min-height: 3.5rem;\n    border-radius: var(--team-radius-btn);\n    color: var(--team-muted);\n  }\n\n  .team-nav-btn svg {\n    width: 1.5rem;\n    height: 1.5rem;\n  }\n\n  .team-nav-btn[aria-current=\"page\"] {\n    color: var(--team-primary);\n    background-color: #E9F1F9;\n    border-bottom: 3.5px solid var(--team-accent);\n    font-weight: 700;\n  }\n\n  /* Modal / Full Screen Sheet for Forms */\n  .team-sheet-overlay {\n    position: fixed;\n    inset: 0;\n    background-color: rgba(15, 23, 42, 0.45);\n    z-index: 50;\n    overflow-y: auto;\n    padding: 1rem;\n    display: flex;\n    justify-content: center;\n  }\n\n  .team-sheet-container {\n    background-color: #FFFFFF;\n    border-radius: var(--team-radius-card);\n    width: 100%;\n    max-width: 42rem;\n    padding: 1.5rem;\n    box-shadow: 0 10px 25px rgba(0, 0, 0, 0.15);\n    margin: auto 0;\n  }\n\n  @media (max-width: 640px) {\n    .team-sheet-overlay {\n      padding: 0;\n    }\n    .team-sheet-container {\n      border-radius: 0;\n      min-height: 100dvh;\n      margin: 0;\n      padding: 1.2rem 1rem calc(5rem + env(safe-area-inset-bottom, 0px));\n    }\n  }\n\n  /* Live Announcements */\n  .team-live-region {\n    min-height: 2rem;\n    font-size: 1.067rem;\n    margin: 0.4rem 0 0.8rem;\n    color: var(--team-primary-dark);\n    font-weight: 700;\n  }\n\n  /* File Upload Preview */\n  .team-file-preview {\n    max-width: 10rem;\n    height: auto;\n    display: block;\n    margin: 0.6rem 0;\n    border-radius: 0.65rem;\n    border: 1px solid var(--team-border);\n  }\n\n  /* Print Styles */\n  @media print {\n    body {\n      background: #FFFFFF;\n      color: #000000;\n    }\n    .team-header,\n    .team-navigation,\n    .team-toolbar,\n    button:not(.team-print-allow) {\n      display: none !important;\n    }\n    .team-main {\n      padding: 0 !important;\n    }\n    .team-summary-box {\n      background: #FFFFFF !important;\n      color: #000000 !important;\n      border: 2px solid #000000;\n    }\n    .team-summary-box .team-money-big {\n      color: #000000 !important;\n    }\n    .team-card, .team-person-card {\n      break-inside: avoid;\n      border: 1px solid #CCCCCC;\n    }\n  }\n\n  @media (prefers-reduced-motion: reduce) {\n    * {\n      animation: none !important;\n      transition: none !important;\n      scroll-behavior: auto !important;\n    }\n  }\n</style>\n\n</head>\n<body>\n  <div id=\"team-app\" lang=\"th\" aria-label=\"ระบบเช็คชื่อพนักงานและคำนวณยอดค่าจ้าง DE TEAM\">\n    <!-- Header -->\n    <header class=\"team-header\">\n      <div class=\"team-brand\" id=\"team-brand-title\">DE <span>TEAM</span></div>\n      <div class=\"team-header-subtitle\" id=\"team-header-info\">ระบบเช็คชื่อและค่าจ้าง</div>\n    </header>\n\n    <!-- Main Dynamic Content -->\n    <main class=\"team-main\">\n      <div id=\"team-content\">\n        <!-- Skeleton / Loading State -->\n        <div class=\"team-card\">\n          <h2>กำลังเชื่อมต่อระบบ...</h2>\n          <p class=\"team-secondary\">กรุณารอสักครู่ ระบบกำลังอ่านข้อมูลล่าสุดจากเซิร์ฟเวอร์</p>\n        </div>\n      </div>\n      <p id=\"team-live\" class=\"team-live-region\" role=\"status\" aria-live=\"polite\"></p>\n    </main>\n\n    <!-- Bottom Navigation Bar -->\n    <nav class=\"team-navigation\" aria-label=\"เมนูหลัก\">\n      <button type=\"button\" class=\"team-nav-btn\" data-tab=\"attendance\" aria-current=\"page\" id=\"nav-btn-attendance\">\n        <svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2.2\" stroke-linecap=\"round\" stroke-linejoin=\"round\" aria-hidden=\"true\">\n          <path d=\"M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2\"></path>\n          <rect x=\"8\" y=\"2\" width=\"8\" height=\"4\" rx=\"1\" ry=\"1\"></rect>\n          <path d=\"m9 14 2 2 4-4\"></path>\n        </svg>\n        <span>เช็คชื่อ</span>\n      </button>\n      <button type=\"button\" class=\"team-nav-btn\" data-tab=\"reports\" id=\"nav-btn-reports\">\n        <svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2.2\" stroke-linecap=\"round\" stroke-linejoin=\"round\" aria-hidden=\"true\">\n          <line x1=\"18\" y1=\"20\" x2=\"18\" y2=\"10\"></line>\n          <line x1=\"12\" y1=\"20\" x2=\"12\" y2=\"4\"></line>\n          <line x1=\"6\" y1=\"20\" x2=\"6\" y2=\"14\"></line>\n        </svg>\n        <span>รายงาน</span>\n      </button>\n      <button type=\"button\" class=\"team-nav-btn\" data-tab=\"settings\" id=\"nav-btn-settings\">\n        <svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2.2\" stroke-linecap=\"round\" stroke-linejoin=\"round\" aria-hidden=\"true\">\n          <circle cx=\"12\" cy=\"12\" r=\"3\"></circle>\n          <path d=\"M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z\"></path>\n        </svg>\n        <span>ตั้งค่า</span>\n      </button>\n    </nav>\n  </div>\n\n  <script>\n/**\n * DE TEAM — Client Application Logic & State Controller\n * Apps Script HTML Service Vanilla JavaScript\n */\n\n(function () {\n  'use strict';\n\n  // --- 1. RPC Client Bridge ---\n  const TeamRpc = (() => {\n    const allowed = new Set([\n      'getBootstrap', 'getDay', 'saveAttendance', 'getMonthReport',\n      'getEmployeeMonth', 'saveEmployee', 'setEmploymentEnd', 'addRate',\n      'saveExtraTemplate', 'prepareMonthExtras', 'saveMonthExtra',\n      'confirmMonthExtras', 'getPhotos', 'savePhoto', 'saveCalendar',\n      'closeMonth', 'reopenMonth'\n    ]);\n\n    function call(method, ...args) {\n      if (!allowed.has(method)) return Promise.reject(new Error('INVALID_METHOD'));\n      const timeoutMs = 25000;\n      return new Promise((resolve, reject) => {\n        let settled = false;\n        const timer = setTimeout(() => {\n          settled = true;\n          reject(new Error('UNKNOWN'));\n        }, timeoutMs);\n\n        const finish = (fn, value) => {\n          clearTimeout(timer);\n          if (!settled) {\n            settled = true;\n            fn(value);\n          }\n        };\n\n        if (typeof google === 'undefined' || !google.script || !google.script.run) {\n          // Dev mock fallback if running directly in a standalone browser\n          settled = true;\n          clearTimeout(timer);\n          reject(new Error('NO_GOOGLE_SCRIPT_RUN'));\n          return;\n        }\n\n        const runner = google.script.run\n          .withSuccessHandler(val => finish(resolve, val))\n          .withFailureHandler(err => finish(reject, err));\n\n        runner[method].apply(runner, args);\n      });\n    }\n\n    return { call };\n  })();\n\n  // --- 2. State & Helpers ---\n  const appRoot = document.getElementById('team-app');\n  const contentEl = document.getElementById('team-content');\n  const liveEl = document.getElementById('team-live');\n  const brandTitleEl = document.getElementById('team-brand-title');\n  const headerInfoEl = document.getElementById('team-header-info');\n\n  const state = {\n    tab: 'attendance',\n    serverToday: '',\n    selectedDate: '',\n    selectedMonth: '',\n    filterAttendance: false, // false = all, true = unmarked only\n    searchQuery: '',\n    shopName: 'DE TEAM',\n    isClosed: false,\n    employees: [],\n    dayAttendance: null,\n    monthReport: null,\n    selectedDetailId: null,\n    editEmployee: null,      // null, 'new', or employeeId\n    editFormDraft: null,\n    photosCache: new Map(),  // employeeId -> { jpegBase64, version }\n    pendingCommands: new Map(), // key -> { command, status: 'SAVING'|'UNKNOWN' }\n    activeModal: null        // null or 'CLEAR_ATTENDANCE'|'CLOSE_MONTH'|'REOPEN_MONTH'|'EDIT_EXTRAS'\n  };\n\n  const STATUS_LABEL = {\n    FULL: 'เต็มวัน',\n    HALF: 'ครึ่งวัน',\n    ABSENT: 'ไม่มา',\n    UNMARKED: 'ยังไม่เช็ค'\n  };\n\n  const ERROR_MESSAGES = {\n    AUTH_REQUIRED: 'ต้องเข้าใช้งานด้วยบัญชี Google ของเจ้าของระบบเท่านั้น',\n    CONFIG_REQUIRED: 'ยังไม่ได้ตั้งค่า SPREADSHEET_ID หรือ OWNER_EMAIL ใน Script Properties',\n    BUSY: 'ระบบกำลังบันทึกข้อมูล กรุณาลองใหม่อีกครั้งใน 2-3 วินาที',\n    CONFLICT: 'ข้อมูลมีการเปลี่ยนแปลงจากเครื่องอื่น กรุณาโหลดข้อมูลใหม่',\n    MONTH_CLOSED: 'เดือนนี้ถูกปิดรอบรายงานแล้ว ไม่สามารถแก้ไขได้',\n    MONTH_ALREADY_CLOSED: 'เดือนนี้ถูกปิดรอบรายงานไปแล้ว',\n    MONTH_ALREADY_OPEN: 'เดือนนี้เปิดอยู่แล้ว ไม่จำเป็นต้องเปิดซ้ำ',\n    MONTH_NOT_ENDED: 'ยังไม่สามารถปิดเดือนได้จนกว่าจะสิ้นสุดเดือนตามเวลาประเทศไทย',\n    UNMARKED_WORKDAYS_REMAIN: 'ยังมีวันทำงานที่ยังไม่ได้เช็คชื่อ กรุณาเช็คชื่อให้ครบก่อนปิดเดือน',\n    EXTRAS_NOT_REVIEWED: 'กรุณาตรวจสอบและยืนยันเงินพิเศษของพนักงานทุกคนก่อนปิดเดือน',\n    ATTENDANCE_AFTER_END_DATE: 'ไม่สามารถบันทึกวันสิ้นสุดการทำงานได้เนื่องจากมีข้อมูลเช็คชื่อหลังจากวันที่ระบุ',\n    ATTENDANCE_ON_HOLIDAY: 'ไม่สามารถเช็คชื่อในวันหยุดตามตารางได้ (กรุณากดบันทึกวันทำงานเพิ่มก่อน)',\n    FUTURE_ATTENDANCE: 'ไม่สามารถเช็คชื่อวันล่วงหน้าได้',\n    PHOTO_TOO_LARGE: 'รูปภาพมีขนาดใหญ่เกินไป (จำกัดไม่เกิน 24 KiB หลังย่อ)',\n    NOT_JPEG: 'ไฟล์ที่ส่งไม่ใช่รูปภาพ JPEG ที่ถูกต้อง',\n    UNKNOWN: 'ยังยืนยันการบันทึกไม่ได้ กรุณากดปุ่มตรวจสอบอีกครั้ง',\n    NO_GOOGLE_SCRIPT_RUN: 'ไม่พบบริการ Google Apps Script (กำลังเปิดในเบราว์เซอร์ธรรมดา)'\n  };\n\n  function announce(msg) {\n    if (liveEl) liveEl.textContent = msg;\n  }\n\n  function formatMoney(satang) {\n    if (typeof satang !== 'number' || !Number.isFinite(satang)) return '0.00';\n    return (satang / 100).toLocaleString('th-TH', {\n      minimumFractionDigits: 2,\n      maximumFractionDigits: 2\n    });\n  }\n\n  function escapeHtml(str) {\n    if (str === null || str === undefined) return '';\n    return String(str).replace(/[&<>\"']/g, c => ({\n      '&': '&amp;', '<': '&lt;', '>': '&gt;', '\"': '&quot;', \"'\": '&#39;'\n    }[c]));\n  }\n\n  function thaiDate(dateStr, options = { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' }) {\n    if (!dateStr) return '';\n    const d = new Date(dateStr + 'T12:00:00+07:00');\n    return new Intl.DateTimeFormat('th-TH', options).format(d);\n  }\n\n  function thaiMonthYear(monthStr) {\n    if (!monthStr) return '';\n    const d = new Date(monthStr + '-01T12:00:00+07:00');\n    return new Intl.DateTimeFormat('th-TH', { month: 'long', year: 'numeric' }).format(d);\n  }\n\n  function uuid() {\n    if (typeof crypto !== 'undefined' && crypto.randomUUID) {\n      return crypto.randomUUID();\n    }\n    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => {\n      const r = Math.random() * 16 | 0;\n      const v = c === 'x' ? r : (r & 0x3 | 0x8);\n      return v.toString(16);\n    });\n  }\n\n  function parseMoneySatang(valStr) {\n    const clean = String(valStr || '').trim();\n    if (!/^(0|[1-9]\\d{0,6})(\\.\\d{1,2})?$/.test(clean)) {\n      throw new Error('กรุณากรอกจำนวนเงินให้ถูกต้อง (ทศนิยมไม่เกิน 2 ตำแหน่ง)');\n    }\n    const parts = clean.split('.');\n    const satang = Number(parts[0]) * 100 + Number(((parts[1] || '') + '00').slice(0, 2));\n    if (!Number.isSafeInteger(satang) || satang > 100000000) {\n      throw new Error('จำนวนเงินสูงเกินกำหนด (สูงสุด 1,000,000.00 บาท)');\n    }\n    return satang;\n  }\n\n  // --- 3. Image Compression Engine (Client-side Canvas) ---\n  async function compressPhotoFile(file) {\n    if (!['image/jpeg', 'image/png', 'image/webp'].includes(file.type)) {\n      throw new Error('กรุณาเลือกไฟล์รูปภาพ JPEG, PNG หรือ WebP');\n    }\n    if (file.size > 5 * 1024 * 1024) {\n      throw new Error('ขนาดไฟล์ต้นฉบับต้องไม่เกิน 5 MB');\n    }\n\n    const objectUrl = URL.createObjectURL(file);\n    try {\n      const img = new Image();\n      img.src = objectUrl;\n      await img.decode();\n\n      if (img.naturalWidth * img.naturalHeight > 20000000) {\n        throw new Error('รูปภาพมีจำนวนพิกเซลมากเกินไป (เกิน 20 ล้านพิกเซล)');\n      }\n\n      // คำนวณขนาดย่อโดยคงสัดส่วนเดิม ด้านยาวสุดไม่เกิน 192px\n      const maxDim = 192;\n      const scale = Math.min(1, maxDim / Math.max(img.naturalWidth, img.naturalHeight));\n      const targetW = Math.max(1, Math.round(img.naturalWidth * scale));\n      const targetH = Math.max(1, Math.round(img.naturalHeight * scale));\n\n      const canvas = document.createElement('canvas');\n      canvas.width = targetW;\n      canvas.height = targetH;\n      const ctx = canvas.getContext('2d');\n\n      // พื้นหลังสีขาวเพื่อรองรับภาพโปร่งใส PNG/WebP ก่อนแปลงเป็น JPEG\n      ctx.fillStyle = '#FFFFFF';\n      ctx.fillRect(0, 0, targetW, targetH);\n      ctx.drawImage(img, 0, 0, targetW, targetH);\n\n      // Re-encode เป็น JPEG โดยปรับ quality จนขนาด base64 <= 32768 ตัวอักษร\n      let finalBase64 = '';\n      const qualities = [0.8, 0.65, 0.5, 0.35];\n      for (const q of qualities) {\n        const dataUrl = canvas.toDataURL('image/jpeg', q);\n        const rawBase64 = dataUrl.split('base64,')[1] || '';\n        if (rawBase64.length <= 32768) {\n          finalBase64 = rawBase64;\n          break;\n        }\n      }\n\n      if (!finalBase64 || finalBase64.length > 32768) {\n        throw new Error('รูปยังมีขนาดใหญ่เกินไป กรุณาเลือกรูปภาพอื่น');\n      }\n\n      return {\n        jpegBase64: finalBase64,\n        dataUrl: 'data:image/jpeg;base64,' + finalBase64,\n        width: targetW,\n        height: targetH\n      };\n    } finally {\n      URL.revokeObjectURL(objectUrl);\n    }\n  }\n\n  // --- 4. Lazy Loading Photos ---\n  async function loadMissingPhotos(employeeIds) {\n    const needed = employeeIds.filter(id => {\n      const emp = state.employees.find(e => e.employeeId === id);\n      const cached = state.photosCache.get(id);\n      return emp && emp.photoVersion > 0 && (!cached || cached.version !== emp.photoVersion);\n    }).slice(0, 8); // โหลดทีละไม่เกิน 8 คนตามข้อกำหนด\n\n    if (needed.length === 0) return;\n\n    try {\n      const known = {};\n      needed.forEach(id => {\n        if (state.photosCache.has(id)) known[id] = state.photosCache.get(id).version;\n      });\n\n      const res = await TeamRpc.call('getPhotos', { employeeIds: needed, knownVersions: known });\n      if (res && res.photos) {\n        Object.keys(res.photos).forEach(empId => {\n          const item = res.photos[empId];\n          state.photosCache.set(empId, {\n            jpegBase64: item.jpegBase64,\n            dataUrl: 'data:image/jpeg;base64,' + item.jpegBase64,\n            version: item.version\n          });\n        });\n        // อัปเดตรูปเฉพาะส่วนที่แสดงอยู่ใน DOM ปัจจุบัน\n        needed.forEach(empId => {\n          const avatarEl = document.querySelector(`[data-avatar-for=\"${empId}\"]`);\n          if (avatarEl && state.photosCache.has(empId)) {\n            const photo = state.photosCache.get(empId);\n            avatarEl.innerHTML = `<img src=\"${photo.dataUrl}\" alt=\"รูปพนักงาน\">`;\n          }\n        });\n      }\n    } catch (e) {\n      // โหลดรูปไม่สำเร็จให้ใช้ชื่อย่อต่อไปโดยไม่ทำให้ระบบหยุดทำงาน\n    }\n  }\n\n  function renderAvatar(emp) {\n    const cached = state.photosCache.get(emp.employeeId);\n    if (cached && cached.dataUrl) {\n      return `<div class=\"team-avatar\" data-avatar-for=\"${emp.employeeId}\"><img src=\"${cached.dataUrl}\" alt=\"รูป${escapeHtml(emp.name)}\"></div>`;\n    }\n    const initial = (emp.nickname || emp.name || '?').slice(0, 1);\n    return `<div class=\"team-avatar\" data-avatar-for=\"${emp.employeeId}\">${escapeHtml(initial)}</div>`;\n  }\n\n  // --- 5. Tab 1: Attendance View (เช็คชื่อ) ---\n  function renderAttendanceView() {\n    const day = state.dayAttendance;\n    if (!day) {\n      contentEl.innerHTML = `<div class=\"team-card\"><p>กำลังโหลดข้อมูลการเช็คชื่อ...</p></div>`;\n      return;\n    }\n\n    const employees = day.employees || [];\n    const attendanceMap = new Map();\n    (day.attendance || []).forEach(r => attendanceMap.set(r.employeeId, r));\n\n    // ตัวกรองการค้นหาและยังไม่เช็ค\n    let displayList = employees;\n    if (state.filterAttendance) {\n      displayList = displayList.filter(e => {\n        const att = attendanceMap.get(e.employeeId);\n        return !att || att.status === 'UNMARKED';\n      });\n    }\n    if (state.searchQuery) {\n      const q = state.searchQuery.toLowerCase().trim();\n      displayList = displayList.filter(e =>\n        e.name.toLowerCase().includes(q) || (e.nickname && e.nickname.toLowerCase().includes(q))\n      );\n    }\n\n    const checkedCount = employees.filter(e => {\n      const a = attendanceMap.get(e.employeeId);\n      return a && a.status && a.status !== 'UNMARKED';\n    }).length;\n\n    const isViewingOtherDay = state.selectedDate !== state.serverToday;\n\n    let html = `\n      <h2>เช็คชื่อ</h2>\n      <div class=\"team-card\">\n        <label for=\"input-att-date\">วันที่\n          <input type=\"date\" id=\"input-att-date\" value=\"${state.selectedDate}\" max=\"${state.serverToday}\">\n        </label>\n        <div style=\"display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 0.5rem;\">\n          <p class=\"team-secondary\" style=\"font-size: 1.15rem; font-weight: 700; color: var(--team-primary-dark);\">\n            ${thaiDate(state.selectedDate)}\n          </p>\n          ${isViewingOtherDay ? `<button type=\"button\" id=\"btn-goto-today\" class=\"team-btn-accent\">ไปที่วันนี้</button>` : ''}\n        </div>\n      </div>\n\n      <div class=\"team-toolbar\">\n        <p class=\"team-secondary\" id=\"att-summary-text\" style=\"font-size: 1.15rem; font-weight: 700;\">\n          เช็คแล้ว ${checkedCount} จาก ${employees.length} คน\n        </p>\n        <button type=\"button\" class=\"team-btn-secondary\" id=\"btn-toggle-filter\" aria-pressed=\"${state.filterAttendance}\">\n          ${state.filterAttendance ? 'แสดงทั้งหมด' : 'แสดงเฉพาะที่ยังไม่เช็ค'}\n        </button>\n      </div>\n    `;\n\n    if (employees.length > 8) {\n      html += `\n        <label for=\"input-att-search\" style=\"margin-bottom: 0.8rem;\">ค้นหาชื่อพนักงาน\n          <input type=\"search\" id=\"input-att-search\" placeholder=\"พิมพ์ชื่อหรือชื่อเล่น...\" value=\"${escapeHtml(state.searchQuery)}\">\n        </label>\n      `;\n    }\n\n    if (!day.isWorkday) {\n      html += `\n        <div class=\"team-notice-box\">\n          <p>วันนี้เป็นวันหยุดตามตาราง</p>\n          <p class=\"team-secondary\" style=\"color: #63480C; margin-top: 0.4rem;\">หากมีพนักงานมาทำงานพิเศษในวันนี้ ให้กดบันทึกวันทำงานเพิ่มเพื่อเริ่มเช็คชื่อ</p>\n          <button type=\"button\" class=\"team-btn-primary\" id=\"btn-add-special-workday\" style=\"margin-top: 0.6rem;\">บันทึกเป็นวันทำงานเพิ่ม</button>\n        </div>\n      `;\n    }\n\n    if (day.isClosed) {\n      html += `\n        <div class=\"team-notice-box team-notice-danger\">\n          <p>เดือนนี้ปิดรอบแล้ว (อ่านอย่างเดียว ไม่สามารถแก้ไขการเช็คชื่อได้)</p>\n        </div>\n      `;\n    }\n\n    if (displayList.length === 0) {\n      if (employees.length === 0) {\n        html += `\n          <div class=\"team-card\" style=\"text-align: center; padding: 2rem 1rem;\">\n            <h3>ยังไม่มีรายชื่อพนักงาน</h3>\n            <p class=\"team-secondary\">กรุณาเพิ่มข้อมูลพนักงานก่อนเริ่มเช็คชื่อ</p>\n            <button type=\"button\" class=\"team-btn-primary\" id=\"btn-empty-add-emp\" style=\"margin-top: 1rem;\">เพิ่มพนักงานคนแรก</button>\n          </div>\n        `;\n      } else {\n        html += `\n          <div class=\"team-card\" style=\"text-align: center; padding: 1.5rem 1rem;\">\n            <p class=\"team-secondary\">เช็คชื่อครบตามตัวกรองแล้ว</p>\n          </div>\n        `;\n      }\n    } else {\n      html += `<div id=\"team-person-cards-container\">`;\n      displayList.forEach(emp => {\n        const att = attendanceMap.get(emp.employeeId);\n        const currentStatus = att ? att.status : 'UNMARKED';\n        const pendingKey = state.selectedDate + '|' + emp.employeeId;\n        const pending = state.pendingCommands.get(pendingKey);\n        const isSaving = pending && pending.status === 'SAVING';\n        const isUnknown = pending && pending.status === 'UNKNOWN';\n\n        let badgeText = 'ยังไม่เช็ค';\n        let badgeClass = '';\n        if (isSaving) {\n          badgeText = 'กำลังบันทึก...';\n          badgeClass = 'status-SAVING';\n        } else if (isUnknown) {\n          badgeText = 'ยังยืนยันการบันทึกไม่ได้';\n          badgeClass = 'status-ABSENT';\n        } else if (currentStatus && currentStatus !== 'UNMARKED') {\n          const timeStr = att && att.updatedAt ? new Date(att.updatedAt).toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' }) : '';\n          badgeText = `${STATUS_LABEL[currentStatus]} · บันทึกแล้ว ${timeStr}`;\n          badgeClass = `status-${currentStatus}`;\n        }\n\n        html += `\n          <article class=\"team-person-card\" data-emp-id=\"${emp.employeeId}\">\n            <div class=\"team-person-head\">\n              ${renderAvatar(emp)}\n              <div class=\"team-person-info\">\n                <div class=\"team-person-name\">${escapeHtml(emp.name)} ${emp.nickname ? `(${escapeHtml(emp.nickname)})` : ''}</div>\n                <div class=\"team-secondary\">${escapeHtml(emp.position)}</div>\n              </div>\n            </div>\n\n            <div class=\"team-statuses\" role=\"group\" aria-label=\"เช็คชื่อ ${escapeHtml(emp.name)}\">\n              <button type=\"button\" class=\"team-status-btn\" data-mark-status=\"FULL\" data-emp-id=\"${emp.employeeId}\"\n                aria-pressed=\"${currentStatus === 'FULL'}\" ${isSaving || day.isClosed ? 'disabled' : ''}>\n                เต็มวัน\n              </button>\n              <button type=\"button\" class=\"team-status-btn\" data-mark-status=\"HALF\" data-emp-id=\"${emp.employeeId}\"\n                aria-pressed=\"${currentStatus === 'HALF'}\" ${isSaving || day.isClosed ? 'disabled' : ''}>\n                ครึ่งวัน\n              </button>\n              <button type=\"button\" class=\"team-status-btn\" data-mark-status=\"ABSENT\" data-emp-id=\"${emp.employeeId}\"\n                aria-pressed=\"${currentStatus === 'ABSENT'}\" ${isSaving || day.isClosed ? 'disabled' : ''}>\n                ไม่มา\n              </button>\n            </div>\n\n            <div class=\"team-saved-row\">\n              <span class=\"team-saved-badge ${badgeClass}\">${badgeText}</span>\n              <div style=\"display: flex; gap: 0.5rem; align-items: center;\">\n                ${isUnknown ? `\n                  <button type=\"button\" class=\"team-btn-accent\" data-retry-emp-id=\"${emp.employeeId}\" style=\"min-height: 2.8rem; padding: 0.3rem 0.6rem;\">\n                    ตรวจสอบอีกครั้ง\n                  </button>\n                ` : ''}\n                ${!day.isClosed && currentStatus !== 'UNMARKED' && !isSaving ? `\n                  <button type=\"button\" class=\"team-btn-text\" data-open-clear=\"${emp.employeeId}\">ล้างรายการ</button>\n                ` : ''}\n              </div>\n            </div>\n          </article>\n        `;\n      });\n      html += `</div>`;\n    }\n\n    contentEl.innerHTML = html;\n    loadMissingPhotos(displayList.map(e => e.employeeId));\n  }\n\n  // --- 6. Tab 2: Reports View (รายงาน) ---\n  function renderReportsView() {\n    if (state.selectedDetailId) {\n      renderEmployeeDetailView(state.selectedDetailId);\n      return;\n    }\n\n    const report = state.monthReport;\n    if (!report) {\n      contentEl.innerHTML = `<div class=\"team-card\"><p>กำลังคำนวณและสรุปรายงาน...</p></div>`;\n      return;\n    }\n\n    const employees = report.employees || [];\n    const isClosed = report.isClosed;\n\n    // ตัวเลือกเดือน: ย้อนหลัง 12 เดือนจนถึงเดือนปัจจุบัน (คำนวณปี-เดือนโดยตรง ปราศจากปัญหา Timezone Shift)\n    const currentMonthKey = (state.serverToday ? state.serverToday.slice(0, 7) : new Date().toISOString().slice(0, 7));\n    const [curYear, curMonthNum] = currentMonthKey.split('-').map(Number);\n    let monthOptions = '';\n    for (let i = 0; i < 12; i++) {\n      let y = curYear;\n      let m = curMonthNum - i;\n      while (m <= 0) {\n        m += 12;\n        y -= 1;\n      }\n      const mKey = `${y}-${String(m).padStart(2, '0')}`;\n      const mText = thaiMonthYear(mKey);\n      monthOptions += `<option value=\"${mKey}\" ${mKey === state.selectedMonth ? 'selected' : ''}>${mText}</option>`;\n    }\n\n    let html = `\n      <h2>รายงานค่าจ้าง</h2>\n      <div class=\"team-card\">\n        <label for=\"select-report-month\">เลือกเดือน\n          <select id=\"select-report-month\">${monthOptions}</select>\n        </label>\n      </div>\n\n      <div class=\"team-summary-box\">\n        <p style=\"font-size: 1.15rem; font-weight: 700;\">${isClosed ? 'ยอดค่าจ้างเดือนนี้ (ปิดรอบแล้ว)' : (state.selectedMonth === currentMonthKey ? `ยอดค่าจ้างสะสมถึง ${thaiDate(state.serverToday, { day: 'numeric', month: 'short' })}` : 'ยอดค่าจ้างรวมเดือนนี้')}</p>\n        <div class=\"team-money-big\">${formatMoney(report.totalSatang)} บาท</div>\n        <p class=\"team-secondary\">\n          ค่าแรงรวม ${formatMoney(report.baseSatang)} บาท + เงินพิเศษรวม ${formatMoney(report.extraSatang)} บาท\n        </p>\n      </div>\n    `;\n\n    if (report.pendingCount > 0) {\n      html += `\n        <div class=\"team-notice-box team-notice-danger\">\n          ยังไม่ได้เช็คชื่ออีก ${report.pendingCount} รายการ (1 คนต่อ 1 วัน = 1 รายการ)\n        </div>\n      `;\n    }\n\n    if (isClosed) {\n      html += `\n        <div class=\"team-notice-box\" style=\"background-color: #F0FDF4; border-color: #86EFAC; color: #166534;\">\n          เดือนนี้ปิดรอบแล้ว (Snapshot เวอร์ชัน ${report.snapshotVersion} · ${thaiDate(report.closedAt ? report.closedAt.slice(0, 10) : '')})\n        </div>\n      `;\n    }\n\n    html += `<div style=\"margin-top: 1rem;\">`;\n    employees.forEach(emp => {\n      html += `\n        <button type=\"button\" class=\"team-report-row\" data-open-detail=\"${emp.employeeId}\">\n          <div class=\"team-report-header\">\n            <div>\n              <h3 style=\"margin: 0; font-size: 1.33rem; color: var(--team-primary-dark);\">${escapeHtml(emp.name)} ${emp.nickname ? `(${escapeHtml(emp.nickname)})` : ''}</h3>\n              <p class=\"team-secondary\" style=\"margin: 0.1rem 0;\">${escapeHtml(emp.position)}</p>\n            </div>\n            <span style=\"font-size: 1.5rem; font-weight: 700; color: var(--team-primary);\" aria-hidden=\"true\">›</span>\n          </div>\n          <div class=\"team-row-money\">${formatMoney(emp.totalSatang)} บาท</div>\n          <div class=\"team-counts-row\">\n            <span>เต็มวัน ${emp.full} วัน</span>\n            <span>ครึ่งวัน ${emp.half} วัน</span>\n            <span>ไม่มา ${emp.absent} วัน</span>\n          </div>\n          ${emp.pending > 0 ? `<p class=\"team-secondary\" style=\"color: #991B1B; font-weight: 700; margin-top: 0.3rem;\">รอเช็ค ${emp.pending} วัน</p>` : ''}\n        </button>\n      `;\n    });\n    html += `</div>`;\n\n    // ปุ่มล่างสุดของรายงาน: ตรวจและปิดเดือน / เปิดเดือน / พิมพ์รายงาน\n    html += `\n      <div style=\"display: grid; gap: 0.8rem; margin-top: 1.5rem;\">\n        ${!isClosed ? `\n          <button type=\"button\" class=\"team-btn-primary team-btn-wide\" id=\"btn-open-close-month-modal\">\n            ตรวจและปิดเดือน\n          </button>\n        ` : `\n          <button type=\"button\" class=\"team-btn-secondary team-btn-wide\" id=\"btn-open-reopen-modal\">\n            เปิดเดือนเพื่อแก้ไข\n          </button>\n        `}\n        <button type=\"button\" class=\"team-btn-secondary team-btn-wide team-print-allow\" id=\"btn-print-report\">\n          พิมพ์รายงานหน้านี้\n        </button>\n      </div>\n    `;\n\n    contentEl.innerHTML = html;\n  }\n\n  // --- 7. Detail View: รายละเอียดเงินเดือนรายบุคคล ---\n  async function renderEmployeeDetailView(employeeId) {\n    contentEl.innerHTML = `<div class=\"team-card\"><p>กำลังโหลดรายละเอียดของพนักงาน...</p></div>`;\n    try {\n      const res = await TeamRpc.call('getEmployeeMonth', employeeId, state.selectedMonth);\n      if (!res || !res.ok) throw new Error('LOAD_FAILED');\n      const d = res.detail;\n\n      let html = `\n        <button type=\"button\" class=\"team-btn-secondary\" id=\"btn-back-to-reports\" style=\"margin-bottom: 1rem;\">\n          ‹ กลับหน้ารายงาน\n        </button>\n\n        <h2>${escapeHtml(d.name)} ${d.nickname ? `(${escapeHtml(d.nickname)})` : ''}</h2>\n        <p class=\"team-secondary\" style=\"font-size: 1.15rem; font-weight: 700; margin-bottom: 1rem;\">\n          ${escapeHtml(d.position)} · ${thaiMonthYear(d.month)}\n        </p>\n\n        <div class=\"team-summary-box\">\n          <p style=\"font-size: 1.15rem; font-weight: 700;\">ยอดค่าจ้างเดือนนี้</p>\n          <div class=\"team-money-big\">${formatMoney(d.totalSatang)} บาท</div>\n          <p class=\"team-secondary\">ค่าแรงรวม ${formatMoney(d.baseSatang)} + เงินพิเศษรวม ${formatMoney(d.extraSatang)}</p>\n        </div>\n\n        <div class=\"team-kpis\">\n          <div class=\"team-kpi-card\">\n            <span class=\"team-secondary\">เต็มวัน</span>\n            <strong>${d.full} วัน</strong>\n          </div>\n          <div class=\"team-kpi-card\">\n            <span class=\"team-secondary\">ครึ่งวัน</span>\n            <strong>${d.half} วัน</strong>\n          </div>\n          <div class=\"team-kpi-card\">\n            <span class=\"team-secondary\">ไม่มา</span>\n            <strong>${d.absent} วัน</strong>\n          </div>\n          <div class=\"team-kpi-card\">\n            <span class=\"team-secondary\">วันคิดค่าจ้าง</span>\n            <strong>${d.paidDayUnits} วัน</strong>\n          </div>\n        </div>\n\n        ${d.pending > 0 ? `\n          <div class=\"team-notice-box team-notice-danger\">\n            ยังมีวันทำงานที่ยังไม่ได้เช็คชื่ออีก ${d.pending} วัน\n          </div>\n        ` : ''}\n\n        <!-- กล่องเงินพิเศษ -->\n        <div class=\"team-card\">\n          <div style=\"display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.8rem;\">\n            <h3 style=\"margin: 0;\">เงินพิเศษประจำเดือนนี้</h3>\n            ${!res.isClosed ? `\n              <button type=\"button\" class=\"team-btn-secondary\" id=\"btn-manage-monthly-extras\" style=\"min-height: 2.8rem; padding: 0.3rem 0.8rem;\">\n                แก้เงินพิเศษเดือนนี้\n              </button>\n            ` : ''}\n          </div>\n          ${(d.extras && d.extras.length > 0) ? d.extras.map(x => `\n            <div class=\"team-line-item\">\n              <span>${escapeHtml(x.label)}</span>\n              <strong>${formatMoney(x.amountSatang)} บาท</strong>\n            </div>\n          `).join('') : '<p class=\"team-secondary\">ไม่มีเงินพิเศษในเดือนนี้</p>'}\n\n          <div style=\"margin-top: 1rem; padding-top: 0.8rem; border-top: 1px solid var(--team-border); display: flex; justify-content: space-between; align-items: center;\">\n            <span class=\"team-secondary\">สถานะการตรวจเงินพิเศษ:</span>\n            <span style=\"font-weight: 700; color: ${d.extrasConfirmed ? '#166534' : '#991B1B'};\">\n              ${d.extrasConfirmed ? 'ตรวจและยืนยันแล้ว' : 'ยังไม่ได้ยืนยัน'}\n            </span>\n          </div>\n          ${!res.isClosed && !d.extrasConfirmed ? `\n            <button type=\"button\" class=\"team-btn-accent team-btn-wide\" id=\"btn-confirm-extras-now\" style=\"margin-top: 0.8rem;\">\n              ยืนยันการตรวจเงินพิเศษคนนี้\n            </button>\n          ` : ''}\n        </div>\n\n        <!-- กล่องอัตราค่าแรง -->\n        <div class=\"team-card\">\n          <h3>อัตราค่าแรงที่ใช้คำนวณ</h3>\n          ${(d.rates && d.rates.length > 0) ? d.rates.map(r => `\n            <div class=\"team-line-item\">\n              <span>เริ่มใช้วันที่ ${thaiDate(r.effectiveFrom, { day: 'numeric', month: 'short', year: 'numeric' })}</span>\n              <strong>${formatMoney(r.dailySatang)} บาท/วัน</strong>\n            </div>\n          `).join('') : '<p class=\"team-secondary\">ไม่มีข้อมูลอัตราค่าแรง</p>'}\n        </div>\n\n        <!-- รายการเช็คชื่อแต่ละวัน -->\n        <div class=\"team-card\">\n          <h3>รายการเช็คชื่อรายวัน</h3>\n          ${(d.days && d.days.length > 0) ? d.days.map(dayRow => `\n            <div class=\"team-line-item\">\n              <div>\n                <div style=\"font-weight: 700;\">${thaiDate(dayRow.dateKey, { day: 'numeric', month: 'short' })}</div>\n                <div class=\"team-secondary\">${STATUS_LABEL[dayRow.status] || dayRow.status}</div>\n              </div>\n              <div>\n                <strong>${dayRow.amountSatang !== null ? formatMoney(dayRow.amountSatang) + ' บาท' : '—'}</strong>\n              </div>\n            </div>\n          `).join('') : '<p class=\"team-secondary\">ไม่มีรายการวันทำงาน</p>'}\n        </div>\n      `;\n\n      contentEl.innerHTML = html;\n    } catch (e) {\n      contentEl.innerHTML = `\n        <div class=\"team-card\">\n          <p style=\"color: #991B1B;\">โหลดรายละเอียดไม่สำเร็จ: ${escapeHtml(ERROR_MESSAGES[e.message] || e.message)}</p>\n          <button type=\"button\" class=\"team-btn-secondary\" id=\"btn-back-to-reports\" style=\"margin-top: 1rem;\">‹ กลับหน้ารายงาน</button>\n        </div>\n      `;\n    }\n  }\n\n  // --- 8. Tab 3: Settings View (ตั้งค่า) ---\n  function renderSettingsView() {\n    if (state.editEmployee !== null) {\n      renderEditEmployeeView();\n      return;\n    }\n\n    const employees = state.employees || [];\n    const activeEmployees = employees.filter(e => !e.endDate);\n    const resignedEmployees = employees.filter(e => Boolean(e.endDate));\n\n    let html = `\n      <h2>ตั้งค่าระบบ</h2>\n      <button type=\"button\" class=\"team-btn-primary team-btn-wide\" id=\"btn-add-new-employee\" style=\"margin-bottom: 1.2rem;\">\n        + เพิ่มพนักงานใหม่\n      </button>\n\n      <!-- รายชื่อพนักงาน -->\n      <div class=\"team-card\">\n        <h3>รายชื่อพนักงานทำงานอยู่ (${activeEmployees.length} คน)</h3>\n        ${activeEmployees.length === 0 ? '<p class=\"team-secondary\">ยังไม่มีพนักงานที่ทำงานอยู่</p>' : ''}\n        ${activeEmployees.map(emp => `\n          <div class=\"team-line-item\" style=\"padding: 0.8rem 0;\">\n            <div style=\"display: flex; align-items: center; gap: 0.8rem;\">\n              ${renderAvatar(emp)}\n              <div>\n                <div style=\"font-weight: 700; font-size: 1.2rem; color: var(--team-primary-dark);\">\n                  ${escapeHtml(emp.name)} ${emp.nickname ? `(${escapeHtml(emp.nickname)})` : ''}\n                </div>\n                <div class=\"team-secondary\">${escapeHtml(emp.position)}</div>\n              </div>\n            </div>\n            <button type=\"button\" class=\"team-btn-secondary\" data-edit-emp=\"${emp.employeeId}\" style=\"min-height: 2.8rem; padding: 0.3rem 0.8rem;\">\n              แก้ไข\n            </button>\n          </div>\n        `).join('')}\n      </div>\n\n      ${resignedEmployees.length > 0 ? `\n        <div class=\"team-card\">\n          <h3 class=\"team-secondary\">พนักงานที่สิ้นสุดการทำงานแล้ว (${resignedEmployees.length} คน)</h3>\n          ${resignedEmployees.map(emp => `\n            <div class=\"team-line-item\" style=\"padding: 0.8rem 0; opacity: 0.8;\">\n              <div>\n                <div style=\"font-weight: 700;\">${escapeHtml(emp.name)} ${emp.nickname ? `(${escapeHtml(emp.nickname)})` : ''}</div>\n                <div class=\"team-secondary\">${escapeHtml(emp.position)} · สิ้นสุด ${thaiDate(emp.endDate, { day: 'numeric', month: 'short', year: 'numeric' })}</div>\n              </div>\n              <button type=\"button\" class=\"team-btn-secondary\" data-edit-emp=\"${emp.employeeId}\" style=\"min-height: 2.8rem; padding: 0.3rem 0.8rem;\">\n                ดูข้อมูล\n              </button>\n            </div>\n          `).join('')}\n        </div>\n      ` : ''}\n\n      <!-- ข้อมูลร้านค้าและวันทำงาน -->\n      <div class=\"team-card\">\n        <h3>ตารางวันทำงานประจำสัปดาห์</h3>\n        <p class=\"team-secondary\">วันทำงานปกติ: จันทร์ – เสาร์ (หยุดวันอาทิตย์)</p>\n      </div>\n\n      <div class=\"team-card\">\n        <h3>ชื่อร้านที่แสดงในระบบ</h3>\n        <p style=\"font-size: 1.2rem; font-weight: 700; color: var(--team-primary);\">${escapeHtml(state.shopName)}</p>\n      </div>\n    `;\n\n    contentEl.innerHTML = html;\n    loadMissingPhotos(employees.map(e => e.employeeId));\n  }\n\n  // --- 9. Full-screen Employee Form (เพิ่ม/แก้ไขพนักงาน) ---\n  function renderEditEmployeeView() {\n    const isNew = state.editEmployee === 'new';\n    const emp = isNew ? null : state.employees.find(e => e.employeeId === state.editEmployee);\n    const draft = state.editFormDraft || {\n      name: emp ? emp.name : '',\n      nickname: emp ? emp.nickname : '',\n      position: emp ? emp.position : '',\n      startDate: emp ? emp.startDate : state.serverToday,\n      endDate: emp ? emp.endDate : '',\n      dailySatang: 50000,\n      dailyWage: '500.00',\n      notes: emp ? emp.notes : '',\n      photoDataUrl: state.photosCache.get(emp ? emp.employeeId : '')?.dataUrl || '',\n      newPhotoBase64: '',\n      newPhotoW: 0,\n      newPhotoH: 0,\n      extraTemplates: []\n    };\n\n    let html = `\n      <button type=\"button\" class=\"team-btn-secondary\" id=\"btn-cancel-edit-emp\" style=\"margin-bottom: 1rem;\">\n        ‹ ยกเลิกและกลับ\n      </button>\n\n      <h2>${isNew ? 'เพิ่มพนักงานใหม่' : 'แก้ไขข้อมูลพนักงาน'}</h2>\n\n      <form id=\"form-employee\" novalidate>\n        <!-- รูปภาพ -->\n        <label>รูปภาพพนักงาน (JPEG, PNG หรือ WebP ขนาดไม่เกิน 5 MB)\n          <input type=\"file\" id=\"input-emp-photo\" accept=\"image/jpeg,image/png,image/webp\">\n        </label>\n        <div id=\"box-photo-preview\" style=\"margin-bottom: 1.2rem;\">\n          ${draft.photoDataUrl ? `<img src=\"${draft.photoDataUrl}\" alt=\"รูปตัวอย่าง\" class=\"team-file-preview\">` : '<p class=\"team-secondary\">ยังไม่ได้เลือกรูปภาพ</p>'}\n        </div>\n\n        <label for=\"emp-name\">ชื่อพนักงาน (จำเป็น)\n          <input type=\"text\" id=\"emp-name\" name=\"name\" required maxlength=\"100\" value=\"${escapeHtml(draft.name)}\">\n        </label>\n\n        <label for=\"emp-nickname\">ชื่อเล่น\n          <input type=\"text\" id=\"emp-nickname\" name=\"nickname\" maxlength=\"40\" value=\"${escapeHtml(draft.nickname)}\">\n        </label>\n\n        <label for=\"emp-position\">ตำแหน่ง (จำเป็น)\n          <input type=\"text\" id=\"emp-position\" name=\"position\" required maxlength=\"80\" value=\"${escapeHtml(draft.position)}\">\n        </label>\n\n        <label for=\"emp-start-date\">วันที่เริ่มงาน (จำเป็น)\n          <input type=\"date\" id=\"emp-start-date\" name=\"startDate\" required value=\"${draft.startDate || state.serverToday}\">\n        </label>\n\n        <label for=\"emp-daily-wage\">ค่าแรงต่อวัน (บาท)\n          <input type=\"text\" id=\"emp-daily-wage\" name=\"dailyWage\" inputmode=\"decimal\" required value=\"${draft.dailyWage || '500.00'}\">\n        </label>\n\n        ${!isNew ? `\n          <label for=\"emp-rate-effective\">เริ่มใช้อัตราค่าแรงนี้ตั้งแต่วันที่\n            <input type=\"date\" id=\"emp-rate-effective\" name=\"rateEffectiveFrom\" value=\"${state.serverToday}\">\n          </label>\n        ` : ''}\n\n        <!-- เงินพิเศษประจำเดือน -->\n        <div class=\"team-card\" style=\"background-color: #F8FAFC; margin: 1.5rem 0;\">\n          <div style=\"display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.8rem;\">\n            <h3 style=\"margin: 0;\">เงินพิเศษประจำเดือน (Template)</h3>\n            <button type=\"button\" class=\"team-btn-secondary\" id=\"btn-add-extra-template-row\" style=\"min-height: 2.8rem; padding: 0.3rem 0.8rem;\">\n              + เพิ่มรายการ\n            </button>\n          </div>\n          <div id=\"extra-templates-list\">\n            ${draft.extraTemplates && draft.extraTemplates.length > 0 ? draft.extraTemplates.map((xt, idx) => `\n              <div class=\"team-card\" style=\"padding: 0.8rem; margin-bottom: 0.6rem;\" data-xt-index=\"${idx}\">\n                <label>ชื่อรายการเงินพิเศษ\n                  <input type=\"text\" class=\"xt-label\" value=\"${escapeHtml(xt.label)}\" placeholder=\"เช่น ค่าเดินทาง, ค่าอาหาร\">\n                </label>\n                <label>จำนวนเงินต่อเดือน (บาท)\n                  <input type=\"text\" class=\"xt-amount\" inputmode=\"decimal\" value=\"${(xt.amountSatang / 100).toFixed(2)}\">\n                </label>\n                <button type=\"button\" class=\"team-btn-danger\" data-remove-xt=\"${idx}\" style=\"min-height: 2.8rem; width: 100%; margin-top: 0.4rem;\">\n                  ลบรายการนี้\n                </button>\n              </div>\n            `).join('') : '<p class=\"team-secondary\">ไม่มีเงินพิเศษประจำ</p>'}\n          </div>\n        </div>\n\n        <label for=\"emp-notes\">รายละเอียดเพิ่มเติม / หมายเหตุ\n          <textarea id=\"emp-notes\" name=\"notes\" maxlength=\"500\">${escapeHtml(draft.notes)}</textarea>\n        </label>\n\n        ${!isNew && !draft.endDate ? `\n          <div style=\"margin: 1.5rem 0; padding: 1rem; border: 1px solid #FCA5A5; border-radius: var(--team-radius-card); background: #FEF2F2;\">\n            <h3 style=\"color: #991B1B;\">สิ้นสุดการทำงาน</h3>\n            <p class=\"team-secondary\" style=\"color: #7F1D1D;\">เมื่อพนักงานลาออกหรือสิ้นสุดการจ้าง ให้กำหนดวันสิ้นสุดการทำงาน (ประวัติการทำงานเดิมจะยังคงอยู่ครบถ้วน)</p>\n            <button type=\"button\" class=\"team-btn-danger team-btn-wide\" id=\"btn-prompt-end-employment\" style=\"margin-top: 0.8rem;\">\n              กำหนดวันสิ้นสุดการทำงาน\n            </button>\n          </div>\n        ` : ''}\n\n        <div style=\"display: grid; gap: 0.8rem; margin-top: 1.5rem;\">\n          <button type=\"submit\" class=\"team-btn-primary team-btn-wide\" id=\"btn-submit-employee\">\n            บันทึกข้อมูล\n          </button>\n          <button type=\"button\" class=\"team-btn-secondary team-btn-wide\" id=\"btn-cancel-employee\">\n            ยกเลิก\n          </button>\n        </div>\n      </form>\n    `;\n\n    contentEl.innerHTML = html;\n  }\n\n  // --- 10. Modals & Sheet Overlays ---\n  function renderClearModal(employeeId) {\n    const emp = state.employees.find(e => e.employeeId === employeeId);\n    const modalHtml = `\n      <div class=\"team-sheet-overlay\" id=\"modal-clear-overlay\">\n        <div class=\"team-sheet-container\" role=\"dialog\" aria-modal=\"true\" aria-labelledby=\"modal-clear-title\">\n          <h2 id=\"modal-clear-title\" style=\"color: #991B1B;\">ยืนยันล้างการเช็คชื่อ</h2>\n          <p>คุณต้องการล้างการเช็คชื่อของ <strong>${escapeHtml(emp ? emp.name : '')}</strong> ในวันที่ <strong>${thaiDate(state.selectedDate)}</strong> ใช่หรือไม่?</p>\n          <p class=\"team-secondary\">สถานะจะกลับเป็น \"ยังไม่เช็ค\" โดยประวัติการแก้ไขเดิมจะยังถูกบันทึกไว้อย่างปลอดภัย</p>\n          <div style=\"display: grid; gap: 0.8rem; margin-top: 1.5rem;\">\n            <button type=\"button\" class=\"team-btn-danger team-btn-wide\" id=\"btn-confirm-clear-action\" data-emp-id=\"${employeeId}\">\n              ยืนยันการล้างข้อมูล\n            </button>\n            <button type=\"button\" class=\"team-btn-secondary team-btn-wide\" id=\"btn-close-modal\">\n              ยกเลิก\n            </button>\n          </div>\n        </div>\n      </div>\n    `;\n    document.body.insertAdjacentHTML('beforeend', modalHtml);\n  }\n\n  function renderCloseMonthModal() {\n    const report = state.monthReport;\n    const canClose = report && report.pendingCount === 0;\n\n    const modalHtml = `\n      <div class=\"team-sheet-overlay\" id=\"modal-close-overlay\">\n        <div class=\"team-sheet-container\" role=\"dialog\" aria-modal=\"true\" aria-labelledby=\"modal-close-title\">\n          <h2 id=\"modal-close-title\">ตรวจและปิดรอบเดือน ${thaiMonthYear(state.selectedMonth)}</h2>\n          <div class=\"team-card\" style=\"margin: 1rem 0;\">\n            <div class=\"team-line-item\">\n              <span>ยอดค่าจ้างรวมทั้งสิ้น:</span>\n              <strong style=\"font-size: 1.3rem; color: var(--team-primary);\">${report ? formatMoney(report.totalSatang) : '0.00'} บาท</strong>\n            </div>\n            <div class=\"team-line-item\">\n              <span>จำนวนวันทำงานที่ยังไม่เช็ค:</span>\n              <strong style=\"color: ${report && report.pendingCount > 0 ? '#991B1B' : '#166534'};\">\n                ${report ? report.pendingCount : 0} วัน\n              </strong>\n            </div>\n          </div>\n\n          ${!canClose ? `\n            <div class=\"team-notice-box team-notice-danger\">\n              ไม่สามารถปิดเดือนได้เนื่องจากยังมีวันทำงานที่ยังไม่ได้เช็คชื่อ กรุณาเช็คชื่อให้ครบก่อนปิดเดือน\n            </div>\n          ` : `\n            <p class=\"team-secondary\">เมื่อปิดเดือนแล้ว ระบบจะสร้าง Snapshot ถาวรและล็อกข้อมูลเดือนนี้เพื่อป้องกันการเปลี่ยนแปลงโดยไม่ตั้งใจ</p>\n          `}\n\n          <div style=\"display: grid; gap: 0.8rem; margin-top: 1.5rem;\">\n            ${canClose ? `\n              <button type=\"button\" class=\"team-btn-primary team-btn-wide\" id=\"btn-execute-close-month\">\n                ยืนยันการปิดเดือน\n              </button>\n            ` : ''}\n            <button type=\"button\" class=\"team-btn-secondary team-btn-wide\" id=\"btn-close-modal\">\n              ปิดหน้าต่าง\n            </button>\n          </div>\n        </div>\n      </div>\n    `;\n    document.body.insertAdjacentHTML('beforeend', modalHtml);\n  }\n\n  function renderReopenMonthModal() {\n    const modalHtml = `\n      <div class=\"team-sheet-overlay\" id=\"modal-reopen-overlay\">\n        <div class=\"team-sheet-container\" role=\"dialog\" aria-modal=\"true\" aria-labelledby=\"modal-reopen-title\">\n          <h2 id=\"modal-reopen-title\">เปิดเดือนเพื่อแก้ไข</h2>\n          <p>คุณกำลังจะเปิดรอบเดือน <strong>${thaiMonthYear(state.selectedMonth)}</strong> เพื่อทำการแก้ไขข้อมูล</p>\n          <label for=\"input-reopen-reason\">กรุณาระบุเหตุผลในการเปิดเดือน (จำเป็น)\n            <textarea id=\"input-reopen-reason\" required maxlength=\"200\" placeholder=\"เช่น แก้ไขรายการเช็คชื่อตกหล่นของช่างสมชาย...\"></textarea>\n          </label>\n          <div style=\"display: grid; gap: 0.8rem; margin-top: 1.5rem;\">\n            <button type=\"button\" class=\"team-btn-danger team-btn-wide\" id=\"btn-execute-reopen-month\">\n              ยืนยันการเปิดเดือน\n            </button>\n            <button type=\"button\" class=\"team-btn-secondary team-btn-wide\" id=\"btn-close-modal\">\n              ยกเลิก\n            </button>\n          </div>\n        </div>\n      </div>\n    `;\n    document.body.insertAdjacentHTML('beforeend', modalHtml);\n  }\n\n  function renderMonthlyExtrasManageModal(employeeId) {\n    const emp = state.employees.find(e => e.employeeId === employeeId);\n    const modalHtml = `\n      <div class=\"team-sheet-overlay\" id=\"modal-extras-overlay\">\n        <div class=\"team-sheet-container\" role=\"dialog\" aria-modal=\"true\">\n          <h2>แก้ไขเงินพิเศษเดือน ${thaiMonthYear(state.selectedMonth)}</h2>\n          <p class=\"team-secondary\">สำหรับ: <strong>${escapeHtml(emp ? emp.name : '')}</strong></p>\n\n          <form id=\"form-manage-month-extra\" style=\"margin-top: 1rem;\">\n            <label for=\"input-me-label\">ชื่อรายการเงินพิเศษ\n              <input type=\"text\" id=\"input-me-label\" required maxlength=\"80\" placeholder=\"เช่น ค่าเดินทางพิเศษ, เงินรางวัล\">\n            </label>\n            <label for=\"input-me-amount\">จำนวนเงิน (บาท)\n              <input type=\"text\" id=\"input-me-amount\" inputmode=\"decimal\" required placeholder=\"0.00\">\n            </label>\n            <div style=\"display: grid; gap: 0.8rem; margin-top: 1.2rem;\">\n              <button type=\"submit\" class=\"team-btn-primary team-btn-wide\">\n                เพิ่มรายการเงินพิเศษเดือนนี้\n              </button>\n              <button type=\"button\" class=\"team-btn-secondary team-btn-wide\" id=\"btn-close-modal\">\n                เสร็จสิ้น / ปิด\n              </button>\n            </div>\n          </form>\n        </div>\n      </div>\n    `;\n    document.body.insertAdjacentHTML('beforeend', modalHtml);\n  }\n\n  function closeModal() {\n    const overlays = document.querySelectorAll('.team-sheet-overlay');\n    overlays.forEach(el => el.remove());\n    state.activeModal = null;\n  }\n\n  // --- 11. Core Navigation & Data Fetching ---\n  function switchTab(newTab) {\n    if (state.editEmployee !== null) {\n      if (!confirm('คุณมีข้อมูลที่ยังไม่ได้บันทึก ต้องการออกจากหน้านี้ใช่หรือไม่?')) return;\n      state.editEmployee = null;\n      state.editFormDraft = null;\n    }\n    state.tab = newTab;\n    state.selectedDetailId = null;\n\n    document.querySelectorAll('.team-nav-btn').forEach(btn => {\n      if (btn.dataset.tab === newTab) {\n        btn.setAttribute('aria-current', 'page');\n      } else {\n        btn.removeAttribute('aria-current');\n      }\n    });\n\n    if (newTab === 'attendance') {\n      renderAttendanceView();\n    } else if (newTab === 'reports') {\n      loadMonthReport(state.selectedMonth);\n    } else if (newTab === 'settings') {\n      renderSettingsView();\n    }\n  }\n\n  async function loadDayAttendance(dateStr) {\n    state.selectedDate = dateStr;\n    renderAttendanceView();\n    try {\n      const res = await TeamRpc.call('getDay', dateStr);\n      if (state.selectedDate === dateStr) {\n        state.dayAttendance = res;\n        renderAttendanceView();\n      }\n    } catch (e) {\n      announce('โหลดข้อมูลเช็คชื่อไม่สำเร็จ: ' + (ERROR_MESSAGES[e.message] || e.message));\n    }\n  }\n\n  async function loadMonthReport(monthKey) {\n    if (!monthKey || typeof monthKey !== 'string' || !/^\\d{4}-\\d{2}$/.test(monthKey)) {\n      monthKey = (state.serverToday || new Date().toISOString().slice(0, 10)).slice(0, 7);\n    }\n    state.selectedMonth = monthKey;\n    renderReportsView();\n    try {\n      const res = await TeamRpc.call('getMonthReport', monthKey);\n      if (state.selectedMonth === monthKey) {\n        state.monthReport = res;\n        renderReportsView();\n      }\n    } catch (e) {\n      announce('โหลดรายงานไม่สำเร็จ: ' + (ERROR_MESSAGES[e.message] || e.message));\n    }\n  }\n\n  // --- 12. Mutation Handlers ---\n  async function handleMarkAttendance(employeeId, newStatus) {\n    const key = state.selectedDate + '|' + employeeId;\n    const day = state.dayAttendance;\n    if (!day) return;\n\n    const currentRecord = (day.attendance || []).find(r => r.employeeId === employeeId);\n    const expectedRevision = currentRecord ? currentRecord.revision : 0;\n\n    let command = state.pendingCommands.get(key)?.command;\n    if (!command) {\n      command = {\n        dateKey: state.selectedDate,\n        employeeId: employeeId,\n        status: newStatus,\n        expectedRevision: expectedRevision,\n        requestId: uuid()\n      };\n    } else {\n      command.status = newStatus;\n      command.expectedRevision = expectedRevision;\n    }\n\n    state.pendingCommands.set(key, { command, status: 'SAVING', timestamp: Date.now() });\n    renderAttendanceView();\n    announce(`กำลังบันทึก ${STATUS_LABEL[newStatus]}...`);\n\n    try {\n      const res = await TeamRpc.call('saveAttendance', command);\n      if (res && res.ok && res.attendance) {\n        state.pendingCommands.delete(key);\n        // อัปเดตข้อมูล attendance ใน local state\n        const attList = day.attendance || [];\n        const idx = attList.findIndex(r => r.employeeId === employeeId);\n        if (idx >= 0) {\n          attList[idx] = res.attendance;\n        } else {\n          attList.push(res.attendance);\n        }\n        day.attendance = attList;\n        renderAttendanceView();\n        announce(`บันทึก ${STATUS_LABEL[newStatus]} เรียบร้อยแล้ว`);\n      }\n    } catch (e) {\n      if (e.message === 'UNKNOWN') {\n        state.pendingCommands.set(key, { command, status: 'UNKNOWN', timestamp: Date.now() });\n        renderAttendanceView();\n        announce('ยังยืนยันผลการบันทึกไม่ได้ กรุณากดตรวจสอบอีกครั้ง');\n      } else {\n        state.pendingCommands.delete(key);\n        renderAttendanceView();\n        alert('บันทึกไม่สำเร็จ: ' + (ERROR_MESSAGES[e.message] || e.message));\n        // ถ้าเกิด conflict ให้โหลดข้อมูลวันล่าสุดใหม่\n        if (e.message === 'CONFLICT') loadDayAttendance(state.selectedDate);\n      }\n    }\n  }\n\n  async function handleClearAttendance(employeeId) {\n    closeModal();\n    await handleMarkAttendance(employeeId, 'UNMARKED');\n  }\n\n  // --- 13. Event Listeners Delegation ---\n  document.addEventListener('click', async event => {\n    const btn = event.target.closest('button');\n    if (!btn) return;\n\n    // แท็บนำทาง\n    if (btn.dataset.tab) {\n      switchTab(btn.dataset.tab);\n      return;\n    }\n\n    // ปุ่มเช็คชื่อ (เต็มวัน / ครึ่งวัน / ไม่มา)\n    if (btn.dataset.markStatus) {\n      const empId = btn.dataset.empId;\n      const status = btn.dataset.markStatus;\n      await handleMarkAttendance(empId, status);\n      return;\n    }\n\n    // ปุ่มลองตรวจสอบอีกครั้งหลัง timeout\n    if (btn.dataset.retryEmpId) {\n      const empId = btn.dataset.retryEmpId;\n      const key = state.selectedDate + '|' + empId;\n      const pending = state.pendingCommands.get(key);\n      if (pending && pending.command) {\n        pending.status = 'SAVING';\n        renderAttendanceView();\n        try {\n          const res = await TeamRpc.call('saveAttendance', pending.command);\n          if (res && res.ok && res.attendance) {\n            state.pendingCommands.delete(key);\n            const attList = state.dayAttendance.attendance || [];\n            const idx = attList.findIndex(r => r.employeeId === empId);\n            if (idx >= 0) attList[idx] = res.attendance;\n            else attList.push(res.attendance);\n            renderAttendanceView();\n            announce('ตรวจสอบและยืนยันการบันทึกเรียบร้อยแล้ว');\n          }\n        } catch (e) {\n          alert('ตรวจสอบไม่สำเร็จ: ' + (ERROR_MESSAGES[e.message] || e.message));\n        }\n      }\n      return;\n    }\n\n    // เปิดโมดอลล้างรายการ\n    if (btn.dataset.openClear) {\n      renderClearModal(btn.dataset.openClear);\n      return;\n    }\n\n    if (btn.id === 'btn-confirm-clear-action') {\n      await handleClearAttendance(btn.dataset.empId);\n      return;\n    }\n\n    if (btn.id === 'btn-close-modal') {\n      closeModal();\n      return;\n    }\n\n    if (btn.id === 'btn-goto-today') {\n      loadDayAttendance(state.serverToday);\n      return;\n    }\n\n    if (btn.id === 'btn-toggle-filter') {\n      state.filterAttendance = !state.filterAttendance;\n      renderAttendanceView();\n      return;\n    }\n\n    if (btn.id === 'btn-add-special-workday') {\n      if (!confirm('ยืนยันกำหนดให้วันนี้เป็นวันทำงานพิเศษหรือไม่?')) return;\n      try {\n        await TeamRpc.call('saveCalendar', {\n          action: 'SET_OVERRIDE',\n          dateKey: state.selectedDate,\n          kind: 'WORKDAY',\n          note: 'วันทำงานพิเศษ',\n          expectedRevision: 0,\n          requestId: uuid()\n        });\n        loadDayAttendance(state.selectedDate);\n      } catch (e) {\n        alert('บันทึกไม่สำเร็จ: ' + (ERROR_MESSAGES[e.message] || e.message));\n      }\n      return;\n    }\n\n    // รายงาน: เปิดดูรายบุคคล\n    if (btn.dataset.openDetail) {\n      state.selectedDetailId = btn.dataset.openDetail;\n      renderEmployeeDetailView(state.selectedDetailId);\n      return;\n    }\n\n    if (btn.id === 'btn-back-to-reports') {\n      state.selectedDetailId = null;\n      renderReportsView();\n      return;\n    }\n\n    if (btn.id === 'btn-open-close-month-modal') {\n      renderCloseMonthModal();\n      return;\n    }\n\n    if (btn.id === 'btn-execute-close-month') {\n      closeModal();\n      try {\n        announce('กำลังคำนวณและบันทึกปิดรอบเดือน...');\n        const res = await TeamRpc.call('closeMonth', {\n          monthKey: state.selectedMonth,\n          expectedRevision: state.monthReport.monthState ? state.monthReport.monthState.revision : 1,\n          requestId: uuid()\n        });\n        alert('ปิดรอบเดือนเรียบร้อยแล้ว');\n        loadMonthReport(state.selectedMonth);\n      } catch (e) {\n        alert('ปิดเดือนไม่สำเร็จ: ' + (ERROR_MESSAGES[e.message] || e.message));\n      }\n      return;\n    }\n\n    if (btn.id === 'btn-open-reopen-modal') {\n      renderReopenMonthModal();\n      return;\n    }\n\n    if (btn.id === 'btn-execute-reopen-month') {\n      const reasonInput = document.getElementById('input-reopen-reason');\n      const reason = reasonInput ? reasonInput.value.trim() : '';\n      if (!reason) {\n        alert('กรุณาระบุเหตุผลในการเปิดเดือน');\n        return;\n      }\n      closeModal();\n      try {\n        announce('กำลังเปิดรอบเดือน...');\n        await TeamRpc.call('reopenMonth', {\n          monthKey: state.selectedMonth,\n          reason: reason,\n          expectedRevision: state.monthReport.monthState ? state.monthReport.monthState.revision : 1,\n          requestId: uuid()\n        });\n        alert('เปิดรอบเดือนเรียบร้อยแล้ว ตอนนี้สามารถแก้ไขข้อมูลได้');\n        loadMonthReport(state.selectedMonth);\n      } catch (e) {\n        alert('เปิดเดือนไม่สำเร็จ: ' + (ERROR_MESSAGES[e.message] || e.message));\n      }\n      return;\n    }\n\n    if (btn.id === 'btn-print-report') {\n      window.print();\n      return;\n    }\n\n    if (btn.id === 'btn-manage-monthly-extras') {\n      renderMonthlyExtrasManageModal(state.selectedDetailId);\n      return;\n    }\n\n    if (btn.id === 'btn-confirm-extras-now') {\n      try {\n        await TeamRpc.call('confirmMonthExtras', {\n          employeeId: state.selectedDetailId,\n          monthKey: state.selectedMonth,\n          requestId: uuid()\n        });\n        alert('ยืนยันการตรวจเงินพิเศษเรียบร้อยแล้ว');\n        renderEmployeeDetailView(state.selectedDetailId);\n      } catch (e) {\n        alert('ยืนยันไม่สำเร็จ: ' + (ERROR_MESSAGES[e.message] || e.message));\n      }\n      return;\n    }\n\n    // ตั้งค่า: เพิ่ม/แก้ไขพนักงาน\n    if (btn.id === 'btn-add-new-employee' || btn.id === 'btn-empty-add-emp') {\n      state.editEmployee = 'new';\n      state.editFormDraft = null;\n      renderEditEmployeeView();\n      return;\n    }\n\n    if (btn.dataset.editEmp) {\n      state.editEmployee = btn.dataset.editEmp;\n      state.editFormDraft = null;\n      renderEditEmployeeView();\n      return;\n    }\n\n    if (btn.id === 'btn-cancel-edit-emp' || btn.id === 'btn-cancel-employee') {\n      state.editEmployee = null;\n      state.editFormDraft = null;\n      renderSettingsView();\n      return;\n    }\n\n    if (btn.id === 'btn-add-extra-template-row') {\n      if (!state.editFormDraft) {\n        const emp = state.employees.find(e => e.employeeId === state.editEmployee);\n        state.editFormDraft = { extraTemplates: [] };\n      }\n      if (!state.editFormDraft.extraTemplates) state.editFormDraft.extraTemplates = [];\n      state.editFormDraft.extraTemplates.push({ label: '', amountSatang: 100000 });\n      renderEditEmployeeView();\n      return;\n    }\n\n    if (btn.dataset.removeXt !== undefined) {\n      const idx = Number(btn.dataset.removeXt);\n      if (state.editFormDraft && state.editFormDraft.extraTemplates) {\n        state.editFormDraft.extraTemplates.splice(idx, 1);\n        renderEditEmployeeView();\n      }\n      return;\n    }\n\n    if (btn.id === 'btn-prompt-end-employment') {\n      const endVal = prompt('กรุณาระบุวันสิ้นสุดการทำงาน (YYYY-MM-DD):', state.serverToday);\n      if (!endVal) return;\n      if (!/^20\\d{2}-\\d{2}-\\d{2}$/.test(endVal)) {\n        alert('รูปแบบวันที่ไม่ถูกต้อง ต้องเป็น YYYY-MM-DD');\n        return;\n      }\n      if (!confirm(`ยืนยันการสิ้นสุดการทำงานของพนักงาน ณ วันที่ ${endVal} ใช่หรือไม่?`)) return;\n\n      const emp = state.employees.find(e => e.employeeId === state.editEmployee);\n      try {\n        await TeamRpc.call('setEmploymentEnd', {\n          employeeId: emp.employeeId,\n          endDate: endVal,\n          expectedRevision: emp.revision,\n          requestId: uuid()\n        });\n        alert('บันทึกการสิ้นสุดการทำงานเรียบร้อยแล้ว');\n        state.editEmployee = null;\n        initApp();\n      } catch (e) {\n        alert('เกิดข้อผิดพลาด: ' + (ERROR_MESSAGES[e.message] || e.message));\n      }\n      return;\n    }\n  });\n\n  // Event: Change\n  document.addEventListener('change', async event => {\n    const target = event.target;\n\n    // เปลี่ยนวันที่เช็คชื่อ\n    if (target.id === 'input-att-date') {\n      if (!target.value || target.value > state.serverToday) {\n        target.value = state.selectedDate;\n        return;\n      }\n      loadDayAttendance(target.value);\n      return;\n    }\n\n    // เปลี่ยนเดือนรายงาน\n    if (target.id === 'select-report-month') {\n      loadMonthReport(target.value);\n      return;\n    }\n\n    // อัปโหลดรูปภาพพนักงาน\n    if (target.id === 'input-emp-photo') {\n      const file = target.files[0];\n      if (!file) return;\n      try {\n        announce('กำลังย่อขนาดรูปภาพ...');\n        const compressed = await compressPhotoFile(file);\n        if (!state.editFormDraft) state.editFormDraft = {};\n        state.editFormDraft.photoDataUrl = compressed.dataUrl;\n        state.editFormDraft.newPhotoBase64 = compressed.jpegBase64;\n        state.editFormDraft.newPhotoW = compressed.width;\n        state.editFormDraft.newPhotoH = compressed.height;\n\n        const previewBox = document.getElementById('box-photo-preview');\n        if (previewBox) {\n          previewBox.innerHTML = `<img src=\"${compressed.dataUrl}\" alt=\"รูปตัวอย่าง\" class=\"team-file-preview\">`;\n        }\n        announce('ย่อขนาดรูปภาพสำเร็จพร้อมบันทึก');\n      } catch (e) {\n        alert(e.message);\n        target.value = '';\n      }\n      return;\n    }\n  });\n\n  // Event: Input (ค้นหาชื่อ)\n  document.addEventListener('input', event => {\n    if (event.target.id === 'input-att-search') {\n      state.searchQuery = event.target.value;\n      renderAttendanceView();\n    }\n  });\n\n  // Event: Submit Forms\n  document.addEventListener('submit', async event => {\n    event.preventDefault();\n\n    // บันทึกฟอร์มพนักงาน\n    if (event.target.id === 'form-employee') {\n      const form = event.target;\n      const name = form.name.value.trim();\n      const position = form.position.value.trim();\n      const nickname = form.nickname.value.trim();\n      const startDate = form.startDate.value;\n      const dailyWage = form.dailyWage.value.trim();\n      const notes = form.notes.value.trim();\n      const rateEffective = form.rateEffectiveFrom ? form.rateEffectiveFrom.value : startDate;\n\n      if (!name || !position || !startDate || !dailyWage) {\n        alert('กรุณากรอกข้อมูลที่จำเป็นให้ครบถ้วน');\n        return;\n      }\n\n      let dailySatang = 0;\n      try {\n        dailySatang = parseMoneySatang(dailyWage);\n      } catch (e) {\n        alert(e.message);\n        return;\n      }\n\n      // รวบรวม ExtraTemplates\n      const xtRows = document.querySelectorAll('#extra-templates-list [data-xt-index]');\n      const extraTemplates = [];\n      for (const row of xtRows) {\n        const lbl = row.querySelector('.xt-label').value.trim();\n        const amtStr = row.querySelector('.xt-amount').value.trim();\n        if (lbl) {\n          try {\n            const satang = parseMoneySatang(amtStr);\n            extraTemplates.push({ label: lbl, amountSatang: satang });\n          } catch (e) {\n            alert(`เงินพิเศษ \"${lbl}\": ${e.message}`);\n            return;\n          }\n        }\n      }\n\n      const isNew = state.editEmployee === 'new';\n      const emp = isNew ? null : state.employees.find(e => e.employeeId === state.editEmployee);\n\n      const payload = {\n        employeeId: isNew ? '' : emp.employeeId,\n        name: name,\n        nickname: nickname,\n        position: position,\n        startDate: startDate,\n        dailySatang: dailySatang,\n        rateEffectiveFrom: rateEffective,\n        extraTemplates: extraTemplates,\n        notes: notes,\n        expectedRevision: isNew ? 0 : emp.revision,\n        requestId: uuid()\n      };\n\n      try {\n        announce('กำลังบันทึกข้อมูลพนักงาน...');\n        const res = await TeamRpc.call('saveEmployee', payload);\n        if (res && res.ok && res.employee) {\n          // หากมีรูปภาพใหม่ที่ย่อแล้ว ให้บันทึกรูปภาพด้วย\n          if (state.editFormDraft && state.editFormDraft.newPhotoBase64) {\n            await TeamRpc.call('savePhoto', {\n              employeeId: res.employee.employeeId,\n              jpegBase64: state.editFormDraft.newPhotoBase64,\n              width: state.editFormDraft.newPhotoW,\n              height: state.editFormDraft.newPhotoH,\n              expectedRevision: 0,\n              requestId: uuid()\n            });\n            state.photosCache.set(res.employee.employeeId, {\n              jpegBase64: state.editFormDraft.newPhotoBase64,\n              dataUrl: state.editFormDraft.photoDataUrl,\n              version: res.employee.photoVersion + 1\n            });\n          }\n\n          alert('บันทึกข้อมูลพนักงานเรียบร้อยแล้ว');\n          state.editEmployee = null;\n          state.editFormDraft = null;\n          await initApp();\n        }\n      } catch (e) {\n        alert('บันทึกไม่สำเร็จ: ' + (ERROR_MESSAGES[e.message] || e.message));\n      }\n      return;\n    }\n\n    // บันทึกเงินพิเศษประจำเดือนใน Detail View\n    if (event.target.id === 'form-manage-month-extra') {\n      const lbl = document.getElementById('input-me-label').value.trim();\n      const amtStr = document.getElementById('input-me-amount').value.trim();\n      if (!lbl) return;\n\n      let satang = 0;\n      try {\n        satang = parseMoneySatang(amtStr);\n      } catch (e) {\n        alert(e.message);\n        return;\n      }\n\n      closeModal();\n      try {\n        announce('กำลังบันทึกเงินพิเศษ...');\n        await TeamRpc.call('saveMonthExtra', {\n          monthKey: state.selectedMonth,\n          employeeId: state.selectedDetailId,\n          label: lbl,\n          amountSatang: satang,\n          expectedRevision: 0,\n          requestId: uuid()\n        });\n        alert('เพิ่มรายการเงินพิเศษเดือนนี้เรียบร้อยแล้ว');\n        renderEmployeeDetailView(state.selectedDetailId);\n      } catch (e) {\n        alert('เกิดข้อผิดพลาด: ' + (ERROR_MESSAGES[e.message] || e.message));\n      }\n      return;\n    }\n  });\n\n  // --- 14. Initial Bootstrap ---\n  async function initApp() {\n    try {\n      announce('กำลังเชื่อมต่อระบบ...');\n      const bootstrap = await TeamRpc.call('getBootstrap');\n      if (!bootstrap) throw new Error('NO_DATA');\n\n      state.serverToday = bootstrap.serverToday || new Date().toISOString().slice(0, 10);\n      state.selectedDate = state.serverToday;\n      state.selectedMonth = bootstrap.monthKey || state.serverToday.slice(0, 7);\n      state.shopName = bootstrap.shopName || 'DE TEAM';\n      state.employees = bootstrap.employees || [];\n      state.dayAttendance = bootstrap.dayAttendance || null;\n\n      if (brandTitleEl) brandTitleEl.innerHTML = `${escapeHtml(state.shopName.slice(0, 2))} <span>${escapeHtml(state.shopName.slice(2))}</span>`;\n      if (headerInfoEl) headerInfoEl.textContent = `เช็คชื่อและสรุปค่าจ้าง (${thaiDate(state.serverToday, { day: 'numeric', month: 'short' })})`;\n\n      // เริ่มต้นที่แท็บเช็คชื่อ\n      switchTab('attendance');\n      announce('พร้อมใช้งาน');\n    } catch (e) {\n      contentEl.innerHTML = `\n        <div class=\"team-card\" style=\"padding: 1.5rem; border: 1.5px solid #FCA5A5;\">\n          <h2 style=\"color: #991B1B;\">เชื่อมต่อระบบไม่สำเร็จ</h2>\n          <p class=\"team-secondary\">${escapeHtml(ERROR_MESSAGES[e.message] || e.message)}</p>\n          <div style=\"margin-top: 1.2rem;\">\n            <button type=\"button\" class=\"team-btn-primary\" onclick=\"location.reload();\">ลองเชื่อมต่อใหม่อีกครั้ง</button>\n          </div>\n        </div>\n      `;\n      announce('การเชื่อมต่อล้มเหลว');\n    }\n  }\n\n  // --- 15. Auto Date Rollover (อัปเดตวันใหม่อัตโนมัติเมื่อข้ามวันหรือเปิดจอกลับมา) ---\n  async function checkDateRollover() {\n    try {\n      const now = new Date();\n      const localToday = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate())}`;\n      if (state.serverToday && localToday > state.serverToday) {\n        const wasToday = state.selectedDate === state.serverToday;\n        state.serverToday = localToday;\n        state.selectedMonth = localToday.slice(0, 7);\n        if (wasToday) {\n          state.selectedDate = localToday;\n        }\n        if (headerInfoEl) {\n          headerInfoEl.textContent = `เช็คชื่อและสรุปค่าจ้าง (${thaiDate(state.serverToday, { day: 'numeric', month: 'short' })})`;\n        }\n        if (state.tab === 'attendance') {\n          await loadDayAttendance(state.selectedDate);\n        }\n      }\n    } catch (e) {\n      // ละเว้นหากเกิดข้อผิดพลาดในการตรวจสอบวันใหม่\n    }\n  }\n\n  document.addEventListener('visibilitychange', () => {\n    if (document.visibilityState === 'visible') {\n      checkDateRollover();\n    }\n  });\n\n  window.addEventListener('focus', checkDateRollover);\n  setInterval(checkDateRollover, 300000);\n\n  // เริ่มรันเมื่อโหลด DOM เสร็จ\n  if (document.readyState === 'loading') {\n    document.addEventListener('DOMContentLoaded', initApp);\n  } else {\n    initApp();\n  }\n})();\n</script>\n\n</body>\n</html>\n";
}
