/**
 * DE TEAM — ระบบเช็คชื่อพนักงานและคำนวณยอดค่าจ้าง (All-in-One Backend Code)
 * รวมโค้ดระบบหลังบ้านและสคริปต์ติดตั้งไว้ในไฟล์เดียวสำหรับวางใน Code.gs
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
  requireValue_(typeof value === 'string' && /^20\d{2}-\d{2}-\d{2}$/.test(value), 'INVALID_DATE');
  var parsed = new Date(value + 'T00:00:00Z');
  requireValue_(Number.isFinite(parsed.getTime()) && parsed.toISOString().slice(0, 10) === value, 'INVALID_DATE');
  return value;
}

function monthKey_(value) {
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
  var template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
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
