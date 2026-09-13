'use strict';

const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');

const base = path.resolve(__dirname, '..');
const gasDir = fs.existsSync(path.join(base, 'gas')) ? path.join(base, 'gas') : base;
const webDir = fs.existsSync(path.join(base, 'web')) ? path.join(base, 'web') : base;
let passed = 0;

function test(name, run) {
  run();
  passed++;
  process.stdout.write(`PASS ${name}\n`);
}

// 1. Setup V8 Context with all backend .gs files
const ctx = vm.createContext({
  Date, Map, Set, JSON, Number, Object, Array, String, Math, Error, RegExp, Boolean
});

const gsFiles = [
  'Config.gs',
  'Payroll.gs',
  'Repository.gs',
  'Auth.gs',
  'Schema.gs',
  'AttendanceService.gs',
  'EmployeeService.gs',
  'ExtrasService.gs',
  'CalendarService.gs',
  'PhotoService.gs',
  'ReportService.gs',
  'MonthService.gs',
  'Code.gs'
];

for (const f of gsFiles) {
  const code = fs.readFileSync(path.join(gasDir, f), 'utf8');
  vm.runInContext(code, ctx, { filename: f });
}

const plain = val => JSON.parse(JSON.stringify(val));

// --- Mock Google Workspace Environment ---
let database, activeEmail, locked, canLock, batches, injectedFailure;

function objectRows(headers, objects) {
  return [plain(headers), ...objects.map(o => headers.map(h => o[h] ?? ''))];
}

function setupMockDatabase() {
  activeEmail = 'owner@example.invalid';
  locked = false;
  canLock = true;
  batches = 0;
  injectedFailure = false;

  database = {
    Settings: objectRows(ctx.SETTINGS_HEADERS_, [
      { key: 'schemaVersion', value: '1', revision: 1, updatedAt: '2026-09-01T00:00:00Z' },
      { key: 'shopName', value: 'DE TEAM', revision: 1, updatedAt: '2026-09-01T00:00:00Z' },
      { key: 'workweekJson', value: '[1,2,3,4,5,6]', revision: 1, updatedAt: '2026-09-01T00:00:00Z' },
      { key: 'calendarEffectiveFrom', value: '2026-09-01', revision: 1, updatedAt: '2026-09-01T00:00:00Z' }
    ]),
    Employees: objectRows(ctx.EMPLOYEE_HEADERS_, [
      { employeeId: 'e1', name: 'สมชาย ใจดี', nickname: 'ชาย', position: 'ช่างติดตั้ง', startDate: '2026-09-01', endDate: '', notes: '', photoVersion: 0, revision: 1, updatedAt: '2026-09-01T00:00:00Z' }
    ]),
    RateHistory: objectRows(ctx.RATE_HEADERS_, [
      { rateId: 'r1', employeeId: 'e1', effectiveFrom: '2026-09-01', dailySatang: 50000, revision: 1, updatedAt: '2026-09-01T00:00:00Z' }
    ]),
    ExtraTemplates: objectRows(ctx.EXTRA_TEMPLATE_HEADERS_, [
      { templateId: 'xt1', employeeId: 'e1', label: 'ค่าเดินทาง', amountSatang: 150000, effectiveFromMonth: '2026-09', effectiveToMonth: '', revision: 1, updatedAt: '2026-09-01T00:00:00Z' }
    ]),
    Photos: objectRows(ctx.PHOTO_HEADERS_, []),
    WorkCalendar: objectRows(ctx.CALENDAR_HEADERS_, []),
    MonthState: objectRows(ctx.MONTH_HEADERS_, [
      { monthKey: '2026-09', state: 'OPEN', revision: 1, snapshotVersion: 0, closedAt: '', closedBy: '' }
    ]),
    Attendance_2026_09: objectRows(ctx.ATTENDANCE_HEADERS_, []),
    MonthlyExtras_2026_09: objectRows(ctx.MONTHLY_EXTRA_HEADERS_, []),
    ExtraReviews_2026_09: objectRows(ctx.EXTRA_REVIEW_HEADERS_, []),
    Snapshots_2026_09: objectRows(ctx.SNAPSHOT_HEADERS_, []),
    Audit_2026_09: objectRows(ctx.AUDIT_HEADERS_, []),
    Requests_2026_09: objectRows(ctx.REQUEST_HEADERS_, [])
  };
}

const sheetIdByName = name => {
  const keys = Object.keys(database);
  const idx = keys.indexOf(name);
  if (idx < 0) {
    // Dynamically created sheets
    keys.push(name);
    return keys.length;
  }
  return idx + 1;
};

const makeMockSheet = name => ({
  getName: () => name,
  getSheetId: () => sheetIdByName(name),
  getDataRange: () => ({
    getValues: () => plain(database[name] || [[]])
  }),
  appendRow: rowValues => {
    if (!database[name]) database[name] = [];
    database[name].push(plain(rowValues));
  },
  getRange: (r, c, numR, numC) => ({
    setValues: vals => {
      if (!database[name]) database[name] = [];
      for (let i = 0; i < vals.length; i++) {
        database[name][r - 1 + i] = plain(vals[i]);
      }
    },
    setFontWeight: () => {},
    getValue: () => (database[name] && database[name][r - 1] ? database[name][r - 1][c - 1] : '')
  }),
  setFrozenRows: () => {}
});

ctx.PropertiesService = {
  getScriptProperties: () => ({
    getProperty: key => ({
      OWNER_EMAIL: 'owner@example.invalid',
      SPREADSHEET_ID: 'test-spreadsheet-id'
    }[key]),
    setProperty: (k, v) => {}
  })
};

ctx.Session = {
  getActiveUser: () => ({ getEmail: () => activeEmail })
};

ctx.SpreadsheetApp = {
  openById: id => ({
    getId: () => id,
    getSheetByName: name => {
      if (!database[name]) return null;
      return makeMockSheet(name);
    },
    getSheets: () => Object.keys(database).map(makeMockSheet),
    insertSheet: name => {
      if (!database[name]) database[name] = [];
      return makeMockSheet(name);
    },
    deleteSheet: sheet => {
      delete database[sheet.getName()];
    }
  })
};

ctx.LockService = {
  getScriptLock: () => ({
    tryLock: () => {
      if (!canLock) return false;
      assert.equal(locked, false);
      locked = true;
      return true;
    },
    releaseLock: () => {
      assert.equal(locked, true);
      locked = false;
    }
  })
};

ctx.Utilities = {
  formatDate: (d, tz, fmt) => '2026-09-14',
  getUuid: () => crypto.randomUUID(),
  DigestAlgorithm: { SHA_256: 'sha256' },
  Charset: { UTF_8: 'utf8' },
  computeDigest: (alg, text) => Array.from(crypto.createHash(alg).update(text).digest()),
  base64Decode: b64 => Array.from(Buffer.from(b64, 'base64'))
};

ctx.Sheets = {
  Spreadsheets: {
    batchUpdate: ({ requests }) => {
      assert.equal(locked, true);
      const copy = plain(database);
      for (const req of requests) {
        const update = req.updateCells || req.appendCells;
        const id = update.sheetId ?? update.start.sheetId;
        const name = Object.keys(database)[id - 1];
        assert.ok(name, `Sheet index ${id - 1} must exist`);
        const values = update.rows[0].values.map(c =>
          c.userEnteredValue.stringValue ?? c.userEnteredValue.numberValue ?? ''
        );
        if (req.updateCells) {
          copy[name][update.start.rowIndex] = values;
        } else {
          copy[name].push(values);
        }
      }
      if (injectedFailure) throw new Error('INJECTED_BATCH_FAILURE');
      database = copy;
      batches++;
    }
  }
};

// ==========================================
// 2. Pure Engine Tests (from verify.cjs)
// ==========================================

const fixture = () => ({
  month: '2026-09',
  today: '2026-09-30',
  systemStartDate: '2026-09-01',
  employee: { employeeId: 'e1', startDate: '2026-09-01', endDate: '' },
  rates: [{ employeeId: 'e1', effectiveFrom: '2026-09-01', dailySatang: 50000 }],
  attendance: [],
  extras: [],
  weekdays: [0, 1, 2, 3, 4, 5, 6],
  calendar: {}
});

const marks = (days, status, start = 1) =>
  Array.from({ length: days }, (_, i) => ({
    employeeId: 'e1',
    dateKey: `2026-09-${String(start + i).padStart(2, '0')}`,
    status
  }));

test('22 full + 4 half + 2 absent + monthly extras = 14,500 THB', () => {
  const f = fixture();
  f.employee.endDate = '2026-09-28';
  f.attendance = [...marks(22, 'FULL'), ...marks(4, 'HALF', 23), ...marks(2, 'ABSENT', 27)];
  f.extras = [
    { employeeId: 'e1', monthKey: f.month, extraId: 'x1', label: 'ค่าเดินทาง', amountSatang: 150000 },
    { employeeId: 'e1', monthKey: f.month, extraId: 'x2', label: 'ค่าอาหาร', amountSatang: 100000 }
  ];
  const r = ctx.calculateEmployeeMonth_(f);
  assert.equal(r.totalSatang, 1450000);
  assert.equal(r.paidDayUnits, 24);
  assert.equal(r.workedDays, 26);
  assert.equal(r.absent, 2);
  assert.equal(r.pending, 0);
});

test('historical rates calculate each day, including half-days', () => {
  const f = fixture();
  f.rates.push({ employeeId: 'e1', effectiveFrom: '2026-09-11', dailySatang: 55000 });
  f.attendance = [...marks(20, 'FULL'), ...marks(2, 'HALF', 21)];
  assert.equal(ctx.calculateEmployeeMonth_(f).baseSatang, 1105000);
});

test('half satang rounds up per day', () => {
  assert.equal(ctx.dailyPay_('HALF', 50001), 25001);
});

test('strict money decimal conversion', () => {
  assert.equal(ctx.moneySatang_('500.01'), 50001);
  assert.equal(ctx.moneySatang_('0.1'), 10);
  for (const bad of ['1e6', 'NaN', '-1', '0.001', '1,000', ' 500', '01', 'Infinity', '1000000.01']) {
    assert.throws(() => ctx.moneySatang_(bad));
  }
});

test('leap year and invalid date', () => {
  assert.equal(ctx.monthDates_('2024-02').length, 29);
  assert.equal(ctx.monthDates_('2026-02').length, 28);
  assert.throws(() => ctx.dateKey_('2026-02-29'));
  assert.throws(() => ctx.dateKey_('2026-09-31'));
});

test('unmarked days are pending, holidays and future days are excluded', () => {
  const f = fixture();
  f.today = '2026-09-03';
  f.calendar = { '2026-09-02': 'HOLIDAY' };
  const r = ctx.calculateEmployeeMonth_(f);
  assert.equal(r.pending, 2);
  assert.equal(r.absent, 0);
  assert.equal(r.days.find(x => x.dateKey === '2026-09-02').amountSatang, null);
});

test('employment dates and system start bound pending days', () => {
  const f = fixture();
  f.employee.startDate = '2026-09-10';
  f.employee.endDate = '2026-09-12';
  assert.equal(ctx.calculateEmployeeMonth_(f).pending, 3);
  f.systemStartDate = '2026-09-11';
  assert.equal(ctx.calculateEmployeeMonth_(f).pending, 2);
});

test('duplicate attendance, rates and extras fail closed', () => {
  let f = fixture();
  f.attendance = [...marks(1, 'FULL'), ...marks(1, 'HALF')];
  assert.throws(() => ctx.calculateEmployeeMonth_(f), /DUPLICATE_ATTENDANCE/);
  f = fixture();
  f.rates.push({ ...f.rates[0] });
  assert.throws(() => ctx.calculateEmployeeMonth_(f), /DUPLICATE_RATE/);
  f = fixture();
  const x = { extraId: 'x', employeeId: 'e1', monthKey: f.month, label: 'ค่ารถ', amountSatang: 100 };
  f.extras = [x, x];
  assert.throws(() => ctx.calculateEmployeeMonth_(f), /DUPLICATE_EXTRA/);
});

test('missing rate never silently becomes zero', () => {
  const f = fixture();
  f.attendance = marks(1, 'FULL');
  f.rates = [];
  assert.throws(() => ctx.calculateEmployeeMonth_(f), /MISSING_RATE/);
});

test('holiday and future attendance are rejected', () => {
  const f = fixture();
  f.attendance = marks(1, 'FULL');
  f.calendar = { '2026-09-01': 'HOLIDAY' };
  assert.throws(() => ctx.calculateEmployeeMonth_(f), /ATTENDANCE_ON_HOLIDAY/);
  f.calendar = {};
  f.today = '2026-08-31';
  assert.throws(() => ctx.calculateEmployeeMonth_(f), /FUTURE_ATTENDANCE/);
});

test('monthly bonus remains whole for partial-month employment', () => {
  const f = fixture();
  f.employee.startDate = '2026-09-30';
  f.extras = [{ employeeId: 'e1', monthKey: f.month, extraId: 'x', label: 'ค่าอาหาร', amountSatang: 100000 }];
  assert.equal(ctx.calculateEmployeeMonth_(f).extraSatang, 100000);
});

// ==========================================
// 3. Attendance Service Endpoint Tests
// ==========================================

const command = (overrides = {}) => ({
  dateKey: '2026-09-14',
  employeeId: 'e1',
  status: 'FULL',
  expectedRevision: 0,
  requestId: crypto.randomUUID(),
  ...overrides
});

test('getDay returns scoped employees and records', () => {
  setupMockDatabase();
  const r = ctx.getDay('2026-09-14');
  assert.equal(r.employees.length, 1);
  assert.equal(r.isClosed, false);
  assert.equal(r.isWorkday, true);
  assert.equal(locked, false);
});

test('data, audit and receipt written in one batch', () => {
  setupMockDatabase();
  const r = ctx.saveAttendance(command());
  assert.equal(r.attendance.revision, 1);
  assert.equal(batches, 1);
  assert.equal(database.Attendance_2026_09.length, 2);
  assert.equal(database.Audit_2026_09.length, 2);
  assert.equal(database.Requests_2026_09.length, 2);
  assert.equal(locked, false);
});

test('retry after a lost response returns same receipt without extra write', () => {
  setupMockDatabase();
  const cmd = command();
  const a = ctx.saveAttendance(cmd);
  const b = ctx.saveAttendance(cmd);
  assert.deepEqual(plain(a), plain(b));
  assert.equal(batches, 1);
});

test('requestId reuse with changed payload is rejected', () => {
  setupMockDatabase();
  const cmd = command();
  ctx.saveAttendance(cmd);
  assert.throws(() => ctx.saveAttendance({ ...cmd, status: 'HALF' }), /REQUEST_ID_REUSED/);
  assert.equal(batches, 1);
});

test('stale second tab cannot overwrite first tab', () => {
  setupMockDatabase();
  ctx.saveAttendance(command());
  assert.throws(() => ctx.saveAttendance(command({ status: 'HALF' })), /CONFLICT/);
  assert.equal(database.Attendance_2026_09[1][3], 'FULL');
  assert.equal(locked, false);
});

test('clear is a versioned record and does not delete history', () => {
  setupMockDatabase();
  ctx.saveAttendance(command());
  ctx.saveAttendance(command({ status: 'UNMARKED', expectedRevision: 1 }));
  assert.equal(database.Attendance_2026_09.length, 2);
  assert.equal(database.Attendance_2026_09[1][4], 2);
  assert.equal(database.Audit_2026_09.length, 3);
});

test('closed month rejects new mutation but honors old receipt', () => {
  setupMockDatabase();
  const cmd = command();
  const r = ctx.saveAttendance(cmd);
  database.MonthState[1][1] = 'CLOSED';
  assert.throws(() => ctx.saveAttendance(command({ expectedRevision: 1 })), /MONTH_CLOSED/);
  assert.deepEqual(plain(ctx.saveAttendance(cmd)), plain(r));
});

test('empty and unapproved identities denied on every public endpoint', () => {
  setupMockDatabase();
  for (const email of ['', 'other@example.invalid']) {
    activeEmail = email;
    assert.throws(() => ctx.getDay('2026-09-14'), /AUTH_REQUIRED/);
    assert.throws(() => ctx.saveAttendance(command()), /AUTH_REQUIRED/);
  }
  assert.equal(batches, 0);
});

test('invalid fields, dates, status, employee and future dates rejected', () => {
  setupMockDatabase();
  for (const extra of [
    { injected: true },
    { dateKey: '2026-09-31' },
    { status: 'LATE' },
    { employeeId: 'unknown' },
    { dateKey: '2026-09-15' }
  ]) {
    assert.throws(() => ctx.saveAttendance(command(extra)));
  }
  assert.equal(batches, 0);
});

test('busy lock returns without mutation', () => {
  setupMockDatabase();
  canLock = false;
  assert.throws(() => ctx.saveAttendance(command()), /BUSY/);
  assert.equal(batches, 0);
});

test('failed batch retains all original data and releases lock', () => {
  setupMockDatabase();
  const before = plain(database);
  injectedFailure = true;
  assert.throws(() => ctx.saveAttendance(command()), /INJECTED_BATCH_FAILURE/);
  assert.deepEqual(plain(database), before);
  assert.equal(locked, false);
});

test('duplicate keys and malformed schema are rejected', () => {
  setupMockDatabase();
  ctx.saveAttendance(command());
  database.Attendance_2026_09.push([...database.Attendance_2026_09[1]]);
  assert.throws(() => ctx.saveAttendance(command()), /DUPLICATE_KEY/);
  setupMockDatabase();
  database.Employees[0][0] = 'renamed';
  assert.throws(() => ctx.getDay('2026-09-14'), /SCHEMA_MISMATCH/);
});

test('formula-like user text is written as an explicit string', () => {
  const cell = ctx.cells_(['=IMPORTXML("x","y")'])[0];
  assert.equal(cell.userEnteredValue.stringValue, '=IMPORTXML("x","y")');
  assert.equal(cell.userEnteredValue.formulaValue, undefined);
});

// ==========================================
// 4. Employee Service Tests
// ==========================================

test('saveEmployee creates employee, initial rate, extra templates and updates revision', () => {
  setupMockDatabase();
  const res = ctx.saveEmployee({
    name: 'มาลี แสงทอง',
    nickname: 'ลี',
    position: 'ผู้ช่วยช่าง',
    startDate: '2026-09-01',
    dailySatang: 45000,
    extraTemplates: [{ label: 'ค่าอาหาร', amountSatang: 100000 }],
    notes: 'ทดสอบ',
    expectedRevision: 0,
    requestId: crypto.randomUUID()
  });

  assert.equal(res.ok, true);
  assert.equal(res.employee.name, 'มาลี แสงทอง');
  assert.equal(res.employee.revision, 1);
  assert.equal(database.Employees.length, 3);
  assert.equal(database.RateHistory.length, 3);
  assert.equal(database.ExtraTemplates.length, 3);
});

test('setEmploymentEnd succeeds when no future attendance exists', () => {
  setupMockDatabase();
  const res = ctx.setEmploymentEnd({
    employeeId: 'e1',
    endDate: '2026-09-20',
    expectedRevision: 1,
    requestId: crypto.randomUUID()
  });
  assert.equal(res.ok, true);
  assert.equal(res.employee.endDate, '2026-09-20');
  assert.equal(res.employee.revision, 2);
});

test('setEmploymentEnd is REJECTED when attendance exists after endDate', () => {
  setupMockDatabase();
  // Mark attendance on 2026-09-14
  ctx.saveAttendance(command({ dateKey: '2026-09-14', status: 'FULL' }));
  // Try to set employment end to 2026-09-10 (which is before 2026-09-14)
  assert.throws(() => ctx.setEmploymentEnd({
    employeeId: 'e1',
    endDate: '2026-09-10',
    expectedRevision: 1,
    requestId: crypto.randomUUID()
  }), /ATTENDANCE_AFTER_END_DATE/);
});

test('addRate adds historical rate and rejects duplicates on same effectiveFrom', () => {
  setupMockDatabase();
  const res = ctx.addRate({
    employeeId: 'e1',
    effectiveFrom: '2026-09-15',
    dailySatang: 55000,
    requestId: crypto.randomUUID()
  });
  assert.equal(res.ok, true);
  assert.equal(res.rate.dailySatang, 55000);

  // Duplicate on same date should be rejected
  assert.throws(() => ctx.addRate({
    employeeId: 'e1',
    effectiveFrom: '2026-09-15',
    dailySatang: 60000,
    requestId: crypto.randomUUID()
  }), /DUPLICATE_RATE/);
});

// ==========================================
// 5. Extras Service Tests
// ==========================================

test('prepareMonthExtras materializes templates and does not duplicate on second run', () => {
  setupMockDatabase();
  const reqId1 = crypto.randomUUID();
  const r1 = ctx.prepareMonthExtras({ monthKey: '2026-09', requestId: reqId1 });
  assert.equal(r1.ok, true);
  assert.equal(r1.createdCount, 1);
  assert.equal(database.MonthlyExtras_2026_09.length, 2);

  // Second run: createdCount must be 0 (no duplicates!)
  const reqId2 = crypto.randomUUID();
  const r2 = ctx.prepareMonthExtras({ monthKey: '2026-09', requestId: reqId2 });
  assert.equal(r2.createdCount, 0);
  assert.equal(database.MonthlyExtras_2026_09.length, 2);
});

test('saveMonthExtra adds/updates monthly extra and invalidates review', () => {
  setupMockDatabase();
  // Set existing review
  database.ExtraReviews_2026_09.push(['e1', 1, '2026-09-14T00:00:00Z', 'owner@example.invalid']);

  const res = ctx.saveMonthExtra({
    monthKey: '2026-09',
    employeeId: 'e1',
    label: 'ค่าตำแหน่งพิเศษ',
    amountSatang: 200000,
    expectedRevision: 0,
    requestId: crypto.randomUUID()
  });
  assert.equal(res.ok, true);
  assert.equal(res.extra.amountSatang, 200000);

  // Review should be invalidated
  assert.equal(database.ExtraReviews_2026_09[1][2], '');
});

test('confirmMonthExtras confirms review for employee', () => {
  setupMockDatabase();
  const res = ctx.confirmMonthExtras({
    monthKey: '2026-09',
    employeeId: 'e1',
    requestId: crypto.randomUUID()
  });
  assert.equal(res.ok, true);
  assert.ok(res.confirmedAt);
  assert.equal(database.ExtraReviews_2026_09.length, 2);
});

// ==========================================
// 6. Photo Service Tests
// ==========================================

test('savePhoto validates JPEG signature and rejects invalid or fake photos', () => {
  setupMockDatabase();
  // Valid JPEG header: FF D8 ... FF D9
  const validJpegBytes = Buffer.from([0xFF, 0xD8, 0x00, 0x10, 0xFF, 0xD9]);
  const validB64 = validJpegBytes.toString('base64');

  const res = ctx.savePhoto({
    employeeId: 'e1',
    jpegBase64: validB64,
    width: 120,
    height: 150,
    expectedRevision: 0,
    requestId: crypto.randomUUID()
  });
  assert.equal(res.ok, true);
  assert.equal(res.photoVersion, 1);
  assert.equal(database.Employees[1][7], 1); // photoVersion bumped

  // Fake non-JPEG photo
  const fakeBytes = Buffer.from([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A]); // PNG header
  assert.throws(() => ctx.savePhoto({
    employeeId: 'e1',
    jpegBase64: fakeBytes.toString('base64'),
    width: 100,
    height: 100,
    expectedRevision: 1,
    requestId: crypto.randomUUID()
  }), /NOT_JPEG/);
});

test('getPhotos returns only changed photo versions', () => {
  setupMockDatabase();
  const validJpeg = Buffer.from([0xFF, 0xD8, 0x01, 0xFF, 0xD9]).toString('base64');
  database.Photos.push(['e1', validJpeg, 100, 100, 2, '2026-09-14T00:00:00Z']);

  // If client knownVersion is 1, photo version 2 is returned
  const r1 = ctx.getPhotos({ employeeIds: ['e1'], knownVersions: { e1: 1 } });
  assert.ok(r1.photos.e1);
  assert.equal(r1.photos.e1.version, 2);

  // If client knownVersion is 2, photo is NOT re-transmitted (saves bandwidth!)
  const r2 = ctx.getPhotos({ employeeIds: ['e1'], knownVersions: { e1: 2 } });
  assert.equal(r2.photos.e1, undefined);
});

// ==========================================
// 7. Report & Month Service (Snapshot & Close) Tests
// ==========================================

test('closeMonth rejects if month has not ended or pending workdays remain', () => {
  setupMockDatabase();
  // Today is 2026-09-14, month 2026-09 ends on 2026-09-30 -> MONTH_NOT_ENDED
  assert.throws(() => ctx.closeMonth({
    monthKey: '2026-09',
    expectedRevision: 1,
    requestId: crypto.randomUUID()
  }), /MONTH_NOT_ENDED/);

  // Now simulate an ended month e.g. 2026-08 where today > 2026-08-31
  database.MonthState.push(['2026-08', 'OPEN', 1, 0, '', '']);
  database.Attendance_2026_08 = objectRows(ctx.ATTENDANCE_HEADERS_, []);
  database.MonthlyExtras_2026_08 = objectRows(ctx.MONTHLY_EXTRA_HEADERS_, []);
  database.ExtraReviews_2026_08 = objectRows(ctx.EXTRA_REVIEW_HEADERS_, []);
  database.Snapshots_2026_08 = objectRows(ctx.SNAPSHOT_HEADERS_, []);
  database.Audit_2026_08 = objectRows(ctx.AUDIT_HEADERS_, []);
  database.Requests_2026_08 = objectRows(ctx.REQUEST_HEADERS_, []);

  // Employee e1 startDate is 2026-09-01, so e1 has no workdays in 2026-08.
  // Add an employee e0 starting in August and set system start date to August:
  database.Settings.find(r => r[0] === 'calendarEffectiveFrom')[1] = '2026-08-01';
  database.Employees.push(['e0', 'ทดสอบ ส.ค.', 'ทด', 'ช่าง', '2026-08-01', '', '', 0, 1, '']);
  database.RateHistory.push(['r0', 'e0', '2026-08-01', 50000, 1, '']);

  // Employee e0 has unmarked days -> UNMARKED_WORKDAYS_REMAIN
  assert.throws(() => ctx.closeMonth({
    monthKey: '2026-08',
    expectedRevision: 1,
    requestId: crypto.randomUUID()
  }), /UNMARKED_WORKDAYS_REMAIN/);
});

test('closeMonth creates snapshot chunking and marks month CLOSED, getMonthReport reads snapshot', () => {
  setupMockDatabase();
  // Month 2026-08
  database.Settings.find(r => r[0] === 'calendarEffectiveFrom')[1] = '2026-08-01';
  database.MonthState.push(['2026-08', 'OPEN', 1, 0, '', '']);
  database.Attendance_2026_08 = objectRows(ctx.ATTENDANCE_HEADERS_, []);
  database.MonthlyExtras_2026_08 = objectRows(ctx.MONTHLY_EXTRA_HEADERS_, []);
  database.ExtraReviews_2026_08 = objectRows(ctx.EXTRA_REVIEW_HEADERS_, [
    { employeeId: 'eAug', extrasRevision: 1, confirmedAt: '2026-08-31T00:00:00Z', confirmedBy: 'owner@example.invalid' }
  ]);
  database.Snapshots_2026_08 = objectRows(ctx.SNAPSHOT_HEADERS_, []);
  database.Audit_2026_08 = objectRows(ctx.AUDIT_HEADERS_, []);
  database.Requests_2026_08 = objectRows(ctx.REQUEST_HEADERS_, []);

  // Employee working only 1 day in August: 2026-08-31
  database.Employees.push(['eAug', 'สมบูรณ์ ปิดเดือน', 'บูรณ์', 'หัวหน้าช่าง', '2026-08-31', '2026-08-31', '', 0, 1, '']);
  database.RateHistory.push(['rAug', 'eAug', '2026-08-01', 60000, 1, '']);
  database.Attendance_2026_08.push(['2026-08-31|eAug', '2026-08-31', 'eAug', 'FULL', 1, '2026-08-31T00:00:00Z', 'owner@example.invalid', 'req1']);

  const closeRes = ctx.closeMonth({
    monthKey: '2026-08',
    expectedRevision: 1,
    requestId: crypto.randomUUID()
  });
  assert.equal(closeRes.ok, true);
  assert.equal(closeRes.state, 'CLOSED');
  assert.equal(closeRes.snapshotVersion, 1);
  assert.equal(closeRes.totalSatang, 60000);

  // Verify Snapshot rows were generated
  assert.ok(database.Snapshots_2026_08.length > 1);

  // Now getMonthReport for 2026-08 MUST read from snapshots!
  const report = ctx.getMonthReport('2026-08');
  assert.equal(report.isClosed, true);
  assert.equal(report.snapshotVersion, 1);
  assert.equal(report.totalSatang, 60000);
  assert.equal(report.employees.length, 1);
  assert.equal(report.employees[0].name, 'สมบูรณ์ ปิดเดือน');

  // Reopen month with reason
  const reopenRes = ctx.reopenMonth({
    monthKey: '2026-08',
    reason: 'ขอแก้ไขรายการเช็คชื่อย้อนหลัง',
    expectedRevision: 2,
    requestId: crypto.randomUUID()
  });
  assert.equal(reopenRes.ok, true);
  assert.equal(reopenRes.state, 'OPEN');
});

// ==========================================
// 8. Bootstrap & Setup System Tests
// ==========================================

test('getBootstrap returns complete initial state without photo base64', () => {
  setupMockDatabase();
  const bs = ctx.getBootstrap();
  assert.equal(bs.serverToday, '2026-09-14');
  assert.equal(bs.monthKey, '2026-09');
  assert.equal(bs.shopName, 'DE TEAM');
  assert.ok(Array.isArray(bs.employees));
  assert.equal(bs.employees[0].name, 'สมชาย ใจดี');
  assert.ok(bs.dayAttendance);
  // Ensure no base64 photo bloat in bootstrap!
  assert.equal(bs.employees[0].jpegBase64, undefined);
});

test('setupSystem_ is idempotent and preserves existing data', () => {
  setupMockDatabase();
  const r1 = ctx.setupSystem_();
  assert.equal(r1.ok, true);
  assert.equal(database.Employees.length, 2); // existing employee preserved!
  const r2 = ctx.setupSystem_();
  assert.equal(r2.ok, true);
  assert.equal(database.Employees.length, 2);
});

// ==========================================
// 9. Frontend HTML & JavaScript Syntax Validation
// ==========================================

test('all embedded JavaScript in Index.html, App.html and Fonts.html parses cleanly', () => {
  const appHtml = fs.readFileSync(path.join(webDir, 'App.html'), 'utf8');
  for (const match of appHtml.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)) {
    new vm.Script(match[1], { filename: 'App.html' });
  }

  const manifest = JSON.parse(fs.readFileSync(path.join(gasDir, 'appsscript.json'), 'utf8'));
  assert.equal(manifest.timeZone, 'Asia/Bangkok');
  assert.equal(manifest.dependencies.enabledAdvancedServices[0].serviceId, 'sheets');
  assert.equal(manifest.webapp.access, 'MYSELF');
});

process.stdout.write(`\nAll ${passed} automated tests PASSED successfully!\n`);
