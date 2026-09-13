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
