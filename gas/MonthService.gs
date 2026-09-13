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
