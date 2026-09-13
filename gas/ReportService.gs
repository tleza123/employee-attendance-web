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
