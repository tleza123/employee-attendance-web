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
