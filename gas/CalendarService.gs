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
