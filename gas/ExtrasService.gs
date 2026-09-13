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
