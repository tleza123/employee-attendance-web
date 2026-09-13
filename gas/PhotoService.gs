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
