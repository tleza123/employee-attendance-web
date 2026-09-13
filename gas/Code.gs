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
