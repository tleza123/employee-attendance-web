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
