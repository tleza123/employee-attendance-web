/**
 * DE TEAM — ระบบเช็คชื่อพนักงานและคำนวณยอดค่าจ้าง
 * Config.gs: ค่าคงที่, โครงสร้าง Headers และการเข้าถึง Script Properties
 */

var SETTINGS_HEADERS_ = ['key', 'value', 'revision', 'updatedAt'];
var EMPLOYEE_HEADERS_ = ['employeeId', 'name', 'nickname', 'position', 'startDate', 'endDate', 'notes', 'photoVersion', 'revision', 'updatedAt'];
var RATE_HEADERS_ = ['rateId', 'employeeId', 'effectiveFrom', 'dailySatang', 'revision', 'updatedAt'];
var EXTRA_TEMPLATE_HEADERS_ = ['templateId', 'employeeId', 'label', 'amountSatang', 'effectiveFromMonth', 'effectiveToMonth', 'revision', 'updatedAt'];
var PHOTO_HEADERS_ = ['employeeId', 'jpegBase64', 'width', 'height', 'revision', 'updatedAt'];
var CALENDAR_HEADERS_ = ['dateKey', 'kind', 'note', 'revision', 'updatedAt'];

var ATTENDANCE_HEADERS_ = ['key', 'dateKey', 'employeeId', 'status', 'revision', 'updatedAt', 'updatedBy', 'requestId'];
var MONTHLY_EXTRA_HEADERS_ = ['extraId', 'employeeId', 'label', 'amountSatang', 'sourceTemplateId', 'sourceTemplateVersion', 'revision', 'updatedAt'];
var EXTRA_REVIEW_HEADERS_ = ['employeeId', 'extrasRevision', 'confirmedAt', 'confirmedBy'];
var MONTH_HEADERS_ = ['monthKey', 'state', 'revision', 'snapshotVersion', 'closedAt', 'closedBy'];
var SNAPSHOT_HEADERS_ = ['snapshotVersion', 'employeeId', 'jsonPartIndex', 'jsonPartCount', 'snapshotJsonPart'];
var AUDIT_HEADERS_ = ['auditId', 'entityKey', 'beforeJson', 'afterJson', 'updatedAt', 'updatedBy', 'requestId'];
var REQUEST_HEADERS_ = ['requestId', 'fingerprint', 'responseJson', 'updatedAt'];

var DEFAULT_WORKWEEK_JSON_ = '[1,2,3,4,5,6]'; // จันทร์-เสาร์ (Sunday=0)
var DEFAULT_SHOP_NAME_ = 'DE TEAM';
var SCHEMA_VERSION_ = '1';
var MAX_SNAPSHOT_PART_LEN_ = 30000;
var MAX_PHOTO_BASE64_LEN_ = 32768;
var MAX_PHOTO_BYTES_ = 24576; // 24 KiB

var DEFAULT_SPREADSHEET_ID_ = '1D9gYkF6uw9UdDDd4mrTBd3W7JU_kRtIawlqwLfce8bM';
var DEFAULT_OWNER_EMAIL_ = 'tlextle23@gmail.com';

function getConfigProperty_(key) {
  var props = PropertiesService.getScriptProperties();
  var val = props ? props.getProperty(key) : null;
  if (!val) {
    if (key === 'SPREADSHEET_ID') return DEFAULT_SPREADSHEET_ID_;
    if (key === 'OWNER_EMAIL') return DEFAULT_OWNER_EMAIL_;
    if (key === 'ADMIN_KEY') return 'DE_TEAM_SECURE_ADMIN_KEY';
  }
  return val;
}

function setConfigProperty_(key, value) {
  var props = PropertiesService.getScriptProperties();
  if (props) {
    props.setProperty(key, value);
  }
}
