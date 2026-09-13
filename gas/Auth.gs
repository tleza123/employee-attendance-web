/**
 * DE TEAM — ระบบเช็คชื่อพนักงานและคำนวณยอดค่าจ้าง
 * Auth.gs: การตรวจสอบสิทธิ์การใช้งานสำหรับผู้ดูแลระบบคนเดียว (Owner Only)
 */

function owner_() {
  var configured = getConfigProperty_('OWNER_EMAIL');
  var activeUser = Session.getActiveUser();
  var email = activeUser ? activeUser.getEmail() : '';

  requireValue_(configured && typeof configured === 'string' && configured.trim().length > 0, 'CONFIG_REQUIRED');
  requireValue_(email && typeof email === 'string' && email.trim().length > 0, 'AUTH_REQUIRED');
  requireValue_(email.trim().toLowerCase() === configured.trim().toLowerCase(), 'AUTH_REQUIRED');

  return email.trim().toLowerCase();
}
