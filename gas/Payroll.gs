/**
 * DE TEAM — ระบบเช็คชื่อพนักงานและคำนวณยอดค่าจ้าง
 * Payroll.gs: Pure Calculation Engine คำนวณค่าจ้าง, ปัดเศษ, วันทำงาน, ตรวจสอบความถูกต้องของข้อมูล
 */

function requireValue_(condition, code) {
  if (!condition) throw new Error(code || 'INVALID_INPUT');
}

function dateKey_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    value = Utilities.formatDate(value, 'Asia/Bangkok', 'yyyy-MM-dd');
  } else if (value && typeof value !== 'string') {
    value = String(value);
  }
  if (typeof value === 'string') {
    value = value.trim();
  }
  requireValue_(typeof value === 'string' && /^20\d{2}-\d{2}-\d{2}$/.test(value), 'INVALID_DATE');
  var parsed = new Date(value + 'T00:00:00Z');
  requireValue_(Number.isFinite(parsed.getTime()) && parsed.toISOString().slice(0, 10) === value, 'INVALID_DATE');
  return value;
}

function monthKey_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    value = Utilities.formatDate(value, 'Asia/Bangkok', 'yyyy-MM');
  } else if (value && typeof value !== 'string') {
    value = String(value);
  }
  if (typeof value === 'string') {
    value = value.trim();
  }
  requireValue_(typeof value === 'string' && /^20\d{2}-(0[1-9]|1[0-2])$/.test(value), 'INVALID_MONTH');
  return value;
}

function moneySatang_(text) {
  requireValue_(typeof text === 'string' && /^(0|[1-9]\d{0,6})(\.\d{1,2})?$/.test(text), 'INVALID_MONEY');
  var parts = text.split('.');
  var amount = Number(parts[0]) * 100 + Number(((parts[1] || '') + '00').slice(0, 2));
  requireValue_(Number.isSafeInteger(amount) && amount <= 100000000, 'INVALID_MONEY');
  return amount;
}

function validateSatang_(amount) {
  requireValue_(Number.isSafeInteger(amount) && amount >= 0 && amount <= 100000000, 'INVALID_MONEY');
  return amount;
}

function monthDates_(month) {
  monthKey_(month);
  var cursor = new Date(month + '-01T00:00:00Z');
  var result = [];
  while (cursor.toISOString().slice(0, 7) === month) {
    result.push(cursor.toISOString().slice(0, 10));
    cursor.setUTCDate(cursor.getUTCDate() + 1);
  }
  return result;
}

function employedOn_(employee, date) {
  return date >= employee.startDate && (!employee.endDate || date <= employee.endDate);
}

function validateEmployeeDates_(employee) {
  requireValue_(employee && typeof employee.employeeId === 'string' && employee.employeeId.length > 0, 'INVALID_EMPLOYEE');
  dateKey_(employee.startDate);
  if (employee.endDate) {
    requireValue_(dateKey_(employee.endDate) >= employee.startDate, 'INVALID_EMPLOYMENT_DATES');
  }
}

function isWorkday_(date, weekdays, overrides) {
  dateKey_(date);
  requireValue_(
    Array.isArray(weekdays) &&
    weekdays.every(function (d) { return Number.isInteger(d) && d >= 0 && d <= 6; }) &&
    new Set(weekdays).size === weekdays.length,
    'INVALID_CALENDAR'
  );
  var kind = overrides && overrides[date];
  if (kind !== undefined) {
    requireValue_(kind === 'WORKDAY' || kind === 'HOLIDAY', 'INVALID_CALENDAR');
    return kind === 'WORKDAY';
  }
  return weekdays.indexOf(new Date(date + 'T00:00:00Z').getUTCDay()) >= 0;
}

function ratesFor_(rates, employeeId) {
  var seen = new Set();
  return rates.filter(function (r) { return r.employeeId === employeeId; }).map(function (r) {
    dateKey_(r.effectiveFrom);
    validateSatang_(r.dailySatang);
    requireValue_(!seen.has(r.effectiveFrom), 'DUPLICATE_RATE');
    seen.add(r.effectiveFrom);
    return r;
  }).sort(function (a, b) { return a.effectiveFrom.localeCompare(b.effectiveFrom); });
}

function rateOn_(sortedRates, date) {
  var selected = null;
  sortedRates.forEach(function (r) {
    if (r.effectiveFrom <= date) selected = r;
  });
  requireValue_(selected !== null, 'MISSING_RATE');
  return selected.dailySatang;
}

function dailyPay_(status, rate) {
  validateSatang_(rate);
  requireValue_(['FULL', 'HALF', 'ABSENT'].indexOf(status) >= 0, 'INVALID_STATUS');
  // กติกาอ้างอิง: ปัดครึ่งสตางค์ขึ้นต่อวัน (เช่น 500.01 บาท ครึ่งวัน = 250.01 บาท)
  return status === 'FULL' ? rate : status === 'HALF' ? Math.floor((rate + 1) / 2) : 0;
}

/**
 * คำนวณสรุปยอดรายเดือนของพนักงาน 1 คน
 * input: month, today, employee, rates[], attendance[], extras[], weekdays[], calendar{}, systemStartDate
 */
function calculateEmployeeMonth_(input) {
  monthKey_(input.month);
  dateKey_(input.today);
  dateKey_(input.systemStartDate);
  validateEmployeeDates_(input.employee);
  var employee = input.employee;
  var rates = ratesFor_(input.rates, employee.employeeId);
  var attendance = new Map();

  input.attendance.filter(function (r) {
    return r.employeeId === employee.employeeId && r.dateKey.slice(0, 7) === input.month;
  }).forEach(function (r) {
    dateKey_(r.dateKey);
    requireValue_(['FULL', 'HALF', 'ABSENT', 'UNMARKED'].indexOf(r.status) >= 0, 'INVALID_STATUS');
    requireValue_(!attendance.has(r.dateKey), 'DUPLICATE_ATTENDANCE');
    requireValue_(employedOn_(employee, r.dateKey), 'OUTSIDE_EMPLOYMENT');
    requireValue_(r.dateKey >= input.systemStartDate, 'BEFORE_SYSTEM_START');
    requireValue_(r.status === 'UNMARKED' || r.dateKey <= input.today, 'FUTURE_ATTENDANCE');
    requireValue_(r.status === 'UNMARKED' || isWorkday_(r.dateKey, input.weekdays, input.calendar), 'ATTENDANCE_ON_HOLIDAY');
    attendance.set(r.dateKey, r);
  });

  var result = {
    employeeId: employee.employeeId,
    month: input.month,
    full: 0,
    half: 0,
    absent: 0,
    pending: 0,
    workedDays: 0,
    paidDayUnits: 0,
    baseSatang: 0,
    extraSatang: 0,
    totalSatang: 0,
    days: [],
    extras: []
  };

  monthDates_(input.month).forEach(function (date) {
    if (!employedOn_(employee, date) || date < input.systemStartDate || date > input.today) return;
    var workday = isWorkday_(date, input.weekdays, input.calendar);
    var row = attendance.get(date);
    var status = row ? row.status : 'UNMARKED';

    if (!workday) {
      result.days.push({ dateKey: date, status: 'HOLIDAY', amountSatang: null });
      return;
    }

    if (status === 'UNMARKED') {
      result.pending += 1;
      result.days.push({ dateKey: date, status: 'UNMARKED', amountSatang: null });
      return;
    }

    var rate = rateOn_(rates, date);
    var amount = dailyPay_(status, rate);

    if (status === 'FULL') result.full += 1;
    if (status === 'HALF') result.half += 1;
    if (status === 'ABSENT') result.absent += 1;

    result.baseSatang += amount;
    result.days.push({ dateKey: date, status: status, dailySatang: rate, amountSatang: amount });
  });

  var extraIds = new Set();
  input.extras.filter(function (x) {
    return x.employeeId === employee.employeeId && x.monthKey === input.month;
  }).forEach(function (x) {
    requireValue_(typeof x.extraId === 'string' && x.extraId.length > 0 && !extraIds.has(x.extraId), 'DUPLICATE_EXTRA');
    requireValue_(typeof x.label === 'string' && x.label.trim().length > 0 && x.label.length <= 80, 'INVALID_EXTRA');
    extraIds.add(x.extraId);
    result.extraSatang += validateSatang_(x.amountSatang);
    result.extras.push({ extraId: x.extraId, label: x.label, amountSatang: x.amountSatang });
  });

  result.workedDays = result.full + result.half;
  result.paidDayUnits = result.full + result.half / 2;
  result.totalSatang = result.baseSatang + result.extraSatang;
  requireValue_(Number.isSafeInteger(result.totalSatang), 'MONEY_OVERFLOW');

  return result;
}
