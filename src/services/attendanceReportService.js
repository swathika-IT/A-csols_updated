const User = require('../models/User');
const Attendance = require('../models/Attendance');
const Holiday = require('../models/Holiday');
const LeaveRequest = require('../models/LeaveRequest');
const { buildStyledWorkbook } = require('../utils/xlsxExport');

function pad2(n) {
  return String(Number(n)).padStart(2, '0');
}

function daysInMonth(year, month) {
  return new Date(year, month, 0).getDate();
}

function monthNameOf(month) {
  return ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'][month - 1];
}

function dateString(year, month, day) {
  return `${year}-${pad2(month)}-${pad2(day)}`;
}

function todayIST() {
  return new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Kolkata' }).format(new Date());
}

function generatedIST() {
  return new Intl.DateTimeFormat('en-GB', { timeZone: 'Asia/Kolkata', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false }).format(new Date());
}

function istHHMM(dateInput) {
  if (!dateInput) return '';
  const d = new Date(dateInput);
  if (Number.isNaN(d.getTime())) return '';
  return new Intl.DateTimeFormat('en-GB', { timeZone: 'Asia/Kolkata', hour: '2-digit', minute: '2-digit', hour12: false }).format(d);
}

function minutesOfHHMM(value) {
  const [h, m] = String(value || '00:00').split(':').map(Number);
  return (h * 60 + m) || 0;
}

function scheduledDurationSeconds(checkInTime, checkOutTime) {
  const start = minutesOfHHMM(checkInTime);
  const end = minutesOfHHMM(checkOutTime);
  if (end === start) return 24 * 60 * 60;
  return (((end - start) % 1440 + 1440) % 1440) * 60;
}

function formatHHMM(totalSeconds) {
  const n = Math.max(0, Math.floor(Number(totalSeconds) || 0));
  const h = Math.floor(n / 3600);
  const m = Math.floor((n % 3600) / 60);
  return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
}

function shiftLabel(user) {
  if (!user || user.attendanceScheduleEnabled === false) return 'Flexible / No fixed time';
  return `${user.checkInTime || '09:00'} - ${user.checkOutTime || '18:00'}`;
}

function expectedHoursPerDay(user) {
  if (user && user.attendanceScheduleEnabled !== false) return formatHHMM(scheduledDurationSeconds(user.checkInTime || '09:00', user.checkOutTime || '18:00'));
  const req = Number(user && user.requiredDailyHours);
  return formatHHMM((Number.isFinite(req) && req > 0 ? req : 9) * 3600);
}

function typeLabel(employeeType) {
  const t = String(employeeType || 'employee').toLowerCase();
  if (t === 'intern') return 'Intern';
  if (t === 'wfh') return 'WFH Employee';
  return 'Employee';
}

function categoryLabel(category) {
  if (category === 'employees') return 'Employees Only';
  if (category === 'interns') return 'Interns Only';
  return 'All Employees & Interns';
}

function normalizeCategory(value) {
  const v = String(value || 'all').toLowerCase().trim();
  if (v === 'employees' || v === 'employee') return 'employees';
  if (v === 'interns' || v === 'intern') return 'interns';
  return 'all';
}

function parseMonthYear(input) {
  const obj = input || {};
  let month = parseInt(obj.month, 10);
  let year = parseInt(obj.year, 10);
  if (typeof obj.month === 'string' && /^\d{4}-\d{2}$/.test(obj.month.trim())) {
    const [y, m] = obj.month.trim().split('-').map(Number);
    year = y;
    month = m;
  }
  if (obj.monthStr && /^\d{4}-\d{2}$/.test(String(obj.monthStr).trim())) {
    const [y, m] = String(obj.monthStr).trim().split('-').map(Number);
    year = y;
    month = m;
  }
  if (!Number.isInteger(month)) month = Number(new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Kolkata', month: 'numeric' }).format(new Date()));
  if (!Number.isInteger(year)) year = Number(new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Kolkata', year: 'numeric' }).format(new Date()));
  if (month < 1 || month > 12) throw Object.assign(new Error('Month must be between 1 and 12.'), { status: 400 });
  if (year < 2000 || year > 2100) throw Object.assign(new Error('Year must be between 2000 and 2100.'), { status: 400 });
  const days = daysInMonth(year, month);
  return {
    month,
    year,
    days,
    monthName: monthNameOf(month),
    prefix: `${year}-${pad2(month)}`,
    start: dateString(year, month, 1),
    end: dateString(year, month, days)
  };
}

function toLeaveDateString(value) {
  if (!value) return '';
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return String(value).slice(0, 10);
  return d.toISOString().slice(0, 10);
}

function holidayApplies(holiday, employeeType) {
  if (!holiday) return false;
  if (holiday.appliesTo === 'all') return true;
  return holiday.appliesTo === employeeType;
}

function isInternSunday(dateStr, employeeType) {
  if (employeeType !== 'intern') return false;
  return new Date(`${dateStr}T12:00:00+05:30`).getDay() === 0;
}

function normalizeSessions(att) {
  if (!att) return [];
  if (Array.isArray(att.sessions) && att.sessions.length) return att.sessions;
  if (att.checkIn) return [{ checkIn: att.checkIn, checkOut: att.checkOut || null, durationSeconds: 0 }];
  return [];
}

function completedSessionSeconds(session) {
  if (!session || !session.checkIn || !session.checkOut) return null;
  const a = new Date(session.checkIn).getTime();
  const b = new Date(session.checkOut).getTime();
  if (Number.isNaN(a) || Number.isNaN(b)) return null;
  const computed = b > a ? Math.floor((b - a) / 1000) : 0;
  const stored = Number(session.durationSeconds || 0);
  return computed > 0 ? computed : (stored > 0 ? stored : computed);
}

function creditedSecondsFor(att, actualSeconds) {
  const credited = Number(att && att.creditedWorkingSeconds || 0);
  if (credited > 0) return credited;
  if (att && att.status === 'half-day') return Math.floor(actualSeconds / 2);
  return actualSeconds;
}

function dailyEntry(att, user, dateStr, context) {
  const { holiday, approvedLeave, isFuture } = context || {};
  if (att) {
    const status = String(att.status || 'present').toLowerCase();
    const sessions = normalizeSessions(att);
    const sorted = [...sessions].sort((a, b) => new Date(a.checkIn).getTime() - new Date(b.checkIn).getTime());
    const hasOpen = sorted.some(s => !s.checkOut);
    const checkInTime = sorted.length ? sorted[0].checkIn : null;
    const closedSessions = sorted.filter(s => s.checkOut);
    const lastClosed = closedSessions.length ? closedSessions[closedSessions.length - 1] : null;
    const checkOutTime = lastClosed ? lastClosed.checkOut : null;
    let actualSeconds = 0;
    let allCompleted = sorted.length > 0 && !hasOpen;
    for (const s of sorted) {
      const secs = completedSessionSeconds(s);
      if (secs == null) {
        allCompleted = false;
      } else {
        actualSeconds += secs;
      }
    }
    if (status === 'leave') return { code: 'L', checkInTime: null, checkOutTime: null, seconds: 0, incomplete: false, attendance: att };
    if (status === 'holiday') return { code: 'H', checkInTime: null, checkOutTime: null, seconds: 0, incomplete: false, attendance: att };
    if (status === 'absent') return { code: 'A', checkInTime: null, checkOutTime: null, seconds: 0, incomplete: false, attendance: att };
    if (!sorted.length || !checkInTime) return { code: 'M', checkInTime: null, checkOutTime: null, seconds: null, incomplete: true, attendance: att };
    if (status === 'half-day') {
      if (!allCompleted) return { code: 'M', checkInTime, checkOutTime: null, seconds: null, incomplete: true, attendance: att };
      const credited = creditedSecondsFor(att, actualSeconds);
      return { code: 'HD', checkInTime, checkOutTime, seconds: credited, actualSeconds, incomplete: false, attendance: att };
    }
    if (!allCompleted) return { code: 'M', checkInTime, checkOutTime: null, seconds: null, incomplete: true, attendance: att };
    const credited = creditedSecondsFor(att, actualSeconds);
    return { code: 'P', checkInTime, checkOutTime, seconds: credited, actualSeconds, incomplete: false, attendance: att };
  }
  if (isFuture) return { code: '', checkInTime: null, checkOutTime: null, seconds: 0, incomplete: false };
  if (approvedLeave) return { code: 'L', checkInTime: null, checkOutTime: null, seconds: 0, incomplete: false };
  if (holiday && holidayApplies(holiday, user && user.employeeType)) return { code: 'H', checkInTime: null, checkOutTime: null, seconds: 0, incomplete: false };
  if (isInternSunday(dateStr, user && user.employeeType)) return { code: 'WO', checkInTime: null, checkOutTime: null, seconds: 0, incomplete: false };
  return { code: 'A', checkInTime: null, checkOutTime: null, seconds: 0, incomplete: false };
}

function scheduledRangeMs(user, dateStr) {
  const start = new Date(`${dateStr}T${user.checkInTime || '09:00'}:00+05:30`).getTime();
  let end = new Date(`${dateStr}T${user.checkOutTime || '18:00'}:00+05:30`).getTime();
  if (!Number.isFinite(start) || !Number.isFinite(end)) return null;
  if (end <= start) end += 24 * 60 * 60 * 1000;
  return { start, end };
}

function isLate(entry, user, dateStr) {
  if (!entry || entry.incomplete || (entry.code !== 'P' && entry.code !== 'HD') || !entry.checkInTime) return false;
  if (!user || user.attendanceScheduleEnabled === false || !dateStr) return false;
  const range = scheduledRangeMs(user, dateStr);
  const actual = new Date(entry.checkInTime).getTime();
  if (!range || !Number.isFinite(actual)) return false;
  const grace = Number(user.graceMinutes != null ? user.graceMinutes : 10);
  return actual > range.start + grace * 60 * 1000;
}

function isEarlyCheckout(entry, user, dateStr) {
  if (!entry || entry.incomplete || (entry.code !== 'P' && entry.code !== 'HD') || !entry.checkOutTime) return false;
  if (!user || user.attendanceScheduleEnabled === false || !dateStr) return false;
  const range = scheduledRangeMs(user, dateStr);
  const actual = new Date(entry.checkOutTime).getTime();
  if (!range || !Number.isFinite(actual)) return false;
  return actual < range.end;
}

function buildLeaveSets(users, leaves) {
  const map = new Map(users.map(u => [String(u._id), new Set()]));
  for (const leave of leaves || []) {
    const uid = String(leave.employee);
    const set = map.get(uid);
    if (!set) continue;
    const from = toLeaveDateString(leave.fromDate);
    const to = toLeaveDateString(leave.toDate);
    if (!from || !to) continue;
    const start = new Date(`${from}T00:00:00Z`).getTime();
    const end = new Date(`${to}T00:00:00Z`).getTime();
    if (Number.isNaN(start) || Number.isNaN(end)) continue;
    for (let t = start; t <= end; t += 24 * 60 * 60 * 1000) {
      set.add(new Date(t).toISOString().slice(0, 10));
    }
  }
  return map;
}

function userStatusMap(attendance) {
  const map = new Map();
  for (const row of attendance || []) {
    const uid = String(row.employee);
    if (!map.has(uid)) map.set(uid, new Map());
    map.get(uid).set(row.date, row);
  }
  return map;
}

function computeMonthly(users, attendance, holidays, leaves, monthYear) {
  const { month, year, days, start, end } = monthYear;
  const todayStr = todayIST();
  const attMap = userStatusMap(attendance);
  const leaveSets = buildLeaveSets(users, leaves);
  const holidayMap = new Map((holidays || []).map(h => [h.date, h]));
  const perUser = users.map(user => {
    const uid = String(user._id);
    const userAtt = attMap.get(uid) || new Map();
    const leaveSet = leaveSets.get(uid) || new Set();
    const entries = [];
    const counts = { P: 0, A: 0, L: 0, HD: 0, WO: 0, H: 0, M: 0 };
    let totalSecs = 0;
    let late = 0;
    let early = 0;
    const notes = [];
    for (let day = 1; day <= days; day++) {
      const dateStr = dateString(year, month, day);
      const att = userAtt.get(dateStr) || null;
      const entry = dailyEntry(att, user, dateStr, {
        holiday: holidayMap.get(dateStr) || null,
        approvedLeave: leaveSet.has(dateStr),
        isFuture: dateStr > todayStr
      });
      entries.push({ date: dateStr, day, ...entry });
      if (counts[entry.code] != null) counts[entry.code]++;
      if (entry.seconds != null) totalSecs += entry.seconds;
      if (isLate(entry, user, dateStr)) late++;
      if (isEarlyCheckout(entry, user, dateStr)) early++;
      if (att && att.note) notes.push(String(att.note).trim());
      if (att && att.autoCheckoutReason) notes.push(String(att.autoCheckoutReason).trim());
    }
    const presentPlusHalf = counts.P + counts.HD;
    const avgSecs = presentPlusHalf > 0 ? totalSecs / presentPlusHalf : null;
    const remarks = [...new Set(notes.filter(Boolean))].join(' | ').slice(0, 500);
    const summary = {
      id: uid,
      employeeId: user.employeeId || '',
      name: user.name || '',
      type: typeLabel(user.employeeType),
      employeeType: user.employeeType || 'employee',
      department: user.department || '',
      shift: shiftLabel(user),
      expectedPerDay: expectedHoursPerDay(user),
      present: counts.P,
      absent: counts.A,
      leave: counts.L,
      half: counts.HD,
      weeklyOffs: counts.WO,
      holidays: counts.H,
      missing: counts.M,
      totalSeconds: totalSecs,
      totalFormatted: formatHHMM(totalSecs),
      averageFormatted: avgSecs == null ? '' : formatHHMM(avgSecs),
      late,
      early,
      remarks
    };
    return { user, entries, summary };
  });
  const totals = perUser.reduce((acc, item) => {
    acc.P += item.summary.present;
    acc.A += item.summary.absent;
    acc.L += item.summary.leave;
    acc.HD += item.summary.half;
    acc.WO += item.summary.weeklyOffs;
    acc.H += item.summary.holidays;
    acc.M += item.summary.missing;
    acc.totalSeconds += item.summary.totalSeconds;
    acc.late += item.summary.late;
    acc.early += item.summary.early;
    return acc;
  }, { P: 0, A: 0, L: 0, HD: 0, WO: 0, H: 0, M: 0, totalSeconds: 0, late: 0, early: 0 });
  const divisor = totals.P + totals.HD;
  totals.totalFormatted = formatHHMM(totals.totalSeconds);
  totals.averageFormatted = divisor > 0 ? formatHHMM(totals.totalSeconds / divisor) : '';
  return { perUser, totals, start, end, days };
}

function computeDaily(users, attendance, holidays, leaves, dateStr) {
  const todayStr = todayIST();
  const attMap = userStatusMap(attendance);
  const leaveSets = buildLeaveSets(users, leaves);
  const holidayMap = new Map((holidays || []).map(h => [h.date, h]));
  return users.map(user => {
    const uid = String(user._id);
    const att = (attMap.get(uid) || new Map()).get(dateStr) || null;
    const entry = dailyEntry(att, user, dateStr, {
      holiday: holidayMap.get(dateStr) || null,
      approvedLeave: (leaveSets.get(uid) || new Set()).has(dateStr),
      isFuture: dateStr > todayStr
    });
    const statusLabel = { P: 'Present', A: 'Absent', L: 'Approved Leave', HD: 'Half Day', WO: 'Weekly Off', H: 'Holiday', M: 'Missing / Incomplete', '': 'No Record' }[entry.code] || entry.code;
    return {
      id: uid,
      employeeId: user.employeeId || '',
      name: user.name || '',
      type: typeLabel(user.employeeType),
      employeeType: user.employeeType || 'employee',
      department: user.department || '',
      shift: shiftLabel(user),
      checkIn: entry.incomplete ? (entry.checkInTime ? istHHMM(entry.checkInTime) : 'M') : (entry.checkInTime ? istHHMM(entry.checkInTime) : ''),
      checkOut: entry.code === 'P' || entry.code === 'HD' ? (entry.checkOutTime ? istHHMM(entry.checkOutTime) : 'M') : '',
      workingHours: (entry.code === 'P' || entry.code === 'HD') && entry.seconds != null ? formatHHMM(entry.seconds) : '',
      seconds: entry.seconds || 0,
      code: entry.code || 'A',
      status: entry.code === '' ? 'A' : entry.code,
      statusLabel,
      remarks: att && att.note ? String(att.note).trim() : ''
    };
  });
}

function cell(v, s) {
  return { v, s };
}

function sheet1(perUser, monthYear, category) {
  const { monthName, year, days } = monthYear;
  const totalCols = 4 + days;
  const lastCol = colNameFor(totalCols);
  const title = `A&C SOLUTIONS — Monthly Attendance Matrix — ${monthName} ${year}`;
  const subtitle = `Report Category: ${categoryLabel(category)} | Users: ${perUser.length} | Generated (IST): ${generatedIST()} | Codes: P=Present A=Absent L=Approved Leave HD=Half Day WO=Weekly Off H=Holiday M=Missing/Incomplete`;
  const rows = [];
  const titleRow = [cell(title, 'title')];
  rows.push(titleRow);
  rows.push([cell(subtitle, 'subtitle')]);
  const header = [cell('Employee ID', 'header'), cell('Name', 'header'), cell('Type', 'header'), cell('Department', 'header')];
  for (let d = 1; d <= days; d++) header.push(cell(String(d), 'header'));
  rows.push(header);
  for (const item of perUser) {
    const row = [cell(item.user.employeeId || '', 'left'), cell(item.user.name || '', 'left'), cell(typeLabel(item.user.employeeType), 'center'), cell(item.user.department || '', 'left')];
    for (const e of item.entries) row.push(cell(e.code || '', e.code && ['P', 'A', 'L', 'HD', 'WO', 'H', 'M'].includes(e.code) ? e.code : 'center'));
    rows.push(row);
  }
  const colWidths = [14, 26, 12, 18];
  for (let d = 0; d < days; d++) colWidths.push(5);
  return {
    name: 'Monthly Attendance Matrix',
    rows,
    colWidths,
    freezeCols: 4,
    freezeRows: 3,
    filterRef: `A3:${lastCol}3`,
    merges: [`A1:${lastCol}1`, `A2:${lastCol}2`]
  };
}

function colNameFor(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function sheet2(perUser, monthYear, category) {
  const { monthName, year } = monthYear;
  const title = `A&C SOLUTIONS — Monthly Summary — ${monthName} ${year}`;
  const subtitle = `Report Category: ${categoryLabel(category)} | Users: ${perUser.length} | Generated (IST): ${generatedIST()} | Working hours use actual check-in/check-out sessions (credited where the system applies it).`;
  const headers = ['Employee ID', 'Name', 'Type', 'Department', 'Shift', 'Expected Hours', 'Present Count', 'Absent Count', 'Approved Leave Count', 'Half-Day Count', 'Weekly-Off Count', 'Holiday Count', 'Total Working Hours', 'Average Working Hours', 'Late Arrivals', 'Early Checkouts', 'Remarks'];
  const rows = [
    [cell(title, 'title')],
    [cell(subtitle, 'subtitle')],
    headers.map(h => cell(h, 'header'))
  ];
  perUser.forEach(item => {
    const s = item.summary;
    rows.push([
      cell(s.employeeId, 'left'),
      cell(s.name, 'left'),
      cell(s.type, 'center'),
      cell(s.department, 'left'),
      cell(s.shift, 'center'),
      cell(s.expectedPerDay, 'center'),
      cell(String(s.present), 'center'),
      cell(String(s.absent), 'center'),
      cell(String(s.leave), 'center'),
      cell(String(s.half), 'center'),
      cell(String(s.weeklyOffs), 'center'),
      cell(String(s.holidays), 'center'),
      cell(s.totalFormatted, 'center'),
      cell(s.averageFormatted, 'center'),
      cell(String(s.late), 'center'),
      cell(String(s.early), 'center'),
      cell(s.remarks, 'left')
    ]);
  });
  return {
    name: 'Monthly Summary',
    rows,
    colWidths: [14, 26, 13, 16, 16, 14, 11, 11, 11, 10, 11, 11, 13, 14, 11, 12, 40],
    freezeCols: 3,
    freezeRows: 3,
    filterRef: 'A3:Q3',
    merges: ['A1:Q1', 'A2:Q2']
  };
}

function sheet3(perUser, monthYear, category) {
  const { monthName, year, days } = monthYear;
  const totalCols = 4 + days * 3;
  const lastCol = colNameFor(totalCols);
  const title = `A&C SOLUTIONS — Daily Check-In / Check-Out / Working Hours Matrix — ${monthName} ${year}`;
  const subtitle = `Report Category: ${categoryLabel(category)} | Users: ${perUser.length} | Generated (IST): ${generatedIST()} | Times in HH:MM (IST). M = Missing/Incomplete (no invented check-out).`;
  const rows = [
    [cell(title, 'title')],
    [cell(subtitle, 'subtitle')]
  ];
  const header = [cell('Employee ID', 'header'), cell('Name', 'header'), cell('Type', 'header'), cell('Department', 'header')];
  for (let d = 1; d <= days; d++) {
    header.push(cell(`${d} In`, 'header'));
    header.push(cell(`${d} Out`, 'header'));
    header.push(cell(`${d} Hours`, 'header'));
  }
  rows.push(header);
  for (const item of perUser) {
    const row = [cell(item.user.employeeId || '', 'left'), cell(item.user.name || '', 'left'), cell(typeLabel(item.user.employeeType), 'center'), cell(item.user.department || '', 'left')];
    for (const e of item.entries) {
      if (e.code === 'P' || e.code === 'HD') {
        row.push(cell(e.checkInTime ? istHHMM(e.checkInTime) : '', 'center'));
        row.push(cell(e.checkOutTime ? istHHMM(e.checkOutTime) : 'M', 'center'));
        row.push(cell(e.seconds != null ? formatHHMM(e.seconds) : 'M', 'center'));
      } else if (e.code === 'M') {
        row.push(cell(e.checkInTime ? istHHMM(e.checkInTime) : 'M', 'M'));
        row.push(cell('M', 'M'));
        row.push(cell('M', 'M'));
      } else if (e.code) {
        row.push(cell(e.code, e.code));
        row.push(cell(e.code, e.code));
        row.push(cell(e.code, e.code));
      } else {
        row.push(cell('', 'center'));
        row.push(cell('', 'center'));
        row.push(cell('', 'center'));
      }
    }
    rows.push(row);
  }
  const colWidths = [14, 26, 10, 16];
  for (let d = 0; d < days * 3; d++) colWidths.push(10);
  return {
    name: 'Daily Check-In-Check-Out-Hours',
    rows,
    colWidths,
    freezeCols: 4,
    freezeRows: 3,
    filterRef: `A3:${lastCol}3`,
    merges: [`A1:${lastCol}1`, `A2:${lastCol}2`]
  };
}

function dailySheet(dailyRows, dateStr, category) {
  const title = `A&C SOLUTIONS — Daily Attendance — ${dateStr}`;
  const subtitle = `Report Category: ${categoryLabel(category)} | Users: ${dailyRows.length} | Generated (IST): ${generatedIST()} | Times in HH:MM (IST).`;
  const headers = ['Employee ID', 'Employee Name', 'Employee Type', 'Department', 'Shift', 'Check-In', 'Check-Out', 'Total Working Hours', 'Attendance Status', 'Remarks'];
  const rows = [
    [cell(title, 'title')],
    [cell(subtitle, 'subtitle')],
    headers.map(h => cell(h, 'header'))
  ];
  for (const r of dailyRows) {
    rows.push([
      cell(r.employeeId, 'left'),
      cell(r.name, 'left'),
      cell(r.type, 'center'),
      cell(r.department, 'left'),
      cell(r.shift, 'center'),
      cell(r.checkIn, 'center'),
      cell(r.checkOut, 'center'),
      cell(r.workingHours, 'center'),
      cell(r.statusLabel, 'center'),
      cell(r.remarks, 'left')
    ]);
  }
  return {
    name: 'Daily Attendance',
    rows,
    colWidths: [14, 26, 13, 18, 16, 12, 12, 15, 16, 40],
    freezeCols: 2,
    freezeRows: 3,
    filterRef: 'A3:J3',
    merges: ['A1:J1', 'A2:J2']
  };
}

async function fetchUsers({ category, employeeId }) {
  const cat = normalizeCategory(category);
  let query = { role: 'employee' };
  if (employeeId) {
    const single = await User.findOne({ _id: employeeId, role: 'employee' }).lean();
    if (!single) return [];
    return [single];
  }
  if (cat === 'employees') query.employeeType = { $in: ['employee', 'wfh'] };
  else if (cat === 'interns') query.employeeType = 'intern';
  const users = await User.find(query).sort({ name: 1 }).lean();
  const seen = new Set();
  return users.filter(u => {
    const k = String(u._id);
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  });
}

function sanitizeFilenamePart(value) {
  return String(value || 'report').replace(/[\\/:*?"<>|]/g, '-').replace(/\s+/g, '_').slice(0, 60);
}

async function buildMonthlyReport(options) {
  const opts = options || {};
  const monthYear = parseMonthYear(opts);
  const category = normalizeCategory(opts.category);
  const users = await fetchUsers({ category, employeeId: opts.employeeId || opts.userId });
  if (!users.length) throw Object.assign(new Error('No employee records found for this selection.'), { status: 404 });
  const ids = users.map(u => u._id);
  const [attendance, holidays, leaves] = await Promise.all([
    Attendance.find({ employee: { $in: ids }, date: { $gte: monthYear.start, $lte: monthYear.end } }).sort({ date: 1 }).lean(),
    Holiday.find({ date: { $gte: monthYear.start, $lte: monthYear.end } }).lean(),
    LeaveRequest.find({ employee: { $in: ids }, status: 'approved', fromDate: { $lte: new Date(`${monthYear.end}T23:59:59.999Z`) }, toDate: { $gte: new Date(`${monthYear.start}T00:00:00Z`) } }).lean()
  ]);
  const { perUser, totals } = computeMonthly(users, attendance, holidays, leaves, monthYear);
  const buffer = buildStyledWorkbook([
    sheet1(perUser, monthYear, category),
    sheet2(perUser, monthYear, category),
    sheet3(perUser, monthYear, category)
  ]);
  const single = users.length === 1 && opts.employeeId;
  const filename = single
    ? `${sanitizeFilenamePart(users[0].employeeId || users[0].name)}_Attendance_${monthYear.prefix}.xlsx`
    : `Attendance_Monthly_${monthYear.prefix}_${category === 'all' ? 'All' : category === 'employees' ? 'Employees' : 'Interns'}.xlsx`;
  return {
    buffer,
    filename,
    meta: { monthYear, category, count: users.length, totals },
    summary: perUser.map(p => p.summary),
    perUser,
    totals
  };
}

async function buildDailyReport(options) {
  const opts = options || {};
  const dateStr = String(opts.date || todayIST()).trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) throw Object.assign(new Error('Date must be in YYYY-MM-DD format.'), { status: 400 });
  const category = normalizeCategory(opts.category);
  const users = await fetchUsers({ category, employeeId: opts.employeeId || opts.userId });
  if (!users.length) throw Object.assign(new Error('No employee records found for this selection.'), { status: 404 });
  const ids = users.map(u => u._id);
  const [attendance, holidays, leaves] = await Promise.all([
    Attendance.find({ employee: { $in: ids }, date: dateStr }).lean(),
    Holiday.find({ date: dateStr }).lean(),
    LeaveRequest.find({ employee: { $in: ids }, status: 'approved', fromDate: { $lte: new Date(`${dateStr}T23:59:59.999Z`) }, toDate: { $gte: new Date(`${dateStr}T00:00:00Z`) } }).lean()
  ]);
  const dailyRows = computeDaily(users, attendance, holidays, leaves, dateStr);
  const buffer = buildStyledWorkbook([dailySheet(dailyRows, dateStr, category)]);
  const filename = `Daily_Attendance_${dateStr}_${category === 'all' ? 'All' : category === 'employees' ? 'Employees' : 'Interns'}.xlsx`;
  return { buffer, filename, meta: { date: dateStr, category, count: users.length }, dailyRows };
}

module.exports = {
  buildMonthlyReport,
  buildDailyReport,
  fetchUsers,
  parseMonthYear,
  computeMonthly,
  computeDaily,
  dailyEntry,
  daysInMonth,
  formatHHMM,
  normalizeCategory,
  categoryLabel,
  typeLabel,
  shiftLabel,
  expectedHoursPerDay,
  istHHMM
};
