const assert = require('assert');
const path = require('path');
const fs = require('fs');
const service = require('../src/services/attendanceReportService');
const { buildStyledWorkbook } = require('../src/utils/xlsxExport');
const { parseXlsx } = require('../src/utils/xlsxLite');

function tmp(name) {
  const dir = path.join(__dirname, '..', 'private_uploads', 'imports');
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  return path.join(dir, name);
}

async function run() {
  console.log('Attendance report unit checks');
  assert.strictEqual(service.parseMonthYear({ month: 2, year: 2024 }).days, 29);
  assert.strictEqual(service.parseMonthYear({ month: 1, year: 2026 }).days, 31);
  assert.strictEqual(service.formatHHMM(9 * 3600), '09:00');
  assert.strictEqual(service.formatHHMM(7 * 3600 + 45 * 60), '07:45');

  const users = [
    { _id: 'u1', employeeId: 'A1', name: 'Alice Employee', employeeType: 'employee', department: 'IT', checkInTime: '09:00', graceMinutes: 10, checkOutTime: '18:00', attendanceScheduleEnabled: true, requiredDailyHours: 9 },
    { _id: 'u2', employeeId: 'I1', name: 'Bob Intern', employeeType: 'intern', department: 'QA', checkInTime: '09:00', graceMinutes: 10, checkOutTime: '18:00', attendanceScheduleEnabled: true, requiredDailyHours: 9 },
    { _id: 'u3', employeeId: 'W1', name: 'Cara WFH', employeeType: 'wfh', department: 'Ops', checkInTime: '09:00', graceMinutes: 10, checkOutTime: '18:00', attendanceScheduleEnabled: false, requiredDailyHours: 8 }
  ];
  const my = service.parseMonthYear({ month: 9, year: 2026 });
  const mk = (date, hours) => {
    const day = date.slice(8);
    const start = new Date(`2026-09-${day}T09:00:00+05:30`);
    const end = new Date(start.getTime() + hours * 3600 * 1000);
    return { checkIn: start.toISOString(), checkOut: end.toISOString() };
  };
  const att = [
    { employee: 'u1', date: '2026-09-01', status: 'present', sessions: [{ ...mk('2026-09-01', 8), durationSeconds: 0 }], note: 'On site' },
    { employee: 'u1', date: '2026-09-02', status: 'present', sessions: [{ checkIn: '2026-09-02T09:15:00+05:30', checkOut: null, durationSeconds: 0 }] },
    { employee: 'u1', date: '2026-09-03', status: 'half-day', sessions: [{ ...mk('2026-09-03', 4), durationSeconds: 0 }], note: 'Left early' },
    { employee: 'u1', date: '2026-09-04', status: 'leave', sessions: [] },
    { employee: 'u1', date: '2026-09-07', status: 'absent', sessions: [] },
    { employee: 'u2', date: '2026-09-06', status: 'present', sessions: [{ ...mk('2026-09-06', 9), durationSeconds: 0 }] },
    { employee: 'u2', date: '2026-09-07', status: 'present', sessions: [{ ...mk('2026-09-07', 9), durationSeconds: 0 }] },
    { employee: 'u3', date: '2026-09-01', status: 'present', sessions: [{ ...mk('2026-09-01', 25), durationSeconds: 0 }] }
  ];
  const holidays = [{ date: '2026-09-05', name: 'Festival', appliesTo: 'all' }];
  const leaves = [{ employee: 'u1', status: 'approved', fromDate: '2026-09-04T00:00:00Z', toDate: '2026-09-04T00:00:00Z' }];
  const { perUser, totals } = service.computeMonthly(users, att, holidays, leaves, my);
  assert.strictEqual(perUser.length, 3);
  const alice = perUser.find(p => p.user.name === 'Alice Employee');
  const bob = perUser.find(p => p.user.name === 'Bob Intern');
  const cara = perUser.find(p => p.user.name === 'Cara WFH');
  assert.strictEqual(alice.summary.present, 1);
  assert.strictEqual(alice.summary.leave, 1);
  assert.strictEqual(alice.summary.half, 1);
  assert.strictEqual(alice.summary.absent, 25);
  assert.strictEqual(alice.summary.missing, 1);
  assert.strictEqual(alice.summary.holidays, 1);
  assert.strictEqual(alice.entries.find(e => e.date === '2026-09-02').code, 'M');
  assert.strictEqual(alice.entries.find(e => e.date === '2026-09-03').code, 'HD');
  assert.strictEqual(alice.entries.find(e => e.date === '2026-09-05').code, 'H');
  assert.strictEqual(bob.entries.find(e => e.date === '2026-09-06').code, 'P');
  assert.strictEqual(bob.entries.find(e => e.date === '2026-09-13').code, 'WO');
  assert.strictEqual(bob.summary.weeklyOffs > 0, true);
  assert.strictEqual(cara.entries.find(e => e.date === '2026-09-01').seconds > 0, true);
  console.log('monthly compute passed');

  const daily = service.computeDaily(users, att, holidays, leaves, '2026-09-01');
  assert.strictEqual(daily.length, 3);
  assert.strictEqual(daily.find(d => d.name === 'Alice Employee').checkIn, '09:00');
  assert.strictEqual(daily.find(d => d.name === 'Alice Employee').workingHours, '08:00');
  console.log('daily compute passed');

  const overnightUsers = [
    { _id: 'n1', employeeId: 'N1', name: 'Night Employee', employeeType: 'employee', department: 'Ops', checkInTime: '22:00', graceMinutes: 10, checkOutTime: '06:00', attendanceScheduleEnabled: true, requiredDailyHours: 8 }
  ];
  const overnightAtt = [
    { employee: 'n1', date: '2026-09-01', status: 'present', sessions: [{ checkIn: '2026-09-01T22:05:00+05:30', checkOut: '2026-09-02T06:30:00+05:30', durationSeconds: 0 }] },
    { employee: 'n1', date: '2026-09-02', status: 'present', sessions: [{ checkIn: '2026-09-02T22:30:00+05:30', checkOut: '2026-09-03T05:00:00+05:30', durationSeconds: 0 }] }
  ];
  const night = service.computeMonthly(overnightUsers, overnightAtt, [], [], my).perUser[0];
  assert.strictEqual(night.summary.expectedPerDay, '08:00');
  assert.strictEqual(night.summary.late, 1);
  assert.strictEqual(night.summary.early, 1);
  assert.strictEqual(night.entries.find(e => e.date === '2026-09-01').seconds, 8 * 3600 + 25 * 60);
  console.log('overnight shift compute passed');

  const futureMonth = service.parseMonthYear({ month: 12, year: 2099 });
  const future = service.computeMonthly(
    users,
    [],
    [{ date: '2099-12-25', name: 'Future Holiday', appliesTo: 'all' }],
    [{ employee: 'u1', status: 'approved', fromDate: '2099-12-10T00:00:00Z', toDate: '2099-12-10T00:00:00Z' }],
    futureMonth
  );
  assert.ok(future.perUser.every(item => item.entries.every(entry => entry.code === '')));
  assert.deepStrictEqual(future.totals, { P: 0, A: 0, L: 0, HD: 0, WO: 0, H: 0, M: 0, totalSeconds: 0, late: 0, early: 0, totalFormatted: '00:00', averageFormatted: '' });
  console.log('future dates blank passed');

  const buf = buildStyledWorkbook([
    { name: 'Monthly Attendance Matrix', rows: [[{ v: 'Title', s: 'title' }], [{ v: 'Sub', s: 'subtitle' }], [{ v: 'Employee', s: 'header' }, { v: 'P', s: 'P' }]] },
    { name: 'Monthly Summary', rows: [[{ v: 'Total: 08:00', s: 'total' }]] },
    { name: 'Daily Check-In/Check-Out/Hours', rows: [[{ v: '09:00', s: 'center' }]] }
  ]);
  assert.ok(buf.length > 500);
  const p = tmp(`attendance-report-unit-${Date.now()}.xlsx`);
  fs.writeFileSync(p, buf);
  const rows = parseXlsx(p);
  assert.ok(rows.length >= 3);
  assert.strictEqual(rows[0][0], 'Title');
  fs.unlinkSync(p);
  console.log('styled workbook passed');

  const html = fs.readFileSync(path.join(__dirname, '..', 'public', 'admin-workspace.html'), 'utf8');
  const js = fs.readFileSync(path.join(__dirname, '..', 'public', 'admin-workspace.js'), 'utf8');
  assert.ok(html.includes('id="reportMonth"'));
  assert.ok(html.includes('id="reportYear"'));
  assert.strictEqual((html.match(/id="reportMonthlyBtn"/g) || []).length, 1);
  for (const removed of ['reportCategory', 'reportPerson', 'reportIndividualBtn', 'dailyReportForm', 'dailyReportBtn', 'dailyReportDate', 'dailyReportCategory']) {
    assert.ok(!html.includes(removed), `Admin report UI still contains ${removed}`);
    assert.ok(!js.includes(removed), `Admin report script still references ${removed}`);
  }
  for (const removed of ['monthly-excel', 'downloadIndividualExcel', 'downloadDailyExcel']) {
    assert.ok(!js.includes(removed), `Admin report script still contains ${removed}`);
  }
  console.log('admin report UI cleanup passed');
  console.log('PASS');
}

run().catch(err => { console.error(err); process.exit(1); });
