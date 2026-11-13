process.env.NODE_ENV = 'test';
if (!process.env.TEST_DB_URI) {
  require('../src/config/environment').mongodbUri = '';
}
const assert = require('assert');
const http = require('http');
const fs = require('fs');
const path = require('path');
const app = require('../src/app');
const { connectDatabase, disconnectDatabase } = require('../src/config/database');
const User = require('../src/models/User');
const Attendance = require('../src/models/Attendance');
const LeaveRequest = require('../src/models/LeaveRequest');
const Holiday = require('../src/models/Holiday');
const { parseXlsx, parseXlsxSheets } = require('../src/utils/xlsxLite');

function request(method, route, options = {}) {
  return new Promise((resolve, reject) => {
    const url = new URL(route, baseUrl);
    const req = http.request({ method, hostname: url.hostname, port: url.port, path: url.pathname + url.search, headers: { 'Content-Type': 'application/json', ...(options.headers || {}) } }, (res) => {
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => {
        const buf = Buffer.concat(chunks);
        let json = null;
        try { json = JSON.parse(buf.toString('utf8')); } catch {}
        resolve({ status: res.statusCode, headers: res.headers, body: json || buf, raw: buf });
      });
    });
    req.on('error', reject);
    if (options.cookie) req.setHeader('Cookie', options.cookie);
    if (options.body) req.write(JSON.stringify(options.body));
    req.end();
  });
}
function cookie(headers) {
  const c = headers['set-cookie'];
  if (!c) return '';
  return (Array.isArray(c) ? c[0] : c).split(';')[0];
}
let server, baseUrl;

async function run() {
  await connectDatabase();
  server = app.listen(0);
  baseUrl = `http://127.0.0.1:${server.address().port}`;

  // seed admin
  await User.deleteMany({ role: 'admin' });
  const admin = await User.create({ name: 'Admin A', email: 'admin@a.local', username: 'admin', passwordHash: await User.hashPassword('AdminSecret2026!'), role: 'admin', isActive: true });
  const login = await request('POST', '/api/auth/login', { body: { username: 'admin', password: 'AdminSecret2026!' } });
  assert.strictEqual(login.status, 200);
  const adminCookie = cookie(login.headers);

  // seed people
  await User.deleteMany({ role: 'employee' });
  await Attendance.deleteMany({});
  await LeaveRequest.deleteMany({});
  await Holiday.deleteMany({});
  const emp = await User.create({ name: 'Alice Employee', email: 'alice@a.local', username: 'alice', passwordHash: await User.hashPassword('Password123!'), role: 'employee', employeeType: 'employee', employeeId: 'A1', department: 'IT', isActive: true, checkInTime: '09:00', checkOutTime: '18:00', attendanceScheduleEnabled: true });
  const intern = await User.create({ name: 'Bob Intern', email: 'bob@a.local', username: 'bob', passwordHash: await User.hashPassword('Password123!'), role: 'employee', employeeType: 'intern', employeeId: 'I1', department: 'QA', isActive: true, checkInTime: '09:00', checkOutTime: '18:00', attendanceScheduleEnabled: true });
  const wfh = await User.create({ name: 'Cara WFH', email: 'cara@a.local', username: 'cara', passwordHash: await User.hashPassword('Password123!'), role: 'employee', employeeType: 'wfh', employeeId: 'W1', department: 'Ops', isActive: true, checkInTime: '09:00', checkOutTime: '18:00', attendanceScheduleEnabled: true });
  const mkSession = (day, hours) => {
    const start = new Date(`2026-09-${day}T09:00:00+05:30`);
    return { checkIn: start, checkOut: new Date(start.getTime() + hours * 3600 * 1000), durationSeconds: hours * 3600 };
  };
  await Attendance.create({ employee: emp._id, date: '2026-09-01', status: 'present', sessions: [mkSession('01', 8)] });
  await Attendance.create({ employee: emp._id, date: '2026-09-02', status: 'present', sessions: [{ checkIn: new Date('2026-09-02T09:15:00+05:30'), checkOut: null, durationSeconds: 0 }] });
  await Attendance.create({ employee: intern._id, date: '2026-09-01', status: 'present', sessions: [mkSession('01', 9)] });
  await Attendance.create({ employee: wfh._id, date: '2026-09-01', status: 'present', sessions: [mkSession('01', 9)] });
  await LeaveRequest.create({ employee: emp._id, type: 'advance', fromDate: new Date('2026-09-03T00:00:00Z'), toDate: new Date('2026-09-03T00:00:00Z'), reason: 'Family', status: 'approved' });
  await LeaveRequest.create({ employee: emp._id, type: 'advance', fromDate: new Date('2026-09-04T00:00:00Z'), toDate: new Date('2026-09-04T00:00:00Z'), reason: 'Pending request', status: 'pending' });
  await Holiday.create({ date: '2026-09-05', name: 'Festival', appliesTo: 'all' });

  // unauthenticated
  const unauth = await request('GET', '/api/workspace-admin/attendance-reports/monthly.xlsx?month=9&year=2026&category=all');
  assert.strictEqual(unauth.status, 401);
  console.log('admin-only (unauth) passed');

  // monthly all: one file with exactly three sheets containing everyone
  const all = await request('GET', '/api/workspace-admin/attendance-reports/monthly.xlsx?month=9&year=2026&category=all', { cookie: adminCookie });
  assert.strictEqual(all.status, 200);
  assert.ok(String(all.headers['content-type']).includes('spreadsheetml'));
  assert.ok(String(all.headers['content-disposition']).includes('.xlsx'));
  const tmp1 = path.join(__dirname, '..', 'private_uploads', 'imports', `att-all-${Date.now()}.xlsx`);
  fs.writeFileSync(tmp1, all.raw);
  const rows1 = parseXlsx(tmp1);
  assert.ok(rows1.length >= 3);
  const workbook = fs.readFileSync(tmp1);
  assert.ok(workbook.length > 2000);
  console.log(`monthly all passed (${rows1.length} rows, ${workbook.length} bytes)`);

  const sheets = parseXlsxSheets(tmp1);
  assert.deepStrictEqual(sheets.map(s => s.name), ['Monthly Attendance Matrix', 'Monthly Summary', 'Daily Check-In-Check-Out-Hours']);
  assert.strictEqual(sheets.length, 3);
  for (const sheet of sheets) {
    const aliceRows = sheet.rows.filter(row => row.includes('Alice Employee'));
    const bobRows = sheet.rows.filter(row => row.includes('Bob Intern'));
    const caraRows = sheet.rows.filter(row => row.includes('Cara WFH'));
    assert.strictEqual(aliceRows.length, 1, `${sheet.name} must contain Alice exactly once`);
    assert.strictEqual(bobRows.length, 1, `${sheet.name} must contain Bob exactly once`);
    assert.strictEqual(caraRows.length, 1, `${sheet.name} must contain Cara exactly once`);
  }
  console.log('exactly three sheets with employees, WFH employees and interns passed');

  const [matrix, summary, dailyMatrix] = sheets;

  assert.deepStrictEqual(matrix.rows[2].slice(0, 4), ['Employee ID', 'Name', 'Type', 'Department']);
  assert.strictEqual(matrix.rows[2].length, 4 + 30);
  assert.strictEqual(matrix.rows[2][4], '1');
  assert.strictEqual(matrix.rows[2][matrix.rows[2].length - 1], '30');

  const aliceMatrix = matrix.rows.find(row => row[1] === 'Alice Employee');
  assert.strictEqual(aliceMatrix[0], 'A1');
  assert.strictEqual(aliceMatrix[2], 'Employee');
  assert.strictEqual(aliceMatrix[3], 'IT');
  assert.strictEqual(aliceMatrix[4], 'P');
  assert.strictEqual(aliceMatrix[5], 'M');
  assert.strictEqual(aliceMatrix[6], 'L');
  assert.strictEqual(aliceMatrix[7], 'A');
  assert.strictEqual(aliceMatrix[8], 'H');
  const bobMatrix = matrix.rows.find(row => row[1] === 'Bob Intern');
  assert.strictEqual(bobMatrix[2], 'Intern');
  assert.strictEqual(bobMatrix[4], 'P');
  assert.strictEqual(bobMatrix[16], 'WO');
  const caraMatrix = matrix.rows.find(row => row[1] === 'Cara WFH');
  assert.strictEqual(caraMatrix[2], 'WFH Employee');
  assert.strictEqual(caraMatrix[4], 'P');
  console.log('monthly matrix values passed');

  const expectedSummaryHeaders = ['Employee ID', 'Name', 'Type', 'Department', 'Shift', 'Expected Hours', 'Present Count', 'Absent Count', 'Approved Leave Count', 'Half-Day Count', 'Weekly-Off Count', 'Holiday Count', 'Total Working Hours', 'Average Working Hours', 'Late Arrivals', 'Early Checkouts', 'Remarks'];
  assert.deepStrictEqual(summary.rows[2], expectedSummaryHeaders);
  const aliceSummary = summary.rows.find(row => row[1] === 'Alice Employee');
  assert.strictEqual(aliceSummary[0], 'A1');
  assert.strictEqual(aliceSummary[2], 'Employee');
  assert.strictEqual(aliceSummary[3], 'IT');
  assert.strictEqual(aliceSummary[6], '1');
  assert.strictEqual(aliceSummary[7], '26');
  assert.strictEqual(aliceSummary[8], '1');
  assert.strictEqual(aliceSummary[9], '0');
  assert.strictEqual(aliceSummary[11], '1');
  const bobSummary = summary.rows.find(row => row[1] === 'Bob Intern');
  assert.strictEqual(bobSummary[2], 'Intern');
  assert.strictEqual(bobSummary[6], '1');
  assert.strictEqual(bobSummary[10], '4');
  assert.strictEqual(bobSummary[11], '1');
  const caraSummary = summary.rows.find(row => row[1] === 'Cara WFH');
  assert.strictEqual(caraSummary[2], 'WFH Employee');
  assert.strictEqual(caraSummary[6], '1');
  assert.strictEqual(caraSummary[7], '28');
  assert.strictEqual(caraSummary[11], '1');
  assert.strictEqual(summary.rows.length, 6);
  assert.strictEqual(matrix.rows.length, 6);
  assert.strictEqual(dailyMatrix.rows.length, 6);
  console.log('monthly summary totals passed');

  assert.deepStrictEqual(dailyMatrix.rows[2].slice(0, 7), ['Employee ID', 'Name', 'Type', 'Department', '1 In', '1 Out', '1 Hours']);
  assert.strictEqual(dailyMatrix.rows[2].length, 4 + 30 * 3);
  const aliceDaily = dailyMatrix.rows.find(row => row[1] === 'Alice Employee');
  assert.deepStrictEqual(aliceDaily.slice(4, 10), ['09:00', '17:00', '08:00', '09:15', 'M', 'M']);
  const bobDaily = dailyMatrix.rows.find(row => row[1] === 'Bob Intern');
  assert.deepStrictEqual(bobDaily.slice(4, 7), ['09:00', '18:00', '09:00']);
  const caraDaily = dailyMatrix.rows.find(row => row[1] === 'Cara WFH');
  assert.deepStrictEqual(caraDaily.slice(0, 7), ['W1', 'Cara WFH', 'WFH Employee', 'Ops', '09:00', '18:00', '09:00']);
  console.log('daily matrix values passed');

  // category employees only
  const empOnly = await request('GET', '/api/workspace-admin/attendance-reports/monthly.xlsx?month=9&year=2026&category=employees', { cookie: adminCookie });
  assert.strictEqual(empOnly.status, 200);
  // interns only
  const intOnly = await request('GET', '/api/workspace-admin/attendance-reports/monthly.xlsx?month=9&year=2026&category=interns', { cookie: adminCookie });
  assert.strictEqual(intOnly.status, 200);
  console.log('category filters passed');

  // individual
  const single = await request('GET', `/api/workspace-admin/attendance-reports/monthly.xlsx?month=9&year=2026&employeeId=${emp._id}`, { cookie: adminCookie });
  assert.strictEqual(single.status, 200);
  assert.ok(single.raw.includes(Buffer.from('Alice Employee')));
  console.log('individual passed');

  // preview JSON
  const prev = await request('GET', '/api/workspace-admin/attendance-reports/preview?month=9&year=2026&category=all', { cookie: adminCookie });
  assert.strictEqual(prev.status, 200);
  assert.strictEqual(prev.body.success, true);
  assert.strictEqual(prev.body.count, 3);
  const alice = prev.body.summary.find(s => s.name === 'Alice Employee');
  assert.strictEqual(alice.present, 1);
  assert.strictEqual(alice.leave, 1);
  console.log('preview passed');

  // daily
  const daily = await request('GET', '/api/workspace-admin/attendance-reports/daily.xlsx?date=2026-09-01&category=all', { cookie: adminCookie });
  assert.strictEqual(daily.status, 200);
  assert.ok(daily.raw.includes(Buffer.from('Daily Attendance')));
  console.log('daily passed');

  fs.unlinkSync(tmp1);
  console.log('PASS');
}

run().then(async () => { if (server) server.close(); await disconnectDatabase(); process.exit(0); })
  .catch(async err => { console.error(err); if (server) server.close(); await disconnectDatabase(); process.exit(1); });
