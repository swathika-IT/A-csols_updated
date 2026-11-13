const assert = require('assert');
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');
const { buildWorkbook, buildStyledWorkbook } = require('../src/utils/xlsxExport');
const { parseXlsx } = require('../src/utils/xlsxLite');

function tmp(name) {
  const dir = path.join(__dirname, '..', 'private_uploads', 'imports');
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  return path.join(dir, name);
}

function readZipEntries(buffer) {
  const entries = new Map();
  let eocd = -1;
  for (let i = buffer.length - 22; i >= Math.max(0, buffer.length - 65557); i--) {
    if (buffer.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  assert.ok(eocd >= 0, 'ZIP end-of-central-directory record is missing');
  const cdSize = buffer.readUInt32LE(eocd + 12);
  const cdOffset = buffer.readUInt32LE(eocd + 16);
  let p = cdOffset;
  const end = cdOffset + cdSize;
  while (p < end) {
    assert.strictEqual(buffer.readUInt32LE(p), 0x02014b50, 'Invalid ZIP central directory entry');
    const method = buffer.readUInt16LE(p + 10);
    const compressedSize = buffer.readUInt32LE(p + 20);
    const nameLen = buffer.readUInt16LE(p + 28);
    const extraLen = buffer.readUInt16LE(p + 30);
    const commentLen = buffer.readUInt16LE(p + 32);
    const localOffset = buffer.readUInt32LE(p + 42);
    const name = buffer.slice(p + 46, p + 46 + nameLen).toString('utf8');
    const localNameLen = buffer.readUInt16LE(localOffset + 26);
    const localExtraLen = buffer.readUInt16LE(localOffset + 28);
    const dataStart = localOffset + 30 + localNameLen + localExtraLen;
    const raw = buffer.slice(dataStart, dataStart + compressedSize);
    entries.set(name, method === 0 ? raw : zlib.inflateRawSync(raw));
    p += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

function sheetNamesFromWorkbook(entries) {
  const workbook = entries.get('xl/workbook.xml').toString('utf8');
  return [...workbook.matchAll(/<sheet\b[^>]*name="([^"]*)"/g)].map(m => m[1]);
}

function xmlOrder(xml, tokens) {
  let last = -1;
  for (const token of tokens) {
    const at = xml.indexOf(token);
    assert.ok(at > last, `Expected ${token} after the previous element in worksheet XML`);
    last = at;
  }
}

function checkWorkbook(buffer) {
  assert.ok(Buffer.isBuffer(buffer) && buffer.length > 200, 'Workbook buffer is empty');
  assert.strictEqual(buffer.readUInt32LE(0), 0x04034b50, 'Workbook is not a ZIP/XLSX archive');
  const entries = readZipEntries(buffer);
  for (const part of ['[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/styles.xml']) {
    assert.ok(entries.has(part), `Missing required workbook part: ${part}`);
  }
  return entries;
}

function checkStylesAndSheetRelationships(entries, expectedSheets) {
  const names = sheetNamesFromWorkbook(entries);
  assert.strictEqual(names.length, expectedSheets, 'Unexpected number of worksheets');
  for (const name of names) {
    assert.ok(name.length >= 1 && name.length <= 31, `Worksheet name length is invalid: ${name}`);
    assert.ok(!/[\[\]:*?\/\\]/.test(name), `Worksheet name contains invalid characters: ${name}`);
  }
  assert.strictEqual(new Set(names.map(n => n.toLowerCase())).size, names.length, 'Worksheet names are not unique');

  const rels = entries.get('xl/_rels/workbook.xml.rels').toString('utf8');
  assert.ok(/relationships\/styles"\s+Target="styles\.xml"/.test(rels), 'styles.xml relationship is missing');
  const styles = entries.get('xl/styles.xml').toString('utf8');
  assert.ok(styles.includes('<cellStyles'), 'styles.xml has no default cell style');
  for (let i = 1; i <= expectedSheets; i++) {
    assert.ok(entries.has(`xl/worksheets/sheet${i}.xml`), `Missing worksheet part ${i}`);
  }
  return names;
}

function run() {
  console.log('XLSX export regression checks');
  const files = [];

  // 1. Plain people/interns workbook (the interns-data.xlsx shape).
  const plain = buildWorkbook([
    {
      name: 'Employees',
      rows: [
        ['Name', 'Employee ID', 'Category'],
        ['A & B <Test> Intern', 'I1', 'Intern'],
        ['Control \u0007 Char', 'I2', 'Intern'],
        ['NaN and Infinity', NaN, Infinity],
        ['Undefined', undefined, new Date('2026-09-01T00:00:00Z')]
      ]
    },
    { name: 'Employees', rows: [['Duplicate sheet name']] },
    { name: 'A/B:C*D?E[F]G', rows: [['Sanitized sheet name']] },
    { name: 'This worksheet name is deliberately longer than thirty one chars', rows: [['Long name']] },
    { name: 'Attendance', rows: [['Employee', 'Date']] },
    { name: 'Performance', rows: [['Employee', 'Score']] },
    { name: 'Leaves', rows: [['Employee', 'From']] }
  ]);
  const plainEntries = checkWorkbook(plain);
  const plainNames = checkStylesAndSheetRelationships(plainEntries, 7);
  assert.strictEqual(new Set(plainNames.map(n => n.toLowerCase())).size, plainNames.length);
  const plainSheet1 = plainEntries.get('xl/worksheets/sheet1.xml').toString('utf8');
  xmlOrder(plainSheet1, ['<dimension', '<sheetViews', '<cols', '<sheetData']);
  assert.ok(!plainSheet1.includes('\u0007'), 'Invalid XML control character was not stripped');
  assert.ok(!/<v>[^<]*(NaN|Infinity)[^<]*<\/v>/.test(plainSheet1), 'Invalid numeric values leaked into the workbook');
  const plainFile = tmp(`xlsx-regression-plain-${Date.now()}.xlsx`);
  files.push(plainFile);
  fs.writeFileSync(plainFile, plain);
  const plainRows = parseXlsx(plainFile);
  assert.strictEqual(plainRows[0][0], 'Name');
  assert.deepStrictEqual(plainRows[1].slice(0, 3), ['A & B <Test> Intern', 'I1', 'Intern']);

  // 2. Styled monthly report with all three preserved sheets.
  const title = { v: 'A&C SOLUTIONS — Monthly Attendance Matrix — September 2026', s: 'title' };
  const subtitle = { v: 'Report Category: All', s: 'subtitle' };
  const header = ['Employee ID', 'Employee Name', 'Type', 'Department', '1'].map(v => ({ v, s: 'header' }));
  const monthly = buildStyledWorkbook([
    {
      name: 'Monthly Attendance Matrix',
      rows: [[title], [subtitle], header, [{ v: 'I1', s: 'left' }, { v: 'Intern', s: 'left' }, { v: 'Intern', s: 'center' }, { v: 'QA', s: 'left' }, { v: 'P', s: 'P' }]],
      colWidths: [14, 26, 12, 18, 5],
      freezeCols: 4,
      freezeRows: 3,
      filterRef: 'A3:E3',
      merges: ['A1:E1', 'A2:E2']
    },
    {
      name: 'Monthly Summary',
      rows: [[title], [subtitle], header, [{ v: 'Total', s: 'totalLeft' }]],
      colWidths: [7, 14, 26, 13, 16],
      freezeCols: 3,
      freezeRows: 3,
      filterRef: 'A3:E3',
      merges: ['A1:E1', 'A2:E2']
    },
    {
      name: 'Daily Check-In/Check-Out/Hours',
      rows: [[title], [subtitle], header, [{ v: '09:00', s: 'center' }]],
      colWidths: [14, 26, 10, 16, 10],
      freezeCols: 4,
      freezeRows: 3,
      filterRef: null,
      merges: ['A1:E1', 'A2:E2']
    }
  ]);
  const monthlyEntries = checkWorkbook(monthly);
  const monthlyNames = checkStylesAndSheetRelationships(monthlyEntries, 3);
  assert.deepStrictEqual(monthlyNames, ['Monthly Attendance Matrix', 'Monthly Summary', 'Daily Check-In-Check-Out-Hours']);
  const matrixXml = monthlyEntries.get('xl/worksheets/sheet1.xml').toString('utf8');
  xmlOrder(matrixXml, ['<sheetPr', '<dimension', '<sheetViews', '<cols', '<sheetData', '<autoFilter', '<mergeCells', '<pageMargins', '<pageSetup']);
  const monthlyFile = tmp(`xlsx-regression-monthly-${Date.now()}.xlsx`);
  files.push(monthlyFile);
  fs.writeFileSync(monthlyFile, monthly);
  const monthlyRows = parseXlsx(monthlyFile);
  assert.strictEqual(monthlyRows[0][0], title.v);
  assert.deepStrictEqual(monthlyRows[2].slice(0, 4), ['Employee ID', 'Employee Name', 'Type', 'Department']);

  // 3. Daily workbook.
  const daily = buildStyledWorkbook([{
    name: 'Daily Attendance',
    rows: [[title], [subtitle], header, [{ v: '09:00', s: 'center' }]],
    colWidths: [14, 26, 13, 18, 12],
    freezeCols: 2,
    freezeRows: 3,
    filterRef: 'A3:E3',
    merges: ['A1:E1', 'A2:E2']
  }]);
  const dailyEntries = checkWorkbook(daily);
  checkStylesAndSheetRelationships(dailyEntries, 1);
  xmlOrder(dailyEntries.get('xl/worksheets/sheet1.xml').toString('utf8'), ['<sheetData', '<autoFilter', '<mergeCells']);
  const dailyFile = tmp(`xlsx-regression-daily-${Date.now()}.xlsx`);
  files.push(dailyFile);
  fs.writeFileSync(dailyFile, daily);
  assert.strictEqual(parseXlsx(dailyFile)[0][0], title.v);

  for (const file of files) fs.unlinkSync(file);
  console.log('XLSX export regression checks passed');
}

run();
