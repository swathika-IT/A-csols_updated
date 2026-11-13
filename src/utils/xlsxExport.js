const zlib = require('zlib');

// XML 1.0 does not allow these characters, and Excel refuses to open a
// workbook whose sheet XML contains them. Lone UTF-16 surrogates are also
// invalid once the document is serialised as UTF-8.
const INVALID_XML_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g;

function cleanXmlText(value) {
  return String(value == null ? '' : value)
    .replace(INVALID_XML_CHARS, '')
    .replace(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])/g, '')
    .replace(/(^|[^\uD800-\uDBFF])[\uDC00-\uDFFF]/g, '$1');
}

function escXml(value) {
  return cleanXmlText(value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

function cellValue(value) {
  if (value == null) return '';
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? '' : value.toISOString();
  if (typeof value === 'number') return Number.isFinite(value) ? String(value) : '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'bigint') return value.toString();
  const text = String(value);
  // Excel's maximum cell text length is 32767 characters.
  return text.length > 32767 ? text.slice(0, 32767) : text;
}

function safeColumnWidth(value) {
  const width = Number(value);
  return Number.isFinite(width) && width > 0 ? Math.min(width, 255) : 14;
}

function truncateSheetName(value) {
  const chars = Array.from(String(value));
  return chars.slice(0, 31).join('');
}

function sanitizeSheetName(value, usedNames) {
  let name = String(value == null ? '' : value)
    .replace(/[\[\]:*?\/\\]/g, '-')
    .replace(/[\u0000-\u001F]/g, ' ')
    .trim()
    .replace(/^'+|'+$/g, '')
    .trim();
  if (!name) name = 'Sheet';
  name = truncateSheetName(name);
  let candidate = name;
  let index = 2;
  while ([...usedNames].some(existing => existing.toLowerCase() === candidate.toLowerCase())) {
    const suffix = ` (${index})`;
    candidate = truncateSheetName(name.slice(0, Math.max(0, 31 - suffix.length)) + suffix);
    index += 1;
  }
  usedNames.add(candidate);
  return candidate;
}

function normalizeSheets(sheets) {
  const list = Array.isArray(sheets) ? sheets : [];
  const usedNames = new Set();
  const normalized = list.map(sheet => ({
    ...sheet,
    name: sanitizeSheetName(sheet && sheet.name, usedNames),
    rows: Array.isArray(sheet && sheet.rows) ? sheet.rows : []
  }));
  if (!normalized.length) normalized.push({ name: 'Sheet1', rows: [] });
  return normalized;
}

function crc32(buf) {
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    crc ^= buf[i];
    for (let j = 0; j < 8; j++) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function u32(n) { const b = Buffer.alloc(4); b.writeUInt32LE(n >>> 0); return b; }
function u16(n) { const b = Buffer.alloc(2); b.writeUInt16LE(n & 0xffff); return b; }

function zipStore(files) {
  const locals = [], centrals = [];
  let offset = 0;
  for (const file of files) {
    const name = Buffer.from(file.name, 'utf8');
    const data = Buffer.isBuffer(file.data) ? file.data : Buffer.from(file.data, 'utf8');
    const crc = crc32(data);
    const local = Buffer.concat([
      Buffer.from([0x50,0x4b,0x03,0x04]), u16(20), u16(0), u16(0), u16(0), u16(0),
      u32(crc), u32(data.length), u32(data.length), u16(name.length), u16(0), name, data
    ]);
    locals.push(local);
    const central = Buffer.concat([
      Buffer.from([0x50,0x4b,0x01,0x02]), u16(20), u16(20), u16(0), u16(0), u16(0), u16(0),
      u32(crc), u32(data.length), u32(data.length), u16(name.length), u16(0), u16(0), u16(0), u16(0), u32(0), u32(offset), name
    ]);
    centrals.push(central);
    offset += local.length;
  }
  const centralData = Buffer.concat(centrals);
  return Buffer.concat([...locals, centralData, Buffer.from([0x50,0x4b,0x05,0x06]), Buffer.alloc(4), u16(files.length), u16(files.length), u32(centralData.length), u32(offset), u16(0)]);
}

function colName(n) {
  let s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function sheetXml(rows) {
  const safeRows = Array.isArray(rows) ? rows : [];
  const cells = [];
  let endCol = 1;
  for (let r = 0; r < safeRows.length; r++) {
    const row = Array.isArray(safeRows[r]) ? safeRows[r] : [];
    endCol = Math.max(endCol, row.length);
    let xml = `<row r="${r + 1}">`;
    for (let c = 0; c < row.length; c++) {
      const ref = `${colName(c + 1)}${r + 1}`;
      xml += `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${escXml(cellValue(row[c]))}</t></is></c>`;
    }
    xml += '</row>'; cells.push(xml);
  }
  const endRow = Math.max(safeRows.length, 1);
  // Element order matters: dimension, sheetViews, cols, sheetData.
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`+
    `<dimension ref="A1:${colName(endCol)}${endRow}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>`+
    `<cols><col min="1" max="${endCol}" width="30" customWidth="1"/></cols><sheetData>${cells.join('')}</sheetData></worksheet>`;
}

function buildWorkbook(sheets) {
  const safeSheets = normalizeSheets(sheets);
  const workbookSheets = safeSheets.map((s, i) => `<sheet name="${escXml(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('');
  const rels = safeSheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') +
    `<Relationship Id="rId${safeSheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`;
  const files = [
    {name:'[Content_Types].xml',data:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${safeSheets.map((_,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`},
    {name:'_rels/.rels',data:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`},
    {name:'xl/workbook.xml',data:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${workbookSheets}</sheets></workbook>`},
    {name:'xl/_rels/workbook.xml.rels',data:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`},
    {name:'xl/styles.xml',data:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`}
  ];
  safeSheets.forEach((s,i)=>files.push({name:`xl/worksheets/sheet${i+1}.xml`,data:sheetXml(s.rows)}));
  return zipStore(files);
}

const STYLES = {
  title: { font: 1, fill: 2, align: 'left', bold: true, size: 14 },
  subtitle: { font: 2, fill: 2, align: 'left' },
  header: { font: 1, fill: 3, align: 'center', bold: true },
  headerId: { font: 1, fill: 3, align: 'center', bold: true },
  P: { font: 0, fill: 4, align: 'center' },
  A: { font: 0, fill: 5, align: 'center' },
  L: { font: 0, fill: 6, align: 'center' },
  HD: { font: 0, fill: 7, align: 'center' },
  WO: { font: 0, fill: 8, align: 'center' },
  H: { font: 0, fill: 9, align: 'center' },
  M: { font: 0, fill: 10, align: 'center' },
  center: { font: 0, fill: 0, align: 'center' },
  left: { font: 0, fill: 0, align: 'left' },
  total: { font: 1, fill: 11, align: 'center', bold: true },
  totalLeft: { font: 1, fill: 11, align: 'left', bold: true }
};
const STYLE_XF_MAP = { title: 1, subtitle: 2, header: 3, headerId: 3, P: 4, A: 5, L: 6, HD: 7, WO: 8, H: 9, M: 10, center: 11, left: 12, total: 13, totalLeft: 14 };
function stylesXml() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<fonts count="3">' +
    '<font><sz val="11"/><name val="Calibri"/></font>' +
    '<font><b/><sz val="14"/><name val="Calibri"/><color rgb="FFFFFFFF"/></font>' +
    '<font><sz val="10"/><name val="Calibri"/><color rgb="FFD9E2EC"/></font>' +
    '</fonts>' +
    '<fills count="12">' +
    '<fill><patternFill patternType="none"/></fill>' +
    '<fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/><bgColor rgb="FF1F3864"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FF2E75B6"/><bgColor rgb="FF2E75B6"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/><bgColor rgb="FFC6EFCE"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor rgb="FFFFC7CE"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/><bgColor rgb="FFFFEB9C"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFF8CBAD"/><bgColor rgb="FFF8CBAD"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/><bgColor rgb="FFD9D9D9"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFB4C6E7"/><bgColor rgb="FFB4C6E7"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFE699"/><bgColor rgb="FFFFE699"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFD9E2EC"/><bgColor rgb="FFD9E2EC"/></patternFill></fill>' +
    '</fills>' +
    '<borders count="2">' +
    '<border><left/><right/><top/><bottom/><diagonal/></border>' +
    '<border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border>' +
    '</borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="15">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="6" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="7" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="8" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="9" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="10" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="1" fillId="11" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="1" fillId="11" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>' +
    '</cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    '</styleSheet>';
}
function styledSheetXml(sheet) {
  const rows = Array.isArray(sheet.rows) ? sheet.rows : [];
  let width = 1;
  for (const row of rows) width = Math.max(width, Array.isArray(row) ? row.length : 0);
  const colWidths = Array.isArray(sheet.colWidths) ? sheet.colWidths : [];
  const freezeCols = Math.max(0, parseInt(sheet.freezeCols, 10) || 0);
  const freezeRows = Math.max(0, parseInt(sheet.freezeRows, 10) || 0);
  const colsXml = colWidths.length
    ? '<cols>' + colWidths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${safeColumnWidth(w)}" customWidth="1"/>`).join('') + '</cols>'
    : `<cols><col min="1" max="${width}" width="14" customWidth="1"/></cols>`;
  let viewsXml = '<sheetViews><sheetView workbookViewId="0">';
  if (freezeCols > 0 || freezeRows > 0) {
    const xSplit = freezeCols;
    const ySplit = freezeRows;
    const topLeft = `${colName(xSplit + 1)}${ySplit + 1}`;
    const activePane = freezeCols > 0 && freezeRows > 0 ? 'bottomRight' : freezeCols > 0 ? 'topRight' : 'bottomLeft';
    viewsXml += `<pane xSplit="${xSplit}" ySplit="${ySplit}" topLeftCell="${topLeft}" activePane="${activePane}" state="frozen"/>`;
    if (freezeCols > 0 && freezeRows > 0) {
      viewsXml += `<selection pane="topRight" activeCell="${colName(xSplit + 1)}1" sqref="${colName(xSplit + 1)}1"/>`;
      viewsXml += `<selection pane="bottomLeft" activeCell="A${ySplit + 1}" sqref="A${ySplit + 1}"/>`;
      viewsXml += `<selection pane="bottomRight" activeCell="${topLeft}" sqref="${topLeft}"/>`;
    } else {
      viewsXml += `<selection pane="${activePane}" activeCell="${topLeft}" sqref="${topLeft}"/>`;
    }
  }
  viewsXml += '</sheetView></sheetViews>';
  const merges = (Array.isArray(sheet.merges) ? sheet.merges : [])
    .map(m => String(m).trim())
    .filter(m => /^[A-Z]{1,3}\d+:[A-Z]{1,3}\d+$/.test(m));
  const mergeXml = merges.length ? `<mergeCells count="${merges.length}">` + merges.map(m => `<mergeCell ref="${m}"/>`).join('') + '</mergeCells>' : '';
  const filterRef = typeof sheet.filterRef === 'string' && /^[A-Z]{1,3}\d+:[A-Z]{1,3}\d+$/.test(sheet.filterRef.trim()) ? sheet.filterRef.trim() : '';
  const filterXml = filterRef ? `<autoFilter ref="${filterRef}"/>` : '';
  const pageXml = '<pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/><pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>';
  const setupXml = '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>';
  const cellsXml = rows.map((row, r) => {
    let xml = `<row r="${r + 1}">`;
    (Array.isArray(row) ? row : []).forEach((cell, c) => {
      let value = '';
      let styleKey = 'left';
      if (cell !== null && typeof cell === 'object' && !Array.isArray(cell) && ('v' in cell || 's' in cell)) {
        value = cell.v;
        styleKey = cell.s || 'left';
      } else {
        value = cell;
      }
      const xf = STYLE_XF_MAP[styleKey] != null ? STYLE_XF_MAP[styleKey] : STYLE_XF_MAP.left;
      xml += `<c r="${colName(c + 1)}${r + 1}" s="${xf}" t="inlineStr"><is><t xml:space="preserve">${escXml(cellValue(value))}</t></is></c>`;
    });
    xml += '</row>';
    return xml;
  }).join('');
  const endRow = Math.max(rows.length, 1);
  // Element order matters: sheetPr, dimension, sheetViews, cols, sheetData,
  // autoFilter, mergeCells, pageMargins, pageSetup.
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    setupXml + `<dimension ref="A1:${colName(width)}${endRow}"/>` + viewsXml + colsXml + '<sheetData>' + cellsXml + '</sheetData>' + filterXml + mergeXml + pageXml + '</worksheet>';
}
function buildStyledWorkbook(sheets) {
  const safeSheets = normalizeSheets(sheets);
  const workbookSheets = safeSheets.map((s, i) => `<sheet name="${escXml(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('');
  const rels = safeSheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') +
    `<Relationship Id="rId${safeSheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`;
  const files = [
    { name: '[Content_Types].xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${safeSheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>` },
    { name: '_rels/.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' },
    { name: 'xl/workbook.xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${workbookSheets}</sheets></workbook>` },
    { name: 'xl/_rels/workbook.xml.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>` },
    { name: 'xl/styles.xml', data: stylesXml() }
  ];
  safeSheets.forEach((s, i) => files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: styledSheetXml(s) }));
  return zipStore(files);
}
module.exports = { buildWorkbook, buildStyledWorkbook };
