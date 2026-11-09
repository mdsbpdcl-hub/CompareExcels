// Opening files, reading rows with real coordinates, and staging them into DuckDB.
import * as XLSX from '@e965/xlsx';
import { unzipSync } from 'fflate';
import { tableToIPC, Table, vectorFromArray, Utf8, Int64 } from 'apache-arrow';
import { colLetter } from './coords.js';
import { parseCsv, detectEncoding, decodeText, detectDelimiter } from './csv.js';

export const EXCEL_EXTS = { xlsx: 'xlsx', xlsm: 'xlsm', xls: 'xls', xlsb: 'xlsb' };
export const CSV_EXTS = { csv: 'csv', txt: 'csv', tsv: 'csv' };
const BATCH = 50000;
const HEADER_SCAN = 30;

export function detectFormat(name) {
  const ext = (name.split('.').pop() || '').toLowerCase();
  if (EXCEL_EXTS[ext]) return EXCEL_EXTS[ext];
  if (CSV_EXTS[ext]) return 'csv';
  throw new Error(`'${name}' is not a supported file. Use .xlsx, .xlsm, .xls, .xlsb, .csv, .txt or .tsv.`);
}

export function csvSheetName(name) {
  const stem = name.replace(/\.[^.]+$/, '') || 'Data';
  return stem.replace(/[[\]*?/\\:]/g, '_').slice(0, 31);
}

export async function sha256(bytes) {
  const h = await globalThis.crypto.subtle.digest('SHA-256', bytes);
  return [...new Uint8Array(h)].map((b) => b.toString(16).padStart(2, '0')).join('');
}

export const kindLabel = (fmt) => ({ xlsx: 'Excel', xlsm: 'Excel (macros)', xls: 'Excel 97-2003',
  xlsb: 'Excel binary', csv: 'CSV' })[fmt];

/** Inspect a file. bytes: Uint8Array (kept untouched). */
export async function openSource(name, bytes, opts = {}) {
  const fmt = detectFormat(name);
  const sf = { name, fmt, bytes, size: bytes.length, sha256: await sha256(bytes), sheets: [], warnings: [],
    encoding: null, delimiter: null, wb: null, text: null, isExcel: fmt !== 'csv' };
  if (fmt === 'csv') {
    sf.encoding = opts.encoding || detectEncoding(bytes);
    sf.text = decodeText(bytes, sf.encoding);
    sf.delimiter = opts.delimiter || detectDelimiter(sf.text.slice(0, 256 * 1024), name);
    sf.sheets = [csvSheetName(name)];
    return sf;
  }
  try {
    sf.wb = XLSX.read(bytes, { type: 'array', cellFormula: true, cellNF: true, cellText: false, cellDates: false,
      dense: true, cellStyles: false });
  } catch (e) {
    throw new Error(`Could not open '${name}' as Excel: ${e.message || e}`);
  }
  sf.sheets = sf.wb.SheetNames.filter((n) => sf.wb.Sheets[n] && !sf.wb.Sheets[n]['!type']);
  if (fmt === 'xlsx' || fmt === 'xlsm') sf.warnings.push(...inspectXlsx(bytes));
  else sf.warnings.push(`${kindLabel(fmt)} files are read as values. Formulas, styles and column widths cannot be `
    + 'carried into the highlighted copy.');
  return sf;
}

function inspectXlsx(bytes) {
  let names;
  try { names = listZip(bytes); } catch { return ['The file does not look like a normal .xlsx package.']; }
  const out = [];
  if (names.some((n) => n.toLowerCase().endsWith('vbaproject.bin'))) out.push('This workbook contains macros. '
    + 'They are never run; the highlighted copy keeps them and is saved as .xlsm.');
  if (names.some((n) => n.startsWith('xl/externalLinks/'))) out.push('This workbook links to other files. Linked '
    + 'values are compared as last saved.');
  if (names.some((n) => n.startsWith('xl/pivotTables/'))) out.push('This workbook has pivot tables. They are kept '
    + 'but not compared.');
  return out;
}

export function listZip(bytes) {
  const names = [];
  unzipSync(bytes, { filter: (f) => { names.push(f.name); return false; } });
  return names;
}

// ---------------------------------------------------------------------------
function pad(n, w = 2) { return String(n).padStart(w, '0'); }

/** Turn a SheetJS cell into the text used for comparison (null = blank). */
export function cellToText(cell, date1904 = false) {
  if (!cell || cell.t === 'z') return null;
  const v = cell.v;
  if (v === undefined || v === null) return null;
  switch (cell.t) {
    case 's': return v === '' ? null : String(v);
    case 'b': return v ? 'TRUE' : 'FALSE';
    case 'e': return errorText(cell);
    case 'd': return v instanceof Date ? v.toISOString().slice(0, 10) : String(v);
    case 'n': {
      if (cell.z && typeof cell.z === 'string' && XLSX.SSF.is_date(cell.z)) {
        const p = XLSX.SSF.parse_date_code(v, { date1904 });
        if (p) {
          const hasTime = p.H || p.M || p.S;
          const f = cell.z.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '');
          if (v >= 0 && v < 1 && !/[yd]/i.test(f) && /[hs]/i.test(f)) return `${pad(p.H)}:${pad(p.M)}:${pad(p.S)}`;
          const d = `${pad(p.y, 4)}-${pad(p.m)}-${pad(p.d)}`;
          return hasTime ? `${d} ${pad(p.H)}:${pad(p.M)}:${pad(p.S)}` : d;
        }
      }
      return numberToText(v);
    }
    default: return String(v);
  }
}

const ERRORS = { 0: '#NULL!', 7: '#DIV/0!', 15: '#VALUE!', 23: '#REF!', 29: '#NAME?', 36: '#NUM!', 42: '#N/A',
  43: '#GETTING_DATA', 44: '#SPILL!', 45: '#CONNECT!', 46: '#BLOCKED!', 47: '#UNKNOWN!', 48: '#FIELD!', 49: '#CALC!' };
/** Excel error cells (#N/A, #DIV/0! …) as the text Excel shows. */
export function errorText(cell) {
  if (typeof cell.w === 'string' && cell.w.startsWith('#')) return cell.w;
  if (typeof cell.v === 'string' && cell.v.startsWith('#')) return cell.v;
  return ERRORS[cell.v] || '#N/A';
}

export function numberToText(v) {
  if (!Number.isFinite(v)) return null;
  if (Number.isInteger(v) && Math.abs(v) < 1e15) return String(v);
  // Excel keeps 15 significant digits: 0.1+0.2 is shown (and exported to CSV) as 0.3
  return String(Number(v.toPrecision(15)));
}

/** Rows of a sheet: { firstCol, lastCol, lastRow, rows: iterator of [rowNo, values[]] }. */
export function sheetReader(sf, sheet) {
  if (sf.fmt === 'csv') {
    return {
      firstCol: 1,
      *rows() { let i = 0; for (const rec of parseCsv(sf.text, sf.delimiter)) { i++; yield [i, rec.map((x) => (x === '' ? null : x))]; } },
    };
  }
  const ws = sf.wb.Sheets[sheet];
  if (!ws) throw new Error(`Sheet '${sheet}' was not found in ${sf.name}`);
  const date1904 = !!(sf.wb.Workbook && sf.wb.Workbook.WBProps && sf.wb.Workbook.WBProps.date1904);
  const ref = ws['!ref'];
  if (!ref) return { firstCol: 1, lastRow: 0, *rows() {} };
  const rg = XLSX.utils.decode_range(ref);
  const data = ws['!data'] || [];
  return {
    firstCol: rg.s.c + 1,
    lastCol: rg.e.c + 1,
    lastRow: rg.e.r + 1,
    *rows() {
      for (let r = 0; r <= rg.e.r; r++) {
        const row = data[r];
        const vals = new Array(rg.e.c - rg.s.c + 1);
        for (let c = rg.s.c; c <= rg.e.c; c++) vals[c - rg.s.c] = row ? cellToText(row[c], date1904) : null;
        yield [r + 1, vals];
      }
    },
  };
}

/** Raw SheetJS cell at (row, col), 1-based. */
export function rawCell(sf, sheet, row, col) {
  const ws = sf.wb.Sheets[sheet];
  const data = ws['!data'] || [];
  return data[row - 1] ? data[row - 1][col - 1] : undefined;
}

export function readTopRows(sf, sheet, n = 80) {
  const rd = sheetReader(sf, sheet);
  const out = [];
  for (const [rno, vals] of rd.rows()) {
    if (rno > n) break;
    out.push([rno, vals]);
  }
  return { rows: out, firstCol: rd.firstCol };
}

const looksNumeric = (s) => /^[-+]?[\d,]*\.?\d+(e[-+]?\d+)?$/i.test(s.trim());

export function detectHeaderRow(rows) {
  const scan = rows.slice(0, HEADER_SCAN);
  if (!scan.length) return 1;
  const widths = scan.map(([, r]) => r.filter((v) => v !== null && v !== '').length);
  const widest = Math.max(...widths);
  if (widest === 0) return scan[0][0];
  for (let i = 0; i < scan.length; i++) {
    const w = widths[i];
    if (w === 0 || (w < Math.max(2, 0.6 * widest) && widest >= 2)) continue;
    const vals = scan[i][1].filter((v) => v !== null && v !== '');
    const texty = vals.filter((v) => !looksNumeric(v)).length;
    if (texty >= 0.6 * vals.length && new Set(vals).size === vals.length) return scan[i][0];
  }
  for (let i = 0; i < scan.length; i++) if (widths[i] === widest) return scan[i][0];
  return scan[0][0];
}

export function buildColumns(headerVals, firstCol, ncols) {
  const cols = [];
  const seen = new Map();
  for (let j = 0; j < ncols; j++) {
    const raw = j < headerVals.length ? headerVals[j] : null;
    const header = raw == null ? '' : String(raw).trim();
    const colno = firstCol + j;
    const letter = colLetter(colno);
    const base = header || `(no heading) ${letter}`;
    const key = base.toLowerCase();
    let name = base;
    if (seen.has(key)) { seen.set(key, seen.get(key) + 1); name = `${base} (${seen.get(key)})`; } else seen.set(key, 1);
    cols.push({ internal: `c${j + 1}`, name, header: raw == null ? '' : String(raw), col: colno, letter });
  }
  return cols;
}

/** Read a whole sheet once: header, column list, data rows (non-empty) and counts. */
export function scanSheet(sf, sheet, headerRow) {
  const rd = sheetReader(sf, sheet);
  let headerVals = [];
  let ncols = 0;
  let lastRow = 0;
  let empty = 0;
  const rowNos = [];
  const data = [];
  for (const [rno, vals] of rd.rows()) {
    lastRow = rno;
    if (rno < headerRow) continue;
    if (rno === headerRow) { headerVals = vals; ncols = Math.max(ncols, vals.length); continue; }
    let n = vals.length;
    while (n && vals[n - 1] === null) n--;
    if (n === 0) { empty++; continue; }
    if (n > ncols) ncols = n;
    rowNos.push(rno);
    data.push(vals);
  }
  if (rd.lastRow) lastRow = Math.max(lastRow, rd.lastRow);
  ncols = Math.max(ncols, 1);
  const columns = buildColumns(headerVals, rd.firstCol, ncols);
  const sourceRows = Math.max(lastRow - headerRow, 0);
  return { columns, rowNos, data, empty, lastRow: Math.max(lastRow, headerRow), sourceRows };
}

export function columnNames(sf, sheet, headerRow) {
  return scanSheet(sf, sheet, headerRow).columns.map((c) => c.name);
}

/** Stage a sheet into DuckDB table `table` (__row BIGINT, c1.. VARCHAR). */
export async function stageSheet(db, sf, sheet, headerRow, table, progress = () => {}) {
  progress(0.05, `Reading ${sheet}`);
  const s = scanSheet(sf, sheet, headerRow);
  const defs = s.columns.map((c) => `${c.internal} VARCHAR`).join(', ');
  await db.exec(`DROP TABLE IF EXISTS ${table}`);
  await db.exec(`CREATE TABLE ${table} (__row BIGINT, ${defs})`);
  const total = s.rowNos.length;
  for (let start = 0; start < total; start += BATCH) {
    const end = Math.min(total, start + BATCH);
    const vecs = { __row: vectorFromArray(s.rowNos.slice(start, end).map(BigInt), new Int64()) };
    s.columns.forEach((c, j) => {
      const arr = new Array(end - start);
      for (let k = start; k < end; k++) { const v = s.data[k][j]; arr[k - start] = v === undefined ? null : v; }
      vecs[c.internal] = vectorFromArray(arr, new Utf8());
    });
    const tmp = `${table}__in`;
    await db.insertIPC(tmp, tableToIPC(new Table(vecs), 'stream'));
    await db.exec(`INSERT INTO ${table} SELECT * FROM ${tmp}`);
    await db.exec(`DROP TABLE ${tmp}`);
    progress(0.1 + 0.85 * (end / Math.max(total, 1)), `Loading ${sheet}: row ${s.rowNos[end - 1].toLocaleString('en-IN')}`);
  }
  progress(1, `Loaded ${sheet}`);
  return { table, fileName: sf.name, sheet, headerRow, columns: s.columns, sourceRows: s.sourceRows,
    stagedRows: total, emptyRows: s.sourceRows - total, lastRow: s.lastRow, fmt: sf.fmt };
}

/** Formula text of an .xlsx/.xlsm sheet as table (__row, col, formula). */
export async function stageFormulas(db, sf, staged, table) {
  await db.exec(`DROP TABLE IF EXISTS ${table}`);
  await db.exec(`CREATE TABLE ${table} (__row BIGINT, col INTEGER, formula VARCHAR)`);
  if (sf.fmt !== 'xlsx' && sf.fmt !== 'xlsm') return 0;
  const ws = sf.wb.Sheets[staged.sheet];
  const data = ws['!data'] || [];
  const rows = []; const cols = []; const forms = [];
  const lo = staged.columns[0].col - 1; const hi = staged.columns[staged.columns.length - 1].col - 1;
  for (let r = staged.headerRow; r < data.length; r++) {
    const row = data[r];
    if (!row) continue;
    for (let c = lo; c <= hi; c++) {
      const cell = row[c];
      if (cell && cell.f) { rows.push(BigInt(r + 1)); cols.push(c + 1); forms.push('=' + cell.f); }
    }
  }
  if (rows.length) {
    const t = new Table({ __row: vectorFromArray(rows, new Int64()), col: vectorFromArray(cols),
      formula: vectorFromArray(forms, new Utf8()) });
    await db.insertIPC(`${table}__in`, tableToIPC(t, 'stream'));
    await db.exec(`INSERT INTO ${table} SELECT __row, col::INTEGER, formula FROM ${table}__in`);
    await db.exec(`DROP TABLE ${table}__in`);
  }
  return rows.length;
}

export function previewRows(sf, sheet, headerRow, n = 50) {
  const { rows, firstCol } = readTopRows(sf, sheet, headerRow + n);
  const hdr = (rows.find(([r]) => r === headerRow) || [0, []])[1];
  const width = Math.max(1, hdr.length, ...rows.map(([, r]) => r.length));
  const cols = buildColumns(hdr, firstCol, width);
  return { columns: cols, rows: rows.filter(([r]) => r > headerRow) };
}
