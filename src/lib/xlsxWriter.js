// A small streaming .xlsx writer for workbooks this app builds from scratch (the report and
// conversions). Every cell uses one of a few fixed styles; sheets are zipped as they are written.
// Text is always an inline string (never a formula), so '=HYPERLINK(...)' or '@SUM(1)' can never run.
import { Zip, ZipDeflate, strToU8 } from 'fflate';
import { EXCEL_MAX_ROWS, EXCEL_MAX_CELL_CHARS, colLetter } from './coords.js';

export const S = { DEFAULT: 0, BODY: 1, BODY_NUM: 2, HEAD: 3, TITLE: 4, SECTION: 5, PASS: 6, FAIL: 7, BODY_INT: 8,
  YELLOW: 9, ORANGE: 10, BODY_INT_LEFT: 11, DATE: 12, DATETIME: 13, Y_DATE: 14, Y_DATETIME: 15, O_DATE: 16, O_DATETIME: 17,
  TIME: 18, Y_TIME: 19, O_TIME: 20 };

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="6"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
<font><sz val="12"/><name val="Times New Roman"/><family val="1"/></font>
<font><b/><sz val="12"/><color rgb="FF1F3864"/><name val="Times New Roman"/><family val="1"/></font>
<font><b/><sz val="14"/><color rgb="FF1F3864"/><name val="Times New Roman"/><family val="1"/></font>
<font><b/><sz val="12"/><color rgb="FF16A34A"/><name val="Times New Roman"/><family val="1"/></font>
<font><b/><sz val="12"/><color rgb="FFDC2626"/><name val="Times New Roman"/><family val="1"/></font></fonts>
<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFF8D6"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFCE4D6"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color rgb="FFCCCCCC"/></left><right style="thin"><color rgb="FFCCCCCC"/></right><top style="thin"><color rgb="FFCCCCCC"/></top><bottom style="thin"><color rgb="FFCCCCCC"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="21">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="49" fontId="1" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="4" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="5" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="3" fontId="1" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="top"/></xf>
<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>
<xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/>
<xf numFmtId="3" fontId="1" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top"/></xf>
<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="14" fontId="0" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/>
<xf numFmtId="22" fontId="0" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/>
<xf numFmtId="14" fontId="0" fillId="4" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/>
<xf numFmtId="22" fontId="0" fillId="4" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/>
<xf numFmtId="21" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="21" fontId="0" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/>
<xf numFmtId="21" fontId="0" fillId="4" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/>
</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
// eslint-disable-next-line no-control-regex
const ILLEGAL = /[\u0000-\u0008\u000b\u000c\u000e-\u001f\ud800-\udfff￾￿]/g;
const xesc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** Make text storable in .xlsx. Returns [text, changed]. Lone surrogates are removed too. */
export function cleanText(v) {
  let s = String(v);
  let changed = false;
  if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/.test(s) || /[\ud800-\udfff]/.test(s)) {
    // keep valid surrogate pairs
    const t = s.replace(/[\ud800-\udbff][\udc00-\udfff]|([\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿\ud800-\udfff])/g,
      (m, bad) => (bad ? '�' : m));
    changed = t !== s; s = t;
  }
  if (s.length > EXCEL_MAX_CELL_CHARS) { s = s.slice(0, EXCEL_MAX_CELL_CHARS - 20) + ' …[shortened]'; changed = true; }
  return [s, changed];
}
void ILLEGAL;

export const INVALID_SHEET_CHARS = /[[\]*?/\\:]/g;
export function safeSheetTitle(title, taken) {
  const base = (title.replace(INVALID_SHEET_CHARS, '_').replace(/^'+|'+$/g, '').slice(0, 31)) || 'Sheet';
  let name = base; let k = 2;
  const lower = () => new Set([...taken].map((t) => t.toLowerCase()));
  while (lower().has(name.toLowerCase())) { const suf = ` (${k})`; name = base.slice(0, 31 - suf.length) + suf; k++; }
  taken.add(name);
  return name;
}

export class FastWorkbook {
  constructor() {
    this.chunks = [];
    this.sheets = [];
    this.zip = new Zip((err, data) => { if (err) throw err; this.chunks.push(data); });
    this._open = null;
  }

  addSheet(title, { widths = null, freeze = true, autofilterCols = null } = {}) {
    if (this._open) this._open.close();
    const sh = new FastSheet(this, title, this.sheets.length, widths, freeze, autofilterCols);
    this.sheets.push(sh);
    this._open = sh;
    return sh;
  }

  _addFile(name, text) {
    const f = new ZipDeflate(name, { level: 6 });
    this.zip.add(f);
    f.push(strToU8(text), true);
  }

  /** Returns an array of Uint8Array chunks (make a Blob in the browser). */
  finish() {
    if (this._open) this._open.close();
    const n = this.sheets.length;
    const ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'];
    for (let i = 1; i <= n; i++) ct.push(`<Override PartName="/xl/worksheets/sheet${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
    ct.push('</Types>');
    const rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
    const sheetsXml = this.sheets.map((s, i) => `<sheet name="${xesc(s.title)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('');
    const names = this.sheets.map((s, i) => (s.autofilterCols
      ? `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">${xesc(`'${s.title.replace(/'/g, "''")}'`)}!$A$1:$${colLetter(s.autofilterCols)}$${Math.max(s.row, 1)}</definedName>`
      : '')).join('');
    const wb = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook ${NS}><bookViews><workbookView activeTab="0"/></bookViews>`
      + `<sheets>${sheetsXml}</sheets>${names ? `<definedNames>${names}</definedNames>` : ''}</workbook>`;
    let wbRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    for (let i = 1; i <= n; i++) wbRels += `<Relationship Id="rId${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i}.xml"/>`;
    wbRels += `<Relationship Id="rId${n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
    this._addFile('[Content_Types].xml', ct.join(''));
    this._addFile('_rels/.rels', rels);
    this._addFile('xl/workbook.xml', wb);
    this._addFile('xl/_rels/workbook.xml.rels', wbRels);
    this._addFile('xl/styles.xml', STYLES);
    this.zip.end();
    return this.chunks;
  }
}

class FastSheet {
  constructor(book, title, index, widths, freeze, autofilterCols) {
    this.book = book; this.title = title; this.index = index; this.autofilterCols = autofilterCols;
    this.row = 0; this.cleaned = 0; this.closed = false; this.buf = [];
    this.file = new ZipDeflate(`xl/worksheets/sheet${index + 1}.xml`, { level: 6 });
    book.zip.add(this.file);
    const sel = index === 0 ? ' tabSelected="1"' : '';
    let head = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet ${NS}>`;
    head += freeze
      ? `<sheetViews><sheetView workbookViewId="0"${sel}><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>`
      : `<sheetViews><sheetView workbookViewId="0"${sel}/></sheetViews>`;
    head += '<sheetFormatPr defaultRowHeight="15"/>';
    if (widths) head += '<cols>' + widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('') + '</cols>';
    head += '<sheetData>';
    this.buf.push(head);
  }

  /** cells: array of [value, style]. string -> text, number -> number, boolean, null -> blank. */
  append(cells) {
    if (this.row >= EXCEL_MAX_ROWS) throw new Error('sheet is full');
    const r = ++this.row;
    let s = `<row r="${r}">`;
    for (let j = 0; j < cells.length; j++) {
      const [v, st] = cells[j];
      if (v === null || v === undefined) { if (st) s += `<c r="${colLetter(j + 1)}${r}" s="${st}"/>`; continue; }
      if (typeof v === 'object' && v.error) s += `<c r="${colLetter(j + 1)}${r}" s="${st}" t="e"><v>${xesc(v.error)}</v></c>`;
      else if (typeof v === 'boolean') s += `<c r="${colLetter(j + 1)}${r}" s="${st}" t="b"><v>${v ? 1 : 0}</v></c>`;
      else if (typeof v === 'number' && Number.isFinite(v)) s += `<c r="${colLetter(j + 1)}${r}" s="${st}"><v>${v}</v></c>`;
      else {
        const [t, ch] = cleanText(v);
        if (ch) this.cleaned++;
        s += `<c r="${colLetter(j + 1)}${r}" s="${st}" t="inlineStr"><is><t xml:space="preserve">${xesc(t)}</t></is></c>`;
      }
    }
    this.buf.push(s + '</row>');
    if (this.buf.length >= 2000) this._flush();
  }

  _flush() { if (this.buf.length) { this.file.push(strToU8(this.buf.join(''))); this.buf = []; } }

  get filterRef() { return `A1:${colLetter(Math.max(this.autofilterCols || 1, 1))}${Math.max(this.row, 1)}`; }

  close() {
    if (this.closed) return;
    this.buf.push('</sheetData>');
    if (this.autofilterCols) this.buf.push(`<autoFilter ref="${this.filterRef}"/>`);
    this.buf.push('<pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>');
    this.file.push(strToU8(this.buf.join('')), true);
    this.buf = [];
    this.closed = true;
    if (this.book._open === this) this.book._open = null;
  }
}

export class TableSheetWriter {
  constructor(wb, title, headers, widths, taken, kinds = null, maxRows = EXCEL_MAX_ROWS) {
    Object.assign(this, { wb, baseTitle: title, headers, widths, taken, kinds: kinds || headers.map(() => 'text'), maxRows });
    this.sheets = []; this.rowsWritten = 0; this.ws = null;
    this._newSheet();
  }
  _newSheet() {
    const part = this.sheets.length + 1;
    const title = safeSheetTitle(part === 1 ? this.baseTitle : `${this.baseTitle.slice(0, 24)} (${part})`, this.taken);
    this.ws = this.wb.addSheet(title, { widths: this.widths, freeze: true, autofilterCols: this.headers.length });
    this.ws.append(this.headers.map((h) => [h, S.HEAD]));
    this.sheets.push(title);
  }
  append(values) {
    if (this.ws.row >= this.maxRows) this._newSheet();
    this.ws.append(values.map((v, k) => (this.kinds[k] === 'num' && typeof v === 'number'
      ? [v, S.BODY_NUM] : [v === null || v === undefined ? null : String(v), S.BODY])));
    this.rowsWritten++;
  }
  close() { this.ws.close(); }
}

export function concatChunks(chunks) {
  const len = chunks.reduce((s, c) => s + c.length, 0);
  const out = new Uint8Array(len);
  let o = 0;
  for (const c of chunks) { out.set(c, o); o += c.length; }
  return out;
}
