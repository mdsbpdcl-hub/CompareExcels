// Output 1: a copy of File B with changed cells filled #FFF2CC.
// .xlsx/.xlsm: the package is copied part by part; only styles.xml (new fill + cloned formats) and the
// compared sheets' cell style indexes change. Values, formulas, widths, merges, images etc. stay as they are.
// .csv/.xls/.xlsb: converted to a new .xlsx (values only).
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import * as XLSX from '@e965/xlsx';
import { cellRef, colNumber, splitLocation } from './coords.js';
import { FastWorkbook, S, safeSheetTitle, concatChunks } from './xlsxWriter.js';
import { parseCsv } from './csv.js';
import { errorText } from './loader.js';

export const YELLOW = 'FFF2CC';
export const ORANGE = 'FCE4D6';
const BLANK = 'Blank key – not compared';

export async function collectTargets(db, run) {
  const yellow = {}; const orange = {};
  const add = (m, sheet, row, col) => { ((m[sheet] ||= {})[row] ||= new Set()).add(col); };
  for (const pr of run.pairs) {
    const sheet = pr.b.sheet;
    yellow[sheet] ||= {};
    for (const r of await db.query(`SELECT b_row, b_col FROM diffs WHERE pair_id = ${pr.pairId}`)) add(yellow, sheet, r.b_row, r.b_col);
    if (run.job.highlightOnlyB && pr.b.columns.length) {
      const lo = pr.b.columns[0].col; const hi = pr.b.columns[pr.b.columns.length - 1].col;
      for (const r of await db.query(`SELECT src_row FROM p${pr.pairId}_ob WHERE status <> '${BLANK}'`)) {
        for (let c = lo; c <= hi; c++) add(orange, sheet, r.src_row, c);
      }
    }
  }
  return { yellow, orange };
}

export function outputName(requested, src) {
  const stem = (requested || '').replace(/\.[^.]+$/, '') || 'File_B_Comparison_Highlighted';
  return stem + (src.fmt === 'xlsm' ? '.xlsm' : '.xlsx');
}

function makeResult(method) {
  return { method, notes: [], yellowCells: 0, orangeCells: 0, sheetParts: {}, sheetTitles: {}, preExistingYellow: 0,
    locate(sheet, row, col) {
      const parts = this.sheetParts[sheet] || [this.sheetTitles[sheet] || sheet];
      const [part, r] = parts.length > 1 ? splitLocation(row) : [1, row];
      return [parts[part - 1], cellRef(col, r)];
    } };
}
const countCells = (m) => Object.values(m).reduce((s, rows) => s + Object.values(rows).reduce((t, c) => t + c.size, 0), 0);

export async function buildHighlighted(db, run, progress = () => {}) {
  const src = run.job.fileB;
  const { yellow, orange } = await collectTargets(db, run);
  let res; let bytes;
  if (src.fmt === 'xlsx' || src.fmt === 'xlsm') {
    try {
      ({ res, bytes } = patchXlsx(src.bytes, yellow, orange, progress));
      res.notes.unshift('Copy made from the original File B; only the fill of highlighted cells changed.');
    } catch (e) {
      ({ res, bytes } = convertToXlsx(src, run, yellow, orange, progress));
      bytes = concatChunks(bytes);
      res.notes.unshift(`The workbook could not be copied part by part (${e.message || e}), so a values-only copy was made `
        + 'instead. Formulas, styles, widths and merged cells are not in this copy.');
    }
  } else {
    ({ res, bytes } = convertToXlsx(src, run, yellow, orange, progress));
    bytes = concatChunks(bytes);
  }
  res.yellowCells = countCells(yellow);
  res.orangeCells = countCells(orange);
  return { res, bytes };
}

// ===================================================================== xlsx patching
const attrRe = (name) => new RegExp(`\\s${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s*=\\s*"([^"]*)"`);
export function getAttr(tag, name) { const m = attrRe(name).exec(tag); return m ? m[1] : null; }
function setAttr(tag, name, value) {
  const re = attrRe(name);
  if (re.test(tag)) return tag.replace(re, ` ${name}="${value}"`);
  const end = tag.endsWith('/>') ? tag.length - 2 : tag.length - 1;
  return `${tag.slice(0, end)} ${name}="${value}"${tag.slice(end)}`;
}
const delAttr = (tag, name) => tag.replace(attrRe(name), '');
const unesc = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const reEsc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

export function resolveSheetParts(files) {
  const wb = strFromU8(files['xl/workbook.xml']);
  const rels = strFromU8(files['xl/_rels/workbook.xml.rels']);
  const relMap = {};
  for (const m of rels.matchAll(/<(?:\w+:)?Relationship\b[^>]*>/g)) {
    const tag = m[0];
    const id = getAttr(tag, 'Id'); const target = getAttr(tag, 'Target'); const type = getAttr(tag, 'Type') || '';
    if (!id || !target) continue;
    let path;
    if (target.startsWith('/')) path = target.slice(1);
    else {
      const parts = ['xl', ...target.split('/')]; const out = [];
      for (const p of parts) { if (p === '..') out.pop(); else if (p !== '.') out.push(p); }
      path = out.join('/');
    }
    relMap[id] = [path, type];
  }
  const sheets = {};
  for (const m of wb.matchAll(/<(?:\w+:)?sheet\b[^>]*>/g)) {
    const tag = m[0];
    const name = getAttr(tag, 'name');
    let rid = null;
    for (const am of tag.matchAll(/\s(\w+:id)\s*=\s*"([^"]*)"/g)) rid = am[2];
    if (name != null && relMap[rid]) sheets[unesc(name)] = relMap[rid][0];
  }
  const styles = Object.values(relMap).find(([, t]) => t.endsWith('/styles'));
  return { sheets, styles: styles ? styles[0] : 'xl/styles.xml' };
}

class StylePatcher {
  constructor(xml) {
    this.xml = xml;
    const m = /<(\w+:)?styleSheet\b/.exec(xml);
    this.p = m && m[1] ? m[1] : '';
    const p = reEsc(this.p);
    this.fillsRe = new RegExp(`<${p}fills\\b[^>]*?(?:/>|>([\\s\\S]*?)</${p}fills>)`);
    this.xfsRe = new RegExp(`<${p}cellXfs\\b[^>]*?(?:/>|>([\\s\\S]*?)</${p}cellXfs>)`);
    const fm = this.fillsRe.exec(xml); const xm = this.xfsRe.exec(xml);
    if (!fm || !xm) throw new Error('styles.xml has no fills or cell formats');
    this.fillItems = (fm[1] || '').match(new RegExp(`<${p}fill\\b[^>]*?/>|<${p}fill\\b[^>]*?>[\\s\\S]*?</${p}fill>`, 'g')) || [];
    this.xfItems = (xm[1] || '').match(new RegExp(`<${p}xf\\b[^>]*?/>|<${p}xf\\b[^>]*?>[\\s\\S]*?</${p}xf>`, 'g')) || [];
    if (!this.xfItems.length) throw new Error('styles.xml has no cell formats');
    this.origXf = this.xfItems.length;
    this.newFills = []; this.fillIds = {}; this.clones = new Map(); this.newXfs = [];
  }
  fillId(rgb) {
    if (!(rgb in this.fillIds)) {
      const p = this.p;
      this.newFills.push(`<${p}fill><${p}patternFill patternType="solid"><${p}fgColor rgb="FF${rgb}"/><${p}bgColor indexed="64"/></${p}patternFill></${p}fill>`);
      this.fillIds[rgb] = this.fillItems.length + this.newFills.length - 1;
    }
    return this.fillIds[rgb];
  }
  clone(s, rgb) {
    const key = `${s}|${rgb}`;
    if (!this.clones.has(key)) {
      const base = s >= 0 && s < this.xfItems.length ? this.xfItems[s] : this.xfItems[0];
      const open = /^<[^>]*?\/?>/.exec(base)[0];
      const newOpen = setAttr(setAttr(open, 'fillId', String(this.fillId(rgb))), 'applyFill', '1');
      this.newXfs.push(newOpen + base.slice(open.length));
      this.clones.set(key, this.origXf + this.newXfs.length - 1);
    }
    return this.clones.get(key);
  }
  render() {
    let xml = this.xml; const p = this.p;
    if (this.newFills.length) {
      const fm = this.fillsRe.exec(xml);
      let open = /^<[^>]*?\/?>/.exec(fm[0])[0];
      if (open.endsWith('/>')) open = open.slice(0, -2) + '>';
      open = setAttr(open, 'count', String(this.fillItems.length + this.newFills.length));
      xml = xml.slice(0, fm.index) + open + this.fillItems.join('') + this.newFills.join('') + `</${p}fills>` + xml.slice(fm.index + fm[0].length);
    }
    if (this.newXfs.length) {
      const xm = this.xfsRe.exec(xml);
      const open = setAttr(/^<[^>]*?\/?>/.exec(xm[0])[0], 'count', String(this.xfItems.length + this.newXfs.length));
      xml = xml.slice(0, xm.index) + open + this.xfItems.join('') + this.newXfs.join('') + `</${p}cellXfs>` + xml.slice(xm.index + xm[0].length);
    }
    return xml;
  }
}

function colStyles(xml, p) {
  const out = {};
  const m = new RegExp(`<${reEsc(p)}cols\\b[^>]*>([\\s\\S]*?)</${reEsc(p)}cols>`).exec(xml);
  if (!m) return out;
  for (const cm of m[1].matchAll(new RegExp(`<${reEsc(p)}col\\b[^>]*>`, 'g'))) {
    const s = getAttr(cm[0], 'style'); const lo = getAttr(cm[0], 'min'); const hi = getAttr(cm[0], 'max');
    if (s && lo && hi && hi - lo < 20000) for (let c = +lo; c <= +hi; c++) out[c] = +s;
  }
  return out;
}

/** targets: Map row -> Map col -> rgb */
function patchSheet(xml, targets, styles) {
  const m = /<(\w+:)?worksheet\b/.exec(xml);
  const p = m && m[1] ? m[1] : '';
  const ep = reEsc(p);
  const cs = colStyles(xml, p);
  const sd = new RegExp(`<${ep}sheetData\\b[^>]*?(/?)>`).exec(xml);
  if (!sd) throw new Error('sheet has no sheetData');
  let head; let tail; let body;
  if (sd[1] === '/') {
    head = xml.slice(0, sd.index) + `<${p}sheetData>`; tail = `</${p}sheetData>` + xml.slice(sd.index + sd[0].length); body = '';
  } else {
    const bs = sd.index + sd[0].length;
    const be = xml.indexOf(`</${p}sheetData>`, bs);
    if (be < 0) throw new Error('sheetData is not closed');
    head = xml.slice(0, bs); tail = xml.slice(be); body = xml.slice(bs, be);
  }
  const rowRe = new RegExp(`<${ep}row\\b([^>]*?)(/>|>([\\s\\S]*?)</${ep}row>)`, 'g');
  const cellRe = new RegExp(`<${ep}c\\b([^>]*?)(/>|>([\\s\\S]*?)</${ep}c>)`, 'g');
  const pending = [...targets.keys()].sort((x, y) => x - y);
  let pi = 0; let pos = 0; let prevRow = 0;
  const out = [];
  const newRow = (r, cols) => {
    const cells = [...cols.keys()].sort((x, y) => x - y).map((c) => `<${p}c r="${cellRef(c, r)}" s="${styles.clone(cs[c] || 0, cols.get(c))}"/>`);
    return `<${p}row r="${r}">${cells.join('')}</${p}row>`;
  };
  const patchRow = (attrs, inner, r, cols) => {
    let rowTag = delAttr(`<${p}row${attrs}>`, 'spans');
    const custom = getAttr(rowTag, 'customFormat');
    const rowS = custom === '1' || custom === 'true' ? getAttr(rowTag, 's') : null;
    const remaining = new Map(cols);
    const parts = []; let cpos = 0; let prevCol = 0;
    cellRe.lastIndex = 0;
    let cm;
    while ((cm = cellRe.exec(inner)) !== null) {
      const cattrs = cm[1];
      const ref = getAttr(' ' + cattrs, 'r');
      const col = ref ? colNumber(/\$?([A-Za-z]+)/.exec(ref)[1]) : prevCol + 1;
      prevCol = col;
      for (const c of [...remaining.keys()].filter((k) => k < col).sort((x, y) => x - y)) {
        const base = rowS != null ? +rowS : (cs[c] || 0);
        parts.push(inner.slice(cpos, cm.index)); cpos = cm.index;
        parts.push(`<${p}c r="${cellRef(c, r)}" s="${styles.clone(base, remaining.get(c))}"/>`);
        remaining.delete(c);
      }
      if (remaining.has(col)) {
        const oldS = getAttr(' ' + cattrs, 's');
        const base = oldS != null ? +oldS : rowS != null ? +rowS : (cs[col] || 0);
        const newOpen = setAttr(`<${p}c${cattrs}>`, 's', String(styles.clone(base, remaining.get(col)))).slice(0, -1);
        parts.push(inner.slice(cpos, cm.index)); parts.push(newOpen);
        cpos = cm.index + `<${p}c`.length + cattrs.length;
        remaining.delete(col);
      }
    }
    parts.push(inner.slice(cpos));
    for (const c of [...remaining.keys()].sort((x, y) => x - y)) {
      const base = rowS != null ? +rowS : (cs[c] || 0);
      parts.push(`<${p}c r="${cellRef(c, r)}" s="${styles.clone(base, remaining.get(c))}"/>`);
    }
    return rowTag + parts.join('') + `</${p}row>`;
  };
  let rm;
  while ((rm = rowRe.exec(body)) !== null) {
    const attrs = rm[1];
    const rAttr = getAttr(' ' + attrs, 'r');
    const r = rAttr ? +rAttr : prevRow + 1;
    prevRow = r;
    while (pi < pending.length && pending[pi] < r) {
      out.push(body.slice(pos, rm.index)); pos = rm.index;
      out.push(newRow(pending[pi], targets.get(pending[pi]))); pi++;
    }
    if (pi < pending.length && pending[pi] === r) {
      out.push(body.slice(pos, rm.index));
      out.push(patchRow(attrs, rm[3] || '', r, targets.get(r)));
      pos = rm.index + rm[0].length; pi++;
    }
  }
  out.push(body.slice(pos));
  while (pi < pending.length) { out.push(newRow(pending[pi], targets.get(pending[pi]))); pi++; }
  return head + out.join('') + tail;
}

export function patchXlsx(bytes, yellow, orange, progress = () => {}) {
  const res = makeResult('patched');
  const files = unzipSync(bytes);
  const { sheets: sheetPaths, styles: stylesPath } = resolveSheetParts(files);
  const styles = new StylePatcher(strFromU8(files[stylesPath]));
  const names = [...new Set([...Object.keys(yellow), ...Object.keys(orange)])].sort();
  const out = { ...files };
  names.forEach((sheet, i) => {
    if (!(sheet in sheetPaths)) throw new Error(`sheet '${sheet}' not found in package`);
    progress(0.1 + 0.7 * i / Math.max(names.length, 1), `Highlighting ${sheet}`);
    const per = new Map();
    const put = (m, rgb) => { for (const [r, cols] of Object.entries(m || {})) for (const c of cols) {
      if (!per.has(+r)) per.set(+r, new Map()); per.get(+r).set(c, rgb); } };
    put(orange[sheet], ORANGE); put(yellow[sheet], YELLOW);
    if (!per.size) return;
    out[sheetPaths[sheet]] = strToU8(patchSheet(strFromU8(files[sheetPaths[sheet]]), per, styles));
  });
  if (styles.newXfs.length) out[stylesPath] = strToU8(styles.render());
  progress(0.85, 'Writing the highlighted copy');
  const zipped = zipSync(Object.fromEntries(Object.entries(out).map(([k, v]) => [k, [v, { level: 6 }]])));
  for (const n of Object.keys(sheetPaths)) { res.sheetTitles[n] = n; res.sheetParts[n] = [n]; }
  return { res, bytes: zipped };
}

// ===================================================================== conversion
function convertToXlsx(src, run, yellow, orange, progress) {
  const res = makeResult('converted');
  const taken = new Set();
  const wb = new FastWorkbook();
  const fill = (yr, or, r, c) => (yr[r] && yr[r].has(c) ? 'Y' : or[r] && or[r].has(c) ? 'O' : '');
  if (src.fmt === 'csv') {
    const sheet = src.sheets[0];
    const headerRow = (run.pairs.find((p) => p.b.sheet === sheet) || { b: { headerRow: 1 } }).b.headerRow;
    const yr = yellow[sheet] || {}; const or = orange[sheet] || {};
    const titles = []; let ws = null; let headerCells = []; let rno = 0;
    for (const rec of parseCsv(src.text, src.delimiter)) {
      rno++;
      const [part] = splitLocation(rno);
      if (part > titles.length) {
        const t = safeSheetTitle(part === 1 ? sheet : `${sheet.slice(0, 24)} (${part})`, taken);
        titles.push(t); ws = wb.addSheet(t, { freeze: false });
        if (part > 1) ws.append(headerCells);
      }
      const st = (c) => ({ Y: S.YELLOW, O: S.ORANGE, '': 0 }[fill(yr, or, rno, c)]);
      const cells = rec.map((v, j) => [v === '' ? null : v, st(j + 1)]);
      const extra = [...new Set([...(yr[rno] || []), ...(or[rno] || [])])].filter((c) => c > rec.length);
      if (extra.length) for (let j = rec.length + 1; j <= Math.max(...extra); j++) cells.push([null, st(j)]);
      if (rno === headerRow) headerCells = cells.map(([v]) => [v, 0]);
      ws.append(cells);
    }
    if (!titles.length) { titles.push(safeSheetTitle(sheet, taken)); wb.addSheet(titles[0], { freeze: false }); }
    res.sheetParts[sheet] = titles; res.sheetTitles[sheet] = titles[0];
    if (titles.length > 1) res.notes.push(`File B has more rows than one Excel sheet can hold, so the copy is split into `
      + `${titles.length} sheets (${titles.join(', ')}). Each later part repeats the heading row; the report shows where every changed cell landed.`);
    res.notes.push('File B was a CSV file, so the copy is a new .xlsx. Every value is kept as text exactly as in the CSV. '
      + 'There were no formulas, styles or widths to keep.');
  } else {
    src.sheets.forEach((sheet, i) => {
      progress(0.1 + 0.7 * i / Math.max(src.sheets.length, 1), `Copying sheet ${sheet}`);
      const t = safeSheetTitle(sheet, taken);
      res.sheetTitles[sheet] = t; res.sheetParts[sheet] = [t];
      const ws = wb.addSheet(t, { freeze: false });
      const yr = yellow[sheet] || {}; const or = orange[sheet] || {};
      const data = src.wb.Sheets[sheet]['!data'] || [];
      const lastData = data.length;
      const extraRows = [...new Set([...Object.keys(yr), ...Object.keys(or)].map(Number))].filter((r) => r > lastData);
      const last = Math.max(lastData, ...extraRows, 0);
      for (let r = 1; r <= last; r++) {
        const row = data[r - 1] || [];
        const marks = new Set([...(yr[r] || []), ...(or[r] || [])]);
        const width = Math.max(row.length, ...[...marks], 0);
        const cells = [];
        for (let c = 1; c <= width; c++) {
          const cell = row[c - 1];
          const f = fill(yr, or, r, c);
          let v = null; let st = f === 'Y' ? S.YELLOW : f === 'O' ? S.ORANGE : 0;
          if (cell && cell.t !== 'z' && cell.v !== undefined && cell.v !== null) {
            if (cell.t === 'n') {
              v = cell.v;
              const f2 = cell.z && typeof cell.z === 'string' ? cell.z.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '') : '';
              if (f2 && XLSX.SSF.is_date(cell.z)) {
                const timeOnly = !/[yd]/i.test(f2) && /[hs]/i.test(f2) && v >= 0 && v < 1;
                const isDt = v % 1 !== 0;
                const kind = timeOnly ? 'TIME' : isDt ? 'DATETIME' : 'DATE';
                st = S[(f === 'Y' ? 'Y_' : f === 'O' ? 'O_' : '') + kind];
              }
            } else if (cell.t === 'b') v = !!cell.v;
            else if (cell.t === 'e') v = { error: errorText(cell) };
            else v = String(cell.v);
          }
          cells.push([v, st]);
        }
        ws.append(cells);
      }
    });
    if (src.fmt === 'xls' || src.fmt === 'xlsb') res.notes.push(`File B was ${src.fmt === 'xls' ? 'Excel 97-2003' : 'Excel binary'} (.${src.fmt}), so the copy is a new .xlsx with `
      + 'the same sheets, names and cell positions. Values are kept; formulas are kept as their last calculated values, '
      + 'and styles, widths, merged cells and macros are not carried over. This is not a full-fidelity copy.');
  }
  const cleaned = wb.sheets.reduce((s, x) => s + x.cleaned, 0);
  if (cleaned) res.notes.push(`${cleaned.toLocaleString('en-IN')} cell(s) held characters Excel cannot store or more than `
    + '32,767 characters; those characters were replaced with � or the text was shortened.');
  progress(0.9, 'Saving the highlighted copy');
  const chunks = wb.finish();
  return { res, bytes: chunks };
}

// ===================================================================== reading highlights back
/** {sheet: Set(refs)} of cells whose fill is rgb. Independent scan of the saved package. */
export function readFilledCells(bytes, rgb = YELLOW) {
  const files = unzipSync(bytes);
  const { sheets, styles } = resolveSheetParts(files);
  const st = strFromU8(files[styles]);
  const fillsBlock = /<(?:\w+:)?fills\b[^>]*>([\s\S]*?)<\/(?:\w+:)?fills>/.exec(st);
  const fills = fillsBlock ? fillsBlock[1].match(/<(?:\w+:)?fill\b[^>]*?\/>|<(?:\w+:)?fill\b[^>]*?>[\s\S]*?<\/(?:\w+:)?fill>/g) || [] : [];
  const targetFills = new Set();
  fills.forEach((f, i) => {
    const fg = /<(?:\w+:)?fgColor\b[^>]*>/.exec(f);
    if (fg && (getAttr(fg[0], 'rgb') || '').toUpperCase().endsWith(rgb.toUpperCase())) targetFills.add(i);
  });
  const xfsBlock = /<(?:\w+:)?cellXfs\b[^>]*>([\s\S]*?)<\/(?:\w+:)?cellXfs>/.exec(st);
  const xfs = xfsBlock ? xfsBlock[1].match(/<(?:\w+:)?xf\b[^>]*?\/?>/g) || [] : [];
  const targetXfs = new Set();
  xfs.forEach((x, i) => { if (targetFills.has(+(getAttr(x, 'fillId') || 0))) targetXfs.add(i); });
  const out = {};
  for (const [name, part] of Object.entries(sheets)) {
    const found = new Set();
    if (targetXfs.size) {
      const xml = strFromU8(files[part]);
      let row = 0; let col = 0;
      for (const m of xml.matchAll(/<(?:\w+:)?(row|c)\b([^>]*?)\/?>/g)) {
        if (m[1] === 'row') { const r = getAttr(' ' + m[2], 'r'); row = r ? +r : row + 1; col = 0; continue; }
        const ref = getAttr(' ' + m[2], 'r');
        if (ref) col = colNumber(/\$?([A-Za-z]+)/.exec(ref)[1]); else col++;
        if (targetXfs.has(+(getAttr(' ' + m[2], 's') || 0))) found.add(ref ? ref.replace(/\$/g, '') : cellRef(col, row));
      }
    }
    out[name] = found;
  }
  return out;
}
