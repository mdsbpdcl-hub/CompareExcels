// Independent checks that prove the outputs are complete and correct (PASS / FAIL).
import * as XLSX from '@e965/xlsx';
import { unzipSync } from 'fflate';
import { cellRef, EXCEL_MAX_ROWS } from './coords.js';
import { sheetReader, cellToText, sha256 } from './loader.js';
import { parseCsv } from './csv.js';
import { readFilledCells, resolveSheetParts } from './highlight.js';

const BLANK = 'Blank key – not compared';
const SAMPLE = 300;
export const check = (name, passed, detail, critical = true) => ({ name, passed, detail, critical,
  status: critical ? (passed ? 'PASS' : 'FAIL') : 'NOTE' });
export const overall = (checks) => checks.every((c) => !c.critical || c.passed);
const fmt = (n) => Number(n).toLocaleString('en-IN');

export async function checkEngine(db, run) {
  const out = [];
  for (const pr of run.pairs) {
    const i = pr.pairId;
    const tag = run.pairs.length > 1 ? ` (pair ${i}: ${pr.label})` : '';
    const [{ n: matched }] = await db.query(`SELECT count(*) AS n FROM p${i}_m`);
    for (const [side, staged] of [['a', pr.a], ['b', pr.b]]) {
      const S = side.toUpperCase();
      const [o] = await db.query(`SELECT count(*) FILTER (WHERE status <> '${BLANK}') AS only, count(*) FILTER (WHERE status = '${BLANK}') AS blank FROM p${i}_o${side}`);
      const [d] = await db.query(`SELECT count(*) FILTER (WHERE __status = 'dup_excluded') AS ex, count(*) AS total FROM p${i}_${side}`);
      const acc = matched + o.only + o.blank + d.ex + staged.emptyRows;
      const ok = acc === staged.sourceRows && d.total === staged.stagedRows;
      out.push(check(`Every row of File ${S} is accounted for${tag}`, ok,
        `${fmt(staged.sourceRows)} rows below the heading = ${fmt(matched)} matched + ${fmt(o.only)} only in ${S} + ${fmt(o.blank)} blank key + `
        + `${fmt(d.ex)} duplicates excluded + ${fmt(staged.emptyRows)} empty rows${ok ? '' : ` → adds up to ${fmt(acc)}`}`));
    }
    const [m] = await db.query(`SELECT count(*) AS n, count(DISTINCT a_row) AS da, count(DISTINCT b_row) AS db FROM p${i}_m`);
    out.push(check(`Matched records are one-to-one${tag}`, m.n === m.da && m.da === m.db,
      `${fmt(m.n)} matched records, ${fmt(m.da)} distinct File A rows, ${fmt(m.db)} distinct File B rows`));
    const [d] = await db.query(`SELECT count(*) AS n, count(DISTINCT (b_row, b_col)) AS nb, count(DISTINCT (a_row, a_col)) AS na,
      count(DISTINCT (a_row, b_row)) AS recs FROM diffs WHERE pair_id = ${i}`);
    const [rw] = await db.query(`SELECT COALESCE(sum(n_changed), 0) AS cells, count(*) AS recs FROM rowwise WHERE pair_id = ${i}`);
    const ok = d.n === d.nb && d.nb === d.na && d.na === Number(rw.cells) && d.recs === rw.recs;
    out.push(check(`Difference counts agree${tag}`, ok, `${fmt(d.n)} changed cells, each File B cell listed once (${fmt(d.nb)} distinct); `
      + `${fmt(rw.recs)} changed records; row-wise total ${fmt(rw.cells)}`));
    const [bad] = await db.query(`SELECT count(*) AS n FROM diffs d WHERE pair_id = ${i} AND NOT EXISTS
      (SELECT 1 FROM p${i}_m m WHERE m.a_row = d.a_row AND m.b_row = d.b_row)`);
    out.push(check(`Differences come only from matched records${tag}`, bad.n === 0,
      bad.n === 0 ? 'All changed cells belong to matched records' : `${fmt(bad.n)} changed cells do not belong to a matched record`));
  }
  return out;
}

/** Read specific cells straight from a source file as comparison text. */
function fetchCells(sf, sheet, wanted) {
  const out = new Map();
  if (sf.fmt === 'csv') {
    const need = new Map();
    for (const [r, c] of wanted) { if (!need.has(r)) need.set(r, []); need.get(r).push(c); }
    const last = Math.max(0, ...need.keys());
    let rno = 0;
    for (const rec of parseCsv(sf.text, sf.delimiter)) {
      rno++;
      if (need.has(rno)) for (const c of need.get(rno)) { const v = rec[c - 1]; out.set(`${rno},${c}`, v === undefined || v === '' ? null : v); }
      if (rno >= last) break;
    }
  } else {
    const ws = sf.wb.Sheets[sheet];
    const date1904 = !!(sf.wb.Workbook && sf.wb.Workbook.WBProps && sf.wb.Workbook.WBProps.date1904);
    for (const [r, c] of wanted) {
      const cell = ws[XLSX.utils.encode_cell({ r: r - 1, c: c - 1 })] || (ws['!data'] && ws['!data'][r - 1] && ws['!data'][r - 1][c - 1]);
      out.set(`${r},${c}`, cellToText(cell, date1904));
    }
  }
  return out;
}

export async function checkCoordinates(db, run) {
  const out = [];
  for (const pr of run.pairs) {
    const i = pr.pairId;
    const tag = run.pairs.length > 1 ? ` (pair ${i})` : '';
    const [{ n: total }] = await db.query(`SELECT count(*) AS n FROM diffs WHERE pair_id = ${i}`);
    const rows = total ? await db.query(`SELECT * FROM (SELECT a_row, a_col, b_row, b_col, before_value, after_value FROM diffs
      WHERE pair_id = ${i}) s USING SAMPLE reservoir(${SAMPLE} ROWS) REPEATABLE (7)`) : [];
    const aWant = [...rows.map((r) => [r.a_row, r.a_col]), ...pr.compareCols.map((c) => [pr.a.headerRow, c.a_col])];
    const bWant = [...rows.map((r) => [r.b_row, r.b_col]), ...pr.compareCols.map((c) => [pr.b.headerRow, c.b_col])];
    const av = fetchCells(run.job.fileA, pr.a.sheet, aWant);
    const bv = fetchCells(run.job.fileB, pr.b.sheet, bWant);
    const bad = rows.filter((r) => av.get(`${r.a_row},${r.a_col}`) !== r.before_value || bv.get(`${r.b_row},${r.b_col}`) !== r.after_value)
      .map((r) => `${cellRef(r.a_col, r.a_row)}/${cellRef(r.b_col, r.b_row)}`);
    out.push(check(`Reported cells hold the reported values${tag}`, !bad.length,
      `Re-read ${fmt(rows.length)} of ${fmt(total)} changed cells from both original files${bad.length ? `; mismatches at ${bad.slice(0, 5).join(', ')}` : ''}`));
    const headBad = pr.compareCols.filter((c) => (av.get(`${pr.a.headerRow},${c.a_col}`) || '').trim() !== (c.a_header || '').trim()
      || (bv.get(`${pr.b.headerRow},${c.b_col}`) || '').trim() !== (c.b_header || '').trim()).map((c) => c.a_name);
    out.push(check(`Headings and column letters are correct${tag}`, !headBad.length,
      `Checked ${pr.compareCols.length} column heading(s) at their reported letters${headBad.length ? `; wrong: ${headBad.slice(0, 5).join(', ')}` : ''}`));
  }
  return out;
}

export async function checkHighlights(db, run, hl, outBytes) {
  const expected = {};
  for (const pr of run.pairs) {
    for (const r of await db.query(`SELECT b_row, b_col FROM diffs WHERE pair_id = ${pr.pairId}`)) {
      const [sheet, ref] = hl.locate(pr.b.sheet, r.b_row, r.b_col);
      (expected[sheet] ||= new Set()).add(ref);
    }
  }
  const found = readFilledCells(outBytes);
  let pre = {};
  if (hl.method !== 'converted') { try { pre = readFilledCells(run.job.fileB.bytes); } catch { pre = {}; } }
  hl.preExistingYellow = Object.values(pre).reduce((s, v) => s + v.size, 0);
  let missing = 0; let extra = 0; const ex = [];
  for (const sheet of new Set([...Object.keys(expected), ...Object.keys(found)])) {
    const exp = expected[sheet] || new Set(); const got = found[sheet] || new Set(); const p = pre[sheet] || new Set();
    for (const x of exp) if (!got.has(x)) { missing++; if (ex.length < 6) ex.push(`${sheet}!${x} not yellow`); }
    for (const x of got) if (!p.has(x) && !exp.has(x)) { extra++; if (ex.length < 6) ex.push(`${sheet}!${x} extra`); }
  }
  const nExp = Object.values(expected).reduce((s, v) => s + v.size, 0);
  const ok = !missing && !extra;
  let detail = `${fmt(nExp)} reported cells, all yellow in the copy, and no other cell was coloured`;
  if (!ok) detail = `${fmt(missing)} reported cell(s) not yellow, ${fmt(extra)} yellow cell(s) not in the report: ${ex.join('; ')}`;
  else if (hl.preExistingYellow) detail += ` (File B already had ${fmt(hl.preExistingYellow)} cell(s) in the same yellow; those were left alone)`;
  return [check('Highlights match the report exactly', ok, detail)];
}

const trim = (v) => { const a = [...v]; while (a.length && a[a.length - 1] === null) a.pop(); return a; };

/** The copy holds exactly the same values as File B. Reads the copy with SheetJS (not our writer). */
export function checkBUnchanged(run, hl, outBytes) {
  const src = run.job.fileB;
  const mism = []; let cells = 0;
  let sheets;
  if (hl.method === 'patched') {
    const compared = new Set(run.pairs.map((p) => p.b.sheet));
    const a = unzipSync(src.bytes); const b = unzipSync(outBytes);
    const na = Object.keys(a).sort(); const nb = Object.keys(b).sort();
    if (na.join('|') !== nb.join('|')) mism.push('the list of parts inside the workbook changed');
    const { sheets: parts, styles } = resolveSheetParts(a);
    const allowed = new Set([styles, ...[...compared].map((s) => parts[s])]);
    for (const n of na) {
      if (allowed.has(n) || !b[n]) continue;
      const x = a[n]; const y = b[n];
      if (x.length !== y.length || x.some((v, k) => v !== y[k])) mism.push(`part ${n} changed`);
    }
    sheets = [...compared];
  } else sheets = [...src.sheets];
  const outWb = XLSX.read(outBytes, { type: 'array', cellNF: true, dense: true });
  const outSf = { fmt: 'xlsx', wb: outWb };
  for (const sheet of sheets) {
    const parts = hl.sheetParts[sheet] || [hl.sheetTitles[sheet] || sheet];
    const srcRows = new Map();
    for (const [r, vals] of sheetReader(src, sheet).rows()) {
      const v = trim(src.fmt === 'csv' ? vals : vals);
      if (v.length) srcRows.set(r, v);
    }
    const outRows = new Map();
    parts.forEach((p, k) => {
      const rd = sheetReader(outSf, p);
      for (const [r, vals] of rd.rows()) {
        if (k > 0 && r === 1) continue;
        const real = k === 0 ? r : EXCEL_MAX_ROWS + (k - 1) * (EXCEL_MAX_ROWS - 1) + (r - 1);
        const v = trim([...new Array(rd.firstCol - 1).fill(null), ...vals]);
        if (v.length) outRows.set(real, v);
      }
    });
    // align source rows to absolute columns
    const firstSrc = sheetReader(src, sheet).firstCol || 1;
    for (const [r, v0] of srcRows) {
      const v = trim([...new Array(firstSrc - 1).fill(null), ...v0]);
      cells += v.length;
      const o = outRows.get(r) || [];
      if (v.length !== o.length || v.some((x, k) => x !== o[k])) { mism.push(`${sheet} row ${r}`); if (mism.length > 5) break; }
      outRows.delete(r);
    }
    for (const r of outRows.keys()) { mism.push(`${sheet} row ${r} (extra)`); if (mism.length > 5) break; }
  }
  const ok = !mism.length;
  return [check('File B values are unchanged in the highlighted copy', ok,
    `Compared every value of ${sheets.length} sheet(s) (${fmt(cells)} cells) with the original${ok ? '' : `; differences at ${mism.slice(0, 5).join(', ')}`}`)];
}

export async function checkSourcesUntouched(run) {
  const out = [];
  for (const [label, sf] of [['A', run.job.fileA], ['B', run.job.fileB]]) {
    const same = (await sha256(sf.bytes)) === sf.sha256;
    out.push(check(`File ${label} itself was not modified`, same, same
      ? 'Fingerprint (SHA-256) is identical before and after the run' : 'The file changed during the run'));
    if (run.job.fileA === run.job.fileB) break;
  }
  return out;
}

export function reopenReport(bytes, expected) {
  const wb = XLSX.read(bytes, { type: 'array', dense: true });
  const problems = [];
  for (const [sheets, n] of Object.values(expected)) {
    let got = 0;
    for (const s of sheets) {
      const ws = wb.Sheets[s];
      const rg = ws && ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : null;
      got += rg ? rg.e.r : 0;      // rows after the heading row
    }
    if (got !== n) problems.push(`${sheets[0]}: ${fmt(got)} rows instead of ${fmt(n)}`);
  }
  return check('Report file re-opened and every row is present', !problems.length,
    problems.length ? problems.join('; ') : 'All sheets re-read from the saved report');
}
