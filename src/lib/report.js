// Output 2 (Detailed_Before_After_Comparison.xlsx) and the export pipeline.
import { FastWorkbook, TableSheetWriter, S, safeSheetTitle, concatChunks } from './xlsxWriter.js';
import { buildHighlighted, outputName } from './highlight.js';
import * as V from './validate.js';
import { POLICIES } from './engine.js';
import { summaryLines } from './normalize.js';
import { kindLabel } from './loader.js';
import { EXCEL_MAX_ROWS } from './coords.js';

const DETAIL_HEADERS = ['Sr. No.', 'Match Key', 'A: File', 'A: Sheet', 'A: Row', 'A: Column', 'A: Col Letter', 'A: Cell', 'Before Value',
  'B: File', 'B: Sheet', 'B: Row', 'B: Column', 'B: Col Letter', 'B: Cell', 'After Value', 'Difference Type', 'Normalization Applied', 'Remarks'];
const DETAIL_KINDS = ['num', 'text', 'text', 'text', 'num', 'text', 'text', 'text', 'text', 'text', 'text', 'num', 'text', 'text', 'text', 'text', 'text', 'text', 'text'];
const DETAIL_WIDTHS = [8, 18, 22, 18, 8, 20, 8, 9, 26, 22, 18, 8, 20, 8, 9, 26, 18, 26, 36];
const PAGE = 20000;
const fmt = (n) => Number(n).toLocaleString('en-IN');

async function* batches(db, sql) {
  for (let off = 0; ; off += PAGE) {
    const rows = await db.query(`${sql} LIMIT ${PAGE} OFFSET ${off}`);
    if (!rows.length) return;
    yield* rows;
    if (rows.length < PAGE) return;
  }
}

/** Build both files, run every check, write the report last. Returns { hl, hlBytes, hlName, reportBytes, checks, passed } */
export async function buildOutputs(db, run, { highlightedName = 'File_B_Comparison_Highlighted.xlsx', progress = () => {} } = {}) {
  const { res: hl, bytes: hlBytes } = await buildHighlighted(db, run, (f, m) => progress(0.05 + 0.35 * f, m));
  progress(0.42, 'Checking the results');
  const checks = [];
  checks.push(...await V.checkEngine(db, run));
  progress(0.48, 'Re-reading reported cells from the original files');
  checks.push(...await V.checkCoordinates(db, run));
  progress(0.55, 'Checking the yellow cells');
  checks.push(...await V.checkHighlights(db, run, hl, hlBytes));
  progress(0.6, 'Checking File B values in the copy');
  checks.push(...V.checkBUnchanged(run, hl, hlBytes));
  checks.push(...await V.checkSourcesUntouched(run));
  progress(0.65, 'Writing the before–after report');
  const { bytes: reportBytes, counts } = await writeReport(db, run, hl, checks, (f, m) => progress(0.65 + 0.3 * f, m));
  progress(0.97, 'Confirming the report file');
  checks.push(V.reopenReport(reportBytes, counts));
  progress(1, 'Files ready');
  return { hl, hlBytes, hlName: outputName(highlightedName, run.job.fileB), reportBytes, checks, passed: V.overall(checks) };
}

async function writeReport(db, run, hl, checks, progress) {
  const wb = new FastWorkbook();
  const taken = new Set();
  const counts = {};
  const pairs = Object.fromEntries(run.pairs.map((p) => [p.pairId, p]));
  const aFile = run.job.fileA.name; const bFile = run.job.fileB.name;
  const [{ n: totalDiffs }] = await db.query('SELECT count(*) AS n FROM diffs');

  let w = new TableSheetWriter(wb, 'Detailed Differences', DETAIL_HEADERS, DETAIL_WIDTHS, taken, DETAIL_KINDS);
  let k = 0;
  for await (const r of batches(db, `SELECT sr, pair_id, key_disp, a_row, a_col_name, a_letter, b_row, b_col, b_col_name, b_letter,
      before_value, after_value, diff_type, rules, remarks FROM diffs ORDER BY sr`)) {
    const p = pairs[r.pair_id];
    let remarks = r.remarks;
    const [ls, lc] = hl.locate(p.b.sheet, r.b_row, r.b_col);
    if (ls !== (hl.sheetTitles[p.b.sheet] || p.b.sheet) || lc !== `${r.b_letter}${r.b_row}`) remarks = (remarks ? `${remarks}; ` : '') + `In the highlighted copy: ${ls}!${lc}`;
    w.append([r.sr, r.key_disp, aFile, p.a.sheet, r.a_row, r.a_col_name, r.a_letter, `${r.a_letter}${r.a_row}`, r.before_value,
      bFile, p.b.sheet, r.b_row, r.b_col_name, r.b_letter, `${r.b_letter}${r.b_row}`, r.after_value, r.diff_type, r.rules, remarks]);
    if (++k % 20000 === 0) progress(0.6 * k / totalDiffs, `Writing changed cells: ${fmt(k)} of ${fmt(totalDiffs)}`);
  }
  w.close(); counts.diffs = [w.sheets, w.rowsWritten];

  progress(0.62, 'Writing changed records');
  w = new TableSheetWriter(wb, 'Row-Wise', ['Sr. No.', 'Match Key', 'A: Sheet', 'A: Row', 'B: Sheet', 'B: Row', 'Changed Columns (count)',
    'Changed Columns', 'Description'], [8, 18, 18, 8, 18, 8, 12, 34, 60], taken, ['num', 'text', 'text', 'num', 'text', 'num', 'num', 'text', 'text']);
  k = 0;
  for await (const r of batches(db, 'SELECT pair_id, key_disp, a_row, b_row, n_changed, changed_columns, description FROM rowwise ORDER BY pair_id, b_row')) {
    const p = pairs[r.pair_id];
    w.append([++k, r.key_disp, p.a.sheet, r.a_row, p.b.sheet, r.b_row, r.n_changed, r.changed_columns, r.description]);
  }
  w.close(); counts.rowwise = [w.sheets, w.rowsWritten];

  for (const [side, label] of [['a', 'Only in A'], ['b', 'Only in B']]) {
    progress(side === 'a' ? 0.7 : 0.78, `Writing ${label.toLowerCase()}`);
    for (const p of run.pairs) {
      const st = side === 'a' ? p.a : p.b;
      const title = run.pairs.length === 1 ? label : `${label} (P${p.pairId})`;
      w = new TableSheetWriter(wb, title, ['Match Key', 'Sheet', 'Original Row', 'Status', ...st.columns.map((c) => c.name)],
        [18, 18, 10, 26, ...st.columns.map(() => 16)], taken, ['text', 'text', 'num', 'text', ...st.columns.map(() => 'text')]);
      const vals = st.columns.map((c) => c.internal).join(', ');
      for await (const r of batches(db, `SELECT key_disp, src_row, status, ${vals} FROM p${p.pairId}_o${side} ORDER BY src_row`)) {
        w.append([r.key_disp, st.sheet, r.src_row, r.status, ...st.columns.map((c) => r[c.internal])]);
      }
      w.close(); counts[`only_${side}_${p.pairId}`] = [w.sheets, w.rowsWritten];
    }
  }

  progress(0.85, 'Writing duplicate keys');
  w = new TableSheetWriter(wb, 'Duplicate Keys', ['File', 'Sheet', 'Match Key', 'Row', 'Occurrence', 'Occurrences', 'Strategy', 'Compared or Excluded'],
    [22, 18, 18, 8, 10, 11, 18, 44], taken, ['text', 'text', 'text', 'num', 'num', 'num', 'text', 'text']);
  for await (const r of batches(db, 'SELECT file_name, sheet, key_disp, src_row, occurrence, occurrences, strategy, outcome FROM dups ORDER BY pair_id, file_side, key_disp, src_row')) {
    w.append([r.file_name, r.sheet, r.key_disp, r.src_row, r.occurrence, r.occurrences, r.strategy, r.outcome]);
  }
  w.close(); counts.dups = [w.sheets, w.rowsWritten];

  const expected = { diffs: totalDiffs, rowwise: (await db.query('SELECT count(*) AS n FROM rowwise'))[0].n,
    dups: (await db.query('SELECT count(*) AS n FROM dups'))[0].n };
  for (const p of run.pairs) for (const s of ['a', 'b']) expected[`only_${s}_${p.pairId}`] = (await db.query(`SELECT count(*) AS n FROM p${p.pairId}_o${s}`))[0].n;
  const short = Object.entries(expected).filter(([key, v]) => counts[key][1] !== v).map(([key, v]) => `${key}: ${fmt(counts[key][1])} of ${fmt(v)}`);
  checks.push(V.check('Report holds every row (nothing cut off)', !short.length, short.length ? short.join('; ') : 'Rows written match the comparison exactly'));
  const split = Object.values(counts).flatMap(([sheets]) => sheets.slice(1));
  if (split.length) checks.push(V.check('Large tables were continued on extra sheets', true, `More rows than one Excel sheet holds; continued on: ${split.join(', ')}`, false));

  progress(0.92, 'Writing the summary');
  writeSummary(wb, taken, run, hl, checks);
  return { bytes: concatChunks(wb.finish()), counts };
}

function writeSummary(wb, taken, run, hl, checks) {
  const ws = wb.addSheet(safeSheetTitle('Summary', taken), { widths: [46, 70, 12], freeze: false });
  const t = run.totals; const job = run.job;
  const item = (label, value, style = S.BODY) => ws.append(typeof value === 'number'
    ? [[label, S.BODY], [value, S.BODY_INT_LEFT]] : [[label, S.BODY], [String(value), style]]);
  const section = (text) => { ws.append([]); ws.append([[text, S.SECTION]]); };
  ws.append([['Before–After Comparison Summary', S.TITLE]]);
  const passed = V.overall(checks);
  item('Validation result', passed ? 'PASS – all critical checks passed' : 'FAIL – see Validation below; do not rely on these results', passed ? S.PASS : S.FAIL);
  section('Files');
  for (const [lab, sf] of [['File A (before)', job.fileA], ['File B (after)', job.fileB]]) {
    const enc = sf.fmt === 'csv' ? `, ${sf.encoding}, delimiter ${JSON.stringify(sf.delimiter)}` : '';
    item(lab, `${sf.name} — ${kindLabel(sf.fmt)}${enc}, ${fmt(sf.size)} bytes`);
    item(`${lab} SHA-256`, sf.sha256);
  }
  section('Sheet pairs');
  for (const p of run.pairs) {
    item(`Pair ${p.pairId}`, `A: ${p.a.sheet} (headings on row ${p.a.headerRow})  ↔  B: ${p.b.sheet} (headings on row ${p.b.headerRow})`);
    item('   Key column(s)', p.keyCols.map((k) => `${k.a_name} ↔ ${k.b_name}`).join(', '));
    item('   Columns compared', `${p.compareCols.length}: ${p.compareCols.map((c) => (c.a_name === c.b_name ? c.a_name : `${c.a_name} ↔ ${c.b_name}`)).join(', ')}`);
    if (job.schemaCompare) { item('   Columns only in A', p.schemaOnlyA.join(', ') || 'None'); item('   Columns only in B', p.schemaOnlyB.join(', ') || 'None'); }
  }
  section('Row counts');
  const rows = [['Rows in File A (below headings)', t.rows_a], ['Rows in File B (below headings)', t.rows_b], ['Matched records', t.matched],
    ['Records with differences', t.changed_records], ['Identical matched records', t.identical], ['Only in A', t.only_a], ['Only in B', t.only_b],
    ['Changed cells', t.changed_cells], ['   Value Changed', t.value_changed], ['   Blank-to-Value', t.blank_to_value],
    ['   Value-to-Blank', t.value_to_blank], ['   Formula Difference', t.formula_diff],
    ['Duplicate keys in A (keys / rows)', `${fmt(t.dup_keys_a)} / ${fmt(t.dup_rows_a)}`], ['Duplicate keys in B (keys / rows)', `${fmt(t.dup_keys_b)} / ${fmt(t.dup_rows_b)}`],
    ['Duplicate rows excluded (A / B)', `${fmt(t.dup_excluded_a)} / ${fmt(t.dup_excluded_b)}`],
    ['Blank keys – not compared (A / B)', `${fmt(t.blank_key_a)} / ${fmt(t.blank_key_b)}`],
    ['Empty rows skipped (A / B)', `${fmt(t.empty_a)} / ${fmt(t.empty_b)}`], ['Cells highlighted yellow in the copy', hl.yellowCells]];
  if (job.highlightOnlyB) rows.push(['Cells filled pale orange (rows only in B)', hl.orangeCells]);
  for (const [l, v] of rows) item(l, v);
  section('Options');
  for (const line of summaryLines(job.rules)) item('Comparison rule', line);
  const kr = job.keyRules;
  item('Key matching', [['trim spaces', kr.trim], ['ignore case', kr.ignore_case], ['ignore leading zeros', kr.leading_zeros]].filter(([, on]) => on).map(([x]) => x).join(', ') || 'exact');
  item('Duplicate keys', POLICIES[job.dupPolicy]);
  item('Compare formulas', job.compareFormulas ? 'Yes' : 'No (values only)');
  item('Schema comparison', job.schemaCompare ? 'Yes' : 'No');
  item('Rows only in B filled orange', job.highlightOnlyB ? 'Yes' : 'No');
  item('Where it ran', 'In the web browser on this computer (files were not uploaded)');
  section('Timing');
  const d = (x) => x.toLocaleString('en-IN', { hour12: false });
  item('Started', d(run.started)); item('Finished', d(run.finished));
  item('Comparison duration', `${run.durationS.toFixed(1)} seconds`); item('Report written', d(new Date()));
  section('Validation');
  for (const c of checks) ws.append([[c.name, S.BODY], [c.detail, S.BODY], [c.status, c.passed ? S.PASS : S.FAIL]]);
  const notes = [...new Set([...run.warnings, ...hl.notes, ...job.fileA.warnings, ...job.fileB.warnings])];
  if (notes.length) { section('Notes'); for (const n of notes) item('Note', n); }
  if (Object.values(hl.sheetParts).some((s) => s.length > 1)) item('Note', `Excel allows ${fmt(EXCEL_MAX_ROWS)} rows per sheet; see the remarks column for cells that moved to a continuation sheet.`);
}
