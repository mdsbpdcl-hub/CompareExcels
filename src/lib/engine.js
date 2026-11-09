// The comparison engine. All heavy work is DuckDB SQL (set-based, no row-by-row loops).
import { stageSheet, stageFormulas } from './loader.js';
import { valueSql, keySql, appliedFor, lit } from './normalize.js';

export const POLICIES = { keep_first: 'Keep first', keep_last: 'Keep last', pair: 'Pair by occurrence',
  exclude: 'Exclude and report' };
const BLANK = 'Blank key – not compared';

// ----------------------------------------------------------------- keys & duplicates
function keyExpr(ints, kr) {
  const parts = ints.map((c) => keySql(c, kr));
  return `(CASE WHEN ${parts.map((p) => `(${p}) IS NULL`).join(' OR ')} THEN NULL ELSE concat_ws(chr(31), ${parts.join(', ')}) END)`;
}
const dispExpr = (ints) => (ints.length === 1 ? `COALESCE(${ints[0]}, '')`
  : `concat_ws(' | ', ${ints.map((c) => `COALESCE(${c}, '')`).join(', ')})`);

export async function prepareSide(db, staged, out, keyInts, kr, policy) {
  const status = {
    keep_first: "CASE WHEN __key IS NULL THEN 'blank_key' WHEN __occ = 1 THEN 'ok' ELSE 'dup_excluded' END",
    keep_last: "CASE WHEN __key IS NULL THEN 'blank_key' WHEN __occ_desc = 1 THEN 'ok' ELSE 'dup_excluded' END",
    pair: "CASE WHEN __key IS NULL THEN 'blank_key' ELSE 'ok' END",
    exclude: "CASE WHEN __key IS NULL THEN 'blank_key' WHEN __cnt > 1 THEN 'dup_excluded' ELSE 'ok' END",
  }[policy];
  if (!status) throw new Error(`Unknown duplicate policy ${policy}`);
  const mocc = policy === 'pair' ? '__occ' : '1';
  await db.exec(`DROP TABLE IF EXISTS ${out}`);
  await db.exec(`CREATE TABLE ${out} AS
    WITH k AS (SELECT *, ${keyExpr(keyInts, kr)} AS __key, ${dispExpr(keyInts)} AS __keydisp FROM ${staged}),
    w AS (SELECT *,
      CASE WHEN __key IS NULL THEN NULL ELSE count(*) OVER (PARTITION BY __key) END AS __cnt,
      CASE WHEN __key IS NULL THEN NULL ELSE row_number() OVER (PARTITION BY __key ORDER BY __row) END AS __occ,
      CASE WHEN __key IS NULL THEN NULL ELSE row_number() OVER (PARTITION BY __key ORDER BY __row DESC) END AS __occ_desc
      FROM k)
    SELECT *, ${status} AS __status, CASE WHEN __key IS NULL THEN NULL ELSE ${mocc} END AS __mocc FROM w ORDER BY __row`);
  const [s] = await db.query(`SELECT count(*) AS n, count(*) FILTER (WHERE __status = 'ok') AS ok,
      count(*) FILTER (WHERE __status = 'blank_key') AS blank, count(*) FILTER (WHERE __status = 'dup_excluded') AS dup_ex,
      count(*) FILTER (WHERE __cnt > 1) AS dup_rows, count(DISTINCT __key) FILTER (WHERE __cnt > 1) AS dup_keys FROM ${out}`);
  const [c] = await db.query(`SELECT count(*) AS n FROM (SELECT __key, __mocc FROM ${out} WHERE __status = 'ok'
      GROUP BY 1, 2 HAVING count(*) > 1)`);
  if (c.n) throw new Error('Internal check failed: duplicate keys would create a many-to-many join.');
  return { rows: s.n, ok: s.ok, blank_key: s.blank, dup_excluded: s.dup_ex, dup_rows: s.dup_rows, dup_keys: s.dup_keys };
}

// ----------------------------------------------------------------- engine
function resolve(a, b, pairs) {
  return pairs.map((p, i) => {
    const ca = a.columns.find((c) => c.name === p.a);
    if (!ca) throw new Error(`Column '${p.a}' was not found in ${a.sheet} (File A, headings on row ${a.headerRow}). Check the heading row.`);
    const cb = b.columns.find((c) => c.name === p.b);
    if (!cb) throw new Error(`Column '${p.b}' was not found in ${b.sheet} (File B, headings on row ${b.headerRow}). Check the heading row.`);
    return { seq: i + 1, a_name: ca.name, a_header: ca.header, a_int: ca.internal, a_col: ca.col, a_letter: ca.letter,
      b_name: cb.name, b_header: cb.header, b_int: cb.internal, b_col: cb.col, b_letter: cb.letter };
  });
}

/**
 * job: { fileA, fileB, pairs: [{a_sheet, b_sheet, a_header, b_header, keys:[{a,b}], columns:[{a,b,compare}]}],
 *        rules, keyRules, dupPolicy, compareFormulas, schemaCompare, highlightOnlyB }
 */
export async function runComparison(db, job, progress = () => {}) {
  const started = new Date();
  const t0 = performance.now();
  const stageTimes = {};
  const warnings = [];
  if (!job.pairs.length) throw new Error('Choose at least one sheet pair.');
  const bSheets = job.pairs.map((p) => p.b_sheet);
  if (new Set(bSheets).size !== bSheets.length) throw new Error('Each File B sheet can be used in only one sheet pair.');
  await db.exec('DROP TABLE IF EXISTS diffs_tmp');
  await db.exec(`CREATE TABLE diffs_tmp (pair_id INTEGER, key_disp VARCHAR, a_row BIGINT, b_row BIGINT, col_seq INTEGER,
    a_col_name VARCHAR, a_col INTEGER, a_letter VARCHAR, b_col_name VARCHAR, b_col INTEGER, b_letter VARCHAR,
    before_value VARCHAR, after_value VARCHAR, diff_type VARCHAR, rules VARCHAR, remarks VARCHAR)`);
  await db.exec('DROP TABLE IF EXISTS dups');
  await db.exec(`CREATE TABLE dups (pair_id INTEGER, file_side VARCHAR, file_name VARCHAR, sheet VARCHAR, key_disp VARCHAR,
    src_row BIGINT, occurrence INTEGER, occurrences INTEGER, strategy VARCHAR, outcome VARCHAR)`);
  const results = [];
  const n = job.pairs.length;
  const r = job.rules;

  for (let i = 1; i <= n; i++) {
    const pc = job.pairs[i - 1];
    const base = (i - 1) / n; const span = 1 / n;
    const sub = (lo, hi) => (f, m) => progress(base + span * (lo + (hi - lo) * f), m);
    if (!pc.keys.length) throw new Error(`Sheet pair ${i}: choose at least one key column.`);
    const p = `p${i}`;
    let ts = performance.now();
    const a = await stageSheet(db, job.fileA, pc.a_sheet, pc.a_header, `${p}_a_stg`, sub(0, 0.25));
    const b = await stageSheet(db, job.fileB, pc.b_sheet, pc.b_header, `${p}_b_stg`, sub(0.25, 0.5));
    stageTimes[`Pair ${i}: read files`] = performance.now() - ts;

    const keyCols = resolve(a, b, pc.keys);
    const keyA = new Set(keyCols.map((k) => k.a_name));
    let cmp = resolve(a, b, pc.columns.filter((c) => c.compare)).filter((c) => !keyA.has(c.a_name));
    cmp.forEach((c, j) => { c.seq = j + 1; });
    const bUsed = cmp.map((c) => c.b_name);
    if (new Set(bUsed).size !== bUsed.length) throw new Error(`Sheet pair ${i}: a File B column is paired with more than one File A column.`);
    const clash = bUsed.filter((x) => keyCols.some((k) => k.b_name === x));
    if (clash.length) throw new Error(`Sheet pair ${i}: ${clash.join(', ')} is a key column and cannot also be compared.`);
    const pr = { pairId: i, config: pc, a, b, compareCols: cmp, keyCols, stats: {}, schemaOnlyA: [], schemaOnlyB: [],
      label: `${a.sheet} ↔ ${b.sheet}` };

    ts = performance.now();
    progress(base + span * 0.5, `Pair ${i}: building keys`);
    pr.aDup = await prepareSide(db, a.table, `${p}_a`, keyCols.map((k) => k.a_int), job.keyRules, job.dupPolicy);
    pr.bDup = await prepareSide(db, b.table, `${p}_b`, keyCols.map((k) => k.b_int), job.keyRules, job.dupPolicy);
    stageTimes[`Pair ${i}: keys`] = performance.now() - ts;

    ts = performance.now();
    progress(base + span * 0.55, `Pair ${i}: matching records`);
    const extra = cmp.length ? ', ' + cmp.map((c) => `a.${c.a_int} AS a_${c.a_int}`).concat(cmp.map((c) => `b.${c.b_int} AS b_${c.b_int}`)).join(', ') : '';
    await db.exec(`DROP TABLE IF EXISTS ${p}_m`);
    await db.exec(`CREATE TABLE ${p}_m AS SELECT a.__key AS mkey, a.__mocc AS mocc, a.__keydisp AS key_disp,
        a.__row AS a_row, b.__row AS b_row${extra}
      FROM (SELECT * FROM ${p}_a WHERE __status = 'ok') a JOIN (SELECT * FROM ${p}_b WHERE __status = 'ok') b
        ON a.__key = b.__key AND a.__mocc = b.__mocc ORDER BY a.__row`);
    const [mm] = await db.query(`SELECT count(*) AS n, count(DISTINCT a_row) AS da, count(DISTINCT b_row) AS db FROM ${p}_m`);
    if (!(mm.n === mm.da && mm.da === mm.db)) throw new Error('Internal check failed: a record was matched more than once.');
    stageTimes[`Pair ${i}: matching`] = performance.now() - ts;

    ts = performance.now();
    const fx = (f) => f === 'xlsx' || f === 'xlsm';
    const useF = job.compareFormulas && fx(job.fileA.fmt) && fx(job.fileB.fmt);
    if (useF) {
      progress(base + span * 0.58, `Pair ${i}: reading formulas`);
      await stageFormulas(db, job.fileA, a, `${p}_fa`);
      await stageFormulas(db, job.fileB, b, `${p}_fb`);
    } else if (job.compareFormulas && i === 1) {
      warnings.push('Formula comparison needs both files to be .xlsx/.xlsm, so only values were compared.');
    }
    for (let j = 0; j < cmp.length; j++) {
      const c = cmp[j];
      progress(base + span * (0.6 + 0.3 * j / Math.max(cmp.length, 1)), `Pair ${i}: comparing column ${c.a_name}`);
      const isNum = r.numeric_columns.includes(c.a_name);
      const isDate = r.date_columns.includes(c.a_name);
      const na = valueSql('va', r, isNum, isDate);
      const nb = valueSql('vb', r, isNum, isDate);
      const rulesTxt = lit(appliedFor(r, c.a_name).join(', ') || 'Exact match');
      const lits = `${i}, key_disp, a_row, b_row, ${c.seq}, ${lit(c.a_name)}, ${c.a_col}, ${lit(c.a_letter)}, ${lit(c.b_name)}, `
        + `${c.b_col}, ${lit(c.b_letter)}, va, vb`;
      await db.exec(`INSERT INTO diffs_tmp SELECT ${lits},
          CASE WHEN va IS NULL THEN 'Blank-to-Value' WHEN vb IS NULL THEN 'Value-to-Blank' ELSE 'Value Changed' END,
          ${rulesTxt},
          CASE WHEN na IS DISTINCT FROM va OR nb IS DISTINCT FROM vb
               THEN 'Compared as: ' || COALESCE(na, '(blank)') || ' → ' || COALESCE(nb, '(blank)') ELSE '' END
        FROM (SELECT key_disp, a_row, b_row, a_${c.a_int} AS va, b_${c.b_int} AS vb FROM ${p}_m) s0,
             LATERAL (SELECT ${na} AS na, ${nb} AS nb) s1
        WHERE na IS DISTINCT FROM nb`);
      if (useF) {
        await db.exec(`INSERT INTO diffs_tmp SELECT ${lits}, 'Formula Difference', ${rulesTxt},
            'Formula A: ' || COALESCE(fa, '(no formula)') || ' | Formula B: ' || COALESCE(fb, '(no formula)')
          FROM (SELECT m.key_disp, m.a_row, m.b_row, m.a_${c.a_int} AS va, m.b_${c.b_int} AS vb, fa.formula AS fa, fb.formula AS fb
                FROM ${p}_m m
                LEFT JOIN ${p}_fa fa ON fa.__row = m.a_row AND fa.col = ${c.a_col}
                LEFT JOIN ${p}_fb fb ON fb.__row = m.b_row AND fb.col = ${c.b_col}
                WHERE fa.formula IS NOT NULL OR fb.formula IS NOT NULL) s0,
               LATERAL (SELECT ${na} AS na, ${nb} AS nb) s1
          WHERE na IS NOT DISTINCT FROM nb AND COALESCE(replace(fa, ' ', ''), '') <> COALESCE(replace(fb, ' ', ''), '')`);
      }
    }
    stageTimes[`Pair ${i}: cell comparison`] = performance.now() - ts;

    ts = performance.now();
    progress(base + span * 0.92, `Pair ${i}: listing unmatched rows`);
    for (const [side, other, staged, fname] of [['a', 'b', a, job.fileA.name], ['b', 'a', b, job.fileB.name]]) {
      const tag = side.toUpperCase();
      const vals = staged.columns.map((c) => c.internal).join(', ');
      await db.exec(`DROP TABLE IF EXISTS ${p}_o${side}`);
      await db.exec(`CREATE TABLE ${p}_o${side} AS SELECT __keydisp AS key_disp, __row AS src_row,
          CASE WHEN __status = 'blank_key' THEN '${BLANK}'
               WHEN ${job.dupPolicy === 'pair' ? 'TRUE' : 'FALSE'} AND __cnt > 1 THEN 'Only in ${tag} (unpaired duplicate)'
               ELSE 'Only in ${tag}' END AS status, ${vals}
        FROM ${p}_${side} s WHERE __status = 'blank_key'
          OR (__status = 'ok' AND NOT EXISTS (SELECT 1 FROM ${p}_m m WHERE m.${side}_row = s.__row))
        ORDER BY __row`);
      await db.exec(`INSERT INTO dups SELECT ${i}, '${tag}', ${lit(fname)}, ${lit(staged.sheet)}, __keydisp, __row, __occ, __cnt,
          ${lit(POLICIES[job.dupPolicy])},
          CASE WHEN __status = 'dup_excluded' THEN 'Excluded – not compared'
               WHEN EXISTS (SELECT 1 FROM ${p}_m m WHERE m.${side}_row = s.__row) THEN 'Compared'
               ELSE 'Compared – no partner in File ${other.toUpperCase()} (listed as Only in ${tag})' END
        FROM ${p}_${side} s WHERE __cnt > 1 ORDER BY __key, __row`);
    }
    stageTimes[`Pair ${i}: unmatched rows`] = performance.now() - ts;
    const [oa] = await db.query(`SELECT count(*) AS n FROM ${p}_oa WHERE status <> '${BLANK}'`);
    const [ob] = await db.query(`SELECT count(*) AS n FROM ${p}_ob WHERE status <> '${BLANK}'`);
    pr.stats = { matched: mm.n, only_a: oa.n, only_b: ob.n };
    if (job.schemaCompare) {
      const mappedA = new Set([...pc.columns.map((c) => c.a), ...keyA]);
      const mappedB = new Set([...pc.columns.map((c) => c.b), ...keyCols.map((k) => k.b_name)]);
      pr.schemaOnlyA = a.columns.map((c) => c.name).filter((x) => !mappedA.has(x));
      pr.schemaOnlyB = b.columns.map((c) => c.name).filter((x) => !mappedB.has(x));
    }
    results.push(pr);
  }

  progress(0.97, 'Numbering differences');
  await db.exec('DROP TABLE IF EXISTS diffs');
  await db.exec(`CREATE TABLE diffs AS SELECT row_number() OVER (ORDER BY pair_id, b_row, col_seq, diff_type) AS sr, *
    FROM diffs_tmp ORDER BY sr`);
  await db.exec('DROP TABLE diffs_tmp');
  await db.exec('DROP TABLE IF EXISTS rowwise');
  await db.exec(`CREATE TABLE rowwise AS SELECT pair_id, key_disp, a_row, b_row, count(*) AS n_changed,
      string_agg(a_col_name, ', ' ORDER BY col_seq) AS changed_columns,
      array_to_string(list(a_col_name || ': ' || COALESCE(left(before_value, 60), '(blank)') || ' → '
        || COALESCE(left(after_value, 60), '(blank)') ORDER BY col_seq)[1:5], '; ')
      || CASE WHEN count(*) > 5 THEN ' … and ' || (count(*) - 5) || ' more' ELSE '' END AS description
    FROM diffs GROUP BY pair_id, key_disp, a_row, b_row ORDER BY pair_id, b_row`);

  for (const pr of results) {
    const [d] = await db.query(`SELECT count(*) AS n, count(DISTINCT (a_row, b_row)) AS recs,
        count(*) FILTER (WHERE diff_type = 'Value Changed') AS vc, count(*) FILTER (WHERE diff_type = 'Blank-to-Value') AS bv,
        count(*) FILTER (WHERE diff_type = 'Value-to-Blank') AS vb, count(*) FILTER (WHERE diff_type = 'Formula Difference') AS fd
      FROM diffs WHERE pair_id = ${pr.pairId}`);
    Object.assign(pr.stats, { changed_cells: d.n, changed_records: d.recs, value_changed: d.vc, blank_to_value: d.bv,
      value_to_blank: d.vb, formula_diff: d.fd, identical: pr.stats.matched - d.recs,
      blank_key_a: pr.aDup.blank_key, blank_key_b: pr.bDup.blank_key, dup_excluded_a: pr.aDup.dup_excluded,
      dup_excluded_b: pr.bDup.dup_excluded, dup_keys_a: pr.aDup.dup_keys, dup_keys_b: pr.bDup.dup_keys,
      dup_rows_a: pr.aDup.dup_rows, dup_rows_b: pr.bDup.dup_rows });
  }
  const keys = Object.keys(results[0].stats);
  const totals = Object.fromEntries(keys.map((k) => [k, results.reduce((s, p) => s + p.stats[k], 0)]));
  totals.rows_a = results.reduce((s, p) => s + p.a.sourceRows, 0);
  totals.rows_b = results.reduce((s, p) => s + p.b.sourceRows, 0);
  totals.empty_a = results.reduce((s, p) => s + p.a.emptyRows, 0);
  totals.empty_b = results.reduce((s, p) => s + p.b.emptyRows, 0);
  stageTimes.Total = performance.now() - t0;
  progress(1, 'Comparison finished');
  return { job, pairs: results, started, finished: new Date(), stageTimes, totals, warnings,
    durationS: (performance.now() - t0) / 1000 };
}
