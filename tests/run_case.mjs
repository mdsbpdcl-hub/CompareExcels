// Test harness: run the browser engine in Node on real files and write both outputs.
// usage: node tests/run_case.mjs config.json outdir
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { wrapBlocking } from '../src/lib/db.js';
import { openSource, columnNames } from '../src/lib/loader.js';
import { suggestPairs } from '../src/lib/mapper.js';
import { runComparison } from '../src/lib/engine.js';
import { buildOutputs } from '../src/lib/report.js';
import { defaultRules, defaultKeyRules } from '../src/lib/normalize.js';

const require = createRequire(import.meta.url);
const duck = require('@duckdb/duckdb-wasm/dist/duckdb-node-blocking.cjs');
const d = path.join(path.dirname(new URL(import.meta.url).pathname), '..', 'node_modules', '@duckdb', 'duckdb-wasm', 'dist') + '/';

const [cfgPath, outDir] = process.argv.slice(2);
const cfg = JSON.parse(fs.readFileSync(cfgPath, 'utf8'));
const db0 = await duck.createDuckDB({ mvp: { mainModule: d + 'duckdb-mvp.wasm', mainWorker: d + 'duckdb-node-mvp.worker.cjs' },
  eh: { mainModule: d + 'duckdb-eh.wasm', mainWorker: d + 'duckdb-node-eh.worker.cjs' } }, new duck.VoidLogger(), duck.NODE_RUNTIME);
await db0.instantiate();
const db = wrapBlocking(db0.connect());

const load = async (p, opts = {}) => openSource(path.basename(p), new Uint8Array(fs.readFileSync(p)), opts);
const t0 = performance.now();
const fa = await load(cfg.a, cfg.a_opts);
const fb = cfg.b === cfg.a ? fa : await load(cfg.b, cfg.b_opts);
let pairs = cfg.pairs;
if (!pairs) {
  const aSheet = cfg.a_sheet || fa.sheets[0]; const bSheet = cfg.b_sheet || fb.sheets[0];
  const ah = cfg.a_header || 1; const bh = cfg.b_header || 1;
  const columns = cfg.columns ? cfg.columns.map(([a, b]) => ({ a, b, compare: true }))
    : suggestPairs(columnNames(fa, aSheet, ah), columnNames(fb, bSheet, bh));
  pairs = [{ a_sheet: aSheet, b_sheet: bSheet, a_header: ah, b_header: bh, keys: cfg.keys.map(([a, b]) => ({ a, b })), columns }];
}
const job = { fileA: fa, fileB: fb, pairs, rules: { ...defaultRules(), ...(cfg.rules || {}) },
  keyRules: { ...defaultKeyRules(), ...(cfg.key_rules || {}) }, dupPolicy: cfg.dup_policy || 'keep_first',
  compareFormulas: !!cfg.compare_formulas, schemaCompare: !!cfg.schema_compare, highlightOnlyB: !!cfg.highlight_only_b };
const out = { ok: true };
try {
  const run = await runComparison(db, job);
  const exp = await buildOutputs(db, run);
  fs.mkdirSync(outDir, { recursive: true });
  const hlPath = path.join(outDir, exp.hlName);
  fs.writeFileSync(hlPath, exp.hlBytes);
  const repPath = path.join(outDir, 'Detailed_Before_After_Comparison.xlsx');
  fs.writeFileSync(repPath, exp.reportBytes);
  Object.assign(out, { totals: run.totals, checks: exp.checks, passed: exp.passed, highlighted: hlPath, report: repPath,
    notes: exp.hl.notes, method: exp.hl.method, seconds: (performance.now() - t0) / 1000,
    pairs: run.pairs.map((p) => ({ schema_only_a: p.schemaOnlyA, schema_only_b: p.schemaOnlyB, a_empty: p.a.emptyRows, b_empty: p.b.emptyRows })) });
} catch (e) {
  Object.assign(out, { ok: false, error: String(e && e.message ? e.message : e), stack: e && e.stack });
}
fs.mkdirSync(outDir, { recursive: true });
fs.writeFileSync(path.join(outDir, 'result.json'), JSON.stringify(out));
