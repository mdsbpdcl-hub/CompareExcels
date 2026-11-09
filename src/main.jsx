import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import './app.css';
import { Header, Stepper, STEPS, PageTitle, Note, Tag, Card, More, Checkbox, Chips, humanSize, fmt } from './ui/common.jsx';
import { Results } from './ui/results.jsx';
import { openSource, readTopRows, detectHeaderRow, scanSheet, previewRows, kindLabel } from './lib/loader.js';
import { DELIMITERS, ENCODINGS } from './lib/csv.js';
import { suggestPairs, suggestKey } from './lib/mapper.js';
import { defaultRules, defaultKeyRules, isExact, summaryLines } from './lib/normalize.js';
import { runComparison, POLICIES } from './lib/engine.js';
import { buildOutputs } from './lib/report.js';
import { openBrowserDb } from './lib/db.js';
import ehWasm from '@duckdb/duckdb-wasm/dist/duckdb-eh.wasm?url';
import ehWorker from '@duckdb/duckdb-wasm/dist/duckdb-browser-eh.worker.js?url';
import mvpWasm from '@duckdb/duckdb-wasm/dist/duckdb-mvp.wasm?url';
import mvpWorker from '@duckdb/duckdb-wasm/dist/duckdb-browser-mvp.worker.js?url';

const NOT_MAPPED = '— not mapped —';
const TPL_KEY = 'xlcompare.templates.v1';
const DEFAULT_HL = 'File_B_Comparison_Highlighted.xlsx';
const DEFAULT_REP = 'Detailed_Before_After_Comparison.xlsx';
const tick = () => new Promise((r) => setTimeout(r, 0));

// ---------------------------------------------------------------- column info cache (one full scan per sheet/heading)
const infoCache = new Map();
function columnInfo(sf, sheet, header) {
  const key = `${sf.sha256}|${sf.encoding}|${sf.delimiter}|${sheet}|${header}`;
  if (!infoCache.has(key)) {
    const s = scanSheet(sf, sheet, header);
    const names = s.columns.map((c) => c.name);
    const profile = {};
    const sample = s.data.slice(0, 600);
    const numRe = /^-?[\d,]*\.?\d+$/;
    const dateRe = /^(\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{1,2}[-/. ](\d{1,2}|[A-Za-z]{3,9})[-/. ]\d{4})/;
    const idLike = /(^|\b|_)(id|no|number|code|ca|account|acct|mobile|phone|contact|pin|pincode|ref|serial)(\b|_|$)/i;
    names.forEach((n, j) => {
      const vals = sample.map((r) => r[j]).filter((v) => v !== null && v !== undefined && v !== '');
      const k = vals.length || 1;
      profile[n] = { unique: vals.length ? new Set(vals).size / k : 0,
        numeric: !idLike.test(n) && vals.length > 0 && vals.filter((v) => numRe.test(v.trim()) && !(v.length > 1 && v.startsWith('0') && !v.includes('.'))).length / k >= 0.9,
        date: vals.length > 0 && vals.filter((v) => dateRe.test(v.trim())).length / k >= 0.9 };
    });
    infoCache.set(key, { names, profile });
  }
  return infoCache.get(key);
}
const autoHeader = (sf, sheet) => detectHeaderRow(readTopRows(sf, sheet, 80).rows);

// ---------------------------------------------------------------- templates (kept in this browser; can be exported as JSON)
const loadTemplates = () => { try { return JSON.parse(localStorage.getItem(TPL_KEY) || '{}'); } catch { return {}; } };
const saveTemplates = (t) => { try { localStorage.setItem(TPL_KEY, JSON.stringify(t)); } catch { /* storage may be blocked */ } };
function download(name, bytes, mime) {
  const url = URL.createObjectURL(new Blob([bytes], { type: mime }));
  const a = document.createElement('a');
  a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}

function App() {
  const [step, setStep] = useState(0);
  const [doneUpto, setDoneUpto] = useState(-1);
  const [files, setFiles] = useState({ A: null, B: null });
  const [fileStatus, setFileStatus] = useState({ A: null, B: null });
  const [sameWb, setSameWb] = useState(false);
  const [showSettings, setShowSettings] = useState({ A: false, B: false });
  const [pairs, setPairs] = useState([]);
  const [maps, setMaps] = useState({});
  const [rules, setRules] = useState(defaultRules());
  const [keyRules, setKeyRules] = useState(defaultKeyRules());
  const [dupPolicy, setDupPolicy] = useState('keep_first');
  const [opts, setOpts] = useState({ compareFormulas: false, schemaCompare: false, highlightOnlyB: false });
  const [run, setRun] = useState(null);
  const [exp, setExp] = useState(null);
  const [error, setError] = useState(null);
  const [prog, setProg] = useState(null);
  const [names, setNames] = useState({ hl: DEFAULT_HL, rep: DEFAULT_REP });
  const [templates, setTemplates] = useState(loadTemplates());
  const [pendingTpl, setPendingTpl] = useState(null);
  const [dbState, setDbState] = useState({ ready: false, error: null });
  const dbRef = useRef(null);

  useEffect(() => {
    openBrowserDb({ mvp: { mainModule: mvpWasm, mainWorker: mvpWorker }, eh: { mainModule: ehWasm, mainWorker: ehWorker } })
      .then((x) => { dbRef.current = x; setDbState({ ready: true, error: null }); })
      .catch((e) => setDbState({ ready: false, error: String(e.message || e) }));
  }, []);

  const fileA = files.A;
  const fileB = sameWb ? files.A : files.B;
  const resetResults = useCallback(() => { setRun(null); setExp(null); setError(null); setDoneUpto((d) => Math.min(d, 3)); }, []);

  // ------------------------------------------------------------ files
  async function addFile(side, file, o = {}) {
    setFileStatus((s) => ({ ...s, [side]: { name: file.name, size: file.size, status: 'Reading' } }));
    await tick();
    try {
      const bytes = o.bytes || new Uint8Array(await file.arrayBuffer());
      const sf = await openSource(file.name, bytes, o);
      setFiles((f) => ({ ...f, [side]: sf }));
      setFileStatus((s) => ({ ...s, [side]: null }));
      setPairs([]); setMaps({}); resetResults(); setDoneUpto(-1);
    } catch (e) {
      setFileStatus((s) => ({ ...s, [side]: { name: file.name, size: file.size, status: 'Error', error: String(e.message || e) } }));
    }
  }
  function removeFile(side) { setFiles((f) => ({ ...f, [side]: null })); setFileStatus((s) => ({ ...s, [side]: null })); setPairs([]); setMaps({}); resetResults(); setDoneUpto(-1); }

  // ------------------------------------------------------------ sheets
  useEffect(() => {
    if (step === 1 && fileA && fileB && !pairs.length) {
      const a = fileA.sheets[0];
      const b = sameWb && fileA.sheets.length > 1 ? fileA.sheets[1] : fileB.sheets.includes(a) ? a : fileB.sheets[0];
      setPairs([{ a_sheet: a, b_sheet: b, a_header: autoHeader(fileA, a), b_header: autoHeader(fileB, b) }]);
    }
  }, [step, fileA, fileB, pairs.length, sameWb]);

  function updatePair(i, patch) {
    setPairs((ps) => ps.map((p, k) => (k === i ? { ...p, ...patch } : p)));
    setMaps((m) => { const n = { ...m }; delete n[i]; return n; });
    resetResults();
  }

  // ------------------------------------------------------------ mapping (built lazily per pair)
  const getMap = useCallback((i) => {
    const p = pairs[i];
    const aCols = columnInfo(fileA, p.a_sheet, p.a_header).names;
    const bCols = columnInfo(fileB, p.b_sheet, p.b_header).names;
    const sig = JSON.stringify([p, aCols, bCols]);
    if (maps[i] && maps[i].sig === sig) return maps[i];
    let sugg = suggestPairs(aCols, bCols);
    let keys = [];
    const tp = pendingTpl && (pendingTpl.pairs || []).find((x) => x.a_sheet === p.a_sheet && x.b_sheet === p.b_sheet);
    if (tp) {
      keys = (tp.keys || []).filter((k) => aCols.includes(k.a) && bCols.includes(k.b));
      const chosen = Object.fromEntries((tp.columns || []).filter((c) => aCols.includes(c.a) && bCols.includes(c.b)).map((c) => [c.a, c]));
      sugg = aCols.filter((a) => chosen[a]).map((a) => ({ a, b: chosen[a].b, quality: 'Saved', compare: chosen[a].compare }));
    }
    if (!keys.length) {
      const prof = columnInfo(fileA, p.a_sheet, p.a_header).profile;
      keys = suggestKey(sugg, Object.fromEntries(Object.entries(prof).map(([k, v]) => [k, v.unique])));
      if (!keys.length) keys = [{ a: aCols[0], b: bCols[0] }];
    }
    const keyA = new Set(keys.map((k) => k.a));
    const byA = Object.fromEntries(sugg.map((s) => [s.a, s]));
    const cols = aCols.filter((a) => !keyA.has(a)).map((a) => (byA[a] ? { a, b: byA[a].b, compare: byA[a].compare, quality: byA[a].quality }
      : { a, b: NOT_MAPPED, compare: false, quality: 'No partner' }));
    return { sig, keys, cols, aCols, bCols };
  }, [pairs, maps, fileA, fileB, pendingTpl]);

  useEffect(() => {
    if (step === 2) {
      const need = pairs.map((_, i) => i).filter((i) => !maps[i] || maps[i].sig !== getMap(i).sig);
      if (need.length) { setMaps((m) => { const n = { ...m }; for (const i of need) n[i] = getMap(i); return n; }); }
    }
  }, [step, pairs, maps, getMap]);

  function setMap(i, patch) { setMaps((m) => ({ ...m, [i]: { ...m[i], ...patch } })); resetResults(); }

  // ------------------------------------------------------------ templates
  function currentSetup() {
    return { version: 1, saved: new Date().toISOString(),
      pairs: pairs.map((p, i) => ({ a_sheet: p.a_sheet, b_sheet: p.b_sheet, a_header_row: p.a_header, b_header_row: p.b_header,
        keys: (maps[i] || { keys: [] }).keys, columns: ((maps[i] || { cols: [] }).cols).filter((c) => c.b !== NOT_MAPPED).map((c) => ({ a: c.a, b: c.b, compare: c.compare })) })),
      rules, key_rules: keyRules, dup_policy: dupPolicy, compare_formulas: opts.compareFormulas, schema_compare: opts.schemaCompare,
      highlight_only_b: opts.highlightOnlyB, filenames: { highlighted: names.hl, report: names.rep } };
  }
  function applySetup(t) {
    const msgs = [];
    const np = [];
    for (const p of t.pairs || []) {
      if (fileA.sheets.includes(p.a_sheet) && fileB.sheets.includes(p.b_sheet)) {
        np.push({ a_sheet: p.a_sheet, b_sheet: p.b_sheet, a_header: +p.a_header_row || 1, b_header: +p.b_header_row || 1 });
      } else msgs.push(`Sheet pair ${p.a_sheet} ↔ ${p.b_sheet} is not in these files and was skipped.`);
    }
    if (np.length) { setPairs(np); setMaps({}); setPendingTpl(t); }
    setRules({ ...defaultRules(), ...(t.rules || {}) });
    setKeyRules({ ...defaultKeyRules(), ...(t.key_rules || {}) });
    setDupPolicy(t.dup_policy || 'keep_first');
    setOpts({ compareFormulas: !!t.compare_formulas, schemaCompare: !!t.schema_compare, highlightOnlyB: !!t.highlight_only_b });
    setNames({ hl: (t.filenames || {}).highlighted || DEFAULT_HL, rep: (t.filenames || {}).report || DEFAULT_REP });
    resetResults();
    return msgs;
  }

  // ------------------------------------------------------------ run
  function makeJob() {
    return { fileA, fileB, rules, keyRules, dupPolicy, compareFormulas: opts.compareFormulas, schemaCompare: opts.schemaCompare,
      highlightOnlyB: opts.highlightOnlyB,
      pairs: pairs.map((p, i) => ({ ...p, keys: maps[i].keys, columns: maps[i].cols.filter((c) => c.b !== NOT_MAPPED) })) };
  }
  async function start() {
    resetResults();
    setProg({ f: 0, msg: 'Starting…' });
    await tick();
    try {
      const db = dbRef.current.api;
      const r = await runComparison(db, makeJob(), (f, msg) => setProg({ f: f * 0.6, msg }));
      const e = await buildOutputs(db, r, { highlightedName: names.hl, progress: (f, msg) => setProg({ f: 0.6 + 0.4 * f, msg }) });
      setRun(r); setExp(e); setProg(null); setDoneUpto((d) => Math.max(d, 4));
    } catch (e) {
      console.error(e);
      setError(String(e.message || e)); setProg(null);
    }
  }

  // ------------------------------------------------------------ validity per step
  const mapsOk = pairs.length > 0 && pairs.every((_, i) => {
    const m = maps[i];
    if (!m) return false;
    const mapped = m.cols.filter((c) => c.b !== NOT_MAPPED);
    const bs = mapped.map((c) => c.b);
    const keyB = new Set(m.keys.map((k) => k.b));
    return m.keys.length && new Set(m.keys.map((k) => k.a)).size === m.keys.length && keyB.size === m.keys.length
      && new Set(bs).size === bs.length && !bs.some((b) => keyB.has(b));
  });
  const sheetsOk = pairs.length > 0 && new Set(pairs.map((p) => p.b_sheet)).size === pairs.length
    && !(sameWb && pairs.some((p) => p.a_sheet === p.b_sheet));
  const canNext = [!!(fileA && fileB && (!sameWb || fileA.isExcel)), sheetsOk, mapsOk, true, !!run, !!run, false][step];

  function go(n) { if (n > step) setDoneUpto((d) => Math.max(d, step)); setStep(n); window.scrollTo({ top: 0 }); }
  useEffect(() => { if (step >= 1 && !(fileA && fileB)) setStep(0); if (step >= 5 && !run) setStep(4); }, [step, fileA, fileB, run]);

  const props = { files, fileStatus, sameWb, setSameWb, addFile, removeFile, showSettings, setShowSettings, fileA, fileB, pairs,
    setPairs, updatePair, maps, setMap, rules, setRules, keyRules, setKeyRules, dupPolicy, setDupPolicy, opts, setOpts, resetResults,
    run, exp, error, prog, start, names, setNames, templates, setTemplates, currentSetup, applySetup, db: dbRef.current, dbState,
    setMaps, setDoneUpto, setStep, setFiles };
  const pages = [StepFiles, StepSheets, StepMatch, StepRules, StepCompare, StepResults, StepDownload];
  const Page = pages[step];
  return (
    <>
      <Header />
      <main>
        <Stepper current={step} doneUpto={doneUpto} />
        {dbState.error && <Note kind="err">The comparison engine could not start in this browser: {dbState.error}. Please use a recent
          Chrome, Edge or Firefox.</Note>}
        <Page {...props} />
      </main>
      <div className="bottombar"><div className="in">
        <div style={{ width: 140 }}>{step > 0 && <button className="btn" onClick={() => go(step - 1)}>← Back</button>}</div>
        <div className="mid">Step {step + 1} of {STEPS.length} · {STEPS[step]}</div>
        <div style={{ width: 170, textAlign: 'right' }}>{step < STEPS.length - 1 && (
          <button className="btn primary" disabled={!canNext || !!prog} onClick={() => go(step + 1)}>
            {{ 4: 'See results', 5: 'Go to downloads' }[step] || 'Next →'}</button>)}</div>
      </div></div>
    </>
  );
}

// =================================================================== step 1 — files
function FileCard({ side, title, sub, p }) {
  const sf = p.files[side];
  const st = p.fileStatus[side];
  const inputRef = useRef(null);
  const [over, setOver] = useState(false);
  const pick = (f) => f && p.addFile(side, f);
  return (
    <Card title={title} sub={sub}>
      {!sf && (!st || st.status === 'Error') && (
        <div className={`drop ${over ? 'over' : ''}`} onClick={() => inputRef.current.click()}
          onDragOver={(e) => { e.preventDefault(); setOver(true); }} onDragLeave={() => setOver(false)}
          onDrop={(e) => { e.preventDefault(); setOver(false); pick(e.dataTransfer.files[0]); }}>
          <div style={{ fontWeight: 550 }}>Drop a file here, or click to choose</div>
          <div className="muted">.xlsx · .xlsm · .xls · .xlsb · .csv · .txt · .tsv</div>
          <input ref={inputRef} type="file" hidden accept=".xlsx,.xlsm,.xls,.xlsb,.csv,.txt,.tsv" data-side={side}
            onChange={(e) => { pick(e.target.files[0]); e.target.value = ''; }} />
        </div>)}
      {st && (
        <div className="file-item"><span className="file-icon">{(st.name.split('.').pop() || '').toUpperCase().slice(0, 4)}</span>
          <div className="grow"><div className="file-name">{st.name}</div><div className="muted">{st.error || humanSize(st.size)}</div></div>
          <Tag kind={st.status === 'Error' ? 'err' : 'busy'}>{st.status === 'Reading' ? <><span className="spinner" /> Reading</> : st.status}</Tag></div>)}
      {sf && (<>
        <div className="file-item">
          <span className="file-icon">{sf.fmt.toUpperCase()}</span>
          <div className="grow"><div className="file-name" title={sf.name}>{sf.name}</div>
            <div className="muted">{kindLabel(sf.fmt)} · {humanSize(sf.size)} · {sf.isExcel ? `${sf.sheets.length} sheet(s)`
              : `${ENCODINGS[sf.encoding] || sf.encoding} · ${DELIMITERS[sf.delimiter] || sf.delimiter}`}</div></div>
          <Tag>Ready</Tag>
          <button className="icon-btn" title="File settings" disabled={sf.isExcel}
            onClick={() => p.setShowSettings((s) => ({ ...s, [side]: !s[side] }))}>⚙</button>
          <button className="icon-btn" title="Remove file" onClick={() => p.removeFile(side)}>×</button>
        </div>
        {!sf.isExcel && p.showSettings[side] && (
          <div className="grid2" style={{ marginTop: 8 }}>
            <div><label className="field">Encoding</label><select value={sf.encoding}
              onChange={(e) => p.addFile(side, { name: sf.name, size: sf.size }, { bytes: sf.bytes, encoding: e.target.value, delimiter: sf.delimiter })}>
              {Object.entries(ENCODINGS).map(([k, v]) => <option key={k} value={k}>{v}</option>)}</select></div>
            <div><label className="field">Separator</label><select value={sf.delimiter}
              onChange={(e) => p.addFile(side, { name: sf.name, size: sf.size }, { bytes: sf.bytes, encoding: sf.encoding, delimiter: e.target.value })}>
              {Object.entries(DELIMITERS).map(([k, v]) => <option key={k} value={k}>{v}</option>)}</select></div>
          </div>)}
        {sf.warnings.map((w) => <Note key={w} kind="warn">{w}</Note>)}
      </>)}
    </Card>
  );
}

function StepFiles(p) {
  return (<div className="fade">
    <PageTitle q="Which two files do you want to compare?" sub="File A is the earlier version (before). File B is the later version (after) — its copy is the one that gets highlighted." />
    <Checkbox checked={p.sameWb} onChange={(v) => { p.setSameWb(v); p.setPairs([]); p.setMaps({}); p.resetResults(); }}>Compare two sheets of the same workbook</Checkbox>
    {p.sameWb ? (<>
      <FileCard side="A" title="Workbook" sub="Pick the sheets to compare on the next step." p={p} />
      {p.files.A && !p.files.A.isExcel && <Note kind="warn">A CSV file has only one sheet. Untick this option and add two files instead.</Note>}
    </>) : (
      <div className="grid2">
        <FileCard side="A" title="File A — before" sub="Excel or CSV" p={p} />
        <FileCard side="B" title="File B — after" sub="Excel or CSV" p={p} />
      </div>)}
    <Note>Everything happens inside this browser tab. Files are never uploaded — you can even disconnect from the internet after the page has loaded.</Note>
  </div>);
}

// =================================================================== step 2 — sheets
function Preview({ sf, sheet, header }) {
  const pv = useMemo(() => previewRows(sf, sheet, header, 50), [sf, sheet, header]);
  return (<div className="table-wrap" style={{ maxHeight: 300 }}><table className="t"><thead><tr><th>Row</th>
    {pv.columns.map((c) => <th key={c.internal}>{c.letter} · {c.name}</th>)}</tr></thead>
  <tbody>{pv.rows.map(([r, vals]) => <tr key={r}><td className="num">{r}</td>{pv.columns.map((c, j) => <td key={j}>{vals[j] ?? ''}</td>)}</tr>)}</tbody></table></div>);
}

function TemplateLoader(p) {
  const fileRef = useRef(null);
  const [pick, setPick] = useState('');
  const [msg, setMsg] = useState(null);
  const nameList = Object.keys(p.templates).sort();
  const use = (t) => { const m = p.applySetup(t); setMsg(m.length ? m.join(' ') : 'Setup applied.'); };
  return (
    <More label="Saved setups">
      <div className="row">
        <select className="grow" value={pick} onChange={(e) => setPick(e.target.value)} style={{ flex: 3 }}>
          <option value="">{nameList.length ? 'Choose a saved setup…' : 'No saved setups in this browser yet'}</option>
          {nameList.map((n) => <option key={n}>{n}</option>)}</select>
        <button className="btn" disabled={!pick} onClick={() => use(p.templates[pick])}>Use</button>
        <button className="btn" disabled={!pick} onClick={() => { const t = { ...p.templates }; delete t[pick]; p.setTemplates(t); saveTemplates(t); setPick(''); setMsg('Setup deleted.'); }}>Delete</button>
        <button className="btn" onClick={() => fileRef.current.click()}>Open a setup file…</button>
        <input ref={fileRef} type="file" accept=".json" hidden onChange={async (e) => {
          const f = e.target.files[0]; e.target.value = '';
          if (!f) return;
          try { use(JSON.parse(await f.text())); } catch { setMsg('That file is not a saved setup.'); }
        }} />
      </div>
      {msg && <div className="muted" style={{ marginTop: 6 }}>{msg}</div>}
    </More>
  );
}

function StepSheets(p) {
  const { fileA, fileB, pairs } = p;
  const [prev, setPrev] = useState({});
  if (!pairs.length) return null;
  const bUsed = pairs.map((x) => x.b_sheet);
  return (<div className="fade">
    <PageTitle q="Which sheets should be compared?" sub="Pick a sheet from each file and check the row that holds the column headings." />
    <TemplateLoader {...p} />
    {pairs.map((pr, i) => (
      <Card key={i} title={pairs.length > 1 ? `Sheet pair ${i + 1}` : 'Sheets'} sub="Rows above the heading row (titles, notes) are ignored."
        right={pairs.length > 1 && <button className="icon-btn" title="Remove this pair" onClick={() => { p.setPairs(pairs.filter((_, k) => k !== i)); p.setMaps({}); p.resetResults(); }}>×</button>}>
        <div className="grid2">
          {[['a', fileA, 'File A'], ['b', fileB, 'File B']].map(([s, sf, label]) => {
            const info = columnInfo(sf, pr[`${s}_sheet`], pr[`${s}_header`]);
            const real = info.names.filter((n) => !n.startsWith('(no heading)'));
            return (<div key={s}>
              <div className="row">
                <div className="grow" style={{ flex: 3 }}><label className="field">{label} sheet</label>
                  <select value={pr[`${s}_sheet`]} disabled={!sf.isExcel} onChange={(e) => p.updatePair(i, { [`${s}_sheet`]: e.target.value, [`${s}_header`]: autoHeader(sf, e.target.value) })}>
                    {sf.sheets.map((n) => <option key={n}>{n}</option>)}</select></div>
                <div style={{ flex: 2 }}><label className="field">Headings on row</label>
                  <input type="number" min={1} max={1000} value={pr[`${s}_header`]} onChange={(e) => p.updatePair(i, { [`${s}_header`]: Math.max(1, +e.target.value || 1) })} /></div>
              </div>
              <div className="muted" style={{ marginTop: 6 }}>{real.length} columns found: {real.slice(0, 8).join(', ')}{real.length > 8 ? ' …' : ''}</div>
            </div>);
          })}
        </div>
        <button className="link-btn" onClick={() => setPrev((x) => ({ ...x, [i]: !x[i] }))}>{prev[i] ? '▾ Hide preview' : '▸ Preview the first 50 rows'}</button>
        {prev[i] && <div className="grid2"><div><div className="muted">File A</div><Preview sf={fileA} sheet={pr.a_sheet} header={pr.a_header} /></div>
          <div><div className="muted">File B</div><Preview sf={fileB} sheet={pr.b_sheet} header={pr.b_header} /></div></div>}
        {p.sameWb && pr.a_sheet === pr.b_sheet && <Note kind="warn">Both sides use the same sheet. Pick two different sheets.</Note>}
      </Card>))}
    {new Set(bUsed).size !== bUsed.length && <Note kind="warn">Each File B sheet can be used in only one pair, so every highlight belongs to one comparison.</Note>}
    {fileA.isExcel && fileB.isExcel && <button className="btn" onClick={() => {
      const aUsed = pairs.map((x) => x.a_sheet);
      // same workbook: never pair a sheet with itself, and prefer sheets not used yet
      const nxt = fileB.sheets.find((s) => !bUsed.includes(s) && !(p.sameWb && aUsed.includes(s)))
        || fileB.sheets.find((s) => !bUsed.includes(s));
      if (!nxt) return;
      const a = p.sameWb
        ? (fileA.sheets.find((s) => s !== nxt && !aUsed.includes(s) && !bUsed.includes(s)) || fileA.sheets.find((s) => s !== nxt))
        : (fileA.sheets.includes(nxt) ? nxt : fileA.sheets[0]);
      if (!a) return;
      p.setPairs([...pairs, { a_sheet: a, b_sheet: nxt, a_header: autoHeader(fileA, a), b_header: autoHeader(fileB, nxt) }]); p.resetResults();
    }}>+ Add another sheet pair</button>}
  </div>);
}

// =================================================================== step 3 — match
function MatchPair({ i, p }) {
  const m = p.maps[i];
  if (!m) return <Card><span className="spinner" /> Reading columns…</Card>;
  const { aCols, bCols } = m;
  const setKeys = (keys) => {
    const keyA = new Set(keys.map((k) => k.a));
    const old = Object.fromEntries(m.cols.map((c) => [c.a, c]));
    const sugg = Object.fromEntries(suggestPairs(aCols, bCols).map((s) => [s.a, s]));
    const cols = aCols.filter((a) => !keyA.has(a)).map((a) => old[a] || (sugg[a] ? { a, b: sugg[a].b, compare: true, quality: sugg[a].quality }
      : { a, b: NOT_MAPPED, compare: false, quality: 'No partner' }));
    p.setMap(i, { keys, cols });
  };
  const setCol = (k, patch) => p.setMap(i, { cols: m.cols.map((c, j) => (j === k ? { ...c, ...patch } : c)) });
  const mapped = m.cols.filter((c) => c.b !== NOT_MAPPED);
  const dupB = [...new Set(mapped.map((c) => c.b).filter((b, k, arr) => arr.indexOf(b) !== k))];
  const keyB = new Set(m.keys.map((k) => k.b));
  const clash = [...new Set(mapped.map((c) => c.b).filter((b) => keyB.has(b)))];
  const nCmp = mapped.filter((c) => c.compare).length;
  return (<>
    <Card title="Key column(s)" sub="Columns that identify the same record in both files, e.g. Consumer ID ↔ CA Number. Add more than one for a combined key.">
      {m.keys.map((k, j) => (
        <div className="row" key={j} style={{ marginBottom: 8 }}>
          <div className="grow">{j === 0 && <label className="field">File A</label>}
            <select value={k.a} onChange={(e) => setKeys(m.keys.map((x, q) => (q === j ? { ...x, a: e.target.value } : x)))}>{aCols.map((c) => <option key={c}>{c}</option>)}</select></div>
          <div className="muted" style={{ paddingTop: j === 0 ? 18 : 0 }}>↔</div>
          <div className="grow">{j === 0 && <label className="field">File B</label>}
            <select value={k.b} onChange={(e) => setKeys(m.keys.map((x, q) => (q === j ? { ...x, b: e.target.value } : x)))}>{bCols.map((c) => <option key={c}>{c}</option>)}</select></div>
          <div style={{ width: 34, paddingTop: j === 0 ? 18 : 0 }}>{m.keys.length > 1 && <button className="icon-btn" title="Remove" onClick={() => setKeys(m.keys.filter((_, q) => q !== j))}>×</button>}</div>
        </div>))}
      <button className="btn" onClick={() => {
        const used = new Set(m.keys.map((k) => k.a));
        const a = aCols.find((x) => !used.has(x)) || aCols[0];
        const c = m.cols.find((x) => x.a === a && x.b !== NOT_MAPPED);
        setKeys([...m.keys, { a, b: c ? c.b : bCols[0] }]);
      }}>+ Add key column</button>
      {(new Set(m.keys.map((k) => k.a)).size !== m.keys.length || keyB.size !== m.keys.length) && <Note kind="warn">A column is used twice in the key.</Note>}
    </Card>
    <Card title="Columns to compare" sub="Matching headings are paired for you. Change a File B column if the name differs, or untick a column to skip it.">
      <div className="table-wrap" style={{ maxHeight: 440 }}><table className="t"><thead><tr><th style={{ width: 80 }}>Compare</th><th>File A column</th><th>File B column</th><th>Match</th></tr></thead>
        <tbody>{m.cols.map((c, k) => (<tr key={c.a}>
          <td><input type="checkbox" checked={c.compare} disabled={c.b === NOT_MAPPED} onChange={(e) => setCol(k, { compare: e.target.checked })} /></td>
          <td>{c.a}</td>
          <td><select value={c.b} onChange={(e) => setCol(k, { b: e.target.value, quality: e.target.value === NOT_MAPPED ? 'No partner' : 'Chosen',
            compare: e.target.value === NOT_MAPPED ? false : (c.b === NOT_MAPPED ? true : c.compare) })}>
            <option>{NOT_MAPPED}</option>{bCols.map((b) => <option key={b}>{b}</option>)}</select></td>
          <td className="muted">{c.quality}</td></tr>))}</tbody></table></div>
      {dupB.length > 0 && <Note kind="warn">These File B columns are paired with more than one File A column: {dupB.join(', ')}</Note>}
      {clash.length > 0 && <Note kind="warn">These File B columns are already used as keys: {clash.join(', ')}</Note>}
      <div className="muted" style={{ marginTop: 8 }}>{nCmp} column(s) will be compared.{m.cols.length - mapped.length ? ` ${m.cols.length - mapped.length} column(s) in File A have no partner and are skipped.` : ''}</div>
    </Card>
  </>);
}

function StepMatch(p) {
  const [tab, setTab] = useState(0);
  const help = { keep_first: 'The first row with that key is compared; later ones are listed as excluded.',
    keep_last: 'The last row with that key is compared; earlier ones are listed as excluded.',
    pair: '1st with 1st, 2nd with 2nd… Extra rows are listed as only in A or B.',
    exclude: 'Rows with a repeated key are not compared; all are listed in the report.' };
  return (<div className="fade">
    <PageTitle q="How should rows be matched?" sub="Choose the column(s) that identify the same record in both files, then check the columns to compare." />
    {p.pairs.length > 1 && <div className="tabs">{p.pairs.map((x, i) => <button key={i} className={`tab ${tab === i ? 'on' : ''}`} onClick={() => setTab(i)}>{x.a_sheet} ↔ {x.b_sheet}</button>)}</div>}
    <MatchPair i={Math.min(tab, p.pairs.length - 1)} p={p} />
    <Card title="If a key appears more than once" sub="Rows are never paired many-to-many. Blank keys never match and are listed separately.">
      <div className="row">{Object.entries(POLICIES).map(([k, v]) => (
        <label key={k} className="check-line" style={{ marginRight: 14 }}><input type="radio" checked={p.dupPolicy === k} onChange={() => { p.setDupPolicy(k); p.resetResults(); }} /> {v}</label>))}</div>
      <div className="muted">{help[p.dupPolicy]}</div>
      <More>
        <div className="row">
          {[['trim', 'Ignore spaces around keys'], ['ignore_case', 'Ignore upper/lower case in keys'], ['leading_zeros', 'Ignore leading zeros in keys']].map(([k, l]) => (
            <Checkbox key={k} checked={p.keyRules[k]} onChange={(v) => { p.setKeyRules({ ...p.keyRules, [k]: v }); p.resetResults(); }}>{l}</Checkbox>))}
        </div>
        <div className="muted">Keys stay text: 00123 is never turned into the number 123 unless you tick “Ignore leading zeros”.</div>
      </More>
    </Card>
  </div>);
}

// =================================================================== step 4 — rules
function StepRules(p) {
  const r = p.rules;
  const set = (patch) => { p.setRules({ ...r, ...patch }); p.resetResults(); };
  const cols = [...new Set(p.pairs.flatMap((_, i) => (p.maps[i] ? p.maps[i].cols.filter((c) => c.compare && c.b !== NOT_MAPPED).map((c) => c.a) : [])))];
  const prof = Object.assign({}, ...p.pairs.map((x) => columnInfo(p.fileA, x.a_sheet, x.a_header).profile));
  const fx = (f) => f === 'xlsx' || f === 'xlsm';
  const bothX = fx(p.fileA.fmt) && fx(p.fileB.fmt);
  return (<div className="fade">
    <PageTitle q="How strict should the comparison be?" sub="Exact match is the default. Tick a rule to ignore that kind of difference. Rules only affect the comparison — your values are never changed." />
    <Card title="Ignore these differences">
      <div className="grid2">
        <div>
          <Checkbox checked={r.trim} onChange={(v) => set({ trim: v })}>Spaces at the start or end</Checkbox>
          <Checkbox checked={r.collapse_spaces} onChange={(v) => set({ collapse_spaces: v })}>Double spaces inside text</Checkbox>
          <Checkbox checked={r.ignore_case} onChange={(v) => set({ ignore_case: v })}>Upper / lower case</Checkbox>
          <Checkbox checked={r.leading_zeros} onChange={(v) => set({ leading_zeros: v })}>Leading zeros (007 = 7)</Checkbox>
          <Checkbox checked={r.blank_equals_empty} onChange={(v) => set({ blank_equals_empty: v })}>Blank and empty cells (spaces only = empty)</Checkbox>
        </div>
        <div>
          <Checkbox checked={r.line_breaks} onChange={(v) => set({ line_breaks: v })}>Line breaks inside a cell</Checkbox>
          <Checkbox checked={r.symbols} onChange={(v) => set({ symbols: v, symbol_chars: v ? (r.symbol_chars || '₹$,-/') : '' })}>Certain symbols</Checkbox>
          {r.symbols && <input type="text" value={r.symbol_chars} onChange={(e) => set({ symbol_chars: e.target.value })} style={{ marginBottom: 6 }} />}
          <Checkbox checked disabled title="Only values are compared">Cell formatting (colours, fonts, number formats)</Checkbox>
        </div>
      </div>
      <Checkbox checked={r.numeric} onChange={(v) => set({ numeric: v, numeric_columns: v ? (r.numeric_columns.length ? r.numeric_columns : cols.filter((c) => prof[c] && prof[c].numeric)) : [] })}>
        Number format (1000 = 1,000 = 1000.00) — only in the columns you pick</Checkbox>
      {r.numeric && <Chips options={cols} value={r.numeric_columns} onChange={(v) => set({ numeric_columns: v })} />}
      <Checkbox checked={r.dates} onChange={(v) => set({ dates: v, date_columns: v ? (r.date_columns.length ? r.date_columns : cols.filter((c) => prof[c] && prof[c].date)) : [] })}>
        Date format (2024-01-05 = 05-Jan-2024) — only in the columns you pick</Checkbox>
      {r.dates && <><Chips options={cols} value={r.date_columns} onChange={(v) => set({ date_columns: v })} />
        <div className="row" style={{ marginTop: 6 }}>{[['none', 'Only unambiguous dates (safe)'], ['dmy', '05/06/2024 means 5 June'], ['mdy', '05/06/2024 means May 6']].map(([k, l]) => (
          <label key={k} className="check-line" style={{ marginRight: 14 }}><input type="radio" checked={r.date_order === k} onChange={() => set({ date_order: k })} /> {l}</label>))}</div></>}
      {isExact(r) ? <Note>No rules ticked: values must match exactly.</Note>
        : <Note kind="ok">Values that are equal after these rules are <b>not</b> highlighted. The report shows which rules were applied to each column.</Note>}
    </Card>
    <More>
      <Checkbox checked={p.opts.compareFormulas} disabled={!bothX} onChange={(v) => { p.setOpts({ ...p.opts, compareFormulas: v }); p.resetResults(); }}>
        Also compare formulas (both files .xlsx) — flags cells whose value is the same but the formula differs</Checkbox>
      <Checkbox checked={p.opts.schemaCompare} onChange={(v) => { p.setOpts({ ...p.opts, schemaCompare: v }); p.resetResults(); }}>Report columns that exist in only one file</Checkbox>
      <Checkbox checked={p.opts.highlightOnlyB} onChange={(v) => { p.setOpts({ ...p.opts, highlightOnlyB: v }); p.resetResults(); }}>Fill rows found only in File B pale orange in the copy</Checkbox>
    </More>
  </div>);
}

// =================================================================== step 5 — compare
function StepCompare(p) {
  const job = { pairs: p.pairs.map((x, i) => ({ ...x, m: p.maps[i] })) };
  return (<div className="fade">
    <PageTitle q="Ready to compare?" sub="Check the plan below, then start. Nothing runs until you press Start." />
    <Card title="Plan">
      <div><b>File A:</b> {p.fileA.name} &nbsp;·&nbsp; <b>File B:</b> {p.fileB.name}</div>
      {job.pairs.map((x, i) => (<div key={i}><b>{x.a_sheet}</b> (row {x.a_header}) ↔ <b>{x.b_sheet}</b> (row {x.b_header}) · key:{' '}
        {x.m.keys.map((k) => (k.a === k.b ? k.a : `${k.a} ↔ ${k.b}`)).join(', ')} · {x.m.cols.filter((c) => c.compare && c.b !== NOT_MAPPED).length} column(s)</div>))}
      <div><b>Rules:</b> {summaryLines(p.rules).slice(0, -1).join('; ')}</div>
      <div><b>Duplicate keys:</b> {POLICIES[p.dupPolicy]}</div>
    </Card>
    {!p.dbState.ready && !p.dbState.error && <Note><span className="spinner" /> Starting the comparison engine…</Note>}
    <button className="btn primary" disabled={!p.dbState.ready || !!p.prog} onClick={p.start}>{p.run ? 'Run again' : 'Start comparison'}</button>
    {p.prog && (<div style={{ marginTop: 12 }}><div className="progress"><div style={{ width: `${Math.round(p.prog.f * 100)}%` }} /></div>
      <div className="muted">{p.prog.msg}</div></div>)}
    {p.error && <Note kind="err">The comparison stopped: {p.error}</Note>}
    {p.run && p.exp && (p.exp.passed
      ? <Note kind="ok">Finished in {p.run.durationS.toFixed(1)} s. {fmt(p.run.totals.matched)} records matched, {fmt(p.run.totals.changed_cells)} changed cells. All checks passed.</Note>
      : <Note kind="err">Finished, but some checks failed. Open Results → Validation before using the files.</Note>)}
  </div>);
}

function StepResults(p) {
  return (<>
    <PageTitle q="What changed?" sub="The key numbers first. Open “All numbers” or the tabs below for details." />
    <Results db={p.db.api} run={p.run} exp={p.exp} />
  </>);
}

// =================================================================== step 7 — download
function StepDownload(p) {
  const { exp } = p;
  const [tplName, setTplName] = useState('');
  const [saved, setSaved] = useState('');
  const hlExt = exp.hlName.endsWith('.xlsm') ? '.xlsm' : '.xlsx';
  const stem = (n, d) => (n || '').replace(/\.[^.]+$/, '') || d;
  return (<div className="fade">
    <PageTitle q="Your files are ready" sub="Rename them if you like, then download. They are created in this browser and saved wherever your browser saves downloads." />
    {!exp.passed && <Note kind="err">Some checks failed, so these files are <b>not</b> marked as a successful comparison.</Note>}
    <div className="grid2">
      <Card title="1 · Highlighted copy of File B" sub={`${fmt(exp.hl.yellowCells)} changed cell(s) filled yellow. Values are unchanged.`}>
        <label className="field">File name</label>
        <input type="text" value={p.names.hl} onChange={(e) => p.setNames({ ...p.names, hl: e.target.value })} />
        <div style={{ margin: '10px 0' }}><button className="btn primary" onClick={() => download(stem(p.names.hl, 'File_B_Comparison_Highlighted') + hlExt, exp.hlBytes,
          hlExt === '.xlsm' ? 'application/vnd.ms-excel.sheet.macroEnabled.12' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}>Download highlighted copy</button></div>
        {exp.hl.notes.map((n) => <div key={n} className="muted">{n}</div>)}
      </Card>
      <Card title="2 · Before–After report" sub="Changed cells, changed records, rows only in A or B, duplicate keys and a summary with validation.">
        <label className="field">File name</label>
        <input type="text" value={p.names.rep} onChange={(e) => p.setNames({ ...p.names, rep: e.target.value })} />
        <div style={{ margin: '10px 0' }}><button className="btn primary" onClick={() => download(stem(p.names.rep, 'Detailed_Before_After_Comparison') + '.xlsx', exp.reportBytes,
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}>Download report</button></div>
      </Card>
    </div>
    <Card title="Save this setup" sub="Keeps sheets, headings, keys, column choices, rules and file names in this browser. Files themselves are never saved.">
      <div className="row">
        <input type="text" className="grow" placeholder="e.g. Monthly consumer check" value={tplName} onChange={(e) => setTplName(e.target.value)} style={{ flex: 3 }} />
        <button className="btn" disabled={!tplName.trim()} onClick={() => {
          const t = { ...p.templates, [tplName.trim()]: p.currentSetup() };
          const existed = tplName.trim() in p.templates;
          p.setTemplates(t); saveTemplates(t); setSaved(existed ? 'Setup updated.' : 'Setup saved.');
        }}>{tplName.trim() in p.templates ? 'Update' : 'Save'}</button>
        <button className="btn" onClick={() => download(`${(tplName.trim() || 'comparison-setup').replace(/[^\w\- ]/g, '_')}.json`,
          new TextEncoder().encode(JSON.stringify(p.currentSetup(), null, 2)), 'application/json')}>Download as file</button>
      </div>
      <div className="muted" style={{ marginTop: 6 }}>{saved || (Object.keys(p.templates).length ? `Saved: ${Object.keys(p.templates).join(', ')}` : '')}</div>
    </Card>
    <button className="btn" onClick={() => { p.setFiles({ A: null, B: null }); p.setPairs([]); p.setMaps({}); p.resetResults(); p.setDoneUpto(-1); p.setStep(0); }}>Start a new comparison</button>
  </div>);
}

createRoot(document.getElementById('root')).render(<App />);
