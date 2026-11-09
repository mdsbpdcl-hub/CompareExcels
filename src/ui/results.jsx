import React, { useMemo, useState } from 'react';
import { SqlTable, Tiles, More, Note, CheckList, fmt } from './common.jsx';
import { lit } from '../lib/normalize.js';

const qi = (n) => `"${String(n).replace(/"/g, '""')}"`;
const TYPES = ['Value Changed', 'Blank-to-Value', 'Value-to-Blank', 'Formula Difference'];

function ChangedCells({ db, run }) {
  const [q, setQ] = useState('');
  const [pair, setPair] = useState('all');
  const [type, setType] = useState('all');
  const [col, setCol] = useState('all');
  const [rFrom, setRFrom] = useState('');
  const [rTo, setRTo] = useState('');
  const [sort, setSort] = useState('report');
  const cols = useMemo(() => [...new Set(run.pairs.flatMap((p) => p.compareCols.map((c) => c.a_name)))].sort(), [run]);
  const aCase = run.pairs.map((p) => `WHEN ${p.pairId} THEN ${lit(p.a.sheet)}`).join(' ');
  const bCase = run.pairs.map((p) => `WHEN ${p.pairId} THEN ${lit(p.b.sheet)}`).join(' ');
  const where = ['1=1'];
  if (q) where.push(`d.key_disp ILIKE ${lit(`%${q}%`)}`);
  if (pair !== 'all') where.push(`d.pair_id = ${+pair}`);
  if (type !== 'all') where.push(`d.diff_type = ${lit(type)}`);
  if (col !== 'all') where.push(`d.a_col_name = ${lit(col)}`);
  if (rFrom) where.push(`d.b_row >= ${+rFrom}`);
  if (rTo) where.push(`d.b_row <= ${+rTo}`);
  const order = { report: 'd.sr', most: 'rw.n_changed DESC, d.sr', key: 'd.key_disp, d.sr' }[sort];
  const sql = `SELECT d.sr AS "No.", d.key_disp AS "Key", d.a_col_name AS "Column",
    (CASE d.pair_id ${aCase} END) || '!' || d.a_letter || d.a_row AS "Before cell", COALESCE(d.before_value, '(blank)') AS "Before",
    COALESCE(d.after_value, '(blank)') AS "After", (CASE d.pair_id ${bCase} END) || '!' || d.b_letter || d.b_row AS "After cell",
    d.diff_type AS "Type", rw.n_changed AS "Changes in record", d.remarks AS "Remarks"
    FROM diffs d JOIN rowwise rw ON rw.pair_id = d.pair_id AND rw.a_row = d.a_row AND rw.b_row = d.b_row
    WHERE ${where.join(' AND ')} ORDER BY ${order}`;
  return (
    <div>
      <div className="row" style={{ marginBottom: 6 }}>
        <input type="text" className="grow" placeholder="Search by key…" value={q} onChange={(e) => setQ(e.target.value)} style={{ flex: 3 }} />
        {run.pairs.length > 1 && <select value={pair} onChange={(e) => setPair(e.target.value)} style={{ flex: 2 }}>
          <option value="all">All sheets</option>{run.pairs.map((p) => <option key={p.pairId} value={p.pairId}>{p.label}</option>)}</select>}
        <select value={type} onChange={(e) => setType(e.target.value)} style={{ flex: 2 }}>
          <option value="all">All types</option>{TYPES.map((t) => <option key={t}>{t}</option>)}</select>
      </div>
      <More label="More filters">
        <div className="row">
          <div className="grow"><label className="field">Column</label><select value={col} onChange={(e) => setCol(e.target.value)}>
            <option value="all">All columns</option>{cols.map((c) => <option key={c}>{c}</option>)}</select></div>
          <div style={{ width: 120 }}><label className="field">B row from</label><input type="number" value={rFrom} onChange={(e) => setRFrom(e.target.value)} /></div>
          <div style={{ width: 120 }}><label className="field">B row to</label><input type="number" value={rTo} onChange={(e) => setRTo(e.target.value)} /></div>
          <div className="grow"><label className="field">Sort by</label><select value={sort} onChange={(e) => setSort(e.target.value)}>
            <option value="report">Report order</option><option value="most">Most differences in the record</option><option value="key">Match key</option></select></div>
        </div>
      </More>
      <SqlTable db={db} sql={sql} numeric={['No.', 'Changes in record']} />
    </div>
  );
}

function ChangedRecords({ db, run }) {
  const [q, setQ] = useState('');
  const [most, setMost] = useState(false);
  const bCase = run.pairs.map((p) => `WHEN ${p.pairId} THEN ${lit(p.b.sheet)}`).join(' ');
  const sql = `SELECT key_disp AS "Key", (CASE pair_id ${bCase} END) AS "Sheet (B)", a_row AS "Row in A", b_row AS "Row in B",
    n_changed AS "Changed columns", changed_columns AS "Which columns", description AS "What changed" FROM rowwise
    WHERE ${q ? `key_disp ILIKE ${lit(`%${q}%`)}` : '1=1'} ORDER BY ${most ? 'n_changed DESC, pair_id, b_row' : 'pair_id, b_row'}`;
  return (<div>
    <div className="row" style={{ marginBottom: 8 }}>
      <input type="text" className="grow" placeholder="Search by key…" value={q} onChange={(e) => setQ(e.target.value)} />
      <select value={most ? 'm' : 'o'} onChange={(e) => setMost(e.target.value === 'm')} style={{ width: 240 }}>
        <option value="o">File B row order</option><option value="m">Most differences first</option></select>
    </div>
    <SqlTable db={db} sql={sql} numeric={['Row in A', 'Row in B', 'Changed columns']} />
  </div>);
}

function OnlyIn({ db, run, side }) {
  const [q, setQ] = useState('');
  const [pid, setPid] = useState(run.pairs[0].pairId);
  const p = run.pairs.find((x) => x.pairId === +pid);
  const st = side === 'a' ? p.a : p.b;
  const sql = `SELECT key_disp AS "Key", src_row AS "Row", status AS "Status", ${st.columns.map((c) => `${c.internal} AS ${qi(c.name)}`).join(', ')}
    FROM p${p.pairId}_o${side} WHERE ${q ? `key_disp ILIKE ${lit(`%${q}%`)}` : '1=1'} ORDER BY src_row`;
  return (<div>
    <div className="row" style={{ marginBottom: 8 }}>
      <input type="text" className="grow" placeholder="Search by key…" value={q} onChange={(e) => setQ(e.target.value)} />
      {run.pairs.length > 1 && <select value={pid} onChange={(e) => setPid(+e.target.value)} style={{ width: 280 }}>
        {run.pairs.map((x) => <option key={x.pairId} value={x.pairId}>{x.label}</option>)}</select>}
    </div>
    <SqlTable db={db} sql={sql} numeric={['Row']} />
  </div>);
}

function Duplicates({ db }) {
  const [q, setQ] = useState('');
  const sql = `SELECT file_side AS "File", sheet AS "Sheet", key_disp AS "Key", src_row AS "Row", occurrence AS "Occurrence",
    occurrences AS "Of", outcome AS "Result" FROM dups WHERE ${q ? `key_disp ILIKE ${lit(`%${q}%`)}` : '1=1'}
    ORDER BY pair_id, file_side, key_disp, src_row`;
  return (<div>
    <input type="text" placeholder="Search by key…" value={q} onChange={(e) => setQ(e.target.value)} style={{ marginBottom: 8 }} />
    <SqlTable db={db} sql={sql} numeric={['Row', 'Occurrence', 'Of']} />
  </div>);
}

function SummaryTable({ run }) {
  return (<div className="table-wrap"><table className="t"><thead><tr>
    {['Sheets', 'Rows A', 'Rows B', 'Matched', 'With differences', 'Identical', 'Changed cells', 'Only in A', 'Only in B',
      'Duplicates excluded (A/B)', 'Blank keys (A/B)'].map((h) => <th key={h}>{h}</th>)}</tr></thead>
  <tbody>{run.pairs.map((p) => { const s = p.stats; return (<tr key={p.pairId}>
    <td>{p.label}</td><td className="num">{fmt(p.a.sourceRows)}</td><td className="num">{fmt(p.b.sourceRows)}</td>
    <td className="num">{fmt(s.matched)}</td><td className="num">{fmt(s.changed_records)}</td><td className="num">{fmt(s.identical)}</td>
    <td className="num">{fmt(s.changed_cells)}</td><td className="num">{fmt(s.only_a)}</td><td className="num">{fmt(s.only_b)}</td>
    <td>{fmt(s.dup_excluded_a)} / {fmt(s.dup_excluded_b)}</td><td>{fmt(s.blank_key_a)} / {fmt(s.blank_key_b)}</td></tr>); })}</tbody>
  </table></div>);
}

export function Results({ db, run, exp }) {
  const [tab, setTab] = useState('Summary');
  const t = run.totals;
  const tabs = ['Summary', 'Changed Cells', 'Changed Records', 'Only in A', 'Only in B', 'Duplicates', 'Validation'];
  const notes = [...new Set([...run.warnings, ...exp.hl.notes, ...run.job.fileA.warnings, ...run.job.fileB.warnings])];
  return (
    <div className="fade">
      {!exp.passed && <Note kind="err"><b>Some checks failed.</b> These results are not marked as successful — see the Validation tab.</Note>}
      <Tiles items={[[fmt(t.matched), 'Matched records'], [fmt(t.changed_records), 'With differences', true],
        [fmt(t.changed_cells), 'Changed cells', true], [`${fmt(t.only_a)} / ${fmt(t.only_b)}`, 'Only in A / Only in B']]} />
      <More label="All numbers">
        <SummaryTable run={run} />
        <div className="grid2" style={{ marginTop: 10 }}>
          <div className="muted">Identical records: <b>{fmt(t.identical)}</b><br />Rows in A / B: <b>{fmt(t.rows_a)} / {fmt(t.rows_b)}</b><br />
            Empty rows skipped (A / B): <b>{fmt(t.empty_a)} / {fmt(t.empty_b)}</b></div>
          <div className="muted">Value changed: <b>{fmt(t.value_changed)}</b> · Blank → value: <b>{fmt(t.blank_to_value)}</b> ·
            Value → blank: <b>{fmt(t.value_to_blank)}</b> · Formula: <b>{fmt(t.formula_diff)}</b><br />
            Duplicate keys A / B: <b>{fmt(t.dup_keys_a)} / {fmt(t.dup_keys_b)}</b><br />Time taken: <b>{run.durationS.toFixed(1)} s</b></div>
        </div>
        {run.pairs.filter((p) => p.schemaOnlyA.length || p.schemaOnlyB.length).map((p) => (
          <div key={p.pairId} className="muted">{p.label}: only in A — {p.schemaOnlyA.join(', ') || 'none'}; only in B — {p.schemaOnlyB.join(', ') || 'none'}</div>))}
      </More>
      <div className="tabs">{tabs.map((x) => <button key={x} className={`tab ${tab === x ? 'on' : ''}`} onClick={() => setTab(x)}>{x}</button>)}</div>
      {tab === 'Summary' && <><SummaryTable run={run} />{notes.map((n) => <Note key={n} kind={/not |lost/i.test(n) ? 'warn' : ''}>{n}</Note>)}</>}
      {tab === 'Changed Cells' && <ChangedCells db={db} run={run} />}
      {tab === 'Changed Records' && <ChangedRecords db={db} run={run} />}
      {tab === 'Only in A' && <OnlyIn db={db} run={run} side="a" />}
      {tab === 'Only in B' && <OnlyIn db={db} run={run} side="b" />}
      {tab === 'Duplicates' && <Duplicates db={db} />}
      {tab === 'Validation' && <><Note kind={exp.passed ? 'ok' : 'err'}>{exp.passed ? <><b>PASS</b> — every critical check passed.</>
        : <><b>FAIL</b> — at least one critical check failed.</>}</Note><CheckList checks={exp.checks} /></>}
    </div>
  );
}
