import React, { useEffect, useMemo, useState } from 'react';

export const STEPS = ['Files', 'Sheets', 'Match', 'Rules', 'Compare', 'Results', 'Download'];
export const fmt = (n) => (typeof n === 'number' ? n.toLocaleString('en-IN') : n);

export function Header() {
  return (
    <div className="header"><div className="header-in">
      <span className="logo" /><span className="app-name">Excel &amp; CSV Compare</span>
      <span className="app-note"><span className="dot" />Files never leave this browser</span>
    </div></div>
  );
}

export function Stepper({ current, doneUpto }) {
  return (
    <div className="stepper">
      {STEPS.map((name, i) => {
        const done = i < current || (i <= doneUpto && i !== current);
        const cls = i === current ? 'active' : done ? 'done' : '';
        return (
          <React.Fragment key={name}>
            <div className={`step ${cls}`}><span className="num">{cls === 'done' ? '✓' : i + 1}</span><span className="lbl">{name}</span></div>
            {i < STEPS.length - 1 && <div className={`line ${i < current ? 'done' : ''}`} />}
          </React.Fragment>
        );
      })}
    </div>
  );
}

export const PageTitle = ({ q, sub }) => (<><div className="q-title">{q}</div><div className="q-sub">{sub}</div></>);
export const Note = ({ kind = '', children }) => <div className={`note ${kind}`}>{children}</div>;
export const Tag = ({ kind = 'ok', children }) => <span className={`tag ${kind}`}>{children}</span>;
export const Card = ({ title, sub, children, right }) => (
  <div className="card fade">
    {(title || right) && <div className="row"><div className="grow">{title && <div className="card-title">{title}</div>}
      {sub && <div className="card-sub">{sub}</div>}</div>{right}</div>}
    {children}
  </div>
);
export const More = ({ label = 'Advanced options', children, open = false }) => (
  <details className="more" open={open}><summary>{label}</summary><div className="inner">{children}</div></details>
);

export function humanSize(n) {
  const u = ['bytes', 'KB', 'MB', 'GB'];
  let i = 0;
  while (n >= 1024 && i < 3) { n /= 1024; i++; }
  return i === 0 ? `${n} bytes` : `${n.toFixed(1)} ${u[i]}`;
}

export function Checkbox({ checked, onChange, children, disabled, title }) {
  return (
    <label className="check-line" title={title} style={disabled ? { opacity: 0.7, cursor: 'default' } : null}>
      <input type="checkbox" checked={checked} disabled={disabled} onChange={(e) => onChange && onChange(e.target.checked)} />
      <span>{children}</span>
    </label>
  );
}

export function Chips({ options, value, onChange }) {
  return (
    <div>{options.map((o) => {
      const on = value.includes(o);
      return <span key={o} className={`chip ${on ? 'on' : ''}`} onClick={() => onChange(on ? value.filter((x) => x !== o) : [...value, o])}>
        {on ? '✓' : '+'} {o}</span>;
    })}</div>
  );
}

export const Tiles = ({ items }) => (
  <div className="tiles">{items.map(([v, l, accent]) => (
    <div key={l} className={`tile ${accent ? 'accent' : ''}`}><div className="v">{v}</div><div className="l">{l}</div></div>))}</div>
);

export function CheckList({ checks }) {
  return <div>{checks.map((c, i) => (
    <div className="check" key={i}><Tag kind={!c.critical ? 'info' : c.passed ? 'ok' : 'err'}>{c.status}</Tag>
      <div><div className="t">{c.name}</div><div className="d">{c.detail}</div></div></div>))}</div>;
}

/** A table whose rows come from SQL, one page at a time (server-side pagination in DuckDB). */
export function SqlTable({ db, sql, columns, numeric = [], deps = [] }) {
  const [size, setSize] = useState(50);
  const [page, setPage] = useState(1);
  const [total, setTotal] = useState(0);
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(false);
  useEffect(() => { setPage(1); }, [sql, size, ...deps]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => {
    let live = true;
    (async () => {
      setLoading(true);
      const [{ n }] = await db.query(`SELECT count(*) AS n FROM (${sql})`);
      const pages = Math.max(1, Math.ceil(n / size));
      const p = Math.min(page, pages);
      const data = n ? await db.query(`${sql} LIMIT ${size} OFFSET ${(p - 1) * size}`) : [];
      if (live) { setTotal(n); setRows(data); setLoading(false); }
    })().catch(() => setLoading(false));
    return () => { live = false; };
  }, [db, sql, size, page]);
  const pages = Math.max(1, Math.ceil(total / size));
  const cols = useMemo(() => columns || (rows[0] ? Object.keys(rows[0]) : []), [columns, rows]);
  if (!loading && total === 0) return <Note>Nothing to show here.</Note>;
  return (
    <div>
      <div className="row" style={{ marginBottom: 8 }}>
        <select value={size} onChange={(e) => setSize(+e.target.value)} style={{ width: 160 }}>
          {[50, 100, 250, 500].map((s) => <option key={s} value={s}>{s} rows per page</option>)}
        </select>
        <button className="btn" disabled={page <= 1} onClick={() => setPage(page - 1)}>‹</button>
        <input type="number" min={1} max={pages} value={page} style={{ width: 90 }}
          onChange={(e) => setPage(Math.min(pages, Math.max(1, +e.target.value || 1)))} />
        <button className="btn" disabled={page >= pages} onClick={() => setPage(page + 1)}>›</button>
        <span className="muted">Page {fmt(page)} of {fmt(pages)} · {fmt(total)} rows {loading && <span className="spinner" />}</span>
      </div>
      <div className="table-wrap">
        <table className="t"><thead><tr>{cols.map((c) => <th key={c}>{c}</th>)}</tr></thead>
          <tbody>{rows.map((r, i) => <tr key={i}>{cols.map((c) => (
            <td key={c} className={numeric.includes(c) ? 'num' : ''}>{r[c] === null || r[c] === undefined ? '' : String(r[c])}</td>))}</tr>)}</tbody>
        </table>
      </div>
    </div>
  );
}
