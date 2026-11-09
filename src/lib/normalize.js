// Comparison rules as DuckDB SQL. Rules change only the comparison copy, never source values.
export const RULE_LABELS = {
  trim: 'Trim spaces', collapse_spaces: 'Collapse multiple spaces', ignore_case: 'Ignore case',
  leading_zeros: 'Ignore leading zeros', numeric: 'Number format', dates: 'Date format',
  blank_equals_empty: 'Blank = empty', symbols: 'Ignore symbols', line_breaks: 'Ignore line breaks',
};

export const defaultRules = () => ({
  trim: false, collapse_spaces: false, ignore_case: false, leading_zeros: false,
  numeric: false, numeric_columns: [], dates: false, date_columns: [], date_order: 'none',
  blank_equals_empty: false, symbols: false, symbol_chars: '', line_breaks: false, ignore_formatting: true,
});
export const defaultKeyRules = () => ({ trim: true, ignore_case: false, leading_zeros: false });

export const lit = (s) => `'${String(s).replace(/'/g, "''")}'`;

export function isExact(r) {
  return !(r.trim || r.collapse_spaces || r.ignore_case || r.leading_zeros || (r.numeric && r.numeric_columns.length)
    || (r.dates && r.date_columns.length) || r.blank_equals_empty || (r.symbols && r.symbol_chars) || r.line_breaks);
}

export function appliedFor(r, aCol) {
  const out = [];
  if (r.line_breaks) out.push(RULE_LABELS.line_breaks);
  if (r.symbols && r.symbol_chars) out.push(`${RULE_LABELS.symbols} (${r.symbol_chars})`);
  if (r.trim) out.push(RULE_LABELS.trim);
  if (r.collapse_spaces) out.push(RULE_LABELS.collapse_spaces);
  if (r.blank_equals_empty) out.push(RULE_LABELS.blank_equals_empty);
  if (r.dates && r.date_columns.includes(aCol)) out.push(RULE_LABELS.dates);
  if (r.numeric && r.numeric_columns.includes(aCol)) out.push(RULE_LABELS.numeric);
  if (r.leading_zeros) out.push(RULE_LABELS.leading_zeros);
  if (r.ignore_case) out.push(RULE_LABELS.ignore_case);
  return out;
}

export function summaryLines(r) {
  const lines = [];
  for (const [k, label] of Object.entries(RULE_LABELS)) {
    if (k === 'numeric') { if (r.numeric && r.numeric_columns.length) lines.push(`${label}: ${r.numeric_columns.join(', ')}`); }
    else if (k === 'dates') {
      if (r.dates && r.date_columns.length) {
        const order = { none: 'unambiguous dates only', dmy: 'day first', mdy: 'month first' }[r.date_order];
        lines.push(`${label} (${order}): ${r.date_columns.join(', ')}`);
      }
    } else if (k === 'symbols') { if (r.symbols && r.symbol_chars) lines.push(`${label}: ${r.symbol_chars}`); }
    else if (r[k]) lines.push(label);
  }
  if (!lines.length) lines.push('Exact match (no rules)');
  lines.push('Cell formatting ignored (only values are compared)');
  return lines;
}

function regexClass(chars) {
  const out = [];
  for (const ch of new Set([...chars])) {
    if (/\s/.test(ch)) continue;
    out.push('\\]^-['.includes(ch) ? `\\${ch}` : ch);
  }
  return out.join('');
}

function dateSql(x, order) {
  const iso = `try_strptime(${x}, ['%Y-%m-%d', '%Y-%m-%d %H:%M:%S', '%Y/%m/%d', '%Y/%m/%d %H:%M:%S', '%Y-%m-%dT%H:%M:%S', '%Y.%m.%d'])`;
  const named = `try_strptime(${x}, ['%d-%b-%Y', '%d %b %Y', '%d-%B-%Y', '%d %B %Y', '%d/%b/%Y', '%b %d, %Y', '%B %d, %Y', '%b %d %Y', '%B %d %Y', '%d-%b-%Y %H:%M:%S', '%d %b %Y %H:%M:%S'])`;
  const pat = String.raw`^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})(?: 00:00(?::00)?)?$`;
  const p1 = `TRY_CAST(regexp_extract(${x}, '${pat}', 1) AS INTEGER)`;
  const p2 = `TRY_CAST(regexp_extract(${x}, '${pat}', 2) AS INTEGER)`;
  const yy = `TRY_CAST(regexp_extract(${x}, '${pat}', 3) AS INTEGER)`;
  // try_strptime gives NULL for impossible dates such as 31/02/2024
  const dmy = `try_strptime(printf('%04d-%02d-%02d', ${yy}, ${p2}, ${p1}), '%Y-%m-%d')`;
  const mdy = `try_strptime(printf('%04d-%02d-%02d', ${yy}, ${p1}, ${p2}), '%Y-%m-%d')`;
  const amb = order === 'dmy' ? dmy : order === 'mdy' ? mdy : 'NULL::TIMESTAMP';
  const numeric = `CASE WHEN ${p1} IS NULL THEN NULL::TIMESTAMP WHEN ${p1} > 12 AND ${p2} BETWEEN 1 AND 12 THEN ${dmy} `
    + `WHEN ${p2} > 12 AND ${p1} BETWEEN 1 AND 12 THEN ${mdy} WHEN ${p1} = ${p2} THEN ${dmy} ELSE ${amb} END`;
  const ts = `COALESCE(${iso}, ${named}, ${numeric})`;
  return `(CASE WHEN ${ts} IS NULL THEN ${x} WHEN ${ts} = date_trunc('day', ${ts}) THEN strftime(${ts}, '%Y-%m-%d') `
    + `ELSE strftime(${ts}, '%Y-%m-%d %H:%M:%S') END)`;
}

function numericSql(x) {
  const dec = `TRY_CAST(replace(replace(${x}, ',', ''), ' ', '') AS DECIMAL(38,10))`;
  const t = `CAST(${dec} AS VARCHAR)`;
  return `(CASE WHEN ${dec} IS NULL THEN ${x} WHEN ${dec} = 0 THEN '0' WHEN position('.' IN ${t}) > 0 `
    + `THEN rtrim(rtrim(${t}, '0'), '.') ELSE ${t} END)`;
}

export function valueSql(col, r, numeric, date) {
  let x = col;
  if (r.line_breaks) x = `regexp_replace(${x}, '\\r\\n|\\r|\\n', ' ', 'g')`;
  if (r.symbols && r.symbol_chars) { const c = regexClass(r.symbol_chars); if (c) x = `regexp_replace(${x}, ${lit(`[${c}]`)}, '', 'g')`; }
  if (r.trim) x = `regexp_replace(${x}, '^\\s+|\\s+$', '', 'g')`;
  if (r.collapse_spaces) x = `regexp_replace(${x}, '\\s+', ' ', 'g')`;
  if (r.blank_equals_empty) x = `(CASE WHEN regexp_matches(${x}, '^\\s*$') THEN NULL ELSE ${x} END)`;
  if (date && r.dates) x = dateSql(x, r.date_order);
  if (numeric && r.numeric) x = numericSql(x);
  if (r.leading_zeros) x = `regexp_replace(${x}, '^0+([0-9])', '\\1')`;
  if (r.ignore_case) x = `lower(${x})`;
  return x;
}

export function keySql(col, kr) {
  let x = col;
  if (kr.trim) x = `regexp_replace(${x}, '^\\s+|\\s+$', '', 'g')`;
  if (kr.leading_zeros) x = `regexp_replace(${x}, '^0+([0-9])', '\\1')`;
  if (kr.ignore_case) x = `lower(${x})`;
  return `NULLIF(regexp_replace(${x}, '^\\s+$', ''), '')`;
}
