// CSV reading that numbers records the way Excel shows rows:
// blank lines count as rows, and a quoted cell with line breaks stays one row.
// Behaves like Python's csv.reader (quote only special at the start of a field, "" = ").

export const DELIMITERS = { ',': 'Comma ( , )', '\t': 'Tab', ';': 'Semicolon ( ; )', '|': 'Pipe ( | )' };
export const ENCODINGS = { 'utf-8': 'UTF-8', 'utf-8-sig': 'UTF-8 with BOM', 'windows-1252': 'Windows-1252' };

export function detectEncoding(bytes) {
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) return 'utf-8-sig';
  if ((bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0xfe && bytes[1] === 0xff)) {
    throw new Error('This file is UTF-16. Please save it as UTF-8 CSV and try again.');
  }
  try {
    new TextDecoder('utf-8', { fatal: true }).decode(bytes);
    return 'utf-8';
  } catch {
    return 'windows-1252';
  }
}

export function decodeText(bytes, encoding) {
  if (encoding === 'windows-1252') return new TextDecoder('windows-1252').decode(bytes);
  // the UTF-8 decoder removes a leading BOM
  return new TextDecoder('utf-8').decode(bytes);
}

// Iterate records: yields arrays of strings ([] for a blank line).
export function* parseCsv(text, delim = ',') {
  const n = text.length;
  let i = 0;
  if (n === 0) return;
  let field = '';
  let rec = [];
  let inQuotes = false;
  let atStart = true;          // at the start of a field
  let sawAny = false;          // record has any content (incl. empty quoted field)
  const d = delim;
  while (i < n) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
        inQuotes = false; i++; continue;
      }
      // fast path: copy until next quote
      const j = text.indexOf('"', i);
      if (j === -1) { field += text.slice(i); i = n; continue; }
      field += text.slice(i, j); i = j; continue;
    }
    if (ch === '"' && atStart) { inQuotes = true; atStart = false; sawAny = true; i++; continue; }
    if (ch === d) { rec.push(field); field = ''; atStart = true; sawAny = true; i++; continue; }
    if (ch === '\n' || ch === '\r') {
      if (sawAny || field !== '') rec.push(field);
      yield rec;
      rec = []; field = ''; atStart = true; sawAny = false;
      if (ch === '\r' && text[i + 1] === '\n') i += 2; else i++;
      continue;
    }
    // plain run of characters
    let j = i;
    while (j < n) {
      const c = text[j];
      if (c === d || c === '\n' || c === '\r') break;
      j++;
    }
    field += text.slice(i, j); atStart = false; sawAny = true; i = j;
  }
  if (inQuotes || sawAny || field !== '') {
    rec.push(field);
    yield rec;
  }
}

export function detectDelimiter(sample, fileName = '') {
  if (fileName.toLowerCase().endsWith('.tsv')) return '\t';
  const lines = sample.split(/\r?\n/).slice(0, 200).join('\n');
  let best = ',';
  let bestScore = [-1, 0];
  for (const d of Object.keys(DELIMITERS)) {
    const counts = [];
    for (const r of parseCsv(lines, d)) if (r.length) counts.push(r.length);
    if (!counts.length) continue;
    const freq = new Map();
    for (const c of counts) freq.set(c, (freq.get(c) || 0) + 1);
    let mode = 0, modeN = 0;
    for (const [k, v] of freq) if (v > modeN || (v === modeN && k > mode)) { mode = k; modeN = v; }
    if (mode < 2) continue;
    const score = [modeN / counts.length, mode];
    if (score[0] > bestScore[0] || (score[0] === bestScore[0] && score[1] > bestScore[1])) { best = d; bestScore = score; }
  }
  return best;
}
