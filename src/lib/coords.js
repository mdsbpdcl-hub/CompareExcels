// Real source coordinates. Rows/columns are 1-based exactly as Excel shows them.
// CSV rows are record numbers as Excel shows them when it opens the CSV.
export const EXCEL_MAX_ROWS = 1048576;
export const EXCEL_MAX_CELL_CHARS = 32767;

export function colLetter(col) {
  if (col < 1) throw new Error(`Column number must be 1 or more, got ${col}`);
  let s = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}

export function colNumber(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) {
    if (ch < 'A' || ch > 'Z') throw new Error(`Not a column letter: ${letters}`);
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n;
}

export const cellRef = (col, row) => `${colLetter(col)}${row}`;

export function parseRef(ref) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref.trim());
  if (!m) throw new Error(`Not a cell reference: ${ref}`);
  return [colNumber(m[1]), Number(m[2])];
}

// Where a source row lands when a very tall sheet is split into parts.
// Part 1 holds rows 1..max unchanged; later parts repeat the heading on row 1.
export function splitLocation(row, max = EXCEL_MAX_ROWS) {
  if (row <= max) return [1, row];
  const per = max - 1;
  const off = row - max - 1;
  return [2 + Math.floor(off / per), 2 + (off % per)];
}
