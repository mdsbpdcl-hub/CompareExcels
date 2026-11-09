// Suggest how columns in File A line up with columns in File B.
const SYNONYMS = [
  ['id', 'no', 'number', 'num', 'code', 'ref', 'reference'],
  ['consumer', 'customer', 'client', 'ca', 'account', 'acct', 'connection'],
  ['name', 'fullname'],
  ['amount', 'amt', 'value', 'total'],
  ['mobile', 'phone', 'contact', 'cell'],
  ['address', 'addr'],
  ['date', 'dt', 'dated'],
  ['qty', 'quantity'],
].map((g) => new Set(g));
const UNIT_WORDS = new Set(['rs', 'inr', 'kw', 'kwh', 'kva', 'the', 'of', 'in', 'and', 'mm', 'cm', 'kg', 'pct']);
const KEY_HINTS = /(^|\b|_)(id|no|number|code|ca|account|acct|ref|key|sl|serial|consumer)(\b|_|$)/i;

const norm = (s) => s.toLowerCase().replace(/[^a-z0-9]/g, '');

function tokens(name) {
  const out = new Set();
  for (const t of name.toLowerCase().match(/[a-z]+|\d+/g) || []) {
    const g = SYNONYMS.find((x) => x.has(t));
    out.add(g ? `~${[...g].sort()[0]}` : t);
  }
  return out;
}

function ratio(a, b) {           // difflib-like similarity: 2*LCS-ish / total (longest common subsequence)
  const m = a.length, n = b.length;
  if (!m || !n) return 0;
  const dp = new Array(n + 1).fill(0);
  for (let i = 1; i <= m; i++) {
    let prev = 0;
    for (let j = 1; j <= n; j++) {
      const tmp = dp[j];
      dp[j] = a[i - 1] === b[j - 1] ? prev + 1 : Math.max(dp[j], dp[j - 1]);
      prev = tmp;
    }
  }
  return (2 * dp[n]) / (m + n);
}

export function similarity(a, b) {
  const na = norm(a), nb = norm(b);
  if (!na || !nb) return 0;
  if (na === nb) return 1;
  const seq = ratio(na, nb) * 0.9;
  let ta = new Set([...tokens(a)].filter((t) => !UNIT_WORDS.has(t)));
  let tb = new Set([...tokens(b)].filter((t) => !UNIT_WORDS.has(t)));
  if (!ta.size || !tb.size) { ta = tokens(a); tb = tokens(b); }
  const inter = [...ta].filter((t) => tb.has(t)).length;
  const uni = new Set([...ta, ...tb]).size;
  const jac = uni ? inter / uni : 0;
  const [small, big] = ta.size <= tb.size ? [ta, tb] : [tb, ta];
  const contained = small.size && [...small].every((t) => big.has(t)) && [...small].some((t) => t.replace(/^~/, '').length >= 4) ? 0.8 : 0;
  return Math.max(seq, jac * 0.95, contained);
}

/** Greedy one-to-one: exact names first, then most similar. Returns [{a, b, quality, compare}] */
export function suggestPairs(aCols, bCols, threshold = 0.72) {
  const pairs = [];
  const usedA = new Set(), usedB = new Set();
  const bByNorm = new Map();
  for (const b of bCols) if (!bByNorm.has(norm(b))) bByNorm.set(norm(b), b);
  for (const a of aCols) {
    if (a.startsWith('(no heading)')) continue;
    const b = bByNorm.get(norm(a));
    if (b && !usedB.has(b)) { pairs.push({ a, b, quality: 'Same name', compare: true }); usedA.add(a); usedB.add(b); }
  }
  const scored = [];
  for (const a of aCols) {
    if (usedA.has(a) || a.startsWith('(no heading)')) continue;
    for (const b of bCols) {
      if (usedB.has(b) || b.startsWith('(no heading)')) continue;
      const s = similarity(a, b);
      if (s >= threshold) scored.push([s, a, b]);
    }
  }
  scored.sort((x, y) => y[0] - x[0]);
  for (const [, a, b] of scored) {
    if (usedA.has(a) || usedB.has(b)) continue;
    pairs.push({ a, b, quality: 'Similar', compare: true }); usedA.add(a); usedB.add(b);
  }
  const order = new Map(aCols.map((a, i) => [a, i]));
  pairs.sort((p, q) => order.get(p.a) - order.get(q.a));
  return pairs;
}

export function suggestKey(pairs, uniqueRatio = {}) {
  if (!pairs.length) return [];
  let best = null, bestScore = null;
  pairs.forEach((p, i) => {
    const u = uniqueRatio[p.a] || 0;
    const score = [u >= 0.98 ? 1 : 0, KEY_HINTS.test(p.a) || KEY_HINTS.test(p.b) ? 1 : 0, u, -i];
    if (!bestScore || compare(score, bestScore) > 0) { best = p; bestScore = score; }
  });
  return [{ a: best.a, b: best.b }];
}

function compare(x, y) {
  for (let i = 0; i < x.length; i++) if (x[i] !== y[i]) return x[i] - y[i];
  return 0;
}
