// Small database adapter so the engine runs the same in the browser (AsyncDuckDB) and in Node tests.
const plain = (row) => {
  const o = row.toJSON();
  for (const k of Object.keys(o)) if (typeof o[k] === 'bigint') o[k] = Number(o[k]);
  return o;
};

export function wrapAsync(conn) {
  return {
    async query(sql) { return (await conn.query(sql)).toArray().map(plain); },
    async exec(sql) { await conn.query(sql); },
    async insertIPC(name, ipc) { await conn.insertArrowFromIPCStream(ipc, { name, create: true }); },
  };
}

export function wrapBlocking(conn) {
  return {
    async query(sql) { return conn.query(sql).toArray().map(plain); },
    async exec(sql) { conn.query(sql); },
    async insertIPC(name, ipc) { conn.insertArrowFromIPCStream(ipc, { name, create: true }); },
  };
}

/** Browser: start DuckDB-Wasm from files bundled with the app (no CDN). */
export async function openBrowserDb(bundles) {
  const duckdb = await import('@duckdb/duckdb-wasm');
  const bundle = await duckdb.selectBundle(bundles);
  const worker = new Worker(bundle.mainWorker);
  const db = new duckdb.AsyncDuckDB(new duckdb.VoidLogger(), worker);
  await db.instantiate(bundle.mainModule);
  const conn = await db.connect();
  return { db, conn, api: wrapAsync(conn) };
}
