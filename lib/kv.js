// Small key-value store shared by every server instance.
//
//   - Turso (libSQL) when TURSO_DATABASE_URL / TURSO_AUTH_TOKEN are set — the shared store in use.
//   - Upstash Redis (REST) when UPSTASH_REDIS_REST_URL/_TOKEN (or the Vercel
//     KV_REST_API_URL/_TOKEN aliases) are set — this is the real, shared store.
//   - A JSON file (.data/kv.json) when running locally, so work survives restarts.
//   - Plain memory on a host with no database configured (the old behaviour);
//     a warning is logged because instances then no longer agree with each other.
//
// Values are JSON-serialisable; every call is async.

const fs = require('fs');
const path = require('path');

const REST_URL = process.env.UPSTASH_REDIS_REST_URL || process.env.KV_REST_API_URL;
const REST_TOKEN = process.env.UPSTASH_REDIS_REST_TOKEN || process.env.KV_REST_API_TOKEN;

function restBackend(url, token) {
  async function command(args) {
    const res = await fetch(url, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(args),
    });
    const body = await res.json().catch(() => ({}));
    if (!res.ok || body.error) throw new Error(`KV ${args[0]} failed: ${body.error || res.status}`);
    return body.result;
  }
  const parse = (v) => (v == null ? null : JSON.parse(v));
  return {
    name: 'upstash',
    async get(key) { return parse(await command(['GET', key])); },
    async set(key, value) { await command(['SET', key, JSON.stringify(value)]); },
    async mget(keys) { return keys.length ? (await command(['MGET', ...keys])).map(parse) : []; },
    async del(key) { await command(['DEL', key]); },
  };
}

// Turso (libSQL) over its HTTP API — no driver needed, so it runs the same on
// Vercel and locally. One table, kv(key, value-as-JSON, updated_at).
function tursoBackend(databaseUrl, authToken) {
  const base = databaseUrl.replace(/^libsql:\/\//, 'https://').replace(/\/$/, '');
  const text = (v) => ({ type: 'text', value: String(v) });
  const cell = (c) => (c && c.type !== 'null' ? c.value : null);

  async function run(statements) {
    const requests = statements.map((s) => ({ type: 'execute', stmt: { sql: s.sql, args: (s.args || []).map(text) } }));
    requests.push({ type: 'close' });
    const res = await fetch(`${base}/v2/pipeline`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${authToken}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ requests }),
    });
    const body = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(`Turso request failed: ${body.error || res.status}`);
    return statements.map((_, i) => {
      const r = body.results && body.results[i];
      if (!r || r.type !== 'ok') throw new Error(`Turso error: ${(r && r.error && r.error.message) || 'no result'}`);
      return r.response.result;
    });
  }

  let ready = null;
  const init = () => ready || (ready = run([{
    sql: 'CREATE TABLE IF NOT EXISTS kv (key TEXT PRIMARY KEY, value TEXT NOT NULL, updated_at TEXT NOT NULL)',
  }]).catch((err) => { ready = null; throw err; }));

  return {
    name: 'turso',
    async get(key) {
      await init();
      const [r] = await run([{ sql: 'SELECT value FROM kv WHERE key = ?', args: [key] }]);
      return r.rows.length ? JSON.parse(cell(r.rows[0][0])) : null;
    },
    async set(key, value) {
      await init();
      await run([{
        sql: 'INSERT INTO kv (key, value, updated_at) VALUES (?, ?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = excluded.updated_at',
        args: [key, JSON.stringify(value), new Date().toISOString()],
      }]);
    },
    async mget(keys) {
      if (!keys.length) return [];
      await init();
      const found = new Map();
      for (let i = 0; i < keys.length; i += 200) {
        const chunk = keys.slice(i, i + 200);
        const [r] = await run([{ sql: `SELECT key, value FROM kv WHERE key IN (${chunk.map(() => '?').join(',')})`, args: chunk }]);
        r.rows.forEach((row) => found.set(cell(row[0]), JSON.parse(cell(row[1]))));
      }
      return keys.map((k) => (found.has(k) ? found.get(k) : null));
    },
    async del(key) {
      await init();
      await run([{ sql: 'DELETE FROM kv WHERE key = ?', args: [key] }]);
    },
  };
}

function fileBackend(file) {
  let data = null;
  const load = () => {
    if (data) return data;
    try { data = JSON.parse(fs.readFileSync(file, 'utf8')); } catch { data = {}; }
    return data;
  };
  return {
    name: 'file',
    async get(key) { return load()[key] ?? null; },
    async set(key, value) {
      load()[key] = value;
      fs.mkdirSync(path.dirname(file), { recursive: true });
      const tmp = `${file}.tmp`;
      fs.writeFileSync(tmp, JSON.stringify(data));
      fs.renameSync(tmp, file); // atomic, so a crash never leaves half a file
    },
    async mget(keys) { return keys.map((k) => load()[k] ?? null); },
    async del(key) { delete load()[key]; fs.mkdirSync(path.dirname(file), { recursive: true }); fs.writeFileSync(file, JSON.stringify(data)); },
  };
}

function memoryBackend() {
  const data = new Map();
  return {
    name: 'memory',
    async get(key) { return data.has(key) ? data.get(key) : null; },
    async set(key, value) { data.set(key, value); },
    async mget(keys) { return keys.map((k) => (data.has(k) ? data.get(k) : null)); },
    async del(key) { data.delete(key); },
  };
}

function chooseBackend() {
  if (process.env.TURSO_DATABASE_URL && process.env.TURSO_AUTH_TOKEN) return tursoBackend(process.env.TURSO_DATABASE_URL, process.env.TURSO_AUTH_TOKEN);
  if (REST_URL && REST_TOKEN) return restBackend(REST_URL, REST_TOKEN);
  if (process.env.VERCEL) {
    console.warn('No KV database configured (UPSTASH_REDIS_REST_URL / _TOKEN): state is per-instance and resets on restart.');
    return memoryBackend();
  }
  return fileBackend(path.join(__dirname, '..', '.data', 'kv.json'));
}

module.exports = chooseBackend();
module.exports.restBackend = restBackend; // exposed for tests
module.exports.tursoBackend = tursoBackend;
