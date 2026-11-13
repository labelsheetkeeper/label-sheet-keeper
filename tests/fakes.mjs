// Stand-ins for Cloudflare's storage, shared by the Worker tests: the D1
// database is Node's built-in SQLite (the same SQL runs), the KV namespace a
// Map. Also a stand-in for Microsoft's sign-in: a signing key published the
// way Microsoft publishes its keys, and ID tokens signed with it.
import { DatabaseSync } from 'node:sqlite';

class Stmt {
  constructor(db, sql, args) { this.db = db; this.sql = sql; this.args = args || []; }
  bind(...args) { return new Stmt(this.db, this.sql, args); }
  runNow() {
    const s = this.db.prepare(this.sql);
    if (/\bRETURNING\b/i.test(this.sql) || /^\s*SELECT\b/i.test(this.sql)) return { success: true, results: s.all(...this.args).map((r) => ({ ...r })), meta: { changes: 0 } };
    const r = s.run(...this.args);
    return { success: true, meta: { changes: Number(r.changes) } };
  }
  async run() { return this.runNow(); }
  async first() { const r = this.db.prepare(this.sql).get(...this.args); return r === undefined ? null : { ...r }; }
  async all() { return { results: this.db.prepare(this.sql).all(...this.args).map((r) => ({ ...r })) }; }
}
export class FakeD1 {
  constructor() { this.db = new DatabaseSync(':memory:'); }
  prepare(sql) { return new Stmt(this.db, sql); }
  async batch(stmts) {
    this.db.exec('BEGIN');
    try { const out = stmts.map((s) => s.runNow()); this.db.exec('COMMIT'); return out; } catch (e) { this.db.exec('ROLLBACK'); throw e; }
  }
  rows(sql, ...args) { return this.db.prepare(sql).all(...args).map((r) => ({ ...r })); }
}
export class FakeKV {
  constructor() { this.m = new Map(); this.hidden = new Set(); }
  async getWithMetadata(key) {
    const v = this.hidden.has(key) ? null : this.m.get(key);
    return v ? { value: v.value, metadata: v.metadata || null } : { value: null, metadata: null };
  }
  async put(key, value, opts = {}) { this.m.set(key, { value, metadata: opts.metadata, ttl: opts.expirationTtl }); }
  async delete(key) { this.m.delete(key); }
  keys(prefix) { return [...this.m.keys()].filter((k) => k.startsWith(prefix)); }
}

const MSA = '9188040d-6c67-4c5b-b112-36a304b66dad';   // Microsoft's "tenant" for personal accounts
const KEYS_URL = 'https://login.microsoftonline.com/common/discovery/v2.0/keys';
const b64u = (x) => Buffer.from(x).toString('base64url');

// -> { keys, fetch(url) for the keys address, idToken(claims) }
export async function fakeMicrosoft(clientId) {
  const k = await crypto.subtle.generateKey({ name: 'RSASSA-PKCS1-v1_5', modulusLength: 2048, publicExponent: new Uint8Array([1, 0, 1]), hash: 'SHA-256' }, true, ['sign', 'verify']);
  const j = await crypto.subtle.exportKey('jwk', k.publicKey);
  const keys = { keys: [{ kid: 'test', kty: 'RSA', use: 'sig', n: j.n, e: j.e, issuer: 'https://login.microsoftonline.com/{tenantid}/v2.0' }] };
  return {
    keys,
    fetch: async (u) => { if (String(u) !== KEYS_URL) throw new Error('unexpected fetch ' + u); return new Response(JSON.stringify(keys)); },
    // oid: which person. Each person is a personal Microsoft account unless tid says otherwise.
    async idToken(person, over = {}) {
      const tid = over.tid || MSA, now = Math.floor(Date.now() / 1000);
      const claims = { aud: clientId, iss: 'https://login.microsoftonline.com/' + tid + '/v2.0', tid, oid: '00000000-0000-0000-0000-' + String(person).padStart(12, '0'), name: 'Person ' + person, preferred_username: 'person' + person + '@example.com', iat: now - 5, nbf: now - 5, exp: now + 3600, ...over };
      const data = b64u(JSON.stringify({ typ: 'JWT', alg: 'RS256', kid: 'test' })) + '.' + b64u(JSON.stringify(claims));
      return data + '.' + b64u(await crypto.subtle.sign('RSASSA-PKCS1-v1_5', k.privateKey, new TextEncoder().encode(data)));
    }
  };
}
