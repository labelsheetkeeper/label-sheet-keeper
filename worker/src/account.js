// Accounts and sync.
//
// Signing in. The panel signs the user in with Microsoft and sends the
// Microsoft ID token to POST /auth/microsoft. The Worker checks Microsoft's
// signature on it and that it was made for this app (MS_CLIENT_ID), then
// hands back a session token of its own. The account is known only by a hash
// of Microsoft's identifiers for the person; no name, email or password is
// stored here. Session tokens are stored hashed.
//
// Syncing. An account holds records: { c: collection, id, m: when it last
// changed (the changing computer's clock, ms), d: deleted, v: its JSON text }.
// POST /sync sends this computer's changes and gets back everyone else's
// since last time. For one record the newest change wins. Every stored change
// gets the account's next "rev" number, so "everything after rev N" is exact.
//
//   POST /sync  { since, push: [{ c, id, m, d, v }] }
//     -> { pull: [{ c, id, m, d, v, r }], next, more }
//   next goes back as "since" next time. more: ask again right away.

import { ALLOWED_ORIGINS, json, randomId, sha256, COUNT_SQL, today } from './shared.js';

const DAY = 24 * 3600 * 1000;
const MAX_BODY = 4 * 1024 * 1024;      // one sync request
const MAX_PUSH = 300;                  // records in one request
const MAX_VALUE = 300 * 1024;          // one record
const MAX_RECORDS = 20000;             // in an account
const MAX_ACCOUNT_BYTES = 20 * 1024 * 1024;
const PAGE = 300;                      // records handed back at a time
const MAX_PUSHES_PER_DAY = 3000;       // requests with changes, per account
const AUTH_PER_DAY = 100;              // sign-ins from one internet address
const MAX_SESSIONS = 30;               // computers signed in to one account
const SESSION_IDLE = 400 * DAY;        // a session unused this long ends

const COLLECTIONS = ['book', 'set', 'sheet'];
const ID_RE = /^[A-Za-z0-9._:-]{1,100}$/;
const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export const ACCOUNT_SCHEMA = [
  'CREATE TABLE IF NOT EXISTS accounts (id TEXT PRIMARY KEY, created INTEGER NOT NULL, rev INTEGER NOT NULL DEFAULT 0, day TEXT NOT NULL DEFAULT \'\', pushes INTEGER NOT NULL DEFAULT 0)',
  'CREATE TABLE IF NOT EXISTS sessions (hash TEXT PRIMARY KEY, account TEXT NOT NULL, created INTEGER NOT NULL, seen INTEGER NOT NULL)',
  'CREATE INDEX IF NOT EXISTS sessions_account ON sessions (account, seen)',
  'CREATE TABLE IF NOT EXISTS records (account TEXT NOT NULL, c TEXT NOT NULL, id TEXT NOT NULL, m INTEGER NOT NULL, d INTEGER NOT NULL DEFAULT 0, v TEXT, rev INTEGER NOT NULL, PRIMARY KEY (account, c, id))',
  'CREATE INDEX IF NOT EXISTS records_rev ON records (account, rev)'
];

// ---------- checking a Microsoft ID token ----------
// Microsoft signs ID tokens with keys it publishes here and changes now and then.
const MS_KEYS_URL = 'https://login.microsoftonline.com/common/discovery/v2.0/keys';
let msKeys = null;   // { at, keys: Map(kid -> key) }

async function msKey(kid) {
  const now = Date.now();
  const stale = !msKeys || now - msKeys.at > DAY;
  // An unknown key may be a brand new one: look again, but not more than every 5 minutes.
  const unknown = msKeys && !msKeys.keys.has(kid) && now - msKeys.at > 5 * 60 * 1000;
  if (stale || unknown) {
    const r = await fetch(MS_KEYS_URL);
    if (!r.ok) throw new Error('Microsoft keys ' + r.status);
    const j = await r.json();
    msKeys = { at: now, keys: new Map(((j && j.keys) || []).map((k) => [k.kid, k])) };
  }
  return msKeys.keys.get(kid) || null;
}

function fromB64url(s) {
  const b = atob(s.replace(/-/g, '+').replace(/_/g, '/') + '==='.slice((s.length + 3) % 4));
  const out = new Uint8Array(b.length);
  for (let i = 0; i < b.length; i++) out[i] = b.charCodeAt(i);
  return out;
}
const fromB64urlJson = (s) => JSON.parse(new TextDecoder().decode(fromB64url(s)));

// The token's claims when it's a genuine, unexpired Microsoft ID token for this app, else null.
export async function verifyMicrosoft(token, clientId) {
  const parts = String(token || '').split('.');
  if (parts.length !== 3) return null;
  let head, claims, sig;
  try { head = fromB64urlJson(parts[0]); claims = fromB64urlJson(parts[1]); sig = fromB64url(parts[2]); } catch (e) { return null; }
  if (!head || head.alg !== 'RS256' || typeof head.kid !== 'string' || !claims) return null;
  const jwk = await msKey(head.kid);
  if (!jwk || jwk.kty !== 'RSA') return null;
  const key = await crypto.subtle.importKey('jwk', { kty: 'RSA', n: jwk.n, e: jwk.e, alg: 'RS256', ext: true }, { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' }, false, ['verify']);
  const good = await crypto.subtle.verify('RSASSA-PKCS1-v1_5', key, sig, new TextEncoder().encode(parts[0] + '.' + parts[1]));
  if (!good) return null;

  const now = Date.now() / 1000;
  if (claims.aud !== clientId) return null;                                   // made for another app
  if (!GUID_RE.test(claims.tid || '') || !GUID_RE.test(claims.oid || '')) return null;
  const iss = 'https://login.microsoftonline.com/' + claims.tid + '/v2.0';
  if (claims.iss !== iss) return null;
  // Each key says which issuer may use it ("{tenantid}" stands for any organization).
  if (jwk.issuer && jwk.issuer.replace('{tenantid}', claims.tid) !== iss) return null;
  if (!(claims.exp > now - 60) || (claims.nbf && claims.nbf > now + 300)) return null;
  return claims;
}

// ---------- sessions ----------
async function sessionOf(request, env) {
  const m = /^Bearer\s+([A-Za-z0-9_-]{40,100})$/.exec(request.headers.get('Authorization') || '');
  if (!m) return null;
  const hash = await sha256(m[1]);
  const row = await env.DB.prepare('SELECT s.hash, s.account, s.seen, a.rev, a.day, a.pushes, a.created FROM sessions s JOIN accounts a ON a.id = s.account WHERE s.hash = ?').bind(hash).first();
  if (!row) return null;
  const now = Date.now();
  if (now - row.seen > SESSION_IDLE) {
    await env.DB.prepare('DELETE FROM sessions WHERE hash = ?').bind(hash).run();
    return null;
  }
  if (now - row.seen > DAY) await env.DB.prepare('UPDATE sessions SET seen = ? WHERE hash = ?').bind(now, hash).run();
  return row;
}

async function signIn(request, env, origin) {
  if (!ALLOWED_ORIGINS.includes(origin)) return json({ error: 'Sign in from the Label Sheet Keeper panel.' }, 403, origin);
  if (!env.MS_CLIENT_ID) return json({ error: 'not set up' }, 503, origin);
  const day = today();
  const who = 'auth:' + (await sha256((request.headers.get('CF-Connecting-IP') || 'unknown') + '|' + day)).slice(0, 24);
  const tries = await env.DB.prepare(COUNT_SQL).bind(day, who).first();
  if (tries && tries.n > AUTH_PER_DAY) return json({ error: 'Too many sign-ins from here today. Try again tomorrow.' }, 429, origin);

  let body = null;
  try { body = JSON.parse((await request.text()).slice(0, 20000)); } catch (e) { body = null; }
  let claims = null;
  try { claims = await verifyMicrosoft(body && body.id_token, env.MS_CLIENT_ID); } catch (e) {
    return json({ error: 'Microsoft couldn’t be reached to check the sign-in. Try again.' }, 503, origin);
  }
  if (!claims) return json({ error: 'That sign-in couldn’t be confirmed. Sign in again.' }, 401, origin);

  const account = (await sha256('ms|' + claims.tid.toLowerCase() + '|' + claims.oid.toLowerCase())).slice(0, 40);
  const now = Date.now(), token = randomId(32);
  const out = await env.DB.batch([
    env.DB.prepare('INSERT INTO accounts (id, created) VALUES (?, ?) ON CONFLICT (id) DO NOTHING').bind(account, now),
    env.DB.prepare('INSERT INTO sessions (hash, account, created, seen) VALUES (?, ?, ?, ?)').bind(await sha256(token), account, now, now),
    // Past the most computers for one account, the ones unused longest are signed out.
    env.DB.prepare('DELETE FROM sessions WHERE account = ?1 AND hash NOT IN (SELECT hash FROM sessions WHERE account = ?1 ORDER BY seen DESC, created DESC LIMIT ?2)').bind(account, MAX_SESSIONS)
  ]);
  return json({ session: token, isNew: !!(out[0].meta && out[0].meta.changes) }, 201, origin);
}

// ---------- sync ----------
function cleanPush(list, now) {
  if (!Array.isArray(list) || list.length > MAX_PUSH) return null;
  const out = [], seen = new Set();
  for (const p of list) {
    if (!p || !COLLECTIONS.includes(p.c) || typeof p.id !== 'string' || !ID_RE.test(p.id)) return null;
    if (!Number.isFinite(p.m) || p.m <= 0) return null;
    const d = p.d ? 1 : 0;
    if (!d && (typeof p.v !== 'string' || !p.v.length || p.v.length > MAX_VALUE)) return null;
    const k = p.c + '/' + p.id;
    if (seen.has(k)) return null;
    seen.add(k);
    // A clock set ahead can't make a change win forever.
    out.push({ c: p.c, id: p.id, m: Math.min(Math.round(p.m), now + 5 * 60 * 1000), d, v: d ? null : p.v });
  }
  return out;
}
const rowOut = (r) => ({ c: r.c, id: r.id, m: r.m, d: r.d ? 1 : 0, v: r.d ? null : r.v, r: r.rev });

async function sync(request, env, origin, me) {
  if (Number(request.headers.get('Content-Length') || 0) > MAX_BODY) return json({ error: 'That’s too much to sync at once.' }, 413, origin);
  const text = await request.text();
  if (text.length > MAX_BODY) return json({ error: 'That’s too much to sync at once.' }, 413, origin);
  let body = null;
  try { body = JSON.parse(text); } catch (e) { body = null; }
  const now = Date.now();
  const since = body && Number.isInteger(body.since) && body.since >= 0 ? body.since : 0;
  const push = cleanPush(body && body.push ? body.push : [], now);
  if (!body || !push) return json({ error: 'The sync request wasn’t understood.' }, 400, origin);
  const n = push.length, db = env.DB, day = today();

  if (n) {
    if (me.day === day && me.pushes >= MAX_PUSHES_PER_DAY) return json({ error: 'Too many changes today. Syncing picks up again tomorrow.' }, 429, origin);
    const adding = push.filter((p) => !p.d);
    if (adding.length) {
      const have = await db.prepare('SELECT COUNT(*) AS n, COALESCE(SUM(LENGTH(v)), 0) AS b FROM records WHERE account = ?').bind(me.account).first();
      const bytes = adding.reduce((t, p) => t + p.v.length, 0);
      if (have.n + adding.length > MAX_RECORDS || have.b + bytes > MAX_ACCOUNT_BYTES) return json({ error: 'Your account is full. Delete addresses you no longer need, then sync again.' }, 413, origin);
    }
  }

  // One transaction: number the changes, keep the newer of each record, read what's new for this computer.
  const batch = [];
  if (n) batch.push(db.prepare('UPDATE accounts SET rev = rev + ?2, pushes = CASE WHEN day = ?3 THEN pushes + 1 ELSE 1 END, day = ?3 WHERE id = ?1').bind(me.account, n, day));
  push.forEach((p, i) => {
    batch.push(db.prepare(
      'INSERT INTO records (account, c, id, m, d, v, rev) VALUES (?1, ?2, ?3, ?4, ?5, ?6, (SELECT rev FROM accounts WHERE id = ?1) - ?7) ' +
      'ON CONFLICT (account, c, id) DO UPDATE SET m = excluded.m, d = excluded.d, v = excluded.v, rev = excluded.rev WHERE excluded.m > records.m'
    ).bind(me.account, p.c, p.id, p.m, p.d, p.v, n - 1 - i));
  });
  // Everything after "since", except what this request just stored (the newest n numbers).
  batch.push(db.prepare(
    'SELECT c, id, m, d, v, rev FROM records WHERE account = ?1 AND rev > ?2 AND rev <= (SELECT rev FROM accounts WHERE id = ?1) - ?3 ORDER BY rev LIMIT ?4'
  ).bind(me.account, since, n, PAGE + 1));
  batch.push(db.prepare('SELECT rev FROM accounts WHERE id = ?').bind(me.account));
  const res = await db.batch(batch);

  const rows = res[res.length - 2].results || [], rev = res[res.length - 1].results[0].rev;
  const more = rows.length > PAGE;
  const pull = rows.slice(0, PAGE).map(rowOut);
  // A change that lost to a newer one: hand back the newer one, whenever it was stored.
  const lost = push.filter((p, i) => !(res[1 + i].meta && res[1 + i].meta.changes));
  if (lost.length) {
    const have = new Set(pull.map((r) => r.c + '/' + r.id));
    const got = await db.batch(lost.filter((p) => !have.has(p.c + '/' + p.id)).map((p) => db.prepare('SELECT c, id, m, d, v, rev FROM records WHERE account = ? AND c = ? AND id = ?').bind(me.account, p.c, p.id)));
    got.forEach((g) => { if (g.results && g.results[0]) pull.push(rowOut(g.results[0])); });
  }
  return json({ pull, next: more ? rows[PAGE - 1].rev : rev, more }, 200, origin);
}

async function account(request, env, origin, me) {
  if (request.method === 'GET') {
    const have = await env.DB.prepare('SELECT COUNT(*) AS n, COALESCE(SUM(LENGTH(v)), 0) AS b FROM records WHERE account = ? AND d = 0').bind(me.account).first();
    const dev = await env.DB.prepare('SELECT COUNT(*) AS n FROM sessions WHERE account = ?').bind(me.account).first();
    return json({ created: me.created, records: have.n, bytes: have.b, computers: dev.n }, 200, origin);
  }
  if (request.method === 'DELETE') {
    await env.DB.batch([
      env.DB.prepare('DELETE FROM records WHERE account = ?').bind(me.account),
      env.DB.prepare('DELETE FROM sessions WHERE account = ?').bind(me.account),
      env.DB.prepare('DELETE FROM accounts WHERE id = ?').bind(me.account)
    ]);
    return json({ ok: true }, 200, origin);
  }
  return json({ error: 'method not allowed' }, 405, origin);
}

// The answer for an account path, or null when the path isn't one.
export async function accountRoute(request, env, parts, origin) {
  const path = parts.join('/');
  if (path === 'auth/microsoft') {
    if (request.method !== 'POST') return json({ error: 'method not allowed' }, 405, origin);
    return signIn(request, env, origin);
  }
  if (path !== 'sync' && path !== 'account' && path !== 'signout') return null;
  const me = await sessionOf(request, env);
  if (!me) return json({ error: 'signed out' }, 401, origin);
  if (path === 'sync') {
    if (request.method !== 'POST') return json({ error: 'method not allowed' }, 405, origin);
    return sync(request, env, origin, me);
  }
  if (path === 'signout') {
    if (request.method !== 'POST') return json({ error: 'method not allowed' }, 405, origin);
    await env.DB.prepare('DELETE FROM sessions WHERE hash = ?').bind(me.hash).run();
    return json({ ok: true }, 200, origin);
  }
  return account(request, env, origin, me);
}

// For tests: forget Microsoft's keys.
export function forgetKeys() { msKeys = null; }
