// Checks the Cloudflare Worker (worker/src/index.js) without deploying it.
// Its D1 database is stood in for by Node's built-in SQLite (the same SQL
// runs), and its KV namespace by a Map.
// Run from the repository folder: node tests/worker-test.mjs
import assert from 'node:assert';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { DatabaseSync } from 'node:sqlite';
import worker from '../worker/src/index.js';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const PANEL = 'https://labelsheetkeeper.github.io';
let n = 0;
const ok = (c, m) => { assert(c, m); n++; };

class Stmt {
  constructor(db, sql, args) { this.db = db; this.sql = sql; this.args = args || []; }
  bind(...args) { return new Stmt(this.db, this.sql, args); }
  runNow() {
    const s = this.db.prepare(this.sql);
    if (/\bRETURNING\b/i.test(this.sql)) return { success: true, results: s.all(...this.args).map((r) => ({ ...r })) };
    const r = s.run(...this.args);
    return { success: true, meta: { changes: Number(r.changes) } };
  }
  async run() { return this.runNow(); }
  async first() { const r = this.db.prepare(this.sql).get(...this.args); return r === undefined ? null : { ...r }; }
  async all() { return { results: this.db.prepare(this.sql).all(...this.args).map((r) => ({ ...r })) }; }
}
class FakeD1 {
  constructor() { this.db = new DatabaseSync(':memory:'); }
  prepare(sql) { return new Stmt(this.db, sql); }
  async batch(stmts) {
    this.db.exec('BEGIN');
    try { const out = stmts.map((s) => s.runNow()); this.db.exec('COMMIT'); return out; } catch (e) { this.db.exec('ROLLBACK'); throw e; }
  }
  rows(sql, ...args) { return this.db.prepare(sql).all(...args).map((r) => ({ ...r })); }
}
class FakeKV {
  constructor() { this.m = new Map(); this.hidden = new Set(); }
  async getWithMetadata(key) {
    const v = this.hidden.has(key) ? null : this.m.get(key);
    return v ? { value: v.value, metadata: v.metadata || null } : { value: null, metadata: null };
  }
  async put(key, value, opts = {}) { this.m.set(key, { value, metadata: opts.metadata, ttl: opts.expirationTtl }); }
  async delete(key) { this.m.delete(key); }
  keys(prefix) { return [...this.m.keys()].filter((k) => k.startsWith(prefix)); }
}
const db = new FakeD1(), kv = new FakeKV();
const env = { DB: db, LABELS: kv };
const call = (p, init = {}) => worker.fetch(new Request('https://api.example.workers.dev' + p, init), env);

// ---- health and browser access ----
let r = await call('/health', { headers: { Origin: PANEL } });
let body = await r.json();
ok(r.status === 200 && body.ok === true && body.service === 'label-sheet-keeper', 'health answers ' + JSON.stringify(body));
ok(r.headers.get('Access-Control-Allow-Origin') === PANEL, 'the panel may call it');
ok(r.headers.get('Cache-Control') === 'no-store', 'answers are not cached');
r = await call('/health', { headers: { Origin: 'https://example.com' } });
ok(r.status === 200 && r.headers.get('Access-Control-Allow-Origin') === null, 'other websites get no browser access');
r = await call('/box/x', { method: 'OPTIONS', headers: { Origin: PANEL, 'Access-Control-Request-Method': 'DELETE' } });
ok(r.status === 204 && /DELETE/.test(r.headers.get('Access-Control-Allow-Methods')) && /Authorization/.test(r.headers.get('Access-Control-Allow-Headers')), 'browser pre-check allows the panel’s requests');
ok((await call('/health', { method: 'POST' })).status === 405, 'health is read-only');
ok((await call('/nothing-here')).status === 404, 'unknown paths are not found');

// ---- linking a phone ----
r = await call('/pair', { method: 'POST', headers: { Origin: 'https://example.com' } });
ok(r.status === 403, 'only the panel can link a phone');
r = await call('/pair', { method: 'POST', headers: { Origin: PANEL } });
const link = await r.json();
ok(r.status === 201 && /^[A-Za-z0-9_-]{22}$/.test(link.box) && /^[A-Za-z0-9_-]{32}$/.test(link.key), 'linking gives a mailbox and a read key ' + JSON.stringify(link));
ok(link.send === 'https://api.example.workers.dev/drop/' + link.box, 'and the phone’s sending address');
const rec = db.rows('SELECT key FROM boxes WHERE box = ?', link.box)[0];
ok(rec && rec.key !== link.key && rec.key.length === 64, 'the read key itself is never stored, only its hash');
const auth = { Authorization: 'Bearer ' + link.key, Origin: PANEL };

// ---- sending labels ----
const pdf = fs.readFileSync(path.join(HERE, 'fixtures/sample-return-label.pdf'));
const png = fs.readFileSync(path.join(HERE, 'fixtures/mark-ship-good.png'));
r = await call('/drop/' + link.box, { method: 'POST', body: pdf, headers: { 'Content-Type': 'application/octet-stream', 'X-File-Name': encodeURIComponent('Amazon return.pdf') } });
ok(r.status === 201 && (await r.json()).ok === true, 'a PDF can be sent');
r = await call('/drop/' + link.box + '?reply=text', { method: 'POST', body: png });
ok(r.status === 201 && (await r.text()) === 'Sent to Label Sheet Keeper.', 'the Shortcut gets a sentence back');
r = await call('/drop/' + link.box + '?reply=text', { method: 'POST', body: 'hello, not a label' });
ok(r.status === 415 && /PDF or a picture/.test(await r.text()), 'other files are refused');
r = await call('/drop/' + link.box, { method: 'POST', body: new Uint8Array(0) });
ok(r.status === 400, 'an empty file is refused');
r = await call('/drop/' + link.box, { method: 'POST', body: new Uint8Array(10 * 1024 * 1024 + 1) });
ok(r.status === 413, 'a file over 10 MB is refused');
r = await call('/drop/AAAAAAAAAAAAAAAAAAAAAA?reply=text', { method: 'POST', body: pdf });
ok(r.status === 404 && /Link your phone again/.test(await r.text()), 'an unknown sending link says how to fix it');
ok(kv.keys('file:' + link.box).length === 2 && kv.keys('file:' + link.box).every((k) => kv.m.get(k).ttl === 7 * 24 * 3600), 'files are kept 7 days at most');
ok(db.rows('SELECT sent FROM boxes WHERE box = ?', link.box)[0].sent === 2, 'sends are counted for the day');

// ---- the panel picks them up ----
r = await call('/box/' + link.box, { headers: { Origin: PANEL } });
ok(r.status === 401, 'listing needs the read key');
r = await call('/box/' + link.box, { headers: { Authorization: 'Bearer ' + 'x'.repeat(32), Origin: PANEL } });
ok(r.status === 401, 'a wrong read key is refused');
r = await call('/box/' + link.box, { headers: auth });
let list = (await r.json()).items;
ok(r.status === 200 && list.length === 2, 'both labels are waiting ' + JSON.stringify(list));
ok(list[0].name === 'Amazon return.pdf' && list[0].type === 'application/pdf' && list[1].type === 'image/png' && list[1].name === 'Label from your phone.png', 'names and real file types are kept');
r = await call('/box/' + link.box + '/' + list[0].id, { headers: auth });
const got = new Uint8Array(await r.arrayBuffer());
ok(r.status === 200 && r.headers.get('Content-Type') === 'application/pdf' && got.length === pdf.length && got[0] === pdf[0], 'the panel gets the file back');
r = await call('/box/' + link.box + '/' + list[0].id, { headers: { Origin: PANEL } });
ok(r.status === 401, 'files need the read key too');
r = await call('/box/' + link.box + '/' + list[0].id, { method: 'DELETE', headers: auth });
ok(r.status === 200 && (await r.json()).waiting === 1, 'a picked-up label is deleted');
ok(!kv.m.has('file:' + link.box + ':' + list[0].id) && db.rows('SELECT id FROM labels WHERE id = ?', list[0].id).length === 0, 'its file and its place in the list are gone');
r = await call('/box/' + link.box + '/' + list[0].id, { headers: auth });
ok(r.status === 404, 'and it can’t be fetched again');

// ---- a label whose file hasn't reached this location yet ----
const late = list[1].id;
kv.hidden.add('file:' + link.box + ':' + late);
r = await call('/box/' + link.box + '/' + late, { headers: auth });
ok(r.status === 409 && /still on its way/.test((await r.json()).error), 'a label just sent but not here yet is “on its way”, not gone');
ok(db.rows('SELECT id FROM labels WHERE id = ?', late).length === 1, 'and stays on the list');
db.db.prepare('UPDATE labels SET at = ? WHERE id = ?').run(Date.now() - 6 * 60 * 1000, late);
r = await call('/box/' + link.box + '/' + late, { headers: auth });
ok(r.status === 404 && db.rows('SELECT id FROM labels WHERE id = ?', late).length === 0, 'one whose file never came is taken off the list');
kv.hidden.clear();

// ---- limits ----
db.db.prepare('INSERT INTO labels (box, id, type, name, size, at) VALUES (?, ?, ?, ?, ?, ?)').run(link.box, 'oldoldoldoldoldo', 'image/png', 'old', 1, Date.now() - 8 * 24 * 3600 * 1000);
list = (await (await call('/box/' + link.box, { headers: auth })).json()).items;
ok(list.length === 0 && db.rows('SELECT id FROM labels WHERE id = ?', 'oldoldoldoldoldo').length === 0, 'labels older than 7 days aren’t listed, and are cleared');
for (let i = 0; i < 20; i++) await call('/drop/' + link.box, { method: 'POST', body: png });
r = await call('/drop/' + link.box + '?reply=text', { method: 'POST', body: png });
ok(r.status === 429 && /Too many labels are waiting/.test(await r.text()), 'at most 20 labels wait at once');
db.db.prepare('DELETE FROM labels WHERE box = ?').run(link.box);
kv.keys('file:' + link.box).forEach((k) => kv.m.delete(k));
db.db.prepare('UPDATE boxes SET day = ?, sent = 60 WHERE box = ?').run(new Date().toISOString().slice(0, 10), link.box);
r = await call('/drop/' + link.box + '?reply=text', { method: 'POST', body: png });
ok(r.status === 429 && /most labels for one day/.test(await r.text()), 'at most 60 labels a day');
db.db.prepare('UPDATE boxes SET day = ? WHERE box = ?').run('2000-01-01', link.box);
r = await call('/drop/' + link.box, { method: 'POST', body: png });
ok(r.status === 201 && db.rows('SELECT sent FROM boxes WHERE box = ?', link.box)[0].sent === 1, 'the count starts over the next day');

// ---- unlinking ----
r = await call('/box/' + link.box, { method: 'DELETE', headers: auth });
ok(r.status === 200 && db.rows('SELECT box FROM boxes WHERE box = ?', link.box).length === 0 && db.rows('SELECT id FROM labels WHERE box = ?', link.box).length === 0, 'unlinking deletes the mailbox and its list');
ok(kv.keys('file:' + link.box).length === 0, 'and every file');
r = await call('/drop/' + link.box, { method: 'POST', body: pdf });
ok(r.status === 404 && db.rows('SELECT box FROM boxes WHERE box = ?', link.box).length === 0, 'the old sending link stops working at once, and doesn’t bring the mailbox back');
r = await call('/box/' + link.box, { headers: auth });
ok(r.status === 401, 'and the old read key no longer works');

// ---- address suggestions for everyone ----
const realFetch = globalThis.fetch;
let geoCalls = [], geoReply = () => new Response(JSON.stringify({ results: [
  { housenumber: '8', street: 'Maple Way', city: 'Boston', state_code: 'MA', postcode: '02108', country_code: 'us', address_line1: '8 Maple Way', address_line2: 'Boston, MA 02108', lat: 42.3, lon: -71.1, place_id: 'x' }] }), { status: 200 });
globalThis.fetch = async (u) => { geoCalls.push(String(u)); return geoReply(); };
const ask = (text, headers = { Origin: PANEL, 'CF-Connecting-IP': '203.0.113.7' }) => call('/suggest?text=' + encodeURIComponent(text), { headers });
r = await ask('8 Maple');
ok(r.status === 503 && geoCalls.length === 0, 'without a Geoapify key set up, suggestions say so and nothing is asked');
env.GEOAPIFY_KEY = 'secret-geo-key';
r = await ask('8 Maple', { Origin: 'https://example.com' });
ok(r.status === 403 && geoCalls.length === 0, 'only the panel can ask for suggestions');
r = await ask('8');
ok(r.status === 400, 'a too-short question is refused');
r = await ask('8   Maple ');
body = await r.json();
ok(r.status === 200 && body.results.length === 1 && body.results[0].street === 'Maple Way' && body.results[0].postcode === '02108', 'suggestions come back ' + JSON.stringify(body));
ok(body.results[0].lat === undefined && body.results[0].place_id === undefined && !JSON.stringify(body).includes('secret-geo-key'), 'with only the address parts, and never the key');
ok(geoCalls.length === 1 && geoCalls[0].includes('text=8%20Maple&') && geoCalls[0].includes('filter=countrycode:us,ca') && geoCalls[0].includes('apiKey=secret-geo-key'), 'Geoapify is asked for US and Canada with the shared key ' + geoCalls[0]);
ok(r.headers.get('Access-Control-Allow-Origin') === PANEL, 'the table window may read the answer');
const used = db.rows('SELECT who, n FROM usage');
ok(used.length === 2 && used.every((u) => u.n === 1) && used.some((u) => u.who === '*') && !used.some((u) => u.who.includes('203.0.113.7')), 'each question is counted, by a hash of the address, not the address itself ' + JSON.stringify(used));
db.db.prepare("UPDATE usage SET n = 300 WHERE who <> '*'").run();
r = await ask('9 Maple');
ok(r.status === 429 && geoCalls.length === 1, 'one internet address gets at most 300 a day');
r = await ask('9 Maple', { Origin: PANEL, 'CF-Connecting-IP': '198.51.100.2' });
ok(r.status === 200, 'others still get theirs');
db.db.prepare("UPDATE usage SET n = 2700 WHERE who = '*'").run();
r = await ask('10 Maple', { Origin: PANEL, 'CF-Connecting-IP': '198.51.100.9' });
ok(r.status === 429, 'everyone together gets at most 2,700 a day, under Geoapify’s free 3,000');
db.db.prepare('DELETE FROM usage').run();
geoReply = () => new Response('{"message":"Invalid apiKey"}', { status: 401 });
r = await ask('11 Maple');
ok(r.status === 503, 'a key Geoapify refuses counts as not set up');
geoReply = () => new Response('{}', { status: 429 });
r = await ask('12 Maple');
ok(r.status === 429, 'Geoapify’s own limit is passed on');
geoReply = () => { throw new Error('offline'); };
r = await ask('13 Maple');
ok(r.status === 503 && /busy/.test((await r.json()).error), 'no connection to Geoapify says busy');
globalThis.fetch = realFetch;

console.log('all', n, 'worker checks passed');
