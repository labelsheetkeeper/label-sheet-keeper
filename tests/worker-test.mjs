// Checks the Cloudflare Worker (worker/src/index.js) without deploying it.
// Its D1 database is stood in for by Node's built-in SQLite (the same SQL
// runs), and its KV namespace by a Map.
// Run from the repository folder: node tests/worker-test.mjs
import assert from 'node:assert';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { FakeD1, FakeKV } from './fakes.mjs';
import { createHash } from 'node:crypto';
import worker from '../worker/src/index.js';
import { forgetKeys } from '../worker/src/account.js';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const PANEL = 'https://labelsheetkeeper.github.io';
let n = 0;
const ok = (c, m) => { assert(c, m); n++; };

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
const okReply = geoReply;
geoReply = () => new Response('{"message":"Invalid apiKey"}', { status: 401 });
r = await ask('11 Maple');
ok(r.status === 503, 'a key Geoapify refuses counts as not set up');
geoReply = () => new Response('{}', { status: 429 });
r = await ask('12 Maple');
ok(r.status === 429, 'Geoapify’s own limit is passed on');
geoReply = () => { throw new Error('offline'); };
r = await ask('13 Maple');
ok(r.status === 503 && /busy/.test((await r.json()).error), 'no connection to Geoapify says busy');
geoReply = okReply;
db.db.prepare("UPDATE usage SET n = 300 WHERE who <> '*'").run();
let before = geoCalls.length;
r = await ask('9 Maple');
ok(r.status === 429, 'one internet address gets at most 300 a day');
r = await ask('9 Maple again');
ok(r.status === 429 && geoCalls.length === before + 1, 'after that, Geoapify isn’t asked for it again today ' + (geoCalls.length - before));
r = await ask('9 Maple', { Origin: PANEL, 'CF-Connecting-IP': '198.51.100.2' });
ok(r.status === 200, 'others still get theirs');
db.db.prepare("UPDATE usage SET n = 2700 WHERE who = '*'").run();
r = await ask('10 Maple', { Origin: PANEL, 'CF-Connecting-IP': '198.51.100.9' });
ok(r.status === 429, 'everyone together gets at most 2,700 a day, under Geoapify’s free 3,000');
before = geoCalls.length;
r = await ask('10 Maple', { Origin: PANEL, 'CF-Connecting-IP': '198.51.100.10' });
ok(r.status === 429 && geoCalls.length === before, 'and then nobody’s question goes to Geoapify until tomorrow');
globalThis.fetch = realFetch;

// ---- signing in with Microsoft ----
// Stand-in for Microsoft: our own signing key, published the way Microsoft publishes its keys.
const CLIENT = '11111111-2222-3333-4444-555555555555';
const MSA = '9188040d-6c67-4c5b-b112-36a304b66dad';   // Microsoft's "tenant" for personal accounts
const WORK = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';
const mkKey = () => crypto.subtle.generateKey({ name: 'RSASSA-PKCS1-v1_5', modulusLength: 2048, publicExponent: new Uint8Array([1, 0, 1]), hash: 'SHA-256' }, true, ['sign', 'verify']);
const k1 = await mkKey(), k2 = await mkKey(), stranger = await mkKey();
const pub = async (k, kid, issuer) => { const j = await crypto.subtle.exportKey('jwk', k.publicKey); return { kid, kty: 'RSA', use: 'sig', n: j.n, e: j.e, issuer }; };
const msKeys = { keys: [
  await pub(k1, 'any', 'https://login.microsoftonline.com/{tenantid}/v2.0'),
  await pub(k2, 'personal', 'https://login.microsoftonline.com/' + MSA + '/v2.0')
] };
let keyFetches = 0, keysDown = false;
globalThis.fetch = async (u) => {
  if (String(u) !== 'https://login.microsoftonline.com/common/discovery/v2.0/keys') throw new Error('unexpected fetch ' + u);
  keyFetches++;
  return keysDown ? new Response('down', { status: 500 }) : new Response(JSON.stringify(msKeys));
};
const b64u = (x) => Buffer.from(x).toString('base64url');
const nowS = () => Math.floor(Date.now() / 1000);
async function idToken(over = {}, o = {}) {
  const tid = over.tid || MSA;
  const claims = { aud: CLIENT, iss: 'https://login.microsoftonline.com/' + tid + '/v2.0', tid, oid: '00000000-0000-0000-1234-56789abcdef0', sub: 'pairwise', name: 'Pat Example', preferred_username: 'pat@example.com', iat: nowS() - 10, nbf: nowS() - 10, exp: nowS() + 3600, ...over };
  const head = { typ: 'JWT', alg: o.alg || 'RS256', kid: o.kid || (tid === MSA ? 'personal' : 'any') };
  const data = b64u(JSON.stringify(head)) + '.' + b64u(JSON.stringify(claims));
  const sig = await crypto.subtle.sign('RSASSA-PKCS1-v1_5', (o.key || (head.kid === 'personal' ? k2 : k1)).privateKey, new TextEncoder().encode(data));
  return data + '.' + b64u(sig);
}
const signIn = (token, headers = { Origin: PANEL, 'CF-Connecting-IP': '203.0.113.50' }) => call('/auth/microsoft', { method: 'POST', headers, body: JSON.stringify({ id_token: token }) });

r = await signIn(await idToken());
ok(r.status === 503 && keyFetches === 0, 'until the app is registered with Microsoft, sign-in says it isn’t set up');
env.MS_CLIENT_ID = CLIENT;
r = await signIn(await idToken(), { Origin: 'https://example.com' });
ok(r.status === 403, 'only the panel can sign in');
ok((await call('/auth/microsoft', { headers: { Origin: PANEL } })).status === 405, 'signing in is a POST');
r = await signIn(await idToken());
body = await r.json();
ok(r.status === 201 && /^[A-Za-z0-9_-]{43}$/.test(body.session) && body.isNew === true, 'a genuine Microsoft sign-in gets a session ' + r.status);
const sessA = body.session;
const dump = JSON.stringify([db.rows('SELECT * FROM accounts'), db.rows('SELECT * FROM sessions')]);
ok(!/pat@example|Pat Example|56789abcdef0|9188040d/.test(dump) && !dump.includes(sessA), 'no name, email, Microsoft ID or session token is stored, only hashes');
ok(db.rows('SELECT id FROM accounts').length === 1 && db.rows('SELECT id FROM accounts')[0].id.length === 40, 'the account is known by a hash');
r = await signIn(await idToken());
body = await r.json();
const sessB = body.session;
ok(r.status === 201 && body.isNew === false && sessB !== sessA && db.rows('SELECT id FROM accounts').length === 1 && db.rows('SELECT hash FROM sessions').length === 2, 'the same person on another computer gets the same account, and a session of their own');
ok(keyFetches === 1, 'Microsoft’s keys are fetched once and remembered ' + keyFetches);

const bad = async (label, token) => { const x = await signIn(token); ok(x.status === 401, label + ' (' + x.status + ')'); };
await bad('a token made for another app is refused', await idToken({ aud: '99999999-2222-3333-4444-555555555555' }));
await bad('an expired token is refused', await idToken({ exp: nowS() - 600 }));
await bad('a token not valid yet is refused', await idToken({ nbf: nowS() + 3600 }));
await bad('a token signed by someone else is refused', await idToken({}, { key: stranger }));
await bad('an unsigned token is refused', (await idToken({}, { alg: 'none' })).replace(/\.[^.]*$/, '.'));
await bad('an issuer that doesn’t match its organization is refused', await idToken({ iss: 'https://login.microsoftonline.com/' + WORK + '/v2.0' }));
await bad('a key for personal accounts can’t vouch for an organization', await idToken({ tid: WORK }, { kid: 'personal', key: k2 }));
await bad('a token with no person in it is refused', await idToken({ oid: undefined }));
await bad('a changed token is refused', (await idToken()).replace(/\.([^.]+)\./, (m, p) => '.' + b64u(JSON.stringify({ ...JSON.parse(Buffer.from(p, 'base64url')), oid: '00000000-0000-0000-9999-999999999999' })) + '.'));
await bad('nonsense is refused', 'not.a.token');
await bad('nothing is refused', '');
ok(db.rows('SELECT id FROM accounts').length === 1, 'and none of those made an account');
r = await signIn(await idToken({ tid: WORK, oid: '00000000-0000-0000-1234-56789abcdef0' }));
const sessWork = (await r.json()).session;
ok(r.status === 201 && db.rows('SELECT id FROM accounts').length === 2, 'the same ID in a different organization is a different person');
forgetKeys(); keysDown = true;
r = await signIn(await idToken());
ok(r.status === 503 && /Microsoft couldn’t be reached/.test((await r.json()).error), 'if Microsoft’s keys can’t be fetched, it says so');
keysDown = false;
db.db.prepare("UPDATE usage SET n = 100 WHERE who LIKE 'auth:%'").run();
r = await signIn(await idToken());
ok(r.status === 429, 'one internet address gets at most 100 sign-ins a day');
r = await signIn(await idToken(), { Origin: PANEL, 'CF-Connecting-IP': '203.0.113.51' });
ok(r.status === 201, 'others can still sign in');
const sessC = (await r.json()).session;

// ---- syncing ----
const acctOf = (sess) => db.rows('SELECT account FROM sessions WHERE hash = ?', createHash('sha256').update(sess).digest('hex'))[0].account;
const acctA = acctOf(sessA);
const sync = (sess, since, push = []) => call('/sync', { method: 'POST', headers: { Origin: PANEL, Authorization: 'Bearer ' + sess }, body: JSON.stringify({ since, push }) });
const item = (c, id, m, v, d) => ({ c, id, m, d: d ? 1 : 0, v: d ? null : JSON.stringify(v) });
ok((await call('/sync', { method: 'POST', headers: { Origin: PANEL }, body: '{}' })).status === 401, 'syncing needs a session');
ok((await sync('x'.repeat(43), 0)).status === 401, 'a made-up session is refused');
const t0 = Date.now() - 100000;
r = await sync(sessA, 0, [item('book', 'c1', t0, { name: 'Ann' }), item('book', 'c2', t0, { name: 'Bob' }), item('set', 'homeCountry', t0, 'US')]);
body = await r.json();
ok(r.status === 200 && body.pull.length === 0 && body.next === 3 && body.more === false, 'the first computer sends its changes and gets nothing back ' + JSON.stringify(body));
r = await sync(sessB, 0);
body = await r.json();
ok(body.pull.length === 3 && body.next === 3 && body.pull.map((x) => x.id).join() === 'c1,c2,homeCountry' && body.pull[0].v === '{"name":"Ann"}', 'the second computer gets them ' + JSON.stringify(body.pull.map((x) => x.id)));
r = await sync(sessB, 3);
ok((await r.json()).pull.length === 0, 'and nothing twice');
r = await sync(sessB, 3, [item('book', 'c1', t0 + 500, { name: 'Ann Lee' })]);
body = await r.json();
ok(body.pull.length === 0 && body.next === 4, 'a change from the second computer is stored');
r = await sync(sessA, 3);
body = await r.json();
ok(body.pull.length === 1 && body.pull[0].v === '{"name":"Ann Lee"}' && body.next === 4, 'and reaches the first');
r = await sync(sessA, 4, [item('book', 'c1', t0 + 100, { name: 'Ann Old' })]);
body = await r.json();
ok(body.pull.length === 1 && body.pull[0].v === '{"name":"Ann Lee"}' && db.rows("SELECT v FROM records WHERE id = 'c1'")[0].v === '{"name":"Ann Lee"}', 'an older change loses, and the newer one is handed back');
r = await sync(sessA, body.next, [item('book', 'c1', t0 + 500, { name: 'Ann Lee' })]);
body = await r.json();
ok(body.pull.length === 1 && db.rows("SELECT rev FROM records WHERE id = 'c1'")[0].rev === 4, 'sending the same change again changes nothing');
r = await sync(sessA, body.next, [item('book', 'c2', t0 + 900, null, true)]);
body = await r.json();
r = await sync(sessB, 4);
body = await r.json();
ok(body.pull.length === 1 && body.pull[0].id === 'c2' && body.pull[0].d === 1 && body.pull[0].v === null, 'a deleted address is passed on as deleted');
r = await sync(sessB, body.next, [item('book', 'c2', t0 + 200, { name: 'Bob again' })]);
ok(db.rows("SELECT d FROM records WHERE id = 'c2'")[0].d === 1, 'and an older edit doesn’t bring it back');
r = await sync(sessB, 0, [item('set', 'clock', Date.now() + 9e9, 'ahead')]);
ok(db.rows("SELECT m FROM records WHERE id = 'clock'")[0].m < Date.now() + 6 * 60 * 1000, 'a clock set far ahead can’t make a change win forever');
ok((await (await sync(sessWork, 0)).json()).pull.length === 0, 'another person’s account sees none of it');

const refuse = async (label, push, status = 400) => { const x = await sync(sessA, 0, push); ok(x.status === status, label + ' (' + x.status + ')'); };
await refuse('an unknown kind of record is refused', [item('other', 'x', t0, 1)]);
await refuse('an odd id is refused', [item('book', 'has space', t0, 1)]);
await refuse('a record with no time is refused', [{ c: 'book', id: 'x', v: '1' }]);
await refuse('a record with no content is refused', [{ c: 'book', id: 'x', m: t0 }]);
await refuse('the same record twice in one request is refused', [item('book', 'x', t0, 1), item('book', 'x', t0 + 1, 2)]);
await refuse('a record over 300 KB is refused', [{ c: 'set', id: 'big', m: t0, v: 'x'.repeat(300 * 1024 + 1) }]);
await refuse('more than 300 records at once are refused', Array.from({ length: 301 }, (_, i) => item('book', 'm' + i, t0, i)));
ok((await call('/sync', { method: 'POST', headers: { Origin: PANEL, Authorization: 'Bearer ' + sessA }, body: 'not json' })).status === 400, 'a request that isn’t JSON is refused');
ok((await call('/sync', { headers: { Origin: PANEL, Authorization: 'Bearer ' + sessA } })).status === 405, 'syncing is a POST');

// Many records come back a page at a time, each exactly once.
for (let b = 0; b < 3; b++) await sync(sessA, 0, Array.from({ length: 250 }, (_, i) => item('book', 'p' + b + '_' + i, t0 + i, { n: i })));
let got1 = new Map(), since = 0, pages = 0, last;
do { last = await (await sync(sessC, since)).json(); last.pull.forEach((x) => got1.set(x.c + '/' + x.id, (got1.get(x.c + '/' + x.id) || 0) + 1)); since = last.next; pages++; } while (last.more && pages < 20);
const total = db.rows('SELECT COUNT(*) AS n FROM records WHERE account = ?', acctA)[0].n;
ok(pages === 3 && got1.size === total && [...got1.values()].every((v) => v === 1) && total === 754, 'a new computer gets everything in pages, each record once ' + JSON.stringify([pages, got1.size, total]));
ok((await (await sync(sessC, since)).json()).pull.length === 0, 'and then it’s up to date');

r = await call('/account', { headers: { Origin: PANEL, Authorization: 'Bearer ' + sessA } });
body = await r.json();
ok(r.status === 200 && body.records === 753 && body.computers === 3 && body.bytes > 0, 'the account says how much it holds ' + JSON.stringify(body));
db.db.prepare('UPDATE accounts SET pushes = 3000, day = ? WHERE id = ?').run(new Date().toISOString().slice(0, 10), acctA);
r = await sync(sessA, 0, [item('set', 'x', Date.now(), 1)]);
ok(r.status === 429, 'at most 3,000 requests with changes a day');
ok((await sync(sessA, since)).status === 200, 'checking for changes still works');
db.db.prepare("UPDATE accounts SET day = '2000-01-01' WHERE id = ?").run(acctA);
ok((await sync(sessA, 0, [item('set', 'x', Date.now(), 1)])).status === 200 && db.rows('SELECT pushes FROM accounts WHERE id = ?', acctA)[0].pushes === 1, 'the count starts over the next day');
const fill = db.db.prepare("INSERT INTO records (account, c, id, m, d, v, rev) VALUES (?, 'book', ?, 1, 0, '1', 0)");
db.db.exec('BEGIN'); for (let i = 0; i < 19300; i++) fill.run(acctA, 'f' + i); db.db.exec('COMMIT');
r = await sync(sessA, 0, [item('book', 'one-too-many', Date.now(), 1)]);
ok(r.status === 413 && /account is full/.test((await r.json()).error), 'an account holds at most 20,000 records');
ok((await sync(sessA, 0, [item('book', 'f1', Date.now(), null, true)])).status === 200, 'deleting still works when it’s full');
db.db.prepare("DELETE FROM records WHERE id LIKE 'f%'").run();

// ---- signing out and deleting the account ----
r = await call('/signout', { method: 'POST', headers: { Origin: PANEL, Authorization: 'Bearer ' + sessC } });
ok(r.status === 200 && (await sync(sessC, 0)).status === 401, 'signing out ends that computer’s session');
ok((await sync(sessA, 0)).status === 200, 'and not the others');
db.db.prepare('UPDATE sessions SET seen = ? WHERE hash = ?').run(Date.now() - 401 * 24 * 3600 * 1000, createHash('sha256').update(sessB).digest('hex'));
ok((await sync(sessB, 0)).status === 401 && (await sync(sessA, 0)).status === 200, 'a session unused for over 400 days ends');
const live = sessA;
r = await call('/account', { method: 'DELETE', headers: { Origin: PANEL, Authorization: 'Bearer ' + live } });
ok(r.status === 200 && db.rows('SELECT id FROM accounts WHERE id = ?', acctA).length === 0 && db.rows('SELECT id FROM records WHERE account = ?', acctA).length === 0 && db.rows('SELECT hash FROM sessions WHERE account = ?', acctA).length === 0, 'deleting the account removes it, its records and its sessions');
ok((await sync(live, 0)).status === 401, 'and its sessions stop working');
ok(db.rows('SELECT id FROM accounts').length === 1 && (await sync(sessWork, 0)).status === 200, 'other people’s accounts are untouched');
for (let i = 0; i < 31; i++) await signIn(await idToken({ tid: WORK }), { Origin: PANEL, 'CF-Connecting-IP': '203.0.113.' + (60 + i) });
ok(db.rows('SELECT hash FROM sessions').length === 30, 'one account keeps at most 30 computers signed in');
globalThis.fetch = realFetch;

console.log('all', n, 'worker checks passed');
