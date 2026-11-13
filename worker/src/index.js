// Label Sheet Keeper server, a Cloudflare Worker at api.<subdomain>.workers.dev.
//
// Phone mailbox. The panel links a phone with POST /pair and gets a mailbox
// ("box") and a read key. The phone only ever knows the box, so it can send
// labels (POST /drop/<box>) but can't read them. The panel lists, fetches and
// deletes them with the read key. Labels are kept 7 days at most.
//
//   GET    /health               status
//   POST   /pair                 new mailbox → { box, key, send }        (panel only)
//   POST   /drop/<box>           send a label (raw PDF or picture body)  (phone)
//   GET    /box/<box>            labels waiting                          (read key)
//   GET    /box/<box>/<id>       one label's file                        (read key)
//   DELETE /box/<box>/<id>       picked up                               (read key)
//   DELETE /box/<box>            unlink: the mailbox and its labels      (read key)
//
// Address suggestions for everyone: the Worker asks Geoapify with the
// GEOAPIFY_KEY secret, so nobody needs a key of their own.
//
//   GET    /suggest?text=…       up to 5 US and Canadian addresses       (panel only)
//
// Accounts and sync (see account.js): sign in with Microsoft, then keep the
// address book and settings the same on every computer.
//
//   POST   /auth/microsoft       a Microsoft ID token → { session }      (panel only)
//   POST   /sync                 send changes, get the others'           (session)
//   GET    /account              how much is stored                      (session)
//   DELETE /account              the account and everything in it        (session)
//   POST   /signout              ends this session                       (session)
//
// Storage: mailboxes and the list of labels waiting are in the DB database
// (D1, always up to date). Each label's bytes are in the LABELS KV namespace
// as "file:<box>:<id>", written once and read once, and expire on their own.

import { ALLOWED_ORIGINS, corsHeaders, json, randomId, sha256, sameText, COUNT_SQL } from './shared.js';
import { ACCOUNT_SCHEMA, accountRoute } from './account.js';

const VERSION = '0.5.0';

const MAX_BYTES = 10 * 1024 * 1024;   // per label
const MAX_WAITING = 20;               // labels waiting in one mailbox
const MAX_PER_DAY = 60;               // labels sent to one mailbox per day
const KEEP_SECONDS = 7 * 24 * 3600;   // a label not picked up is deleted after this
const ARRIVING_MS = 5 * 60 * 1000;    // a new label's file can take a moment to reach every Cloudflare location

const SUGGEST_PER_DAY = 2700;         // for everyone together: Geoapify's free plan allows 3,000 a day
const SUGGEST_PER_PERSON = 300;       // for one internet address a day
const GEO_URL = 'https://api.geoapify.com/v1/geocode/autocomplete?format=json&lang=en&limit=5&filter=countrycode:us,ca&text=';
const GEO_KEEP = ['housenumber', 'street', 'address_line1', 'address_line2', 'city', 'state_code', 'postcode', 'country_code'];
const outOf = new Map();              // "*" or a who -> the day its suggestions ran out

const BOX_RE = /^[A-Za-z0-9_-]{22}$/;
const ID_RE = /^[A-Za-z0-9_-]{16}$/;

const EXT = { 'application/pdf': 'pdf', 'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/webp': 'webp', 'image/heic': 'heic' };

const SCHEMA = [
  'CREATE TABLE IF NOT EXISTS boxes (box TEXT PRIMARY KEY, key TEXT NOT NULL, created INTEGER NOT NULL, day TEXT NOT NULL DEFAULT \'\', sent INTEGER NOT NULL DEFAULT 0)',
  'CREATE TABLE IF NOT EXISTS labels (box TEXT NOT NULL, id TEXT NOT NULL, type TEXT NOT NULL, name TEXT NOT NULL, size INTEGER NOT NULL, at INTEGER NOT NULL, PRIMARY KEY (box, id))',
  // Suggestions asked today: who is "*" for everyone, or a one-day hash of an internet address.
  'CREATE TABLE IF NOT EXISTS usage (day TEXT NOT NULL, who TEXT NOT NULL, n INTEGER NOT NULL, PRIMARY KEY (day, who))',
  ...ACCOUNT_SCHEMA
];
let schemaReady = null;
function ready(env) {
  if (!schemaReady) schemaReady = env.DB.batch(SCHEMA.map((s) => env.DB.prepare(s))).catch((e) => { schemaReady = null; throw e; });
  return schemaReady;
}

// ---------- helpers ----------
// The iPhone Shortcut asks for ?reply=text so its notification shows a sentence.
function message(text, status, origin, asText) {
  if (!asText) return json(status < 300 ? { ok: true, message: text } : { error: text }, status, origin);
  return new Response(text, { status, headers: { 'Content-Type': 'text/plain; charset=utf-8', 'Cache-Control': 'no-store', ...corsHeaders(origin) } });
}

// What the file really is, from its first bytes (the phone's own label isn't trusted).
function sniff(b) {
  const ascii = (from, to) => String.fromCharCode(...b.subarray(from, to));
  if (b.length >= 8 && b[0] === 0x89 && ascii(1, 4) === 'PNG') return 'image/png';
  if (b.length >= 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return 'image/jpeg';
  if (b.length >= 6 && /^GIF8[79]a$/.test(ascii(0, 6))) return 'image/gif';
  if (b.length >= 12 && ascii(0, 4) === 'RIFF' && ascii(8, 12) === 'WEBP') return 'image/webp';
  if (b.length >= 12 && ascii(4, 8) === 'ftyp' && /^(heic|heix|hevc|hevx|mif1|msf1|heim|heis)$/.test(ascii(8, 12))) return 'image/heic';
  if (ascii(0, Math.min(b.length, 1024)).indexOf('%PDF-') >= 0) return 'application/pdf';
  return '';
}

function fileName(request, type) {
  let name = '';
  try { name = decodeURIComponent(request.headers.get('X-File-Name') || ''); } catch (e) { name = ''; }
  name = name.replace(/^.*[\\/]/, '').replace(/[\u0000-\u001f\u007f]/g, '').trim().slice(0, 80);
  return name || 'Label from your phone.' + EXT[type];
}

const fileKey = (box, id) => 'file:' + box + ':' + id;
const oldest = () => Date.now() - KEEP_SECONDS * 1000;

async function getBox(env, box) {
  if (!BOX_RE.test(box)) return null;
  return env.DB.prepare('SELECT box, key, day, sent FROM boxes WHERE box = ?').bind(box).first();
}

async function withKey(request, env, box) {
  const rec = await getBox(env, box);
  if (!rec) return null;
  const m = /^Bearer\s+([A-Za-z0-9_-]{20,})$/.exec(request.headers.get('Authorization') || '');
  if (!m) return null;
  return sameText(await sha256(m[1]), rec.key) ? rec : null;
}

// ---------- routes ----------
async function pair(request, env, url, origin) {
  if (!ALLOWED_ORIGINS.includes(origin)) return json({ error: 'Link your phone from the Label Sheet Keeper panel.' }, 403, origin);
  const box = randomId(16), key = randomId(24);
  await env.DB.prepare('INSERT INTO boxes (box, key, created) VALUES (?, ?, ?)').bind(box, await sha256(key), Date.now()).run();
  return json({ box, key, send: url.origin + '/drop/' + box }, 201, origin);
}

async function drop(request, env, box, origin, asText) {
  const rec = await getBox(env, box);
  if (!rec) return message('This sending link isn’t active. Link your phone again from the panel’s Settings.', 404, origin, asText);
  if (Number(request.headers.get('Content-Length') || 0) > MAX_BYTES) return message('That file is over 10 MB.', 413, origin, asText);
  const buf = await request.arrayBuffer();
  if (!buf.byteLength) return message('The file was empty.', 400, origin, asText);
  if (buf.byteLength > MAX_BYTES) return message('That file is over 10 MB.', 413, origin, asText);
  const type = sniff(new Uint8Array(buf, 0, Math.min(buf.byteLength, 1024)));
  if (!type) return message('Send a PDF or a picture (PNG, JPEG or HEIC).', 415, origin, asText);

  const waiting = await env.DB.prepare('SELECT COUNT(*) AS n FROM labels WHERE box = ? AND at > ?').bind(box, oldest()).first();
  if (waiting && waiting.n >= MAX_WAITING) return message('Too many labels are waiting. Open the panel to pick them up first.', 429, origin, asText);
  const today = new Date().toISOString().slice(0, 10);
  const sent = rec.day === today ? rec.sent : 0;
  if (sent >= MAX_PER_DAY) return message('That’s the most labels for one day. Try again tomorrow.', 429, origin, asText);

  const id = randomId(12), name = fileName(request, type);
  await env.LABELS.put(fileKey(box, id), buf, { expirationTtl: KEEP_SECONDS, metadata: { type, name } });
  await env.DB.batch([
    env.DB.prepare('INSERT INTO labels (box, id, type, name, size, at) VALUES (?, ?, ?, ?, ?, ?)').bind(box, id, type, name, buf.byteLength, Date.now()),
    env.DB.prepare('UPDATE boxes SET day = ?, sent = ? WHERE box = ?').bind(today, sent + 1, box)
  ]);
  return message('Sent to Label Sheet Keeper.', 201, origin, asText);
}

async function boxRoute(request, env, box, id, origin) {
  const rec = await withKey(request, env, box);
  if (!rec) return json({ error: 'not linked' }, 401, origin);

  if (!id && request.method === 'GET') {
    await env.DB.prepare('DELETE FROM labels WHERE box = ? AND at <= ?').bind(box, oldest()).run();
    const { results } = await env.DB.prepare('SELECT id, type, name, size, at FROM labels WHERE box = ? ORDER BY at, id').bind(box).all();
    return json({ items: results || [] }, 200, origin);
  }
  if (!id && request.method === 'DELETE') {
    const { results } = await env.DB.prepare('SELECT id FROM labels WHERE box = ?').bind(box).all();
    await env.DB.batch([
      env.DB.prepare('DELETE FROM labels WHERE box = ?').bind(box),
      env.DB.prepare('DELETE FROM boxes WHERE box = ?').bind(box)
    ]);
    await Promise.all((results || []).map((it) => env.LABELS.delete(fileKey(box, it.id))));
    return json({ ok: true }, 200, origin);
  }
  if (!id || !ID_RE.test(id)) return json({ error: 'not found' }, 404, origin);

  if (request.method === 'GET') {
    const row = await env.DB.prepare('SELECT type, at FROM labels WHERE box = ? AND id = ? AND at > ?').bind(box, id, oldest()).first();
    if (!row) return json({ error: 'This label is no longer waiting.' }, 404, origin);
    const got = await env.LABELS.getWithMetadata(fileKey(box, id), 'arrayBuffer');
    if (!got || !got.value) {
      // Just sent from far away: its file may not have reached this location yet.
      if (row.at > Date.now() - ARRIVING_MS) return json({ error: 'This label is still on its way. Try again in a moment.' }, 409, origin);
      await env.DB.prepare('DELETE FROM labels WHERE box = ? AND id = ?').bind(box, id).run();
      return json({ error: 'This label is no longer waiting.' }, 404, origin);
    }
    return new Response(got.value, { status: 200, headers: { 'Content-Type': row.type, 'Cache-Control': 'no-store', ...corsHeaders(origin) } });
  }
  if (request.method === 'DELETE') {
    await env.DB.prepare('DELETE FROM labels WHERE box = ? AND id = ?').bind(box, id).run();
    await env.LABELS.delete(fileKey(box, id));
    const left = await env.DB.prepare('SELECT COUNT(*) AS n FROM labels WHERE box = ? AND at > ?').bind(box, oldest()).first();
    return json({ ok: true, waiting: left ? left.n : 0 }, 200, origin);
  }
  return json({ error: 'method not allowed' }, 405, origin);
}

async function suggest(request, env, url, origin, ctx) {
  if (!ALLOWED_ORIGINS.includes(origin)) return json({ error: 'Address suggestions are for the Label Sheet Keeper panel.' }, 403, origin);
  if (!env.GEOAPIFY_KEY) return json({ error: 'not set up' }, 503, origin);
  const text = (url.searchParams.get('text') || '').trim().replace(/\s+/g, ' ').slice(0, 200);
  if (text.length < 3) return json({ error: 'Type a little more.' }, 400, origin);

  const today = new Date().toISOString().slice(0, 10);
  const who = (await sha256((request.headers.get('CF-Connecting-IP') || 'unknown') + '|' + today)).slice(0, 32);
  const usedUp = () => json({ error: 'Address suggestions are used up for today.' }, 429, origin);
  if (outOf.get('*') === today || outOf.get(who) === today) return usedUp();

  // Count the question and ask Geoapify at the same time, so the answer comes sooner.
  const count = COUNT_SQL;
  const asking = fetch(GEO_URL + encodeURIComponent(text) + '&apiKey=' + encodeURIComponent(env.GEOAPIFY_KEY)).catch(() => null);
  const [all, mine] = await env.DB.batch([env.DB.prepare(count).bind(today, '*'), env.DB.prepare(count).bind(today, who)]);
  if (ctx && Math.random() < 0.02) ctx.waitUntil(env.DB.prepare('DELETE FROM usage WHERE day < ?').bind(today).run());
  const overAll = all.results[0].n > SUGGEST_PER_DAY, overMine = mine.results[0].n > SUGGEST_PER_PERSON;
  if (overAll || overMine) {
    // Remembered by this copy of the Worker, so Geoapify isn't asked for them again today.
    if (outOf.size > 5000) outOf.clear();
    outOf.set(overAll ? '*' : who, today);
    return usedUp();
  }

  const r = await asking;
  if (!r) return json({ error: 'Address suggestions are busy.' }, 503, origin);
  if (r.status === 401 || r.status === 403) return json({ error: 'not set up' }, 503, origin);
  if (r.status === 429) return json({ error: 'Address suggestions are used up for today.' }, 429, origin);
  if (!r.ok) return json({ error: 'Address suggestions are busy.' }, 503, origin);
  const j = await r.json().catch(() => ({}));
  const results = ((j && j.results) || []).slice(0, 5).map((x) => {
    const o = {};
    GEO_KEEP.forEach((k) => { if (x[k] !== undefined && x[k] !== null) o[k] = x[k]; });
    return o;
  });
  return json({ results }, 200, origin);
}

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    const origin = request.headers.get('Origin') || '';
    const parts = url.pathname.split('/').filter(Boolean);

    if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: corsHeaders(origin) });

    try {
      if (url.pathname === '/' || url.pathname === '/health') {
        if (request.method !== 'GET') return json({ error: 'method not allowed' }, 405, origin);
        return json({ ok: true, service: 'label-sheet-keeper', version: VERSION }, 200, origin);
      }
      await ready(env);
      if (parts[0] === 'pair' && parts.length === 1) {
        if (request.method !== 'POST') return json({ error: 'method not allowed' }, 405, origin);
        return await pair(request, env, url, origin);
      }
      if (parts[0] === 'drop' && parts.length === 2) {
        if (request.method !== 'POST') return json({ error: 'method not allowed' }, 405, origin);
        return await drop(request, env, parts[1], origin, url.searchParams.get('reply') === 'text');
      }
      if (parts[0] === 'box' && (parts.length === 2 || parts.length === 3)) {
        return await boxRoute(request, env, parts[1], parts[2] || '', origin);
      }
      if (parts[0] === 'suggest' && parts.length === 1) {
        if (request.method !== 'GET') return json({ error: 'method not allowed' }, 405, origin);
        return await suggest(request, env, url, origin, ctx);
      }
      const acct = await accountRoute(request, env, parts, origin);
      if (acct) return acct;
      return json({ error: 'not found' }, 404, origin);
    } catch (e) {
      return json({ error: 'Something went wrong on the server. Try again.' }, 500, origin);
    }
  }
};
