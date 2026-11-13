// Checks syncing end to end without a network: several pretend computers,
// each with its own storage and the panel's sync code (lsk-account.js), talk
// to the real Worker code (worker/) running here on a stand-in database.
// Run from the repository folder: node tests/sync-test.mjs
import assert from 'node:assert';
import { createRequire } from 'node:module';
import worker from '../worker/src/index.js';
import { forgetKeys } from '../worker/src/account.js';
import { FakeD1, FakeKV, fakeMicrosoft } from './fakes.mjs';

const require = createRequire(import.meta.url);
const A = require('../lsk-account.js');
const PANEL = 'https://labelsheetkeeper.github.io', API = 'https://api.example.workers.dev', CLIENT = '11111111-2222-3333-4444-555555555555';
let n = 0;
const ok = (c, m) => { assert(c, m); n++; };
const J = JSON.stringify;

const db = new FakeD1(), env = { DB: db, LABELS: new FakeKV(), MS_CLIENT_ID: CLIENT };
const ms = await fakeMicrosoft(CLIENT);
forgetKeys();
globalThis.fetch = ms.fetch;                     // the Worker fetching Microsoft's keys

class Storage {
  constructor() { this.m = new Map(); }
  get length() { return this.m.size; }
  key(i) { return [...this.m.keys()][i] ?? null; }
  getItem(k) { return this.m.has(k) ? this.m.get(k) : null; }
  setItem(k, v) { this.m.set(k, String(v)); }
  removeItem(k) { this.m.delete(k); }
}
let clock = Date.now() - 3600 * 1000;            // every computer's clock; tests move it forward
const tick = (ms2 = 1000) => { clock += ms2; };
// One pretend computer.
function computer(name) {
  const c = { name, storage: new Storage(), offline: false, calls: [], changed: [], statuses: [] };
  const fetch = async (u, init = {}) => {
    if (c.offline) throw new TypeError('offline');
    const body = init.body ? JSON.parse(init.body) : null;
    c.calls.push({ path: new URL(u).pathname, method: init.method, push: body && body.push ? body.push.length : 0 });
    return worker.fetch(new Request(u, { ...init, headers: { ...init.headers, Origin: PANEL, 'CF-Connecting-IP': '203.0.113.' + (name.charCodeAt(0) % 200) } }), env);
  };
  c.engine = (id = 'x') => A.make({ storage: c.storage, api: API, fetch, now: () => clock, timers: false, changed: (keys) => c.changed.push(keys), status: (s) => c.statuses.push(s.state) });
  c.acct = c.engine();
  c.book = () => JSON.parse(c.storage.getItem('lsk.addressBook') || '[]');
  c.setBook = (b) => { c.storage.setItem('lsk.addressBook', J(b)); c.acct.touch('lsk.addressBook'); };
  c.set = (k, v) => { if (v === null) c.storage.removeItem(k); else c.storage.setItem(k, v); c.acct.touch(k); };
  c.names = () => c.book().map((e) => e.name).join(', ');
  c.pushes = () => c.calls.filter((x) => x.push > 0).length;
  return c;
}
const contact = (id, name, street, more = {}) => ({ id, name, lines: [name, street, 'Boston, MA 02108'], ...more });
const stored = (c = '%') => db.rows('SELECT c, id, d, v FROM records WHERE c LIKE ? ORDER BY c, id', c);

// ---- the first computer turns on sync ----
const mac = computer('mac');
mac.storage.setItem('lsk.addressBook', J([contact('a1', 'Ann Lee', '8 Maple Way'), contact('b1', 'Bob Ray', '1 Oak Ave')]));
mac.storage.setItem('lsk.homeCountry', 'CA');
mac.storage.setItem('lsk.fromAddress', 'Me\n5 Pine Rd\nSalem, MA 01970');
mac.storage.setItem('lsk.stock.avery-5160', J({ on: 12 }));
mac.storage.setItem('lsk.geoKey', 'my-own-geoapify-key');
mac.storage.setItem('lsk.nudge.avery-5160', J({ x: 0.03, y: 0 }));
mac.storage.setItem('lsk.phone', J({ box: 'b', key: 'k' }));
mac.storage.setItem('lsk.sheets', J([{ id: 's' }]));
ok(mac.acct.status().on === false && (await mac.acct.sync()).length === 0 && mac.calls.length === 0, 'signed out, nothing is sent anywhere');
let who = await mac.acct.signIn(await ms.idToken(1));
ok(who.isNew === true && who.email === 'person1@example.com' && mac.acct.status().on && mac.acct.status().email === 'person1@example.com', 'signing in starts an account, and the panel knows who it is ' + J(who));
await mac.acct.sync();
ok(J(stored().map((r) => r.c + '/' + r.id)) === J(['book/a1', 'book/b1', 'set/fromAddress', 'set/homeCountry', 'set/stock.avery-5160']), 'the address book and chosen settings go to the account ' + J(stored().map((r) => r.id)));
ok(!/my-own-geoapify-key|nudge|"box"|sheets/.test(J(stored())), 'the Geoapify key, printer adjustment, phone link and sheet list stay on this computer');
ok(!J(db.rows('SELECT * FROM accounts')).includes('person1'), 'the server doesn’t keep the name or email');
ok(mac.statuses.join() === 'idle,syncing,ok' && mac.acct.status().at === clock, 'the status goes syncing, then ok ' + mac.statuses.join());
let before = mac.pushes();
await mac.acct.sync();
ok(mac.pushes() === before, 'with nothing changed, nothing is sent again');

// ---- a second computer joins, with addresses and settings of its own ----
const laptop = computer('laptop');
laptop.storage.setItem('lsk.addressBook', J([contact('z9', 'Ann Lee', '8 Maple Way', { phone: '617-555-0100' }), contact('c1', 'Cy Fox', '3 Elm St'), { name: 'No Id', lines: ['No Id', '4 Ash St'] }]));
laptop.storage.setItem('lsk.homeCountry', 'US');
laptop.storage.setItem('lsk.feedArrow', 'off');
who = await laptop.acct.signIn(await ms.idToken(1));
ok(who.isNew === false, 'the same person on another computer gets the same account');
tick();
await laptop.acct.sync();
ok(laptop.names() === 'Ann Lee, Bob Ray, Cy Fox, No Id', 'it gets the account’s addresses and keeps its own ' + laptop.names());
const ann = laptop.book().find((e) => e.name === 'Ann Lee');
ok(ann.id === 'a1' && ann.phone === '617-555-0100', 'the same address on both becomes one contact, with the phone only this computer had ' + J(ann));
ok(laptop.storage.getItem('lsk.homeCountry') === 'CA', 'the account’s settings win over a new computer’s');
ok(laptop.storage.getItem('lsk.feedArrow') === 'off' && laptop.storage.getItem('lsk.fromAddress').startsWith('Me\n5 Pine'), 'settings only one side had end up on both');
ok(laptop.book().every((e) => /^[A-Za-z0-9._:-]+$/.test(e.id)), 'a contact saved without an id gets one');
ok(laptop.changed.length === 1 && laptop.changed[0].includes('lsk.addressBook') && laptop.changed[0].includes('lsk.homeCountry'), 'the panel is told what to redraw ' + J(laptop.changed));
tick();
await mac.acct.sync();
ok(mac.names() === 'Ann Lee, Bob Ray, Cy Fox, No Id' && mac.book().find((e) => e.id === 'a1').phone === '617-555-0100' && mac.storage.getItem('lsk.feedArrow') === 'off', 'the first computer gets the second one’s additions ' + mac.names());
ok(mac.storage.getItem('lsk.geoKey') === 'my-own-geoapify-key' && laptop.storage.getItem('lsk.geoKey') === null, 'and its own key is still only there');
ok(stored('book').filter((r) => !r.d).length === 4, 'the account has each contact once');

// ---- everyday changes ----
tick();
mac.setBook(mac.book().map((e) => (e.id === 'b1' ? { ...e, name: 'Bob Ray Jr', lines: ['Bob Ray Jr', '1 Oak Ave', 'Boston, MA 02108'] } : e)));
await mac.acct.sync(); await laptop.acct.sync();
ok(laptop.book().find((e) => e.id === 'b1').name === 'Bob Ray Jr', 'an edit on one computer shows up on the other');
ok(laptop.names() === 'Ann Lee, Bob Ray Jr, Cy Fox, No Id', 'the address book stays in name order');
tick();
laptop.setBook(laptop.book().filter((e) => e.id !== 'c1'));
await laptop.acct.sync(); await mac.acct.sync();
ok(!mac.book().some((e) => e.id === 'c1') && stored('book').find((r) => r.id === 'c1').d === 1, 'deleting an address deletes it everywhere');
tick();
mac.set('lsk.feedArrow', null);
mac.set('lsk.tripLimit', '5');
await mac.acct.sync(); await laptop.acct.sync();
ok(laptop.storage.getItem('lsk.feedArrow') === null && laptop.storage.getItem('lsk.tripLimit') === '5', 'settings changes and resets carry over');

// ---- the same contact changed on both: the later change wins, whoever syncs first ----
tick();
mac.setBook(mac.book().map((e) => (e.id === 'a1' ? { ...e, email: 'early@example.com' } : e)));     // earlier
tick();
laptop.setBook(laptop.book().map((e) => (e.id === 'a1' ? { ...e, email: 'later@example.com' } : e)));  // later
tick();
await laptop.acct.sync(); await mac.acct.sync(); await laptop.acct.sync();
ok(mac.book().find((e) => e.id === 'a1').email === 'later@example.com' && laptop.book().find((e) => e.id === 'a1').email === 'later@example.com', 'the later change wins on both');
tick();
laptop.setBook(laptop.book().map((e) => (e.id === 'a1' ? { ...e, email: 'first@example.com' } : e)));  // earlier this time
tick();
mac.setBook(mac.book().map((e) => (e.id === 'a1' ? { ...e, email: 'second@example.com' } : e)));       // later
tick();
await laptop.acct.sync(); await mac.acct.sync(); await laptop.acct.sync();
ok(mac.book().find((e) => e.id === 'a1').email === 'second@example.com' && laptop.book().find((e) => e.id === 'a1').email === 'second@example.com', 'also when the earlier one syncs first');

// ---- no connection ----
tick();
mac.offline = true;
mac.setBook(mac.book().concat([contact('d1', 'Di Oz', '6 Fir Ln')]));
let failed = null;
try { await mac.acct.sync(); } catch (e) { failed = e; }
ok(failed && mac.acct.status().state === 'error' && /couldn’t be reached/.test(mac.acct.status().error), 'offline, syncing says the server couldn’t be reached');
ok(mac.names().includes('Di Oz'), 'and the panel keeps working with what’s on this computer');
mac.offline = false;
tick(60000);
await mac.acct.sync(); await laptop.acct.sync();
ok(laptop.names().includes('Di Oz') && mac.acct.status().state === 'ok', 'back online, the change made offline goes through');

// ---- an address book that goes missing on one computer doesn't empty everyone's ----
tick();
const hadBook = laptop.book().length;
laptop.storage.removeItem('lsk.addressBook');
await laptop.acct.sync(); await mac.acct.sync();
ok(mac.book().length === hadBook && stored('book').filter((r) => !r.d).length === hadBook, 'if one computer loses its address book, nothing is deleted from the account or the others');
ok(laptop.book().length === hadBook, 'and that computer gets its addresses back from the account');
tick();
laptop.storage.setItem('lsk.addressBook', '{not json');
await laptop.acct.sync(); await mac.acct.sync();
ok(mac.book().length === hadBook && laptop.book().length === hadBook, 'the same goes for an address book that can’t be read');
tick();
laptop.setBook(laptop.book().filter((e) => e.id !== 'd1'));
await laptop.acct.sync(); await mac.acct.sync();
ok(mac.book().length === hadBook - 1, 'while deleting an address on purpose still deletes it everywhere');
tick();
mac.setBook(mac.book().concat([contact('d1', 'Di Oz', '6 Fir Ln')]));
await mac.acct.sync(); await laptop.acct.sync();

// ---- limits and odd records ----
tick();
mac.set('lsk.logo', J({ b64: 'x'.repeat(400 * 1024), w: 10, h: 10 }));
mac.set('lsk.design', 'bold');
await mac.acct.sync(); await laptop.acct.sync();
ok(laptop.storage.getItem('lsk.design') === 'bold' && laptop.storage.getItem('lsk.logo') === null && mac.acct.status().state === 'ok', 'something too big to sync stays on its computer without holding up the rest');
mac.set('lsk.logo', J({ b64: 'small', w: 10, h: 10 }));
await mac.acct.sync(); await laptop.acct.sync();
ok(JSON.parse(laptop.storage.getItem('lsk.logo')).b64 === 'small', 'and syncs once it’s small enough');
const acctId = db.rows('SELECT id FROM accounts')[0].id;
db.db.prepare("UPDATE accounts SET rev = rev + 3 WHERE id = ?").run(acctId);
const rev = db.rows('SELECT rev FROM accounts')[0].rev;
const plant = db.db.prepare('INSERT INTO records (account, c, id, m, d, v, rev) VALUES (?, ?, ?, ?, 0, ?, ?)');
plant.run(acctId, 'sheet', 's1', clock, '{"name":"Sheet 1"}', rev - 2);
plant.run(acctId, 'set', 'acct', clock, '{"session":"stolen"}', rev - 1);
plant.run(acctId, 'set', 'geoKey', clock, 'planted', rev);
const sessionBefore = laptop.storage.getItem('lsk.acct');
await laptop.acct.sync();
ok(JSON.parse(laptop.storage.getItem('lsk.acct')).session === JSON.parse(sessionBefore).session && laptop.storage.getItem('lsk.geoKey') === null, 'records for anything but the address book and the listed settings are never written to a computer');
ok(laptop.acct.status().state === 'ok' && stored('sheet').length === 1, 'and kinds a newer version syncs are left alone');

// ---- a big address book, and a third computer ----
tick();
mac.setBook(mac.book().concat(Array.from({ length: 450 }, (_, i) => contact('big' + i, 'Person ' + String(i).padStart(3, '0'), i + ' Long Rd'))));
before = mac.calls.length;
await mac.acct.sync();
ok(mac.calls.slice(before).filter((x) => x.push > 0).map((x) => x.push).join() === '200,200,50', 'a big address book goes up in parts ' + mac.calls.slice(before).map((x) => x.push).join());
const desk = computer('desk');
await desk.acct.signIn(await ms.idToken(1));
await desk.acct.sync();
ok(desk.book().length === mac.book().length && desk.book().length === 454 && desk.pushes() === 0, 'a third computer gets all of it and has nothing to send ' + desk.book().length);
ok(desk.calls.filter((x) => x.path === '/sync').length === 2, 'in pages');

// ---- two panels on one computer (one per Word window) take turns ----
const second = mac.engine();
tick();
mac.setBook(mac.book().map((e) => (e.id === 'd1' ? { ...e, phone: '555-0199' } : e)));
before = mac.calls.length;
const [r1, r2] = await Promise.all([mac.acct.sync(), second.sync()]);
ok(mac.calls.length - before === 2 && Array.isArray(r1) && r2.length === 0, 'two panels on one computer don’t sync at the same moment ' + (mac.calls.length - before));
await second.sync();
ok(second.status().state === 'ok', 'the other panel syncs on its next turn');

// ---- another person ----
const other = computer('other');
other.storage.setItem('lsk.addressBook', J([contact('q1', 'Quinn Other', '9 Private Rd')]));
await other.acct.signIn(await ms.idToken(2));
await other.acct.sync();
ok(other.names() === 'Quinn Other' && db.rows('SELECT id FROM accounts').length === 2, 'another person has their own account and sees only their own addresses');
await mac.acct.sync();
ok(!mac.names().includes('Quinn'), 'and theirs never reach anyone else');

// ---- turning sync off, being signed out, deleting the account ----
await laptop.acct.signOut();
ok(laptop.acct.status().on === false && laptop.storage.getItem('lsk.acct') === null && laptop.storage.getItem('lsk.acct.shadow') === null, 'turning sync off forgets the session on that computer');
ok(laptop.book().length > 3 && laptop.storage.getItem('lsk.homeCountry') === 'CA', 'and leaves its addresses and settings where they are');
ok(db.rows('SELECT hash FROM sessions WHERE account = ?', acctId).length === 2, 'the server forgets that computer’s session too');
tick();
laptop.setBook(laptop.book().concat([contact('l9', 'Later Local', '2 Solo St')]));
await laptop.acct.sync(); await mac.acct.sync();
ok(!mac.names().includes('Later Local') && laptop.calls.slice(-1)[0].path === '/signout', 'changes made after that stay on that computer');

db.db.prepare('DELETE FROM sessions WHERE hash IN (SELECT hash FROM sessions WHERE account = ? ORDER BY created DESC LIMIT 1)').run(acctId);   // the desk is signed out from elsewhere
failed = null;
try { await desk.acct.sync(); } catch (e) { failed = e; }
ok(failed && failed.signedOut && desk.acct.status().on === false && /signed out/.test(desk.statuses.length && desk.acct.status().error), 'a computer signed out by the server says so and stops syncing');
ok(desk.book().length === 454, 'with its addresses still there');

const use = await mac.acct.usage();
ok(use.records > 450 && use.computers === 1, 'the account says how much it holds ' + J(use));
await mac.acct.remove();
ok(db.rows('SELECT id FROM accounts WHERE id = ?', acctId).length === 0 && db.rows('SELECT id FROM records WHERE account = ?', acctId).length === 0, 'deleting the account removes everything of it from the server');
ok(mac.acct.status().on === false && mac.book().length === 454, 'and leaves this computer’s addresses alone');
ok(db.rows('SELECT id FROM accounts').length === 1 && (await other.acct.sync(), other.acct.status().state === 'ok'), 'the other person’s account is untouched');

// A fresh start after deleting: signing in again makes a new, empty account and fills it from this computer.
who = await mac.acct.signIn(await ms.idToken(1));
await mac.acct.sync();
ok(who.isNew === true && db.rows('SELECT COUNT(*) AS n FROM records WHERE c = ? AND d = 0', 'book')[0].n === 455, 'signing in again starts over from what’s on this computer');

ok(A.isSynced('lsk.addressBook') && A.isSynced('lsk.stock.custom-1') && !A.isSynced('lsk.geoKey') && !A.isSynced('lsk.acct') && !A.isSynced('lsk.phone') && !A.isSynced('lsk.nudge.avery-5160') && !A.isSynced('lsk.sheets'), 'only the listed things ever sync');
console.log('all', n, 'sync checks passed');
