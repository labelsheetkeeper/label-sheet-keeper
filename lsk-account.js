/*
  Label Sheet Keeper — your account and sync.

  Signed in, the address book and a few settings are kept the same on every
  computer, through the Label Sheet Keeper server (worker/src/account.js).

  What syncs: each address-book contact (one record per contact), and the
  settings listed in SET below (one record each). Label sheets, printer
  adjustments, the Geoapify key and the phone link stay on this computer.

  How: everything still lives in this computer's storage, and the panel keeps
  working without a connection. This file compares storage with what was last
  synced (the "shadow"), notes when each thing changed ("pending"), sends
  those changes and brings in the others'. For one record, the newest change
  wins. The first time a computer syncs, the account's settings win over this
  computer's, and contacts with the same address are joined into one.

  Kept on this computer only (never synced, never in backups):
    lsk.acct          { session, name, email, since, at }
    lsk.acct.shadow   { "book/<id>": { h: hash, m: when } | { d: 1, m } }
    lsk.acct.pending  { "book/<id>": { h, m } | { d: 1, m } }
*/
(function (root) {
  'use strict';

  var KEY = { acct: 'lsk.acct', shadow: 'lsk.acct.shadow', pending: 'lsk.acct.pending', lock: 'lsk.acct.lock' };
  var BOOK = 'lsk.addressBook';
  // Settings that follow the account. Each is one stored value.
  var SET = ['lsk.fromAddress', 'lsk.markWithoutAsking', 'lsk.feedArrow', 'lsk.tripLimit', 'lsk.warnPassHidden', 'lsk.design', 'lsk.logo',
    'lsk.customFormats', 'lsk.homeCountry', 'lsk.suggest', 'lsk.codes', 'lsk.recent', 'lsk.showSaved'];
  var SET_PREFIX = ['lsk.stock.'];     // label stock, one per size
  var ID_OK = /^[A-Za-z0-9._:-]{1,100}$/;
  var MAX_VALUE = 300 * 1024;          // the server's limit for one record
  var CHUNK = 200, CHUNK_BYTES = 700 * 1024;
  var FIELDS = ['phone', 'phone2', 'email', 'email2', 'lines2', 'groups', 'show'];

  function isSynced(key) {
    if (key === BOOK || SET.indexOf(key) >= 0) return true;
    return SET_PREFIX.some(function (p) { return String(key).indexOf(p) === 0; });
  }
  // Short fingerprint of a record's text, to tell whether it changed.
  function hash(s) {
    var h = 0x811c9dc5;
    for (var i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = (h + ((h << 1) + (h << 4) + (h << 7) + (h << 8) + (h << 24))) >>> 0; }
    return h.toString(36) + '.' + s.length;
  }
  function newId() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 8); }
  function addressKey(e) { return (e && Array.isArray(e.lines) ? e.lines : []).join('\n').toLowerCase(); }
  // Name, email from a Microsoft ID token (already checked by the server).
  function profileOf(idToken) {
    try {
      var p = String(idToken).split('.')[1].replace(/-/g, '+').replace(/_/g, '/');
      var j = JSON.parse(decodeURIComponent(escape(root.atob ? root.atob(p) : Buffer.from(p, 'base64').toString('binary'))));
      return { name: String(j.name || ''), email: String(j.email || j.preferred_username || '') };
    } catch (e) { return { name: '', email: '' }; }
  }

  // opts: { storage, api, fetch, now, changed(keys), status(s) }
  function make(opts) {
    var S = opts.storage, api = opts.api, now = opts.now || Date.now;
    var doFetch = opts.fetch || function (u, o) { return root.fetch(u, o); };
    var me = newId(), running = null, again = false, timer = 0, stampTimer = 0, state = { state: 'idle', error: '' };

    function get(k) { try { return S.getItem(k); } catch (e) { return null; } }
    function put(k, v) { try { if (v === null) S.removeItem(k); else S.setItem(k, v); } catch (e) { /* storage full or blocked */ } }
    function getJson(k, d) { try { var v = JSON.parse(get(k) || 'null'); return v === null || typeof v !== 'object' ? d : v; } catch (e) { return d; } }
    function putJson(k, v) { put(k, JSON.stringify(v)); }
    function acct() { var a = getJson(KEY.acct, null); return a && a.session ? a : null; }
    function setState(s, error) { state = { state: s, error: error || '' }; if (opts.status) opts.status(status()); }
    function status() { var a = acct(); return { on: !!a, state: a ? state.state : 'off', error: state.error, at: a ? a.at || 0 : 0, name: a ? a.name : '', email: a ? a.email : '' }; }

    function readBook() { var b = getJson(BOOK, []); return Array.isArray(b) ? b : []; }
    function bookReadable() { try { return Array.isArray(JSON.parse(get(BOOK) || 'null')); } catch (e) { return false; } }
    function writeBook(b) {
      b.sort(function (x, y) { return String(x.name || '').localeCompare(String(y.name || '')); });
      putJson(BOOK, b);
    }
    // Everything on this computer that syncs: { "book/<id>": text, "set/<name>": text }
    function local() {
      var out = {}, b = readBook(), fixed = false, seen = {};
      b.forEach(function (e) {
        if (!e || typeof e !== 'object') return;
        if (!e.id || !ID_OK.test(String(e.id)) || seen[e.id]) { e.id = newId(); fixed = true; }
        seen[e.id] = true;
        out['book/' + e.id] = JSON.stringify(e);
      });
      if (fixed) putJson(BOOK, b);
      var n = 0;
      try { n = S.length; } catch (e) { n = 0; }
      for (var i = 0; i < n; i++) {
        var k = S.key(i);
        if (!k || k === BOOK || !isSynced(k)) continue;
        var id = k.slice(4), v = get(k);
        if (ID_OK.test(id) && typeof v === 'string' && v.length) out['set/' + id] = v;
      }
      return out;
    }

    // Notes what changed on this computer since the last sync, and when.
    function stamp() {
      var a = acct();
      if (!a) return null;
      var shadow = getJson(KEY.shadow, {}), pending = getJson(KEY.pending, {}), t = now(), dirty = false;
      // An address book that has gone missing or unreadable here is never taken to mean
      // "delete every address everywhere": this computer gets the account's copy again.
      if (!bookReadable() && Object.keys(shadow).some(function (k) { return k.indexOf('book/') === 0; })) {
        Object.keys(shadow).forEach(function (k) { if (k.indexOf('book/') === 0) delete shadow[k]; });
        Object.keys(pending).forEach(function (k) { if (k.indexOf('book/') === 0) delete pending[k]; });
        putJson(KEY.shadow, shadow); putJson(KEY.pending, pending);
        a.since = 0; putJson(KEY.acct, a);
      }
      var recs = local();
      // A computer syncing for the first time: the account's settings win over the ones here.
      var first = !a.since && !Object.keys(shadow).length;
      Object.keys(recs).forEach(function (k) {
        var h = hash(recs[k]), sh = shadow[k], p = pending[k];
        if (sh && !sh.d && sh.h === h) { if (p) { delete pending[k]; dirty = true; } return; }
        if (p && !p.d && p.h === h) return;
        pending[k] = { h: h, m: first && k.indexOf('set/') === 0 ? 1 : t };
        dirty = true;
      });
      Object.keys(shadow).forEach(function (k) {
        if (k in recs || shadow[k].d || (pending[k] && pending[k].d)) return;
        pending[k] = { d: 1, m: t };
        dirty = true;
      });
      // Something added and removed again before it ever synced.
      Object.keys(pending).forEach(function (k) { if (!(k in recs) && !shadow[k]) { delete pending[k]; dirty = true; } });
      if (dirty) putJson(KEY.pending, pending);
      return { recs: recs, pending: pending };
    }

    // Brings in records from the server. Returns the storage keys it changed.
    function applyRows(rows) {
      if (!rows || !rows.length) return [];
      var shadow = getJson(KEY.shadow, {}), pending = getJson(KEY.pending, {}), b = null, changed = {};
      var byId = function () { var m = {}; b.forEach(function (e, i) { if (e && e.id) m[e.id] = i; }); return m; };
      rows.forEach(function (r) {
        if (r.c !== 'book' && r.c !== 'set') return;        // a kind a newer version syncs; it stays on the server
        if (r.c === 'set' && (!isSynced('lsk.' + r.id) || 'lsk.' + r.id === BOOK)) return;   // only the settings listed above are ever written
        var k = r.c + '/' + r.id, p = pending[k];
        if (p && p.m > r.m) return;                         // this computer's change is newer; it goes up next
        if (r.c === 'set') {
          put('lsk.' + r.id, r.d ? null : r.v);
          changed['lsk.' + r.id] = true;
        } else {
          if (!b) b = readBook();
          var at = byId()[r.id], e = null;
          if (!r.d) { try { e = JSON.parse(r.v); } catch (err) { e = null; } if (!e || typeof e !== 'object' || Array.isArray(e)) return; e.id = r.id; }
          if (r.d) { if (at !== undefined) b.splice(at, 1); }
          else {
            // The same address saved here before this computer synced: one contact, with both sets of details.
            var key = addressKey(e);
            for (var i = b.length - 1; key && i >= 0; i--) {
              var o = b[i];
              if (!o || o.id === r.id || shadow['book/' + o.id] || addressKey(o) !== key) continue;
              FIELDS.forEach(function (f) { if (o[f] && !e[f]) e[f] = o[f]; });
              b.splice(i, 1);
            }
            at = byId()[r.id];
            if (at === undefined) b.push(e); else b[at] = e;
          }
          changed[BOOK] = true;
        }
        shadow[k] = r.d ? { d: 1, m: r.m } : { h: hash(r.v), m: r.m };
        delete pending[k];
      });
      if (b) writeBook(b);
      putJson(KEY.shadow, shadow);
      putJson(KEY.pending, pending);
      return Object.keys(changed);
    }

    function post(path, method, body) {
      var a = acct();
      if (!a) return Promise.reject(new Error('signed out'));
      var init = { method: method, headers: { Authorization: 'Bearer ' + a.session } };
      if (body !== undefined) { init.headers['Content-Type'] = 'application/json'; init.body = JSON.stringify(body); }
      return doFetch(api + path, init).then(function (r) {
        return r.json().catch(function () { return {}; }).then(function (j) {
          if (r.ok) return j;
          var e = new Error((j && j.error) || 'HTTP ' + r.status);
          e.status = r.status;
          // The server no longer knows this session (signed out elsewhere, or the account was deleted).
          if (r.status === 401) { forget(); e.signedOut = true; }
          throw e;
        });
      });
    }
    function saveSince(next) { var a = acct(); if (a && typeof next === 'number') { a.since = next; putJson(KEY.acct, a); } }
    function forget() { put(KEY.acct, null); put(KEY.shadow, null); put(KEY.pending, null); put(KEY.lock, null); }

    // One computer's panels take turns (each Word window has its own panel).
    function takeTurn() {
      var l = getJson(KEY.lock, null), t = now();
      if (l && l.by !== me && t - l.at < 20000) return false;
      putJson(KEY.lock, { by: me, at: t });
      return true;
    }

    function cycle() {
      var changed = {}, guard = 0;
      var note = function (keys) { keys.forEach(function (k) { changed[k] = true; }); };
      var pullAll = function () {
        var a = acct();
        return post('/sync', 'POST', { since: (a && a.since) || 0, push: [] }).then(function (res) {
          note(applyRows(res.pull)); saveSince(res.next); putJson(KEY.lock, { by: me, at: now() });
          return res.more && ++guard < 400 ? pullAll() : null;
        });
      };
      var pushAll = function () {
        var s = stamp();
        if (!s) return null;
        var list = [], bytes = 0;
        Object.keys(s.pending).some(function (k) {
          var p = s.pending[k], i = k.indexOf('/'), item = { c: k.slice(0, i), id: k.slice(i + 1), m: p.m, d: p.d ? 1 : 0 };
          if (!p.d) {
            var v = s.recs[k];
            if (v === undefined || v.length > MAX_VALUE) return false;     // too big to sync: stays on this computer
            item.v = v; bytes += v.length;
          }
          list.push(item);
          return list.length >= CHUNK || bytes >= CHUNK_BYTES;
        });
        if (!list.length || ++guard > 400) return null;
        var a = acct();
        return post('/sync', 'POST', { since: (a && a.since) || 0, push: list }).then(function (res) {
          note(applyRows(res.pull));
          // What went up is now what the account holds, unless it changed here again meanwhile.
          var shadow = getJson(KEY.shadow, {}), pending = getJson(KEY.pending, {}), back = {};
          (res.pull || []).forEach(function (r) { back[r.c + '/' + r.id] = true; });
          list.forEach(function (it) {
            var k = it.c + '/' + it.id, p = pending[k];
            if (back[k]) return;
            shadow[k] = it.d ? { d: 1, m: it.m } : { h: hash(it.v), m: it.m };
            if (p && p.m === it.m && !!p.d === !!it.d) delete pending[k];
          });
          putJson(KEY.shadow, shadow); putJson(KEY.pending, pending);
          saveSince(res.next); putJson(KEY.lock, { by: me, at: now() });
          return (res.more ? pullAll() : Promise.resolve()).then(pushAll);
        });
      };
      stamp();
      return pullAll().then(pushAll).then(function () {
        var a = acct();
        if (a) { a.at = now(); putJson(KEY.acct, a); }
        return Object.keys(changed);
      });
    }

    // Syncs now. Resolves to the storage keys other computers changed. One at a time.
    function sync() {
      if (!acct()) return Promise.resolve([]);
      if (running) { again = true; return running; }
      if (!takeTurn()) { soon(5000); return Promise.resolve([]); }
      setState('syncing');
      running = cycle().then(function (keys) {
        running = null; put(KEY.lock, null); setState('ok');
        if (keys.length && opts.changed) opts.changed(keys);
        if (again) { again = false; soon(500); }
        return keys;
      }, function (e) {
        running = null; put(KEY.lock, null); again = false;
        setState(e && e.signedOut ? 'off' : 'error', e && e.signedOut ? 'You were signed out. Turn on sync again to keep syncing.' : (e && e.status ? e.message : 'The server couldn’t be reached. Syncing will try again.'));
        if (e && e.signedOut && opts.changed) opts.changed([]);
        throw e;
      });
      return running;
    }
    function soon(ms) {
      if (opts.timers === false) return;
      clearTimeout(timer);
      timer = setTimeout(function () { sync().catch(function () { /* shown as the sync status */ }); }, ms);
    }
    // Something was saved on this computer: note the time now, and sync shortly.
    function touch(key) {
      if (!acct() || (key && !isSynced(key))) return;
      if (opts.timers === false) { stamp(); return; }
      clearTimeout(stampTimer);
      stampTimer = setTimeout(stamp, 250);
      soon(2000);
    }

    // Exchanges a Microsoft ID token for a session and starts syncing.
    function signIn(idToken) {
      return doFetch(api + '/auth/microsoft', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ id_token: idToken }) }).then(function (r) {
        return r.json().catch(function () { return {}; }).then(function (j) {
          if (!r.ok || !j.session) { var e = new Error((j && j.error) || 'Signing in didn’t work. Try again.'); e.status = r.status; throw e; }
          var p = profileOf(idToken);
          forget();
          putJson(KEY.acct, { session: j.session, name: p.name, email: p.email, since: 0, at: 0 });
          setState('idle');
          return { isNew: !!j.isNew, name: p.name, email: p.email };
        });
      });
    }
    // Stops syncing on this computer. What's here stays here.
    function signOut() {
      clearTimeout(timer);
      var done = function () { forget(); setState('off'); };
      if (!acct()) { done(); return Promise.resolve(); }
      return post('/signout', 'POST').then(done, done);
    }
    // Deletes the account and everything in it on the server. What's on this computer stays.
    function remove() {
      clearTimeout(timer);
      return post('/account', 'DELETE').then(function () { forget(); setState('off'); });
    }
    function usage() { return post('/account', 'GET'); }

    return { status: status, sync: sync, touch: touch, stamp: stamp, signIn: signIn, signOut: signOut, remove: remove, usage: usage, soon: soon, isSynced: isSynced };
  }

  var api = { make: make, isSynced: isSynced, hash: hash, profileOf: profileOf, KEY: KEY };
  root.LSK_ACCOUNT = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
