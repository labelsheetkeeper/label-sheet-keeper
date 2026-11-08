/*
  Label Sheet Keeper — the Word side panel.

  Words used throughout (code and screen):
    label            one spot on the sheet, numbered 1, 2, 3… across each row
    open             blank and still on the sheet
    used             already peeled off (locked and left blank in the document)
    ready to print   has text or a picture that hasn't been printed yet
    trip             one pass of the physical sheet through the printer
*/
(function () {
  'use strict';

  var VERSION = '1.12.1';
  var ADDIN_ID = '8d491c85-a76d-4d26-a6f5-b7b2dd4a1711';   // must match manifest.xml <Id>
  var F = window.LSK_FORMATS;
  var S = window.LSK_SHEET;
  var T = window.LSK_TOOLS;

  // Stored in the document (Office settings travel with the file)
  var DOC = {
    history: 'lskHistory',     // [{at, nums, text}]
    lastPrint: 'lskLastPrint', // last "printed" time the panel has seen
    trips: 'lskTrips',         // trips this sheet has made through the printer
    spots: 'lskSpots',         // {n: {at, t, m}} when each used label was printed or marked
    snap: 'lskSnapshot',       // {u: [used], r: {n: first line}} for repairs
    id: 'lskSheetId',
    onLabel: 'lskOnLabel',     // {phone, email}: on this sheet's address labels to start with
    autoShow: 'Office.AutoShowTaskpaneWithDocument'
  };
  // Stored on this Mac
  var LS = {
    from: 'lsk.fromAddress', autoMark: 'lsk.markWithoutAsking', side: 'lsk.side.', fillTab: 'lsk.fillTab', shipTab: 'lsk.shipTab',
    addinRef: 'lsk.addinRef', feedArrow: 'lsk.feedArrow', tripLimit: 'lsk.tripLimit', nudge: 'lsk.nudge.', testCheck: 'lsk.testCheck.',
    warnPass: 'lsk.warnPassHidden', last: 'lsk.last.', book: 'lsk.addressBook', design: 'lsk.design', logo: 'lsk.logo',
    custom: 'lsk.customFormats', sheets: 'lsk.sheets', tableInfo: 'lsk.tableInfo', tableDraft: 'lsk.tableDraft',
    home: 'lsk.homeCountry', geoKey: 'lsk.geoKey', suggest: 'lsk.suggest',
    stock: 'lsk.stock.', fillJob: 'lsk.fillJob', testNudge: 'lsk.testNudge.', codes: 'lsk.codes',
    recent: 'lsk.recent', saved: 'lsk.savedLabels', showSaved: 'lsk.showSaved', phone: 'lsk.phone', preview: 'lsk.preview'
  };
  // The Label Sheet Keeper server (a Cloudflare Worker, see worker/).
  var API = 'https://api.labelsheetkeeper.workers.dev';
  // A contact's phones and emails, in the order they go under an address.
  var XF = ['phone', 'phone2', 'email', 'email2'];
  var USED_TAG = 'lsk-used';

  // A missing element (for example while Word still shows an older copy of
  // the panel page after an update) gets a stand-in so the panel keeps working.
  function $(id) { return document.getElementById(id) || document.createElement('div'); }

  var st = { kind: 'loading', fmt: null, used: {}, ready: {}, texts: {}, pics: {}, over: {}, here: 0, header: null, margins: null, applied: null };
  var undo = null;            // last print marked used: { nums, cells: [{n, ooxml}], at, trips }
  var pendingConfirm = null;  // action for the confirm bar's main button
  var pendingAlt = null;      // action for its second button, when there is one
  var choosing = false;       // size chooser opened from a sheet
  var chosenId = F.formats.length ? F.formats[0].id : '';
  var chain = Promise.resolve();
  var busy = false;

  // ---------- small helpers ----------
  function store(key, value) {
    try {
      if (value === undefined) return localStorage.getItem(key);
      if (value === null) localStorage.removeItem(key); else localStorage.setItem(key, value);
    } catch (e) { /* storage full or blocked */ }
    return null;
  }
  function storeJson(key, value) {
    if (value === undefined) { try { return JSON.parse(store(key) || 'null'); } catch (e) { return null; } }
    store(key, value === null ? null : JSON.stringify(value));
    return null;
  }
  function plural(n, w, suffix) { return n + ' ' + w + (n === 1 ? '' : (suffix || 's')); }
  function cap(t) { return t.charAt(0).toUpperCase() + t.slice(1); }
  function ranges(list) {
    var a = list.slice().sort(function (x, y) { return x - y; }), out = [], i = 0;
    while (i < a.length) {
      var j = i;
      while (j + 1 < a.length && a[j + 1] === a[j] + 1) j++;
      out.push(i === j ? String(a[i]) : a[i] + '–' + a[j]);
      i = j + 1;
    }
    return out.join(', ');
  }
  function labelList(list) { return (list.length === 1 ? 'label ' : 'labels ') + ranges(list); }
  function total() { return st.fmt ? st.fmt.across * st.fmt.down : 0; }
  function keys(obj) { return Object.keys(obj || {}).map(Number).sort(function (a, b) { return a - b; }); }
  function homeCountry() { return store(LS.home) === 'CA' ? 'CA' : 'US'; }
  function openList() { var a = []; for (var n = 1; n <= total(); n++) if (!st.used[n] && !st.ready[n]) a.push(n); return a; }
  function nextOpen() { var o = openList(); return o.length ? o[0] : 0; }
  function userError(msg) { var e = new Error(msg); e.userMessage = msg; return e; }
  function fmtDate(ms) {
    try { return new Date(ms).toLocaleString(undefined, { month: 'short', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit' }); }
    catch (e) { return String(new Date(ms)); }
  }
  function fmtDay(ms) {
    try { return new Date(ms).toLocaleDateString(undefined, { month: 'short', day: 'numeric' }); } catch (e) { return ''; }
  }
  function timeOnly(ms) {
    try { return new Date(ms).toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' }); } catch (e) { return ''; }
  }
  function firstLine(t) {
    var l = String(t || '').split(/[\r\n\u000b]+/).map(function (s) { return s.trim(); }).filter(Boolean);
    return l.length ? l[0].slice(0, 48) : '';
  }
  function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 8); }
  function el(tag, cls, text) { var e = document.createElement(tag); if (cls) e.className = cls; if (text !== undefined) e.textContent = text; return e; }
  function settings() { try { return Office.context.document.settings; } catch (e) { return null; } }
  function getSetting(k) { var s = settings(); try { return s ? s.get(k) : null; } catch (e) { return null; } }
  function setSetting(k, v) { var s = settings(); try { if (s) { s.set(k, v); s.saveAsync(function () {}); } } catch (e) { /* ignore */ } }
  // Where the document is saved. Word gives the address it had when the panel
  // opened; the file properties are asked again every few seconds, so a sheet
  // saved later gets its address too.
  var liveUrl = null, urlAsked = 0;
  function docUrl() {
    if (liveUrl !== null) return liveUrl;
    try { return Office.context.document.url || ''; } catch (e) { return ''; }
  }
  function refreshUrl() {
    if (Date.now() - urlAsked < 4000) return;
    urlAsked = Date.now();
    try {
      Office.context.document.getFilePropertiesAsync(function (r) {
        if (r.status !== Office.AsyncResultStatus.Succeeded || !r.value) return;
        liveUrl = r.value.url || '';
        if (st.kind !== 'sheet') return;
        var mine = registry()[getSetting(DOC.id)];
        if (mine && mine.url !== liveUrl) updateRegistry();
        showSaveHint();
      });
    } catch (e) { /* older Word: keep the address from when the panel opened */ }
  }
  function fileName(url) {
    url = url || docUrl();
    if (!url) return '';
    var name = url.split(/[\\/]/).pop();
    try { name = decodeURIComponent(name); } catch (e) { /* keep as is */ }
    return name;
  }

  // ---------- label sizes ----------
  function customFormats() {
    var list = storeJson(LS.custom) || [];
    return list.filter(function (f) { return f && !F.problems(f).length; }).map(function (f) { f.custom = true; return F.withDefaults(f); });
  }
  function allFormats() { return F.formats.concat(customFormats()); }
  function findFormat(id) { return allFormats().filter(function (f) { return f.id === id; })[0] || null; }
  // A sheet's own copy of its measurements (stored when it was made).
  function formatFromSpec(meta) {
    var json = '';
    for (var i = 1; meta[S.SPEC_PROP + i] !== undefined; i++) json += meta[S.SPEC_PROP + i];
    if (!json) return null;
    try { var f = JSON.parse(json); return F.problems(f).length ? null : F.withDefaults(f); } catch (e) { return null; }
  }
  function sheetLayout(meta, f) {
    try {
      var l = JSON.parse(meta[S.LAYOUT_PROP] || 'null');
      if (l && typeof l.top === 'number') return { margins: { top: l.top, left: l.left }, applied: { x: l.nx || 0, y: l.ny || 0 } };
    } catch (e) { /* older sheet */ }
    // Sheets from before version 1.3 used the first label's position as the margins.
    return { margins: { top: f.firstLabel.top, left: f.firstLabel.left }, applied: { x: 0, y: 0 } };
  }

  // Per-Mac settings
  function feedArrowOn() { return store(LS.feedArrow) !== 'off'; }
  function autoMarkOn() { return store(LS.autoMark) === 'on'; }
  // Labels printed on sheets that already had used labels. The count keeps
  // going while it's hidden, so it's right when it's shown again.
  function savedCount() { var c = Number(store(LS.saved)); return isFinite(c) && c > 0 ? Math.round(c) : 0; }
  function addSaved(n) { store(LS.saved, String(Math.max(0, savedCount() + n))); }
  function showSavedOn() { return store(LS.showSaved) !== 'off'; }
  function tripLimit() { var n = Math.round(Number(store(LS.tripLimit))); return n >= 1 && n <= 20 ? n : 3; }
  function nudgeFor(id) { var v = storeJson(LS.nudge + id); return v && typeof v.x === 'number' ? v : { x: 0, y: 0 }; }
  function trips() { return Number(getSetting(DOC.trips)) || 0; }
  function spots() { var s = getSetting(DOC.spots); return s && typeof s === 'object' ? s : {}; }
  function history() { var h = getSetting(DOC.history); return Array.isArray(h) ? h : []; }

  // Label stock, per size: { on: sheets on hand, pack: sheets per pack,
  // warn: remind at this many, link: where to buy, packs: packs to order }.
  // A sheet comes off the count on its first trip through the printer.
  function stockFor(id) { var s = storeJson(LS.stock + id); return s && typeof s.on === 'number' ? s : null; }
  function stockLow(s) { return !!s && s.on <= (s.warn || 0); }
  function takeFromStock(f) {
    var s = stockFor(f.id);
    if (!s) return false;
    s.on = Math.max(0, s.on - 1);
    storeJson(LS.stock + f.id, s);
    return true;
  }
  function putBackInStock(f) { var s = stockFor(f.id); if (s) { s.on += 1; storeJson(LS.stock + f.id, s); } }
  // "Avery 5160 (laser), 8160 (inkjet)" in a size's description -> product numbers.
  function productNumbers(f) {
    var d = String((f && f.detail) || ''), l = /(\d{3,5}[A-Z]?) \(laser\)/.exec(d), i = /(\d{3,5}[A-Z]?) \(inkjet\)/.exec(d);
    return { laser: l ? l[1] : '', inkjet: i ? i[1] : '' };
  }

  // Sideways 2-per-sheet labels can be called "top" and "bottom" once the
  // person says which half of the physical sheet label 1 printed on.
  function sideApplies(f) { return !!f && f.page.width > f.page.height && f.across === 2 && f.down === 1; }
  function getSide(f) { return sideApplies(f) ? (store(LS.side + f.id) || '') : ''; }
  function labelName(n) {
    var f = st.fmt;
    if (!sideApplies(f)) return 'label ' + n;
    var side = getSide(f);
    if (!side) return 'label ' + n + (n === 1 ? ' (left)' : ' (right)');
    return n === (side === 'top' ? 1 : 2) ? 'top label' : 'bottom label';
  }
  function theName(n) { var nm = labelName(n); return /^label/.test(nm) ? nm : 'the ' + nm; }
  function sideWord(n) { var side = getSide(st.fmt); return side ? (n === (side === 'top' ? 1 : 2) ? 'Top' : 'Bottom') : ''; }

  // ---------- status, errors, confirm ----------
  function say(msg, tone) {
    var e = $('say');
    e.textContent = msg || '';
    e.className = 'say' + (tone === 'warn' ? ' warn' : '');
  }
  function showError(e) {
    var box = $('error');
    if (e && e.userMessage) { say(e.userMessage, 'warn'); box.hidden = true; return; }
    box.textContent = 'Word reported a problem' + (e && e.message ? ': ' + e.message : '') + '. Try again, or close and reopen the panel.';
    box.hidden = false;
    if (window.console) console.error(e, e && e.debugInfo);
  }
  // alt: { label, action } for a second choice next to the main button.
  function ask(msg, yesLabel, action, danger, alt, noLabel) {
    pendingConfirm = action;
    $('confirm-no').textContent = noLabel || 'Cancel';
    pendingAlt = alt ? alt.action : null;
    $('confirm-alt').hidden = !alt;
    $('confirm-alt').textContent = alt ? alt.label : '';
    $('confirm-msg').textContent = msg;
    $('confirm-yes').textContent = yesLabel;
    $('confirm-yes').className = 'btn sm ' + (danger ? 'danger' : 'primary');
    $('confirm').className = 'confirm' + (danger ? ' danger' : '');
    $('confirm').hidden = false;
    try { $('confirm').scrollIntoView({ block: 'nearest' }); } catch (e) { /* old browser */ }
    $('confirm-no').focus();
  }
  function closeAsk() { pendingConfirm = null; pendingAlt = null; $('confirm').hidden = true; }
  // Web pages open in the Mac's own browser.
  function openLink(url) {
    try { if (Office.context.ui && typeof Office.context.ui.openBrowserWindow === 'function') { Office.context.ui.openBrowserWindow(url); return; } } catch (e) { /* fall back */ }
    try { window.open(url, '_blank', 'noopener'); } catch (e) { /* blocked */ }
  }

  // Word tasks run one at a time. If one fails, the error is shown and the
  // caller's follow-up steps don't run.
  function queue(task) {
    var p = chain.then(function () { busy = true; $('error').hidden = true; return task(); });
    chain = p.then(function () { busy = false; }, function (e) { busy = false; showError(e); });
    return p.then(function (v) { return v; }, function () { return new Promise(function () {}); });
  }
  // Same, for tasks that change the document: save afterwards.
  function queueChange(task) { return queue(task).then(function (v) { autosave(); return v; }); }

  // ---------- saving ----------
  var saveTimer = 0, needsFirstSave = false;
  // A document that has never been saved is left alone, because saving it
  // would ask for a name; the panel shows a reminder instead.
  function showSaveHint() { $('save-hint').hidden = !(needsFirstSave && !docUrl()); }
  function autosave() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(function () {
      if (!docUrl()) { needsFirstSave = true; $('save-hint').hidden = false; return; }
      Word.run(function (ctx) { ctx.document.save(); return ctx.sync(); }).catch(function () {});
    }, 700);
  }

  // ---------- opening the panel automatically ----------
  // Word opens a panel by itself only in a document that has the add-in
  // attached and tagged. This document gets tagged when the panel reads it.
  // New sheets are built from scratch, so the panel copies the add-in
  // reference Word wrote into this document and builds it into each new sheet.
  function tagForAutoOpen() { if (getSetting(DOC.autoShow) !== true) setSetting(DOC.autoShow, true); }

  var refState = 'unknown';
  function cachedRef() { return storeJson(LS.addinRef); }
  function readOwnFile() {
    return new Promise(function (resolve, reject) {
      Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, function (res) {
        if (res.status !== Office.AsyncResultStatus.Succeeded) { reject(res.error); return; }
        var file = res.value, count = file.sliceCount, parts = [];
        (function next(i) {
          file.getSliceAsync(i, function (r) {
            if (r.status !== Office.AsyncResultStatus.Succeeded) { file.closeAsync(); reject(r.error); return; }
            parts.push(r.value.data);
            if (i + 1 < count) { next(i + 1); return; }
            file.closeAsync();
            var size = parts.reduce(function (a, p) { return a + p.length; }, 0), out = new Uint8Array(size), off = 0;
            parts.forEach(function (p) { out.set(p, off); off += p.length; });
            resolve(out);
          });
        })(0);
      });
    });
  }
  function findAddinRef() {
    var c = cachedRef();
    if (c) { refState = 'found'; return Promise.resolve(c); }
    if (!window.JSZip || !window.DOMParser) { refState = 'failed: file tools didn’t load'; return Promise.resolve(null); }
    return readOwnFile().then(function (bytes) { return window.JSZip.loadAsync(bytes); }).then(function (zip) {
      var names = Object.keys(zip.files).filter(function (n) { return /^word\/webextensions\/webextension\d*\.xml$/i.test(n); });
      return Promise.all(names.map(function (n) { return zip.file(n).async('string'); }));
    }).then(function (xmls) {
      for (var i = 0; i < xmls.length; i++) {
        var doc = new DOMParser().parseFromString(xmls[i], 'application/xml');
        var r = doc.getElementsByTagNameNS(S.WE_NS, 'reference')[0];
        if (r && (r.getAttribute('id') || '').toLowerCase().indexOf(ADDIN_ID) >= 0) {
          var ref = { id: r.getAttribute('id'), version: r.getAttribute('version') || '1.0.0.0', store: r.getAttribute('store') || '', storeType: r.getAttribute('storeType') || '' };
          storeJson(LS.addinRef, ref);
          refState = 'found';
          return ref;
        }
      }
      refState = 'missing';
      return null;
    }).catch(function (e) {
      refState = 'failed: ' + ((e && e.message) || 'couldn’t read this document');
      return null;
    });
  }

  // ---------- will the text fit? ----------
  var measureCtx = null;
  function measure(text, css) {
    if (!measureCtx) measureCtx = document.createElement('canvas').getContext('2d');
    measureCtx.font = css;
    return measureCtx.measureText(text).width * 0.75;   // CSS px -> points
  }
  function rowHeightIn(f) {
    var minRow = Infinity;
    S.layout(f).rows.forEach(function (r) { if (r.label) minRow = Math.min(minRow, r.h / 1440); });
    return minRow;
  }
  function textBox(f) {
    return {
      width: (f.label.width - f.padding.left - f.padding.right) * 72 - 1,
      height: rowHeightIn(f) * 72 - (f.align === 'top' ? 0.12 * 72 : 4)
    };
  }
  var PARA_PROPS = 'items/text,items/font/name,items/font/size,items/font/bold,items/font/italic,items/alignment,items/spaceBefore,items/spaceAfter,items/leftIndent,items/rightIndent,items/firstLineIndent,items/lineSpacing';
  function paraSpecs(items, f) {
    return items.map(function (p) {
      return {
        text: p.text, font: p.font.name || f.font.name, size: p.font.size || f.font.size, bold: !!p.font.bold, italic: !!p.font.italic,
        align: p.alignment || 'Left', before: p.spaceBefore || 0, after: p.spaceAfter || 0, indent: Math.max(0, p.leftIndent || 0),
        right: Math.max(0, p.rightIndent || 0), first: p.firstLineIndent || 0, lineSpacing: p.lineSpacing || 12
      };
    });
  }

  // ---------- reading the document ----------
  function readState() {
    return Word.run(function (ctx) {
      var props = ctx.document.properties.customProperties;
      props.load('items/key,items/value');
      var tables = ctx.document.body.tables;
      tables.load('items/rowCount');
      var paras = ctx.document.body.paragraphs;
      paras.load('items/tableNestingLevel,items/text');
      var bodyCCs = ctx.document.body.contentControls;
      bodyCCs.load('items/tag');
      var header = ctx.document.sections.getFirst().getHeader('Primary');
      header.load('text');
      return ctx.sync().then(function () {
        var meta = {};
        props.items.forEach(function (p) { meta[p.key] = String(p.value); });
        var id = meta[S.PROP];
        if (!id) return { kind: meta[S.BACKUP_PROP] ? 'backup' : 'none' };
        var fmt = findFormat(id) || formatFromSpec(meta);
        if (!fmt) {
          return { kind: 'broken', fixable: false, reason: 'This sheet uses a label size (“' + id + '”) that isn’t on this Mac anymore. Make a new sheet, or add the size again.' };
        }
        var lay = sheetLayout(meta, fmt);
        var shape = S.tableShape(fmt);
        var spacers = bodyCCs.items.filter(function (cc) { return cc.tag === S.SPACER_TAG; }).length;
        var outside = paras.items.filter(function (p) { return p.tableNestingLevel === 0; });
        var damage = '';
        if (!tables.items.length) damage = 'The labels are missing from this document.';
        else if (tables.items[0].rowCount !== shape.rows) damage = 'Rows of labels were added or removed, so the labels no longer line up with your sheet.';
        else if (outside.length > spacers + 1 || outside.some(function (p) { return p.text.trim(); })) damage = 'There’s text or a blank line outside the labels, which pushes them out of place.';
        else if (paras.items.some(function (p) { return p.tableNestingLevel > 1; })) damage = 'Some labels have a table inside them, which hides their text.';
        var base = { fmt: fmt, margins: lay.margins, applied: lay.applied, header: /▲/.test(header.text || ''), headerOther: !!(header.text || '').trim() && !/▲/.test(header.text || '') };
        if (damage) { base.kind = 'broken'; base.fixable = true; base.reason = damage; return base; }

        var table = tables.items[0];
        table.load('values');
        table.rows.load('items/cellCount');
        var pics = ctx.document.body.inlinePictures; pics.load('items');
        var selCell = ctx.document.getSelection().parentTableCellOrNullObject; selCell.load('rowIndex,cellIndex');
        return ctx.sync().then(function () {
          if (table.rows.items.some(function (r) { return r.cellCount !== shape.cols; })) {
            base.kind = 'broken'; base.fixable = true; base.reason = 'Labels were added, removed or merged in a row, so they no longer line up with your sheet.';
            return base;
          }
          var usedCells = [], picCells = [];
          bodyCCs.items.forEach(function (cc) {
            if ((cc.tag || '').indexOf(USED_TAG) === 0) { var c = cc.parentTableCellOrNullObject; c.load('rowIndex,cellIndex'); usedCells.push(c); }
          });
          pics.items.forEach(function (p) { var c = p.parentTableCellOrNullObject; c.load('rowIndex,cellIndex'); picCells.push(c); });
          return ctx.sync().then(function () {
            var used = {}, ready = {}, texts = {}, picsAt = {};
            usedCells.forEach(function (c) { if (!c.isNullObject) { var n = S.labelAt(fmt, c.rowIndex, c.cellIndex); if (n) used[n] = true; } });
            var vals = table.values || [];
            for (var n = 1; n <= fmt.across * fmt.down; n++) {
              if (used[n]) continue;
              var ci = S.cellOf(fmt, n);
              var t = String((vals[ci.row] && vals[ci.row][ci.cell]) || '').replace(/ /g, ' ').trim();
              if (t) { ready[n] = true; texts[n] = t; }
            }
            picCells.forEach(function (c) {
              if (c.isNullObject) return;
              var k = S.labelAt(fmt, c.rowIndex, c.cellIndex);
              if (k && !used[k]) { ready[k] = true; picsAt[k] = true; texts[k] = texts[k] || '(picture)'; }
            });
            var here = selCell.isNullObject ? 0 : S.labelAt(fmt, selCell.rowIndex, selCell.cellIndex);
            // Check whether text fits on each label that's ready to print.
            var check = keys(ready).filter(function (k) { return !picsAt[k]; });
            var lists = check.map(function (k) {
              var cell = S.cellOf(fmt, k), p = table.getCell(cell.row, cell.cell).body.paragraphs;
              p.load(PARA_PROPS);
              return p;
            });
            return ctx.sync().then(function () {
              var over = {}, specs = {}, box = textBox(fmt);
              check.forEach(function (k, i) {
                try {
                  specs[k] = paraSpecs(lists[i].items, fmt);
                  if (!T.fits(specs[k], box, measure)) over[k] = true;
                } catch (e) { /* skip */ }
              });
              base.kind = 'sheet';
              base.used = used; base.ready = ready; base.texts = texts; base.pics = picsAt; base.over = over; base.here = here; base.specs = specs;
              return base;
            });
          });
        });
      });
    });
  }

  var lastSnap = '';
  function apply(s) {
    st = s;
    if (s.kind === 'sheet') {
      tagForAutoOpen();
      if (!getSetting(DOC.id)) setSetting(DOC.id, uid());
      var snap = JSON.stringify({ u: keys(st.used), r: keys(st.ready).reduce(function (o, n) { o[n] = firstLine(st.texts[n]); return o; }, {}) });
      if (snap !== lastSnap) { lastSnap = snap; setSetting(DOC.snap, JSON.parse(snap)); }
      updateRegistry();
      syncHeader();
      tidyEdges();
      setTimeout(checkFillJob, 300);
      previewSoon();
    } else if (s.kind !== 'broken') st.fmt = null;
    render();
  }
  function refresh() { return queue(function () { return readState().then(apply); }); }

  var selTimer = 0;
  function onSelectionChanged() {
    clearTimeout(selTimer);
    selTimer = setTimeout(function () { if (!busy) refresh(); }, 400);
    docLineSoon();
  }

  // ---------- the "▲ top" mark in the sheet's blank top strip ----------
  function arrowFits() { return !!st.fmt && F.hasStrip(st.fmt) && st.margins && st.margins.top >= S.STRIP_MARGIN - 0.001; }
  function setHeader(on) {
    return Word.run(function (ctx) {
      var h = ctx.document.sections.getFirst().getHeader('Primary');
      h.clear();
      var p = h.paragraphs.getFirst();
      p.spaceAfter = 0;
      if (on) {
        p.insertText(S.ARROW_TEXT, 'Start');
        p.alignment = 'Centered';
        p.spaceBefore = S.ARROW_BEFORE_PT;
        p.font.set({ name: 'Arial', size: S.ARROW_SIZE_PT, bold: true, italic: false, color: '#555555' });
      } else {
        // An empty header paragraph at normal size is taller than the top
        // margin and would push the labels down, so make it tiny.
        p.spaceBefore = 0;
        p.font.size = 1;
      }
      return ctx.sync();
    });
  }
  // Sheets made before version 1.3.1 have no header or footer parts. With top
  // or bottom margins of 0, Word makes room for an empty one anyway, which
  // pushes full-height labels (shipping, full sheet) onto a second page. Once
  // per panel load, shrink an empty header or footer to a 1-point line.
  var edgesChecked = false;
  function tidyEdges() {
    if (edgesChecked || st.kind !== 'sheet') return;
    edgesChecked = true;
    var withHeader = !st.header && !st.headerOther;
    queue(function () {
      return Word.run(function (ctx) {
        var sec = ctx.document.sections.getFirst();
        var parts = [sec.getFooter('Primary')];
        if (withHeader) parts.push(sec.getHeader('Primary'));
        var lists = parts.map(function (h) { var ps = h.paragraphs; ps.load('items/text,items/font/size,items/spaceBefore'); return ps; });
        return ctx.sync().then(function () {
          var changed = false;
          lists.forEach(function (ps, i) {
            var items = ps.items;
            if (items.some(function (p) { return p.text.trim(); })) return;   // someone's own header or footer
            if (items.length === 1 && items[0].font.size <= 2 && !items[0].spaceBefore) return;   // already tiny
            parts[i].clear();
            var p = parts[i].paragraphs.getFirst();
            p.spaceBefore = 0; p.spaceAfter = 0; p.font.size = 1;
            changed = true;
          });
          return changed ? ctx.sync().then(autosave) : null;
        });
      }).catch(function () { /* older Word: leave it */ });
    });
  }

  var headerBusy = false;
  function syncHeader() {
    if (headerBusy || !arrowFits() || st.headerOther) return;
    var want = feedArrowOn();
    if (want === st.header) return;
    headerBusy = true;
    queueChange(function () { return setHeader(want); }).then(function () { st.header = want; headerBusy = false; render(); });
  }

  // ---------- your sheets (kept on this Mac) ----------
  function registry() { return storeJson(LS.sheets) || {}; }
  // A sheet that was never saved can't be opened again once it's closed, so
  // it drops off the list: when its panel closes, or when nothing has been
  // heard from it for a while (an open sheet's panel checks in every minute).
  var UNSAVED_LIFE = 10 * 60000, lastCheckIn = 0;
  function updateRegistry() {
    var id = getSetting(DOC.id);
    if (!id || !st.fmt) return;
    var reg = registry(), open = openList().length, ready = keys(st.ready).length;
    if (!open && !ready) delete reg[id];
    else reg[id] = { id: id, fmtId: st.fmt.id, fmtName: st.fmt.name, file: fileName(), url: docUrl(), open: open, ready: ready, trips: trips(), at: Date.now(), seen: Date.now() };
    storeJson(LS.sheets, reg);
    lastCheckIn = Date.now();
  }
  function checkIn() {
    if (st.kind !== 'sheet') return;
    refreshUrl();
    if (Date.now() - lastCheckIn > 60000) updateRegistry();
  }
  function forgetIfUnsaved() {
    var id = getSetting(DOC.id);
    if (st.kind !== 'sheet' || !id || docUrl()) return;
    var reg = registry();
    if (reg[id]) { delete reg[id]; storeJson(LS.sheets, reg); }
  }
  function sheetEntries(exceptCurrent) {
    var reg = registry(), me = getSetting(DOC.id), now = Date.now(), dropped = false;
    Object.keys(reg).forEach(function (k) {
      var e = reg[k];
      if (e && !e.url && e.id !== me && now - (e.seen || e.at || 0) > UNSAVED_LIFE) { delete reg[k]; dropped = true; }
    });
    if (dropped) storeJson(LS.sheets, reg);
    return Object.keys(reg).map(function (k) { return reg[k]; })
      .filter(function (e) { return e && (e.open > 0 || e.ready > 0) && !(exceptCurrent && e.id === me); })
      .sort(function (a, b) { return a.fmtName < b.fmtName ? -1 : a.fmtName > b.fmtName ? 1 : b.at - a.at; });
  }
  function forgetSheet(id) { var reg = registry(); delete reg[id]; storeJson(LS.sheets, reg); }
  // Opens a saved sheet (Word asks you to allow it first). Resolves true, or
  // false if it didn't open; a sheet whose file is gone comes off the list.
  var lastOpenGone = false;
  function tryOpenSheet(entry) {
    lastOpenGone = false;
    if (!entry.url) return Promise.resolve(false);
    return Word.run(function (ctx) {
      if (typeof ctx.application.openDocument !== 'function') throw new Error('no openDocument');
      ctx.application.openDocument(entry.url);
      return ctx.sync();
    }).then(function () { return true; }, function (e) {
      var why = ((e && e.code) || '') + ' ' + ((e && e.message) || '');
      if (/not.?found|exist|missing|ItemNotFound|FileNotFound/i.test(why)) { lastOpenGone = true; forgetSheet(entry.id); }
      return false;
    });
  }
  function openSheetFile(entry) {
    tryOpenSheet(entry).then(function (ok) {
      if (ok) return;
      render();
      say(lastOpenGone ? 'Word couldn’t find “' + (entry.file || 'that sheet') + '”, so it’s off the list. If you moved it, open it in Word and it comes back.'
        : 'Word didn’t open “' + (entry.file || 'that sheet') + '”. Open it from File › Open Recent, or click Remove if you don’t need it.', 'warn');
    });
  }
  function renderSheetList(ul, entries, showCurrent) {
    ul.innerHTML = '';
    var me = getSetting(DOC.id), limit = tripLimit();
    entries.forEach(function (e) {
      var li = el('li', showCurrent && e.id === me ? 'here' : '');
      li.appendChild(el('b', '', e.fmtName));
      var t = Math.min(e.trips || 0, limit);
      li.appendChild(el('span', 'meta', plural(e.open, 'open label') + (e.ready ? ' · ' + e.ready + ' ready to print' : '') +
        ' · ' + t + ' of ' + limit + ' trips · ' + (showCurrent && e.id === me ? 'this sheet' : e.url ? (e.file || 'saved') + ' · ' + fmtDay(e.at) : 'open in Word, not saved yet')));
      var acts = el('span', 'acts');
      if (!(showCurrent && e.id === me)) {
        var open = el('button', 'btn ghost sm', 'Open'); open.type = 'button';
        open.addEventListener('click', function () { openSheetFile(e); });
        open.hidden = !e.url;
        var x = el('button', 'btn link', 'Remove'); x.type = 'button'; x.title = 'Remove from this list';
        x.addEventListener('click', function () { forgetSheet(e.id); render(); });
        acts.appendChild(open); acts.appendChild(x);
      }
      li.appendChild(acts);
      ul.appendChild(li);
    });
  }

  // ---------- changing the document ----------
  function cellBody(ctx, n) {
    var ci = S.cellOf(st.fmt, n);
    return ctx.document.body.tables.getFirst().getCell(ci.row, ci.cell).body;
  }

  // A used label is emptied and locked with a tagged content control, so it
  // prints nothing and can't be typed in by accident.
  function lockAsUsed(ctx, nums) {
    var made = nums.map(function (n) {
      var body = cellBody(ctx, n);
      body.clear();
      var cc = body.insertText(' ', 'Start').insertContentControl();
      cc.tag = USED_TAG;
      cc.title = 'Used';
      cc.cannotEdit = true;
      cc.cannotDelete = true;
      return cc;
    });
    return ctx.sync().then(function () {
      // On-screen look only (tags and colors never print). Skipped if this Word can't show them.
      made.forEach(function (cc) { cc.appearance = 'Tags'; cc.color = '#7F8C8D'; });
      return ctx.sync().catch(function () {});
    });
  }
  function unlock(ctx, nums) {
    var lists = nums.map(function (n) { var cs = cellBody(ctx, n).contentControls; cs.load('items/tag'); return cs; });
    return ctx.sync().then(function () {
      var mine = [];
      lists.forEach(function (cs) {
        cs.items.forEach(function (cc) {
          if ((cc.tag || '').indexOf(USED_TAG) === 0) { cc.cannotDelete = false; cc.cannotEdit = false; mine.push(cc); }
        });
      });
      return ctx.sync().then(function () { mine.forEach(function (cc) { cc.delete(false); }); return ctx.sync(); });
    });
  }
  function setSpots(nums, rec) {
    var s = spots();
    nums.forEach(function (n) { if (rec) s[n] = rec(n); else delete s[n]; });
    setSetting(DOC.spots, s);
  }
  function markUsed(nums) {
    return Word.run(function (ctx) { return lockAsUsed(ctx, nums); }).then(function () {
      var at = Date.now();
      setSpots(nums, function () { return { at: at, m: 1 }; });
    });
  }
  function markOpen(nums) {
    return Word.run(function (ctx) { return unlock(ctx, nums); }).then(function () { setSpots(nums, null); });
  }

  // A label's content: its paragraphs as OOXML, plus the layout of its last
  // paragraph. (A cell body's own getOoxml() is the whole table row, so the
  // content range is used. Word gives pasted content's last paragraph the
  // layout of the paragraph it lands in, so that layout is put back after.)
  var PARA_LAYOUT = ['alignment', 'leftIndent', 'rightIndent', 'firstLineIndent', 'spaceBefore', 'spaceAfter'];
  function grabContent(body) {
    var ox = body.getRange('Content').getOoxml(), ps = body.paragraphs;
    ps.load('items/' + PARA_LAYOUT.join(',items/'));
    return function () {   // call after ctx.sync()
      var last = ps.items[ps.items.length - 1], lay = {};
      if (last) PARA_LAYOUT.forEach(function (k) { lay[k] = last[k]; });
      return { ooxml: ox.value, count: ps.items.length, last: lay };
    };
  }
  // items: [{ n, c }] where c came from grabContent.
  function putContents(ctx, items) {
    var lists = items.map(function (it) {
      var b = cellBody(ctx, it.n);
      b.clear();
      b.insertOoxml(T.unwrapTable(it.c.ooxml), 'Replace');
      var ps = b.paragraphs; ps.load('items/text');
      return ps;
    });
    return ctx.sync().then(function () {
      items.forEach(function (it, i) {
        var ps = lists[i].items, last = ps[ps.length - 1];
        // Some versions of Word add an empty paragraph after pasted content.
        if (ps.length > it.c.count && last && !last.text) { last.delete(); last = ps[ps.length - 2]; }
        if (last) PARA_LAYOUT.forEach(function (k) { if (it.c.last && it.c.last[k] !== undefined) last[k] = it.c.last[k]; });
      });
      return ctx.sync();
    });
  }

  // Copy one label's content (by number) to other labels.
  function copyFrom(src, targets) {
    return Word.run(function (ctx) {
      var get = grabContent(cellBody(ctx, src));
      return ctx.sync().then(function () {
        var c = get();
        return putContents(ctx, targets.map(function (n) { return { n: n, c: c }; }));
      });
    });
  }
  // The label the cursor is in, if it's ready to print.
  function cursorLabel() {
    return Word.run(function (ctx) {
      var c = ctx.document.getSelection().parentTableCellOrNullObject;
      c.load('rowIndex,cellIndex');
      return ctx.sync().then(function () {
        if (c.isNullObject) throw userError('Click inside a label in the document first.');
        var n = S.labelAt(st.fmt, c.rowIndex, c.cellIndex);
        if (!n) throw userError('The cursor is in the gap between labels. Click inside a label.');
        if (st.used[n]) throw userError(cap(labelName(n)) + ' is used. Click inside the label that has your text.');
        if (!st.ready[n]) throw userError(cap(labelName(n)) + ' is empty. Type your label there first.');
        return n;
      });
    });
  }
  // The label's text and basic formatting, for building new sheets.
  function readLabelParas(n) {
    return Word.run(function (ctx) {
      var p = cellBody(ctx, n).paragraphs;
      p.load('items/text,items/alignment,items/font/name,items/font/size,items/font/bold,items/font/italic');
      return ctx.sync().then(function () {
        return p.items.map(function (q) {
          var a = String(q.alignment || '').toLowerCase();
          return { t: q.text, size: q.font.size || st.fmt.font.size, b: !!q.font.bold, i: !!q.font.italic, font: q.font.name || undefined, align: a === 'centered' ? 'center' : a === 'right' ? 'right' : 'left' };
        });
      });
    });
  }

  // Write lines of text (and optionally a picture) into labels.
  // paras: [{ t, size, b, i, color, align, indent, before }]; pic: { b64, w, h, align } in points
  function writeLabels(nums, paras, pic) {
    return writeEach(nums.map(function (n) { return { n: n, paras: paras, pic: pic }; }));
  }
  // items: [{ n, paras, pic }] — each label gets its own content, in one pass.
  function writeEach(items) {
    var f = st.fmt, topPt = f.align === 'top' ? f.topSpace * 72 : 0;
    var alignOf = function (a) { return a === 'center' ? 'Centered' : a === 'right' ? 'Right' : 'Left'; };
    return Word.run(function (ctx) {
      items.forEach(function (item) {
        var paras = item.paras || [], pic = item.pic;
        var body = cellBody(ctx, item.n);
        body.clear();
        // One slot per paragraph: the picture (if any), then each line of text.
        var slots = (pic ? [{ pic: pic }] : []).concat(paras.map(function (p, i) { return { p: p, i: i }; }));
        // Word for Mac joins paragraphs added at the end of a cell, so new
        // paragraphs go in before the cell's own (last) paragraph, which takes
        // the last slot.
        var cellPara = body.paragraphs.getFirst();
        slots.forEach(function (s, k) {
          var text = s.p ? String(s.p.t || '') : '';
          var para = k === slots.length - 1 ? cellPara : cellPara.insertParagraph(text, 'Before');
          if (para === cellPara && text) para.insertText(text, 'Start');
          para.spaceAfter = 0;
          para.firstLineIndent = 0;
          if (s.pic) {
            var img = para.insertInlinePictureFromBase64(s.pic.b64, 'Start');
            img.width = s.pic.w; img.height = s.pic.h;
            if (s.pic.alt) img.altTextDescription = s.pic.alt;
            para.alignment = alignOf(s.pic.align);
            para.spaceBefore = topPt; para.spaceAfter = 2; para.leftIndent = 0;
            return;
          }
          var p = s.p;
          para.alignment = alignOf(p.align);
          para.spaceBefore = p.before !== undefined ? p.before : (s.i === 0 && !pic ? topPt : 0);
          para.leftIndent = p.indent || 0;
          para.font.set({ name: p.font || f.font.name, size: p.size || f.font.size, bold: !!p.b, italic: !!p.i, color: p.color || '#000000' });
        });
      });
      return ctx.sync();
    });
  }

  // Mark every ready-to-print label used: one more trip for the sheet.
  // A label's text and its look, compact enough to keep in the print history
  // so an older print can be undone. Labels with a picture keep nothing.
  var SNAP_PROPS = 'items/text,items/alignment,items/leftIndent,items/spaceBefore,items/font/name,items/font/size,items/font/bold,items/font/italic,items/font/color';
  function snapParas(items) {
    return items.map(function (q) {
      var a = String(q.alignment || '').toLowerCase(), p = { t: q.text, size: q.font.size || undefined, align: a === 'centered' ? 'center' : a === 'right' ? 'right' : 'left' };
      if (q.font.bold) p.b = true;
      if (q.font.italic) p.i = true;
      if (q.font.name) p.font = q.font.name;
      if (q.font.color && !/^#0{6}$/i.test(q.font.color)) p.color = q.font.color;
      if (q.leftIndent) p.indent = q.leftIndent;
      p.before = q.spaceBefore || 0;
      return p;
    });
  }
  var KEEP_SNAPS = 8;   // newest prints that can be undone with their text
  function markPrinted() {
    var nums = keys(st.ready), firstTrip = trips() === 0, snaps = {}, saved = keys(st.used).length ? nums.length : 0;
    return Word.run(function (ctx) {
      var gets = nums.map(function (n) { return grabContent(cellBody(ctx, n)); });
      var lists = nums.map(function (n) { if (st.pics[n]) return null; var ps = cellBody(ctx, n).paragraphs; ps.load(SNAP_PROPS); return ps; });
      return ctx.sync().then(function () {
        undo = { nums: nums, cells: nums.map(function (n, i) { return { n: n, c: gets[i]() }; }), at: Date.now() };
        nums.forEach(function (n, i) { if (lists[i]) snaps[n] = snapParas(lists[i].items); });
        return lockAsUsed(ctx, nums);
      });
    }).then(function () {
      var at = undo.at, texts = [];
      nums.forEach(function (n) { var t = firstLine(st.texts[n]); if (t && texts.indexOf(t) < 0) texts.push(t); });
      undo.stock = firstTrip && takeFromStock(st.fmt);
      undo.saved = saved;
      if (saved) addSaved(saved);
      var h = history();
      h.unshift({ at: at, nums: nums, text: texts.length ? texts[0] + (texts.length > 1 ? ' and ' + plural(texts.length - 1, 'other') : '') : '', stock: undo.stock || undefined, saved: saved || undefined, labels: snaps });
      h.forEach(function (e, i) { if (i >= KEEP_SNAPS) delete e.labels; });
      setSetting(DOC.history, h.slice(0, 200));
      setSetting(DOC.trips, trips() + 1);
      setSpots(nums, function (n) { return { at: at, t: firstLine(st.texts[n]) }; });
      // Keep the last label for "Reprint last label".
      var last = undo.cells[undo.cells.length - 1];
      if (last) {
        store(LS.last + st.fmt.id, null);
        storeJson(LS.last + st.fmt.id, { v: 2, c: last.c.ooxml.length < 3e6 ? last.c : null, text: firstLine(st.texts[last.n]) || 'picture', at: at });
      }
      return nums;
    });
  }

  function undoPrinted() {
    var u = undo;
    return Word.run(function (ctx) {
      return unlock(ctx, u.nums).then(function () { return putContents(ctx, u.cells); });
    }).then(function () {
      var h = history();
      if (h.length && h[0].at === u.at) { h.shift(); setSetting(DOC.history, h); }
      setSetting(DOC.trips, Math.max(0, trips() - 1));
      if (u.stock) putBackInStock(st.fmt);
      if (u.saved) addSaved(-u.saved);
      setSpots(u.nums, null);
      undo = null;
    });
  }

  // Undoing an older print from the history: its labels must still be used
  // by that print (not reset or marked again since).
  function canUndoPrint(e) {
    var sp = spots();
    return !!(e && e.nums && e.nums.length) && e.nums.every(function (n) { return st.used[n] && sp[n] && sp[n].at === e.at; });
  }
  function undoPrint(e) {
    if (undo && undo.at === e.at) return undoPrinted();
    var items = e.nums.filter(function (n) { return e.labels && e.labels[n]; }).map(function (n) { return { n: n, paras: e.labels[n] }; });
    return Word.run(function (ctx) { return unlock(ctx, e.nums); }).then(function () {
      return items.length ? writeEach(items) : null;
    }).then(function () {
      setSetting(DOC.history, history().filter(function (x) { return x.at !== e.at; }));
      setSetting(DOC.trips, Math.max(0, trips() - 1));
      if (e.stock) putBackInStock(st.fmt);
      if (e.saved) addSaved(-e.saved);
      setSpots(e.nums, null);
      return items.length;
    });
  }

  function reprintLast() {
    var last = storeJson(LS.last + st.fmt.id), n = nextOpen();
    if (!last || last.v !== 2 || !last.c) return Promise.reject(userError('There’s no saved label to reprint yet.'));
    if (!n) return Promise.reject(userError('There are no open labels on this sheet.'));
    return Word.run(function (ctx) { return putContents(ctx, [{ n: n, c: last.c }]); }).then(function () { return n; });
  }

  // ---------- making documents ----------
  function createDoc(f, opts) {
    return S.buildDocx(f, null, opts).then(function (b64) {
      return Word.run(function (ctx) {
        var doc = ctx.application.createDocument(b64);
        return ctx.sync().then(function () { doc.open(); return ctx.sync(); });
      });
    });
  }
  // prefills: one { n: [paras] } object per new sheet
  function createSheets(f, prefills) {
    return findAddinRef().then(function (ref) {
      var list = prefills && prefills.length ? prefills : [null];
      return list.reduce(function (p, pre) {
        return p.then(function () { return createDoc(f, { addinRef: ref, feedArrow: feedArrowOn(), nudge: nudgeFor(f.id), prefill: pre || undefined }); });
      }, Promise.resolve()).then(function () { return !!ref; });
    });
  }
  // Split many labels' paragraphs into new sheets.
  function spillPrefills(f, labels) {
    var per = f.across * f.down, out = [];
    for (var i = 0; i < labels.length; i += per) {
      var pre = {};
      labels.slice(i, i + per).forEach(function (paras, j) { if (paras) pre[j + 1] = paras; });
      out.push(pre);
    }
    return out;
  }
  // ---------- filling your other sheets first ----------
  // When there's more to print than this sheet holds, labels can go on your
  // other part-used sheets of the same size before any new sheet is made.
  // The list waits on this Mac; each sheet's panel fills its open labels,
  // then opens the next sheet. What's left at the end goes on new sheets.
  // job: { id, fmtId, at, items: [paras per label], sheets: [ids], next, done: [ids] }
  var JOB_LIFE = 30 * 60000;
  function partSheets(f) {
    var lim = tripLimit();
    return sheetEntries(true).filter(function (e) { return e.fmtId === f.id && e.open > 0 && (e.trips || 0) < lim && e.url; })
      .sort(function (a, b) { return b.open - a.open; });
  }
  function roomOn(list) { return list.reduce(function (a, e) { return a + e.open; }, 0); }
  function nextJobSheet(job) {
    var reg = registry(), lim = tripLimit();
    for (var i = 0; i < job.sheets.length; i++) {
      var e = reg[job.sheets[i]];
      if (job.done.indexOf(job.sheets[i]) < 0 && e && e.url && e.open > 0 && (e.trips || 0) < lim) return e;
    }
    return null;
  }
  // Sends what's left of the list on: to the next sheet Word can open, or onto
  // new sheets when there's none. Resolves with what happened, in words.
  function passJobOn(f, job) {
    var e = nextJobSheet(job), n = job.items.length;
    if (!e) {
      storeJson(LS.fillJob, null);
      var sheets = Math.ceil(n / (f.across * f.down));
      return queue(function () { return createSheets(f, spillPrefills(f, job.items)); }).then(function () {
        return 'made ' + plural(sheets, 'new sheet') + ' for the other ' + plural(n, 'label') + '. ' + (sheets === 1 ? 'It opened in its own window.' : 'They opened in their own windows.');
      });
    }
    job.next = e.id; job.at = Date.now();
    storeJson(LS.fillJob, job);
    return tryOpenSheet(e).then(function (ok) {
      if (ok) return 'opened “' + (e.file || 'your other sheet') + '” for the other ' + plural(n, 'label') + '; its panel fills them. If the panel doesn’t open there, click Label Sheet on its Home tab.';
      job.done.push(e.id);
      return passJobOn(f, job);
    });
  }
  function startFillJob(f, items, others) {
    var job = { id: uid(), fmtId: f.id, at: Date.now(), items: items, sheets: others.map(function (e) { return e.id; }), next: null, done: [getSetting(DOC.id)] };
    return passJobOn(f, job);
  }
  var jobRunning = false;
  function checkFillJob() {
    if (jobRunning || busy || st.kind !== 'sheet') return;
    var job = storeJson(LS.fillJob);
    if (!job || !Array.isArray(job.items)) return;
    if (Date.now() - job.at > JOB_LIFE) { storeJson(LS.fillJob, null); return; }
    var me = getSetting(DOC.id), f = st.fmt;
    if (!me || job.next !== me || job.fmtId !== f.id) return;
    jobRunning = true;
    var open = openList(), take = job.items.slice(0, open.length), rest = job.items.slice(take.length);
    var items = take.map(function (paras, i) { return { n: open[i], paras: paras }; }), nums = items.map(function (it) { return it.n; });
    // Claim these labels before writing, so no other panel fills them too.
    job.done = (job.done || []).concat([me]);
    job.items = rest; job.next = null; job.at = Date.now();
    storeJson(LS.fillJob, rest.length ? job : null);
    var here = nums.length ? 'Filled ' + labelList(nums) + ' with labels from your list' : 'This sheet had no open labels left';
    var fill = items.length ? queueChange(function () { return writeEach(items); }).then(refresh) : Promise.resolve();
    fill.then(function () {
      if (!rest.length) { say(here + '. That was the last of them.'); return null; }
      return passJobOn(f, job).then(function (msg) { say(here + '. ' + cap(msg)); });
    }).then(function () { jobRunning = false; });
  }
  // The question to ask when labels don't all fit on this sheet.
  // here: labels going on this sheet; extra: [paras] for the rest; what: 'address', 'copy';
  // spots: the rest in their table positions (null = leave that label open), for new sheets.
  function askOverflow(f, here, extra, what, fillHere, done, spots) {
    var count = extra.length, noun = what === 'copy' ? plural(count, 'copy', 'ies').replace('copyies', 'copies') : plural(count, 'address', 'es');
    spots = spots || extra;
    var others = partSheets(f), room = roomOn(others), sheets = Math.ceil(spots.length / total());
    var hereNoun = what === 'copy' ? plural(here, 'copy', 'ies').replace('copyies', 'copies') : plural(here, 'address', 'es');
    var lead = here ? hereNoun + (here === 1 ? ' goes' : ' go') + ' on this sheet. ' : 'This sheet has no open labels. ';
    var makeNew = function () {
      fillHere().then(function () { return queue(function () { return createSheets(f, spillPrefills(f, spots)); }); }).then(function () {
        done('made ' + plural(sheets, 'new sheet') + ' for the other ' + noun + '. ' + (sheets === 1 ? 'It opened in its own window.' : 'They opened in their own windows.'));
      });
    };
    if (!others.length) {
      ask(lead + 'Put the other ' + noun + ' on ' + plural(sheets, 'new sheet') + '? ' + (sheets === 1 ? 'It opens in its own window.' : 'Each opens in its own window.'), (here ? 'Fill and make ' : 'Make ') + plural(sheets, 'sheet'), makeNew);
      return;
    }
    var which = others.length === 1 ? 'Your other sheet of this size has ' : 'Your other ' + others.length + ' sheets of this size have ';
    ask(lead + which + plural(room, 'open label') + '. Put the other ' + noun + ' there first' + (room < count ? ', and the rest on new sheets' : '') + '?',
      'Fill my other sheets first', function () {
        fillHere().then(function () { return startFillJob(f, extra, others); }).then(done);
      }, false, { label: 'Make new sheets', action: makeNew });
  }

  function makeTestPage(f) {
    var nudge = nudgeFor(f.id);
    return createDoc(f, { test: true, nudge: nudge }).then(function () { store(LS.testCheck + f.id, 'pending'); storeJson(LS.testNudge + f.id, nudge); });
  }

  // ---------- repair / rebuild in place ----------
  // Rebuilds the label table (and the spacer above it), keeping used labels
  // and labels that are ready to print. Also applies the printer adjustment.
  function rebuildSheet(f, margins, nudge) {
    var report = { used: 0, ready: 0, lost: 0 };
    return Word.run(function (ctx) {
      var body = ctx.document.body;
      var tables = body.tables; tables.load('items/rowCount');
      var ccs = body.contentControls; ccs.load('items/tag');
      return ctx.sync().then(function () {
        var shape = S.tableShape(f), t = tables.items[0];
        if (t) { t.load('values'); t.rows.load('items/cellCount'); }
        return ctx.sync().then(function () {
          var intact = !!t && t.rowCount === shape.rows && t.rows.items.every(function (r) { return r.cellCount === shape.cols; });
          var snap = getSetting(DOC.snap) || { u: [], r: {} };
          var used = {}, grabbed = [];
          var usedCells = [];
          ccs.items.forEach(function (cc) {
            if ((cc.tag || '').indexOf(USED_TAG) === 0) { var c = cc.parentTableCellOrNullObject; c.load('rowIndex,cellIndex'); usedCells.push(c); }
          });
          return ctx.sync().then(function () {
            if (intact) usedCells.forEach(function (c) { if (!c.isNullObject) { var n = S.labelAt(f, c.rowIndex, c.cellIndex); if (n) used[n] = true; } });
            else (snap.u || []).forEach(function (n) { used[n] = true; });
            // Collect content from cells that have text.
            var vals = t ? (t.values || []) : [];
            if (t) {
              vals.forEach(function (row, ri) {
                (row || []).forEach(function (v, ci) {
                  var text = String(v || '').replace(/ /g, ' ').trim();
                  if (!text) return;
                  var n = intact ? S.labelAt(f, ri, ci) : 0;
                  if (intact && (!n || used[n])) return;
                  grabbed.push({ n: n, text: text, get: grabContent(t.getCell(ri, ci).body) });
                });
              });
            }
            return ctx.sync();
          }).then(function () {
            grabbed.forEach(function (g) { g.c = g.get(); });
            // Without an intact table, match content to labels by its first line.
            var taken = {};
            grabbed.forEach(function (g) {
              if (g.n) { taken[g.n] = true; return; }
              var line = firstLine(g.text);
              Object.keys(snap.r || {}).some(function (k) {
                var n = Number(k);
                if (!taken[n] && !used[n] && snap.r[k] === line) { g.n = n; taken[n] = true; return true; }
                return false;
              });
            });
            grabbed.forEach(function (g) {
              if (g.n) return;
              for (var n = 1; n <= f.across * f.down; n++) if (!used[n] && !taken[n]) { g.n = n; taken[n] = true; return; }
              report.lost++;
            });
            ccs.items.forEach(function (cc) { cc.cannotDelete = false; cc.cannotEdit = false; });
            return ctx.sync();
          }).then(function () {
            body.clear();
            return ctx.sync();
          }).then(function () {
            body.insertOoxml(S.tablePackage(f, { margins: margins, nudge: nudge }), 'Start');
            return ctx.sync();
          }).then(function () {
            // Keep only one (tiny) paragraph after the table.
            var paras = body.paragraphs; paras.load('items/tableNestingLevel,items/text');
            return ctx.sync().then(function () {
              var items = paras.items, lastIn = -1;
              items.forEach(function (p, i) { if (p.tableNestingLevel > 0) lastIn = i; });
              var after = items.filter(function (p, i) { return i > lastIn && p.tableNestingLevel === 0; });
              after.slice(0, -1).forEach(function (p) { if (!p.text) p.delete(); });
              if (after.length) { var lastP = after[after.length - 1]; lastP.font.size = 1; lastP.spaceBefore = 0; lastP.spaceAfter = 0; }
              return ctx.sync();
            });
          }).then(function () {
            var prev = st.fmt; st.fmt = f;
            var usedNums = keys(used);
            return lockAsUsed(ctx, usedNums).then(function () {
              return putContents(ctx, grabbed.filter(function (g) { return g.n; }));
            }).then(function () {
              ctx.document.properties.customProperties.add(S.LAYOUT_PROP, JSON.stringify({ v: 2, top: margins.top, left: margins.left, nx: nudge.x, ny: nudge.y }));
              report.used = usedNums.length;
              report.ready = grabbed.filter(function (g) { return g.n; }).length;
              return ctx.sync();
            }).then(function () { st.fmt = prev; }, function (e) { st.fmt = prev; throw e; });
          });
        });
      });
    }).then(function () { return report; });
  }

  // ---------- noticing prints ----------
  // Word stamps a "last printed" time on the document when it sends it to
  // print (that includes Save as PDF, and a print the printer then drops).
  // The panel checks it every few seconds; when it moves forward, it asks
  // whether the labels that were ready came out, unless the user chose to
  // mark them without asking.
  var lastPrintSeen = null, watchTimer = 0;
  function readPrintTime() {
    return Word.run(function (ctx) {
      var props = ctx.document.properties;
      props.load('lastPrintDate');
      return ctx.sync().then(function () {
        var ms = props.lastPrintDate ? new Date(props.lastPrintDate).getTime() : 0;
        return isFinite(ms) && ms > 0 && new Date(ms).getFullYear() > 1990 ? ms : 0;
      });
    });
  }
  function checkPrinted() {
    checkFillJob();
    checkIn();
    if (busy || st.kind !== 'sheet') return;
    readPrintTime().then(function (ms) {
      if (lastPrintSeen === null) {
        // First check: a print counts as new only if it's later than the last
        // one this document recorded.
        var saved = Number(getSetting(DOC.lastPrint)) || 0;
        lastPrintSeen = saved || ms;
        if (!saved) setSetting(DOC.lastPrint, ms);
        showLastPrint(ms);
        if (!(saved && ms > saved)) return;
      }
      if (ms <= lastPrintSeen) return;
      lastPrintSeen = ms;
      setSetting(DOC.lastPrint, ms);
      showLastPrint(ms);
      onPrinted(ms);
    }).catch(function () { /* best effort; the button still works */ });
  }
  function showLastPrint(ms) { $('print-watch').textContent = ms ? 'Last printed ' + fmtDate(ms) + '.' : 'Not printed yet.'; }
  function onPrinted(ms) {
    queue(function () { return readState().then(apply); }).then(function () {
      var nums = keys(st.ready);
      if (!nums.length) return;
      var when = timeOnly(ms) ? ' at ' + timeOnly(ms) : '';
      var doMark = function (asked) {
        queueChange(markPrinted).then(function (done) {
          return refresh().then(function () {
            var are = done.length === 1 ? ' is' : ' are';
            say(asked ? cap(labelList(done)) + are + ' now marked used.'
              : 'Word sent this sheet to print' + when + ', so ' + labelList(done) + are + ' now marked used.');
          });
        });
      };
      if (autoMarkOn()) doMark(false);
      else ask('Word sent this sheet to print' + when + '. Did ' + labelList(nums) + ' print?', 'Yes, mark used', function () { doMark(true); }, false, null, 'No');
    });
  }
  function startPrintWatch() {
    clearInterval(watchTimer);
    watchTimer = setInterval(checkPrinted, 2500);
    document.addEventListener('visibilitychange', function () { if (!document.hidden) checkPrinted(); });
    setTimeout(checkPrinted, 800);
  }

  // ---------- shipping label files ----------
  var DPI = 300;
  var ship = { doc: null, page: 1, pages: 1, src: null, crop: null, turn: 0, found: false, name: '', waiting: [] };

  // Usable area inside one label, in inches.
  function contentBox(f) {
    var rowH = rowHeightIn(f), top = f.align === 'top' ? f.topSpace : 0.3;
    return { w: f.label.width - f.padding.left - f.padding.right, h: Math.max(0.5, rowH - top - 0.3), top: top, rowH: rowH };
  }

  var pdfLib = null;
  function loadPdfLib() {
    if (window.pdfjsLib) return Promise.resolve(window.pdfjsLib);
    if (pdfLib) return pdfLib;
    var fail = function () { pdfLib = null; throw userError('The PDF reader didn’t load. Check your internet connection and try again.'); };
    pdfLib = new Promise(function (resolve, reject) {
      var s = document.createElement('script');
      s.src = 'vendor/pdfjs/pdf.js';
      s.onload = function () { if (window.pdfjsLib) resolve(window.pdfjsLib); else reject(); };
      s.onerror = reject;
      document.head.appendChild(s);
    }).then(function (lib) { lib.GlobalWorkerOptions.workerSrc = 'vendor/pdfjs/pdf.worker.js'; return lib; }, fail);
    return pdfLib;
  }
  function fileError(e) {
    say((e && e.userMessage) || 'That file couldn’t be opened. Try a PDF, PNG, or JPEG.', 'warn');
    if (window.console && e && !e.userMessage) console.error(e);
  }
  // Reading files: a PDF page or a picture -> a canvas.
  function readFileAs(file, how) {
    return new Promise(function (resolve, reject) {
      var r = new FileReader();
      r.onload = function () { resolve(r.result); };
      r.onerror = function () { reject(userError('That file couldn’t be read.')); };
      if (how === 'buffer') r.readAsArrayBuffer(file); else r.readAsDataURL(file);
    });
  }
  function isPdfFile(file) { return /pdf$/i.test(file.type || '') || /\.pdf$/i.test(file.name || ''); }
  function openPdf(file) {
    return Promise.all([readFileAs(file, 'buffer'), loadPdfLib()]).then(function (a) {
      return a[1].getDocument({ data: new Uint8Array(a[0]), standardFontDataUrl: 'vendor/pdfjs/standard_fonts/', isEvalSupported: false }).promise;
    });
  }
  function pdfPageCanvas(doc, n, dpi) {
    return doc.getPage(n).then(function (page) {
      var one = page.getViewport({ scale: 1 });
      var scale = dpi / 72, maxPx = 14e6;
      if (one.width * one.height * scale * scale > maxPx) scale = Math.sqrt(maxPx / (one.width * one.height));
      var vp = page.getViewport({ scale: scale });
      var c = document.createElement('canvas');
      c.width = Math.floor(vp.width); c.height = Math.floor(vp.height);
      var ctx = c.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
      return page.render({ canvasContext: ctx, viewport: vp }).promise.then(function () { return c; });
    });
  }
  function imageCanvas(file, maxW) {
    return readFileAs(file, 'url').then(function (url) {
      return new Promise(function (resolve, reject) {
        var img = new Image();
        img.onload = function () {
          var k = maxW && img.naturalWidth > maxW ? maxW / img.naturalWidth : 1;
          var c = document.createElement('canvas');
          c.width = Math.round(img.naturalWidth * k); c.height = Math.round(img.naturalHeight * k);
          var x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.drawImage(img, 0, 0, c.width, c.height);
          resolve(c);
        };
        img.onerror = function () { reject(userError('That picture couldn’t be opened. Try a PDF, PNG, or JPEG.')); };
        img.src = url;
      });
    });
  }

  function handleFile(file) {
    if (!file) return;
    say('Opening ' + (file.name || 'the picture') + '…');
    ship.name = file.name || 'Pasted picture';
    if (isPdfFile(file)) {
      openPdf(file).then(function (doc) {
        ship.doc = doc; ship.pages = doc.numPages; ship.page = 1;
        return renderPdfPage();
      }).then(afterSource).catch(fileError);
    } else {
      imageCanvas(file).then(function (c) {
        ship.doc = null; ship.pages = 1; ship.page = 1; ship.src = c;
        afterSource();
      }).catch(fileError);
    }
  }
  function renderPdfPage() { return pdfPageCanvas(ship.doc, ship.page, DPI).then(function (c) { ship.src = c; }); }

  // ---------- several labels at once ----------
  // Labels waiting to go on the sheet, in order: { name, b64, w, h (points), found }.
  function handleFiles(list) {
    var files = Array.prototype.slice.call(list || []).filter(Boolean);
    if (files.length === 1) { handleFile(files[0]); return; }
    if (!files.length) return;
    say('Reading ' + plural(files.length, 'file') + '…');
    var added = 0, guessed = [], failed = [];
    files.reduce(function (p, file) {
      return p.then(function () {
        return labelFromFile(file).then(function (item) {
          ship.waiting.push(item); added++;
          if (!item.found) guessed.push(item.name);
        }, function () { failed.push(file.name || 'a file'); });
      });
    }, Promise.resolve()).then(function () {
      setFillTab('file');
      var msg = 'Added ' + plural(added, 'label') + ' to the waiting list.';
      if (guessed.length) msg += ' No printed border was found in ' + guessed.join(', ') + ', so the whole page was used; check ' + (guessed.length === 1 ? 'it' : 'them') + ' on the sheet before printing.';
      if (failed.length) msg += ' Couldn’t open ' + failed.join(', ') + '.';
      say(msg, guessed.length || failed.length ? 'warn' : '');
    });
  }
  // The first page (of up to four) with a printed label border; otherwise page 1 trimmed.
  function labelFromFile(file) {
    var name = file.name || 'Label';
    if (!isPdfFile(file)) return imageCanvas(file).then(function (c) { var found = findLabel(c); return waitingItem(name, c, found || trimBox(c), !!found); });
    return openPdf(file).then(function (doc) {
      var last = Math.min(doc.numPages, 4);
      var tryPage = function (n) {
        return pdfPageCanvas(doc, n, DPI).then(function (c) {
          var found = findLabel(c);
          if (found) return waitingItem(name, c, found, true);
          if (n < last) return tryPage(n + 1);
          return pdfPageCanvas(doc, 1, DPI).then(function (c1) { return waitingItem(name, c1, trimBox(c1), false); });
        });
      };
      return tryPage(1);
    });
  }
  function waitingItem(name, src, crop, found) {
    var out = outputOf(src, crop, 0);
    return { name: name, b64: out.canvas.toDataURL('image/png').split(',')[1], w: out.wIn * 72, h: out.hIn * 72, found: found };
  }
  // Every page of the open PDF that has a label border.
  function usePdfPages() {
    if (!ship.doc) return;
    var doc = ship.doc, pages = [], n = 0;
    say('Looking for labels on ' + plural(doc.numPages, 'page') + '…');
    var next = function () {
      n++;
      if (n > doc.numPages) return Promise.resolve();
      return pdfPageCanvas(doc, n, DPI).then(function (c) {
        var found = findLabel(c);
        if (found) { ship.waiting.push(waitingItem(ship.name + ' · page ' + n, c, found, true)); pages.push(n); }
        return next();
      });
    };
    next().then(function () {
      renderWaiting();
      if (!pages.length) { say('No page has a printed label border. Go to the page with the label, drag across it, and put it on the sheet.', 'warn'); return; }
      say('Added ' + plural(pages.length, 'label') + ' from page' + (pages.length === 1 ? ' ' : 's ') + ranges(pages) + ' to the waiting list.' + (pages.length < doc.numPages ? ' Pages without a label border were skipped.' : ''));
    }).catch(fileError);
  }
  function renderWaiting() {
    var w = ship.waiting, box = $('waiting'), n = openList().length;
    box.hidden = !w.length;
    if (!w.length) return;
    $('wait-count').textContent = w.length;
    var ul = $('wait-list'); ul.innerHTML = '';
    w.forEach(function (it, i) {
      var li = el('li'), span = el('span', '', it.name + (it.found ? '' : ' (whole page)'));
      var x = el('button', '', '×'); x.type = 'button'; x.setAttribute('aria-label', 'Remove ' + it.name);
      x.addEventListener('click', function () { ship.waiting.splice(i, 1); renderWaiting(); });
      li.appendChild(span); li.appendChild(x); ul.appendChild(li);
    });
    var k = Math.min(w.length, n);
    $('wait-place').textContent = k ? 'Put the next ' + (k === 1 ? 'label' : k) + ' on this sheet' : 'Fresh sheet loaded? Put the next ' + (Math.min(w.length, total()) === 1 ? 'label' : Math.min(w.length, total())) + ' on it';
  }
  function placeWaiting() {
    var open = openList();
    if (!ship.waiting.length) return;
    if (!open.length) {
      if (keys(st.ready).length) { say('Print this sheet first, then click I printed these. After that, load a fresh sheet.', 'warn'); return; }
      ask('This sheet is used up. Loaded a fresh one? All its labels become open again and the trip count starts over, then the next labels go on.', 'Fresh sheet loaded', function () {
        var used = keys(st.used);
        undo = null;
        queueChange(function () { return markOpen(used).then(function () { setSetting(DOC.trips, 0); setSetting(DOC.spots, {}); }); }).then(refresh).then(placeWaiting);
      });
      return;
    }
    var take = ship.waiting.slice(0, open.length);
    var items = take.map(function (it, i) { return { n: open[i], paras: [], pic: { b64: it.b64, w: it.w, h: it.h, align: 'center', alt: 'Shipping label' } }; });
    queueChange(function () { return writeEach(items); }).then(refresh).then(function () {
      ship.waiting = ship.waiting.slice(take.length);
      renderWaiting();
      var names = items.map(function (it) { return theName(it.n); });
      say('Placed ' + plural(take.length, 'label') + ' on ' + names.join(' and ') + '. Print with ⌘P.' +
        (ship.waiting.length ? ' ' + plural(ship.waiting.length, 'more label') + (ship.waiting.length === 1 ? ' is' : ' are') + ' waiting: after printing, load a fresh sheet and click the button again.' : ''));
    });
  }

  function afterSource() {
    ship.turn = 0;
    var found = findLabel(ship.src);
    ship.crop = found || trimBox(ship.src);
    ship.found = !!found;
    setFillTab('file');
    $('crop').hidden = false;
    drawCrop();
    if (ship.pages > 1 && !found) say('This PDF has ' + ship.pages + ' pages. Use the arrows to find the page with the label.');
    else if (found) say('Found the label' + (ship.pages > 1 ? ' on page ' + ship.page + ' of ' + ship.pages : '') + '. Check the preview, then put it on the sheet.');
    else say('Drag across the label if the page has anything else on it.');
  }
  // Smart crop: the label's printed border, if it has one.
  function findLabel(src) {
    var w = Math.min(900, src.width), k = src.width / w, h = Math.max(1, Math.round(src.height / k));
    var c = document.createElement('canvas'); c.width = w; c.height = h;
    var x = c.getContext('2d'); x.drawImage(src, 0, 0, w, h);
    var d = x.getImageData(0, 0, w, h).data, gray = new Uint8Array(w * h);
    for (var i = 0; i < w * h; i++) gray[i] = (d[i * 4] * 299 + d[i * 4 + 1] * 587 + d[i * 4 + 2] * 114) / 1000;
    var r = T.findLabelBorder(gray, w, h);
    if (!r) return null;
    var pad = 3;
    var bx = Math.max(0, Math.floor((r.x - pad) * k)), by = Math.max(0, Math.floor((r.y - pad) * k));
    return { x: bx, y: by, w: Math.min(src.width - bx, Math.ceil((r.w + 2 * pad) * k)), h: Math.min(src.height - by, Math.ceil((r.h + 2 * pad) * k)) };
  }
  // Box around everything that isn't white, with a small margin.
  function trimBox(src) {
    var w = Math.min(500, src.width), k = src.width / w, h = Math.max(1, Math.round(src.height / k));
    var c = document.createElement('canvas'); c.width = w; c.height = h;
    var x = c.getContext('2d'); x.drawImage(src, 0, 0, w, h);
    var d = x.getImageData(0, 0, w, h).data, minX = w, minY = h, maxX = -1, maxY = -1;
    for (var y = 0; y < h; y++) {
      for (var i = 0; i < w; i++) {
        var o = (y * w + i) * 4;
        if (d[o] + d[o + 1] + d[o + 2] < 690) { if (i < minX) minX = i; if (i > maxX) maxX = i; if (y < minY) minY = y; if (y > maxY) maxY = y; }
      }
    }
    if (maxX < 0) return { x: 0, y: 0, w: src.width, h: src.height };
    var m = 4;
    minX = Math.max(0, minX - m); minY = Math.max(0, minY - m); maxX = Math.min(w - 1, maxX + m); maxY = Math.min(h - 1, maxY + m);
    var bx = Math.floor(minX * k), by = Math.floor(minY * k);
    return { x: bx, y: by, w: Math.min(src.width - bx, Math.ceil((maxX - minX + 1) * k)), h: Math.min(src.height - by, Math.ceil((maxY - minY + 1) * k)) };
  }
  function placeBox(cr) {
    var box = $('crop-box');
    box.hidden = false;
    box.style.left = (cr.x / ship.src.width * 100) + '%';
    box.style.top = (cr.y / ship.src.height * 100) + '%';
    box.style.width = (cr.w / ship.src.width * 100) + '%';
    box.style.height = (cr.h / ship.src.height * 100) + '%';
  }
  function drawCrop() {
    if (!ship.src) return;
    var cv = $('crop-canvas'), dpr = window.devicePixelRatio || 1;
    var dispW = $('crop-wrap').clientWidth || 280;
    cv.width = Math.round(dispW * dpr);
    cv.height = Math.round(dispW * ship.src.height / ship.src.width * dpr);
    var x = cv.getContext('2d');
    x.fillStyle = '#fff'; x.fillRect(0, 0, cv.width, cv.height);
    x.drawImage(ship.src, 0, 0, cv.width, cv.height);
    placeBox(ship.crop);
    $('page-nav').hidden = ship.pages < 2;
    $('pg-all').hidden = ship.pages < 2;
    $('pg-label').textContent = 'Page ' + ship.page + ' of ' + ship.pages;
    $('pg-prev').disabled = ship.page <= 1;
    $('pg-next').disabled = ship.page >= ship.pages;
    drawFit();
  }
  // The chosen area, turned to fit the label best, at up to about 2400 px.
  function output() { return outputOf(ship.src, ship.crop, ship.turn); }
  function outputOf(src, cr, turn) {
    var box = contentBox(st.fmt);
    var straight = Math.min(box.w / cr.w, box.h / cr.h), turned = Math.min(box.w / cr.h, box.h / cr.w);
    var rot = ((turned > straight * 1.05 ? 90 : 0) + (turn || 0)) % 360;
    var k = Math.min(1, 2400 / Math.max(cr.w, cr.h));
    var w = Math.round(cr.w * k), h = Math.round(cr.h * k), swap = rot % 180 !== 0;
    var c = document.createElement('canvas');
    c.width = swap ? h : w; c.height = swap ? w : h;
    var x = c.getContext('2d');
    x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height);
    x.translate(c.width / 2, c.height / 2);
    x.rotate(rot * Math.PI / 180);
    x.drawImage(src, cr.x, cr.y, cr.w, cr.h, -w / 2, -h / 2, w, h);
    var fit = Math.min(box.w / c.width, box.h / c.height);
    return { canvas: c, wIn: c.width * fit, hIn: c.height * fit, box: box };
  }
  function drawFit() {
    if (!ship.src || !st.fmt) return;
    var out = output(), f = st.fmt, box = out.box;
    var cv = $('fit-canvas'), pxPerIn = 40;
    cv.width = Math.round(f.label.width * pxPerIn); cv.height = Math.round(box.rowH * pxPerIn);
    var x = cv.getContext('2d');
    x.fillStyle = '#fff'; x.fillRect(0, 0, cv.width, cv.height);
    var left = (f.padding.left + (box.w - out.wIn) / 2) * pxPerIn;
    var top = (f.align === 'top' ? box.top : (box.rowH - out.hIn) / 2) * pxPerIn;
    x.drawImage(out.canvas, left, top, out.wIn * pxPerIn, out.hIn * pxPerIn);
  }

  // ---------- lining up from a scanned test page ----------
  // Scan the printed test page (the whole Letter page, at 100%); the panel
  // finds the label boxes on it and works out the printer adjustment.
  var scanResult = null;
  function scanWorks(f) { return !!f && f.page.height >= f.page.width; }
  function inches(v) { return Math.abs(v).toFixed(2).replace(/^0/, '') + ' in'; }
  function handleScan(file) {
    var f = st.fmt;
    if (!file || !f) return;
    say('Reading the scan…');
    scanResult = null; renderScan();
    var get = isPdfFile(file) ? openPdf(file).then(function (doc) { return pdfPageCanvas(doc, 1, 150); }) : imageCanvas(file, 1700);
    get.then(function (c) {
      var w = c.width, h = c.height, d = c.getContext('2d').getImageData(0, 0, w, h).data, gray = new Uint8Array(w * h);
      for (var i = 0; i < w * h; i++) gray[i] = (d[i * 4] * 299 + d[i * 4 + 1] * 587 + d[i * 4 + 2] * 114) / 1000;
      var made = storeJson(LS.testNudge + f.id) || nudgeFor(f.id);
      // Test pages carry five small squares; older ones only the label boxes.
      var long = Math.max(f.page.width, f.page.height), short = Math.min(f.page.width, f.page.height);
      var r = T.alignFromMarks(T.findMarks(gray, w, h, w / (w > h ? long : short)), w, h, f, made);
      if (!r) {
        var grid = scanWorks(f) ? T.findGrid(gray, w, h) : null;
        if (!grid) { say('The alignment squares weren’t found on the scan. Make a new test page from this panel, print it at 100%, and scan the whole page.', 'warn'); return; }
        r = T.alignFromScan(grid, w, h, f, made);
      }
      if (r.problem) { say(r.problem, 'warn'); return; }
      scanResult = r;
      renderScan();
      try { $('scan-result').scrollIntoView({ block: 'nearest' }); } catch (e) { /* old browser */ }
      say('');
    }).catch(fileError);
  }
  function renderScan() {
    var r = scanResult, box = $('scan-result');
    box.hidden = !r || !st.fmt;
    if (!r) return;
    var e = r.error, small = Math.abs(e.x) < 0.02 && Math.abs(e.y) < 0.02;
    var off = [];
    if (Math.abs(e.x) >= 0.01) off.push(inches(e.x) + (e.x > 0 ? ' too far right' : ' too far left'));
    if (Math.abs(e.y) >= 0.01) off.push(inches(e.y) + (e.y > 0 ? ' too low' : ' too high'));
    $('scan-msg').textContent = small
      ? 'The test page lines up' + (off.length ? ' within 1/50 inch (' + off.join(', ') + ')' : '') + '. No adjustment needed.'
      : 'On the scan, the boxes printed ' + off.join(' and ') + '. The new adjustment is: move right ' + r.nudge.x + ' in, move down ' + r.nudge.y + ' in.';
    $('scan-save').hidden = small;
    $('scan-done').hidden = !small;
  }

  // ---------- addresses ----------
  function book() { var b = storeJson(LS.book); return Array.isArray(b) ? b : []; }
  // The last ten addresses put on a label, newest first: { lines, at }
  function recent() { var r = storeJson(LS.recent); return Array.isArray(r) ? r : []; }
  // x: { id, which, vals, on } — the contact, phones and emails that went with it.
  function noteRecent(lines, x) {
    lines = (lines || []).map(function (l) { return String(l).trim(); }).filter(Boolean);
    if (!lines.length) return;
    var key = lines.join('\n').toLowerCase();
    var r = recent().filter(function (e) { return e.lines.join('\n').toLowerCase() !== key; });
    var item = { lines: lines, at: Date.now() };
    if (x && (x.id || XF.some(function (k) { return x.vals && x.vals[k]; }))) item.x = x;
    r.unshift(item);
    storeJson(LS.recent, r.slice(0, 10));
  }
  function groupsOf(e) { return Array.isArray(e.groups) ? e.groups : []; }
  function allGroups() {
    var seen = {}, out = [];
    book().forEach(function (e) { groupsOf(e).forEach(function (g) { var k = g.toLowerCase(); if (!seen[k]) { seen[k] = true; out.push(g); } }); });
    return out.sort(function (a, b) { return a.localeCompare(b); });
  }
  function inGroup(e, g) { return groupsOf(e).some(function (x) { return x.toLowerCase() === g.toLowerCase(); }); }
  var groupShown = '';         // group the address book list is showing
  // A saved address being changed in the address book list. Kept across
  // redraws so a refresh doesn't lose what's being typed.
  var bookEdit = null;   // { id, text }
  // A saved contact: { id, name, lines, lines2?, phone?, phone2?, email?, email2?,
  // groups?, show? }. lines2 is a second address (a PO box, work...). show is
  // what goes on its labels when that was remembered for it:
  // { addr: 1|2, phone, phone2, email, email2 }.
  function keyOf(lines) { return lines.join('\n').toLowerCase(); }
  // Adds contacts ({ lines, lines2?, phone?... }). One already saved (same
  // address) gets the details it's missing instead. -> { added, updated }
  function addContacts(list) {
    var b = book(), byKey = {}, added = 0, updated = 0;
    b.forEach(function (e) { byKey[keyOf(e.lines)] = e; });
    list.forEach(function (c) {
      var lines = (c.lines || []).map(function (s) { return String(s).trim(); }).filter(Boolean);
      if (!lines.length) return;
      var e = byKey[keyOf(lines)], lines2 = (c.lines2 || []).map(function (s) { return String(s).trim(); }).filter(Boolean);
      if (e) {
        var changed = false;
        XF.forEach(function (k) { if (c[k] && !e[k]) { e[k] = String(c[k]).trim(); changed = true; } });
        if (lines2.length && !e.lines2) { e.lines2 = lines2; changed = true; }
        if (changed) updated++;
        return;
      }
      e = { id: uid(), name: lines[0], lines: lines };
      if (lines2.length) e.lines2 = lines2;
      XF.forEach(function (k) { if (c[k]) e[k] = String(c[k]).trim(); });
      b.push(e); byKey[keyOf(lines)] = e; added++;
    });
    b.sort(function (x, y) { return x.name.localeCompare(y.name); });
    storeJson(LS.book, b);
    return { added: added, updated: updated };
  }
  function contactById(id) { return id ? book().filter(function (e) { return e.id === id; })[0] || null : null; }
  // Phone and email on this sheet's labels unless a contact says otherwise. Kept in the document.
  function sheetOn() { var s = getSetting(DOC.onLabel); return { phone: !!(s && s.phone), email: !!(s && s.email) }; }
  function defaultOn(e) {
    var s = sheetOn(), r = e && e.show;
    return { phone: r ? !!r.phone : s.phone, phone2: !!(r && r.phone2), email: r ? !!r.email : s.email, email2: !!(r && r.email2) };
  }
  // The ticked phones and emails that have something in them, as label lines.
  function extraLines(vals, on) { return XF.filter(function (k) { return on[k] && vals[k]; }).map(function (k) { return String(vals[k]).trim(); }); }
  // A saved contact's label: the address it uses, then its phone and email as chosen.
  function contactLines(e) {
    var base = e.show && e.show.addr === 2 && e.lines2 ? e.lines2 : e.lines;
    return base.concat(extraLines(e, defaultOn(e)));
  }
  // A second address typed without the name gets it, so the label still says who.
  function withName(lines2, name) {
    if (!lines2.length || !name) return lines2;
    return /^(\d|p\.?\s*o\.?\s*box|box\s|#)/i.test(lines2[0]) ? [name].concat(lines2) : lines2;
  }
  function textLines(t) { return String(t || '').split(/\r?\n/).map(function (s) { return s.trim(); }).filter(Boolean); }
  function fromAddress() { return store(LS.from) || getSetting(LS.from) || ''; }
  function saveFrom(v) { store(LS.from, v); setSetting(LS.from, v); }

  var DESIGNS = [
    { id: 'plain', name: 'Plain' },
    { id: 'bold', name: 'Bold name' },
    { id: 'returnto', name: 'Return to' },
    { id: 'logo', name: 'With logo' }
  ];
  function currentDesign() {
    var d = store(LS.design) || 'plain';
    if (d === 'logo' && !logoFits(st.fmt)) d = 'plain';
    return DESIGNS.some(function (x) { return x.id === d; }) ? d : 'plain';
  }
  function logoFits(f) { return !!f && rowHeightIn(f) >= 0.9; }
  // Lines in a design -> paragraphs (and a picture for the logo design).
  function designParas(lines, design, f) {
    var base = f.font.size;
    if (design === 'bold') return { paras: lines.map(function (t, i) { return { t: t, size: i ? base : base + 1, b: !i }; }) };
    if (design === 'returnto') return { paras: [{ t: 'RETURN TO', size: Math.max(5, Math.round(base * 0.6)), b: true, color: '#444444' }].concat(lines.map(function (t) { return { t: t, size: base }; })) };
    if (design === 'logo') {
      var logo = storeJson(LS.logo), paras = lines.map(function (t) { return { t: t, size: Math.max(6, base - 1) }; });
      if (!logo || !logo.b64) return { paras: paras };
      var maxH = Math.min(rowHeightIn(f) * 0.38, 0.6) * 72, maxW = (f.label.width - f.padding.left - f.padding.right) * 72;
      var s = Math.min(maxH / logo.h, maxW / logo.w);
      return { paras: paras, pic: { b64: logo.b64, w: logo.w * s, h: logo.h * s } };
    }
    return { paras: lines.map(function (t) { return { t: t, size: base }; }) };
  }
  // Shipping layout: FROM block at the top, larger SHIP TO block below.
  function shippingParas(fromText, toText, f) {
    var topPt = f.align === 'top' ? f.topSpace * 72 : 0, fromLines = textLines(fromText), paras = [];
    if (fromLines.length) {
      paras.push({ t: 'FROM', size: 8, b: true, before: topPt });
      fromLines.forEach(function (l) { paras.push({ t: l, size: 11 }); });
    }
    paras.push({ t: 'SHIP TO', size: 9, b: true, before: fromLines.length ? 90 : topPt + 90, indent: 36 });
    textLines(toText).forEach(function (l, i) { paras.push({ t: l, size: 16, b: i === 0, indent: 36 }); });
    return paras;
  }

  // ---------- numbers, QR codes and barcodes ----------
  // Settings for the Numbers tab, kept on this Mac: { text, pad, kind, data, next }
  var qrLib = null;
  function loadQrLib() {
    if (window.qrcode) return Promise.resolve();
    if (qrLib) return qrLib;
    qrLib = new Promise(function (resolve, reject) {
      var sc = document.createElement('script');
      sc.src = 'vendor/qrcode-generator/qrcode.js?t=' + (window.LSK_T || Date.now());
      sc.onload = function () { if (window.qrcode) resolve(); else reject(); };
      sc.onerror = reject;
      document.head.appendChild(sc);
    }).catch(function () { qrLib = null; throw userError('The QR code maker didn’t load. Check your internet connection and try again.'); });
    return qrLib;
  }
  function codeSettings() {
    var c = storeJson(LS.codes) || {};
    return { text: typeof c.text === 'string' ? c.text : 'Box {n} of {total}', pad: [0, 2, 3, 4].indexOf(c.pad) >= 0 ? c.pad : 0,
      kind: ['none', 'qr', 'bar'].indexOf(c.kind) >= 0 ? c.kind : 'none', data: c.data || '', next: c.next >= 0 ? c.next : 1 };
  }
  function codesFromForm() {
    var c = codeSettings(), start = Math.round(Number($('num-start').value)), count = Math.round(Number($('num-count').value)), tot = $('num-total').value.trim();
    return {
      text: $('num-text').value, pad: Number($('num-pad').value) || 0, kind: c.kind, data: $('code-data').value.trim(),
      start: start, count: count, total: tot === '' ? start + count - 1 : Math.round(Number(tot))
    };
  }
  // What each label gets: its text and what its code holds.
  function codeLabel(o, i) {
    var n = o.start + i;
    var text = T.numberText(o.text, n, o.total, o.pad).trim();
    return { n: n, text: text, data: o.kind === 'none' ? '' : (o.data ? T.numberText(o.data, n, o.total, o.pad) : text) };
  }
  // The picture for a code, sized for this label size. { canvas, w, h } in points, or { problem }.
  function codePicture(kind, data, f, captionPt) {
    var box = textBox(f), roomH = box.height - (captionPt ? captionPt * 1.3 : 0) - 4;
    if (!data) return { problem: 'Type the text for the code.' };
    if (kind === 'qr') {
      var m;
      try { m = T.qrMatrix(data); } catch (e) { return { problem: 'That’s too much text for a QR code on a label.' }; }
      var side = Math.min(roomH, box.width * 0.62, 1.4 * 72);
      if (side < 0.4 * 72) return { problem: 'These labels are too small for a QR code that phones can read. Use a bigger label, or just numbers.' };
      var q = 2, cells = m.length + 2 * q, px = Math.max(4, Math.ceil(600 / cells)), c = document.createElement('canvas');
      c.width = c.height = cells * px;
      var x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.fillStyle = '#000';
      m.forEach(function (row, r) { row.forEach(function (dark, k) { if (dark) x.fillRect((k + q) * px, (r + q) * px, px, px); }); });
      if (side / cells < 0.012 * 72) return { problem: 'That’s too much text for a QR code this small. Use fewer characters.' };
      return { canvas: c, w: side, h: side };
    }
    var bits = T.code128(data);
    if (!bits) return { problem: 'A barcode can hold only letters, numbers, spaces and common symbols. Use a QR code for other characters.' };
    var quiet = 10, mods = bits.length + 2 * quiet;
    var w = Math.min(box.width, mods * 0.015 * 72), h = Math.min(0.55 * 72, roomH);
    if (w / mods < 0.0075 * 72) return { problem: 'That’s too much text for a barcode this wide. Use fewer characters, or a QR code.' };
    if (h < 0.25 * 72) return { problem: 'These labels are too short for a barcode. Use a bigger label, or just numbers.' };
    var mpx = 4, cv = document.createElement('canvas');
    cv.width = mods * mpx; cv.height = Math.max(20, Math.round(cv.width * h / w));
    var bx = cv.getContext('2d'); bx.fillStyle = '#fff'; bx.fillRect(0, 0, cv.width, cv.height); bx.fillStyle = '#000';
    for (var i = 0; i < bits.length; i++) if (bits.charAt(i) === '1') bx.fillRect((i + quiet) * mpx, 0, mpx, cv.height);
    return { canvas: cv, w: w, h: h };
  }
  // Text size: big for numbers alone, smaller under a code; shrinks to fit the width.
  function codeTextSize(text, f, withCode) {
    var box = textBox(f), size = withCode ? Math.max(6, f.font.size) : Math.max(f.font.size, Math.min(28, Math.floor(box.height * 0.5)));
    while (size > 6 && measure(text, 'bold ' + size + 'pt ' + (f.font.name || 'Arial')) > box.width) size -= 0.5;
    return size;
  }
  function codeItem(o, i, f) {
    var lab = codeLabel(o, i), size = codeTextSize(lab.text, f, o.kind !== 'none');
    var paras = lab.text ? [{ t: lab.text, size: size, b: true, align: 'center' }] : [];
    if (o.kind === 'none') return { paras: paras, label: lab };
    var pic = codePicture(o.kind, lab.data, f, lab.text ? size : 0);
    if (pic.problem) return { problem: pic.problem, label: lab };
    var alt = (o.kind === 'qr' ? 'QR code: ' : 'Barcode: ') + lab.data;
    return { paras: paras, pic: { b64: pic.canvas.toDataURL('image/png').split(',')[1], w: pic.w, h: pic.h, align: 'center', alt: alt }, label: lab, canvas: pic.canvas };
  }
  var codesShown = false;
  function renderCodes() {
    var c = codeSettings(), open = openList().length;
    if (!codesShown) {
      codesShown = true;
      $('num-text').value = c.text; $('num-pad').value = String(c.pad); $('code-data').value = c.data;
      $('num-start').value = c.next; $('num-count').value = open || 1;
    }
    Array.prototype.forEach.call($('code-kind').querySelectorAll('button'), function (b) { b.setAttribute('aria-checked', String(b.dataset.kind === c.kind)); });
    $('code-data-row').hidden = c.kind === 'none';
    var o = codesFromForm();
    $('num-fill').disabled = !open;
    $('num-fill').textContent = open ? 'Fill ' + plural(Math.min(Math.max(1, o.count || 1), open), 'open label') : 'No open labels';
    drawCodePreview(o);
  }
  function drawCodePreview(o) {
    var f = st.fmt, msg = $('code-msg'), cv = $('code-canvas'), pxPerIn = 60;
    var go = function () {
      var it = o.count >= 1 && o.start >= 0 ? codeItem(o, 0, f) : { problem: 'Enter the first number and how many labels to fill.' };
      var rowH = rowHeightIn(f);
      cv.width = Math.round(f.label.width * pxPerIn); cv.height = Math.round(rowH * pxPerIn);
      var x = cv.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, cv.width, cv.height);
      if (it.problem) { msg.textContent = it.problem; msg.className = 'warn'; return; }
      msg.className = 'hint';
      msg.textContent = 'How label ' + it.label.n + ' will look' + (o.count > 1 ? '. The last one says “' + codeLabel(o, o.count - 1).text + '”.' : '.');
      var k = pxPerIn / 72, picH = it.pic ? it.pic.h * k + 2 * k : 0, textH = it.paras.length ? it.paras[0].size * 1.2 * k : 0;
      var top = f.align === 'top' ? f.topSpace * pxPerIn : (cv.height - picH - textH) / 2;
      if (it.pic) x.drawImage(it.canvas, (cv.width - it.pic.w * k) / 2, top, it.pic.w * k, it.pic.h * k);
      if (textH) {
        x.fillStyle = '#000'; x.textAlign = 'center'; x.textBaseline = 'top';
        x.font = 'bold ' + (it.paras[0].size * k) + 'px ' + (f.font.name || 'Arial');
        x.fillText(it.paras[0].t, cv.width / 2, top + picH);
      }
    };
    if (o.kind === 'qr' && !window.qrcode) { msg.textContent = 'Loading the QR code maker…'; loadQrLib().then(function () { drawCodePreview(codesFromForm()); }, function (e) { msg.textContent = e.userMessage; msg.className = 'warn'; }); return; }
    go();
  }
  function saveCodes(extra) {
    var c = codeSettings(), o = codesFromForm();
    c.text = o.text; c.pad = o.pad; c.data = o.data;
    Object.keys(extra || {}).forEach(function (k) { c[k] = extra[k]; });
    storeJson(LS.codes, c);
  }
  function fillCodes() {
    var f = st.fmt, o = codesFromForm(), open = openList();
    if (!(o.start >= 0)) { say('Enter the first number.', 'warn'); $('num-start').focus(); return; }
    if (!(o.count >= 1)) { say('Enter how many labels to fill.', 'warn'); $('num-count').focus(); return; }
    if (!(o.total >= 0)) { say('Enter the total as a whole number, or leave it empty.', 'warn'); return; }
    if (!open.length) { say('There are no open labels on this sheet.', 'warn'); return; }
    var count = Math.min(o.count, open.length);
    var ready = o.kind === 'qr' ? loadQrLib() : Promise.resolve();
    ready.then(function () {
      var items = [];
      for (var i = 0; i < count; i++) {
        var it = codeItem(o, i, f);
        if (it.problem) { say(it.problem, 'warn'); return; }
        items.push({ n: open[i], paras: it.paras, pic: it.pic });
      }
      saveCodes({ next: o.start + count });
      queueChange(function () { return writeEach(items); }).then(refresh).then(function () {
        $('num-start').value = o.start + count;
        say('Filled ' + labelList(items.map(function (x) { return x.n; })) + ' with ' + codeLabel(o, 0).text + (count > 1 ? ' to ' + codeLabel(o, count - 1).text : '') + '.' +
          (o.count > count ? ' The sheet ran out of open labels. To go on, print it, then fill your next sheet; numbering starts at ' + (o.start + count) + '.' : ''));
        renderCodes();
      });
    }, function (e) { say(e.userMessage || 'The QR code maker didn’t load.', 'warn'); });
  }

  // ---------- fit text ----------
  function fitLabels(nums) {
    var f = st.fmt, box = textBox(f), tooLong = [], fitted = [];
    var minPt = Math.max(5, Math.min(7, f.font.size * 0.7));   // smallest text Fit text will use
    return Word.run(function (ctx) {
      var lists = nums.map(function (n) { var p = cellBody(ctx, n).paragraphs; p.load(PARA_PROPS); return p; });
      return ctx.sync().then(function () {
        nums.forEach(function (n, i) {
          var specs = paraSpecs(lists[i].items, f);
          var s = T.fitScale(specs, box, measure);
          if (s !== null && s >= 0.999) return;
          // Don't shrink the main text below a readable size; ask for shorter text instead.
          var main = Math.max.apply(null, specs.map(function (p) { return p.size; }));
          if (s === null || main * s < minPt - 0.01) { tooLong.push(n); return; }
          lists[i].items.forEach(function (p, j) { p.font.size = Math.max(Math.min(specs[j].size, minPt), Math.floor(specs[j].size * s * 2) / 2); });
          fitted.push(n);
        });
        return ctx.sync();
      });
    }).then(function () { return { fitted: fitted, tooLong: tooLong }; });
  }

  // ---------- table window ----------
  var tableDlg = null;
  function openTableWindow() {
    if (!Office.context.ui || !Office.context.ui.displayDialogAsync) { say('This version of Word can’t open the table window. Update Word and try again.', 'warn'); return; }
    if (tableDlg) { say('The table window is already open.'); return; }
    storeJson(LS.tableInfo, { open: openList().length, labels: openList(), name: st.fmt.name, per: total(), onLabel: sheetOn() });
    var url = new URL('table.html', location.href.split('?')[0]).href + '?t=' + Date.now();
    Office.context.ui.displayDialogAsync(url, { height: 70, width: 60, displayInIframe: false }, function (res) {
      if (res.status !== Office.AsyncResultStatus.Succeeded) { say('The table window couldn’t open' + (res.error && res.error.message ? ': ' + res.error.message : '') + '.', 'warn'); return; }
      tableDlg = res.value;
      tableDlg.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
        var msg = null;
        try { msg = JSON.parse(arg.message); } catch (e) { return; }
        if (msg && msg.type === 'fill') { closeTable(); fillFromRows(msg.rows || [], msg.show); }
        else if (msg && msg.type === 'close') closeTable();
      });
      tableDlg.addEventHandler(Office.EventType.DialogEventReceived, function () { tableDlg = null; });
    });
  }
  function closeTable() { try { if (tableDlg) tableDlg.close(); } catch (e) { /* already closed */ } tableDlg = null; }

  // ---------- the preview window ----------
  // Shows the page as it will print: each label's edge (which doesn't print),
  // the used labels shaded, and the text and pictures where Word puts them.
  // While it's open it follows changes to the sheet.
  var previewDlg = null, previewTimer = 0, previewPics = {};
  // A picture from the document, made small enough to keep on this Mac for the window.
  function shrinkPicture(b64, max) {
    var type = /^\/9j\//.test(b64) ? 'jpeg' : /^R0lG/.test(b64) ? 'gif' : 'png';
    return new Promise(function (resolve) {
      var img = new Image();
      img.onload = function () {
        try {
          var w = img.naturalWidth || 1, h = img.naturalHeight || 1, k = Math.min(1, max / Math.max(w, h));
          var c = document.createElement('canvas');
          c.width = Math.max(1, Math.round(w * k)); c.height = Math.max(1, Math.round(h * k));
          var x = c.getContext('2d');
          x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height);
          x.drawImage(img, 0, 0, c.width, c.height);
          resolve(c.toDataURL('image/jpeg', 0.85));
        } catch (e) { resolve(''); }
      };
      img.onerror = function () { resolve(''); };
      img.src = 'data:image/' + type + ';base64,' + b64;
    });
  }
  // Every label that's ready to print: its lines as Word has them, and its picture.
  function previewLabels() {
    var f = st.fmt, nums = keys(st.ready);
    return Word.run(function (ctx) {
      var table = ctx.document.body.tables.getFirst();
      var got = nums.map(function (n) {
        var c = S.cellOf(f, n), body = table.getCell(c.row, c.cell).body;
        var ps = body.paragraphs; ps.load(PARA_PROPS);
        var pics = body.inlinePictures; pics.load('items/width,items/height');
        return { n: n, ps: ps, pics: pics };
      });
      return ctx.sync().then(function () {
        got.forEach(function (g) { g.src = g.pics.items.length ? g.pics.items[0].getBase64ImageSrc() : null; });
        return ctx.sync().then(function () { return got; });
      });
    }).then(function (got) {
      var seen = {};
      return Promise.all(got.map(function (g) {
        var specs = paraSpecs(g.ps.items, f), out = { n: g.n, specs: specs };
        if (!g.src) return out;
        var pic = g.pics.items[0], b64 = String(g.src.value || ''), id = b64.length + ':' + b64.slice(0, 40) + b64.slice(-40);
        var spot = 0;
        for (var i = 0; i < specs.length; i++) { if (!cleanPara(specs[i].text)) { spot = i; break; } }
        var url = previewPics[id] ? Promise.resolve(previewPics[id]) : shrinkPicture(b64, 700);
        return url.then(function (src) {
          seen[id] = previewPics[id] = src;
          if (src && specs[spot]) specs[spot].pic = { src: src, w: pic.width, h: pic.height };
          else out.picture = true;
          return out;
        });
      })).then(function (labels) {
        Object.keys(previewPics).forEach(function (id) { if (!seen[id]) delete previewPics[id]; });
        return labels;
      });
    });
  }
  function previewData() {
    var f = st.fmt;
    return previewLabels().then(function (labels) {
      return {
        at: Date.now(), name: f.name, fmt: {
          page: f.page, label: f.label, firstLabel: f.firstLabel, pitch: f.pitch, across: f.across, down: f.down,
          corner: f.corner || 0, padding: f.padding, align: f.align, font: f.font
        },
        used: keys(st.used), labels: labels, arrow: st.header ? S.ARROW_TEXT : '', arrowPt: { before: S.ARROW_BEFORE_PT, size: S.ARROW_SIZE_PT },
        adjusted: !!(st.applied && (Math.abs(st.applied.x || 0) > 0.001 || Math.abs(st.applied.y || 0) > 0.001))
      };
    });
  }
  // Kept for the window to read. Pictures are left out if they're too big to keep.
  function savePreview(data) {
    try { localStorage.setItem(LS.preview, JSON.stringify(data)); return true; } catch (e) { /* too big */ }
    data.labels.forEach(function (l) { l.specs.forEach(function (p) { if (p.pic) { delete p.pic; l.picture = true; } }); });
    try { localStorage.setItem(LS.preview, JSON.stringify(data)); return true; } catch (e) { return false; }
  }
  function openPreview() {
    if (!Office.context.ui || !Office.context.ui.displayDialogAsync) { say('This version of Word can’t open the preview. Update Word and try again.', 'warn'); return; }
    closePreview();
    queue(function () { return readState().then(apply); }).then(previewData).then(function (data) {
      if (!savePreview(data)) { say('The preview couldn’t be made: this Mac’s storage for the panel is full.', 'warn'); return; }
      var land = st.fmt.page.width > st.fmt.page.height;
      var url = new URL('preview.html', location.href.split('?')[0]).href + '?t=' + Date.now();
      Office.context.ui.displayDialogAsync(url, { height: land ? 75 : 90, width: land ? 70 : 45, displayInIframe: false }, function (res) {
        if (res.status !== Office.AsyncResultStatus.Succeeded) { say('The preview couldn’t open' + (res.error && res.error.message ? ': ' + res.error.message : '') + '.', 'warn'); return; }
        previewDlg = res.value;
        previewDlg.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
          var msg = null;
          try { msg = JSON.parse(arg.message); } catch (e) { return; }
          if (msg && msg.type === 'close') closePreview();
        });
        previewDlg.addEventHandler(Office.EventType.DialogEventReceived, function () { previewDlg = null; });
      });
    });
  }
  function closePreview() { clearTimeout(previewTimer); try { if (previewDlg) previewDlg.close(); } catch (e) { /* already closed */ } previewDlg = null; }
  // The sheet changed while the window is open: give it the new page.
  function previewSoon() {
    if (!previewDlg) return;
    clearTimeout(previewTimer);
    previewTimer = setTimeout(function () {
      if (!previewDlg || st.kind !== 'sheet') return;
      previewData().then(savePreview).catch(function () { /* the window keeps the last page */ });
    }, 700);
  }

  // rows: [[name, company, street, apt, city, state, zip], ...], one per label
  // in order: row 1 goes on the first open label, and so on. An empty row
  // leaves its label open. Rows past the open labels go on new sheets.
  // Rows from the table window: name, company, street, apt, city, state, zip,
  // phone, email. show says whether the phone and email columns go on the labels.
  function fillFromRows(rows, show) {
    show = show || sheetOn();
    var lines = (rows || []).map(function (r) {
      r = r || [];
      var l = T.addressLines({ name: r[0], company: r[1], street: r[2], apt: r[3], city: r[4], state: r[5], zip: r[6] }, { home: homeCountry() });
      return l.length ? l.concat(extraLines({ phone: r[7], email: r[8] }, show)) : l;
    });
    while (lines.length && !lines[lines.length - 1].length) lines.pop();
    var count = lines.filter(function (l) { return l.length; }).length;
    if (!count) { say('The table is empty. Add at least one address.', 'warn'); return; }
    fillAddresses(lines, 'from the table');
  }
  // Put a list of addresses on the open labels, in order (an empty entry
  // leaves a label empty), and offer more sheets for the rest.
  function fillAddresses(lines, from) {
    var f = st.fmt, fromText = fromAddress();
    var design = currentDesign() === 'bold' ? 'bold' : 'plain';
    var parasFor = f.shipping ? function (l) { return shippingParas(fromText, l.join('\n'), f); } : function (l) { return designParas(l, design, f).paras; };
    var open = openList();
    var items = [];
    lines.slice(0, open.length).forEach(function (l, i) { if (l.length) items.push({ n: open[i], paras: parasFor(l) }); });
    var nums = items.map(function (it) { return it.n; });
    var spots = lines.slice(open.length), rest = spots.filter(function (l) { return l.length; });
    var fillHere = function () {
      if (!items.length) return Promise.resolve();
      return queueChange(function () { return writeEach(items); }).then(refresh);
    };
    if (!rest.length) {
      fillHere().then(function () { say('Filled ' + labelList(nums) + ' ' + from + '.'); askShrink(nums); });
      return;
    }
    askOverflow(f, nums.length, rest.map(parasFor), 'address', fillHere, function (msg) {
      say((nums.length ? 'Filled ' + labelList(nums) + ' here and ' : 'Nothing went on this sheet; ') + msg);
    }, spots.map(function (l) { return l.length ? parasFor(l) : null; }));
  }

  // ---------- backing up and restoring settings ----------
  // A backup is a Word document holding this Mac's settings (address book,
  // return address, sizes, adjustments, stock, logo...). Things that only make
  // sense on this Mac, or for a moment, stay out.
  var BACKUP_SKIP = /^lsk\.(tableInfo|tableDraft|addinRef|fillJob|sheets|testCheck\.|testNudge\.|last\.|phone$|preview$)/;
  function backupKeys(withKey) {
    var out = {};
    try {
      for (var i = 0; i < localStorage.length; i++) {
        var k = localStorage.key(i);
        if (k && k.indexOf('lsk.') === 0 && !BACKUP_SKIP.test(k) && (withKey || k !== LS.geoKey)) out[k] = localStorage.getItem(k);
      }
    } catch (e) { /* storage blocked */ }
    return out;
  }
  function toB64(text) { return window.btoa(unescape(encodeURIComponent(text))); }
  function fromB64(b64) { return decodeURIComponent(escape(window.atob(b64))); }
  function backupSummary(data) {
    var parse = function (k) { try { return JSON.parse(data[k] || 'null'); } catch (e) { return null; } };
    var withPrefix = function (prefix) { return Object.keys(data).filter(function (k) { return k.indexOf(prefix) === 0; }); };
    var b = parse(LS.book), sizes = parse(LS.custom);
    var nudges = withPrefix(LS.nudge).filter(function (k) { var v = parse(k); return v && (v.x || v.y); }).length;
    var stocks = withPrefix(LS.stock).length;
    var lines = ['Address book: ' + plural(Array.isArray(b) ? b.length : 0, 'address', 'es'), 'Return address: ' + (data[LS.from] ? 'saved' : 'none')];
    if (Array.isArray(sizes) && sizes.length) lines.push('Your own label sizes: ' + sizes.length);
    lines.push('Printer adjustments: ' + (nudges ? plural(nudges, 'label size') : 'none'));
    if (stocks) lines.push('Label stock and reorder links: ' + plural(stocks, 'label size'));
    if (data[LS.logo]) lines.push('Logo: saved');
    var rc = parse(LS.recent);
    if (Array.isArray(rc) && rc.length) lines.push('Recent addresses: ' + rc.length);
    if (Number(data[LS.saved]) > 0) lines.push('Labels saved count: ' + Number(data[LS.saved]).toLocaleString());
    lines.push('Geoapify key: ' + (data[LS.geoKey] ? 'included' : 'not included'));
    return lines;
  }
  function makeBackup() {
    var withKey = $('backup-key').checked && !!store(LS.geoKey);
    var payload = { app: 'Label Sheet Keeper', v: 1, at: Date.now(), version: VERSION, data: backupKeys(withKey) };
    var lines = ['Made ' + fmtDate(payload.at) + ' with version ' + VERSION + '. It has:'].concat(backupSummary(payload.data).map(function (l) { return '•  ' + l; })).concat(['',
      'To restore, open this file in Word and click Restore these settings in the Label Sheet Keeper panel. Or choose this file with Restore from a file in the panel’s backup settings.',
      'This file holds your addresses' + (withKey ? ' and your Geoapify key' : '') + '. Keep it somewhere private.']);
    return findAddinRef().then(function (ref) {
      return S.buildBackupDocx(null, { data: toB64(JSON.stringify(payload)), lines: lines, addinRef: ref });
    }).then(function (b64) {
      return Word.run(function (ctx) {
        var d = ctx.application.createDocument(b64);
        return ctx.sync().then(function () { d.open(); return ctx.sync(); });
      });
    });
  }
  function backupFromXml(xml) {
    var m = /<(?:[\w.-]+:)?lskBackup\b[^>]*>([A-Za-z0-9+/=\s]+)<\/(?:[\w.-]+:)?lskBackup>/.exec(String(xml || ''));
    if (!m) return null;
    try {
      var p = JSON.parse(fromB64(m[1].replace(/\s+/g, '')));
      return p && p.app === 'Label Sheet Keeper' && p.data && typeof p.data === 'object' ? p : null;
    } catch (e) { return null; }
  }
  function backupFromZip(zip) {
    var names = Object.keys(zip.files).filter(function (n) { return /^customXml\/item\d+\.xml$/i.test(n); });
    return Promise.all(names.map(function (n) { return zip.file(n).async('string'); })).then(function (xmls) {
      for (var i = 0; i < xmls.length; i++) { var p = backupFromXml(xmls[i]); if (p) return p; }
      return null;
    });
  }
  // The backup in the open document: through Word's custom XML parts, or by
  // reading the document's own file.
  function readBackupHere() {
    var viaParts = new Promise(function (resolve, reject) {
      try {
        Office.context.document.customXmlParts.getByNamespaceAsync(S.BACKUP_NS, function (res) {
          if (res.status !== Office.AsyncResultStatus.Succeeded || !res.value || !res.value.length) { reject(); return; }
          res.value[0].getXmlAsync(function (r) {
            var p = r.status === Office.AsyncResultStatus.Succeeded ? backupFromXml(r.value) : null;
            if (p) resolve(p); else reject();
          });
        });
      } catch (e) { reject(e); }
    });
    return viaParts.catch(function () {
      if (!window.JSZip) return null;
      return readOwnFile().then(function (bytes) { return window.JSZip.loadAsync(bytes); }).then(backupFromZip);
    });
  }
  function restoreBackup(p) {
    var keepKey = !p.data[LS.geoKey];
    try {
      var old = [];
      for (var i = 0; i < localStorage.length; i++) {
        var k = localStorage.key(i);
        if (k && k.indexOf('lsk.') === 0 && !BACKUP_SKIP.test(k) && !(keepKey && k === LS.geoKey)) old.push(k);
      }
      old.forEach(function (k2) { localStorage.removeItem(k2); });
      Object.keys(p.data).forEach(function (k3) {
        if (k3.indexOf('lsk.') === 0 && !BACKUP_SKIP.test(k3) && typeof p.data[k3] === 'string') localStorage.setItem(k3, p.data[k3]);
      });
    } catch (e) { throw userError('The settings couldn’t be saved on this Mac.'); }
    codesShown = false;
  }
  function askRestore(p, after) {
    ask('Restore the settings from the backup made ' + fmtDate(p.at) + '? They replace the Label Sheet Keeper settings on this Mac.', 'Restore', function () {
      try { restoreBackup(p); } catch (e) { showError(e); return; }
      if (after) after();
      say('Restored your settings from the backup made ' + fmtDate(p.at) + '.');
    });
  }
  var backupHere = null;   // null: not read yet, false: reading, 'bad', or the backup
  function renderBackupView() {
    show('v-backup');
    if (backupHere === null) {
      backupHere = false;
      $('backup-made').textContent = 'Reading the backup…';
      $('backup-restore').hidden = true;
      readBackupHere().then(function (p) { backupHere = p || 'bad'; renderBackupView(); }, function () { backupHere = 'bad'; renderBackupView(); });
      return;
    }
    if (!backupHere) return;
    var ul = $('backup-sum'), bad = backupHere === 'bad';
    ul.innerHTML = '';
    $('backup-restore').hidden = bad;
    if (bad) { $('backup-made').textContent = 'This backup couldn’t be read here. Open the panel in another document and use Restore from a file.'; return; }
    $('backup-made').textContent = 'Made ' + fmtDate(backupHere.at) + (backupHere.version ? ' with version ' + backupHere.version : '') + '. It has:';
    backupSummary(backupHere.data).forEach(function (l) { ul.appendChild(el('li', '', l)); });
  }

  // ---------- rendering ----------
  function show(view) {
    ['v-loading', 'v-outside', 'v-old', 'v-none', 'v-broken', 'v-sheet', 'v-backup'].forEach(function (id) { $(id).hidden = id !== view; });
  }
  function render() {
    // The backup tools sit in Settings on a sheet, and in their own box elsewhere.
    var slot = st.kind === 'sheet' && !choosing ? $('backup-sheet-slot') : $('backup-none-slot');
    if ($('backup-block').parentNode !== slot) slot.appendChild($('backup-block'));
    $('backup-key-row').hidden = !store(LS.geoKey);
    if (st.kind === 'backup') { renderBackupView(); return; }
    if (st.kind === 'none' || choosing) { renderChooser(); show('v-none'); return; }
    if (st.kind === 'broken') {
      $('broken-msg').textContent = st.reason;
      $('repair-btn').hidden = !st.fixable;
      $('repair-note').hidden = !st.fixable;
      show('v-broken');
      return;
    }
    if (st.kind !== 'sheet') return;
    var f = st.fmt;
    show('v-sheet');
    $('fmt-name').textContent = f.name;
    $('fmt-detail').textContent = f.detail;
    renderMap();
    renderTally();
    $('first-row').hidden = total() <= 2;
    $('first-n').max = String(total());
    renderFill();
    renderPrint();
    showSaveHint();
    var entries = sheetEntries(false);
    $('sheets-count').textContent = entries.length;
    renderSheetList($('sheets-list'), entries, true);
    renderHistory();
    renderStock();
    renderSettings();
    renderAbout();
    renderPhone();
    renderPhoneIn();
  }

  function thumb(f, width) {
    var t = el('span', 'thumb');
    width = width || 30;
    t.style.width = width + 'px';
    t.style.height = Math.round(width * f.page.height / f.page.width) + 'px';
    t.setAttribute('aria-hidden', 'true');
    for (var n = 0; n < f.across * f.down; n++) {
      var r = Math.floor(n / f.across), c = n % f.across, i = document.createElement('i');
      i.style.left = ((f.firstLabel.left + c * f.pitch.across) / f.page.width * 100) + '%';
      i.style.top = ((f.firstLabel.top + r * f.pitch.down) / f.page.height * 100) + '%';
      i.style.width = (f.label.width / f.page.width * 100) + '%';
      i.style.height = (f.label.height / f.page.height * 100) + '%';
      t.appendChild(i);
    }
    return t;
  }

  function renderChooser() {
    var list = $('fmt-list'), fmts = allFormats();
    if (!fmts.some(function (f) { return f.id === chosenId; })) chosenId = fmts.length ? fmts[0].id : '';
    list.innerHTML = '';
    fmts.forEach(function (f) {
      var lab = el('label', 'fmt');
      var inp = document.createElement('input'); inp.type = 'radio'; inp.name = 'fmt'; inp.value = f.id; inp.checked = f.id === chosenId;
      inp.addEventListener('change', function () { chosenId = f.id; });
      var text = el('span');
      text.appendChild(el('b', '', f.name));
      text.appendChild(el('span', 'd', f.detail));
      lab.appendChild(inp); lab.appendChild(thumb(f)); lab.appendChild(text);
      if (f.custom) {
        var x = el('button', 'x', '×'); x.type = 'button'; x.title = 'Remove this size'; x.setAttribute('aria-label', 'Remove ' + f.name);
        x.addEventListener('click', function (e) {
          e.preventDefault();
          ask('Remove the size “' + f.name + '” from this Mac? Sheets you already made with it keep working.', 'Remove size', function () {
            storeJson(LS.custom, (storeJson(LS.custom) || []).filter(function (c) { return c.id !== f.id; }));
            render();
          }, true);
        });
        lab.appendChild(x);
      } else lab.appendChild(el('span'));
      list.appendChild(lab);
    });
    var links = $('tpl-links'); links.innerHTML = '';
    F.formats.forEach(function (f, i) {
      if (i) links.appendChild(document.createTextNode(' · '));
      var a = el('a', '', f.name); a.href = 'templates/' + f.id + '.docx'; a.target = '_blank'; a.rel = 'noopener';
      links.appendChild(a);
    });
    $('fmt-problems').hidden = !F.invalid.length;
    $('fmt-problems').textContent = F.invalid.map(function (x) { return 'The size “' + x.id + '” in formats.js was skipped: ' + x.problems.join('; ') + '.'; }).join(' ');
    $('choose-title').textContent = choosing ? 'Make another sheet' : 'Make a label sheet';
    $('choose-lead').textContent = choosing ? 'It opens as a new Word document.' : 'Choose your labels. The sheet opens as a new Word document.';
    $('back-btn').hidden = !(choosing && st.kind === 'sheet');
    var entries = sheetEntries(true);
    $('finish-first').hidden = !entries.length;
    renderSheetList($('finish-list'), entries, false);
    renderSizeForm();
  }

  var mapBuiltFor = '';
  function renderMap() {
    var f = st.fmt, map = $('map'), sp = spots();
    if (mapBuiltFor !== f.id + total()) {
      map.innerHTML = '';
      map.style.aspectRatio = f.page.width + ' / ' + f.page.height;
      for (var n = 1; n <= total(); n++) {
        var r = Math.floor((n - 1) / f.across), c = (n - 1) % f.across;
        var b = el('button', 'cell'); b.type = 'button'; b.dataset.n = n;
        b.style.left = ((f.firstLabel.left + c * f.pitch.across) / f.page.width * 100) + '%';
        b.style.top = ((f.firstLabel.top + r * f.pitch.down) / f.page.height * 100) + '%';
        b.style.width = (f.label.width / f.page.width * 100) + '%';
        b.style.height = (f.label.height / f.page.height * 100) + '%';
        b.style.borderRadius = f.corner ? (f.corner / f.page.width * 100) + 'cqw' : '0';
        b.innerHTML = '<span class="t"></span><span class="num" aria-hidden="true">' + n + '</span><span class="side" aria-hidden="true"></span>';
        map.appendChild(b);
      }
      mapBuiltFor = f.id + total();
    }
    Array.prototype.forEach.call(map.children, function (b) {
      var n = +b.dataset.n;
      var state = st.used[n] ? 'used' : st.ready[n] ? 'ready' : 'open';
      b.className = 'cell ' + state + (st.over[n] ? ' over' : '') + (st.here === n ? ' here' : '');
      drawCellText(b.querySelector('.t'), state === 'ready' ? n : 0);
      b.querySelector('.side').textContent = sideWord(n);
      var tip = '';
      if (state === 'used') {
        var rec = sp[n];
        tip = rec ? (rec.m ? 'Marked used ' + fmtDate(rec.at) : 'Printed ' + fmtDate(rec.at) + (rec.t ? ' · “' + rec.t + '”' : '')) : 'Used';
      } else if (state === 'ready') tip = st.over[n] ? 'Ready to print, but the text may be cut off' : 'Ready to print';
      else tip = 'Open';
      b.title = tip;
      b.setAttribute('aria-label', cap(labelName(n)) + ', ' + (state === 'ready' ? 'ready to print: ' + firstLine(st.texts[n]) + (st.over[n] ? ', text may be cut off' : '') : tip.toLowerCase()));
    });
  }

  // The label's text where Word puts it: the lines as they're set in the document, to scale.
  function drawCellText(t, n) {
    var specs = n && st.specs && st.specs[n];
    t.className = 't';
    t.removeAttribute('style');
    if (specs) { T.drawLabelText(t, specs, st.fmt); return; }
    t.textContent = n ? String(st.texts[n] || '').split(/[\r\n\u000b]+/).slice(0, 3).join('\n') : '';
  }

  function renderTally() {
    var t = $('tally'); t.innerHTML = '';
    var counts = { open: openList().length, used: keys(st.used).length, ready: keys(st.ready).length };
    [['open', 'open'], ['used', 'used'], ['ready', 'ready to print']].forEach(function (p) {
      var s = el('span');
      s.innerHTML = '<i class="sw ' + p[0] + '"></i><b>' + counts[p[0]] + '</b>';
      s.appendChild(document.createTextNode(p[1]));
      t.appendChild(s);
    });
    var lim = tripLimit(), tr = trips();
    var tripsEl = el('span', 'trips' + (tr >= lim - 1 && tr > 0 ? ' last' : ''), tr + ' of ' + lim + ' trips');
    tripsEl.title = 'Times this sheet has gone through the printer';
    t.appendChild(tripsEl);
    if (showSavedOn()) {
      var c = savedCount(), sheets = Math.floor(c / total());
      var sv = el('span', 'saved', c.toLocaleString() + (c === 1 ? ' label saved' : ' labels saved'));
      sv.title = 'Labels printed on partly used sheets instead of fresh ones, on this Mac' + (sheets ? ': about ' + plural(sheets, 'sheet') + ' of this size' : '') + '. Turn this off in Settings.';
      t.appendChild(sv);
    }
  }

  function renderHere() {
    var box = $('here'), n = st.here;
    box.innerHTML = '';
    if (!n) { box.textContent = 'The cursor isn’t in a label.'; return; }
    box.appendChild(document.createTextNode('The cursor is in '));
    box.appendChild(el('b', '', labelName(n)));
    box.appendChild(document.createTextNode(st.used[n] ? ', which is used.' : st.ready[n] ? (st.over[n] ? '. Its text may be cut off.' : '.') : ', which is empty.'));
  }

  function fillTabs() {
    return st.fmt.shipping ? [['file', 'Label file'], ['addr', 'Address']] : [['copy', 'Copy'], ['addr', 'Address'], ['table', 'Table'], ['codes', 'Numbers']];
  }
  function currentTab() {
    var tabs = fillTabs(), saved = store(st.fmt.shipping ? LS.shipTab : LS.fillTab);
    return tabs.some(function (t) { return t[0] === saved; }) ? saved : tabs[0][0];
  }
  function setFillTab(id) { store(st.fmt.shipping ? LS.shipTab : LS.fillTab, id); if (st.kind === 'sheet') renderFill(); }

  function renderFill() {
    var f = st.fmt, tab = currentTab(), n = nextOpen();
    $('s2-title').textContent = f.shipping ? 'Add a shipping label' : 'Fill labels';
    var tabsEl = $('fill-tabs');
    tabsEl.innerHTML = '';
    fillTabs().forEach(function (t) {
      var b = el('button', '', t[1]); b.type = 'button'; b.setAttribute('role', 'tab');
      b.setAttribute('aria-selected', String(t[0] === tab));
      b.setAttribute('aria-controls', 'pane-' + t[0]);
      b.addEventListener('click', function () { setFillTab(t[0]); });
      tabsEl.appendChild(b);
    });
    ['copy', 'file', 'addr', 'table', 'codes'].forEach(function (id) { $('pane-' + id).hidden = id !== tab; });

    // Shipping sheets: which label is next, and which half is which.
    var applies = sideApplies(f), side = getSide(f);
    $('ship-target').hidden = !f.shipping;
    $('ship-target').textContent = n ? 'Next open: ' + theName(n) + '.' : 'There are no open labels on this sheet. When you load a fresh sheet, click Reset for a fresh sheet.';
    $('side-row').hidden = !applies;
    $('side-q').textContent = side
      ? 'Label 1 prints on the ' + side + ' half of your sheet. Change this if it’s wrong:'
      : 'Print the test page once. Which half of your sheet did Label 1 come out on?';
    $('side-top').setAttribute('aria-pressed', String(side === 'top'));
    $('side-bottom').setAttribute('aria-pressed', String(side === 'bottom'));

    if (tab === 'copy') renderHere();
    if (tab === 'file') { $('ship-place').disabled = !n; if (ship.src) drawFit(); renderWaiting(); }
    if (tab === 'addr') renderAddress();
    if (tab === 'codes') renderCodes();
    if (tab === 'table') {
      $('table-open').disabled = !n;
      $('table-msg').textContent = n ? plural(openList().length, 'open label') + ' on this sheet.' : 'There are no open labels on this sheet.';
    }

    // Text that may not fit
    var over = keys(st.over), warn = $('fit-warn');
    warn.hidden = !over.length || f.shipping;
    warn.innerHTML = '';
    if (over.length) {
      warn.appendChild(document.createTextNode('Text may be cut off on ' + labelList(over) + '. '));
      var fit = el('button', 'btn link inline', over.length === 1 ? 'Fit it' : 'Fit them'); fit.type = 'button';
      fit.addEventListener('click', function () { doFit(over); });
      warn.appendChild(fit);
    }
    // Reprint last label
    var last = storeJson(LS.last + f.id), rp = $('reprint-btn');
    rp.hidden = !(last && last.v === 2 && last.c && n);
    if (!rp.hidden) rp.textContent = 'Reprint last label (“' + last.text + '”)';
  }

  // Recent addresses and the address book, in one dropdown.
  function pickerOptions(select, includeMine) {
    var b = book(), mine = fromAddress(), r = recent(), keep = select.value;
    select.innerHTML = '';
    var any = b.length || r.length || (includeMine && mine);
    var o0 = el('option', '', any ? 'Choose an address…' : 'No saved or recent addresses yet'); o0.value = '';
    select.appendChild(o0);
    if (includeMine && mine) { var om = el('option', '', 'My return address'); om.value = 'mine'; select.appendChild(om); }
    var label = function (lines) { return lines[0] + (lines[1] ? ' · ' + lines[1] : ''); };
    if (r.length) {
      var gr = document.createElement('optgroup'); gr.label = 'Recent';
      r.forEach(function (e, i) { var o = el('option', '', label(e.lines)); o.value = 'recent:' + i; gr.appendChild(o); });
      select.appendChild(gr);
    }
    if (b.length) {
      var gb = document.createElement('optgroup'); gb.label = 'Address book';
      b.forEach(function (e) { var o = el('option', '', label(e.lines)); o.value = e.id; gb.appendChild(o); });
      select.appendChild(gb);
    }
    if (keep && select.querySelector('option[value="' + keep.replace(/"/g, '') + '"]')) select.value = keep;
  }
  function pickedLines(v) {
    if (v === 'mine') return textLines(fromAddress());
    if (/^recent:\d+$/.test(v)) { var r = recent()[Number(v.slice(7))]; return r ? r.lines : null; }
    var e = book().filter(function (k) { return k.id === v; })[0];
    return e ? e.lines : null;
  }

  // ---------- phone and email under the address ----------
  // The address box in use (Address, or Ship to on shipping sheets) and what
  // came with it: the saved contact it was picked from and which address.
  var xs = { id: '', which: 1 };
  function addrBox() { return st.fmt && st.fmt.shipping ? $('to-addr') : $('addr-text'); }
  function xVals() { var o = {}; XF.forEach(function (k) { o[k] = $('x-' + k).value.trim(); }); return o; }
  function xOn() { var o = {}; XF.forEach(function (k) { o[k] = $('x-' + k + '-on').checked; }); return o; }
  function setExtras(vals, on) {
    XF.forEach(function (k) { $('x-' + k).value = (vals && vals[k]) || ''; $('x-' + k + '-on').checked = !!(on && on[k]); });
    renderExtras();
  }
  function clearExtras() { xs = { id: '', which: 1 }; setExtras(null, defaultOn(null)); }
  // The address box's lines plus the ticked phones and emails.
  function boxLines() { return textLines(addrBox().value).concat(extraLines(xVals(), xOn())); }
  function renderExtras() {
    var e = contactById(xs.id), vals = xVals(), on = xOn(), s = sheetOn();
    if (xs.id && !e) xs = { id: '', which: 1 };
    $('x-phone2-row').hidden = !(vals.phone2 || (e && e.phone2));
    $('x-email2-row').hidden = !(vals.email2 || (e && e.email2));
    XF.forEach(function (k) { $('x-' + k + '-on').disabled = !vals[k]; });
    $('addr-which').hidden = !(e && e.lines2);
    $('addr-which-1').setAttribute('aria-checked', String(xs.which !== 2));
    $('addr-which-2').setAttribute('aria-checked', String(xs.which === 2));
    // Offer to keep a changed choice, for the sheet or for the contact.
    $('x-sheet').hidden = !((vals.phone && on.phone !== s.phone) || (vals.email && on.email !== s.email));
    var contactDiff = false;
    if (e) {
      var d = defaultOn(e), r = e.show || {};
      contactDiff = XF.some(function (k) { return (vals[k] || '') !== (e[k] || '') || (vals[k] && on[k] !== d[k]); }) || (xs.which === 2) !== (r.addr === 2);
      $('x-contact').textContent = 'Remember for ' + e.name;
    }
    $('x-contact').hidden = !contactDiff;
    $('x-keep').hidden = $('x-sheet').hidden && $('x-contact').hidden;
  }
  // A choice from the Recent and saved addresses list goes into the address box.
  function loadPicked(v, box) {
    var e = contactById(v), r = /^recent:\d+$/.test(v) ? recent()[Number(v.slice(7))] : null;
    if (r && r.x && contactById(r.x.id)) e = null;
    if (e) {
      xs = { id: e.id, which: e.show && e.show.addr === 2 && e.lines2 ? 2 : 1 };
      box.value = (xs.which === 2 ? e.lines2 : e.lines).join('\n');
      setExtras(e, defaultOn(e));
      return;
    }
    var lines = pickedLines(v);
    if (!lines) return;
    box.value = lines.join('\n');
    if (r && r.x) { xs = { id: contactById(r.x.id) ? r.x.id : '', which: r.x.which === 2 ? 2 : 1 }; setExtras(r.x.vals, r.x.on); } else clearExtras();
  }
  function switchWhich(n) {
    var e = contactById(xs.id);
    if (!e || (n === 2 && !e.lines2) || xs.which === n) return;
    xs.which = n;
    addrBox().value = (n === 2 ? e.lines2 : e.lines).join('\n');
    renderExtras();
  }
  // What to remember in Recent: the address and what went with it.
  function extrasNote() { return { id: xs.id, which: xs.which, vals: xVals(), on: xOn() }; }
  // After addresses go on labels: if text may be cut off, say so and offer to shrink it.
  function askShrink(nums) {
    if (!st.fmt || st.fmt.shipping) return;
    var over = nums.filter(function (n) { return st.over[n]; });
    if (!over.length) return;
    ask('The text may be cut off on ' + labelList(over) + '. Shrink it to fit?', 'Shrink', function () { doFit(over); });
  }
  // ---------- address suggestions in the address boxes ----------
  // Typing a house number and street on a line of an address box suggests
  // addresses (the same ones as the table window); picking one fills that
  // line and the city line under it.
  // picked: the street just chosen, so typing ", Apt 4" after it stays quiet.
  var AS = null, as = { box: null, line: -1, items: [], active: -1, timer: 0, seq: 0, picking: false, picked: null };
  function suggester() {
    if (!AS) AS = T.makeSuggester({ key: function () { return store(LS.geoKey) || ''; }, notice: function (t) { say(t, 'warn'); } });
    return AS;
  }
  function caretLine(box) {
    var v = box.value, pos = box.selectionStart || 0, start = v.lastIndexOf('\n', pos - 1) + 1;
    var end = v.indexOf('\n', pos);
    if (end < 0) end = v.length;
    return { index: v.slice(0, start).split('\n').length - 1, text: v.slice(start, end) };
  }
  // A city line: "Boston, MA 02108", "Toronto ON  M5V 3L9".
  function cityLike(s) { return /,\s*[A-Za-z]{2}\b|\b\d{5}(-\d{4})?\s*$|[A-Za-z]\d[A-Za-z][ -]?\d[A-Za-z]\d\s*$/.test(String(s || '')); }
  function closeAs() { clearTimeout(as.timer); as.seq++; $('asugg').hidden = true; as.box = null; as.items = []; as.active = -1; }
  function showAs(box, line, items) {
    if (!items.length) { closeAs(); return; }
    as.box = box; as.line = line; as.items = items; as.active = -1;
    var list = $('asugg-list');
    list.innerHTML = '';
    items.forEach(function (it, i) {
      var b = el('button', '', it.label);
      b.type = 'button'; b.tabIndex = -1; b.setAttribute('role', 'option'); b.setAttribute('aria-selected', 'false');
      b.appendChild(el('small', '', it.sub));
      b.addEventListener('mousedown', function (e) { e.preventDefault(); pickAs(i); });
      list.appendChild(b);
    });
    $('asugg-credit').textContent = suggester().credit(items);
    var r = box.getBoundingClientRect(), d = $('asugg');
    d.style.left = (r.left + window.scrollX) + 'px';
    d.style.top = (r.bottom + window.scrollY + 2) + 'px';
    d.style.width = r.width + 'px';
    d.hidden = false;
  }
  function markAs(i) {
    as.active = i;
    Array.prototype.forEach.call($('asugg-list').children, function (b, j) { b.setAttribute('aria-selected', j === i ? 'true' : 'false'); });
  }
  function pickAs(i) {
    var it = as.items[i], box = as.box;
    if (!it || !box) return;
    var lines = box.value.split('\n'), at = as.line;
    var addr = T.addressLines({ street: it.street, city: it.city, state: it.state, zip: it.zip, country: it.country }, { home: homeCountry() });
    // Replace the street line, and the city (and country) line under it if there is one.
    var n = 1;
    if (cityLike(lines[at + 1])) { n = 2; if (/^(usa|u\.s\.a\.|united states|canada)$/i.test(String(lines[at + 2] || '').trim())) n = 3; }
    Array.prototype.splice.apply(lines, [at, n].concat(addr));
    box.value = lines.join('\n');
    var pos = lines.slice(0, at).join('\n').length + (at ? 1 : 0) + addr[0].length;
    closeAs();
    as.picked = { box: box, line: at, street: addr[0] };
    as.picking = true;
    box.dispatchEvent(new Event('input', { bubbles: true }));
    as.picking = false;
    box.focus();
    box.setSelectionRange(pos, pos);
  }
  function boxChanged(box) {
    clearTimeout(as.timer);
    if (store(LS.suggest) === 'off') { closeAs(); return; }
    var c = caretLine(box), text = c.text.trim(), S = suggester(), p = as.picked;
    if (p && p.box === box && p.line === c.index && text.indexOf(p.street) === 0) return;
    as.picked = null;
    if (!/^\d/.test(text) || !S.ready(text)) { if (!$('asugg').hidden) closeAs(); return; }
    var next = box.value.split('\n')[c.index + 1], q = cityLike(next) ? text + ', ' + next.trim() : text;
    var seq = ++as.seq, now = S.quick(q);
    // What's already known and still fits shows at once; a fresher answer follows.
    if (now.length) showAs(box, c.index, now);
    as.timer = setTimeout(function () {
      S.ask(q).then(function (items) {
        if (seq === as.seq && document.activeElement === box) showAs(box, c.index, items);
      }).catch(function () { /* the suggester says what happened */ });
    }, S.delay());
  }
  function attachSuggest(box) {
    box.setAttribute('aria-autocomplete', 'list');
    box.setAttribute('aria-controls', 'asugg-list');
    box.addEventListener('input', function () { if (!as.picking) boxChanged(box); });
    // While suggestions are showing: arrows choose, Enter or Tab picks, Escape closes.
    box.addEventListener('keydown', function (e) {
      if (as.box !== box || $('asugg').hidden) return;
      var k = e.key;
      if (k === 'ArrowDown') markAs(Math.min(as.items.length - 1, as.active + 1));
      else if (k === 'ArrowUp' && as.active >= 0) markAs(as.active - 1);
      else if ((k === 'Enter' || k === 'Tab') && as.active >= 0) pickAs(as.active);
      else if (k === 'Escape') closeAs();
      else return;
      e.preventDefault();
      e.stopImmediatePropagation();
    }, true);
    box.addEventListener('blur', function () { setTimeout(function () { if (as.box === box && document.activeElement !== box) closeAs(); }, 150); });
  }

  // ---------- address suggestions while typing on a label in Word ----------
  // Word doesn't let add-ins show anything inside the document, so when the
  // cursor is on a street line of a label, the suggestions show at the
  // bottom of the panel. Picking one fills that line and the city line under
  // it in the document.
  // at: { n: label, index: paragraph in the label, text: the street typed }
  // picked/closed: the street just filled, or the line whose card was closed, so they stay quiet.
  var ds = { timer: 0, seq: 0, at: null, items: [], picked: null, closed: '' };
  function cleanPara(t) { return String(t || '').replace(/[\u0000-\u001f]/g, ' ').replace(/\s+/g, ' ').trim(); }
  function closeDocSugg() { ds.seq++; clearTimeout(ds.timer); ds.at = null; ds.items = []; $('dsugg').hidden = true; }
  function docLineSoon() {
    clearTimeout(ds.timer);
    if (store(LS.suggest) === 'off' || st.kind !== 'sheet' || !st.fmt) { if (!$('dsugg').hidden) closeDocSugg(); return; }
    ds.timer = setTimeout(readDocLine, 150);
  }
  function readDocLine() {
    var seq = ++ds.seq, f = st.fmt;
    return Word.run(function (ctx) {
      var sel = ctx.document.getSelection();
      var cell = sel.parentTableCellOrNullObject; cell.load('rowIndex,cellIndex');
      var para = sel.paragraphs.getFirst(); para.load('text');
      return ctx.sync().then(function () {
        if (cell.isNullObject) return null;
        var n = S.labelAt(f, cell.rowIndex, cell.cellIndex);
        if (!n || st.used[n]) return null;
        var ps = cell.body.paragraphs; ps.load('items/text');
        return ctx.sync().then(function () {
          var lines = ps.items.map(function (p) { return cleanPara(p.text); }), text = cleanPara(para.text);
          var index = lines.indexOf(text);
          return index < 0 ? null : { n: n, index: index, text: text, next: lines[index + 1] || '' };
        });
      });
    }).then(function (at) {
      if (seq !== ds.seq) return;
      var S2 = suggester(), p = ds.picked;
      if (!at || !/^\d/.test(at.text) || !S2.ready(at.text)) { if (!$('dsugg').hidden) closeDocSugg(); return; }
      if (p && p.n === at.n && p.index === at.index && at.text.indexOf(p.street) === 0) return;
      ds.picked = null;
      var key = at.n + ':' + at.index;
      if (ds.closed && ds.closed !== key) ds.closed = '';
      if (ds.closed) return;
      var q = cityLike(at.next) ? at.text + ', ' + at.next : at.text, now = S2.quick(q);
      if (now.length) showDocSugg(at, now);
      ds.timer = setTimeout(function () {
        S2.ask(q).then(function (items) { if (seq === ds.seq) showDocSugg(at, items); }).catch(function () { /* the suggester says what happened */ });
      }, S2.delay());
    }).catch(function () { /* best effort: the label can still be typed by hand */ });
  }
  function showDocSugg(at, items) {
    if (!items.length) { $('dsugg').hidden = true; ds.items = []; return; }
    ds.at = at; ds.items = items;
    $('dsugg-title').textContent = 'Suggestions for ' + labelName(at.n);
    var list = $('dsugg-list');
    list.innerHTML = '';
    items.forEach(function (it, i) {
      var b = el('button', '', it.label);
      b.type = 'button'; b.setAttribute('role', 'option');
      b.appendChild(el('small', '', it.sub));
      b.addEventListener('click', function () { pickDocSugg(i); });
      list.appendChild(b);
    });
    $('dsugg-credit').textContent = suggester().credit(items);
    $('dsugg').hidden = false;
  }
  function pickDocSugg(i) {
    var it = ds.items[i], at = ds.at;
    if (!it || !at) return;
    var addr = T.addressLines({ street: it.street, city: it.city, state: it.state, zip: it.zip, country: it.country }, { home: homeCountry() });
    closeDocSugg();
    queueChange(function () {
      return Word.run(function (ctx) {
        var c = S.cellOf(st.fmt, at.n), table = ctx.document.body.tables.getFirst();
        var ps = table.getCell(c.row, c.cell).body.paragraphs;
        ps.load('items/text,items/font/name,items/font/size,items/spaceBefore');
        return ctx.sync().then(function () {
          var items = ps.items, street = items[at.index];
          // The line changed since the suggestions came up: leave it alone.
          if (!street || cleanPara(street.text) !== at.text) return false;
          // Replace the street line, and the city (and country) line under it if there is one.
          var take = 1, under = items[at.index + 1] ? cleanPara(items[at.index + 1].text) : null;
          if (under !== null && (cityLike(under) || !under)) {
            take = 2;
            if (items[at.index + 2] && /^(usa|u\.s\.a\.|united states|canada)$/i.test(cleanPara(items[at.index + 2].text))) take = 3;
          }
          var slots = items.slice(at.index, at.index + take), font = { name: street.font.name, size: street.font.size };
          var line0 = street, k;
          if (addr.length <= take) {
            addr.forEach(function (line, j) { slots[j].insertText(line, 'Replace'); });
            for (k = addr.length; k < take; k++) slots[k].delete();
          } else if (at.index + take < items.length) {
            slots.forEach(function (p, j) { p.insertText(addr[j], 'Replace'); });
            var after = slots[take - 1];
            addr.slice(take).forEach(function (line) { after = after.insertParagraph(line, 'After'); after.font.set(font); after.spaceBefore = 0; });
          } else {
            // These are the label's last lines. Word for Mac joins a line added after
            // a cell's last line onto it, so new lines go in above the last one,
            // which then takes the last line of the address.
            var lastSlot = slots[take - 1], added = [];
            slots.slice(0, take - 1).forEach(function (p, j) { p.insertText(addr[j], 'Replace'); });
            addr.slice(take - 1, addr.length - 1).forEach(function (line, j) {
              var np = lastSlot.insertParagraph(line, 'Before');
              np.font.set(font);
              np.spaceBefore = j === 0 && take === 1 ? lastSlot.spaceBefore || 0 : 0;
              added.push(np);
            });
            lastSlot.insertText(addr[addr.length - 1], 'Replace');
            if (take === 1) { lastSlot.spaceBefore = 0; line0 = added[0]; }
          }
          try { line0.getRange('End').select(); } catch (e) { /* the cursor stays where it was */ }
          return ctx.sync().then(function () { return true; });
        });
      });
    }).then(function (done) {
      if (!done) { say('That line changed in the meantime, so it was left as it is.', 'warn'); return; }
      ds.picked = { n: at.n, index: at.index, street: addr[0] };
      return refresh().then(function () { askShrink([at.n]); });
    });
  }

  // ---------- choosing contacts to import ----------
  var imp = null;   // { name, list: [contact], picked: {index: true}, filter, skipped }
  function startImport(fileName, list) {
    if (!list.length) {
      imp = null; renderImport();
      say('No contacts with a street address were found in that file.', 'warn');
      return;
    }
    imp = { name: fileName, list: list, picked: {}, filter: '', skipped: list.skipped || 0 };
    $('imp-filter').value = '';
    var sec = $('imp').closest('details');
    if (sec) sec.open = true;
    renderImport();
    try { $('imp').scrollIntoView({ block: 'nearest' }); } catch (e) { /* old browser */ }
    say('Tick the contacts to import, then click Import.');
  }
  function importText(c) { return [c.lines.join(' '), (c.lines2 || []).join(' '), c.phone, c.phone2, c.email, c.email2].join(' ').toLowerCase(); }
  function shownImport() {
    var q = imp ? imp.filter.trim().toLowerCase() : '';
    return imp ? imp.list.map(function (c, i) { return i; }).filter(function (i) { return !q || importText(imp.list[i]).indexOf(q) >= 0; }) : [];
  }
  function renderImport() {
    $('imp').hidden = !imp;
    if (!imp) return;
    var have = {}, ul = $('imp-list');
    book().forEach(function (e) { have[keyOf(e.lines)] = true; });
    ul.innerHTML = '';
    var shown = shownImport();
    shown.forEach(function (i) {
      var c = imp.list[i], li = el('li'), lab = el('label'), cb = document.createElement('input');
      cb.type = 'checkbox'; cb.checked = !!imp.picked[i];
      cb.addEventListener('change', function () { if (cb.checked) imp.picked[i] = true; else delete imp.picked[i]; renderImportCount(); });
      var extra = [c.lines2 ? '2 addresses' : '', c.phone ? c.phone : '', c.email ? c.email : '', have[keyOf(c.lines)] ? 'already saved' : ''].filter(Boolean);
      lab.appendChild(cb);
      lab.appendChild(el('span', '', c.lines[0]));
      lab.appendChild(el('small', '', c.lines.slice(1).join(', ') + (extra.length ? ' · ' + extra.join(' · ') : '')));
      li.appendChild(lab); ul.appendChild(li);
    });
    if (!shown.length) ul.appendChild(el('li', 'none', 'No contacts match.'));
    $('imp-head').textContent = 'Choose contacts to import from ' + imp.name + ' (' + imp.list.length + ')';
    $('imp-note').textContent = imp.skipped ? plural(imp.skipped, 'contact') + ' without a street address ' + (imp.skipped === 1 ? 'was' : 'were') + ' left out.' : '';
    renderImportCount();
  }
  function renderImportCount() {
    var k = imp ? Object.keys(imp.picked).length : 0;
    $('imp-go').textContent = k ? 'Import ' + plural(k, 'contact') : 'Import';
    $('imp-go').disabled = !k;
  }

  function wireExtras() {
    XF.forEach(function (k) {
      $('x-' + k).addEventListener('input', function () {
        // A phone or email typed in: tick it if this sheet (or contact) puts it on labels.
        var on = $('x-' + k + '-on'), v = $('x-' + k).value.trim();
        if (v && on.disabled) on.checked = !!defaultOn(contactById(xs.id))[k];
        renderExtras();
      });
      $('x-' + k + '-on').addEventListener('change', renderExtras);
    });
    $('addr-which-1').addEventListener('click', function () { switchWhich(1); });
    $('addr-which-2').addEventListener('click', function () { switchWhich(2); });
    [$('addr-text'), $('to-addr')].forEach(function (box) {
      box.addEventListener('input', function () { if (!box.value.trim()) clearExtras(); });
    });
    $('x-sheet').addEventListener('click', function () {
      var on = xOn(), vals = xVals(), s = sheetOn();
      var next = { phone: vals.phone ? on.phone : s.phone, email: vals.email ? on.email : s.email };
      setSetting(DOC.onLabel, next);
      renderExtras();
      say('Saved for this sheet: ' + (next.phone ? 'phone' : 'no phone') + ' and ' + (next.email ? 'email' : 'no email') + ' on new labels.');
    });
    $('x-contact').addEventListener('click', function () {
      var b = book(), vals = xVals(), on = xOn(), name = '';
      b.forEach(function (e) {
        if (e.id !== xs.id) return;
        XF.forEach(function (k) { if (vals[k]) e[k] = vals[k]; else delete e[k]; });
        e.show = { addr: xs.which === 2 && e.lines2 ? 2 : 1 };
        XF.forEach(function (k) { e.show[k] = !!(vals[k] && on[k]); });
        name = e.name;
      });
      if (!name) return;
      storeJson(LS.book, b);
      renderAddress();
      say('Saved for ' + name + '. Next time, their phone, email and choices come back.');
    });
  }
  function renderAddress() {
    var f = st.fmt, n = nextOpen();
    $('addr-plain').hidden = !!f.shipping;
    $('addr-ship').hidden = !f.shipping;
    // The address switch goes above the address box in use, phone and email below it.
    var abox = addrBox();
    if (abox.previousElementSibling !== $('addr-which')) abox.parentNode.insertBefore($('addr-which'), abox);
    if (abox.nextElementSibling !== $('addr-extras')) abox.parentNode.insertBefore($('addr-extras'), abox.nextSibling);
    renderExtras();
    renderImport();
    if (f.shipping) {
      pickerOptions($('to-pick'), false);
      if (document.activeElement !== $('from-addr')) $('from-addr').value = fromAddress();
      $('ship-addr-place').disabled = !n;
    } else {
      pickerOptions($('addr-pick'), true);
      var box = $('designs'), d = currentDesign();
      box.innerHTML = '';
      DESIGNS.forEach(function (x) {
        if (x.id === 'logo' && !logoFits(f)) return;
        var b = el('button', 'design'); b.type = 'button'; b.setAttribute('role', 'radio'); b.setAttribute('aria-checked', String(x.id === d));
        var mini = el('span', 'mini');
        if (x.id === 'returnto') mini.appendChild(el('small', '', 'RETURN TO'));
        if (x.id === 'logo') mini.appendChild(el('i'));
        var nm = x.id === 'bold' ? el('b', '', 'Your Name') : document.createTextNode('Your Name');
        mini.appendChild(nm); mini.appendChild(el('br')); mini.appendChild(document.createTextNode('123 Main St'));
        b.appendChild(mini); b.appendChild(document.createTextNode(x.name));
        b.addEventListener('click', function () { store(LS.design, x.id); renderAddress(); });
        box.appendChild(b);
      });
      var logo = storeJson(LS.logo);
      $('logo-row').hidden = d !== 'logo';
      $('logo-label').textContent = logo ? 'Change logo' : 'Choose a logo';
      $('logo-msg').textContent = logo ? 'Logo saved on this Mac.' : 'PNG or JPEG. Saved on this Mac.';
      $('addr-one').disabled = !n; $('addr-all').disabled = !n;
    }
    var b2 = book(), groups = allGroups();
    $('book-count').textContent = b2.length;
    // Groups: a filter above the list, shown once an address has a group.
    if (groupShown && !groups.some(function (g) { return g.toLowerCase() === groupShown.toLowerCase(); })) groupShown = '';
    $('group-row').hidden = !groups.length;
    var gp = $('group-pick'); gp.innerHTML = '';
    var g0 = el('option', '', 'All addresses'); g0.value = ''; gp.appendChild(g0);
    groups.forEach(function (g) {
      var o = el('option', '', g + ' (' + b2.filter(function (e) { return inGroup(e, g); }).length + ')'); o.value = g; gp.appendChild(o);
    });
    gp.value = groupShown;
    $('group-fill').hidden = !groupShown;
    $('group-fill').disabled = !n;
    var dl = $('group-names'); dl.innerHTML = '';
    groups.forEach(function (g) { var o = document.createElement('option'); o.value = g; dl.appendChild(o); });
    if (groupShown) b2 = b2.filter(function (e) { return inGroup(e, groupShown); });
    var ul = $('book-list'); ul.innerHTML = '';
    if (!b2.length) ul.appendChild(el('li', '', 'No saved addresses yet.'));
    b2.forEach(function (e) {
      var li = el('li');
      if (bookEdit && bookEdit.id === e.id) { li.className = 'editing'; li.appendChild(bookEditor(e)); ul.appendChild(li); return; }
      var name = el('button', 'addr', e.lines.join(', ')); name.type = 'button';
      var more = (e.lines2 ? '\nAddress 2: ' + e.lines2.join(', ') : '') + XF.filter(function (k) { return e[k]; }).map(function (k) { return '\n' + e[k]; }).join('');
      name.title = e.lines.join('\n') + more + (groupsOf(e).length ? '\nGroups: ' + groupsOf(e).join(', ') : '') + '\n\nClick to edit';
      name.setAttribute('aria-label', 'Edit ' + e.name);
      name.addEventListener('click', function () {
        bookEdit = { id: e.id, text: e.lines.join('\n'), text2: (e.lines2 || []).join('\n'), groups: groupsOf(e).join(', ') };
        XF.forEach(function (k) { bookEdit[k] = e[k] || ''; });
        renderAddress(); var ta = $('book-edit'); ta.focus();
      });
      var x = el('button', '', '×'); x.type = 'button'; x.setAttribute('aria-label', 'Delete ' + e.name);
      x.addEventListener('click', function () { storeJson(LS.book, book().filter(function (k) { return k.id !== e.id; })); renderAddress(); say('Deleted ' + e.name + ' from the address book.'); });
      li.appendChild(name); li.appendChild(x); ul.appendChild(li);
    });
  }
  function bookEditor(e) {
    var box = el('div', 'book-edit');
    var ta = document.createElement('textarea'); ta.id = 'book-edit'; ta.rows = Math.max(3, e.lines.length + 1); ta.spellcheck = false;
    ta.value = bookEdit.text; ta.setAttribute('aria-label', 'Address for ' + e.name);
    ta.addEventListener('input', function () { bookEdit.text = ta.value; });
    var ta2 = document.createElement('textarea'); ta2.id = 'book-edit-2'; ta2.rows = e.lines2 ? e.lines2.length + 1 : 2; ta2.spellcheck = false;
    ta2.value = bookEdit.text2 || ''; ta2.placeholder = 'Second address, like a PO box (optional)'; ta2.setAttribute('aria-label', 'Second address for ' + e.name);
    ta2.addEventListener('input', function () { bookEdit.text2 = ta2.value; });
    var field = function (k, type, ph) {
      var i = document.createElement('input'); i.type = type; i.id = 'book-edit-' + k; i.placeholder = ph; i.value = bookEdit[k] || '';
      i.spellcheck = false; i.autocomplete = 'off'; i.setAttribute('aria-label', ph + ' for ' + e.name);
      i.addEventListener('input', function () { bookEdit[k] = i.value; });
      return i;
    };
    var phones = el('div', 'pair2'), fp = field('phone', 'tel', 'Phone'), fp2 = field('phone2', 'tel', 'Second phone');
    phones.appendChild(fp); phones.appendChild(fp2);
    var fe = field('email', 'email', 'Email'), fe2 = field('email2', 'email', 'Second email');
    var gi = document.createElement('input'); gi.type = 'text'; gi.id = 'book-edit-groups'; gi.className = 'groups';
    gi.placeholder = 'Groups, like Family, Work'; gi.value = bookEdit.groups; gi.spellcheck = false; gi.autocomplete = 'off';
    gi.setAttribute('aria-label', 'Groups for ' + e.name + ', separated by commas'); gi.setAttribute('list', 'group-names');
    gi.addEventListener('input', function () { bookEdit.groups = gi.value; });
    attachSuggest(ta); attachSuggest(ta2);
    [ta, ta2, fp, fp2, fe, fe2, gi].forEach(function (x) { x.addEventListener('keydown', function (k) { if (k.key === 'Escape') { bookEdit = null; renderAddress(); } }); });
    [fp, fp2, fe, fe2].forEach(function (x) { x.addEventListener('keydown', function (k) { if (k.key === 'Enter') { k.preventDefault(); saveBookEdit(e.id); } }); });
    gi.addEventListener('keydown', function (k) { if (k.key === 'Enter') { k.preventDefault(); saveBookEdit(e.id); } });
    var row = el('div', 'row');
    var save = el('button', 'btn primary sm', 'Save'); save.type = 'button';
    save.addEventListener('click', function () { saveBookEdit(e.id); });
    var cancel = el('button', 'btn ghost sm', 'Cancel'); cancel.type = 'button';
    cancel.addEventListener('click', function () { bookEdit = null; renderAddress(); });
    row.appendChild(save); row.appendChild(cancel);
    [ta, ta2, phones, fe, fe2, gi, row].forEach(function (x) { box.appendChild(x); });
    return box;
  }
  function parseGroups(text) {
    var seen = {}, out = [];
    String(text || '').split(/[,;]/).forEach(function (g) {
      g = g.replace(/\s+/g, ' ').trim().slice(0, 40);
      var k = g.toLowerCase();
      if (g && !seen[k]) {
        // Use the spelling already in the address book.
        var have = allGroups().filter(function (x) { return x.toLowerCase() === k; })[0];
        seen[k] = true; out.push(have || g);
      }
    });
    return out;
  }
  function saveBookEdit(id) {
    var lines = textLines(bookEdit && bookEdit.text);
    if (!lines.length) { say('Type the address, or click × to delete it.', 'warn'); return; }
    var key = lines.join('\n').toLowerCase(), b = book(), groups = parseGroups(bookEdit.groups);
    var same = b.filter(function (k) { return k.id !== id && k.lines.join('\n').toLowerCase() === key; })[0];
    var setGroups = function (k, g) { if (g.length) k.groups = g; else delete k.groups; };
    if (same) {
      b = b.filter(function (k) { return k.id !== id; });
      setGroups(same, parseGroups(groupsOf(same).concat(groups).join(',')));
      XF.forEach(function (f) { var v = String(bookEdit[f] || '').trim(); if (v && !same[f]) same[f] = v; });
      var l2 = withName(textLines(bookEdit.text2), lines[0]);
      if (l2.length && !same.lines2) same.lines2 = l2;
    } else b.forEach(function (k) {
      if (k.id !== id) return;
      k.lines = lines; k.name = lines[0]; setGroups(k, groups);
      var lines2 = withName(textLines(bookEdit.text2), lines[0]);
      if (lines2.length) k.lines2 = lines2; else { delete k.lines2; if (k.show) k.show.addr = 1; }
      XF.forEach(function (f) { var v = String(bookEdit[f] || '').trim(); if (v) k[f] = v; else delete k[f]; });
    });
    b.sort(function (x, y) { return x.name.localeCompare(y.name); });
    storeJson(LS.book, b);
    bookEdit = null;
    renderAddress();
    say(same ? 'That address was already in the address book, so the two are now one.' : 'Saved the changes to ' + lines[0] + '.');
  }

  function renderPrint() {
    var f = st.fmt, ready = keys(st.ready).length, open = openList().length, tr = trips(), lim = tripLimit();
    // Feeding guide
    var side = getSide(f), text;
    if (sideApplies(f)) {
      text = side ? 'Feed the sheet print side up. Label 1 prints on the ' + side + ' half. Choose Labels as the paper type.'
        : 'Feed the sheet print side up and choose Labels as the paper type. Print a test page to see which half is label 1.';
    } else {
      text = 'Feed the sheet top edge first, print side up, and choose Labels as the paper type.' +
        (arrowFits() && feedArrowOn() ? ' The small ▲ in the top strip marks the top.' : '');
    }
    $('feed-text').textContent = text;
    // Avery's one-pass note: shown before a sheet's second trip.
    $('warn-pass').hidden = !(tr >= 1 && ready > 0 && store(LS.warnPass) !== '1');
    // After a test page
    $('test-check').hidden = store(LS.testCheck + f.id) !== 'pending';
    // Last trip: fill the rest so nothing goes to waste
    var lastTrip = !f.shipping && tr === lim - 1 && open > 0 && ready > 0;
    $('use-up').hidden = !lastTrip;
    if (lastTrip) $('use-up-msg').textContent = 'This is this sheet’s last trip through the printer. Fill the ' + plural(open, 'open label') + ' too, so nothing goes to waste?';
    $('use-up-mine').hidden = !fromAddress();
    // Trip limit reached
    $('trips-out').hidden = tr < lim;
    $('trips-out').textContent = 'This sheet has gone through the printer ' + plural(tr, 'time') + ', your limit. Use the rest of its labels by hand, or print anyway at your own risk.';
    $('printed-btn').textContent = ready ? 'I printed these (' + ready + ')' : 'I printed these';
    $('printed-btn').disabled = !ready;
    $('undo-row').hidden = !undo;
    if (undo) $('undo-msg').textContent = cap(labelList(undo.nums)) + (undo.nums.length === 1 ? ' was' : ' were') + ' marked used.';
    // Lining up from a scan
    renderScan();
    // Running low on labels
    var sk = stockFor(f.id), low = stockLow(sk);
    $('stock-low').hidden = !low;
    if (low) {
      $('stock-low-msg').textContent = (sk.on ? plural(sk.on, 'sheet') + ' of ' : 'No sheets of ') + 'these labels left' + (sk.on ? '.' : ' after this one.');
      $('stock-low-order').textContent = sk.link ? 'Reorder' : 'Set up reorder';
      $('stock-low-got').hidden = !sk.pack;
    }
  }

  function renderStock() {
    var f = st.fmt, sk = storeJson(LS.stock + f.id) || {}, counting = typeof sk.on === 'number';
    $('stock-sum').textContent = counting ? '(' + plural(sk.on, 'sheet') + ')' : '';
    var put = function (id, v) { if (document.activeElement !== $(id)) $(id).value = v === undefined || v === null ? '' : v; };
    put('stock-on', counting ? sk.on : ''); put('stock-pack', sk.pack); put('stock-warn', sk.warn === undefined ? '' : sk.warn);
    put('stock-packs', sk.packs || 1); put('stock-link', sk.link || '');
    $('stock-got').hidden = !(counting && sk.pack);
    $('stock-order').hidden = !sk.link;
    $('stock-off').hidden = !counting;
    var pn = productNumbers(f);
    $('stock-info').setAttribute('data-tip', 'Count your sheets of these labels. A sheet comes off the count the first time it goes through the printer, and you get a reminder when you’re low. ' +
      (pn.laser ? 'Buy Avery ' + pn.laser + ' for a laser printer' + (pn.inkjet ? ' or ' + pn.inkjet + ' for an inkjet' : '') + '.' : 'Buy labels made for your kind of printer.'));
    renderStockLink();
  }
  function renderStockLink() {
    var v = $('stock-link').value.trim(), r = v ? T.reorderLink(v, 1) : null, m = $('stock-link-msg');
    m.className = v && !r ? 'warn' : 'hint';
    m.hidden = !v;
    m.textContent = !v ? ''
      : !r ? 'Paste the whole web address, starting with https://'
      : r.cart ? 'Reorder adds the labels to your ' + r.store + ' cart.' : 'Reorder opens this page at ' + r.store + '. Add the labels to your cart there.';
  }
  function reorder() {
    var f = st.fmt, sk = storeJson(LS.stock + f.id) || {};
    var r = sk.link ? T.reorderLink(sk.link, sk.packs || 1) : null;
    if (!r) {
      $('stock-box').open = true;
      try { $('stock-box').scrollIntoView({ block: 'nearest' }); } catch (e) { /* old browser */ }
      $('stock-link').focus();
      say('Paste a link to the labels you buy, then click Save.');
      return;
    }
    openLink(r.url);
    var packs = plural(sk.packs || 1, 'pack');
    say(r.cart ? 'Opened ' + r.store + ' to add ' + packs + ' to your cart. Check out there.' : 'Opened the product page at ' + r.store + '. Add ' + packs + ' to your cart there.');
  }
  function gotPack() {
    var f = st.fmt, sk = stockFor(f.id);
    if (!sk || !sk.pack) { $('stock-box').open = true; $('stock-pack').focus(); say('Enter how many sheets come in a pack, then click Save.'); return; }
    sk.on += sk.pack;
    storeJson(LS.stock + f.id, sk);
    render();
    say('Added ' + plural(sk.pack, 'sheet') + '. You have ' + plural(sk.on, 'sheet') + ' of these labels.');
  }

  function renderHistory() {
    var ul = $('hist'); ul.innerHTML = '';
    var h = history();
    if (!h.length) { ul.appendChild(el('li', 'txt', 'Nothing printed yet.')); return; }
    h.slice(0, 20).forEach(function (e) {
      var li = el('li');
      var top = el('div', 'hist-top');
      top.appendChild(el('span', 'when', fmtDate(e.at)));
      if (canUndoPrint(e)) {
        var u = el('button', 'btn link', 'Undo'); u.type = 'button';
        u.setAttribute('aria-label', 'Undo the print of ' + labelList(e.nums));
        u.addEventListener('click', function () { askUndoPrint(e); });
        top.appendChild(u);
      }
      li.appendChild(top);
      li.appendChild(el('span', '', cap(labelList(e.nums || []))));
      if (e.text) li.appendChild(el('span', 'txt', '“' + e.text + '”'));
      ul.appendChild(li);
    });
  }
  function askUndoPrint(e) {
    var exact = undo && undo.at === e.at;
    var withText = exact || e.nums.some(function (n) { return e.labels && e.labels[n]; });
    var lost = exact ? [] : e.nums.filter(function (n) { return !(e.labels && e.labels[n]); });
    var what = cap(labelList(e.nums)) + (e.nums.length === 1 ? ' becomes' : ' become') + ' open again' + (withText ? ' with ' + (e.nums.length === 1 ? 'its' : 'their') + ' text back' : '');
    var note = lost.length && withText ? ' Pictures and older labels (' + ranges(lost) + ') come back empty.' : !withText ? ' ' + (e.nums.length === 1 ? 'It comes' : 'They come') + ' back empty.' : '';
    ask('Undo this print? ' + what + ', and the trip count goes down by one.' + note, 'Undo print', function () {
      queueChange(function () { return undoPrint(e); }).then(refresh).then(function () {
        say('Undone. ' + cap(labelList(e.nums)) + (e.nums.length === 1 ? ' is' : ' are') + ' open again' + (withText ? (lost.length ? ', with the text that could be kept.' : ' with ' + (e.nums.length === 1 ? 'its' : 'their') + ' text back.') : '.'));
      });
    });
  }

  function renderSettings() {
    var f = st.fmt, nd = nudgeFor(f.id);
    $('auto-mark').checked = autoMarkOn();
    $('show-saved').checked = showSavedOn();
    $('arrow-row').hidden = !(F.hasStrip(f));
    $('feed-arrow').checked = feedArrowOn();
    $('feed-arrow').disabled = !arrowFits() || !!st.headerOther;
    if (document.activeElement !== $('trip-limit')) $('trip-limit').value = tripLimit();
    if (document.activeElement !== $('nudge-x')) $('nudge-x').value = nd.x;
    if (document.activeElement !== $('nudge-y')) $('nudge-y').value = nd.y;
    var ap = st.applied || { x: 0, y: 0 };
    $('nudge-apply').hidden = Math.abs(ap.x - nd.x) < 0.001 && Math.abs(ap.y - nd.y) < 0.001;
    $('home-country').value = homeCountry();
    $('suggest-on').checked = store(LS.suggest) !== 'off';
    renderGeo();
  }
  function renderGeo(note) {
    var key = store(LS.geoKey);
    $('geo-remove').hidden = !key;
    $('geo-save').textContent = key ? 'Replace key' : 'Save key';
    $('geo-status').textContent = note || (key ? 'Using your Geoapify key (…' + key.slice(-4) + ').' : 'Using Photon. No key needed.');
  }

  function renderAbout() {
    var tagged = getSetting(DOC.autoShow) === true, saved = !!docUrl(), ref = cachedRef(), nd = nudgeFor(st.fmt.id), ap = st.applied || { x: 0, y: 0 };
    var rows = [
      ['Version', VERSION],
      ['This sheet', tagged ? (saved ? 'Opens with the panel' : 'Opens with the panel once you save it') : 'Not set to open with the panel yet'],
      ['New sheets', ref ? 'Open with the panel' : refState === 'missing' ? 'Need one more step: save this document, close it, open it again, then make the sheet'
        : refState.indexOf('failed') === 0 ? 'Can’t open with the panel (' + refState.slice(8) + ')' : 'Checking…'],
      ['Trips', trips() + ' of ' + tripLimit()],
      ['Adjustment', nd.x || nd.y ? 'right ' + nd.x + ' in, down ' + nd.y + ' in' + (ap.x === nd.x && ap.y === nd.y ? '' : ' (not yet on this sheet)') : 'None'],
      ['Add-in link', ref ? (ref.storeType || '?') + ' · ' + (ref.store || '?') : '—']
    ];
    var dl = $('about-list'); dl.innerHTML = '';
    rows.forEach(function (r) { dl.appendChild(el('dt', '', r[0])); dl.appendChild(el('dd', '', r[1])); });
  }

  // ---------- labels sent from a phone ----------
  // Linking gives this Mac a mailbox on the server and a read key ({box, key}).
  // The phone only gets the mailbox (in the QR code), so it can send labels
  // but not read them. The panel checks for labels while it's open.
  var phone = { items: [], qr: false, problem: '', checking: false, using: false };
  function phoneLink() { var l = storeJson(LS.phone); return l && l.box && l.key ? l : null; }
  function phonePageUrl(box) { return new URL('phone.html', location.href).href.split('#')[0] + '#' + box; }
  function phoneFetch(path, opts) {
    opts = opts || {};
    var link = phoneLink(), headers = {};
    if (link) headers.Authorization = 'Bearer ' + link.key;
    return fetch(API + path, { method: opts.method || 'GET', headers: headers, cache: 'no-store' }).then(function (r) {
      if (r.ok) return opts.raw ? r : r.json();
      var e = new Error('HTTP ' + r.status); e.status = r.status; throw e;
    });
  }
  function checkPhone() {
    var link = phoneLink();
    if (!link || phone.checking || phone.using || document.visibilityState === 'hidden') return Promise.resolve();
    phone.checking = true;
    return phoneFetch('/box/' + link.box).then(function (b) {
      phone.items = (b && b.items) || [];
      phone.problem = '';
    }, function (e) {
      if (e && e.status === 401) { storeJson(LS.phone, null); phone.items = []; phone.qr = false; phone.problem = 'Your phone was unlinked. Link it again to send labels.'; }
      else phone.problem = 'Couldn’t reach the server. Trying again shortly.';
    }).then(function () { phone.checking = false; renderPhone(); renderPhoneIn(); });
  }
  function startPhoneWatch() {
    checkPhone();
    setInterval(checkPhone, 20000);
    document.addEventListener('visibilitychange', function () { if (document.visibilityState === 'visible') checkPhone(); });
    window.addEventListener('focus', function () { checkPhone(); });
  }
  function renderPhone() {
    var link = phoneLink();
    $('phone-link').hidden = !!link;
    $('phone-show').hidden = !link;
    $('phone-show').textContent = phone.qr ? 'Hide the QR code' : 'Show the QR code';
    $('phone-unlink').hidden = !link;
    $('phone-qr-box').hidden = !link || !phone.qr;
    $('phone-status').textContent = phone.problem || (link ? 'Linked. Labels you send show up on shipping-label sheets.' : '');
  }
  function renderPhoneIn() {
    var k = phone.items.length, box = $('phone-in');
    box.hidden = !k || st.kind !== 'sheet' || !st.fmt;
    if (box.hidden) return;
    var isShip = !!st.fmt.shipping, them = k === 1 ? 'it' : 'them';
    $('phone-in-msg').textContent = isShip ? plural(k, 'label') + ' from your phone.'
      : plural(k, 'label') + ' from your phone ' + (k === 1 ? 'is' : 'are') + ' waiting. Open a shipping-label sheet to use ' + them + '.';
    $('phone-in-use').hidden = !isShip;
    $('phone-in-use').textContent = 'Use ' + them;
    $('phone-in-use').disabled = phone.using;
  }
  function drawPhoneQr() {
    var link = phoneLink();
    if (!link) return Promise.resolve();
    return loadQrLib().then(function () {
      var url = phonePageUrl(link.box), m = T.qrMatrix(url), n = m.length, quiet = 4, scale = 5;
      var size = (n + 2 * quiet) * scale, dpr = window.devicePixelRatio || 1, c = $('phone-qr');
      c.width = Math.round(size * dpr); c.height = c.width;
      c.style.width = size + 'px'; c.style.height = size + 'px';
      c.setAttribute('data-url', url);
      var x = c.getContext('2d');
      x.setTransform(dpr, 0, 0, dpr, 0, 0);
      x.fillStyle = '#fff'; x.fillRect(0, 0, size, size);
      x.fillStyle = '#000';
      for (var r = 0; r < n; r++) for (var col = 0; col < n; col++) if (m[r][col]) x.fillRect((col + quiet) * scale, (r + quiet) * scale, scale, scale);
    });
  }
  function linkPhone() {
    $('phone-link').disabled = true;
    fetch(API + '/pair', { method: 'POST', cache: 'no-store' }).then(function (r) {
      if (!r.ok) throw new Error('HTTP ' + r.status);
      return r.json();
    }).then(function (b) {
      if (!b || !b.box || !b.key) throw new Error('No mailbox in the answer');
      storeJson(LS.phone, { box: b.box, key: b.key, at: Date.now() });
      phone.items = []; phone.problem = ''; phone.qr = true;
      renderPhone();
      return drawPhoneQr().then(function () { say('Scan the QR code with your phone’s camera.'); });
    }).catch(function (e) {
      say(e && e.userMessage ? e.userMessage : 'Couldn’t reach the Label Sheet Keeper server. Check your internet connection and try again.', 'warn');
    }).then(function () { $('phone-link').disabled = false; });
  }
  function unlinkPhone() {
    ask('Unlink your phone? Its sending address stops working, and labels not picked up yet are deleted.', 'Unlink', function () {
      var link = phoneLink();
      var done = function () { storeJson(LS.phone, null); phone.items = []; phone.qr = false; phone.problem = ''; renderPhone(); renderPhoneIn(); say('Your phone is unlinked.'); };
      if (!link) { done(); return; }
      phoneFetch('/box/' + link.box, { method: 'DELETE' }).then(done, function (e) {
        if (e && e.status === 401) { done(); return; }
        say('Couldn’t reach the server, so your phone is still linked. Try again when you’re online.', 'warn');
      });
    }, true);
  }
  // Fetch the waiting labels, delete them from the server, and open them like
  // chosen files: one goes to the crop view, several to the waiting list.
  function usePhoneLabels() {
    var link = phoneLink(), items = phone.items.slice();
    if (!link || !items.length || phone.using || !st.fmt || !st.fmt.shipping) return;
    phone.using = true; renderPhoneIn();
    say('Getting ' + plural(items.length, 'label') + ' from your phone…');
    var files = [], taken = {}, failed = false, arriving = false;
    items.reduce(function (p, it) {
      return p.then(function () {
        if (failed) return null;
        return phoneFetch('/box/' + link.box + '/' + it.id, { raw: true }).then(function (r) { return r.blob(); }).then(function (blob) {
          files.push(new File([blob], it.name || 'Label from your phone', { type: it.type || blob.type }));
          taken[it.id] = true;
          return phoneFetch('/box/' + link.box + '/' + it.id, { method: 'DELETE' }).catch(function () { /* it expires on its own */ });
        }, function (e) {
          if (e && e.status === 409) { arriving = true; return null; }   // just sent; try again soon
          if (!e || e.status !== 404) { failed = true; return null; }
          // Expired on the server: take it off the server's list too.
          taken[it.id] = true;
          return phoneFetch('/box/' + link.box + '/' + it.id, { method: 'DELETE' }).catch(function () { /* next check */ });
        });
      });
    }, Promise.resolve()).then(function () {
      phone.items = phone.items.filter(function (it) { return !taken[it.id]; });
      phone.using = false;
      renderPhoneIn();
      if (!files.length) {
        say(failed ? 'Couldn’t get the labels from your phone. Check your internet connection and try again.'
          : arriving ? 'That label is still on its way. Try again in a moment.'
            : 'That label is no longer waiting. Send it again from your phone.', 'warn');
        return;
      }
      setFillTab('file');
      handleFiles(files);
    });
  }
  function wirePhone() {
    $('phone-link').addEventListener('click', linkPhone);
    $('phone-show').addEventListener('click', function () {
      phone.qr = !phone.qr;
      renderPhone();
      if (phone.qr) drawPhoneQr().catch(function (e) { say(e && e.userMessage ? e.userMessage : 'The QR code couldn’t be drawn.', 'warn'); });
    });
    $('phone-unlink').addEventListener('click', unlinkPhone);
    $('phone-in-use').addEventListener('click', usePhoneLabels);
  }

  // ---------- add your own size ----------
  function sizeFromForm() {
    var num = function (id) { return Number($(id).value); };
    var w = num('sz-w'), h = num('sz-h'), sideways = $('sz-sideways').checked, shipping = $('sz-ship').checked;
    var detail = (w + '″ × ' + h + '″') + ($('sz-detail').value.trim() ? ' · ' + $('sz-detail').value.trim() : '');
    var font = h < 0.6 ? 7 : h < 1.2 ? 10 : h < 2.5 ? 12 : 14;
    return {
      id: 'custom-' + uid(), name: $('sz-name').value.trim() || 'My labels', detail: detail,
      page: sideways ? { width: 11, height: 8.5 } : { width: 8.5, height: 11 },
      across: Math.round(num('sz-across')), down: Math.round(num('sz-down')),
      label: { width: w, height: h },
      firstLabel: { left: num('sz-left'), top: num('sz-top') },
      pitch: { across: w + num('sz-gapx'), down: h + num('sz-gapy') },
      padding: { left: Math.min(0.1, w * 0.05), right: Math.min(0.1, w * 0.05) },
      align: shipping ? 'top' : 'center', topSpace: 0.45,
      font: { name: 'Arial', size: font }, corner: w < 5 ? 0.0625 : 0, shipping: shipping
    };
  }
  function sizeProblems(f) {
    var bad = [];
    if (!(f.across >= 1 && f.down >= 1)) bad.push('Enter how many labels go across and down.');
    if (!(f.label.width > 0 && f.label.height > 0)) bad.push('Enter the label width and height.');
    [f.firstLabel.left, f.firstLabel.top, f.pitch.across - f.label.width, f.pitch.down - f.label.height].forEach(function (v) { if (!(v >= 0)) bad.push('Distances can’t be negative.'); });
    if (bad.length) return bad.filter(function (v, i, a) { return a.indexOf(v) === i; });
    return F.problems(f).map(function (p) {
      if (/right edge/.test(p)) return 'The labels run past the right edge of the page. Check the widths and gaps.';
      if (/bottom edge/.test(p)) return 'The labels run past the bottom of the page. Check the heights and gaps.';
      return cap(p) + '.';
    });
  }
  function renderSizeForm() {
    var f = sizeFromForm(), bad = sizeProblems(f), holder = $('sz-thumb');
    holder.innerHTML = '';
    if (!bad.length) holder.appendChild(thumb(F.withDefaults(JSON.parse(JSON.stringify(f))), 84));
    $('sz-msg').textContent = bad.length ? bad.join(' ') : (f.across * f.down) + ' labels, ' + f.label.width + '″ × ' + f.label.height + '″. Print a test page to check it before saving.';
    $('sz-msg').className = bad.length ? 'warn' : 'hint';
    $('sz-save').disabled = !!bad.length;
    $('sz-test').disabled = !!bad.length;
  }

  // ---------- wiring ----------
  function wire() {
    // Choosing and making sheets
    $('create-btn').addEventListener('click', function () {
      var f = findFormat(chosenId);
      if (!f) { say('Choose a label size first.', 'warn'); return; }
      var make = function () {
        queue(function () { return createSheets(f, null); }).then(function (withPanel) {
          var msg = 'Your new sheet (' + f.name + ') opened in its own window' +
            (withPanel ? ', with this panel.' : '. Click Label Sheet on its Home tab to open the panel there.');
          if (st.kind === 'sheet') { choosing = false; render(); say(msg); }
          else { $('made').hidden = false; $('made-msg').textContent = msg; say(''); }
        });
      };
      var lim = tripLimit();
      var left = sheetEntries(true).filter(function (e) { return e.fmtId === f.id && e.open > 0 && (e.trips || 0) < lim; })[0];
      if (left) ask('You still have ' + plural(left.open, 'open label') + ' on “' + (left.file || 'a sheet you haven’t saved') + '” (' + f.name + '). Finish that sheet first?', 'Make a new sheet anyway', make);
      else make();
    });
    $('broken-new').addEventListener('click', function () { choosing = true; render(); });
    $('back-btn').addEventListener('click', function () { choosing = false; render(); });
    $('other-btn').addEventListener('click', function () { choosing = true; $('made').hidden = true; render(); });
    $('refresh-btn').addEventListener('click', function () { refresh().then(function () { say('Up to date.'); }); });
    $('reset-btn').addEventListener('click', function () {
      var used = keys(st.used);
      if (!used.length && !trips()) { say('Every label on this sheet is already open.'); return; }
      ask('Loaded a fresh sheet? All ' + plural(used.length, 'used label') + ' become open again and the trip count starts over. Labels ready to print stay as they are.', 'Reset', function () {
        undo = null;
        queueChange(function () { return markOpen(used).then(function () { setSetting(DOC.trips, 0); setSetting(DOC.spots, {}); }); }).then(refresh).then(function () { say('Reset. Every label without text is open.'); });
      });
    });
    var repair = function (nudge) {
      var f = st.fmt, margins = st.margins || { top: f.firstLabel.top, left: f.firstLabel.left };
      queueChange(function () { return rebuildSheet(f, margins, nudge); }).then(function (r) {
        return refresh().then(function () {
          say('Rebuilt the labels. Kept ' + plural(r.used, 'used label') + ' and ' + plural(r.ready, 'label') + ' ready to print.' + (r.lost ? ' ' + plural(r.lost, 'label') + ' with text didn’t fit back on the sheet.' : ''), r.lost ? 'warn' : '');
        });
      });
    };
    $('repair-btn').addEventListener('click', function () { repair(st.applied || { x: 0, y: 0 }); });
    $('rebuild-btn').addEventListener('click', function () {
      ask('Repair rebuilds this sheet’s label layout. Used labels and labels ready to print are kept. Continue?', 'Repair', function () { repair(st.applied || { x: 0, y: 0 }); });
    });
    $('test-btn').addEventListener('click', function () {
      if (!st.fmt) return;
      var f = st.fmt;
      queue(function () { return makeTestPage(f); }).then(function () {
        render();
        say('A test page opened in its own window. Print it on scrap paper at 100% scale.');
      });
    });

    // Step 1: mark used labels
    $('map').addEventListener('click', function (ev) {
      var b = ev.target.closest('.cell'); if (!b || busy) return;
      var n = +b.dataset.n;
      if (st.used[n]) {
        queueChange(function () { return markOpen([n]); }).then(refresh).then(function () { say(cap(labelName(n)) + ' is open again.'); });
      } else if (st.ready[n]) {
        ask(cap(labelName(n)) + ' has text on it. Clear it and mark it used?', 'Clear and mark used', function () {
          queueChange(function () { return markUsed([n]); }).then(refresh).then(function () { say(cap(labelName(n)) + ' marked used.'); });
        }, true);
      } else {
        queueChange(function () { return markUsed([n]); }).then(refresh).then(function () { say(cap(labelName(n)) + ' marked used.'); });
      }
    });
    $('first-btn').addEventListener('click', function () {
      var max = total(), k = Math.round(Number($('first-n').value));
      if (!(k >= 1 && k <= max)) { say('Enter a number from 1 to ' + max + '.', 'warn'); return; }
      var nums = []; for (var n = 1; n <= k; n++) if (!st.used[n]) nums.push(n);
      if (!nums.length) { say((k === 1 ? 'Label 1 is' : 'Labels 1–' + k + ' are') + ' already marked used.'); return; }
      var run = function () { queueChange(function () { return markUsed(nums); }).then(refresh).then(function () { say('Marked ' + labelList(nums) + ' used.'); }); };
      var withText = nums.filter(function (x) { return st.ready[x]; });
      if (withText.length) ask(cap(labelList(withText)) + (withText.length === 1 ? ' has' : ' have') + ' text on it. Clear and mark used?', 'Clear and mark used', run, true);
      else run();
    });

    // Step 2: copy
    function doCopy(k, all) {
      queue(function () { return readState().then(apply).then(cursorLabel); }).then(function (src) {
        var open = openList(), targets = open.slice(0, k), short = all ? 0 : Math.max(0, k - open.length);
        if (!targets.length && !short) { say('There are no open labels left on this sheet.', 'warn'); return; }
        var copyHere = function () {
          if (!targets.length) return Promise.resolve();
          return queueChange(function () { return copyFrom(src, targets); }).then(refresh);
        };
        if (!short) { copyHere().then(function () { say('Copied ' + labelName(src) + ' to ' + labelList(targets) + '.'); }); return; }
        queue(function () { return readLabelParas(src); }).then(function (paras) {
          var copies = []; for (var i = 0; i < short; i++) copies.push(paras);
          askOverflow(st.fmt, targets.length, copies, 'copy', copyHere, function (msg) {
            say('Copied ' + labelName(src) + (targets.length ? ' to ' + labelList(targets) + ' and ' : '; ') + msg);
          });
        });
      });
    }
    $('copy-btn').addEventListener('click', function () {
      var k = Math.round(Number($('copies').value));
      if (!(k >= 1)) { say('Enter how many copies you want.', 'warn'); return; }
      doCopy(k, false);
    });
    $('fill-all').addEventListener('click', function () { doCopy(total(), true); });
    $('fit-btn').addEventListener('click', function () {
      queue(function () { return readState().then(apply); }).then(function () {
        var n = st.here && st.ready[st.here] ? st.here : 0;
        var nums = n ? [n] : keys(st.over);
        if (!nums.length) { say('Click inside a label with text first. Labels whose text may be cut off get a red outline on the map.', 'warn'); return; }
        doFit(nums);
      });
    });

    // Step 2: shipping label files, address, table
    wireShip();
    wireAddress();
    wirePhone();
    $('table-open').addEventListener('click', openTableWindow);
    $('reprint-btn').addEventListener('click', function () {
      queueChange(reprintLast).then(function (n) { return refresh().then(function () { say('Put the last label on ' + theName(n) + ' again.'); }); });
    });

    // Step 3: print
    $('printed-btn').addEventListener('click', function () {
      if (!keys(st.ready).length) { say('No labels are ready to print.', 'warn'); return; }
      queueChange(markPrinted).then(function (nums) {
        return refresh().then(function () { say('Marked ' + labelList(nums) + ' used. ' + plural(openList().length, 'open label') + ' left.'); });
      });
      readPrintTime().then(function (ms) { if (ms) { lastPrintSeen = Math.max(lastPrintSeen || 0, ms); setSetting(DOC.lastPrint, lastPrintSeen); } }).catch(function () {});
    });
    $('undo-btn').addEventListener('click', function () {
      if (!undo) return;
      queueChange(undoPrinted).then(refresh).then(function () { say('Undone. Those labels have their text back.'); });
    });
    $('warn-pass-hide').addEventListener('click', function () { store(LS.warnPass, '1'); render(); });
    $('test-yes').addEventListener('click', function () { store(LS.testCheck + st.fmt.id, null); render(); say('You’re set for this printer and these labels.'); });
    $('test-no').addEventListener('click', function () {
      var box = $('settings'); box.open = true;
      var nd = $('nudge'); nd.classList.remove('flash'); void nd.offsetWidth; nd.classList.add('flash');
      try { nd.scrollIntoView({ block: 'center' }); } catch (e) { /* old browser */ }
      $('nudge-x').focus();
      say('Enter how far the boxes were off, then click Save. Print another test page to check.');
    });
    // Label stock
    $('stock-link').addEventListener('input', renderStockLink);
    // Numbers & codes
    ['num-text', 'num-start', 'num-count', 'num-total', 'num-pad', 'code-data'].forEach(function (id) {
      var keep = function () {
        // A first number typed in is where numbering starts next time too.
        var n = Math.round(Number($('num-start').value));
        saveCodes(id === 'num-start' && $('num-start').value !== '' && n >= 0 ? { next: n } : null);
        renderCodes();
      };
      $(id).addEventListener('input', keep);
      $(id).addEventListener('change', keep);
    });
    Array.prototype.forEach.call($('code-kind').querySelectorAll('button'), function (b) {
      b.addEventListener('click', function () { saveCodes({ kind: b.dataset.kind }); renderCodes(); });
    });
    $('num-fill').addEventListener('click', fillCodes);

    // Back up and restore
    $('backup-make').addEventListener('click', function () {
      queue(makeBackup).then(function () {
        say('Your backup opened in its own window. Save it (⌘S) somewhere safe, such as iCloud Drive. It holds your addresses, so keep it private.');
      });
    });
    var restoreIn = $('restore-file');
    restoreIn.addEventListener('change', function () {
      var file = restoreIn.files && restoreIn.files[0]; restoreIn.value = '';
      if (!file) return;
      if (!window.JSZip) { say('The file reader didn’t load. Close and reopen the panel.', 'warn'); return; }
      readFileAs(file, 'buffer').then(function (buf) { return window.JSZip.loadAsync(buf); }).then(backupFromZip).then(function (p) {
        if (!p) { say('That file isn’t a Label Sheet Keeper backup. Choose the backup document you saved.', 'warn'); return; }
        askRestore(p, render);
      }).catch(function () { say('That file couldn’t be read. Choose the backup document (.docx) you saved.', 'warn'); });
    });
    $('backup-restore').addEventListener('click', function () {
      if (backupHere && backupHere !== 'bad') askRestore(backupHere, function () { $('backup-done').hidden = false; });
    });
    $('stock-save').addEventListener('click', function () {
      var f = st.fmt, num = function (id) { var v = $(id).value.trim(); return v === '' ? null : Math.round(Number(v)); };
      var on = num('stock-on'), pack = num('stock-pack'), warn = num('stock-warn'), packs = num('stock-packs'), link = $('stock-link').value.trim();
      if (on === null && !link) {
        if (storeJson(LS.stock + f.id)) { storeJson(LS.stock + f.id, null); render(); say('Cleared the label stock for this size.'); return; }
        say('Enter how many sheets you have now, or a link to where you buy them.', 'warn'); $('stock-on').focus(); return;
      }
      if ((on !== null && !(on >= 0 && on <= 9999)) || (pack !== null && !(pack >= 1 && pack <= 9999)) || (warn !== null && !(warn >= 0 && warn <= 9999)) || (packs !== null && !(packs >= 1 && packs <= 99))) {
        say('Check the numbers: whole numbers only, and at least 1 sheet in a pack.', 'warn'); return;
      }
      if (link && !T.reorderLink(link, 1)) { say('That link doesn’t look right. Paste the whole web address, starting with https://', 'warn'); $('stock-link').focus(); return; }
      var rec = { pack: pack || 0, warn: warn === null ? Math.min(5, pack || 5) : warn, packs: packs || 1, link: link };
      if (on !== null) rec.on = on;
      storeJson(LS.stock + f.id, rec);
      render();
      say(on !== null ? 'Saved. The count goes down by one each time a fresh sheet goes through the printer.' : 'Saved the link. Enter how many sheets you have to count them too.');
    });
    $('stock-got').addEventListener('click', gotPack);
    $('stock-low-got').addEventListener('click', gotPack);
    $('stock-order').addEventListener('click', reorder);
    $('stock-low-order').addEventListener('click', reorder);
    $('stock-off').addEventListener('click', function () {
      var f = st.fmt, sk = storeJson(LS.stock + f.id) || {};
      delete sk.on;
      storeJson(LS.stock + f.id, sk.link || sk.pack ? sk : null);
      render();
      say('Stopped counting sheets of these labels. Your link is kept for reordering.');
    });

    var scanIn = $('scan-file');
    scanIn.addEventListener('change', function () { var file = scanIn.files && scanIn.files[0]; scanIn.value = ''; handleScan(file); });
    $('scan-save').addEventListener('click', function () {
      if (!scanResult || !st.fmt) return;
      storeJson(LS.nudge + st.fmt.id, scanResult.nudge);
      store(LS.testCheck + st.fmt.id, null);
      scanResult = null;
      render();
      say('Saved. New sheets and test pages for this size use it. Print another test page to check it.' + ($('nudge-apply').hidden ? '' : ' To move the labels on this sheet too, click Apply to this sheet in Settings.'));
    });
    $('scan-done').addEventListener('click', function () {
      if (st.fmt) store(LS.testCheck + st.fmt.id, null);
      scanResult = null; render(); say('You’re set for this printer and these labels.');
    });
    $('use-up-mine').addEventListener('click', function () {
      var lines = textLines(fromAddress()), open = openList();
      if (!lines.length) { setFillTab('addr'); say('Enter your return address first.', 'warn'); return; }
      var d = designParas(lines, currentDesign(), st.fmt);
      queueChange(function () { return writeLabels(open, d.paras, d.pic); }).then(refresh).then(function () { say('Filled ' + labelList(open) + ' with your return address.'); });
    });
    $('use-up-copy').addEventListener('click', function () {
      var src = st.here && st.ready[st.here] ? st.here : keys(st.ready)[0], open = openList();
      queueChange(function () { return copyFrom(src, open); }).then(refresh).then(function () { say('Filled ' + labelList(open) + ' with copies of ' + labelName(src) + '.'); });
    });

    // Settings
    $('auto-mark').addEventListener('change', function () { store(LS.autoMark, $('auto-mark').checked ? 'on' : 'off'); });
    $('show-saved').addEventListener('change', function () { store(LS.showSaved, $('show-saved').checked ? 'on' : 'off'); if (st.kind === 'sheet') renderTally(); });
    $('group-pick').addEventListener('change', function () { groupShown = $('group-pick').value; bookEdit = null; renderAddress(); });
    $('group-fill').addEventListener('click', function () {
      var g = groupShown, list = book().filter(function (e) { return inGroup(e, g); }).map(contactLines);
      if (!list.length) return;
      fillAddresses(list, 'with ' + g);
    });
    $('feed-arrow').addEventListener('change', function () { store(LS.feedArrow, $('feed-arrow').checked ? 'on' : 'off'); syncHeader(); render(); });
    $('trip-limit').addEventListener('change', function () {
      var n = Math.round(Number($('trip-limit').value));
      if (!(n >= 1 && n <= 20)) { say('Enter a number of trips from 1 to 20.', 'warn'); return; }
      store(LS.tripLimit, String(n)); updateRegistry(); render();
    });
    $('nudge-save').addEventListener('click', function () {
      var x = Number($('nudge-x').value), y = Number($('nudge-y').value);
      if (!(Math.abs(x) <= 0.5 && Math.abs(y) <= 0.5)) { say('Enter adjustments between -0.5 and 0.5 inches.', 'warn'); return; }
      storeJson(LS.nudge + st.fmt.id, { x: Math.round(x * 1000) / 1000, y: Math.round(y * 1000) / 1000 });
      store(LS.testCheck + st.fmt.id, null);
      render();
      say('Saved. New sheets and test pages for this size use it.' + ($('nudge-apply').hidden ? '' : ' Click Apply to this sheet to move these labels too.'));
    });
    $('home-country').addEventListener('change', function () {
      store(LS.home, $('home-country').value === 'CA' ? 'CA' : null);
      say(homeCountry() === 'CA' ? 'Addresses in the US get “USA” as a last line.' : 'Addresses in Canada get “CANADA” as a last line.');
    });
    $('geo-save').addEventListener('click', function () {
      var key = $('geo-key').value.trim();
      if (!key) { renderGeo('Paste your key in the box first.'); $('geo-key').focus(); return; }
      if (!/^[A-Za-z0-9_-]{16,64}$/.test(key)) { renderGeo('That doesn’t look like a Geoapify key. Copy the whole key from your Geoapify project.'); return; }
      $('geo-save').disabled = true;
      renderGeo('Checking the key with Geoapify…');
      T.suggestAddresses('1 Main St', key).then(function () {
        store(LS.geoKey, key); $('geo-key').value = '';
        renderGeo('Key saved. Suggestions now use Geoapify.');
      }, function (e) {
        if (e && (e.status === 401 || e.status === 403)) renderGeo('Geoapify didn’t accept that key. Check that you copied all of it, and that the key allows this add-in’s web address.');
        else { store(LS.geoKey, key); $('geo-key').value = ''; renderGeo('Saved, but Geoapify couldn’t be reached to check the key. Suggestions will start once it can.'); }
      }).then(function () { $('geo-save').disabled = false; });
    });
    $('geo-remove').addEventListener('click', function () { store(LS.geoKey, null); renderGeo('Key removed. Suggestions use Photon again.'); });
    $('suggest-on').addEventListener('change', function () { store(LS.suggest, $('suggest-on').checked ? null : 'off'); if (!$('suggest-on').checked) closeDocSugg(); });
    $('nudge-apply').addEventListener('click', function () {
      var nd = nudgeFor(st.fmt.id), L = S.layout(st.fmt, { margins: st.margins, nudge: nd });
      if (L.clampedUp) say('These labels can’t move that far up on this sheet. They were moved as far as possible; new sheets get the full adjustment.', 'warn');
      repair(nd);
    });

    $('confirm-yes').addEventListener('click', function () { var a = pendingConfirm; closeAsk(); if (a) a(); });
    $('confirm-no').addEventListener('click', closeAsk);
    $('preview-btn').addEventListener('click', openPreview);
    $('dsugg-close').addEventListener('click', function () { if (ds.at) ds.closed = ds.at.n + ':' + ds.at.index; closeDocSugg(); });
    $('confirm-alt').addEventListener('click', function () { var a = pendingAlt; closeAsk(); if (a) a(); });
    wireFileLabels();

    // Add your own size
    ['sz-name', 'sz-detail', 'sz-across', 'sz-down', 'sz-w', 'sz-h', 'sz-top', 'sz-left', 'sz-gapx', 'sz-gapy', 'sz-sideways', 'sz-ship'].forEach(function (id) {
      $(id).addEventListener('input', renderSizeForm);
      $(id).addEventListener('change', renderSizeForm);
    });
    $('sz-test').addEventListener('click', function () {
      var f = F.withDefaults(sizeFromForm());
      queue(function () { return createDoc(f, { test: true }); }).then(function () { say('A test page for your size opened in its own window. Print it on scrap paper at 100% scale.'); });
    });
    $('sz-save').addEventListener('click', function () {
      var f = sizeFromForm();
      if (sizeProblems(f).length) return;
      var list = storeJson(LS.custom) || []; list.push(f); storeJson(LS.custom, list);
      chosenId = f.id; $('add-size').open = false;
      render();
      say('Saved “' + f.name + '” on this Mac. It’s selected; click Make the sheet.');
    });
  }

  function doFit(nums) {
    queueChange(function () { return fitLabels(nums); }).then(function (r) {
      return refresh().then(function () {
        var parts = [];
        if (r.fitted.length) parts.push('Shrank the text on ' + labelList(r.fitted) + ' to fit.');
        if (r.tooLong.length) parts.push(cap(labelList(r.tooLong)) + (r.tooLong.length === 1 ? ' has' : ' have') + ' too much text to fit at a readable size. Shorten ' + (r.tooLong.length === 1 ? 'it' : 'them') + ' or remove a line.');
        if (!parts.length) parts.push('The text already fits.');
        say(parts.join(' '), r.tooLong.length ? 'warn' : '');
      });
    });
  }

  function wireShip() {
    [$('side-top'), $('side-bottom')].forEach(function (btn) {
      btn.addEventListener('click', function () {
        if (!st.fmt) return;
        store(LS.side + st.fmt.id, btn.dataset.side);
        render();
        say('Got it. Label 1 is the ' + btn.dataset.side + ' label on your sheet.');
      });
    });
    var input = $('ship-file');
    input.addEventListener('change', function () { handleFiles(input.files); input.value = ''; });
    var card = $('pane-file');
    card.addEventListener('dragover', function (e) { e.preventDefault(); card.classList.add('drop'); });
    card.addEventListener('dragleave', function () { card.classList.remove('drop'); });
    card.addEventListener('drop', function (e) {
      e.preventDefault(); card.classList.remove('drop');
      handleFiles(e.dataTransfer && e.dataTransfer.files);
    });
    // Paste a screenshot: ⌘V anywhere in the panel (outside text boxes).
    document.addEventListener('paste', function (e) {
      if (!st.fmt || !st.fmt.shipping || st.kind !== 'sheet') return;
      var tgt = e.target;
      if (tgt && /^(TEXTAREA|INPUT|SELECT)$/.test(tgt.tagName)) return;
      var items = (e.clipboardData && e.clipboardData.items) || [];
      for (var i = 0; i < items.length; i++) {
        if (items[i].kind === 'file' && /^image\//.test(items[i].type)) {
          var file = items[i].getAsFile();
          if (file) { e.preventDefault(); handleFile(file); return; }
        }
      }
      say('There’s no picture on the clipboard. Copy a label image first, then press ⌘V here.', 'warn');
    });

    $('pg-prev').addEventListener('click', function () { if (ship.doc && ship.page > 1) { ship.page--; renderPdfPage().then(afterSource).catch(fileError); } });
    $('pg-next').addEventListener('click', function () { if (ship.doc && ship.page < ship.pages) { ship.page++; renderPdfPage().then(afterSource).catch(fileError); } });
    $('crop-find').addEventListener('click', function () {
      if (!ship.src) return;
      var found = findLabel(ship.src);
      ship.crop = found || trimBox(ship.src);
      drawCrop();
      say(found ? 'Found the label’s border.' : 'No printed border found, so the white edges were trimmed instead. Drag across the label if needed.', found ? '' : 'warn');
    });
    $('pg-all').addEventListener('click', usePdfPages);
    $('wait-place').addEventListener('click', placeWaiting);
    $('wait-clear').addEventListener('click', function () { ship.waiting = []; renderWaiting(); say('Cleared the waiting labels.'); });
    $('crop-all').addEventListener('click', function () { if (ship.src) { ship.crop = { x: 0, y: 0, w: ship.src.width, h: ship.src.height }; drawCrop(); } });
    $('crop-turn').addEventListener('click', function () { ship.turn = (ship.turn + 90) % 360; drawFit(); });

    // Drag a rectangle on the page preview to choose the label area.
    var wrap = $('crop-wrap'), drag = null;
    function frac(e) {
      var r = wrap.getBoundingClientRect();
      return { x: Math.min(1, Math.max(0, (e.clientX - r.left) / r.width)), y: Math.min(1, Math.max(0, (e.clientY - r.top) / r.height)) };
    }
    function rectFrom(a, b) {
      return { x: Math.min(a.x, b.x) * ship.src.width, y: Math.min(a.y, b.y) * ship.src.height, w: Math.abs(a.x - b.x) * ship.src.width, h: Math.abs(a.y - b.y) * ship.src.height };
    }
    wrap.addEventListener('pointerdown', function (e) {
      if (!ship.src) return;
      drag = { a: frac(e), b: frac(e) };
      try { wrap.setPointerCapture(e.pointerId); } catch (err) { /* not supported */ }
      e.preventDefault();
    });
    wrap.addEventListener('pointermove', function (e) { if (!drag) return; drag.b = frac(e); placeBox(rectFrom(drag.a, drag.b)); });
    wrap.addEventListener('pointerup', function (e) {
      if (!drag) return;
      drag.b = frac(e);
      var r = rectFrom(drag.a, drag.b);
      drag = null;
      if (r.w > ship.src.width * 0.03 && r.h > ship.src.height * 0.03) ship.crop = { x: Math.round(r.x), y: Math.round(r.y), w: Math.round(r.w), h: Math.round(r.h) };
      drawCrop();
    });
    wrap.addEventListener('pointercancel', function () { drag = null; drawCrop(); });

    $('ship-place').addEventListener('click', function () {
      if (!ship.src) { say('Choose a PDF or picture first.', 'warn'); return; }
      var n = nextOpen();
      if (!n) { say('There are no open labels on this sheet.', 'warn'); return; }
      var out = output(), b64 = out.canvas.toDataURL('image/png').split(',')[1];
      queueChange(function () { return writeLabels([n], [], { b64: b64, w: out.wIn * 72, h: out.hIn * 72, align: 'center', alt: 'Shipping label' }); }).then(refresh).then(function () {
        say('Placed on ' + theName(n) + '. Print with ⌘P.');
      });
    });
  }

  function wireAddress() {
    wireExtras();
    ['addr-text', 'from-addr', 'to-addr'].forEach(function (id) { attachSuggest($(id)); });
    window.addEventListener('resize', function () { if (!$('asugg').hidden) closeAs(); });
    // Address labels (with a design)
    $('addr-pick').addEventListener('change', function () {
      loadPicked($('addr-pick').value, $('addr-text'));
      $('addr-pick').value = '';
    });
    function placeAddr(all) {
      var addr = textLines($('addr-text').value), lines = boxLines();
      if (!addr.length) { say('Type an address or choose a saved one first.', 'warn'); $('addr-text').focus(); return; }
      var nums = all ? openList() : [nextOpen()].filter(Boolean);
      if (!nums.length) { say('There are no open labels on this sheet.', 'warn'); return; }
      var d = designParas(lines, currentDesign(), st.fmt);
      if (currentDesign() === 'logo' && !d.pic) { say('Choose a logo first, or pick another design.', 'warn'); return; }
      var note = extrasNote();
      queueChange(function () { return writeLabels(nums, d.paras, d.pic); }).then(function () { noteRecent(addr, note); return refresh(); }).then(function () {
        say('Put the address on ' + labelList(nums) + '.');
        askShrink(nums);
      });
    }
    // Saves the address box (with its phones, emails) as a contact, or adds them to the one already saved.
    function saveBox(box, what) {
      var lines = textLines(box.value);
      if (!lines.length) { say('Type ' + what + ' first.', 'warn'); return; }
      var c = xVals(); c.lines = lines;
      var e = contactById(xs.id);
      if (e && e.lines2 && xs.which === 2) { c.lines = e.lines; c.lines2 = lines; }
      var r = addContacts([c]);
      say(r.added ? 'Saved ' + lines[0] + ' in the address book.' : r.updated ? 'Added the new details to ' + lines[0] + ' in the address book.' : 'That address is already in the address book.');
      if (r.added) xs = { id: (book().filter(function (k) { return keyOf(k.lines) === keyOf(c.lines); })[0] || {}).id || '', which: 1 };
      renderAddress();
    }
    $('addr-one').addEventListener('click', function () { placeAddr(false); });
    $('addr-all').addEventListener('click', function () { placeAddr(true); });
    $('addr-save').addEventListener('click', function () { saveBox($('addr-text'), 'an address'); });
    $('addr-mine').addEventListener('click', function () {
      var v = $('addr-text').value.replace(/\s+$/, '');
      if (!v.trim()) { say('Type your return address first.', 'warn'); return; }
      saveFrom(v); renderAddress(); say('Saved as your return address.');
    });
    var logoIn = $('logo-file');
    logoIn.addEventListener('change', function () {
      var file = logoIn.files && logoIn.files[0]; logoIn.value = '';
      if (!file) return;
      var reader = new FileReader();
      reader.onload = function () {
        var img = new Image();
        img.onload = function () {
          var s = Math.min(1, 600 / Math.max(img.naturalWidth, img.naturalHeight));
          var c = document.createElement('canvas'); c.width = Math.round(img.naturalWidth * s); c.height = Math.round(img.naturalHeight * s);
          var x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.drawImage(img, 0, 0, c.width, c.height);
          storeJson(LS.logo, { b64: c.toDataURL('image/png').split(',')[1], w: c.width, h: c.height });
          renderAddress(); say('Logo saved.');
        };
        img.onerror = function () { say('That picture couldn’t be opened. Try a PNG or JPEG.', 'warn'); };
        img.src = reader.result;
      };
      reader.readAsDataURL(file);
    });

    // Shipping address labels (From / Ship to)
    var from = $('from-addr');
    from.value = fromAddress();
    from.addEventListener('input', function () { saveFrom(from.value); });
    $('to-pick').addEventListener('change', function () {
      loadPicked($('to-pick').value, $('to-addr'));
      $('to-pick').value = '';
    });
    $('to-save').addEventListener('click', function () { saveBox($('to-addr'), 'the Ship to address'); });
    $('ship-addr-place').addEventListener('click', function () {
      var to = $('to-addr').value.replace(/\s+$/, '');
      if (!to.trim()) { say('Type the Ship to address first.', 'warn'); $('to-addr').focus(); return; }
      var n = nextOpen();
      if (!n) { say('There are no open labels on this sheet.', 'warn'); return; }
      var fromText = from.value.replace(/\s+$/, ''), toText = boxLines().join('\n'), note = extrasNote();
      saveFrom(fromText);
      queueChange(function () { return writeLabels([n], shippingParas(fromText, toText, st.fmt)); }).then(function () { noteRecent(textLines(to), note); return refresh(); }).then(function () {
        $('to-addr').value = '';
        clearExtras();
        say('Address label placed on ' + theName(n) + '. You can edit it in the document.');
      });
    });

    // Address book import: read the file, then choose which contacts to import.
    var bookIn = $('book-file');
    bookIn.addEventListener('change', function () {
      var file = bookIn.files && bookIn.files[0]; bookIn.value = '';
      if (!file) return;
      var name = (file.name || '').toLowerCase(), reader = new FileReader(), opts = { home: homeCountry() };
      var done = function (list) { startImport(file.name, list); };
      if (/\.xlsx$/.test(name)) {
        reader.onload = function () { T.readXlsx(reader.result).then(function (rows) { done(T.rowsToContacts(rows, opts)); }).catch(function () { say('That Excel file couldn’t be read.', 'warn'); }); };
        reader.readAsArrayBuffer(file);
      } else {
        reader.onload = function () {
          var text = String(reader.result || '');
          if (/\.vcf$/.test(name) || /BEGIN:VCARD/i.test(text)) done(T.parseVCard(text, opts));
          else done(T.rowsToContacts(T.parseDelimited(text), opts));
        };
        reader.readAsText(file);
      }
    });
    $('imp-filter').addEventListener('input', function () { if (imp) { imp.filter = $('imp-filter').value; renderImport(); } });
    $('imp-all').addEventListener('click', function () { if (imp) { shownImport().forEach(function (i) { imp.picked[i] = true; }); renderImport(); } });
    $('imp-none').addEventListener('click', function () { if (imp) { imp.picked = {}; renderImport(); } });
    $('imp-cancel').addEventListener('click', function () { imp = null; renderImport(); say('Nothing was imported.'); });
    $('imp-go').addEventListener('click', function () {
      if (!imp) return;
      var list = Object.keys(imp.picked).map(function (i) { return imp.list[Number(i)]; });
      var r = addContacts(list);
      imp = null;
      renderAddress();
      var parts = [];
      if (r.added) parts.push('Imported ' + plural(r.added, 'contact') + '.');
      if (r.updated) parts.push('Added new details to ' + plural(r.updated, 'contact') + ' already saved.');
      say(parts.join(' ') || 'Those contacts were already in the address book.');
    });
    $('book-export').addEventListener('click', function () {
      var b = book();
      if (!b.length) { say('The address book is empty.', 'warn'); return; }
      var q = function (s) { return '"' + String(s).replace(/"/g, '""') + '"'; };
      var csv = 'Name,Address\n' + b.map(function (e) { return q(e.name) + ',' + q(e.lines.slice(1).join('\n')); }).join('\n');
      var fallback = function () {
        var ta = $('book-csv'); ta.hidden = false; ta.value = csv; ta.focus(); ta.select();
        say('Press ⌘C to copy the selected addresses.');
      };
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(csv).then(function () { say('Copied ' + plural(b.length, 'address', 'es') + '. Paste them into Excel or a text file to keep a backup.'); }, fallback);
      } else fallback();
    });
  }

  // Labels that stand in for file pickers open them with Enter or Space too.
  function wireFileLabels() {
    Array.prototype.forEach.call(document.querySelectorAll('label[for][tabindex]'), function (lab) {
      lab.addEventListener('keydown', function (e) {
        if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); var inp = document.getElementById(lab.getAttribute('for')); if (inp) inp.click(); }
      });
    });
  }

  // ---------- little (i) explanations ----------
  // Pointing at an (i) shows its explanation in a small box; it goes away when
  // the pointer leaves. Keyboard focus shows it too. A click (or tap) keeps it
  // open until the next click anywhere, or Escape.
  var tipFor = null, tipPinned = false;
  function showTip(btn) {
    var box = $('tipbox'), text = btn.getAttribute('data-tip');
    if (!text) return;
    if (tipFor && tipFor !== btn) tipFor.removeAttribute('aria-describedby');
    tipFor = btn;
    box.textContent = text;
    box.hidden = false;
    btn.setAttribute('aria-describedby', 'tipbox');
    var vw = document.documentElement.clientWidth || window.innerWidth, vh = window.innerHeight;
    box.style.maxWidth = Math.min(280, vw - 16) + 'px';
    box.style.left = '0px'; box.style.top = '0px';
    var r = btn.getBoundingClientRect(), w = box.offsetWidth, h = box.offsetHeight;
    var left = Math.min(Math.max(8, r.left + r.width / 2 - w / 2), Math.max(8, vw - w - 8));
    var top = r.bottom + 6;
    if (top + h > vh - 8 && r.top - h - 6 >= 8) top = r.top - h - 6;
    box.style.left = Math.round(left) + 'px'; box.style.top = Math.round(top) + 'px';
  }
  function hideTip() {
    if (tipFor) tipFor.removeAttribute('aria-describedby');
    tipFor = null; tipPinned = false;
    $('tipbox').hidden = true;
  }
  function infoFrom(node) { return node && node.closest ? node.closest('.info') : null; }
  function wireTips() {
    document.addEventListener('mouseover', function (e) { var b = infoFrom(e.target); if (b && b !== tipFor) { tipPinned = false; showTip(b); } });
    document.addEventListener('mouseout', function (e) { var b = infoFrom(e.target); if (b && b === tipFor && !tipPinned && !b.contains(e.relatedTarget)) hideTip(); });
    document.addEventListener('focusin', function (e) { var b = infoFrom(e.target); if (b) showTip(b); });
    document.addEventListener('focusout', function (e) { if (infoFrom(e.target) === tipFor && !tipPinned) hideTip(); });
    // A click must not open a section or focus a field.
    document.addEventListener('click', function (e) {
      var b = infoFrom(e.target);
      if (!b) { if (tipFor) hideTip(); return; }
      e.preventDefault(); e.stopPropagation();
      showTip(b);
      tipPinned = true;
    }, true);
    document.addEventListener('keydown', function (e) { if (e.key === 'Escape' && tipFor) hideTip(); });
    window.addEventListener('scroll', function () { if (tipFor) hideTip(); }, true);
    window.addEventListener('resize', function () { if (tipFor) hideTip(); });
  }

  // ---------- start ----------
  function demo() {
    // Preview outside Word: taskpane.html?demo (address sheet), ?demo=ship, ?demo&choose
    var shipDemo = /demo=ship/.test(location.search);
    var f = shipDemo ? F.formats.filter(function (x) { return x.shipping; })[0] : F.formats[0];
    st = { kind: 'sheet', fmt: f, used: {}, ready: {}, texts: {}, pics: {}, over: {}, here: shipDemo ? 1 : 8, header: true, margins: { top: S.STRIP_MARGIN, left: f.firstLabel.left }, applied: { x: 0, y: 0 } };
    if (!shipDemo) {
      for (var n = 1; n <= 7; n++) st.used[n] = true;
      [8, 9, 10].forEach(function (k) { st.ready[k] = true; st.texts[k] = 'Sample Name\n123 Main St\nBoston, MA 02101'; });
      st.over[10] = true;
    }
    if (/choose/.test(location.search)) st.kind = 'none';
    window.LSK_TEST = { ship: ship, output: output, handleFile: handleFile, st: function () { return st; } };
    render();
  }

  function start() {
    if (typeof Office === 'undefined') { show('v-outside'); return; }
    Office.onReady(function (info) {
      if (!info || info.host !== Office.HostType.Word) {
        if (/[?&]demo\b/.test(location.search)) { wire(); wireTips(); demo(); } else show('v-outside');
        return;
      }
      if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) { show('v-old'); return; }
      wire();
      wireTips();
      refresh();
      startPrintWatch();
      startPhoneWatch();
      // Learn how Word attaches this add-in, so new sheets can open with the panel.
      findAddinRef().then(function () { if (st.kind === 'sheet') renderAbout(); });
      try { Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged); } catch (e) { /* older Word */ }
      window.addEventListener('pagehide', forgetIfUnsaved);
      window.LSK_TEST = { st: function () { return st; }, ship: ship, output: output, handleFile: handleFile, fillFromRows: fillFromRows, checkPrinted: checkPrinted, checkFillJob: checkFillJob, scan: function () { return scanResult; }, checkPhone: checkPhone, phone: phone, backupKeys: backupKeys };
    });
  }
  start();
})();
