/*
  Label Sheet Keeper — the table window.
  One row per label, in order: row 1 goes on the first open label, and so on.
  The first column shows the label each row goes on; a "New sheet" divider
  marks where rows run past the open labels and onto new sheets. An empty
  row leaves its label open.
*/
(function () {
  'use strict';

  var T = window.LSK_TOOLS;
  var COLS = ['name', 'company', 'street', 'apt', 'city', 'state', 'zip', 'phone', 'email'];
  var DRAFT = 'lsk.tableDraft', INFO = 'lsk.tableInfo', BOOK = 'lsk.addressBook', GEO_KEY = 'lsk.geoKey', SUGGEST = 'lsk.suggest';
  var rowsEl = document.getElementById('rows');

  function $(id) { return document.getElementById(id); }
  function store(k, v) { try { if (v === undefined) return localStorage.getItem(k); localStorage.setItem(k, v); } catch (e) { /* blocked */ } return null; }
  function json(k) { try { return JSON.parse(store(k) || 'null'); } catch (e) { return null; } }
  function plural(n, w, s) { return n + ' ' + w + (n === 1 ? '' : (s || 's')); }
  function msg(text, warn) { $('msg').textContent = text || ''; $('msg').className = 'msg' + (warn ? ' warn' : ''); }

  // ---------- rows ----------
  function dataRows() { return Array.prototype.filter.call(rowsEl.children, function (tr) { return tr.classList.contains('row'); }); }
  function addRow(values) {
    var tr = document.createElement('tr');
    tr.className = 'row';
    var num = document.createElement('td');
    tr.appendChild(num);
    COLS.forEach(function (c, i) {
      var td = document.createElement('td'), inp = document.createElement('input');
      inp.type = 'text'; inp.value = (values && values[i]) || ''; inp.setAttribute('aria-label', c);
      inp.autocomplete = 'off'; inp.spellcheck = false;
      td.appendChild(inp); tr.appendChild(td);
    });
    var xd = document.createElement('td'), x = document.createElement('button');
    x.type = 'button'; x.className = 'x'; x.textContent = '×'; x.setAttribute('aria-label', 'Delete row');
    x.addEventListener('click', function () { tr.remove(); if (!dataRows().length) addRows(1); update(); });
    xd.appendChild(x); tr.appendChild(xd);
    rowsEl.appendChild(tr);
    return tr;
  }
  function addRows(n) { for (var i = 0; i < n; i++) addRow(); }
  function rowValues(tr) { return Array.prototype.map.call(tr.querySelectorAll('input'), function (i) { return i.value.trim(); }); }
  // All rows in order, without the empty ones at the end.
  function readRows() {
    var rows = dataRows().map(rowValues);
    while (rows.length && !rows[rows.length - 1].some(Boolean)) rows.pop();
    return rows;
  }
  function filledRows() { return readRows().filter(function (r) { return r.some(Boolean); }); }
  function setRows(rows) {
    var info = json(INFO) || {};
    rowsEl.innerHTML = '';
    rows.forEach(function (r) { addRow(r); });
    // Start with one row per open label (or 10 for a full sheet), plus a few spare.
    var want = rows.length ? Math.max(rows.length + 3, info.open || 0) : (info.open || 10);
    addRows(want - rows.length);
    update();
  }
  // Put new rows after the last row with something in it.
  function appendRows(rows) { setRows(readRows().concat(rows)); }

  function update() {
    var info = json(INFO) || {}, open = info.open, per = info.per || 30, labels = info.labels || [];
    var known = typeof open === 'number';
    Array.prototype.slice.call(rowsEl.querySelectorAll('tr.divider')).forEach(function (d) { d.remove(); });
    var filled = 0, last = -1;
    dataRows().forEach(function (tr, i) {
      if (rowValues(tr).some(Boolean)) { filled++; last = i; }
      var extra = known ? i - open : -1;
      tr.firstChild.textContent = !known ? String(i + 1) : extra < 0 ? String(labels[i] || i + 1) : String(extra % per + 1);
      tr.classList.toggle('spill', extra >= 0);
      if (extra >= 0 && extra % per === 0) {
        var d = document.createElement('tr'), td = document.createElement('td');
        d.className = 'divider';
        td.colSpan = COLS.length + 2;
        td.textContent = 'Extra sheet ' + (extra / per + 1) + ' · rows below don’t fit on this sheet';
        td.title = 'These go on your other part-used sheets of this size, or on a new sheet. The panel asks which.';
        d.appendChild(td);
        rowsEl.insertBefore(d, tr);
      }
    });
    var extraSlots = known ? Math.max(0, last + 1 - open) : 0;
    $('count').textContent = plural(filled, 'address', 'es') + (known ? ' · ' + plural(open, 'open label') + ' on your sheet' : '') +
      (extraSlots ? ' · ' + (extraSlots === 1 ? '1 row doesn’t' : extraSlots + ' rows don’t') + ' fit on it' : '');
    $('fill').disabled = !filled;
    store(DRAFT, JSON.stringify(readRows()));
  }

  // ---------- turning other data into rows ----------
  // Address fields with a contact's phone and email.
  function withContact(fields, c) { fields.phone = c.phone || ''; fields.email = c.email || ''; return fields; }
  function fieldsToRow(a) {
    if (a.extra && a.extra.length && !a.street && !a.city) {
      var f = T.linesToFields([a.name].concat(a.extra));
      f.company = a.company || f.company;
      f.phone = a.phone || ''; f.email = a.email || '';
      a = f;
    }
    return COLS.map(function (c) { return c === 'zip' ? T.fixZip(a.zip, a.state) : (a[c] || ''); });
  }
  function rowsFromGrid(grid) {
    // A header row is recognized by its names; otherwise columns go in order.
    if (T.headerMap(grid[0])) return T.rowsToAddresses(grid).map(fieldsToRow);
    return grid.map(function (r) { return COLS.map(function (c, i) { return String(r[i] === undefined ? '' : r[i]).trim(); }); });
  }

  // Paste: a block of cells fills the grid from the cell you're in.
  rowsEl.addEventListener('paste', function (e) {
    var target = e.target;
    if (!target || target.tagName !== 'INPUT') return;
    var text = e.clipboardData && e.clipboardData.getData('text/plain');
    if (!text || (text.indexOf('\t') < 0 && text.indexOf('\n') < 0)) return;   // a single value: normal paste
    e.preventDefault();
    var grid = T.parseDelimited(text.replace(/\r?\n$/, ''), text.indexOf('\t') >= 0 ? '\t' : ',');
    if (T.headerMap(grid[0])) { appendRows(rowsFromGrid(grid)); msg('Pasted ' + plural(grid.length - 1, 'row') + ' and matched the columns by their headings.'); return; }
    var tr = target.closest('tr'), startCol = Array.prototype.indexOf.call(tr.querySelectorAll('input'), target);
    var rows = dataRows(), rowIndex = rows.indexOf(tr);
    grid.forEach(function (cells, r) {
      var row = rows[rowIndex + r] || addRow();
      var inputs = row.querySelectorAll('input');
      cells.forEach(function (v, c) { if (inputs[startCol + c]) inputs[startCol + c].value = String(v).trim(); });
    });
    if (dataRows().length - rowIndex - grid.length < 3) addRows(5);
    update();
    msg('Pasted ' + plural(grid.length, 'row') + '.');
  });

  // ---------- filling in addresses ----------
  var S = T.makeSuggester({ key: function () { return store(GEO_KEY) || ''; }, notice: function (t) { msg(t, true); } });
  function geoKey() { return S.keyOk() ? (store(GEO_KEY) || '') : ''; }
  function suggestOn() { return store(SUGGEST) !== 'off'; }
  function cellOf(tr, field) { return tr.querySelectorAll('input')[COLS.indexOf(field)]; }
  function setIfEmpty(inp, v) { if (inp && v && !inp.value.trim()) inp.value = v; }

  // A ZIP or postal code fills in an empty City and State/Prov.
  var postalTimer = 0;
  function postalChanged(inp) {
    clearTimeout(postalTimer);
    var tr = inp.closest('tr'), code = inp.value.trim();
    if (!/^\d{5}(-\d{4})?$|^[A-Za-z]\d[A-Za-z][ -]?\d[A-Za-z]\d$/.test(code)) return;
    if (cellOf(tr, 'city').value.trim() && cellOf(tr, 'state').value.trim()) return;
    postalTimer = setTimeout(function () {
      T.lookupPostal(code, geoKey()).then(function (r) {
        if (!r || inp.value.trim() !== code) return;
        setIfEmpty(cellOf(tr, 'city'), r.city); setIfEmpty(cellOf(tr, 'state'), r.state);
        if (/^[A-Za-z]/.test(code)) inp.value = T.fixPostal(code);
        update();
      });
    }, 350);
  }

  // Address suggestions while typing a street: Geoapify with the user's own
  // key, else Geoapify through the Label Sheet Keeper server (its key is shared
  // by everyone), else the free Photon service.
  var sugg = $('sugg'), suggList = $('sugg-list'), suggFor = null, suggItems = [], suggActive = -1, suggTimer = 0, suggSeq = 0;
  var CREDITS = {
    photon: [['Photon by Komoot', 'https://photon.komoot.io/'], ['© OpenStreetMap contributors', 'https://www.openstreetmap.org/copyright']],
    geoapify: [['Powered by Geoapify', 'https://www.geoapify.com/'], ['© OpenStreetMap contributors', 'https://www.openstreetmap.org/copyright']]
  };
  function setCredit(source) {
    var c = $('sugg-credit');
    c.innerHTML = '';
    (CREDITS[source] || []).forEach(function (pair, i) {
      if (i) c.appendChild(document.createTextNode(' · '));
      var a = document.createElement('a'); a.href = pair[1]; a.target = '_blank'; a.rel = 'noopener'; a.textContent = pair[0];
      c.appendChild(a);
    });
  }
  function closeSugg() { sugg.hidden = true; suggFor = null; suggItems = []; suggActive = -1; }
  function showSugg(inp, items) {
    suggList.innerHTML = '';
    suggItems = items; suggActive = -1; suggFor = inp;
    if (!items.length) { closeSugg(); return; }
    items.forEach(function (it, i) {
      var b = document.createElement('button');
      b.type = 'button'; b.setAttribute('role', 'option'); b.tabIndex = -1;
      b.textContent = it.label;
      var sm = document.createElement('small'); sm.textContent = it.sub; b.appendChild(sm);
      b.addEventListener('mousedown', function (e) { e.preventDefault(); chooseSugg(i); });
      suggList.appendChild(b);
    });
    setCredit(items[0].source);
    var r = inp.getBoundingClientRect();
    sugg.style.left = Math.max(8, Math.min(r.left, window.innerWidth - 300)) + 'px';
    sugg.style.top = (r.bottom + 2) + 'px';
    sugg.hidden = false;
  }
  function markSugg(i) {
    suggActive = i;
    Array.prototype.forEach.call(suggList.children, function (b, j) { b.setAttribute('aria-selected', j === i ? 'true' : 'false'); });
  }
  function chooseSugg(i) {
    var it = suggItems[i], inp = suggFor;
    if (!it || !inp) return;
    var tr = inp.closest('tr');
    inp.value = it.street;
    if (it.city) cellOf(tr, 'city').value = it.city;
    if (it.state) cellOf(tr, 'state').value = it.state;
    if (it.zip) cellOf(tr, 'zip').value = it.zip;
    closeSugg();
    update();
    cellOf(tr, 'apt').focus();
  }
  function streetChanged(inp) {
    clearTimeout(suggTimer);
    var text = inp.value.trim();
    // Only ask once there's a house number and a few letters of the street.
    if (!suggestOn() || !S.ready(text)) { if (suggFor === inp) closeSugg(); return; }
    var tr = inp.closest('tr'), city = cellOf(tr, 'city').value.trim(), state = cellOf(tr, 'state').value.trim();
    var q = [text, city, state].filter(Boolean).join(', '), seq = ++suggSeq;
    // What's already known and still fits shows at once; a fresher answer follows.
    var now = S.quick(q);
    if (now.length && document.activeElement === inp) showSugg(inp, now);
    suggTimer = setTimeout(function () {
      S.ask(q).then(function (items) {
        if (seq === suggSeq && document.activeElement === inp) showSugg(inp, items);
      }).catch(function () { /* the suggester says what happened */ });
    }, S.delay());
  }
  rowsEl.addEventListener('input', function (e) {
    var inp = e.target;
    if (inp.tagName !== 'INPUT') return;
    var field = inp.getAttribute('aria-label');
    if (field === 'zip') postalChanged(inp);
    else if (field === 'street') streetChanged(inp);
  });
  rowsEl.addEventListener('focusout', function (e) { if (e.target === suggFor) setTimeout(function () { if (document.activeElement !== suggFor) closeSugg(); }, 120); });
  $('grid').parentNode.addEventListener('scroll', closeSugg);
  window.addEventListener('resize', closeSugg);

  // Arrow keys and Enter pick a suggestion; otherwise Enter moves down a row, like a spreadsheet.
  rowsEl.addEventListener('keydown', function (e) {
    if (!sugg.hidden && e.target === suggFor) {
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        e.preventDefault();
        var n = suggItems.length;
        markSugg(e.key === 'ArrowDown' ? (suggActive + 1) % n : (suggActive + n - 1) % n);
        return;
      }
      if (e.key === 'Enter' && suggActive >= 0) { e.preventDefault(); chooseSugg(suggActive); return; }
      if (e.key === 'Escape') { e.preventDefault(); closeSugg(); return; }
    }
    if (e.key !== 'Enter' || e.target.tagName !== 'INPUT') return;
    e.preventDefault();
    var tr = e.target.closest('tr'), col = Array.prototype.indexOf.call(tr.querySelectorAll('input'), e.target);
    var rows = dataRows(), next = rows[rows.indexOf(tr) + 1] || addRow();
    next.querySelectorAll('input')[col].focus();
    update();
  });
  rowsEl.addEventListener('input', update);

  // Open a file
  $('file').addEventListener('change', function () {
    var file = $('file').files && $('file').files[0]; $('file').value = '';
    if (!file) return;
    var name = (file.name || '').toLowerCase(), reader = new FileReader();
    var done = function (rows) {
      rows = rows.filter(function (r) { return r.some(Boolean); });
      if (!rows.length) { msg('No addresses were found in that file.', true); return; }
      appendRows(rows);
      msg('Added ' + plural(rows.length, 'row') + ' from ' + file.name + '. Check the columns before filling.');
    };
    if (/\.xlsx$/.test(name)) {
      reader.onload = function () { T.readXlsx(reader.result).then(function (grid) { done(rowsFromGrid(grid)); }).catch(function () { msg('That Excel file couldn’t be read. Try saving it as CSV.', true); }); };
      reader.readAsArrayBuffer(file);
    } else {
      reader.onload = function () {
        var text = String(reader.result || '');
        if (/\.vcf$/.test(name) || /BEGIN:VCARD/i.test(text)) done(T.parseVCard(text).map(function (c) { return fieldsToRow(withContact(c.fields, c)); }));
        else done(rowsFromGrid(T.parseDelimited(text)));
      };
      reader.readAsText(file);
    }
  });
  $('file-label').addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('file').click(); } });

  // Saved addresses
  function fillBookPicker() {
    var b = json(BOOK) || [], sel = $('book');
    sel.innerHTML = '';
    var o = document.createElement('option'); o.value = ''; o.textContent = b.length ? 'Add a saved address…' : 'No saved addresses';
    sel.appendChild(o);
    if (b.length) { var all = document.createElement('option'); all.value = '*'; all.textContent = 'Add all ' + b.length; sel.appendChild(all); }
    var groups = {};
    b.forEach(function (e) { (Array.isArray(e.groups) ? e.groups : []).forEach(function (g) { var k = g.toLowerCase(); if (!groups[k]) groups[k] = { name: g, n: 0 }; groups[k].n++; }); });
    Object.keys(groups).sort().forEach(function (k) {
      var op = document.createElement('option'); op.value = 'group:' + k; op.textContent = 'Add group: ' + groups[k].name + ' (' + groups[k].n + ')'; sel.appendChild(op);
    });
    b.forEach(function (e) { var op = document.createElement('option'); op.value = e.id; op.textContent = e.name; sel.appendChild(op); });
    sel.disabled = !b.length;
  }
  $('book').addEventListener('change', function () {
    var b = json(BOOK) || [], v = $('book').value;
    var g = v.indexOf('group:') === 0 ? v.slice(6) : null;
    var inG = function (e) { return (Array.isArray(e.groups) ? e.groups : []).some(function (x) { return x.toLowerCase() === g; }); };
    var pick = v === '*' ? b : g !== null ? b.filter(inG) : b.filter(function (e) { return e.id === v; });
    if (pick.length) { appendRows(pick.map(function (e) { return fieldsToRow(withContact(T.linesToFields(e.show && e.show.addr === 2 && e.lines2 ? e.lines2 : e.lines), e)); })); msg('Added ' + plural(pick.length, 'saved address', 'es') + '.'); }
    $('book').value = '';
  });

  $('more').addEventListener('click', function () { addRows(10); update(); });
  $('clear').addEventListener('click', function () {
    if ($('clear').dataset.sure !== '1') { $('clear').dataset.sure = '1'; $('clear').textContent = 'Click again to clear'; setTimeout(function () { $('clear').dataset.sure = ''; $('clear').textContent = 'Clear'; }, 3000); return; }
    $('clear').dataset.sure = ''; $('clear').textContent = 'Clear';
    setRows([]); msg('Cleared.');
  });

  function send(message) {
    var ui = window.Office && Office.context && Office.context.ui;
    if (ui && typeof ui.messageParent === 'function') ui.messageParent(JSON.stringify(message));
    else msg('This window works with the Label Sheet Keeper panel in Word.', true);
  }
  $('fill').addEventListener('click', function () {
    var rows = readRows();
    if (!filledRows().length) { msg('Add at least one address first.', true); return; }
    store(DRAFT, '[]');
    send({ type: 'fill', rows: rows, show: { phone: $('on-phone').checked, email: $('on-email').checked } });
  });
  $('cancel').addEventListener('click', function () { send({ type: 'close' }); });

  // ---------- start ----------
  function start() {
    var info = json(INFO) || {};
    $('info').textContent = (info.name ? info.name + '. ' : '') + 'One row per label; the first column shows which label. Type addresses, paste cells from Excel or Numbers, or open a file. ' +
      (suggestOn() ? 'Type a house number and street for address suggestions.' : 'A ZIP or postal code fills in the city.');
    var on = info.onLabel || {};
    $('on-phone').checked = !!on.phone; $('on-email').checked = !!on.email;
    var draft = json(DRAFT);
    setRows(Array.isArray(draft) ? draft : []);
    fillBookPicker();
    var first = rowsEl.querySelector('input'); if (first) first.focus();
  }
  if (window.Office && Office.onReady) Office.onReady(function () { start(); });
  else start();
  window.LSK_TABLE = { readRows: filledRows, rows: readRows, setRows: setRows };
})();
