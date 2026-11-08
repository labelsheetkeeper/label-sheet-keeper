/* A small in-memory stand-in for Office.js / Word, enough for the panel's flows.
   Operations apply immediately; load() and sync() are no-ops. */
(function () {
  'use strict';
  var params = new URLSearchParams(location.search);
  var M = window.MOCK = {
    settings: {}, props: {}, header: '', lastPrint: 0, selection: null, created: [], saved: 0, url: params.has('unsaved') ? '' : (params.get('url') || '/Users/test/Labels.docx'),
    dialog: null, blocks: [], calls: [], opened: [], browser: []
  };

  function para(text, level) {
    return { text: text || '', level: level || 0, alignment: 'Left', spaceBefore: 0, spaceAfter: 0, leftIndent: 0, rightIndent: 0, firstLineIndent: 0, lineSpacing: 12,
      font: { name: null, size: null, bold: false, italic: false, color: '#000000' } };
  }
  function newCell() { return { paras: [para('', 1)], ccs: [], pics: 0 }; }
  M.newTable = function (rows, cols) {
    var t = { type: 'table', rows: [] };
    for (var r = 0; r < rows; r++) { var row = []; for (var c = 0; c < cols; c++) row.push(newCell()); t.rows.push(row); }
    return t;
  };
  M.table = function () { return M.blocks.filter(function (b) { return b.type === 'table'; })[0]; };
  M.cellText = function (cell) { return cell.paras.map(function (p) { return p.text; }).join('\n'); };

  M.init = function () {
    var F = window.LSK_FORMATS, S = window.LSK_SHEET;
    if (params.get('backup')) { M.props[S.BACKUP_PROP] = '1'; M.blocks = [{ type: 'para', p: para('Label Sheet Keeper backup') }]; return; }
    if (params.get('blank')) return;
    var f = F.getFormat(params.get('fmt') || 'avery-5160');
    var legacy = params.get('legacy');
    M.props[S.PROP] = f.id;
    var L = legacy ? S.layout(f, { margins: { top: f.firstLabel.top, left: f.firstLabel.left } }) : S.layout(f);
    if (!legacy) {
      M.props[S.LAYOUT_PROP] = JSON.stringify({ v: 2, top: L.margins.top, left: L.margins.left, nx: 0, ny: 0 });
      var spec = JSON.stringify(S.specOf(f));
      for (var i = 0; i * 240 < spec.length; i++) M.props[S.SPEC_PROP + (i + 1)] = spec.slice(i * 240, i * 240 + 240);
    }
    if (L.spacer) M.blocks.push({ type: 'spacer', cc: { tag: S.SPACER_TAG, cannotDelete: true, cannotEdit: true } });
    M.blocks.push(M.newTable(L.rows.length, L.cols.length));
    M.blocks.push({ type: 'para', p: para('') });
  };
  M.labelCell = function (n) {
    var S = window.LSK_SHEET, F = window.LSK_FORMATS, f = F.getFormat(M.props[S.PROP]) || M.fmt;
    var ci = S.cellOf(f, n); return M.table().rows[ci.row][ci.cell];
  };
  M.setLabelText = function (n, text, size) {
    var c = M.labelCell(n);
    c.ccs = []; c.pics = 0;
    c.paras = text.split('\n').map(function (t) { var p = para(t, 1); p.font.size = size || 10; p.font.name = 'Arial'; return p; });
  };

  // A 2x2 grey PNG.
  M.GREY_PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAADklEQVR4nGPoAAMGCAUALK4GYdPeFKwAAAAASUVORK5CYII=';

  // ---------- object wrappers ----------
  function collection(items, extra) {
    var c = { items: items, load: function () { return c; } };
    c.getFirst = function () { return items[0]; };
    c.getLast = function () { return items[items.length - 1]; };
    Object.keys(extra || {}).forEach(function (k) { c[k] = extra[k]; });
    return c;
  }
  function paraObj(p, owner) {
    var o = {
      load: function () { return o; },
      get text() { return p.text; },
      get tableNestingLevel() { return p.level; },
      get alignment() { return p.alignment; }, set alignment(v) { p.alignment = v; },
      get spaceBefore() { return p.spaceBefore; }, set spaceBefore(v) { p.spaceBefore = v; },
      get spaceAfter() { return p.spaceAfter; }, set spaceAfter(v) { p.spaceAfter = v; },
      get leftIndent() { return p.leftIndent; }, set leftIndent(v) { p.leftIndent = v; },
      get firstLineIndent() { return p.firstLineIndent; }, set firstLineIndent(v) { p.firstLineIndent = v; },
      get rightIndent() { return p.rightIndent; }, set rightIndent(v) { p.rightIndent = v; },
      get lineSpacing() { return p.lineSpacing; }, set lineSpacing(v) { p.lineSpacing = v; },
      font: {
        get name() { return p.font.name; }, set name(v) { p.font.name = v; },
        get size() { return p.font.size; }, set size(v) { p.font.size = v; },
        get bold() { return p.font.bold; }, set bold(v) { p.font.bold = v; },
        get italic() { return p.font.italic; }, set italic(v) { p.font.italic = v; },
        get color() { return p.font.color; }, set color(v) { p.font.color = v; },
        set: function (o2) { Object.keys(o2).forEach(function (k) { p.font[k] = o2[k]; }); }
      },
      insertText: function (t, where) { p.text = where === 'Replace' ? t : (where === 'Start' ? t + p.text : p.text + t); return rangeObj(owner, p); },
      insertInlinePictureFromBase64: function (b64) { var img = { width: 0, height: 0, b64: b64, altTextDescription: '' }; if (owner && owner.cell) { owner.cell.pics++; p.pic = img; } return img; },
      insertParagraph: function (t, where) {
        if (!owner || !owner.cell) throw new Error('mock: insertParagraph outside a cell');
        var list = owner.cell.paras, i = list.indexOf(p), q = para(t, p.level);
        // Word for Mac joins a paragraph added after a cell's last one onto it.
        if (where === 'After' && i === list.length - 1) { p.text += t; list.push(para('', p.level)); M.calls.push('para.insertParagraph.AfterLast'); return paraObj(p, owner); }
        list.splice(where === 'Before' ? i : i + 1, 0, q);
        return paraObj(q, owner);
      },
      getRange: function () { return { select: function () { M.calls.push('select'); } }; },
      delete: function () { if (owner && owner.remove) owner.remove(p); }
    };
    return o;
  }
  function rangeObj(owner, p) {
    return {
      insertContentControl: function () {
        var cc = { tag: '', title: '', cannotEdit: false, cannotDelete: false, appearance: '', color: '' };
        if (owner && owner.cell) owner.cell.ccs.push(cc);
        return ccObj(cc, owner && owner.cell);
      }
    };
  }
  function findCellPos(cell) {
    var t = M.table(); if (!t) return null;
    for (var r = 0; r < t.rows.length; r++) for (var c = 0; c < t.rows[r].length; c++) if (t.rows[r][c] === cell) return { row: r, cell: c };
    return null;
  }
  function cellRef(cell) {
    var pos = cell ? findCellPos(cell) : null;
    return { isNullObject: !pos, rowIndex: pos ? pos.row : -1, cellIndex: pos ? pos.cell : -1, load: function () { return this; } };
  }
  function ccObj(cc, cell) {
    var o = {
      get tag() { return cc.tag; }, set tag(v) { cc.tag = v; },
      get title() { return cc.title; }, set title(v) { cc.title = v; },
      get cannotEdit() { return cc.cannotEdit; }, set cannotEdit(v) { cc.cannotEdit = v; },
      get cannotDelete() { return cc.cannotDelete; }, set cannotDelete(v) { cc.cannotDelete = v; },
      set appearance(v) { cc.appearance = v; }, set color(v) { cc.color = v; },
      get parentTableCellOrNullObject() { return cellRef(cell); },
      delete: function () { if (cc.cannotDelete) throw new Error('Locked'); if (cell) cell.ccs = cell.ccs.filter(function (x) { return x !== cc; }); }
    };
    return o;
  }
  function serializeCell(cell) {
    var data = JSON.stringify({ paras: cell.paras, ccs: cell.ccs, pics: cell.pics });
    return '<pkg:package><!--MOCK:' + btoa(unescape(encodeURIComponent(data))) + '-->' + cell.paras.map(function () { return '<w:p>'; }).join('') + '</pkg:package>';
  }
  function cellBodyObj(cell) {
    var owner = { cell: cell, remove: function (p) { cell.paras = cell.paras.filter(function (x) { return x !== p; }); if (!cell.paras.length) cell.paras.push(para('', 1)); } };
    var b = {
      clear: function () {
        if (cell.ccs.some(function (c) { return c.cannotDelete; })) throw new Error('Can’t clear: a locked content control is in the way');
        cell.paras = [para('', 1)]; cell.ccs = []; cell.pics = 0;
      },
      insertText: function (t, where) { var p = cell.paras[0]; p.text = where === 'Start' ? t + p.text : p.text + t; return rangeObj(owner, p); },
      // Word for Mac joins a paragraph added at the end of a cell onto the last one.
      insertParagraph: function (t, where) {
        if (where === 'Start') { var p = para(t, 1); cell.paras.unshift(p); return paraObj(p, owner); }
        var last = cell.paras[cell.paras.length - 1]; last.text += t; cell.paras.push(para('', 1)); M.calls.push('cell.insertParagraph.End');
        return paraObj(last, owner);
      },
      insertInlinePictureFromBase64: function () { cell.pics++; return { width: 0, height: 0 }; },
      // Word returns the whole table row for a cell body.
      getOoxml: function () { return { value: '<w:tbl><w:tr>' + serializeCell(cell) + '</w:tr></w:tbl>' }; },
      getRange: function () { return { getOoxml: function () { return { value: serializeCell(cell) }; } }; },
      insertOoxml: function (xml) {
        M.calls.push('cell.insertOoxml');
        var m = /<!--MOCK:([^-]+)-->/.exec(xml);
        if (/<w:tbl>/.test(xml)) { cell.nested = true; M.calls.push('cell.nestedTable'); }
        if (m) {
          var data = JSON.parse(decodeURIComponent(escape(atob(m[1]))));
          cell.paras = data.paras.map(function (p) { return JSON.parse(JSON.stringify(p)); });
          // Word for Mac: the last pasted paragraph takes the layout of the paragraph it lands in.
          var lp = cell.paras[cell.paras.length - 1], d = para('', 1);
          if (lp) ['alignment', 'spaceBefore', 'spaceAfter', 'leftIndent', 'firstLineIndent'].forEach(function (k) { lp[k] = d[k]; });
          cell.ccs = data.ccs; cell.pics = data.pics;
        } else {
          var texts = []; xml.replace(/<w:t[^>]*>([^<]*)<\/w:t>/g, function (_, t) { texts.push(t); });
          cell.paras = (texts.length ? texts : ['']).map(function (t) { return para(t, 1); });
        }
      },
      get paragraphs() { return collection(cell.paras.map(function (p) { return paraObj(p, owner); })); },
      // The label's pictures: ones placed through a paragraph keep their data; others are a small grey square.
      get inlinePictures() {
        var list = cell.paras.filter(function (p) { return p.pic; }).map(function (p) { return p.pic; });
        while (list.length < cell.pics) list.push({ width: 72, height: 72, b64: M.GREY_PNG });
        return collection(list.slice(0, cell.pics).map(function (img) {
          return { load: function () { return this; }, width: img.width, height: img.height, getBase64ImageSrc: function () { return { value: img.b64 || M.GREY_PNG }; } };
        }));
      },
      get contentControls() { return collection(cell.ccs.map(function (c) { return ccObj(c, cell); })); },
      get text() { return M.cellText(cell); },
      load: function () { return b; }
    };
    return b;
  }
  function tableObj(t) {
    var o = {
      load: function () { return o; },
      get rowCount() { return t.rows.length; },
      get values() { return t.rows.map(function (r) { return r.map(function (c) { return M.cellText(c); }); }); },
      get rows() { return collection(t.rows.map(function (r) { return { cellCount: r.length }; })); },
      getCell: function (r, c) {
        if (!t.rows[r] || !t.rows[r][c]) throw new Error('No such cell ' + r + ',' + c);
        return { body: cellBodyObj(t.rows[r][c]) };
      }
    };
    return o;
  }
  function allBodyParas() {
    var out = [];
    M.blocks.forEach(function (b) {
      if (b.type === 'spacer') out.push({ p: para('', 0), block: b });
      else if (b.type === 'para') out.push({ p: b.p, block: b });
      else if (b.type === 'table') b.rows.forEach(function (r) { r.forEach(function (c) { c.paras.forEach(function (p) { out.push({ p: p, cell: c }); }); }); });
    });
    return out;
  }
  function bodyObj() {
    var body = {
      get tables() { return collection(M.blocks.filter(function (b) { return b.type === 'table'; }).map(tableObj)); },
      get paragraphs() {
        return collection(allBodyParas().map(function (x) {
          return paraObj(x.p, { remove: function () { if (x.block) M.blocks = M.blocks.filter(function (b) { return b !== x.block; }); } });
        }));
      },
      get contentControls() {
        var list = [];
        M.blocks.forEach(function (b) {
          if (b.type === 'spacer') list.push(ccObj(b.cc, null));
          if (b.type === 'table') b.rows.forEach(function (r) { r.forEach(function (c) { c.ccs.forEach(function (cc) { list.push(ccObj(cc, c)); }); }); });
        });
        return collection(list);
      },
      get inlinePictures() {
        var list = [], t = M.table();
        if (t) t.rows.forEach(function (r) { r.forEach(function (c) { for (var i = 0; i < c.pics; i++) list.push({ get parentTableCellOrNullObject() { return cellRef(c); } }); }); });
        return collection(list);
      },
      clear: function () {
        var locked = false;
        M.blocks.forEach(function (b) {
          if (b.type === 'spacer' && b.cc.cannotDelete) locked = true;
          if (b.type === 'table') b.rows.forEach(function (r) { r.forEach(function (c) { if (c.ccs.some(function (x) { return x.cannotDelete; })) locked = true; }); });
        });
        if (locked) throw new Error('Can’t clear: a locked content control is in the way');
        M.blocks = [{ type: 'para', p: para('') }];
      },
      insertOoxml: function (xml, where) {
        M.calls.push('body.insertOoxml');
        var add = [];
        if (/lsk-spacer/.test(xml)) add.push({ type: 'spacer', cc: { tag: 'lsk-spacer', cannotDelete: false, cannotEdit: false }, height: /w:line="(\d+)"/.exec(xml)[1] });
        if (/<w:tbl>/.test(xml)) {
          var rows = (xml.match(/<w:tr>/g) || []).length, cols = (xml.match(/<w:gridCol /g) || []).length;
          var t = M.newTable(rows, cols); t.indent = Number((/<w:tblInd w:w="(-?\d+)"/.exec(xml) || [0, 0])[1]); add.push(t);
        }
        add.push({ type: 'para', p: para('') });
        M.blocks = where === 'Start' ? add.concat(M.blocks) : M.blocks.concat(add);
      }
    };
    return body;
  }

  window.Word = {
    run: function (cb) {
      var ctx = {
        sync: function () { return Promise.resolve(); },
        document: {
          get body() { return bodyObj(); },
          properties: {
            load: function () { return this; },
            get lastPrintDate() { return M.lastPrint ? new Date(M.lastPrint) : new Date(0); },
            customProperties: {
              load: function () { return this; },
              get items() { return Object.keys(M.props).map(function (k) { return { key: k, value: M.props[k] }; }); },
              add: function (k, v) { M.props[k] = v; }
            }
          },
          sections: {
            getFirst: function () {
              return {
                getHeader: function () {
                  if (!M.hp) { M.hp = para(M.header); M.hp.font.size = 12; }
                  M.hp.text = M.header;
                  var h = {
                    load: function () { return h; },
                    get text() { return M.header; },
                    clear: function () { M.header = ''; M.hp = para(''); M.hp.font.size = 12; },
                    get paragraphs() {
                      var o = paraObj(M.hp, null), orig = o.insertText;
                      o.insertText = function (t, w) { var r = orig(t, w); M.header = M.hp.text; return r; };
                      return collection([o]);
                    }
                  };
                  return h;
                },
                getFooter: function () {
                  if (!M.fp) { M.fp = para(''); M.fp.font.size = 12; }
                  var f = {
                    load: function () { return f; },
                    clear: function () { M.fp = para(''); M.fp.font.size = 12; },
                    get paragraphs() { return collection([paraObj(M.fp, null)]); }
                  };
                  return f;
                }
              };
            }
          },
          // M.selection: { row, cell, para } puts the cursor in that cell, on that paragraph (default the first).
          getSelection: function () {
            var selCell = function () { var s = M.selection, t = M.table(); return s && t && t.rows[s.row] ? t.rows[s.row][s.cell] : null; };
            return {
              get parentTableCellOrNullObject() {
                var s = M.selection, c = selCell();
                return s ? { isNullObject: false, rowIndex: s.row, cellIndex: s.cell, body: c ? cellBodyObj(c) : null, load: function () { return this; } } : { isNullObject: true, load: function () { return this; } };
              },
              paragraphs: {
                getFirst: function () {
                  var c = selCell();
                  if (!c) return paraObj(para('', 0), null);
                  return paraObj(c.paras[Math.min(M.selection.para || 0, c.paras.length - 1)], { cell: c });
                }
              }
            };
          },
          save: function () { M.saved++; }
        },
        application: {
          createDocument: function (b64) { var d = { b64: b64, opened: false, open: function () { d.opened = true; } }; M.created.push(d); return d; },
          openDocument: function (url) { if (/missing/.test(url)) throw new Error('mock: file not found'); if (/declined/.test(url)) throw new Error('mock: the user said no'); M.opened.push(url); }
        }
      };
      try { return Promise.resolve(cb(ctx)); } catch (e) { return Promise.reject(e); }
    }
  };

  window.Office = {
    HostType: { Word: 'Word' },
    EventType: { DocumentSelectionChanged: 'sel', DialogMessageReceived: 'msg', DialogEventReceived: 'evt' },
    AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
    FileType: { Compressed: 'compressed' },
    onReady: function (cb) { setTimeout(function () { M.init(); cb({ host: 'Word' }); }, 0); },
    context: {
      requirements: { isSetSupported: function () { return true; } },
      document: {
        get url() { return M.url; },
        settings: {
          get: function (k) { return M.settings[k] === undefined ? null : JSON.parse(JSON.stringify(M.settings[k])); },
          set: function (k, v) { M.settings[k] = JSON.parse(JSON.stringify(v)); },
          saveAsync: function (cb) { if (cb) cb({ status: 'succeeded' }); }
        },
        addHandlerAsync: function (type, fn) { M.selHandler = fn; },
        getFilePropertiesAsync: function (cb) { setTimeout(function () { cb({ status: 'succeeded', value: { url: M.url } }); }, 0); },
        customXmlParts: {
          getByNamespaceAsync: function (ns, cb) {
            var xml = window.MOCK_CUSTOM_XML;
            setTimeout(function () {
              cb({ status: 'succeeded', value: xml ? [{ getXmlAsync: function (c2) { c2({ status: 'succeeded', value: xml }); } }] : [] });
            }, 0);
          }
        },
        getFileAsync: function (type, opts, cb) {
          var zip = new window.JSZip();
          zip.file('word/webextensions/webextension1.xml', '<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11"><we:reference id="8d491c85-a76d-4d26-a6f5-b7b2dd4a1711" version="1.1.0.0" store="developer" storeType="Registry"/></we:webextension>');
          zip.generateAsync({ type: 'uint8array' }).then(function (bytes) {
            cb({ status: 'succeeded', value: { sliceCount: 1, getSliceAsync: function (i, c2) { c2({ status: 'succeeded', value: { data: bytes } }); }, closeAsync: function () {} } });
          });
        }
      },
      ui: {
        openBrowserWindow: function (url) { M.browser.push(url); },
        displayDialogAsync: function (url, opts, cb) {
          var handlers = {};
          M.dialog = { url: url, closed: false, handlers: handlers,
            addEventHandler: function (t, h) { handlers[t] = h; }, close: function () { M.dialog.closed = true; },
            send: function (obj) { handlers.msg({ message: JSON.stringify(obj) }); } };
          cb({ status: 'succeeded', value: M.dialog });
        }
      }
    }
  };
})();
